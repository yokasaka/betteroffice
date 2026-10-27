import { expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { createSessionClient } from '../../../../shared/office-session/client';
import { createSessionHost } from '../../../../shared/office-session/host';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import * as wasm from '../wasm/loader';
import type { TablePrimitive } from '../types';
import { registerPresentationCommandHandlers } from './commandHost';
import { PRESENTATION_COMMAND_METHODS, type PresentationCommandAck, type PresentationCommandEnvelope,
  type PresentationCommandEvents, type PresentationCommandIntent, type PresentationCommandMethods } from './commandProtocol';

const identity = { session: 'table-session', incarnation: 'table-worker' };
const input = { inputTime: 1, undoGroup: 'table-session' };

async function open() {
  const pair = createInProcessPair();
  const registration = registerPresentationCommandHandlers({ initWasm: async () => {
    await wasm.initWasm(await readFile(new URL('../wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
  } });
  let handle: wasm.PresentationHandle;
  const original = wasm.openPresentation;
  const opened = spyOn(wasm, 'openPresentation').mockImplementation((...args) => {
    handle = original(...args);
    return handle;
  });
  createSessionHost<PresentationCommandMethods, PresentationCommandEvents, null>(pair.host, {
    handlers: registration.handlers, policies: registration.policies, context: null, onDispose: registration.dispose,
  });
  const rpc = createSessionClient<PresentationCommandMethods, PresentationCommandEvents>(pair.client,
    { methods: PRESENTATION_COMMAND_METHODS });
  try {
    const bytes = await readFile(new URL('../../../../crates/pptx-render/tests/fixtures/table-basic.pptx', import.meta.url));
    const font = await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
    const state = await rpc.call.openCommandEditor(new Uint8Array(bytes).buffer,
      { ...identity, fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font).buffer }] });
    const slide = handle!.snapshot().slides[0]!;
    const table = slide.shapes.find((shape) => shape.kind === 'graphicFrame')!;
    return { state, rpc, slideId: slide.id, shapeId: table.id, handle: () => handle!,
      command(state: PresentationCommandAck, seq: number, intent: unknown) {
        return rpc.call.command({ ...identity, seq, basis: { epoch: state.epoch, version: state.version },
          ...(intent as PresentationCommandIntent) } as PresentationCommandEnvelope);
      },
      async dispose() { try { await rpc.dispose(); } finally { opened.mockRestore(); registration.dispose(); pair.client.close(); } },
    };
  } catch (error) {
    opened.mockRestore();
    registration.dispose();
    pair.client.close();
    throw error;
  }
}

function sizes(handle: wasm.PresentationHandle) {
  const shape = handle.snapshot().slides[0]!.shapes.find((entry) => entry.kind === 'graphicFrame')!;
  const table = (shape as { graphic?: { grid: number[]; rows: { height: number }[] } }).graphic!;
  return { widths: table.grid, heights: table.rows.map((row) => row.height) };
}

test('set_table_sizes_command_applies_undoes_redoes_and_saves', async () => {
  const session = await open();
  try {
    const before = sizes(session.handle());
    const widths = [2000000, 1000000, 2500000];
    const heights = [400000, 500000, 600000];
    const applied = await session.command(session.state, 1, { kind: 'invoke', payload: {
      ...input, method: 'setTableSizes', args: [session.slideId, session.shapeId, widths, heights] } });
    expect(applied.outcomes).toEqual([{ seq: 1, status: 'applied', committed: true }]);
    expect(applied.editState.canUndo).toBe(true);
    expect(sizes(session.handle())).toEqual({ widths, heights });
    const undone = await session.command(applied, 2, { kind: 'undo', payload: input });
    expect(sizes(session.handle())).toEqual(before);
    const redone = await session.command(undone, 3, { kind: 'redo', payload: input });
    expect(redone.editState.canUndo).toBe(true);
    expect(sizes(session.handle())).toEqual({ widths, heights });
    const saved = await session.rpc.call.saveCommandEditor({ ...identity, seq: 3 });
    const slide = await (await JSZip.loadAsync(saved.bytes)).file('ppt/slides/slide1.xml')!.async('string');
    expect(slide).toContain('<a:gridCol w="2000000"/><a:gridCol w="1000000"/><a:gridCol w="2500000"/>');
    expect(slide).toContain('<a:tr h="600000">');
    expect(slide).toContain('<a:ext cx="5500000" cy="1500000"/>');
  } finally { await session.dispose(); }
});

test('layout_reports_the_table_grid_lines_after_set_table_sizes', async () => {
  const session = await open();
  try {
    const table = () => session.handle().layoutSlide(0).primitives.find((entry) => entry.kind === 'table') as TablePrimitive;
    const before = table();
    expect(before.columns).toHaveLength(4);
    expect(before.rows).toHaveLength(4);
    expect([before.columns![0], before.rows![0]]).toEqual([before.x, before.y]);
    await session.command(session.state, 1, { kind: 'invoke', payload: {
      ...input, method: 'setTableSizes', args: [session.slideId, session.shapeId, [2000000, 1000000, 2500000], [400000, 500000, 600000]] } });
    const after = table();
    // 9525 EMU per slide pixel
    expect(after.columns!.map((x) => Math.round(x - after.columns![0]!))).toEqual([0, 210, 315, 577]);
    expect(after.rows).toHaveLength(4);
    expect(after.rows!.every((y, i) => i === 0 || y > after.rows![i - 1]!)).toBe(true);
  } finally { await session.dispose(); }
});

test.each([
  ['a count mismatch', [1, 1], [1, 1, 1]],
  ['all zero widths', [0, 0, 0], [1, 1, 1]],
  ['a negative height', [1, 1, 1], [1, -1, 1]],
  ['a fractional height', [1, 1, 1], [1, 1.5, 1]],
  ['a non-array', 'wide', [1, 1, 1]],
])('set_table_sizes_refuses_%s_without_changes', async (_, widths, heights) => {
  const session = await open();
  try {
    const before = sizes(session.handle());
    const ack = await session.command(session.state, 1, { kind: 'invoke', payload: {
      ...input, method: 'setTableSizes', args: [session.slideId, session.shapeId, widths, heights] } });
    expect(ack.outcomes[0]).toMatchObject({ seq: 1, status: 'refused', committed: false });
    expect(ack.version).toBe(session.state.version);
    expect(ack.editState.canUndo).toBe(false);
    expect(sizes(session.handle())).toEqual(before);
    expect(session.rpc.failure).toBeUndefined();
  } finally { await session.dispose(); }
});
