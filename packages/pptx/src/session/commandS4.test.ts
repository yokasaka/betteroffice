import { expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { createSessionClient } from '../../../../shared/office-session/client';
import { createSessionHost } from '../../../../shared/office-session/host';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import type { PptxStoryText } from '../edits';
import * as wasm from '../wasm/loader';
import { registerPresentationCommandHandlers } from './commandHost';
import { PRESENTATION_COMMAND_METHODS, type PresentationCommandAck, type PresentationCommandEnvelope,
  type PresentationCommandEvents, type PresentationCommandIntent, type PresentationCommandMethods,
  type PresentationCommandSelection } from './commandProtocol';

const identity = { session: 's4', incarnation: 's4-worker' };
const input = { inputTime: 1, undoGroup: 'typing' };

function selection(story: PptxStoryText, anchor = 0, focus = anchor): PresentationCommandSelection {
  return { slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId,
    anchor, focus, affinity: 'downstream' };
}

async function deck(): Promise<Uint8Array> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ns = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${rel}"`;
  const ids = [1, 2];
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${ids.map((id) => `<Override PartName="/ppt/slides/slide${id}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')}</Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldIdLst>${ids.map((id) => `<p:sldId id="${255 + id}" r:id="rId${id}"/>`).join('')}</p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${ids.map((id) => `<Relationship Id="rId${id}" Type="${rel}/slide" Target="slides/slide${id}.xml"/>`).join('')}</Relationships>`);
  for (const id of ids) zip.file(`ppt/slides/slide${id}.xml`, `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="6000000" cy="1200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:latin typeface="Liberation Sans"/></a:rPr><a:t>alpha</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  return zip.generateAsync({ type: 'uint8array' });
}

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
    const font = await readFile(new URL('../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
    const state = await rpc.call.openCommandEditor(new Uint8Array(await deck()).buffer,
      { ...identity, fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font).buffer }] });
    const result = (await rpc.call.readEditor({ ...identity, seq: 0 })).result;
    if (!result.ok) throw new Error(result.failure.message);
    return { state, stories: result.stories, handle: () => handle!,
      command(state: PresentationCommandAck, seq: number, intent: unknown, afterSeq?: number) {
        return rpc.call.command({ ...identity, seq, basis: { epoch: state.epoch, version: state.version },
          ...(afterSeq === undefined ? {} : { afterSeq }), ...(intent as PresentationCommandIntent) } as PresentationCommandEnvelope);
      },
      read: () => rpc.call.readEditor({ ...identity, seq: 0 }),
      thumbnail: (slideId: string, knownRevision?: number) => rpc.call.thumbnailFrame({ ...identity, slideId, knownRevision }),
      async dispose() { try { await rpc.dispose(); } finally { opened.mockRestore(); registration.dispose(); pair.client.close(); } },
    };
  } catch (error) {
    opened.mockRestore(); registration.dispose(); pair.client.close();
    throw error;
  }
}

test('typing_group', async () => {
  const session = await open();
  try {
    const story = session.stories[0]!;
    const first = await session.command(session.state, 1, { kind: 'insertText', payload: {
      ...input, at: { source: 'confirmed', selection: selection(story) }, text: 'A' } });
    const rapid = await session.command(first, 2, { kind: 'insertText', payload: {
      ...input, inputTime: 100, at: { source: 'afterSeq', seq: 1 }, text: 'B' } }, 1);
    const paused = await session.command(rapid, 3, { kind: 'insertText', payload: {
      ...input, inputTime: 1000, at: { source: 'afterSeq', seq: 2 }, text: 'C' } }, 2);
    const undone = await session.command(paused, 4, { kind: 'undo', payload: input });
    expect(undone.outcomes).toEqual([{ seq: 4, status: 'applied', committed: true }]);
    expect(session.handle().story(story.storyId).paragraphs[0]!.runs.map((run) => run.text).join('')).toBe('ABalpha');
    const all = await session.command(undone, 5, { kind: 'undo', payload: input });
    expect(session.handle().story(story.storyId).paragraphs[0]!.runs.map((run) => run.text).join('')).toBe('alpha');
    expect(all.editState).toMatchObject({ canUndo: false, canRedo: true });
    const redone = await session.command(all, 6, { kind: 'redo', payload: input });
    expect(session.handle().story(story.storyId).paragraphs[0]!.runs.map((run) => run.text).join('')).toBe('ABalpha');
    expect(redone.editState).toMatchObject({ canUndo: true, canRedo: true });
  } finally { await session.dispose(); }
});

test('empty_history_keeps_typing_selection_and_marks', async () => {
  const session = await open();
  try {
    const at = selection(session.stories[0]!, 2);
    const selected = await session.command(session.state, 1, { kind: 'setSelection', payload: {
      ...input, selection: at } });
    const formatted = await session.command(selected, 2, { kind: 'formatText', payload: {
      ...input, at: { source: 'confirmed', selection: at }, patch: { bold: true } } });
    const undone = await session.command(formatted, 3, { kind: 'undo', payload: input });
    expect(undone.outcomes[0]).toMatchObject({ status: 'noop', committed: false });
    expect(undone.selection).toEqual(at);
    expect(undone.editState.toolbar.storedMarks).toEqual({ bold: true });
    expect(undone.dirtySlideIds).toEqual([]);
    const redone = await session.command(undone, 4, { kind: 'redo', payload: input });
    expect(redone.outcomes[0]).toMatchObject({ status: 'noop', committed: false });
    expect(redone.selection).toEqual(at);
    expect(redone.editState.toolbar.storedMarks).toEqual({ bold: true });
    const typed = await session.command(redone, 5, { kind: 'insertText', payload: {
      ...input, at: { source: 'afterSeq', seq: 4 }, text: 'Z' } }, 4);
    expect(typed.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    const story = session.handle().story(at.storyId);
    expect(story.paragraphs[0]!.runs.map((run) => run.text).join('')).toBe('alZpha');
    expect(story.paragraphs[0]!.runs.find((run) => run.text === 'Z')?.style.bold).toBe(true);
  } finally { await session.dispose(); }
});

test('empty_redo_keeps_typing_group', async () => {
  const session = await open();
  try {
    const first = await session.command(session.state, 1, { kind: 'insertText', payload: {
      ...input, at: { source: 'confirmed', selection: selection(session.stories[0]!) }, text: 'A' } });
    const redone = await session.command(first, 2, { kind: 'redo', payload: input });
    expect(redone.outcomes[0]).toMatchObject({ status: 'noop', committed: false });
    const typed = await session.command(redone, 3, { kind: 'insertText', payload: {
      ...input, inputTime: 100, at: { source: 'afterSeq', seq: 2 }, text: 'B' } }, 2);
    await session.command(typed, 4, { kind: 'undo', payload: input });
    expect(session.handle().story(session.stories[0]!.storyId).paragraphs[0]!.runs.map((run) => run.text).join('')).toBe('alpha');
  } finally { await session.dispose(); }
});

test('discrete_groups', async () => {
  const session = await open();
  try {
    const story = session.stories[0]!;
    let state = session.state as PresentationCommandAck;
    let seq = 0;
    const command = async (intent: unknown) => { state = await session.command(state, ++seq, intent); return state; };
    const at = (anchor: number, focus = anchor) => ({ source: 'confirmed', selection: selection(story, anchor, focus) });
    await command({ kind: 'insertText', payload: { ...input, at: at(0), text: 'A' } });
    await command({ kind: 'formatText', payload: { ...input, at: at(0, 1), patch: { bold: true } } });
    await command({ kind: 'replaceText', payload: { ...input, at: at(1), text: 'paste' } });
    for (const [id, offset, text] of [['ime-1', 6, 'X'], ['ime-2', 7, 'Y']] as const) {
      await command({ kind: 'compositionStart', payload: { ...input, at: at(offset), compositionId: id } });
      await command({ kind: 'compositionCommit', payload: { ...input, at: at(offset), compositionId: id, text } });
      await command({ kind: 'compositionEnd', payload: { ...input, compositionId: id } });
    }
    for (const text of ['ApasteXalpha', 'Apastealpha', 'Aalpha', 'Aalpha', 'alpha']) {
      const ack = await command({ kind: 'undo', payload: input });
      expect(ack.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
      expect(session.handle().story(story.storyId).paragraphs[0]!.runs.map((run) => run.text).join('')).toBe(text);
      if (seq === 13) expect(session.handle().story(story.storyId).paragraphs[0]!.runs[0]!.style.bold).not.toBe(true);
    }
    expect(state.editState).toMatchObject({ canUndo: false, canRedo: true });
  } finally { await session.dispose(); }
});

test('history_refresh', async () => {
  const session = await open();
  try {
    const [first, second] = session.stories as [PptxStoryText, PptxStoryText];
    const edited = await session.command(session.state, 1, { kind: 'insertText', payload: {
      ...input, at: { source: 'confirmed', selection: selection(first) }, text: 'Z' } });
    const active = await session.command(edited, 2, { kind: 'setSelection', payload: {
      ...input, selection: selection(second, 5) } });
    const before = await session.thumbnail(first.slideId);
    const undone = await session.command(active, 3, { kind: 'undo', payload: input });
    expect(undone.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect(undone.dirtySlideIds).toContain(first.slideId);
    expect(undone.frame?.slideId).toBe(second.slideId);
    expect(undone.frame?.coversThroughSeq).toBe(3);
    expect(undone.frame!.slideRevision).toBe(active.frame!.slideRevision);
    expect(undone.selection).toBeNull();
    const thumbnail = await session.thumbnail(first.slideId, before.slideRevision);
    expect(thumbnail.status).toBe('frame');
    expect(thumbnail.slideRevision).toBeGreaterThan(before.slideRevision);
    const read = (await session.read()).result;
    expect(read.ok && read.stories.find((story) => story.storyId === first.storyId)?.text).toBe('alpha');
    const redone = await session.command(undone, 4, { kind: 'redo', payload: input });
    expect(redone.dirtySlideIds).toContain(first.slideId);
    expect(redone.editState).toMatchObject({ canUndo: true, canRedo: false });
    session.handle().addUndoBoundary();
    session.handle().deleteSlide(second.slideId);
    session.handle().addUndoBoundary();
    const restored = await session.command(redone, 5, { kind: 'undo', payload: input });
    const removed = await session.command(restored, 6, { kind: 'redo', payload: input });
    expect(removed.selection).toBeNull();
    expect(removed.activeSlideId).toBe(first.slideId);
    expect(removed.slides?.map((slide) => slide.id)).toEqual([first.slideId]);
    expect(removed.frame?.slideId).toBe(first.slideId);
    expect(removed.editState.toolbar.target).toBeNull();
    const refused = await session.command(removed, 7, { kind: 'insertText', payload: {
      ...input, at: { source: 'confirmed', selection: selection(second) }, text: 'invalid' } });
    expect(refused.outcomes[0]).toMatchObject({ status: 'refused', reason: 'bad-target' });
    const anchor = await session.command(refused, 8, { kind: 'insertText', payload: {
      ...input, at: { source: 'afterSeq', seq: 2 }, text: 'invalid' } }, 2);
    expect(anchor.outcomes[0]).toMatchObject({ status: 'refused', reason: 'stale-basis' });
    const stale = await session.command(edited, 9, { kind: 'insertText', payload: {
      ...input, at: { source: 'confirmed', selection: selection(first, 6) }, text: 'stale' } });
    expect(stale.outcomes[0]).toMatchObject({ status: 'refused', reason: 'stale-basis' });
  } finally { await session.dispose(); }
});

test('history_slide_order_preserves_unchanged_slide_cache', async () => {
  const session = await open();
  const read = spyOn(session.handle(), 'readCommandContent');
  const storyRead = spyOn(session.handle(), 'story');
  try {
    const [first, last] = session.stories as [PptxStoryText, PptxStoryText];
    const before = await session.thumbnail(first.slideId);
    session.handle().addUndoBoundary();
    session.handle().deleteSlide(last.slideId);
    session.handle().addUndoBoundary();
    read.mockClear(); storyRead.mockClear();
    const restored = await session.command(session.state, 1, { kind: 'undo', payload: input });
    expect(restored.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect(restored.dirtySlideIds).toEqual([last.slideId]);
    expect(read).toHaveBeenCalledWith({ slideIds: [last.slideId] });
    expect(storyRead.mock.calls.some(([id]) => id === first.storyId)).toBe(false);
    expect(await session.thumbnail(first.slideId, before.slideRevision)).toMatchObject({ status: 'unchanged' });
  } finally { read.mockRestore(); storyRead.mockRestore(); await session.dispose(); }
});

test('history_keeps_unaffected_thumbnail_and_story_cache', async () => {
  const session = await open();
  const read = spyOn(session.handle(), 'readCommandContent');
  const story = spyOn(session.handle(), 'story');
  const snapshot = spyOn(session.handle(), 'snapshot');
  try {
    const [first, second] = session.stories as [PptxStoryText, PptxStoryText];
    const edited = await session.command(session.state, 1, { kind: 'insertText', payload: {
      ...input, at: { source: 'confirmed', selection: selection(first) }, text: 'Z' } });
    const before = await session.thumbnail(second.slideId);
    read.mockClear(); story.mockClear(); snapshot.mockClear();
    const undone = await session.command(edited, 2, { kind: 'undo', payload: input });
    expect(undone.dirtySlideIds).toEqual([first.slideId]);
    expect(await session.thumbnail(second.slideId, before.slideRevision)).toMatchObject({
      status: 'unchanged', slideRevision: before.slideRevision });
    expect(read).toHaveBeenCalledWith({ slideIds: [first.slideId] });
    expect(story.mock.calls.every(([id]) => id !== second.storyId)).toBe(true);
    expect(snapshot).not.toHaveBeenCalled();
    const selected = await session.command(undone, 3, { kind: 'setSelection', payload: {
      ...input, selection: selection(second, 2) } });
    expect(selected.outcomes[0]).toMatchObject({ status: 'applied', committed: false });
    expect(selected.selection).toEqual(selection(second, 2));
    read.mockClear(); story.mockClear();
    const redone = await session.command(selected, 4, { kind: 'redo', payload: input });
    expect(redone.dirtySlideIds).toEqual([first.slideId]);
    expect(await session.thumbnail(second.slideId, before.slideRevision)).toMatchObject({
      status: 'unchanged', slideRevision: before.slideRevision });
    expect(read).toHaveBeenCalledWith({ slideIds: [first.slideId] });
    expect(story.mock.calls.every(([id]) => id !== second.storyId)).toBe(true);
    expect(snapshot).not.toHaveBeenCalled();
    const restored = await session.command(redone, 5, { kind: 'setSelection', payload: {
      ...input, selection: selection(second, 2) } });
    expect(restored.outcomes[0]).toMatchObject({ status: 'applied', committed: false });
    expect(restored.selection).toEqual(selection(second, 2));
  } finally { read.mockRestore(); story.mockRestore(); snapshot.mockRestore(); await session.dispose(); }
});

test.each([
  ['setShapeFill', (slide: string, text: string) => [slide, text, '#336699']],
  ['setShapeStroke', (slide: string, text: string) => [slide, text, { color: '#336699', widthPt: 2 }]],
  ['setShapeAdjust', (slide: string, _: string, box: string) => [slide, box, { adj: 30000 }]],
  ['removeShape', (slide: string, text: string) => [slide, text]],
  ['bringShapeToFront', (slide: string, text: string) => [slide, text]],
  ['sendShapeToBack', (slide: string, _: string, box: string) => [slide, box]],
  ['bringShapeForward', (slide: string, text: string) => [slide, text]],
  ['sendShapeBackward', (slide: string, _: string, box: string) => [slide, box]],
  ['setSlideNotes', (slide: string) => [slide, 'Speaker note']],
  ['addComment', (slide: string) => [slide, { author: 'Reviewer', text: 'Check this', created: '2026-01-01T00:00:00Z' }]],
])('invoking_%s_keeps_the_other_slides_thumbnail_and_story_cache', async (method, args) => {
  const session = await open();
  const read = spyOn(session.handle(), 'readCommandContent');
  try {
    const [first, second] = session.stories as [PptxStoryText, PptxStoryText];
    const opened = await session.command(session.state, 1, { kind: 'activateSlide', payload: { slideId: second.slideId } });
    const inserted = await session.command(opened, 2, { kind: 'insertShape', payload: {
      ...input, slideId: second.slideId, draft: { name: 'Box', geometry: 'roundRect', rect: { x: 200000, y: 1600000, width: 1500000, height: 900000 } } } });
    expect(inserted.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    const back = await session.command(inserted, 3, { kind: 'activateSlide', payload: { slideId: first.slideId } });
    const before = await session.thumbnail(first.slideId);
    read.mockClear();
    const edited = await session.command(back, 4, { kind: 'invoke', payload: {
      ...input, method, args: args(second.slideId, second.shapeId, inserted.creation!.shapeId) } });
    expect(edited.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect(edited.dirtySlideIds).toEqual([second.slideId]);
    expect(read).toHaveBeenCalledWith({ slideIds: [second.slideId] });
    expect(await session.thumbnail(first.slideId, before.slideRevision)).toMatchObject({
      status: 'unchanged', slideRevision: before.slideRevision });
  } finally { read.mockRestore(); await session.dispose(); }
});

test('accepting_a_proposal_redraws_the_slide_it_edits', async () => {
  const session = await open();
  try {
    const [, second] = session.stories as [PptxStoryText, PptxStoryText];
    const proposed = await session.command(session.state, 1, { kind: 'invoke', payload: {
      ...input, method: 'propose', args: ['Writer', null, [{ type: 'replaceText', storyId: second.storyId, start: 0, end: 5, text: 'omega' }]] } });
    expect(proposed.outcomes[0]).toMatchObject({ status: 'applied' });
    const before = await session.thumbnail(second.slideId);
    const accepted = await session.command(proposed, 2, { kind: 'invoke', payload: {
      ...input, method: 'acceptProposal', args: [(proposed.mutationResult as { id: string }).id] } });
    expect(accepted.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect(accepted.dirtySlideIds).toContain(second.slideId);
    expect(await session.thumbnail(second.slideId, before.slideRevision)).toMatchObject({ status: 'frame' });
  } finally { await session.dispose(); }
});

test.each([
  ['moveSlide', (_: string, second: string) => [second, 0], (first: string) => first],
  ['deleteSlide', (first: string) => [first], (_: string, second: string) => second],
  ['insertSlide', () => [0], (first: string) => first],
])('invoking_%s_redraws_the_slides_it_moves', async (method, args, moved) => {
  const session = await open();
  try {
    const [first, second] = session.stories as [PptxStoryText, PptxStoryText];
    const slide = moved(first.slideId, second.slideId);
    const before = await session.thumbnail(slide);
    const reordered = await session.command(session.state, 1, { kind: 'invoke', payload: {
      ...input, method, args: args(first.slideId, second.slideId) } });
    expect(reordered.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect(reordered.dirtySlideIds).toContain(slide);
    expect(await session.thumbnail(slide, before.slideRevision)).toMatchObject({ status: 'frame' });
  } finally { await session.dispose(); }
});

test('redoing_an_inserted_slide_reselects_it', async () => {
  const session = await open();
  try {
    const inserted = await session.command(session.state, 1, { kind: 'invoke', payload: {
      ...input, method: 'insertSlide', args: [1] } });
    expect(inserted.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    const insertedId = (inserted.mutationResult as { slideId: string }).slideId;
    expect(inserted.slides?.map((slide) => slide.id)).toHaveLength(3);
    expect(inserted.slides?.[1]?.id).toBe(insertedId);
    expect(inserted.activeSlideId).toBe(insertedId);
    const undone = await session.command(inserted, 2, { kind: 'undo', payload: input });
    expect(undone.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect(undone.slides?.map((slide) => slide.id)).toHaveLength(2);
    const redone = await session.command(undone, 3, { kind: 'redo', payload: input });
    expect(redone.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect(redone.slides?.map((slide) => slide.id)).toHaveLength(3);
    expect(redone.slides?.[1]?.id).toBe(insertedId);
    expect(redone.activeSlideId).toBe(insertedId);
    expect(redone.frame?.slideId).toBe(insertedId);
  } finally { await session.dispose(); }
});

test('undoing_a_deleted_slide_reselects_it', async () => {
  const session = await open();
  try {
    const [first, second] = session.stories as [PptxStoryText, PptxStoryText];
    const activated = await session.command(session.state, 1, { kind: 'activateSlide', payload: { slideId: second.slideId } });
    const deleted = await session.command(activated, 2, { kind: 'invoke', payload: {
      ...input, method: 'deleteSlide', args: [second.slideId] } });
    expect(deleted.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect(deleted.activeSlideId).toBe(first.slideId);
    const restored = await session.command(deleted, 3, { kind: 'undo', payload: input });
    expect(restored.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect(restored.slides?.map((slide) => slide.id)).toEqual([first.slideId, second.slideId]);
    expect(restored.activeSlideId).toBe(second.slideId);
    expect(restored.frame?.slideId).toBe(second.slideId);
    const redone = await session.command(restored, 4, { kind: 'redo', payload: input });
    expect(redone.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect(redone.activeSlideId).toBe(first.slideId);
  } finally { await session.dispose(); }
});

test('slide_changing_invocations_drop_the_previous_slides_selection', async () => {
  const session = await open();
  try {
    const [first, second] = session.stories as [PptxStoryText, PptxStoryText];
    const object = await session.command(session.state, 1, { kind: 'selectObject', payload: {
      ...input, frame: session.state.frame, shapeId: first.shapeId } });
    expect(object.selection).toMatchObject({ kind: 'object', slideId: first.slideId, shapeId: first.shapeId });
    const moved = await session.command(object, 2, { kind: 'invoke', payload: {
      ...input, method: 'moveSlide', args: [second.slideId, 0] } });
    expect(moved.activeSlideId).toBe(first.slideId);
    expect(moved.selection).toMatchObject({ kind: 'object', slideId: first.slideId, shapeId: first.shapeId });
    const inserted = await session.command(moved, 3, { kind: 'invoke', payload: {
      ...input, method: 'insertSlide', args: [2] } });
    expect(inserted.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect(inserted.activeSlideId).toBe((inserted.mutationResult as { slideId: string }).slideId);
    expect(inserted.selection).toBeNull();
    expect(inserted.editState.object ?? null).toBeNull();

    const caret = await session.command(inserted, 4, { kind: 'setSelection', payload: {
      ...input, selection: selection(first, 1, 3) } });
    expect(caret.activeSlideId).toBe(first.slideId);
    const again = await session.command(caret, 5, { kind: 'invoke', payload: {
      ...input, method: 'insertSlide', args: [2] } });
    expect(again.activeSlideId).toBe((again.mutationResult as { slideId: string }).slideId);
    expect(again.selection).toBeNull();

    const back = await session.command(again, 6, { kind: 'setSelection', payload: {
      ...input, selection: selection(first, 2) } });
    expect(back.activeSlideId).toBe(first.slideId);
    const deleted = await session.command(back, 7, { kind: 'invoke', payload: {
      ...input, method: 'deleteSlide', args: [first.slideId] } });
    expect(deleted.outcomes[0]).toMatchObject({ status: 'applied', committed: true });
    expect(deleted.activeSlideId).not.toBe(first.slideId);
    expect(deleted.selection).toBeNull();
  } finally { await session.dispose(); }
});
