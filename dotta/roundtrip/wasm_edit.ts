import { readFile, writeFile } from 'node:fs/promises';
import { initWasm, openPresentation } from '../../packages/pptx/src/wasm/loader';

const [input, output, sizesJson, layoutSlide] = process.argv.slice(2);
if (!input || !output || !sizesJson || !layoutSlide) {
  throw new Error('usage: bun wasm_edit.ts <in.pptx> <out.pptx> <{"slide":n,"widths":[],"heights":[]}> <layout slide index>');
}
const sizes = JSON.parse(sizesJson) as { slide: number; widths: number[]; heights: number[] };

await initWasm(await readFile(new URL('../../packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm', import.meta.url)));
const font = await readFile(new URL('../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url));
const handle = openPresentation(new Uint8Array(await readFile(input)),
  { fonts: [{ family: 'Liberation Sans', bytes: new Uint8Array(font) }] });
try {
  const slide = handle.snapshot().slides[sizes.slide]!;
  const table = slide.shapes.find((shape) => shape.kind === 'graphicFrame')!;
  handle.setTableSizes(slide.id, table.id, sizes.widths, sizes.heights);
  const layout = handle.layoutSlide(Number(layoutSlide));
  await writeFile(output, handle.save());
  const geometries = layout.primitives.flatMap((primitive) => primitive.kind === 'shape' ? [primitive.geometry] : []);
  process.stdout.write(JSON.stringify({ geometries }));
} finally {
  handle.dispose();
}
