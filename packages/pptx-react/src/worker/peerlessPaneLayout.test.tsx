import { expect, test } from 'bun:test';
import { open } from './peerlessTestHarness';

const { waitFor } = await import('@testing-library/react');

function strip(container: HTMLElement): HTMLElement {
  return container.querySelector('aside')!.parentElement!;
}

test('the slide strip, notes and stage keep the default layout without pane props', async () => {
  const value = await open();
  await waitFor(() => expect(value.view.container.querySelector('aside')).not.toBeNull());
  const rail = strip(value.view.container);
  expect(rail.nextElementSibling).toBe(value.stage);
  expect(rail.style.flexDirection).toBe('column');
  expect(value.view.getByTestId('pptx-notes-textarea')).toBeDefined();
});

test('slideStrip bottom lays the thumbnails out in a row under the stage', async () => {
  const value = await open({ slideStrip: 'bottom' });
  await waitFor(() => expect(value.view.container.querySelector('aside')).not.toBeNull());
  const rail = strip(value.view.container);
  expect(value.stage.compareDocumentPosition(rail) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
  expect(value.stage.contains(rail)).toBe(false);
  expect(rail.style.flexDirection).toBe('row');
  const aside = rail.querySelector('aside')!;
  expect(aside.style.overflowX).toBe('auto');
  expect(aside.querySelector('button[aria-current="page"]')).not.toBeNull();
});

test('slideStrip hidden renders no slide strip', async () => {
  const value = await open({ slideStrip: 'hidden' });
  expect(value.view.container.querySelector('aside')).toBeNull();
  expect(value.view.queryByText('Add slide')).toBeNull();
});

test('showNotes false renders no notes field', async () => {
  const value = await open({ showNotes: false });
  expect(value.view.queryByTestId('pptx-notes-textarea')).toBeNull();
});
