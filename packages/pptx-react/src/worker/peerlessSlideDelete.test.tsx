import { expect, test } from 'bun:test';
import { open, replyControl } from './peerlessTestHarness';

const { act, fireEvent, waitFor } = await import('@testing-library/react');

function thumbnails(container: HTMLElement): HTMLButtonElement[] {
  return Array.from(container.querySelectorAll('aside button'));
}

test('a clicked thumbnail keeps focus and Delete removes that slide', async () => {
  const value = await open();
  const first = value.client.state.slides[0]!.id;
  const items = thumbnails(value.view.container);
  act(() => { fireEvent.click(items[1]!, { detail: 1 }); });
  expect(document.activeElement).toBe(items[1]!);
  act(() => { fireEvent.keyDown(items[1]!, { key: 'Delete' }); });
  await waitFor(() => expect(value.client.state.slides.map((slide) => slide.id)).toEqual([first]));
  await waitFor(() => expect(document.activeElement).toBe(thumbnails(value.view.container)[0]!));
});

test('Backspace on a focused thumbnail removes it and focus moves to the one in its place', async () => {
  const value = await open();
  const second = value.client.state.slides[1]!.id;
  const items = thumbnails(value.view.container);
  act(() => { items[0]!.focus(); fireEvent.keyDown(items[0]!, { key: 'Backspace' }); });
  await waitFor(() => expect(value.client.state.slides.map((slide) => slide.id)).toEqual([second]));
  await waitFor(() => expect(document.activeElement).toBe(thumbnails(value.view.container)[0]!));
});

test('Delete with a modifier leaves the slides alone and keyboard activation moves focus to the stage', async () => {
  const value = await open();
  const items = thumbnails(value.view.container);
  act(() => { fireEvent.keyDown(items[1]!, { key: 'Delete', metaKey: true }); });
  act(() => { fireEvent.click(items[1]!, { detail: 0 }); });
  expect(document.activeElement).toBe(value.stage);
  expect(value.client.state.slides).toHaveLength(2);
});

test('the context menu deletes the thumbnail it opened on', async () => {
  const value = await open();
  const first = value.client.state.slides[0]!.id;
  act(() => { fireEvent.contextMenu(thumbnails(value.view.container)[1]!, { clientX: 40, clientY: 60 }); });
  const item = value.view.getByRole('menuitem', { name: 'Delete slide' });
  expect(document.activeElement).toBe(item);
  act(() => { fireEvent.click(item); });
  await waitFor(() => expect(value.client.state.slides.map((slide) => slide.id)).toEqual([first]));
  expect(value.view.queryByRole('menu')).toBeNull();
});

test('Escape closes the context menu, keeps the slides and returns focus to the thumbnail', async () => {
  const value = await open();
  const items = thumbnails(value.view.container);
  act(() => { fireEvent.contextMenu(items[0]!); });
  act(() => { fireEvent.keyDown(value.view.getByRole('menuitem'), { key: 'Escape' }); });
  expect(value.view.queryByRole('menu')).toBeNull();
  expect(document.activeElement).toBe(items[0]!);
  expect(value.client.state.slides).toHaveLength(2);
});

test('right-clicking another thumbnail retargets the open menu', async () => {
  const value = await open();
  const second = value.client.state.slides[1]!.id;
  const items = thumbnails(value.view.container);
  act(() => { fireEvent.contextMenu(items[1]!); });
  act(() => { fireEvent.contextMenu(items[0]!); });
  act(() => { fireEvent.click(value.view.getByRole('menuitem', { name: 'Delete slide' })); });
  await waitFor(() => expect(value.client.state.slides.map((slide) => slide.id)).toEqual([second]));
});

test('the strip keeps the last slide and ignores a held Delete', async () => {
  const value = await open();
  const items = thumbnails(value.view.container);
  act(() => { fireEvent.keyDown(items[0]!, { key: 'Delete', repeat: true }); });
  expect(value.client.state.slides).toHaveLength(2);
  act(() => { fireEvent.keyDown(items[0]!, { key: 'Delete' }); });
  await waitFor(() => expect(value.client.state.slides).toHaveLength(1));
  const [only] = thumbnails(value.view.container);
  act(() => { fireEvent.keyDown(only!, { key: 'Delete' }); });
  expect(fireEvent.contextMenu(only!)).toBe(true);
  expect(value.view.queryByRole('menu')).toBeNull();
  await act(async () => { await value.api.flushPendingInput(); });
  expect(value.client.state.slides).toHaveLength(1);
});

test('a second Delete sent before the first is acknowledged still keeps the last slide', async () => {
  const value = await open();
  const items = thumbnails(value.view.container);
  replyControl.hold();
  act(() => { fireEvent.keyDown(items[0]!, { key: 'Delete' }); fireEvent.keyDown(items[1]!, { key: 'Delete' }); });
  await act(async () => { replyControl.release(); await value.api.flushPendingInput(); });
  await waitFor(() => expect(value.client.state.slides).toHaveLength(1));
});

test('after a pointer click, PageDown, Home and End move between slides from the thumbnail', async () => {
  const value = await open();
  const [first, second] = value.client.state.slides.map((slide) => slide.id);
  const items = thumbnails(value.view.container);
  act(() => { fireEvent.click(items[0]!, { detail: 1 }); });
  act(() => { fireEvent.keyDown(items[0]!, { key: 'PageDown' }); });
  await waitFor(() => expect(value.client.state.activeSlideId).toBe(second!));
  expect(document.activeElement).toBe(items[1]!);
  act(() => { fireEvent.keyDown(items[1]!, { key: 'Home' }); });
  await waitFor(() => expect(value.client.state.activeSlideId).toBe(first!));
  expect(document.activeElement).toBe(items[0]!);
  act(() => { fireEvent.keyDown(items[0]!, { key: 'End' }); });
  await waitFor(() => expect(value.client.state.activeSlideId).toBe(second!));
});

test('a read-only editor neither deletes on Delete nor opens the menu', async () => {
  const value = await open({ readOnly: true });
  const items = thumbnails(value.view.container);
  act(() => { fireEvent.keyDown(items[0]!, { key: 'Delete' }); });
  expect(fireEvent.contextMenu(items[0]!)).toBe(true);
  expect(value.view.queryByRole('menu')).toBeNull();
  expect(value.client.state.slides).toHaveLength(2);
});
