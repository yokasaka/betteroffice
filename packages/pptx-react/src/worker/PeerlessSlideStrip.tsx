import { memo, useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';
import type { CSSProperties } from 'react';
import type { SlideFit } from '../fitSlideScale';
import { useTranslation } from '../i18n';
import type { PeerlessConfirmedFrame, PeerlessViewSnapshot } from './peerlessViewStore';
import { peerlessDeleteSlide, peerlessInsertSlide, peerlessMoveSlide } from './peerlessSlideCommands';
import { activatePeerlessSlide } from './peerlessNavigation';
import { peerlessPainterFactories, type PeerlessPresentation } from './usePeerlessPresentation';

const Preview = memo(function Preview({ frame, run, fit, previewFit, current }: {
  frame: PeerlessConfirmedFrame; run: PeerlessPresentation; fit: SlideFit;
  previewFit(frame: PeerlessConfirmedFrame, run: PeerlessPresentation): SlideFit | null;
  current(token: SlideFit, run: PeerlessPresentation): boolean;
}) {
  const canvas = useRef<HTMLCanvasElement>(null);
  const latest = useRef({ previewFit, current });
  latest.current = { previewFit, current };
  const resume = useRef<((current: () => boolean) => Promise<void>) | null>(null);
  useEffect(() => {
    const valid = () => {
      const token = latest.current.previewFit(frame, run);
      return Boolean(token && token.scale === fit.scale && token.dpr === fit.dpr && latest.current.current(token, run));
    };
    if (!canvas.current) return;
    const painter = peerlessPainterFactories.frame(canvas.current, { ...fit, images: run.images, current: valid });
    const view: PeerlessViewSnapshot = { confirmed: frame, selection: null, focused: false, caret: null,
      selectionRects: [], composition: null, echo: null, activeLine: null, activeContextBytes: 0,
      pending: [], waitingInputCount: 0, appliedThroughSeq: frame.frame.coversThroughSeq,
      presentedThroughSeq: frame.frame.coversThroughSeq,
      editState: { readOnly: true, canUndo: false, canRedo: false,
        toolbar: { target: null, storedMarks: {}, formatting: null, alignment: null,
          hyperlink: null, table: null, canEditText: false } },
      thumbnails: new Map(), dirtyThumbnailIds: new Set<string>() };
    painter.paintFrame(frame, view);
    void painter.whenPainted().catch(() => {});
    resume.current = async (current) => {
      if (!await painter.whenPainted() && current() && valid()) {
        painter.paintFrame(frame, view);
        await painter.whenPainted();
      }
    };
    return () => { resume.current = null; painter.dispose(); };
  }, [frame, run, fit.scale, fit.dpr]);
  useEffect(() => {
    let alive = true;
    void resume.current?.(() => alive).catch(() => {});
    return () => { alive = false; };
  }, [frame, run, fit.scale, fit.dpr, fit.generation]);
  return <canvas ref={canvas} style={{ width: 128, display: 'block' }} aria-hidden="true" />;
});

const Slide = memo(function Slide({ id, name, index, active, disabled, frame, run, previewFit, current, onNavigate, onDelete, onStep, onMenu, row }: {
  id: string; name: string | null; index: number; active: boolean; disabled: boolean; row: boolean;
  frame: PeerlessConfirmedFrame | null | undefined; run: PeerlessPresentation;
  previewFit(frame: PeerlessConfirmedFrame, run: PeerlessPresentation): SlideFit | null;
  current(token: SlideFit, run: PeerlessPresentation): boolean;
  onNavigate(): void;
  onDelete(id: string, index: number): void;
  onStep(index: number): void;
  onMenu(menu: SlideMenu): void;
}) {
  const { t } = useTranslation();
  const fit = useMemo(() => frame ? previewFit(frame, run) : null, [frame, run, previewFit]);
  return <button type="button" disabled={disabled}
    aria-current={active ? 'page' : undefined} style={row ? styles.rowSlide : styles.slide}
    draggable={!run.store.getSnapshot().editState.readOnly}
    onDragStart={(event) => event.dataTransfer.setData('text/plain', id)}
    onDragOver={(event) => event.preventDefault()}
    onDrop={(event) => {
      event.preventDefault();
      const source = event.dataTransfer.getData('text/plain');
      if (run.client.state.slides.some((slide) => slide.id === source)) void peerlessMoveSlide(run, source, index).catch(() => {});
    }}
    onKeyDown={(event) => {
      if (!event.altKey && !event.ctrlKey && !event.metaKey && !event.shiftKey) {
        if (event.key === 'Delete' || event.key === 'Backspace') {
          event.preventDefault();
          if (!event.repeat) onDelete(id, index);
          return;
        }
        const last = run.client.state.slides.length - 1;
        const next = ['PageUp', 'ArrowUp', 'ArrowLeft'].includes(event.key) ? Math.max(0, index - 1)
          : ['PageDown', 'ArrowDown', 'ArrowRight'].includes(event.key) ? Math.min(last, index + 1)
          : event.key === 'Home' ? 0 : event.key === 'End' ? last : null;
        if (next === null) return;
        event.preventDefault();
        onStep(next);
        return;
      }
      if (!event.altKey || !['ArrowUp', 'ArrowDown'].includes(event.key)) return;
      event.preventDefault();
      const next = index + (event.key === 'ArrowUp' ? -1 : 1);
      if (next >= 0 && next < run.client.state.slides.length) void peerlessMoveSlide(run, id, next).catch(() => {});
    }}
    onContextMenu={(event) => {
      if (run.store.getSnapshot().editState.readOnly || run.client.state.slides.length < 2) return;
      event.preventDefault();
      onMenu({ id, index, x: event.clientX, y: event.clientY, trigger: event.currentTarget });
    }}
    onClick={(event) => {
      void activatePeerlessSlide(run, id).catch(() => {});
      // A pointer click keeps focus on the thumbnail so Delete reaches the strip; Enter and Space move to the slide.
      if (event.detail === 0) onNavigate();
      else event.currentTarget.focus();
    }}>
    <span>{index + 1}</span>
    {frame && fit ? <Preview frame={frame} run={run} fit={fit} previewFit={previewFit} current={current} /> : <span>{name ?? t('slides.fallbackTitle', { number: index + 1 })}</span>}
  </button>;
});

type SlideMenu = { id: string; index: number; x: number; y: number; trigger: HTMLElement };

function SlideContextMenu({ menu, onDelete, onClose }: {
  menu: SlideMenu; onDelete(id: string, index: number): void; onClose(): void;
}) {
  const { t } = useTranslation();
  const root = useRef<HTMLDivElement>(null);
  const item = useRef<HTMLButtonElement>(null);
  const [position, setPosition] = useState({ left: menu.x, top: menu.y });
  // Open upward or leftward when the pointer is near the bottom or right edge of the window.
  useLayoutEffect(() => {
    const rect = root.current!.getBoundingClientRect();
    setPosition({ left: Math.max(0, Math.min(menu.x, window.innerWidth - rect.width)),
      top: menu.y + rect.height > window.innerHeight ? Math.max(0, menu.y - rect.height) : menu.y });
  }, [menu.x, menu.y]);
  useEffect(() => { item.current?.focus(); }, []);
  return <div ref={root} role="menu" aria-label={t('slides.panelLabel')} style={{ ...styles.menu, ...position }}
    onMouseDown={(event) => event.preventDefault()}
    onKeyDown={(event) => { if (event.key === 'Escape') { event.preventDefault(); menu.trigger.focus(); onClose(); } }}>
    <button ref={item} type="button" role="menuitem" style={styles.menuItem} onBlur={onClose}
      onClick={() => { onClose(); onDelete(menu.id, menu.index); }}>{t('toolbar.deleteSlide')}</button>
  </div>;
}

function SlideActions({ run, active, disabled }: { run: PeerlessPresentation; active?: string; disabled: boolean }) {
  const { t } = useTranslation();
  const subscribe = useCallback((listener: () => void) => run.store.subscribeEditState(listener), [run]);
  const snapshot = useCallback(() => run.store.getSnapshot().editState.readOnly, [run]);
  const readOnly = useSyncExternalStore(subscribe, snapshot, snapshot);
  return readOnly ? null : <div>
    <button type="button" disabled={disabled} onClick={() => { void peerlessInsertSlide(run).catch(() => {}); }}>{t('toolbar.addSlide')}</button>
    <button type="button" disabled={disabled || !active} onClick={() => { if (active) void peerlessDeleteSlide(run, active).catch(() => {}); }}>{t('toolbar.deleteSlide')}</button>
  </div>;
}

export const PeerlessSlideStrip = memo(function PeerlessSlideStrip({ run, previewFit, current, onNavigate, placement = 'left' }: {
  run: PeerlessPresentation; previewFit(frame: PeerlessConfirmedFrame, run: PeerlessPresentation): SlideFit | null;
  current(token: SlideFit, run: PeerlessPresentation): boolean;
  onNavigate(): void;
  placement?: 'left' | 'bottom';
}) {
  const { store, client } = run;
  const { t } = useTranslation();
  const subscribe = useCallback((listener: () => void) => {
    const offStructure = store.subscribeStructure(listener);
    const offPresentation = store.subscribePresentation(listener);
    return () => { offStructure(); offPresentation(); };
  }, [store]);
  const snapshot = useMemo(() => {
    let previous: { active: string | undefined; disabled: boolean; slides: { id: string; name: string | null;
      frame: PeerlessConfirmedFrame | null | undefined }[] } | undefined;
    return () => {
      const view = store.getSnapshot();
      const active = view.confirmed?.frame.slideId;
      const disabled = !run.current || Boolean(run.failure);
      const slides = client.state.slides.map((slide) => ({ id: slide.id, name: slide.name,
        frame: slide.id === active ? view.confirmed : view.thumbnails.get(slide.id) }));
      if (previous && previous.active === active && previous.disabled === disabled && previous.slides.length === slides.length &&
        slides.every((slide, index) => slide.id === previous!.slides[index]!.id &&
          slide.name === previous!.slides[index]!.name && slide.frame === previous!.slides[index]!.frame)) return previous;
      previous = { active, disabled, slides };
      return previous;
    };
  }, [store, client, run]);
  const view = useSyncExternalStore(subscribe, snapshot, snapshot);
  const row = placement === 'bottom';
  const list = useRef<HTMLElement>(null);
  // After a delete, focus the thumbnail that took the deleted one's place so Delete can repeat.
  const refocus = useRef<{ index: number; count: number } | null>(null);
  const [menu, setMenu] = useState<SlideMenu | null>(null);
  // The strip never deletes the last slide: with none left it unmounts and focus is lost.
  const deleting = useRef(0);
  const remove = useCallback((id: string, index: number) => {
    const count = run.client.state.slides.length;
    if (count - deleting.current < 2 || run.store.getSnapshot().editState.readOnly) return;
    const target = { index, count };
    const clear = () => { if (refocus.current === target) refocus.current = null; };
    refocus.current = target;
    deleting.current += 1;
    try {
      void peerlessDeleteSlide(run, id, { reportRefusal: false })
        .then((ack) => { if (!ack.outcomes.some((outcome) => outcome.committed)) clear(); }, clear)
        .finally(() => { deleting.current -= 1; });
    } catch { deleting.current -= 1; clear(); }
  }, [run]);
  const step = useCallback((index: number) => {
    const slide = run.client.state.slides[index];
    if (!slide) return;
    void activatePeerlessSlide(run, slide.id).catch(() => {});
    list.current?.querySelectorAll('button')[index]?.focus();
  }, [run]);
  const closeMenu = useCallback(() => setMenu(null), []);
  useLayoutEffect(() => {
    const target = refocus.current;
    if (!target || view.slides.length >= target.count) return;
    refocus.current = null;
    const buttons = list.current?.querySelectorAll('button');
    if (buttons?.length) buttons[Math.min(target.index, buttons.length - 1)]!.focus();
  }, [view.slides]);
  return <div style={row ? styles.row : styles.strip}><aside ref={list} style={row ? styles.rowList : styles.list} aria-label={t('slides.panelLabel')}>
    {view.slides.map((slide, index) => <Slide key={slide.id} {...slide} index={index} active={slide.id === view.active}
      disabled={view.disabled} run={run} previewFit={previewFit} current={current} onNavigate={onNavigate}
      onDelete={remove} onStep={step} onMenu={setMenu} row={row} />)}
  </aside><SlideActions run={run} active={view.active} disabled={view.disabled} />
  {menu && !view.disabled && <SlideContextMenu menu={menu} onDelete={remove} onClose={closeMenu} />}</div>;
});

const styles: Record<string, CSSProperties> = {
  strip: { width: 160, display: 'flex', flexDirection: 'column', padding: 8, background: '#f1f5f9' },
  list: { overflowY: 'auto', flex: 1 },
  slide: { display: 'flex', flexDirection: 'column', gap: 4, width: '100%', marginBottom: 8 },
  row: { display: 'flex', flexDirection: 'row', alignItems: 'center', gap: 8, padding: 8, background: '#f1f5f9' },
  rowList: { display: 'flex', gap: 8, overflowX: 'auto', overflowY: 'hidden', flex: 1, minWidth: 0 },
  rowSlide: { display: 'flex', flexDirection: 'row', alignItems: 'flex-start', gap: 4, flex: 'none' },
  menu: { position: 'fixed', zIndex: 10, padding: 4, background: '#fff', border: '1px solid #cbd5e1', borderRadius: 4,
    boxShadow: '0 4px 12px rgba(15, 23, 42, 0.15)' },
  menuItem: { display: 'block', width: '100%', padding: '4px 12px', textAlign: 'start', background: 'none', border: 0 },
};
