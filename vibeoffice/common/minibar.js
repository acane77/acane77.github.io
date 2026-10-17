/* The formatting bar that appears after a text selection (Office's mini toolbar).
   Quire and Lectern call mount() once with the buttons, then show() when a range is made.
   It sits above and to the right of the pointer, inside the window. Opacity starts at 0.7.
   On the bar, or on a menu it opened, it is fully opaque. Elsewhere it follows the distance
   from the point where it opened: 0.7 at that point, 0 at 50px, and once it reaches 0 it
   stays gone until the next show(). fade() and place() are pure (tests/minibar.test.mjs). */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const FADE = 50, REST = 0.7, GAP = 4, EDGE = 2;

  /** opacity at distance d from the open point; over means the pointer is on the bar or its menu.
      latched: it already faded out, so the pointer cannot bring it back. */
  function fade(d, over, latched) {
    if (latched) return { opacity: 0, latched: true };
    const dist = typeof d === 'number' && d > 0 ? d : 0;
    const opacity = over ? 1 : (dist >= FADE ? 0 : REST * (1 - dist / FADE));
    return { opacity, latched: !over && opacity <= 0 };
  }
  /** top-left of a bar of size {w,h} placed up-right of anchor, gap px clear of it, inside view {w,h} */
  function place(anchor, size, view, gap) {
    const g = gap == null ? GAP : gap;
    let x = anchor.x + g;
    let y = anchor.y - size.h - g;
    const maxX = Math.max(EDGE, view.w - size.w - EDGE);
    const maxY = Math.max(EDGE, view.h - size.h - EDGE);
    if (x < EDGE) x = EDGE;
    if (y < EDGE) y = EDGE;
    if (x > maxX) x = maxX;
    if (y > maxY) y = maxY;
    return { x, y };
  }

  L.minibar = { FADE, REST, GAP, fade, place };
  if (!root.document) return;

  let build = null, bar = null, anchor = null;
  let latched = false, shown = false, over = false;
  let pointer = { x: 0, y: 0 };
  let quietUntil = 0;   /* ignore the scroll that follows a selection into view */

  const ui = () => L.ui;
  const onBar = (node) => !!(node && node.closest && (node.closest('.minibar') || (ui() && ui().menuOpen() && node.closest('.menu'))));

  L.minibar.mount = (fn) => { build = fn; };
  L.minibar.pointer = () => ({ x: pointer.x, y: pointer.y });
  L.minibar.visible = () => shown;
  L.minibar.latched = () => latched;

  function remember(e) { pointer = { x: e.clientX, y: e.clientY }; }

  function covers(el, x, y) {
    const r = el.getBoundingClientRect();
    return x >= r.left && x < r.right && y >= r.top && y < r.bottom;
  }

  L.minibar.show = function (x, y) {
    if (!build || (ui() && ui().dialogOpen())) return;
    if (x == null) { x = pointer.x; y = pointer.y; }
    if (!bar) {
      bar = L.h('div', { class: 'minibar', role: 'toolbar', 'aria-label': 'Formatting' });
      for (const row of build()) bar.appendChild(row);
      root.document.body.appendChild(bar);
    }
    anchor = { x, y };
    latched = false;
    shown = true;
    over = false;
    quietUntil = Date.now() + 120;
    bar.hidden = false;
    const rect = bar.getBoundingClientRect();
    const p = place(anchor, { w: rect.width, h: rect.height }, { w: root.innerWidth, h: root.innerHeight });
    bar.style.left = p.x + 'px';
    bar.style.top = p.y + 'px';
    over = covers(bar, x, y);
    const state = fade(0, over, false);
    bar.style.opacity = String(state.opacity);
    if (ui()) ui().refresh();
  };

  function conceal() {
    shown = false;
    over = false;
    if (bar) { bar.hidden = true; bar.style.opacity = ''; }
  }
  L.minibar.hide = function () { latched = false; conceal(); };
  function latch() {
    latched = true;
    conceal();
    if (ui() && ui().menuOpen()) ui().closeMenus();
  }

  function paint(node) {
    if (!shown || !anchor) return;
    if (ui() && ui().dialogOpen()) { L.minibar.hide(); return; }
    over = onBar(node) || !!(bar && bar.matches(':hover'));
    const d = Math.hypot(pointer.x - anchor.x, pointer.y - anchor.y);
    const state = fade(d, over, false);
    if (state.latched) latch();
    else if (bar) bar.style.opacity = String(state.opacity);
  }

  root.addEventListener('pointermove', (e) => { remember(e); if (shown) paint(e.target); }, true);
  root.addEventListener('pointerdown', remember, true);
  root.addEventListener('pointerup', remember, true);

  function typing(e) {
    if (e.ctrlKey || e.metaKey || e.altKey || e.key === 'Escape') return false;
    return e.key === 'Backspace' || e.key === 'Delete' || e.key === 'Enter' || e.key === 'Tab' || e.key.length === 1;
  }
  root.addEventListener('keydown', (e) => {
    if (!shown) return;
    if (onBar(e.target)) {
      if (e.key === 'Escape' && !(ui() && ui().menuOpen())) L.minibar.hide();
      return;
    }
    if (e.key === 'Escape' && !(ui() && ui().menuOpen())) L.minibar.hide();
    else if (typing(e)) L.minibar.hide();
  }, true);
  root.addEventListener('compositionstart', (e) => { if (shown && !onBar(e.target)) L.minibar.hide(); }, true);

  function scrollingBox(t) {
    return !!(t && t.closest && (t.closest('.menu') || t.closest('.minibar') || t.classList && (t.classList.contains('menu') || t.classList.contains('minibar'))));
  }
  root.addEventListener('scroll', (e) => {
    if (!shown || Date.now() < quietUntil || scrollingBox(e.target)) return;
    L.minibar.hide();
  }, true);
  root.addEventListener('wheel', (e) => { if (shown && !onBar(e.target)) L.minibar.hide(); }, { capture: true, passive: true });
  root.addEventListener('resize', () => { if (shown) L.minibar.hide(); });
  root.addEventListener('click', () => {
    if (!shown) return;
    root.requestAnimationFrame(() => { if (shown) paint(root.document.elementFromPoint(pointer.x, pointer.y)); });
  }, true);
})(typeof window !== 'undefined' ? window : globalThis);
