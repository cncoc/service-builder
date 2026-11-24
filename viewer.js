/* Hymn and chorus viewer, and the full-screen presenter, for the Service Builder.
 *
 * Viewer (a dialog): Slides · Page · Words · Play for a hymn; Slides · Words for a chorus.
 * Presenter: shows a hymn's slides (or a chorus's) full screen straight away, for songs sung on the spot.
 *   Next: → ↓ Space Enter PageDown, a click/tap on the right, or a swipe left (clickers send PageDown).
 *   Back: ← ↑ PageUp Backspace, a tap on the left. 1–9 jump to that stanza, B or . blacks the screen,
 *   F toggles full screen, Esc stops. The bar's stanza buttons turn stanzas on and off while presenting.
 *   On a laptop with an extended projector screen it can open on the projector.
 *
 * Created by index.html with what it needs from the builder:
 *   HymnViewer({ api, hymn(n), chorus(n), hymnSlides(item), picCache, ccli, add(kind, n, opts), locked() })
 */
(function () {
  "use strict";

  window.HymnViewer = function create(deps) {
    const { api, hymn, chorus, hymnSlides, picCache, ccli, add, locked } = deps;
    const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
    const b64ToBytes = (s) => { const bin = atob(s); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; };
    const pad = (n, w = 2) => String(n).padStart(w, "0");
    const M = window.HymnMidi;

    /* ---------- Pictures from the church Drive ---------- */
    const urls = new Map(), inflight = new Map();
    function picUrl(name) {
      if (urls.has(name)) return urls.get(name);
      const b = picCache.get(name); if (!b) return null;
      const u = URL.createObjectURL(new Blob([b], { type: "image/gif" })); urls.set(name, u); return u;
    }
    function loadPics(n, names) {
      const need = names.filter((nm) => !picCache.has(nm) && !inflight.has(nm));
      if (need.length) {
        const pr = api({ action: "pics", n, names: need.join(",") }).then((r) => {
          for (const [k, v] of Object.entries(r.pics)) picCache.set(k, b64ToBytes(v));
        });
        const done = () => need.forEach((nm) => inflight.delete(nm));
        pr.then(done, done);
        need.forEach((nm) => inflight.set(nm, pr));
      }
      return Promise.all(names.map((nm) => inflight.get(nm)).filter(Boolean));
    }

    /* ---------- Words, music and pages (need the Media.gs part of the church script) ---------- */
    const mediaCache = new Map(), pageCache = new Map();
    function media(n) {
      if (!mediaCache.has(n)) {
        const pr = api({ action: "media", what: "words", n }).then((r) => {
          const w = r.words; w.song = M.parse(b64ToBytes(w.midi)); delete w.midi; return w;
        });
        pr.catch(() => mediaCache.delete(n));
        mediaCache.set(n, pr);
      }
      return mediaCache.get(n);
    }
    function page(n) {
      if (!pageCache.has(n)) {
        const pr = api({ action: "media", what: "page", n }).then((r) => URL.createObjectURL(new Blob([b64ToBytes(r.page)], { type: "image/gif" })));
        pr.catch(() => pageCache.delete(n));
        pageCache.set(n, pr);
      }
      return pageCache.get(n);
    }
    function why(e) {
      if (e && e.locked) { locked(); return ""; }
      const m = (e && e.message) || String(e);
      if (/unknown action/i.test(m)) return "This part needs the latest update to the church's Apps Script (the Media.gs file). The slides and presenting work without it.";
      if (/not on the church drive/i.test(m)) return esc(m) + ".";
      if (/load failed|failed to fetch|networkerror/i.test(m)) return "The connection to the church Drive dropped. Check the internet connection and try again.";
      return "Couldn't load it: " + esc(m);
    }

    /* ---------- Shared pieces ---------- */
    const stLabel = (code) => (code === "OE" ? "Ending" : code === "OI" ? "Intro" : "Stanza " + Number(code));
    function hymnSeq(h, sel) {
      return hymnSlides({ n: h.n, st: sel.st, oe: sel.oe, oi: sel.oi }).map((s) => ({ type: "pic", name: s.name, code: s.name.split("-")[1] }));
    }
    function chorusSeq(c) { return c.slides.map((blocks, i) => ({ type: "chorus", blocks, i, total: c.slides.length })); }
    // A chorus slide drawn like the PowerPoint one (gen.js): sizes are fractions of a 16:9 slide
    function chorusSlideHtml(c, it, showCcli) {
      const sz = window.ServiceDeck.chorusFontSize(it.blocks), f = (pt) => `${(pt * 100 / 960).toFixed(3)}cqw`;
      const words = it.blocks.map((b, k) =>
        `<div class="cs-block" style="margin-top:${k ? f(0.55 * sz) : 0}">${b.label ? `<div class="cs-label" style="font-size:${f(0.62 * sz)}">${esc(b.label)}</div>` : ""}` +
        b.lines.map((l) => `<div class="cs-line" style="font-size:${f(sz)}">${esc(l)}</div>`).join("") + "</div>").join("");
      return `<div class="cs"><span class="cs-num">${c.n}</span><span class="cs-title">${esc(c.t)}${c.ref ? `<small>${esc(c.ref)}</small>` : ""}</span>
        <span class="cs-cue">${it.i + 1} of ${it.total}</span><i class="cs-rule"></i><div class="cs-words">${words}</div>
        <span class="cs-foot">CNCOC Songbook ${c.n}${showCcli && ccli ? ` · CCLI # ${esc(ccli)}` : ""}</span></div>`;
    }

    /* ======================= Presenter ======================= */
    const P = { open: false };
    // A plain full-window layer, not a <dialog>: Chrome won't put a modal dialog into full screen
    const pel = document.createElement("div");
    pel.className = "presenter"; pel.hidden = true; pel.tabIndex = -1;
    pel.setAttribute("role", "dialog"); pel.setAttribute("aria-modal", "true"); pel.setAttribute("aria-label", "Presenting");
    pel.innerHTML = `<div class="p-stage" id="p-stage"></div>
      <div class="p-hud" id="p-hud">
        <button type="button" class="p-btn" data-p="close" aria-label="Stop presenting" title="Stop presenting (Esc)">✕</button>
        <div class="p-info"><b id="p-title"></b><span id="p-where"></span></div>
        <div class="p-jump" id="p-jump" role="group" aria-label="Stanzas to show"></div>
        <button type="button" class="p-btn" data-p="prev" aria-label="Previous slide">◀</button>
        <button type="button" class="p-btn" data-p="next" aria-label="Next slide">▶</button>
        <button type="button" class="p-btn" data-p="black" aria-label="Black screen" title="Black screen (B)">◼</button>
        <button type="button" class="p-btn" data-p="full" aria-label="Full screen" aria-pressed="false" title="Full screen (F)">⛶</button>
      </div>`;
    document.body.appendChild(pel);
    const stage = pel.querySelector("#p-stage"), hud = pel.querySelector("#p-hud");

    async function present(opts) {
      // opts: {title, kind:'hymn'|'chorus', n, seq, start, chorus}
      Object.assign(P, { h: null, sel: null, chorus: null, error: "", note: "" }, opts,
        { i: Math.max(0, Math.min(opts.start || 0, opts.seq.length - 1)), black: false, open: true, wasFull: false, leavingFull: false });
      P.seq = [...opts.seq, { type: "end" }];
      P.returnFocus = document.activeElement;
      pel.querySelector("#p-title").textContent = opts.title;
      renderPicker();
      if (vel.open) { P.reopenViewer = true; vel.close(); } // the viewer is a dialog, which would sit on top
      pel.hidden = false; setInert(true);
      pel.focus({ preventScroll: true }); // not the ✕ button, so Space and Enter move on instead of stopping
      document.body.classList.add("presenting");
      show(); wake(true);
      goFull(false);
      if (opts.kind === "hymn") {
        // The first slides come first so presenting can start at once; the rest follow in a second request
        const names = opts.seq.filter((x) => x.type === "pic").map((x) => x.name);
        const from = names.indexOf((P.seq[P.i] || {}).name);
        const first = names.slice(Math.max(0, from), Math.max(0, from) + 2), rest = names.filter((x) => !first.includes(x));
        const redraw = () => { if (P.open) show(); };
        const fail = (e) => { if (!P.open) return; const msg = why(e); if (msg) { P.error = msg; show(); } };
        loadPics(opts.n, first).then(redraw, fail);
        if (rest.length) loadPics(opts.n, rest).then(redraw, fail);
      }
      flashHud();
    }

    function show() {
      const it = P.seq[P.i];
      let html = "";
      if (P.black || it.type === "end") html = it.type === "end" && !P.black ? `<div class="p-endnote">End · press → to finish</div>` : "";
      else if (it.type === "pic") {
        const u = picUrl(it.name);
        const firstPic = P.seq.findIndex((x) => x.type === "pic") === P.i;
        html = u ? `<div class="p-fit"><img src="${u}" alt="${esc(P.title)}, ${stLabel(it.code)}">${firstPic && ccli ? `<span class="p-ccli">CCLI # ${esc(ccli)}</span>` : ""}</div>`
          : `<div class="p-wait">${P.error ? P.error : "Loading the slides…"}</div>`;
        // get the next picture decoded so the change is instant
        const nx = P.seq[P.i + 1]; if (nx && nx.type === "pic" && picUrl(nx.name)) { const im = new Image(); im.src = picUrl(nx.name); }
      } else if (it.type === "chorus") html = `<div class="p-fit wide">${chorusSlideHtml(P.chorus, it, it.i === 0)}</div>`;
      else if (it.type === "page") html = `<div class="p-page"><img src="${it.url}" alt="${esc(P.title)}, full hymnal page"></div>`;
      stage.innerHTML = html;
      stage.classList.toggle("light", it.type === "page");
      const real = P.seq.length - 1;
      pel.querySelector("#p-where").textContent = P.note ? P.note : it.type === "end" ? "End" : P.black ? "Screen blacked out (B)" :
        `${it.code ? stLabel(it.code) + " · " : ""}${P.i + 1} of ${real}`;
      pel.querySelectorAll("[data-st]").forEach((b) => b.classList.toggle("cur", it.code === b.dataset.st));
      pel.querySelector('[data-p="black"]').setAttribute("aria-pressed", String(P.black));
    }
    function go(i) { if (i < 0) i = 0; if (i >= P.seq.length) return close(); P.i = i; P.black = false; P.note = ""; show(); }
    const next = () => go(P.i + 1), prev = () => go(P.i - 1);
    function jump(code) { const k = P.seq.findIndex((x) => x.code === code); if (k >= 0) go(k); }
    function say(text) { P.note = text; show(); flashHud(); clearTimeout(P.noteTimer); P.noteTimer = setTimeout(() => { P.note = ""; if (P.open) show(); }, 5000); }

    // Stanza picker in the bar: each stanza (and Intro/Ending) turns on or off; the slides follow at once
    function renderPicker() {
      const box = pel.querySelector("#p-jump");
      if (!P.h) { box.innerHTML = ""; return; }
      const h = P.h, s = P.sel;
      const chip = (code, label, on) => `<button type="button" class="p-chip" data-st="${code}" aria-pressed="${on}" title="${stLabel(code)}: ${on ? "showing, tap to leave out" : "left out, tap to show"}">${label}</button>`;
      box.innerHTML = (h.st.length > 1 || h.oi || h.oe ? `<span class="p-label">Stanzas</span>` : "") +
        (h.oi ? chip("OI", "Intro", s.oi) : "") +
        (h.st.length > 1 ? h.st.map((k) => chip(pad(k), k, s.st.includes(k))).join("") : "") +
        (h.oe ? chip("OE", "End", s.oe) : "");
    }
    function toggleStanza(code) {
      const h = P.h, s = P.sel; if (!h) return;
      P.note = "";
      const cur = P.seq[P.i], atEnd = cur.type === "end";
      if (code === "OE") s.oe = !s.oe;
      else if (code === "OI") s.oi = !s.oi;
      else {
        const k = Number(code);
        if (s.st.includes(k)) { if (s.st.length === 1) { say("At least one stanza has to stay in"); return; } s.st = s.st.filter((x) => x !== k); }
        else s.st = [...s.st, k].sort((a, z) => a - z);
      }
      const seq = hymnSeq(h, s), all = hymnSeq(h, { st: h.st, oe: true, oi: true }).map((x) => x.name);
      P.seq = [...seq, { type: "end" }];
      if (atEnd) P.i = seq.length;
      else {
        // stay on the same slide; if its stanza was just left out, go to the next stanza still shown
        const at = all.indexOf(cur.name);
        let i = seq.findIndex((x) => x.name === cur.name);
        if (i < 0) i = seq.findIndex((x) => all.indexOf(x.name) > at);
        P.i = i < 0 ? Math.max(0, seq.length - 1) : i;
      }
      renderPicker(); show();
      const missing = seq.map((x) => x.name).filter((nm) => !picCache.has(nm));
      if (missing.length) loadPics(h.n, missing).then(() => { if (P.open) show(); }, () => {});
    }

    // Full screen hides the browser's tabs and address bar (and the taskbar) so only the slides show. On a laptop
    // whose projector is set up as a second, "extended" screen, the slides go to the projector once Chrome has been
    // allowed to use both screens; it asks the first time the ⛶ button is pressed.
    async function goFull(asked) {
      if (document.fullscreenElement) return true;
      let opt;
      if (window.screen.isExtended && "getScreenDetails" in window) {
        let state = "prompt";
        try { state = (await navigator.permissions.query({ name: "window-management" })).state; } catch (e) { /* older Chrome */ }
        if (state === "granted" || asked) {
          try { const sd = await window.getScreenDetails(); const other = sd.screens.find((x) => x !== sd.currentScreen); if (other) opt = { screen: other }; }
          catch (e) { /* not allowed: use this screen */ }
        }
      }
      if (pel.requestFullscreen) {
        for (const o of opt ? [opt, undefined] : [undefined]) {
          try { await pel.requestFullscreen(o); P.wasFull = true; return true; } catch (e) { /* try the next way */ }
        }
      }
      if (asked) say(document.fullscreenEnabled === false || !pel.requestFullscreen
        ? "This browser can't go full screen here (iPhones can't); turn the device sideways for the biggest slides"
        : "Full screen didn't start. Press ⛶ or F once more");
      return false;
    }
    function leaveFull() { P.leavingFull = true; document.exitFullscreen().catch(() => { P.leavingFull = false; }); }
    document.addEventListener("fullscreenchange", () => {
      const full = !!document.fullscreenElement;
      const b = pel.querySelector('[data-p="full"]');
      b.setAttribute("aria-pressed", String(full)); b.title = full ? "Leave full screen (F); the slides keep showing" : "Full screen (F)";
      b.setAttribute("aria-label", full ? "Leave full screen" : "Full screen");
      if (!P.open) return;
      if (full) { P.wasFull = true; return; }
      if (P.leavingFull) { P.leavingFull = false; return; } // the ⛶ button: keep presenting in the window
      if (P.wasFull) close(); // Esc in full screen: stop, like PowerPoint
    });
    async function wake(on) {
      try {
        if (on && navigator.wakeLock) P.lock = await navigator.wakeLock.request("screen");
        else if (!on && P.lock) { await P.lock.release(); P.lock = null; }
      } catch (e) { /* not supported */ }
    }
    function close() {
      if (!P.open) return;
      P.open = false; P.wasFull = false; P.leavingFull = false; clearTimeout(P.noteTimer);
      if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
      pel.hidden = true; setInert(false); document.body.classList.remove("presenting"); wake(false);
      if (P.reopenViewer) { P.reopenViewer = false; vel.showModal(); }
      if (P.returnFocus && P.returnFocus.focus) P.returnFocus.focus({ preventScroll: true });
    }
    // While presenting, nothing behind the slides can be tabbed to or clicked
    function setInert(on) { for (const el of document.body.children) if (el !== pel && el.tagName !== "SCRIPT") el.inert = on; }

    document.addEventListener("keydown", (e) => {
      if (!P.open || e.altKey || e.ctrlKey || e.metaKey) return;
      const k = e.key;
      if (["ArrowRight", "ArrowDown", "PageDown", " ", "Enter", "n", "N"].includes(k)) { if (e.target.closest && e.target.closest(".p-hud button") && (k === " " || k === "Enter")) return; next(); }
      else if (["ArrowLeft", "ArrowUp", "PageUp", "Backspace", "p", "P"].includes(k)) prev();
      else if (k === "Home") go(0);
      else if (k === "End") go(P.seq.length - 2);
      else if (k === "Escape") close();
      else if (k === "b" || k === "B" || k === "." || k === "w" || k === "W" || k === ",") { P.black = !P.black; show(); }
      else if (/^[1-9]$/.test(k)) jump(pad(k));
      else if (k === "f" || k === "F") document.fullscreenElement ? leaveFull() : goFull(true);
      else if (k === "F5") goFull(true);
      else return;
      e.preventDefault(); e.stopPropagation();
    }, true);

    // Tap/click: right two-thirds forward, left third back; a sideways swipe also works
    let down = null;
    stage.addEventListener("pointerdown", (e) => { down = { x: e.clientX, y: e.clientY }; });
    stage.addEventListener("pointerup", (e) => {
      if (!down) return; const dx = e.clientX - down.x, dy = e.clientY - down.y; down = null;
      if (Math.abs(dx) > 50 && Math.abs(dx) > Math.abs(dy)) { dx < 0 ? next() : prev(); return; }
      if (Math.abs(dx) > 10 || Math.abs(dy) > 10) return;
      if (e.clientX < stage.clientWidth / 3) prev(); else next();
      flashHud();
    });
    hud.addEventListener("click", (e) => {
      const b = e.target.closest("button"); if (!b) return;
      if (b.dataset.st) toggleStanza(b.dataset.st);
      else ({ close, prev, next, black: () => { P.black = !P.black; show(); }, full: () => (document.fullscreenElement ? leaveFull() : goFull(true)) })[b.dataset.p]();
      flashHud();
    });
    let hudTimer = null;
    function flashHud(ms = 2600) { pel.classList.add("hud-on"); clearTimeout(hudTimer); hudTimer = setTimeout(() => { if (!hud.matches(":hover")) pel.classList.remove("hud-on"); }, ms); }
    hud.addEventListener("pointerdown", () => flashHud(5000));
    pel.addEventListener("pointermove", (e) => { if (e.pointerType === "mouse") flashHud(); });

    /* ======================= Viewer ======================= */
    const vel = document.createElement("dialog");
    vel.className = "viewer"; vel.setAttribute("aria-labelledby", "v-title");
    vel.innerHTML = `<div class="v-wrap">
      <header class="v-head"><span class="v-num" id="v-num"></span><div class="v-titles"><h2 id="v-title"></h2><div class="v-sub" id="v-sub"></div></div>
        <button type="button" class="v-close" id="v-close" aria-label="Close">✕</button></header>
      <div class="v-bar"><div class="chips" id="v-chips"></div>
        <div class="v-actions"><button type="button" class="btn" id="v-add">Add to service</button>
          <button type="button" class="btn btn-primary v-present" id="v-present">▶ Present</button></div></div>
      <div class="tabs v-tabs" role="tablist" id="v-tabs"></div>
      <div class="v-body" id="v-body"></div></div>`;
    document.body.appendChild(vel);
    const $v = (s) => vel.querySelector(s);
    const V = { kind: null, n: 0, sel: null, tab: "slides" };
    const player = new M.Player();

    function openHymn(n, opts = {}) {
      const h = hymn(n); if (!h) return;
      if (V.kind !== "hymn" || V.n !== n) player.load(null);
      Object.assign(V, { kind: "hymn", n, sel: { st: [...(opts.st || h.st)], oe: opts.oe != null ? !!opts.oe : h.oe > 0, oi: !!opts.oi } });
      if (!["slides", "page", "words", "play"].includes(V.tab)) V.tab = "slides";
      openDialog();
    }
    function openChorus(n) {
      if (!chorus(n)) return;
      player.load(null);
      Object.assign(V, { kind: "chorus", n, sel: null });
      if (!["slides", "words"].includes(V.tab)) V.tab = "slides";
      openDialog();
    }
    function openDialog() {
      const isH = V.kind === "hymn", x = isH ? hymn(V.n) : chorus(V.n);
      $v("#v-num").textContent = x.n; $v("#v-num").classList.toggle("ch", !isH);
      $v("#v-title").textContent = x.t;
      $v("#v-sub").textContent = isH ? `Songs of Faith and Praise · ${x.st.length} stanza${x.st.length > 1 ? "s" : ""}${x.rf ? " · refrain" : ""}` : `CNCOC Songbook chorus${x.ref ? " · " + x.ref : ""}`;
      const tabs = isH ? [["slides", "Slides"], ["page", "Page"], ["words", "Words"], ["play", "Play"]] : [["slides", "Slides"], ["words", "Words"]];
      $v("#v-tabs").innerHTML = tabs.map(([k, l]) => `<button type="button" role="tab" data-vtab="${k}" aria-selected="${V.tab === k}">${l}</button>`).join("");
      renderChips(); renderTab();
      if (!vel.open) vel.showModal();
      $v("#v-present").focus({ preventScroll: true });
    }
    function renderChips() {
      if (V.kind !== "hymn") { $v("#v-chips").innerHTML = ""; return; }
      const h = hymn(V.n), s = V.sel;
      $v("#v-chips").innerHTML = `<span class="label">Stanzas</span>` +
        h.st.map((k) => `<button type="button" class="chip" data-vst="${k}" aria-pressed="${s.st.includes(k)}" aria-label="Stanza ${k}">${k}</button>`).join("") +
        (h.oi || h.oe ? '<span class="sep"></span>' : "") +
        (h.oi ? `<button type="button" class="chip opt" data-vopt="oi" aria-pressed="${s.oi}">Intro</button>` : "") +
        (h.oe ? `<button type="button" class="chip opt" data-vopt="oe" aria-pressed="${s.oe}">Ending</button>` : "");
      $v("#v-present").disabled = !s.st.length;
    }
    function renderTab() {
      vel.querySelectorAll("[data-vtab]").forEach((b) => b.setAttribute("aria-selected", String(b.dataset.vtab === V.tab)));
      const body = $v("#v-body"), n = V.n, kind = V.kind;
      body.dataset.tab = V.tab;
      const still = () => vel.open && V.n === n && V.kind === kind && body.dataset.tab === V.tab;
      if (V.tab === "slides") return slidesTab(body, still);
      if (V.tab === "page") return pageTab(body, still);
      if (V.tab === "words") return wordsTab(body, still);
      return playTab(body, still);
    }

    function slidesTab(body, still) {
      if (V.kind === "chorus") {
        const c = chorus(V.n), seq = chorusSeq(c);
        body.innerHTML = `<p class="v-hint">Tap a slide to present from it.</p><ol class="v-grid wide">${seq.map((it, i) =>
          `<li><button type="button" class="v-slide" data-from="${i}" aria-label="Present from slide ${i + 1}"><div class="p-fit wide mini">${chorusSlideHtml(c, it, i === 0)}</div></button></li>`).join("")}</ol>`;
        return;
      }
      const h = hymn(V.n), seq = hymnSeq(h, V.sel);
      if (!seq.length) { body.innerHTML = `<p class="v-hint">Choose at least one stanza.</p>`; return; }
      const draw = () => {
        const groups = [];
        seq.forEach((it, i) => { if (!groups.length || groups[groups.length - 1].code !== it.code) groups.push({ code: it.code, items: [] }); groups[groups.length - 1].items.push([it, i]); });
        body.innerHTML = `<p class="v-hint">Tap a slide to present from it. These are the slides for the stanzas chosen above.</p>` + groups.map((g) =>
          `<section class="v-group"><h3 class="v-st">${stLabel(g.code)}</h3><ol class="v-grid">${g.items.map(([it, i]) => {
            const u = picUrl(it.name);
            return `<li><button type="button" class="v-slide" data-from="${i}" aria-label="Present from ${stLabel(it.code)}, slide ${i + 1}">${u ? `<img src="${u}" alt="" loading="lazy">` : `<span class="v-ph">${i + 1}</span>`}</button></li>`;
          }).join("")}</ol></section>`).join("");
      };
      draw();
      const missing = seq.map((x) => x.name).filter((nm) => !picCache.has(nm));
      if (missing.length) loadPics(h.n, missing).then(() => { if (still()) draw(); }, (e) => { if (still()) body.insertAdjacentHTML("afterbegin", `<p class="v-err">${why(e)}</p>`); });
    }

    function pageTab(body, still) {
      const h = hymn(V.n);
      body.innerHTML = `<p class="v-hint">The whole hymn as printed in the hymnal. <button type="button" class="linkbtn" data-pagefull>Show full screen</button> · tap the page to zoom.</p><div class="v-page"><div class="v-wait">Loading the hymnal page…</div></div>`;
      page(h.n).then((u) => {
        if (!still()) return;
        body.querySelector(".v-page").innerHTML = `<img src="${u}" alt="Hymn ${h.n}, ${esc(h.t)}: the full page from the hymnal">`;
      }, (e) => { if (still()) body.querySelector(".v-page").innerHTML = `<p class="v-err">${why(e)}</p>`; });
    }

    const person = (s) => String(s || "").split("@").map((x) => {
      const m = x.trim().match(/^([^,()]+),\s*([^()]+?)\s*(\(.*\))?$/);
      return m ? `${m[2]} ${m[1]}${m[3] ? " " + m[3] : ""}` : x.trim();
    }).filter(Boolean).join("; ");
    const lines = (t) => t.split("\n").map((l) => esc(l.trim())).join("<br>");
    function wordsTab(body, still) {
      if (V.kind === "chorus") {
        const c = chorus(V.n);
        body.innerHTML = `<div class="v-words">${c.slides.flat().map((b) => `<div class="w-st">${b.label ? `<span class="w-n">${esc(b.label)}</span>` : ""}<p>${b.lines.map(esc).join("<br>")}</p></div>`).join("")}</div>`;
        return;
      }
      const h = hymn(V.n);
      body.innerHTML = `<div class="v-wait">Loading the words…</div>`;
      media(h.n).then((w) => {
        if (!still()) return;
        const blocks = w.lyrics.split(/\n\s*\n/).map((b) => b.trim()).filter(Boolean);
        const refrain = (w.refrain || "").replace(/^refrain\s*:?\s*\n?/i, "").trim();
        let html = "", count = 0;
        blocks.forEach((b) => {
          const m = b.match(/^(\d+)\.\s*([\s\S]*)$/), num = m ? +m[1] : null;
          html += `<div class="w-st${num && !V.sel.st.includes(num) ? " off" : ""}"><span class="w-n">${num || ""}</span><p>${lines(m ? m[2] : b)}</p></div>`;
          count++;
          if (count === 1 && refrain) html += `<div class="w-st w-rf"><span class="w-n">Refrain</span><p>${lines(refrain)}</p></div>`;
        });
        const note = h.st.length > count ? `<p class="v-hint">The hymnal prints ${h.st.length} stanzas; this text has ${count}. The Page view shows them all.</p>` : "";
        const c = w.credit || {}, rows = [
          ["Words", [person(c.words_by), c.words_date].filter(Boolean).join(", ") + (c.translator ? `; translated by ${person(c.translator)}` : "")],
          ["Music", [person(c.music_by), c.music_date].filter(Boolean).join(", ")],
          ["Arranged", person(c.arranger)],
          ["Tune", [c.hymn_tune, c.meter].filter(Boolean).join(" · ")],
          ["Key", M.prettyKey(c.key)],
          ["Tempo", [c.tempo ? c.tempo + " beats a minute" : "", c.style].filter(Boolean).join(" · ")],
          ["Scripture", c.scripture],
          ["Copyright", (c.copyright || "").replace(/^copyright\s*/i, "")],
        ].filter((r) => r[1]);
        body.innerHTML = `<div class="v-words">${html}</div>${note}<dl class="v-credit">${rows.map(([k, v]) => `<dt>${k}</dt><dd>${esc(v)}</dd>`).join("")}</dl>`;
      }, (e) => { if (still()) body.innerHTML = `<p class="v-err">${why(e)}</p>`; });
    }

    /* ----- Play ----- */
    const PART = ["", "Soprano", "Alto", "Tenor", "Bass"];
    const fmt = (s) => `${Math.floor(s / 60)}:${pad(Math.floor(s % 60))}`;
    let raf = 0;
    function playTab(body, still) {
      const h = hymn(V.n);
      body.innerHTML = `<div class="v-wait">Loading the music…</div>`;
      media(h.n).then((w) => {
        if (!still()) return;
        if (player.song !== w.song) player.load(w.song);
        const song = w.song, key = (w.credit || {}).key || "";
        const names = song.parts.map((pt) => PART[pt.track] || (w.descant ? "Descant" : "Part " + pt.track));
        const opening = song.parts.map((pt, i) => `${names[i]} ${M.noteName(pt.notes[0].n, key)}`).join(" · ");
        const ios = /iPhone|iPad|iPod/.test(navigator.userAgent) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
        body.innerHTML = `<div class="v-play">
          <div class="pl-row"><button type="button" class="btn btn-primary pl-go" id="pl-go" aria-label="Play">▶ Play</button>
            <span class="pl-time" id="pl-time">0:00</span><input type="range" id="pl-seek" min="0" max="${song.duration.toFixed(2)}" step="0.1" value="0" aria-label="Position">
            <span class="pl-time">${fmt(song.duration)}</span></div>
          <div class="pl-row"><span class="label">Tempo</span><input type="range" id="pl-rate" min="50" max="150" step="5" value="${Math.round(player.rate * 100)}" aria-label="Tempo">
            <span id="pl-bpm" class="pl-bpm"></span><button type="button" class="linkbtn" id="pl-rate0">Reset</button></div>
          <div class="pl-row"><span class="label">Parts</span><div class="chips">${names.map((nm, i) =>
            `<button type="button" class="chip opt" data-part="${i}" aria-pressed="${!player.muted.has(i)}">${nm}</button>`).join("")}</div>
            <span class="v-hint">Tap a part to turn it off, e.g. to practise your own line.</span></div>
          <div class="pl-row"><button type="button" class="btn" id="pl-pitch">♪ Give the starting pitch</button>
            <span class="v-hint">${key ? `Key of ${esc(M.prettyKey(key))} · ` : ""}opening notes: ${esc(opening)}</span></div>
          <p class="v-hint" id="pl-sound"></p>${ios ? `<p class="v-hint">No sound on an iPhone or iPad? Check the silent switch is off.</p>` : ""}</div>`;
        const book = (w.credit || {}).tempo;
        const tick = () => {
          if (!still()) { cancelAnimationFrame(raf); return; }
          const pos = player.position();
          body.querySelector("#pl-time").textContent = fmt(pos);
          const sk = body.querySelector("#pl-seek"); if (!sk.matches(":active")) sk.value = pos.toFixed(2);
          if (player.playing) raf = requestAnimationFrame(tick);
        };
        player.onchange = () => {
          if (!still()) return;
          const go = body.querySelector("#pl-go");
          go.textContent = player.playing ? "❚❚ Pause" : player.sound === "loading" ? "Loading…" : "▶ Play";
          go.setAttribute("aria-label", player.playing ? "Pause" : "Play");
          body.querySelector("#pl-bpm").textContent = `${Math.round(player.rate * 100)}% speed${book ? ` · hymnal marking ${book}` : ""}`;
          body.querySelector("#pl-sound").textContent = player.sound === "tone" ? "Sound: a simple built-in tone (the piano sound couldn't be loaded)." : player.sound === "loading" ? "Loading the piano sound…" : "";
          body.querySelectorAll("[data-part]").forEach((b) => b.setAttribute("aria-pressed", String(!player.muted.has(+b.dataset.part))));
          cancelAnimationFrame(raf); tick();
        };
        player.onchange();
      }, (e) => { if (still()) body.innerHTML = `<p class="v-err">${why(e)}</p>`; });
    }

    vel.addEventListener("click", (e) => {
      const b = e.target.closest("button") || e.target.closest(".v-page img");
      if (e.target === vel) { closeViewer(); return; } // a click on the dimmed backdrop
      if (!b) return;
      const h = V.kind === "hymn" ? hymn(V.n) : null;
      if (b.id === "v-close") return closeViewer();
      if (b.dataset.vtab) { V.tab = b.dataset.vtab; renderTab(); return; }
      if (b.dataset.vst) { const k = +b.dataset.vst; V.sel.st = V.sel.st.includes(k) ? V.sel.st.filter((x) => x !== k) : [...V.sel.st, k].sort((a, z) => a - z); renderChips(); if (V.tab !== "play") renderTab(); return; }
      if (b.dataset.vopt) { V.sel[b.dataset.vopt] = !V.sel[b.dataset.vopt]; renderChips(); if (V.tab === "slides") renderTab(); return; }
      if (b.id === "v-present") return presentCurrent(0);
      if (b.dataset.from != null) return presentCurrent(+b.dataset.from);
      if (b.id === "v-add") { add(V.kind, V.n, V.kind === "hymn" ? { ...V.sel, st: [...V.sel.st] } : null); closeViewer(); return; }
      if (b.dataset.pagefull != null) { page(V.n).then((u) => present({ title: `${h.n} ${h.t}`, kind: "page", n: h.n, seq: [{ type: "page", url: u }] }), () => {}); return; }
      if (b.tagName === "IMG" && b.closest(".v-page")) { b.closest(".v-page").classList.toggle("zoom"); return; }
      if (b.id === "pl-go") { player.playing ? player.pause() : player.play().catch((err) => alertIn(err)); return; }
      if (b.id === "pl-pitch") { player.pitch().catch((err) => alertIn(err)); return; }
      if (b.id === "pl-rate0") { player.setRate(1); vel.querySelector("#pl-rate").value = 100; return; }
      if (b.dataset.part != null) { const i = +b.dataset.part; player.setMuted(i, !player.muted.has(i)); return; }
    });
    vel.addEventListener("input", (e) => {
      if (e.target.id === "pl-rate") player.setRate(+e.target.value / 100);
      if (e.target.id === "pl-seek") player.seek(+e.target.value);
    });
    vel.addEventListener("close", () => { player.pause(); });
    function alertIn(err) { const s = vel.querySelector("#pl-sound"); if (s) s.textContent = "Couldn't play: " + (err && err.message ? err.message : err); }
    function closeViewer() { player.pause(); vel.close(); }

    function presentCurrent(start) {
      if (V.kind === "chorus") return presentChorus(V.n, start);
      if (!V.sel.st.length) return;
      presentHymn(V.n, V.sel, start);
    }
    function presentHymn(n, sel, start = 0) {
      const h = hymn(n); if (!h) return;
      sel = sel || { st: [...h.st], oe: h.oe > 0, oi: false };
      player.pause();
      sel = { st: [...sel.st], oe: !!sel.oe, oi: !!sel.oi };
      present({ title: `${h.n} ${h.t}`, kind: "hymn", n: h.n, h, sel, seq: hymnSeq(h, sel), start });
    }
    function presentChorus(n, start = 0) {
      const c = chorus(n); if (!c) return;
      present({ title: `Chorus ${c.n} ${c.t}`, kind: "chorus", n: c.n, chorus: c, seq: chorusSeq(c), start });
    }

    return { openHymn, openChorus, presentHymn, presentChorus };
  };
})();
