/* Hymn MIDI player for the Service Builder.
 *
 * The SOFTPraise MIDI files have one track per voice part (soprano, alto, tenor, bass, and for 46 hymns a
 * descant). This reads them and plays them through Web Audio: a piano sound when it can be loaded, otherwise
 * a simple built-in tone. Tempo, which parts sound and the starting pitch can be controlled.
 *
 *   const song = HymnMidi.parse(bytes)      -> {parts:[{track, notes:[{t,d,n,v}]}], duration, bpm}
 *   const p = new HymnMidi.Player(); p.load(song); p.play(); p.pause(); p.seek(s); p.setRate(1.1);
 *   p.setMuted(partIndex, true); p.pitch(); p.position(); p.onchange = () => {...}
 */
(function () {
  "use strict";
  const PIANO_URL = "https://cdn.jsdelivr.net/npm/smplr@1.2.0/+esm";

  /* ---------- Reading a standard MIDI file ---------- */
  function parse(bytes) {
    const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let p = 0;
    const id = () => String.fromCharCode(u8[p], u8[p + 1], u8[p + 2], u8[p + 3]);
    const u32 = () => { const v = dv.getUint32(p); p += 4; return v; };
    const vlq = () => { let v = 0, b; do { b = u8[p++]; v = v * 128 + (b & 0x7f); } while (b & 0x80 && p < u8.length); return v; };
    if (id() !== "MThd") throw new Error("This isn't a MIDI file");
    p += 4; const hlen = u32();
    const ntracks = dv.getUint16(p + 2), division = dv.getUint16(p + 4);
    p += hlen;
    if (division & 0x8000) throw new Error("This MIDI file uses a timing format that isn't supported");
    const tempos = [], tracks = [];
    while (p + 8 <= u8.length && tracks.length < ntracks) {
      const chunk = id(); p += 4; const len = u32(), end = p + len;
      if (chunk !== "MTrk") { p = end; continue; }
      let tick = 0, status = 0, name = "";
      const open = new Map(), notes = [];
      while (p < end) {
        tick += vlq();
        if (u8[p] & 0x80) status = u8[p++];
        if (status === 0xff) {
          const type = u8[p++], l = vlq();
          if (type === 0x51 && l === 3) tempos.push([tick, (u8[p] << 16) | (u8[p + 1] << 8) | u8[p + 2]]);
          else if (type === 0x03) name = String.fromCharCode(...u8.subarray(p, p + l));
          p += l; status = 0;
        } else if (status === 0xf0 || status === 0xf7) {
          p += vlq(); status = 0;
        } else {
          const hi = status & 0xf0, key = ((status & 0x0f) << 7) | u8[p];
          const a = u8[p++], b = hi === 0xc0 || hi === 0xd0 ? 0 : u8[p++];
          if (hi === 0x90 && b > 0) {
            if (!open.has(key)) open.set(key, []);
            open.get(key).push({ tick, n: a, v: b });
          } else if (hi === 0x80 || hi === 0x90) {
            const q = open.get(key);
            if (q && q.length) { const o = q.shift(); notes.push({ s: o.tick, e: tick, n: o.n, v: o.v }); }
          }
        }
      }
      p = end;
      tracks.push({ name, notes });
    }
    // Ticks to seconds through the tempo map (default 120 bpm until the first tempo change)
    tempos.sort((x, y) => x[0] - y[0]);
    const map = [{ tick: 0, sec: 0, us: 500000 }];
    for (const [tk, us] of tempos) {
      const last = map[map.length - 1];
      const sec = last.sec + ((tk - last.tick) * last.us) / division / 1e6;
      if (tk === last.tick) last.us = us; else map.push({ tick: tk, sec, us });
    }
    const toSec = (tk) => {
      let k = map.length - 1; while (k > 0 && map[k].tick > tk) k--;
      return map[k].sec + ((tk - map[k].tick) * map[k].us) / division / 1e6;
    };
    // track: position in the file (0 is the tempo track, then soprano, alto, tenor, bass, descant)
    const parts = tracks.map((t, k) => ({ ...t, track: k })).filter((t) => t.notes.length).map((t) => ({
      name: t.name, track: t.track,
      notes: t.notes.map((x) => { const t0 = toSec(x.s); return { t: t0, d: Math.max(0.05, toSec(x.e) - t0), n: x.n, v: x.v }; }).sort((x, y) => x.t - y.t),
    }));
    const duration = parts.reduce((m, pt) => pt.notes.reduce((mm, x) => Math.max(mm, x.t + x.d), m), 0);
    return { parts, duration, bpm: Math.round(60e6 / map[0].us) };
  }

  /* ---------- Note names ---------- */
  const SHARP = ["C", "C♯", "D", "D♯", "E", "F", "F♯", "G", "G♯", "A", "A♯", "B"];
  const FLAT = ["C", "D♭", "D", "E♭", "E", "F", "G♭", "G", "A♭", "A", "B♭", "B"];
  function noteName(n, key) {
    const flats = /^[A-G]b/.test(key || "") || /^(F|Dm|Gm|Cm|Fm)$/.test(key || "");
    return (flats ? FLAT : SHARP)[((n % 12) + 12) % 12];
  }
  const prettyKey = (k) => String(k || "").replace(/^([A-G])b/, "$1♭").replace(/^([A-G])#/, "$1♯").replace(/m$/, " minor");

  /* ---------- Player ---------- */
  class Player {
    constructor() {
      this.ctx = null; this.song = null; this.rate = 1; this.muted = new Set();
      this.playing = false; this.pos = 0; this.voices = []; this.timer = null;
      this.sound = "none"; // "loading" | "piano" | "tone"
      this.onchange = null;
    }
    emit() { if (this.onchange) try { this.onchange(); } catch (e) { /* the page's problem */ } }

    async audio() {
      if (!this.ctx) {
        const AC = window.AudioContext || window.webkitAudioContext;
        if (!AC) throw new Error("This browser can't play sound");
        this.ctx = new AC();
        const comp = this.ctx.createDynamicsCompressor();
        this.out = this.ctx.createGain(); this.out.gain.value = 0.9;
        this.out.connect(comp); comp.connect(this.ctx.destination);
      }
      if (this.ctx.state !== "running") await this.ctx.resume();
      return this.ctx;
    }
    // Piano samples come from the web; if they can't be had within 10 seconds, the built-in tone is used
    piano() {
      if (this.pianoReady) return this.pianoReady;
      this.sound = "loading"; this.emit();
      this.pianoReady = (async () => {
        try {
          const lib = await import(PIANO_URL);
          const inst = lib.Soundfont(this.ctx, { instrument: "acoustic_grand_piano", destination: this.out });
          await Promise.race([inst.ready || inst.load, new Promise((_, no) => setTimeout(() => no(new Error("timeout")), 10000))]);
          const lp = inst.loadProgress; // samples that didn't arrive would play as silence
          if (lp && lp.total && lp.loaded < lp.total) throw new Error("piano samples missing");
          this.inst = inst; this.sound = "piano";
        } catch (e) { this.inst = null; this.sound = "tone"; }
        this.emit();
      })();
      return this.pianoReady;
    }

    load(song) { this.stop(); this.song = song; this.pos = 0; this.muted = new Set(); this.emit(); }
    position() {
      if (!this.song) return 0;
      if (!this.playing) return this.pos;
      return Math.min(this.song.duration, Math.max(this.from, this.from + (this.ctx.currentTime - this.t0) * this.rate));
    }

    note(n, v, when, dur) {
      const c = this.ctx;
      if (this.inst) {
        const stop = this.inst.start({ note: n, velocity: Math.min(127, Math.round(v * 0.9)), time: when, duration: dur });
        return () => { try { if (typeof stop === "function") stop(); } catch (e) { /* already ended */ } };
      }
      // Built-in tone: a soft triangle with a quiet octave, shaped like a struck note
      const f = 440 * Math.pow(2, (n - 69) / 12), g = c.createGain(), o = c.createOscillator(), o2 = c.createOscillator(), g2 = c.createGain();
      o.type = "triangle"; o.frequency.value = f; o2.type = "sine"; o2.frequency.value = f * 2; g2.gain.value = 0.18;
      o.connect(g); o2.connect(g2); g2.connect(g); g.connect(this.out);
      const peak = 0.16 * (v / 127), end = when + dur;
      g.gain.setValueAtTime(0, when);
      g.gain.linearRampToValueAtTime(peak, when + 0.015);
      g.gain.setTargetAtTime(peak * 0.55, when + 0.015, 0.25);
      g.gain.setTargetAtTime(0, end, 0.05);
      o.start(when); o2.start(when); o.stop(end + 0.35); o2.stop(end + 0.35);
      return () => {
        try { const t = c.currentTime; g.gain.cancelScheduledValues(t); g.gain.setTargetAtTime(0, t, 0.015); o.stop(t + 0.08); o2.stop(t + 0.08); } catch (e) { /* not started or ended */ }
      };
    }

    async play() {
      if (!this.song || this.playing) return;
      await this.audio();
      await this.piano();
      if (this.playing) return;
      if (this.pos >= this.song.duration - 0.05) this.pos = 0;
      this.queue = this.song.parts.flatMap((pt, i) => (this.muted.has(i) ? [] : pt.notes)).filter((x) => x.t >= this.pos - 0.001).sort((a, b) => a.t - b.t);
      this.qi = 0; this.from = this.pos; this.t0 = this.ctx.currentTime + 0.08; this.playing = true;
      this.timer = setInterval(() => this.pump(), 120);
      this.pump(); this.emit();
    }
    // Schedules the notes due in the next moment, so pausing or changing tempo takes effect straight away
    pump() {
      const now = this.ctx.currentTime, horizon = now + 1.0;
      this.voices = this.voices.filter((x) => x.end > now);
      while (this.qi < this.queue.length) {
        const x = this.queue[this.qi], when = this.t0 + (x.t - this.from) / this.rate;
        if (when > horizon) break;
        if (when >= now - 0.03) { const d = x.d / this.rate; this.voices.push({ stop: this.note(x.n, x.v, Math.max(when, now), d), end: when + d + 0.4 }); }
        this.qi++;
      }
      if (this.position() >= this.song.duration) { this.stop(); this.pos = 0; this.emit(); }
    }
    silence() { clearInterval(this.timer); this.timer = null; this.voices.forEach((x) => x.stop()); this.voices = []; }
    pause() { if (!this.playing) return; this.pos = this.position(); this.playing = false; this.silence(); this.emit(); }
    stop() { this.silence(); this.playing = false; }
    async restartIf(change) { const was = this.playing; if (was) this.pause(); change(); if (was) await this.play(); else this.emit(); }
    seek(sec) { return this.restartIf(() => { this.pos = Math.max(0, Math.min(this.song ? this.song.duration : 0, sec)); }); }
    setRate(r) { return this.restartIf(() => { this.rate = Math.max(0.4, Math.min(1.8, r)); }); }
    setMuted(i, m) { return this.restartIf(() => { if (m) this.muted.add(i); else this.muted.delete(i); }); }

    // The opening note of each sounding part, held together for two seconds
    async pitch() {
      if (!this.song) return [];
      this.pause();
      await this.audio(); await this.piano();
      const firsts = this.song.parts.map((pt, i) => (this.muted.has(i) ? null : pt.notes[0])).filter(Boolean);
      const now = this.ctx.currentTime + 0.05;
      firsts.forEach((x) => this.voices.push({ stop: this.note(x.n, 90, now, 2.2), end: now + 2.6 }));
      return firsts;
    }
  }

  window.HymnMidi = { parse, Player, noteName, prettyKey };
})();
