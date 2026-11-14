/* Service deck generator.
 * Builds a .pptx from the church's transition-slide deck (the base package) plus
 * SOFTPraise-style hymn slides. Runs in the browser (window.ServiceDeck) and in Node (module.exports).
 *
 * buildDeck({JSZip, base, tpl, items, getImage, ccli}) -> Promise<Uint8Array>
 *   base     : bytes of Transition_Slides.pptx
 *   tpl      : {title, first, plain, end} slide XML strings from a SOFTPraise export
 *   items    : [{type:'transition', key}, {type:'reading', ref}, {type:'teacher', label},
 *               {type:'hymn', number, title, slides:[{name, path}]},
 *               {type:'chorus', number, title, ref, slides:[[{label, lines:[...]}, ...], ...]}]
 *   getImage : async ({name, path}) -> {bytes: Uint8Array, ext: 'gif'|'png'|'jpg'}
 *   date     : service date 'YYYY-MM-DD' (optional). The Bible Class room slides (bibleSun, bibleWed)
 *              are only shown on the last Sunday/Wednesday of the month, i.e. when the date is in the
 *              month's last week. Without a date they are always shown.
 *
 * A hymn gets a black title slide before it only when it follows another hymn (or starts the deck),
 * not when it directly follows a transition, reading, chorus or teacher's slides.
 * Choruses (CNCOC Songbook) are drawn as text slides: number, title, "1 of 3" cue, words.
 */
(function (root) {
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const REL_SLIDE = R + "/slide", REL_LAYOUT = R + "/slideLayout", REL_IMAGE = R + "/image";
  const CT_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
  const MIME = { gif: "image/gif", png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg" };
  // Order of slides in Transition_Slides.pptx
  const TRANSITION_INDEX = { welcome: 0, supper: 1, reading: 2, closing: 3, bibleSun: 4, bibleWed: 5, prayer: 6 };

  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const dirOf = (p) => p.slice(0, p.lastIndexOf("/"));
  function resolve(fromPart, target) {
    const parts = (dirOf(fromPart) + "/" + target).split("/");
    const out = [];
    for (const s of parts) { if (s === "..") out.pop(); else if (s && s !== ".") out.push(s); }
    return out.join("/");
  }
  function relTo(fromPart, absPath) { // ppt/slides/x.xml -> ../media/y
    return "../" + absPath.replace(/^ppt\//, "");
  }
  function relative(fromPart, toPart) { // ppt/slides/a.xml, ppt/media/b.png -> ../media/b.png
    const from = dirOf(fromPart).split("/"), to = toPart.split("/");
    let i = 0;
    while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
    return [...Array(from.length - i).fill(".."), ...to.slice(i)].join("/");
  }
  const relsPathOf = (p) => dirOf(p) + "/_rels/" + p.split("/").pop() + ".rels";
  function parseRels(xml) {
    const rels = [];
    xml.replace(/<Relationship\b[^>]*>/g, (tag) => {
      const g = (a) => (tag.match(new RegExp(a + '="([^"]*)"')) || [])[1];
      rels.push({ id: g("Id"), type: g("Type"), target: g("Target"), mode: g("TargetMode") });
      return tag;
    });
    return rels;
  }
  function relsXml(rels) {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      rels.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${esc(r.target)}"${r.mode ? ` TargetMode="${r.mode}"` : ""}/>`).join("") +
      "</Relationships>";
  }

  /* ----- Chorus slides (CNCOC Songbook) ----- */
  const C = { maroon: "6E1E2B", soft: "F4E7E9", ink: "1C2130", muted: "6B7280", rule: "D9DCE4", font: "Calibri" };
  const W = 12192000, M = 640080;
  const BOX = { x: M, y: 1234440, w: W - 2 * M, h: 5120640 };
  // Largest font size (pt) at which the words fit the box, allowing for wrapped lines
  function chorusFontSize(blocks) {
    const widthPt = (BOX.w - 182880) / 12700, heightPt = (BOX.h - 91440) / 12700;
    for (let sz = 44; sz >= 20; sz -= 2) {
      const perLine = widthPt / (0.5 * sz);
      let h = 0;
      blocks.forEach((b, i) => {
        if (i) h += 0.55 * sz * 1.15;
        if (b.label) h += 0.62 * sz * 1.25;
        for (const l of b.lines) h += Math.max(1, Math.ceil(l.length / perLine)) * sz * 1.15;
      });
      if (h <= heightPt) return sz;
    }
    return 20;
  }
  const run = (text, { sz, b, i, color, spc, caps } = {}) =>
    `<a:r><a:rPr lang="en-AU" sz="${Math.round(sz * 100)}"${b ? ' b="1"' : ""}${i ? ' i="1"' : ""}${caps ? ' cap="all"' : ""}${spc ? ` spc="${spc}"` : ""} dirty="0"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="${C.font}"/><a:cs typeface="${C.font}"/></a:rPr><a:t>${esc(text)}</a:t></a:r>`;
  const box = (id, name, x, y, w, h, body, { fill, round, ph, anchor = "ctr", inset = 91440 } = {}) =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr${ph ? '><a:spLocks noGrp="1"/></p:cNvSpPr>' : ' txBox="1"/>'}<p:nvPr>${ph ? '<p:ph type="title"/>' : ""}</p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="${round ? "roundRect" : "rect"}"><a:avLst/></a:prstGeom>${fill ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>` : "<a:noFill/>"}<a:ln><a:noFill/></a:ln></p:spPr>` +
    `<p:txBody><a:bodyPr wrap="square" lIns="${inset}" rIns="${inset}" tIns="45720" bIns="45720" anchor="${anchor}"><a:noAutofit/></a:bodyPr><a:lstStyle/>${body}</p:txBody></p:sp>`;
  const para = (runs, { algn = "l", sz = 18, before = 0 } = {}) =>
    `<a:p><a:pPr algn="${algn}">${before ? `<a:spcBef><a:spcPts val="${Math.round(before * 100)}"/></a:spcBef>` : ""}<a:lnSpc><a:spcPct val="100000"/></a:lnSpc></a:pPr>${runs}<a:endParaRPr lang="en-AU" sz="${Math.round(sz * 100)}"/></a:p>`;

  function chorusSlideXml(ch, blocks, i, total, ccli) {
    const sz = chorusFontSize(blocks);
    let words = "";
    blocks.forEach((b, k) => {
      const gap = k ? 0.55 * sz : 0;
      if (b.label) words += para(run(b.label, { sz: 0.62 * sz, i: true, color: C.maroon, spc: 50 }), { algn: "ctr", sz: 0.62 * sz, before: gap });
      b.lines.forEach((l, q) => { words += para(run(l, { sz, color: C.ink }), { algn: "ctr", sz, before: q === 0 && !b.label ? gap : 0 }); });
    });
    const title = run(ch.title, { sz: 24, b: true, color: C.ink, spc: 80, caps: true }) + (ch.ref ? run("   " + ch.ref, { sz: 16, color: C.muted }) : "");
    const shapes = [
      box(10, "Chorus number", M, 384048, 1005840, 548640, para(run(String(ch.number), { sz: 24, b: true, color: "FFFFFF" }), { algn: "ctr", sz: 24 }), { fill: C.maroon, round: true, inset: 0 }),
      box(2, "Title", M + 1188720, 338328, 7863840, 640080, para(title, { sz: 24 }), { ph: true, inset: 0 }),
      box(11, "Slide cue", W - M - 1554480, 402336, 1554480, 512064, para(run(`${i + 1} of ${total}`, { sz: 20, b: true, color: C.maroon }), { algn: "ctr", sz: 20 }), { fill: C.soft, round: true, inset: 0 }),
      `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="12" name="Rule"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="${M}" y="1106424"/><a:ext cx="${W - 2 * M}" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="12700"><a:solidFill><a:srgbClr val="${C.rule}"/></a:solidFill></a:ln></p:spPr></p:cxnSp>`,
      box(13, "Words", BOX.x, BOX.y, BOX.w, BOX.h, words),
      box(14, "Footer", M, 6400800, W - 2 * M, 320040,
        para(run(`CNCOC Songbook ${ch.number}${i === 0 && ccli ? `   ·   CCLI # ${ccli}` : ""}`, { sz: 11, color: C.muted }), { algn: "r", sz: 11 }), { anchor: "b", inset: 0 }),
    ];
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
      shapes.join("") + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
  }

  // True when the date is the last of its weekday in the month (a week later is next month)
  function isLastOfMonth(date) {
    const d = new Date(date + "T12:00:00Z");
    if (isNaN(d)) throw new Error("Service date should look like 2026-11-15, not " + date);
    const later = new Date(d); later.setUTCDate(d.getUTCDate() + 7);
    return later.getUTCMonth() !== d.getUTCMonth();
  }

  async function buildDeck({ JSZip, base, tpl, items, getImage, ccli, date }) {
    const zip = await JSZip.loadAsync(base);
    const read = (p) => zip.file(p).async("string");
    let ct = await read("[Content_Types].xml");
    let pres = await read("ppt/presentation.xml");
    const presRelsPath = "ppt/_rels/presentation.xml.rels";
    let presRels = parseRels(await read(presRelsPath));

    // Original slides in presentation order
    const order = [...pres.matchAll(/<p:sldId\b[^>]*r:id="([^"]+)"[^>]*\/>/g)].map((m) => m[1]);
    const srcSlides = [];
    for (const rid of order) {
      const rel = presRels.find((r) => r.id === rid);
      const path = resolve("ppt/presentation.xml", rel.target);
      const relsPath = dirOf(path) + "/_rels/" + path.split("/").pop() + ".rels";
      srcSlides.push({ path, xml: await read(path), rels: zip.file(relsPath) ? parseRels(await read(relsPath)) : [] });
    }
    // Content-type overrides for every part (to copy for duplicated diagram parts)
    const overrides = {};
    ct.replace(/<Override PartName="([^"]+)" ContentType="([^"]+)"\/>/g, (_, p, t) => { overrides[p] = t; return _; });

    // Remove original slides
    for (const s of srcSlides) {
      zip.remove(s.path);
      zip.remove(dirOf(s.path) + "/_rels/" + s.path.split("/").pop() + ".rels");
      ct = ct.replace(`<Override PartName="/${s.path}" ContentType="${CT_SLIDE}"/>`, "");
    }
    presRels = presRels.filter((r) => r.type !== REL_SLIDE);

    // Layout for hymn slides: "Title Only"
    let hymnLayout = null;
    for (const f of Object.keys(zip.files).filter((f) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(f)).sort()) {
      if (/<p:cSld name="Title Only"/.test(await read(f))) { hymnLayout = f; break; }
    }
    if (!hymnLayout) throw new Error("No 'Title Only' layout in the transition deck");

    const newSlides = []; // {path}
    const usedCopies = {}; // part path -> count used
    const addedMedia = {}; // key -> abs path
    let n = 0;

    function addSlide(xml, rels) {
      n += 1;
      const path = `ppt/slides/slide${n}.xml`;
      zip.file(path, xml);
      zip.file(`ppt/slides/_rels/slide${n}.xml.rels`, relsXml(rels));
      ct = ct.replace("</Types>", `<Override PartName="/${path}" ContentType="${CT_SLIDE}"/></Types>`);
      newSlides.push(path);
    }

    async function cloneTransition(key, edit) {
      const src = srcSlides[TRANSITION_INDEX[key]];
      if (!src) throw new Error("Transition slide missing: " + key);
      let xml = src.xml;
      if (edit) xml = edit(xml);
      const rels = [];
      for (const r of src.rels) {
        if (/\/diagram/.test(r.type)) {
          // Diagram parts can't be shared between slides: copy them on second and later use
          const abs = resolve(src.path, r.target);
          usedCopies[abs] = (usedCopies[abs] || 0) + 1;
          let target = abs;
          if (usedCopies[abs] > 1) {
            target = abs.replace(/\.xml$/, `c${usedCopies[abs]}.xml`);
            zip.file(target, await read(abs));
            const t = overrides["/" + abs];
            if (t) ct = ct.replace("</Types>", `<Override PartName="/${target}" ContentType="${t}"/></Types>`);
          }
          rels.push({ ...r, target: relTo(src.path, target) });
        } else rels.push({ ...r });
      }
      addSlide(xml, rels);
    }

    async function addImage(slide) {
      if (addedMedia[slide.name]) return addedMedia[slide.name];
      const img = await getImage(slide);
      const ext = img.ext.toLowerCase();
      const abs = `ppt/media/hymn-${slide.name}.${ext}`;
      zip.file(abs, img.bytes);
      if (!new RegExp(`<Default Extension="${ext}"`, "i").test(ct))
        ct = ct.replace("<Default ", `<Default Extension="${ext}" ContentType="${MIME[ext]}"/><Default `);
      addedMedia[slide.name] = abs;
      return abs;
    }

    /* ----- Importing a teacher's PowerPoint ----- */
    // IDs for slide masters and layouts share one number space (>= 2147483648) and must be unique
    let nextMasterId = 2147483648;
    const bumpIds = (xml) => { for (const m of xml.matchAll(/<p:(?:sldMasterId|sldLayoutId) id="(\d+)"/g)) nextMasterId = Math.max(nextMasterId, +m[1] + 1); };
    bumpIds(pres);
    for (const f of Object.keys(zip.files).filter((f) => /^ppt\/slideMasters\/[^/]+\.xml$/.test(f))) bumpIds(await read(f));
    const defaults = {};
    ct.replace(/<Default Extension="([^"]+)" ContentType="([^"]+)"\/>/g, (_, e, t) => { defaults[e.toLowerCase()] = t; return _; });
    let notesMaster = null; // only one notes master is allowed per presentation
    const extraPresRels = [];
    const masterEntries = [];
    let importCount = 0;
    const DROP_REL = /\/(comments|commentAuthors|authors)$|\/relationships\/comments$/;

    async function importDeck(bytes) {
      importCount += 1;
      const tag = `t${importCount}_`;
      const tz = await JSZip.loadAsync(bytes);
      const tread = (p) => tz.file(p).async("string");
      const tct = await tread("[Content_Types].xml");
      const tDefaults = {}, tOverrides = {};
      tct.replace(/<Default Extension="([^"]+)" ContentType="([^"]+)"\/>/g, (_, e, t) => { tDefaults[e.toLowerCase()] = t; return _; });
      tct.replace(/<Override PartName="([^"]+)" ContentType="([^"]+)"\/>/g, (_, p, t) => { tOverrides[p] = t; return _; });
      const mainRel = parseRels(await tread("_rels/.rels")).find((r) => /officeDocument$/.test(r.type));
      const tPresPath = resolve("x", mainRel.target.replace(/^\//, ""));
      const tPres = await tread(tPresPath);
      const tPresRels = parseRels(await tread(relsPathOf(tPresPath)));
      const tOrder = [...tPres.matchAll(/<p:sldId\b[^>]*r:id="([^"]+)"[^>]*\/>/g)]
        .map((m) => resolve(tPresPath, tPresRels.find((r) => r.id === m[1]).target));
      const map = {};
      tOrder.forEach((p, i) => { map[p] = `ppt/slides/slide${n + 1 + i}.xml`; });
      const done = new Set();

      async function copyPart(p) {
        if (done.has(p)) return map[p];
        if (!tz.file(p)) return null;
        const type = tOverrides["/" + p] || tDefaults[(p.split(".").pop() || "").toLowerCase()];
        if (/notesMaster\+xml$/.test(type) && notesMaster && !map[p]) { map[p] = notesMaster; done.add(p); return notesMaster; }
        done.add(p);
        const np = map[p] || (map[p] = `${dirOf(p)}/${tag}${p.split("/").pop()}`);
        const rp = relsPathOf(p);
        if (tz.file(rp)) {
          const rels = [];
          for (const r of parseRels(await tread(rp))) {
            if (DROP_REL.test(r.type)) continue;
            if (r.mode === "External") { rels.push(r); continue; }
            const tgt = resolve(p, r.target);
            if (r.type === REL_SLIDE && !map[tgt]) continue; // link to a slide we don't have
            const ntgt = r.type === REL_SLIDE ? map[tgt] : await copyPart(tgt);
            if (!ntgt) continue;
            rels.push({ ...r, target: relative(np, ntgt) });
          }
          zip.file(relsPathOf(np), relsXml(rels));
        }
        zip.file(np, await tz.file(p).async("uint8array"));
        const ext = (np.split(".").pop() || "").toLowerCase();
        if (tOverrides["/" + p]) ct = ct.replace("</Types>", `<Override PartName="/${np}" ContentType="${tOverrides["/" + p]}"/></Types>`);
        else if (!defaults[ext] && tDefaults[ext]) { defaults[ext] = tDefaults[ext]; ct = ct.replace("<Default ", `<Default Extension="${ext}" ContentType="${tDefaults[ext]}"/><Default `); }
        if (/slideMaster\+xml$/.test(type)) masterEntries.push(np);
        if (/notesMaster\+xml$/.test(type)) notesMaster = np;
        return np;
      }

      for (const p of tOrder) {
        await copyPart(p);
        n += 1;
        newSlides.push(map[p]);
      }
      const size = (tPres.match(/<p:sldSz cx="(\d+)" cy="(\d+)"/) || []).slice(1).map(Number);
      return { slides: tOrder.length, size };
    }
    const importNotes = [];

    const layoutRel = { id: "rId1", type: REL_LAYOUT, target: relTo("ppt/slides/x.xml", hymnLayout) };
    const setTitle = (xml, text) => xml.replace(/(<p:sp>(?:(?!<\/p:sp>).)*?<a:t>)[^<]*(<\/a:t>)/s, `$1${esc(text)}$2`);

    const skipped = [];
    const showRooms = !date || isLastOfMonth(date);
    let prev = null; // previous item actually placed in the deck
    for (const it of items) {
      if (it.type === "transition" && (it.key === "bibleSun" || it.key === "bibleWed") && !showRooms) { skipped.push(it.key); continue; }
      const afterTransition = !!prev && prev.type !== "hymn";
      prev = it;
      if (it.type === "transition") await cloneTransition(it.key);
      else if (it.type === "reading") {
        const ref = (it.ref || "").trim();
        await cloneTransition("reading", (x) => x.replace("<a:t>Reading:</a:t>", `<a:t>Reading:${ref ? " " + esc(ref) : ""}</a:t>`));
      } else if (it.type === "teacher" && it.file) {
        let res;
        try { res = await importDeck(it.file); }
        catch (e) { throw new Error(`Couldn't read the ${it.label || "teacher's"} PowerPoint. Make sure it's a .pptx file (in PowerPoint: File > Save As > PowerPoint Presentation).`); }
        importNotes.push({ label: it.label, ...res });
      } else if (it.type === "teacher") {
        addSlide(setTitle(tpl.title, it.label || "Teacher's PowerPoint")
          .replace(/<\/p:spTree>/, `<p:sp><p:nvSpPr><p:cNvPr id="9" name="Note"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="6096000"/><a:ext cx="12192000" cy="461665"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-AU" sz="1400"><a:solidFill><a:srgbClr val="404040"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>${esc(it.label || "Teacher's PowerPoint")}: switch to the teacher's slides</a:t></a:r></a:p></p:txBody></p:sp></p:spTree>`),
          [layoutRel]);
      } else if (it.type === "chorus") {
        it.slides.forEach((blocks, i) => addSlide(chorusSlideXml(it, blocks, i, it.slides.length, ccli), [layoutRel]));
      } else if (it.type === "hymn") {
        if (!afterTransition) addSlide(setTitle(tpl.title, `${String(it.number).padStart(4, "0")}     ${it.title}`), [layoutRel]);
        for (let i = 0; i < it.slides.length; i++) {
          const s = it.slides[i];
          const abs = await addImage(s);
          let xml = setTitle(i === 0 ? tpl.first : tpl.plain, `Slide ${s.name}`).replace(/r:embed="[^"]+"/, 'r:embed="rId2"');
          if (i === 0 && ccli) xml = xml.replace("<a:t>CCLI # </a:t>", `<a:t>CCLI # ${esc(ccli)}</a:t>`);
          addSlide(xml, [layoutRel, { id: "rId2", type: REL_IMAGE, target: relTo("ppt/slides/x.xml", abs) }]);
        }
      }
    }
    addSlide(setTitle(tpl.end, "End"), [layoutRel]);

    // Presentation slide list
    const ids = newSlides.map((p, i) => {
      const rid = "rIdS" + (i + 1);
      presRels.push({ id: rid, type: REL_SLIDE, target: p.replace(/^ppt\//, "") });
      return `<p:sldId id="${256 + i}" r:id="${rid}"/>`;
    }).join("");
    pres = pres.replace(/<p:sldIdLst>.*?<\/p:sldIdLst>/s, `<p:sldIdLst>${ids}</p:sldIdLst>`);
    // Register imported slide masters (renumbering their layout IDs) and the notes master
    let mi = 0;
    for (const mp of masterEntries) {
      const rid = "rIdM" + (++mi);
      presRels.push({ id: rid, type: R + "/slideMaster", target: mp.replace(/^ppt\//, "") });
      pres = pres.replace("</p:sldMasterIdLst>", `<p:sldMasterId id="${nextMasterId++}" r:id="${rid}"/></p:sldMasterIdLst>`);
      let mx = await read(mp);
      mx = mx.replace(/(<p:sldLayoutId id=")(\d+)(")/g, (_, a, __, c) => a + (nextMasterId++) + c);
      zip.file(mp, mx);
    }
    if (notesMaster && !/<p:notesMasterIdLst>/.test(pres)) {
      presRels.push({ id: "rIdNM1", type: R + "/notesMaster", target: notesMaster.replace(/^ppt\//, "") });
      pres = pres.replace("</p:sldMasterIdLst>", `</p:sldMasterIdLst><p:notesMasterIdLst><p:notesMasterId r:id="rIdNM1"/></p:notesMasterIdLst>`);
    }
    zip.file("ppt/presentation.xml", pres);
    zip.file(presRelsPath, relsXml(presRels));

    // Drop media and diagram parts nothing refers to any more
    const referenced = new Set();
    for (const f of Object.keys(zip.files).filter((f) => f.endsWith(".rels"))) {
      const owner = f.replace(/_rels\/([^/]+)\.rels$/, "$1");
      for (const r of parseRels(await read(f))) if (r.mode !== "External") referenced.add(resolve(owner === "" ? "x" : owner, r.target));
    }
    for (const f of Object.keys(zip.files)) {
      if (/^ppt\/(media|diagrams)\/[^/]+$/.test(f) && !referenced.has(f)) {
        zip.remove(f);
        ct = ct.replace(`<Override PartName="/${f}" ContentType="${overrides["/" + f]}"/>`, "");
      }
    }
    zip.file("[Content_Types].xml", ct);
    zip.remove("docProps/app.xml");
    ct = ct.replace(/<Override PartName="\/docProps\/app.xml"[^>]*\/>/, "");
    zip.file("[Content_Types].xml", ct);
    const rootRels = parseRels(await read("_rels/.rels")).filter((r) => !/app\.xml$/.test(r.target));
    zip.file("_rels/.rels", relsXml(rootRels));

    // Write a clean package: files only (no folder entries, which stop Office opening or co-editing the file
    // from cloud storage such as Google Drive) and [Content_Types].xml first, as Office itself writes it.
    const clean = new JSZip();
    const names = Object.keys(zip.files).filter((f) => !zip.files[f].dir);
    names.sort((a, b) => (a === "[Content_Types].xml" ? -1 : b === "[Content_Types].xml" ? 1 : 0));
    for (const f of names) clean.file(f, await zip.file(f).async("uint8array"), { createFolders: false });
    const out = await clean.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 1 } });
    out.imports = importNotes;
    out.skipped = skipped;
    return out;
  }

  const api = { buildDeck, TRANSITION_INDEX, isLastOfMonth, chorusFontSize };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.ServiceDeck = api;
})(typeof window !== "undefined" ? window : globalThis);
