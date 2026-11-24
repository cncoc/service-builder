# CNCOC Service Builder

The web page the song leader uses to put together the Sunday and Wednesday service PowerPoint for Central North Church of Christ.

This repository holds only the page itself (`index.html`, the deck generator `gen.js`, the church's transition slides in `deck-data.js`, the hymn viewer and presenter `viewer.js` / `viewer.css`, and the MIDI player `midi-player.js`). The hymn pictures, chorus words and hymn catalogue are private: the page fetches them from the church's Google Apps Script after it is opened with the access key from the members' page of the church website.

Published with GitHub Pages: Settings → Pages → Deploy from a branch → `main`, `/ (root)`.

## Hymn viewer and presenting

Tap a hymn's title to open it: **Slides**, **Page** (the full hymnal page), **Words** and **Play** (the hymn's MIDI, with
tempo, voice parts and starting pitch). **▶ Present** shows a hymn or chorus full screen straight away, for songs sung on
the spot. Keys: → / Space / PageDown next, ← / PageUp back, 1–9 jump to a stanza, B black screen, F full screen, Esc stop.

Words, music and pages come from the church Apps Script's `media` action (`apps-script/Media.gs`), which reads
`hymn-media.json` and the `Hymn Pages` folder on the church Drive. Slides and presenting work without it.
