# Transcript Editor

Standalone Next.js app hosting a ProseMirror transcript editor for Steora Nexus Typist Studio. It has Word-like pagination, track changes (insertions and deletions per author), and DOCX import/export with real Word revisions. It runs on its own and is meant to be embedded in other apps (such as `av-tracking-project`) through an iframe.

The editor never loads or saves documents itself. The host page owns the document: it sends plain text or a DOCX in and reads plain text, JSON or a DOCX back over `postMessage`. That keeps this app free of auth, S3 and database access.

## Running

```bash
pnpm install
pnpm dev        # http://localhost:3001
```

| Command | Description |
|---------|-------------|
| `pnpm dev` | Dev server on port 3001 (av-tracking uses 3000) |
| `pnpm build` / `pnpm start` | Production build / serve on port 3001 |
| `pnpm test` | Vitest (track changes, page breaks, DOCX round trip, serializers, protocol) |
| `pnpm typecheck` | `next typegen` + `tsc --noEmit` |
| `pnpm lint` | ESLint |

## Routes

| Route | Purpose |
|-------|---------|
| `/` | Demo: the editor with a control panel (read-only, track changes, authors, changes list, DOCX open/save, handle methods, event log) |
| `/embed` | Bare studio editor for iframes. Talks to its parent only via the message protocol below |
| `/harness` | Fake host page that frames `/embed` and drives it, to test the embedding contract here |

## Layout

```
app/                       routes (/, /embed, /harness)
components/editor/         TranscriptEditor (React wrapper), studio chrome, page sheets, ChangesPanel, CSS
components/demo/           demo page UI
components/embed/          EmbedBridge (postMessage <-> editor handle)
components/harness/        HarnessHost (example host implementation)
lib/pm/schema.ts           document schema: paragraph, heading, lists, tables, image, page_break, hard_break; marks
lib/pm/controller.ts       EditorController: owns the EditorView, tracked dispatch, commands
lib/pm/plugins/            keymaps, input rules, history, placeholder, paste, picture/link views
lib/pm/tables.ts           table commands and plugins (prosemirror-tables)
lib/pm/objects.ts          hyperlink and picture commands
lib/pm/find.ts             find and replace plugin and commands
lib/pm/stats.ts, outline.ts  word counts, headings outline
lib/pm/comments.ts         comment threads (doc attribute) and commands
lib/pm/track-changes/      tracking plugin, transaction rewriting, accept/reject, change list
lib/pm/pagination/         page setup, pure page-break algorithm, measuring plugin
lib/docx/                  importDocx / exportDocx (Word revisions, page setup, page numbers)
lib/serializers/           plain text <-> ProseMirror JSON
lib/bridge/                message protocol types and allowed-origin parsing
```

## Editor architecture

The editor is plain ProseMirror (no TipTap). `TranscriptEditor` creates one `EditorController` in a ref callback; React only renders the chrome around it and re-renders on transactions (throttled to one frame).

### Paragraph formatting

- Toolbar and Format menu: alignment, line spacing (1.0–3.0, plus add/remove space before and after), bullets and numbering, increase/decrease indent, hanging indent, and a Word-style Paragraph dialog (alignment, left/right indent, first-line/hanging, spacing before/after, line spacing incl. "at least"/"exactly"/"multiple", "don't add space between paragraphs of the same style", widow/orphan control, keep with next, keep lines together, page break before, suppress line numbers).
- Word shortcuts: `Ctrl+L/E/R/J` align, `Ctrl+M` / `Ctrl+Shift+M` indent, `Ctrl+1/5/2` line spacing, `Ctrl+0` space before, `Ctrl+Shift+L` bullets. Some may be taken by the browser (e.g. `Ctrl+Shift+W` / `Ctrl+Tab` outside the iframe).
- Commands live in `lib/pm/paragraph.ts` and edit each paragraph's `format` attribute, so pagination and DOCX export pick the changes up. With tracking on, a format change becomes a paragraph revision (see below).
- Contextual spacing is a decoration (`lib/pm/plugins/contextualSpacing.ts`) that zeroes the space before/after a paragraph next to one with the same `styleId`.

### Styles and lists

- Styles (`lib/pm/styles.ts`): a document's styles are its imported `styles.xml` definitions (`layout.styles`), styles edited in the editor (the doc's `styles` attribute) and Word's built-in quick styles it lacks (Normal, No Spacing, Heading 1–6, Title, Subtitle, Quote, List Paragraph, Strong, Emphasis); built-ins inherit a Word document's Normal style. The toolbar's styles gallery applies one: a paragraph style sets the paragraph's `format` (with `styleId`), turns it into a heading for Heading 1–6 (a paragraph otherwise), and swaps the old style's character formatting (e.g. bold) for the new style's; a character style applies its formatting to the selection. "Update <style> to match selection" redefines the style from the paragraph and restyles every paragraph using it; "Clear all" applies Normal and clears character formatting.
- With tracking on, a style change is a paragraph revision ("Style: Heading 2, …"); rejecting it restores the old format and block type. Enter at the end of a paragraph whose style names a following style (headings → Normal) starts a paragraph in that style.
- Lists: `ordered_list` has `order` (start) and `numFormat` (decimal, lower/upper letter, lower/upper roman, legal `1.1.1.`; `null` follows Word's default 1. a. i. by depth); `bullet_list` has `bullet` (disc, circle, square, dash, arrow, check, diamond). The bullet and numbering split buttons offer the libraries, three multilevel lists, Restart at 1 (splits the list at the selected item), Continue numbering (from the previous list at the same level) and Set numbering value. Tab at the start of a list item demotes it.
- DOCX: list levels are read from `numbering.xml` (format, bullet character including Symbol/Wingdings, start, `isLgl`/"%1.%2." legal numbering) and numbering continues across lists sharing an abstract definition until a `startOverride`, as in Word. Export writes one abstract numbering per distinct set of levels and one `w:num` per list (with its start). styles.xml redefines Heading 1–6 and Title from the document (instead of docx's blue defaults) and adds every other paragraph style in use; paragraphs reference their style with `w:pStyle` and keep their formatting as direct formatting.

### Character formatting

- Marks besides `strong`/`em`/`textStyle` (font, size): `underline` (single, double, thick, dotted, dashed, wavy, words), `strike` (single/double; not a deletion), `textColor`, `highlight` (Word highlight colours or any run shading), `caps` (all/small) and `vertAlign` (super/subscript).
- Text colours are drawn for the dark page: dark colours have their lightness inverted (as Word's dark mode does), and printing uses the real colour. Super/subscript is drawn at 2/3 size without growing the line, like Word.
- Toolbar (Word's Home ▸ Font group): font family, font size (type any size or pick from Word's list), grow/shrink, clear formatting, bold, italic, underline (with a style menu), strikethrough, sub/superscript, text highlight and font colour (split buttons that remember the last colour). Format ▸ Font… (`Ctrl+D`) has the same options plus double strikethrough and small/all caps, and applies only the fields you changed, as one transaction.
- Word shortcuts: `Ctrl+B/I/U`, `Ctrl+Shift+D` double underline, `Ctrl+Shift+W` words-only underline, `Ctrl+=` subscript, `Ctrl+Shift+=` superscript, `Ctrl+Shift+A` all caps, `Ctrl+Shift+K` small caps, `Ctrl+Shift+>` / `<` next size in Word's list, `Ctrl+]` / `[` one point, `Ctrl+Space` clear formatting. `Shift+Tab` / `Ctrl+Tab` outdent/indent a list item.
- Commands live in `lib/pm/character.ts`. They work on the selection or, with an empty selection, on the stored marks for the next typed text. A font or size equal to the paragraph's is stored as "no override". With tracking on, every change is a formatting revision with Word's wording.
- Paste (`lib/pm/plugins/paste.ts`): Word clipboard HTML is cleaned (conditional comments, `<style>`, `<o:p>`, `mso-*` styles). Edit menu paste options mirror Word's: keep source formatting, merge formatting (keeps bold/italic/underline etc., drops fonts, colours, highlight and paragraph formatting) and keep text only. "Automatic" (the default) keeps formatting from this editor and Word and merges content from web pages, whose theme colours would otherwise be invisible in print. Pasted tracked deletions are dropped and pasted text never carries revision marks (with tracking on it becomes a new insertion). Alignment and left/first-line indents in pasted HTML become paragraph formatting.

### Find, navigation and view

- Find and Replace (`lib/pm/find.ts`, `Ctrl+F` / `Ctrl+H`, Edit menu, toolbar): a floating panel that leaves the document editable. Matches are highlighted as decorations (the current one in orange) and recomputed after every edit; options are Match case and Whole words only. A selection of up to one line is the initial search text. Enter / Shift+Enter (or the arrows) go to the next/previous match, wrapping around. Replace replaces the selected match and selects the next one (the first press only selects, as in Word); Replace All replaces every match as one undo step. Matches stay within a paragraph and skip tracked deletions; with tracking on, replacements are tracked like typing.
- Navigation pane (View ▸ Navigation pane): Headings lists the document's headings (indented by level, the one holding the cursor highlighted) and moves the cursor to the clicked heading, scrolling it to the top. Pages shows a thumbnail per page sketched from the paginated line positions; clicking one scrolls to that page.
- Zoom: 50–200% from the status bar slider, the View menu or Ctrl+mouse wheel; kept for the browser session (`sessionStorage`). The zoom is a CSS scale on the laid-out pages, so pagination and line breaks don't change with it.
- Status bar: "Page X of Y" (X is the page holding the cursor), word and character counts (of the selection when there is one; tracked deletions don't count), a Word Count dialog, the Track Changes state (click to toggle) and the zoom controls. Counts come from `lib/pm/stats.ts`.
- Ruler (View ▸ Ruler): sticky above the page and aligned with it at any zoom. It shows the margins, the cursor paragraph's indents and its tab stops. Drag the first-line (top triangle), hanging (bottom triangle, keeps the first line in place) and left indent (square, moves both) markers or the right indent; positions snap to 1/16" (hold Alt for free placement). Click the ruler to add a tab of the kind shown in the tab selector at its left end (click it to cycle left, center, right, decimal), drag a tab to move it, or drag it off the ruler to remove it. Changes apply to every selected paragraph.

### Track changes

- Three marks, `insertion`, `deletion` and `formatChange`, each with `id`, `authorId`, `authorName` and `date`. A `formatChange` also stores the old character formatting (`old`, a `RunFormat`) and Word's wording (`description`, e.g. "Font: Italic, Font color: Red").
- Paragraph revisions live on the paragraph's `paraChange` attribute (old `format` plus Word's wording, e.g. "Centered, Indent: Left: 0.5\""). They are listed as `type: "paragraph"` changes, get balloons and change bars, and rejecting one restores the old format.
- Displayed like Word's All Markup: insertions underlined and deletions struck through in the author's colour (red, blue, green… in order of first appearance), a change bar in the left margin beside every changed line, and formatting changes as balloons in a markup area to the right of the page. Simple Markup hides deletions, balloons and colours and shows red change bars.
- When tracking is on, every user transaction goes through `trackTransaction` before it is applied. Each replace step is rebuilt: deleted text is kept and marked `deletion`, inserted text is marked `insertion`. Deleting your own insertion removes it for real. Splitting or joining paragraphs is not tracked.
- Consecutive edits by the same author join the adjacent change, so a typing run is one change.
- Undo/redo, accept/reject, `setPlainText` and anything tagged with the `SKIP_TRACKING` meta are never rewritten.
- Typing through an IME is applied as-is and marked as an insertion when composition ends.
- Character formatting changes (every mark in `FORMAT_MARK_NAMES`) become formatting revisions that keep the original formatting, so toggling back removes the revision and rejecting restores it. Formatting your own insertion is not a separate revision.
- `getPlainText({ mode: "accepted" })` (the default) returns the text with all changes accepted; `"original"` returns it with all changes rejected.

### Review and comments

- Review menu: Track changes (`Ctrl+Shift+E`), Lock tracking (keeps tracking on; the Track button, status bar toggle and shortcut can't turn it off until unlocked), Previous/Next change, Accept/Reject and move to next (the changes in the selection, or the one at the cursor), Accept/Reject all, New comment, Previous/Next comment, Delete all comments, Simple markup, and Word's Show Markup filters: Insertions and deletions, Formatting, Comments.
- Turning a filter off removes its balloons and change bars (`MarkupFilter`, passed to pagination through `EditorController.setMarkupFilter`). Hiding insertions and deletions also shows the text as if accepted (deletions hidden, insertions plain), as Word does; hiding comments removes their highlight. Simple markup keeps comment balloons.
- Balloons: hovering a change balloon shows ✓ Accept / ✕ Reject; a comment balloon shows the comment, its replies and Reply / Resolve (or Reopen) / Delete. Resolved comments are dimmed. Right-click a balloon, a revision or commented text for the same actions in a context menu. Clicking a balloon selects its text.
- Comments (`lib/pm/comments.ts`): a `comment` mark (`id`, `authorId`, drawn in the author's colour) anchors each thread; overlapping comments are allowed. Threads (`text`, author, date, `resolved`, `replies`) live in the doc's `comments` attribute and change through `setDocAttribute`, so they undo with the document. New comment (`Ctrl+Alt+M`, toolbar, Review menu) anchors to the selection, or to the word at the cursor. Comments are never tracked as revisions.

### Tables, hyperlinks, and pictures

- Tables (`lib/pm/tables.ts`, `prosemirror-tables`): Insert ▸ Table or the Table menu inserts a grid sized to the text column. Tab moves between cells and adds a row after the last cell. Column borders drag to resize; row breaks across pages use a spacer row inside the table. Table menu: insert/delete rows and columns, merge/split cells, view gridlines. Right-click in a cell for the same layout commands.
- Hyperlinks (`lib/pm/objects.ts`): `link` mark; Insert ▸ Link or `Ctrl+K`, right-click on a link (edit, open, copy, remove). Ctrl+Click follows external links in a new tab. DOCX round-trips `w:hyperlink`, field hyperlinks, and internal anchors.
- Pictures: inline `image` node; Insert ▸ Picture, paste, or drop image files (PNG/JPEG/GIF/BMP). Selected images show a corner resize handle. DOCX imports `w:drawing` / VML and exports embedded media.
- Footnotes and endnotes are **not** implemented (see Out of scope).

### Layout, sections, headers and footers

- **Page Setup** (File or Layout menu): paper size presets (Letter, A4, Legal), portrait/landscape, margins. Updates screen pagination and, for Word documents, `layout.page` so DOCX export stays aligned.
- **Section breaks** (Insert menu): next page, continuous, even page, odd page. Each break stores that section's `SectionLayout` and how the following section starts.
- **Headers and footers**: double-click the header or footer band on a page sheet (Word layout only) to edit plain text; use `{PAGE}`, `{NUMPAGES}`, or `{SECTIONPAGES}` for fields. Changes are undoable and relayout the page frame.
- **Columns** (Layout ▸ Columns): sets `section.columns` on the current section for export. The editor still draws a single column on screen.

### Transcript blocks

- `qa_line` (Insert ▸ Q line / A line), `speaker_line` (speaker label + text), and inline `timestamp` nodes for deposition-style content.
- Plain-text export: `Q. …` / `A. …`, `THE COURT: …`, and timestamp text inline. Blank lines still separate paragraphs.

### Pagination

- Page geometry comes from a `PageSetup` (px at 96 dpi). The default is US Letter with 1 inch margins. An imported DOCX carries its own page setup in the document's `layout` attribute, unless the `pageSetup` prop is set.
- After each change, the pagination plugin measures every rendered text line and runs `computePageBreaks`, a pure function. A line that would cross the bottom of the page's text area moves to the next page, so paragraphs split between lines like in Word. Widow/orphan control, keep-with-next and keep-lines-together move the break earlier, as in Word.
- With multiple line spacing (e.g. 1.5 lines), Word puts the extra space under the text, and only the text has to fit on the page. Lines carry that extra leading as `slack`, which may hang past the bottom of the text area.
- Word documents render with `white-space: pre-wrap` (trailing spaces hang past the margin instead of wrapping) and `font-kerning: none` (Word only kerns runs with `w:kern`). A `<wbr>` decoration after a hyphen before a digit lets lines break in "25-127116" as Word does.
- Each page's text area is grown to fit its header and footer, as Word does: the body starts below `headerDistance + header height` and ends above `footerDistance + footer height` when those exceed the margins. A title page (`titlePage`) uses the first-page header and footer; with `evenAndOddHeaders`, even pages use the even header and footer.
- Sections: a `section_break` node ends a section and stores that section's layout (`SectionLayout`: margins, headers/footers, title page, line numbers, page-number restart, columns) plus how the next section starts (next page, continuous, even page, odd page). Import reads `w:type` from the same inline `w:sectPr` as Word; **continuous** breaks do not force a new page. The last section's layout is the document's `layout` attribute. Every measured line carries its section index; `computePageBreaks` reports each page's section, whether it is the section's first page, and its displayed page number, and leaves a blank page when an odd/even section break needs one. A paragraph with "page break before" starts a new page.
- Breaks are drawn as block widget decorations whose height pushes the next line to the top of the next page's text area. Tab widths are node decorations, measured so each tab reaches its Word tab stop (custom stops, the hanging indent, then the default stop grid).
- Page sheets are drawn behind the text: the page's section's header and footer (with live `PAGE` / `NUMPAGES` fields), framed header/footer paragraphs (`w:framePr`) at their anchored position, and Word line numbers (`countBy`, restart per page, per section or continuous). Documents without a Word layout get a "Page N" footer.
- `Ctrl+Enter` (Insert → Page break) inserts a hard page break.
- Known limits: all sections share the last section's page width, height and left/right margins (top/bottom margins and headers vary per section); columns are stored and exported but drawn as one column. Justified lines that need more than Word's 25% space shrink can still break one word differently on rare lines, which can move a page break by a line or two in long documents.
- **Print / Save as PDF**: File ▸ Print or Save as PDF sets a `printing` class on `<html>`, hides chrome and markup balloons, resets zoom to 1, and prints page sheets plus body (browser "Save as PDF" uses the same path). Custom paper sizes may not match every browser's `@page` box; use Save .docx for exact Word print when needed.

### DOCX

- `importDocx` resolves Word styles (document defaults, the paragraph style `basedOn` chain, character styles, direct formatting, theme fonts) into a `format` attribute per paragraph: alignment, indents (first-line and hanging), spacing before/after, line spacing (auto multiple, exact, at least), paragraph-mark font and size, tab stops, widow control, keep-with-next. Run fonts that differ from the paragraph become `textStyle` marks; `<w:tab>` becomes a `tab` node.
- It also reads every section (`w:sectPr`: page size and margins, header/footer distances, default/first/even headers and footers with Word's inheritance from the previous section, `titlePg`, `lnNumType`, `pgNumType start`, `cols`, break type), `evenAndOddHeaders` and `defaultTabStop`, header/footer text, tabs, frames and `PAGE` / `NUMPAGES` / `SECTIONPAGES` fields, plus headings, character formatting (bold, italic, underline, strikethrough, colour, highlight/shading, caps, super/subscript), contextual spacing, page break before, line and page breaks, lists, tables (with merged cells, `tblW`/`jc`/`tblInd`, per-cell `tcBorders`/`tcMar`, row `trHeight`/`cantSplit`, and cell shading), paragraph box borders (`w:pBdr`), body `w:framePr` paragraphs (positioned, zero layout height), hyperlinks and inline pictures (including legacy VML `w:pict`), `<w:ins>`/`<w:del>` (and moves), `<w:rPrChange>` and `<w:pPrChange>` revisions, and a summary of `styles.xml` (`layout.styles`).
- Batch QA: `node scripts/build-batch-fixtures.mjs`, `node scripts/batch-layout-scan.mjs <folder>`, `node scripts/layout-audit.mjs <file.docx>`, `npx tsx scripts/batch-layout-coverage.ts <folder>`. Golden fixtures live under `public/fixtures/batch/`.
- Word's "single" line height is rendered as 1.15 × the font size (Times New Roman and Arial). Paragraph spacing is CSS padding, so it adds up like Word's instead of collapsing.
- Word 2013+ documents (compatibility mode 15) let justified lines shrink their spaces by up to 25% to fit another word. The importer stores this as `layout.justifyShrink`; justified paragraphs get a matching negative `word-spacing` (through the `--pm-justify-shrink` variable on the editor), so the browser breaks lines where Word does and justification stretches the spaces back.
- `exportDocx` writes the same formatting, tabs, fonts, sections, headers/footers (including even pages and frames), title page and line numbers back. Plain-text documents (no Word layout) use Times New Roman 12 pt, double spacing and a centred `Page N` footer. Insertions become `<w:ins>`, deletions `<w:del>`/`<w:delText>`, formatting changes `<w:rPrChange>`, paragraph revisions `<w:pPrChange>`, with author and date, so Word shows them as revisions. If tracking is on, the file opens in Word with Track Changes on.
- Word stores only an author name, so imported changes use the name as `authorId`.
- Comments: `comments.xml` threads are anchored with `commentRangeStart`/`commentRangeEnd` (ranges may span paragraphs); `commentsExtended.xml` supplies replies (`paraIdParent`) and the Resolved state (`done`). Export writes the ranges, a `commentReference` after each, and replies as threaded comments sharing the parent's range. The docx library only writes `commentsExtended.xml` when some comment has a reply, so Resolved is lost on export for documents without replies.

## Plain text semantics

- Paragraph break: blank line (`\n\n`), aligned with `lib/studio-docx-merge.js` in av-tracking-project.
- A single `\n` inside a block becomes a line break (`hard_break`) and is exported as `\n`.
- Headings and list items are exported as separate paragraphs; page breaks are dropped.

## Message protocol

Every message is an object with a `type` prefixed `te:`. Types live in [`lib/bridge/protocol.ts`](lib/bridge/protocol.ts). DOCX payloads are `ArrayBuffer`s; pass them in the transfer list.

Host to editor:

| Message | Effect |
|---------|--------|
| `te:load { plainText, editable? }` | Replace the document (fresh editor, undo history cleared) |
| `te:load-docx { buffer, editable? }` | Replace the document with a DOCX, keeping its revisions and page setup |
| `te:get-plain { requestId, mode? }` | Reply with `te:plain`; `mode` is `"accepted"` (default) or `"original"` |
| `te:get-json { requestId }` | Reply with `te:json` (ProseMirror JSON, including change marks) |
| `te:get-docx { requestId }` | Reply with `te:docx` |
| `te:set-plain { plainText }` | Replace content in place (undoable, not tracked) |
| `te:set-editable { editable }` | Toggle read-only |
| `te:set-author { id, name }` | Who new changes are attributed to |
| `te:set-track-changes { enabled }` | Turn tracking on or off |
| `te:focus` | Focus the editor |
| `te:load-template { styles, requestId? }` | Merge `StyleDefinition[]` into the document styles (gallery / CRM templates); optional `requestId` gets `te:loaded` when done |

Editor to host:

| Message | When |
|---------|------|
| `te:ready` | `/embed` mounted and listening; host should send `te:load` or `te:load-docx` |
| `te:loaded` | A loaded document is rendered (also acknowledges `te:load-template` when `requestId` was sent) |
| `te:plain { requestId, text }` | Reply to `te:get-plain` |
| `te:json { requestId, json }` | Reply to `te:get-json` |
| `te:docx { requestId, buffer }` | Reply to `te:get-docx` |
| `te:change { pageCount, changeCount }` | User edited the document (debounced 400 ms) |
| `te:layout { pageCount }` | Page count changed |
| `te:action { name: "playPause" \| "print" }` | Studio toolbar/menu actions the host should handle |
| `te:error { message, requestId? }` | A request failed (for example a request before load, or an unreadable DOCX) |

### Allowed origins

`NEXT_PUBLIC_ALLOWED_PARENT_ORIGINS` (comma-separated, see `.env.example`) controls who may embed the editor. Defaults to `http://localhost:3000,http://localhost:3001`.

- `/embed` is served with `Content-Security-Policy: frame-ancestors 'self' <origins>`.
- The bridge ignores messages not sent by its parent window from an allowed origin, and only posts back to that origin.

The value is read at build time, so rebuild after changing it.

## Future av-tracking integration

Not implemented yet. The plan is a `TranscriptIframePane` in av-tracking-project modelled on `components/studio/edentext-editor-pane.jsx`:

1. Render `<iframe src="{EDITOR_URL}/embed">`, with av-tracking's origin added to `NEXT_PUBLIC_ALLOWED_PARENT_ORIGINS` here.
2. On `te:ready`, post `te:set-author` with the signed-in user and `te:set-track-changes` for review stages, then fetch the stored DOCX and post `te:load-docx` (or `te:load` with plain text).
3. Saving: post `te:get-docx` with a fresh `requestId`, resolve on the matching `te:docx` with a timeout (as `HarnessHost` does), and upload the buffer. `te:get-plain` remains available for plain-text exports.
4. `te:action playPause` calls `postStudioSync(subtaskId, { type: "toggle" })`; `te:change` can drive autosave, word counts and "N changes pending" badges.

Always check `event.source === iframe.contentWindow` and `event.origin === EDITOR_ORIGIN` before trusting a message.

## Out of scope

- VBA / macros, mail merge
- Fields beyond PAGE, NUMPAGES, SECTIONPAGES (and footnotes / endnotes)
- SmartArt, charts, OMML math, OLE
- Real-time co-authoring
- Full Compare / legal blackline
- Binary-identical DOCX round-trip
- Per-section page width and left/right margins (top/bottom and headers vary per section; screen width follows the document's last section page size)

## Test plan (embed hosts)

- Load plain text and DOCX; confirm `te:change` page count matches the status bar.
- `te:get-plain` accepted vs original with track changes on.
- `te:load-template` then open the styles gallery in the iframe.
- Print preview: chrome hidden, page count unchanged from paginated view.
- Section break + Page Setup round-trip via Save/Open .docx on the demo page.

## Deferred

- Document Compare.
