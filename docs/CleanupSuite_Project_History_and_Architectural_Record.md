# CleanupSuite Project History and Architectural Record

**Project:** CleanupSuite For Word
**Publisher:** MasseysLab
**First repository commit:** June 6, 2026
**Record established:** July 18, 2026
**Stable installer release:** 0.9.3 (retained internal/tag identifier:
`0.9.3-beta-vba-final`)
**Final standalone VBA-only release:** `0.9.3-beta-vba-final`
**Current development branch:** `codex/0.9.5-beta`
**Next public milestone:** `0.9.5` Official Beta
**Approved architectural direction:** C# analysis engine with VBA as the Word mediator

## Purpose of This Record

CleanupSuite began as a broad VBA cleanup toolkit and, in a short period, became a packaged, preview-first Word product with a global template, installer, documentation, tests, and a growing safety model. By July 18, 2026, the project had reached another consequential design point: whether to continue expanding entirely in VBA or retain VBA as the Word-facing mediator while moving analysis into a separately programmed engine.

This document preserves that history and the reasoning behind the emerging architecture. It is intentionally different from release notes and `WHERE_WE_LEFT_OFF.md`:

- release notes describe what shipped in a particular version;
- `WHERE_WE_LEFT_OFF.md` describes the exact current working state;
- this file explains how the product evolved, why major decisions were made, and which proposed decisions may define its future.

The record uses four status labels:

- **Published:** released and represented by a Git tag or published artifact.
- **Local:** implemented in the current working tree but not published.
- **Approved, paused:** accepted as the intended correction, but deliberately waiting for Chris to say `go`.
- **Authorized:** Chris has opened the implementation gate; publication remains separately controlled.
- **Proposed:** investigated and recommended, but not yet accepted as an implementation decision.

Difficulty and desirability percentages are practical decision-support estimates. They are not scientific survey results. Difficulty represents combined engineering effort, regression risk, Word-integration risk, testing, and installer impact. Professional desirability estimates how strongly careful professional users would likely value the behavior.

## The Original Product Problem

The earliest project planning recognized that CleanupSuite did not primarily lack tools. By June 10, 2026, it already had approximately nineteen Word cleanup tools. Its central problem was product flow:

- users needed a coherent way to find the right tool;
- Preview needed to become the normal path rather than an incidental checkbox;
- destructive actions needed to be deliberate;
- the source, generated VBA, test documents, macro documents, template, and installer needed one authoritative workflow;
- the product needed to work on ordinary documents without modifying its own macro host;
- the eventual destination was a global Word template rather than a collection of pasted macros.

The project priority list from June 10 stated the direction clearly: protect the work, redesign the launcher, establish the Preview/Apply model, create a shared form language, modernize complex tools, modernize simple tools, and only then move toward a global `.dotm` release. See [CleanupSuite Project Priority List](CleanupSuite_Project_Priority_List_2026-06-10.md).

## Chronological History

### June 6-9, 2026: Source Pipeline and Installation Foundation

The repository began on June 6 with commit `948fe84`, described as “Initial commit — pipeline working, all validators clean.” The initial foundation already treated the project as source-first rather than as an opaque Word file.

Early work concentrated on:

- assembling readable source into distributable VBA;
- validating the assembled result;
- repairing array-continuation syntax;
- making the installer report failures without triggering Word error 75;
- assigning captions to the generated controls;
- correcting Preview control text and installation behavior.

This established the source pipeline that remains fundamental:

```text
src/ → assemble.py → VBA_Cleanup_tool.txt → synchronized Word artifacts
```

The governing rule became: edit structured files under `src/`; do not hand-edit the generated `VBA_Cleanup_tool.txt` as the normal development surface.

### June 10, 2026: The Product Becomes a Planned Suite

June 10 was the first major product-definition day. The project gained:

- an approved form-redesign direction;
- refined launcher and Preview-flow specifications;
- revised categories for document tools;
- a documented sub-tool placement review;
- a formal project priority list;
- a pre-1.0 version ladder;
- the `0.5.0` working baseline.

The launcher was redesigned for `0.6.0`, synchronized into the working documents, and compacted into two columns. This was the point where CleanupSuite began to behave like a unified Word utility rather than a flat collection of macros.

### June 11-12, 2026: Preview Becomes the Product Spine

The Preview Actions panel was introduced, made modal for stability, refined, and finalized. The new workflow separated configuration from review:

1. configure a tool;
2. run Preview;
3. review the document with a small Preview Actions panel;
4. Apply or return to the options.

The `0.6.4` code-freeze candidate finalized the compact Preview Actions bar, preserved panel position, added return-to-main behavior, and synchronized the generated artifacts. `0.6.5` then became the documentation-complete release for the same programming baseline.

This established a durable product principle: **CleanupSuite accelerates careful human review; it does not replace it.**

### June 12-15, 2026: The 0.7 Line and Release Discipline

The `0.7.x` line began after the `0.6.5` documentation checkpoint. It corrected launcher settings overlap and refined Preview-off wording. The line then progressed through:

- `0.7.4` programming-side code freeze;
- `0.7.4.5`, a small bridge for release-blocking Preview Actions and compile fixes;
- `0.7.5`, the documentation-complete release.

By this stage the project was maintaining version-matched user manuals, programmer guides, tool references, and release notes. Documentation-only milestones were deliberately separated from feature milestones so behavior would not drift during a documentation release.

### June 15-25, 2026: Guided Forms and Alpha Preparation

The `0.8.0` milestone moved the suite into the guided-form redesign. The project standardized:

- body-owned visible titles;
- consistent instructions and action placement;
- scope language;
- dynamic two-column guided information;
- contextual risk explanations;
- compact forms within the supported display envelope;
- common layout helpers instead of one-off positioning.

The `0.8.5` Alpha-preparation work refined the two-column information box, simplified option captions, suppressed the Preview form title, improved long-caption sizing, and preserved a compact Office-adjacent visual style.

### July 15, 2026: 0.9.0 Alpha — First Packaged Public Alpha

Commit `4a1f074` released CleanupSuite `0.9.0-alpha`. This was the first packaged public Alpha and included:

- twenty focused cleanup tools;
- the global `CleanupSuite.dotm` template;
- a per-user Windows installer;
- a registered uninstall entry;
- installer backup retention;
- periodic update checks with an opt-out and manual check;
- Reading View Preview with the Preview Actions panel;
- MetaDataSuite auditing and safe editing;
- the upgraded Capitalization repair baseline;
- source assembly, regression tests, Word smoke tests, and full user documentation.

The installer deliberately avoided administrator rights, `Normal.dotm`, machine-wide installation, and changes to Word macro-security policy.

### July 15, 2026: 0.9.1 Alpha — OneDrive and MetaDataSuite

The first Alpha hotfix addressed a real synchronized OneDrive document. Word represented the document as a `d.docs.live.net` URL and represented the extra period before the extension with a caret. MetaDataSuite could not reliably resolve the physical package.

The fix:

- mapped the Word URL to the existing synchronized local path;
- repaired the caret representation only when the physical file existed;
- unified current-document detection across safe editing and grouped editing;
- prevented duplicate document opens.

The affected test document was `Arguement for and against God..docx`. This document later became the primary large-document diagnostic for Preview performance, highlighting, tables, duplicates, and structural safety.

### July 18, 2026: 0.9.2 Alpha — Punctuation Performance and Preview Stability

The final planned Alpha hotfix focused on Punctuation Normalizer behavior in the large “Argument for God” document.

The original symptom was that Punctuation could blink heavily or appear to run indefinitely on a large document while some simpler punctuation families completed. Capitalization did not exhibit the same repeated blinking.

The resulting release:

- moved punctuation Preview to Word's display-only `HitHighlight` path;
- grouped Apply into wildcard `ReplaceAll` passes;
- temporarily disabled and restored Word's automatic smart-quote replacement during Apply;
- removed unnecessary Preview repainting and screen-updating toggles;
- kept Preview in Reading View;
- restored the original view when Preview ended;
- changed paragraph-spacing Preview from full-paragraph shading to a boundary marker;
- made Setup bypass stale cached release metadata.

This became an important performance rule: exact deterministic character conversion should use Word's fast bulk path, while context-sensitive analysis may use a richer candidate path.

## The Local Post-0.9.2 Hotfix and Beta Preparation

**Status: Local, not published.**

The `codex/0.9.5-beta` branch advanced from published `0.9.2-alpha` through three approved Beta tasks. Chris explicitly authorized the resulting final standalone `.dotm` checkpoint to be committed, tagged, pushed, and published before hybrid-dependent tool integration begins.

### Minimal Preview Markers

The early spacing Preview attempt shaded whole paragraphs bright green. On the large article this overwhelmed the page and made the actual location of a spacing change difficult to understand. Chris rejected the full-area result and chose a minimal marker as close as possible to the changed boundary.

The resulting local rule is:

- highlight one visible character nearest the actual change;
- remain within the target range whenever possible;
- do not cross a table-cell marker into an adjacent cell;
- use the nearest safe visible character for an invisible boundary;
- fall back to temporary paragraph shading only when no safe visible character exists.

This rule was extended locally to structural Preview results in Break Normalizer, Final Review, Header/Footer Standardizer, List Normalizer, Object Remover, Paragraph Fixer, Soft Return Converter, Style Cleanup, Table Cleaner, and Duplicate Paragraph Remover.

### Duplicate Paragraph Remover Redesign

**Status: Local.**

The tool evolved from a detector with redundant Highlight/Remove choices into a removal-oriented, Preview-first tool:

- title: **Duplicate Paragraph Remover**;
- Exact match: left column, row 1;
- Normalized match: left column, row 2;
- Fuzzy match: right column, row 1;
- Loose, Medium, and Strict appear vertically beneath Fuzzy;
- a divider separates matching choices from **Include empty paragraphs**;
- Exact is the default;
- Medium is the default Fuzzy threshold;
- Preview is the review step and Apply removes later duplicates while keeping the first survivor.

The large article produced 96 exact duplicate paragraphs whether Include empty paragraphs was enabled or disabled. Investigation showed that this was correct, not a failure:

- 980 total paragraphs;
- 140 visually empty paragraph objects;
- 139 table-contained empty objects;
- 62 required empty-cell markers;
- 77 table row-end markers, one for each row across five tables;
- one required final-document paragraph.

There were no additional safely removable empty duplicate paragraphs. This established a central structural lesson: a visually blank Word paragraph is not necessarily disposable content.

## Approved Corrective Safety Package

**Status: Authorized for implementation on July 18, 2026; this is the first active Beta task.**

The following corrective behavior has been accepted as the intended direction:

1. Create a shared blank classifier that distinguishes removable body blanks, removable extra blanks inside populated cells, protected empty-cell markers, protected row markers, protected final-document markers, eligible wholly empty rows, and ambiguous/unsupported structures.
2. Spacing Fixer must never delete table rows.
3. Paragraph Fixer, Duplicate Remover, Document Trim, and other blank-aware tools must use the same classifier.
4. Preview and Apply must use the same candidate scan and revalidate those candidates before editing.
5. Table Cleaner is the sole owner of row/column deletion and table-to-text behavior.
6. Structural Table Cleaner options start OFF after initialization and Reset and are excluded from All, Recommended, and Select All.
7. Empty-row detection must consider pictures, fields, controls, bookmarks, hidden content, nested tables, nonbreaking content, merged cells, header rows, partial-scope rows, the only remaining row, and implicit whole-table deletion.
8. Deletion proceeds from the bottom upward after candidate revalidation.
9. Table-to-text Apply must enforce the same single-column eligibility that Preview reports.
10. Structural Apply requires a successful Preview, a count-specific confirmation with the safe answer as default, and one Word Undo record.
11. Break Normalizer's page-break and section-break structural choices start OFF after initialization and Reset.
12. Preview navigation uses the approved layout:

```text
[Change] [Page] ◀ Previous | Next ▶ [Page] [Change]
```

All four navigation buttons have the same size. Page buttons use a pale-blue treatment; Change buttons use a pale-gold treatment; disabled buttons use gray. Page moves exactly one page. Change jumps to the nearest page containing a Preview candidate.

Estimated combined difficulty: **58–70%**.
Estimated professional desirability: **97–99%**.

## Beta Product Contract

The existing [0.9.5 Beta Contract](0.9.5-beta-contract.md) defines Beta as more than general polish. Its required direction includes:

- full MetaDataSuite incorporation while keeping Final Review separate;
- the Capitalization follow-up upgrades;
- Page and Change Preview navigation;
- a review path for nonhighlighted findings;
- carefully designed inverse transformations, beginning with straight-to-curly quotes;
- one state-aware Setup executable supporting Install, Update, Repair, and Uninstall;
- enough consistency and validation for serious Beta testing.

The installer requirement is specifically state-aware:

- absent: offer Install;
- older version present: offer Update, Repair, and Uninstall;
- current version present: offer Repair/Reinstall and Uninstall.

Repair must verify hashes, safely replace the Startup template, preserve backup retention, and leave Word security and `Normal.dotm` untouched.

## Professional Suite-Wide Audit

**Status: Proposed.**

On July 18, 2026, the approved structural package prompted a wider question: should the same professional expectations be extended across all tools? The answer was yes.

The proposed common standard is:

- Preview must be truthful and non-destructive.
- Existing document formatting, including user highlights, must survive Preview.
- Low-confidence changes require review.
- Structural or content-removing choices start OFF.
- Preview and Apply use the same candidate list.
- Protected and skipped items are explained.
- Exact deterministic replacements retain the fast path.
- Apply reports changed, protected, skipped, failed, and remaining items.

### Existing Highlight Preservation

The audit identified a critical shared concern: several Preview paths apply real Word highlight formatting and later clear highlighting across a range or entire document. This can erase or overwrite highlights that existed before CleanupSuite began. That is especially important for documents such as the “Argument for God” article, where yellow and green highlighting carries document meaning.

The proposed correction is a shared Preview marker store that captures and restores the original highlight or shading at every marker location. Clear Preview, Reconfigure, Apply, Cancel, errors, switched windows, and closed documents must restore the exact original state. A whole-document assignment to `wdNoHighlight` must not be used as Preview cleanup.

Estimated combined difficulty: **45–60%**.
Estimated professional desirability: **99–100%**.

### Shared Candidate Manifest

Each finding should be represented by a durable record containing:

- tool and candidate type;
- story, range, and page;
- before and after values;
- confidence;
- highlightability;
- protected/skipped reason;
- fingerprint for Apply-time revalidation.

Preview summaries, nonhighlighted details, Page/Change navigation, Apply, and final reporting should all consume the same manifest.

Estimated combined difficulty: **65–80% in VBA alone**.
Estimated professional desirability: **97–99%**.

### Tool-Specific Recommendations

| Tool | Principal proposed professional correction | VBA-only difficulty | Professional desirability |
| --- | --- | ---: | ---: |
| Punctuation Normalizer | Explicit plain-text and typographic directions; context-aware straight-to-curly conversion; ambiguous cases review-only | 58–72% | 92–97% |
| Invisible Unicode Cleaner | Stop labeling all controls safe; protect intentional NBSP, nonbreaking hyphens, ZWJ/ZWNJ, and soft-hyphen semantics | 35–50% | 98–100% |
| Capitalization Fixer | Divide high-, medium-, and low-confidence repairs; batch only high-confidence findings | 50–68% | 93–97% |
| List Normalizer | Stop imposing universal hard-coded list indents and spacing; learn from a reference list/style | 58–75% | 96–99% |
| Font Normalizer | Operate on exact direct-formatting runs; protect character styles, symbols, equations, fields, links, and mixed formatting | 50–65% | 97–99% |
| Formatting Cleaner | Preview every property that reset will remove; paragraph reset OFF by default; preserve mixed-run emphasis and highlighting | 48–65% | 97–99% |
| Hyperlink Cleaner | Separate appearance and target; add Restore Hyperlink style; report field/shape/unsupported links | 28–42% | 87–94% |
| Soft Return Converter | Make paragraph-to-soft-return selection-first and structurally protected; add imported-text heuristics | 42–58% | 97–99% |
| Document Trim | Use the shared object-aware blank classifier and report removable/protected trailing blanks | 25–38% | 94–98% |
| Final Review | Preserve OFF defaults; report revisions/comments by type and author; expose tracking state | 35–50% | 94–98% |
| Style Cleanup | Protect style dependencies and types; replace fixed remapping with explicit source-to-target mapping | 65–82% | 96–99% |
| Footnote/Endnote Remover | Both removals OFF; show note excerpts; count-specific confirmation; disclose flattening in Keep Inline | 32–48% | 98–100% |
| Header/Footer Standardizer | Operate on existing unique stories; account for LinkToPrevious; use Header/Footer styles or a reference section | 48–65% | 96–99% |
| Object Remover | Keep all defaults OFF; count every hidden-text run; distinguish controls and objects; report failures/skips | 45–60% | 97–99% |
| MetaDataSuite | Consolidate metadata surfaces; require backups; exact before/after preview; atomic package replacement and validation | 55–72% | 97–99% |

The complete VBA-only professional package was estimated at:

Estimated combined difficulty: **82–92%**.
Estimated professional desirability: **98–100%**.

## Architectural Turning Point: VBA Mediator Plus External Engine

**Status: Approved architectural direction and authorized for incremental implementation on July 18, 2026.**

After seeing the size of the professional audit, Chris asked for a recalculation assuming the software's analysis could be programmed in another language while VBA remained the mediator. This reframed the problem.

After reviewing the recalculation and the implications for adding future tools, Chris accepted the recommended C#/VBA hybrid design. This is now a product architecture decision rather than a tentative idea. The approval does not authorize implementation, commit, push, publication, or a rewrite-before-Beta. It establishes the direction in which new architecture should be designed when implementation is authorized.

The recommended long-term architecture is a hybrid:

```text
Microsoft Word
      ↕
VBA mediator
UI • scope • Word ranges • Preview • navigation • Undo • Apply
      ↕ versioned request/response contract
C# analysis engine
classification • language rules • matching • candidate manifests • reports • Open XML
```

The governing boundary is: **move the thinking out of VBA, but keep Word manipulation in VBA.**

### Responsibilities That Stay in VBA

- launcher and Word UserForms;
- active-document and selection handling;
- story, table, field, object, and formatting extraction;
- Reading View and Preview lifecycle;
- minimal markers and restoration;
- Page and Change navigation;
- protected-document checks;
- candidate revalidation;
- Word edits and Undo records;
- view, scroll, selection, and screen-state restoration.

Estimated combined difficulty: **30–42%**.
Estimated professional desirability: **97–100%**.

### Responsibilities That Move to C#

- Unicode and punctuation classification;
- context-aware inverse punctuation;
- Capitalization confidence and language rules;
- exact, normalized, and fuzzy duplicate matching;
- manual-list and soft-return heuristics;
- shared candidate records and reports;
- style dependency graphs after Word facts are extracted;
- metadata and Open XML package inspection;
- performance-intensive scans;
- unit-testable rule engines.

Estimated combined difficulty: **52–67%**.
Estimated professional desirability: **97–100%**.

### Why a Separate C# Process Is Preferred

A separate C# executable is preferred over loading a managed COM DLL inside Word because it:

- avoids Office 32/64-bit in-process matching problems;
- isolates engine failures from Word;
- can be self-contained for installation;
- supports modern Unicode, regular expressions, collections, cancellation, and testing;
- works naturally with Open XML for closed-file metadata analysis;
- does not require the external engine to hold live Word Range objects.

The engine should return candidates, not directly edit the open document. VBA should restore Preview, revalidate the candidate fingerprint, and perform the edit.

Microsoft documents both the ability to expose modern .NET components to COM and the associated COM-host bitness and deployment constraints: <https://learn.microsoft.com/en-us/dotnet/core/native-interop/expose-components-to-com>. Microsoft documents Open XML as a strongly typed way to inspect and manipulate Office Open XML packages: <https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk>.

Recommended C# sidecar architecture:

Estimated combined difficulty: **48–62%**.
Estimated professional desirability: **96–99%**.

### Communication Contract

The first practical bridge could use restricted, short-lived UTF-8 job/result files. A later bridge could use named pipes for live progress, cancellation, and no persistent document-text payload.

Every request and response should contain a schema version. A candidate response should include fields equivalent to:

```text
ToolId
StoryId
Start
Length
Category
Before
After
Confidence
Highlightable
Protected
Reason
Fingerprint
```

Temporary files must be current-user-only, stored outside ordinary shared temporary locations, and deleted on success, error, startup recovery, repair, and uninstall.

Job/result bridge:

Estimated combined difficulty: **28–40%**.
Estimated professional desirability: **91–96%**.

Named-pipe follow-up:

Estimated combined difficulty: **48–63%**.
Estimated professional desirability: **84–92%**.

### Recalculated Tool Difficulty Under the Hybrid Model

| Tool | VBA-only estimate | Hybrid estimate after foundation |
| --- | ---: | ---: |
| Punctuation Normalizer | 58–72% | 36–50% |
| Unicode Cleaner | 35–50% | 20–32% |
| Capitalization Fixer | 50–68% | 31–45% |
| Duplicate Remover | 58–70% | 30–44% |
| List Normalizer | 58–75% | 46–61% |
| Font Normalizer | 50–65% | 43–57% |
| Formatting Cleaner | 48–65% | 42–56% |
| Hyperlink Cleaner | 28–42% | 24–36% |
| Soft Return Converter | 42–58% | 33–46% |
| Document Trim | 25–38% | 21–31% |
| Final Review | 35–50% | 31–43% |
| Style Cleanup | 65–82% | 49–65% |
| Footnote/Endnote Remover | 32–48% | 28–40% |
| Header/Footer Standardizer | 48–65% | 42–56% |
| Object Remover | 45–60% | 38–52% |
| MetaDataSuite | 55–72% | 29–44% |

The hybrid model helps text analysis, fuzzy matching, reporting, and metadata most. It helps Word formatting and structural tools less because the Word object model remains the difficult part.

### Approved Migration Direction

Three paths were considered:

1. **Rewrite everything before Beta:** clean eventual architecture but excessive simultaneous risk. Difficulty 89–96%; desirability 62–76%.
2. **Keep everything in VBA:** fastest for individual hotfixes but increasingly difficult to maintain. Complete-vision difficulty 82–92%; desirability 70–82%.
3. **Incremental hybrid migration:** build the contract around real tools, then migrate the parts that benefit. Complete-package difficulty 72–84%; desirability 98–100%.

The approved path is incremental hybrid migration. It must not become a big-bang rewrite before Beta. The first proof should use real CleanupSuite needs rather than an abstract framework:

1. Unicode Cleaner;
2. Duplicate Paragraph Remover;
3. Punctuation Normalizer;
4. candidate manifest and reporting;
5. Capitalization Fixer;
6. MetaDataSuite;
7. Soft Return and List analysis;
8. Style dependency analysis.

The first milestone costs more than one VBA hotfix, but later ambitious tools are expected to become roughly 20–40% easier and considerably more testable.

## What the Hybrid Architecture Means for Adding Tools

**Status: Accepted as the extensibility direction within the approved hybrid architecture. Exact framework details remain subject to implementation validation.**

Today, a new VBA tool tends to duplicate scanning, collections, Preview logic, highlight handling, summary counts, Apply behavior, errors, progress, and form layout. Under the hybrid architecture, an ordinary new tool would usually provide:

```text
tool definition + analyzer + tests
```

The shared system would provide Preview, navigation, reporting, Apply mediation, safety, and lifecycle behavior.

### Five Parts of a Future Tool

1. **Tool definition:** identity, category, risk, choices, defaults, scopes, and Preview requirement.
2. **C# analyzer:** finds candidates, assigns confidence, and supplies protected/skipped reasons.
3. **VBA Apply adapter:** performs a standard Word operation.
4. **Generic Preview presentation:** counts, markers, details, Page/Change navigation, and final results.
5. **Tests:** analyzer unit tests, contract tests, Preview/Apply parity, and Word integration.

### Standard Operation Vocabulary

The mediator should support a deliberately small operation vocabulary:

```text
ReplaceText
DeleteText
InsertText
DeleteParagraph
SetFontProperty
ResetDirectFormatting
SetParagraphProperty
ConvertList
DeleteObject
DeleteTableRow
ConvertTableToText
SetDocumentProperty
PackageMetadataChange
ReviewOnly
```

If a new analyzer uses an existing operation, it should require little or no new VBA Apply code. Only genuinely new Word behavior adds a mediator operation.

Estimated combined difficulty: **42–57%**.
Estimated professional desirability: **97–100%**.

### Generic Tool Form

A configurable guided form could display a tool's introduction, presets, custom choices, risk labels, scope, Preview, Reset, and explanations from its definition. Specialized tools such as Duplicate Remover and MetaDataSuite could retain custom forms.

Estimated combined difficulty: **45–60%**.
Estimated professional desirability: **94–98%**.

### Expected Effort for New Tool Categories

| New tool category | Current VBA-only difficulty | Hybrid difficulty after foundation | Professional desirability |
| --- | ---: | ---: | ---: |
| Deterministic text repair | 35–50% | 12–25% | 95–99% |
| Context-sensitive language analysis | 60–75% | 25–42% | 95–99% |
| Opposite mode in an existing tool | 50–65% | 18–32% | 91–97% |
| Word formatting tool | 45–65% | 32–50% | 89–96% |
| Structural Word tool | 60–80% | 48–68% | 88–96% |
| Metadata or closed-file tool | 55–75% | 20–38% | 94–99% |

Estimated reduction in effort for ordinary new tools: **35–60%**.
Estimated reduction in defects: **40–65%**.
Estimated professional desirability of the extensibility model: **97–100%**.

### Opposite Tools Become Modes

Many requested “opposite tools” should become clear directions within the existing tool rather than separate launcher entries:

```text
Punctuation Normalizer
├─ Typographic → plain text
└─ Plain text → typographic

Soft Return Converter
├─ Soft return → paragraph
└─ Paragraph → soft return
```

They reuse the same candidates, Preview, navigation, and Apply adapter. Ambiguous reverse transformations remain review-first.

### Capability Discovery and Version Handshake

The VBA mediator should ask the engine which tools, schema versions, options, and Apply operations it supports. Unsupported tools remain hidden or disabled. If the engine and `.dotm` are incompatible, Apply is disabled and CleanupSuite recommends Repair.

Estimated combined difficulty: **34–48%**.
Estimated professional desirability: **96–99%**.

The engine should be modular internally, but arbitrary third-party DLL plugins should not be accepted initially. Third-party plugins would introduce signing, trust, compatibility, and support problems before there is proven demand.

Internal modularity:

Estimated combined difficulty: **32–45%**.
Estimated professional desirability: **97–100%**.

Third-party plugin system now:

Estimated combined difficulty: **72–88%**.
Estimated professional desirability: **48–66%**.

## Installer Evolution Under the Hybrid Model

The same state-aware Setup requirement remains, but Setup would manage a matched product set:

- `CleanupSuite.dotm`;
- CleanupSuite engine executable;
- signed rule/configuration assets;
- version/protocol information;
- per-user installation state;
- backups and user settings;
- install, update, repair, and uninstall behavior.

Repair must verify both the template and engine. Uninstall should remove program files and temporary payloads while offering to preserve intentional user settings. The version handshake prevents a partially updated installation from applying changes.

A self-contained engine would increase installer size but avoid a separate runtime prerequisite. Signing the installer, executable, and distributable VBA publisher identity becomes increasingly important as the product becomes significant.

Estimated combined difficulty: **58–72%**.
Estimated professional desirability: **97–99%**.

## Development Continuity and Context Compaction

On July 18, 2026, the C#/VBA hybrid decision was reached near the practical context limit of an unusually long design session. Chris asked what happens when the development conversation is compacted and specifically requested that the explanation become part of CleanupSuite's history.

There is no human-like sensation associated with compaction. Operationally, however, the change is significant: the complete working conversation is replaced by a structured briefing of the facts and decisions judged most important. Work can continue immediately, but the briefing may not preserve every detail of the original exchange. In particular, compaction can lose:

- the exact path by which a decision was reached;
- small preferences that were never promoted into a formal requirement;
- the emphasis behind a correction or objection;
- the distinction between an idea being discussed, recommended, approved, implemented, or published;
- technical details buried in an older diagnostic exchange.

The most serious risk is not forgetting a large headline. It is allowing a compressed summary to turn a proposal into an approval, erase a safety restriction, or blur the difference between local work and a published release.

CleanupSuite therefore treats durable repository documentation as part of the engineering control system. The conversation can support exploration, but accepted decisions must be written into an authoritative project document. The two complementary records are:

- this history and architectural record, which preserves consequential decisions and the reasoning behind them;
- `WHERE_WE_LEFT_OFF.md`, which preserves the precise current state, active restrictions, approval gates, validation evidence, and next safe action.

Before a foreseeable compaction, consequential decisions should be promoted from conversational context into these records. After compaction—or at the beginning of a resumed development session—the records should be read before implementation begins. Source code, tests, release artifacts, and Git history remain the final evidence of what was actually implemented or published.

This practice is analogous to a professional engineering handoff: continuity comes from the quality of the maintained record, not from assuming an uninterrupted internal memory. It becomes more important as CleanupSuite gains additional tools, a C# engine, a versioned VBA/engine protocol, and a state-aware installer.

Estimated combined difficulty: **4–9%**.
Estimated professional desirability: **98–100%**.

## The Collaboration That Shaped CleanupSuite

On July 18, 2026, after asking for a deliberate rereading of the project records, Chris asked that the nature of the collaboration itself be preserved for posterity. He summarized the necessary honesty this way: **“We are who we are, but we aren't who we aren't too.”**

That sentence describes an important project asset. CleanupSuite has not improved because a human and an AI are interchangeable. It has improved because they are not.

### What Chris Contributes

Chris is the product owner, author, and final decision-maker. He contributes things that cannot be manufactured from source inspection alone:

- the reason CleanupSuite should exist and the standard it is meant to reach;
- lived experience of using Word and CleanupSuite on real documents;
- visual and interaction judgment about whether a technically functioning result is actually understandable;
- sensitivity to when a feature is too intrusive, too destructive, too confusing, or insufficiently useful;
- the willingness to stop, reconsider, and reject an answer that does not feel right in practice;
- priorities, taste, personal responsibility for the product, and authority to decide what is approved or published.

The project record contains concrete examples. Full-paragraph green Preview shading made a change highly visible, but Chris recognized that it overwhelmed the document and obscured the location that mattered. The accepted one-character marker came from that judgment. Punctuation logic could be described as working on small documents, but Chris noticed the blinking and long-running behavior in the large “Argument for God” article and distinguished it from Capitalization. He also challenged an unchanged Duplicate Remover count when `Include empty paragraphs` was selected. That challenge led to the deeper discovery that the apparent blanks were required table-cell, row-end, and final-document markers rather than disposable paragraphs.

These were not cosmetic comments added after engineering was complete. They changed the engineering model.

### What Codex Contributes

Codex contributes a different kind of leverage:

- rapid inspection and synthesis across source, generated artifacts, tests, documentation, Git history, and live application behavior;
- the ability to translate an observation into explicit invariants, ownership boundaries, candidate classifications, regression tests, and release evidence;
- sustained attention to repetitive implementation and verification work;
- comparison with professional product, safety, installer, and architecture patterns;
- practical difficulty and desirability estimates to make tradeoffs visible;
- the ability to connect a local defect to a reusable system correction;
- durable handoffs that distinguish proposed, approved, local, validated, committed, and published states.

From the Codex side, there is no human experience of Word, personal ownership of MasseysLab, continuous autobiographical memory, or independent authority to redefine the product. There is also no honest basis for pretending that a plausible implementation is correct merely because it compiles or passes a test. The useful contribution is analysis, construction, verification, and candid recommendation—not impersonation of the person who owns and uses the result.

### What Neither Side Should Pretend to Be

Chris should not have to function as the project's compiler, exhaustive dependency tracker, regression harness, or mechanical memory of every generated artifact. Codex should not pretend to possess Chris's lived judgment, intentions, responsibility, or final authority. Neither should treat the other as infallible.

The connection needed for this project is therefore not blind trust or automatic agreement. It is **trust calibrated by evidence**:

- Chris's observations are treated as product evidence, not dismissed because the code appears correct.
- Codex's recommendations are accompanied by reasoning, consequences, uncertainty, and verification—not presented as authority that replaces a decision.
- Disagreement is useful when it exposes a hidden assumption.
- Approval gates remain explicit; enthusiasm is not silently converted into authorization.
- Tests settle mechanical claims where tests can reach them; real Word use settles experience claims that source inspection cannot prove.
- Important conclusions move from conversation into source, tests, contracts, release notes, the handoff, and this history.

### How the Combination Produces a Better Product

The strongest CleanupSuite pattern is a recurring loop:

1. Chris identifies a real user-facing need or an outcome that does not look or behave correctly.
2. Codex investigates the implementation, reproduces or classifies the behavior, and explains the underlying constraint.
3. Both sides refine the intended behavior until the product rule is precise.
4. Codex implements or records the rule, adds proportionate verification, and preserves the state needed for continuity.
5. Chris evaluates the result in the real application and either accepts it or supplies the next piece of evidence.

This loop has repeatedly converted a narrow symptom into a stronger system:

- intrusive spacing shading became the shared minimal-marker concept;
- punctuation flicker became the fast-path rule for deterministic transformations;
- visually empty table paragraphs became a shared structural-classification requirement;
- redundant Duplicate Detector choices became a clearer Preview-first remover;
- a large suite-wide professional audit became the approved C#/VBA division of responsibility;
- context compaction became a durable documentation and handoff discipline.

Neither side working alone would likely have produced the same result. Human-only development would carry a much larger burden of exhaustive code generation, dependency tracking, repetitive validation, and cross-document consistency. AI-only development could produce substantial machinery quickly while missing the meaning of visual overload, practical trust, intentional document structure, or an unspoken product boundary. The combination works when speed remains answerable to judgment and judgment is converted into durable engineering.

### Working Compact for the Future

For future CleanupSuite work, this relationship should preserve the following practices:

1. **Tell the truth about capability and uncertainty.** Do not invent certainty, memory, validation, or human experience.
2. **Lead with the observed outcome.** A screenshot, real document, surprising count, or awkward interaction can be more valuable than a speculative redesign.
3. **Investigate before defending.** A user objection is a reason to inspect the system, not a reason for reflexive agreement or resistance.
4. **Make recommendations candidly.** Codex should say what it actually recommends, including tradeoffs and the requested difficulty/desirability estimates.
5. **Keep authority clear.** Chris decides product intent, accepts consequential design choices, and authorizes commits and publication. Codex may act autonomously only within the authorized implementation scope.
6. **Use professional evidence without surrendering product taste.** External conventions inform a decision; they do not automatically overrule the intended CleanupSuite experience.
7. **Automate the mechanical burden.** Source generation, candidate analysis, tests, synchronization, hashes, and installer checks should free human attention for judgment rather than consume it.
8. **Reserve human involvement for meaningful gates.** Program-level work and automated verification should come first; live Word review belongs at visual, behavioral, safety, and release boundaries.
9. **Preserve reversibility and provenance.** Backups, Preview, Undo, state restoration, candidate revalidation, and explicit release status protect both the document and the collaboration.
10. **Write down what matters.** Compaction, time, and project growth must not be allowed to turn an accepted principle back into an ambiguous conversation.

### Note for Posterity

CleanupSuite's history is not only a sequence of versions. It is the record of a product becoming more deliberate through observation, challenge, explanation, revision, and proof. The human contribution is not a ceremonial approval at the end, and the AI contribution is not merely typing code after receiving a complete specification. The specification itself is often discovered through the collaboration.

The durable aspiration is not for either participant to imitate what the other is. It is to understand the boundary well enough that each can contribute what the other lacks. If that honesty is maintained, CleanupSuite can continue becoming better than either participant would have made it alone.

Estimated combined difficulty of maintaining this working discipline: **8–16%**.
Estimated professional desirability: **98–100%**.

## Durable Product Principles

The following principles have emerged repeatedly across releases, diagnostics, and design decisions:

1. **Preview first.** Preview is the primary safety and comprehension system.
2. **Undo is necessary but not sufficient.** Structural, privacy, and removal operations may need confirmations, backups, or protected defaults.
3. **Do not confuse visual emptiness with structural emptiness.** Word markers can be required content.
4. **Never make advanced destructive behavior accidental.** Keep it OFF and outside All/Recommended/Reset.
5. **Preview and Apply must agree.** They should consume the same candidates.
6. **Explain protected and skipped items.** A count that did not change should have a reason.
7. **Use minimal markers.** Show the location without visually overwhelming the document.
8. **Preserve the user's document state.** Highlights, formatting, selection, view, zoom, scrolling, and Word options must be restored.
9. **Use the fastest semantically correct path.** Bulk exact replacement is good; blind contextual replacement is not.
10. **Keep ownership clear.** Table Cleaner owns table structure; Final Review owns review markup; MetaDataSuite owns metadata/privacy.
11. **Keep the source authoritative.** Generated artifacts are outputs, not the normal editing surface.
12. **Make tests part of the architecture.** Tests document the contract and protect the generated Word artifacts.
13. **Do not broaden scope silently.** Main body, selection, and all stories have different meanings.
14. **Do not publish without explicit approval.** Local work remains local until Chris authorizes the release action.
15. **Promote decisions into durable records.** Do not rely on conversational context alone to preserve approvals, restrictions, implementation state, or release state across compaction and handoff.
16. **Use complementary judgment honestly.** Let human product judgment and real-use evidence direct intent; let analysis, automation, tests, and documentation make that intent durable and verifiable.

## Decision Register as of July 26, 2026

| Decision | Status |
| --- | --- |
| CleanupSuite is a preview-first professional cleanup assistant, not an automatic replacement for human review | Established |
| `0.9.2-alpha` is the final published Alpha hotfix before Beta development | Published |
| Current local post-`0.9.2-alpha` work becomes the `0.9.5 Beta` baseline; no additional Alpha hotfix will be prepared | Approved July 18, 2026 |
| Minimal one-character markers are preferred over full-paragraph shading | Implemented and included in `v0.9.3-beta-vba-final` |
| Duplicate Detector becomes Duplicate Paragraph Remover with Exact/Normalized/Fuzzy structure | Implemented and included in `v0.9.3-beta-vba-final` |
| Shared blank classification and structural protections | Implemented and passed source, generated-fixture, Word-native runtime, visual Preview, confirmation, Apply, and one-step Undo checks July 18, 2026 |
| Spacing Fixer must not delete table rows | Implemented locally, protected by a negative regression contract, and included in the Word structural gate July 18, 2026 |
| Four equal Page/Change navigation buttons with the approved layout and colors | Implemented and live-validated July 26, 2026 |
| Change navigation means nearest page containing changes, not next individual change | Implemented and live-validated July 26, 2026 |
| Final standalone `.dotm` release boundary | Frozen as `v0.9.3-beta-vba-final` after Task 3 and before C# tool dependency |
| State-aware single Setup for Install/Update/Repair/Uninstall | Beta contract |
| Suite-wide preservation of pre-existing highlights | Implemented through preview-owned restoration records in Task 3 |
| Versioned hybrid candidate manifest and operation vocabulary | Contract 1.0 defined and tested July 26, 2026 |
| Tool-specific professional corrections from the July 18 audit | Proposed |
| C# sidecar engine with VBA as Word mediator | Approved and authorized; scheduled after immediate safety and contract work |
| Incremental migration rather than a pre-Beta rewrite | Approved and authorized direction |
| Generic tool definitions and standard Apply operation vocabulary | Accepted extensibility direction; details to be proven during implementation |
| Reverse/inverse modes in the final VBA-only release | Explicitly deferred |
| Third-party engine plugins | Deferred/not recommended now |
| Repository history and handoff documents are part of the development-continuity control system | Established |
| Human product judgment and AI engineering leverage operate as complementary, explicitly bounded responsibilities | Established collaboration principle |
| `last-backup-before-hybrid_20260718_205525` is the frozen recovery checkpoint before C# engine implementation | Created and verified July 18, 2026 |
| Beta execution minimizes paid user interaction, batches questions, and continues non-Word work during locked or unavailable interactive periods | Established July 18, 2026 |

## Immediate Historical Checkpoint

### First Beta implementation pass

After the explicit `go`, work began with the accepted structural-safety package. The first local pass established a shared, fail-closed classifier for Word blank-paragraph and table structures and routed Spacing Fixer, Paragraph Structure Fixer, Document Trim, Duplicate Paragraph Remover, and Table Cleaner through the appropriate shared rules. It removed Spacing Fixer's former ability to delete table rows, made Table Cleaner the sole owner of row and column deletion, protected nested and merged structures, aligned single-column conversion between Preview and Apply, and changed Break Normalizer's structural collapse defaults to OFF.

The implementation also added a deterministic `Regression - Structural Safety.docx` fixture, mirrored it into the practice package, added static safety contracts, and added a Word-native smoke entry point that exercises classification, revalidation, final-marker protection, NBSP protection, and nested/merged table guards using real Word objects. The focused source gate passed 74 tests and 223 subtests, and all 29 generated-VBA builders passed. Word-side compilation, runtime execution, visual fixture review, Preview/Apply parity, and Undo verification remained the final gate at the time of this entry.

The only full-suite failure at this checkpoint was the intentionally stale release-manifest hash: local Beta Word artifacts no longer match the published Alpha manifest, and the manifest was deliberately left unchanged because publication had not been authorized.

### Word fixture corrected an over-protective table assumption

The final live Task 2 review found an important difference between tables created through Word VBA and tables generated as fixed-geometry Open XML fixtures. `Table.Uniform` returned `False` for every generated fixture table, including ordinary rectangular tables with consistent rows and columns. Treating that property as a safety gate therefore protected everything: Preview reported zero eligible rows, zero eligible columns, and zero conversion candidates.

The correction removed `Table.Uniform` as the deciding test. CleanupSuite now verifies a rectangular table by successfully addressing every row/column cell; merged or otherwise nonrectangular grids fail closed when that addressing is invalid. Nested-table detection uses Word's end-of-cell marker invariant: an ordinary cell range contains one `Chr(7)` marker, while a cell containing a nested table contains additional cell markers. This is substantially faster than serializing each cell's full `WordOpenXML`, and it behaved consistently in both runtime-created tables and the generated fixture.

The smoke gate was strengthened so Word opens and validates the real fixture in addition to creating its own runtime tables. After the correction, live Table Cleaner Preview reported 1 removable empty row, 2 removable empty columns, and 1 eligible single-column conversion while keeping merged/nested structures in protected counts. The count-specific confirmation repeated the same totals and focused **No** by default. Break Normalizer's structural choices were also visually confirmed OFF.

Automated evidence at this checkpoint was 218 passing tests, 3 skips, and 1,095 passing subtests. The only full-run failure was intentional: the unchanged published-Alpha manifest no longer matches the synchronized local Beta template. Excluding that known condition yielded 218 passed, 3 skipped, 1 deselected, and 1,095 passing subtests. Word's temporary fixture owner file disappeared normally when Word closed. The remaining Task 2 live gate was one disposable structural Apply followed by a one-step Undo; it was not forced after the protected confirmation default because the Windows-control connection stopped accepting focus.

### One-step Undo exposed and corrected a shared lifecycle boundary

The first live Apply/Undo attempt did not pass. Table Cleaner removed 1 eligible row, 2 eligible columns, and converted 1 eligible table, but one Ctrl+Z restored only the table conversion. A fresh Preview reported 0 eligible rows, 0 eligible columns, and 1 conversion. This was important evidence that the primary mutations were safe but were not yet presented to Word as one complete reversible user action.

An isolated Word diagnostic proved that `UndoRecord.StartCustomRecord` correctly grouped the same row, column, and conversion operations. The failure therefore came from CleanupSuite's surrounding lifecycle. The shared Preview Actions form restored formatting before Apply, then could restore it again while unloading. More importantly, Table Cleaner ended its custom Undo record before the shared cleanup-finished/applied lifecycle ran. That lifecycle can retry preview-format restoration, so document writes could be recorded above the structural custom record.

The correction established three reusable rules:

1. Preview Actions restores the editable preview document only once per panel instance.
2. Repeated applied-state notification is idempotent and cannot repeat transient document cleanup.
3. Table Cleaner's custom Undo record remains open until all document-affecting cleanup-finished and applied lifecycle work has completed.

After rebuilding and reinstalling the Startup template, the exact live sequence was repeated. Initial Preview again reported 1 eligible row, 2 eligible columns, and 1 eligible conversion. Apply performed those changes after the default-No count confirmation. Exactly one Ctrl+Z was issued. A new Preview then returned to **1 row / 2 columns / 1 conversion**, with **3 protected rows / 2 protected columns / 5 skipped multi-column tables / 2 protected conversions** intact. The fixture was closed without saving.

The completed Task 2 automated evidence is 220 passing tests, 3 skips, and 1,095 passing subtests, with the single intentionally stale published-Alpha manifest comparison as the only failure. Excluding that one known condition yields 220 passed, 3 skipped, 1 deselected, and 1,095 passing subtests. The focused structural, preview-lifecycle, and Undo-boundary gate passes 79 tests and 19 subtests. The final Word-native structural report is `docs/test-results/structural_safety_smoke_20260718_225756.txt`.

This episode reinforced a product principle: a tool's Undo boundary must include the complete document-affecting lifecycle, not merely the visibly primary mutations. Preview restoration, cleanup markers, and teardown are part of the operation whenever they can write to the document.

Estimated combined difficulty: **28–40%**.
Estimated professional desirability: **99–100%**.

Estimated combined difficulty: **58–70%**.
Estimated professional desirability: **98–100%**.

Estimated combined difficulty: **58–70%**.
Estimated professional desirability: **97–99%**.

Chris gave the explicit `go` command on July 18, 2026. The approved structural package is the first implementation task; the hybrid contract and engine follow in the recorded dependency order. No hybrid-engine code had been created at the moment the gate opened, and no suite-wide professional-audit package had been implemented. The current local post-Alpha work remains uncommitted and unpublished, but Chris has accepted it as the starting baseline for `0.9.5 Beta`; development will not return to an additional Alpha hotfix track.

Before hybrid implementation, a major persistent directory snapshot and matching ZIP archive were created under `C:\Users\Chris\PROGRAMMING\VBA\CleanupSuite external backups` with the label `last-backup-before-hybrid_20260718_205525`. The snapshot includes the ordered pre-hybrid task list and recovery metadata. The ZIP SHA-256 is `5630503AD738DEB665D79FC2572175B2666AE43A9D6C28922F21D255DBCA5868`.

The architecture decision is now settled: CleanupSuite will evolve incrementally toward a C# analysis engine with VBA retaining Word-facing mediation, Preview, navigation, Apply, and Undo. The implementation sequence must preserve the working product and avoid a rewrite-before-Beta. The current recommended sequence is:

1. implement the approved immediate safety corrections in a way that is compatible with a future candidate contract;
2. define and version the engine request/response schema, tool definition, candidate record, and standard Apply-operation vocabulary;
3. create the isolated C# sidecar and prove the bridge with Unicode Cleaner, Duplicate Paragraph Remover, and Punctuation Normalizer;
4. add the VBA/engine version handshake and extend the state-aware installer to manage the matched `.dotm` and engine;
5. migrate additional analysis incrementally while leaving live Word manipulation in VBA.

The exact scheduling within `0.9.5` remains an implementation-planning question, not an architectural question. The implementation command was given on July 18, 2026. Chris separately authorized publication of the final standalone VBA-only checkpoint after Task 3; official hybrid Beta publication remains a later gate.

### Task 3 completed the final standalone VBA-only product

Task 3 corrected a second shared lifecycle boundary: Preview formatting and
navigation are product infrastructure, not per-tool decoration. CleanupSuite now
records only the temporary Preview formatting it owns and restores those exact
records, avoiding damage to formatting that existed before Preview.

The two-button Preview navigator became four equal controls:

`[Change] [Page] ◀ Previous | Next ▶ [Page] [Change]`

The inner pale-blue controls turn one page. The outer pale-gold controls use a
page registry built during Preview and jump to the nearest page that contains one
or more changes. Chris clarified that “Change” must never mean stepping through
every individual result; twenty results on the current page must not require
twenty clicks before reaching the next changed page.

The most stubborn regression occurred in Word Reading View. Backward navigation
could update `Selection.Start` without repainting the visible page. The stable
solution moves Word's internal page and then gives the Reading View canvas the
same focused Left input that causes Word to repaint, while the modeless Preview
panel is temporarily hidden. This retained the fast display-only punctuation path
and avoided the heavy blinking seen before the earlier performance fix.

The final live gate used a disposable copy of the 53-page Argument for God
document. Punctuation Preview completed in approximately 35 seconds, showed
minimal markers, moved forward to the next changed page, moved visibly backward
to the prior changed page, and kept Reading View and the panel stable. The copy
was closed without saving.

This result was frozen as `0.9.3-beta-vba-final`, the final complete release that
does not require the future C# engine. It is a manual-install `.dotm` prerelease;
the `0.9.2-alpha` installer and stable update manifest remain unchanged. Reverse
tool modes were deliberately left for reconsideration after the hybrid foundation.

Estimated combined difficulty: **55–70%**.
Estimated professional desirability: **98–100%**.

### July 26, 2026: Hybrid risk gate and change-control thresholds

Before Task 4 began, Chris asked for a fresh reconsideration of the hybrid
decision. The two most consequential concerns were identified as:

1. the document can change between engine analysis and VBA Apply;
2. antivirus and Windows reputation systems may scrutinize an executable more
   heavily than a VBA-only template.

The decision remained in favor of an incremental hybrid, but both concerns became
hard contract requirements rather than later implementation details.

For stale analysis, VBA must validate the complete analyzed scope before accepting
an engine result into Preview and again before Apply. It must then validate every
candidate's exact range and any required structural fingerprint. A mismatch aborts
the entire Apply before the first mutation and instructs the user to run Preview
again. Destructive candidates are never automatically relocated or approximately
matched.

For executable trust, the engine must run locally, offline, on demand, without
administrator rights or a Windows service. VBA must verify the matched engine
version and published hash before launch. Job paths and filenames are constrained,
logs exclude document content by default, the engine never edits Word, and the
official hybrid Beta requires a signed and hash-verified distribution path. The
final VBA-only release remains available for environments that prohibit companion
executables.

Chris also established a durable implementation-governance rule:

- minor implementation adjustments may be made on the fly;
- medium changes require at least **95%** estimated professional desirability;
- large changes require at least **99%**;
- huge changes require at least **99.9%**, high certainty, or necessity for
  continuation of the project;
- correctness is mandatory even when speed is preferred;
- extended rabbit holes require an explicit reconsideration of value and scope.

These percentages are practical product judgments, not scientific measurements.
Their purpose is to prevent scope drift while still allowing efficient engineering.

Estimated combined difficulty: **18–28%**.
Estimated professional desirability: **99–100%**.

### Task 4: Contract before engine

Task 4 deliberately added no C# executable and connected no Word tool to an
engine. It first converted the architectural promises into Contract 1.0 under
`contracts/hybrid/v1`.

The contract defines:

- exact UTF-8 snapshots without Word newline or marker normalization;
- Word-compatible UTF-16 Range offsets;
- versioned requests, responses, capabilities, and installation manifests;
- applicable, review-only, protected, and skipped candidate states;
- deterministic/high/medium/low confidence;
- exact text, surrounding-context, paragraph, and structural fingerprints;
- a bounded Apply-operation vocabulary whose operations remain proposals until
  independently revalidated and executed by VBA;
- fixed owner-only job filenames and path-traversal/reparse-point rejection;
- no network, service, elevation, Word automation, or document-content logging;
- no partial-result reuse, approximate relocation, or silent behavior-changing
  fallback;
- an official-Beta Authenticode requirement and a matched component hash manifest.

A live Word probe verified the most important text-coordinate assumption. In a
story containing `A😀B` and Word's final paragraph mark, Word treated the emoji as
one `Characters` item but advanced its Range positions from 1 to 3. Contract 1.0
therefore correctly defines all offsets in UTF-16 code units and includes a
supplementary-Unicode regression fixture.

The contract gate adds JSON Schema validation and deterministic positive and
negative fixtures. At completion, the focused contract suite passed 13 tests and
the complete repository suite passed 227 tests with 3 expected skips. All 29 VBA
assembly builders remained clean. No Word runtime or release artifact changed.

Estimated combined difficulty: **42–57%**.
Estimated professional desirability: **97–100%**.

### Task 5: Isolated engine before Word integration

Task 5 created the first C# component without crossing the release boundary. The
new `CleanupSuite.Engine` is an offline, on-demand `net8.0-windows` command-line
process that implements Contract 1.0. It is not installed, not shipped, not called
by Word, and not used by a real CleanupSuite tool.

The engine accepts only a lowercase-UUID job directory directly beneath
CleanupSuite's fixed LocalAppData job root. It requires an owner-only protected
directory, rejects reparse points and unsafe paths, reads only fixed filenames,
strictly validates UTF-8 JSON and exact snapshot hashes and lengths, and writes a
result atomically without overwriting an existing result. Cancellation and every
failure path return zero candidates. Logs contain event codes and counts rather
than document text or source paths.

The Task 5 analyzer is deliberately artificial:
`contract-fixture/replace-literal`. Its purpose is to prove deterministic,
Word-compatible UTF-16 offsets, surrogate-safe context, fingerprints, capability
discovery, cancellation, failure behavior, and schema-valid result transport. It
is not a migrated product tool.

The focused engine gate passed 12 end-to-end tests. The actual binary's capability
output matches Contract 1.0 and its actual two-candidate result validates against
the Contract 1.0 JSON Schema. The complete repository suite passed 233 tests with
3 expected skips, all 29 VBA builders remained clean, and C# formatting required
no changes. Local Microsoft Defender scans reported no new detections for either
the framework build or the self-contained proof. This is early local evidence,
not a substitute for Authenticode signing or reputation testing on clean machines.

The development publish proof produced a single-file self-contained `win-x64`
engine of 35,106,239 bytes with SHA-256
`32BD795E91FDB722BF6CD35045B9E00740239B79545B9F887C0ACC56B9AB824D`.
The proof is unsigned, ignored, and nondistributable. Chris clarified the durable
packaging requirement: using .NET internally is acceptable, but requiring users
to install, update, select, or maintain .NET separately is undesirable. The
official product must therefore ship the matched engine with its runtime
self-contained.

Task 5 changed no VBA runtime, installed template, release artifact, installer, or
stable manifest. Task 6 remains the hard boundary where a real pilot tool first
depends on a matched engine and VBA bridge.

Estimated combined difficulty: **48–62%**.
Estimated professional desirability: **96–99%**.

### Task 6 checkpoint: the first real hybrid pilot

Task 6 crossed the source-code boundary deliberately left untouched by Task 5.
Invisible Unicode Cleaner is now the first real tool whose analysis can be
performed by the matched C# engine while VBA continues to own Word, Preview,
navigation, Apply, and Undo.

Engine 0.2.0 analyzes only the explicitly selected invisible Unicode categories.
It returns deterministic, document-order, one-UTF-16-unit replacement candidates
for nonbreaking spaces, zero-width spaces, zero-width nonjoiners, zero-width
joiners, byte-order marks, soft hyphens, and nonbreaking hyphens. The engine never
opens Word or mutates the document.

The VBA bridge is intentionally strict. It:

- creates and hardens a per-job owner-only directory through the trusted engine;
- writes exact UTF-8 snapshots and requests without a shell;
- launches the engine directly with `CreateProcessW`;
- validates engine identity, version, capability allowlists, protocol fields,
  response shape, candidate ranges, fingerprints, replacements, and summaries;
- rejects duplicate or unknown JSON fields and invalid Unicode;
- revalidates the complete snapshot, scope, each exact candidate, surrounding
  context, and paragraph immediately before Apply;
- aborts before the Undo record and before the first mutation if anything is
  stale;
- applies accepted replacements from the end of the document toward the start in
  one VBA-owned Word Undo record.

Preview uses the existing minimal-highlight infrastructure. Visible characters
receive exact one-character highlighting. Characters that Word cannot visibly
shade use the nearest safe one-character cue rather than a whole paragraph or
page. The legacy full-scan Unicode Preview and Apply paths are no longer active.

Development trust is explicit and intentionally inconvenient to choose
accidentally: it requires both an absolute engine path and its exact SHA-256 in
process-local development environment variables, followed by the same strict
capability handshake. Ordinary product operation has no unsigned or
version-mismatched fallback. Official trust remains dependent on the matched
installer manifest, hash verification, and Authenticode.

The final development package proof is a self-contained `win-x64` executable of
35,108,802 bytes with SHA-256
`BDEB2CD8578B31946DF0731B80FF3EC8CC2336C807726E5112E5A7C64E36DA44`.
CleanupSuite may use .NET internally, but requiring users to install, update,
select, or maintain .NET separately is a rejected product design. The engine and
its runtime must be installed and serviced as one CleanupSuite component.

Automated evidence at this checkpoint:

- 15 isolated engine contract and analyzer tests passed;
- 249 repository tests passed with 3 expected skips and 1,105 passing subtests;
- all 31 VBA builders passed, producing 19,717 lines and 1,361,767 bytes;
- strict C# formatting verification passed for both engine projects;
- disposable Word automation passed primitive bridge, seven-candidate analysis,
  stale-scope rejection, reverse-order Apply, exact one-step Undo restoration, and
  a 200,000-character/400-candidate performance probe;
- the large bridge probe completed analysis in approximately 21.75 seconds and
  then passed full pre-Apply revalidation;
- a local Microsoft Defender scan of the current self-contained proof reported no
  threats.

The installed Startup template, final VBA-only release template, stable installer,
and stable update manifest remain unchanged by the hybrid pilot. The source is not
yet an official distributable hybrid build.

The final human-visible Word gate passed on July 26, 2026. A disposable macro
document containing one example of every supported category opened the real
Invisible Unicode Cleaner form through the CleanupSuite launcher. Preview reported
exactly one candidate in each of the seven rows. Word showed narrow
one-character—or, where the character itself was not safely visible,
nearest-one-character—green cues at the seven actual boundaries. It did not shade
whole paragraphs, pages, or unrelated text. The probe was closed without saving,
and no Trust Center setting, installed template, release artifact, or user
document was changed.

That visual result completes Task 6: the first hybrid pilot now has source,
contract, performance, Apply/Undo, stale-state, security, packaging, antivirus,
and human-visible Preview evidence. Distribution remains intentionally deferred
until the state-aware installer can place and verify the matched signed engine.

The release-integrity regression was updated to reflect the intentional two-track
state: the installer manifest remains pinned to stable `0.9.2-alpha`, while the
repository's manual-install `CleanupSuite.dotm` is the separately hash-documented
`0.9.3-beta-vba-final` checkpoint.

Estimated combined difficulty: **55–68%**.
Estimated professional desirability: **97–99%**.

### Task 7 checkpoint: unhighlighted results became reviewable

The Preview table had long been able to label a count as `unhighlighted`, but a
count alone did not let a serious Beta tester inspect the underlying condition.
Task 7 added a compact review lane without altering the accepted four-button Page
and Change navigation layout.

When an active unhighlighted row has a nonzero count, Preview Actions now shows a
subtle pale-blue **Review details** control and a concise context area. Tools may
register exact review findings with a document range, category, and explanation.
Each activation moves to the next locatable finding and reports its ordinal,
category, and context. Findings such as document properties that have no honest
text location instead receive a document-level explanation; CleanupSuite does not
invent a misleading highlight.

Duplicate Paragraph Remover is the first tool to provide detailed protected
locations. When **Include empty paragraphs** is active, required empty-cell
markers, table-row markers, the final-document marker, and unsupported blank
structures can be reviewed one at a time. The tool still leaves them unchanged,
and table-row deletion remains exclusively owned by Table Cleaner.

The shared review collection is preview-scoped, document-bound, deduplicated, and
cleared on Apply, Reconfigure, Preview teardown, errors, and form termination.
The disposable Word lifecycle smoke verified both branches: a locatable finding
moved Selection to the registered range and displayed context, while a
nonlocatable document-property count displayed its scope explanation without
offering false navigation. A human-visible probe confirmed that the extra control
is compact, the summary remains readable, and the existing Page/Change controls
retain their equal size and accepted colors.

Task 7 also reconciled the Beta records with Chris's accepted deferral of inverse
tool modes. Straight-to-curly quotes remain the first sensible future candidate,
but reverse modes are not a `0.9.5-beta` release gate and will not be rushed into
the hybrid installer transition.

Automated evidence at this checkpoint:

- 250 repository tests passed with 3 expected skips and 1,118 passing subtests;
- all 31 VBA builders passed;
- the isolated Word wiring smoke passed for Capitalization, MetaDataSuite, Final
  Review, Preview Actions, guided risk controls, and representative tool forms;
- the disposable Preview lifecycle smoke passed Reading View, Page/Change
  navigation, review-detail navigation/context, formatting restoration, switched
  windows, closed documents, and progress cleanup;
- both tracked development `.docm` files, the installed Startup template, release
  `.dotm`, Setup, and stable update manifest remained unchanged.

Estimated combined difficulty: **28–40%**.
Estimated professional desirability: **96–99%**.

### Task 8 checkpoint: Setup became the hybrid maintenance boundary

The first hybrid installer implementation was completed on July 26, 2026 without
changing the published Alpha Setup, stable release manifest, release template, or
the user's installed Startup template.

The same Setup executable now inspects the per-user installation and presents
Install when absent; Update, Repair/Reinstall, and Uninstall for an older or
legacy installation; Repair/Reinstall and Uninstall for the current package; and
a protected no-downgrade state when a newer package is present. The downloaded
Setup is an offline maintenance unit containing the matched Word template,
self-contained engine, protocol definitions, operation rules, and a manifest
binding all four by version, approved relative path, byte length, and SHA-256.
The user never installs or maintains .NET separately.

Maintenance is transactional. Setup stages and verifies every component, backs
up the prior Word template, retains only the newest three installer backups,
captures the affected installation files for rollback, installs the manifest
last, verifies the final package, and removes staging material. User settings,
`Normal.dotm`, Word security settings, Trust Center configuration, and trusted
locations remain outside Setup's authority.

The production VBA bridge now uses that installed manifest when no explicit
repository-development engine override is present. It accepts only the matched
suite/protocol versions and approved component identities and paths, verifies
file lengths and SHA-256 hashes, and performs the existing capability handshake.
A mismatch fails closed before analysis or Apply and recommends Setup
Repair/Reinstall.

The real Setup executable passed an isolated state matrix covering absent,
legacy, older, current healthy, current damaged, Update, Repair/Reinstall, and
Uninstall. An intentional failure after engine replacement restored the exact
pre-repair damaged engine and left no staging directory. Repeated repairs retained
exactly three backups, and Uninstall removed runtime components while preserving
backups and a user-settings sentinel.

A separate disposable Word run then removed the development engine override and
proved the production installed-manifest path end to end. Word verified the
matched package, completed the capability handshake, received all seven Unicode
candidates, and rejected a deliberately stale scope before Apply. During that
gate, two harness-relevant defects were corrected: `.dotm` synchronization now
preserves the template extension instead of temporarily renaming it `.docm`, and
the already-loaded Word template is hash-read without requesting a conflicting
write lock while engine and protocol files retain the stricter lock.

The Setup UI was also reviewed live at the active Windows display scale. The
absent, older, current, and damaged states were clean, readable, and exposed only
their intended actions.

Official distribution is still gated on Authenticode. A release build requires a
signing certificate, signs the self-contained engine before the package hash is
created, then signs the completed Setup. Clean-machine signature/reputation
testing and final publication remain release work.

Automated evidence at this checkpoint:

- 256 repository tests passed with 3 expected skips and 1,131 passing subtests;
- all 31 VBA builders passed after the production manifest resolver was added;
- the isolated installer matrix passed its state, rollback, backup, settings, and
  cleanup checks;
- the disposable production-manifest Word smoke passed without a development
  engine override;
- the absent, older, current, and damaged Setup layouts passed live review;
- the stable release directory and both tracked `.docm` files remained unchanged.

Estimated combined difficulty: **72–84%**.
Estimated professional desirability: **98–99.7%**.

### The first signed candidate must not also be the publication

The release-engineering closeout on July 26, 2026 found one remaining avoidable
risk: the original hybrid build switch signed the package only while writing
directly into the official `release` directory. That made it impossible to
inspect the first signed engine and Setup independently before those artifacts
replaced the stable release.

Candidate signing and publication were therefore separated. `-SignBuild` now
creates `build\installer-signed-candidate`, signs the staged self-contained engine
before its package hash is calculated, signs the completed Setup, and exposes a
copy of the exact matched payload for independent signature, hash, and security
inspection. It does not modify `release`. `-PublishRelease` remains the later,
intentional action that writes the official package and release pointer after the
candidate has passed every gate. Publication itself compiles and signs outside
`release`, captures the prior template, Setup, and release pointer, and restores
all three if artifact replacement fails.

The correction was followed by the complete repository gate (257 passed, 3
expected skips, and 1,131 passing subtests), all 31 VBA builders, the full
installer maintenance matrix, and the production installed-manifest Word smoke.
Microsoft Defender, enabled with current definitions, found no threats in the
exact unsigned development Setup and self-contained engine. The stable release
remained unchanged.

No code-signing certificate with a usable private key was present in either the
current-user or local-machine certificate store, and both candidate binaries
correctly reported `NotSigned`. The project therefore held publication at the
security boundary instead of weakening it. Remaining official-Beta work is
Authenticode signing, repeating the complete gate against the signed candidate,
and clean-machine signature/reputation, installation, maintenance, Word
Preview/Apply, and Undo validation.

Estimated combined difficulty: **18–30% remaining after a certificate is
available**.
Estimated professional desirability: **99–100%**.

### Free software should pursue a free trust path first

On July 27, 2026, Chris declined to pay a continuing code-signing fee for a free
application. Research found a credible alternative: SignPath Foundation offers
sponsored public-trust signing without charge to accepted open-source projects.
CleanupSuite appears to satisfy the principal published conditions—MIT-licensed
public source, existing releases, active maintenance, documented behavior,
per-user installation and uninstall, local document processing, and an
automated Windows validation workflow—but acceptance remains SignPath
Foundation's decision.

The project prepared the policy material needed for a responsible application:
a public code-signing policy with authorship, review, and signing-approval roles;
a privacy policy accurately describing local hybrid jobs and the optional
GitHub-hosted update check; and a plan for repository-linked, manually approved
engine-then-Setup signing. Sponsored signatures would display **SignPath
Foundation** as the Windows certificate publisher while the CleanupSuite product
publisher remains **MasseysLab**.

A safe unsigned-install guide was also added for the waiting period or a rejected
application. It requires the official repository, complete SHA-256 comparison,
and Microsoft Defender scanning before the user considers Windows' one-time
**More info** / **Run anyway** choice. It explicitly forbids disabling
SmartScreen, Defender, Word macro security, the Trust Center, or organization
policy. Template-only manual installation is limited to releases explicitly
labeled standalone or VBA-only; a hybrid release must keep its template, engine,
protocol, rules, and manifest matched.

The stronger signed hybrid release contract remains intact while the free path is
evaluated. An unsigned hybrid release would be a separate, explicit distribution
decision rather than an accidental weakening of the existing gate.

Estimated combined difficulty: **28–42%**.
Estimated professional desirability: **98–99.5%**.

### Initial channel clarification (superseded)

Chris's July 27, 2026 installation screenshots exposed a documentation failure:
the repository correctly listed `0.9.3-beta-vba-final` as the newest manual
prerelease, but the prominent generic Setup link still downloaded the stable
`0.9.2-alpha` installer. Setup therefore correctly offered 0.9.2, and Word
correctly reported 0.9.2 as current against the stable manifest, but the page had
not made that channel boundary clear before the click.

The README was reorganized so the first download decision explicitly names both
channels:

- stable Setup installs `0.9.2-alpha`; and
- the `0.9.3-beta-vba-final` checkpoint has no new Setup and requires its release
  asset to be installed manually.

The prominent 0.9.3 choice now links directly to its `CleanupSuite.dotm` asset
and manual-install notes. The update-check source now says **current on the
stable installer channel** and explains that manual-install prereleases are not
offered by that check. Frozen 0.9.2 and 0.9.3 release artifacts were not altered.

Estimated combined difficulty: **5–10%**.
Estimated professional desirability: **98–99%**.

### 0.9.3 becomes the single stable VBA-only release

Chris then made the product decision explicit: **0.9.3 is the stable release**.
The earlier attempt to preserve 0.9.2 as a separate stable installer channel
created exactly the confusion that the release page should prevent. A user who
selects the prominent Setup download must receive the current stable 0.9.3
product, not an older 0.9.2 package followed by an apparently contradictory
“up to date” message.

The `0.9.3-beta-vba-final` tag and internal source version remain unchanged
because they identify the already-frozen, hash-verified template. They are
technical identity, not a second public distribution channel. Public wording
identifies the product as **0.9.3 Stable — Final VBA-Only Release**.

A dedicated standalone Setup was added rather than weakening or repurposing the
future hybrid installer. It:

- embeds the exact frozen 0.9.3 template;
- refuses to build if the supplied template SHA-256 differs;
- verifies that embedded hash again before installation;
- installs per user without administrator rights or security-policy changes;
- keeps the three newest exact backups;
- detects whether CleanupSuite is absent, current, or another version;
- presents Install, Update, Repair/Reinstall, and Uninstall according to that
  state; and
- clearly labels the unsigned release and directs users to verify its published
  SHA-256 and scan it with Microsoft Defender.

The stable manifest, default Setup link, release notes, versioning text, safety
guide, and update message now describe one stable 0.9.3 channel. The manual
template remains an alternative installation method for environments where
Setup cannot be used. The hybrid 0.9.5 installer and its stronger matched,
signed-package release gate remain separate.

Validation of the exact stable artifacts completed with **265 passed, 3
skipped**, 1,131 subtests, 31 clean builders, the isolated installer
install/reinstall/uninstall matrix, exact template and Setup hash checks, and a
Microsoft Defender custom scan with protection enabled.

Estimated combined difficulty: **18–28%**.
Estimated professional desirability: **99–100%**.

## Primary Historical Sources

- [Current handoff](../WHERE_WE_LEFT_OFF.md)
- [Versioning and roadmap](../VERSIONING.md)
- [0.9.5 Beta contract](0.9.5-beta-contract.md)
- [0.9.5 Jump Board](0.9.5-jump-board.md)
- [Project priority list — June 10, 2026](CleanupSuite_Project_Priority_List_2026-06-10.md)
- [0.6.4 to 0.6.5 release notes](Release_Notes_0.6.4_to_0.6.5.md)
- [0.7.4 to 0.7.5 release notes](Release_Notes_0.7.4_to_0.7.5.md)
- [0.7.5 to 0.8.0 release notes](Release_Notes_0.7.5_to_0.8.0.md)
- [0.8.0 to 0.8.5 release notes](Release_Notes_0.8.0_to_0.8.5.md)
- [0.9.0 Alpha release notes](Release_Notes_v0.9.0-alpha.md)
- [0.9.1 Alpha release notes](Release_Notes_v0.9.1-alpha.md)
- [0.9.2 Alpha release notes](Release_Notes_v0.9.2-alpha.md)
- [0.9.3 Final VBA-only release notes](Release_Notes_v0.9.3-beta-vba-final.md)
- [MetaData incorporation map](metadata-incorporation-map.md)
- [Capitalization 0.9.5 checklist](capitalization-0.9.5-checklist.md)
- [Programmer's guide for adding tools](CleanupSuite_Programmers_Guide_For_Adding_Tools_v0.9.0-alpha.md)

This record should be updated when a consequential product, safety, installer, release, or architectural decision becomes accepted—not only when code is published.
