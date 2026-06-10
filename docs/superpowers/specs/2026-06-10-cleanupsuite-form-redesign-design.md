# CleanupSuite Form Redesign

Date: 2026-06-10

## Goal

Redesign the CleanupSuite Word VBA tool forms so they feel more useful, fluid, and trustworthy without drifting into a visual style that is unrealistic for Word UserForms.

The current suite works, but many forms feel old, repetitive, and slightly counter-intuitive because they rely on:

- long vertical checkbox stacks
- fixed-width layouts with weak visual hierarchy
- tiny `?` help buttons separated from the main action flow
- repeated scope and preview controls that all read the same way
- forms that feel like control collections instead of guided tasks

The redesign should keep the underlying tool logic intact while making the forms easier to scan, easier to trust, and more modern in tone.

## Current Constraints

This design assumes the product remains a Word VBA solution built with UserForms.

That means:

- layout is limited compared with a web app or desktop framework
- advanced effects should be avoided or simplified
- the design must survive the realities of Word fonts, control rendering, and fixed-size form behavior
- the suite should remain understandable to users who expect Office-like tooling

The visual direction can improve substantially, but the implementation should prioritize layout, grouping, control language, and button flow over decorative styling.

## Design Directions Explored

### Direction A: Structured Utility

This approach keeps the forms recognizably Office-like, but much cleaner.

Characteristics:

- stronger heading hierarchy
- clearer section grouping
- more intentional spacing
- better labels and button order
- calmer visual tone

Pros:

- easiest transition from the current suite
- low implementation risk in VBA
- improves clarity immediately

Cons:

- still reads as conservative
- does not go far enough to solve the "archaic" feeling

### Direction B: Guided Workflow

This approach treats each tool form as a short decision flow instead of a static pile of options.

Characteristics:

- modes come first
- scope and preview become distinct steps
- recommended defaults are visually elevated
- the user is guided toward the next action

Pros:

- directly improves counter-intuitive flows
- makes complex forms easier to use
- matches how people actually approach cleanup tasks

Cons:

- can still feel slightly old-fashioned if the styling remains too traditional
- requires more careful restructuring of each form

### Direction C: Compact Modern Panel

This approach uses denser grouping, stronger panels, clearer primary actions, and a more contemporary utility-tool visual language.

Characteristics:

- panel-based layout
- stronger top header and subtitle
- grouped decisions instead of long flat control lists
- clear action zone with preview/help/run treated as related controls
- a more intentional, app-like feel

Pros:

- strongest improvement in perceived quality
- best answer to the "Win95-ish" complaint
- makes the suite feel designed rather than assembled

Cons:

- easiest direction to overdo
- must be restrained enough to remain believable inside Word/VBA

## Chosen Direction

The selected direction is:

`Direction C as the base visual language, with the discipline of Direction B and the restraint of Direction A.`

In practice, that means:

- the suite should feel modern and panel-based
- forms should still guide users through a clear action sequence
- visual styling should stop short of pretending to be a full custom desktop app

This is the right balance between freshness, usability, and implementation realism.

## Example Interpretation

The strongest test case is `Invisible Unicode Cleaner`.

The approved mockup shows the right direction:

- a strong title and short explanatory subtitle
- a first-class mode section
- a dedicated custom-options panel
- a clearer scope panel
- a distinct run-mode panel for preview behavior
- a visually obvious action area

What changed conceptually from the current form:

- the form reads as a sequence of decisions
- defaults are easier to understand
- custom options feel secondary instead of mixed into the same visual tier
- help feels adjacent to action instead of detached
- the form is easier to trust at a glance

## Form System Rules

These rules should apply across the suite.

### 1. Every form gets a strong top identity

Each form should begin with:

- a clear tool title
- a one-sentence purpose line
- optional lightweight status or recommendation text when useful

This replaces the current pattern where the user often lands directly in controls before understanding the tool's intent.

### 2. Controls should be grouped by decision, not by control type

The redesign should organize forms into logical blocks such as:

- what kind of cleanup to run
- what content to include
- what the preview/run mode does
- what the scope is

This is more natural than grouping by radio buttons, checkboxes, and frames separately.

### 3. Primary actions must be visually obvious

Every form should have a clearly defined action area containing:

- `Help`
- `Preview`
- `Run`

`Run` should always be the dominant action.

`Preview only` should stop being treated as just another checkbox buried in the control list. It should become part of the action decision.

### 4. Recommended defaults should look recommended

When a form has a safest or most common option, that state should be visually reinforced.

Examples:

- `All invisible / problem characters`
- `Entire document`
- `Preview` as the first pass

This reduces hesitation and makes the forms feel more helpful.

### 5. Simple tools and complex tools should not be designed identically

The same visual language should be used across the suite, but its density should vary.

Use a lighter version for:

- `Document Trim`
- `Hyperlink Remover`
- `Soft Return Converter`

Use a fuller panel layout for:

- `Unicode Cleanup`
- `Formatting Stripper`
- `Duplicate Detector`
- `Object Remover`
- `Header/Footer Standardizer`

This keeps the suite coherent without making every tool feel oversized.

### 6. Help should support decisions, not interrupt them

Help content is valuable, but the current `?` pattern feels detached and tiny.

The redesign should treat help as:

- a visible nearby action
- optionally a short inline summary or tip area
- a supplement to the form, not the only explanation

The recent cleanup of forced line breaks in the help screens supports this direction.

## Proposed Form Architecture

Each tool form should be built from the same high-level structure:

1. Header
2. Decision panels
3. Context panel
4. Action area

### Header

Contains:

- tool title
- short explanation
- optional help entry point

### Decision panels

Contain:

- mode selection
- custom options when relevant
- thresholds or variants when relevant

### Context panel

Contains:

- scope
- preview meaning
- caveats when needed

### Action area

Contains:

- help
- preview
- run

This system should make the suite feel consistent while allowing each tool to vary based on complexity.

## Launcher Redesign

The launcher needs separate treatment.

Current issues:

- paired buttons plus tiny `?` buttons make the grid feel old
- tool names are short but not very descriptive in context
- the launcher feels more like a control board than a starting point

Recommended direction:

- convert the launcher from paired button rows into compact tool cards or list panels
- show each tool with title plus one short description line
- fold help into the card instead of relying on isolated tiny buttons
- preserve density, because the suite contains many tools

The launcher should feel like a curated toolbox, not a spreadsheet of commands.

## Tool Categories

The original category idea was useful, but it is too coarse for a few specific tools.

The main friction points are:

- `Table Cleaner` is partly structure cleanup and partly formatting cleanup
- `Break Normalizer` and `Soft Return Converter` both change structure, but they feel very different from paragraph/list cleanup
- `Footnote / Endnote Remover` is not really metadata, but it is also not purely text cleanup
- `Header / Footer Standardizer` is document structure plus formatting, not just formatting
- `Style Cleanup` is document-wide hygiene and does not sit naturally beside local formatting tools
- `Document Trim` and `Metadata Scrubber` should not share a generic `Document Cleanup` category, because one is layout/structure cleanup and the other is privacy/final-review cleanup

To make the launcher feel more intuitive, use categories that reflect the user's mental model of the task rather than the implementation type.

Recommended launcher categories:

- Text and Characters
- Paragraphs, Breaks, and Lists
- Layout and Document Structure
- Formatting, Links, and Styles
- Review, Privacy, and Removals

Recommended tool placement:

- Text and Characters
  - `Invisible Unicode Cleaner`
  - `Punctuation Normalizer`
  - `Spacing Fixer`
  - `Capitalization Fixer`

- Paragraphs, Breaks, and Lists
  - `List Normalizer`
  - `Paragraph Structure Fixer`
  - `Soft Return Converter`
  - `Duplicate Paragraph Detector`

- Layout and Document Structure
  - `Table Cleaner`
  - `Break Normalizer`
  - `Header / Footer Standardizer`
  - `Document Trim`

- Formatting, Links, and Styles
  - `Font Normalizer`
  - `Formatting Stripper`
  - `Style Cleanup`
  - `Hyperlink Remover`

- Review, Privacy, and Removals
  - `Metadata Scrubber`
  - `Footnote / Endnote Remover`
  - `Object Remover`

Why this is better:

- `Table Cleaner` is treated as a layout tool, which is how most users will think about it
- `Break Normalizer` sits near other page-structure tools instead of text tools
- `Soft Return Converter` has a better home because the category now explicitly includes breaks
- `Document Trim` sits with visible document structure tools instead of being paired with metadata
- `Metadata Scrubber` sits with privacy and removal tools, which better matches the user's intention before sharing
- `Footnote / Endnote Remover` stops being hidden in a category that feels too technical or purely metadata-oriented
- the categories are more descriptive of what the user is trying to clean up

What can be done in the launcher:

- show category bands or collapsible sections
- let the launcher open on the last-used category
- surface 3 to 5 "most used" tools above the categories
- allow search or quick filtering later if the suite keeps growing

This should make the launcher feel less like a flat command inventory and more like an organized workshop.

## Sub-tool Placement Review

Most sub-tools are in the right general area, but several options should not be treated as ordinary checkboxes in the redesigned UI.

The design should distinguish between:

- safe cleanup options
- conversion actions
- destructive removal actions
- final-review/privacy actions

This matters because the current forms sometimes put very different levels of consequence side by side.

### Keep inside the current tool, but redesign as clearer modes

These options belong with their current parent tool, but should be presented as modes or grouped panels instead of flat checkbox rows.

- `Table Cleaner`: keep `Remove empty rows`, `Remove empty columns`, `Normalize cell padding`, `Strip direct formatting`, `Normalize border styles`, and `Remove all borders` together. These all operate on table structure or presentation and make sense as table cleanup.
- `Punctuation Normalizer`: keep quotes, dashes, ellipses, and custom punctuation choices together. They share the same mental model.
- `Unicode Cleanup`: keep NBSP, zero-width characters, BOM, soft hyphen, and non-breaking hyphen together. The tool is coherent.
- `Spacing Fixer`: keep double spaces, trim spaces, punctuation spacing, and blank-line cleanup together. These are all text spacing hygiene.
- `Capitalization Fixer`: keep sentence/title/upper/lower/smart sentence options together, but present them as capitalization modes rather than a busy control cluster.
- `Formatting Stripper`: keep emphasis-preservation modes and preserve-highlight/drop-cap options together because they are all part of the same formatting reset decision.

### Keep inside the current tool, but make the risk visually obvious

These options belong near the current parent tool, but the redesigned UI should elevate the consequence.

- `Table Cleaner`: `Convert single-column tables to plain text` is a conversion action, not simple cleaning. Keep it in the table tool for now, but put it in a distinct `Convert` panel and require the preview/control-panel flow before applying.
- `Header / Footer Standardizer`: `Clear all header/footer content` is destructive. Keep it in the same tool as an explicit mode, but visually separate it from `Standardize formatting`.
- `Duplicate Paragraph Detector`: `Remove duplicate paragraphs` should become a post-preview action, not a first-step option. The tool should detect and show duplicates first; the compact preview panel can then offer `Remove duplicates`.
- `Style Cleanup`: `Remove unused styles` and `Remap style variants` can stay together, but the UI should explain that one deletes unused custom styles and the other changes text style assignments before cleanup.

### Consider splitting into separate tools later

These sub-tools are conceptually strong enough that they may deserve their own tool, especially after the launcher redesign makes the suite easier to browse.

- `Metadata Scrubber`: split the current tool into `Privacy Scrubber` and `Review Markup Finalizer`.
- `Privacy Scrubber`: document properties and personal information.
- `Review Markup Finalizer`: comments and tracked changes.
- `Object Remover`: move `Hidden text` out of object removal and into the privacy/review family. Hidden text is document content/privacy, not an embedded object.
- `Object Remover`: move `Tables -- DELETES tables and all their contents` out of the general object-removal list. If table deletion remains available, it should become a dedicated `Delete Tables` action near table/layout tools, with much stronger warnings.
- `Footnote / Endnote Remover`: consider renaming it to `Footnote / Endnote Tools` if it keeps both delete and inline-convert behavior. `Keep note text inline` is a conversion path, not just a removal option.

### Recommended near-term decision

For the launcher redesign, do not create new tools yet.

Instead:

- keep the top-level tool list stable
- improve categories and descriptions
- mark risky sub-actions in the spec so the later form redesign handles them correctly
- use the compact preview/control panel to make destructive follow-up actions feel deliberate

This gives the launcher a cleaner structure now while avoiding a premature split of the VBA code.

## Action Model After Preview

The current preview and run model is awkward if the main form remains open and in the way while the user examines the document.

That awkwardness becomes more visible as the forms become more polished, because the user expects the interaction model to be polished too.

The redesign should introduce a second-stage action pattern:

- the main form gathers decisions
- preview runs and then gets out of the way
- a smaller control surface remains available while the user inspects the document

This is the right place to use a compact control panel.

### Recommended interaction flow

1. User configures the tool in the main form
2. User clicks `Preview`
3. The main form closes or minimizes out of the way
4. A compact floating control panel appears
5. The user reviews the document and then decides what to do next

### Compact control panel actions

The small panel should contain only the actions that matter after preview.

Recommended controls:

- `Undo`
- `Apply` or `Do It`
- `Close Preview`
- `Stop` when a tool may run long or iterate heavily
- optional `Back to Options` when returning to the full form is useful

Naming guidance:

- use `Apply` for the general case
- use `Do It` only if the suite intentionally adopts a more conversational tone everywhere
- use `Close Preview` instead of a vague `Cancel`
- use `Stop` only for tools that genuinely need interruption support

### Which tools benefit most from the compact panel

This pattern is most valuable for:

- `Duplicate Paragraph Detector`
- `Formatting Stripper`
- `Object Remover`
- `Unicode Cleanup`
- `Paragraph Structure Fixer`
- `Table Cleaner`

It is less necessary for very simple tools, but using the same pattern selectively is still fine if the compact panel remains minimal.

### Why this improves usability

- the document becomes the main thing on screen during review
- the suite stops blocking the user's view
- preview feels like a real review state rather than a checkbox mode
- the user gets an immediate path to undo, stop, or apply without reopening the full form

### VBA realism

Inside Word/VBA, the compact panel should be simple and robust.

It should probably be:

- a smaller secondary UserForm
- modeless where reliable
- visually reduced to a narrow action strip or small floating panel
- consistent across tools, even if not every button is enabled for every tool

This should not try to become a full inspector pane.

The implementation should favor:

- consistency
- low obstruction
- very clear verbs

over visual novelty.

## Modernization Without Breaking Trust

The forms should feel modern, but not playful or overdesigned.

Visual tone should emphasize:

- clean spacing
- stronger headings
- restrained color accents
- clearly separated panels
- better action emphasis

Avoid:

- decorative gradients that feel fake in VBA
- overly round card-heavy design
- novelty layouts that reduce clarity
- hiding important options behind too much interaction

The ideal result is "modern Office-adjacent utility tool" rather than "fake web app inside Word."

## Implementation Strategy

This design should be implemented in stages.

### Stage 1: system pass

Define the shared form language:

- sizing
- title treatment
- section spacing
- button order
- panel patterns
- preview/action treatment

### Stage 2: launcher redesign

Update the suite launcher first so the whole product feels changed immediately.

This stage should include:

- new category structure
- compact tool-card or tool-row presentation
- integrated help access
- category ordering tuned to the actual tool set

### Stage 3: post-preview control panel

Build the compact review/action panel pattern before redesigning every individual tool.

This creates the suite-wide interaction model for:

- preview review
- apply after preview
- undo
- stop for long-running tools

Once this exists, the individual forms can be designed around it instead of continuing to treat preview as a buried checkbox.

### Stage 4: complex tools

Apply the full `C` treatment to:

- `frmUnicodeCleanup`
- `frmDuplicateDetector`
- `frmFormattingStripper`
- `frmObjectRemover`
- `frmHeaderFooterStandardizer`
- `frmTableCleaner`

### Stage 5: simple tools

Apply the lighter version to the simpler forms.

### Stage 6: polish pass

Normalize:

- help wording
- caption consistency
- focus order
- default selections
- button naming

## Testing Expectations

Redesign work should be tested for:

- control visibility and overlap
- text fit at Word form scale
- tab order and keyboard flow
- selection-enabled versus selection-disabled scope states
- preview clarity
- consistency of action placement across forms

Because this is VBA UI work, visual regression is mostly manual. That makes consistency rules especially important.

## Recommendation Summary

Use the `C` vibe as the suite-wide direction.

Do not apply it as pure styling only.

The real value comes from combining:

- the visual clarity of `C`
- the decision flow of `B`
- the restraint of `A`

That combination gives CleanupSuite the best chance of feeling modern, useful, and trustworthy inside the constraints of Word VBA.
