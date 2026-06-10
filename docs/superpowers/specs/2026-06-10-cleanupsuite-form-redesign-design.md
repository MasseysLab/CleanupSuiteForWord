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

To make the suite more fluid conceptually, group tools visually into categories on the launcher:

- Text cleanup
- Structure cleanup
- Formatting cleanup
- Document cleanup
- Metadata and objects

This makes the suite easier to learn and reduces the feeling that every tool is unrelated.

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

### Stage 3: complex tools

Apply the full `C` treatment to:

- `frmUnicodeCleanup`
- `frmDuplicateDetector`
- `frmFormattingStripper`
- `frmObjectRemover`
- `frmHeaderFooterStandardizer`
- `frmTableCleaner`

### Stage 4: simple tools

Apply the lighter version to the simpler forms.

### Stage 5: polish pass

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
