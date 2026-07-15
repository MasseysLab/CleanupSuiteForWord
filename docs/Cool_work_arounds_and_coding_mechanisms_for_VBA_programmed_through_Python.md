# Cool work arounds and coding mechanisms for VBA programmed through Python.

This document records the special techniques, constraints, and project-specific workarounds used in CleanupSuite for MS Word. It is meant for future Codex, ChatGPT, and human maintainers who need to understand why the project is shaped this way.

## 1. Python builds a pasteable VBA installer

CleanupSuite is maintained as normal source files under `src/`, but distributed as one large pasteable VBA installer file:

- `src/manifest.txt` controls assembly order.
- `assemble.py` reads the manifest and emits `VBA_Cleanup_tool.txt`.
- `VERBATIM` manifest entries are copied as ordinary VBA.
- `BUILDER` entries are converted into VBA string-builder functions.
- `XML_BUILDER` entries do the same for Ribbon XML.
- The final output is CRLF-normalized because VBA import and copy/paste workflows are sensitive to line endings.

The unusual part is that plain form/module source is turned into functions such as `frmCapitalizationCleanup_Code()` and `modCleanupHelpers_Code()`. The installer then injects those strings into Word's VBProject.

Why it matters:

- Source remains readable and diffable.
- The user can still install from a single `.txt` bundle.
- Generated output can be validated before it is written.
- The same source can be synced into both working `.docm` files.

## 2. The installer creates and styles MSForms at runtime

The installer does more than insert code. It programmatically creates UserForms and controls:

- `CreateOrReplaceModule` replaces standard modules.
- `CreateOrReplaceForm` creates UserForms and injects their code.
- `AttemptCreateControls` creates the form controls.
- `ControlsForForm` is the per-form control manifest.
- `ControlTypeByName` maps naming conventions to MSForms control classes.
- `ApplyFormStyle` sets shared form styling.
- `ApplyControlStyle` sets shared control styling.
- Form-specific layout routines such as `LayoutLauncherControls`, `LayoutPreviewActionsControls`, and `LayoutCapitalizationControls` position controls after creation.

This is a workaround for the fact that generated `.frm` designer files are not the main source of truth here. Instead, the installer rebuilds the forms from code, which makes the suite easier to ship as one macro installer.

## 3. Build-gate validators prevent fragile VBA output

`assemble.py` runs `vbaeval.py` before writing the distributable bundle. The validators catch problems that are easy to miss in generated VBA:

- unbalanced blocks
- unsafe single-line `If` traps
- broken form/module wiring
- malformed string-builder output
- round-trip conversion failures

The most important project rule is documented in the installer header: never put block constructs inside a single-line `If`. VBA can accept some single-line syntax that later breaks badly once generated into string builders.

## 4. Forms need explicit OptionButton groups

MSForms frames do not reliably group option buttons when the controls are created flat by installer code. CleanupSuite sets `GroupName` explicitly in `UserForm_Initialize`.

Examples:

- `CapMode` and `CapScope`
- `PunctMode` and `PunctScope`
- `DupAction`, `DupMatch`, and `DupThreshold`

This is why new option buttons must be wired in code, not only placed near each other visually.

## 5. Two methods for removing double titles

The project now has a shared title strategy in `modCleanupHelpers`:

- `ApplyMSFormTitleStrategy`
- `HideLegacyTitleControls`
- `IsLegacyTitleControlName`

There are two valid methods.

Method 1: blank the native UserForm caption when the form body already owns the title.

Use this when the form has a clear in-body title, heading, or compact action identity. Current examples:

- `frmCleanupSuiteLauncher`
- `frmPreviewActions`
- `frmCapitalizationCleanup`
- all generic cleanup tool submenus via `LayoutCleanupToolForm`

The installer mirrors this with `FormBodyOwnsTitle`, which blanks the design-time caption for those forms during generated installs.

Method 2: keep the native UserForm caption and hide legacy in-body title labels.

Use this as a fallback if a future form cannot own its title in the body, but duplicate `lblTitle` controls should not appear. The helper remains because it is a useful escape hatch for older or unusual MSForms behavior.

Why both methods are recorded:

- Some modernized forms feel cleaner when the body owns the title and the native caption is blank.
- Some older forms still benefit from a normal title bar.
- If a future MSForms/Word behavior changes, either method can be used without rediscovering the problem.

## 6. Help windows do not share the same title problem

The Help surfaces are usually native `MsgBox` dialogs. A `MsgBox` has one native title and a body, but it is not the same custom MSForms UserForm surface. That is why the Help window did not show the same duplicate-title behavior as the Capitalization form.

## 7. Preview Actions is a compact modeless action bar

`frmPreviewActions` is intentionally different from the older control-panel style forms. It works as a compact modeless action bar:

- it hides the originating tool form
- it shows preview state and action buttons
- it allows the document to remain visible
- it can re-open the source form for reconfiguration
- it can apply the selected cleanup from the panel

The shared helpers keep the panel usable:

- `ShowPreviewActions` owns creation and fallback behavior.
- `RememberPreviewActionPanelPosition` stores the user's moved location.
- `ApplyPreviewActionPanelPosition` restores that location for the next panel instance.
- `DocumentHasVisibleContent` prevents applying cleanup to a blank document.

The position-memory workaround exists because toggling preview can recreate or re-show the panel. Without storing `Left` and `Top`, Word/MSForms can snap the window back to its original startup position.

## 8. Preview OFF uses editing language, not stale findings

When preview is OFF, the action bar should not keep reporting the old preview result. It now says:

`You may now edit your document if you wish.`

This is a small UI rule with a big usability effect. Preview findings belong to the highlighted-preview state. Editing guidance belongs to the off state.

## 9. Capitalization custom mode avoids frame layering problems

The first guided Capitalization form exposed a useful MSForms limitation: controls inside or near a frame can be hidden or layered unexpectedly when created and repositioned programmatically.

The final workaround:

- keep `fraCustom` as a generated compatibility/control placeholder
- hide it at runtime
- position the custom checkboxes directly on the form
- bring the real checkboxes and buttons forward with `ZOrder 0`

This avoids depending on the frame as a visible container while preserving installer/source compatibility.

## 10. Generic submenus share a guided runtime layout

The non-Capitalization cleanup forms are intentionally not hand-tuned one by one. They all call `LayoutCleanupToolForm Me`, which now does the common modern submenu work:

- blanks the native title with `ApplyMSFormTitleStrategy toolForm, True`
- creates runtime guidance labels and uses an intro-first guided layout
- maps form names to friendly titles and short instructions
- hides old visible frame boxes while preserving their controls and grouping compatibility
- styles labels, options, checkboxes, and command buttons consistently
- places scope controls above a full-width Preview button
- places Apply, Reset, and Help in one compact bottom row

The matching installer routine is `LayoutGenericToolControls`. It reserves top space for the runtime title and intro, hides frame placeholders, and generates the same Preview / Apply / Reset / Help action pattern. This prevents generated installs from drifting away from the live runtime layout.

The generic layout is now a **two-column guided layout**. Normal `opt` and `chk` choices are paired left/right as they appear in `ControlsForForm`, and the same rule applies after a `Custom` choice reveals its detail controls. If a row ends with only one remaining choice, single leftover choices are centered instead of being left-anchored. This is important because the approved Capitalization form established a compact two-column reading pattern; future tools should follow that pattern unless they have a deliberately custom layout helper.

The key design decision is centralization: the individual tools keep their existing cleanup behavior, while shared form structure lives in one helper. That gives the suite a consistent interface without rewriting every cleanup algorithm.

## 11. Global settings are stored in document custom properties

CleanupSuite uses Word `CustomDocumentProperties` for global settings:

- `CleanupSuiteAutoSave`
- `CleanupSuiteReturnToMainAfterApply`
- `CleanupSuiteReturnToMainAfterClose`
- `CleanupSuiteInProgress`

This keeps settings with the active document rather than relying on a separate config file. The helper pattern deletes and recreates the property so values are normalized to `"True"` or `"False"`.

## 12. Safety guards protect user documents

Before running cleanup, `GuardBeforeCleanup` checks the document state:

- unsaved documents prompt the user to save
- auto-save can save silently when enabled
- macro-host documents are protected from accidental self-save behavior

`MarkCleanupStart` and `MarkCleanupEnd` use a custom property so the document can record that a tool was in progress. This is useful for recovery thinking and future diagnostics.

## 13. The `.docm` sync process needs a path workaround

Word COM automation has been unreliable with the long project path that contains spaces. The successful sync pattern is:

1. Copy the target `.docm` to a short temporary path.
2. Open the temporary copy through Word COM.
3. Replace the relevant VBA modules/forms.
4. Save and close the temporary copy.
5. Copy the temporary file back over the real `.docm`.
6. Verify by reading the embedded VBA back from the saved document.

This avoids hangs seen when automating the real long-path project files directly.

## 14. `.docm` state must be verified separately from source and bundle state

CleanupSuite has four states that can drift:

- local source files under `src/`
- generated `VBA_Cleanup_tool.txt`
- practice generated bundle
- embedded VBA inside both `.docm` files

Passing source tests is not enough if Chris is testing a stale `.docm`. Release and UI work must verify source, generated bundle, practice bundle, and embedded `.docm` state.

## 15. Runtime audits use temporary VBA modules

When visual behavior matters, the project has used temporary audit code inserted into a temporary synced `.docm` copy. The audit can instantiate forms, read captions, inspect control visibility, and confirm layout metrics.

Important cleanup rule:

- remove temporary `modCodex*` modules after the audit
- verify no temporary modules remain in the real `.docm` files
- close Word automation sessions and clear lock files when necessary

This gives evidence from the real Word/VBA runtime without permanently adding audit code to the project.

## 16. Source tests act as contract tests for generated VBA

Many Python tests read VBA source and installer text directly. They are not traditional unit tests of Word behavior. They are contract tests that ensure important strings, wiring, and layout calls remain present.

Examples:

- launcher category and setting wiring
- preview action panel dimensions and position memory
- capitalization guided-form controls and help text
- installer control manifests
- expected version metadata

This style is useful because much of the project is code generation and Word automation, where normal isolated unit tests are hard.

## 17. Main Menu layout is generated, not hand-drawn

The Main Menu is not a static designer artifact. The installer lays it out with row and column helper routines:

- category labels
- help buttons
- tool buttons
- descriptions
- global settings
- reset button

This lets the project adjust the menu consistently without editing a fragile `.frm` designer file by hand.

## 17. Reinstall is intentionally split into remove and install phases

The installer includes `UninstallCleanupSuite` and `ReinstallCleanupSuite`. The comments explain the key VBA limitation: removing and recreating components in the same run can create duplicate forms or fail because the VBA editor defers removal.

The safer model is:

- remove suite components
- let that macro run finish
- run install separately

`ReinstallCleanupSuite` uses `Application.OnTime` as a convenience, but direct install has sometimes been more reliable during automation.

## 18. Ribbon XML is injected through a macro-enabled package edit

The suite includes Ribbon customization logic that writes CustomUI XML into `.docm` or `.dotm` files. This is another reason the project has Python/VBA split responsibilities:

- Python assembles and validates source.
- VBA installs modules/forms and can write package parts where Word expects them.

Macro-enabled files are required because ordinary `.docx` files cannot hold VBA.

## 19. Practical future guidance

For future work:

- Prefer shared helpers over one-form fixes.
- Add source-contract tests before changing generated behavior.
- Rebuild with `assemble.py` after source changes.
- Sync both generated text bundles.
- Sync both `.docm` files when runtime behavior matters.
- Verify embedded `.docm` contents by read-back.
- Use runtime audits for MSForms layout issues.
- Record every unusual workaround in `AI_HANDOFF.md` and this document when it becomes durable project knowledge.
