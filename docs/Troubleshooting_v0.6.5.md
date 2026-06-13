# CleanupSuite Troubleshooting Guide

**Documentation release target:** 0.6.5  
**Codebase documented:** 0.6.4 code-freeze candidate  
**Commit documented:** `c787a70c8b78a6150a2fe1a94b3760a8fde3e6cc`

> Documentation basis: written from the pushed GitHub `0.6.4` code-freeze candidate, commit `c787a70c8b78a6150a2fe1a94b3760a8fde3e6cc`. The planned `0.6.5` release should use the same codebase plus this documentation, except for documentation-blocking bug fixes if required.


## Installer says controls or forms could not be created

Enable **Trust access to the VBA project object model** in Word Trust Center settings. Also check that the VBA Extensibility reference is available when doing developer installation.

## The launcher or a tool form looks stale

Close Word completely and reopen the macro-enabled document/template. Reinstall from the current `VBA_Cleanup_tool.txt` if the embedded `.docm` is out of sync with the source.

## Preview highlighting remains in the document

Use **Reconfigure** or Reset All when available. You can also clear yellow highlighting manually in Word. Save a copy first if the document matters.

## Apply does nothing

Possible causes:

- The preview found no matching content.
- The current selection does not contain matches.
- The document is blank.
- The document is protected.
- The tool settings do not target the issue you expected.

Try whole-document scope, use preview, and confirm the document is not protected.

## Word crashes or hangs

1. Do not overwrite your only copy.
2. Reopen Word.
3. Check AutoRecover if needed.
4. Use the latest saved backup or Ctrl+Z if available.
5. Record the tool, settings, document type, Word version, Windows scaling, and screenshot.

## Preview Actions bar moves unexpectedly

In 0.6.4, the preview panel is designed to remember its dragged position while toggling preview off and back on during the same session. If it still jumps, record the exact steps and whether you clicked **Preview is ON/OFF**, **Reconfigure**, or closed the bar.

## Button or label clipping

Record screen resolution and Windows scaling. The 0.6.4 compact Preview Actions bar is deliberately small, so visual problems should be checked in live Word, not only through automated screenshots.

## Developer sync problems

Compare these items when behavior in Word does not match source:

- `src/forms/*.bas`
- `src/modules/*.bas`
- `src/installer/installer.bas`
- `VBA_Cleanup_tool.txt`
- `Practice - Try CleanupSuite Here/VBA_Cleanup_tool.txt`
- the embedded `.docm` files

Then run:

```text
python assemble.py
python -m unittest discover -s tests
```
