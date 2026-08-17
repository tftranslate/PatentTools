# Building PatentTools

This project builds a macro-enabled Word template (`.dotm`) from exported VBA sources and RibbonX resources.

## Prerequisites

- Microsoft Word for Windows.
- Windows PowerShell 5.1 (`powershell.exe`).
- The project files checked out locally.

### Enable VBA-project access

The build script imports VBA components through Word's automation interface. In Word, enable this setting for the Windows user running the build:

1. Open **Word**.
2. Go to **File → Options → Trust Center → Trust Center Settings**.
3. Open **Macro Settings**.
4. Enable **Trust access to the VBA project object model**.
5. Close Word completely before building.

Only enable this option on a trusted development machine.

## Build

1. Ensure that all Word windows and VBA Editor windows are closed.
2. Double-click (or open a CMD terminal and execute):

   ```text
   Build-PatentTools.cmd
   ```

3. Wait for the command window to report `BUILD SUCESSFUL`.

The generated template is written to:

```text
build\PatentTools.dotm
```

## What the build does

- Copies `template/PatentTools.base.dotm` to the build output.
- Imports VBA modules and UserForms from `source/vba/`.
- Embeds the RibbonX XML, image resources, and package metadata from `source/ribbon/` and `source/package/`.

Do not open the generated `.dotm` in Word while a build is running.