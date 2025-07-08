# Zotero Word Integration

This repository provides advanced Microsoft Word integration for Zotero, focusing exclusively on the VBA macro code and Word template (dotm) components. It enables citation, bibliography, and reference management directly within Word, with a modern ribbon UI and custom macros.

## Cross-Platform Support

This release includes separate templates optimized for:
- **Windows**: Full-featured with UserForm dialog for multi-reference selection
- **macOS**: Compatible version with text-based reference selection

## Legal & Trademark Notice

Zotero and the Zotero logo are registered trademarks of the Corporation for Digital Scholarship. Use of the name and logo in this project is solely for the purpose of indicating compatibility. This project is not affiliated with or endorsed by Zotero or the Corporation for Digital Scholarship.

All original icons or modified graphics in this project that reference Zotero use the official Zotero icon **unmodified**, in accordance with Zotero's [trademark policy](https://www.zotero.org/support/trademark_policy), and are used strictly for descriptive, nominative purposes.

## Features

- **Go To Zotero**: Instantly navigate from a Word citation to the corresponding Zotero item(s) in your Zotero library, using the `zotero://` protocol. Supports both single and multi-reference citations with a user-friendly selection dialog.
- **Modern Ribbon UI**: Redesigned ribbon layout and icons, with groups for Citations, Bibliography, and Tools, inspired by EndNote for a familiar experience.
- **Multi-reference Support**: When a citation contains multiple references, a selection interface allows you to open one or more items in Zotero.
- **UserForm Selection Dialog**: A modern, resizable dialog for selecting items, with Select All and Clear All options. See the build guide below. (Windows only)
- **Compatibility**: Built and tested on Word 2010 and later.
- Original Features:
  - **Add/Edit Citations**: Insert or edit citations at the current cursor position in Word.
  - **Add/Edit Bibliography**: Insert or update a bibliography in your document.
  - **Insert Note**: Add notes linked to your references.
  - **Unlink Citations**: Remove all Zotero field codes and unlink from the Zotero library.
  - **Document Preferences**: Change citation style or locale.
  - **Refresh Citations**: Update all citations to reflect changes in your Zotero library.

## Installation

### Step 1: Choose the Right Template
- For **Windows**: Use `Zotero.dotm` from the `install` folder
- For **macOS**: Use `Zotero.dotm` from the `install/mac` folder

### Step 2: Install the Template

Replace the existing `Zotero.dotm` file in your Word startup folder with the appropriate version from this repository.

**Windows Installation:**
1. Close Word completely
2. Navigate to `%APPDATA%\Microsoft\Word\STARTUP`
3. Backup the existing `Zotero.dotm` if present
4. Copy the new `Zotero.dotm` file from this repository to the STARTUP folder
5. Start Word and verify that the Zotero tab appears with the new "Go To Zotero" button

**macOS Installation:**
1. Close Word completely
2. Navigate to `~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word`
   - If you can't find this path, open Finder, press Cmd+Shift+G, and paste the path
3. Backup the existing `Zotero.dotm` if present
4. Copy the `Zotero.dotm` file from the `install/mac` folder in this repository to the Word Startup folder
5. Start Word and verify that the Zotero tab appears with the new "Go To Zotero" button

### Step 3: Test the Installation

1. Create a new document in Word
2. Add a Zotero citation using the Add/Edit Citation button
3. Place your cursor within the citation
4. Click the "Go To Zotero" button in the ribbon
5. Verify that Zotero opens and displays the cited item(s)

## Build and Test Environment

- The Windows template was compiled under **Microsoft Office Word 2010 14.0.4760.1000 (32-bit)**
- The macOS template was tested on **Word for Mac 16.78.23100802**
- Tested on:
  - Windows XP (32-bit)
    - Compile environment (Word 2010 32-bit)
    - Zotero 5.0.77
  - Windows 10/11 (64-bit)
    - Word for Microsoft 365 MSO 2506 Build 16.0.18925.20076 (64-bit)
    - Zotero 7.0.16 (Latest)
  - macOS Sonoma (14.6.1)
    - Microsoft Word for Mac 16.78.23100802
    - Zotero 7.0.18 (Latest)

## Known Issues

- On macOS, the multi-reference selection uses a text-based interface instead of a graphical UserForm dialog
- The multi-reference selection dialog may only display items that belong to the first field code in the document when multiple field codes are selected
- The `Go To Zotero` feature may not work correctly if the Zotero item has been deleted or moved

## Upstream C++ Source Reference

This repository includes VBA code modifications only. However, the Microsoft Word add-ins are designed to work with the original C++ integration code, which is responsible for the communication between Zotero and Word.

The original C++ integration code (not maintained here) is available from the [upstream Zotero repository](https://github.com/zotero/zotero-word-for-windows-integration).  
If you require the C++ component (such as `zoteroWinWordIntegration.cpp`), please refer to the [original source file](https://github.com/zotero/zotero-word-for-windows-integration/blob/master/build/zoteroWinWordIntegration/zoteroWinWordIntegration.cpp).

**Note:**  
This repository does not track or update the C++ integration code. For the latest version, bug fixes, or to build from source, always consult the official Zotero repository.

If you copy any files from the upstream project, please ensure you respect the original licensing and include a reference to the original commit/version.

## Template Build Requirements

- Templates should be built with the oldest version of Word to be supported. Otherwise, older versions of Word may fail to function properly. This is currently:
  - Word 2007 (for the ribbonized dotm template)
  - Word 2003 (for the old dot template)

## To Modify/Build the Templates

- Open the template from inside Microsoft Word
- Go to View -> Macros -> View Macros (Ribbonized Word) or Tools -> Macros -> View Macros (Word 2003) and click "Edit" for one of the Zotero macros
- Edit/replace code as desired
- Go to Debug -> Compile Project to ensure there are no code errors

## To Unpack the Template

- Run `build/template/unpack_templates.sh`
- Prerequisites: `libxml2`, `unzip`

## UserForm Build Guide

To enable the multi-reference selection dialog (Windows only), follow the instructions in [`build/template/Zotero.dotm/word/UserForm_Creation_Guide.md`](build/template/Zotero.dotm/word/UserForm_Creation_Guide.md) to create and configure the required UserForm in the template.

## Development Starter's Guide

Start by opening the dotm/dot template in Word. Word templates support custom macros and UI elements to call the macros, which is how the extension is implemented. The Ribbon UI can be edited by extracting the dotm file or using the [Custom UI editor](http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2009/08/06/7293.aspx). To edit the .dot template UI, Word for Windows 2003 is needed (no longer supported).

In VBA macro code, [SendMessage](https://msdn.microsoft.com/en-us/library/windows/desktop/ms644950(v=vs.85).aspx) is used to issue commands to the Zotero process from Word. These commands are received in [zotero-service.js](https://github.com/zotero/zotero/blob/eaf8d3696359dcea0edaa2fd9bc1e4cf5d985014/components/zotero-service.js#L516-L516) and passed to integration.js.

Zotero talks to Word via [js-ctype bindings](https://github.com/zotero/zotero-word-for-windows-integration/blob/4f07be4bfaa3f37897a5af5371ea20353214f23e/components/zoteroWinWordIntegration.js#L52-L52) to a C++ OLE Automation based [library](https://github.com/zotero/zotero-word-for-windows-integration/blob/8d1807584d02f3b10715dd9895413c04349d45e8/build/zoteroWinWordIntegration/zoteroWinWordIntegration.h). To generate new interfaces for Word interop communications, use the Add New Class wizard in Visual Studio and select 'MFC Class from Typelib'. The interop API docs can be found in the [MSDN](https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word._document?view=word-pia).

The plugin should technically work with Word versions starting with 2003, but support for versions below Word 2010 has ended due to unfixable bugs and Microsoft's own compatibility changes. Some API calls are on a deprecation path, so future changes may require splitting the library into multiple DLLs.

## Credits

- Original Zotero Word for Windows Integration by the Zotero team and contributors
- Major enhancements, ribbon redesign, and Go To Zotero feature by stanwsh (2025)
- Modern icon set and UI improvements by stanwsh (2025)

**Note:** The Zotero icon is © Corporation for Digital Scholarship and used here under the terms of the AGPL and Zotero's trademark policy for nominative, unmodified use. All other icons and graphics in this project are licensed under the terms stated in the `NOTICE.txt`.

---
