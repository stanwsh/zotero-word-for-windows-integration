# How to Create the Zotero Item Selection Form

This document provides detailed instructions for creating the "Select Zotero Items" UserForm used in the Go To Zotero feature. This form allows users to select which Zotero items they want to open when clicking on a citation that contains multiple references.

## Disclaimer

Zotero and the Zotero logo are registered trademarks of the Corporation for Digital Scholarship. Use of the name and logo here does not imply endorsement.

## Creating the UserForm

1. Open Word and load the Zotero template (Zotero.dotm)
2. Press **Alt+F11** to open the Visual Basic Editor
3. In the Project Explorer (top-left pane), right-click on the project name
4. Select **Insert > UserForm**
5. In the Properties window (usually bottom-left):
   - Change the UserForm's **Name** property to: `frmZoteroSelectItems`
   - Change the **Caption** property to: `Select Zotero Items to Open`
   - Set the **Width** to approximately: `400`
   - Set the **Height** to approximately: `300`

## Adding Controls to the Form

Add the following controls to the form from the Toolbox (if the Toolbox isn't visible, select View > Toolbox):

### 1. ListBox

- Draw a large ListBox on the form
- Set its properties:
  - **Name**: `lstItems`
  - **MultiSelect**: `1 - fmMultiSelectMulti`
  - Position it to take up most of the form space

### 2. OK Button

- Add a CommandButton
- Set its properties:
  - **Name**: `btnOK`
  - **Caption**: `OK`
  - Position it in the bottom-right area of the form

### 3. Cancel Button

- Add a CommandButton
- Set its properties:
  - **Name**: `btnCancel`
  - **Caption**: `Cancel`
  - Position it next to the OK button

### 4. Select All Button

- Add a CommandButton
- Set its properties:
  - **Name**: `btnSelectAll`
  - **Caption**: `Select All`
  - Position it in the bottom area of the form

### 5. Clear All Button

- Add a CommandButton
- Set its properties:
  - **Name**: `btnClearAll`
  - **Caption**: `Clear All`
  - Position it next to the Select All button

## Adding Code to the UserForm

1. Double-click on the UserForm itself (or right-click and select "View Code")
2. Copy and paste the code from the `UserForm_Implementation_Instructions.bas` file into the code window
3. The code includes:
   - Dialog result handling
   - Button event handlers
   - Form positioning logic
   - Form initialization

## Testing the Form

1. Save the project
2. Test the "Go To Zotero" feature with a citation that contains multiple references
3. The form should appear showing all items in the citation
4. Test the Select All and Clear All buttons
5. Test selection behavior (clicking OK should open selected items, Cancel should do nothing)

## Troubleshooting

If the form doesn't appear:

1. Make sure the form is named exactly `frmZoteroSelectItems`
2. Check if all controls are named correctly (especially `lstItems`)
3. Verify the code was pasted into the form's code window
4. Check for any error messages in the VBA Editor (Debug > Compile Project)

If the form appears but doesn't function correctly:

1. Check that all button events are properly implemented
2. Verify that the `DialogResult` property is set correctly
3. Test in a new Word document

## Notes

- The form is designed to be resizable, and controls should reposition appropriately
- All items are selected by default for user convenience
- If no form is available, the code will fall back to a text-based selection interface
