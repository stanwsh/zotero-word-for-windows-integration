# UserForm Implementation for Go To Zotero Feature

The "Go To Zotero" feature is designed to work with two selection methods:

1. **UserForm Selection (Preferred)**: A graphical multi-select list that allows users to easily select which Zotero items to open.
2. **Simple Text Selection (Fallback)**: A text-based InputBox that asks users to enter numbers of items to open.

## Current Status

The code is set up to try to use a UserForm called `frmZoteroSelectItems` for item selection. If this form is not available, it automatically falls back to the text-based selection method.

## Creating the UserForm

To implement the full UserForm functionality, a form named `frmZoteroSelectItems` needs to be created in the VBA project with the following elements:

1. **Form Properties**:
   - Name: `frmZoteroSelectItems`
   - Caption: (Will be set by code)
   - A property/variable named `DialogResult` to store the result (vbOK or vbCancel)

2. **Controls**:
   - A ListBox named `lstItems` for displaying and selecting items
   - "OK" and "Cancel" buttons
   - Optional: "Select All" and "Clear All" buttons for convenience

3. **Button Code**:
   - OK button should set `DialogResult = vbOK` and then `Me.Hide`
   - Cancel button should set `DialogResult = vbCancel` and then `Me.Hide`
   - Select All button should select all items in the list
   - Clear All button should deselect all items in the list

## How to Create the UserForm

To create this UserForm:

1. Open the Zotero.dotm template in Word
2. Press Alt+F11 to open the VBA Editor
3. Right-click on the project in the Project Explorer
4. Choose "Insert" -> "UserForm"
5. In the Properties panel, set the UserForm's name to `frmZoteroSelectItems`
6. Add the required controls to the form
7. Double-click each button to add the necessary code
8. Add a public variable `Public DialogResult As VbMsgBoxResult` to the form

## Sample Button Code

For the OK button:

```vb
Private Sub btnOK_Click()
    DialogResult = vbOK
    Me.Hide
End Sub
```

For the Cancel button:

```vb
Private Sub btnCancel_Click()
    DialogResult = vbCancel
    Me.Hide
End Sub
```

For the Select All button:

```vb
Private Sub btnSelectAll_Click()
    Dim i As Long
    For i = 0 To lstItems.ListCount - 1
        lstItems.Selected(i) = True
    Next i
End Sub
```

For the Clear All button:

```vb
Private Sub btnClearAll_Click()
    Dim i As Long
    For i = 0 To lstItems.ListCount - 1
        lstItems.Selected(i) = False
    Next i
End Sub
```

## Note

Even without this UserForm, the feature will work using the text-based selection method. The UserForm simply provides a more user-friendly interface.
