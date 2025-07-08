Attribute VB_Name = "frmZoteroSelectItems"
Option Explicit
' ***** BEGIN LICENSE BLOCK *****
'
' Copyright (c) 2025 stanwsh
' SPDX-License-Identifier: AGPL-3.0-only
'
' ***** END LICENSE BLOCK *****

' To create a UserForm:
' 1. In the VBA Editor (Alt+F11), select Insert > UserForm
' 2. Rename the UserForm to "frmZoteroSelectItems" in the Properties window
' 3. Add the following controls to the form:
'    - ListBox named "lstItems" (set MultiSelect property to 1 - fmMultiSelectMulti)
'    - CommandButton named "btnOK" with Caption "OK"
'    - CommandButton named "btnCancel" with Caption "Cancel"
'    - CommandButton named "btnSelectAll" with Caption "Select All"
'    - CommandButton named "btnClearAll" with Caption "Clear All"
' 4. Copy and paste this entire code into the UserForm's code window

' Private variable to store dialog result

Private m_DialogResult As VbMsgBoxResult

' Property to get/set dialog result
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = m_DialogResult
End Property

Public Property Let DialogResult(value As VbMsgBoxResult)
    m_DialogResult = value
End Property


' Form initialization
Private Sub UserForm_Initialize()
    ' Set default dialog result to Cancel
    DialogResult = vbCancel
    
    ' Position buttons appropriately
    PositionControls
    
    ' Set Fonts
    SetUserFormFont Me
    
    ' Center the form on screen
    Me.StartUpPosition = 0 ' Manual
    CenterForm
End Sub

' Position controls based on form size
Private Sub PositionControls()
    ' Set list box to fill most of the form
    With lstItems
        .Left = 5
        .Top = 5
        .Width = Me.InsideWidth - 20 + 10
        .Height = Me.InsideHeight - 80
        .IntegralHeight = False  ' Allow partial rows
    End With
    
    ' Position buttons at bottom
    Dim buttonTop As Long
    buttonTop = Me.InsideHeight - 40
    
    btnOK.Top = buttonTop
    btnCancel.Top = buttonTop
    btnSelectAll.Top = buttonTop
    btnClearAll.Top = buttonTop
    
    ' Space buttons evenly
    Dim buttonWidth As Long, spacing As Long
    buttonWidth = 60
    spacing = (Me.InsideWidth - (4 * buttonWidth)) / 5
    
    btnOK.Left = spacing
    btnSelectAll.Left = (2 * spacing) + buttonWidth
    btnClearAll.Left = (3 * spacing) + (2 * buttonWidth)
    btnCancel.Left = (4 * spacing) + (3 * buttonWidth)
    
    ' Set button widths
    btnOK.Width = buttonWidth
    btnCancel.Width = buttonWidth
    btnSelectAll.Width = buttonWidth
    btnClearAll.Width = buttonWidth
End Sub

' Center form on screen
Private Sub CenterForm()
    Me.Left = (Application.Width - Me.Width) / 2
    Me.Top = (Application.Height - Me.Height) / 2
End Sub

Sub SetUserFormFont(uf As Object)
    Dim ctrl As Control
    Dim fontName As String
    Dim fontFound As Boolean
    Dim i As Integer
    Dim fontList As Variant
    
    ' Default font fallback list
    fontList = Array( _
    "Noto Sans", _
    "Segoe UI", _
    "San Francisco", _
    "Aptos", _
    "Calibri", _
    "Arial", _
    "Liberation Sans", _
    "DejaVu Sans", _
    "Times New Roman" _
    )
    
    fontFound = False
    ' Find the first available font
    For i = LBound(fontList) To UBound(fontList)
        Debug.Print "Trying font: [" & fontList(i) & "]"
        If FontExists(CStr(fontList(i))) Then
            fontName = fontList(i)
            fontFound = True
            Debug.Print "Font found: [" & fontName & "]"
            Exit For
        End If
    Next i
    
    ' fallback Font
    If Not fontFound Then
        fontName = "sans-serif"
        Debug.Print "Fallback to: [" & fontName & "]"
    End If
    
    ' Loop through all controls
    For Each ctrl In uf.Controls
        On Error Resume Next
        ' Some controls don't have a Font property (e.g., frames, images)
        ctrl.fontName = fontName
        ctrl.FontSize = 9
        Debug.Print "Set font for: " & ctrl.name
        Debug.Print "Set font as: " & ctrl.fontName
        Debug.Print "Set font size as: " & ctrl.FontSize
        On Error GoTo 0
    Next ctrl
End Sub

' Function to check if font is installed
Function FontExists(fontName As String) As Boolean
    On Error Resume Next
    FontExists = False
    Dim f As Variant
    For Each f In Application.FontNames
        If StrComp(f, fontName, vbTextCompare) = 0 Then
            FontExists = True
            Exit Function
        End If
    Next
    On Error GoTo 0
End Function

' Handle window resizing
Private Sub UserForm_Resize()
    PositionControls
End Sub

' Handle OK button click
Private Sub btnOK_Click()
    DialogResult = vbOK
    Me.Hide
End Sub

' Handle Cancel button click
Private Sub btnCancel_Click()
    DialogResult = vbCancel
    Me.Hide
End Sub

' Handle Select All button click
Private Sub btnSelectAll_Click()
    Dim i As Long
    For i = 0 To lstItems.ListCount - 1
        lstItems.Selected(i) = True
    Next i
End Sub

' Handle Clear All button click
Private Sub btnClearAll_Click()
    Dim i As Long
    For i = 0 To lstItems.ListCount - 1
        lstItems.Selected(i) = False
    Next i
End Sub

' Handle X button click
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        ' User clicked X button - treat as cancel
        DialogResult = vbCancel
        Cancel = True ' Prevent actual closing
        Me.Hide ' Hide instead
    End If
End Sub
