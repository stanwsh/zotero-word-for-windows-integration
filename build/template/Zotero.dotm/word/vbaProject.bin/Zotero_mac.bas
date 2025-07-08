Attribute VB_Name = "Zotero"
' ***** BEGIN LICENSE BLOCK *****
'
' Copyright (c) 2015  Zotero
'                     Center for History and New Media
'                     George Mason University, Fairfax, Virginia, USA
'                     http://zotero.org
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
' ***** END LICENSE BLOCK *****

Sub ZoteroInsertCitation()
CallZotero ("addCitation")
End Sub

Sub ZoteroAddNote()
CallZotero ("addNote")
End Sub

Sub ZoteroEditCitation()
CallZotero ("editCitation")
End Sub

Sub ZoteroAddEditCitation()
CallZotero ("addEditCitation")
End Sub

Sub ZoteroAddEditBibliography()
CallZotero ("addEditBibliography")
End Sub

Sub ZoteroSetDocPrefs()
CallZotero ("setDocPrefs")
End Sub

Sub ZoteroRefresh()
CallZotero ("refresh")
End Sub

Sub ZoteroRemoveCodes()
CallZotero ("removeCodes")
End Sub

Sub CallZotero(func)
Dim zoteroUrl As String
If zoteroUrl = "" Then
    zoteroUrl = "http://127.0.0.1:23119/integration/macWordCommand?"
    FileNum = FreeFile()
    On Error GoTo customUrlNotSet
        Open Application.StartupPath & Application.PathSeparator & "ZoteroPort.txt" For Input As #FileNum
        If Not EOF(FileNum) Then
            Line Input #FileNum, DataLine
            zoteroUrl = "http://127.0.0.1:" & DataLine & "/integration/macWordCommand?"
        End If
End If

customUrlNotSet:
On Error GoTo -1
nl$ = Chr$(10)
templateVersion$ = "2"
wordVersion$ = "MacWord2016"
onFailMessage$ = "Word could not communicate with Zotero. Please ensure that Zotero is open and try again."

On Error GoTo catchWordPath
    wordAppPath$ = AppleScriptTask("Zotero.scpt", "getPath", "")
    GoTo okWordPath
catchWordPath:
    wordAppPath$ = MacScript("POSIX path of (path to current application)")
    
okWordPath:
On Error GoTo -1
wordAppPath$ = Replace(wordAppPath$, " ", "%20")
#If VBA6 Then
     Dim majorVersion As Integer
     majorVersion = Split(Application.Version, ".")(0)
     If majorVersion >= 16 Then
         wordVersion$ = "MacWord16"
     End If
#End If

httpShellScript$ = "curl -s -o /dev/null -I -m 2 -w '%{http_code}' -X GET '" & zoteroUrl & "agent=" & wordVersion$ & "&command=" & func & "&document=" & wordAppPath$ & "&templateVersion=" & templateVersion$ & "' | grep -q '200' || exit 1"

On Error GoTo catchMacScriptHttp
    MacScript "try" & nl$ & "do shell script """ & httpShellScript$ & """" & nl$ & "on error" & nl$ & "display alert """ & onFailMessage$ & """  as critical" & nl$ & "end try"
Exit Sub

catchMacScriptHttp:
On Error GoTo -1
On Error GoTo catchAppleScriptTaskHttp
    Result$ = AppleScriptTask("Zotero.scpt", "callZotero", httpShellScript$)
Exit Sub
    
catchAppleScriptTaskHttp:
On Error GoTo -1
On Error GoTo catchMacScriptPipeWrite
    pipeLocation$ = "PIPE=\""$CFFIXED_USER_HOME/.zoteroIntegrationPipe\""; if [ ! -e \""$PIPE\"" ]; then PIPE=\""/Users/$USER/Library/Containers/com.microsoft.Word/Data/.zoteroIntegrationPipe\""; fi; if [ ! -e \""$PIPE\"" ]; then PIPE=.zoteroIntegrationPipe; fi"
    MacScript "try" & nl$ & "do shell script """ & pipeLocation$ & "; if [ -e \""$PIPE\"" ]; then echo '" & wordVersion$ & " " & func & " "" & POSIX path of (path to current application) & "" " & templateVersion & "' > \""$PIPE\""; else exit 1; fi;""" & nl$ & "on error" & nl$ & "display alert """ & onFailMessage$ & """  as critical" & nl$ & "end try"
Exit Sub

catchMacScriptPipeWrite:
On Error GoTo -1
    pipeLocation$ = "PIPE=""$HOME/.zoteroIntegrationPipe""; if [ ! -e ""$PIPE"" ]; then PIPE=""$HOME/Library/Containers/com.microsoft.Word/Data/.zoteroIntegrationPipe""; fi"
    shellScript$ = pipeLocation$ & "; if [ -e ""$PIPE"" ]; then echo '" & wordVersion$ & " " & func & " " & Application.Path & Application.PathSeparator & Application.Name & ".app/ " & templateVersion & "' > ""$PIPE""; else exit 1; fi;"
    Result$ = AppleScriptTask("Zotero.scpt", "callZotero", shellScript$)
End Sub

Public Sub ZoteroGoToZotero()
    ' Check if the selection is in a valid Zotero citation field
    ' If Selection.Fields.Count = 0 Then
    '     MsgBox "Please place the cursor inside a Zotero citation to use Go To Zotero.", vbInformation, "Zotero"
    '     Exit Sub
    ' End If

    Dim zoteroField As Field
    Set zoteroField = Nothing

    Dim rng as Range
    Set rng = Selection.Range

    ' Step 1: First check all fields in the selection
    Debug.Print "Selection.Range.Start: " & Selection.Range.Start & ", End: " & Selection.Range.End

    Dim fld As Field
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldAddin Then ' Most Zotero fields are Addin fields
            ' Check if the cursor is inside the field code or result
            If rng.Start >= fld.Code.Start And rng.End <= fld.Result.End Then
                If InStr(1, fld.Code.Text, "ZOTERO_ITEM") > 0 Then
                    Set zoteroField = fld
                    Debug.Print " This is a Zotero field."
                    Exit For
                End If
            End If
        End If
    Next

    ' Step 2: If still nothing, try original selection-based method (just in case)
    If zoteroField Is Nothing And Selection.Fields.Count > 0 Then
        For Each fld In Selection.Fields
            If InStr(1, fld.Code.Text, "ZOTERO_ITEM") > 0 Then
                Set zoteroField = fld
                Exit For
            End If
        Next
    End If

    If zoteroField Is Nothing Then
        MsgBox "Please place the cursor in a Zotero citation before using Go To Zotero.", vbExclamation, "Zotero"
        Exit Sub
    End If

    ' Get field code text (the JSON)
    Dim fieldCode As String
    fieldCode = zoteroField.Code.text
    
    ' The JSON is enclosed in braces; find the JSON content
    Dim jsonStart As Long, jsonEnd As Long
    jsonStart = InStr(fieldCode, "{")
    jsonEnd = InStrRev(fieldCode, "}")
    Dim jsonText As String
    jsonText = Mid$(fieldCode, jsonStart, jsonEnd - jsonStart + 1)
    
    ' Create collections to store citation items and their metadata
    Dim itemKeys As New Collection
    Dim itemDetails As New Collection
    Dim displayNames As New Collection
    Dim groupIDs As New Collection
    
    ' Extract citation items from JSON
    ' Look for the "citationItems" array which contains all cited works
    Dim citItemsPos As Long
    citItemsPos = InStr(1, jsonText, """citationItems"":[")
    If citItemsPos = 0 Then
        ' If we can't find the citationItems array, try the old extraction method
        ExtractCitationItemKeysLegacy jsonText, itemKeys, groupIDs
    Else
        ' Parse the citation items array to extract keys and metadata
        ExtractCitationItems jsonText, itemKeys, itemDetails, displayNames, groupIDs
    End If
    
    If itemKeys.Count = 0 Then
        MsgBox "No Zotero items found in this citation.", vbExclamation, "Zotero"
        Exit Sub
    End If
    
    ' If multiple items, show selection dialog
    Dim selectedKeys As New Collection
    If itemKeys.Count > 1 Then
        Dim selectionResult As Variant
        selectionResult = ShowItemSelectionDialog(itemKeys, displayNames)
        
        ' If user canceled, exit
        If IsEmpty(selectionResult) Then Exit Sub
        
        ' Process selected items
        ' Handle both array and Collection return types
        If IsArray(selectionResult) Then
            ' Handle array return (from SimpleItemSelection)
            Dim i As Long
            For i = LBound(selectionResult) To UBound(selectionResult)
                Dim indexValue As Integer
                indexValue = selectionResult(i)
                If indexValue >= 1 And indexValue <= itemKeys.Count Then
                    selectedKeys.Add itemKeys(indexValue)
                End If
            Next i
        Else
            ' Handle Collection return (from UserForm selection)
            Dim idx As Variant
            For Each idx In selectionResult
                selectedKeys.Add itemKeys(CInt(idx))
            Next
        End If
    Else
        ' Single item - no dialog needed
        selectedKeys.Add itemKeys(1)
    End If
    
    ' Open the selected items in Zotero
    If selectedKeys.Count > 0 Then
        OpenItemsInZotero selectedKeys, groupIDs
    End If
End Sub

' Extract keys using the legacy method (direct string search)
Private Sub ExtractCitationItemKeysLegacy(jsonText As String, ByRef itemKeys As Collection, ByRef groupIDs As Collection)
    Dim pos As Long, key As String
    pos = InStr(1, jsonText, "/items/")
    
    Do While pos > 0
        ' Find start of key and end quote
        Dim keyStart As Long, keyEnd As Long
        keyStart = pos + Len("/items/")
        keyEnd = InStr(keyStart, jsonText, """")
        If keyEnd = 0 Then keyEnd = InStr(keyStart, jsonText, "}") ' fallback, in case no quote
        key = Mid$(jsonText, keyStart, keyEnd - keyStart)
        If key <> "" Then
            On Error Resume Next
            ' Check if this is a group item
            Dim groupPos As Long, groupID As String
            groupPos = InStrRev(jsonText, "/groups/", pos)
            If groupPos > 0 And pos - groupPos < 30 Then ' Assume group ID is within 30 chars of the item key
                Dim groupStart As Long, groupEnd As Long
                groupStart = groupPos + Len("/groups/")
                groupEnd = InStr(groupStart, jsonText, "/")
                If groupEnd > groupStart Then
                    groupID = Mid$(jsonText, groupStart, groupEnd - groupStart)
                    groupIDs.Add groupID, key
                End If
            End If
            ' Add the key to our collection
            itemKeys.Add key, key ' Use key as key to avoid duplicates
            On Error GoTo 0
        End If
        pos = InStr(keyEnd + 1, jsonText, "/items/")
    Loop
End Sub

' Extract citation items with metadata from JSON
Private Sub ExtractCitationItems(jsonText As String, ByRef itemKeys As Collection, ByRef itemDetails As Collection, _
                                ByRef displayNames As Collection, ByRef groupIDs As Collection)
    Dim citItemsStart As Long, citItemsEnd As Long
    Dim currentPos As Long, itemStart As Long, itemEnd As Long
    
    ' Find the citationItems array
    citItemsStart = InStr(1, jsonText, """citationItems"":[") + Len("""citationItems"":[")
    citItemsEnd = FindMatchingBracket(jsonText, citItemsStart, "[", "]")
    
    If citItemsStart = 0 Or citItemsEnd = 0 Then Exit Sub
    
    ' Extract the array content
    Dim citItemsArray As String
    citItemsArray = Mid$(jsonText, citItemsStart, citItemsEnd - citItemsStart)
    
    ' Process each item object in the array
    currentPos = 1
    Do While currentPos < Len(citItemsArray)
        ' Find start of item object
        itemStart = InStr(currentPos, citItemsArray, "{")
        If itemStart = 0 Then Exit Do
        
        ' Find end of item object
        itemEnd = FindMatchingBracket(citItemsArray, itemStart, "{", "}")
        If itemEnd = 0 Then Exit Do
        
        ' Extract the item JSON
        Dim itemJson As String
        itemJson = Mid$(citItemsArray, itemStart, itemEnd - itemStart + 1)
        
        ' Extract URI, key, and other metadata
        Dim itemKey As String, displayName As String, groupID As String
        ExtractItemData itemJson, itemKey, displayName, groupID
        
        ' Add to collections if key found
        If itemKey <> "" Then
            On Error Resume Next
            itemKeys.Add itemKey, itemKey
            itemDetails.Add itemJson, itemKey
            displayNames.Add displayName, itemKey
            If groupID <> "" Then
                groupIDs.Add groupID, itemKey
            End If
            On Error GoTo 0
        End If
        
        ' Move to next item
        currentPos = itemEnd + 1
    Loop
End Sub

' Find matching closing bracket
Private Function FindMatchingBracket(text As String, startPos As Long, openBracket As String, closeBracket As String) As Long
    Dim depth As Long, i As Long
    depth = 1
    
    For i = startPos + 1 To Len(text)
        If Mid$(text, i, 1) = openBracket Then
            depth = depth + 1
        ElseIf Mid$(text, i, 1) = closeBracket Then
            depth = depth - 1
            If depth = 0 Then
                FindMatchingBracket = i
                Exit Function
            End If
        End If
    Next i
    
    FindMatchingBracket = 0 ' No match found
End Function

' Extract item data from JSON
Private Sub ExtractItemData(itemJson As String, ByRef itemKey As String, ByRef displayName As String, ByRef groupID As String)
    itemKey = ""
    displayName = ""
    groupID = ""
    
    ' Extract item key from URI
    Dim uriPos As Long, uriStart As Long, uriEnd As Long
    uriPos = InStr(1, itemJson, """uri""")
    If uriPos = 0 Then
        uriPos = InStr(1, itemJson, """uris""")
        If uriPos > 0 Then
            ' Find first URI in the array
            uriStart = InStr(uriPos, itemJson, "http")
            If uriStart > 0 Then
                uriEnd = InStr(uriStart, itemJson, """")
                If uriEnd > uriStart Then
                    Dim uri As String
                    uri = Mid$(itemJson, uriStart, uriEnd - uriStart)
                    
                    ' Check for group ID
                    Dim groupPos As Long
                    groupPos = InStr(1, uri, "/groups/")
                    If groupPos > 0 Then
                        Dim groupStart As Long, groupEnd As Long
                        groupStart = groupPos + Len("/groups/")
                        groupEnd = InStr(groupStart, uri, "/")
                        If groupEnd > groupStart Then
                            groupID = Mid$(uri, groupStart, groupEnd - groupStart)
                        End If
                    End If
                    
                    ' Extract key from URI
                    Dim keyPos As Long
                    keyPos = InStr(1, uri, "/items/")
                    If keyPos > 0 Then
                        itemKey = Mid$(uri, keyPos + Len("/items/"))
                    End If
                End If
            End If
        End If
    Else
        ' Direct URI extraction
        uriStart = InStr(uriPos, itemJson, "http")
        If uriStart > 0 Then
            uriEnd = InStr(uriStart, itemJson, """")
            If uriEnd > uriStart Then
                Dim directUri As String
                directUri = Mid$(itemJson, uriStart, uriEnd - uriStart)
                
                ' Check for group
                Dim directGroupPos As Long
                directGroupPos = InStr(1, directUri, "/groups/")
                If directGroupPos > 0 Then
                    Dim directGroupStart As Long, directGroupEnd As Long
                    directGroupStart = directGroupPos + Len("/groups/")
                    directGroupEnd = InStr(directGroupStart, directUri, "/")
                    If directGroupEnd > directGroupStart Then
                        groupID = Mid$(directUri, directGroupStart, directGroupEnd - directGroupStart)
                    End If
                End If
                
                ' Extract key
                Dim directKeyPos As Long
                directKeyPos = InStr(1, directUri, "/items/")
                If directKeyPos > 0 Then
                    itemKey = Mid$(directUri, directKeyPos + Len("/items/"))
                End If
            End If
        End If
    End If
    
    ' Try to extract title, author, year for display
    Dim title As String, author As String, year As String
    
    ' Extract title
    Dim titlePos As Long, titleStart As Long, titleEnd As Long
    titlePos = InStr(1, itemJson, """title""")
    If titlePos > 0 Then
        titleStart = InStr(titlePos, itemJson, ":")
        If titleStart > 0 Then
            titleStart = InStr(titleStart, itemJson, """")
            If titleStart > 0 Then
                titleStart = titleStart + 1
                titleEnd = InStr(titleStart, itemJson, """")
                If titleEnd > titleStart Then
                    title = Mid$(itemJson, titleStart, titleEnd - titleStart)
                End If
            End If
        End If
    End If
    
    ' Extract author (family name of first author)
    Dim authorPos As Long, authorStart As Long, authorEnd As Long, familyPos As Long
    authorPos = InStr(1, itemJson, """author""")
    If authorPos > 0 Then
        familyPos = InStr(authorPos, itemJson, """family""")
        If familyPos > 0 Then
            authorStart = InStr(familyPos, itemJson, ":")
            If authorStart > 0 Then
                authorStart = InStr(authorStart, itemJson, """")
                If authorStart > 0 Then
                    authorStart = authorStart + 1
                    authorEnd = InStr(authorStart, itemJson, """")
                    If authorEnd > authorStart Then
                        author = Mid$(itemJson, authorStart, authorEnd - authorStart)
                    End If
                End If
            End If
        End If
    End If
    
    ' Extract year
    Dim yearPos As Long, issuedPos As Long, dateParts As Long
    issuedPos = InStr(1, itemJson, """issued""")
    If issuedPos > 0 Then
        dateParts = InStr(issuedPos, itemJson, """date-parts""")
        If dateParts > 0 Then
            yearPos = InStr(dateParts, itemJson, "[")
            If yearPos > 0 Then
                yearPos = InStr(yearPos, itemJson, "[") ' Find nested array
                If yearPos > 0 Then
                    yearPos = InStr(yearPos, itemJson, """")
                    If yearPos > 0 Then
                        yearPos = yearPos + 1
                        Dim yearEnd As Long
                        yearEnd = InStr(yearPos, itemJson, """")
                        If yearEnd > yearPos Then
                            year = Mid$(itemJson, yearPos, yearEnd - yearPos)
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    ' Build display name
    If title <> "" Or author <> "" Or year <> "" Then
        If author <> "" Then
            displayName = author
            If year <> "" Then
                displayName = displayName & " (" & year & ")"
            End If
            If title <> "" Then
                displayName = displayName & " - " & title
            End If
        ElseIf title <> "" Then
            displayName = title
            If year <> "" Then
                displayName = displayName & " (" & year & ")"
            End If
        ElseIf year <> "" Then
            displayName = "Item from " & year
        End If
    End If
    
    ' If no display name could be constructed, use the key
    If displayName = "" And itemKey <> "" Then
        displayName = "Item: " & itemKey
    End If
End Sub

' Show a dialog for the user to select items
Private Function ShowItemSelectionDialog(itemKeys As Collection, displayNames As Collection) As Variant
    ' Try to create the UserForm
    Dim frm As Object
    
    On Error Resume Next
    ' Use CreateObject rather than UserForms.Add for better reliability
    Set frm = UserForms.Add("frmZoteroSelectItems")
    
    If Err.Number <> 0 Then
        ' UserForm approach failed, fallback to simplified selection approach
        Debug.Print "UserForm not available. Error: " & Err.Description & " (" & Err.Number & ")"
        ShowItemSelectionDialog = SimpleItemSelection(itemKeys, displayNames)
        Exit Function
    End If
    On Error GoTo 0
    
    ' Configure the form
    With frm
        ' Set caption
        .Caption = "Select Zotero Items to Open"
        
        ' Resize form appropriately
        .Width = 400
        .Height = 300
        
        ' Add items to the list
        Dim i As Long
        For i = 1 To itemKeys.Count
            .lstItems.AddItem displayNames(i)
            ' Select all items by default
            .lstItems.Selected(i - 1) = True
        Next i
        
        ' Show the form modally
        .Show vbModal
        
        ' Process results based on DialogResult
        If .DialogResult = vbOK Then
            ' Count selected items
            Dim selectedCount As Integer
            selectedCount = 0
            For i = 0 To .lstItems.ListCount - 1
                If .lstItems.Selected(i) Then
                    selectedCount = selectedCount + 1
                End If
            Next i
            
            ' Create array with selected indexes
            If selectedCount > 0 Then
                Dim selectedArr() As Integer
                ReDim selectedArr(1 To selectedCount)
                
                Dim arrIndex As Integer
                arrIndex = 1
                
                For i = 0 To .lstItems.ListCount - 1
                    If .lstItems.Selected(i) Then
                        selectedArr(arrIndex) = i + 1 ' Make 1-based
                        arrIndex = arrIndex + 1
                    End If
                Next i
                
                ShowItemSelectionDialog = selectedArr
            Else
                ' No items selected
                MsgBox "No items were selected. Operation canceled.", vbInformation, "Zotero"
                ShowItemSelectionDialog = Empty
            End If
        Else
            ' User canceled
            ShowItemSelectionDialog = Empty
        End If
    End With
    
    ' Clean up
    Unload frm
    Set frm = Nothing
End Function

' Simple item selection when UserForms are not available
Private Function SimpleItemSelection(itemKeys As Collection, displayNames As Collection) As Variant
    ' Just display a list to the user and let them choose a number
    Dim msg As String
    msg = "Found " & itemKeys.Count & " Zotero items in this citation:" & vbCrLf & vbCrLf
    
    Dim i As Long
    For i = 1 To itemKeys.Count
        msg = msg & i & ") " & displayNames(i) & vbCrLf
    Next i
    
    msg = msg & vbCrLf & "Enter the numbers to open (separated by commas), or type 'all' for all items."
    msg = msg & vbCrLf & "Press Cancel to abort."
    
    ' Ask user
    Dim userInput As String
    userInput = InputBox(msg, "Select Zotero Items", "all")
    
    ' Check for cancel
    If userInput = "" Then
        SimpleItemSelection = Empty
        Exit Function
    End If
    
    ' Create array instead of Collection for more reliable return value
    Dim selectedArr() As Integer
    Dim selectedCount As Integer
    selectedCount = 0
    
    ' Process selection
    If LCase(Trim(userInput)) = "all" Then
        ' All items selected
        ReDim selectedArr(1 To itemKeys.Count)
        For i = 1 To itemKeys.Count
            selectedArr(i) = i
            selectedCount = selectedCount + 1
        Next i
    Else
        ' Parse numbers - first count how many valid selections we have
        Dim parts As Variant
        parts = Split(userInput, ",")
        Dim validCount As Integer
        validCount = 0
        
        Dim part As Variant
        For Each part In parts
            Dim num As Long
            num = Val(Trim(part))
            If num >= 1 And num <= itemKeys.Count Then
                validCount = validCount + 1
            End If
        Next part
        
        ' Now create the array with the right size
        If validCount > 0 Then
            ReDim selectedArr(1 To validCount)
            selectedCount = 0
            ' Fill the array
            For Each part In parts
                num = Val(Trim(part))
                If num >= 1 And num <= itemKeys.Count Then
                    selectedCount = selectedCount + 1
                    selectedArr(selectedCount) = num
                End If
            Next part
        End If
    End If
    
    ' Check if any items were selected
    If selectedCount = 0 Then
        MsgBox "No valid items were selected. Operation canceled.", vbInformation, "Zotero"
        SimpleItemSelection = Empty
    Else
        ' Return the array directly
        SimpleItemSelection = selectedArr
    End If
End Function

' Open the selected items in Zotero
Private Sub OpenItemsInZotero(selectedKeys As Collection, groupIDs As Collection)
    ' Handle both single-item and multi-item cases
    If selectedKeys.Count = 1 Then
        ' Single item - simple URI
        Dim key As String
        key = selectedKeys(1)
        
        Dim zoteroLink As String
        
        ' Check if it belongs to a group
        On Error Resume Next
        Dim groupID As String
        groupID = groupIDs(key)
        On Error GoTo 0
        
        If groupID <> "" Then
            zoteroLink = "zotero://select/groups/" & groupID & "/items/" & key
        Else
            zoteroLink = "zotero://select/library/items/" & key
        End If
        
        ' Open the URI
        OpenZoteroUrl zoteroLink
        
        ' MsgBox "Zotero item opened in Zotero library.", vbInformation, zoteroLink
    Else
        ' Multiple items - use item list in query parameter
        ' Group items by library (personal vs. each group)
        Dim personalItems As String
        personalItems = ""
        
        ' Collections for group items (key = groupID, value = comma-separated keys)
        Dim groupItems As New Collection
        
        Dim i As Integer
        For i = 1 To selectedKeys.Count
            key = selectedKeys(i)
            
            ' Check if it belongs to a group
            On Error Resume Next
            groupID = groupIDs(key)
            On Error GoTo 0
            
            If groupID <> "" Then
                ' Group item
                On Error Resume Next
                Dim groupKeyList As String
                groupKeyList = groupItems(groupID)
                
                If Err.Number <> 0 Then
                    ' First item for this group
                    groupItems.Add key, groupID
                Else
                    ' Add to existing group items
                    groupItems.Remove groupID
                    groupItems.Add groupKeyList & "," & key, groupID
                End If
                On Error GoTo 0
            Else
                ' Personal library item
                If personalItems = "" Then
                    personalItems = key
                Else
                    personalItems = personalItems & "," & key
                End If
            End If
        Next i
        
        ' Open personal library items
        If personalItems <> "" Then
            zoteroLink = "zotero://select/library/items?itemKey=" & personalItems
            OpenZoteroUrl zoteroLink
        End If
        
        ' Open group library items
        If groupItems.Count > 0 Then
            Dim g As Variant
            For Each g In groupItems.Keys
                groupID = g
                Dim groupKeysList As String
                groupKeysList = groupItems(groupID)
                
                zoteroLink = "zotero://select/groups/" & groupID & "/items?itemKey=" & groupKeysList
                OpenZoteroUrl zoteroLink
            Next g
        End If
        
        ' MsgBox selectedKeys.Count & " Zotero items opened in Zotero library.", vbInformation, zoteroLink
    End If
End Sub

Private Sub OpenZoteroUrl(zoteroLink As String)
    ' Open the Zotero link using the default application
    On Error Resume Next
    Shell "open '" & zoteroLink & "'", vbNormalFocus
    On Error GoTo 0
End Sub
