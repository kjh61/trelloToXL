Attribute VB_Name = "TrelloImport"
'    Tool for importing Trello JSON files into Excel.
'    Written by Kevin Harper, 16-Mar-17.
'
'    This spreadsheet uses the JSON export from Trello to import a rich content into Excel as a list of cards or a list of actions.
'    No Chrome or other browser extensions are required for the export/import.
'    The import scripts use the JSON parser capability developed by Tim Hall, available at:
'    https://github.com/VBA-tools/VBA-JSON
'    The VBA-JSON parser is already included in the spreadsheet file
'
'    From within the spreadsheet, there are two import options
'    1. The "ImportedCards" sheet runs vbscript ImportMyTrello - this creates a list of all the cards on your Trello board,
'    excluding those that have been archived.
'    2. The "ImportedActions" sheet runs vbscript ImportActionsFromTrello - this creates a list of
'    all the actions that have been carried out on your Trello board -
'    by applying Excel filters, you can create a list of specific actions;  eg, such as a record of who and when moved cards into a particular column.
'
'    To export a board from Trello into Excel using the spreadsheet, carry out the following steps:
'
'    (1) In Trello, for your chosen board and using the right-hand side menu options, select:
'           More / Print and Export / Export to JSON
'
'    o  In Chrome, this will display the JSON code in the open tab.  Save this as a local file on your computer,
'    by right clicking and selecting "Save as..."
'    o  In Internet Explorer, this will download the JSON export as a local file on your computer.
'
'    (2) Using the spreadsheet, click the "Import" button on either worksheet and select your downloaded JSON file
'    ....and the import should proceed.
'
' Copyright (c) 2017, Kevin Harper
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.


' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Option Explicit
Public myFile As String
Public cardMovedColumn As Boolean

Public Sub ImportMyTrello()

    Dim JsonText As String
    Dim Parsed As New Dictionary
    
    Const StartRow = 4    'Row to contain column headers, should be greater than 3
    
    ' Fetch the JSON string and then parse it to Dictionary, or Exit if file selection is cancelled:
    JsonText = GetJsonText()
    If JsonText = "" Then Exit Sub
    
    Set Parsed = JsonConverter.ParseJson(JsonText)
    
'    Clear the current content from the spreadsheet:
'    Sheets("ImportedCards").Cells.ClearContents
    Range(Sheets("ImportedCards").Cells(1, 1), Sheets("ImportedCards").Cells(1, 1).SpecialCells(xlLastCell)).ClearContents

'    and insert the Board name and URL
    Sheets("ImportedCards").Cells(1, 3) = Parsed("name")
    Sheets("ImportedCards").Cells(2, 1) = "URL for imported Trello board:"
    Sheets("ImportedCards").Hyperlinks.Add Anchor:=Cells(3, 1), Address:=Parsed("url"), TextToDisplay:=Parsed("url")
        
'   =======================================================================================

'    Create an array called Columns, of content (id, name);  this is used later to lookup
'    column names based on their id

    Dim Columns As Variant
    Dim NumberOfColumns As Integer
    Dim NumberOfArchivedColumns As Integer
    Call FetchColumns(Parsed, Columns, NumberOfColumns, NumberOfArchivedColumns)

'   =======================================================================================

'    Create an array called Members, of content (id, fullName);  this is used later to lookup
'    people's names based on their id

    Dim Members As Variant
    Dim NumberOfMembers As Integer
    Call FetchMembers(Parsed, Members, NumberOfMembers)

'   =======================================================================================
    
'    Create an array called Cards, of content aligned to the following headers:
   
    Const CollectedColumns = 10
    Dim LastColumnIndex As Integer
    LastColumnIndex = CollectedColumns - 1
    Dim Header(CollectedColumns) As String
    Header(0) = "Column"
    Header(1) = "Card title"
    Header(2) = "Description"
    Header(3) = "Date of last update"
    Header(4) = "Due date"
    Header(5) = "Card Members"
    Header(6) = "Labels"
    Header(7) = "Card ID"
    Header(8) = "Attachments"
    Header(9) = "URL"

    Dim myObject As String
    myObject = "cards"
    
    Dim Card As Dictionary
    Dim Label As Dictionary
    Dim CardMember As Variant
    Dim Cards As Variant
    Dim Filename As Variant
    Dim NumberOfCards As Integer
    Dim NumberOfArchivedCards As Integer
    Dim i As Integer
    Dim FirstPass As Boolean

    
'    Redim the upper array bounds, where lowest bounds are 0
    ReDim Cards(Parsed(myObject).Count - 1, LastColumnIndex)
    
'   Loop through each card in the list, skipping it if it has been closed in Trello (ie archived):
    i = 0
    For Each Card In Parsed(myObject)
        If (Card("closed") <> "True") Then
            
'           First do all the card fields that are single values:
            Cards(i, 0) = Application.VLookup(Card("idList"), Columns, 2, False)        ' Lookup the list id to extract it's name, stored in the list created earlier
            Cards(i, 1) = Card("name")
            Cards(i, 2) = Card("desc")
            Cards(i, 3) = Left(Card("dateLastActivity"), 10)
            Cards(i, 4) = Left(Card("due"), 10)
            Cards(i, 7) = Card("idShort")
            Cards(i, 9) = Card("shortUrl")
            
'           The next three fields may have more than one value for the card; where multiple values exist, a comma separated list is created
            FirstPass = True
            For Each CardMember In Card("idMembers")
     
                If FirstPass Then
                    Cards(i, 5) = Application.VLookup(CardMember, Members, 2, False)    ' Lookup the person's id to extract their name, stored in the list created earlier
                    FirstPass = False
                Else
                    Cards(i, 5) = Cards(i, 5) & ", " & Application.VLookup(CardMember, Members, 2, False)
                End If
            Next CardMember
           
            FirstPass = True
            For Each Label In Card("labels")
                If FirstPass Then
                    Cards(i, 6) = Label("name")
                    FirstPass = False
                Else
                    Cards(i, 6) = Cards(i, 6) & ", " & Label("name")
                End If
            Next Label
    
            FirstPass = True
            For Each Filename In Card("attachments")
                If FirstPass Then
                    Cards(i, 8) = Filename("name")
                    FirstPass = False
                Else
                    Cards(i, 8) = Cards(i, 8) & ", " & Filename("name")
                End If
            Next Filename
                      
            i = i + 1
        Else
            NumberOfArchivedCards = NumberOfArchivedCards + 1
            
        End If
    Next Card
    
    NumberOfCards = i
'   =======================================================================================

'    Put the Cards array into the spreadsheet:
    Sheets("ImportedCards").Range(Cells(StartRow, 1), Cells(StartRow, CollectedColumns)) = Header
    Sheets("ImportedCards").Range(Cells(StartRow + 1, 1), Cells(StartRow + 1 + NumberOfCards - 1, CollectedColumns)) = Cards
    
'   Add some stats in the header:
    
    Sheets("ImportedCards").Cells(2, 4) = NumberOfColumns
    Sheets("ImportedCards").Cells(2, 5) = "columns (lists) imported from Trello"
    Sheets("ImportedCards").Cells(3, 4) = NumberOfCards
    Sheets("ImportedCards").Cells(3, 5) = "cards"
    
    Sheets("ImportedCards").Cells(2, 7) = NumberOfArchivedColumns
    Sheets("ImportedCards").Cells(2, 8) = "archived column(s) NOT imported"
    Sheets("ImportedCards").Cells(3, 7) = NumberOfArchivedCards
    Sheets("ImportedCards").Cells(3, 8) = "archived card(s) NOT imported"
    
    Call FormatSheet("ImportedCards", StartRow)

End Sub

Public Sub ImportActionsFromTrello()

    Dim JsonText As String
    Dim Parsed As New Dictionary

    Const StartRow = 4    'Row to contain column headers, should be greater than 3
    
    ' Fetch the JSON string and then parse it to Dictionary, or Exit if file selection is cancelled:
    JsonText = GetJsonText()
    If JsonText = "" Then Exit Sub
    
    Set Parsed = JsonConverter.ParseJson(JsonText)
    
    
'    Clear the current content from the spreadsheet:
'    Sheets("ImportedActions").Cells.ClearContents
    Range(Sheets("ImportedActions").Cells(1, 1), Sheets("ImportedActions").Cells(1, 1).SpecialCells(xlLastCell)).ClearContents
    
'    and insert the Board name and URL
    Sheets("ImportedActions").Cells(1, 3) = Parsed("name")
    Sheets("ImportedActions").Cells(2, 1) = "URL for imported Trello board:"
    Sheets("ImportedActions").Hyperlinks.Add Anchor:=Cells(3, 1), Address:=Parsed("url"), TextToDisplay:=Parsed("url")
        
'   =======================================================================================

'    Create an array called Members, of content (id, fullName);  this is used later to lookup
'    people's names based on their id

    Dim Members As Variant
    Dim NumberOfMembers As Integer
    Call FetchMembers(Parsed, Members, NumberOfMembers)
    
'   =======================================================================================
    
'    Create an array called Cards, of content aligned to the following headers:

    Const CollectedColumns = 6
    Dim LastColumnIndex As Integer
    LastColumnIndex = CollectedColumns - 1
    Dim Header(CollectedColumns) As String
    Header(0) = "Column"
    Header(1) = "Card title"
    Header(2) = "Member"
    Header(3) = "Action"
    Header(4) = "Date"
    Header(5) = "Details"
    
    Dim myObject As String
    myObject = "actions"
    Dim Action As Dictionary
    Dim Card As Variant
    Dim Actions As Variant
    Dim NumberOfActions As Integer
    Dim i As Integer
    Dim FirstPass As Boolean
    
'    Redim the upper array bounds, where lowest bounds are 0
    ReDim Actions(Parsed(myObject).Count - 1, LastColumnIndex)
    
'   Loop through each card in the list, skipping it if it has been closed in Trello (ie archived):
    i = 0
    For Each Action In Parsed(myObject)
            
            Actions(i, 0) = FindActionListName(Action)
            Actions(i, 1) = FindActionCardName(Action)

            Actions(i, 2) = Application.VLookup(Action("idMemberCreator"), Members, 2, False)        ' Lookup the person's id to extract their name, stored in the list created earlier
            Actions(i, 3) = Action("type")
            Actions(i, 4) = Left(Action("date"), 10)
            
            If Action("type") = "addAttachmentToCard" Then
                Actions(i, 5) = "Attached file: " & Action("data")("attachment")("name")
                
            ElseIf Action("type") = "deleteAttachmentFromCard" Then
                Actions(i, 5) = "Removed file: " & Action("data")("attachment")("name")
'                Actions(i, 0) =
                
            ElseIf Action("type") = "commentCard" Then
                Actions(i, 5) = "Added comment: " & Action("data")("text")
                
            ElseIf Action("type") = "createCard" Then
                Actions(i, 5) = "New card added"
                
            ElseIf Action("type") = "copyCard" Then
                Actions(i, 5) = "Copied from: " & Action("data")("cardSource")("name")
                
            ElseIf Action("type") = "addChecklistToCard" Then
                Actions(i, 5) = "Added checklist: " & Action("data")("checklist")("name")
                
            ElseIf Action("type") = "removeChecklistFromCard" Then
                Actions(i, 5) = "Removed checklist: " & Action("data")("checklist")("name")
                
            ElseIf Action("type") = "updateChecklist" Then
                Actions(i, 5) = "Changed from [" & Action("data")("old")("name") _
                    & "] to [" & Action("data")("checklist")("name") & "]"
                    
            ElseIf Action("type") = "updateCheckItemStateOnCard" Then
                Actions(i, 5) = "Checklist item [" & Action("data")("checkItem")("name") & "] set to <" & _
                    Action("data")("checkItem")("state") & ">"
                
            ElseIf Action("type") = "createList" Then
                Actions(i, 5) = "Added new column: " & Action("data")("list")("name")
                
            ElseIf Action("type") = "updateList" Then
                If Action("data")("old")("name") <> "" Then
                    Actions(i, 5) = "Column changed from: " & Action("data")("old")("name")
                    
                ElseIf Action("data")("old")("pos") > Action("data")("list")("pos") Then
                    Actions(i, 5) = "Column moved left"
                    
                ElseIf Action("data")("old")("pos") < Action("data")("list")("pos") Then
                    Actions(i, 5) = "Column moved right"
                    
                End If
                
            ElseIf Action("type") = "createBoard" Then
                Actions(i, 5) = "Added new board: " & Action("data")("board")("name")
                
            ElseIf Action("type") = "updateBoard" Then
                Actions(i, 5) = "Background changed from [" & Action("data")("old")("prefs")("background") _
                    & "] to [" & Action("data")("board")("prefs")("background") & "]"
                
            ElseIf Action("type") = "updateCard" Then
                If Action("data")("old")("due") <> "" Then
                    Actions(i, 5) = "Due changed from: " & Left(Action("data")("old")("due"), 10)
                    
                ElseIf Action("data")("old")("name") <> "" Then
                    Actions(i, 5) = "Name changed from: " & Action("data")("old")("name")
                    
                ElseIf Action("data")("old")("pos") > Action("data")("card")("pos") Then
                    Actions(i, 5) = "Card moved up in column"
                    
                ElseIf Action("data")("old")("pos") < Action("data")("card")("pos") Then
                    Actions(i, 5) = "Card moved down in column"
                    
                ElseIf cardMovedColumn Then
                    Actions(i, 5) = "Moved from column: " & Action("data")("listBefore")("name")
                    
                ElseIf Len(Action("data")("old")("desc")) > 0 Then
                    Actions(i, 5) = "Card description amended"
                    
                ElseIf Action("data")("old")("desc") = "" Then
                    Actions(i, 5) = "Card description added"
                    
                End If

            ElseIf Action("type") = "addMemberToBoard" Then
                Actions(i, 5) = "Added new member: " & Application.VLookup(Action("data")("idMemberAdded"), Members, 2, False)
                
            ElseIf Action("type") = "addMemberToCard" Then
                Actions(i, 5) = "Added to card: " & Application.VLookup(Action("data")("idMember"), Members, 2, False)
                
            ElseIf Action("type") = "removeMemberFromCard" Then
                Actions(i, 5) = "Removed: " & Application.VLookup(Action("data")("idMember"), Members, 2, False)
                
            ElseIf Action("type") = "addToOrganizationBoard" Then
                Actions(i, 5) = "Added to org: " & Action("data")("organization")("name")
                               
            
            End If
                     
                      
            i = i + 1

    Next Action
    
    NumberOfActions = i
'   =======================================================================================

'    Put the Cards array into the spreadsheet:
    Sheets("ImportedActions").Range(Cells(StartRow, 1), Cells(StartRow, CollectedColumns)) = Header
    Sheets("ImportedActions").Range(Cells(StartRow + 1, 1), Cells(StartRow + 1 + NumberOfActions - 1, CollectedColumns)) = Actions
    
'   Add some stats in the header:
    
    Sheets("ImportedActions").Cells(3, 4) = NumberOfActions
    Sheets("ImportedActions").Cells(3, 5) = "actions"

    Call FormatSheet("ImportedActions", StartRow)

End Sub


Private Function GetJsonText()
'   Open the filechooser, select file and read .json file
    Dim FSO As New FileSystemObject
    Dim JsonTS As TextStream
    
    myFile = Application.GetOpenFilename("JSON Files (*.json),*.json", 1, "Open JSON file for import")
    If myFile <> "False" Then
        Set JsonTS = FSO.OpenTextFile(myFile, ForReading)
        GetJsonText = JsonTS.ReadAll
        JsonTS.Close
    Else
        GetJsonText = ""
    End If
    
End Function


Private Sub FetchColumns(Parsed As Dictionary, Columns As Variant, NumberOfColumns As Integer, NumberOfArchivedColumns As Integer)

    Dim Column As Dictionary
    Dim myObject As String
    Dim i As Integer
    myObject = "lists"
'    Dim Columns As Variant
'    Dim NumberOfColumns As Integer
'    Dim NumberOfArchivedColumns As Integer
'    Redim the upper array bounds, where lowest bounds are 0
    ReDim Columns(Parsed(myObject).Count - 1, 1)
    
    i = 0
    For Each Column In Parsed(myObject)
        If (Column("closed") <> "True") Then
            Columns(i, 0) = Column("id")
            Columns(i, 1) = Column("name")
    
            i = i + 1
        Else
            NumberOfArchivedColumns = NumberOfArchivedColumns + 1
        End If
    Next Column
    NumberOfColumns = i

End Sub

Private Sub FetchMembers(Parsed As Dictionary, Members As Variant, NumberOfMembers As Integer)


    Dim Member As Dictionary
    Dim myObject As String
    Dim i As Integer
    myObject = "members"
    
'    Dim Members As Variant
'    Dim NumberOfMembers As Integer
'    Redim the upper array bounds, where lowest bounds are 0
    ReDim Members(Parsed(myObject).Count - 1, 1)
      
    i = 0
    For Each Member In Parsed(myObject)
        Members(i, 0) = Member("id")
        Members(i, 1) = Member("fullName")

        i = i + 1
    Next Member
    NumberOfMembers = i
    
End Sub





Private Function FindActionListName(Action As Dictionary)
    cardMovedColumn = False
    On Error GoTo ErrHandler1
    FindActionListName = Action("data")("list")("name")
    Exit Function
    
TryNext1:
    On Error GoTo ErrHandler2
    FindActionListName = Action("data")("listAfter")("name")
    cardMovedColumn = True
    Exit Function
    
TryNext2:
    FindActionListName = "not found"
    Exit Function
    
ErrHandler1:
    Resume TryNext1
ErrHandler2:
    Resume TryNext2
        
End Function


Private Function FindActionCardName(Action As Dictionary)
    On Error GoTo ErrHandler1
    FindActionCardName = Action("data")("card")("name")
    Exit Function
    
TryNext1:
    FindActionCardName = "not found"
    Exit Function
        
ErrHandler1:
    Resume TryNext1
    
End Function

Private Sub FormatSheet(Sheet As String, StartRow As Integer)
'   Apply formatting to the header:
    Sheets(Sheet).Rows("1:3").Style = "Accent1"
    
    With Sheets(Sheet).Rows("2:3")
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    With Sheets(Sheet).Range("C1")
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .RowHeight = 50
        .Font.Bold = True
        .Font.Size = 26
    End With
    
    With Sheets(Sheet).Range("D2:D3")
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    With Sheets(Sheet).Range("G2:G3")
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    With Sheets(Sheet).Rows(StartRow)
        .Style = "Accent1"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .RowHeight = 40
        .Font.Bold = True
    End With
    
    With Sheets(Sheet).Range("A3")
        .Style = "Hyperlink"
    End With
    
End Sub
