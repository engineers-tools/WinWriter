Attribute VB_Name = "mWordFunctions"
'The MIT License (MIT)
'
'Copyright (c) 2018 Carlos Gamez
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.
'
' Description: This module contains code of generic functions
'              controling a Word file.
'
' Authors:      Carlos Gamez
'
Option Explicit

' **************************************************************
' Module Constant Declarations Follow
' **************************************************************
Private Const mstrMODULE As String = "mWordFunctions"

Sub UpdateAllFields(objCurrentDocument As Word.document)
' --------------------------------------------------------------
' Comments: This routine updates all automaticaly caclculated
'           fields in Headers and Footers.
'
' Arguments:    objCurrentDocument = Reference to the currently
'                                    opened Word Document.
'
' Date          Developer       Comment
' --------------------------------------------------------------
' 05/11/11      Carlos Gamez    Initial version
' 26/01/2018    Carlos Gamez    Open source release
'
    Dim oStory As Word.Range
    Dim oField As Word.Field
  
  'Open each range that might contain fields
  For Each oStory In objCurrentDocument.StoryRanges
    'Update each field in that range
    For Each oField In oStory.fields
      oField.Update
    Next oField
  Next oStory
  
End Sub
Function InsertBlockMarker(objWordApplication As Word.Application, _
            strMarker As String, strMarkerType As String, _
            strTemplateName As String, strBlockName As String, _
            Optional objWordDocument As Word.document) As Boolean
' --------------------------------------------------------------
' Comments: This routine inserts a word block based on a marker.
'
' ARGUMENTS:
'
' objWordApplication (Word.Application) = Reference to the current
'                                       Word application instance.
'
' objWordDocument (Word.Document) = Reference to the current
'                                   Word document.
'
' strMarker (String) = String containing the marker to be searched
'                      for inserting the text.
'
' strMarkerType (String) = String containing the type of marker
'                      possible values are: Bookmark or Tag.
'
' strTemplateName (String) = name of the word template containing
'                           block to be inserted
'
' strBlockName (String) = String containing the name of the block
'                           to be inserted
'
' Date          Developer       Comment
' --------------------------------------------------------------
' 05/11/11      Carlos Gamez    Initial version
' 19/02/13      Carlos Gamez    Added error handling to detect when
'                               requested block doesn't exist in the
'                               word template
' 26/01/2018    Carlos Gamez    Open source release
'
On Error GoTo ErrHandler

    'Select insert method based on the Marker Type
    Select Case UCase(strMarkerType)
        
        Case "BOOKMARK"
            If objWordDocument.Bookmarks.Exists(strMarker) Then
                
                objWordDocument.Bookmarks(strMarker).Select
                
                With objWordApplication
                    .Selection.GoToPrevious What:=wdGoToLine
                    .Selection.TypeText Text:=vbCrLf
                    .Templates(strTemplateName). _
                        BuildingBlockEntries(strBlockName).Insert _
                        Where:=objWordApplication.Selection.Range, _
                        RichText:=True
                    .Selection.TypeText Text:=vbCrLf
                End With
                
                InsertBlockMarker = True
            Else
                InsertBlockMarker = False
            End If
        
        Case "TAG"
        
        With objWordApplication
    
            'Clear the find object from previous search terms
            .Selection.Find.ClearFormatting
            
            'Define the marker to look for
            .Selection.Find.Text = strMarker
            
            'Search for marker below the current selection postion
            .Selection.Find.Forward = True
            
            If .Selection.Find.Execute() = True Then
                'Insert into document
                objWordApplication.Templates(strTemplateName). _
                    BuildingBlockEntries(strBlockName).Insert _
                    Where:=objWordApplication.Selection.Range, _
                    RichText:=True
                    objWordApplication.Selection.TypeText Text:=vbCrLf
                'Signal success
                InsertBlockMarker = True
            Else 'try to find the text above the current selection position
                .Selection.Find.Forward = False
                If .Selection.Find.Execute() = True Then
                    'Insert into document
                    objWordApplication.Templates(strTemplateName). _
                    BuildingBlockEntries(strBlockName).Insert _
                    Where:=objWordApplication.Selection.Range, _
                    RichText:=True
                    objWordApplication.Selection.TypeText Text:=vbCrLf
                    'Signal success
                    InsertBlockMarker = True
                Else
                    'Signal failure
                    InsertBlockMarker = False
                End If
            End If
        End With
    
        Case Else
            InsertBlockMarker = False
        
    End Select
    
Exit_ErrHandler:
    Exit Function

ErrHandler:
    Select Case Err.Number
    'If block wasn't found
    Case Is = 5941
        sDisplayUnexpectedError Err.Number, "The block [" & strBlockName & "] could not be found " & _
                                "on the [" & strTemplateName & "] template file." & vbCrLf & _
                                "Please verify your settings.", mstrMODULE
    'Defalut error handle
    Case Else
        sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    End Select
    
    Resume Exit_ErrHandler
    Resume
    
End Function
Function InsertCollectionMarker(objWordApplication As Word.Application, _
            strMarker As String, strMarkerType As String, ctContentType As enmContentType, _
            colContent As Collection, Optional objWordDocument As Word.document _
            ) As Boolean
' --------------------------------------------------------------
' Comments: This routine inserts text content based on a marker.
'
' ARGUMENTS:
'
' objWordApplication (Word.Application) = Reference to the current
'                                       Word application instance.
'
' objWordDocument (Word.Document) = Reference to the current
'                                   Word document.
'
' strMarker (String) = String containing the marker to be searched
'                      for inserting the text.
'
' strMarkerType (String) = String containing the type of marker
'                      possible values are: Bookmark or Tag.
'
' colContent (Collection) = Collection that contains the content
'                           to be inserted
'
' ctContentType (enmContentType) = A constant identifying the type of
'                                   content that wants to be inserted
'
' Date           Developer       Comment
' --------------------------------------------------------------
' 17/07/12       Carlos Gamez    Initial version
' 26/01/2018     Carlos Gamez    Open source release
'
Dim Element As Variant
Dim strFileName As Variant
Dim intElementCount As Integer
Dim intLastElement As Integer

intLastElement = colContent.Count
intElementCount = 1
'Select insert method based on the Marker Type
Select Case UCase(strMarkerType)
    
    Case "BOOKMARK"
        If objWordDocument.Bookmarks.Exists(strMarker) Then
            'If content to be inserted is text
            If ctContentType = ctText Then
                For Each Element In colContent
                    objWordDocument.Bookmarks(strMarker).Select
                    objWordApplication.Selection.TypeText Text:=Element.value
                Next Element
                InsertCollectionMarker = True
            ElseIf ctContentType = ctHyperlink Then 'If content is a hyperlink
                
                objWordDocument.Bookmarks(strMarker).Select
                objWordApplication.Selection.TypeText Text:=vbCrLf
                
                For Each Element In colContent
                    
                    strFileName = ExtractFilenameFromFullPath(CStr(Element))
                    
                    objWordDocument.Bookmarks(strMarker).Select
                    
                    objWordDocument.Hyperlinks.Add Anchor:=objWordApplication.Selection.Range, _
                                                   Address:=CStr(Element), _
                                                   TextToDisplay:=strFileName
                    objWordApplication.Selection.MoveStart Unit:=wdLine, Count:=-1
                    objWordApplication.Selection.Collapse
                    If intElementCount <> intLastElement Then
                        objWordApplication.Selection.TypeText Text:=vbCrLf
                    End If
                    intElementCount = intElementCount + 1
                Next Element
                InsertCollectionMarker = True
            Else 'This function does not cater for inserting other type of content
                InsertCollectionMarker = False
            End If
        Else
            InsertCollectionMarker = False
        End If
    
    Case "TAG"
        'Not implemented
    Case Else
        InsertCollectionMarker = False
End Select
    
End Function

Function InsertTextMarker(objWordApplication As Word.Application, _
            strMarker As String, strMarkerType As String, _
            strContent As String, Optional objWordDocument As Word.document _
            ) As Boolean
' --------------------------------------------------------------
' Comments: This routine inserts text content based on a marker.
'
' ARGUMENTS:
'
' objWordApplication (Word.Application) = Reference to the current
'                                       Word application instance.
'
' objWordDocument (Word.Document) = Reference to the current
'                                   Word document.
'
' strMarker (String) = String containing the marker to be searched
'                      for inserting the text.
'
' strMarkerType (String) = String containing the type of marker
'                      possible values are: Bookmark or Tag.
'
' strContent (String) = String containing the text to be inserted
'
' Date          Developer       Comment
' --------------------------------------------------------------
' 05/11/11      Carlos Gamez    Initial version
' 26/01/2018    Carlos Gamez    Open source release
'
'
    'Select insert method based on the Marker Type
    Select Case UCase(strMarkerType)
        
        Case "BOOKMARK"
            If objWordDocument.Bookmarks.Exists(strMarker) Then
                objWordDocument.Bookmarks(strMarker).Select
                If strContent = "" Or strContent = vbNullString Then
                    objWordApplication.Selection.TypeText Text:=Chr(32)
                Else
                    objWordApplication.Selection.TypeText Text:=strContent
                End If
                InsertTextMarker = True
            Else
                InsertTextMarker = False
            End If
        
        Case "TAG"
        
            With objWordApplication
        
                'Clear the find object from previous search terms
                .Selection.Find.ClearFormatting
                
                'Make sure it searches for exact word (avoids replacing similar text)
                .Selection.Find.MatchWholeWord = True
                
                'Define the marker to look for
                .Selection.Find.Text = strMarker
                
                'Search for marker below the current selection postion
                .Selection.Find.Forward = True
                
                If .Selection.Find.Execute() = True Then
                    'Insert into document
                    .Selection.TypeText Text:=strContent
                    'Signal success
                    InsertTextMarker = True
                Else 'try to find the text above the current selection position
                    .Selection.Find.Forward = False
                    If .Selection.Find.Execute() = True Then
                        'Insert into document
                        .Selection.TypeText Text:=strContent
                        'Signal success
                        InsertTextMarker = True
                    Else
                        'Signal failure
                        InsertTextMarker = False
                    End If
                End If
            End With
    
        Case Else
            InsertTextMarker = False
        
    End Select

End Function

Function InsertImageMarker(objWordApplication As Word.Application, strMarker As String, _
                            strMarkerType As String, strContent As String, _
                            Optional objWordDocument As Word.document, _
                            Optional strSize As Variant) As Boolean
' --------------------------------------------------------------
' Comments: This routine inserts text content based on a marker.
'
' ARGUMENTS:
'
' objWordApplication (Word.Application) = Reference to the current
'                                       Word application instance.
'
' objWordDocument (Word.Document) = Reference to the current
'                                   Word document.
'
' strMarker (String) = String containing the marker to be searched
'                      for inserting the text.
'
' strMarkerType (String) = String containing the type of marker
'                      possible values are: Bookmark or Tag.
'
' strContent (String) = String containing the name of the image
'                        to be inserted, including full path
'
' strSize (String) = Height of image in pixels. If not included
'                       Image won't be resized
'
'
'
' Date          Developer       Comment
' --------------------------------------------------------------
' 05/11/11      Carlos Gamez    Initial version
' 26/01/2018    Carlos Gamez    Open source release
'
On Error GoTo ErrHandler

    'Select insert method based on the Marker Type
    Select Case UCase(strMarkerType)
        
        Case "BOOKMARK"
            If objWordDocument.Bookmarks.Exists(strMarker) Then
                objWordDocument.Bookmarks(strMarker).Select
                With objWordApplication
                    .Selection.InlineShapes.AddPicture Filename:=strContent, _
                                                       LinkToFile:=False, _
                                                       SaveWithDocument:=True
                    .Selection.ExtendMode = True
                    .Selection.MoveLeft
                    .Selection.ExtendMode = False
                    If Not IsMissing(strSize) Then
                        .Selection.InlineShapes(1).LockAspectRatio = msoTrue
                        .Selection.InlineShapes(1).Height = CSng(strSize)
                    End If
                    .Selection.MoveRight
                End With
                InsertImageMarker = True
            Else
                InsertImageMarker = False
            End If
        
        Case "TAG"
        
        With objWordApplication
    
            'Clear the find object from previous search terms
            .Selection.Find.ClearFormatting
            
            'Define the marker to look for
            .Selection.Find.Text = strMarker
            
            'Search for marker below the current selection postion
            .Selection.Find.Forward = True
            
            If .Selection.Find.Execute() = True Then
                'Insert into document
                With objWordApplication
                    .Selection.InlineShapes.AddPicture Filename:=strContent, _
                                                       LinkToFile:=False, _
                                                       SaveWithDocument:=True
                    .Selection.ExtendMode = True
                    .Selection.MoveLeft
                    .Selection.ExtendMode = False
                    If Not IsMissing(strSize) Then
                        .Selection.InlineShapes(1).LockAspectRatio = msoTrue
                        .Selection.InlineShapes(1).Height = CSng(strSize)
                    End If
                    .Selection.MoveRight
                End With
                'Signal success
                InsertImageMarker = True
            Else 'try to find the text above the current selection position
                .Selection.Find.Forward = False
                If .Selection.Find.Execute() = True Then
                    'Insert into document
                    With objWordApplication
                    .Selection.InlineShapes.AddPicture Filename:=strContent, _
                                                       LinkToFile:=False, _
                                                       SaveWithDocument:=True
                    .Selection.ExtendMode = True
                    .Selection.MoveLeft
                    .Selection.ExtendMode = False
                    If Not IsMissing(strSize) Then
                        .Selection.InlineShapes(1).LockAspectRatio = msoTrue
                        .Selection.InlineShapes(1).Height = CSng(strSize)
                    End If
                    .Selection.MoveRight
                End With
                    'Signal success
                    InsertImageMarker = True
                Else
                    'Signal failure
                    InsertImageMarker = False
                End If
            End If
        End With
    
        Case Else
            InsertImageMarker = False
        
    End Select

Exit_ErrHandler:
    Exit Function

ErrHandler:
    If Err.Number = 5152 Then
        ThisWorkbook.Activate
        MsgBox "The image >> " & strContent & " << was not found." & vbCrLf & _
                "The application will continue to run and place " & _
                "a message where the photo would normally be inserted.", vbInformation, gstrAPP_NAME
        objWordApplication.Activate
    Else
        MsgBox Err.Number & "-" & Err.Description, , gstrAPP_NAME
    End If
    InsertImageMarker = False
    Resume Exit_ErrHandler


End Function
Function HyperlinksFromRange(rngSource As Range, strFolderPath As String) As Collection
' --------------------------------------------------------------
' Comments: This function takes filenames from a range of Excel
'           cells and a folder path and returns a collection
'           with hyperlinks based on the concatenation of the
'           folder. The hyperlinks are returned as strings in the
'           collection with the key being the name of the file and
'           and the value the fully qualified path.
'
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 17/07/12      Carlos Gamez        Initial version
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler
    
    'Variable used to hold the content of each cell during the iteration cycle
    Dim Element As Variant
    Dim hypLinkToFile As String
    Dim colHyperlinksCollection As Collection
    
    Set colHyperlinksCollection = New Collection
    
    'Build the array from the range
    If Not rngSource Is Nothing Then
        For Each Element In rngSource
            If Element.Text <> "" And Element.Text <> "(blank)" Then
                'Create hyperlink
                colHyperlinksCollection.Add strFolderPath & Element.Text, Element.Text
            End If
        Next Element
    End If

ExitErrHandler:
    Set HyperlinksFromRange = colHyperlinksCollection
    Exit Function

ErrHandler:
    'If error is a duplicate hyperlink, ignore and continue processing
    Resume Next
    
End Function

Function ContentFromRange(rngSource As Range, varSeparator As Variant) As String
' --------------------------------------------------------------
' Comments: This function takes content from a range of Excel
'           cells and returns a string which is the concatenation
'           of the content of each cells. The function also takes
'           a character to be used as the separator between the
'           elements of this string.
'
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 21/12/11      Carlos Gamez        Initial version
' 26/01/2018    Carlos Gamez        Open source release
'
    'Array that will contain all the elements found in the cell range
    Dim ContentList() As String
    Dim ContentListUniqueItems() As String
    'Variable used to hold the content of each cell during the iteration cycle
    Dim Element As Variant
    
    ReDim ContentList(0)
    ReDim ContentListUniqueItems(0)
    
    ContentFromRange = ""
    
    'Build the array from the range
    If Not rngSource Is Nothing Then
        For Each Element In rngSource
            If Element.Text <> "" And Element.Text <> "(blank)" Then
                ContentList(UBound(ContentList)) = Element.Text
                ReDim Preserve ContentList(UBound(ContentList) + 1)
                
            End If
        Next Element
    End If
    
    'Eliminate the duplicated elements in the array
    EliminateDuplicates ContentList, ContentListUniqueItems
    
    'Concatenate the array elements into the final string
    For Each Element In ContentListUniqueItems
        If Element <> "" Then
            ContentFromRange = ContentFromRange & Element & varSeparator
        End If
    Next Element

End Function
Sub EliminateDuplicates(ByRef arrList() As String, ByRef arrUniques() As String)
' --------------------------------------------------------------
' Comments: This function takes an one-dimension string array and eliminates
'           all the duplicated elements. It returns the array without
'           duplicates.
'
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 21/12/11      Carlos Gamez        Initial version
' 26/01/2018    Carlos Gamez        Open source release
'
    Dim arrResult() As String
    Dim varElement As Variant
    Dim intElements As Integer
    Dim intCounter As Integer
    Dim intCompareCounter As Integer
    Dim boolDuplicate As Boolean
    
    ReDim arrResult(0)
    
    boolDuplicate = False
    
    intElements = UBound(arrList)
    
    For intCounter = 0 To intElements
        'Get nth element
        varElement = arrList(intCounter)
        
        'Compare that against the rest of the array
        For intCompareCounter = intCounter + 1 To intElements
            
            If varElement = arrList(intCompareCounter) Then boolDuplicate = True
            
        Next intCompareCounter
        
        If boolDuplicate = False Then
            arrResult(UBound(arrResult)) = varElement
            ReDim Preserve arrResult(UBound(arrResult) + 1)
        End If
        
        boolDuplicate = False
    
    Next intCounter
    
    ReDim arrUniques(UBound(arrResult))
    arrUniques = arrResult

End Sub

Function InsertTableRowsMarker(objWordApplication As Word.Application, strMarker As String, _
                            strMarkerType As String, intRows As Integer, _
                            Optional objWordDocument As Word.document, _
                            Optional strDirection As Variant = "BELOW" _
                            ) As Boolean
' --------------------------------------------------------------
' Comments:
'   Inserts a table on the word document at the specified marker
'   location.
'
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
    'Remove one row to account for the row with the marker
    intRows = intRows - 1
    
    'Select insert method based on the Marker Type
    Select Case UCase(strMarkerType)
        
        Case "BOOKMARK"
            'If bookmark is found
            If objWordDocument.Bookmarks.Exists(strMarker) Then
                'Select the bookmark
                objWordDocument.Bookmarks(strMarker).Select
                
                'If there's more than one row to insert
                If intRows > 0 Then
                    'Insert rows in desired direction
                    If UCase(strDirection) = "BELOW" Then
                        objWordApplication.Selection.InsertRowsBelow intRows
                        'Leave selection on row where marker was found
                        objWordApplication.Selection.GoToPrevious What:=wdGoToLine
                        InsertTableRowsMarker = True
                    ElseIf UCase(strDirection) = "ABOVE" Then
                        objWordApplication.Selection.InsertRowsAbove intRows
                        'Leave selection on row where marker was found
                        objWordApplication.Selection.GoToNext What:=wdGoToLine
                        InsertTableRowsMarker = True
                    Else
                        InsertTableRowsMarker = False
                    End If
                End If
            
            Else
                'Bookmark wasn't found
                InsertTableRowsMarker = False
            End If
        
        Case "TAG"
        
        With objWordApplication
    
            'Clear the find object from previous search terms
            .Selection.Find.ClearFormatting
            
            'Define the marker to look for
            .Selection.Find.Text = strMarker
            
            'Search for marker below the current selection postion
            .Selection.Find.Forward = True
            
            If .Selection.Find.Execute() = True Then
                'Clear marker
                .Selection.TypeText Text:=Chr(32)
                
                'If there's more than one row to insert
                If intRows > 0 Then
                    'Insert rows in desired direction
                    If UCase(strDirection) = "BELOW" Then
                        objWordApplication.Selection.InsertRowsBelow intRows
                        'Leave selection on row where marker was found
                        objWordApplication.Selection.GoToPrevious What:=wdGoToLine
                        InsertTableRowsMarker = True
                    ElseIf UCase(strDirection) = "ABOVE" Then
                        objWordApplication.Selection.InsertRowsAbove intRows
                        'Leave selection on row where marker was found
                        objWordApplication.Selection.GoToNext What:=wdGoToLine
                        InsertTableRowsMarker = True
                    Else
                        InsertTableRowsMarker = False
                    End If
                Else 'No need to add rows, just signal success
                    InsertTableRowsMarker = True
                End If
                
            Else 'try to find the text above the current selection position
                .Selection.Find.Forward = False
                If .Selection.Find.Execute() = True Then
                    'Clear marker
                    .Selection.TypeText Text:=Chr(32)
                    
                    'If there's more than one row to insert
                    If intRows > 0 Then
                        'Insert rows in desired direction
                        If UCase(strDirection) = "BELOW" Then
                            objWordApplication.Selection.InsertRowsBelow intRows
                            'Leave selection on row where marker was found
                            objWordApplication.Selection.GoToPrevious What:=wdGoToLine
                            InsertTableRowsMarker = True
                        ElseIf UCase(strDirection) = "ABOVE" Then
                            objWordApplication.Selection.InsertRowsAbove intRows
                            'Leave selection on row where marker was found
                            objWordApplication.Selection.GoToNext What:=wdGoToLine
                            InsertTableRowsMarker = True
                        Else
                            InsertTableRowsMarker = False
                        End If
                    End If
                    
                Else
                    'Marker couldn't be found
                    InsertTableRowsMarker = False
                End If
            End If
        End With
        
        'Markers other than bookmark and tag are not handled
        Case Else
            InsertTableRowsMarker = False
        
    End Select

End Function

Function fInsertTable(ByRef objWordApplication As Word.Application, _
                      strTemplateName As String, _
                      strTableBlockName As String, _
                      arrRowHeaders() As String, _
                      arrColumnHeaders() As String) As Boolean
' --------------------------------------------------------------
' Description:
'
' This function inserts a table with variable number of rows
' and columns. If the operation is successful, the function returns
' true. The table is inserted at the current cursor position.
'
' Arguments:
'
' objWordApplication (Object) = Object referencing the Word application instance
' strTemplateName (String) = name of the word template containing
'                           block to be inserted
' strTableBlockName (String) = Name of the table block to insert
' arrRowHeaders (Array) = Array containing the names of the row headers
' arrColumnHeaders (Array) = Array containing the names of the column headers
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 01/02/12      Carlos Gamez        Initial version
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler

    Dim intRows As Long
    Dim intColumns As Long
    Dim i As Long
    Dim strRowHeader As Variant
    Dim strColumnHeader As Variant
    
    intRows = UBound(arrRowHeaders)
    intColumns = UBound(arrColumnHeaders)

    With objWordApplication
    
        'If the selection is inside a table, move until it's out
        Do While .Selection.Information(wdWithInTable)
            .Selection.MoveDown
        Loop
        
        'Insert a new line
        .Selection.TypeText Text:=vbCrLf
        
        'Insert dynamic table block
        .Templates(strTemplateName). _
            BuildingBlockEntries(strTableBlockName).Insert _
            Where:=.Selection.Range, _
            RichText:=True
        
        'Make sure selection is inside inserted table
        Do While Not .Selection.Information(wdWithInTable)
            .Selection.MoveUp
        Loop
        
        'Move selection to second row / first column table cell
        .Selection.Tables(1).Cell(Row:=2, Column:=1).Range.Select
        .Selection.Collapse
        
        'Insert Rows
        If intRows > 0 Then .Selection.InsertRowsBelow intRows
        
        'Move selection to second row / first column table cell
        .Selection.Tables(1).Cell(Row:=2, Column:=1).Range.Select
        .Selection.Collapse
        
        'Insert Row Header Names
        For Each strRowHeader In arrRowHeaders
            .Selection.TypeText strRowHeader
            .Selection.MoveDown
        Next strRowHeader
        
        'Move selection to first row / second column table cell
        .Selection.Tables(1).Cell(Row:=1, Column:=2).Range.Select
        .Selection.Collapse
        
        'Insert Columns
        For i = 1 To intColumns
            .Selection.InsertColumnsRight
        Next i
        
        'Move selection to first row / second column table cell
        .Selection.Tables(1).Cell(Row:=1, Column:=2).Range.Select
        .Selection.Collapse
        
        'Insert Column Header Names
        For Each strColumnHeader In arrColumnHeaders
            .Selection.TypeText strColumnHeader
            .Selection.MoveRight
        Next strColumnHeader
        
        'Format table to 100% of page witdth
        .Selection.Tables(1).PreferredWidthType = wdPreferredWidthPercent
        .Selection.Tables(1).PreferredWidth = 100
        
    End With

Exit_ErrHandler:
    Exit Function

ErrHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    fInsertTable = False
    
    Resume Exit_ErrHandler
    Resume

End Function

Public Function fInsertPhoto(ByRef objWordApp As Word.Application, _
                             strFileName As String, _
                             Optional strSize As String = vbNullString, _
                             Optional wdWrappingOption As WdWrapType = wdWrapSquare, _
                             Optional boolConvertToInline As Boolean = True) As Word.Shape
' --------------------------------------------------------------
' Comments:
'   Inserts a new photo/image on the document
'
' Returns:
'   A reference to the photo Shape object
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 19/02/13      Carlos Gamez        Initial version
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler
    
    Dim shpImage As Word.Shape

    With objWordApp
        
        Set shpImage = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, _
                                         Filename:=strFileName, _
                                         LinkToFile:=False, _
                                         SaveWithDocument:=True)
        
        With shpImage
            
            'Make inline if requested
            If boolConvertToInline Then .ConvertToInlineShape
            
            'Resize if requested
            If strSize <> vbNullString Then
                .LockAspectRatio = msoTrue
                .Height = CSng(strSize)
            End If
        
            'Apply wrapping format
            .WrapFormat.Type = wdWrappingOption
        
        End With
        
    End With
    
    Set fInsertPhoto = shpImage

Exit_ErrHandler:
    Set shpImage = Nothing
    Exit Function

ErrHandler:
    Select Case Err.Number
        'Image file couldn't be found
        Case Is = 5152
            Set fInsertPhoto = Nothing
            ThisWorkbook.Activate
            sDisplayUnexpectedError Err.Number, "The image [" & strFileName & "] was not found." & vbCrLf & _
                    "The application will continue to run and place " & _
                    "a message where the photo would normally be inserted.", mstrMODULE
            objWordApp.Activate
        'Wrapping option could not be applied
        Case Is = 438
            Resume Next
    'Defalut error handle
    Case Else
        sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    End Select
    
    Resume Exit_ErrHandler
    Resume

End Function

Public Sub InsertTask(ByRef objWordApp As Word.Application, _
                      rngTask As Range, strTaskCase As String, _
                      Optional strSize As String = "80", _
                      Optional strTaskNo As String = vbNullString)
' --------------------------------------------------------------
' Comments:
'   Inserts a task in the document
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler
 
    Dim currRow As Integer
    Dim currColumn As Integer
    Dim currPicture As Word.Shape
    Dim strAcceptableConditions As String
    Dim strRecommendedAction As String
    Const strAC As String = "AC"
    Const strRA As String = "RA"
     
    strAcceptableConditions = ""
    strRecommendedAction = ""
     
    With objWordApp
    
        'Insert task ID if applicable
        If fGetSetting("InsertTaskNumberOption") = 1 Then
            If strTaskNo <> vbNullString Then
                currRow = .Selection.Cells(1).RowIndex
                currColumn = .Selection.Cells(1).ColumnIndex
                If currColumn > 1 Then currColumn = 1
                .Selection.Tables(1).Cell(Row:=currRow, Column:=currColumn).Range.Select
                .Selection.Collapse
                .Selection.TypeText Text:=strTaskNo
                .Selection.Tables(1).Cell(Row:=currRow, Column:=currColumn + 1).Range.Select
                .Selection.Collapse
            End If
        Else
            '.Selection.Tables(1).Cell(Row:=.Selection.Cells(1).RowIndex, _
                                    Column:=.Selection.Cells(1).ColumnIndex + 1 _
                                    ).Range.Select
        End If
     
        'Insert task
        If Left(strTaskCase, 1) = "1" Then
            .Selection.TypeText Text:=rngTask(, 1).Text
        End If
        
        'Insert in-task photo if applicable
        If Right(strTaskCase, 1) = "1" Then
            .Selection.TypeText vbCrLf
            Set currPicture = fInsertPhoto(objWordApp, ImageFolder & rngTask(, 4).Text, strSize)
            If currPicture Is Nothing Then
                'Place a message on the document
                .Selection.TypeText "[NOT FOUND] : " & ImageFolder & rngTask(, 4).Text
            Else
                'Position reduced photo on the top right-hand side of the cell
                currPicture.RelativeHorizontalPosition = wdRelativeHorizontalPositionRightMarginArea
                currPicture.RelativeVerticalPosition = wdRelativeVerticalPositionTopMarginArea
                currPicture.Left = wdShapeRight
                currPicture.Top = wdShapeTop
            End If
        End If
        
        'Insert acceptable conditions and recommended actions
        If Mid(strTaskCase, 2, 1) = "1" Then
            
            'If Acceptable Conditions exist
            If InStr(rngTask(, 2).Text, strAC) > 0 Then
                strAcceptableConditions = Trim(Mid(rngTask(, 2).Text, InStr(rngTask(, 2).Text, strAC) + Len(strAC) + 1, _
                                        InStr(rngTask(, 2).Text, Chr(13)) - InStr(rngTask(, 2).Text, strAC) - Len(strAC)))
            End If
            
            'If Recommended Action exists
            If InStr(rngTask(, 2).Text, strRA) > 0 Then
                strRecommendedAction = Trim(Mid(rngTask(, 2).Text, InStr(rngTask(, 2).Text, strRA) + Len(strRA) + 1, _
                    Len(rngTask(, 2).Text) - InStr(rngTask(, 2).Text, strRA) + 1))
            End If
            
            If fGetSetting("BundleAcceptableConditionsWithTask") = 1 Then
                .Selection.TypeText Text:=vbCrLf
                '.Selection.Font.Italic = wdToggle
                '.Selection.Font.Bold = wdToggle
                '.Selection.TypeText Text:=strAC & ": "
                '.Selection.Font.Bold = wdToggle
                .Selection.TypeText Text:=strAcceptableConditions
                '.Selection.Font.Bold = wdToggle
                '.Selection.TypeText Text:=strRA & ": "
                '.Selection.Font.Bold = wdToggle
                .Selection.TypeText Text:=strRecommendedAction
                '.Selection.Font.Italic = wdToggle
            Else
                If strAcceptableConditions <> "" Then
                    'Insert Acceptable Conditions on the column to the right
                    'of the task.
                    .Selection.Tables(1).Cell(Row:=.Selection.Cells(1).RowIndex, _
                                              Column:=.Selection.Cells(1).ColumnIndex + 1 _
                                              ).Range.Select
                    .Selection.TypeText Text:=strAcceptableConditions
                    .Selection.Tables(1).Cell(Row:=.Selection.Cells(1).RowIndex, _
                                              Column:=.Selection.Cells(1).ColumnIndex - 1 _
                                              ).Range.Select
                End If
                
                If strRecommendedAction <> "" Then
                    'Insert Corrective Action on the column to the right
                    'of the Accepatble Conditions
                    .Selection.Tables(1).Cell(Row:=.Selection.Cells(1).RowIndex, _
                                              Column:=.Selection.Cells(1).ColumnIndex + 2 _
                                              ).Range.Select
                    .Selection.TypeText Text:=strRecommendedAction
                    .Selection.Tables(1).Cell(Row:=.Selection.Cells(1).RowIndex, _
                                              Column:=.Selection.Cells(1).ColumnIndex - 2 _
                                              ).Range.Select
                End If
            End If
        End If
            
        'Insert warning label/photo
        If Mid(strTaskCase, 3, 1) = "1" Then
        
            If fGetSetting("BundleAcceptableConditionsWithTask") = 1 Then
                .Selection.MoveStart Unit:=wdCell, Count:=-1
                .Selection.Collapse
            Else
                .Selection.Collapse
            End If
            
            .Selection.TypeText Text:=vbCrLf & vbCrLf
            .Selection.MoveStart Unit:=wdLine, Count:=-2
            .Selection.Collapse
            .Selection.InlineShapes.AddPicture Filename:=ImageFolder & rngTask(, 3).Text, _
                                                       LinkToFile:=False, _
                                                       SaveWithDocument:=True
        End If
        
    End With
 
Exit_ErrHandler:
    Set currPicture = Nothing
    Exit Sub
 
ErrHandler:
    Select Case Err.Number
    Case Is = 5941
        sDisplayUnexpectedError Err.Number, "There are not enough columns on the task table to insert the content " & _
                                "as per the selected settings. The document will not be generated correctly.", mstrMODULE
    Case Else
        sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    End Select
    
    Resume Exit_ErrHandler
    Resume
    
End Sub

Function fInsertBlankTable(ByRef objWordApplication As Word.Application, _
                      strTemplateName As String, _
                      strTableBlockName As String, _
                      Optional intRows As Long = 2, _
                      Optional intColumns As Long = 2) As Boolean
' --------------------------------------------------------------
' Description:
'
' This function inserts a blank table with variable number of rows
' and columns. If the operation is successful, the function returns
' true. The table is inserted at the current cursor position.
'
' Arguments:
'
' objWordApplication (Object) = Object referencing the Word application instance
' strTemplateName (String) = name of the word template containing
'                           block to be inserted
' strTableBlockName (String) = Name of the table block to insert
' intRows (Long) = Number of rows
' intColumns (Long) = Number of columns
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 11/09/12      Carlos Gamez        Initial version
' 26/01/2018    Carlos Gamez        Open source release

On Error GoTo ErrHandler

    Dim i As Long

    With objWordApplication
    
        'If the selection is inside a table, move until it's out
        Do While .Selection.Information(wdWithInTable)
            .Selection.MoveDown
        Loop
        
        'Insert a new line
        .Selection.TypeText Text:=vbCrLf
        
        'Insert dynamic table block
        .Templates(strTemplateName). _
            BuildingBlockEntries(strTableBlockName).Insert _
            Where:=.Selection.Range, _
            RichText:=True
        
        'Make sure selection is inside inserted table
        Do While Not .Selection.Information(wdWithInTable)
            .Selection.MoveUp
        Loop
        
        'Move selection to second row / first column table cell
        .Selection.Tables(1).Cell(Row:=2, Column:=1).Range.Select
        .Selection.Collapse
        
        'Insert Rows
        If intRows > 1 Then .Selection.InsertRowsBelow intRows - 1
        
        
        'Move selection to first row / second column table cell
        .Selection.Tables(1).Cell(Row:=1, Column:=2).Range.Select
        .Selection.Collapse
        
        'Insert Columns
        For i = 2 To intColumns
            .Selection.InsertColumnsRight
        Next i
        
        'Format table to 100% of page witdth
        .Selection.Tables(1).PreferredWidthType = wdPreferredWidthPercent
        .Selection.Tables(1).PreferredWidth = 100
        
    End With
    
    fInsertBlankTable = True

Exit_ErrHandler:
    Exit Function

ErrHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    fInsertBlankTable = False
    
    Resume Exit_ErrHandler
    Resume

End Function

Public Sub sInsertComponentTable(ByRef objWordApplication As Word.Application, _
                                 ByRef objWordDocument As Word.document)
' --------------------------------------------------------------
' Comments:
'   This method inserts the component summary table into the
'   document using the nominated bookmark
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 20/02/13      Carlos Gamez        Initial version, moved out of
'                                   main WIN writing routine
' 27/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler

    Dim intComponentCounter As Integer
    
    Dim colComponents As Collection
    
    Dim rngComponentNumbers As Excel.Range
    Dim rngComponentDescriptions As Excel.Range
    
    Dim bolArrayNonEmpty As Boolean
    
    Dim strEquipmentTableBookmark As String
    Dim strEquipmentTableBlock As String
    Dim strEquipmentNumberColumnName As String
    Dim strEquipmentDescriptionColumnName As String
    Dim arrComponentNumbers() As String
    Dim arrComponentDescriptions() As String
    
    bolArrayNonEmpty = False
    
    'Get the name of the table block in the word template
    strEquipmentTableBlock = fGetSetting("PMEquipmentSummaryTableBlockName")
    
    'Get the names of the columns where the component list is located
    strEquipmentNumberColumnName = fGetSetting("EquipmentSummaryTableHeader1")
    strEquipmentDescriptionColumnName = fGetSetting("EquipmentSummaryTableHeader2")
    
    'Get the name of the bookmark
    strEquipmentTableBookmark = fGetSetting("EquipmentSummaryTableBookmarkName")
    
    'Load the table into a collection
    Set colComponents = fCollectColumns(strEquipmentNumberColumnName & _
                                        "," & strEquipmentDescriptionColumnName, 2, shtWorkInstructions.Name)
    
    If Not colComponents Is Nothing Then
    
        Set rngComponentNumbers = colComponents.Item(1)
        Set rngComponentDescriptions = colComponents.Item(2)
    
        'Collect all components/equuipment into two arrays
        For intComponentCounter = 1 To rngComponentNumbers.Cells.Count
            'Store the current component in the array
            If rngComponentNumbers(intComponentCounter, 1).Text <> "" And _
               rngComponentNumbers(intComponentCounter, 1).Text <> vbNullString Then
        
                ReDim Preserve arrComponentNumbers(intComponentCounter)
                ReDim Preserve arrComponentDescriptions(intComponentCounter)
        
                arrComponentNumbers(intComponentCounter) = rngComponentNumbers(intComponentCounter, 1).value
                arrComponentDescriptions(intComponentCounter) = rngComponentDescriptions(intComponentCounter, 1).value
        
                bolArrayNonEmpty = True
        
            Else
                Exit For
            End If
        Next intComponentCounter
    
        If bolArrayNonEmpty Then 'If the arrays have content
            'If bookmark exists
            If objWordDocument.Bookmarks.Exists(strEquipmentTableBookmark) Then
    
                'Locate appropiate bookmark
                objWordDocument.Bookmarks(strEquipmentTableBookmark).Select
            
                'Insert text
                objWordApplication.Selection.TypeText Text:="This document relates to:"
        
                 EliminateDuplicates arrComponentNumbers, arrComponentNumbers
                 EliminateDuplicates arrComponentDescriptions, arrComponentDescriptions
        
                On Error Resume Next
        
                'Insert table with summary of components of document
                fInsertBlankTable objWordApplication, PMTemplateName, PMBlankTableBlockName, UBound(arrComponentNumbers), 1
        
                If Err.Number = 5941 Then
                    sDisplayUnexpectedError Err.Number, _
                                            "The block [" & PMBlankTableBlockName & "] could not " & _
                                            "be found in the [" & PMTemplateName & "] gallery." & Chr(13) & _
                                            "This part of the document will be skipped. " & _
                                            "Please check the Word Template file and ensure it contains " & Chr(13) & _
                                            "the required elements.", mstrMODULE
                    Err.Clear
        
                Else
            
                    With objWordApplication
                        'Move to first row and first column
                        .Selection.Tables(1).Cell(Row:=1, Column:=1).Range.Select
                        .Selection.Collapse
                
                        'Insert headers
                        .Selection.TypeText Text:=strEquipmentNumberColumnName
                        .Selection.Tables(1).Cell(Row:=1, Column:=2).Range.Select
                        .Selection.Collapse
                        .Selection.TypeText Text:=strEquipmentDescriptionColumnName
                
                        'Size table
                        .Selection.Tables(1).PreferredWidthType = wdPreferredWidthPoints
                        .Selection.Tables(1).PreferredWidth = fGetSetting("DefaultTableWidth") * 28.35
                
                        For intComponentCounter = 1 To UBound(arrComponentNumbers)
                            .Selection.Tables(1).Cell(Row:=intComponentCounter + 1, Column:=1).Range.Select
                            .Selection.Collapse
                            .Selection.TypeText Text:=arrComponentNumbers(intComponentCounter)
                            .Selection.Tables(1).Cell(Row:=intComponentCounter + 1, Column:=2).Range.Select
                            .Selection.Collapse
                            .Selection.TypeText Text:=arrComponentDescriptions(intComponentCounter)
                        Next intComponentCounter
                
                    End With
                End If
                On Error GoTo ErrHandler
            Else
                Err.Raise 514, mstrMODULE, _
                          "The bookmark [" & strEquipmentTableBookmark & "] could not be found in the [" & PMTemplateName & "] gallery." & Chr(13) & _
                          "This part of the document will be skipped. Please check the Word Template file and ensure it contains " & Chr(13) & _
                          "the required elements."
            End If
        End If
    End If

Exit_ErrHandler:
    Exit Sub

ErrHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    
    Resume Exit_ErrHandler
    Resume

End Sub

Public Function fMoveCursorToFirstMulticolumnRow(ByRef objWordApp As Word.Application) As Boolean
' --------------------------------------------------------------
' Comments:
'   This funciton moves the cursor (assumed to be already inside a table)
'   to the first cell (top-left) of the first multi-column row.
'   Returns true if successful.
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 20/02/13      Carlos Gamez        Initial version
' 27/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler
    
    With objWordApp
        
        If .Selection.Information(wdWithInTable) Then
            'Move selection to first row / first column table cell
            .Selection.Tables(1).Cell(Row:=1, Column:=1).Range.Select
            .Selection.Collapse
            
            'Keep moving the cursor down while the row is not multicolumn
            Do While .Selection.Tables(1).Columns.Count < 2
                .Selection.MoveDown
            Loop
            
            fMoveCursorToFirstMulticolumnRow = True

        Else
            fMoveCursorToFirstMulticolumnRow = False
        End If
    
    End With

Exit_ErrHandler:
    Exit Function

ErrHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    
    Resume Exit_ErrHandler
    Resume

End Function
