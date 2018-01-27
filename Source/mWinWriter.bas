Attribute VB_Name = "mWinWriter"
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
' Description:  This module contains routines necessary to produce
'               the Work Instruction documents.
'               Like many other projects that start simple and become more
'               complex over time, this module truned into a good example
'               of how not to do things. It ended up as this huge procedure
'               that works but is extremely fragile and difficult to maintain.
'               I'd like to think that I've continued learning in the last
'               8 years or so and today I would have architected this application
'               better. I'd still like to one day break this down into
'               manageable components.
'               This application served its purpose at the time and I'm
'               releasing it with the hope that it might help other people.
'
' Authors:      Carlos Gamez
'
' Options:
    Option Explicit
    Option Base 1
' **************************************************************
' Module Constant Declarations Follow
' **************************************************************
    Private Const mstrMODULE = "mWinWriter"
' **************************************************************
' Module Dependencies
' **************************************************************
'   mWordFunctions, mErrorHandling, mFileRegistry, mGlobals
' **************************************************************

    '** Word Application and Document variables **
    Dim objWordApplication As Word.Application
    Dim objWordDocument As Word.document
    Dim DocumentFileName As String
    Dim DocumentOpen As Boolean
    Dim GeneratedContent As Boolean
    Dim FrontPageInserted As Boolean
    Dim DocumentCounter As Integer

    '** tblWorkInstructions Variables **
    '**** Document level variables ***
    Dim Aim As String
    Dim Instructions As String
    Dim Location As String
    Dim SiteOrArea As String
    Dim PermitsAndIsolations As String
    Dim PPE As String
    Dim HazardsAndSafetyNotes As String
    Dim OtherSafetyEquipment As String
    Dim References As String
    Dim Resources As String

    Dim ToolRange As Range
    Dim Tools As String
    Dim MaterialsRange As Range
    Dim Materials As String
    Dim SpecialNotesRange As Range
    Dim SpecialNotes As String
    Dim SpecialPermitsRange As Range
    Dim SpecialPermits As String
    Dim SpecialPPERange As Range
    Dim SpecialPPE As String
    Dim SpecialHazardsRange As Range
    Dim SpecialHazards As String
    Dim SpecialReferencesRange As Range
    'Dim SpecialReferences As String
    Dim SpecialReferences As Collection

    Dim DocumentTitle As String
    Dim DocumentNumber As String
    Dim UnitAbbreviation As String
    Dim EquipmentStatusAbbreviation As String
    Dim TradeAbbreviation As String

    '**** Task level variables ***
    Dim DocTitle As PivotItem
    Dim Interval As PivotItem
    Dim Unit As PivotItem
    Dim EquipmentStatus As PivotItem
    Dim Trade As PivotItem
    Dim ComponentDescription As PivotItem
    Dim ComponentNumber As PivotItem
    Dim RowHeaders As PivotItem
    Dim ColumnHeaders As PivotItem
    Dim TaskPhoto As PivotItem
    Dim ComponentPhotoLink As String
    Dim TaskRange As Range
    Dim TaskCase As String
    Dim Task As Range
    Dim TaskDescription As Variant

    '**** Local Variables ****
    Dim boolInitializeVariables As Boolean
    Dim strSelectedTaskTable As String
    Dim rngIntersectTest As Range
    Dim dteStartTime As Date
    Dim dteFinishTime As Date
    Dim strTimeTaken As String
    Dim arrComponentNumbers() As String
    Dim arrComponentDescriptions() As String
    Dim intComponentCounter As Integer
    Dim intTaskCounter As Integer

Public Sub GenerateWordDocument(Optional butRibbon As IRibbonControl)
' --------------------------------------------------------------
' Comments:
'   This routine takes information from a Pivot Table containing the summary of
'   all tasks contained in the tblWorkInstructions. According to a number of settings, it then
'   outputs information to a word document based on a template assigned in one of
'   the settings.
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 03/02/2012    Carlos Gamez        Added functionality to add dynamically sizing
'                                   table below the tasks for multiple components
' 22/05/2012    Carlos Gamez        Incorporated changes made by Adrian to speed up the generation
'                                   of documents, by avoiding unnecessary loops through the pivot
'                                   table.
' 06/07/2012    Carlos Gamez        Included feature that logs generated files on a Log sheet in the
'                                   workbook.
' 25/08/2012    Carlos Gamez        Moved the insertion of the tasks to a function, which now includes
'                                   top photo, in-task photo and warning labels.
' 18/02/2013    Carlos Gamez        Added the capability to use a column in the Pivot Table to be used
'                                   as the title of the document.
' 26/02/2018    Carlos Gamez        Open source release
'
'
Dim strMethod As String
strMethod = "GenerateWordDocument"

On Error GoTo ErrHandler

    'Variables initialization
    boolInitializeVariables = InitializeVariables("GenerateWordDocument")
    If boolInitializeVariables = False Then Exit Sub
    DocumentCounter = 0
    intComponentCounter = 0
    intTaskCounter = 0

    'Refresh the pivot table data
    Excel.Workbooks(ThisWorkbook.Name).RefreshAll

    'Start timer
    dteStartTime = Now

    ' *****************************
    ' *** START OF MAIN ROUTINE ***
    ' *****************************

    'Check that selected Pivot Table exists or that it hasn't changed names
    If DoesPivotTableExist(shtTaskSummary, PivotTableName) = False Then
        Err.Raise 513, mstrMODULE, "The Pivot Table [" & PivotTableName & "] could not be found." & _
                vbCrLf & "Please check your settings."
    End If

    'Start scanning the PivotTable containing the summary of activities
    With shtTaskSummary.PivotTables(PivotTableName)

        '>>Document title grouping level
        For Each DocTitle In .VisibleFields(DocTitleFieldName).PivotItems

                '>>Interval grouping level
                For Each Interval In .VisibleFields(IntervalFieldName).PivotItems
                    
                    'Check intersection, avoid loop if unnecessary
                    If DoRangesIntersect( _
                        .VisibleFields(DocTitleFieldName).PivotItems(DocTitle.Name).DataRange.EntireRow, _
                        .VisibleFields(IntervalFieldName).PivotItems(Interval.Name).DataRange _
                    ) = False Then GoTo NextInterval

                    '>>Interval unit grouping level
                    For Each Unit In .VisibleFields(IntervalUnitsFieldName).PivotItems

                        'Check intersection, avoid loop if unnecessary
                        If DoRangesIntersect( _
                            .VisibleFields(DocTitleFieldName).PivotItems(DocTitle.Name).DataRange.EntireRow, _
                            .VisibleFields(IntervalFieldName).PivotItems(Interval.Name).DataRange.EntireRow, _
                            .VisibleFields(IntervalUnitsFieldName).PivotItems(Unit.Name).DataRange _
                        ) = False Then GoTo NextUnit

                        '>>Equipment Status grouping level
                        For Each EquipmentStatus In .VisibleFields(EquipmentStatusFieldName).PivotItems

                            'Check intersection, avoid loop if unnecessary
                            If DoRangesIntersect( _
                                .VisibleFields(DocTitleFieldName).PivotItems(DocTitle.Name).DataRange.EntireRow, _
                                .VisibleFields(IntervalFieldName).PivotItems(Interval.Name).DataRange.EntireRow, _
                                .VisibleFields(IntervalUnitsFieldName).PivotItems(Unit.Name).DataRange.EntireRow, _
                                .VisibleFields(EquipmentStatusFieldName).PivotItems(EquipmentStatus.Name).DataRange _
                            ) = False Then GoTo NextEquipmentStatus

                            '>> Trade grouping level
                            For Each Trade In .VisibleFields(TradeFieldName).PivotItems

                                'Check intersection, avoid loop if unnecessary
                                If DoRangesIntersect( _
                                    .VisibleFields(DocTitleFieldName).PivotItems(DocTitle.Name).DataRange.EntireRow, _
                                    .VisibleFields(IntervalFieldName).PivotItems(Interval.Name).DataRange.EntireRow, _
                                    .VisibleFields(IntervalUnitsFieldName).PivotItems(Unit.Name).DataRange.EntireRow, _
                                    .VisibleFields(EquipmentStatusFieldName).PivotItems(EquipmentStatus.Name).DataRange.EntireRow, _
                                    .VisibleFields(TradeFieldName).PivotItems(Trade.Name).DataRange _
                                ) = False Then GoTo NextTrade

                                'Initialize document cycle variables
                                DocumentTitle = ""
                                UnitAbbreviation = ""
                                EquipmentStatusAbbreviation = ""
                                TradeAbbreviation = ""
                                DocumentNumber = ""
                                FrontPageInserted = False
                                GeneratedContent = False

                                'At this level, generate the document title based on frequency, status and trade
                                
                                'Get abbreviations
                                UnitAbbreviation = Excel.WorksheetFunction.VLookup(Unit.Name, Application.Range("IntervalUnitsAbbreviations"), 2, False)
                                
                                EquipmentStatusAbbreviation = Excel.WorksheetFunction.VLookup(EquipmentStatus.Name, Application.Range("EquipmentStatusAbbreviation"), 2, False)
                                
                                TradeAbbreviation = Excel.WorksheetFunction.VLookup(Trade.Name, Application.Range("TradeAbbreviations"), 2, False)
                                
                                'Check prefixes
                                If CInt(IncludeNumberPrefix) <> 1 Then DocumentNumberPrefix = ""
                                    If CInt(IncludeNumberSuffix) <> 1 Then DocumentNumberSuffix = ""

                                    'Build Document Number
                                    DocumentNumber = DocumentNumberPrefix & EquipmentNumber & "-" & Interval.Name & UnitAbbreviation & "-" & _
                                                TradeAbbreviation & "-" & EquipmentStatusAbbreviation & "-" & DocumentNumberSuffix & _
                                                CStr(CInt(DocumentConsecutiveStartingNumber) + DocumentCounter)

                                    'Build Document Title
                                    If CInt(IncludeDescription) = 4 Then
                                        DocumentTitle = CStr(DocTitle.value)
                                    Else
                                        DocumentTitle = EquipmentName & " " & Interval & " " & Unit & _
                                                    " " & Trade & " " & EquipmentStatus & " Inspection"
                                    End If

                                    'Build the DocumentFileName variable
                                    If CInt(IncludeDescription) = 1 Then
                                        DocumentFileName = DocumentNumber & " - " & DocumentTitle
                                    ElseIf CInt(IncludeDescription) = 2 Then
                                        DocumentFileName = DocumentNumber
                                    ElseIf CInt(IncludeDescription) = 3 Then
                                        DocumentFileName = DocumentTitle
                                    ElseIf CInt(IncludeDescription) = 4 Then
                                        DocumentFileName = DocumentTitle
                                    End If

                                    '>> Component Number grouping level
                                    For Each ComponentNumber In .VisibleFields(ComponentNumberFieldName).PivotItems

                                        'Check intersection, avoid loop if unnecessary
                                        If DoRangesIntersect( _
                                            .VisibleFields(DocTitleFieldName).PivotItems(DocTitle.Name).DataRange.EntireRow, _
                                            .VisibleFields(IntervalFieldName).PivotItems(Interval.Name).DataRange.EntireRow, _
                                            .VisibleFields(IntervalUnitsFieldName).PivotItems(Unit.Name).DataRange.EntireRow, _
                                            .VisibleFields(EquipmentStatusFieldName).PivotItems(EquipmentStatus.Name).DataRange.EntireRow, _
                                            .VisibleFields(TradeFieldName).PivotItems(Trade.Name).DataRange.EntireRow, _
                                            .VisibleFields(ComponentNumberFieldName).PivotItems(ComponentNumber.Name).DataRange _
                                        ) = False Then GoTo NextComponentNumber

                                        '>> Component Description grouping level
                                        For Each ComponentDescription In .VisibleFields(ComponentDescriptionFieldName).PivotItems

                                            'Check intersection, avoid loop if unnecessary
                                            If DoRangesIntersect( _
                                                .VisibleFields(DocTitleFieldName).PivotItems(DocTitle.Name).DataRange.EntireRow, _
                                                .VisibleFields(IntervalFieldName).PivotItems(Interval.Name).DataRange.EntireRow, _
                                                .VisibleFields(IntervalUnitsFieldName).PivotItems(Unit.Name).DataRange.EntireRow, _
                                                .VisibleFields(EquipmentStatusFieldName).PivotItems(EquipmentStatus.Name).DataRange.EntireRow, _
                                                .VisibleFields(TradeFieldName).PivotItems(Trade.Name).DataRange.EntireRow, _
                                                .VisibleFields(ComponentNumberFieldName).PivotItems(ComponentNumber.Name).DataRange.EntireRow, _
                                                .VisibleFields(ComponentDescriptionFieldName).PivotItems(ComponentDescription.Name).DataRange _
                                            ) = False Then GoTo NextComponentDescription

                                            '>> Row Header grouping level
                                            For Each RowHeaders In .VisibleFields(RowHeadersFieldName).PivotItems

                                                'Check intersection, avoid loop if unnecessary
                                                If DoRangesIntersect( _
                                                    .VisibleFields(DocTitleFieldName).PivotItems(DocTitle.Name).DataRange.EntireRow, _
                                                    .VisibleFields(IntervalFieldName).PivotItems(Interval.Name).DataRange.EntireRow, _
                                                    .VisibleFields(IntervalUnitsFieldName).PivotItems(Unit.Name).DataRange.EntireRow, _
                                                    .VisibleFields(EquipmentStatusFieldName).PivotItems(EquipmentStatus.Name).DataRange.EntireRow, _
                                                    .VisibleFields(TradeFieldName).PivotItems(Trade.Name).DataRange.EntireRow, _
                                                    .VisibleFields(ComponentNumberFieldName).PivotItems(ComponentNumber.Name).DataRange.EntireRow, _
                                                    .VisibleFields(ComponentDescriptionFieldName).PivotItems(ComponentDescription.Name).DataRange.EntireRow, _
                                                    .VisibleFields(RowHeadersFieldName).PivotItems(RowHeaders.Name).DataRange _
                                                ) = False Then GoTo NextRowHeaders

                                                '>> Column Header grouping level
                                                For Each ColumnHeaders In .VisibleFields(ColumnHeadersFieldName).PivotItems

                                                    'Check intersection, avoid loop if unnecessary
                                                    If DoRangesIntersect(.VisibleFields(DocTitleFieldName).PivotItems(DocTitle.Name).DataRange.EntireRow, _
                                                                        .VisibleFields(IntervalFieldName).PivotItems(Interval.Name).DataRange.EntireRow, _
                                                                        .VisibleFields(IntervalUnitsFieldName).PivotItems(Unit.Name).DataRange.EntireRow, _
                                                                        .VisibleFields(EquipmentStatusFieldName).PivotItems(EquipmentStatus.Name).DataRange.EntireRow, _
                                                                        .VisibleFields(TradeFieldName).PivotItems(Trade.Name).DataRange.EntireRow, _
                                                                        .VisibleFields(ComponentNumberFieldName).PivotItems(ComponentNumber.Name).DataRange.EntireRow, _
                                                                        .VisibleFields(ComponentDescriptionFieldName).PivotItems(ComponentDescription.Name).DataRange.EntireRow, _
                                                                        .VisibleFields(RowHeadersFieldName).PivotItems(RowHeaders.Name).DataRange.EntireRow, _
                                                                        .VisibleFields(ColumnHeadersFieldName).PivotItems(ColumnHeaders.Name).DataRange _
                                                    ) = False Then GoTo NextColumnHeaders

                                                    '>> Task Photo grouping level
                                                    For Each TaskPhoto In .VisibleFields(TaskPhotoFieldName).PivotItems

                                                        'Check intersection, avoid loop if unnecessary
                                                        If DoRangesIntersect( _
                                                            .VisibleFields(DocTitleFieldName).PivotItems(DocTitle.Name).DataRange.EntireRow, _
                                                            .VisibleFields(IntervalFieldName).PivotItems(Interval.Name).DataRange.EntireRow, _
                                                            .VisibleFields(IntervalUnitsFieldName).PivotItems(Unit.Name).DataRange.EntireRow, _
                                                            .VisibleFields(EquipmentStatusFieldName).PivotItems(EquipmentStatus.Name).DataRange.EntireRow, _
                                                            .VisibleFields(TradeFieldName).PivotItems(Trade.Name).DataRange.EntireRow, _
                                                            .VisibleFields(ComponentNumberFieldName).PivotItems(ComponentNumber.Name).DataRange.EntireRow, _
                                                            .VisibleFields(ComponentDescriptionFieldName).PivotItems(ComponentDescription.Name).DataRange.EntireRow, _
                                                            .VisibleFields(RowHeadersFieldName).PivotItems(RowHeaders.Name).DataRange.EntireRow, _
                                                            .VisibleFields(ColumnHeadersFieldName).PivotItems(ColumnHeaders.Name).DataRange.EntireRow, _
                                                            .VisibleFields(TaskPhotoFieldName).PivotItems(TaskPhoto.Name).DataRange _
                                                        ) = False Then GoTo NextTaskPhoto
                                                        
                                                        'Find the applicable tasks and conditions for this document, if nothing found, skip to next document
                                                        Set TaskRange = Nothing
                                                        Set TaskRange = Intersect(.VisibleFields(DocTitleFieldName).PivotItems(DocTitle.Name).DataRange.EntireRow, _
                                                                    .VisibleFields(IntervalFieldName).PivotItems(Interval.Name).DataRange.EntireRow, _
                                                                    .VisibleFields(IntervalUnitsFieldName).PivotItems(Unit.Name).DataRange.EntireRow, _
                                                                    .VisibleFields(EquipmentStatusFieldName).PivotItems(EquipmentStatus.Name).DataRange.EntireRow, _
                                                                    .VisibleFields(TradeFieldName).PivotItems(Trade.Name).DataRange.EntireRow, _
                                                                    .VisibleFields(ComponentNumberFieldName).PivotItems(ComponentNumber.Name).DataRange.EntireRow, _
                                                                    .VisibleFields(ComponentDescriptionFieldName).PivotItems(ComponentDescription.Name).DataRange.EntireRow, _
                                                                    .VisibleFields(RowHeadersFieldName).PivotItems(RowHeaders.Name).DataRange.EntireRow, _
                                                                    .VisibleFields(ColumnHeadersFieldName).PivotItems(ColumnHeaders.Name).DataRange.EntireRow, _
                                                                    .VisibleFields(TaskPhotoFieldName).PivotItems(TaskPhoto.Name).DataRange.EntireRow, _
                                                                    .VisibleFields(TaskFieldName).DataRange)

                                                        'If tasks are found
                                                        If Not TaskRange Is Nothing Then
                                                            If DocumentOpen = False Then
                                                                OpenDocument
                                                            End If

                                                            'Signal that content will be created
                                                            GeneratedContent = True

                                                            'Insert front page (once per document)
                                                            If FrontPageInserted = False Then
                                                                InsertFrontPage
                                                                'Signal successful creation of the front page
                                                                FrontPageInserted = True
                                                        
                                                            End If

                                                            'Ensure the document view is PrintLayout
                                                            objWordApplication.ActiveWindow.View.ReadingLayout = False
                                                            objWordApplication.ActiveWindow.View.Type = wdPrintView

                                                            InsertStep
                                                        
                                                        'Finish document writing routine
                                                        End If
NextTaskPhoto:
                                                    Next TaskPhoto
NextColumnHeaders:
                                                Next ColumnHeaders
NextRowHeaders:
                                            Next RowHeaders
NextComponentDescription:
                                        Next ComponentDescription
NextComponentNumber:
                                    Next ComponentNumber

                                    'If tasks were found Save and Close word document, otherwise, discard it
                                    If GeneratedContent = True Then
                                        'Insert component summary table
                                        If fGetSetting("IncludeComponentSummaryTable") = 1 Then
                                            sInsertComponentTable objWordApplication, objWordDocument
                                        End If
                                        
                                        FinaliseDocument
                                    End If

NextTrade:
                            Next Trade

NextEquipmentStatus:
                        Next EquipmentStatus

NextUnit:
                    Next Unit

NextInterval:
                Next Interval
        
NextDocTitle:
        Next DocTitle

    End With

    ' *****************************
    ' **** FINISH MAIN ROUTINE ****
    ' *****************************
    
Exit_ErrHandler:
    'Finish timer
    dteFinishTime = Now
    
    'Calculate total time and form string
    strTimeTaken = CStr(Int(DateDiff("s", dteStartTime, dteFinishTime) / 60) & " minutes and " & DateDiff("s", dteStartTime, dteFinishTime) Mod 60 & " seconds.")
    
    'A bit of housekeeping
    If DocumentOpen = True Then objWordDocument.Close False
    If Not objWordApplication Is Nothing Then objWordApplication.Quit
    
    'Clean objects
    Set objWordApplication = Nothing
    Set objWordDocument = Nothing
    Set ToolRange = Nothing
    Set MaterialsRange = Nothing
    Set SpecialNotesRange = Nothing
    Set SpecialPermitsRange = Nothing
    Set SpecialPPERange = Nothing
    Set SpecialHazardsRange = Nothing
    Set SpecialReferencesRange = Nothing
    Set SpecialReferences = Nothing
    Set DocTitle = Nothing
    Set Interval = Nothing
    Set Unit = Nothing
    Set EquipmentStatus = Nothing
    Set Trade = Nothing
    Set ComponentDescription = Nothing
    Set ComponentNumber = Nothing
    Set RowHeaders = Nothing
    Set ColumnHeaders = Nothing
    Set TaskPhoto = Nothing
    Set TaskRange = Nothing
    Set Task = Nothing
    Set rngIntersectTest = Nothing
    
    
    'Report number of generated documents
    If DocumentCounter > 0 Then
        MsgBox DocumentCounter & " Word documents were generated and saved to:" & vbCrLf & _
                                        FolderToSaveFilesTo & vbCrLf & _
                                        "In total, it took " & strTimeTaken, _
                                        vbInformation, gstrAPP_NAME
    End If
    Exit Sub

ErrHandler:
    Select Case Err.Number
        Case Is = 1004
            Resume Next
        Case Is = 5152
            sDisplayUnexpectedError Err.Number, "Invalid characters have been used for the file name.", mstrMODULE & " : " & strMethod
        Case Else
            sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE & " : " & strMethod
    End Select
    
    Resume Exit_ErrHandler
    Resume

End Sub

Private Function DoRangesIntersect(rngOne As Excel.Range, rngTwo As Excel.Range, _
                                   Optional rngThree As Excel.Range, Optional rngFour As Excel.Range, _
                                   Optional rngFive As Excel.Range, Optional rngSix As Excel.Range, _
                                   Optional rngSeven As Excel.Range, Optional rngEight As Excel.Range, _
                                   Optional rngNine As Excel.Range, Optional rngTen As Excel.Range, _
                                   Optional rngEleven As Excel.Range, Optional rngTwelve As Excel.Range _
                                   ) As Boolean
' --------------------------------------------------------------
' Comments:
'   Recursive function that interstects ranges and returns whether
'   the intersection exists or not
'
' Arguments:
'   rng[One-Ten] (Excel.Range) = The ranges that will be intersected
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
    Dim rngResult As Excel.Range
    
    DoRangesIntersect = False
    
    If rngThree Is Nothing Then
        Set rngResult = Intersect(rngOne, rngTwo)
    ElseIf rngFour Is Nothing Then
        Set rngResult = Intersect(rngOne, rngTwo, rngThree)
    ElseIf rngFive Is Nothing Then
        Set rngResult = Intersect(rngOne, rngTwo, rngThree, rngFour)
    ElseIf rngSix Is Nothing Then
        Set rngResult = Intersect(rngOne, rngTwo, rngThree, rngFour, rngFive)
    ElseIf rngSeven Is Nothing Then
        Set rngResult = Intersect(rngOne, rngTwo, rngThree, rngFour, rngFive, rngSix)
    ElseIf rngEight Is Nothing Then
        Set rngResult = Intersect(rngOne, rngTwo, rngThree, rngFour, rngFive, rngSix, rngSeven)
    ElseIf rngNine Is Nothing Then
        Set rngResult = Intersect(rngOne, rngTwo, rngThree, rngFour, rngFive, rngSix, rngSeven, rngEight)
    ElseIf rngTen Is Nothing Then
        Set rngResult = Intersect(rngOne, rngTwo, rngThree, rngFour, rngFive, rngSix, rngSeven, rngEight, rngNine)
    ElseIf rngEleven Is Nothing Then
        Set rngResult = Intersect(rngOne, rngTwo, rngThree, rngFour, rngFive, rngSix, rngSeven, rngEight, rngNine, rngTen)
    ElseIf rngTwelve Is Nothing Then
        Set rngResult = Intersect(rngOne, rngTwo, rngThree, rngFour, rngFive, rngSix, rngSeven, rngEight, rngNine, rngTen, rngEleven)
    Else
        Set rngResult = Intersect(rngOne, rngTwo, rngThree, rngFour, rngFive, rngSix, rngSeven, rngEight, rngNine, rngTen, rngEleven, rngTwelve)
    End If
    
    DoRangesIntersect = Not rngResult Is Nothing
    
    Set rngResult = Nothing
    
End Function

Private Sub InsertFrontPage()
' --------------------------------------------------------------
' Comments:
'   Inserts the front page of each document
'
' Arguments:
'   None
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler

    'Insert front page photo
    InsertImageMarker objWordApplication, _
                      PMFrontPagePhotoBookmarkName, _
                      "Bookmark", _
                      ImageFolder & EquipmentPhoto, _
                      objWordDocument, _
                      PhotoHeight
    
    'Insert front page block
    InsertBlockMarker objWordApplication, _
                      PMFrontPageSectionBookmarkName, _
                      "Bookmark", _
                      PMTemplateName, _
                      PMFrontPageBlankBlockName, _
                      objWordDocument
    
    'Find and assign value to the SITE OR AREA field
    SiteOrArea = Excel.WorksheetFunction.HLookup(EquipmentStatus.Name, Application.Range("FrontPageContentTable"), 3, 0)
    
    'Insert Front Page Fields
    'Try to insert these fields in case the bookmarks are in the header
    If InsertTextMarker(objWordApplication, PMSiteOrAreaBookmarkName, "Bookmark", SiteOrArea, objWordDocument) = False _
    Or InsertTextMarker(objWordApplication, PMIntervalBookmarkName, "Bookmark", Interval.Name, objWordDocument) = False _
    Or InsertTextMarker(objWordApplication, PMIntervalUnitBookmarkName, "Bookmark", Unit.Name, objWordDocument) = False _
    Or InsertTextMarker(objWordApplication, PMEquipmentStatusBookmarkName, "Bookmark", EquipmentStatus.Name, objWordDocument) = False _
    Or InsertTextMarker(objWordApplication, PMEquipmentNumberBookmarkName, "Bookmark", EquipmentNumber & " " & EquipmentName, objWordDocument) = False _
    Or InsertTextMarker(objWordApplication, PMTradeBookmarkName, "Bookmark", Trade.Name, objWordDocument) = False _
    Or InsertTextMarker(objWordApplication, PMDocumentTitleBookmarkName, "Bookmark", DocumentTitle, objWordDocument) = False _
    Then
        'Finaly try to insert these fields in case they're stated as tags and not bookmarks
        InsertTextMarker objWordApplication, PMSiteOrAreaBookmarkName, "Tag", SiteOrArea, objWordDocument
        InsertTextMarker objWordApplication, PMIntervalBookmarkName, "Tag", Interval.Name, objWordDocument
        InsertTextMarker objWordApplication, PMIntervalUnitBookmarkName, "Tag", Unit.Name, objWordDocument
        InsertTextMarker objWordApplication, PMEquipmentStatusBookmarkName, "Tag", EquipmentStatus.Name, objWordDocument
        InsertTextMarker objWordApplication, PMEquipmentNumberBookmarkName, "Tag", EquipmentNumber & " " & EquipmentName, objWordDocument
        InsertTextMarker objWordApplication, PMTradeBookmarkName, "Tag", Trade.Name, objWordDocument
        InsertTextMarker objWordApplication, PMDocumentTitleBookmarkName, "Tag", DocumentTitle, objWordDocument
    End If
    
    'Find and assign value to the AIM field
    Aim = Excel.WorksheetFunction.Index(Application.Range("AimContentsArea"), Excel.WorksheetFunction.Match(Trade.Name, Application.Range("AimTradeNamesColumn"), 0), Excel.WorksheetFunction.Match(EquipmentStatus.Name, Application.Range("AimEquipmentStatusRow"), 0))
    InsertTextMarker objWordApplication, PMFrontPageAimBookmarkName, "Bookmark", Aim, objWordDocument
    
    'Find and assign value to the INSTRUCTIONS field
    Instructions = Excel.WorksheetFunction.HLookup(EquipmentStatus.Name, Application.Range("FrontPageContentTable"), 2, 0)
    InsertTextMarker objWordApplication, PMFrontPageInstructionsBookmarkName, "Bookmark", Instructions, objWordDocument
    
    'Insert LOCATION field
    Location = Excel.WorksheetFunction.HLookup(EquipmentStatus.Name, Application.Range("FrontPageContentTable"), 4, 0)
    InsertTextMarker objWordApplication, PMFrontPageLocationBookmarkName, "Bookmark", Location, objWordDocument
    
    'Find and assign value to the PERMITS AND ISOLATIONS field
    PermitsAndIsolations = Excel.WorksheetFunction.HLookup(EquipmentStatus.Name, Application.Range("FrontPageContentTable"), 5, 0)
    InsertTextMarker objWordApplication, PMFrontPagePermitsAndIsolationsBookmarkName, "Bookmark", PermitsAndIsolations, objWordDocument
    
    'Find and assign value to the PPE field
    PPE = Excel.WorksheetFunction.HLookup(EquipmentStatus.Name, Application.Range("FrontPageContentTable"), 6, 0)
    InsertTextMarker objWordApplication, PMFrontPagePPEBookmarkName, "Bookmark", PPE, objWordDocument
    
    'Find and assign value to the HAZARDS AND SAFETY NOTES field
    HazardsAndSafetyNotes = Excel.WorksheetFunction.HLookup(EquipmentStatus.Name, Application.Range("FrontPageContentTable"), 7, 0)
    InsertTextMarker objWordApplication, PMFrontPageHazardsAndSafetyNotesBookmarkName, "Bookmark", HazardsAndSafetyNotes, objWordDocument
    
    'Find and assign value to the OTHER SAFETY EQUIPMENT field
    OtherSafetyEquipment = Excel.WorksheetFunction.HLookup(EquipmentStatus.Name, Application.Range("FrontPageContentTable"), 8, 0)
    InsertTextMarker objWordApplication, PMFrontPageOtherSafetyEquipmentBookmarkName, "Bookmark", OtherSafetyEquipment, objWordDocument
    
    'Find and assign value to the REFERENCES field
    References = Excel.WorksheetFunction.HLookup(EquipmentStatus.Name, Application.Range("FrontPageContentTable"), 9, 0)
    InsertTextMarker objWordApplication, PMFrontPageReferencesBookmarkName, "Bookmark", References, objWordDocument
    
    'Find and assign value to the RESOURCES field
    Resources = Excel.WorksheetFunction.HLookup(EquipmentStatus.Name, Application.Range("FrontPageContentTable"), 10, 0) & " " & Trade.Name
    InsertTextMarker objWordApplication, PMFrontPageResourcesBookmarkName, "Bookmark", Resources, objWordDocument
    
    'Find the TOOLS, MATERIALS, SPECIAL NOTES, SPECIAL PERMITS,
    'SPECIAL PPE, SPECIAL HAZARDS and SPECIAL REFERENCES for this document
    'These come from the "FrontPageTools" Pivot Table
    With shtToolsAndMaterialsSummary.PivotTables("FrontPageTools")
        Set ToolRange = Nothing
        Set ToolRange = Intersect(.VisibleFields(IntervalFieldName).PivotItems(Interval.Name).DataRange.EntireRow, _
                    .VisibleFields(IntervalUnitsFieldName).PivotItems(Unit.Name).DataRange.EntireRow, _
                    .VisibleFields(EquipmentStatusFieldName).PivotItems(EquipmentStatus.Name).DataRange.EntireRow, _
                    .VisibleFields(TradeFieldName).PivotItems(Trade.Name).DataRange.EntireRow, _
                    .VisibleFields(TaskToolsFieldName).DataRange)
        Set MaterialsRange = Nothing
        Set MaterialsRange = Intersect(.VisibleFields(IntervalFieldName).PivotItems(Interval.Name).DataRange.EntireRow, _
                    .VisibleFields(IntervalUnitsFieldName).PivotItems(Unit.Name).DataRange.EntireRow, _
                    .VisibleFields(EquipmentStatusFieldName).PivotItems(EquipmentStatus.Name).DataRange.EntireRow, _
                    .VisibleFields(TradeFieldName).PivotItems(Trade.Name).DataRange.EntireRow, _
                    .VisibleFields(TaskMaterialsFieldName).DataRange)
        Set SpecialNotesRange = Nothing
        Set SpecialNotesRange = Intersect(.VisibleFields(IntervalFieldName).PivotItems(Interval.Name).DataRange.EntireRow, _
                    .VisibleFields(IntervalUnitsFieldName).PivotItems(Unit.Name).DataRange.EntireRow, _
                    .VisibleFields(EquipmentStatusFieldName).PivotItems(EquipmentStatus.Name).DataRange.EntireRow, _
                    .VisibleFields(TradeFieldName).PivotItems(Trade.Name).DataRange.EntireRow, _
                    .VisibleFields(TaskSpecialNotesFieldName).DataRange)
        Set SpecialPermitsRange = Nothing
        Set SpecialPermitsRange = Intersect(.VisibleFields(IntervalFieldName).PivotItems(Interval.Name).DataRange.EntireRow, _
                    .VisibleFields(IntervalUnitsFieldName).PivotItems(Unit.Name).DataRange.EntireRow, _
                    .VisibleFields(EquipmentStatusFieldName).PivotItems(EquipmentStatus.Name).DataRange.EntireRow, _
                    .VisibleFields(TradeFieldName).PivotItems(Trade.Name).DataRange.EntireRow, _
                    .VisibleFields(TaskSpecialPermitsFieldName).DataRange)
        Set SpecialPPERange = Nothing
        Set SpecialPPERange = Intersect(.VisibleFields(IntervalFieldName).PivotItems(Interval.Name).DataRange.EntireRow, _
                    .VisibleFields(IntervalUnitsFieldName).PivotItems(Unit.Name).DataRange.EntireRow, _
                    .VisibleFields(EquipmentStatusFieldName).PivotItems(EquipmentStatus.Name).DataRange.EntireRow, _
                    .VisibleFields(TradeFieldName).PivotItems(Trade.Name).DataRange.EntireRow, _
                    .VisibleFields(TaskSpecialPPEFieldName).DataRange)
        Set SpecialHazardsRange = Nothing
        Set SpecialHazardsRange = Intersect(.VisibleFields(IntervalFieldName).PivotItems(Interval.Name).DataRange.EntireRow, _
                    .VisibleFields(IntervalUnitsFieldName).PivotItems(Unit.Name).DataRange.EntireRow, _
                    .VisibleFields(EquipmentStatusFieldName).PivotItems(EquipmentStatus.Name).DataRange.EntireRow, _
                    .VisibleFields(TradeFieldName).PivotItems(Trade.Name).DataRange.EntireRow, _
                    .VisibleFields(TaskSpecialHazardsFieldName).DataRange)
        Set SpecialReferencesRange = Nothing
        Set SpecialReferencesRange = Intersect(.VisibleFields(IntervalFieldName).PivotItems(Interval.Name).DataRange.EntireRow, _
                    .VisibleFields(IntervalUnitsFieldName).PivotItems(Unit.Name).DataRange.EntireRow, _
                    .VisibleFields(EquipmentStatusFieldName).PivotItems(EquipmentStatus.Name).DataRange.EntireRow, _
                    .VisibleFields(TradeFieldName).PivotItems(Trade.Name).DataRange.EntireRow, _
                    .VisibleFields(TaskSpecialReferencesFieldName).DataRange)
    End With
    
    'Convert ranges to strings and insert in document
    Tools = ContentFromRange(ToolRange, vbCrLf)
    InsertTextMarker objWordApplication, PMFrontPageToolsBookmarkName, "Bookmark", Tools, objWordDocument
    
    Materials = ContentFromRange(MaterialsRange, vbCrLf)
    InsertTextMarker objWordApplication, PMFrontPagePartsAndMaterialsBookmarkName, "Bookmark", Materials, objWordDocument
    
    SpecialNotes = ContentFromRange(SpecialNotesRange, vbCrLf)
    InsertTextMarker objWordApplication, PMFrontPageSpecialNotesBookmarkName, "Bookmark", SpecialNotes, objWordDocument
    
    SpecialPermits = ContentFromRange(SpecialPermitsRange, vbCrLf)
    InsertTextMarker objWordApplication, PMFrontPagePermitsAndIsolationsBookmarkName, "Bookmark", SpecialPermits, objWordDocument
    
    SpecialPPE = ContentFromRange(SpecialPPERange, vbCrLf)
    InsertTextMarker objWordApplication, PMFrontPagePPEBookmarkName, "Bookmark", SpecialPPE, objWordDocument
    
    SpecialHazards = ContentFromRange(SpecialHazardsRange, vbCrLf)
    InsertTextMarker objWordApplication, PMFrontPageHazardsAndSafetyNotesBookmarkName, "Bookmark", SpecialHazards, objWordDocument
    
    'Code to insert references just as text
    'SpecialReferences = ContentFromRange(SpecialReferencesRange, vbCrLf)
    'InsertTextMarker objWordApplication, PMFrontPageReferencesBookmarkName, "Bookmark", SpecialReferences, objWordDocument
    
    'Code to insert references as hyperlinks
    Set SpecialReferences = HyperlinksFromRange(SpecialReferencesRange, ManualsFolder)
    InsertCollectionMarker objWordApplication, PMFrontPageReferencesBookmarkName, "Bookmark", ctHyperlink, SpecialReferences, objWordDocument
    
    'Insert tblWorkInstructions Author
    InsertTextMarker objWordApplication, PMRESDAuthorBookmarkName, "Bookmark", DocumentAuthor, objWordDocument
    
Exit_ErrHandler:
    Exit Sub

ErrHandler:
    Select Case Err.Number
        Case Is = 5941
            MsgBox "The block " & PMFrontPageBlankBlockName & " could not be found in the " & PMTemplateName & " gallery." & Chr(13) & _
                    "This part of the document will be skipped. Please check the Word Template file and ensure it contains " & Chr(13) & _
                    "the required elements.", vbCritical, gstrAPP_NAME
            Err.Clear
            Resume Next

        Case Else
            sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    End Select
    
    Resume Exit_ErrHandler
    Resume
    
End Sub

Private Sub InsertStep()
' --------------------------------------------------------------
' Comments:
'   Inserts a new step block with tasks in it
'
' Arguments:
'   None
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
    'Decide which task table to insert
    If TaskPhoto.Name <> "" And TaskPhoto.Name <> "(blank)" Then
        strSelectedTaskTable = PMInstructionBlockName
    Else
        strSelectedTaskTable = PMInstructionTasksOnlyBlockName
    End If
    
    'Insert a new task block table
    If InsertBlockMarker(objWordApplication, PMSectionBookmarkName, _
        "Bookmark", PMTemplateName, strSelectedTaskTable, objWordDocument) = True Then
    
        'Place cursor in first multicolumn row
        fMoveCursorToFirstMulticolumnRow objWordApplication
        
        'Instruction writing block
        With objWordApplication
        
            intTaskCounter = 0
            
            'Insert Equipment Photo
            If TaskPhoto.Name <> "" And TaskPhoto.Name <> "(blank)" Then
                ComponentPhotoLink = ImageFolder & TaskPhoto.Name
                If InsertImageMarker(objWordApplication, InsertEquipmentPhotoMarker, "Tag", ComponentPhotoLink, objWordDocument, PhotoHeight) = False Then
                    'Insert text in document explaining which photo could not be found
                    .Selection.TypeText Text:="Image " & ComponentPhotoLink & " could not be found."
                End If
            Else
                InsertTextMarker objWordApplication, InsertEquipmentPhotoMarker, "Tag", Chr(32), objWordDocument
            End If
            
            'Insert Equipment/Component Number or Functional Location in the table's header
            If ComponentNumber.Name <> "" And ComponentNumber.Name <> "(blank)" Then
                If InsertTextMarker(objWordApplication, FunctionalLocationMarker, "Tag", ComponentNumber.Name) = False Then
                    Err.Raise 515, mstrMODULE, _
                                "Error During Tasks Insertion:" & vbCrLf & _
                                "The Tag [" & FunctionalLocationMarker & "] could not be found on the template document." & vbCrLf & _
                                "Verify that it exists or the name has been entered correctly in the Settings." & vbCrLf & _
                                "The document cannot be generated correctly."
                End If
            Else
                InsertTextMarker objWordApplication, FunctionalLocationMarker, "Tag", Chr(32), objWordDocument
            End If
        
            'Insert the Equipment/Component Description in the table's header
            If ComponentDescription.Name <> "" And ComponentDescription.Name <> "(blank)" Then
                If InsertTextMarker(objWordApplication, InsertEquipmentNameMarker, "Tag", ComponentDescription.Name, objWordDocument) = False Then
                    Err.Raise 515, mstrMODULE, _
                                "Error During Tasks Insertion:" & vbCrLf & _
                                "The Tag [" & InsertEquipmentNameMarker & "] could not be found on the template document." & vbCrLf & _
                                "Verify that it exists or the name has been entered correctly in the Settings." & vbCrLf & _
                                "The document cannot be generated correctly."
                End If
            Else
                InsertTextMarker objWordApplication, InsertEquipmentNameMarker, "Tag", Chr(32), objWordDocument
            End If
        
            'Insert rows in table
            If InsertTableRowsMarker(objWordApplication, InsertInstructionMarker, "Tag", TaskRange.Rows.Count, objWordDocument) = False Then
                Err.Raise 515, mstrMODULE, _
                            "Error During Tasks Insertion:" & vbCrLf & _
                            "The Tag [" & InsertInstructionMarker & "] could not be found on the template document." & vbCrLf & _
                            "Verify that it exists or the name has been entered correctly in the Settings." & vbCrLf & _
                            "The document cannot be generated correctly."
            End If
            
            'Insert task group
            For Each Task In TaskRange
                'Build a TaskCase code depending on the content
                If Task(, 1).Text <> "" And Task(, 1).Text <> "(blank)" Then TaskCase = "1" Else TaskCase = "0"
                If Task(, 2).Text <> "" And Task(, 2).Text <> "(blank)" Then TaskCase = TaskCase & "1" Else TaskCase = TaskCase & "0"
                If Task(, 3).Text <> "" And Task(, 3).Text <> "(blank)" Then TaskCase = TaskCase & "1" Else TaskCase = TaskCase & "0"
                If Task(, 4).Text <> "" And Task(, 4).Text <> "(blank)" Then TaskCase = TaskCase & "1" Else TaskCase = TaskCase & "0"
                
                intTaskCounter = intTaskCounter + 1
                
                InsertTask objWordApplication, Task, TaskCase, TaskPhotoHeight, CStr(intTaskCounter)
                
                'Move to next row
                .Selection.Tables(1).Cell(Row:=.Selection.Cells(1).RowIndex + 1, _
                                        Column:=.Selection.Cells(1).ColumnIndex _
                                        ).Range.Select
            Next Task
            
            'Size table
            .Selection.Tables(1).PreferredWidth = CSng(fGetSetting("DefaultTableWidth")) * 28.35
        'Finish instruction writing block
        End With
    End If
    
    'If multiple component table needs to be inserted and there's content
    If (RowHeaders.Name <> vbNullString And RowHeaders.Name <> "(blank)") Or _
       (ColumnHeaders.Name <> vbNullString And ColumnHeaders.Name <> "(blank)") Then
        
        Dim arrRowHeaders() As String
        Dim arrColumnHeaders() As String
        arrRowHeaders = Split(RowHeaders.value, ",")
        arrColumnHeaders = Split(ColumnHeaders.value, ",")
        If UBound(arrRowHeaders) > 0 And UBound(arrColumnHeaders) > 0 Then
            fInsertTable objWordApplication, PMTemplateName, PMBlankTableBlockName, arrRowHeaders, arrColumnHeaders
        End If
        
    End If
End Sub

Private Sub FinaliseDocument()
' --------------------------------------------------------------
' Comments:
'   Finalises the document and saves it or discards it
'
' Arguments:
'   None
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
    'Save file
    objWordDocument.SaveAs (FolderToSaveFilesTo & DocumentFileName & ".docx")
    
    'Update fields
    'Call UpdateAllFields(objWordDocument)
    objWordDocument.Windows(1).View.Type = wdPrintView
    
    'Open Footer to update fields
    objWordDocument.Windows(1).View.SeekView = wdSeekPrimaryFooter
    objWordApplication.Selection.WholeStory
    objWordApplication.Selection.fields.Update
    
    'Close footer, go back to main document
    objWordDocument.Windows(1).View.SeekView = wdSeekMainDocument
    
    'If Screen Updating option has been turned off, turn on again
    If objWordApplication.ScreenUpdating = False Then objWordApplication.ScreenUpdating = True
    
    'Assign document properties (metadata)
    objWordDocument.BuiltinDocumentProperties(wdPropertyAuthor).value = DocumentAuthor
    objWordDocument.BuiltinDocumentProperties(wdPropertyTitle).value = DocumentTitle
    objWordDocument.BuiltinDocumentProperties(wdPropertySubject).value = "tblWorkInstructions Analysis / Preventive Maintenance Documentation"
    objWordDocument.BuiltinDocumentProperties(wdPropertyCompany).value = "Assetivity"
    objWordDocument.BuiltinDocumentProperties(wdPropertyComments).value = "Generated by: " & gstrAPP_NAME & vbCrLf & _
                                            "Version :" & gstrAPP_VERSION & vbCrLf & _
                                            "Copyright of Assetivity Pty. Ltd."
    
    'Save field updates, increase document counter and close it
    objWordDocument.Save
    DocumentCounter = DocumentCounter + 1
    objWordDocument.Close
    
    'Register the document in the log sheet
    If fGetSetting("LogFileCreation") = 1 Then
        sRegisterFile DocumentFileName, CStr(Now), DocumentAuthor, FolderToSaveFilesTo, PMTemplateName
    End If
    
    'Signal that document has been closed
    DocumentOpen = False
End Sub

Private Sub OpenDocument()
' --------------------------------------------------------------
' Comments:
'   Opens a new document in preparation for content insertion
'
' Arguments:
'   None
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
    'Check if word is running, if not create a new instance
    If IsAppRunning("Word.Application") Then
        Set objWordApplication = GetObject(, "Word.Application")
    Else
        Set objWordApplication = CreateObject("Word.Application")
    End If
    
    'Get document ready to start inserting content
    With objWordApplication
        .Visible = True
        .WindowState = wdWindowStateMaximize
        .Activate
        Set objWordDocument = .Documents.Add(Template:=PMTemplateName)
        .Caption = DocumentTitle
        If Not CInt(ScreenUpdatingOption) = 1 Then
            .ScreenUpdating = False
        Else
            .ScreenUpdating = True
        End If
    End With
    
    'Signal that document is open
    DocumentOpen = True
End Sub
