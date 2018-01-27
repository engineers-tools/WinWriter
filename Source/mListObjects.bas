Attribute VB_Name = "mListObjects"
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
' Description: This model contains functions to manage and control ListObjects
'              or Tables
'
' Authors:      Carlos Gamez
'
Option Explicit

' **************************************************************
' Module Declarations Follow
' **************************************************************
' *** CONSTANTS ***
Private Const mstrMODULE As String = "mListObjects"
Private mstrColumnsNotFound As String

Public Function fCollectColumns(strColumnsList As String, Optional lngHeadersRow As Long = 1, _
                                Optional strWorksheet As String = vbNullString) As Collection
' --------------------------------------------------------------
' Comments:
'   This function finds the columns nominated in the strColumnsList
'   and returns them in a collection containing a range object
'   for each of the columns requested
'
' Arguments:
'   strColumnList (String) = A list of comma-separated column names
'   lngHeadersRow (Long)   = The row number where the columns names
'                            are. Defaults to 1
'   strWorksheet  (String) = The name of the worksheet to look for.
'                            If no worksheet is specified it defaults
'                            to looking in the current sheet.
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 23/08/2012    Carlos Gamez        Initial version
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler

    Dim i As Integer
    Dim lngColumn As Long
    Dim lngRow As Long
    Dim strElement As String
    Dim rngSearch As Excel.Range
    Dim rngColumn As Variant
    Dim colRanges As Collection
    
    i = 1
    Set colRanges = New Collection
    mstrColumnsNotFound = vbNullString

    With ActiveWorkbook
        Do
            strElement = fExtractFromList(strColumnsList, ",", i)
            If strElement <> vbNullString Then
                If strWorksheet = vbNullString Then
                    If Not .ActiveSheet Is Nothing Then
                        lngColumn = fGetColumnNumber(strElement, lngHeadersRow)
                        If lngColumn <> 0 Then
                            lngRow = fGetLastNonEmptyRow(lngColumn)
                            Set rngSearch = Range(.ActiveSheet.Cells(1, lngColumn), _
                                                  .ActiveSheet.Cells(3500, lngColumn))
                            colRanges.Add rngSearch, strElement
                        Else
                            If mstrColumnsNotFound = vbNullString Then
                                mstrColumnsNotFound = strElement
                            Else
                                mstrColumnsNotFound = mstrColumnsNotFound & ", " & strElement
                            End If
                        End If
                    End If
                Else
                    lngColumn = fGetColumnNumber(strElement, lngHeadersRow, strWorksheet)
                    If lngColumn <> 0 Then
                            lngRow = fGetLastNonEmptyRow(lngColumn, strWorksheet)
                            Set rngSearch = Range(.Sheets(strWorksheet).Cells(lngHeadersRow + 1, lngColumn), _
                                                  .Sheets(strWorksheet).Cells(lngRow, lngColumn))
                            colRanges.Add rngSearch, strElement
                        Else
                            If mstrColumnsNotFound = vbNullString Then
                                mstrColumnsNotFound = strElement
                            Else
                                mstrColumnsNotFound = mstrColumnsNotFound & ", " & strElement
                            End If
                        End If
                End If
            End If
            i = i + 1
        Loop Until strElement = vbNullString
    End With
            
    Set fCollectColumns = colRanges

Exit_ErrHandler:
    'Cleanup
    Set rngSearch = Nothing
    Set colRanges = Nothing
    Exit Function

ErrHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    
    Resume Exit_ErrHandler
    Resume

End Function

Public Function fGetColumnNumber(strColumnHeader As String, Optional lngHeadersRow As Long = 1, _
                                Optional strWorksheet As String = vbNullString) As Long
' --------------------------------------------------------------
' Comments:
'   Goes through a range of headers and finds the number of a
'   specific column
'
' Arguments:
'   strColumnHeader (String) = The name of the column to look for
'   lngHeadersRow   (Long)   = The row number where the columns names
'                              are. Defaults to 1
'   strWorksheet    (String) = The name of the worksheet to look for.
'                              If no worksheet is specified it defaults
'                              to looking in the current sheet.
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler

    Dim rngSearch As Excel.Range
    Dim rngFound As Excel.Range
    Dim shtWorksheet As Excel.Worksheet
    
    If strWorksheet = vbNullString Then
        Set shtWorksheet = ActiveWorkbook.ActiveSheet
    Else
        Set shtWorksheet = ActiveWorkbook.Sheets(strWorksheet)
    End If
    
    With shtWorksheet
        Set rngSearch = .Range(lngHeadersRow & ":" & lngHeadersRow)
        Set rngFound = rngSearch.Find(strColumnHeader)
        If rngFound Is Nothing Then
            fGetColumnNumber = 0
        Else
            fGetColumnNumber = rngFound.Column
        End If
    End With
    
Exit_ErrHandler:
    Set rngSearch = Nothing
    Set rngFound = Nothing
    Exit Function

ErrHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    
    Resume Exit_ErrHandler
    Resume
    
End Function

Public Function fGetLastNonEmptyRow(Optional varColumn As Variant = 1, _
                                    Optional strWorksheet As String = vbNullString) As Long
' --------------------------------------------------------------
' Comments:
'   Finds the last row to be non-empty in a worksheet by looking
'   on values on a specific column. Returns the number of the row.
'
' Arguments:
'   varColumn    (Variant) = Number of column to use for looking.
'   strWorksheet (String)  = Name of worksheet to look in. If no
'                            sheet is provided the current sheet
'                            is used.
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler

    Dim rngLastCell As Excel.Range
    Dim shtWorksheet As Excel.Worksheet
    
    If strWorksheet = vbNullString Then
        Set shtWorksheet = ActiveWorkbook.ActiveSheet
    Else
        Set shtWorksheet = ActiveWorkbook.Sheets(strWorksheet)
    End If
    
    With shtWorksheet
        Set rngLastCell = .Cells(.Rows.Count, varColumn).End(xlUp)
        fGetLastNonEmptyRow = rngLastCell.Row
    End With
    
Exit_ErrHandler:
    Set rngLastCell = Nothing
    Set shtWorksheet = Nothing
    Exit Function

ErrHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    
    Resume Exit_ErrHandler
    Resume
    
End Function

Public Function fGetLastNonEmptyColumn(Optional lngRow As Long = 1, _
                                        Optional shtWorksheet As Excel.Worksheet = Nothing) As Long
' --------------------------------------------------------------
' Comments:
'   Finds the last column to be non-empty in a worksheet by looking
'   on values on a specific row. Returns the number of the column.
'
' Arguments:
'   lngRow       (Long)    = Number of row to use for looking.
'   strWorksheet (String)  = Name of worksheet to look in. If no
'                            sheet is provided the current sheet
'                            is used.
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
'
On Error GoTo ErrHandler

    Dim rngLastCell As Excel.Range
    
    If shtWorksheet Is Nothing Then
        Set shtWorksheet = ActiveWorkbook.ActiveSheet
    End If
    
    With shtWorksheet
        If .Cells(lngRow, .Columns.Count) <> vbNullString Then
            Set rngLastCell = .Cells(lngRow, .Columns.Count)
            fGetLastNonEmptyColumn = rngLastCell.Column
        Else
            Set rngLastCell = .Cells(lngRow, .Columns.Count).End(xlToLeft)
            fGetLastNonEmptyColumn = rngLastCell.Column
        End If
    End With
    
Exit_ErrHandler:
    Set rngLastCell = Nothing
    Set shtWorksheet = Nothing
    Exit Function

ErrHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    
    Resume Exit_ErrHandler
    Resume
    
End Function

Public Function DoesPivotTableExist(ByRef shtWorksheet As Excel.Worksheet, _
                                    strPivotName As String) As Boolean
' --------------------------------------------------------------
' Comments:
'   Validates if a Pivot Table exists in the nominated spreadsheet
'   Returns false if not found.
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 19/02/2013    Carlos Gamez        Initial version
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler

    Dim pvtCurrent As Excel.PivotTable
    
    DoesPivotTableExist = False
    
    For Each pvtCurrent In shtWorksheet.PivotTables
        If pvtCurrent.Name = strPivotName Then
            DoesPivotTableExist = True
            Exit For
        End If
    Next pvtCurrent

Exit_ErrHandler:
    Set pvtCurrent = Nothing
    Exit Function

ErrHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    
    Resume Exit_ErrHandler
    Resume

End Function

