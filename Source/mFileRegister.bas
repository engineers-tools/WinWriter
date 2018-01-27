Attribute VB_Name = "mFileRegister"
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
'
' Description: This module contains code related to the generation
'              and management of a file registry, which contains a
'              list of all the files generated by this spreadsheet.
'
' Authors:      Carlos Gamez
'
Option Explicit

' **************************************************************
' Module Constant Declarations Follow
' **************************************************************
Private Const mstrMODULE As String = "mFileRegistry"
Private Const mszREGISTRY_SHEET As String = "Generated Files Log"
Private Const mszREGISTRY_TABLE As String = "tblGeneratedFilesLog"

Public Sub sFindLastRow(strGeneratedFilesLogSheet As String, _
                        strGeneratedFilesLogTable As String)

' --------------------------------------------------------------
' Comments: This routine finds the next blank column on a range
'           being used as a file register
'
' Arguments:
'   strGeneratedFilesLogSheet(String)   =   Name of the table holding
'                                       the record of the generated
'                                       files.
'   strGeneratedFilesLogTable(String)   = Name of the table
'
' Date          Developer       Comment
' --------------------------------------------------------------
' 04/07/2012    Carlos Gamez    Initial version
' 20/01/2018    Carlos Gamez    Open source release, added error handling
'
On Error GoTo ErrHandler
    
    ' --------------------------------------------------------------
    ' Variables Declarations
    ' --------------------------------------------------------------
        Dim shtRegister As Excel.Worksheet
        Dim tblRegister As Excel.Range
        Dim rngCell As Excel.Range
    
    ' --------------------------------------------------------------
    ' Set Objects
    ' --------------------------------------------------------------
        Set shtRegister = ThisWorkbook.Worksheets(strGeneratedFilesLogSheet)
        Set tblRegister = shtRegister.Range(strGeneratedFilesLogTable)
    
    ' --------------------------------------------------------------
    ' Place cursor in next available blank row
    ' --------------------------------------------------------------
        With shtRegister
            .Activate
            If tblRegister.Cells(1, 1) = "" Then
                Set rngCell = tblRegister.Cells(1, 1)
                rngCell.Select
            ElseIf tblRegister.Cells(1, 1) <> "" Then
                If tblRegister.Rows.Count > 1 Then
                    Set rngCell = tblRegister.Cells(1, 1)
                    Do
                        Set rngCell = rngCell.Offset(1, 0)
                        rngCell.Select
                    Loop Until rngCell.value = ""
                Else
                    Set rngCell = tblRegister.Cells(1, 1).Offset(1, 0)
                    rngCell.Select
                End If
            End If
        End With

Exit_ErrHandler:
    Set shtRegister = Nothing
    Set tblRegister = Nothing
    Exit Sub

ErrHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    
    Resume Exit_ErrHandler
    Resume

End Sub

Public Sub sRegisterFile(strFileName As String, strGeneratedOn As String, _
                         strGeneratedBy As String, strSavedTo As String, strTemplate As String)
' --------------------------------------------------------------
' Comments: This routine stores a new row with information about
'           the new file.
'
' Arguments:
'   strFileName(String)     =   Name of the file
'   strGeneratedOn(String)  =   Date on which the file was generated
'   strGeneratedBy(String)  =   Author that generated the file
'   strSavedTo(String)      =   Path to the folder where the file was saved to
'   strTemplate(String)     =   Path and name of template file used
'
' Date          Developer       Comment
' --------------------------------------------------------------
' 04/07/2012    Carlos Gamez    Initial version
' 26/01/2018    Carlos Gamez    Open source release, added error handling
'
On Error GoTo ErrHandler

    ' --------------------------------------------------------------
    ' Variables Declarations
    ' --------------------------------------------------------------
        Dim shtCurrentSheet As Excel.Worksheet
        Dim shtRegister As Excel.Worksheet
        Dim tblRegister As Excel.ListObject
    
        Set shtCurrentSheet = ActiveSheet
    
    ' --------------------------------------------------------------
    ' Find next available row
    ' --------------------------------------------------------------
        sFindLastRow mszREGISTRY_SHEET, mszREGISTRY_TABLE
    
    ' --------------------------------------------------------------
    ' Insert new row to avoid auto-total showing up
    ' --------------------------------------------------------------
        Set shtRegister = ThisWorkbook.Worksheets(mszREGISTRY_SHEET)
        Set tblRegister = shtRegister.ListObjects(mszREGISTRY_TABLE)
        
        tblRegister.ListRows.Add AlwaysInsert:=True
        
    ' --------------------------------------------------------------
    ' Log new file
    ' --------------------------------------------------------------
        'Disable automatic formula propagation
        Application.AutoCorrect.AutoFillFormulasInLists = False
        
        'Save values
        Selection.Formula = "=HYPERLINK(" & Chr(34) & strSavedTo & strFileName & ".docx" & _
                            Chr(34) & "," & Chr(34) & strFileName & Chr(34) & ")"
        ActiveCell.Offset(0, 1).Select
        Selection.value = strGeneratedOn
        ActiveCell.Offset(0, 1).Select
        Selection.value = strGeneratedBy
        ActiveCell.Offset(0, 1).Select
        Selection.Formula = "=HYPERLINK(" & Chr(34) & strSavedTo & Chr(34) & _
                                "," & Chr(34) & strSavedTo & Chr(34) & ")"
        ActiveCell.Offset(0, 1).Select
        Selection.Formula = "=HYPERLINK(" & Chr(34) & strTemplate & Chr(34) & _
                                "," & Chr(34) & strTemplate & Chr(34) & ")"
        
        'Re-enable automatic formula propagation (default Excel behaviour)
        Application.AutoCorrect.AutoFillFormulasInLists = True
        
    ' --------------------------------------------------------------
    ' Return to original sheet
    ' --------------------------------------------------------------
        shtCurrentSheet.Activate
    
Exit_ErrHandler:
    Set shtRegister = Nothing
    Set tblRegister = Nothing
    Exit Sub

ErrHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    
    Resume Exit_ErrHandler
    Resume
    
End Sub