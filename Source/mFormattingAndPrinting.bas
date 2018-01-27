Attribute VB_Name = "mFormattingAndPrinting"
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
' Description: This module encapsulates functions that format
'              and print various reports.
'
' Authors:      Carlos Gamez
'
Option Explicit

' **************************************************************
' Module Declarations Follow
' **************************************************************
' *** CONSTANTS ***
Private Const mstrMODULE As String = "mFormattingAndPrinting"

Sub sSetPrintArea(shtSheet As Worksheet, strPivotTable As String)
Attribute sSetPrintArea.VB_ProcData.VB_Invoke_Func = " \n14"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This routine takes a pivot table reference and
'           sets the printing area to that pivot table.
'
' Arguments:    strPivotTable as String
'
' Date          Developer       Comment
' --------------------------------------------------------------
' 18/11/2011    Carlos Gamez    Initial version
' 26/01/2018    Carlos Gamez    Open source release, added error handling

On Error GoTo ErrHandler
 
    Dim rngPivotTable As Range
    
    Application.ScreenUpdating = False
    
    With shtSheet
        Set rngPivotTable = .Range(.PivotTables(strPivotTable))
        .PageSetup.PrintArea = rngPivotTable
    End With
    
    Application.ScreenUpdating = True

Exit_ErrHandler:
    Exit Sub

ErrHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    
    Resume Exit_ErrHandler
    Resume
    
End Sub
