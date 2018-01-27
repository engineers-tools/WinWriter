Attribute VB_Name = "mHelpers"
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
'              that might be called from any other module.
'
' Authors:      Carlos Gamez
'
Option Explicit

' **************************************************************
' Module Declarations Follow
' **************************************************************
' *** CONSTANTS ***
Private Const mstrMODULE As String = "mHelpers"

Public Function IsAppRunning(ByVal strAppName As String) As Boolean
' --------------------------------------------------------------
' Comments:
'   Check if an application is running and returns a boolean
'   representing the status of that application
'
' Arguments:
'   strAppName (String) = The name of the application to check
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
    Dim oApp As Object
    
    On Error Resume Next
    
    IsAppRunning = False
    
    Set oApp = GetObject(, strAppName)
    
    If Not oApp Is Nothing Then
        Set oApp = Nothing
        IsAppRunning = True
    End If
    
End Function

Public Sub sInsertFormula(rngToInsert As Excel.Range, strFormula As String)
' --------------------------------------------------------------
' Comments:
'   Inserts a formula on any Range. If the Range spans more than
'   one cell, it will insert the formula on every cell on that
'   range.
'
' Arguments:
'   rngToInsert (Range) = Range where the formula will be inserted
'   strFormula  (String)= Formula to be inserted
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
    Dim rngCurrent As Excel.Range
    
    For Each rngCurrent In rngToInsert
        rngCurrent.Formula = strFormula
    Next rngCurrent
    
End Sub

Public Function fLastSheetName() As String
' --------------------------------------------------------------
' Comments:
'   Returns the name of the last sheet in the workbook
'
' Arguments:
'   None
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
    Application.Volatile True
    
    With Application.Caller.Parent.Parent.Worksheets
        fLastSheetName = .Item(.Count).Name
    End With
    
End Function
