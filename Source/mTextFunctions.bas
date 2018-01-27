Attribute VB_Name = "mTextFunctions"
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
' Description:  Functions used to manipulate text strings and lists
'
' Authors:      Carlos Gamez
'
' Options:
Option Explicit

' **************************************************************
' Module Constant Declarations Follow
' **************************************************************
Private Const mstrMODULE As String = "USysModTextFunctions"

Public Function fExtractFromList(strList As String, strSeparator As String, intElement As Integer) As String
' --------------------------------------------------------------
' Comments:
'   Returns an element indicated by intElement from the strList
'   of strings separated by the strSeparator character
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 12/07/12      Carlos Gamez      Initial version
' 26/01/2018    Carlos Gamez      Open source release
'
On Error GoTo ErrHandler
    
    Dim sarrList() As String
    
    sarrList = Split(strList, strSeparator)
    
    If intElement <= UBound(sarrList) + 1 Then
        fExtractFromList = Trim(sarrList(intElement - 1))
        If Left(fExtractFromList, 1) = "[" Then
            fExtractFromList = Right(fExtractFromList, Len(fExtractFromList) - 1)
        ElseIf Right(fExtractFromList, 1) = "]" Then
            fExtractFromList = Left(fExtractFromList, Len(fExtractFromList) - 1)
        End If
    Else
        fExtractFromList = vbNullString
    End If

    'If all the string is returned the strSeparator character was
    'not found, return null
    If fExtractFromList <> vbNullString And Len(fExtractFromList) = Len(strList) Then
        fExtractFromList = vbNullString
    End If

Exit_ErrHandler:
    Exit Function

ErrHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    
    Resume Exit_ErrHandler
    Resume

End Function

Public Function StripItem(startStrg As String, parser As String) As String
' --------------------------------------------------------------
' Comments:
'   This takes a string separated by the chr passed in parser,
'   splits off 1 item, and shortens startStrg so that the next
'   item is ready for removal.
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez      Open source release
'
   Dim cnt As Integer
   Dim Item As String
   
   cnt = 1
   
   Do
      If Mid(startStrg, cnt, 1) = parser Then
         Item = Mid(startStrg, 1, cnt - 1)
         startStrg = Mid(startStrg, cnt + 1, Len(startStrg))
         StripItem = Item
         Exit Function
      End If
      cnt = cnt + 1
   Loop

End Function
