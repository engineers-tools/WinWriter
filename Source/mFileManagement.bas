Attribute VB_Name = "mFilemanagement"
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
' Description: Contains methods that manipulate files and interact with the
'              file system.
'
' Authors:     Carlos Gamez
'
Option Explicit

' **************************************************************
' Module Constant Declarations Follow
' **************************************************************
Private Const mstrMODULE As String = "mFileManagement"

Public Function ExtractFilenameFromFullPath(strPath As String) As String
' --------------------------------------------------------------
' Comments:
'   This function takes a full path and returns the name of the file
'
' Arguments:
'   strPath (String) = The full path that contains the file name
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
    If Not strPath = vbNullString Then
        ExtractFilenameFromFullPath = StrReverse(Left(StrReverse(Trim(strPath)), _
                                    InStr(StrReverse(Trim(strPath)), "\") - 1))
    Else
        ExtractFilenameFromFullPath = strPath
    End If

End Function

Public Function CheckIfFileExists(strFilePathAndName As String) As Boolean
' --------------------------------------------------------------
' Comments:
'   This function takes the Path and Name of a file and
'   returns true if it exists.
'
' Arguments:
'   strFilePathAndName (String) = String containing the Path
'                                 and Name of file that needs
'                                 to be checked.The full path
'                                 that contains the file name
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 08/11/2011    Carlos Gamez        Initial version
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler

    Dim fs As Object
    Dim f As Object
    
    CheckIfFileExists = True
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(strFilePathAndName)

Exit_ErrHandler:
    Set fs = Nothing
    Set f = Nothing
    Exit Function

ErrHandler:
    Select Case Err.Number
        Case 53
            CheckIfFileExists = False
        Case Else
            sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    End Select
    
    Resume Exit_ErrHandler
    Resume
    
End Function

Public Function CheckIfFolderExists(strFolderPath As String) As Boolean
' --------------------------------------------------------------
' Comments: This function takes the Path of a folder and
'           returns true if it exists.
'
' Arguments:    strFolderPath = String containing the Path
'                                    and Name of file that needs
'                                    to be checked.
'
' Date          Developer       Comment
' --------------------------------------------------------------
' 08/11/2011    Carlos Gamez    Initial version
' 26/01/2018    Carlos Gamez    Open source release

On Error GoTo ErrHandler

    Dim fs As Object
    Dim f As Object
    
    CheckIfFolderExists = True
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(strFolderPath)

Exit_ErrHandler:
    Set fs = Nothing
    Set f = Nothing
    Exit Function

ErrHandler:
    Select Case Err.Number
        Case 76
            CheckIfFolderExists = False
        Case Else
            sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    End Select
    
    Resume Exit_ErrHandler
    Resume

End Function

Public Function fCheckTrailingBackslash(strFolderPath As String) As String
' --------------------------------------------------------------
' Comments: Validates if the folder path has a trailing backslash "\"
'           it adds the backslash if it's missing or return the same
'           path if it exists
'
' Arguments:    strFolderPath = String containing the Path
'                                    and Name of file that needs
'                                    to be checked.
'
' Date          Developer       Comment
' --------------------------------------------------------------
' 17/11/2011    Carlos Gamez    Initial version
' 26/01/2018    Carlos Gamez    Open source release
'
On Error GoTo ErrHandler

    Dim strLastCharacter As String
    
    strLastCharacter = Right(strFolderPath, 1)
    
    If strLastCharacter = "\" Then
        fCheckTrailingBackslash = strFolderPath
    Else
        fCheckTrailingBackslash = strFolderPath & "\"
    End If

Exit_ErrHandler:
    Exit Function

ErrHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    Resume Exit_ErrHandler
    Resume

End Function
