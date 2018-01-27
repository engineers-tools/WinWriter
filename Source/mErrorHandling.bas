Attribute VB_Name = "mErrorHandling"
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
' Description:  Used to handle unexpected errors
'
' Authors:      Carlos Gamez
'
' Options:
    Option Explicit
'
' **************************************************************
' Module Constant Declarations Follow
' **************************************************************
    Private Const mstrMODULE As String = "mErrorHandling"
'
' **************************************************************
' List of Custom/Application Specific Errors
' **************************************************************
'
    'Err.Raise 513, mstrMODULE, _
                    "The Pivot Table [" & <NAME OF PIVOT TABLE> & "] could not be found." & _
                    vbCrLf & "Please check your settings."
    
    'Err.Raise 514, mstrMODULE, _
                    "The bookmark [" & <NAME OF BOOKMARK VARIABLE> & "] could not be found in the [" & PMTemplateName & "] gallery." & Chr(13) & _
                    "This part of the document will be skipped. Please check the Word Template file and ensure it contains " & Chr(13) & _
                    "the required elements."
    
    'Err.Raise 515, mstrMODULE, _
                    "Error During Tasks Insertion:" & vbCrLf & _
                    "The Tag [" & <NAME OF TAG VARIABLE> & "] could not be found on the template document." & vbCrLf & _
                    "Verify that it exists or the name has been entered correctly in the Settings." & vbCrLf & _
                    "The document cannot be generated correctly."
'
' **************************************************************

Public Sub sDisplayUnexpectedError(strErrorNumber As String, strErrorDescription As String, _
                                    Optional strModuleName As String = vbNullString)

' --------------------------------------------------------------
' Comments:
'   This procedure takes an error number and description and presents
'   a standardized error message
'   Use this line to call for this centralised error handling routine
'
'   sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
'   sDisplayUnexpectedError Err.Number, Err.Description, cstrCLASS
'
' Arguments:
'   The variable names are self-explanatory
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 11/07/2012    Carlos Gamez        Initial version
' 26/01/2018    Carlos Gamez        Open source release

    Dim strMessage As String
    Dim strMessageWithDeveloper As String
    Dim strAuthor As String
    Dim strEmail As String
    Dim strAppName As String
    Dim strAppVersion As String
    
    InitGlobals
    strAuthor = gstrAUTHOR
    strEmail = gstrEMAIL
    strAppName = gstrAPP_NAME
    strAppVersion = gstrAPP_VERSION
            
    strMessage = "An error has occurred in this application." & vbCrLf & vbCrLf & _
            "[Error Number]: " & strErrorNumber & vbCrLf & vbCrLf & _
            "[Error Description]: " & vbCrLf & strErrorDescription & vbCrLf

    If strModuleName <> vbNullString Then
        strMessage = strMessage & vbCrLf & "[Module]: " & strModuleName
    End If

    If MsgBox(strMessage, Buttons:=vbRetryCancel, title:=strAppName & " - " & strAppVersion) = vbCancel Then
        End
    End If

End Sub

