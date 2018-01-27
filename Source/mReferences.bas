Attribute VB_Name = "mReferences"
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
' Description: This module deals with functions relating to Reference management
'              in this VBProject
'
' Authors:      Carlos Gamez
'
Option Explicit

' **************************************************************
' Module Declarations Follow
' **************************************************************
' *** CONSTANTS ***
Private Const mstrMODULE As String = "mReferences"
Private mstrColumnsNotFound As String

' **************************************************************
' Global Constant Declarations Follow
' **************************************************************
Public Const mstrREFERENCES As String = "MSWORD.OLB"

' *** FUNCTION DECLARATIONS ***
#If Win64 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
#Else
    Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
#End If

' *** TYPES ***
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public Function CreateGUID() As String
' --------------------------------------------------------------
' Comments:
'   Generates a GUID string. I'm documenting this now as part of
'   the open source release of this project and I cannot honestly
'   remember where I obtained it a few years ago. If anyone reading
'   this comment identifies the author, please let me know to give
'   proper credit.
'
' Arguments:
'   None
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
    Dim G As GUID
    
    With G
        If (CoCreateGuid(G) = 0) Then
        CreateGUID = _
            String$(8 - Len(Hex$(.Data1)), "0") & Hex$(.Data1) & _
            String$(4 - Len(Hex$(.Data2)), "0") & Hex$(.Data2) & _
            String$(4 - Len(Hex$(.Data3)), "0") & Hex$(.Data3) & _
            IIf((.Data4(0) < &H10), "0", "") & Hex$(.Data4(0)) & _
            IIf((.Data4(1) < &H10), "0", "") & Hex$(.Data4(1)) & _
            IIf((.Data4(2) < &H10), "0", "") & Hex$(.Data4(2)) & _
            IIf((.Data4(3) < &H10), "0", "") & Hex$(.Data4(3)) & _
            IIf((.Data4(4) < &H10), "0", "") & Hex$(.Data4(4)) & _
            IIf((.Data4(5) < &H10), "0", "") & Hex$(.Data4(5)) & _
            IIf((.Data4(6) < &H10), "0", "") & Hex$(.Data4(6)) & _
            IIf((.Data4(7) < &H10), "0", "") & Hex$(.Data4(7))
        End If
    End With
End Function

Public Sub sDebugReferences()
' --------------------------------------------------------------
' Comments:
'   Prints all references in the current VBProject to the
'   Immediate window
'
' Arguments:
'   None
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
    Dim n As Integer
     
    On Error Resume Next
    
    For n = 1 To ThisWorkbook.VBProject.References.Count
        Debug.Print "Name: " & ThisWorkbook.VBProject.References.Item(n).Name
        Debug.Print "Description: " & ThisWorkbook.VBProject.References.Item(n).Description
        Debug.Print "GUID: " & ThisWorkbook.VBProject.References.Item(n).GUID
        Debug.Print "Major: " & ThisWorkbook.VBProject.References.Item(n).Major
        Debug.Print "Minor: " & ThisWorkbook.VBProject.References.Item(n).Minor
        Debug.Print "Fullpath: " & ThisWorkbook.VBProject.References.Item(n).fullpath
        Debug.Print "--------------------------------------------------------------------------------"
    Next n
     
End Sub

Public Sub sRemoveBrokenReferences()
' --------------------------------------------------------------
' Comments:
'   Remove any missing references from the current VBProject
'
' Arguments:
'   None
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
    Dim i As Integer
    Dim theRef As Variant
    
    For i = ThisWorkbook.VBProject.References.Count To 1 Step -1
        Set theRef = ThisWorkbook.VBProject.References.Item(i)
        If theRef.isbroken = True Then
            ThisWorkbook.VBProject.References.Remove theRef
        End If
    Next i

End Sub

Public Sub RemoveReference(strGUID As String)
' --------------------------------------------------------------
' Comments:
'   Remove a specific reference from the current VBProject.
'
' Arguments:
'   strGUID (String) = The GUID of the reference being removed.
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler

    Dim theRef As Variant, i As Long
    Dim errDescription As String
     
    'Clear any errors so that error trapping for GUID additions can be evaluated
    Err.Clear
     
    'Remove reference
    For i = ThisWorkbook.VBProject.References.Count To 1 Step -1
        Set theRef = ThisWorkbook.VBProject.References.Item(i)
        If theRef.GUID = strGUID Then
            ThisWorkbook.VBProject.References.Remove theRef
        End If
    Next i

Exit_ErrHandler:
    Exit Sub

ErrHandler:
    Select Case Err.Number
        Case Is = 32813
             'Reference already in use.  No action necessary
             Resume Next
        Case Is = 1004
            'VBA is not allowed to run, modify the trust center
            errDescription = "The security settings do not allow to load the needed references." & vbCrLf & _
                    "You can still work with this file, but the PRT generation and word" & vbCrLf & _
                    "exporting functions will not be available." & vbCrLf & _
                    vbCrLf & _
                    "Please tick the 'Trust access to the VBA object model' checkbox" & vbCrLf & _
                    "under 'File > Options > Trust Center > Trust Center Settings > Macro Settings' and reload this file."
                    
            sDisplayUnexpectedError Err.Number, errDescription, mstrMODULE
            Err.Clear
            Exit Sub
        Case Is = vbNullString
             'Reference added without issue
        Case Else
            errDescription = "A problem was encountered trying to" & vbNewLine _
            & "add or remove a reference in this file" & vbNewLine & "Please check the " _
            & "references in your VBA project!"
            sDisplayUnexpectedError Err.Number, errDescription, mstrMODULE
    End Select
    
    Resume Exit_ErrHandler
    Resume
    
End Sub

Public Sub AddReference(strGUID As String)
' --------------------------------------------------------------
' Comments:
'   Add a specific reference to the current VBProject.
'
' Arguments:
'   strGUID (String) = The GUID of the reference being added.
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler

    Dim theRef As Variant, i As Long
    Dim errDescription As String
    
    'Add the reference, depending on the type of reference passed
    If Left(strGUID, 1) = "{" Then 'It's a GUID
        ThisWorkbook.VBProject.References.AddFromGuid _
                    GUID:=strGUID, Major:=1, Minor:=0
    Else 'It's a file
        ThisWorkbook.VBProject.References.AddFromFile strGUID
    End If
     
Exit_ErrHandler:
    Exit Sub

ErrHandler:
    Select Case Err.Number
        Case Is = 32813
            'Reference already in use.  No action necessary
            Resume Next
        Case Is = 1004
            'VBA is not allowed to run, modify the trust center
            errDescription = "The security settings do not allow to load the needed references." & vbCrLf & _
                    "You can still work with this file, but the PRT generation and word" & vbCrLf & _
                    "exporting functions will not be available." & vbCrLf & _
                    vbCrLf & _
                    "Please tick the 'Trust access to the VBA object model' checkbox" & vbCrLf & _
                    "under 'File > Options > Trust Center > Trust Center Settings > Macro Settings' and reload this file."
                    
            sDisplayUnexpectedError Err.Number, errDescription, mstrMODULE
            Err.Clear
            Exit Sub
        Case Is = vbNullString
             'Reference added without issue
        Case Else
            errDescription = "A problem was encountered trying to" & vbNewLine _
            & "add or remove a reference in this file" & vbNewLine & "Please check the " _
            & "references in your VBA project!"
            sDisplayUnexpectedError Err.Number, errDescription, mstrMODULE
    End Select
    
    Resume Exit_ErrHandler
    Resume

End Sub

Public Sub LoadReferences()
' --------------------------------------------------------------
' Comments:
'   Adds all references necessary for this application to run.
'   References are defined on the module level constant:
'           mstrREFERENCES
'
' Arguments:
'   None
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler

    Dim sExcelVersion As Single
  
    sRemoveBrokenReferences
    
    'Load references
    AddReference fCheckTrailingBackslash(Application.Path) & mstrREFERENCES
  
    dbReferencesLoaded = True

    If gbolDEBUG = True Then MsgBox "References loaded successfuly", vbInformation, gstrAPP_NAME
    
Exit_ErrHandler:
    Exit Sub

ErrHandler:
    Select Case Err.Number
        Case 32813
          Resume Next
        Case Else
            sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    End Select
    
    Resume Exit_ErrHandler
    Resume
    
End Sub

Public Sub UnloadReferences()
' --------------------------------------------------------------
' Comments:
'   Unload references that were added in this project
'
' Arguments:
'   None
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler
  
    Dim sExcelVersion As Single
  
    'Check Excel Version
    sExcelVersion = Excel.Application.Version
    
    'Select the appropiate references to load according to the version
    Select Case sExcelVersion
        'Case 14 'Excel 2010
        '    RemoveReference "{00020905-0000-0000-C000-000000000046}"
    End Select
    
Exit_ErrHandler:
    Exit Sub

ErrHandler:
    Select Case Err.Number
        Case 32813
          Resume Next
        Case Else
            sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    End Select
    
    Resume Exit_ErrHandler
    Resume
    
End Sub
