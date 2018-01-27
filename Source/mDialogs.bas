Attribute VB_Name = "mDialogs"
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
' Description: This module contains code to manage
'              Windows dialogs
'
' Authors:      Carlos Gamez
'
Option Explicit

' **************************************************************
' Module Constant Declarations Follow
' **************************************************************
Private Const mstrMODULE As String = "mDialogs"

Public Sub BrowseForFolder(ByRef tbCallingControl As Control)
' --------------------------------------------------------------
' Comments:
'   Opens the File dialog box and returns the selected
'   folder to the calling control.
'
' Arguments:
'   tbCallingControl (Control) = Control calling for the filename
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "Select Folder ..."
        .AllowMultiSelect = False
        .Show
    
        If .SelectedItems.Count > 0 Then tbCallingControl.Text = .SelectedItems(1) & "\"
    
    End With

Exit_ErrHandler:
    Exit Sub

ErrHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    
    Resume Exit_ErrHandler
    Resume

End Sub

Public Sub BrowseForImage(ByRef tbCallingControl As Control, Optional bolRemoveFolderPath As Boolean = True)
' --------------------------------------------------------------
' Comments:
'   Opens the Open Filename dialog box and returns the selected
'   file to the calling control. It uses the ImageFolder setting
'   as the default starting location.
'
' Arguments:
'   tbCallingControl (Control) = Control calling for the filename
'   bolRemoveFilePath(Boolean) = Flag indicating whether the folder
'                                should be removed from the path.
'                                Defaults to True
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .title = "Select Image ..."
        .Filters.Add "JPEG(*.jpg;*.jpeg)", "*.jpg;*.jpeg"
        .Filters.Add "Bitmap(*.bmp)", "*.bmp"
        .Filters.Add "Portable Graphics(*.png)", "*.png"
        .InitialFileName = fGetSetting("ImageFolder")
        .AllowMultiSelect = False
        .Show
    
        If .SelectedItems.Count > 0 Then
            If bolRemoveFolderPath = True Then
                tbCallingControl.Text = ExtractFilenameFromFullPath(.SelectedItems(1))
            Else
                tbCallingControl.Text = .SelectedItems(1)
            End If
        End If
    
    End With

Exit_ErrHandler:
    Exit Sub

ErrHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    
    Resume Exit_ErrHandler
    Resume

End Sub

Public Sub BrowseForWordFile(tbCallingControl As Control, Optional bolRemoveFolderPath As Boolean = True)
' --------------------------------------------------------------
' Comments:
'   Opens the Open Filename dialog box and returns the selected
'   file to the calling control. It uses the ImageFolder setting
'   as the default starting location.
'
' Arguments:
'   tbCallingControl (Control) = Control calling for the filename
'   bolRemoveFilePath(Boolean) = Flag indicating whether the folder
'                                should be removed from the path.
'                                Defaults to True
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .title = "Select Word File ..."
        .Filters.Add "Word 2007/2010(*.dotx)", "*.dotx"
        .Filters.Add "Word 2007/2010 (Macro Enabled)(*.dotm)", "*.dotm"
        .Filters.Add "Windows XP/2003(*.dot)", "*.dot"
        .AllowMultiSelect = False
        .Show
    
        If .SelectedItems.Count > 0 Then
            If bolRemoveFolderPath = True Then
                tbCallingControl.Text = ExtractFilenameFromFullPath(.SelectedItems(1))
            Else
                tbCallingControl.Text = .SelectedItems(1)
            End If
        End If
    
    End With

Exit_ErrHandler:
    Exit Sub

ErrHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    
    Resume Exit_ErrHandler
    Resume

End Sub
