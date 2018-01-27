Attribute VB_Name = "mRibbonHandler"
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
' Description: This module manages the custom Ribbon
'
' Authors:      Carlos Gamez
'
Option Explicit

' **************************************************************
' Module Declarations Follow
' **************************************************************
' *** CONSTANTS ***
Private Const mstrMODULE As String = "mRibbonHandler"
Private mobjCustomRibbon As IRibbonUI

' *** PROPERTIES ***
Public Property Get WinWriterTab() As IRibbonUI
' --------------------------------------------------------------
' Comments:
'   Global property that re-exposes the ribbon system-wide
'   so consumers can call the IRibbonUI.Invalidate method as required
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
  Set WinWriterTab = mobjCustomRibbon
End Property

' *** CUSTOM RIBBON XML ***
'<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnMainLoad">
'    <ribbon>
'        <tabs>
'            <tab id="tabWinWriter" label="WORK INSTRUCTION WRITER" insertAfterMso="TabHome">
'                <group id="grpOptions" label="Options">
'                    <button id="butSettings" label="Settings" size="large" onAction="OpenSettings" imageMso="AccessNavigationOptions" />
'                </group>
'                <group id="grpFileManagement" label="File Management">
'                    <button id="butGenerateWordFiles" label="Generate Documents" size="large" onAction="GenerateWordDocument" imageMso="ExportWord"/>
'                </group>
'
'                <group id="grpTools" label="Tools">
'                    <button id="butInsertTaskPhoto" label="Insert Task Photo" size="large" onAction="InsertTaskPhoto" imageMso="PictureReflectionGallery" />
'                    <button id="butInsertManualReference" label="Insert Manual Reference" size="large" onAction="InsertManualReference" imageMso="FunctionsLookupReferenceInsertGallery" />
'                </group>
'            </tab>
'        </tabs>
'    </ribbon>
'</customUI>

' *** METHODS ***
Public Sub OnMainLoad(Ribbon As IRibbonUI)
' --------------------------------------------------------------
' Comments:
'   Handles the OnLoad callback of the customUI element
'   and saves the IRibbonUI object reference to mobjCustomRibbon
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
  Set mobjCustomRibbon = Ribbon
End Sub

Public Sub OpenSettings(Optional butRibbon As IRibbonControl)
' --------------------------------------------------------------
' Comments:
'   Opens the main Settings form
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
    fSettings.Show
End Sub

Public Sub InsertTaskPhoto(Optional butRibbon As IRibbonControl)
' --------------------------------------------------------------
' Comments:
'   Inserts a task Hyperlink to an Image in the currently selected
'   Cell.
'   TODO: Remove hard-coded references
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
    Dim NewFile As String
    
    'Variables initialization
    Call InitializeVariables("InsertTaskPhoto")
    
    With shtWorkInstructions
        If .ListObjects("tblWorkInstructions").ListColumns(TaskPhotoFieldName).Range.Column = ActiveCell.Column Or _
           .ListObjects("tblWorkInstructions").ListColumns(InTaskPhotoFieldName).Range.Column = ActiveCell.Column Then
        
            NewFile = Application.GetOpenFilename( _
                      "All Files (*.*),*.*," & _
                      "JPEG(*.jpg *.jpeg),*.jpg;*.jpeg," & _
                      "Portable Graphics (*png),*.png," & _
                      "Bitmap(*.bmp),*.bmp", , , , False)
            
            NewFile = ExtractFilenameFromFullPath(NewFile)
            
            If Not NewFile = vbNullString Then
                'Disable automatic formula propagation
                Application.AutoCorrect.AutoFillFormulasInLists = False
                ActiveCell.Formula = "=HYPERLINK(ImageFolder & " _
                                     & Chr(34) & NewFile & Chr(34) & _
                                     "," & Chr(34) & NewFile & Chr(34) & ")"
                'Re-enable automatic formula propagation (default Excel behaviour)
                Application.AutoCorrect.AutoFillFormulasInLists = True
            End If
            
        Else
            MsgBox "Please select a cell within the " & TaskPhotoFieldName & " or " & _
                   InTaskPhotoFieldName & " column.", , gstrAPP_NAME
        End If
        
    End With

End Sub

Public Sub InsertManualReference(Optional butRibbon As IRibbonControl)
' --------------------------------------------------------------
' Comments:
'   Inserts a task Hyperlink to a File in the currently selected
'   Cell.
'   TODO: Remove hard-coded references
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
    Dim NewFile As String
    
    'Variables initialization
    Call InitializeVariables("InsertManualReference")
    
    With shtWorkInstructions
        If .ListObjects("tblWorkInstructions").ListColumns(TaskSpecialReferencesFieldName).Range.Column = ActiveCell.Column Then
        
            NewFile = Application.GetOpenFilename( _
                      "All Files (*.*),*.*," & _
                      "PDF (*.pdf),*.pdf," & _
                      "Word 2003 (*.doc),*.doc," & _
                      "Word 2007-2010 (*docx),*.docx", , , , False)
                      
            NewFile = ExtractFilenameFromFullPath(NewFile)
            
            If Not NewFile = vbNullString Then
                'Disable automatic formula propagation
                Application.AutoCorrect.AutoFillFormulasInLists = False
                ActiveCell.Formula = "=HYPERLINK(ManualsFolder & " & Chr(34) & NewFile & Chr(34) & _
                                     "," & Chr(34) & NewFile & Chr(34) & ")"
                'Re-enable automatic formula propagation (default Excel behaviour)
                Application.AutoCorrect.AutoFillFormulasInLists = True
            End If
            
        Else
            MsgBox "Please select a cell within the " & TaskSpecialReferencesFieldName & " column.", , gstrAPP_NAME
        End If
        
    End With

End Sub

