VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fSettings 
   Caption         =   "WIN Writer - Settings"
   ClientHeight    =   9795.001
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   11220
   OleObjectBlob   =   "fSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Description: This UserForm manages all the displaying, retrieving and
'              assignment of user selectable options.
'
' Dependencies:
'               mSettings => Module containig all helper functions & methods
'
' Authors:      Carlos Gamez
'

Option Explicit

Private Sub butBrowseEquipmentPhoto_Click()
    BrowseForImage Me.tbEquipmentPhoto
End Sub

Private Sub butBrowseFolderToSaveFiles_Click()
    BrowseForFolder Me.tbFolderToSaveFilesTo
End Sub

Private Sub butBrowseImageFolder_Click()
    BrowseForFolder Me.tbImageFolder
End Sub

Private Sub butBrowseManualsFolder_Click()
    BrowseForFolder Me.tbManualsFolder
End Sub

Private Sub butBrowseWordTemplate_Click()
    BrowseForWordFile Me.tbPMTemplateName, False
End Sub

Private Sub butClose_Click()
    Unload Me
End Sub

Private Sub butSave_Click()
    SaveSettings Me
End Sub

Private Sub cbPivotTableName_Change()
    sLoadControls Me, "PivotFields"
End Sub

Private Sub UserForm_Initialize()

InitGlobals

'Insert form caption
Me.Caption = gstrAPP_NAME & " Version: " & gstrAPP_VERSION & " - Settings"

'Control numbers of available tabs
'Basic->0
'Advanced RESD->1
'Advanced Word->2

'Move to default tab
Me.mpgSettings.value = 0

'Load current settings.
sLoadControls Me

End Sub
