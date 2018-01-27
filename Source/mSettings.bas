Attribute VB_Name = "mSettings"
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
' Description:  This module contains procedures to set and retrieve
'               application settings
'
' Authors:      Carlos Gamez
'
' Options:
    Option Explicit
    
' **************************************************************
' Module Constant Declarations Follow
' **************************************************************
    Private Const mstrMODULE As String = "mSettings"

' **************************************************************
'Global variables related to the Document Writing procedures
' **************************************************************

    ' ** Category: tblWorkInstructions **
    Public DocTitleFieldName As String
    Public EquipmentNumber As String
    Public EquipmentName As String
    Public EquipmentPhoto As String
    Public EquipmentStatusOnlineWord As String
    Public EquipmentStatusOfflineWord As String
    Public ComponentSummaryRangeName As String
    Public PivotTableName As String
    Public IntervalFieldName As String
    Public IntervalUnitsFieldName As String
    Public EquipmentStatusFieldName As String
    Public TradeFieldName As String
    Public ComponentNumberFieldName As String
    Public ComponentDescriptionFieldName As String
    Public RowHeadersFieldName As String
    Public ColumnHeadersFieldName As String
    Public TaskFieldName As String
    Public TaskPhotoFieldName As String
    Public InTaskPhotoFieldName As String
    Public TaskSafetyWarningLabelFieldName As String
    Public TaskToolsFieldName As String
    Public TaskMaterialsFieldName As String
    Public TaskSpecialNotesFieldName As String
    Public TaskSpecialPermitsFieldName As String
    Public TaskSpecialPPEFieldName As String
    Public TaskSpecialHazardsFieldName As String
    Public TaskSpecialReferencesFieldName As String
    Public AcceptableConditionsFieldName As String
    Public DocumentAuthor As String
    Public ManualsFolder As String
    
    ' ** Category: Images **
    Public ImageFolder As String
    Public PhotoHeight As String
    Public TaskPhotoHeight As String
    
    ' ** Category: Word Application Options **
    Public ScreenUpdatingOption As String
    
    ' ** Category: Word Document - File Options **
    Public PMTemplateName As String
    Public FolderToSaveFilesTo As String
    Public IncludeDescription As String
    Public DocumentNumberPrefix As String
    Public IncludeNumberPrefix As String
    Public DocumentNumberSuffix As String
    Public IncludeNumberSuffix As String
    Public DocumentConsecutiveStartingNumber As String
    
    ' ** Category: Word Document - Bookmarks **
    Public PMDocumentTitleBookmarkName As String
    Public PMSiteOrAreaBookmarkName As String
    Public PMIntervalBookmarkName As String
    Public PMIntervalUnitBookmarkName As String
    Public PMEquipmentStatusBookmarkName As String
    Public PMEquipmentNumberBookmarkName As String
    Public PMTradeBookmarkName As String
    Public PMFrontPageSectionBookmarkName As String
    Public PMFrontPagePhotoBookmarkName As String
    Public PMFrontPageAimBookmarkName As String
    Public PMFrontPageInstructionsBookmarkName As String
    Public PMFrontPageLocationBookmarkName As String
    Public PMFrontPagePermitsAndIsolationsBookmarkName As String
    Public PMFrontPagePPEBookmarkName As String
    Public PMFrontPageOtherSafetyEquipmentBookmarkName As String
    Public PMFrontPageHazardsAndSafetyNotesBookmarkName As String
    Public PMFrontPageToolsBookmarkName As String
    Public PMFrontPagePartsAndMaterialsBookmarkName As String
    Public PMFrontPageSpecialNotesBookmarkName As String
    Public PMFrontPageReferencesBookmarkName As String
    Public PMFrontPageResourcesBookmarkName As String
    Public PMSectionBookmarkName As String
    Public PMRESDAuthorBookmarkName As String
    
    ' ** Category: Word Document - Block Names **
    Public PMInstructionBlockName As String
    Public PMInstructionTasksOnlyBlockName As String
    Public PMFrontPageBlankBlockName As String
    Public PMBlankTableBlockName As String
    
    ' ** Category: Word Document - Markers **
    Public InsertInstructionMarker As String
    Public InsertEquipmentNameMarker As String
    Public InsertEquipmentPhotoMarker As String
    Public FunctionalLocationMarker As String

Public Function fGetSetting(strSettingName As String) As Variant
' ---------------------------------------------------------------------------
' Purpose:
'   Takes the name of a setting and returns it's value
'
' Date          Developer           Comment
' ---------------------------------------------------------------------------
' 07/08/12      Carlos Gamez        Initial version
' 26/01/2018    Carlos Gamez        Open source release

On Error GoTo ErrHandler
    
    fGetSetting = shtDocumentControl.Range(strSettingName).value

Exit_ErrHandler:
    Exit Function

ErrHandler:
    Select Case Err.Number
        Case Is = 1004
            fGetSetting = vbNullString
        Case Else
            sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    End Select

    Resume Exit_ErrHandler
    Resume

End Function

Public Function fAssignSetting(strSettingName As String, varValue As Variant) As Boolean
' ---------------------------------------------------------------------------
' Purpose:
'   Takes the name of a setting and its value and returns true if setting is
'   set correctly
'
' Date          Developer           Comment
' ---------------------------------------------------------------------------
' 19/02/13      Carlos Gamez        Initial version
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler
    
    shtDocumentControl.Range(strSettingName).value = varValue
    fAssignSetting = True

Exit_ErrHandler:
    Exit Function

ErrHandler:
    Select Case Err.Number
        Case Is = 1004
            fAssignSetting = False
        Case Else
            fAssignSetting = False
            sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    End Select

    Resume Exit_ErrHandler
    Resume

End Function

Public Function InitializeVariables(Optional strCallingModule As String) As Boolean
' ---------------------------------------------------------------------------
' Purpose:
'   Initialise all global variables that control the document generation process
'   using stored settings.
'
' Date          Developer           Comment
' ---------------------------------------------------------------------------
' 19/02/13      Carlos Gamez        Initial version
'
    Dim strErrors As String
    
    InitializeVariables = True
    
    ' ** Category: tblWorkInstructions **
    DocTitleFieldName = fGetSetting("DocumentTitle")
    EquipmentNumber = fGetSetting("EquipmentNumber")
    EquipmentName = fGetSetting("EquipmentName")
    EquipmentPhoto = fGetSetting("EquipmentPhoto")
    EquipmentStatusOnlineWord = fGetSetting("EquipmentStatusOnlineWord")
    EquipmentStatusOfflineWord = fGetSetting("EquipmentStatusOfflineWord")
    ComponentSummaryRangeName = fGetSetting("ComponentSummaryRangeName")
    PivotTableName = fGetSetting("PivotTableName")
    IntervalFieldName = fGetSetting("IntervalFieldName")
    IntervalUnitsFieldName = fGetSetting("IntervalUnitsFieldName")
    EquipmentStatusFieldName = fGetSetting("EquipmentStatusFieldName")
    TradeFieldName = fGetSetting("TradeFieldName")
    ComponentNumberFieldName = fGetSetting("ComponentNumberFieldName")
    ComponentDescriptionFieldName = fGetSetting("ComponentDescriptionFieldName")
    RowHeadersFieldName = fGetSetting("RowHeadersFieldName")
    ColumnHeadersFieldName = fGetSetting("ColumnHeadersFieldName")
    TaskFieldName = fGetSetting("TaskFieldName")
    TaskPhotoFieldName = fGetSetting("TaskPhotoFieldName")
    InTaskPhotoFieldName = fGetSetting("InTaskPhotoFieldName")
    TaskSafetyWarningLabelFieldName = fGetSetting("TaskSafetyWarningLabelFieldName")
    TaskToolsFieldName = fGetSetting("TaskToolsFieldName")
    TaskMaterialsFieldName = fGetSetting("TaskMaterialsFieldName")
    TaskSpecialNotesFieldName = fGetSetting("TaskSpecialNotesFieldName")
    TaskSpecialPermitsFieldName = fGetSetting("TaskSpecialPermitsFieldName")
    TaskSpecialPPEFieldName = fGetSetting("TaskSpecialPPEFieldName")
    TaskSpecialHazardsFieldName = fGetSetting("TaskSpecialHazardsFieldName")
    TaskSpecialReferencesFieldName = fGetSetting("TaskSpecialReferencesFieldName")
    AcceptableConditionsFieldName = fGetSetting("AcceptableConditionsFieldName")
    DocumentAuthor = fGetSetting("DocumentAuthor")
    ManualsFolder = fGetSetting("ManualsFolder")
    
    ' ** Category: Images **
    ImageFolder = fGetSetting("ImageFolder")
    PhotoHeight = fGetSetting("PhotoHeight")
    TaskPhotoHeight = fGetSetting("TaskPhotoHeight")
    
    ' ** Category: Word Application Options **
    ScreenUpdatingOption = fGetSetting("ScreenUpdatingOption")
    
    ' ** Category: Word Document - File Options **
    PMTemplateName = fGetSetting("PMTemplateName")
    FolderToSaveFilesTo = fGetSetting("FolderToSaveFilesTo")
    IncludeDescription = fGetSetting("IncludeDescription")
    DocumentNumberPrefix = fGetSetting("DocumentNumberPrefix")
    IncludeNumberPrefix = fGetSetting("IncludeNumberPrefix")
    DocumentNumberSuffix = fGetSetting("DocumentNumberSuffix")
    IncludeNumberSuffix = fGetSetting("IncludeNumberSuffix")
    DocumentConsecutiveStartingNumber = fGetSetting("DocumentConsecutiveStartingNumber")
    
    ' ** Category: Word Document - Bookmarks **
    PMDocumentTitleBookmarkName = fGetSetting("PMDocumentTitleBookmarkName")
    PMSiteOrAreaBookmarkName = fGetSetting("PMSiteOrAreaBookmarkName")
    PMIntervalBookmarkName = fGetSetting("PMIntervalBookmarkName")
    PMIntervalUnitBookmarkName = fGetSetting("PMIntervalUnitBookmarkName")
    PMEquipmentStatusBookmarkName = fGetSetting("PMEquipmentStatusBookmarkName")
    PMEquipmentNumberBookmarkName = fGetSetting("PMEquipmentNumberBookmarkName")
    PMTradeBookmarkName = fGetSetting("PMTradeBookmarkName")
    PMFrontPageSectionBookmarkName = fGetSetting("PMFrontPageSectionBookmarkName")
    PMFrontPagePhotoBookmarkName = fGetSetting("PMFrontPagePhotoBookmarkName")
    PMFrontPageAimBookmarkName = fGetSetting("PMFrontPageAimBookmarkName")
    PMFrontPageInstructionsBookmarkName = fGetSetting("PMFrontPageInstructionsBookmarkName")
    PMFrontPageLocationBookmarkName = fGetSetting("PMFrontPageLocationBookmarkName")
    PMFrontPagePermitsAndIsolationsBookmarkName = fGetSetting("PMFrontPagePermitsAndIsolationsBookmarkName")
    PMFrontPagePPEBookmarkName = fGetSetting("PMFrontPagePPEBookmarkName")
    PMFrontPageOtherSafetyEquipmentBookmarkName = fGetSetting("PMFrontPageOtherSafetyEquipmentBookmarkName")
    PMFrontPageHazardsAndSafetyNotesBookmarkName = fGetSetting("PMFrontPageHazardsAndSafetyNotesBookmarkName")
    PMFrontPageToolsBookmarkName = fGetSetting("PMFrontPageToolsBookmarkName")
    PMFrontPagePartsAndMaterialsBookmarkName = fGetSetting("PMFrontPagePartsAndMaterialsBookmarkName")
    PMFrontPageSpecialNotesBookmarkName = fGetSetting("PMFrontPageSpecialNotesBookmarkName")
    PMFrontPageReferencesBookmarkName = fGetSetting("PMFrontPageReferencesBookmarkName")
    PMFrontPageResourcesBookmarkName = fGetSetting("PMFrontPageResourcesBookmarkName")
    PMSectionBookmarkName = fGetSetting("PMSectionBookmarkName")
    PMRESDAuthorBookmarkName = fGetSetting("PMRESDAuthorBookmarkName")
    
    ' ** Category: Word Document - Block Names **
    PMInstructionBlockName = fGetSetting("PMInstructionBlockName")
    PMInstructionTasksOnlyBlockName = fGetSetting("PMInstructionTasksOnlyBlockName")
    PMFrontPageBlankBlockName = fGetSetting("PMFrontPageBlankBlockName")
    PMBlankTableBlockName = fGetSetting("PMBlankTableBlockName")
    
    ' ** Category: Word Document - Markers **
    InsertInstructionMarker = fGetSetting("InsertInstructionMarker")
    InsertEquipmentNameMarker = fGetSetting("InsertEquipmentNameMarker")
    InsertEquipmentPhotoMarker = fGetSetting("InsertEquipmentPhotoMarker")
    FunctionalLocationMarker = fGetSetting("FunctionalLocationMarker")
    
    If strCallingModule = "GenerateWordDocument" Then
        
        strErrors = fCheckFolders
        
        If strErrors <> vbNullString Then
            InitializeVariables = False
        End If
        
    End If
    
    'Report errors if any
    If InitializeVariables = False Then
        sDisplayUnexpectedError "File Locations", strErrors & vbCrLf & "Documents cannot be generated. Please check the settings before continuing.", mstrMODULE
    End If

End Function

Public Sub sLoadControls(ByRef frmUserForm As UserForm, _
                         Optional strControlTag As String = "All")
' --------------------------------------------------------------
' Comments:
'   Dynamically load the value of all the controls. The convention
'   is that the name of cell holding the setting matches the name
'   of the control.
'   The optional strControlTag parameter, indicates if only certain
'   types of controls are to be loaded. Useful for refreshing operations.
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 19/02/13      Carlos Gamez        Initial version
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler
    
Dim ctlCurrent As MSForms.Control
    
For Each ctlCurrent In frmUserForm.Controls
    
    If TypeOf ctlCurrent Is MSForms.TextBox Then
        ctlCurrent = fGetSetting(Right(ctlCurrent.Name, Len(ctlCurrent.Name) - 2))
    
    ElseIf TypeOf ctlCurrent Is MSForms.ComboBox Then
        
        If strControlTag = "All" Then
        
            Select Case ctlCurrent.Tag
                Case Is = "PivotFields"
                    ctlCurrent.Clear
                    ctlCurrent = fGetSetting(Right(ctlCurrent.Name, Len(ctlCurrent.Name) - 2))
                    sFillPivotFieldsComboBox frmUserForm, ctlCurrent.Name, shtTaskSummary, fSettings.cbPivotTableName.value
                Case Is = "PivotTables"
                    ctlCurrent.Clear
                    ctlCurrent = fGetSetting(Right(ctlCurrent.Name, Len(ctlCurrent.Name) - 2))
                    'Fill Pivot Table Name ComboBox
                    sFillPivotTablesComboBox frmUserForm, shtTaskSummary, ctlCurrent.Name
                Case Is = "FileNamingOptions"
                    ctlCurrent.Clear
                    'Fill File Naming Options ComboBox
                    With ctlCurrent
                        .AddItem 1
                        .List(.ListCount - 1, 1) = "<Document Number> - <Document Description>"
                        .AddItem 2
                        .List(.ListCount - 1, 1) = "<Document Number>"
                        .AddItem 3
                        .List(.ListCount - 1, 1) = "<Document Description>"
                        .AddItem 4
                        .List(.ListCount - 1, 1) = "<Table Field>"
                    End With
                    ctlCurrent = CInt(fGetSetting(Right(ctlCurrent.Name, Len(ctlCurrent.Name) - 2)))
                Case Is = "ResdTableFields"
                    ctlCurrent.Clear
                    ctlCurrent = fGetSetting(Right(ctlCurrent.Name, Len(ctlCurrent.Name) - 2))
                    'Fill Pivot Table Name ComboBox
                    sFillFieldsComboBox frmUserForm, ctlCurrent.Name, shtWorkInstructions, shtWorkInstructions.ListObjects(1).Name
            End Select
        
        ElseIf strControlTag = "PivotFields" And ctlCurrent.Tag = "PivotFields" Then
            ctlCurrent.Clear
            ctlCurrent = fGetSetting(Right(ctlCurrent.Name, Len(ctlCurrent.Name) - 2))
            sFillPivotFieldsComboBox frmUserForm, ctlCurrent.Name, shtTaskSummary, fSettings.cbPivotTableName.value
        
        ElseIf strControlTag = "PivotTables" And ctlCurrent.Tag = "PivotTables" Then
            ctlCurrent.Clear
            ctlCurrent = fGetSetting(Right(ctlCurrent.Name, Len(ctlCurrent.Name) - 2))
            sFillPivotTablesComboBox frmUserForm, shtTaskSummary, ctlCurrent.Name
        
        ElseIf strControlTag = "FileNamingOptions" And ctlCurrent.Tag = "FileNamingOptions" Then
            ctlCurrent.Clear
            With ctlCurrent
                .AddItem 1
                .List(.ListCount - 1, 1) = "<Document Number> - <Document Description>"
                .AddItem 2
                .List(.ListCount - 1, 1) = "<Document Number>"
                .AddItem 3
                .List(.ListCount - 1, 1) = "<Document Description>"
                .AddItem 4
                .List(.ListCount - 1, 1) = "<Table Field>"
            End With
            ctlCurrent = fGetSetting(Right(ctlCurrent.Name, Len(ctlCurrent.Name) - 2))
        End If
        
    ElseIf TypeOf ctlCurrent Is MSForms.CheckBox Then
        If fGetSetting(Right(ctlCurrent.Name, Len(ctlCurrent.Name) - 3)) = 1 Then
            ctlCurrent = True
        Else
            ctlCurrent = False
        End If
    End If

Next ctlCurrent

Exit_ErrHandler:
    Set ctlCurrent = Nothing
    Exit Sub

ErrHandler:
    Select Case Err.Number
    
    'Defalut error handle
    Case Else
        sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    End Select
    
    Resume Exit_ErrHandler
    Resume

End Sub

Public Sub sAssignSettings(ByRef frmUserForm As UserForm)
' --------------------------------------------------------------
' Comments:
'   Dynamically load the value of all the controls to the appropiate
'   cell in the settings worksheet
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 19/02/13      Carlos Gamez        Initial version
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler
    
Dim ctlCurrent As MSForms.Control
    
For Each ctlCurrent In frmUserForm.Controls
    
    If TypeOf ctlCurrent Is MSForms.TextBox Then
        fAssignSetting Right(ctlCurrent.Name, Len(ctlCurrent.Name) - 2), frmUserForm.Controls(ctlCurrent.Name).value
    ElseIf TypeOf ctlCurrent Is MSForms.ComboBox Then
        fAssignSetting Right(ctlCurrent.Name, Len(ctlCurrent.Name) - 2), frmUserForm.Controls(ctlCurrent.Name).value
    ElseIf TypeOf ctlCurrent Is MSForms.CheckBox Then
    
        If frmUserForm.Controls(ctlCurrent.Name).value = True Then
            fAssignSetting Right(ctlCurrent.Name, Len(ctlCurrent.Name) - 3), 1
        Else
            fAssignSetting Right(ctlCurrent.Name, Len(ctlCurrent.Name) - 3), 0
        End If
    End If

Next ctlCurrent

Exit_ErrHandler:
    Set ctlCurrent = Nothing
    Exit Sub

ErrHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    
    Resume Exit_ErrHandler
    Resume

End Sub

Public Sub sFillPivotTablesComboBox(ByRef frmUserForm As UserForm, _
                                    ByRef wksPivotTables As Excel.Worksheet, _
                                    strComboBoxName As String)
' --------------------------------------------------------------
' Comments:
'   Fill a ComboBox control in a UserForm with all the PivotTables
'   contained in a WorkSheet (wksPivotTables)
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 19/02/13      Carlos Gamez        Initial version
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler
    
    Dim intNumberOfPivots As Integer
    
    intNumberOfPivots = 0
    
    With frmUserForm.Controls(strComboBoxName)
        
        For intNumberOfPivots = 1 To wksPivotTables.PivotTables.Count
            .AddItem wksPivotTables.PivotTables(intNumberOfPivots).Name
        Next intNumberOfPivots

    End With
    
Exit_ErrHandler:
    Exit Sub

ErrHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    
    Resume Exit_ErrHandler
    Resume

End Sub

Public Sub SaveSettings(ByRef frmUser As UserForm)
' --------------------------------------------------------------
' Comments:
'   Save all the settings contained in a UserForm to the settings
'   storage table.
'
' Arguments:
'   frmUser (UserForm) = The form that the settings are taken from
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 26/01/2018    Carlos Gamez        Open source release
'
    Dim strErrors As String
    
    'Assign current settings.
    sAssignSettings frmUser
    
    'Initialise global variables
    InitializeVariables
    
    'Check for errors on the folders/tempalte names
    strErrors = fCheckFolders
    
    If strErrors <> vbNullString Then
        sDisplayUnexpectedError "File Locations", strErrors & vbCrLf & _
                                "Settings will not be saved. Please check the settings before continuing.", _
                                mstrMODULE
    Else
        ThisWorkbook.Save
    End If

End Sub

Public Sub sFillFieldsComboBox(ByRef frmUserForm As UserForm, _
                                    strComboBoxName As String, _
                                    ByRef wksFields As Excel.Worksheet, _
                                    strTableName As String)
' --------------------------------------------------------------
' Comments:
'   Fill a ComboBox control in a UserForm with all the Fields
'   contained in a Table of a specific WorkSheet (wksPivotTables)
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 19/02/13      Carlos Gamez        Initial version
' 26/01/2018    Carlos Gamez        Open source release
'
On Error GoTo ErrHandler
    
    Dim intNumberOfFields As Integer
    
    intNumberOfFields = 0
    
    With frmUserForm.Controls(strComboBoxName)
        
        For intNumberOfFields = 1 To wksFields.ListObjects(strTableName).ListColumns.Count
            .AddItem wksFields.ListObjects(strTableName).ListColumns(intNumberOfFields).Name
        Next intNumberOfFields

    End With
    
Exit_ErrHandler:
    Exit Sub

ErrHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    
    Resume Exit_ErrHandler

End Sub

Public Sub sFillPivotFieldsComboBox(ByRef frmUserForm As UserForm, _
                                    strComboBoxName As String, _
                                    ByRef wksFields As Excel.Worksheet, _
                                    strTableName As String)
' --------------------------------------------------------------
' Comments:
'   Fill a ComboBox control in a UserForm with all the Fields
'   contained in a PivotTable of a specific WorkSheet (wksPivotTables)
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 19/02/13      Carlos Gamez        Initial version
' 26/01/2018    Carlos Gamez        Open source version
'
On Error GoTo ErrHandler
    
    Dim intNumberOfFields As Integer
    Dim pvtCurrent As PivotTable
    Dim fldCurrent As PivotFields
    
    intNumberOfFields = 0
    
    With frmUserForm.Controls(strComboBoxName)
        
        Set pvtCurrent = wksFields.PivotTables(strTableName)
        Set fldCurrent = pvtCurrent.VisibleFields
        
        For intNumberOfFields = 1 To fldCurrent.Count
            .AddItem fldCurrent.Item(intNumberOfFields).Name
        Next intNumberOfFields

    End With
    
Exit_ErrHandler:
    Set pvtCurrent = Nothing
    Set fldCurrent = Nothing
    Exit Sub

ErrHandler:
    Select Case Err.Number
    'Pivot Table does not exits
    Case Is = 1004
        frmUserForm.Controls(strComboBoxName).Clear
    'Defalut error handle
    Case Else
        sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    End Select
    
    Resume Exit_ErrHandler
    Resume

End Sub

Public Function fCheckFolders() As String
' --------------------------------------------------------------
' Comments:
'   This function checks wheather the various folders and the
'   template file exists in the designated location. If all
'   checks are ok it returns a nullstring, otherwise returns a
'   string with the error descriptions.
'
' Date          Developer           Comment
' --------------------------------------------------------------
' 06/03/13      Carlos Gamez        Initial version
' 26/01/2018    Carlos Gamez        Open source version
'
On Error GoTo ErrHandler
    
    Dim strErrors As String
    
    fCheckFolders = vbNullString
    
    'Check if word template file exists
    If CheckIfFileExists(PMTemplateName) = False Then
        strErrors = "Template: The Word Template " & PMTemplateName & " could not be located." & vbCrLf
    End If
    
    'Check if image folder exists
    If CheckIfFolderExists(ImageFolder) = False Then
        strErrors = strErrors & vbCrLf & "Images: The Image Folder " & ImageFolder & " could not be located." & vbCrLf
    End If
    
    'Check if folder to save files to exists
    If CheckIfFolderExists(FolderToSaveFilesTo) = False Then
        strErrors = strErrors & vbCrLf & "Output: The Folder to save files to " & FolderToSaveFilesTo & " could not be located." & vbCrLf
    End If
    
    'Check if folder with OEM manuals exists
    If CheckIfFolderExists(ManualsFolder) = False Then
        strErrors = strErrors & vbCrLf & "Manuals: The Folder with OEM manuals " & ManualsFolder & " could not be located." & vbCrLf
    End If
    
    fCheckFolders = strErrors
    
Exit_ErrHandler:
    Exit Function

ErrHandler:
    Select Case Err.Number
    'Use cases to handle different types of errors
    
    'Defalut error handle
    Case Else
        sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    End Select
    
    Resume Exit_ErrHandler
    Resume
End Function
