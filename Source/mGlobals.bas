Attribute VB_Name = "MGlobals"
'
' Description:  This module holds declarations for global constants,
'               variables, type structures, and DLLs.
'
' Authors:      Carlos Gamez
'
Option Explicit
Option Private Module

' **************************************************************
' Module Constant Declarations Follow
' **************************************************************
Private Const mstrMODULE As String = "mGlobals"
' **************************************************************
' Global Variable Declarations Follow
' **************************************************************
    Public dbReferencesLoaded  As Boolean
    Public gbolShutdownInProgress As Boolean
    Public gstrAPP_NAME As String
    Public gstrAPP_VERSION As String
    Public gstrDATE_LAST_CHANGE As String
    Public gstrAUTHOR As String
    Public gstrEMAIL As String
    Public gbolDEBUG As Boolean

    Public Enum enmContentType
        ctText = 2
        ctParagraph = 4
        ctHyperlink = 8
        ctImage = 16
        ctSection = 32
        ctTable = 64
        ctTemplateBlock = 128
    End Enum
    
    Public Enum enmContentLocation
        clHeader = 2
        clBody = 4
        clFooter = 8
        clTable = 16
    End Enum
    
    Public Enum ContentMarker
        cmBookmark
        cmTag
    End Enum
    
    Public Enum enmContentInsertion
        ciBefore
        ciAfter
    End Enum
    
Public Sub InitGlobals()
    ' ---------------------------------------------------------------------------
    ' Purpose:
    '   This routine initialises all global variables
    '
    ' Date          Developer           Comment
    ' ---------------------------------------------------------------------------
    ' 04/02/12      Carlos Gamez        Added App Name and Version
    ' 07/08/12      Carlos Gamez        Added Author, Email, Debug
    ' 19/02/13      Carlos Gamez        Added if statements
    
    gbolShutdownInProgress = False
    
    If gstrAPP_NAME = vbNullString Then gstrAPP_NAME = fGetSetting("rngAPP_NAME")
    If gstrAPP_VERSION = vbNullString Then gstrAPP_VERSION = fGetSetting("rngAPP_VERSION")
    If gstrDATE_LAST_CHANGE = vbNullString Then gstrDATE_LAST_CHANGE = fGetSetting("rngDATE_LAST_CHANGE")
    If gstrAUTHOR = vbNullString Then gstrAUTHOR = fGetSetting("rngAUTHOR")
    If gstrEMAIL = vbNullString Then gstrEMAIL = fGetSetting("rngEMAIL")
    gbolDEBUG = fGetSetting("rngDEBUG")
    
    If gbolDEBUG = True Then
        Debug.Print gstrAPP_NAME, gstrAPP_VERSION, gstrDATE_LAST_CHANGE
    End If
    
End Sub
