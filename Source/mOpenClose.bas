Attribute VB_Name = "mOpenClose"
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
' Description:  This module contains the application startup and shutdown code.
'               Many of the techniques I've used in this application
'               came from the concepts I learned on the book "Professional Excel
'               Development". I highly recommend it for any Excel developer that
'               wants to take her applications to a more professional level.
'
' Authors:      Carlos Gamez
'
'
Option Explicit
Option Private Module

' **************************************************************
' Module Constant Declarations Follow
' **************************************************************
Private Const mstrMODULE As String = "mOpenClose"

Public Sub Auto_Open()
' --------------------------------------------------------------
' Comments: This routine is run every time the application is
'           run. It handles initialization of the application.
'
' Arguments:    None
'
' Date          Developer       Comment
' --------------------------------------------------------------
' 05/11/11      Carlos Gamez    Initial version
' 17/11/11      Carlos Gamez    Replaced CustomToolbars for Custom Ribbon Tab
' 26/01/2018    Carlos Gamez    Open source release
'
    Const szSOURCE As String = "Auto_Open"
    
    Dim bOpenFailed As Boolean
    
    On Error GoTo ErrorHandler
    
        Application.ScreenUpdating = False
        
        'Initialize global variables.
        InitGlobals
        
        'Load references
        LoadReferences
    
        'Select whether to build hierarchy by hand or import from database
    
        'Goto tblWorkInstructions tab by default
        shtWorkInstructions.Activate
        
        Application.ScreenUpdating = True

Exit_ErrHandler:

    ' Reset critical application properties.
    ResetAppProperties
    If bOpenFailed Then ShutdownApplication
    Exit Sub
    
ErrorHandler:
    sDisplayUnexpectedError Err.Number, Err.Description, mstrMODULE
    
    Resume Exit_ErrHandler
    Resume
    
End Sub

Public Sub Auto_Close()
' --------------------------------------------------------------
' Comments: This routine runs automatically every time the
'           application workbook is closed. If the application
'           shutdown code is not already running (as the result
'           of a call from the Exit button), this procedure
'           calls the application shutdown procedure.
'
' Arguments:    None
'
' Date          Developer       Comment
' --------------------------------------------------------------
' 04/29/04      Rob Bovey       Initial version
'
' Call standard shutdown code if it isn't already running.
    If Not gbolShutdownInProgress Then ShutdownApplication
End Sub

Public Sub ShutdownApplication()
' --------------------------------------------------------------
' Comments: This routine shuts down the application.
'
' Arguments:    None
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 03/15/04      Rob Bovey       Ch05    Initial version
' 17/11/11      Carlos Gamez    Replaced CustomToolbars for Custom Ribbon Tab
'
On Error Resume Next
    
    Application.ScreenUpdating = False
    
    ' This flag prevents this routine from being called a second time
    ' by Auto_Close if has already been called by another procedure.
    gbolShutdownInProgress = True
    
    UnloadReferences
    
End Sub

Public Sub ResetAppProperties()
' --------------------------------------------------------------
' Comments: Called by all entry point procedures before exiting
'           to ensure that all application properties are
'           correctly restored.
'
' Arguments:    None
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 03/15/04      Rob Bovey       Ch08    Initial version
'
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.EnableCancelKey = xlInterrupt
    Application.Cursor = xlDefault
    
End Sub

