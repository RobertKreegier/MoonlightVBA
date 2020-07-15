VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSuppression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************
'*** Gratuitous Best Practices Notes ***************************************************
'***************************************************************************************
' You can use this class module as a reference for VBA best practices (as I see fit,
' anyway).
'
' At the top of every module, you should declare Option Explicit (which you'll see
' below). Without it VBA allows you to declare a variable simply by using it. With
' Option Explicit, we're telling VBA that we're always going to deliberately declare our
' variables with a Dim. This is good practice as it avoids errors and bugs from typos.

' On the next line, I declare a constant flag blnDeveloperMode to tell the method
' SupThrowError how to handle errors (explained in a second).
'
' At the top of nearly every procedure, is the line:
'
'    On Error GoTo ProcException
'
' And at the bottom of the procedure:
'
'   ExitProc:
'        Exit Sub
'   ProcException:
'        SupThrowError strModuleName, "Class_Initialize"
'        Resume ExitProc
'
' "On Error GoTo ProcException" informs the compiler that we want to divert execution to
' the label ProcException when an error is raised. SupThrowError uses the flag
' blnDeveloperMode to determine if it should handle the error, or let VBA do the error
' handling. SupThrowError is designed to be user friendly, and handling the error
' ourselves allows execution to Resume to the label ExitProc, which allows us to deal
' with any clean up at the end of the method, keeping execution moving. Without that,
' VBA can only stop execution altogether, or allow the user to see where the error is in
' code (not good).
'
' Dealing with errors elegantly in your code is especially important when using this
' class. If the computer encounters an error and just stops execution on the spot, we
' could potentially be left without events or calculation re-enabled. That puts the user
' in the annoying position of having Excel basically broken to them.
'
' Some form of error handling should be used by default in all procedures, as a best
' practice, even when the procedure is as simple as a single line that does division...
' a divide by zero could be a game-ender. You never know what kind of errors will pop
' up.
'
' The next line is a constant string letting SupThrowError know what the name of the
' module is (SupThrowError takes two parameters: the module name, and the procedure
' name). There's no way in code to reference the module of where code execution is
' currently taking place. For that matter, there's no way to know the procedure name
' either. So we use strModuleName to tell SupThrowError which module we're in, and we
' explicitly tell SupThrowError which procedure we're in.
'
' After that, you should have a header comment that designates the author and purpose of
' the module. I also include usage notes here, so I have a quick line to copy/paste into
' new procedures.
'
' Just the same, before each procedure there's an attribution along with notes. It
' seems redundant here because every procedure is written by me, but there will
' definitely be times where you're either collaborating on code with others or copying
' code off the internet. I prefer to know who wrote what, just in case. If the code
' came from the internet, I'll also include the URL where I found it.

'***************************************************************************************
Option Explicit
Private Const blnDeveloperMode As Boolean = False
Private Const strModuleName As String = "CSuppression"
'**** Author  : Robert M Kreegier
'**** Purpose : Suppress Excel's automatic event processing, screen updating,
'****           automatic calculation, and alerts. This also sets the window mode to
'****           normal and removes page breaks. All these things are to increase speed
'****           and reduce the chances that Excel's processing will interfere with our
'****           own code.
'****           It should be noted that the process of suppressing these things and then
'****           setting Excel back to normal takes it's own amount of time. Tests run on
'****           my computer show that running a procedure with only suppression
'****           activated then deactivated takes about 45ms. This is in contrast to
'****           most procedures that can execute in sub-millisecond times.
'****           Therefor in implementing this object, you should consider the needs and
'****           trade-offs: On the one hand, the suppression might be needed in order
'****           to avoid infinite loops or minimize screen flickering. Just as well, if
'****           there's a procedure that modifies a lot of cells, it's almost always
'****           beneficial to enable suppression, as modifying anything in the interface
'****           takes a lot of time. 45ms up front might save you minutes of processing
'****           time down the line. On the other hand, if the user is doing something
'****           over and over again (like quickly entering data into a spreadsheet),
'****           45ms actually becomes noticable.
'****
'****           In general, however, I always activate suppression at the top of
'****           "external" procedures. In other words, by default I put suppression in
'****           procedures that are the first to be executed from the user interface.
'****           For instance, I'll put suppression at the top of an event handler, or a
'****           Button_Click() handler. This ensures that any processing done "down the
'****           line" is ready to go with the interface suppressed. If any procedures
'****           try activating suppression when suppression is already activated, there
'****           is very little effect on performance.
'****
'**** Use     : To use, just initialize a CSuppression object at the top of a procedure:
'****
'****               Dim Suppression As Object
'****               Set Suppression = New CSuppression
'****               Suppression.Activate
'****
'****           I often put these on a single line...
'****
'****           Dim Suppression As Object: Set Suppression = New CSuppression: Suppression.Activate
'****
'****           Suppression takes place with Suppression.Activate. When the object loses
'****           scope, garbage collection destroys it and Class_Terminate returns
'****           Excel's settings to the way they were when the object was created.
'****           Because each instance remembers the settings from before its creation
'****           and resets to them, we don't have to worry about suppression objects in
'****           different procedures stepping on toes. Every object gets destroyed in
'****           the reverse order it was created.
'***************************************************************************************

' These are my prefered defaults for Excel. If we're resetting things, we'll reset to
' these values.
Private Const blnDefaultEvents = True
Private Const lngDefaultCalc = xlCalculationAutomatic
Private Const blnDefaultScreen = True
Private Const blnDefaultAlerts = True
Private Const vntDefaultView = xlNormalView
Private Const blnDefaultBreaks = True

' These properties save Excel's current state when suppression gets activated.
Private blnOldEvents As Boolean
Private lngOldCalc As Long
Private blnOldScreen As Boolean
Private blnOldAlerts As Boolean
Private vntOldView As Variant
Private blnOldBreaks As Boolean

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Use this method to activate the suppression. Use the parameters to
'           selectively suppress certain things.
'*******************************************************************************
Public Sub Activate(Optional ByVal blnEvents As Boolean = False, _
                    Optional ByVal lngCalc As Long = xlCalculationManual, _
                    Optional ByVal blnScreen As Boolean = False, _
                    Optional ByVal blnAlerts As Boolean = False, _
                    Optional ByVal vntView As Variant = xlNormalView, _
                    Optional ByVal blnBreaks As Boolean = False)
    On Error GoTo ProcException
    '*******************************************************************************

    With Application
        ' Suppress everything.
        If .EnableEvents <> blnEvents Then .EnableEvents = blnEvents
        If .Calculation <> lngCalc Then .Calculation = lngCalc
        If .ScreenUpdating <> blnScreen Then .ScreenUpdating = blnScreen
        If .DisplayAlerts <> blnAlerts Then .DisplayAlerts = blnAlerts
    End With

    If ActiveWindow.View <> vntView Then ActiveWindow.View = vntView
    If ActiveSheet.DisplayPageBreaks <> blnBreaks Then ActiveSheet.DisplayPageBreaks = blnBreaks
    
    ' Clear out the status bar.
    Application.StatusBar = False
    '*******************************************************************************
ExitProc:
    Exit Sub
ProcException:
    SupThrowError strModuleName, "Activate"
    Resume ExitProc
End Sub

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : This method, as the name implies, deactivates the suppression.
'*******************************************************************************
Public Sub Deactivate()
    On Error GoTo ProcException
    '*******************************************************************************
    
    With Application
        If .EnableEvents <> blnOldEvents Then .EnableEvents = blnOldEvents
        If .Calculation <> lngOldCalc Then .Calculation = lngOldCalc
        If .ScreenUpdating <> blnOldScreen Then .ScreenUpdating = blnOldScreen
        If .DisplayAlerts <> blnOldAlerts Then .DisplayAlerts = blnOldAlerts
    End With
    
    If ActiveWindow.View <> vntOldView Then ActiveWindow.View = vntOldView
    If ActiveSheet.DisplayPageBreaks <> blnOldBreaks Then ActiveSheet.DisplayPageBreaks = blnOldBreaks
    
    '*******************************************************************************
ExitProc:
    Exit Sub
ProcException:
    SupThrowError strModuleName, "Deactivate"
    Resume ExitProc
End Sub

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : In certain situations, we may want to process events and update the
'           screen with the results in the middle of processing things. We can
'           use this method to do just that.
'*******************************************************************************
Public Sub Do_Events()
    On Error GoTo ProcException
    '*******************************************************************************
    
    Dim blnSavedState As Boolean
    
    With Application
        blnSavedState = .ScreenUpdating
        
        .ScreenUpdating = blnOldScreen
        DoEvents
        .ScreenUpdating = blnSavedState
    End With
    
    '*******************************************************************************
ExitProc:
    Exit Sub
ProcException:
    SupThrowError strModuleName, "Do_Events"
    Resume ExitProc
End Sub

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Reset the environment back to pre-defined defaults (at the top of
'           the module). Note that this deactivates suppression (which allows
'           calculation to happen and events to trigger). If you want
'           suppression start back up after this, you have to Activate again.
'*******************************************************************************
Public Sub Reset()
    On Error GoTo ProcException
    '*******************************************************************************
    
    blnOldEvents = blnDefaultEvents
    lngOldCalc = lngDefaultCalc
    blnOldScreen = blnDefaultScreen
    blnOldAlerts = blnDefaultAlerts
    vntOldView = vntDefaultView
    blnOldBreaks = blnDefaultBreaks
    
    Deactivate
    
    '*******************************************************************************
ExitProc:
    Exit Sub
ProcException:
    SupThrowError strModuleName, "Reset"
    Resume ExitProc
End Sub

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Our error handling procedure.
'*******************************************************************************
Private Sub SupThrowError(ByVal strModule As String, ByVal strProcedure As String)
    If Not blnDeveloperMode Then
        If strModule <> vbNullString Then
            strModule = strModule & ": "
        End If
        
        If strProcedure <> vbNullString Then
            strModule = strModule & strProcedure & ", "
            
            If Erl > 0 Then
                strModule = strModule & "Line: " & Erl
            End If
        End If
        
        If strModule <> vbNullString Then
            strModule = strModule & Chr(10) & Chr(10)
        End If
        
        Dim strErrorMessage As String
        MsgBox strModule & Err.Description
    Else
        Err.Raise vbObjectError + 1000, "ERROR", strErrorMessage
    End If
End Sub

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : An Undo specifically geared so Undo actions don't trigger worksheet
'           events.
'*******************************************************************************
Public Sub Undo(Optional ByVal strMessage As String = vbNullString)
    On Error GoTo ProcException
    '*******************************************************************************
    
    On Error Resume Next
    If Application.EnableEvents Then
        Application.EnableEvents = False
        Application.Undo
        Application.EnableEvents = True
    Else
        Application.Undo
    End If
    
    If strMessage <> vbNullString Then Application.StatusBar = "Suppression: Undo: " & strMessage
    
    '*******************************************************************************
ExitProc:
    Exit Sub
ProcException:
    SupThrowError strModuleName, "Undo"
    Resume ExitProc
End Sub

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Initialize, remembering the state we found everything in.
'*******************************************************************************
Private Sub Class_Initialize()
    On Error GoTo ProcException
    '*******************************************************************************
    
    ' save old states
    With Application
        blnOldEvents = .EnableEvents
        lngOldCalc = .Calculation
        blnOldScreen = .ScreenUpdating
        blnOldAlerts = .DisplayAlerts
    End With
    
    vntOldView = ActiveWindow.View
    blnOldBreaks = ActiveSheet.DisplayPageBreaks
    
    '*******************************************************************************
ExitProc:
    Exit Sub
ProcException:
    SupThrowError strModuleName, "Class_Initialize"
    Resume ExitProc
End Sub

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : When the object is destroyed, Class_Terminate is automatically
'           called. We use this as an opportunity to return Excel's environment
'           back to how we found it when Activate was called.
'*******************************************************************************
Private Sub Class_Terminate()
    On Error GoTo ProcException
    '*******************************************************************************
    
    Deactivate
    
    '*******************************************************************************
ExitProc:
    Exit Sub
ProcException:
    SupThrowError strModuleName, "Class_Terminate"
    Resume ExitProc
End Sub
