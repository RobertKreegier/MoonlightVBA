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
Option Explicit
#Const blnDeveloperMode = False
Private Const strModuleName As String = "CSuppression"
'**** Author  : Robert M Kreegier
'**** Purpose : Suppress Excel's automatic event processing, screen updating,
'****           calculation, and alerts. This also sets the window mode to normal and
'****           removes page breaks. All these things are to increase speed and reduce
'****           the chances that Excel's processing will interfere with our own code.
'**** Use     : To use, just initialize a CSuppression object at the top of a procedure:
'****
'****               Dim Suppression As Object: Set Suppression = New CSuppression: Suppression.Activate
'****
'****           When the object is initialized, the suppression takes place. When
'****           the object loses scope, garbage collection destroys it and
'****           Class_Terminate returns Excel's settings to the way they were when the
'****           object was created at the start of the procedure.
'****           Because each instance remembers the settings from before its creation
'****           and resets them, we don't have to worry about suppression objects in
'****           different procedures stepping on toes.
'***************************************************************************************

Private Const blnDefaultEvents = True
Private Const lngDefaultCalc = xlCalculationAutomatic
Private Const blnDefaultScreen = True
Private Const blnDefaultAlerts = True
Private Const vntDefaultView = xlNormalView
Private Const blnDefaultBreaks = True
    
Private blnOldEvents As Boolean
Private lngOldCalc As Long
Private blnOldScreen As Boolean
Private blnOldAlerts As Boolean
Private vntOldView As Variant
Private blnOldBreaks As Boolean

Private blnActive As Boolean

' Initialize and suppress everything, remembering the state we found everything in
Private Sub Class_Initialize()
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
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
    #If Not blnDeveloperMode Then
        ThrowError Err.Description & Chr(10) & "thrown from " & strModuleName & ": Class_Initialize, Line: " & Erl
        Resume ExitProc
    #End If
End Sub

Public Sub Activate(Optional ByVal blnEvents As Boolean = False, _
                    Optional ByVal lngCalc As Long = xlCalculationManual, _
                    Optional ByVal blnScreen As Boolean = False, _
                    Optional ByVal blnAlerts As Boolean = False, _
                    Optional ByVal vntView As Variant = xlNormalView, _
                    Optional ByVal blnBreaks As Boolean = False)
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '*******************************************************************************

    With Application
        ' suppress everything
        If .EnableEvents <> blnEvents Then .EnableEvents = blnEvents
        If .Calculation <> lngCalc Then .Calculation = lngCalc
        If .ScreenUpdating <> blnScreen Then .ScreenUpdating = blnScreen
        If .DisplayAlerts <> blnAlerts Then .DisplayAlerts = blnAlerts
    End With

    On Error Resume Next

    If ActiveWindow.View <> vntView Then ActiveWindow.View = vntView
    If ActiveSheet.DisplayPageBreaks <> blnBreaks Then ActiveSheet.DisplayPageBreaks = blnBreaks
    
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #Else
        On Error GoTo 0
    #End If

    blnActive = True
    
    '*******************************************************************************
ExitProc:
    ' clear out the status bar to give an indication of if the program freezes or not
    Application.StatusBar = False
    Exit Sub
ProcException:
    #If Not blnDeveloperMode Then
        ThrowError Err.Description & Chr(10) & "thrown from " & strModuleName & ": Activate, Line: " & Erl
        Resume ExitProc
    #End If
End Sub

Public Property Let Active(ByVal blnValue As Boolean)
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '*******************************************************************************
    If blnValue = True Then
        Activate
    Else
        Deactivate
    End If
    '*******************************************************************************
ExitProc:
    Exit Property
ProcException:
    #If Not blnDeveloperMode Then
        ThrowError Err.Description & Chr(10) & "thrown from " & strModuleName & ": Let Active, Line: " & Erl
        Resume ExitProc
    #End If
End Property

Public Property Get Active() As Boolean
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '*******************************************************************************
    Active = blnActive
    '*******************************************************************************
ExitProc:
    Exit Property
ProcException:
    #If Not blnDeveloperMode Then
        ThrowError Err.Description & Chr(10) & "thrown from " & strModuleName & ": Get Active, Line: " & Erl
        Resume ExitProc
    #End If
End Property

Public Sub Do_Events()
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
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
    #If Not blnDeveloperMode Then
        ThrowError Err.Description & Chr(10) & "thrown from " & strModuleName & ": Do_Events, Line: " & Erl
        Resume ExitProc
    #End If
End Sub

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : An Undo specifically geared toward this application so that Undo actions don't trigger
'           worksheet events.
'*******************************************************************************
Public Sub Undo(Optional ByVal strMessage As String = vbNullString)
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
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
    #If Not blnDeveloperMode Then
        ThrowError Err.Description & Chr(10) & "thrown from " & strModuleName & ": Undo, Line: " & Erl
        Resume ExitProc
    #End If
End Sub

Private Sub ThrowError(Optional ByVal strErrorMessage As String = "")
    #If Not blnDeveloperMode Then
        MsgBox strErrorMessage
    #Else
        Err.Raise vbObjectError + 1000, "ERROR", strErrorMessage
    #End If
End Sub

' Reset the environment back to some defaults
Public Sub Reset()
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
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
    #If Not blnDeveloperMode Then
        ThrowError Err.Description & Chr(10) & "thrown from " & strModuleName & ": Reset, Line: " & Erl
        Resume ExitProc
    #End If
End Sub

Public Sub Deactivate()
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '*******************************************************************************
    
    With Application
        If .EnableEvents <> blnOldEvents Then .EnableEvents = blnOldEvents
        If .Calculation <> lngOldCalc Then .Calculation = lngOldCalc
        If .ScreenUpdating <> blnOldScreen Then .ScreenUpdating = blnOldScreen
        If .DisplayAlerts <> blnOldAlerts Then .DisplayAlerts = blnOldAlerts
    End With
    
    If ActiveWindow.View <> vntOldView Then ActiveWindow.View = vntOldView
    If ActiveSheet.DisplayPageBreaks <> blnOldBreaks Then ActiveSheet.DisplayPageBreaks = blnOldBreaks
    
    blnActive = False
    '*******************************************************************************
ExitProc:
    Exit Sub
ProcException:
    #If Not blnDeveloperMode Then
        ThrowError Err.Description & Chr(10) & "thrown from " & strModuleName & ": Deactivate, Line: " & Erl
        Resume ExitProc
    #End If
End Sub

' Return everything to the way it was before this class was initialized
Private Sub Class_Terminate()
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '*******************************************************************************
    Deactivate
    '*******************************************************************************
ExitProc:
    Exit Sub
ProcException:
    #If Not blnDeveloperMode Then
        ThrowError Err.Description & Chr(10) & "thrown from " & strModuleName & ": Class_Terminate, Line: " & Erl
        Resume ExitProc
    #End If
End Sub
