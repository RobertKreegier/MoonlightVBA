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
'****               Dim Suppression As Object: Set Suppression = New CSuppression
'****
'****           When the object is initialized, the suppression takes place. When
'****           the object loses scope, garbage collection destroys it and
'****           Class_Terminate returns Excel's settings to the way they were when the
'****           object was created at the start of the procedure.
'****           Because each instance remembers the settings from before its creation
'****           and resets them, we don't have to worry about suppression objects in
'****           different procedures stepping on toes.
'***************************************************************************************

Private blnOldEvents As Boolean
Private blnOldScreen As Boolean
Private blnOldDisplay As Boolean
Private lngOldCalc As Long
Private vntOldWindowView As Variant
Private blnOldPages As Boolean

Private blnActive As Boolean

' Initialize and suppress everything, remembering the state we found everything in
Private Sub Class_Initialize()

    With Application
        ' save old states
        blnOldEvents = .EnableEvents
        lngOldCalc = .Calculation
        blnOldScreen = .ScreenUpdating
        blnOldDisplay = .DisplayAlerts
        vntOldWindowView = ActiveWindow.View
        blnOldPages = ActiveSheet.DisplayPageBreaks
    End With
    
    Activate
End Sub

Private Sub Activate()
    #If Not blnDeveloperMode Then
        On Error GoTo Whoa
    #End If
    '***********************************************************************************************

    With Application
        ' suppress everything
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .DisplayAlerts = False
        ActiveWindow.View = xlNormalView
        ActiveSheet.DisplayPageBreaks = False
    End With

    blnActive = True
    '***********************************************************************************************
Letscontinue:
    Exit Sub
Whoa:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": Activate", True
        Resume Letscontinue
    #End If
End Sub

Public Property Let Active(ByVal blnValue As Boolean)
    #If Not blnDeveloperMode Then
        On Error GoTo Whoa
    #End If
    '***********************************************************************************************
    
    If blnValue = True Then
        Activate
    Else
        Deactivate
    End If
    
    '***********************************************************************************************
Letscontinue:
    ' clear out the status bar to give an indication of if the program freezes or not
    Application.StatusBar = False
    Exit Property
Whoa:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": Let Active", True
        Resume Letscontinue
    #End If
End Property

Public Property Get Active() As Boolean
    #If Not blnDeveloperMode Then
        On Error GoTo Whoa
    #End If
    '***********************************************************************************************
    
    With Application
        ' save old states
        If blnActive Then
            Active = True
        Else
            Active = False
        End If
    End With
    
    '***********************************************************************************************
Letscontinue:
    ' clear out the status bar to give an indication of if the program freezes or not
    Application.StatusBar = False
    Exit Property
Whoa:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": Let Active", True
        Resume Letscontinue
    #End If
End Property

' Reset the environment back to some defaults
Public Sub Reset()
    blnOldEvents = True
    blnOldScreen = True
    blnOldDisplay = True
    lngOldCalc = xlCalculationAutomatic
    vntOldWindowView = xlNormalView
    blnOldPages = True
    
    With Application
        ActiveWindow.View = vntOldWindowView
        ActiveSheet.DisplayPageBreaks = blnOldPages
        .ScreenUpdating = blnOldScreen
        .DisplayAlerts = blnOldDisplay
        .Calculation = lngOldCalc
        .EnableEvents = blnOldEvents
    End With
End Sub

Private Sub Deactivate()
    #If Not blnDeveloperMode Then
        On Error GoTo Whoa
    #End If
    '***********************************************************************************************
    
    With Application
        ActiveWindow.View = vntOldWindowView
        ActiveSheet.DisplayPageBreaks = blnOldPages
        .ScreenUpdating = blnOldScreen
        .DisplayAlerts = blnOldDisplay
        .Calculation = lngOldCalc
        .EnableEvents = blnOldEvents
    End With
    
    blnActive = False
    '***********************************************************************************************
Letscontinue:
    ' clear out the status bar to give an indication of if the program freezes or not
    Application.StatusBar = False
    Exit Sub
Whoa:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": Deactivate", True
        Resume Letscontinue
    #End If
End Sub

' Return everything to the way it was before this class was initialized
Private Sub Class_Terminate()
    Deactivate
End Sub
