VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStatProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************
Option Explicit
#Const blnDeveloperMode = False
Private Const strModuleName As String = "CStatProg"
'**** Author : Marcus Mangelsdorf (https://stackoverflow.com/users/2822719/marcus-mangelsdorf), Robert M Kreegier
'**** Source : https://stackoverflow.com/questions/5181164/progress-bar-in-vba-excel
'**** Purpose: Class for a progress bar in the status bar
'***************************************************************************************

Private statusBarState As Boolean
Private Const NUM_BARS As Integer = 50
Private Const MAX_LENGTH As Integer = 255
Private BAR_CHAR As String
Private SPACE_CHAR As String

Private Sub Class_Initialize()
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '*******************************************************************************
    ' Init the class
    
    ' Save the state of the status bar
    statusBarState = Application.DisplayStatusBar

    ' set the progress bar chars (should be equal size)
    BAR_CHAR = ChrW(9608)
    SPACE_CHAR = ChrW(9620)
    
    ' Set the desired state of the status bar
    Application.DisplayStatusBar = True

    '*******************************************************************************
ExitProc:
    Exit Sub
ProcException:
    #If Not blnDeveloperMode Then
        ThrowError Err.Description & Chr(10) & "thrown from " & strModuleName & ": Class_Initialize, Line: " & Erl
        Resume ExitProc
    #End If
End Sub

Private Sub Class_Terminate()
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '*******************************************************************************
    ' Clean up
    
    Application.DisplayStatusBar = statusBarState
    Application.StatusBar = False
    
    '*******************************************************************************
ExitProc:
    Exit Sub
ProcException:
    #If Not blnDeveloperMode Then
        ThrowError Err.Description & Chr(10) & "thrown from " & strModuleName & ": Class_Terminate, Line: " & Erl
        Resume ExitProc
    #End If
End Sub

Public Sub Update(ByVal value As Long, _
                  Optional ByVal MaxValue As Long = 0, _
                  Optional ByVal Status As String = "", _
                  Optional ByVal DisplayPercent As Boolean = True)
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '*******************************************************************************
    
    ' Value          : 0 to 100 (if no max is set)
    ' Value          : >=0 (if max is set)
    ' MaxValue       : >= 0
    ' Status         : optional message to display for user
    ' DisplayPercent : Display the percent complete after the status bar

    ' <Status> <Progress Bar> <Percent Complete>

    ' Validate entries
    If value < 0 Or MaxValue < 0 Or (value > 100 And MaxValue = 0) Then Exit Sub

    ' If the maximum is set then adjust value to be in the range 0 to 100
    If MaxValue > 0 Then value = WorksheetFunction.RoundUp((value * 100) / MaxValue, 0)

    ' Message to set the status bar to
    Dim display As String
    display = Status & "  "

    ' Set bars
    display = display & String(Int(value / (100 / NUM_BARS)), BAR_CHAR)
    ' set spaces
    display = display & String(NUM_BARS - Int(value / (100 / NUM_BARS)), SPACE_CHAR)

    ' Closing character to show end of the bar
    display = display & BAR_CHAR

    If DisplayPercent = True Then display = display & "  (" & value & "%)  "

    ' chop off to the maximum length if necessary
    If Len(display) > MAX_LENGTH Then display = Right(display, MAX_LENGTH)

    Application.StatusBar = display
    
    '*******************************************************************************
ExitProc:
    Exit Sub
ProcException:
    #If Not blnDeveloperMode Then
        ThrowError Err.Description & Chr(10) & "thrown from " & strModuleName & ": Update, Line: " & Erl
        Resume ExitProc
    #End If
End Sub

