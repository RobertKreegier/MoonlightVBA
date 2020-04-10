VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CProgIndicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************
Option Explicit
#Const blnDeveloperMode = False
Private Const strModuleName As String = "CProgIndicator"
'**** Author  : Robert M Kreegier
'**** Purpose : Show progress
'**** Use     : To use, just initialize a CProgIndicator object at the top of a procedure:
'****
'****               Dim objProg As Object: Set objProg = New CProgIndicator: objProg.Init 7
'****
'****               objProg.Update "Initializing...", "Initializing styles..."
'***************************************************************************************

Private dblProgTick As Double
Private dblProgTotal As Double
Private strProgCaption As String
Private strProgStatus As String

Public Sub Init(ByVal dblTotal As Double, Optional ByVal strCaption As String = "", Optional ByVal strStatus As String = "")
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '*******************************************************************************
    
    dblProgTick = 0
    dblProgTotal = dblTotal
    strProgCaption = strCaption
    strProgStatus = strStatus
    
    ' If a caption was given, then we'll update it, otherwise there's no need to
    ' update a caption if a null string was given.
    If strCaption <> vbNullString Then
        ProgIndicatorForm.Caption = strCaption
    End If

    ' Same for the status.
    If strStatus <> vbNullString Then
        ProgIndicatorForm.Status = strStatus
    End If
    
    '*******************************************************************************
ExitProc:
    Exit Sub
ProcException:
    #If Not blnDeveloperMode Then
        ThrowError Err.Description & Chr(10) & "thrown from " & strModuleName & ": Init, Line: " & Erl
        Resume ExitProc
    #End If
End Sub

Public Sub Update(Optional ByVal strCaption As String = "", Optional ByVal strStatus As String = "")
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '*******************************************************************************

    ' If a caption was given, then we'll update it, otherwise there's no need to
    ' update a caption if a null string was given.
    If strCaption <> vbNullString Then
        strProgCaption = strCaption
        ProgIndicatorForm.Caption = strCaption
    End If

    ' Same for the status.
    If strStatus <> vbNullString Then
        strProgStatus = strStatus
        ProgIndicatorForm.Status = strStatus
    End If

    ' if the given total is above zero, then we can do some things
    If dblProgTotal > 0 Then
        dblProgTick = dblProgTick + 1
        
        ' calculate the percentage and show it in the text
        ProgIndicatorForm.Text.Caption = strProgCaption & ": " & Math.Round((dblProgTick / dblProgTotal) * 100) & "%"

        ' update the width of the progress bar
        ProgIndicatorForm.Bar.Top = 0
        ProgIndicatorForm.Bar.Left = 0
        ProgIndicatorForm.Bar.Height = ProgIndicatorForm.Frame1.Height - 4
        ProgIndicatorForm.Bar.Width = Math.Round((dblProgTick / dblProgTotal) * (ProgIndicatorForm.Frame1.Width - 4))

        ' if the tick count equals the total, then we know the progress is done for this task
        If dblProgTick >= dblProgTotal Then
            ' make sure tick and total are the same
            dblProgTick = dblProgTotal

            ProgIndicatorForm.Hide
            ProgIndicatorForm.Text.Caption = strProgCaption & ": 0%"
            ProgIndicatorForm.Bar.Width = 0
            ProgIndicatorForm.Caption = vbNullString
            ProgIndicatorForm.Status = vbNullString

        ' the tick does not equal the total
        Else
            ' and make sure the window is showing
            ProgIndicatorForm.Show vbModeless
        End If
    End If
    
    ProgIndicatorForm.Repaint

    '*******************************************************************************
ExitProc:
    Exit Sub
ProcException:
    #If Not blnDeveloperMode Then
        ThrowError Err.Description & Chr(10) & "thrown from " & strModuleName & ": Update, Line: " & Erl
        Resume ExitProc
    #End If
End Sub

Private Sub Deactivate()
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '*******************************************************************************
    
    dblProgTick = dblProgTotal
    ProgIndicatorForm.Hide
    ProgIndicatorForm.Text.Caption = strProgCaption & ": 0%"
    ProgIndicatorForm.Bar.Width = 0
    ProgIndicatorForm.Caption = vbNullString
    ProgIndicatorForm.Status = vbNullString
    
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
    Deactivate
End Sub
