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
Private Const blnDeveloperMode As Boolean = False
Private Const strModuleName As String = "CProgIndicator"
'**** Author  : Robert M Kreegier
'**** Purpose : Show progress
'**** Use     : To use, just initialize a CProgIndicator object at the top of a procedure:
'****
'****               Dim objProg As Object
'****               Set objProg = New CProgIndicator
'****
'****               ' Initialize the progress indicator to expect 3 updates, and start it off with a caption and status
'****               objProg.Init 3,"Initializing...", "Initializing things..."
'****
'****               ' The first update with a caption and status
'****               objProg.Update "Progress...", "Doing thing 1..."
'****
'****               ' The second update with a caption and status
'****               objProg.Update "Progress...", "Doing thing 2..."
'****
'****               ' The third update, which closes the progress
'****               objProg.Update "Progress...", "Done!"
'***************************************************************************************

Private dblProgTick As Double
Private dblProgTotal As Double
Private strProgCaption As String
Private strProgStatus As String

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Return everything to the way it was before this class was
'           initialized.
'*******************************************************************************
Private Sub Deactivate()
    On Error GoTo ProcException
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
    ProgThrowError strModuleName, "Deactivate"
    Resume ExitProc
End Sub

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Initialize the progress indicator by telling it how many updates
'           there will be, what the initial dialog caption will be, and what
'           the initial status string will be.
'*******************************************************************************
Public Sub Init(ByVal dblTotal As Double, Optional ByVal strCaption As String = "", Optional ByVal strStatus As String = "")
    On Error GoTo ProcException
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
    ProgThrowError strModuleName, "Init"
    Resume ExitProc
End Sub

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Our error handling procedure.
'*******************************************************************************
Private Sub ProgThrowError(ByVal strModule As String, ByVal strProcedure As String)
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
' Purpose : Update the progress indicator with new info. This also increments
'           dblProgTick so we keep track of how many updates there's been to the
'           progess.
'*******************************************************************************
Public Sub Update(Optional ByVal strCaption As String = "", Optional ByVal strStatus As String = "")
    On Error GoTo ProcException
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
    ProgThrowError strModuleName, "Update"
    Resume ExitProc
End Sub

Private Sub Class_Terminate()
    Deactivate
End Sub
