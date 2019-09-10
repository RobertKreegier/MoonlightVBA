'***************************************************************************************
Option Explicit
#Const blnDeveloperMode = False
Private Const strModuleName As String = "PERFORMANCE"
'**** Author : Robert M Kreegier
'**** Purpose: Procedures geared toward testing the performance of processes in VBA
'***************************************************************************************

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Private StartTime As Long
Private EndTime As Long

Sub StartTimer()
    StartTime = GetTickCount
    Debug.Print "Start: " & StartTime
End Sub

Sub EndTimer()
    EndTime = GetTickCount
    Debug.Print "End: " & EndTime
    Debug.Print "Total execution time: " & EndTime - StartTime
End Sub

Function TestProc(ByVal strProcedureName As String, Optional ByVal lngIterations As Long = 10) As Double
    ' we'll start the timer, run the procedure a bunch of times, then stop the timer
    ' we subtract the start and stop times, then divide by the number of times we ran the procedure
    ' this should get us a good average with sub-millisecond accuracy (significant figures dictated by lngIterations)
    Dim Suppression As Object: Set Suppression = New CSuppression
    
    Debug.Print "Testing " & strProcedureName & " with " & lngIterations & " iterations."
    
    StartTimer
    
    Dim lngCount As Long
    For lngCount = 1 To lngIterations
        Application.Run strProcedureName
    Next
    
    EndTimer

    Debug.Print "Average execution time: " & (EndTime - StartTime) / lngIterations & "ms"
    TestProc = (EndTime - StartTime) / lngIterations
End Function

Sub CountCodeLines(Optional ByVal HideFromMacroList As Boolean = False)
    Dim VBCodeModule As Object
    Dim NumLines As Long, N As Long
    With ThisWorkbook
          For N = 1 To .VBProject.VBComponents.Count
                Set VBCodeModule = .VBProject.VBComponents(N).CodeModule
                NumLines = NumLines + VBCodeModule.CountOfLines
          Next
    End With
    Debug.Print "Total number of lines of code in the project = " & NumLines
    Set VBCodeModule = Nothing
End Sub

' ?TestProc("TestContainer", 100)
Sub TestContainer(Optional ByVal HideFromMacroList As Boolean = False)
'    Dim Suppression As Object: Set Suppression = New CSuppression

    Dim strTemp As String
    
    strTemp = GetVar("V_SAVE_AS").Value
End Sub
