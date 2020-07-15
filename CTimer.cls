VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************
Option Explicit
Private Const blnDeveloperMode As Boolean = False
Private Const strModuleName As String = "CTimer"
'**** Author  : Robert M Kreegier
'**** Purpose : A simple class to encapsulate QueryPerformanceCounter
'**** Use     : To use, just initialize a CTimer object at the top of a procedure.
'****               Dim cClock As Object: Set cClock = New CTimer
'****           It will automatically record the dblBegin time. Use Begin to set another
'****           dblBegin time.
'****           Use Lap to register each lap. This will also record dblEnd and
'****           calculate the dblAverage at that moment.
'****           Use Report to write out data to the Immediate window.
'****           GetTick will return the current clock tick, in milliseconds.
'***************************************************************************************

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

#If VBA7 Then
    Private Declare PtrSafe Function getTime Lib "kernel32" Alias "QueryPerformanceCounter" (lpPerformanceCount As LARGE_INTEGER) As Long
    Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (lpFrequency As LARGE_INTEGER) As Long
#Else
    Private Declare Function getTime Lib "kernel32" Alias "QueryPerformanceCounter" (lpPerformanceCount As LARGE_INTEGER) As Long
    Private Declare Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (lpFrequency As LARGE_INTEGER) As Long
#End If

Private Const TWO_32 = 4294967296# ' = 256# * 256# * 256# * 256#

Private dlbFreq As Double

Public dblBegin As Double
Public dblEnd As Double
Public lngLaps As Long
Public dblAverage As Double

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : "Starts" the timer, recording the current time.
'*******************************************************************************
Public Sub Begin()
    On Error GoTo ProcException
    '*******************************************************************************
    
    ' Get and convert the frequency
    Dim PerfFrequency As LARGE_INTEGER
    getFrequency PerfFrequency
    dlbFreq = LI2Double(PerfFrequency)
    
    ' Start the timer
    dblBegin = GetTick

    ' Init the end
    dblEnd = dblBegin

    lngLaps = 0
    dblAverage = 0
    
    '*******************************************************************************
ExitProc:
    Exit Sub
ProcException:
    TimerThrowError strModuleName, "Begin"
    Resume ExitProc
End Sub

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Returns the current tick.
'*******************************************************************************
Public Function GetTick() As Double
    On Error GoTo ProcException
    '*******************************************************************************
    
    Dim dblTime As LARGE_INTEGER
    getTime dblTime
    GetTick = 1000# * LI2Double(dblTime) / dlbFreq
    
    '*******************************************************************************
ExitProc:
    Exit Function
ProcException:
    TimerThrowError strModuleName, "GetTick"
    Resume ExitProc
End Function

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Increments the lap counter and records the current tick.
'*******************************************************************************
Public Sub Lap()
    On Error GoTo ProcException
    '*******************************************************************************
    
    lngLaps = lngLaps + 1
    
    dblEnd = GetTick
    
    dblAverage = (dblEnd - dblBegin) / lngLaps
    
    '*******************************************************************************
ExitProc:
    Exit Sub
ProcException:
    TimerThrowError strModuleName, "Lap"
    Resume ExitProc
End Sub

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Class_Initialize
'*******************************************************************************
Private Sub Class_Initialize()
    On Error GoTo ProcException
    '*******************************************************************************
    
    Begin
    
    '*******************************************************************************
ExitProc:
    Exit Sub
ProcException:
    TimerThrowError strModuleName, "Class_Initialize"
    Resume ExitProc
End Sub

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Converts what QueryPerformanceCounter returns into milliseconds.
'*******************************************************************************
Private Function LI2Double(LI As LARGE_INTEGER) As Double
    On Error GoTo ProcException
    '*******************************************************************************
    
    Dim Low As Double
    Low = LI.lowpart
    If Low < 0 Then
        Low = Low + TWO_32
    End If
    LI2Double = LI.highpart * TWO_32 + Low
    
    '*******************************************************************************
ExitProc:
    Exit Function
ProcException:
    TimerThrowError strModuleName, "LI2Double"
    Resume ExitProc
End Function

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Writes out a report to the Immediate Window.
'*******************************************************************************
Public Sub Report()
    On Error GoTo ProcException
    '*******************************************************************************
    
    If lngLaps < 1 Then Lap

    Debug.Print "Begin: " & dblBegin
    Debug.Print "End: " & dblEnd
    Debug.Print "Laps: " & lngLaps
    
    Dim dblTotalms As Double: dblTotalms = dblEnd - dblBegin
    
    If dblTotalms > 1000 Then
        Debug.Print "Total Time: " & Round(dblTotalms / 1000, 3) & "sec" & " (" & Round(dblTotalms, 0) & "ms)"
    ElseIf dblTotalms < 1 Then
        Debug.Print "Total Time: " & Round(dblTotalms * 1000, 0) & "us" & " (" & Round(dblTotalms, 3) & "ms)"
    Else
        Debug.Print "Total Time: " & Round(dblTotalms, 3) & "ms"
    End If
    
    If dblAverage > 1000 Then
        Debug.Print "Average Lap: " & Round(dblAverage / 1000, 3) & "sec" & " (" & Round(dblAverage, 0) & "ms)"
    ElseIf dblAverage < 1 Then
        Debug.Print "Average Lap: " & Round(dblAverage * 1000, 0) & "us" & " (" & Round(dblAverage, 3) & "ms)"
    Else
        Debug.Print "Average Lap: " & Round(dblAverage, 3) & "ms"
    End If
    
    '*******************************************************************************
ExitProc:
    Exit Sub
ProcException:
    TimerThrowError strModuleName, "Report"
    Resume ExitProc
End Sub

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Our error handling procedure.
'*******************************************************************************
Private Sub TimerThrowError(ByVal strModule As String, ByVal strProcedure As String)
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
