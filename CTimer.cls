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
#Const blnDeveloperMode = False
Private Const strModuleName As String = "CTimer"
'**** Author  : Robert M Kreegier
'**** Purpose : A simple class tp encapsulate QueryPerformanceCounter
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

Private Sub Class_Initialize()
    Begin
End Sub

Private Function LI2Double(LI As LARGE_INTEGER) As Double
    Dim Low As Double
    Low = LI.lowpart
    If Low < 0 Then
        Low = Low + TWO_32
    End If
    LI2Double = LI.highpart * TWO_32 + Low
End Function

Public Function GetTick() As Double
    Dim dblTime As LARGE_INTEGER
    getTime dblTime
    GetTick = 1000# * LI2Double(dblTime) / dlbFreq
End Function

Public Sub Begin()
    ' get and convert the frequency
    Dim PerfFrequency As LARGE_INTEGER
    getFrequency PerfFrequency
    dlbFreq = LI2Double(PerfFrequency)
    
    ' Start the timer
    dblBegin = GetTick

    ' init the ends
    dblEnd = dblBegin

    lngLaps = 0
    dblAverage = 0
End Sub

Public Sub Sleep(ByVal dblMilliseconds As Double)
    Dim dblStartSleep As Double: dblStartSleep = GetTick
    
    If dblMilliseconds > 0 Then
        While (GetTick - dblStartSleep) < dblMilliseconds
        Wend
    End If
End Sub

Public Sub Lap()
    lngLaps = lngLaps + 1
    
    dblEnd = GetTick
    
    dblAverage = (dblEnd - dblBegin) / lngLaps
End Sub

Public Sub Report()
    If lngLaps < 1 Then Lap

    Debug.Print "Begin: " & dblBegin
    Debug.Print "End: " & dblEnd
    Debug.Print "Laps: " & lngLaps
    Debug.Print "Total Time: " & dblEnd - dblBegin & "ms"
    Debug.Print "Average Lap: " & dblAverage & "ms"
End Sub
