VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTiming"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CTiming - (c) 2004 by Donald Lessau, www.xbeat.net
' total rewrite of old CTimingPC
' created: 20040614
' updated: 20040914, 20050826

' usage:
' Module declare: Private tmr as CTiming
' Make Instance:  Set tmr = New CTiming
' start timer:    tmr.Reset
' stop/get time:  dTime = tmr.Elapsed   'returns Double millisec with microsec as fraction
'                 sTime = tmr.sElapsed  'returns String millisec with microsec as fraction and " msec"

Option Explicit

' LARGE_INTEGER is faster than Currency type
' Currency requires CPU to execute slow floating-point instructions
Private Type LARGE_INTEGER
  Lo As Long
  Hi As Long
End Type

Private Declare Function QueryPerformanceCounter Lib "kernel32" ( _
    lpPerformanceCount As LARGE_INTEGER) As Long

Private Declare Function QueryPerformanceFrequency Lib "kernel32" ( _
    lpFrequency As LARGE_INTEGER) As Long

Private uFreq           As LARGE_INTEGER
Private uStart          As LARGE_INTEGER
Private uEnd            As LARGE_INTEGER
Private uPauseStart     As LARGE_INTEGER

Private dPauseDuration  As Double
Private dOverhead       As Double

Private fAvailable As Boolean

Private Sub Class_Initialize()
  Const overheadLoopCount As Long = 100
  Dim i As Long
  
  ' returns ticks/sec
  If QueryPerformanceFrequency(uFreq) = 0& Then
    
    ' some CPUs do NOT support API QueryPerformanceCounter
    MsgBox "Performance Counter not available", vbExclamation
  
  Else
    
    fAvailable = True
    
    ' determine API overhead
    QueryPerformanceCounter uStart
    For i = 1 To overheadLoopCount
      QueryPerformanceCounter uEnd
    Next
    dOverhead = (LI2Dbl(uEnd) - LI2Dbl(uStart)) / overheadLoopCount
    ' 20040614: AMD Athlon XP 2000+
    ' frequency: 3579545          overhead: ca. 2,92 ticks
    ''Debug.Print "frequency:"; LI2Dbl(uFreq), "overhead:"; dOverhead; "ticks"
    
  End If
  
End Sub

Friend Sub Reset()
  dPauseDuration = 0
  QueryPerformanceCounter uStart
End Sub

Friend Function Elapsed() As Double
' return elapsed time in milliseconds
  QueryPerformanceCounter uEnd
  If fAvailable Then
    Elapsed = 1000 * (LI2Dbl(uEnd) - LI2Dbl(uStart) - dOverhead - dPauseDuration) / LI2Dbl(uFreq)
  End If
End Function

Friend Function sElapsed() As String
' returns a nicely formatted string
  sElapsed = Format$(Elapsed, "#,0.000") & " msec"
End Function

Friend Sub PauseStart()
' begin pause
  QueryPerformanceCounter uPauseStart
End Sub
Friend Sub PauseEnd()
' end pause: pause duration will be subtracted from elapsed time
  QueryPerformanceCounter uEnd
  ' add 2 * dOverhead: the API calls are part of the pause
  dPauseDuration = dPauseDuration + (LI2Dbl(uEnd) - LI2Dbl(uPauseStart)) + 2 * dOverhead
End Sub

Friend Sub Wait(dMsec As Double, Optional fDoEvents As Boolean)
' returns after dMsec milliseconds
' fDoEvents = False:  total suspend, all CPU blocked
  Reset
  Do
    If fDoEvents Then
      DoEvents
    End If
  Loop While fAvailable And Elapsed < dMsec
End Sub

' 20050826
Private Function LI2Dbl(uLi As LARGE_INTEGER) As Double
  Const OFFSET_4 As Double = 4294967296#  ' 2^32
  Dim Low As Double, High As Double

  If uLi.Lo < 0 Then
    Low = uLi.Lo + OFFSET_4
  Else
    Low = uLi.Lo
  End If

  If uLi.Hi < 0 Then
    High = uLi.Hi + OFFSET_4
  Else
    High = uLi.Hi
  End If

  LI2Dbl = Low + High * OFFSET_4

End Function

