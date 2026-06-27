Attribute VB_Name = "modTimer"
'====================================================================
' modTimer - high-resolution timing primitives + workbook-level calc.
' MicroTimer uses QueryPerformanceCounter (~microsecond resolution).
'====================================================================
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
#End If

' Seconds since boot, microsecond resolution.
Public Function MicroTimer() As Double
    Dim c As Currency, f As Currency
    QueryPerformanceFrequency f
    QueryPerformanceCounter c
    If f = 0 Then MicroTimer = Timer: Exit Function
    MicroTimer = c / f
End Function

' Time CalculateFull, averaging iters runs (first run discarded as warm-up).
Public Function TimeFullCalc(app As Application, ByVal iters As Long) As Double
    Dim i As Long, t As Double, s As Double, total As Double, n As Long
    For i = 1 To iters
        t = MicroTimer
        app.CalculateFull
        s = MicroTimer - t
        If i > 1 Then total = total + s: n = n + 1
    Next i
    If n = 0 Then n = 1
    TimeFullCalc = total / n
End Function

' Time Application.Calculate (recalc: volatiles + dirty + dependents).
Public Function TimeRecalc(app As Application, ByVal iters As Long) As Double
    Dim i As Long, t As Double, s As Double, total As Double, n As Long
    For i = 1 To iters
        t = MicroTimer
        app.Calculate
        s = MicroTimer - t
        If i > 1 Then total = total + s: n = n + 1
    Next i
    If n = 0 Then n = 1
    TimeRecalc = total / n
End Function

' Time Range.Calculate (single-threaded), averaging iters runs (drop warm-up).
Public Function TimeRangeCalc(rng As Range, ByVal iters As Long) As Double
    Dim i As Long, t As Double, s As Double, total As Double, n As Long
    For i = 1 To iters
        t = MicroTimer
        rng.Calculate
        s = MicroTimer - t
        If i > 1 Then total = total + s: n = n + 1
    Next i
    If n = 0 Then n = 1
    TimeRangeCalc = total / n
End Function
