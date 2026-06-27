Attribute VB_Name = "modCollect"
'====================================================================
' modCollect - pattern collection (native FormulaR1C1 grouping),
' classification, and batch-and-subtract per-pattern measurement.
'====================================================================
Option Explicit

Public Const SCRATCH As String = "_perf_scratch"

' --- classification -------------------------------------------------
Private Function HasFn(ByVal f As String, ByVal fnName As String) As Boolean
    HasFn = (InStr(1, f, fnName & "(", vbTextCompare) > 0)
End Function

Public Function IsVolatileF(ByVal f As String) As Boolean
    IsVolatileF = HasFn(f, "OFFSET") Or HasFn(f, "INDIRECT") Or HasFn(f, "NOW") _
        Or HasFn(f, "TODAY") Or HasFn(f, "RAND") Or HasFn(f, "RANDBETWEEN") _
        Or HasFn(f, "CELL") Or HasFn(f, "INFO")
End Function

Public Function IsSingleThreadedF(ByVal f As String) As Boolean
    IsSingleThreadedF = HasFn(f, "INDIRECT") Or HasFn(f, "GETPIVOTDATA") _
        Or HasFn(f, "CELL") Or HasFn(f, "INFO")
End Function

Public Function ClassifyF(ByVal f As String) As String
    If IsVolatileF(f) Then ClassifyF = "volatile": Exit Function
    If HasFn(f, "SUMPRODUCT") Or HasFn(f, "MMULT") Or HasFn(f, "DSUM") _
        Or HasFn(f, "TRANSPOSE") Then ClassifyF = "array": Exit Function
    If HasFn(f, "LAMBDA") Or HasFn(f, "REDUCE") Or HasFn(f, "MAP") _
        Or HasFn(f, "SCAN") Then ClassifyF = "lambda": Exit Function
    If HasFn(f, "FILTER") Or HasFn(f, "SORT") Or HasFn(f, "UNIQUE") _
        Or HasFn(f, "SEQUENCE") Then ClassifyF = "dynamic-array": Exit Function
    If HasFn(f, "VLOOKUP") Or HasFn(f, "HLOOKUP") Or HasFn(f, "XLOOKUP") _
        Or HasFn(f, "MATCH") Or HasFn(f, "XMATCH") Or HasFn(f, "INDEX") _
        Then ClassifyF = "lookup": Exit Function
    ClassifyF = "cheap"
End Function

Public Function HasFullColumn(ByVal f As String) As Boolean
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}"
        re.Global = False
    End If
    HasFullColumn = re.Test(f)
End Function

' --- collection -----------------------------------------------------
' Returns a Scripting.Dictionary keyed "Sheet|FormulaR1C1".
' Each item is an array: Array(count, sheet, r1c1, sampleAddr, sampleA1,
'                               class, isVolatile, isSingleThreaded, fullCol)
Public Function CollectPatterns(wb As Workbook, ByVal capPerSheet As Long) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet, fcells As Range, ar As Range, c As Range
    Dim key As String, r1c1 As String, scanned As Long, arr As Variant
    For Each ws In wb.Worksheets
        If ws.Name = SCRATCH Then GoTo NextWs
        Set fcells = Nothing
        On Error Resume Next
        Set fcells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If Not fcells Is Nothing Then
            scanned = 0
            For Each ar In fcells.Areas
                For Each c In ar.Cells
                    r1c1 = c.FormulaR1C1
                    key = ws.Name & "|" & r1c1
                    If dict.Exists(key) Then
                        arr = dict(key)
                        arr(0) = arr(0) + 1
                        dict(key) = arr
                    Else
                        dict(key) = Array(1, ws.Name, r1c1, c.Address, c.Formula, _
                            ClassifyF(c.Formula), IsVolatileF(c.Formula), _
                            IsSingleThreadedF(c.Formula), HasFullColumn(c.Formula))
                    End If
                    scanned = scanned + 1
                    If scanned >= capPerSheet Then Exit For
                Next c
                If scanned >= capPerSheet Then Exit For
            Next ar
        End If
NextWs:
    Next ws
    Set CollectPatterns = dict
End Function

' Batch-and-subtract: copy sampleCell's R1C1 into N scratch cells, time
' Range.Calculate, subtract the same N cells filled with "=1". Returns us/cell.
Public Function MeasureUsPerCell(scratch As Worksheet, ByVal r1c1 As String, _
        ByVal n As Long, ByVal iters As Long, ByRef stdevOut As Double) As Double
    Dim rng As Range, tPat As Double, tOver As Double
    Dim samples() As Double, i As Long, mean As Double, v As Double

    ' Clear any prior spill residue across the whole scratch sheet.
    On Error Resume Next
    scratch.UsedRange.ClearContents
    On Error GoTo 0

    Set rng = scratch.Range(scratch.Cells(1, 1), scratch.Cells(n, 1))
    rng.ClearContents
    rng.FormulaR1C1 = r1c1

    ' multiple measured passes for stdev
    ReDim samples(1 To iters - 1)
    Dim k As Long: k = 0
    For i = 1 To iters
        Dim t As Double: t = MicroTimer
        rng.Calculate
        v = MicroTimer - t
        If i > 1 Then k = k + 1: samples(k) = v
    Next i

    rng.ClearContents
    rng.FormulaR1C1 = "=1"
    tOver = TimeRangeCalc(rng, iters)
    rng.ClearContents

    For i = 1 To k: mean = mean + samples(i): Next i
    mean = mean / k
    Dim perCell As Double
    perCell = (mean - tOver) / n * 1000000#
    If perCell < 0 Then perCell = 0

    Dim sd As Double
    For i = 1 To k
        Dim pc As Double: pc = (samples(i) - tOver) / n * 1000000#
        If pc < 0 Then pc = 0
        sd = sd + (pc - perCell) * (pc - perCell)
    Next i
    stdevOut = Sqr(sd / k)
    MeasureUsPerCell = perCell
End Function
