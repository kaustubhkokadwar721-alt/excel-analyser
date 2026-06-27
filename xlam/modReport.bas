Attribute VB_Name = "modReport"
'====================================================================
' modReport - entry point, orchestration, safety envelope, 4-sheet
' report output, and the evidence-bound recommendation engine.
'
' Public entry: RunPerfDiagnostic  (Alt+F8 or a ribbon/button).
'====================================================================
Option Explicit

Private Const MAX_MEASURE As Long = 60      ' deep-time at most this many patterns
Private Const CAP_PER_SHEET As Long = 200000
Private Const NCOPIES As Long = 1000
Private Const ITERS As Long = 5
Private Const BUDGET_S As Double = 90       ' safety: stop deep-timing after this

Public PerfDiagPartial As Boolean

' Set True for headless/automated runs to suppress MsgBox dialogs.
Public PerfDiagSilent As Boolean
Public PerfDiagLastError As String
' Set True to write phase logs to %TEMP%\perfdiag_log.txt (diagnostics only).
Public PerfDiagDebug As Boolean

' Headless verification entry: runs the diagnostic without dialogs.
Public Function RunPerfDiagnosticTest(Optional ByVal wbName As String = "") As String
    PerfDiagSilent = True
    PerfDiagDebug = True          ' enable phase logging for automated verification
    PerfDiagLastError = ""
    Dim t As Workbook
    On Error Resume Next
    If wbName <> "" Then Set t = Application.Workbooks(wbName)
    If t Is Nothing Then Set t = Application.ActiveWorkbook
    On Error GoTo 0
    RunPerfDiagnosticCore t
    RunPerfDiagnosticTest = "workbooks=" & Application.Workbooks.Count & _
        IIf(PerfDiagPartial, " partial", " complete") & _
        IIf(PerfDiagLastError = "", " ok", " ERR:" & PerfDiagLastError)
End Function

Public Sub PLog(ByVal s As String)
    If Not PerfDiagDebug Then Exit Sub
    On Error Resume Next
    Dim ff As Integer: ff = FreeFile
    Open Environ$("TEMP") & "\perfdiag_log.txt" For Append As #ff
    Print #ff, Format(Now, "hh:nn:ss") & "  " & s
    Close #ff
End Sub

' Interactive entry (ribbon / Alt+F8): diagnose the active workbook.
Public Sub RunPerfDiagnostic()
    Dim wb As Workbook: Set wb = Application.ActiveWorkbook
    If wb Is Nothing Then MsgBox "No active workbook.", vbExclamation: Exit Sub
    RunPerfDiagnosticCore wb
End Sub

' Core engine (takes an explicit workbook).
Public Sub RunPerfDiagnosticCore(ByVal target As Workbook)
    Dim app As Application: Set app = Application
    PLog "START " & target.Name

    ' ---- save state (safety envelope) ----
    Dim sCalc As Long, sScreen As Boolean, sEvents As Boolean, sAlerts As Boolean
    Dim sMT As Boolean, scratchAdded As Boolean
    sCalc = app.Calculation: sScreen = app.ScreenUpdating
    sEvents = app.EnableEvents: sAlerts = app.DisplayAlerts
    On Error Resume Next
    sMT = app.MultiThreadedCalculation.Enabled
    On Error GoTo CleanFail

    app.Calculation = xlCalculationManual
    app.ScreenUpdating = False
    app.EnableEvents = False
    app.DisplayAlerts = False

    ' ---- scratch sheet in target ----
    Dim scratch As Worksheet
    On Error Resume Next
    Set scratch = target.Worksheets(modCollect.SCRATCH)
    On Error GoTo CleanFail
    If scratch Is Nothing Then
        Set scratch = target.Worksheets.Add
        scratch.Name = modCollect.SCRATCH
        scratchAdded = True
    End If

    ' ---- workbook-level timing ----
    Dim multiMs As Double, recalcMs As Double, singleMs As Double, eff As Double
    On Error Resume Next
    app.MultiThreadedCalculation.Enabled = True
    On Error GoTo CleanFail
    ' Adaptive iteration count: heavy files get fewer passes so the
    ' (uncapped) workbook-level timing cannot run away.
    multiMs = modTimer.TimeFullCalc(app, 2) * 1000      ' 1 warm-up + 1 measured
    Dim wkIters As Long
    If multiMs < 300 Then
        wkIters = ITERS
    ElseIf multiMs < 1500 Then
        wkIters = 3
    Else
        wkIters = 2
    End If
    multiMs = modTimer.TimeFullCalc(app, wkIters) * 1000
    recalcMs = modTimer.TimeRecalc(app, wkIters) * 1000
    If multiMs < 2500 Then
        On Error Resume Next
        app.MultiThreadedCalculation.Enabled = False
        singleMs = modTimer.TimeFullCalc(app, IIf(wkIters > 3, 3, 2)) * 1000
        app.MultiThreadedCalculation.Enabled = True
        On Error GoTo CleanFail
    Else
        singleMs = multiMs                              ' too slow to double-measure
    End If
    If multiMs > 0 Then eff = singleMs / multiMs Else eff = 1
    Dim volPct As Double
    If multiMs > 0 Then volPct = 100# * recalcMs / multiMs
    PLog "wkTiming done multi=" & Format(multiMs, "0") & " recalc=" & Format(recalcMs, "0")

    ' ---- collect + measure patterns ----
    Dim dict As Object: Set dict = modCollect.CollectPatterns(target, CAP_PER_SHEET)
    PLog "collected patterns=" & dict.Count

    ' rank keys by weight = count * classWeight
    Dim keys() As String, wts() As Double, nP As Long
    nP = dict.Count
    ReDim keys(1 To Application.Max(nP, 1))
    ReDim wts(1 To Application.Max(nP, 1))
    Dim k As Variant, i As Long: i = 0
    For Each k In dict.keys
        i = i + 1
        keys(i) = k
        Dim a As Variant: a = dict(k)
        wts(i) = a(0) * ClassWeight(CStr(a(5)))
        If a(6) Then wts(i) = wts(i) * 3      ' volatile
        If a(8) Then wts(i) = wts(i) * 2      ' full column
    Next k
    SortDescIdx keys, wts, nP
    PLog "ranked nP=" & nP

    ' measure top patterns (bounded by count AND a wall-clock budget)
    Dim res As Object: Set res = CreateObject("Scripting.Dictionary") ' key -> result array
    Dim measured As Long, tStart As Double: tStart = modTimer.MicroTimer
    PerfDiagPartial = False
    For i = 1 To nP
        If measured >= MAX_MEASURE Then Exit For
        If (modTimer.MicroTimer - tStart) > BUDGET_S Then PerfDiagPartial = True: Exit For
        Dim arr As Variant: arr = dict(keys(i))
        Dim occ As Long: occ = arr(0)
        Dim r1c1 As String: r1c1 = arr(2)
        Dim nCopy As Long: nCopy = NCOPIES
        If occ < nCopy Then nCopy = Application.Max(occ, 200)
        ' Spill-safety: never replicate a spilling formula into many cells
        ' (1000 copies of FILTER/SORT/structured-column = millions of cells -> crash).
        Dim isSpill As Boolean: isSpill = False
        On Error Resume Next
        isSpill = target.Worksheets(arr(1)).Range(arr(3)).HasSpill
        On Error GoTo CleanFail
        If isSpill Or arr(5) = "dynamic-array" Or arr(5) = "lambda" Then nCopy = 1
        ' Full-column / array patterns are costly per cell - fewer copies suffice
        ' and avoid replicating a 10k-row scan thousands of times.
        If arr(8) Or arr(5) = "array" Then If nCopy > 100 Then nCopy = 100
        Dim sd As Double, us As Double
        PLog "measure i=" & i & " occ=" & occ & " nCopy=" & nCopy & " cls=" & arr(5) & " " & Left$(r1c1, 40)
        On Error Resume Next
        us = modCollect.MeasureUsPerCell(scratch, r1c1, nCopy, ITERS, sd)
        On Error GoTo CleanFail
        ' result: Array(sheet,r1c1,a1,class,occ,us,stdev,total_ms,vol,st,fullcol)
        res(keys(i)) = Array(arr(1), arr(2), arr(4), arr(5), occ, us, sd, _
            us * occ / 1000#, arr(6), arr(7), arr(8), arr(3))
        measured = measured + 1
    Next i

    ' ---- write report ----
    PLog "measured " & measured & " -> writing report"
    WriteReport target, dict, res, multiMs, recalcMs, singleMs, eff, volPct
    PLog "report written"

    ' ---- cleanup ----
    If scratchAdded Then
        app.DisplayAlerts = False
        scratch.Delete
    End If
    RestoreState app, sCalc, sScreen, sEvents, sAlerts, sMT

    If Not PerfDiagSilent Then _
        MsgBox "Performance diagnostic complete." & vbCrLf & _
        "Full-calc: " & Format(multiMs, "0") & " ms" & vbCrLf & _
        "Volatility: " & Format(volPct, "0.0") & "%   MT eff: " & Format(eff, "0.00") & "x" & vbCrLf & _
        "Patterns timed: " & measured, vbInformation, "PerfDiag"
    Exit Sub

CleanFail:
    Dim em As String: em = Err.Description
    PerfDiagLastError = em
    On Error Resume Next
    If scratchAdded And Not scratch Is Nothing Then app.DisplayAlerts = False: scratch.Delete
    RestoreState app, sCalc, sScreen, sEvents, sAlerts, sMT
    If Not PerfDiagSilent Then MsgBox "PerfDiag error: " & em, vbCritical
End Sub

Private Sub RestoreState(app As Application, ByVal c As Long, ByVal scr As Boolean, _
        ByVal ev As Boolean, ByVal al As Boolean, ByVal mt As Boolean)
    On Error Resume Next
    app.Calculation = c
    app.ScreenUpdating = scr
    app.EnableEvents = ev
    app.DisplayAlerts = al
    app.MultiThreadedCalculation.Enabled = mt
End Sub

Private Function ClassWeight(ByVal cls As String) As Double
    Select Case cls
        Case "volatile": ClassWeight = 8
        Case "array": ClassWeight = 6
        Case "lookup": ClassWeight = 4
        Case "dynamic-array", "lambda": ClassWeight = 3
        Case Else: ClassWeight = 1
    End Select
End Function

' descending sort of keys() by wts() (insertion sort; pattern counts are small)
Private Sub SortDescIdx(keys() As String, wts() As Double, ByVal n As Long)
    Dim i As Long, j As Long, kw As Double, ks As String
    For i = 2 To n
        kw = wts(i): ks = keys(i): j = i - 1
        Do While j >= 1
            If wts(j) >= kw Then Exit Do
            wts(j + 1) = wts(j): keys(j + 1) = keys(j): j = j - 1
        Loop
        wts(j + 1) = kw: keys(j + 1) = ks
    Next i
End Sub

'====================================================================
' Report output
'====================================================================
Private Sub WriteReport(target As Workbook, dict As Object, res As Object, _
        ByVal multiMs As Double, ByVal recalcMs As Double, ByVal singleMs As Double, _
        ByVal eff As Double, ByVal volPct As Double)
    Dim app As Application: Set app = Application
    Dim prevScreen As Boolean: prevScreen = app.ScreenUpdating
    Dim wbR As Workbook: Set wbR = app.Workbooks.Add

    ' build a sortable array of measured results
    Dim m() As Variant, nm As Long
    nm = res.Count
    ReDim m(1 To Application.Max(nm, 1))
    Dim k As Variant, i As Long: i = 0
    For Each k In res.keys
        i = i + 1: m(i) = res(k)
    Next k
    SortResByTotal m, nm

    ' ---- Sheet: Summary ----
    Dim ws As Worksheet: Set ws = wbR.Worksheets(1): ws.Name = "Summary"
    ws.Range("A1").Value = "Excel Performance Diagnostic - Summary"
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 15
    ws.Range("A2").Value = "File: " & target.FullName
    ws.Range("A3").Value = "Excel " & app.Version & "  -  " & Format(Now, "yyyy-mm-dd hh:nn")

    ws.Range("A5").Value = "Health strip"
    ws.Range("A5").Font.Bold = True
    HealthCell ws.Range("A6"), "Volatility %", Format(volPct, "0") & "%", Band(volPct, 40, 60, False)
    HealthCell ws.Range("B6"), "MT efficiency", Format(eff, "0.00") & "x", Band(eff, 2.5, 1.5, True)

    Dim r As Long: r = 8
    ws.Cells(r, 1).Value = "Metric": ws.Cells(r, 2).Value = "Value"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Font.Bold = True
    AddMetric ws, r, "Full-calc (multi-thread) ms", Format(multiMs, "0")
    AddMetric ws, r, "Recalc (volatiles+deps) ms", Format(recalcMs, "0")
    AddMetric ws, r, "Single-thread full-calc ms", Format(singleMs, "0")
    AddMetric ws, r, "Volatility %", Format(volPct, "0.0")
    AddMetric ws, r, "Multi-thread efficiency", Format(eff, "0.00") & "x"
    AddMetric ws, r, "External links", CStr(LinkCount(target))
    AddMetric ws, r, "Defined names", CStr(target.Names.Count)
    AddMetric ws, r, "Styles", CStr(target.Styles.Count)

    ' Top costs
    r = r + 2
    ws.Cells(r, 1).Value = "Top costs": ws.Cells(r, 2).Value = "Total ms": ws.Cells(r, 3).Value = "Why"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Font.Bold = True
    Dim shown As Long
    For i = 1 To nm
        If shown >= 5 Then Exit For
        r = r + 1
        ws.Cells(r, 1).Value = m(i)(0) & "!" & "  " & m(i)(3)
        ws.Cells(r, 2).Value = Round(m(i)(7), 1)
        ws.Cells(r, 3).Value = FlagReason(m(i))
        shown = shown + 1
    Next i
    ws.Columns("A").ColumnWidth = 32: ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 60

    ' ---- Sheet: Formula Cost ----
    Dim fc As Worksheet: Set fc = wbR.Worksheets.Add(After:=ws): fc.Name = "Formula Cost"
    Dim hdr As Variant
    hdr = Array("Sheet", "Pattern (R1C1)", "Class", "Occurrences", "us/occ", _
                "Stdev us", "Total ms", "Cumulative %", "Volatile?", "Single-thread?", "Flag + reason")
    For i = 0 To UBound(hdr): fc.Cells(1, i + 1).Value = hdr(i): fc.Cells(1, i + 1).Font.Bold = True: Next i
    Dim grand As Double, cum As Double
    For i = 1 To nm: grand = grand + m(i)(7): Next i
    If grand = 0 Then grand = 1
    For i = 1 To nm
        cum = cum + m(i)(7)
        Dim rw As Long: rw = i + 1
        fc.Cells(rw, 1).Value = m(i)(0)
        fc.Cells(rw, 2).Value = Left$(CStr(m(i)(1)), 120)
        fc.Cells(rw, 3).Value = m(i)(3)
        fc.Cells(rw, 4).Value = m(i)(4)
        fc.Cells(rw, 5).Value = Round(m(i)(5), 2)
        fc.Cells(rw, 6).Value = Round(m(i)(6), 2)
        fc.Cells(rw, 7).Value = Round(m(i)(7), 2)
        fc.Cells(rw, 8).Value = Round(100# * cum / grand, 1)
        fc.Cells(rw, 9).Value = IIf(m(i)(8), "Y", "")
        fc.Cells(rw, 10).Value = IIf(m(i)(9), "Y", "")
        fc.Cells(rw, 11).Value = FlagReason(m(i))
    Next i
    If nm >= 1 Then
        On Error Resume Next
        fc.Range("G2:G" & (nm + 1)).FormatConditions.AddDatabar
        fc.Range("A1:K1").AutoFilter
        On Error GoTo 0
    End If
    fc.Columns("B").ColumnWidth = 42: fc.Columns("K").ColumnWidth = 50
    On Error Resume Next
    fc.Activate: ActiveWindow.FreezePanes = False: fc.Range("A2").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    ' ---- Sheet: Structure ----
    Dim st As Worksheet: Set st = wbR.Worksheets.Add(After:=fc): st.Name = "Structure"
    st.Range("A1:D1").Value = Array("Sheet", "Used cells", "Used-range cells", "Waste %")
    st.Range("A1:D1").Font.Bold = True
    Dim wsT As Worksheet, rr As Long: rr = 2
    For Each wsT In target.Worksheets
        If wsT.Name <> modCollect.SCRATCH Then
            Dim usedCells As Double, dimCells As Double, waste As Double
            dimCells = CDbl(wsT.UsedRange.Rows.Count) * wsT.UsedRange.Columns.Count
            usedCells = CountFormulasAndConstants(wsT)
            If dimCells > 0 Then waste = (1 - usedCells / dimCells) * 100
            st.Cells(rr, 1).Value = wsT.Name
            st.Cells(rr, 2).Value = usedCells
            st.Cells(rr, 3).Value = dimCells
            st.Cells(rr, 4).Value = Round(waste, 1)
            st.Cells(rr, 4).Interior.Color = Band(waste, 50, 90, False)
            rr = rr + 1
        End If
    Next wsT
    st.Columns("A").ColumnWidth = 20

    ' ---- Sheet: Actions ----
    Dim ac As Worksheet: Set ac = wbR.Worksheets.Add(After:=st): ac.Name = "Actions"
    ac.Range("A1").Value = "Actions - ranked by measured ROI (each anchored to evidence)"
    ac.Range("A1").Font.Bold = True
    ac.Range("A2:F2").Value = Array("Anchor", "Measured cost", "Why", "Fix", "Effort", "ROI")
    ac.Range("A2:F2").Font.Bold = True
    WriteActions ac, m, nm, eff
    ac.Columns("A").ColumnWidth = 28: ac.Columns("C").ColumnWidth = 40
    ac.Columns("D").ColumnWidth = 46

    app.ScreenUpdating = prevScreen
    ws.Activate
End Sub

' total-ms descending (index 7)
Private Sub SortResByTotal(m() As Variant, ByVal n As Long)
    Dim i As Long, j As Long, key As Variant
    For i = 2 To n
        key = m(i): j = i - 1
        Do While j >= 1
            If m(j)(7) >= key(7) Then Exit Do
            m(j + 1) = m(j): j = j - 1
        Loop
        m(j + 1) = key
    Next i
End Sub

Private Function FlagReason(a As Variant) As String
    Dim s As String
    If a(8) Then
        If InStr(1, a(2), "INDIRECT", vbTextCompare) > 0 Then
            s = "INDIRECT - volatile + single-threaded + ref-resolution cost"
        Else
            s = "volatile - recalcs on every change"
        End If
    End If
    If a(10) Then s = s & IIf(s = "", "", "; ") & "full-column reference - scans whole column"
    If a(9) And InStr(1, a(2), "INDIRECT", vbTextCompare) = 0 Then _
        s = s & IIf(s = "", "", "; ") & "single-threaded - blocks other cores"
    If a(3) = "array" Then s = s & IIf(s = "", "", "; ") & "array - cost scales with cells touched"
    FlagReason = s
End Function

Private Sub WriteActions(ac As Worksheet, m() As Variant, ByVal n As Long, ByVal eff As Double)
    Dim r As Long, i As Long, fixTxt As String
    r = 3
    ' volatile high-cost
    For i = 1 To n
        If m(i)(8) And m(i)(7) > 0 Then
            If InStr(1, m(i)(2), "OFFSET", 1) > 0 Then
                fixTxt = "Replace OFFSET with INDEX (non-volatile)."
            ElseIf InStr(1, m(i)(2), "INDIRECT", 1) > 0 Then
                fixTxt = "Replace INDIRECT with CHOOSE / structured refs."
            ElseIf InStr(1, m(i)(2), "RAND", 1) > 0 Then
                fixTxt = "Freeze random values (paste-as-values)."
            Else
                fixTxt = "Remove the volatile function; recalcs on every change."
            End If
            EmitAction ac, r, m(i)(0) & "!" & m(i)(11) & " (" & m(i)(3) & ")", _
                Format(m(i)(7), "0") & " ms total", "Volatile: recalcs on every change.", _
                fixTxt, "medium", m(i)(7) * 0.7
        End If
    Next i
    ' full column
    For i = 1 To n
        If m(i)(10) And m(i)(7) > 0 Then
            EmitAction ac, r, m(i)(0) & "!" & m(i)(11) & " (" & m(i)(3) & ")", _
                Format(m(i)(7), "0") & " ms total", "Full-column ref scans whole column.", _
                "Bound the range to used data or convert to a Table.", "low", m(i)(7) * 0.7
        End If
    Next i
    ' single-threaded when eff poor
    If eff < 1.5 Then
        For i = 1 To n
            If m(i)(9) And m(i)(7) > 0 Then
                EmitAction ac, r, m(i)(0) & "!" & m(i)(11) & " (" & m(i)(3) & ")", _
                    "MT eff " & Format(eff, "0.00") & "x", "Single-threaded blocks cores.", _
                    "Replace INDIRECT/GETPIVOTDATA/CELL/INFO.", "medium", m(i)(7) * 0.5
            End If
        Next i
    End If
    If r = 3 Then ac.Cells(3, 1).Value = "No high-ROI actions - healthy on measured axes."
End Sub

Private Sub EmitAction(ac As Worksheet, ByRef r As Long, anchor As String, cost As String, _
        whyTxt As String, fixTxt As String, effort As String, ByVal gain As Double)
    Dim roi As Double, ef As Double
    ef = IIf(effort = "low", 1, IIf(effort = "high", 4, 2))
    roi = Round(gain / ef, 2)
    ac.Cells(r, 1).Value = anchor
    ac.Cells(r, 2).Value = cost
    ac.Cells(r, 3).Value = whyTxt
    ac.Cells(r, 4).Value = fixTxt
    ac.Cells(r, 5).Value = effort
    ac.Cells(r, 6).Value = roi
    r = r + 1
End Sub

' ---- helpers ----
Private Sub HealthCell(tgt As Range, label As String, valTxt As String, ByVal clr As Long)
    tgt.Value = label: tgt.Font.Bold = True
    tgt.Offset(1, 0).Value = valTxt
    tgt.Offset(1, 0).Interior.Color = clr
    tgt.Offset(1, 0).Font.Bold = True
End Sub

Private Sub AddMetric(ws As Worksheet, ByRef r As Long, k As String, v As String)
    r = r + 1: ws.Cells(r, 1).Value = k: ws.Cells(r, 2).Value = v
End Sub

Private Function Band(ByVal v As Double, ByVal warn As Double, ByVal bad As Double, _
        ByVal reverse As Boolean) As Long
    Dim GREEN As Long, AMBER As Long, RED As Long
    GREEN = RGB(198, 239, 206): AMBER = RGB(255, 235, 156): RED = RGB(255, 199, 206)
    If reverse Then
        If v <= bad Then
            Band = RED
        ElseIf v <= warn Then
            Band = AMBER
        Else
            Band = GREEN
        End If
    Else
        If v >= bad Then
            Band = RED
        ElseIf v >= warn Then
            Band = AMBER
        Else
            Band = GREEN
        End If
    End If
End Function

Private Function LinkCount(wb As Workbook) As Long
    Dim v As Variant
    On Error Resume Next
    v = wb.LinkSources(xlExcelLinks)
    If IsArray(v) Then LinkCount = UBound(v) - LBound(v) + 1
    On Error GoTo 0
End Function

Private Function CountFormulasAndConstants(ws As Worksheet) As Double
    Dim n As Double, rng As Range
    On Error Resume Next
    Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    If Not rng Is Nothing Then n = n + rng.Cells.Count
    Set rng = Nothing
    Set rng = ws.UsedRange.SpecialCells(xlCellTypeConstants)
    If Not rng Is Nothing Then n = n + rng.Cells.Count
    On Error GoTo 0
    CountFormulasAndConstants = n
End Function
