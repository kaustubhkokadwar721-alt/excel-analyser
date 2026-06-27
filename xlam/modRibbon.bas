Attribute VB_Name = "modRibbon"
'====================================================================
' modRibbon - ribbon callbacks for the "Perf Diagnostic" tab.
' The ribbon XML lives in customUI14.xml, injected into the .xlam by
' inject_ribbon.py. Callbacks use As Object to avoid an Office-library
' reference dependency.
'====================================================================
Option Explicit

Public gRibbon As Object

Public Sub Ribbon_OnLoad(ribbonUI As Object)
    Set gRibbon = ribbonUI
End Sub

' Big button: run the diagnostic on the active workbook.
Public Sub Ribbon_RunDiag(control As Object)
    On Error Resume Next
    RunPerfDiagnostic
    If Err.Number <> 0 Then MsgBox "PerfDiag could not run: " & Err.Description, vbExclamation
End Sub

' Big button: about / help.
Public Sub Ribbon_About(control As Object)
    MsgBox "Excel Performance Diagnostic" & vbCrLf & vbCrLf & _
        "Measures what slows THIS workbook - per formula pattern, per sheet, " & _
        "per structure - with exact microsecond figures, then ranks the cost " & _
        "and writes a report with evidence-anchored fixes." & vbCrLf & vbCrLf & _
        "Click 'Diagnose Workbook' to produce the report.", _
        vbInformation, "PerfDiag"
End Sub
