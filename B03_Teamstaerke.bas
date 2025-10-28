Attribute VB_Name = "B03_Teamstaerke"
'Attribute VB_Name = "B03_Teamstaerke"
Option Explicit

' ----------------------------- Öffentliche Einstiege --------------------------

Public Sub B03_TeamstaerkeAlle()
    Dim ws As Worksheet, ok As Long, fail As Long
    BeginFastOpsLocal
    On Error GoTo Done

    For Each ws In ThisWorkbook.Worksheets
        If Z_Konfiguration.CFG_IsMonatsblattName(ws.Name) Then
            On Error Resume Next
            SetzeTeamStaerkeFormeln ws
            If Err.Number = 0 Then ok = ok + 1 Else fail = fail + 1
            Err.Clear
            On Error GoTo 0
        End If
    Next

Done:
    EndFastOpsLocal
    Debug.Print "[B03] TeamstaerkeAlle: OK=" & ok & " / Fehlend=" & fail
End Sub

Public Sub B03_TeamstaerkeAktiv()
    If Not Z_Konfiguration.CFG_IsMonatsblattName(ActiveSheet.Name) Then
        MsgBox "Kein Monatsblatt aktiv (Jan–Dez).", vbExclamation
        Exit Sub
    End If
    BeginFastOpsLocal
    SetzeTeamStaerkeFormeln ActiveSheet
    EndFastOpsLocal
    MsgBox "Teamstärke-Formeln aktualisiert für '" & ActiveSheet.Name & "'.", vbInformation
End Sub

' --------------------------------- Kernlogik ---------------------------------

Public Sub SetzeTeamStaerkeFormeln(ByVal ws As Worksheet)
    Dim r0 As Long: r0 = Z_Konfiguration.CFG_ErsteDatenZeile
    Dim rLast As Long: rLast = M_Basis.GetLetztePersonenzeile(ws)
    Dim cPers As Long: cPers = Z_Konfiguration.CFG_Spalte_Personen
    Dim cTag As Long: cTag = Z_Konfiguration.CFG_ErsteTagSpalte
    Dim cTagLast As Long: cTagLast = Z_Konfiguration.CFG_LetzteTagSpalte

    Dim r As Long
    For r = r0 To rLast
        If IsNumeric(ws.Cells(r, cPers).Value) And ws.Cells(r, cPers).Value > 0 Then
            Dim teamGroesse As Long: teamGroesse = ws.Cells(r, cPers).Value
            Dim c As Long
            For c = cTag To cTagLast Step 2
                If Not IsEmpty(ws.Cells(5, c).Value) Then
                    SetzeTeamStaerkeFormelFuerSpalte ws, r, c, teamGroesse
                End If
            Next c
        End If
    Next r
End Sub

Private Sub SetzeTeamStaerkeFormelFuerSpalte(ByVal ws As Worksheet, ByVal teamZeile As Long, _
                                             ByVal spalte As Long, ByVal teamGroesse As Long)
    Dim startZeile As Long: startZeile = teamZeile + 1
    Dim endZeile As Long:   endZeile = teamZeile + teamGroesse
    Dim bereich As String
    bereich = ws.Cells(startZeile, spalte).Address(False, False) & ":" & ws.Cells(endZeile, spalte).Address(False, False)

    Dim f As String
    f = "=COUNTIF(" & bereich & ","""")" & _
        "+COUNTIF(" & bereich & ",""TA"")" & _
        "+COUNTIF(" & bereich & ",""Z"")" & _
        "+COUNTIF(" & bereich & ",""P"")" & _
        "+COUNTIF(" & bereich & ",""S"")"
    ws.Cells(teamZeile, spalte).Formula = f
End Sub

' --------------------------------- Guards ------------------------------------

Private Sub BeginFastOpsLocal()
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Private Sub EndFastOpsLocal()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

