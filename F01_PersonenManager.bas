Attribute VB_Name = "F01_PersonenManager"
'Attribute VB_Name = "F01_PersonenManager"
Option Explicit

' ------------------------------------------ Öffentliche Einstiege ------------

Public Sub F01_PersonenAktualisierenAlle()
    Dim ws As Worksheet, ok As Long, fail As Long
    BeginFastOpsLocal
    Dim personenDaten As Variant: personenDaten = LesePersonenDatenLocal()
    If IsEmpty(personenDaten) Then EndFastOpsLocal: Exit Sub

    For Each ws In ThisWorkbook.Worksheets
        If Z_Konfiguration.CFG_IsMonatsblattName(ws.Name) Then
            On Error Resume Next
            AktualisiereMonatsblattPersonenSchnell ws, personenDaten
            If Err.Number = 0 Then ok = ok + 1 Else fail = fail + 1
            Err.Clear
            On Error GoTo 0
        End If
    Next

    EndFastOpsLocal
    Debug.Print "[F01] PersonenAktualisierenAlle: OK=" & ok & " / Fehlend=" & fail
End Sub

Public Sub F01_PersonenAktualisierenAktiv()
    If Not Z_Konfiguration.CFG_IsMonatsblattName(ActiveSheet.Name) Then
        MsgBox "Kein Monatsblatt aktiv (Jan–Dez).", vbExclamation
        Exit Sub
    End If
    BeginFastOpsLocal
    Dim personenDaten As Variant: personenDaten = LesePersonenDatenLocal()
    If Not IsEmpty(personenDaten) Then
        AktualisiereMonatsblattPersonenSchnell ActiveSheet, personenDaten
    End If
    EndFastOpsLocal
    MsgBox "Personen-Struktur aktualisiert für '" & ActiveSheet.Name & "'.", vbInformation
End Sub

' ------------------------------------------ Kernlogik ------------------------

Private Sub AktualisiereMonatsblattPersonenSchnell(ByVal ws As Worksheet, ByVal personenDaten As Variant)
    Dim backup As Variant
    backup = SichereVorhandeneEintragungenLocal(ws)

    LoeschePersonenBereichLocal ws
    ErstelleNeuePersonenStrukturLocal ws, personenDaten
    RestoreVorhandeneEintragungenLocal ws, backup

    On Error Resume Next
    B03_Teamstaerke.SetzeTeamStaerkeFormeln ws
    B02_Dropdowns.B02_DropdownsAktiv ' oder ErzeugeDropDownsInBlatt indirekt
    On Error GoTo 0
End Sub

' ------------------------------------------ Daten I/O ------------------------

Private Function LesePersonenDatenLocal() As Variant
    On Error GoTo EH
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Worksheets(Z_Konfiguration.CFG_Sheet_Personen)
    Dim lo As ListObject: Set lo = wsP.ListObjects(Z_Konfiguration.CFG_Table_Personen)
    If lo Is Nothing Or lo.ListRows.Count = 0 Then Exit Function
    LesePersonenDatenLocal = lo.DataBodyRange.Value
    Exit Function
EH:
    Debug.Print "[F01] FEHLER LesePersonenDatenLocal: " & Err.Number & " - " & Err.Description
End Function

Private Function SichereVorhandeneEintragungenLocal(ByVal ws As Worksheet) As Variant
    Dim out() As String, n As Long
    ReDim out(1 To 2000, 1 To 2)

    Dim r As Long, c As Long, key As String, val As String
    Dim r0 As Long: r0 = Z_Konfiguration.CFG_ErsteDatenZeile
    Dim rLast As Long: rLast = M_Basis.GetLetztePersonenzeile(ws)
    Dim cPers As Long: cPers = Z_Konfiguration.CFG_Spalte_Personen
    Dim cTag As Long: cTag = Z_Konfiguration.CFG_ErsteTagSpalte
    Dim cTagLast As Long: cTagLast = Z_Konfiguration.CFG_LetzteTagSpalte

    For r = r0 + 1 To rLast
        Dim kuerzel As String: kuerzel = Trim$(CStr(ws.Cells(r, cPers).Value))
        If Len(kuerzel) > 0 And Not IsNumeric(ws.Cells(r, cPers).Value) Then
            For c = cTag To cTagLast
                val = Trim$(CStr(ws.Cells(r, c).Value))
                If Len(val) > 0 Then
                    n = n + 1
                    If n > UBound(out, 1) Then ReDim Preserve out(1 To n + 500, 1 To 2)
                    key = kuerzel & "|" & CStr(c)
                    out(n, 1) = key
                    out(n, 2) = val
                End If
            Next c
        End If
    Next r

    If n > 0 Then
        ReDim Preserve out(1 To n, 1 To 2)
        SichereVorhandeneEintragungenLocal = out
    End If
End Function

Private Sub RestoreVorhandeneEintragungenLocal(ByVal ws As Worksheet, ByVal backup As Variant)
    If IsEmpty(backup) Then Exit Sub
    Dim i As Long, parts As Variant, key As String, val As String, c As Long, z As Long, kz As String
    For i = 1 To UBound(backup, 1)
        key = backup(i, 1): val = backup(i, 2)
        parts = Split(key, "|")
        If UBound(parts) = 1 Then
            kz = parts(0): c = CLng(parts(1))
            z = FindePersonenZeileLocal(ws, kz)
            If z > 0 Then ws.Cells(z, c).Value = val
        End If
    Next i
End Sub

Private Function FindePersonenZeileLocal(ByVal ws As Worksheet, ByVal kuerzel As String) As Long
    Dim r As Long, r0 As Long, rLast As Long, cPers As Long
    r0 = Z_Konfiguration.CFG_ErsteDatenZeile
    rLast = M_Basis.GetLetztePersonenzeile(ws)
    cPers = Z_Konfiguration.CFG_Spalte_Personen
    For r = r0 + 1 To rLast
        If Trim$(CStr(ws.Cells(r, cPers).Value)) = kuerzel _
           And Not IsNumeric(ws.Cells(r, cPers).Value) Then
            FindePersonenZeileLocal = r: Exit Function
        End If
    Next r
End Function

Private Sub LoeschePersonenBereichLocal(ByVal ws As Worksheet)
    Dim r0 As Long: r0 = Z_Konfiguration.CFG_ErsteDatenZeile + 1
    Dim rTo As Long: rTo = Application.Max(r0, 50)
    Dim cLast As Long: cLast = Z_Konfiguration.CFG_LetzteTagSpalte
    On Error Resume Next
    ws.Range(ws.Cells(r0, Z_Konfiguration.CFG_Spalte_Personen), ws.Cells(rTo, cLast)).ClearContents
    ws.Range(ws.Cells(r0, Z_Konfiguration.CFG_Spalte_Personen), ws.Cells(rTo, cLast)).ClearFormats
    On Error GoTo 0
End Sub

Private Sub ErstelleNeuePersonenStrukturLocal(ByVal ws As Worksheet, ByVal personenDaten As Variant)
    Dim rOut As Long: rOut = Z_Konfiguration.CFG_ErsteDatenZeile + 1
    Dim i As Long
    Dim grp As String, grpPrev As String
    Dim baoTeam As String
    Dim cPers As Long: cPers = Z_Konfiguration.CFG_Spalte_Personen
    Dim cTeam As Long: cTeam = Z_Konfiguration.CFG_Spalte_Team

    For i = 1 To UBound(personenDaten, 1)
        grp = personenDaten(i, 1) ' Gruppierung
        If grp <> grpPrev Then
            baoTeam = Trim$(CStr(personenDaten(i, 9))) ' BAO-Team
            ws.Cells(rOut, cPers).Formula = "=COUNTIFS(tbl_Personen[Gruppierung],""" & grp & """,tbl_Personen[Aktiv],""Ja"")"
            ws.Cells(rOut, cTeam).Value = personenDaten(i, 3) ' Teamname
            rOut = rOut + 1
            grpPrev = grp
        End If

        If UCase$(Trim$(CStr(personenDaten(i, 8)))) = "JA" Then
            ws.Cells(rOut, cPers).Value = personenDaten(i, 6) ' Kürzel
            ws.Cells(rOut, cTeam).Value = personenDaten(i, 7) ' Funktion
            rOut = rOut + 1
        End If

        Dim nextGrp As String
        If i < UBound(personenDaten, 1) Then nextGrp = personenDaten(i + 1, 1) Else nextGrp = "#END#"
        If (nextGrp <> grp Or nextGrp = "#END#") And Len(baoTeam) > 0 Then
            ws.Cells(rOut, cPers).Value = vbNullString
            ws.Cells(rOut, cTeam).Value = baoTeam
            rOut = rOut + 1
        End If
    Next i
End Sub

' ------------------------------------------ Guards ---------------------------

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

