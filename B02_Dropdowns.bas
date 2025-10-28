Attribute VB_Name = "B02_Dropdowns"
'Attribute VB_Name = "B02_Dropdowns"
Option Explicit

' =============================================================================
' B02_Dropdowns – Anwesenheit/Aufgaben-Validierungen für Monatsblätter
' - Zentrale Konfiguration (Z_Konfiguration) & Helfer (M_Basis, M_SafeApp)
' - Idempotent (.Validation.Delete), robust, batch-fähig
' - 255-Zeichen-Sicherheit: lange Listen -> Named Range im VeryHidden-"KonfigCache"
' =============================================================================

' ----------------------------- Öffentliche Einstiege --------------------------

Public Sub B02_DropdownsAlle()
    M_SafeApp.BeginFastOps True, True, True
    On Error GoTo Clean

    Dim ws As Worksheet, mon, ok&, fail&, nm$
    mon = Z_Konfiguration.CFG_MonatsNamen

    Dim srcAnw As String, srcAuf As String
    srcAnw = ResolveListFormula("Anwesenheit", Z_Konfiguration.GetAnwesenheitsCodesForValidation)
    srcAuf = ResolveListFormula("Aufgaben", Z_Konfiguration.GetAufgabenCodesForValidation)

    For Each ws In ThisWorkbook.Worksheets
        If CFG_IsMonatsblattName(ws.Name) Then
            If ErzeugeDropdownsInBlatt(ws, srcAnw, srcAuf) Then
                ok = ok + 1
            Else
                fail = fail + 1
            End If
        End If
    Next

    Debug.Print "[B02] DropdownsAlle: OK=" & ok & " / Fehlend=" & fail

Clean:
    M_SafeApp.EndFastOps
    If Err.Number <> 0 Then Debug.Print "[B02] FEHLER B02_DropdownsAlle: "; Err.Number; " - "; Err.Description
End Sub

Public Sub B02_DropdownsAktiv()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If Not CFG_IsMonatsblattName(ws.Name) Then
        MsgBox "Kein Monatsblatt aktiv (Jan–Dez).", vbExclamation
        Exit Sub
    End If

    M_SafeApp.BeginFastOps True, True, True
    On Error GoTo Clean

    Dim ok As Boolean
    ok = ErzeugeDropdownsInBlatt(ws, _
         ResolveListFormula("Anwesenheit", Z_Konfiguration.GetAnwesenheitsCodes), _
         ResolveListFormula("Aufgaben", Z_Konfiguration.GetAufgabenCodes))

Clean:
    M_SafeApp.EndFastOps
    If Err.Number <> 0 Then
        Debug.Print "[B02] FEHLER B02_DropdownsAktiv: "; Err.Number; " - "; Err.Description
        MsgBox "Fehler beim Erstellen der Dropdowns: " & Err.Description, vbExclamation
    ElseIf ok Then
        MsgBox "Dropdowns aktualisiert für '" & ws.Name & "'.", vbInformation
    End If
End Sub

Public Sub B02_DropdownsResetAlle()
    ' Entfernt alle Validierungen im Datenbereich aller Monatsblätter
    M_SafeApp.BeginFastOps True, True, True
    On Error GoTo Clean

    Dim ws As Worksheet, ok&, fail&
    For Each ws In ThisWorkbook.Worksheets
        If CFG_IsMonatsblattName(ws.Name) Then
            If EntferneValidierungen(ws) Then ok = ok + 1 Else fail = fail + 1
        End If
    Next
    Debug.Print "[B02] ResetAlle: OK=" & ok & " / Fehlend=" & fail

Clean:
    M_SafeApp.EndFastOps
    If Err.Number <> 0 Then Debug.Print "[B02] FEHLER B02_DropdownsResetAlle: "; Err.Number; " - "; Err.Description
End Sub

' ------------------------------- Kernfunktionen -------------------------------

Private Function ErzeugeDropdownsInBlatt(ByVal ws As Worksheet, _
                                         ByVal srcAnwesenheit As String, _
                                         ByVal srcAufgaben As String) As Boolean
    On Error GoTo ErrH

    Dim r0&, cPers&, cTeam&, cTag&, cTagLast&, rLast&
    r0 = Z_Konfiguration.CFG_ErsteDatenZeile
    cPers = Z_Konfiguration.CFG_Spalte_Personen
    cTeam = Z_Konfiguration.CFG_Spalte_Team
    cTag = Z_Konfiguration.CFG_ErsteTagSpalte
    cTagLast = Z_Konfiguration.CFG_LetzteTagSpalte
    rLast = M_Basis.GetLetztePersonenzeile(ws)

    If rLast < r0 Or cTagLast < cTag Then GoTo Done

    Dim r&, c&, z As Range
    ' vor der Schleife – Quellen einmal “US-ready” machen
    Dim srcAnwUS As String, srcAufUS As String
    srcAnwUS = ToUSList(srcAnwesenheit)
    srcAufUS = ToUSList(srcAufgaben)

For r = r0 To rLast
    If Len(Trim$(CStr(ws.Cells(r, cPers).Value))) > 0 _
       And Not IsNumeric(ws.Cells(r, cPers).Value) Then

        For c = cTag To cTagLast Step 2
            Set z = ws.Cells(r, c)

            ' linke Tagesspalte: Anwesenheit
            With z.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                     Operator:=xlBetween, Formula1:=srcAnwUS
                .IgnoreBlank = True
                .InCellDropdown = True
            End With

            ' rechte Tagesspalte: Aufgaben
            With z.Offset(0, 1).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                     Operator:=xlBetween, Formula1:=srcAufUS
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
        Next c
    End If
Next r

Done:
    ErzeugeDropdownsInBlatt = True
    Exit Function

ErrH:
    Debug.Print "[B02] FEHLER ErzeugeDropdownsInBlatt(" & ws.Name & "): " & Err.Number & " - " & Err.Description
    ErzeugeDropdownsInBlatt = False
End Function

Private Function EntferneValidierungen(ByVal ws As Worksheet) As Boolean
    On Error GoTo ErrH
    Dim r0&, cTag&, cTagLast&, rLast&
    r0 = Z_Konfiguration.CFG_ErsteDatenZeile
    cTag = Z_Konfiguration.CFG_ErsteTagSpalte
    cTagLast = Z_Konfiguration.CFG_LetzteTagSpalte
    rLast = M_Basis.GetLetztePersonenzeile(ws)
    If rLast < r0 Then GoTo Done

    ws.Range(ws.Cells(r0, cTag), ws.Cells(rLast, cTagLast)).Validation.Delete

Done:
    EntferneValidierungen = True
    Exit Function
ErrH:
    Debug.Print "[B02] FEHLER EntferneValidierungen(" & ws.Name & "): " & Err.Number & " - " & Err.Description
    EntferneValidierungen = False
End Function

' ------------------------ 255-Zeichen- & Cache-Helfer -------------------------

Private Function ResolveListFormula(ByVal key As String, ByVal rawList As String) As String
    If Len(rawList) <= 255 Then
        ResolveListFormula = ToUSList(rawList)   ' <- WICHTIG: Komma-Liste
    Else
        Dim nm$: nm = "valListe_" & key
        EnsureNamedList nm, rawList              ' schreibt Werte spaltenweise in KonfigCache
        ResolveListFormula = "=" & nm            ' benannter Bereich
    End If
End Function

' Schreibt lange Listen einmalig in ein VeryHidden-Blatt "KonfigCache" und legt einen Namen an.
Private Sub EnsureNamedList(ByVal nm As String, ByVal rawList As String)
    Dim ws As Worksheet, parts As Variant, i&, rng As Range, ref$

    ' VeryHidden Cache-Blatt anlegen/holen
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("KonfigCache")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "KonfigCache"
        ws.Visible = xlSheetVeryHidden
    End If

    ' Listeneinträge spaltenweise schreiben
    parts = Split(rawList, Z_Konfiguration.CFG_ListSep)
    ws.Columns(1).ClearContents
    For i = LBound(parts) To UBound(parts)
        ws.Cells(i + 1, 1).Value = Trim$(CStr(parts(i)))
    Next i

    ' Bereich festlegen und Namen (neu) anlegen
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(UBound(parts) + 1, 1))

    On Error Resume Next
    ThisWorkbook.names(nm).Delete
    On Error GoTo 0

    ' WICHTIG: RefersTo MUSS eine String-Formel mit "=" und externer Adresse sein
    ref = "=" & rng.Address(RowAbsolute:=True, ColumnAbsolute:=True, External:=True)
    ThisWorkbook.names.Add Name:=nm, RefersTo:=ref
End Sub

' Wandelt eine ggf. lokal getrennte Liste in eine Komma-Liste für Validation.Add
Private Function ToUSList(ByVal s As String) As String
    Dim Sep As String: Sep = Z_Konfiguration.CFG_ListSep
    Dim t As String: t = Trim$(s)
    If Left$(t, 1) = "=" Then
        ' Bereits ein benannter Bereich/Bezug: unverändert zurück
        ToUSList = t
        Exit Function
    End If
    If Sep <> "," Then t = Replace(t, Sep, ",")
    ' Sicherheitsnetz: auch Semikolon -> Komma ersetzen
    ToUSList = Replace(t, ";", ",")
End Function


