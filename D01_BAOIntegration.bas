Attribute VB_Name = "D01_BAOIntegration"
'Attribute VB_Name = "D01_BAOIntegration"

'=== Am Kopf des Moduls D01_BAOIntegration einfügen/prüfen ===
Option Explicit

'=== Robuste Füll-Helfer (vermeiden 438) =======================================
Private Sub SafeFillSolid(ByVal rng As Range, ByVal clr As Long)
    On Error Resume Next
    With rng.Interior
        .Pattern = xlSolid
        .Color = clr
        .TintAndShade = 0
    End With
    On Error GoTo 0
End Sub

Private Sub SafeClearIfColor(ByVal cell As Range, ByVal clr As Long)
    On Error Resume Next
    If cell.Interior.Color = clr Then
        cell.Value = vbNullString
        With cell.Interior
            .Pattern = xlSolid
            .ColorIndex = xlColorIndexNone
        End With
        ResetZelle cell
    End If
    On Error GoTo 0
End Sub

Private Function MVL_RowName() As String
    MVL_RowName = Z_Konfiguration.CFG_MVL_Zeilenname()
End Function

' -------------------------------------------------------------------
' Öffentlicher Einstieg bleibt gleich:
'   - D01_BAOIntegrationAlle
'   - D01_BAOIntegrationAktiv
'   - AktualisiereMonatsblaetterNachBAO -> ruft D01_BAOIntegrationAlle
'   - IntegriereBAODatenKomplett (unverändert)
' -------------------------------------------------------------------

'===================== MVL-Integration ===============================

Public Sub IntegriereMVLDaten(ByVal ws As Worksheet)
    On Error GoTo EH
    Debug.Print "[D01] MVL-Integration START: " & ws.Name

    Dim wsMVL As Worksheet: Set wsMVL = Nothing
    On Error Resume Next
    Set wsMVL = ThisWorkbook.Worksheets(Z_Konfiguration.CFG_Sheet_Bereitschaften)
    On Error GoTo EH
    If wsMVL Is Nothing Then
        Debug.Print "[D01] MVL: Blatt '" & Z_Konfiguration.CFG_Sheet_Bereitschaften & "' nicht gefunden."
        Exit Sub
    End If

    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsMVL.ListObjects(Z_Konfiguration.CFG_Table_MVL)
    On Error GoTo EH
    If lo Is Nothing Or lo.ListRows.Count = 0 Then
        Debug.Print "[D01] MVL: Tabelle '" & Z_Konfiguration.CFG_Table_MVL & "' fehlt/leer."
        Exit Sub
    End If

    ' 1) Zielzeile (MVL) sicherstellen
    Dim zMVL As Long
    zMVL = FindeOderErzeugeMVLZeile(ws)
    If zMVL <= 0 Then
        Debug.Print "[D01] MVL: Zielzeile nicht auffindbar."
        Exit Sub
    End If

    ' 2) alte MVL-Markierungen entfernen (Solid-Farbe)
    EntferneAlteMVLMarkierungen ws

    ' 3) Monatsbereich bestimmen
    Dim cStart As Long: cStart = Z_Konfiguration.CFG_ErsteTagSpalte
    Dim cEnd As Long:   cEnd = Z_Konfiguration.CFG_LetzteTagSpalte
    If Not IsDate(ws.Cells(5, cStart).Value) Then Exit Sub
    If Not IsDate(ws.Cells(5, cEnd).Value) Then
        cEnd = ws.Cells(5, ws.Columns.Count).End(xlToLeft).Column
        If cEnd < cStart Or Not IsDate(ws.Cells(5, cEnd).Value) Then Exit Sub
    End If
    Dim dMin As Date: dMin = ws.Cells(5, cStart).Value
    Dim dMax As Date: dMax = ws.Cells(5, cEnd).Value

    ' 4) MVL Zeiträume anwenden
    Dim r As Long, dBeg As Date, dEnd As Date
    For r = 1 To lo.ListRows.Count
        If IsDate(lo.DataBodyRange(r, 2).Value) And IsDate(lo.DataBodyRange(r, 3).Value) Then
            dBeg = lo.DataBodyRange(r, 2).Value
            dEnd = lo.DataBodyRange(r, 3).Value
            If dBeg <= 0 Or dEnd < dBeg Then GoTo NextR

            Dim dFrom As Date, dTo As Date
            dFrom = IIf(dBeg < dMin, dMin, dBeg)
            dTo = IIf(dEnd > dMax, dMax, dEnd)
            If dFrom > dTo Then GoTo NextR

            Dim c As Long, d As Date
            For c = cStart To cEnd Step 2 ' nur linke (Anwesenheits-)Spalten
                If IsDate(ws.Cells(5, c).Value) Then
                    d = ws.Cells(5, c).Value
                    If d >= dFrom And d <= dTo Then
                        ' Nur setzen, wenn noch nichts drin (BAO nicht überschreiben)
                        If Len(Trim$(CStr(ws.Cells(zMVL, c).Value))) = 0 Then
                            ' SOLID-Füllung robust auf beide Zellen (c und c+1)
                            SafeFillSolid ws.Range(ws.Cells(zMVL, c), ws.Cells(zMVL, c + 1)), Z_Konfiguration.CFG_Farbe_MVL()

                            ws.Cells(zMVL, c).Value = "BE"
                            ws.Cells(zMVL, c + 1).Value = vbNullString
                        End If
                    End If
                End If
            Next c
        End If
NextR:
    Next r

    Debug.Print "[D01] MVL-Integration ENDE: " & ws.Name
    Exit Sub
EH:
    Debug.Print "[D01] FEHLER IntegriereMVLDaten(" & ws.Name & "): " & Err.Number & " - " & Err.Description
End Sub

'— Entfernt nur MVL-Muster (nicht BAO):
Private Sub EntferneAlteMVLMarkierungen(ByVal ws As Worksheet)
    Dim r0 As Long: r0 = Z_Konfiguration.CFG_ErsteDatenZeile
    Dim rLast As Long: rLast = M_Basis.GetLetztePersonenzeile(ws)
    Dim cStart As Long: cStart = Z_Konfiguration.CFG_ErsteTagSpalte
    Dim cEnd As Long:   cEnd = Z_Konfiguration.CFG_LetzteTagSpalte
    Dim r As Long, c As Long, mvColor As Long
    mvColor = Z_Konfiguration.FarbeBereitschaftMuster

    For r = r0 To rLast
        For c = cStart To cEnd
            SafeClearIfColor ws.Cells(r, c), mvColor
        Next c
    Next r
End Sub

Private Function FindeOderErzeugeMVLZeile(ByVal ws As Worksheet) As Long
    Dim r As Long, s As String, mvName As String
    mvName = MVL_RowName()

    ' 1) Suchen (Personen-Spalte leer, Team-Spalte = MVL)
    For r = Z_Konfiguration.CFG_ErsteDatenZeile + 1 To M_Basis.GetLetztePersonenzeile(ws)
        s = Trim$(CStr(ws.Cells(r, Z_Konfiguration.CFG_Spalte_Team).Value))
        If StrComp(s, mvName, vbTextCompare) = 0 _
           And Len(Trim$(CStr(ws.Cells(r, Z_Konfiguration.CFG_Spalte_Personen).Value))) = 0 Then
            FindeOderErzeugeMVLZeile = r
            Exit Function
        End If
    Next r

    ' 2) Nicht gefunden ? unterhalb Urlaubssperre (Zeile 6 + 1) einfügen
    Dim insertRow As Long: insertRow = Z_Konfiguration.CFG_ErsteDatenZeile + 1
    ws.Rows(insertRow).Insert Shift:=xlDown
    ws.Cells(insertRow, Z_Konfiguration.CFG_Spalte_Personen).Value = vbNullString
    ws.Cells(insertRow, Z_Konfiguration.CFG_Spalte_Team).Value = mvName
    With ws.Range(ws.Cells(insertRow, 2), ws.Cells(insertRow, 3))
        .Font.Italic = True
        .Interior.Color = Z_Konfiguration.GetBAOZeilenFormatierung
    End With
    Debug.Print "[D01] MVL-Zeile automatisch eingefügt in " & ws.Name & " (Zeile " & insertRow & ")"
    FindeOderErzeugeMVLZeile = insertRow
End Function

' ============================================================================
' D01_BAOIntegration
' - Integration von BAO- und MVL-Zeiträumen in alle Monatsblätter
' - Entfernt alte Markierungen und trägt neue ein
' - Nutzt zentrale Konfiguration (Z_Konfiguration, M_Basis)
' - Enthält Admin-Wrapper für M_Admin-Kompatibilität
' ============================================================================

' -------------------- Öffentliche Einstiege ---------------------------------

Public Sub D01_BAOIntegrationAlle()
    Dim ws As Worksheet, ok As Long, fail As Long
    M_SafeApp.BeginFastOps True, True, True
    For Each ws In ThisWorkbook.Worksheets
        If Z_Konfiguration.CFG_IsMonatsblattName(ws.Name) Then
            On Error Resume Next
            IntegriereBAODatenKomplett ws
            IntegriereMVLDaten ws
            If Err.Number = 0 Then
                ok = ok + 1
            Else
                fail = fail + 1
                Debug.Print "[D01] FEHLER in Blatt " & ws.Name & ": " & Err.Description
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next
    M_SafeApp.EndFastOps
    Debug.Print "[D01] BAO+MVL Integration: OK=" & ok & " / Fehlend=" & fail
End Sub

Public Sub D01_BAOIntegrationAktiv()
    If Not Z_Konfiguration.CFG_IsMonatsblattName(ActiveSheet.Name) Then
        MsgBox "Kein Monatsblatt aktiv (Jan–Dez).", vbExclamation
        Exit Sub
    End If
    M_SafeApp.BeginFastOps True, True, True
    On Error GoTo Clean
    IntegriereBAODatenKomplett ActiveSheet
    IntegriereMVLDaten ActiveSheet
Clean:
    M_SafeApp.EndFastOps
    If Err.Number = 0 Then
        MsgBox "BAO-/MVL-Integration aktualisiert für '" & ActiveSheet.Name & "'.", vbInformation
    Else
        MsgBox "Fehler in D01_BAOIntegrationAktiv: " & Err.Description, vbCritical
    End If
End Sub

' Wrapper für Admin/Events (Kompatibilität)
Public Sub AktualisiereMonatsblaetterNachBAO()
    D01_BAOIntegrationAlle
End Sub

' -------------------- BAO-Integration ---------------------------------------

Public Sub IntegriereBAODatenKomplett(ws As Worksheet)
    On Error GoTo EH
    Debug.Print "[D01] === BAO-Integration START: " & ws.Name & " ==="
    EntferneAlteBAOMarkierungen ws
    IntegriereBAOZeitraeume ws
    Debug.Print "[D01] === BAO-Integration ENDE: " & ws.Name & " ==="
    Exit Sub
EH:
    Debug.Print "[D01] FEHLER IntegriereBAODatenKomplett(" & ws.Name & "): " & Err.Number & " - " & Err.Description
End Sub

Private Sub IntegriereBAOZeitraeume(ws As Worksheet)
    Dim wsBAO As Worksheet: Set wsBAO = Nothing
    On Error Resume Next
    Set wsBAO = ThisWorkbook.Worksheets(Z_Konfiguration.CFG_Sheet_BAO)
    On Error GoTo 0
    If wsBAO Is Nothing Then Exit Sub
    
    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsBAO.ListObjects(Z_Konfiguration.CFG_Table_BAO)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    
    Dim lr As Long: lr = lo.ListRows.Count
    Dim r As Long, beg As Date, en As Date
    For r = 1 To lr
        If IsDate(lo.DataBodyRange(r, 2).Value) And IsDate(lo.DataBodyRange(r, 3).Value) Then
            beg = lo.DataBodyRange(r, 2).Value
            en = lo.DataBodyRange(r, 3).Value
            If beg > 0 And en >= beg Then
                Dim c As Long, lastDataCol As Long
                lastDataCol = lo.ListColumns.Count
                For c = 4 To lastDataCol
                    Dim hdr As String, txt As String
                    hdr = Trim$(CStr(lo.HeaderRowRange.Cells(1, c).Value))
                    txt = Trim$(CStr(lo.DataBodyRange(r, c).Value))
                    If Len(hdr) > 0 And Len(txt) > 0 Then
                        TrageBAOZeitraumEin ws, hdr, beg, en, txt
                    End If
                Next c
            End If
        End If
    Next r
End Sub

Private Sub TrageBAOZeitraumEin(ws As Worksheet, teamSpaltenName As String, beginn As Date, ende As Date, zText As String)
    Dim z As Long
    If UCase$(teamSpaltenName) = "URLAUBSSPERRE" Then
        z = Z_Konfiguration.CFG_ErsteDatenZeile
    Else
        z = FindeBAOZeile(ws, teamSpaltenName)
    End If
    If z <= 0 Then Exit Sub
    
    Dim sp As Long, d As Date
    For sp = Z_Konfiguration.CFG_ErsteTagSpalte To Z_Konfiguration.CFG_LetzteTagSpalte Step 2
        If IsDate(ws.Cells(5, sp).Value) Then
            d = ws.Cells(5, sp).Value
            If d >= beginn And d <= ende Then
                SetzeBAOFormatierung ws, z, sp, zText
            End If
        End If
    Next sp
End Sub

Private Sub SetzeBAOFormatierung(ws As Worksheet, zeile As Long, spalte As Long, zeitraumText As String)
    Dim zTxt As String: zTxt = Trim$(ws.Cells(zeile, Z_Konfiguration.CFG_Spalte_Team).Value)
    If Not (UCase$(zTxt) Like "*BAO*" Or UCase$(zTxt) Like "*EA/F*" Or UCase$(zTxt) = "URLAUBSSPERRE" Or UCase$(zTxt) = "FUNK") Then
        Exit Sub
    End If
    
    ws.Cells(zeile, spalte).Value = zeitraumText
    ws.Cells(zeile, spalte + 1).Value = vbNullString
    
    With ws.Range(ws.Cells(zeile, spalte), ws.Cells(zeile, spalte + 1))
        SafeFillSolid .Cells, Z_Konfiguration.FarbeBAOMuster
    End With
    
    With ws.Cells(zeile, spalte)
        .HorizontalAlignment = xlLeft
        .WrapText = False
        .ShrinkToFit = False
        .IndentLevel = 0
    End With
    
    With ws.Cells(zeile, spalte + 1)
        .Value = vbNullString
        .HorizontalAlignment = xlLeft
        .WrapText = False
        .ShrinkToFit = False
        .IndentLevel = 0
    End With
End Sub

Private Function FindeBAOZeile(ws As Worksheet, teamName As String) As Long
    Dim r As Long, s As String
    For r = Z_Konfiguration.CFG_ErsteDatenZeile + 1 To M_Basis.GetLetztePersonenzeile(ws)
        s = Trim$(CStr(ws.Cells(r, Z_Konfiguration.CFG_Spalte_Team).Value))
        If UCase$(s) = UCase$(teamName) Then
            If Trim$(CStr(ws.Cells(r, Z_Konfiguration.CFG_Spalte_Personen).Value)) = "" Then
                FindeBAOZeile = r
                Exit Function
            End If
        End If
    Next r
End Function

Private Sub EntferneAlteBAOMarkierungen(ws As Worksheet)
    Dim r As Long, c As Long
    Dim r0 As Long: r0 = Z_Konfiguration.CFG_ErsteDatenZeile
    Dim rLast As Long: rLast = M_Basis.GetLetztePersonenzeile(ws)
    Dim cStart As Long: cStart = Z_Konfiguration.CFG_ErsteTagSpalte
    Dim cEnd As Long:   cEnd = Z_Konfiguration.CFG_LetzteTagSpalte
    Dim baoColor As Long: baoColor = Z_Konfiguration.FarbeBAOMuster

    For r = r0 To rLast
        For c = cStart To cEnd
            SafeClearIfColor ws.Cells(r, c), baoColor
        Next c
    Next r
End Sub

'— ResetZelle
'===================================================================

Private Sub ResetZelle(cell As Range)
    With cell
        .Interior.Pattern = xlSolid
        .TintAndShade = 0
        .Interior.ColorIndex = xlColorIndexNone
        .Font.Bold = False
        .Font.Color = Z_Konfiguration.CFG_Farbe_Text_Schwarz
        .HorizontalAlignment = xlCenter
    End With
End Sub



