Attribute VB_Name = "C01_Formatierung"
'Attribute VB_Name = "C01_Formatierung"
' =====================================================
' OPTIMIERTE FORMATIERUNG
' Reduziert von 800+ auf 300 Zeilen Code
' =====================================================

Sub InitialisiereGrundformatierungFinal(Optional ws As Worksheet)
    ' Hauptfunktion für vollständige Formatierung
    If ws Is Nothing Then Set ws = ActiveSheet
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Formatierung in optimaler Reihenfolge
    Call SetzeGrundformatierung(ws)           ' Grundformatierung + Rahmen
    Call FormatiereKalenderHeader(ws)         ' Header-Formatierung
    Call MarkiereKalenderElemente(ws)         ' Wochenenden/Feiertage
    Call FormatiereAnwesenheitscodes(ws)      ' Anwesenheitscodes
    Call BerechneTagesstaerke(ws)            ' Tagesstärke
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If Err.Number <> 0 Then Err.Clear
End Sub

Sub SchnellAktualisierung(Target As Range)
    ' Schnelle Aktualisierung bei Änderungen
    Static isRunning As Boolean
    If isRunning Or Target.Cells.Count > 10 Then Exit Sub
    
    isRunning = True
    Application.ScreenUpdating = False
    
    On Error GoTo Cleanup
    
    ' Nur geänderte Zellen formatieren
    Dim cell As Range
    For Each cell In Target
        If cell.Column >= 4 And cell.Column <= 66 Then
            If Not IsEmpty(cell.Value) Then
                Call FormatiereSingleZelle(cell)
            Else
                Call ResetSingleZelle(cell)
            End If
        End If
    Next cell
    
    ' Tagesstärke nur bei linken Spalten aktualisieren
    Call AktualisiereTagesstaerkeSchnell(Target)
    
Cleanup:
    Application.ScreenUpdating = True
    isRunning = False
    If Err.Number <> 0 Then Err.Clear
End Sub

' ===== GRUNDFORMATIERUNG =====

Private Sub SetzeGrundformatierung(ws As Worksheet)
    ' Grundformatierung mit zentraler Konfiguration
    Dim letzteZeile As Long
    letzteZeile = M_Basis.M_Basis.GetLetztePersonenzeile(ws)
    
    With ws
        ' Gitternetzlinien ausblenden
        ActiveWindow.DisplayGridlines = False
        
        ' Fenster fixieren bei C6
        .Range("C6").Select
        ActiveWindow.FreezePanes = True
        
        ' Globale Schriftart aus Konfiguration
        .Range("A1:BO" & letzteZeile).Font.Name = GetStandardSchriftart()
        .Range("A1:BO" & letzteZeile).Font.Size = GetStandardSchriftgroesse()
        .Range("A1:BO" & letzteZeile).VerticalAlignment = xlCenter
        
        ' STANDARD-AUSRICHTUNG: Zentriert für alle
        .Range("A1:BO" & letzteZeile).HorizontalAlignment = GetAusrichtungStandard()
        
        ' SPEZIELLE AUSRICHTUNG: Spalte B rechtsausgerichtet
        .Range("B1:B" & letzteZeile).HorizontalAlignment = GetAusrichtungSpalteB()
        
        ' Spaltenbreiten aus Konfiguration
        .Columns("A:A").ColumnWidth = SpaltenbreiteA()
        .Columns("B:B").ColumnWidth = SpaltenbreiteB()
        .Columns("C:C").ColumnWidth = SpaltenbreiteC()
        .Range("D:BO").ColumnWidth = SpaltenbreiteTage()
        
        ' Formatierung anwenden
        Call SetzeAbwechselndeZeilenfarben(ws, letzteZeile)
        Call FormatiereGruppenzeilen(ws, letzteZeile)
        Call SetzeRahmenOptimiert(ws, letzteZeile)
        
    End With
End Sub

Private Sub SetzeAbwechselndeZeilenfarben(ws As Worksheet, letzteZeile As Long)
    ' Abwechselnde Zeilenfärbung mit zentraler Konfiguration
    Dim i As Long
    
    ' Header-Bereich C4:BO4 UND C5:BO5 (BEIDE Zeilen)
    ws.Range("C4:BO4").Interior.Color = FarbeZeileGerade()
    ws.Range("C5:BO5").Interior.Color = FarbeZeileUngerade()
    
    ' Datenbereich B6:BO[letzteZeile]
    For i = 6 To letzteZeile
        Dim farbe As Long
        If (i - 6) Mod 2 = 0 Then
            farbe = FarbeZeileGerade()
        Else
            farbe = FarbeZeileUngerade()
        End If
        ws.Range("B" & i & ":BO" & i).Interior.Color = farbe
    Next i
End Sub

Private Sub FormatiereGruppenzeilen(ws As Worksheet, letzteZeile As Long)
    ' Gruppenzeilen mit zentraler Konfiguration
    Dim i As Long
    For i = 6 To letzteZeile
        If IsNumeric(ws.Cells(i, 2).Value) And ws.Cells(i, 2).Value > 0 Then
            With ws.Range("B" & i & ":BO" & i)
                .Interior.Color = FarbeGruppe()
                .Font.Bold = True
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Color = RahmenFarbeSchwarz()
                .Borders(xlEdgeTop).Weight = RahmenStaerkeDuenn()
            End With
        End If
    Next i
End Sub

Private Sub SetzeRahmenOptimiert(ws As Worksheet, letzteZeile As Long)
    ' Rahmen mit zentraler Konfiguration
    On Error Resume Next
    
    With ws
        ' Vertikale Rahmen zwischen Tagen
        Dim col As Long
        For col = 3 To 66 Step 2
            .Range(.Cells(5, col), .Cells(letzteZeile, col)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(5, col), .Cells(letzteZeile, col)).Borders(xlEdgeRight).Color = RahmenFarbeGrau()
            .Range(.Cells(5, col), .Cells(letzteZeile, col)).Borders(xlEdgeRight).Weight = RahmenStaerkeDuenn()
        Next col
        
        ' Horizontale Rahmen
        .Range("B6:BO" & letzteZeile).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Range("B6:BO" & letzteZeile).Borders(xlInsideHorizontal).Color = RahmenFarbeGrau()
        .Range("B6:BO" & letzteZeile).Borders(xlInsideHorizontal).Weight = RahmenStaerkeHaar()
        
        ' Äußere Rahmen
        '.Range("B5:BO" & letzteZeile).Borders(xlEdgeTop).LineStyle = xlContinuous
        '.Range("B5:BO" & letzteZeile).Borders(xlEdgeTop).Color = RahmenFarbeSchwarz()
        '.Range("B5:BO" & letzteZeile).Borders(xlEdgeTop).Weight = RahmenStaerkeMittel()
        
        '.Range("B5:BO" & letzteZeile).Borders(xlEdgeBottom).LineStyle = xlContinuous
        '.Range("B5:BO" & letzteZeile).Borders(xlEdgeBottom).Color = RahmenFarbeSchwarz()
        '.Range("B5:BO" & letzteZeile).Borders(xlEdgeBottom).Weight = RahmenStaerkeDuenn()
        
        '.Range("B5:BO" & letzteZeile).Borders(xlEdgeLeft).LineStyle = xlContinuous
        '.Range("B5:BO" & letzteZeile).Borders(xlEdgeLeft).Color = RahmenFarbeSchwarz()
        '.Range("B5:BO" & letzteZeile).Borders(xlEdgeLeft).Weight = RahmenStaerkeDuenn()
        
        '.Range("B5:BO" & letzteZeile).Borders(xlEdgeRight).LineStyle = xlContinuous
        '.Range("B5:BO" & letzteZeile).Borders(xlEdgeRight).Color = RahmenFarbeSchwarz()
        '.Range("B5:BO" & letzteZeile).Borders(xlEdgeRight).Weight = RahmenStaerkeDuenn()
    End With
    
    On Error GoTo 0
End Sub

Private Sub MarkiereWochenendFeiertag(ws As Worksheet, spalte As Long, istHeute As Boolean, kommentarText As String)
    ' Wochenend-/Feiertagsmarkierung mit zentraler Konfiguration
    
    If istHeute Then
        ' Heute: NUR Zeile 5 markieren, Zeile 4 bleibt normal
        With ws.Cells(4, spalte)
            ' Zeile 4 bleibt unverändert (normale Wochenend-/Feiertagsfarbe falls zutreffend)
            If Weekday(ws.Cells(5, spalte).Value) = 1 Or Weekday(ws.Cells(5, spalte).Value) = 7 Then
                ' Falls heute ein Wochenende ist, normale Wochenendfarbe in Zeile 4
                .Interior.Color = FarbeWochenendeDunkel()
                .Font.Color = SchriftfarbeWeiss()
                .Font.Bold = False
            End If
            ' Sonst bleibt Zeile 4 in Standardformatierung
        End With
        
        ' NUR Zeile 5 mit Heute-Farbe markieren
        With ws.Cells(5, spalte)
            .Interior.Color = FarbeHeuteHell()  ' Nur helle Heute-Farbe in Zeile 5
            .Font.Color = SchriftfarbeWeiss()
            .Font.Bold = False
        End With
    Else
        ' Wochenende/Feiertag: Beide Zeilen wie bisher
        With ws.Cells(4, spalte)
            .Interior.Color = FarbeWochenendeDunkel()
            .Font.Color = SchriftfarbeWeiss()
            .Font.Bold = False
        End With
        
        With ws.Cells(5, spalte)
            .Interior.Color = FarbeWochenendeHell()
            .Font.Color = SchriftfarbeSchwarz()
            .Font.Bold = False
        End With
    End If
    
    ' Kommentar hinzufügen
    If kommentarText <> "" Then
        Call FuegeKommentarHinzu(ws.Cells(5, spalte), kommentarText)
    End If
End Sub

Private Sub FormatiereKalenderHeader(ws As Worksheet)
    ' Header-Formatierung mit zentraler Konfiguration
    With ws
        ' Monat und Jahr
        .Range("C4:C5").HorizontalAlignment = xlCenter
        .Range("C4:C5").VerticalAlignment = xlCenter
        .Range("C4").Font.Bold = True
        .Range("C4:C5").Font.Size = GetHeaderSchriftgroesse()
        .Range("C4").Font.Size = GetMonatSchriftgroesse()
        
        ' Kalender-Bereich
        .Range("D3:BO5").HorizontalAlignment = xlCenter
        .Range("D3:BO5").VerticalAlignment = xlCenter
        
        ' Kalenderwochen
        .Range("D3:BO3").Font.Size = 10
        .Range("D3:BO3").Font.Bold = False
        .Range("D3:BO3").Font.Color = SchriftfarbeGrau()
        .Range("D3:BO3").Interior.Color = FarbeKalenderHeader()
        
        ' Wochentage und Datum
        .Range("D4:BO4").Font.Bold = False
        .Range("D4:BO4").Font.Size = GetStandardSchriftgroesse()
        
        .Range("D5:BO5").Font.Bold = False
        .Range("D5:BO5").Font.Size = 10
    End With
End Sub

' ===== KALENDER-FORMATIERUNG =====

Private Sub MarkiereKalenderElemente(ws As Worksheet)
    ' Korrigierte Kalender-Markierung: Zeile 5 bekommt abwechselnde Farben + Ferien, Sa/So nur in Zeile 4
    Dim wsFeiertage As Worksheet, wsFerien As Worksheet
    On Error Resume Next
    Set wsFeiertage = ThisWorkbook.Worksheets("Feiertage")
    Set wsFerien = ThisWorkbook.Worksheets("Ferien")
    On Error GoTo 0
    
    ' SCHRITT 1: Erst Grundformatierung für Zeile 5 (abwechselnde Farben)
    'Call SetzeZeile5Grundformatierung(ws)
    
    ' SCHRITT 2: Dann spezielle Kalender-Markierungen
    Dim rngTag As Range, cell As Range
    Set rngTag = ws.Range("D5:BM5")
    
    For Each cell In rngTag
        If Not IsEmpty(cell.Value) And IsDate(cell.Value) Then
            Dim istWochenende As Boolean, istFeiertag As Boolean, istFerien As Boolean, istHeute As Boolean
            Dim feiertagsName As String
            
            istWochenende = (Weekday(cell.Value) = 1 Or Weekday(cell.Value) = 7)
            istFeiertag = PruefeFeiertag(cell.Value, wsFeiertage, feiertagsName)
            istHeute = (CLng(cell.Value) = CLng(Date))  ' Heute-Prüfung (ohne Uhrzeit)
            istFerien = PruefeFerien(cell.Value, wsFerien)
            
            ' GETRENNTE Behandlung für bessere Kontrolle
            If istHeute Then
            ' Heute: Nur Zeile 5 markieren
                Call MarkiereNurHeute(ws, cell.Column, feiertagsName)
            ElseIf istWochenende Or istFeiertag Then
            ' VEREINFACHT: Heute-Behandlung komplett entfernt
            If istWochenende Or istFeiertag Then
                ' Wochenende/Feiertag: NUR Zeile 4 markieren, Zeile 5 bleibt Grundfarbe
                Call MarkiereNurWochenendeFeiertag(ws, cell.Column, feiertagsName)
            ElseIf istFerien Then
                ' Ferien: Nur Zeile 5 markieren
                Call MarkiereNurFerien(ws, cell.Column)
            End If
        End If
    Next cell
End Sub

Private Sub SetzeZeile5Grundformatierung(ws As Worksheet)
    ' Setzt die abwechselnde Grundformatierung für Zeile 5 (D5:BM5)
    Dim spalte As Long
    
    ' Zeile 5 bekommt abwechselnde Farben wie Zeile 4
    For spalte = 4 To 66 Step 2  ' D, F, H, J, K, M, ... (alle 2 Spalten für Tage)
        If Not IsEmpty(ws.Cells(5, spalte).Value) Then
            ' Berechne welche "Tag-Nummer" das ist für abwechselnde Farben
            Dim tagNummer As Long
            tagNummer = (spalte - 4) / 2  ' 0, 1, 2, 3, ...
            
            Dim grundfarbe As Long
            If tagNummer Mod 2 = 0 Then
                grundfarbe = FarbeZeileGerade()
            Else
                grundfarbe = FarbeZeileUngerade()
            End If
            
            ' Verbundene Zelle (D&E, F&G, etc.) formatieren
            With ws.Range(ws.Cells(5, spalte), ws.Cells(5, spalte + 1))
                .Interior.Color = grundfarbe
                .Font.Color = SchriftfarbeSchwarz()
                .Font.Bold = False
            End With
        End If
    Next spalte
End Sub

Private Sub MarkiereNurHeute(ws As Worksheet, spalte As Long, feiertagsName As String)
    ' Heute: NUR Zeile 5 markieren, Zeile 4 normal lassen
    
    ' Zeile 4: Falls heute Wochenende/Feiertag ist, dann markieren
    If Weekday(ws.Cells(5, spalte).Value) = 1 Or Weekday(ws.Cells(5, spalte).Value) = 7 Or feiertagsName <> "" Then
        With ws.Cells(4, spalte)
            .Interior.Color = FarbeWochenendeDunkel()
            .Font.Color = SchriftfarbeWeiss()
            .Font.Bold = True
        End With
    End If
    
    ' Zeile 5: Heute-Farbe
    With ws.Range(ws.Cells(5, spalte), ws.Cells(5, spalte + 1))
        .Interior.Color = FarbeHeuteHell()
        .Font.Color = SchriftfarbeSchwarz()
        .Font.Bold = True
    End With
    
    ' Kommentar
    Dim kommentarText As String
    If feiertagsName <> "" Then
        kommentarText = "Heute: " & feiertagsName
    Else
        kommentarText = "Heute"
    End If
    
    Call FuegeKommentarHinzu(ws.Cells(5, spalte), kommentarText)
End Sub

Private Sub MarkiereNurWochenendeFeiertag(ws As Worksheet, spalte As Long, feiertagsName As String)
    ' Wochenende/Feiertag: NUR Zeile 4 markieren, Zeile 5 behält Grundfarbe
    
    ' Zeile 4: Wochenend-/Feiertagsfarbe
    With ws.Cells(4, spalte)
        .Interior.Color = FarbeWochenendeDunkel()
        .Font.Color = SchriftfarbeWeiss()
        .Font.Bold = True
    End With
    
    ' Zeile 5: BLEIBT in Grundformatierung (wurde bereits in SetzeZeile5Grundformatierung gesetzt)
    ' Nur Kommentar hinzufügen falls Feiertag
    If feiertagsName <> "" Then
        Call FuegeKommentarHinzu(ws.Cells(5, spalte), feiertagsName)
    End If
End Sub

Private Sub MarkiereNurFerien(ws As Worksheet, spalte As Long)
    ' Ferien: Nur Zeile 5 markieren
    
    ' Zeile 4: Bleibt normal
    ' Zeile 5: Ferienfarbe
    With ws.Range(ws.Cells(5, spalte), ws.Cells(5, spalte + 1))
        .Interior.Color = FarbeFerien()
        .Font.Color = SchriftfarbeSchwarz()
        .Font.Bold = False
    End With
    
    'Call FuegeKommentarHinzu(ws.Cells(5, spalte), "Ferien")
End Sub

Private Sub MarkiereTagMitKommentar(ws As Worksheet, spalte As Long, farbe As Long, weisseSchrift As Boolean, kommentarText As String)
    ' Markiert einen Tag mit Kommentar (aus Original-Code)
    
    ' Wochentag-Zelle (Zeile 4) formatieren
    With ws.Cells(4, spalte)
        .Interior.Color = farbe
        .HorizontalAlignment = xlCenter  ' ZENTRIERT
        .VerticalAlignment = xlCenter
        If weisseSchrift Then
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End If
    End With
    
    ' Tag-Zelle (Zeile 5) formatieren
    With ws.Cells(5, spalte)
        '.Interior.Color = farbe
        .HorizontalAlignment = xlCenter  ' ZENTRIERT
        .VerticalAlignment = xlCenter
        If weisseSchrift Then
            .Font.Color = RGB(0, 0, 0)
            .Font.Bold = False
        End If
        
        ' Kommentar hinzufügen (falls Text vorhanden)
        If kommentarText <> "" Then
            Call FuegeKommentarHinzu(ws.Cells(5, spalte), kommentarText)
        End If
    End With
End Sub

Private Sub FuegeKommentarHinzu(cell As Range, kommentarText As String)
    ' Fügt Kommentar hinzu (aus Original A03-Modul)
    On Error Resume Next
    
    ' Lösche vorhandenen Kommentar
    If Not cell.Comment Is Nothing Then
        cell.Comment.Delete
    End If
    
    ' Neuen Kommentar hinzufügen
    cell.AddComment kommentarText
    
    ' Kommentar formatieren
    With cell.Comment.Shape.TextFrame.Characters
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Bold = True
    End With
    
    On Error GoTo 0
End Sub

' ===== ANWESENHEITSCODE-FORMATIERUNG =====

Private Sub FormatiereAnwesenheitscodes(ws As Worksheet)
    ' Vereinfachte Anwesenheitscode-Formatierung
    Dim letzteZeile As Long
    letzteZeile = M_Basis.M_Basis.GetLetztePersonenzeile(ws)
    
    Dim i As Long, j As Long, code As String
    For i = 6 To letzteZeile
        ' Nur Personenzeilen (keine Gruppenzeilen)
        If Not IsNumeric(ws.Cells(i, 2).Value) Or ws.Cells(i, 2).Value = 0 Then
            For j = 4 To 66
                If Not IsEmpty(ws.Cells(i, j).Value) Then
                    code = UCase(Trim(CStr(ws.Cells(i, j).Value)))
                    Call FormatiereSingleCode(ws.Cells(i, j), code)
                End If
            Next j
        End If
    Next i
End Sub

Private Sub FormatiereSingleCode(cell As Range, code As String)
    ' Anwesenheitscode-Formatierung mit zentraler Konfiguration
    On Error Resume Next
    
    Select Case code
        Case "P", "S", "TA"
            cell.Interior.Color = FarbeAnwesenheit()
        Case "Z"
            cell.Interior.Color = FarbeAnwesenheitZ()
        Case "UR"
            cell.Interior.Color = FarbeUrlaub()
        Case "UV"
            cell.Interior.Color = FarbeUrlaubVorschuss()
        Case "ABW", "GL"
            cell.Interior.Color = FarbeAbwesenheit()
        Case "SU"
            cell.Interior.Color = FarbeSonderurlaub()
        Case "BE", "BE-D"
            With cell.Interior
                .Pattern = xlDown
                .PatternColor = CFG_Farbe_MVL()
            End With
        Case "BA-B", "BA-D", "BAO"
            With cell.Interior
                .Pattern = xlLightDown
                .PatternColor = FarbeBAOMuster()
            End With
    End Select
    
    cell.Font.Color = SchriftfarbeSchwarz()
    On Error GoTo 0
End Sub

Private Sub FormatiereSingleZelle(cell As Range)
    ' Formatiert eine einzelne geänderte Zelle
    If Not IsNumeric(cell.Value) And cell.Value <> "" Then
        Call FormatiereSingleCode(cell, UCase(Trim(CStr(cell.Value))))
    End If
End Sub

Private Sub ResetSingleZelle(cell As Range)
    ' Setzt eine Zelle auf Grundformatierung zurück
    Dim zeile As Long
    zeile = cell.Row
    
    With cell
        If zeile Mod 2 = 0 Then
            .Interior.Color = FarbeZeileGerade()
        Else
            .Interior.Color = FarbeZeileUngerade()
        End If
        .Interior.Pattern = xlSolid
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = False
    End With
End Sub

' ===== TAGESSTÄRKE-BERECHNUNG =====

Private Sub BerechneTagesstaerke(ws As Worksheet)
    ' Vereinfachte Tagesstärke-Berechnung
    Dim letzteZeile As Long
    letzteZeile = M_Basis.M_Basis.GetLetztePersonenzeile(ws)
    
    Dim i As Long, j As Long, teamGroesse As Long, anwesend As Long
    For i = 6 To letzteZeile
        If IsNumeric(ws.Cells(i, 2).Value) And ws.Cells(i, 2).Value > 0 Then
            teamGroesse = ws.Cells(i, 2).Value
            
            ' Nur linke Spalten (Anwesenheit)
            For j = 4 To 66 Step 2
                If Not IsEmpty(ws.Cells(5, j).Value) Then
                    anwesend = ZaehleAnwesende(ws, i, teamGroesse, j)
                    ws.Cells(i, j).Value = anwesend
                    Call FormatiereTagesstaerke(ws.Cells(i, j), anwesend, teamGroesse)
                End If
            Next j
        End If
    Next i
End Sub

Private Function ZaehleAnwesende(ws As Worksheet, gruppenZeile As Long, teamGroesse As Long, spalte As Long) As Long
    ' Zählt anwesende Personen in einer Gruppe
    Dim anwesend As Long, i As Long, code As String
    
    For i = 1 To teamGroesse
        code = UCase(Trim(CStr(ws.Cells(gruppenZeile + i, spalte).Value)))
        
        Select Case code
            Case "", "P", "S", "TA", "Z"
                anwesend = anwesend + 1
            Case Else
                If IsNumeric(ws.Cells(gruppenZeile + i, spalte).Value) And ws.Cells(gruppenZeile + i, spalte).Value > 0 Then
                    anwesend = anwesend + 1
                End If
        End Select
    Next i
    
    ZaehleAnwesende = anwesend
End Function

Private Sub FormatiereTagesstaerke(cell As Range, istWert As Long, sollWert As Long)
    ' Einfache Tagesstärke-Formatierung mit Farbverlauf
    If sollWert = 0 Then Exit Sub
    
    Dim ratio As Double
    ratio = CDbl(istWert) / CDbl(sollWert)
    If ratio > 1 Then ratio = 1
    
    ' Einfacher Farbverlauf: Rot -> Gelb -> Grün
    Dim farbe As Long
    If ratio >= 0.8 Then
        farbe = RGB(255, 255, 230)  ' Hellgelb
    ElseIf ratio >= 0.6 Then
        farbe = RGB(255, 242, 121)  ' Gelb
    ElseIf ratio >= 0.4 Then
        farbe = RGB(255, 200, 50)   ' Orange
    Else
        farbe = RGB(255, 150, 150)  ' Hellrot
    End If
    
    With cell
        .Interior.Color = farbe
        .Font.Bold = True
        If ratio < 0.5 Then
            .Font.Color = RGB(255, 255, 255)
        Else
            .Font.Color = RGB(0, 0, 0)
        End If
    End With
End Sub

Private Sub AktualisiereTagesstaerkeSchnell(Target As Range)
    ' Schnelle Tagesstärke-Aktualisierung
    If Target.Cells.Count > 5 Then Exit Sub
    
    Dim cell As Range, gruppenZeile As Long, teamGroesse As Long
    For Each cell In Target
        ' Nur bei linken Spalten (Anwesenheit)
        If cell.Column >= 4 And cell.Column <= 66 And (cell.Column - 4) Mod 2 = 0 Then
            gruppenZeile = FindeGruppenzeile(cell.Worksheet, cell.Row)
            If gruppenZeile > 0 Then
                teamGroesse = cell.Worksheet.Cells(gruppenZeile, 2).Value
                If teamGroesse > 0 Then
                    Dim anwesend As Long
                    anwesend = ZaehleAnwesende(cell.Worksheet, gruppenZeile, teamGroesse, cell.Column)
                    cell.Worksheet.Cells(gruppenZeile, cell.Column).Value = anwesend
                    Call FormatiereTagesstaerke(cell.Worksheet.Cells(gruppenZeile, cell.Column), anwesend, teamGroesse)
                End If
            End If
        End If
    Next cell
End Sub

' ===== HILFSFUNKTIONEN =====

Private Function PruefeFeiertag(datum As Date, wsFeiertage As Worksheet, ByRef feiertagsName As String) As Boolean
    ' Korrigierte Feiertag-Prüfung MIT Namen-Rückgabe
    On Error Resume Next
    feiertagsName = ""
    
    If wsFeiertage Is Nothing Then
        PruefeFeiertag = False
        Exit Function
    End If
    
    Dim feiertagTabelle As ListObject
    Set feiertagTabelle = wsFeiertage.ListObjects("tbl_Feiertage")
    
    If feiertagTabelle Is Nothing Then
        PruefeFeiertag = False
        Exit Function
    End If
    
    Dim feiertagRange As Range
    Set feiertagRange = feiertagTabelle.ListColumns("Datum").DataBodyRange
    
    ' Prüfe alle Feiertage
    Dim i As Long
    For i = 1 To feiertagRange.Rows.Count
        If feiertagRange.Cells(i, 1).Value = datum Then
            feiertagsName = feiertagRange.Cells(i, 1).Offset(0, -1).Value  ' Name aus Spalte A
            PruefeFeiertag = True
            Exit Function
        End If
    Next i
    
    PruefeFeiertag = False
    On Error GoTo 0
End Function

Private Function PruefeFerien(datum As Date, wsFerien As Worksheet) As Boolean
    ' Verbesserte Ferien-Prüfung mit Fehlerbehandlung
    On Error Resume Next
    
    PruefeFerien = False
    
    If wsFerien Is Nothing Then Exit Function
    
    Dim ferienTabelle As ListObject
    Set ferienTabelle = wsFerien.ListObjects("tbl_Ferien")
    
    If ferienTabelle Is Nothing Then Exit Function
    
    Dim ferienRange As Range
    Set ferienRange = ferienTabelle.DataBodyRange
    
    If ferienRange Is Nothing Then Exit Function
    
    Dim i As Long
    For i = 1 To ferienRange.Rows.Count
        Dim ferienBeginn As Date, ferienEnde As Date
        ferienBeginn = ferienRange.Cells(i, 2).Value  ' Spalte "Beginn"
        ferienEnde = ferienRange.Cells(i, 3).Value    ' Spalte "Ende"
        
        If datum >= ferienBeginn And datum <= ferienEnde Then
            PruefeFerien = True
            Exit Function
        End If
    Next i
    
    On Error GoTo 0
End Function

Private Function FindeGruppenzeile(ws As Worksheet, zeile As Long) As Long
    ' Findet die Gruppenzeile für eine bestimmte Zeile
    Dim i As Long
    For i = zeile To 6 Step -1
        If IsNumeric(ws.Cells(i, 2).Value) And ws.Cells(i, 2).Value > 0 Then
            FindeGruppenzeile = i
            Exit Function
        End If
    Next i
    FindeGruppenzeile = 0
End Function

' ===== PERFORMANCE-OPTIMIERUNG =====

Private Sub OptimizeCodeBegin()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Private Sub OptimizeCodeEnd()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

' ===== ÖFFENTLICHE SCHNITTSTELLEN =====

Sub ManuelleVollformatierung()
    ' Manuelle Vollformatierung für Benutzer
    Dim antwort As VbMsgBoxResult
    antwort = MsgBox("Vollständige Formatierung für aktuelles Blatt anwenden?", vbYesNo + vbQuestion)
    
    If antwort = vbYes Then
        Call InitialisiereGrundformatierungFinal(ActiveSheet)
        MsgBox "Vollständige Formatierung wurde angewendet!", vbInformation
    End If
End Sub

' ===== TEST-FUNKTIONEN FÜR FORMATIERUNG =====

Sub TestFormatierungAktuellesBlatt()
    ' Testet die komplette Formatierung für das aktuell aktive Blatt
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Prüfe ob es ein Monatsblatt ist
    If Not IstMonatsblatt(ws.Name) Then
        MsgBox "Bitte wechseln Sie zu einem Monatsblatt (Jan, Feb, etc.)", vbExclamation, "Formatierungs-Test"
        Exit Sub
    End If
    
    Dim startZeit As Double
    startZeit = Timer
    
    ' Formatierung anwenden
    Call InitialisiereGrundformatierungFinal(ws)
    
    Dim endZeit As Double
    endZeit = Timer
    
    MsgBox "Formatierung für '" & ws.Name & "' abgeschlossen!" & vbCrLf & _
           "?Dauer: " & Format((endZeit - startZeit), "0.00") & " Sekunden", _
           vbInformation, "Formatierungs-Test"
End Sub

Sub TestFormatierungAlleMonatsblaetter()
    ' Testet die komplette Formatierung für ALLE Monatsblätter
    Dim antwort As VbMsgBoxResult
    antwort = MsgBox("?FORMATIERUNG ALLER MONATSBLÄTTER TESTEN" & vbCrLf & vbCrLf & _
                     "Dies wendet die komplette Formatierung auf alle Monatsblätter an." & vbCrLf & _
                     "Vorhandene Daten bleiben erhalten!" & vbCrLf & vbCrLf & _
                     "Fortfahren?", _
                     vbYesNo + vbQuestion, "Formatierungs-Test alle Blätter")
    
    If antwort = vbNo Then Exit Sub
    
    Call OptimizeCodeBegin
    
    Dim startZeit As Double
    startZeit = Timer
    
    Dim monate As Variant
    monate = Array("Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez")
    
    Dim ws As Worksheet
    Dim verarbeiteteBlätter As Long
    Dim i As Long
    
    For i = 0 To UBound(monate)
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(monate(i))
        On Error GoTo NextBlatt
        
        If Not ws Is Nothing Then
            ' Formatierung für dieses Blatt anwenden
            Call InitialisiereGrundformatierungFinal(ws)
            verarbeiteteBlätter = verarbeiteteBlätter + 1
            Set ws = Nothing
        End If
        
NextBlatt:
        If Err.Number <> 0 Then Err.Clear
    Next i
    
    Call OptimizeCodeEnd
    
    Dim endZeit As Double
    endZeit = Timer
    
    MsgBox "FORMATIERUNG ABGESCHLOSSEN!" & vbCrLf & vbCrLf & _
           "?Verarbeitete Blätter: " & verarbeiteteBlätter & " von 12" & vbCrLf & _
           "?Gesamtdauer: " & Format((endZeit - startZeit), "0.00") & " Sekunden" & vbCrLf & _
           "Pro Blatt: " & Format((endZeit - startZeit) / verarbeiteteBlätter, "0.00") & " Sekunden", _
           vbInformation, "Formatierungs-Test Ergebnis"
End Sub

Sub TestNurKalenderFormatierung()
    ' Testet nur die Kalender-Formatierung (Wochenenden, Feiertage, heute)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    If Not IstMonatsblatt(ws.Name) Then
        MsgBox "Bitte wechseln Sie zu einem Monatsblatt (Jan, Feb, etc.)", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Nur Kalender-Elemente neu formatieren
    Call FormatiereKalenderHeader(ws)
    Call MarkiereKalenderElemente(ws)
    
    Application.ScreenUpdating = True
    
    MsgBox "?Kalender-Formatierung für '" & ws.Name & "' aktualisiert!", vbInformation
End Sub

Sub TestNurTagesstaerke()
    ' Testet nur die Tagesstärke-Berechnung und -Formatierung
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    If Not IstMonatsblatt(ws.Name) Then
        MsgBox "Bitte wechseln Sie zu einem Monatsblatt (Jan, Feb, etc.)", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Nur Tagesstärke neu berechnen und formatieren
    Call BerechneTagesstaerke(ws)
    
    Application.ScreenUpdating = True
    
    MsgBox "?Tagesstärke für '" & ws.Name & "' neu berechnet!", vbInformation
End Sub

Private Function IstMonatsblatt(blattName As String) As Boolean
    ' Hilfsfunktion: Prüft ob ein Blatt ein Monatsblatt ist
    Dim monate As Variant
    monate = Array("Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez")
    
    Dim i As Integer
    For i = 0 To UBound(monate)
        If blattName = monate(i) Then
            IstMonatsblatt = True
            Exit Function
        End If
    Next i
    
    IstMonatsblatt = False
End Function

Sub ResetteFormatierung()
    ' Setzt die Formatierung eines Blatts zurück (für Debugging)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    If Not IstMonatsblatt(ws.Name) Then
        MsgBox "Bitte wechseln Sie zu einem Monatsblatt (Jan, Feb, etc.)", vbExclamation
        Exit Sub
    End If
    
    Dim antwort As VbMsgBoxResult
    antwort = MsgBox("?FORMATIERUNG ZURÜCKSETZEN" & vbCrLf & vbCrLf & _
                     "Dies entfernt alle Formatierungen von '" & ws.Name & "'." & vbCrLf & _
                     "Daten bleiben erhalten!" & vbCrLf & vbCrLf & _
                     "Fortfahren?", _
                     vbYesNo + vbExclamation, "Formatierung zurücksetzen")
    
    If antwort = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Dim letzteZeile As Long
    letzteZeile = M_Basis.M_Basis.GetLetztePersonenzeile(ws)
    
    With ws
        ' Alle Formatierungen entfernen
        .Range("A1:BO" & letzteZeile).ClearFormats
        
        ' Basis-Schriftart wieder setzen
        .Range("A1:BO" & letzteZeile).Font.Name = GetStandardSchriftart()
        .Range("A1:BO" & letzteZeile).Font.Size = GetStandardSchriftgroesse()
        
        ' Gitternetzlinien wieder einblenden (für Test)
        ActiveWindow.DisplayGridlines = True
    End With
    
    Application.ScreenUpdating = True
    
    MsgBox "?Formatierung wurde zurückgesetzt!" & vbCrLf & _
           "Verwenden Sie 'TestFormatierungAktuellesBlatt' zum erneuten Formatieren.", _
           vbInformation
End Sub

' --- Batch-Formatierung (C01) ---
Public Sub C01_FormatierungAktuell()
    If Not IstMonatsblatt(ActiveSheet.Name) Then
        MsgBox "Bitte ein Monatsblatt aktivieren.", vbExclamation
        Exit Sub
    End If
    InitialisiereGrundformatierungFinal ActiveSheet
End Sub

Public Sub C01_FormatierungAlle()
    Dim mon As Variant, ws As Worksheet, cnt As Long
    mon = Array("Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez")
    M_SafeApp.BeginFastOps True, True, True
    On Error GoTo Done
    For Each monName In mon
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(monName)
        On Error GoTo 0
        If Not ws Is Nothing Then
            InitialisiereGrundformatierungFinal ws
            cnt = cnt + 1
            Set ws = Nothing
        End If
    Next
Done:
    M_SafeApp.EndFastOps
    Debug.Print "C01_FormatierungAlle: OK=" & cnt & " / Fehlend=" & 12 - cnt
End Sub

