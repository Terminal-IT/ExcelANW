Attribute VB_Name = "B01_Monatsblaetter"
'Attribute VB_Name = "B01_Monatsblaetter"
' =====================================================
' OPTIMIERTE MONATSBLÄTTER-ERSTELLUNG
' Reduziert von 500+ auf 200 Zeilen Code
' =====================================================

Sub ErstelleMonatsblaetter()
    ' Performance-Settings aktivieren
    Call OptimizeCodeBegin
    
    On Error GoTo ErrorHandler
    
    Dim monate As Variant
    monate = Array("Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez")
    
    Dim jahr As Long
    jahr = ThisWorkbook.Worksheets("Anleitung").Range("C2").Value
    
    ' Alte Monatsblätter löschen
    Call LoescheAlteMonatsblaetter(monate)
    
    ' Neue Monatsblätter erstellen
    Dim i As Long, ws As Worksheet
    For i = 0 To 11
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = monate(i)
        
        Call ErstelleMonatsstruktur(ws, i + 1, jahr)
        Call FuellePersonenUndDropdowns(ws)
        Call IntegriereDatenUndFormatierung(ws)
                
        ws.Range("C4").Select
    Next i
    
    ' NEU: Registerfarben setzen
    Call SetzeRegisterfarben
    
    Call OptimizeCodeEnd
    MsgBox "Alle Monatsblätter für " & jahr & " wurden erstellt!", vbInformation
    Exit Sub
    
ErrorHandler:
    Call OptimizeCodeEnd
    MsgBox "Fehler beim Erstellen der Monatsblätter: " & Err.Description, vbCritical
End Sub

Private Sub LoescheAlteMonatsblaetter(monate As Variant)
    ' Optimierte Löschung alter Monatsblätter
    Application.DisplayAlerts = False
    
    Dim ws As Worksheet, i As Long
    For Each ws In ThisWorkbook.Worksheets
        For i = 0 To UBound(monate)
            If ws.Name = monate(i) Then
                ws.Delete
                Exit For
            End If
        Next i
    Next ws
    
    Application.DisplayAlerts = True
End Sub

Private Sub ErstelleMonatsstruktur(ws As Worksheet, monat As Long, jahr As Long)
    With ws
        ' Basis-Einstellungen
        ActiveWindow.DisplayGridlines = False
        
        ' Header erstellen (unverändert)
        .Range("C4").Formula = "=$D$5"
        .Range("C4").NumberFormat = "mmmm"
        .Range("C5").Formula = "=Anleitung!$C$2"
        .Range("D5").Formula = "=DATE($C$5," & monat & ",1)"
        
        ' Tage generieren
        Dim tageImMonat As Long
        tageImMonat = Day(DateSerial(jahr, monat + 1, 0))
        
        Dim tag As Long, spalte As Long
        For tag = 1 To tageImMonat - 1
            spalte = 4 + tag * 2
            .Cells(5, spalte).Formula = "=" & .Cells(5, spalte - 2).Address & "+1"
        Next tag
        
        ' Grundformatierung
        Call SetzeGrundformatierung(ws, tageImMonat)
    End With
End Sub

Private Sub SetzeGrundformatierung(ws As Worksheet, tageImMonat As Long)
    ' Grundformatierung mit zentraler Konfiguration
    With ws
        ' Spaltenbreiten aus Konfiguration
        .Columns("A:A").ColumnWidth = SpaltenbreiteA()
        .Columns("B:B").ColumnWidth = SpaltenbreiteB()
        .Columns("C:C").ColumnWidth = SpaltenbreiteC()
        .Range("D:BM").ColumnWidth = SpaltenbreiteTage()
        
        ' Schriftart aus Konfiguration
        .Cells.Font.Name = GetStandardSchriftart()
        .Cells.Font.Size = GetStandardSchriftgroesse()
        .Cells.VerticalAlignment = xlCenter
        
        ' STANDARD-AUSRICHTUNG: Zentriert
        .Cells.HorizontalAlignment = GetAusrichtungStandard()
        
        ' SPEZIELLE AUSRICHTUNG: Spalte B rechtsausgerichtet
        .Range("B:B").HorizontalAlignment = GetAusrichtungSpalteB()
        
        ' Zellen verbinden und Formeln für Kalender
        Dim tag As Long, spalte As Long
        For tag = 0 To tageImMonat - 1
            spalte = 4 + tag * 2
            
            ' Zellen verbinden (D&E, F&G, etc.)
            .Range(.Cells(3, spalte), .Cells(3, spalte + 1)).Merge
            .Range(.Cells(4, spalte), .Cells(4, spalte + 1)).Merge
            .Range(.Cells(5, spalte), .Cells(5, spalte + 1)).Merge
            
            ' Kalenderwoche (nur bei Montag)
            .Cells(3, spalte).Formula = "=IF(WEEKDAY(" & .Cells(5, spalte).Address & ",2)=1,WEEKNUM(" & .Cells(5, spalte).Address & ",21),"""")"
            
            ' Wochentag
            .Cells(4, spalte).Formula = "=WEEKDAY(" & .Cells(5, spalte).Address & ",1)"
            .Cells(4, spalte).NumberFormat = "ddd"
            
            ' Datum formatieren
            .Cells(5, spalte).NumberFormat = "dd"
        Next tag
        
        ' Rahmen mit zentraler Konfiguration
        Call SetzeRahmen(ws, tageImMonat)
    End With
End Sub

Private Sub SetzeRahmen(ws As Worksheet, tageImMonat As Long)
    ' Rahmen mit zentraler Konfiguration
    On Error Resume Next
    
    With ws
        ' Vertikale Rahmen alle 2 Spalten
        Dim col As Long
        For col = 3 To (4 + tageImMonat * 2) Step 2
            .Range(.Cells(5, col), .Cells(70, col)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(5, col), .Cells(70, col)).Borders(xlEdgeRight).Color = RahmenFarbeGrau()
            .Range(.Cells(5, col), .Cells(70, col)).Borders(xlEdgeRight).Weight = RahmenStaerkeDuenn()
        Next col
        
        ' Oberer Rahmen für Datenbereich
        .Range("B6:BO6").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("B6:BO6").Borders(xlEdgeTop).Color = RahmenFarbeSchwarz()
        .Range("B6:BO6").Borders(xlEdgeTop).Weight = RahmenStaerkeMittel()
    End With
    
    On Error GoTo 0
End Sub

Private Sub FuellePersonenUndDropdowns(ws As Worksheet)
    ' Kombiniert: Personen eintragen + Dropdowns erstellen + BAO-Zeilen
    
    ' SCHRITT 1: Urlaubssperre-Zeile erstellen (Zeile 6)
    ws.Cells(6, PERSONEN_SPALTE).Value = ""
    ws.Cells(6, TEAM_SPALTE).Value = "Urlaubssperre"
    
    Dim wsPersonen As Worksheet
    Set wsPersonen = ThisWorkbook.Worksheets("Personen")
    
    Dim lastRow As Long
    lastRow = wsPersonen.Cells(wsPersonen.Rows.Count, "I").End(xlUp).Row
    
    ' Listen für Dropdowns
    Dim anwesenheitListe As String, aufgabenListe As String
    anwesenheitListe = GetAnwesenheitsCodes()
    aufgabenListe = GetAufgabenCodes()
    
    Dim zeile As Long, i As Long
    zeile = 7
    
    ' Teams und Personen eintragen
    Dim aktuelleGruppe As String, vorherigeGruppe As String
    Dim aktuellesBaoTeam As String  ' NEU: BAO-Team merken
    
    For i = 2 To lastRow
        aktuelleGruppe = wsPersonen.Cells(i, 1).Value
        
        ' Neue Gruppe - Team-Stärke-Zeile erstellen
        If aktuelleGruppe <> vorherigeGruppe Then
            ' BAO-Team für dieses Team merken (aus Spalte I)
            aktuellesBaoTeam = Trim(wsPersonen.Cells(i, 9).Value)
            
            ws.Cells(zeile, 2).Value = "=COUNTIFS(Personen!A:A,""" & aktuelleGruppe & """,Personen!H:H,""Ja"")"
            ws.Cells(zeile, 3).Value = wsPersonen.Cells(i, 3).Value
            zeile = zeile + 1
            vorherigeGruppe = aktuelleGruppe
        End If
        
        ' Person eintragen (nur aktive)
        If UCase(Trim(wsPersonen.Cells(i, 8).Value)) = "JA" Then
            ws.Cells(zeile, 2).Value = wsPersonen.Cells(i, 6).Value ' Kürzel
            ws.Cells(zeile, 3).Value = wsPersonen.Cells(i, 7).Value ' Zuständigkeit
            
            ' Dropdowns für diese Person erstellen
            Call ErstelleDropdownsFuerZeile(ws, zeile, anwesenheitListe, aufgabenListe)
            
            zeile = zeile + 1
        End If
        
        ' Prüfe ob nächste Zeile eine neue Gruppe ist oder Ende erreicht
        Dim naechsteGruppe As String
        If i < lastRow Then
            naechsteGruppe = wsPersonen.Cells(i + 1, 1).Value
        Else
            naechsteGruppe = "ENDE"
        End If
        
        ' BAO-Zeile hinzufügen wenn Gruppe wechselt oder Ende erreicht und BAO-Team definiert ist
        If (naechsteGruppe <> aktuelleGruppe Or naechsteGruppe = "ENDE") And aktuellesBaoTeam <> "" Then
            ws.Cells(zeile, PERSONEN_SPALTE).Value = ""
            ws.Cells(zeile, TEAM_SPALTE).Value = aktuellesBaoTeam
            
            ' BAO-Zeile formatieren
            With ws.Range("B" & zeile & ":C" & zeile)
                .Font.Italic = True
                .Interior.Color = GetBAOZeilenFormatierung()
            End With
            
            Debug.Print "BAO-Zeile hinzugefügt: " & aktuellesBaoTeam & " in Zeile " & zeile
            zeile = zeile + 1
        End If
    Next i
End Sub

Private Sub ErstelleDropdownsFuerZeile(ws As Worksheet, zeile As Long, anwesenheitListe As String, aufgabenListe As String)
    ' Vereinfachte Dropdown-Erstellung pro Zeile
    If IsNumeric(ws.Cells(zeile, 2).Value) Then Exit Sub ' Gruppenzeile überspringen
    
    Dim spalte As Long
    For spalte = 4 To 66 Step 2
        On Error Resume Next
        ' Anwesenheit (linke Spalte)
        With ws.Cells(zeile, spalte).Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:=anwesenheitListe
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        
        ' Aufgaben (rechte Spalte)
        With ws.Cells(zeile, spalte + 1).Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:=aufgabenListe
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        On Error GoTo 0
    Next spalte
End Sub

' In B01_Monatsblaetter
Private Sub IntegriereDatenUndFormatierung(ws As Worksheet)
    ' Tagesstärke-Formeln
    Call SetzeTagesstaerkeFormeln(ws)

    ' BAO + MVL aus D01 eindeutig aufrufen
    Call D01_BAOIntegration.IntegriereBAODatenKomplett(ws)
    Call D01_BAOIntegration.IntegriereMVLDaten(ws)

    ' Formatierung
    Call C01_Formatierung.InitialisiereGrundformatierungFinal(ws)
End Sub


Private Sub SetzeTagesstaerkeFormeln(ws As Worksheet)
    ' Vereinfachte Tagesstärke-Formeln
    Dim zeile As Long, spalte As Long, teamGroesse As Long
    
    For zeile = 7 To 50
        If IsNumeric(ws.Cells(zeile, 2).Value) And ws.Cells(zeile, 2).Value > 0 Then
            teamGroesse = ws.Cells(zeile, 2).Value
            
            ' Nur für Anwesenheitsspalten (linke Spalten)
            For spalte = 4 To 66 Step 2
                If Not IsEmpty(ws.Cells(5, spalte).Value) Then
                    Dim bereich As String
                    bereich = ws.Cells(zeile + 1, spalte).Address & ":" & ws.Cells(zeile + teamGroesse, spalte).Address
                    
                    ws.Cells(zeile, spalte).Formula = "=COUNTIF(" & bereich & ","""")" & _
                                                     "+COUNTIF(" & bereich & ",""TA"")" & _
                                                     "+COUNTIF(" & bereich & ",""Z"")"
                End If
            Next spalte
        End If
    Next zeile
End Sub

Sub SetzeRegisterfarben()
    ' Setzt Registerfarben für alle Monatsblätter
    On Error Resume Next
    
    Dim monate As Variant
    monate = Array("Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez")
    
    Dim aktuellerMonat As Integer
    aktuellerMonat = Month(Date)  ' Aktueller Monat (1-12)
    
    Dim i As Integer
    For i = 0 To UBound(monate)
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(monate(i))
        
        If Not ws Is Nothing Then
            If i + 1 = aktuellerMonat Then
                ' Aktueller Monat: Orange
                ws.Tab.Color = RGB(237, 125, 49)  ' #ED7D31
                Debug.Print "Aktueller Monat " & monate(i) & ": Orange"
            ElseIf i Mod 2 = 0 Then
                ' Gerade Monate (Jan, Mrz, Mai, Jul, Sep, Nov): Hellblau
                ws.Tab.Color = RGB(202, 216, 227)  ' #CAD8E3
                Debug.Print monate(i) & ": Hellblau"
            Else
                ' Ungerade Monate (Feb, Apr, Jun, Aug, Okt, Dez): Mittelblau
                ws.Tab.Color = RGB(180, 198, 231)  ' #B4C6E7
                Debug.Print monate(i) & ": Mittelblau"
            End If
        End If
    Next i
    
    Debug.Print "Registerfarben gesetzt für aktuellen Monat: " & aktuellerMonat
End Sub

' ===== PERFORMANCE-OPTIMIERUNG =====

Private Sub OptimizeCodeBegin()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
End Sub

Private Sub OptimizeCodeEnd()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

