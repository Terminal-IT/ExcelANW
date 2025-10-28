Attribute VB_Name = "A07_BAO"
'Attribute VB_Name = "A07_BAO"
Sub EinrichtenBAO()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("BAO")
    
    ws.Activate
    
    With ws
        ' Alte Inhalte löschen
        .Cells.Clear
        
        ' Tabellenüberschriften erstellen
        .Range("A1").Value = "KW"
        .Range("B1").Value = "Beginn"
        .Range("C1").Value = "Ende"
        .Range("D1").Value = "Urlaubssperre"
        .Range("E1").Value = "EA/F Technik"
        .Range("F1").Value = "BAO DV"
        .Range("G1").Value = "BAO Funk"
        
        ' Formatierung der Überschriften
        .Range("A1:G1").Font.Bold = True
        .Range("A1:G1").Interior.Color = RGB(180, 198, 231)
        .Range("A1:G1").Font.Color = RGB(0, 0, 0)
        
        ' Beispieldaten eintragen (basierend auf Ihrem Bild)
        Dim beispielDaten As Variant
        beispielDaten = Array( _
            Array("", "01.01.2024", "01.01.2024", "Neujahr", "Neujahr", "Neujahr", "Neujahr"), _
            Array("", "01.01.2025", "01.01.2025", "Neujahr", "Neujahr", "Neujahr", "Neujahr") _
        )
        
        ' Beispieldaten in die Tabelle eintragen
        Dim i As Integer
        For i = 0 To UBound(beispielDaten)
            Dim zeile As Integer
            zeile = i + 2
            
            ' Datum eintragen und formatieren
            .Cells(zeile, 2).Value = CDate(beispielDaten(i)(1)) ' Beginn
            .Cells(zeile, 3).Value = CDate(beispielDaten(i)(2)) ' Ende
            .Cells(zeile, 2).NumberFormat = "dd.mm.yyyy"
            .Cells(zeile, 3).NumberFormat = "dd.mm.yyyy"
            
            ' Teamdaten eintragen
            .Cells(zeile, 4).Value = beispielDaten(i)(3) ' Urlaubssperre
            .Cells(zeile, 5).Value = beispielDaten(i)(4) ' EA/F Technik
            .Cells(zeile, 6).Value = beispielDaten(i)(5) ' BAO DV
            .Cells(zeile, 7).Value = beispielDaten(i)(6) ' BAO Funk
        Next i
        
        ' Als strukturierte Tabelle formatieren
        On Error Resume Next
        .ListObjects("tbl_BAO").Delete
        On Error GoTo 0
        
        Dim lastRow As Integer
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
        
        Dim tblRange As Range
        Set tblRange = .Range("A1:G" & lastRow)
        .ListObjects.Add(xlSrcRange, tblRange, , xlYes).Name = "tbl_BAO"
        
        ' JETZT die Formel für Kalenderwoche setzen (nach Tabellenerstellung!)
        .ListObjects("tbl_BAO").ListColumns("KW").DataBodyRange.FormulaLocal = "=KALENDERWOCHE([@Beginn];21)"
        
        ' Spaltenbreiten anpassen
        .Columns("A:A").ColumnWidth = 8   ' KW
        .Columns("B:B").ColumnWidth = 12  ' Beginn
        .Columns("C:C").ColumnWidth = 12  ' Ende
        .Columns("D:D").ColumnWidth = 15  ' Urlaubssperre
        .Columns("E:E").ColumnWidth = 15  ' EA/F Technik
        .Columns("F:F").ColumnWidth = 12  ' BAO DV
        .Columns("G:G").ColumnWidth = 12  ' BAO Funk
        
        ' Zentrierte Ausrichtung für KW
        .ListObjects("tbl_BAO").ListColumns("KW").DataBodyRange.HorizontalAlignment = xlCenter
        
    End With
    
    MsgBox "BAO-Tabelle mit Beispieldaten wurde erstellt!"
End Sub

' Hilfsfunktion zum Hinzufügen weiterer Teamspalten
Sub TeamspaltenHinzufuegen()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("BAO")
    
    Dim teamName As String
    teamName = InputBox("Name des neuen Teams:", "Team hinzufügen")
    
    If teamName <> "" Then
        ' Letzte Spalte der Tabelle finden
        Dim lastCol As Integer
        lastCol = ws.ListObjects("tbl_BAO").Range.Columns.Count
        
        ' Tabelle erweitern
        ws.ListObjects("tbl_BAO").Resize ws.ListObjects("tbl_BAO").Range.Resize(, lastCol + 1)
        
        ' Neue Spalte benennen
        ws.ListObjects("tbl_BAO").HeaderRowRange.Cells(1, lastCol + 1).Value = teamName
        
        ' Spaltenbreite anpassen
        ws.Columns(lastCol + 1).ColumnWidth = 15
        
        MsgBox "Team '" & teamName & "' wurde hinzugefügt!"
    End If
End Sub

' Funktion zum Sortieren der BAO-Einträge nach Datum
Sub SortiereBAONachDatum()
    ' Sortiert tbl_BAO automatisch nach Spalte B (Beginn-Datum)
    On Error GoTo ErrorHandler
    
    Dim wsBAO As Worksheet
    Set wsBAO = ThisWorkbook.Worksheets("BAO")
    
    If wsBAO Is Nothing Then
        Debug.Print "BAO-Blatt nicht gefunden!"
        Exit Sub
    End If
    
    Dim tblBAO As ListObject
    Set tblBAO = wsBAO.ListObjects("tbl_BAO")
    
    If tblBAO Is Nothing Then
        Debug.Print "tbl_BAO nicht gefunden!"
        Exit Sub
    End If
    
    ' Prüfen ob Tabelle Daten hat
    If tblBAO.ListRows.Count = 0 Then
        Debug.Print "tbl_BAO ist leer - keine Sortierung nötig"
        Exit Sub
    End If
    
    ' Sortierung nach Spalte B (Beginn) aufsteigend
    With tblBAO.Sort
        .SortFields.Clear
        .SortFields.Add key:=tblBAO.ListColumns("Beginn").Range, _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Debug.Print "tbl_BAO wurde nach Datum sortiert (" & tblBAO.ListRows.Count & " Einträge)"
    Exit Sub
    
ErrorHandler:
    Debug.Print "FEHLER beim Sortieren der tbl_BAO: " & Err.Description
End Sub

