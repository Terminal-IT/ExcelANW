Attribute VB_Name = "A05_Personen"
'Attribute VB_Name = "A05_Personen"
Sub EinrichtenPersonen()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Personen")
    
    ' Alte Tabellen löschen
    On Error Resume Next
    ws.ListObjects("tbl_Personen").Delete
    On Error GoTo 0
    
    With ws
        ' Alte Inhalte löschen
        .Cells.Clear
        
        ' Überschriften
        .Range("A1").Value = "Gruppierung"
        .Range("B1").Value = "Sortierung"
        .Range("C1").Value = "Teamname"
        .Range("D1").Value = "Vorname"
        .Range("E1").Value = "Nachname"
        .Range("F1").Value = "Kürzel"
        .Range("G1").Value = "Funktion"
        .Range("H1").Value = "Aktiv"
        .Range("I1").Value = "BAO_Team"
        .Range("J1").Value = "Vollname"
        
        ' Formatierung
        .Range("A1:J1").Font.Bold = True
        .Range("A1:J1").Interior.Color = RGB(200, 200, 200)
        
        ' KORRIGIERT: Beispieldaten-Funktion aufrufen
        Call FuelleOptimierteBeispieldaten(ws)
        
        ' Als Tabelle formatieren
        Dim lastRow As Long
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    
    ' Alte Tabellen löschen
    On Error Resume Next
    ws.ListObjects("tbl_Personen").Delete
    On Error GoTo 0

        .ListObjects.Add(xlSrcRange, .Range("A1:J" & lastRow), , xlYes).Name = "tbl_Personen"
        
        ' Spaltenbreiten
        .Columns("A:B").ColumnWidth = 8
        .Columns("C:C").ColumnWidth = 12
        .Columns("D:E").ColumnWidth = 12
        .Columns("F:F").ColumnWidth = 6
        .Columns("G:G").ColumnWidth = 15
        .Columns("H:H").ColumnWidth = 6
        .Columns("I:I").ColumnWidth = 15
        .Columns("J:J").ColumnWidth = 20
    End With
    
    MsgBox "Optimierte Personen-Tabelle wurde erstellt!", vbInformation
End Sub

Private Sub FuelleOptimierteBeispieldaten(ws As Worksheet)
    ' KORRIGIERT: Vollständige 14 Personen + MVL Bereitschaft
    Dim daten As Variant
    daten = Array( _
        Array(1, "A", "ZA", "Max", "Mustermann", "MM", "Leiter", "Ja", "EA/F TECHNIK"), _
        Array(1, "B", "ZA", "Anna", "Beispiel", "AB", "ISB", "Ja", "EA/F TECHNIK"), _
        Array(2, "A", "ZA P/K", "Peter", "Test1", "PT", "P/K", "Ja", ""), _
        Array(2, "B", "ZA P/K", "Peter", "Test2", "PT", "P/K", "Ja", ""), _
        Array(3, "A", "DV", "Peter", "Test3", "PT", "DV/MobiKom", "Ja", "BAO DV"), _
        Array(3, "B", "DV", "Peter", "Test4", "PT", "DV/MobiKom", "Ja", "BAO DV"), _
        Array(3, "C", "DV", "Peter", "Test5", "PT", "DV", "Ja", "BAO DV"), _
        Array(4, "A", "Funk", "Peter", "Test6", "PT", "Funk", "Ja", "BAO FUNK"), _
        Array(4, "B", "Funk", "Peter", "Test7", "PT", "Funk", "Ja", "BAO FUNK"), _
        Array(5, "A", "Azubi", "Peter", "Test8", "PT", "Azubi", "Ja", ""), _
        Array(6, "A", "MVL", "Tom", "Tom", "TO", "Sys", "Nein", "MVL Bereitschaft"), _
        Array(6, "B", "MVL", "Fab", "Fab", "FA", "Sys", "Nein", "MVL Bereitschaft"), _
        Array(6, "C", "MVL", "Mel", "Mel", "ME", "Sys", "Ja", "MVL Bereitschaft") _
    )
    
    Dim i As Long, j As Long
    For i = 0 To UBound(daten)
        Dim zeile As Long
        zeile = i + 2
        
        ' Daten eintragen (Spalte A-I)
        For j = 0 To UBound(daten(i))
            ws.Cells(zeile, j + 1).Value = daten(i)(j)
        Next j
        
        ' Vollname-Formel (Spalte J)
        ws.Cells(zeile, 10).Formula = "=D" & zeile & "&"" ""&E" & zeile
    Next i
End Sub

