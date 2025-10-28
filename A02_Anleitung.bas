Attribute VB_Name = "A02_Anleitung"
'Attribute VB_Name = "A02_Anleitung"
Public Sub EinrichtenAnleitung()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Anleitung")

    With ws
        ' Überschrift
        .Range("A1").Value = "Anwesenheitsverwaltung"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True

        ' Jahr
        .Range("B2").Value = "Jahr:"
        If Len(.Range("C2").Value) = 0 Or .Range("C2").Value = 2025 Then
            .Range("C2").Value = Year(Date)
        End If
        .Range("C2").Font.Bold = True

        ' Bundesland
        .Range("B3").Value = "Bundesland:"
        If Len(.Range("C3").Value) = 0 Then
            .Range("C3").Value = CFG_Bundesland_Default() & " – Nordrhein-Westfalen"
        End If
        
        ' MVL-Farbton (optional)
        .Range("B4").Value = "MVL-Farbton:"
        If Len(Trim(.Range("C4").Value)) = 0 Then
            .Range("C4").Value = "#B4C6E7"   ' Startwert = aktueller Standard
        End If
        ' (Optional) einfache Validierung mit Beispielen:
        With .Range("C4").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, _
                 Formula1:="#B4C6E7" & CFG_ListSep() & "#ED7D31" & CFG_ListSep() & "180,198,231" & CFG_ListSep() & "237,125,49"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With

        ' Dropdown (Validierung)
        With .Range("C3").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:=Replace(CFG_BundeslandListeCSV(), ",", CFG_ListSep())
            .IgnoreBlank = True
            .InCellDropdown = True
        End With

        ' Abschnitt "Anleitung"
        .Range("A4").Value = "Anleitung zur Bedienung"
        .Range("A4").Font.Size = 14
        .Range("A4").Font.Bold = True

        .Range("A6").Value = "1. Personen im Blatt 'Personen' pflegen"
        .Range("A7").Value = "2. Bundesland wählen & Jahr prüfen"
        .Range("A8").Value = "3. Feiertage & Ferien erstellen/aktualisieren"
        .Range("A9").Value = "4. Monatsblätter erstellen"
        .Range("A10").Value = "5. BAO/Bereitschaften integrieren"

        .Columns("A:A").ColumnWidth = 40
        .Columns("B:B").ColumnWidth = 12
        .Columns("C:C").ColumnWidth = 28
    End With
End Sub

'' Diese Funktion ins Klassen-Modul des Blattes "Anleitung" einfügen
'Private Sub Worksheet_Change(ByVal Target As Range)
'    ' Wenn das Jahr in C2 geändert wird, Feiertage automatisch aktualisieren
'    If Not Intersect(Target, Range("C2")) Is Nothing Then
'        Application.EnableEvents = False
'        Call AktualisiereFeiertage
'        Application.EnableEvents = True
'    End If
'End Sub


