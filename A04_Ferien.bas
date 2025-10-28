Attribute VB_Name = "A04_Ferien"
'Attribute VB_Name = "A04_Ferien"
Option Explicit

' ============================================================================
' A04_Ferien
'  - Räumt Blatt "Ferien" robust auf (erst Tabellen löschen, dann Inhalte)
'  - Erstellt Kopf: Ferienart | Beginn | Ende | Bundesland | Hinweis
'  - Liest Jahr aus Anleitung!C2, Bundesland-Code aus CFG_GetBundeslandCode()
'  - Fügt optional Beispielzeilen (NRW) für das Jahr ein – leicht entfernbar
'  - Legt ListObject "tbl_Ferien" an und sortiert nach Beginn
'  - Spaltenbreiten & Formate
' ============================================================================

Public Sub EinrichtenFerien()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Ferien")

    ' Jahr defensiv lesen
    Dim jahr As Long
    On Error Resume Next
    jahr = CLng(ThisWorkbook.Worksheets("Anleitung").Range("C2").Value)
    On Error GoTo 0
    If jahr < 1900 Or jahr > 2100 Then jahr = Year(Date)

    Dim blCode As String
    blCode = CFG_GetBundeslandCode()   ' z. B. "NW" für NRW

    ' --- Blatt robust bereinigen ---
    Call ClearSheetFerien(ws)

    With ws
        ' Kopf schreiben
        .Range("A1").Value = "Ferienart"
        .Range("B1").Value = "Beginn"
        .Range("C1").Value = "Ende"
        .Range("D1").Value = "Bundesland"
        .Range("E1").Value = "Hinweis"
        .Range("A1:E1").Font.Bold = True
        .Range("A1:E1").Interior.Color = RGB(200, 200, 200)

        ' --- Beispiel-/Startdaten (NRW) optional ---
        ' Diese Routine liefert nur einen sinnvollen Start; echte Daten bitte ersetzen.
        Call BeispielFerienEintragen(ws, jahr, blCode)

        ' Tabelle anlegen
        Dim lastRow As Long
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        If lastRow < 2 Then lastRow = 1 ' falls keine Beispielzeilen

        Dim lo As ListObject, rng As Range
        Set rng = .Range("A1:E" & lastRow)
        Set lo = .ListObjects.Add(xlSrcRange, rng, , xlYes)
        lo.Name = "tbl_Ferien"

        ' Sortierung: Beginn aufsteigend
        With lo.Sort
            .SortFields.Clear
            If lastRow >= 2 Then
                .SortFields.Add key:=lo.ListColumns("Beginn").DataBodyRange, _
                                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            End If
            .Header = xlYes
            .Apply
        End With

        ' Formate & Spalten
        .Columns("A").ColumnWidth = 22
        .Columns("B").ColumnWidth = 14
        .Columns("C").ColumnWidth = 14
        .Columns("D").ColumnWidth = 12
        .Columns("E").ColumnWidth = 24

        If lastRow >= 2 Then
            .Range("B2:C" & lastRow).NumberFormat = "dd.mm.yyyy"
        End If
    End With

    MsgBox "Ferien " & jahr & " (" & blCode & ") wurden eingerichtet.", vbInformation
End Sub

Private Sub ClearSheetFerien(ByVal ws As Worksheet)
    ' WICHTIG: Erst vorhandene Tabellen löschen, dann Inhalte.
    On Error Resume Next
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        lo.Delete
    Next lo
    On Error GoTo 0

    ws.Cells.Clear
End Sub

Private Sub BeispielFerienEintragen(ByVal ws As Worksheet, ByVal jahr As Long, ByVal blCode As String)
    ' Minimalbeispiel: NRW (NW) – bitte bei Bedarf durch echte Daten ersetzen.
    ' Für andere Bundesländer wird ein Platzhalter erzeugt.
    Dim z As Long: z = 2

    Select Case UCase$(blCode)
        Case "NW", "NW.", "NRW"
            ' Beispiel NRW (ungefähr, bitte bei Bedarf ersetzen):
            z = AddFerien(ws, z, "Osterferien", DateSerial(jahr, 4, 14), DateSerial(jahr, 4, 25), blCode, "")
            z = AddFerien(ws, z, "Sommerferien", DateSerial(jahr, 7, 28), DateSerial(jahr, 9, 9), blCode, "")
            z = AddFerien(ws, z, "Herbstferien", DateSerial(jahr, 10, 27), DateSerial(jahr, 10, 31), blCode, "")
            z = AddFerien(ws, z, "Weihnachtsferien", DateSerial(jahr, 12, 22), DateSerial(jahr + 1, 1, 3), blCode, "")
        Case Else
            ' Platzhalter für andere BL – User kann die Zeilen direkt überschreiben
            z = AddFerien(ws, z, "Ferienblock 1", DateSerial(jahr, 2, 10), DateSerial(jahr, 2, 14), blCode, "Bitte echte Daten eintragen")
            z = AddFerien(ws, z, "Ferienblock 2", DateSerial(jahr, 4, 1), DateSerial(jahr, 4, 12), blCode, "Bitte echte Daten eintragen")
            z = AddFerien(ws, z, "Ferienblock 3", DateSerial(jahr, 7, 15), DateSerial(jahr, 8, 25), blCode, "Bitte echte Daten eintragen")
            z = AddFerien(ws, z, "Ferienblock 4", DateSerial(jahr, 10, 21), DateSerial(jahr, 10, 25), blCode, "Bitte echte Daten eintragen")
            z = AddFerien(ws, z, "Ferienblock 5", DateSerial(jahr, 12, 23), DateSerial(jahr + 1, 1, 3), blCode, "Bitte echte Daten eintragen")
    End Select
End Sub

Private Function AddFerien(ByVal ws As Worksheet, ByVal z As Long, _
                           ByVal art As String, ByVal beg As Date, ByVal ende As Date, _
                           ByVal blCode As String, Optional ByVal hinweis As String = vbNullString) As Long
    ' Einfügen ohne Duplikate (Art+Beginn+Ende)
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 1

    For r = 2 To lastRow
        If ws.Cells(r, 1).Value = art _
        And CLng(ws.Cells(r, 2).Value) = CLng(beg) _
        And CLng(ws.Cells(r, 3).Value) = CLng(ende) Then
            AddFerien = z
            Exit Function
        End If
    Next r

    ws.Cells(z, 1).Value = art
    ws.Cells(z, 2).Value = beg
    ws.Cells(z, 3).Value = ende
    ws.Cells(z, 4).Value = blCode
    If Len(hinweis) > 0 Then ws.Cells(z, 5).Value = hinweis

    AddFerien = z + 1
End Function

' Komfort: Schnelltest für das aktive Jahr/BL
Public Sub Test_A04_EinrichtenFerien()
    Call EinrichtenFerien
End Sub


