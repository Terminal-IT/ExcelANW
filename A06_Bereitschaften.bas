Attribute VB_Name = "A06_Bereitschaften"
'Attribute VB_Name = "A06_Bereitschaften"
Sub EinrichtenBereitschaften()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Bereitschaften")
    
    ws.Activate
    
    With ws
        ' Alte Inhalte löschen
        .Cells.Clear
        
        ' Überschrift
        .Range("B1").Value = "MVL Bereitschaft"
        .Range("B1").Font.Bold = True
        .Range("B1").Font.Size = 12
        .Range("B1").Interior.Color = RGB(180, 198, 231)
        
        ' Hilfszahlen-Bereich
        .Range("B2").Value = "Hilfszahlen:"
        .Range("B2").Font.Italic = True
        
        ' Bezugsdaten aus Anleitung lesen
        Dim bezugsjahr As Integer
        bezugsjahr = ThisWorkbook.Worksheets("Anleitung").Range("C2").Value
        
        ' Hilfszahlen anzeigen
        .Range("C2").Value = DateSerial(bezugsjahr, 1, 1)
        .Range("C2").NumberFormat = "dd.mm.yyyy"
        
        .Range("D2").Value = 35
        .Range("D3").Value = DateSerial(2013, 2, 25)
        .Range("D3").NumberFormat = "dd.mm.yyyy"
        
        ' Tabellenüberschriften
        .Range("B5").Value = "KW"
        .Range("C5").Value = "Beginn"
        .Range("D5").Value = "Ende"
        
        ' Formatierung der Überschriften
        .Range("B5:D5").Font.Bold = True
        .Range("B5:D5").Interior.Color = RGB(180, 198, 231)
        
        ' Bereitschaftsdaten berechnen
        Call BerechneBereitschaftszyklen(ws, bezugsjahr)
        
        ' Als strukturierte Tabelle formatieren
        On Error Resume Next
        .ListObjects("tbl_MVL").Delete
        On Error GoTo 0
        
        Dim tblRange As Range
        Set tblRange = .Range("B5:D17")
        .ListObjects.Add(xlSrcRange, tblRange, , xlYes).Name = "tbl_MVL"
        
        ' Spaltenbreiten anpassen
        .Columns("B:B").ColumnWidth = 18
        .Columns("C:C").ColumnWidth = 18
        .Columns("D:D").ColumnWidth = 18
    
        ' Rahmen um Hilfszahlen
        '.Range("C2:D3").Borders.LineStyle = xlContinuous
        '.Range("C2:D3").Borders.Weight = xlThin
        '.Range("C2:D3").Interior.Color = RGB(240, 240, 240)
        
    End With
    
    MsgBox "MVL-Bereitschaften für " & bezugsjahr & " wurden berechnet!"
End Sub

Sub BerechneBereitschaftszyklen(ws As Worksheet, bezugsjahr As Integer)
    ' Konstanten für die Berechnung
    Dim jahresbeginn As Date
    Dim referenzdatum As Date
    Dim zykluslaenge As Integer
    Dim ersterBeginn As Date
    
    jahresbeginn = DateSerial(bezugsjahr, 1, 1)
    referenzdatum = DateSerial(2013, 2, 25)
    zykluslaenge = 35
    
    ' Erste Bereitschaft berechnen - Ihre original Formel in VBA übersetzt:
    ' =WENN(($C$2-7)<$D$3+$D$2;$D$3+$D$2;GANZZAHL(AUFRUNDEN((($C$2-7)-$D$3)/$D$2;0)*$D$2+$D$3))
    If (jahresbeginn - 7) < (referenzdatum + zykluslaenge) Then
        ersterBeginn = referenzdatum + zykluslaenge
    Else
        Dim anzahlZyklen As Long
        anzahlZyklen = Int(Application.RoundUp(((jahresbeginn - 7) - referenzdatum) / zykluslaenge, 0))
        ersterBeginn = referenzdatum + (anzahlZyklen * zykluslaenge)
    End If
    
    ' Bereitschaftszyklen für 12 Einträge eintragen
    Dim zeile As Integer
    Dim aktuellerBeginn As Date
    Dim aktuellesEnde As Date
    
    aktuellerBeginn = ersterBeginn
    
    For zeile = 6 To 17
        aktuellesEnde = aktuellerBeginn + 7
        
        ' Kalenderwoche berechnen - nur wenn im Bezugsjahr
        If Year(aktuellerBeginn) = bezugsjahr Then
            ws.Cells(zeile, 2).Value = Application.WeekNum(aktuellerBeginn, 21)
        Else
            ' Wenn außerhalb des Bezugsjahres, das Datum anzeigen
            ws.Cells(zeile, 2).Value = aktuellerBeginn
            ws.Cells(zeile, 2).NumberFormat = "dd.mm.yyyy"
        End If
        
        ' Beginn eintragen
        ws.Cells(zeile, 3).Value = aktuellerBeginn
        ws.Cells(zeile, 3).NumberFormat = "dd.mm.yyyy"
        
        ' Ende eintragen
        ws.Cells(zeile, 4).Value = aktuellesEnde
        ws.Cells(zeile, 4).NumberFormat = "dd.mm.yyyy"
        
        ' Nächster Zyklus (+ 35 Tage)
        aktuellerBeginn = aktuellerBeginn + 35
    Next zeile
End Sub

' Zusätzliche Funktion zum Aktualisieren der Bereitschaften
Sub AktualisiereBereitschaften()
    Dim jahr As Integer
    jahr = ThisWorkbook.Worksheets("Anleitung").Range("C2").Value
    
    If jahr < 1900 Or jahr > 2100 Then
        MsgBox "Bitte geben Sie ein gültiges Jahr zwischen 1900 und 2100 ein!"
        Exit Sub
    End If
    
    EinrichtenBereitschaften
End Sub
