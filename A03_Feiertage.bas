Attribute VB_Name = "A03_Feiertage"
'Attribute VB_Name = "A03_Feiertage"
Public Sub EinrichtenFeiertage()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Feiertage")

    Dim jahr As Long
    On Error Resume Next
    jahr = CLng(ThisWorkbook.Worksheets("Anleitung").Range("C2").Value)
    On Error GoTo 0
    If jahr < 1900 Or jahr > 2100 Then jahr = Year(Date)

    Dim blCode As String
    blCode = NormalizeBL(CFG_GetBundeslandCode()) ' z. B. "NRW" -> "NW"

    With ws
        .Cells.Clear
        Dim tbl As ListObject
        For Each tbl In .ListObjects
            tbl.Delete
        Next tbl

        ' Kopf
        .Range("A1").Value = "Feiertag"
        .Range("B1").Value = "Datum"
        .Range("C1").Value = "Bundesland"
        .Range("A1:C1").Font.Bold = True
        .Range("A1:C1").Interior.Color = RGB(200, 200, 200)

        ' Eintragen
        Call EintragenFeiertage_Bundesland(ws, jahr, blCode)

        ' Tabelle
        Dim lastRow As Long
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        Dim tblRange As Range
        Set tblRange = .Range("A1:C" & lastRow)
        .ListObjects.Add(xlSrcRange, tblRange, , xlYes).Name = "tbl_Feiertage"

        ' Spaltenbreiten
        .Columns("A:A").ColumnWidth = 28
        .Columns("B:B").ColumnWidth = 15
        .Columns("C:C").ColumnWidth = 22

        ' Sortierung nach Datum (nur wenn Datenzeilen existieren)
        Dim lo As ListObject
        Set lo = .ListObjects("tbl_Feiertage")
        If lo.ListRows.Count > 0 Then
            With lo.Sort
                .SortFields.Clear
                .SortFields.Add key:=lo.ListColumns("Datum").DataBodyRange, _
                    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                .Header = xlYes
                .Apply
            End With
        End If
    End With  ' <- wichtig: schließt With ws

    MsgBox "Feiertage " & jahr & " (" & blCode & ") wurden erstellt!", vbInformation
End Sub

Private Sub EintragenFeiertage_Bundesland(ByVal ws As Worksheet, ByVal jahr As Long, ByVal blCode As String)
    ' Gemeinsame Basis + länderspezifische Feiertage
    Dim ostersonntag As Date: ostersonntag = BerechneOstersonntag(jahr)
    Dim z As Long: z = 2

    ' --- Feste (bundesweit) ---
    z = AddFeiertag(ws, z, "Neujahr", DateSerial(jahr, 1, 1), blCode)
    z = AddFeiertag(ws, z, "Tag der Arbeit", DateSerial(jahr, 5, 1), blCode)
    z = AddFeiertag(ws, z, "Tag der Deutschen Einheit", DateSerial(jahr, 10, 3), blCode)
    z = AddFeiertag(ws, z, "1. Weihnachtsfeiertag", DateSerial(jahr, 12, 25), blCode)
    z = AddFeiertag(ws, z, "2. Weihnachtsfeiertag", DateSerial(jahr, 12, 26), blCode)

    ' --- Bewegliche (bundesweit) ---
    z = AddFeiertag(ws, z, "Karfreitag", ostersonntag - 2, blCode)
    z = AddFeiertag(ws, z, "Ostermontag", ostersonntag + 1, blCode)
    z = AddFeiertag(ws, z, "Christi Himmelfahrt", ostersonntag + 39, blCode)
    z = AddFeiertag(ws, z, "Pfingstmontag", ostersonntag + 50, blCode)

    ' --- Länderspezifisch ---

    ' Heilige Drei Könige (BW, BY, ST)
    Select Case blCode
        Case "BW", "BY", "ST"
            z = AddFeiertag(ws, z, "Heilige Drei Könige", DateSerial(jahr, 1, 6), blCode)
    End Select

    ' Fronleichnam (BW, BY, HE, NW, RP, SL)
    Select Case blCode
        Case "BW", "BY", "HE", "NW", "RP", "SL"
            z = AddFeiertag(ws, z, "Fronleichnam", ostersonntag + 60, blCode)
    End Select

    ' Reformationstag (BB, MV, SN, ST, TH, HB, HH, NI, SH, BE)
    Select Case blCode
        Case "BB", "MV", "SN", "ST", "TH", "HB", "HH", "NI", "SH", "BE"
            z = AddFeiertag(ws, z, "Reformationstag", DateSerial(jahr, 10, 31), blCode)
    End Select

    ' Allerheiligen (BW, BY, NW, RP, SL)
    Select Case blCode
        Case "BW", "BY", "NW", "RP", "SL"
            z = AddFeiertag(ws, z, "Allerheiligen", DateSerial(jahr, 11, 1), blCode)
    End Select

    ' Mariä Himmelfahrt (BY*, SL) – hier pragmatisch für BY/SL
    Select Case blCode
        Case "BY", "SL"
            z = AddFeiertag(ws, z, "Mariä Himmelfahrt", DateSerial(jahr, 8, 15), blCode)
    End Select

    ' Buß- und Bettag (SN)
    If blCode = "SN" Then
        z = AddFeiertag(ws, z, "Buß- und Bettag", BussUndBettag(jahr), blCode)
    End If

    ' Datumsformat
    ws.Range("B2:B" & z - 1).NumberFormat = "dd.mm.yyyy"
End Sub

Private Function AddFeiertag(ByVal ws As Worksheet, ByVal z As Long, _
                             ByVal nameFeiertag As String, ByVal dt As Date, _
                             ByVal blCode As String) As Long
    ' Duplikate prüfen (Name + Datum)
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 1

    For r = 2 To lastRow
        If ws.Cells(r, 1).Value = nameFeiertag And _
           CLng(ws.Cells(r, 2).Value) = CLng(dt) Then
            AddFeiertag = z
            Exit Function
        End If
    Next r

    ' Eintragen
    ws.Cells(z, 1).Value = nameFeiertag
    ws.Cells(z, 2).Value = dt
    ws.Cells(z, 3).Value = blCode   ' <— neu
    AddFeiertag = z + 1
End Function

' Zusätzliche Funktion zum Aktualisieren der Feiertage
Sub AktualisiereFeiertage()
    Dim jahr As Integer
    jahr = ThisWorkbook.Worksheets("Anleitung").Range("C2").Value
    
    If jahr < 1900 Or jahr > 2100 Then
        MsgBox "Bitte geben Sie ein gültiges Jahr zwischen 1900 und 2100 ein!"
        Exit Sub
    End If
    
    EinrichtenFeiertage
End Sub

Private Function BussUndBettag(ByVal jahr As Long) As Date
    ' Buß- und Bettag = Mittwoch vor dem 23.11.
    Dim d As Date
    d = DateSerial(jahr, 11, 23)
    ' Rückwärts bis zum Mittwoch
    Do While Weekday(d, vbMonday) <> 3 ' 3 = Mittwoch bei Montag=1
        d = d - 1
    Loop
    BussUndBettag = d
End Function

Public Function BerechneOstersonntag(jahr As Long) As Date
    ' Gauß
    Dim a&, b&, c&, d&, e&, f&, g&, h&, i&, k&, l&, m&, n&, p&
    a = jahr Mod 19
    b = jahr \ 100
    c = jahr Mod 100
    d = b \ 4
    e = b Mod 4
    f = (b + 8) \ 25
    g = (b - f + 1) \ 3
    h = (19 * a + b - d - g + 15) Mod 30
    i = c \ 4
    k = c Mod 4
    l = (32 + 2 * e + 2 * i - h - k) Mod 7
    m = (a + 11 * h + 22 * l) \ 451
    n = (h + l - 7 * m + 114) \ 31
    p = (h + l - 7 * m + 114) Mod 31
    BerechneOstersonntag = DateSerial(jahr, n, p + 1)
End Function

Private Function NormalizeBL(ByVal inputCode As String) As String
    Dim c As String: c = UCase$(Trim$(inputCode))
    If c = "NRW" Then c = "NW"
    NormalizeBL = c
End Function

Public Sub Test_EintragenFeiertage_Bundesland()
    Dim ws As Worksheet, y As Long, bl As String
    Set ws = ThisWorkbook.Worksheets("Feiertage")
    y = ThisWorkbook.Worksheets("Anleitung").Range("C2").Value
    bl = NormalizeBL("NRW")
    ws.Cells.Clear
    ws.Range("A1").Value = "Feiertag": ws.Range("B1").Value = "Datum"
    EintragenFeiertage_Bundesland ws, y, bl
    ' Einfach sortieren
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If lastRow >= 2 Then
        ws.Range("A1:B" & lastRow).Sort Key1:=ws.Range("B2"), Order1:=xlAscending, Header:=xlYes
    End If
End Sub

