Attribute VB_Name = "Z_Konfiguration"
'Attribute VB_Name = "Z_Konfiguration"
Option Explicit

'===============================================================================
' Z_Konfiguration – zentrale Einstellungen / Namen / Grenzen / Farben
' Excel 2019 DE – alle Module greifen über diese API zu (keine hart codierten Literale)
' Enthält zusätzlich Wrapper mit bisherigen Funktionsnamen (Kompatibilität).
'===============================================================================

' --- Arbeitsblätter (Namen) ---------------------------------------------------
Public Function CFG_Sheet_Admin() As String:           CFG_Sheet_Admin = "Administration": End Function
Public Function CFG_Sheet_Personen() As String:        CFG_Sheet_Personen = "Personen": End Function
Public Function CFG_Sheet_BAO() As String:             CFG_Sheet_BAO = "BAO": End Function
Public Function CFG_Sheet_Feiertage() As String:       CFG_Sheet_Feiertage = "Feiertage": End Function
Public Function CFG_Sheet_Ferien() As String:          CFG_Sheet_Ferien = "Ferien": End Function
Public Function CFG_Sheet_Anleitung() As String:       CFG_Sheet_Anleitung = "Anleitung": End Function
Public Function CFG_Sheet_Bereitschaften() As String:  CFG_Sheet_Bereitschaften = "Bereitschaften": End Function
Public Function CFG_Sheet_Legende() As String:         CFG_Sheet_Legende = "Legende": End Function
Public Function CFG_Sheet_Information() As String:     CFG_Sheet_Information = "Information": End Function

' Monatsblätter in Reihenfolge 1..12
Public Function CFG_MonatsNamen() As Variant
    CFG_MonatsNamen = Array("Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez")
End Function

' --- Tabellen (ListObjects) ---------------------------------------------------
Public Function CFG_Table_Personen() As String:        CFG_Table_Personen = "tbl_Personen": End Function
Public Function CFG_Table_BAO() As String:             CFG_Table_BAO = "tbl_BAO": End Function
Public Function CFG_Table_Feiertage() As String:       CFG_Table_Feiertage = "tbl_Feiertage": End Function
Public Function CFG_Table_Ferien() As String:          CFG_Table_Ferien = "tbl_Ferien": End Function
Public Function CFG_Table_MVL() As String:             CFG_Table_MVL = "tbl_MVL": End Function

' --- Spalten-/Zeilen-Definitionen --------------------------------------------
Public Function CFG_ErsteDatenZeile() As Long:         CFG_ErsteDatenZeile = 6: End Function
Public Function CFG_Spalte_Personen() As Long:         CFG_Spalte_Personen = 2 ' B
End Function
Public Function CFG_Spalte_Team() As Long:             CFG_Spalte_Team = 3 ' C
End Function
Public Function CFG_ErsteTagSpalte() As Long:          CFG_ErsteTagSpalte = 4 ' D
End Function
Public Function CFG_LetzteTagSpalte() As Long:         CFG_LetzteTagSpalte = 66 ' BM
End Function

' --- Länderspezifisches -------------------------------------------------------
Public Function CFG_ListSep() As String:               CFG_ListSep = M_SafeApp.ListSep(): End Function

' --- Spaltenbreiten / Schrift / Ausrichtung -----------------------------------
Public Function SpaltenbreiteA() As Double:            SpaltenbreiteA = 2      ' Leer-/Puffer
End Function
Public Function SpaltenbreiteB() As Double:            SpaltenbreiteB = 6      ' Personen-/Stärke
End Function
Public Function SpaltenbreiteC() As Double:            SpaltenbreiteC = 16     ' Team / Funktion
End Function
Public Function SpaltenbreiteTage() As Double:         SpaltenbreiteTage = 3.5 ' Tagesspalten
End Function

Public Function GetStandardSchriftart() As String:     GetStandardSchriftart = "Calibri"
End Function
Public Function GetStandardSchriftgroesse() As Double: GetStandardSchriftgroesse = 10
End Function
Public Function GetHeaderSchriftgroesse() As Double:   GetHeaderSchriftgroesse = 11
End Function
Public Function GetMonatSchriftgroesse() As Double:    GetMonatSchriftgroesse = 14
End Function

Public Function GetAusrichtungStandard() As Long:      GetAusrichtungStandard = xlCenter
End Function
Public Function GetAusrichtungSpalteB() As Long:       GetAusrichtungSpalteB = xlRight
End Function

' --- Rahmenfarben/-stärken ----------------------------------------------------
Public Function RahmenFarbeGrau() As Long:             RahmenFarbeGrau = RGB(200, 200, 200)
End Function
Public Function RahmenFarbeSchwarz() As Long:          RahmenFarbeSchwarz = RGB(0, 0, 0)
End Function
Public Function RahmenStaerkeHaar() As XlBorderWeight: RahmenStaerkeHaar = xlHairline
End Function
Public Function RahmenStaerkeDuenn() As XlBorderWeight: RahmenStaerkeDuenn = xlThin
End Function
Public Function RahmenStaerkeMittel() As XlBorderWeight: RahmenStaerkeMittel = xlMedium
End Function

' --- Farben -------------------------------------------------------------------
Public Function CFG_Farbe_Heute() As Long:             CFG_Farbe_Heute = RGB(237, 125, 49)    ' Orange
End Function
Public Function CFG_Farbe_Text_Heute() As Long:        CFG_Farbe_Text_Heute = RGB(255, 255, 255)
End Function

Public Function CFG_Farbe_WeekendHell() As Long:       CFG_Farbe_WeekendHell = RGB(217, 217, 217)
End Function
Public Function CFG_Farbe_WeekendDunkel() As Long:     CFG_Farbe_WeekendDunkel = RGB(191, 191, 191)
End Function
Public Function CFG_Farbe_Text_Weekend() As Long:      CFG_Farbe_Text_Weekend = RGB(0, 0, 0)
End Function

Public Function CFG_Farbe_ZeileGerade() As Long:       CFG_Farbe_ZeileGerade = RGB(242, 242, 242)
End Function
Public Function CFG_Farbe_ZeileUngerade() As Long:     CFG_Farbe_ZeileUngerade = RGB(255, 255, 255)
End Function
Public Function CFG_Farbe_Text_Schwarz() As Long:      CFG_Farbe_Text_Schwarz = RGB(0, 0, 0)
End Function
Public Function CFG_Farbe_Text_Weiss() As Long:        CFG_Farbe_Text_Weiss = RGB(255, 255, 255)
End Function
Public Function CFG_Farbe_Text_Grau() As Long:         CFG_Farbe_Text_Grau = RGB(120, 120, 120)
End Function

Public Function CFG_Farbe_Ferien() As Long:            CFG_Farbe_Ferien = RGB(255, 242, 204)
End Function
Public Function FarbeKalenderHeader() As Long:         FarbeKalenderHeader = RGB(235, 241, 222)
End Function

' Codespezifische Farben/Muster
Public Function FarbeAnwesenheit() As Long:            FarbeAnwesenheit = RGB(219, 238, 243)
End Function
Public Function FarbeAnwesenheitZ() As Long:           FarbeAnwesenheitZ = RGB(197, 224, 180)
End Function
Public Function FarbeUrlaub() As Long:                 FarbeUrlaub = RGB(255, 230, 153)
End Function
Public Function FarbeUrlaubVorschuss() As Long:        FarbeUrlaubVorschuss = RGB(255, 192, 0)
End Function
Public Function FarbeAbwesenheit() As Long:            FarbeAbwesenheit = RGB(255, 199, 206)
End Function
Public Function FarbeSonderurlaub() As Long:           FarbeSonderurlaub = RGB(204, 192, 218)
End Function
Public Function FarbeBAOMuster() As Long:              FarbeBAOMuster = RGB(237, 125, 49)
End Function
' --- MVL-Farbe zentral --------------------------------------------------------
' Quelle: Anleitung!C4 (optional)
'   - Akzeptiert: "#RRGGBB" oder "R,G,B" (z. B. "180,198,231")
'   - Sonst Default
Public Function CFG_Farbe_MVL() As Long
    Dim s As String, col As Long
    On Error Resume Next
    s = CStr(ThisWorkbook.Worksheets("Anleitung").Range("C4").Value)
    On Error GoTo 0
    s = Trim$(s)

    If ParseColorSpec(s, col) Then
        CFG_Farbe_MVL = col
    Else
        ' Default wie bisheriges Blauton für Bereitschaft
        CFG_Farbe_MVL = RGB(180, 198, 231)
    End If
End Function

' Rückwärtskompatibel: alter Name zeigt auf neue Konfiguration
Public Function FarbeBereitschaftMuster() As Long
    FarbeBereitschaftMuster = CFG_Farbe_MVL()
End Function

' Hilfsparser: "#RRGGBB" oder "R,G,B" -> Long
Private Function ParseColorSpec(ByVal s As String, ByRef outColor As Long) As Boolean
    Dim r As Long, g As Long, b As Long, parts As Variant, hexv As Long

    If Len(s) = 7 And Left$(s, 1) = "#" Then
        On Error Resume Next
        hexv = CLng("&H" & Mid$(s, 2))
        On Error GoTo 0
        If hexv >= 0 Then
            r = (hexv \ &H10000) And &HFF
            g = (hexv \ &H100&) And &HFF
            b = hexv And &HFF
            outColor = RGB(r, g, b)
            ParseColorSpec = True
            Exit Function
        End If
    End If

    If InStr(1, s, ",") > 0 Then
        parts = Split(s, ",")
        If UBound(parts) = 2 Then
            r = val(Trim$(parts(0)))
            g = val(Trim$(parts(1)))
            b = val(Trim$(parts(2)))
            If r >= 0 And r <= 255 And g >= 0 And g <= 255 And b >= 0 And b <= 255 Then
                outColor = RGB(r, g, b)
                ParseColorSpec = True
                Exit Function
            End If
        End If
    End If
End Function

' Zeilen-/Header-Spezialfarben
Public Function FarbeGruppe() As Long:                 FarbeGruppe = RGB(220, 230, 241) ' Team-Stärke-Zeile
End Function
Public Function GetBAOZeilenFormatierung() As Long:    GetBAOZeilenFormatierung = RGB(255, 242, 204)
End Function

' --- Bereiche der Tagesköpfe (Zeile 4/5) -------------------------------------
Public Function CFG_Range_Tageszeile4(ByVal ws As Worksheet) As Range
    Set CFG_Range_Tageszeile4 = ws.Range(ws.Cells(4, CFG_ErsteTagSpalte), ws.Cells(4, CFG_LetzteTagSpalte))
End Function
Public Function CFG_Range_Tageszeile5(ByVal ws As Worksheet) As Range
    Set CFG_Range_Tageszeile5 = ws.Range(ws.Cells(5, CFG_ErsteTagSpalte), ws.Cells(5, CFG_LetzteTagSpalte))
End Function

' --- Validierungsquellen ------------------------------------------------------
' Standardlisten, bis dedizierte Funktionen bereitstehen.
Public Function GetAnwesenheitsCodes() As String
    ' Beispiele: P(Anwesend), S(Schicht), TA(Tagesaufgabe), Z(Home/zentral), UR(Url.), UV(Url.-Vorsch.), ABW(Abw.), GL(Gleitzeit), SU(Sonderurl.)
    Dim s As String: s = Join(Array("", "P", "S", "TA", "Z", "UR", "UV", "ABW", "GL", "SU", "BE", "BE-D", "BA-B", "BA-D", "BAO"), CFG_ListSep())
    GetAnwesenheitsCodes = s
End Function

Public Function GetAufgabenCodes() As String
    ' Aufgaben rechts – leere Auswahl erlaubt
    Dim s As String: s = Join(Array("", "Disp", "Proj", "Doku", "Schul", "Backlog", "Meeting"), CFG_ListSep())
    GetAufgabenCodes = s
End Function

'===============================================================================
' KOMPATIBILITÄTS-WRAPPER (alte Funktions-/Konstantennamen weiter nutzbar)
'===============================================================================
Public Property Get ERSTE_DATEN_ZEILE() As Long:       ERSTE_DATEN_ZEILE = CFG_ErsteDatenZeile(): End Property
Public Property Get PERSONEN_SPALTE() As Long:         PERSONEN_SPALTE = CFG_Spalte_Personen(): End Property
Public Property Get TEAM_SPALTE() As Long:             TEAM_SPALTE = CFG_Spalte_Team(): End Property
Public Property Get ERSTE_TAG_SPALTE() As Long:        ERSTE_TAG_SPALTE = CFG_ErsteTagSpalte(): End Property

' Farben (bestehende Bezeichner)
Public Function FarbeWochenendeDunkel() As Long:       FarbeWochenendeDunkel = CFG_Farbe_WeekendDunkel(): End Function
Public Function FarbeWochenendeHell() As Long:         FarbeWochenendeHell = CFG_Farbe_WeekendHell(): End Function
Public Function FarbeZeileGerade() As Long:            FarbeZeileGerade = CFG_Farbe_ZeileGerade(): End Function
Public Function FarbeZeileUngerade() As Long:          FarbeZeileUngerade = CFG_Farbe_ZeileUngerade(): End Function
Public Function SchriftfarbeSchwarz() As Long:         SchriftfarbeSchwarz = CFG_Farbe_Text_Schwarz(): End Function
Public Function SchriftfarbeWeiss() As Long:           SchriftfarbeWeiss = CFG_Farbe_Text_Weiss(): End Function
Public Function SchriftfarbeGrau() As Long:            SchriftfarbeGrau = CFG_Farbe_Text_Grau(): End Function
Public Function FarbeFerien() As Long:                 FarbeFerien = CFG_Farbe_Ferien(): End Function
Public Function FarbeHeuteHell() As Long:              FarbeHeuteHell = CFG_Farbe_Heute(): End Function

' --- MVL-Zeilenname zentral konfigurierbar ---
Public Function CFG_MVL_Zeilenname() As String
    ' Standard: exakt so, wie die BAO-/MVL-Zeile im Monatsblatt heißen soll
    CFG_MVL_Zeilenname = "MVL Bereitschaft"
End Function

' --- Wrapper wie bei ERSTE_TAG_SPALTE, der bislang fehlte ---
Public Property Get LETZTE_TAG_SPALTE() As Long
    LETZTE_TAG_SPALTE = CFG_LetzteTagSpalte()
End Property

' --- Bundesland-Konfiguration -------------------------------------------------
Public Function CFG_Bundesland_Default() As String
    ' NRW als Default
    CFG_Bundesland_Default = "NW" ' ISO-16 Bundeslandcodes: BW, BY, BE, BB, HB, HH, HE, MV, NI, NW, RP, SL, SN, ST, SH, TH
End Function

Public Function CFG_BundeslandListeCSV() As String
    ' Anzeige "Code – Name" für Dropdown
    CFG_BundeslandListeCSV = "BW – Baden-Württemberg," & _
                             "BY – Bayern," & _
                             "BE – Berlin," & _
                             "BB – Brandenburg," & _
                             "HB – Bremen," & _
                             "HH – Hamburg," & _
                             "HE – Hessen," & _
                             "MV – Mecklenburg-Vorpommern," & _
                             "NI – Niedersachsen," & _
                             "NW – Nordrhein-Westfalen," & _
                             "RP – Rheinland-Pfalz," & _
                             "SL – Saarland," & _
                             "SN – Sachsen," & _
                             "ST – Sachsen-Anhalt," & _
                             "SH – Schleswig-Holstein," & _
                             "TH – Thüringen"
End Function

Public Function CFG_GetBundeslandCode() As String
    ' Liest den gewählten Code aus Anleitung!C3 (Format "NW – Nordrhein-Westfalen")
    Dim v As String
    On Error Resume Next
    v = CStr(ThisWorkbook.Worksheets("Anleitung").Range("C3").Value)
    On Error GoTo 0
    v = Trim(v)
    If Len(v) = 0 Then
        CFG_GetBundeslandCode = CFG_Bundesland_Default()
        Exit Function
    End If
    ' Erwartetes Format "XY – Name": nimm die ersten 2 Zeichen als Code
    CFG_GetBundeslandCode = UCase(Left$(v, 2))
End Function

' --- Z_Konfiguration: zentrale Monatsblatt-Erkennung ---
Public Function CFG_IsMonatsblattName(ByVal blattName As String) As Boolean
    Dim monate As Variant, i As Long
    monate = Array("Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez")
    For i = LBound(monate) To UBound(monate)
        If StrComp(blattName, monate(i), vbTextCompare) = 0 Then
            CFG_IsMonatsblattName = True
            Exit Function
        End If
    Next i
End Function

' --- Kompatibilität: zentraler Wrapper ---
Public Function IstMonatsblatt(ByVal blattName As String) As Boolean
    IstMonatsblatt = CFG_IsMonatsblattName(blattName)
End Function

Public Function GetAnwesenheitsCodesForValidation() As String
    ' US-kompatible Komma-Liste für Excel-Validierung
    Dim arr As Variant
    arr = Array("", "P", "S", "TA", "Z", "UR", "UV", "ABW", "GL", "SU", "BE", "BE-D", "BA-B", "BA-D", "BAO")
    GetAnwesenheitsCodesForValidation = Join(arr, ",")  ' ? Immer Komma!
End Function

Public Function GetAufgabenCodesForValidation() As String
    ' US-kompatible Komma-Liste für Excel-Validierung
    Dim arr As Variant
    arr = Array("", "Disp", "Proj", "Doku", "Schul", "Backlog", "Meeting")
    GetAufgabenCodesForValidation = Join(arr, ",")
End Function
