Attribute VB_Name = "M_Basis"
'Attribute VB_Name = "M_Basis"
Option Explicit

'===============================================================================
' M_Basis – kleine, zentrale Helferfunktionen
'===============================================================================

' Liefert die letzte relevante Datenzeile eines Monatsblatts.
' Logik:
'   1) Suche von unten (Zeile 200 ? 6) die letzte belegte Zelle in Spalte C (TEAM_SPALTE)
'   2) Fallback: mindestens ERSTE_DATEN_ZEILE + 50 (robust für leere Blätter)
Public Function GetLetztePersonenzeile(ByVal ws As Worksheet) As Long
    On Error GoTo fallback
    Dim r As Long, c As Long
    c = Z_Konfiguration.CFG_Spalte_Team()  ' ? FIX: Konstante ersetzt
    r = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
    If r < Z_Konfiguration.CFG_ErsteDatenZeile() + 1 Then  ' ? FIX
        r = Z_Konfiguration.CFG_ErsteDatenZeile() + 50      ' ? FIX
    End If
    GetLetztePersonenzeile = r
    Exit Function
fallback:
    GetLetztePersonenzeile = Z_Konfiguration.CFG_ErsteDatenZeile() + 50  ' ? FIX
End Function

' Sicherer Range-Getter (verhindert Laufzeitfehler bei ungültigen Adressen)
Public Function SafeRange(ByVal ws As Worksheet, ByVal addr As String) As Range
    On Error Resume Next
    Set SafeRange = ws.Range(addr)
End Function

' Prüft, ob ein Blattname ein Monatsblatt ist (zentral verwendbar)
Public Function IstMonatsblattName(ByVal blattName As String) As Boolean
    Dim m As Variant, i As Long
    m = Z_Konfiguration.CFG_MonatsNamen
    For i = LBound(m) To UBound(m)
        If StrComp(blattName, CStr(m(i)), vbTextCompare) = 0 Then
            IstMonatsblattName = True
            Exit Function
        End If
    Next
End Function
