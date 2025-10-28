VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Jan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
' Jan – Monatsblatt-Event-Handler
' Duplizierbar für Feb, Mrz, Apr, Mai, Jun, Jul, Aug, Sep, Okt, Nov, Dez
' - Worksheet_Change: Schnellformatierung bei Dateneingabe
' - Worksheet_Activate: Intelligente Cursor-Positionierung
' - Worksheet_Deactivate: Optional - Cleanup/Statistik
'===============================================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Schnelle Formatierung bei Änderungen im Datenbereich
    On Error GoTo Done
    
    ' Nur reagieren wenn Target im relevanten Bereich
    Dim cTag As Long: cTag = Z_Konfiguration.CFG_ErsteTagSpalte()
    Dim cTagLast As Long: cTagLast = Z_Konfiguration.CFG_LetzteTagSpalte()
    Dim r0 As Long: r0 = Z_Konfiguration.CFG_ErsteDatenZeile()
    
    ' Prüfe ob Änderung im Datenbereich (ab Zeile 6, Spalte D bis BM)
    If Target.Row < r0 Or Target.Column < cTag Or Target.Column > cTagLast Then
        GoTo Done
    End If
    
    ' Prüfe ob mehr als 10 Zellen geändert (dann Skip für Performance)
    If Target.Cells.Count > 10 Then GoTo Done
    
    ' Schnellformatierung anwenden
    Application.EnableEvents = False
    C01_Formatierung.SchnellAktualisierung Target
    Application.EnableEvents = True
    
Done:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Activate()
    ' Intelligente Cursor-Positionierung beim Aktivieren des Blatts
    On Error Resume Next
    
    ' Standard: C4 (Monat/Jahr-Bereich)
    Dim zielZelle As Range
    Set zielZelle = Me.Range("C4")
    
    ' Optional: Springe zum heutigen Tag (falls im aktuellen Monat)
    Dim heute As Date: heute = Date
    Dim monatDieserSheet As Long
    monatDieserSheet = MonatsIndexAusName(Me.Name)
    
    If Month(heute) = monatDieserSheet And Year(heute) = Me.Range("C5").Value Then
        ' Finde Spalte mit heutigem Datum
        Dim sp As Long
        For sp = Z_Konfiguration.CFG_ErsteTagSpalte() To Z_Konfiguration.CFG_LetzteTagSpalte() Step 2
            If IsDate(Me.Cells(5, sp).Value) Then
                If CLng(Me.Cells(5, sp).Value) = CLng(heute) Then
                    ' Springe zur ersten Personenzeile beim heutigen Tag
                    Set zielZelle = Me.Cells(Z_Konfiguration.CFG_ErsteDatenZeile() + 1, sp)
                    Exit For
                End If
            End If
        Next sp
    End If
    
    ' Aktiviere Zielzelle
    zielZelle.Select
    
    On Error GoTo 0
End Sub

Private Sub Worksheet_Deactivate()
    ' Optional: Cleanup beim Verlassen des Blatts
    On Error Resume Next
    
    ' Hier könnten z.B. Statistiken aktualisiert werden
    ' oder temporäre Markierungen entfernt werden
    
    On Error GoTo 0
End Sub

'--- Hilfsfunktion ---
Private Function MonatsIndexAusName(ByVal blattName As String) As Long
    ' Gibt Monatsnummer zurück (Jan=1, Feb=2, ..., Dez=12)
    Select Case UCase$(blattName)
        Case "JAN": MonatsIndexAusName = 1
        Case "FEB": MonatsIndexAusName = 2
        Case "MRZ", "MÄR": MonatsIndexAusName = 3
        Case "APR": MonatsIndexAusName = 4
        Case "MAI": MonatsIndexAusName = 5
        Case "JUN": MonatsIndexAusName = 6
        Case "JUL": MonatsIndexAusName = 7
        Case "AUG": MonatsIndexAusName = 8
        Case "SEP": MonatsIndexAusName = 9
        Case "OKT": MonatsIndexAusName = 10
        Case "NOV": MonatsIndexAusName = 11
        Case "DEZ": MonatsIndexAusName = 12
        Case Else: MonatsIndexAusName = 0
    End Select
End Function
