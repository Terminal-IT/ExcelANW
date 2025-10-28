Attribute VB_Name = "A00_MasterGrundstruktur"
'Attribute VB_Name = "A00_MasterGrundstruktur"
' =====================================================
' MASTER-MODUL F�R VOLLST�NDIGE GRUNDSTRUKTUR
' Ruft alle A01-A07 Module in korrekter Reihenfolge auf
' =====================================================

Sub ErstelleKompletteGrundstruktur()
' Erstellt die vollst�ndige Grundstruktur der Anwesenheitsverwaltung
    
    Dim startZeit As Double
    startZeit = Timer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    On Error GoTo ErrorHandler
    
    Dim antwort As VbMsgBoxResult
    antwort = MsgBox("VOLLST�NDIGE GRUNDSTRUKTUR ERSTELLEN" & vbCrLf & vbCrLf & _
                     "Dies erstellt:" & vbCrLf & _
                     "� Alle Grundbl�tter (Anleitung, Personen, etc.)" & vbCrLf & _
                     "� Feiertage f�r NRW" & vbCrLf & _
                     "� Ferienzeiten" & vbCrLf & _
                     "� Personen-Struktur" & vbCrLf & _
                     "� Bereitschaftsplanung" & vbCrLf & _
                     "� BAO-Tabelle" & vbCrLf & _
                     "� Administration und Legende" & vbCrLf & vbCrLf & _
                     "Fortfahren?", vbYesNo + vbQuestion, "Grundstruktur erstellen")
    
    If antwort = vbNo Then Exit Sub
    
    Debug.Print "=== VOLLST�NDIGE GRUNDSTRUKTUR-ERSTELLUNG GESTARTET ==="
    Debug.Print "Startzeit: " & Now
    
    ' SCHRITT 1: Grundbl�tter erstellen
    Debug.Print "Schritt 1/7: Erstelle Grundbl�tter..."
    Call ErstelleGrundstruktur  ' A01
    Debug.Print "Grundbl�tter erstellt"
    
    ' SCHRITT 2: Anleitung einrichten
    Debug.Print "Schritt 2/7: Richte Anleitung ein..."
    Call EinrichtenAnleitung    ' A02
    Debug.Print "Anleitung eingerichtet"
    
    ' SCHRITT 3: Feiertage erstellen
    Debug.Print "Schritt 3/7: Erstelle Feiertage..."
    Call EinrichtenFeiertage    ' A03
    Debug.Print "Feiertage f�r NRW erstellt"
    
    ' SCHRITT 4: Ferien einrichten
    Debug.Print "Schritt 4/7: Richte Ferien ein..."
    Call EinrichtenFerien       ' A04
    Debug.Print "Ferienzeiten eingetragen"
    
    ' SCHRITT 5: Personen-Struktur
    Debug.Print "Schritt 5/7: Richte Personen ein..."
    Call EinrichtenPersonen     ' A05
    Debug.Print "Personen-Struktur erstellt"
    
    ' SCHRITT 6: Bereitschaften
    Debug.Print "Schritt 6/7: Richte Bereitschaften ein..."
    Call EinrichtenBereitschaften  ' A06
    Debug.Print "MVL-Bereitschaften berechnet"
    
    ' SCHRITT 7: BAO-Struktur
    Debug.Print "Schritt 7/7: Richte BAO ein..."
    Call EinrichtenBAO          ' A07
    Debug.Print "BAO-Tabelle erstellt"
    
    ' OPTIONAL: Zus�tzliche Konfiguration
    Debug.Print "Zusatz: F�hre finale Konfiguration durch..."
    Call FinaleGrundkonfiguration
    Debug.Print "Finale Konfiguration abgeschlossen"
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    Dim endZeit As Double
    endZeit = Timer
    
    Debug.Print "=== GRUNDSTRUKTUR-ERSTELLUNG ABGESCHLOSSEN ==="
    Debug.Print "Endzeit: " & Now
    Debug.Print "Dauer: " & Format((endZeit - startZeit) / 86400, "hh:mm:ss")
    
    ' Erfolgs-Meldung mit Details
    MsgBox "[OK]  GRUNDSTRUKTUR ERFOLGREICH ERSTELLT!" & vbCrLf & vbCrLf & _
           "Erstellte Komponenten:" & vbCrLf & _
           "[OK] Grundbl�tter (Anleitung, Personen, etc.)" & vbCrLf & _
           "[OK] Feiertage f�r NRW " & Year(Now) & vbCrLf & _
           "[OK] Ferienzeiten" & vbCrLf & _
           "[OK] Personen-Beispiele" & vbCrLf & _
           "[OK] MVL-Bereitschaftsplanung" & vbCrLf & _
           "[OK] BAO-Tabelle" & vbCrLf & _
           "[OK] Administration & Legende" & vbCrLf & vbCrLf & _
           "[OK] Dauer: " & Format((endZeit - startZeit) / 86400, "mm:ss") & " Minuten" & vbCrLf & vbCrLf & _
           "[OK] N�chster Schritt: Monatsbl�tter erstellen mit 'ErstelleMonatsblaetter'", _
           vbInformation, "Grundstruktur fertig"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    Debug.Print "FEHLER in ErstelleKompletteGrundstruktur: " & Err.Description
    
    MsgBox "FEHLER beim Erstellen der Grundstruktur:" & vbCrLf & vbCrLf & _
           "Fehler-Nr: " & Err.Number & vbCrLf & _
           "Beschreibung: " & Err.Description & vbCrLf & vbCrLf & _
           "Details im Direktfenster (Strg+G)", vbCritical, "Fehler"
    
End Sub

Sub FinaleGrundkonfiguration()
' F�hrt finale Konfigurationsschritte durch
    
    On Error Resume Next
    
    Debug.Print "Starte finale Grundkonfiguration..."
    
    ' Aktiviere das Anleitung-Blatt als Standard
    ThisWorkbook.Worksheets("Anleitung").Activate
    
    ' Setze das Jahr in der Anleitung auf aktuelles Jahr (falls leer)
    If ThisWorkbook.Worksheets("Anleitung").Range("C2").Value = "" Or _
       ThisWorkbook.Worksheets("Anleitung").Range("C2").Value = 2025 Then
        ThisWorkbook.Worksheets("Anleitung").Range("C2").Value = Year(Now)
        Debug.Print "Jahr in Anleitung auf " & Year(Now) & " gesetzt"
    End If
    
    ' Pr�fe alle erstellten Abh�ngigkeiten
    Call PruefeGrundstrukturAbhaengigkeiten
    
    ' Speichere die Arbeitsmappe
    On Error Resume Next
    ThisWorkbook.Save
    If Err.Number = 0 Then
        Debug.Print "Arbeitsmappe erfolgreich gespeichert"
    Else
        Debug.Print "Warnung: Arbeitsmappe konnte nicht gespeichert werden: " & Err.Description
    End If
    
    On Error GoTo 0
    
End Sub

Sub PruefeGrundstrukturAbhaengigkeiten()
' �berpr�ft ob alle Grundstruktur-Komponenten korrekt erstellt wurden
    
    Debug.Print "Pr�fe Grundstruktur-Abh�ngigkeiten..."
    
    Dim fehlerListe As String
    fehlerListe = ""
    
    ' Pr�fe Arbeitsbl�tter
    Dim benoetigteBl�tter As Variant
    benoetigteBl�tter = Array("Anleitung", "Personen", "Feiertage", "Ferien", "Bereitschaften", "BAO", "Administration", "Legende", "Information")
    
    Dim i As Integer
    For i = 0 To UBound(benoetigteBl�tter)
        On Error Resume Next
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(benoetigteBl�tter(i))
        On Error GoTo 0
        
        If ws Is Nothing Then
            fehlerListe = fehlerListe & "- Arbeitsblatt '" & benoetigteBl�tter(i) & "' fehlt" & vbCrLf
        Else
            Debug.Print "Blatt '" & benoetigteBl�tter(i) & "' vorhanden"
        End If
        Set ws = Nothing
    Next i
    
    ' Pr�fe wichtige Tabellen
    On Error Resume Next
    
    ' Feiertage-Tabelle
    Dim tblFeiertage As ListObject
    Set tblFeiertage = ThisWorkbook.Worksheets("Feiertage").ListObjects("tbl_Feiertage")
    If tblFeiertage Is Nothing Then
        fehlerListe = fehlerListe & "- Tabelle 'tbl_Feiertage' fehlt" & vbCrLf
    Else
        Debug.Print "Tabelle 'tbl_Feiertage' vorhanden (" & tblFeiertage.ListRows.Count & " Eintr�ge)"
    End If
    
    ' Ferien-Tabelle
    Dim tblFerien As ListObject
    Set tblFerien = ThisWorkbook.Worksheets("Ferien").ListObjects("tbl_Ferien")
    If tblFerien Is Nothing Then
        fehlerListe = fehlerListe & "- Tabelle 'tbl_Ferien' fehlt" & vbCrLf
    Else
        Debug.Print "Tabelle 'tbl_Ferien' vorhanden (" & tblFerien.ListRows.Count & " Eintr�ge)"
    End If
    
    ' MVL-Tabelle
    Dim tblMVL As ListObject
    Set tblMVL = ThisWorkbook.Worksheets("Bereitschaften").ListObjects("tbl_MVL")
    If tblMVL Is Nothing Then
        fehlerListe = fehlerListe & "- Tabelle 'tbl_MVL' fehlt" & vbCrLf
    Else
        Debug.Print "Tabelle 'tbl_MVL' vorhanden (" & tblMVL.ListRows.Count & " Eintr�ge)"
    End If
    
    On Error GoTo 0
    
    ' Ergebnis
    If fehlerListe = "" Then
        Debug.Print "Alle Grundstruktur-Abh�ngigkeiten erfolgreich erstellt!"
    Else
        Debug.Print "?Fehlende Komponenten:"
        Debug.Print fehlerListe
    End If
    
End Sub

' ===== HILFSFUNKTIONEN =====

Sub SchnellGrundstruktur()
' Schnelle Grundstruktur ohne R�ckfragen (f�r Entwicklung/Tests)
    
    Debug.Print "=== SCHNELLE GRUNDSTRUKTUR (ohne R�ckfragen) ==="
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    On Error Resume Next
    
    Call ErstelleGrundstruktur
    Call EinrichtenAnleitung
    Call EinrichtenFeiertage
    Call EinrichtenFerien
    Call EinrichtenPersonen
    Call EinrichtenBereitschaften
    Call EinrichtenBAO
    Call FinaleGrundkonfiguration
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    Debug.Print "=== SCHNELLE GRUNDSTRUKTUR ABGESCHLOSSEN ==="
    
    If Err.Number = 0 Then
        MsgBox "Schnelle Grundstruktur erfolgreich erstellt!", vbInformation
    Else
        MsgBox "Fehler bei schneller Grundstruktur: " & Err.Description, vbExclamation
    End If
    
End Sub

Sub ResetteGrundstruktur()
' L�scht alle Grundstruktur-Bl�tter f�r Neustart
    
    Dim antwort As VbMsgBoxResult
    antwort = MsgBox("?ACHTUNG: ALLE DATEN L�SCHEN" & vbCrLf & vbCrLf & _
                     "Dies l�scht ALLE Arbeitsbl�tter au�er dem ersten!" & vbCrLf & _
                     "Alle Daten gehen verloren!" & vbCrLf & vbCrLf & _
                     "Wirklich fortfahren?", vbYesNo + vbCritical, "Alle Daten l�schen")
    
    If antwort = vbNo Then Exit Sub
    
    Application.DisplayAlerts = False
    
    ' L�sche alle Bl�tter au�er dem ersten
    Do While ThisWorkbook.Worksheets.Count > 1
        ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count).Delete
    Loop
    
    ' Benenne erstes Blatt um
    ThisWorkbook.Worksheets(1).Name = "Start"
    ThisWorkbook.Worksheets(1).Cells.Clear
    ThisWorkbook.Worksheets(1).Range("A1").Value = "Bereit f�r neue Grundstruktur"
    
    Application.DisplayAlerts = True
    
    MsgBox "Grundstruktur zur�ckgesetzt!" & vbCrLf & _
           "Verwenden Sie 'ErstelleKompletteGrundstruktur' f�r Neustart.", vbInformation
    
End Sub

