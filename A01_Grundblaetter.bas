Attribute VB_Name = "A01_Grundblaetter"
'Attribute VB_Name = "A01_Grundblaetter"
Option Explicit

'===============================================================================
' A01_Grundblätter
' - Nicht-destruktiv: legt fehlende Basisblätter an, vorhandene werden nicht gelöscht
' - Erstellt/initialisiert: Administration, Anleitung, BAO, Personen,
'   Bereitschaften, Feiertage, Ferien, Legende, Information
' - Idempotent & schnell (Intersect-Checks, kein Select)
'===============================================================================

Public Sub ErstelleGrundstruktur()
    Dim t0 As Double: t0 = Timer
    Debug.Print "A01: ErstelleGrundstruktur (nicht-destruktiv)  ", Format(Now, "dd.mm.yyyy HH:nn:ss")

    M_SafeApp.BeginFastOps True, True, True
    On Error GoTo Clean

    Dim names As Variant
    names = Array("Administration", "Anleitung", "BAO", "Personen", _
                  "Bereitschaften", "Feiertage", "Ferien", "Legende", "Information")

    Dim i As Long
    For i = LBound(names) To UBound(names)
        EnsureSheet CStr(names(i))
    Next i

    ' Minimal-Initialisierung spezieller Blätter
    InitSheet_Administration ThisWorkbook.Worksheets("Administration")
    InitSheet_Information ThisWorkbook.Worksheets("Information")

    Debug.Print "A01: OK  (", Format(Timer - t0, "0.00"), " s)"

Clean:
    M_SafeApp.EndFastOps
End Sub

'--- Helfer --------------------------------------------------------------------

Private Sub EnsureSheet(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        Debug.Print "A01: Blatt erstellt -> "; sheetName
    Else
        Debug.Print "A01: Blatt vorhanden -> "; sheetName
    End If
End Sub

Private Sub InitSheet_Administration(ByVal ws As Worksheet)
    ' Sehr schlanke, nicht-invasive Startseite
    With ws
        If .Range("A1").Value <> "Administration" Then
            .Cells.Clear
            .Range("A1").Value = "Administration"
            .Range("A1").Font.Bold = True
            .Range("A1").Font.Size = 16

            Dim r As Range
            Set r = .Range("A3")
            r.Value = "Startpunkte:"
            r.Font.Bold = True

            .Range("A5").Value = "A-Masterlauf (A00)"
            .Range("B5").Value = "Admin_Starte_A_Master"
            .Range("A6").Value = "A-Einzelmodule (A01–A07)"
            .Range("B6").Value = "Admin_Starte_A_Einzel"
            .Range("A7").Value = "B-Gruppe (Monatssetup)"
            .Range("B7").Value = "Admin_Starte_B_Monatssetup"
            .Range("A8").Value = "C-Gruppe (Formatierung)"
            .Range("B8").Value = "Admin_Starte_C_Formatierung"
            .Range("A9").Value = "D-Gruppe (BAO-Integration)"
            .Range("B9").Value = "Admin_Starte_D_BAOIntegration"
            .Range("A10").Value = "F-Gruppe (Personen-Tools)"
            .Range("B10").Value = "Admin_Starte_F_PersonenTools"
            .Columns("A:B").AutoFit
        End If
    End With
End Sub

Private Sub InitSheet_Information(ByVal ws As Worksheet)
    With ws
        If .Range("A1").Value <> "Information" Then
            .Cells.Clear
            .Range("A1").Value = "Information"
            .Range("A1").Font.Bold = True
            .Range("A1").Font.Size = 16

            .Range("A3").Value = "Version"
            .Range("B3").Value = Format(Now, "yyyy.mm.dd") & " (Build " & Format(Now, "HHnn") & ")"

            .Range("A5").Value = "Hinweise"
            .Range("A6").Value = "• Alle Namen/Parameter zentral in Z_Konfiguration pflegen."
            .Range("A7").Value = "• A00_MasterGrundstruktur erzeugt die komplette Basis."
            .Range("A8").Value = "• B01/B02/B03 bauen Monatsblätter, Dropdowns & Teamstärke."
            .Range("A9").Value = "• C01 formatiert, D01 integriert BAO, F01 synchronisiert Personen."
            .Columns("A:B").AutoFit
        End If
    End With
End Sub


