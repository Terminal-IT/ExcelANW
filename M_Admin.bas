Attribute VB_Name = "M_Admin"
'Attribute VB_Name = "M_Admin"
Option Explicit

'===============================================================================
' M_Admin: führt Gruppen A–F aus.
' NEU:
'   - A-Gruppe hat zwei Modi: Master ODER Einzelmodule (exklusiv)
'   - Keine doppelten Läufe mehr
'   - Saubere Konsolen-Headlines und SafeRunEx-Verwendung
'===============================================================================

Private Sub RunList(ByRef procNames As Variant, ByVal headline As String, _
                    Optional ByVal moduleName As String = vbNullString)
    Dim i As Long, ok As Boolean, p As String
    Debug.Print String(60, "-")
    Debug.Print headline & "  (" & Format(Now, "dd.mm.yyyy HH:nn:ss") & ")"
    For i = LBound(procNames) To UBound(procNames)
        p = CStr(procNames(i))
        If Len(p) > 0 Then
            ok = M_SafeApp.SafeRunEx(p, moduleName)
            Debug.Print IIf(ok, "OK   ", "SKIP ") & p
        End If
    Next i
End Sub

'=============================== A-GRUPPE ======================================
' Öffentlicher Einstieg mit Auswahl: Masterlauf ODER Einzelmodule.
Public Sub Admin_Starte_A_Einrichten()
    Dim modus As VbMsgBoxResult
    modus = MsgBox( _
        "A-Gruppe ausführen:" & vbCrLf & vbCrLf & _
        "Ja  = Masterlauf (A00_MasterGrundstruktur)" & vbCrLf & _
        "Nein= Einzelmodule (A01…A07)", _
        vbYesNo + vbQuestion, "A-Gruppe starten")
    
    If modus = vbYes Then
        Admin_Starte_A_Master
    Else
        Admin_Starte_A_Einzel
    End If
End Sub

' Nur Masterlauf (A00)
Public Sub Admin_Starte_A_Master()
    M_SafeApp.BeginFastOps True, True, True
    On Error GoTo Clean

    RunList Array( _
        "ErstelleKompletteGrundstruktur", _
        "FinaleGrundkonfiguration", _
        "PruefeGrundstrukturAbhaengigkeiten" _
    ), "Starte A00 (Master-Grundstruktur)", "A00_MasterGrundstruktur"

Clean:
    M_SafeApp.EndFastOps
End Sub

' Nur die Einzelmodule A01…A07 (ohne Master)
Public Sub Admin_Starte_A_Einzel()
    M_SafeApp.BeginFastOps True, True, True
    On Error GoTo Clean

    ' A01
    RunList Array("ErstelleGrundstruktur"), "Starte A01 (Grundblätter)", "A01_Grundblaetter"

    ' A02
    RunList Array("EinrichtenAnleitung"), "Starte A02 (Anleitung)", "A02_Anleitung"

    ' A03
    RunList Array("EinrichtenFeiertage", "AktualisiereFeiertage"), _
        "Starte A03 (Feiertage)", "A03_Feiertage"

    ' A04
    RunList Array("EinrichtenFerien"), "Starte A04 (Ferien)", "A04_Ferien"

    ' A05
    RunList Array("EinrichtenPersonen"), "Starte A05 (Personen)", "A05_Personen"

    ' A06
    RunList Array("EinrichtenBereitschaften", "AktualisiereBereitschaften"), _
        "Starte A06 (Bereitschaften)", "A06_Bereitschaften"

    ' A07
    RunList Array("EinrichtenBAO", "TeamspaltenHinzufuegen", "SortiereBAONachDatum"), _
        "Starte A07 (BAO)", "A07_BAO"

Clean:
    M_SafeApp.EndFastOps
End Sub

'=============================== B-GRUPPE ======================================

Public Sub Admin_Starte_B_Monatssetup()
    M_SafeApp.BeginFastOps True, True, True
    On Error GoTo Clean
    RunList Array("ErstelleMonatsblaetter"), "Starte B01 (Monatsblätter)", "B01_Monatsblaetter"
    RunList Array("B02_DropdownsAlle"), "Starte B02 (Dropdowns)", "B02_Dropdowns"
    RunList Array("B03_TeamstaerkeAlle"), "Starte B03 (Teamstärke)", "B03_Teamstaerke"
Clean:
    M_SafeApp.EndFastOps
End Sub

'=============================== C/D/F =========================================
Public Sub Admin_Starte_C_Formatierung()
    M_SafeApp.BeginFastOps True, True, True
    On Error GoTo Clean
    RunList Array("InitialisiereGrundformatierungFinal"), "Starte C01 (Formatierung)", "C01_Formatierung"
Clean:
    M_SafeApp.EndFastOps
End Sub

Public Sub Admin_Starte_D_BAOIntegration()
    M_SafeApp.BeginFastOps True, True, True
    On Error GoTo Clean
    RunList Array("AktualisiereMonatsblaetterNachBAO"), "Starte D01 (BAO-Integration)", "D01_BAOIntegration"
Clean:
    M_SafeApp.EndFastOps
End Sub

Public Sub Admin_Starte_F_PersonenTools()
    M_SafeApp.BeginFastOps True, True, True
    On Error GoTo Clean
    RunList Array("AktualisierePersonenInMonatsblaettern"), "Starte F01 (Personen-Tools)", "F01_PersonenManager"
Clean:
    M_SafeApp.EndFastOps
End Sub

Public Sub Admin_Starte_ALLES()
    ' Sinnvolle Reihenfolge ohne Doppel-Läufe
    Admin_Starte_A_Master          ' Master baut Grundstruktur komplett
    Admin_Starte_B_Monatssetup     ' Monatsblätter + Dropdowns + Teamstärke
    Admin_Starte_C_Formatierung    ' Formatierung
    Admin_Starte_D_BAOIntegration  ' BAO auf Monatsblätter
    Admin_Starte_F_PersonenTools   ' Personen ggf. erneut auf alle Blätter
End Sub


