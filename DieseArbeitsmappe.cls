VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Excel.Range)
    On Error GoTo Done

    ' 1) Nur reagieren, wenn ein Monatsblatt betroffen ist
    If Not Z_Konfiguration.CFG_IsMonatsblattName(CStr(Sh.Name)) Then
        ' Sonderfall: Anleitung – Jahr/Bundesland löst Neuaufbau aus
        If Sh.Name = Z_Konfiguration.CFG_Sheet_Anleitung Then
            Dim hitYear As Boolean, hitBL As Boolean
            hitYear = Not Intersect(Target, Sh.Range("C2")) Is Nothing
            hitBL = Not Intersect(Target, Sh.Range("C3")) Is Nothing
            If hitYear Or hitBL Then
                M_SafeApp.BeginFastOps True, True, True
                On Error Resume Next
                A03_Feiertage.EinrichtenFeiertage
                A04_Ferien.EinrichtenFerien
                ' BAO/MVL erneut auf alle Monatsblätter anwenden
                D01_BAOIntegration.D01_BAOIntegrationAlle
                ' (Optional) Dropdowns/Teamstärke nachziehen
                B02_Dropdowns.B02_DropdownsAlle
                B03_Teamstaerke.B03_TeamstaerkeAlle
                On Error GoTo 0
                M_SafeApp.EndFastOps
            End If
        End If
        GoTo Done
    End If

    ' 2) Für Monatsblätter: schnelle Re-Formatierung + Tagesstärke aktualisieren
    C01_Formatierung.SchnellAktualisierung Target

Done:
End Sub

