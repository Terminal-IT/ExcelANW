Attribute VB_Name = "M_SafeApp"
'Attribute VB_Name = "M_SafeApp"
Option Explicit

'===============================================================================
' Zweck: Gemeinsame, defensive Helfer für robuste & schnelle Makros.
'  - BeginFastOps / EndFastOps: stapelbares Umschalten von ScreenUpdating, Events, Calc
'  - SafeRun: sicheres Application.Run mit Fehlerfang (true/false)
'  - ListSep: länderspezifischer Listentrenner
'  - Throttle: einfache Drosselung (z. B. bei Blattwechseln)
'
' Excel: 2019 DE kompatibel
'===============================================================================

Private m_stackDepth As Long
Private m_prevCalc As XlCalculation
Private m_prevEvents As Boolean
Private m_prevScreen As Boolean
Private m_prevStatusBar As Variant

Public Sub BeginFastOps(Optional ByVal setManualCalc As Boolean = True, _
                        Optional ByVal hideScreen As Boolean = True, _
                        Optional ByVal disableEvents As Boolean = True)
    On Error Resume Next

    If m_stackDepth = 0 Then
        m_prevCalc = Application.Calculation
        m_prevEvents = Application.EnableEvents
        m_prevScreen = Application.ScreenUpdating
        m_prevStatusBar = Application.StatusBar

        If setManualCalc Then Application.Calculation = xlCalculationManual
        If hideScreen Then Application.ScreenUpdating = False
        If disableEvents Then Application.EnableEvents = False
        Application.StatusBar = False
    End If

    m_stackDepth = m_stackDepth + 1
End Sub

Public Sub EndFastOps()
    On Error Resume Next

    If m_stackDepth > 0 Then
        m_stackDepth = m_stackDepth - 1
        If m_stackDepth = 0 Then
            Application.Calculation = m_prevCalc
            Application.EnableEvents = m_prevEvents
            Application.ScreenUpdating = m_prevScreen
            Application.StatusBar = m_prevStatusBar
        End If
    End If
End Sub

Public Function SafeRun(ByVal procName As String, Optional ByVal arg1 As Variant, _
                        Optional ByVal arg2 As Variant, Optional ByVal arg3 As Variant, _
                        Optional ByVal arg4 As Variant, Optional ByVal arg5 As Variant) As Boolean
    ' Ruft Prozedur per Name auf, ohne Fehler durchzulassen. Gibt Erfolg zurück.
    On Error GoTo fail
    Select Case True
        Case IsMissing(arg1): Application.Run procName
        Case IsMissing(arg2): Application.Run procName, arg1
        Case IsMissing(arg3): Application.Run procName, arg1, arg2
        Case IsMissing(arg4): Application.Run procName, arg1, arg2, arg3
        Case IsMissing(arg5): Application.Run procName, arg1, arg2, arg3, arg4
        Case Else:             Application.Run procName, arg1, arg2, arg3, arg4, arg5
    End Select
    SafeRun = True
    Exit Function
fail:
    SafeRun = False
End Function

Public Function ListSep() As String
    On Error Resume Next
    ListSep = Application.International(xlListSeparator)
    If Len(ListSep) = 0 Then ListSep = ";"
End Function

Public Function Throttle(ByRef lastStamp As Date, ByVal seconds As Long) As Boolean
    ' Gibt True zurück, wenn der letzte Aufruf > seconds her ist (und aktualisiert lastStamp)
    If DateDiff("s", lastStamp, Now) >= seconds Then
        lastStamp = Now
        Throttle = True
    Else
        Throttle = False
    End If
End Function

' === Ergänzung: flexibler Multi-Versuch für Application.Run ===
Public Function SafeRunEx(ByVal baseName As String, _
                          Optional ByVal moduleName As String = vbNullString) As Boolean
    ' Probiert mehrere Varianten:
    '   - baseName
    '   - 'WorkbookName'!baseName
    '   - moduleName.baseName
    '   - 'WorkbookName'!moduleName.baseName
    '   - Variante mit/ohne Unterstrich an typischen Stellen
    Dim wbName As String: wbName = ThisWorkbook.Name
    Dim tries As New Collection, cand As String
    On Error Resume Next
    
    tries.Add baseName
    tries.Add "'" & wbName & "'!" & baseName
    If Len(moduleName) > 0 Then
        tries.Add moduleName & "." & baseName
        tries.Add "'" & wbName & "'!" & moduleName & "." & baseName
    End If
    ' Heuristik: Alternativ mit/ohne Unterstrich nach Präfixen
    If InStr(1, baseName, "_") = 0 Then
        tries.Add Replace(baseName, "Erstelle", "Erstelle_")
        tries.Add Replace(baseName, "Einrichten", "Einrichten_")
        tries.Add Replace(baseName, "Setze", "Setze_")
    Else
        tries.Add Replace(baseName, "_", "")
    End If
    
    Dim i As Long
    For i = 1 To tries.Count
        cand = CStr(tries(i))
        Err.Clear
        Application.Run cand
        If Err.Number = 0 Then SafeRunEx = True: Exit Function
    Next i
End Function

