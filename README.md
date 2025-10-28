# ExcelANW
Excel Projekt - Anwesenheitsliste

## 🧩 Projekt-Startprompt: *Excel Anwesenheit – KI-Entwicklung*

**Ziel:**
Fertigstellung der intelligenten Excel-Anwesenheitsliste mit VBA-Automatisierung.
Die KI arbeitet als **Hauptentwickler**, du als **Projektleiter**, der Zielrichtung, Logik und Designentscheidungen vorgibt.

---

### 🎯 Projektkontext

* Projektname: **Excel Anwesenheit**
* Ziel: Vollautomatische Verwaltung von Anwesenheiten, Aufgaben, Personen und Monatsübersichten mit hoher Performance und klarer Code-Struktur.
* Aktueller Stand:
  Basisfunktionen aus den Modulen

  * Die Module, Klassenmodule und Tabellen liegen hier in diesem Repository
 
Diese bilden den **funktionalen Kern der Anwesenheitslogik**.
Zukünftige Module oder Arbeitsblätter werden als Textdateien (`.cls`, `.bas`, `.frm`) bzw. als Screenshots eingebracht.

---

### 👥 Rollenverteilung

* **Projektleiter:**

  * Liefert Screenshots der Excel-Blätter zur Analyse von Layout, Tabellenstruktur, Farben, Zelllogik.
  * Stellt fachliche Anforderungen, Designrichtlinien und gewünschte Workflows bereit.
  * Entscheidet über Umsetzungsrichtung und Funktionsumfang.
* **KI (Hauptentwickler):**

  * Analysiert und dokumentiert den vorhandenen Code.
  * Entwickelt neue Module, Funktionen und Event-Routinen.
  * Optimiert Performance, Code-Architektur und Fehlerbehandlung.
  * Erstellt Vorschläge für UI/UX-Verbesserungen in Excel (z. B. Dropdowns, Formate, Tabellenstruktur).
  * Stimmt Implementierungen mit dem Projektleiter ab, bevor Code integriert wird.

---

### 📦 Arbeitsweise

1. **Eingabe durch den Projektleiter:**

   * Screenshots von Arbeitsblättern (zur visuellen Analyse).
   * `.cls`, `.bas`, `.frm` Dateien (bestehender oder neuer Code) aus diesem Repo.
   * Beschreibung von Problemen, Änderungswünschen oder Zielen.

2. **Ausgabe durch die KI:**

   * Vollständige, kommentierte Code-Vorschläge.
   * Strukturdiagramme (Beschreibung der Modul-Interaktion).
   * Vorschläge für Refactoring oder Erweiterung.
   * Debugging- und Testanleitungen.

3. **Zielarchitektur:**

   * Klare Trennung zwischen **Datenblättern (Personen, BAO, Monatsblätter)** und **Steuerlogik (DieseArbeitsmappe, Module)**.
   * Einheitliche Ereignissteuerung (`Worksheet_Change`, `Deactivate`, `Activate`).
   * Hohe Wartbarkeit und Erweiterbarkeit.

---

### 🧠 Startanweisung für die KI

> Du bist der Hauptentwickler dieses Projekts.
> Arbeite eng mit mir als Projektleiter zusammen, um die **Excel-Anwesenheitsliste** aus den bereitgestellten Modulen und Screenshots **fertigzustellen, zu optimieren und zu dokumentieren**.
> Nutze die hochgeladenen `.cls`-Dateien als letzten Entwicklungsstand.
> Wenn ich neue Dateien oder Screenshots einbringe, sollst du:
>
> 1. Den Code oder das Layout analysieren,
> 2. den aktuellen Stand beschreiben,
> 3. notwendige Änderungen oder Erweiterungen vorschlagen,
> 4. anschließend den optimierten Code vollständig liefern.
>
> Du darfst eigenständig strukturelle Verbesserungen und saubere VBA-Refactorings vorschlagen, **musst aber jede größere Änderung vorher zur Abstimmung stellen**.
>
> Beginne jedes neue Arbeitspaket mit einer **kurzen technischen Einschätzung** (Was passiert, wo es greift, welche Abhängigkeiten bestehen).

---

Perfekt — das ist eine wichtige Grundlage, um das Projekt „Excel Anwesenheit“ zu verstehen und gezielt weiterzuentwickeln.
Ich erkläre dir zuerst **den Gesamtaufbau der Excel-Anwesenheitsliste** (die gesamte Arbeitsmappe), und danach **den spezifischen Aufbau eines Monatsblatts** wie „Jan“, „Feb“ usw.

Die Beschreibung basiert auf deiner aktuellen Codebasis (aus den `.cls`-Dateien), gängigen Excel-Logik-Patterns und der Funktionalität, die sich daraus ergibt.

---

## 🧩 1. Gesamtaufbau der Excel-Anwesenheitsliste (nicht zwingend final, sondern Momentaufnahme!)

Die Arbeitsmappe ist modular aufgebaut – jede Komponente hat eine klar abgegrenzte Rolle.
Das Ziel: **vollautomatische Synchronisierung** zwischen Personen, Anwesenheiten, Aufgaben und Monatsübersichten.

### 🔧 Strukturübersicht

| Ebene                        | Element                            | Zweck / Inhalt                                                                              |
| ---------------------------- | ---------------------------------- | ------------------------------------------------------------------------------------------- |
| **Arbeitsmappe**             | `DieseArbeitsmappe.cls`            | Zentrale Steuerung: Startlogik, Markierungen, Monatswechsel, Farbaktualisierung             |
| **Tabelle: Personen**        | `Personen.cls`                     | Stammdaten der Mitarbeitenden, synchronisiert mit Monatsblättern                            |
| **Tabelle: BAO**             | `BAO.cls`                          | Übersicht über „Besondere Abwesenheiten / Organisation“ (z. B. Schulung, Urlaub, Krankheit) |
| **Tabelle: Monatsblatt(e)**  | `Jan.cls`, `Feb.cls`, …            | Monatliche Anwesenheitsdarstellung inkl. Dropdowns und Teamstärke-Berechnung                |
| **Hilfstabellen (optional)** | `Feiertage`, `Ferien`, `Anleitung` | Referenzdaten für automatische Markierungen und Farblogik                                   |

---

### 🧠 Funktionslogik im Überblick

1. **Beim Öffnen der Arbeitsmappe (`DieseArbeitsmappe.Workbook_Open`)**

   * Aktualisiert Registerfarben der Monatsblätter (aktueller Monat orange markiert).
   * Markiert den heutigen Tag in allen Monatsblättern farblich.
   * Aktiviert automatisch das aktuelle Monatsblatt.

2. **Personenverwaltung (`Personen.cls`)**

   * Enthält Tabelle `tbl_Personen` mit Stammdaten (z. B. Name, Rolle, Abteilung).
   * Bei Änderungen werden alle Monatsblätter automatisch aktualisiert:

     * Neue Personen werden eingefügt.
     * Gelöschte Personen entfernt.
     * Namensänderungen synchronisiert.
   * Cursor wird bei Aktivierung automatisch auf die nächste freie Zeile gesetzt.

3. **BAO-Tabelle (`BAO.cls`)**

   * Enthält besondere Termine (Beginn/Ende von Abwesenheiten).
   * Jede Änderung löst automatische Sortierung und ggf. Aktualisierung der Monatsblätter aus.
   * Bei Massenänderungen erfolgt ein verzögerter Refresh (zur Performance-Steigerung).

4. **Monatsblätter (z. B. `Jan.cls`)**

   * Enthalten die tägliche Anwesenheitsdarstellung für alle Personen.
   * Dropdowns für Anwesenheitsstatus und Aufgaben werden dynamisch erstellt.
   * Änderungen werden sofort visuell formatiert (Farblogik, Kürzel etc.).
   * Automatische Cursor-Positionierung beim Öffnen.

5. **Hilfstabellen**

   * **`Feiertage`**: Enthält `tbl_Feiertage`, wird genutzt für Wochenend-/Feiertagsfärbung.
   * **`Ferien`**: Enthält `tbl_Ferien`, sorgt für spezielle Markierungen während Ferienzeiten.

---

### 🪄 Automatische Aktualisierungen (zentraler Datenfluss)

```
Personenblatt → Monatsblätter → BAO-Tabelle → Monatsblätter
                              ↘ Feiertage / Ferien (Farblogik)
```

**Beispiel:**

* Eine neue Person wird auf dem Blatt *Personen* angelegt →
  → Alle Monatsblätter bekommen automatisch eine neue Zeile mit Dropdowns.
* Ein Feiertag in „Feiertage“ wird geändert →
  → Tageszellen in allen Monatsblättern werden automatisch farblich angepasst.
* Heute ist der 24.10. →
  → Der entsprechende Tag in allen Monatsblättern wird orange markiert.

---

## 📅 2. Aufbau eines Monatsblattes (z. B. „Jan“)

Das Monatsblatt ist das **Herzstück** der Anwendung.
Hier werden Teamstärken, Aufgaben und Anwesenheitsinformationen verwaltet.

### 🧩 Grundstruktur (vereinfacht)

| Zeile | Inhalt / Funktion                                   | Bereich                   |
| ----- | --------------------------------------------------- | ------------------------- |
| 1–3   | Kopfzeilen, Titel, evtl. Teamname oder Monatsname   | A1–BM3                    |
| 4     | **Wochentage** (z. B. Mo, Di, Mi …)                 | D4–BM4                    |
| 5     | **Datum** (z. B. 01.01., 02.01. …)                  | D5–BM5                    |
| 6–70  | **Personenzeilen** (automatisch aus `tbl_Personen`) | B6–BM70                   |
| 71+   | **Teamstärke-Zeile(n)**                             | Berechnete Summen pro Tag |

---

### 🔽 2.1. Anwesenheits- und Aufgaben-Dropdowns

Jeder Tag besteht aus **zwei Spalten**:

| Spalte            | Inhalt                          | Beispiel                                  |
| ----------------- | ------------------------------- | ----------------------------------------- |
| **Linke Spalte**  | Anwesenheitscode (Dropdown)     | `A`, `U`, `K`, `F`, `D`, `S` ...          |
| **Rechte Spalte** | Aufgabenbeschreibung (Dropdown) | `Büro`, `Home`, `Projekt`, `Schulung` ... |

Das bedeutet:

* Tag 1 = Spalten D (Anwesenheit) und E (Aufgabe)
* Tag 2 = Spalten F (Anwesenheit) und G (Aufgabe)
* usw. bis maximal Spalte **BM** (entspricht 31 Tagen × 2 Spalten).

Die Dropdowns werden **automatisch erstellt**, sobald eine neue Person erkannt wird:

```vba
Call ErstelleDropdownsFuerPersonenzeile(zeile)
```

Diese Funktion ruft:

```vba
anwesenheitListe = GetAnwesenheitsCodes()
aufgabenListe = GetAufgabenCodes()
```

Diese Funktionen liefern vordefinierte Listen (z. B. aus einem Modul `mod_Listen`).

---

### 📊 2.2. Teamstärke und Tagesauswertung

Am unteren Rand des Monatsblatts (z. B. Zeile 72) befindet sich die **Teamstärke-Übersicht**.
Diese zeigt pro Tag:

* Wie viele Personen **anwesend** sind,
* Wie viele **abwesend** sind (Urlaub, Krankheit, etc.),
* Wie viele **im Homeoffice oder auf Projekten** sind.

Berechnung (typisch mit Formeln oder VBA):

* Zähle Einträge mit bestimmten Codes (z. B. `A`, `U`, `K`).
* Summiere diese pro Spaltentag.
* Formatiere Zellen farblich (z. B. grün = vollzählig, rot = wenige anwesend).

---

### 🎨 2.3. Formatierungslogik

1. **Abwechselnde Zellfarben:**
   Gerade und ungerade Tage wechseln zwischen hellgrau und weiß.

2. **Wochenenden / Feiertage:**

   * Hintergrundfarbe dunkelgrau oder rot.
   * Schriftfarbe weiß.
   * Keine Dropdown-Auswahl nötig (optional gesperrt).

3. **Ferien:**

   * Sonderfarbe (z. B. hellgelb).
   * Anwesenheitszellen optional deaktiviert.

4. **Heute-Markierung:**

   * Zeile 5 (Datum) orange hinterlegt.
   * Zeile 4 (Wochentag) normal formatiert oder farblich abgesetzt.

Diese Logik wird über `DieseArbeitsmappe.AktualisiereTodayMarkierung` gesteuert, die durchläuft:

```vba
For Each ws In Worksheets("Jan" to "Dez")
    Call AktualisiereHeuteMarkierungAlleZellen(ws, heute)
Next
```

---

### 🧩 2.4. Ereignissteuerung des Monatsblatts

Das Modul `Jan.cls` steuert das Verhalten des Monatsblatts über Events:

| Event                  | Auslöser                                       | Wirkung                                                                 |
| ---------------------- | ---------------------------------------------- | ----------------------------------------------------------------------- |
| `Worksheet_Change`     | Änderung in Anwesenheits- oder Aufgabenbereich | Schnelle Aktualisierung der Formate (z. B. Farbwechsel bei Codeeingabe) |
| `Worksheet_Change`     | Neue Person in Spalte B                        | Automatische Dropdown-Erstellung für neue Zeile                         |
| `Worksheet_Activate`   | Aktivierung des Blatts                         | Cursor auf sinnvolle Position (z. B. `C4`)                              |
| `Worksheet_Deactivate` | Verlassen des Blatts                           | (optional) Aktualisierung von Statistiken                               |

---

## 🧮 Zusammenfassung (Visuelles Schema)

```
DieseArbeitsmappe
├── Feiertage / Ferien → Farben & Markierungen
├── Personenblatt → Stammdaten → Monatsblätter
├── BAO-Tabelle → Sonderabwesenheiten → Monatsblätter
└── Monatsblätter (Jan–Dez)
    ├── Kopfzeilen (Datum, Wochentag)
    ├── Personenzeilen (Dropdowns)
    ├── Tagescodes + Aufgaben
    ├── Teamstärke (Summen)
    └── Heute-Markierung
```

---
