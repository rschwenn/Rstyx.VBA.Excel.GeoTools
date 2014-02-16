VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CdatDatenpuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2004-2014  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'====================================================================================
'Klasse CdatDatenpuffer
'====================================================================================
'Repräsentiert eine virtuelle Tabelle als Datenpuffer für Import-/Export-
'Operationen sowie für die Bearbeitung der Daten (bestimmte Berechnungen).
'Stellt Eigenschaften und Methoden zur Verwaltung und Bearbeitung zur Verfügung.
'Wird verwendet von: CdatExpim
'
'Datenstruktur:
 ' - Spaltenorientiertes Vorhalten einer Tabelle im Dictionary "oDaten":
 '     Key  = Spaltenname mit Status-Präfix (ohne Einheits-Suffix)
 '     Item = 2d-String-Array (1 to Zeilenanzahl, DPidxWert to DPidxFormel)
 '            mit: Array(i, DPidxWert)   = Wert   der Tabellenzelle
 '                 Array(i, DPidxFormel) = Formel der Tabellenzelle
 '     
 '     Hinweise zu den Daten:
 '     - Das Array darf auch 1d sein (dann ohne Formeln)!
 '     - Nicht vorhandener Wert = "".
 '     - Größenwerte (Länge, Winkel) sind immer normiert, d.h. in der Einheit
 '       mit dem Faktor 1.0 abgelegt (Meter bzw. Gon)
 '     - Spaltennamen des Datenpuffers enthalten deshalb kein Einheiten-Suffix.
 '     - Formeln: Excel-Zellbezüge sind ersetzt durch Addin-interne Bezüge:
 '                - Excel-Spalte => @@<Spaltenname>@@
 '                - Excel-Zeile  => Index i des o.g. Arrays
 '
 ' - Methoden zum Verwalten der Daten (Spalte hinzufügen/löschen).
 ' - Methoden zum Bearbeiten der Daten.
 ' - Eigenschaften als Optionen für Bearbeitungsmethoden.
'
'Hinweise:
 ' - Nach Hinzufügen oder Löschen von Spalten muss Daten_GetSpaltenInfo() aufgerufen werden, das
 '   Informationen zu vorhandenen Spalten im Dictionary oSpaltenStatus ablegt. Bei Verwendung
 '   der Methoden Me.Daten_AddSpalte und Me.Daten_LoeschenSpalte erfolgt dies automatisch.
 ' - Die Methode Me.Daten_Bearbeiten führt im Stapel alle Bearbeitungsmethoden aus,
 '   die durch die entsprechende Option festgelegt sind.
 ' - Me.Daten_SchreibenErsatzSpalte() schreibt Inhalte einer Spalte in die konfigurierte
'    Ersatzspalte, und zwar zellenorientiert, falls die Ersatzspalte bereits existiert.
'
'Historie:
 '08.11.2008  - Beschreiben der Ersatzspalte erfgolgt jetzt zellenorientiert, falls nötig.
 '17.05.2004  - Berechnung von Fehlern und Verbesserungen
 '            - Ist-Überhöhung aus Bemerkung ermitteln
 '            - Transfo' Trassenkoo' => Gleissystem (Zwangspunktreduktion)
 '            - Transfo' Gleissystem => Trassenkoo' (umgekehrte Zwangspunktreduktion)
 '14.04.2008  - Transfo' Trassenkoo' <=> Gleissystem: Basis für Überhöhung konfigurierbar.
 '            - Ist-Überhöhung aus Bemerkung: Schalter für "streng" (u=xxx) eingeführt
'====================================================================================


Option Explicit

'Konstanten
Const strKlasse_Name                    As String = "CdatDatenpuffer"

'Eigenschaften
Dim bln_VorhWerteUeberschreiben         As Boolean
Dim bln_FehlerVerbesserungen            As Boolean
Dim bln_UeberhoehungAusBemerkung        As Boolean
Dim bln_Transfo_Tk2Gls                  As Boolean
Dim bln_UebAusInfoStreng                As Boolean
Dim dbl_TkBasisUeberhoehung             As Double

'Objekte
Dim oDaten                              As Scripting.Dictionary
                                        
Dim oSpaltenStatus                      As Scripting.Dictionary
Dim oSpaltenIst                         As Scripting.Dictionary
Dim oSpaltenSoll                        As Scripting.Dictionary
Dim oSpaltenFehler                      As Scripting.Dictionary
Dim oSpaltenVerbesserung                As Scripting.Dictionary
'



Private Sub Class_Initialize()
  '
  On Error GoTo Fehler
  DebugEcho "CdatDatenpuffer\Class_Initialize() beginnt."
  
  Dim pi        As Double
  Dim rho       As Double
  
  Set oDaten = New Scripting.Dictionary
  
  Set oSpaltenIst = New Scripting.Dictionary
  Set oSpaltenSoll = New Scripting.Dictionary
  Set oSpaltenFehler = New Scripting.Dictionary
  Set oSpaltenVerbesserung = New Scripting.Dictionary
  Set oSpaltenStatus = New Scripting.Dictionary
  
  oSpaltenStatus.Add StatusBez_Ist, oSpaltenIst
  oSpaltenStatus.Add StatusBez_Soll, oSpaltenSoll
  oSpaltenStatus.Add StatusBez_Fehler, oSpaltenFehler
  oSpaltenStatus.Add StatusBez_Verbesserung, oSpaltenVerbesserung
  
  
  'Mathematische Konstanten
  pi = 4 * Atn(1)
  rho = 200 / pi
  
  'Standard-Einstellungen
  Call SetzeStandardEinstellungen
  
  DebugEcho "CdatDatenpuffer\Class_Initialize() beendet."
  Exit Sub
  
Fehler:
  FehlerNachricht "CdatDatenpuffer.Initialize()"
  Set oDaten = Nothing
  Set oSpaltenIst = Nothing
  Set oSpaltenSoll = Nothing
  Set oSpaltenFehler = Nothing
  Set oSpaltenVerbesserung = Nothing
  Set oSpaltenStatus = Nothing
End Sub



Private Sub Class_Terminate()
  Set oDaten = Nothing
  Set oSpaltenIst = Nothing
  Set oSpaltenSoll = Nothing
  Set oSpaltenFehler = Nothing
  Set oSpaltenVerbesserung = Nothing
  Set oSpaltenStatus = Nothing
End Sub




'==== Eigenschaften ===========================================================

Property Get Daten() As Scripting.Dictionary
  'Bietet Zugriff auf die Daten des Datenpuffers.
  '=> Da nur der Inhalt des Dictionary geändert wird, erfolgt auch der
  'schreibende Zugriff auf den Datenpuffer mit dieser Methode!
  Set Daten = oDaten
End Property


Property Get Opt_VorhWerteUeberschreiben() As Boolean
  'Liefert die Option "Vorhandene Werte überschreiben".
  Opt_VorhWerteUeberschreiben = bln_VorhWerteUeberschreiben
End Property

Property Let Opt_VorhWerteUeberschreiben(inp_VorhWerteUeberschreiben As Boolean)
  'Setzt die Option "Vorhandene Werte überschreiben".
  bln_VorhWerteUeberschreiben = inp_VorhWerteUeberschreiben
End Property


Property Get Opt_UebAusInfoStreng() As Boolean
  'Liefert die Einstellung, ob die Angabe der Ist-Überhöhung in der Punktinfo'
  'in der Form u=xxx vorliegen muss oder ob anderenfalls auch die erste Zahl verwendet wird.
  Opt_UebAusInfoStreng = bln_UebAusInfoStreng
End Property

Property Let Opt_UebAusInfoStreng(inp_UebAusInfoStreng As Boolean)
  'Setzt die Einstellung, ob die Angabe der Ist-Überhöhung in der Punktinfo'
  'in der Form u=xxx vorliegen muss oder ob anderenfalls auch die erste Zahl verwendet wird.
  bln_UebAusInfoStreng = inp_UebAusInfoStreng
End Property


Property Get TkBasisUeberhoehung() As Double
  'Liefert "Basis für Überhöhung" in [m] als Vorgabe für Transfo' Trassensystem <=> Gleissystem.
  TkBasisUeberhoehung = dbl_TkBasisUeberhoehung
End Property

Property Let TkBasisUeberhoehung(inp_TkBasisUeberhoehung As Double)
  'Setzt "Basis für Überhöhung" in [m] als Vorgabe für Transfo' Trassensystem <=> Gleissystem.
  dbl_TkBasisUeberhoehung = inp_TkBasisUeberhoehung
End Property



'Die folgenden Eigenschaften legen fest, welche der möglichen Datenbearbeitungsfunktionen
'mit der Methode "Daten_Bearbeiten" im Stapel ausgeführt werden.

Property Get Opt_FehlerVerbesserungen() As Boolean
  'Liefert die Option "Fehler und Verbesserungen berechnen".
  Opt_FehlerVerbesserungen = bln_FehlerVerbesserungen
End Property

Property Let Opt_FehlerVerbesserungen(inp_FehlerVerbesserungen As Boolean)
  'Setzt die Option "Fehler und Verbesserungen berechnen".
  bln_FehlerVerbesserungen = inp_FehlerVerbesserungen
End Property


Property Get Opt_UeberhoehungAusBemerkung() As Boolean
  'Liefert die Option "Überhöhung aus Bemerkung ermitteln".
  Opt_UeberhoehungAusBemerkung = bln_UeberhoehungAusBemerkung
End Property

Property Let Opt_UeberhoehungAusBemerkung(inp_UeberhoehungAusBemerkung As Boolean)
  'Setzt die Option "Überhöhung aus Bemerkung ermitteln".
  bln_UeberhoehungAusBemerkung = inp_UeberhoehungAusBemerkung
End Property


Property Get Opt_Transfo_Tk2Gls() As Boolean
  'Liefert die Option "Zwangspunkt-Reduktion".
  Opt_Transfo_Tk2Gls = bln_Transfo_Tk2Gls
End Property

Property Let Opt_Transfo_Tk2Gls(inp_Transfo_Tk2Gls As Boolean)
  'Setzt die Option "Zwangspunkt-Reduktion".
  bln_Transfo_Tk2Gls = inp_Transfo_Tk2Gls
End Property



'==== Methoden ================================================================

Public Sub SetzeStandardEinstellungen()
  'Alle Einstellungen auf Standardwerte zurücksetzen.
  On Error GoTo Fehler
  
  DebugEcho "CdatDatenpuffer.SetzeStandardEinstellungen(): Einstellungen werden zurückgesetzt."
  
  'Für einige Einstellungen: Konfiguration übernehmen, falls verfügbar
  If (Not ThisWorkbook.Konfig Is Nothing) Then
    bln_VorhWerteUeberschreiben = String2Bool(ThisWorkbook.Konfig.Expim_Ueberschreiben)
    bln_UebAusInfoStreng        = String2Bool(ThisWorkbook.Konfig.UebAusInfoStreng)
    dbl_TkBasisUeberhoehung     = ThisWorkbook.Konfig.TkBasisUeberhoehung  'gibt garantiert einen Double-Wert zurück.
  Else
    DebugEcho "CdatDatenpuffer.SetzeStandardEinstellungen(): Keine Konfiguration verfügbar => Setze hart kodierte Standard-Einstellungen."
    bln_VorhWerteUeberschreiben = Std_VorhWerteUeberschreiben
    bln_UebAusInfoStreng        = Std_UebAusInfoStreng
    dbl_TkBasisUeberhoehung     = Std_TkBasisUeberhoehung
  End If
  
  'Anzuwendende Bearbeitungsmodule
  bln_FehlerVerbesserungen = True
  bln_UeberhoehungAusBemerkung = True
  bln_Transfo_Tk2Gls = True
  
  Exit Sub
  
Fehler:
  FehlerNachricht "CdatDatenpuffer.SetzeStandardEinstellungen()"
End Sub


Public Sub Daten_Bearbeiten()
  'Alle durch Optionen geforderten Manipulation am Datenpuffer ausführen.
  'On Error GoTo Fehler
  
  'SpaltenInfo aktualisieren.
  'Call Daten_GetSpaltenInfo
  
  If (Me.Opt_UeberhoehungAusBemerkung) Then Call Mod_UeberhoehungAusBemerkung
  If (Me.Opt_Transfo_Tk2Gls) Then Call Mod_Transfo_Tk2Gls
  If (Me.Opt_FehlerVerbesserungen) Then Call Mod_FehlerVerbesserung

  Exit Sub

Fehler:
  FehlerNachricht "CdatDatenpuffer.Daten_Bearbeiten()"
End Sub


Public Sub Daten_SchreibenErsatzSpalte(byVal QuellSpalte as String)
  'Die Inhalte der angegebenen Quellspalte werden in die konfigurierte
  'Ersatzspalte geschrieben, falls möglich.
  'Parameter: QuellSpalte ... Name der zu ersetzenden Spalte.
  
  Dim i                  As Long
  Dim QuellFormel        As String
  Dim ZielFormel         As String
  Dim SpalteErsatz       As String
  Dim SpalteOhneStatus   As String
  Dim oSpNameAttr        As Scripting.Dictionary
  
  On Error GoTo Fehler
  DebugEcho "Daten_SchreibenErsatzSpalte(): Ersatzziel einrichten für Spalte '" & QuellSpalte & "'"
  
  If (Not oDaten.Exists(QuellSpalte)) Then
    DebugEcho "Daten_SchreibenErsatzSpalte(): Quellspalte '" & QuellSpalte & "' existiert nicht! => Nichts zu tun."
    
  Else
    'Ersatzspalte laut Konfiguration ermitteln.
    SpalteErsatz = ""
    Set oSpNameAttr = ThisWorkbook.Konfig.SpNameAttr(QuellSpalte)
    SpalteOhneStatus = oSpNameAttr("NamePur")
    If (SpalteOhneStatus <> SpName_unbekannt) Then
      'Für die Spalte existiert eine Konfiguration.
      If (ThisWorkbook.Konfig.SpaltenErsatzZiel(SpalteOhneStatus) <> "") Then
        'Für "Spalte" ist ein Ersatzziel konfiguriert.
        SpalteErsatz = oSpNameAttr("StatusPrefix") & ThisWorkbook.Konfig.SpaltenErsatzZiel(SpalteOhneStatus)
        DebugEcho "Daten_SchreibenErsatzSpalte(): Ersatzspalte ist '" & SpalteErsatz & "'."
      End If
    End If
    
    'Ersatzspalte einrichten.
    If (SpalteErsatz <> "") Then
      'Formeln: Bezug auf die "QuellSpalte" ändern auf "SpalteErsatz"
      Call Formel_AendernSpaltenbezug(QuellSpalte, SpalteErsatz)
      
      'Quelldaten umschaufeln.
      Call Vektor2Datenuffer(oDaten(QuellSpalte), SpalteErsatz, false)
      
      'SpaltenInfo aktualisieren.
      Call Daten_GetSpaltenInfo
    End If
  End If
  
  Set oSpNameAttr = Nothing
  Exit Sub
  
Fehler:
  Set oSpNameAttr = Nothing
  FehlerNachricht "CdatDatenpuffer.Daten_SchreibenErsatzSpalte()"
End Sub


Sub Daten_AddSpalte(ByVal SpName As String, Vektor)
  'Eine Spalte wird dem Datenpuffer hinzugefügt.
  On Error GoTo Fehler
  oDaten.Add SpName, Vektor
  'SpaltenInfo aktualisieren.
  Call Daten_GetSpaltenInfo
  Exit Sub
Fehler:
  FehlerNachricht "CdatDatenpuffer.Daten_AddSpalte()"
End Sub


Sub Daten_LoeschenSpalte(ByVal SpName As String, Vektor)
  'Eine Spalte wird aus dem Datenpuffer entfernt.
  On Error GoTo Fehler
  oDaten.Remove SpName
  'SpaltenInfo aktualisieren.
  Call Daten_GetSpaltenInfo
  Exit Sub
Fehler:
  FehlerNachricht "CdatDatenpuffer.Daten_LoeschenSpalte()"
End Sub


Public Sub Daten_LoeschenAlleSpalten()
  'Datenpuffer wird komplett gelöscht.
  On Error GoTo Fehler
  oDaten.RemoveAll
  'SpaltenInfo aktualisieren.
  Call Daten_GetSpaltenInfo
  Exit Sub
Fehler:
  FehlerNachricht "CdatDatenpuffer.Daten_LoeschenAlleSpalten()"
End Sub


Sub Mod_FehlerVerbesserung()
  'Manipulation am Datenpuffer: Berechnen von Fehlern und Verbesserungen.
  
  'On Error GoTo Fehler
  
  Dim lb                   As Long
  Dim ub                   As Long
  Dim i                    As Long
  Dim AnzD                 As Long
  Dim AnzD2                As Long
  Dim Vektor()             As String
  Dim Vektor2()            As String
  Dim SpName_Ist           As String
  Dim SpName_Fehler        As String
  Dim SpName_Soll          As String
  Dim SpName_Verbesserung  As String
  Dim SpNameOhneStatus     As Variant
  Dim ZielWert             As Variant
  Dim QuellWert            As Variant
  Dim Quellwert2           As Variant
  
  
  Echo vbnewline & "Berechnen von Fehlern und Verbesserungen."
  Echo "  Einstellung: Vorhandene Werte überschreiben = " & bool2String(Me.Opt_VorhWerteUeberschreiben)
  Application.StatusBar = "Berechnen von Fehlern und Verbesserungen."
  
  '1. Vorhandene Fehlerspalten als Verbesserungen eintragen (unter Umkehrung des Vorzeichens).
  If (Not oSpaltenStatus(StatusBez_Fehler) Is Nothing) Then
    For Each SpNameOhneStatus In oSpaltenStatus(StatusBez_Fehler)
      'Alle vorhandenen Fehlerwerte durchgehen.
      SpName_Fehler = ThisWorkbook.Konfig.StatusPrefix(StatusBez_Fehler) & SpNameOhneStatus
      SpName_Verbesserung = ThisWorkbook.Konfig.StatusPrefix(StatusBez_Verbesserung) & SpNameOhneStatus
      lb = LBound(oDaten(SpName_Fehler), 1)
      ub = UBound(oDaten(SpName_Fehler), 1)
      ReDim Vektor(lb To ub)
      AnzD = CountDim(oDaten(SpName_Fehler))
      For i = lb To ub
        If (AnzD = 1) Then
          QuellWert = oDaten(SpName_Fehler)(i)
        Else
          QuellWert = oDaten(SpName_Fehler)(i, DPidxWert)
        End If
        If (IsNumeric(QuellWert)) Then
          ZielWert = QuellWert * (-1)
          Vektor(i) = ZielWert
        End If
      Next
      Call Vektor2Datenuffer(Vektor, SpName_Verbesserung, Me.Opt_VorhWerteUeberschreiben)
      'SpaltenInfo nicht aktualisieren, da neu erzeugte Spalten nicht für Berechnungen dienen sollen.
      'oSpaltenStatus(StatusBez_Verbesserung).add SpNameOhneStatus, SpName_Verbesserung
    Next
  End If
  
  '2. Vorhandene Verbesserungsspalten als Fehler eintragen (unter Umkehrung des Vorzeichens).
  If (Not oSpaltenStatus(StatusBez_Verbesserung) Is Nothing) Then
    For Each SpNameOhneStatus In oSpaltenStatus(StatusBez_Verbesserung)
      'Alle vorhandenen Verbesserungen durchgehen.
      SpName_Fehler = ThisWorkbook.Konfig.StatusPrefix(StatusBez_Fehler) & SpNameOhneStatus
      SpName_Verbesserung = ThisWorkbook.Konfig.StatusPrefix(StatusBez_Verbesserung) & SpNameOhneStatus
      lb = LBound(oDaten(SpName_Verbesserung), 1)
      ub = UBound(oDaten(SpName_Verbesserung), 1)
      ReDim Vektor(lb To ub)
      AnzD = CountDim(oDaten(SpName_Verbesserung))
      For i = lb To ub
        If (AnzD = 1) Then
          QuellWert = oDaten(SpName_Verbesserung)(i)
        Else
          QuellWert = oDaten(SpName_Verbesserung)(i, DPidxWert)
        End If
        If (IsNumeric(QuellWert)) Then
          ZielWert = QuellWert * (-1)
          Vektor(i) = ZielWert
        End If
      Next
      Call Vektor2Datenuffer(Vektor, SpName_Fehler, Me.Opt_VorhWerteUeberschreiben)
    Next
  End If
  
  '3. Verbesserungen und Fehler berechnen aus Soll- und Istwerten.
  If (Not oSpaltenStatus(StatusBez_Soll) Is Nothing) Then
    For Each SpNameOhneStatus In oSpaltenStatus(StatusBez_Soll)
      'Alle vorhandenen Sollwerte durchgehen.
      If (oSpaltenStatus(StatusBez_Ist).Exists(SpNameOhneStatus)) Then
        'Istwert existiert.
        SpName_Ist = ThisWorkbook.Konfig.StatusPrefix(StatusBez_Ist) & SpNameOhneStatus
        SpName_Soll = ThisWorkbook.Konfig.StatusPrefix(StatusBez_Soll) & SpNameOhneStatus
        SpName_Fehler = ThisWorkbook.Konfig.StatusPrefix(StatusBez_Fehler) & SpNameOhneStatus
        SpName_Verbesserung = ThisWorkbook.Konfig.StatusPrefix(StatusBez_Verbesserung) & SpNameOhneStatus
        lb = LBound(oDaten(SpName_Soll), 1)
        ub = UBound(oDaten(SpName_Soll), 1)
        ReDim Vektor(lb To ub)
        ReDim Vektor2(lb To ub)
        AnzD = CountDim(oDaten(SpName_Ist))
        AnzD2 = CountDim(oDaten(SpName_Soll))
        For i = lb To ub
          If (AnzD = 1) Then QuellWert = oDaten(SpName_Ist)(i) Else QuellWert = oDaten(SpName_Ist)(i, DPidxWert)
          If (AnzD2 = 1) Then Quellwert2 = oDaten(SpName_Soll)(i) Else Quellwert2 = oDaten(SpName_Soll)(i, DPidxWert)
          If (IsNumeric(QuellWert) And IsNumeric(Quellwert2)) Then
            'Fehler.
            ZielWert = QuellWert - Quellwert2
            Vektor(i) = ZielWert
            'Verbesserung.
            ZielWert = Quellwert2 - QuellWert
            Vektor2(i) = ZielWert
          End If
        Next
        Call Vektor2Datenuffer(Vektor, SpName_Fehler, Me.Opt_VorhWerteUeberschreiben)
        Call Vektor2Datenuffer(Vektor2, SpName_Verbesserung, Me.Opt_VorhWerteUeberschreiben)
      End If
    Next
  End If
  
  'SpaltenInfo aktualisieren.
  Call Daten_GetSpaltenInfo
  
  Exit Sub
  
Fehler:
  FehlerNachricht "CdatDatenpuffer.Mod_FehlerVerbesserung()"
End Sub


Sub Mod_Transfo_Tk2Gls()
  'Manipulation am Datenpuffer: Berechnen von Koordinaten im gedrehten Gleissystem
  '(Zwangspunkt-Reduktion) ==> Berechnung erfolgt für alle WerteStati.
  
  On Error GoTo Fehler
  
  Dim lb                   As Long
  Dim ub                   As Long
  Dim i                    As Long
  Dim AbstRed()            As String
  Dim dHRed()              As String
  Dim Radius               As Variant
  Dim Abst                 As Variant
  Dim dH                   As Variant
  Dim u                    As Variant
  Dim SpN_Abst             As String
  Dim SpN_dH               As String
  Dim SpN_AbstRed          As String
  Dim SpN_dHRed            As String
  Dim WertStatus           As Variant
  Dim Prefix               As String
  
  Echo vbnewline & "Koordinatentransfo' Trasse => Gleis."
  Echo "  Einstellung: Basis für Überhöhung = " & Format(Me.TkBasisUeberhoehung, "#.000")
  Echo "  Einstellung: Vorhandene Werte überschreiben = " & bool2String(Me.Opt_VorhWerteUeberschreiben)
  Application.StatusBar = "Transformation von Trassenkoo' ins gedrehte Gleissystem (Zwangspunkt-Reduktion)."
  
  If (oDaten.Exists(SpN_S_Tra_Radius) And oDaten.Exists(SpN_S_Tra_u)) Then
    For Each WertStatus In ThisWorkbook.Konfig.StatusPrefix.Keys
      Prefix = ThisWorkbook.Konfig.StatusPrefix(WertStatus)
      SpN_Abst = Prefix & SpN_TK_Q
      SpN_dH = Prefix & SpN_TK_HSOK
      SpN_AbstRed = Prefix & SpN_TK_QG
      SpN_dHRed = Prefix & SpN_TK_HG
      If (oDaten.Exists(SpN_Abst) And oDaten.Exists(SpN_dH)) Then
        lb = LBound(oDaten(SpN_Abst), 1)
        ub = UBound(oDaten(SpN_Abst), 1)
        ReDim AbstRed(lb To ub)
        ReDim dHRed(lb To ub)
        For i = lb To ub
          If (CountDim(oDaten(SpN_S_Tra_Radius)) = 1) Then Radius = oDaten(SpN_S_Tra_Radius)(i) Else Radius = oDaten(SpN_S_Tra_Radius)(i, DPidxWert)
          If (CountDim(oDaten(SpN_S_Tra_u)) = 1) Then u = oDaten(SpN_S_Tra_u)(i) Else u = oDaten(SpN_S_Tra_u)(i, DPidxWert)
          If (CountDim(oDaten(SpN_Abst)) = 1) Then Abst = oDaten(SpN_Abst)(i) Else Abst = oDaten(SpN_Abst)(i, DPidxWert)
          If (CountDim(oDaten(SpN_dH)) = 1) Then dH = oDaten(SpN_dH)(i) Else dH = oDaten(SpN_dH)(i, DPidxWert)
          If (IsNumeric(Radius) And IsNumeric(Abst) And IsNumeric(dH) And IsNumeric(u)) Then
            Call Transfo_Tk2Gls(Radius, u, Me.TkBasisUeberhoehung, Abst, dH, AbstRed(i), dHRed(i))
          End If
        Next
        Call Vektor2Datenuffer(AbstRed, SpN_AbstRed, Me.Opt_VorhWerteUeberschreiben)
        Call Vektor2Datenuffer(dHRed, SpN_dHRed, Me.Opt_VorhWerteUeberschreiben)
      End If
    Next
  End If
  
  'SpaltenInfo aktualisieren.
  Call Daten_GetSpaltenInfo
  
  Exit Sub
  
Fehler:
  FehlerNachricht "CdatDatenpuffer.Mod_Transfo_Tk2Gls()"
End Sub


Sub Mod_Transfo_Gls2Tk()
  '--------------------------------------------------------------------------------------------------------'
  'Modifikation am Datenpuffer: Transformation von Koordinaten des gedrehten Gleissystems
  '                             in normale Trassenkoordinaten (Berechnung erfolgt für alle WerteStati).
  '--------------------------------------------------------------------------------------------------------'
  
  On Error GoTo Fehler
  
  Dim lb                   As Long
  Dim ub                   As Long
  Dim i                    As Long
  Dim Abst()               As String
  Dim dH()                 As String
  Dim Radius               As Variant
  Dim AbstRed              As Variant
  Dim dHRed                As Variant
  Dim u                    As Variant
  Dim SpN_Abst             As String
  Dim SpN_dH               As String
  Dim SpN_AbstRed          As String
  Dim SpN_dHRed            As String
  Dim WertStatus           As Variant
  Dim Prefix               As String
  Dim keineFormelnInRadius As Boolean
  Dim keineFormelnInUeb    As Boolean
  Dim keineFormelnInQG     As Boolean
  Dim keineFormelnInHG     As Boolean
  
  Echo vbnewline & "Koordinatentransfo' Gleis => Trasse."
  Echo "  Einstellung: Basis für Überhöhung = " & Format(Me.TkBasisUeberhoehung, "#.000")
  Echo "  Einstellung: Vorhandene Werte überschreiben = " & bool2String(Me.Opt_VorhWerteUeberschreiben)
  Application.StatusBar = "Transformation von Koo' des gedrehten Gleissystems in normale Trassenkoo'."
  
  If ((oDaten.Exists(SpN_S_Tra_Radius)) And (oDaten.Exists(SpN_S_Tra_u))) Then
    For Each WertStatus In ThisWorkbook.Konfig.StatusPrefix.Keys
      Prefix = ThisWorkbook.Konfig.StatusPrefix(WertStatus)
      SpN_Abst = Prefix & SpN_TK_Q
      SpN_dH = Prefix & SpN_TK_HSOK
      SpN_AbstRed = Prefix & SpN_TK_QG
      SpN_dHRed = Prefix & SpN_TK_HG
      If (oDaten.Exists(SpN_AbstRed) And (oDaten.Exists(SpN_dHRed))) Then
        lb = LBound(oDaten(SpN_AbstRed), 1)
        ub = UBound(oDaten(SpN_AbstRed), 1)
        ReDim Abst(lb To ub)
        ReDim dH(lb To ub)
        
        keineFormelnInRadius = (CountDim(oDaten(SpN_S_Tra_Radius)) = 1)
        keineFormelnInUeb    = (CountDim(oDaten(SpN_S_Tra_u)) = 1)
        keineFormelnInQG     = (CountDim(oDaten(SpN_AbstRed)) = 1)
        keineFormelnInHG     = (CountDim(oDaten(SpN_dHRed)) = 1)
        
        For i = lb To ub
          If keineFormelnInRadius Then Radius = oDaten(SpN_S_Tra_Radius)(i) Else Radius = oDaten(SpN_S_Tra_Radius)(i, DPidxWert)
          If keineFormelnInUeb    Then u = oDaten(SpN_S_Tra_u)(i) Else u = oDaten(SpN_S_Tra_u)(i, DPidxWert)
          If keineFormelnInQG     Then AbstRed = oDaten(SpN_AbstRed)(i) Else AbstRed = oDaten(SpN_AbstRed)(i, DPidxWert)
          If keineFormelnInHG     Then dHRed = oDaten(SpN_dHRed)(i) Else dHRed = oDaten(SpN_dHRed)(i, DPidxWert)
          If (IsNumeric(Radius) And IsNumeric(AbstRed) And IsNumeric(dHRed) And IsNumeric(u)) Then
            Call Transfo_Gls2Tk(Radius, u, Me.TkBasisUeberhoehung, AbstRed, dHRed, Abst(i), dH(i))
          End If
        Next
        Call Vektor2Datenuffer(Abst, SpN_Abst, Me.Opt_VorhWerteUeberschreiben)
        Call Vektor2Datenuffer(dH, SpN_dH, Me.Opt_VorhWerteUeberschreiben)
      End If
    Next
  End If
  
  'SpaltenInfo aktualisieren.
  Call Daten_GetSpaltenInfo
  
  Exit Sub
  
Fehler:
  FehlerNachricht "CdatDatenpuffer.Mod_Transfo_Gls2Tk()"
End Sub


Sub Mod_UeberhoehungAusBemerkung()
  '--------------------------------------------------------------------------------------------------------'
  'Modifikation am Datenpuffer: Ist-Überhöhung anhand der Bemerkung ermitteln und aus Bemerkung entfernen.
  '--------------------------------------------------------------------------------------------------------'
  ' ==> Es wird versucht, der Punktinfo die gemessene Ist-Überhöhung
  '     nach folgenden Regeln zu entnehmen:
  '     1. Falls die Zeichenkette "u= xxx" (an irgendeiner Stelle) enthalten
  '        ist, so wird "xxx" als Ist-Überhöhung angesehen.
  '     2. Falls Variante 1 nicht zum Erfolg führt und in den Einstellungen
  '        nicht nur die strenge Variante erlaubt ist, wird:
  '        => die erste Zahl als Ist-Überhöhung verwendet.
  '--------------------------------------------------------------------------------------------------------'
  
  On Error GoTo Fehler
  
  Dim lb                   As Long
  Dim ub                   As Long
  Dim i                    As Long
  Dim u_ist()              As String
  Dim BemerkNeu()          As String
  Dim BemerkAlt            As String
  Dim ZielWert             As Variant
  Dim QuellWert            As Variant
  Dim keineFormeln         As Boolean
  
  Echo vbnewline & "Ist-Überhöhung anhand der Bemerkung ermitteln (und aus Bemerkung entfernen)."
  Echo "  Einstellung: Überhöhung muss in der Form 'u=xxx' vorliegen = " & bool2String(Me.Opt_UebAusInfoStreng)
  Echo "  Einstellung: Vorhandene Werte überschreiben = " & bool2String(Me.Opt_VorhWerteUeberschreiben)
  Application.StatusBar = "Ist-Überhöhung anhand der Bemerkung ermitteln und aus Bemerkung entfernen."
  
  If (oDaten.Exists(SpN_Pkt_Erl_L)) Then
    lb = LBound(oDaten(SpN_Pkt_Erl_L), 1)
    ub = UBound(oDaten(SpN_Pkt_Erl_L), 1)
    ReDim u_ist(lb To ub)
    ReDim BemerkNeu(lb To ub)
    keineFormeln = (CountDim(oDaten(SpN_Pkt_Erl_L)) = 1)
    
    For i = lb To ub
      'Vorhandene Bemerkung lesen
      If (keineFormeln) Then
        BemerkAlt = oDaten(SpN_Pkt_Erl_L)(i)
      Else
        BemerkAlt = oDaten(SpN_Pkt_Erl_L)(i, DPidxWert)
      End If
      
      'Ist-Überhöhung aus Bemerkung extrahieren
      QuellWert = Ueberhoehung(BemerkAlt, Me.Opt_UebAusInfoStreng)  'u in [mm]
      If (IsNumeric(QuellWert)) Then
        ZielWert = QuellWert / 1000
        u_ist(i) = ZielWert
      End If
      
      '(strenge) Überhöhungsangabe aus Bemerkung entfernen.
      BemerkNeu(i) = substitute("u *= *[-|+]?[0-9]+", "", BemerkAlt, False, False)
    Next
    
    'Neue Werte in den Datenpuffer schreiben.
    Call Vektor2Datenuffer(u_ist, SpN_Tra_u, Me.Opt_VorhWerteUeberschreiben)
    Call Vektor2Datenuffer(BemerkNeu, SpN_Pkt_Erl_L, Me.Opt_VorhWerteUeberschreiben)
  End If
  
  'SpaltenInfo aktualisieren.
  Call Daten_GetSpaltenInfo
  
  Exit Sub
  
Fehler:
  FehlerNachricht "CdatDatenpuffer.Mod_UeberhoehungAusBemerkung()"
End Sub


Sub ZeigeEinstellungen()
  'Zeigt für Kontrollzwecke alle aktiven Einstellungen für Import/Export an.
  'Dim Spalte As Variant
  Dim Message As String
  '
  'Message = "Dialog bestätigt: " & vbTab & CStr(Me.Dialog_OK)
  'Message = Message & vbNewLine
  '
  'Message = Message & vbNewLine & "Quelle Typ: " & vbTab & Me.Quelle_Typ
  'Message = Message & vbNewLine & "Quelle Format: " & vbTab & Me.Quelle_FormatID
  'Message = Message & vbNewLine & "Quelle Datei: " & vbTab & Me.Quelle_AsciiDatei_Name
  'Message = Message & vbNewLine
  '
  'Message = Message & vbNewLine & "Ziel Typ: " & vbTab & vbTab & Me.Ziel_Typ
  'Message = Message & vbNewLine & "Ziel Format: " & vbTab & Me.Ziel_FormatID
  'Message = Message & vbNewLine & "Ziel Datei: " & vbTab & vbTab & Me.Ziel_AsciiDatei_Name
  'Message = Message & vbNewLine & "Ziel Dateimodus: " & vbTab & Me.Ziel_AsciiDatei_Modus
  'Message = Message & vbNewLine
  '
  Message = "Anzeige der Einstellungen ist noch nicht implementiert"
  MsgBox Message
End Sub




'==== interne Routinen ========================================================

Private Sub Daten_LoeschenLeereSpalten()
  'Manipulation am Datenpuffer: Leere Spalten werden gelöscht.
  On Error GoTo Fehler
  Dim Spalte    As Variant
  For Each Spalte In oDaten
    If (isSpalteLeer(oDaten(Spalte))) Then
      oDaten.Remove (Spalte)
    End If
  Next
  'SpaltenInfo nicht aktualisieren, da Aufruf aus Daten_GetSpaltenInfo heraus erfolgt!
  Exit Sub
Fehler:
  FehlerNachricht "CdatDatenpuffer.Daten_LoeschenLeereSpalten()"
End Sub


Private Sub Daten_GetSpaltenInfo()
  'Analysiert alle Spalten des Datenpuffers und speichert die entsprechenden Informationen
  'zwecks Verwendung für Manipulationen am Datenpuffer in folgenden Dictionaries:
  '  oSpaltenStatus("Ist")          ... (Key=Name, Item=Status.Name); Alle Spalten mit WertStatus "Fehler".
  '  oSpaltenStatus("Soll")         ... (Key=Name, Item=Status.Name); Alle Spalten mit WertStatus "Soll".
  '  oSpaltenStatus("Fehler")       ... (Key=Name, Item=Status.Name); Alle Spalten mit WertStatus "Fehler".
  '  oSpaltenStatus("Verbesserung") ... (Key=Name, Item=Status.Name); Alle Spalten mit WertStatus "Verbesserung".
  
  On Error GoTo Fehler
  
  Dim SpNameMitStatus    As Variant
  Dim SpNameOhneStatus   As String
  Dim StatusBez          As String
  Dim oSpNameAttr        As Scripting.Dictionary
  
  'Überflüssige (weil leere) Spalten aus dem Datenpuffer entfernen.
  Call Daten_LoeschenLeereSpalten
  
  'Alle Dictionaries werden zunächst geleert.
  oSpaltenStatus(StatusBez_Ist).RemoveAll
  oSpaltenStatus(StatusBez_Soll).RemoveAll
  oSpaltenStatus(StatusBez_Fehler).RemoveAll
  oSpaltenStatus(StatusBez_Verbesserung).RemoveAll
  
  '"Daten_GetSpaltenInfo" wird während der Initialisierung des Add-Ins verwendet (ThisWorkbook.Konfig.initialize).
  'Zu diesem Zeitpunkt existiert ThisWorkbook.Konfig noch nicht:
  If (Not ThisWorkbook.Konfig Is Nothing) Then
    For Each SpNameMitStatus In oDaten
      Set oSpNameAttr = ThisWorkbook.Konfig.SpNameAttr(SpNameMitStatus)
      StatusBez = oSpNameAttr("StatusBez")
      If (StatusBez <> "") Then
        SpNameOhneStatus = oSpNameAttr("NamePur")
        oSpaltenStatus(StatusBez).Add SpNameOhneStatus, SpNameMitStatus
      End If
    Next
  End If
  
  Set oSpNameAttr = Nothing
  Exit Sub
  
Fehler:
  Set oSpNameAttr = Nothing
  FehlerNachricht "CdatDatenpuffer.Daten_GetSpaltenInfo()"
End Sub


Private Sub Vektor2Datenuffer(Vektor, SpName As String, byVal blnUeberschreiben as Boolean)
  'Schreibt die Werte des Vektorfeldes in die angegebene Spalte des Datenpuffers.
  'Existiert die Spalte im Datenpuffer noch nicht, so wird sie angelegt.
  'Existiert die Spalte im Datenpuffer bereits, so werden vorhandene Werte/Formeln nur dann
  'überschrieben, wenn "blnUeberschreiben" gesetzt ist. Anderenfalls
  'werden nur leere Felder (weder Wert noch Formel) im Datenpuffer beschrieben.
  '
  'Parameter: Vektor ... 1d- oder 2d-Array mit Werten und evtl. Formeln.
  '           SpName ... Name der Zielspalte (ohne Prefix "Spalte.", mit Wertstatus, ohne Einheit)
  '           blnUeberschreiben ... Zielspalte wird zunächst geleert.
  '
  'ACHTUNG: Formeln werden unverändert übernommen - und nur, wenn der Vektor als Ganzes
  '         übernommen wird!
  '         Die Behandlung von Formeln ist demnach auf genau zwei Anwendungsfälle abgestimmt:
  '         1. Es gibt keine ;-)
  '         2. Einrichten einer Ersatzspalte
  
  On Error GoTo Fehler
  
  Dim lb            As Long
  Dim ub            As Long
  Dim ub2           As Long
  Dim i             As Long
  Dim AnzD          As Long
  Dim QuellWert     As Variant
  'Dim QuellFormel   As String
  Dim ZielVektor()  As String
  
  If (Not oDaten.Exists(SpName)) Then
    'Zielspalte existiert noch nicht.
    '=> Quellvektor im Ganzen als Spalte neu hinzufügen.
    oDaten.Add SpName, Vektor
    
  ElseIf (blnUeberschreiben) Then
    'Zielspalte existiert bereits, jeder vorhandene Wert soll aber (auch mit leeren Werten) überschrieben werden.
    '=> Spalte löschen und Quellvektor im Ganzen als Spalte neu hinzufügen.
    oDaten.Remove SpName
    oDaten.Add SpName, Vektor
    
  Else
    'Zielspalte existiert bereits, vorhandene Inhalte sollen erhalten bleiben.
    '=> Werte des Quellvektors nur in leere Zielzellen schreiben.
    '=> Formeln des Quellvektors werden NICHT geschrieben, da sie potentiell
    '   gefährlich sind in dieser "Mischkalkulation".
    AnzD = CountDim(Vektor)
    lb   = LBound(Vektor)
    ub   = UBound(Vektor)
    ub2  = UBound(oDaten(SpName), 1)
    ReDim ZielVektor(lb To ub, DP2lb To DP2ub)
    i = lb
    'Werte und Formeln der Zielspalte auslesen und damit einen temporären
    'Vektor belegen, da direktes Beschreiben des Datenpuffers nicht funktioniert.
    Do
      If (CountDim(oDaten(SpName)) = 1) Then
        ZielVektor(i, DPidxWert)   = oDaten(SpName)(i)
        ZielVektor(i, DPidxFormel) = ""
      Else
        ZielVektor(i, DPidxWert)   = oDaten(SpName)(i, DPidxWert)
        ZielVektor(i, DPidxFormel) = oDaten(SpName)(i, DPidxFormel)
      End If
      i = i + 1
    Loop Until ((i > ub) Or (i > ub2))
    
    'Quelldaten in die ZielSpalte schreiben, falls die jeweilige ZielZelle leer ist.
    i = lb
    Do
      If ((ZielVektor(i, DPidxWert) = "") And (ZielVektor(i, DPidxFormel) = "")) Then
        If (AnzD = 1) Then
          QuellWert = Vektor(i)
          'QuellFormel = ""
        Else
          QuellWert = Vektor(i, DPidxWert)
          'QuellFormel = Vektor(i, DPidxFormel)
        End If
        
        ZielVektor(i, DPidxWert) = QuellWert
        'ZielVektor(i, DPidxFormel) = QuellFormel
      End If
      i = i + 1
    Loop Until ((i > ub) Or (i > ub2))
    
    oDaten.Remove SpName
    oDaten.Add SpName, ZielVektor
  End If
  
  Exit Sub
  
Fehler:
  FehlerNachricht "CdatDatenpuffer.Vektor2Datenuffer()"
End Sub


Private Sub Formel_AendernSpaltenbezug(byVal QuellSpalte as String, byVal ZielSpalte as String)
  'In den Formeln der angegebenen Quellspalte werden Bezüge auf die eigene Spalte
  'geändert in Bezüge auf die Zielspalte.
  'Parameter: QuellSpalte ... Name der zu behandelnden Spalte.
  '           Zielspalte  ... Name der Spalte, die als Ersatzbezug dient.
  
  Dim lb            As Long
  Dim ub            As Long
  Dim i             As Long
  Dim QuellFormel   As String
  Dim ZielFormel    As String
  Dim ZielVektor()  As String
  
  On Error GoTo Fehler
  DebugEcho "Formel_AendernSpaltenbezug(): In Formeln Spaltenbezüge ändern von '" & QuellSpalte & "' auf '" & ZielSpalte & "'"
  
  If (Not oDaten.Exists(QuellSpalte)) Then
    DebugEcho "Formel_AendernSpaltenbezug(): Quellspalte '" & QuellSpalte & "' existiert nicht! => Nichts zu tun."
    
  Else
    if (CountDim(oDaten(QuellSpalte)) > 1) then
      'Spalte könnte Formeln enthalten
      'Änderungen erfolgen zunächst in einem temporären Vektor,
      'da direktes Beschreiben des Datenpuffers nicht funktioniert.
      lb = LBound(oDaten(QuellSpalte))
      ub = UBound(oDaten(QuellSpalte))
      ReDim ZielVektor(lb To ub, DP2lb To DP2ub)
      
      for i = lb to ub
        ZielVektor(i, DPidxWert)   = oDaten(QuellSpalte)(i, DPidxWert)
        ZielVektor(i, DPidxFormel) = ""
        QuellFormel = oDaten(QuellSpalte)(i, DPidxFormel)
        
        if (QuellFormel <> "") then
          ZielFormel = replace(QuellFormel, "@@" & QuellSpalte & "@@", "@@" & ZielSpalte & "@@")
          ZielVektor(i, DPidxFormel) = ZielFormel
        end if
      next
      
      oDaten(QuellSpalte) = ZielVektor
    End If
  End If
  
  Exit Sub
  
Fehler:
  FehlerNachricht "CdatDatenpuffer.Formel_AendernSpaltenbezug()"
End Sub


Private Function isSpalteLeer(Vektor) As Boolean
  'Parameter: Vektor ... ein- oder zweidimensionales Array (Werte oder Werte und Formeln)
  'Rückgabe:  false, wenn nicht alle Werte (und Formeln) des Vektors leer sind bzw. "",
  '           true in allen anderen Fällen (auch bei Fehler).
  On Error GoTo Fehler
  Dim i         As Long
  Dim k         As Long
  
  isSpalteLeer = True
  For i = LBound(Vektor, 1) To UBound(Vektor, 1)
    If (CountDim(Vektor) = 1) Then
      If (Vektor(i) <> "") Then
        isSpalteLeer = False
        Exit For
      End If
    Else
      For k = LBound(Vektor, 2) To UBound(Vektor, 2)
        If (Vektor(i, k) <> "") Then
          isSpalteLeer = False
          Exit For
        End If
      Next
    End If
  Next

  Exit Function
Fehler:
  FehlerNachricht "CdatDatenpuffer.isSpalteLeer()"
End Function


'für jEdit:  :folding=indent::collapseFolds=1:
