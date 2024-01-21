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
' Copyright © 2004-2024  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
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
 '   Informationen zu vorhandenen Spalten im Dictionary oSpaltenStatus ablegt.
 '   Bei Verwendung der Methoden Me.Daten_AddSpalte und Me.Daten_LoeschenSpalte erfolgt dies 
 '   automatisch. Deshalb sollten diese Methoden nicht vielfach hintereinander aufgerufen werden.
 ' - Die Methode Me.Daten_Bearbeiten führt im Stapel alle Bearbeitungsmethoden aus,
 '   die durch die entsprechende Option festgelegt sind.
 ' - Me.Daten_SchreibenErsatzSpalte() schreibt Inhalte einer Spalte in die konfigurierte
'    Ersatzspalte, und zwar zellenorientiert, falls die Ersatzspalte bereits existiert.
'
'Historie:
 '17.05.2004  - Berechnung von Fehlern und Verbesserungen
 '            - Ist-Überhöhung aus Bemerkung ermitteln
 '            - Transfo' Trassenkoo' => Gleissystem (Zwangspunktreduktion)
 '            - Transfo' Gleissystem => Trassenkoo' (umgekehrte Zwangspunktreduktion)
 '14.04.2008  - Transfo' Trassenkoo' <=> Gleissystem: Basis für Überhöhung konfigurierbar.
 '            - Ist-Überhöhung aus Bemerkung: Schalter für "streng" (u=xxx) eingeführt
 '08.11.2008  - Beschreiben der Ersatzspalte erfgolgt jetzt zellenorientiert, falls nötig.
 '27.02.2021  - Neue Modifikation Mod_iTrassenCodeAusBemerkung()
 '02.03.2021  - Transfo' Trassenkoo' <=> Gleissystem arbeitet jetzt mit mehreren
 '              möglichen Sätzen gegebener Werte in dieser Reihenfolge:
 '              1. Soll: Absolute Überhöhung (S.Tra.ua)
 '              2. Soll: Radius und relative Überhöhung (S.Tra.Radius, S.Tra.u)
 '              3. Ist : Absolute Überhöhung (Tra.ua)
 '              4. Ist : Radius und relative Überhöhung (Tra.Radius, Tra.u)
 '10.04.2022  - Parse_iTC() generalüberholt: Interpretation der iTrassen-Codierung
 '              an iGeo und Rstyx.Utilities 2.8 angeglichen.
 '            - Mod_iTrassenCodeAusBemerkung() und Mod_UeberhoehungAusBemerkung()
 '              zusammengefasst zu Mod_InfoTextAuswerten()
 '08.04.2023  - Unterstützung für erweiterte iTrassen-Codierung:
 '              - Neue Punktarten: "Messpunkt" (1,2) und "SOK" (1,2)
 '              - Spurweite als neues Attribut für Gleispunkt
 '              - Alternative Attribut-Schreibweisen "-u*" and "-ueb*" (w/o "i")
 '              - Überhöhung und Spurweite werden auch mit NKS erkannt und geschrieben.
 '            - InfoText: Angaben für Überhöhung werden auch mit NKS erkannt.
 '            - Punktarterkennung jetzt auch von PArtText (KindText)
 '21.01.2024  - Transformationen Tk <-> Gls berücksichtigen nun auch die
 '              am Punkt gespeicherte Überhöhungsbasis (Soll und Ist). Falls die Berechnung
 '              für Istwerte erfolgt, dann hat die Ist-Überhöhungsbasis Vorrang, sonst Soll.
 '21.01.2024  - Daten_GetSpaltenInfo() als Public, damit Daten_AddSpalte() bei 
 '              Massenoperationen vermieden werden kann (siehe CdatExpim).
'====================================================================================


Option Explicit

'Konstanten
Const strKlasse_Name                    As String = "CdatDatenpuffer"

'Pseudo-Konstanten
Dim iTC_Pattern                         As String

'Eigenschaften
Dim bln_VorhWerteUeberschreiben         As Boolean
Dim bln_FehlerVerbesserungen            As Boolean
Dim bln_UeberhoehungAusBemerkung        As Boolean
Dim bln_iTrassenCodeAusBemerkung        As Boolean
Dim bln_Transfo_Tk2Gls                  As Boolean
Dim dbl_TkBasisUeberhoehung             As Double

'Objekte
Dim oDaten                              As Scripting.Dictionary
                                        
Dim oSpaltenStatus                      As Scripting.Dictionary
Dim oSpaltenIst                         As Scripting.Dictionary
Dim oSpaltenSoll                        As Scripting.Dictionary
Dim oSpaltenFehler                      As Scripting.Dictionary
Dim oSpaltenVerbesserung                As Scripting.Dictionary
'


' Supported point kinds.
Public Enum GeoPointKind
    
    ''' <summary> Not a supported point kind. </summary>
    None = 0
    
    ''' <summary> Rails axis. </summary>
    Rails = 1
    
    ''' <summary> Platform edge. </summary>
    Platform = 2
            
    ''' <summary> General fix point. </summary>
    FixPoint = 3
            
    ''' <summary> Height fix point. </summary>
    FixPoint1D = 4
            
    ''' <summary> Position fix point. </summary>
    FixPoint2D = 5
            
    ''' <summary> Position and Height fix point. </summary>
    FixPoint3D = 6
    
    ''' <summary> Rails fix point. </summary>
    RailsFixPoint = 7
        
    ''' <summary> Measure point. </summary>
    MeasurePoint = 8
    
    ''' <summary> Measure point 1. </summary>
    MeasurePoint1 = 9
    
    ''' <summary> Measure point 2. </summary>
    MeasurePoint2 = 10
    
    ''' <summary> Top of rail. </summary>
    RailTop = 11
    
    ''' <summary> Top of rail 1. </summary>
    RailTop1 = 12
    
    ''' <summary> Top of rail 2. </summary>
    RailTop2 = 13
    
End Enum

Dim MarkType2Kind           As Scripting.Dictionary
Dim KindText2Kind           As Scripting.Dictionary
Dim DefaultKindTextShort    As Scripting.Dictionary
Dim DefaultKindTextLong     As Scripting.Dictionary
Dim InfoCantPatterns        As Scripting.Dictionary
Dim InfoKindPatterns        As Scripting.Dictionary


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
  
  ' Mapping:  MarkType => Kind (String, GeoPointKind)
  Set MarkType2Kind = New Scripting.Dictionary
  MarkType2Kind.Add "0" , GeoPointKind.None            ' unvermarkt aus Ril 808 
  MarkType2Kind.Add "1" , GeoPointKind.RailsFixPoint   ' Bolzen am Fahrleitungsmast / Fundament 
  MarkType2Kind.Add "2" , GeoPointKind.RailsFixPoint   ' Stehbolzen, Nagel, Bolzen in Mauer 
  MarkType2Kind.Add "3" , GeoPointKind.RailsFixPoint   ' Bodenvermarkung, Tiefpunkt 
  MarkType2Kind.Add "4" , GeoPointKind.RailsFixPoint   ' Tunnelvermarkung 
  MarkType2Kind.Add "5" , GeoPointKind.FixPoint2D      ' Lochstein Grundlagen-Vermessung 
  MarkType2Kind.Add "6" , GeoPointKind.FixPoint2D      ' Dränrohr 
  MarkType2Kind.Add "7" , GeoPointKind.FixPoint2D      ' Eisenrohr 
  MarkType2Kind.Add "8" , GeoPointKind.None            ' indirekte Vermarkung 
  MarkType2Kind.Add "9" , GeoPointKind.FixPoint1D      ' Höhen- bzw. NivP-Bolzen 
  MarkType2Kind.Add "10", GeoPointKind.FixPoint2D      ' TP - Stein 
  MarkType2Kind.Add "11", GeoPointKind.FixPoint2D      ' TP - Platte 
  MarkType2Kind.Add "12", GeoPointKind.None            ' Grenzstein Kataster-Vermessung 
  MarkType2Kind.Add "13", GeoPointKind.None            ' Grenzmarke 
  MarkType2Kind.Add "14", GeoPointKind.FixPoint2D      ' Nagel 
  MarkType2Kind.Add "15", GeoPointKind.FixPoint2D      ' Kreuz 
  MarkType2Kind.Add "16", GeoPointKind.FixPoint3D      ' Bolzen mit eingravierter ID DB-Referenzsystem
                
  ' Mapping:  KindText => Kind.
  Set KindText2Kind = New Scripting.Dictionary
  KindText2Kind.Add "Gls" , GeoPointKind.Rails
  KindText2Kind.Add "Bstg", GeoPointKind.Platform
  KindText2Kind.Add "DBRF", GeoPointKind.FixPoint3D
  KindText2Kind.Add "GPSC", GeoPointKind.FixPoint3D
  KindText2Kind.Add "GVPV", GeoPointKind.RailsFixPoint
  KindText2Kind.Add "GVP" , GeoPointKind.RailsFixPoint
  KindText2Kind.Add "HBH" , GeoPointKind.FixPoint1D
  KindText2Kind.Add "HFP" , GeoPointKind.FixPoint1D
  KindText2Kind.Add "LFP" , GeoPointKind.FixPoint2D
  KindText2Kind.Add "NXO" , GeoPointKind.FixPoint3D
  KindText2Kind.Add "PPB" , GeoPointKind.FixPoint2D
  KindText2Kind.Add "PS0" , GeoPointKind.FixPoint3D
  KindText2Kind.Add "PS1" , GeoPointKind.FixPoint3D
  KindText2Kind.Add "PS2" , GeoPointKind.FixPoint2D 
  KindText2Kind.Add "PS3" , GeoPointKind.FixPoint1D
  KindText2Kind.Add "PS4" , GeoPointKind.RailsFixPoint
  KindText2Kind.Add "PSx" , GeoPointKind.FixPoint
  KindText2Kind.Add "PP"  , GeoPointKind.FixPoint
  KindText2Kind.Add "AP"  , GeoPointKind.FixPoint
  KindText2Kind.Add "SOK" , GeoPointKind.RailTop
  KindText2Kind.Add "SOK1", GeoPointKind.RailTop1
  KindText2Kind.Add "SOK2", GeoPointKind.RailTop2
  KindText2Kind.Add "MP"  , GeoPointKind.MeasurePoint
  KindText2Kind.Add "MP1" , GeoPointKind.MeasurePoint1
  KindText2Kind.Add "MP2" , GeoPointKind.MeasurePoint2
  
  ' Mapping:  GeoPointKind => Short Text (GeoPointKind, String)
  Set DefaultKindTextShort = New Scripting.Dictionary
  DefaultKindTextShort.Add GeoPointKind.None          , PArt1_None 
  DefaultKindTextShort.Add GeoPointKind.FixPoint      , PArt1_FixPoint
  DefaultKindTextShort.Add GeoPointKind.FixPoint1D    , PArt1_FixPoint1D
  DefaultKindTextShort.Add GeoPointKind.FixPoint2D    , PArt1_FixPoint2D
  DefaultKindTextShort.Add GeoPointKind.FixPoint3D    , PArt1_FixPoint3D
  DefaultKindTextShort.Add GeoPointKind.Platform      , PArt1_Platform
  DefaultKindTextShort.Add GeoPointKind.Rails         , PArt1_Rails
  DefaultKindTextShort.Add GeoPointKind.RailsFixPoint , PArt1_RailsFixPoint
  DefaultKindTextShort.Add GeoPointKind.RailTop       , PArt1_RailTop
  DefaultKindTextShort.Add GeoPointKind.RailTop1      , PArt1_RailTop1
  DefaultKindTextShort.Add GeoPointKind.RailTop2      , PArt1_RailTop2
  DefaultKindTextShort.Add GeoPointKind.MeasurePoint  , PArt1_MeasurePoint
  DefaultKindTextShort.Add GeoPointKind.MeasurePoint1 , PArt1_MeasurePoint1
  DefaultKindTextShort.Add GeoPointKind.MeasurePoint2 , PArt1_MeasurePoint2
  
  ' Mapping:  GeoPointKind => Long Text (GeoPointKind, String)
  Set DefaultKindTextLong = New Scripting.Dictionary
  DefaultKindTextLong.Add GeoPointKind.None          , ""  ' PArt2_None 
  DefaultKindTextLong.Add GeoPointKind.FixPoint      , PArt2_FixPoint
  DefaultKindTextLong.Add GeoPointKind.FixPoint1D    , PArt2_FixPoint1D
  DefaultKindTextLong.Add GeoPointKind.FixPoint2D    , PArt2_FixPoint2D
  DefaultKindTextLong.Add GeoPointKind.FixPoint3D    , PArt2_FixPoint3D
  DefaultKindTextLong.Add GeoPointKind.Platform      , PArt2_Platform
  DefaultKindTextLong.Add GeoPointKind.Rails         , PArt2_Rails
  DefaultKindTextLong.Add GeoPointKind.RailsFixPoint , PArt2_RailsFixPoint
  DefaultKindTextLong.Add GeoPointKind.RailTop       , PArt2_RailTop
  DefaultKindTextLong.Add GeoPointKind.RailTop1      , PArt2_RailTop1
  DefaultKindTextLong.Add GeoPointKind.RailTop2      , PArt2_RailTop2
  DefaultKindTextLong.Add GeoPointKind.MeasurePoint  , PArt2_MeasurePoint
  DefaultKindTextLong.Add GeoPointKind.MeasurePoint1 , PArt2_MeasurePoint1
  DefaultKindTextLong.Add GeoPointKind.MeasurePoint2 , PArt2_MeasurePoint2
  
  ' Patterns for recognizing actual cant from info text.
  ' 08.04.2023: Cant value now may be a decimal instead of only integer.
  Const RegExDecimal_Info As String = " *= *([+-]?([0-9]*[.])?[0-9]+)\s*"
  Set InfoCantPatterns = New Scripting.Dictionary
  InfoCantPatterns.Add "(u)"   & RegExDecimal_Info, "relative or indefinite cant"
  InfoCantPatterns.Add "(ueb)" & RegExDecimal_Info, "absolute cant"
  'InfoCantPatterns.Add "(u) *= *([+-]? *[0-9]+)\s*"  , "relative or indefinite cant"
  'InfoCantPatterns.Add "(ueb) *= *([+-]? *[0-9]+)\s*", "absolute cant"
  
  ' Patterns for recognizing point kind from info text.
  Set InfoKindPatterns  = New Scripting.Dictionary
  InfoKindPatterns.Add "\bGls\b|\bGleis\b"             , GeoPointKind.Rails
  InfoKindPatterns.Add "\bBst\b|\bBstg\b|\bBahnst"     , GeoPointKind.Platform
  InfoKindPatterns.Add "\bBst_|\bBstg_"                , GeoPointKind.Platform
  InfoKindPatterns.Add "\bPS4\b|\bGVP"                 , GeoPointKind.RailsFixPoint
  InfoKindPatterns.Add "\bPS3\b|\bHFP\b|\bHB\b|\bHP\b" , GeoPointKind.FixPoint1D
  InfoKindPatterns.Add "\bPS2\b|\bLFP\b|\bPPB\b"       , GeoPointKind.FixPoint2D
  InfoKindPatterns.Add "\bPS1\b|\bGPSC\b|\bLHFP\b"     , GeoPointKind.FixPoint3D
  InfoKindPatterns.Add "\bPS0\b|\bNXO\b|\bDBRF\b"      , GeoPointKind.FixPoint3D
  InfoKindPatterns.Add "\bPP\b|\bAP\b|\bPSx\b"         , GeoPointKind.FixPoint
  InfoKindPatterns.Add "\bSOK1\b"                      , GeoPointKind.RailTop1
  InfoKindPatterns.Add "\bSOK2\b"                      , GeoPointKind.RailTop2
  InfoKindPatterns.Add "\bSOK\b"                       , GeoPointKind.RailTop
  InfoKindPatterns.Add "\bMesspkt1\b|\bMesspunkt1\b"   , GeoPointKind.MeasurePoint1
  InfoKindPatterns.Add "\bMesspkt2\b|\bMesspunkt2\b"   , GeoPointKind.MeasurePoint2
  InfoKindPatterns.Add "\bMesspkt\b|\bMesspunkt\b"     , GeoPointKind.MeasurePoint
  
  
  ' Pattern to recognize a text as iGeo "iTrassen-Codierung".
   Const RegExDecimal_iTC As String = " *=? *([+-]?([0-9]*[.])?[0-9]+)"
   
   iTC_Pattern = "^ *"
   iTC_Pattern = iTC_Pattern & "(\w)?"            ' PointKindAB
   
   iTC_Pattern = iTC_Pattern & "("                ' Start attributes
   
   iTC_Pattern = iTC_Pattern & "(-b)"             ' Platform
   iTC_Pattern = iTC_Pattern & "|"
   iTC_Pattern = iTC_Pattern & "(-v)([0-9]+)?"    ' Rails fixpoint
   iTC_Pattern = iTC_Pattern & "|"
   iTC_Pattern = iTC_Pattern & "(-f)([0-9]+)?"    ' Fixpoint
   iTC_Pattern = iTC_Pattern & "|"
   iTC_Pattern = iTC_Pattern & "(-s1)"            ' Rail point 1
   iTC_Pattern = iTC_Pattern & "|"
   iTC_Pattern = iTC_Pattern & "(-s2)"            ' Rail point 2
   iTC_Pattern = iTC_Pattern & "|"
   iTC_Pattern = iTC_Pattern & "(-m1)"            ' Measure point 1
   iTC_Pattern = iTC_Pattern & "|"
   iTC_Pattern = iTC_Pattern & "(-m2)"            ' Measure point 2
   iTC_Pattern = iTC_Pattern & "|"
   
   ' Rails may have and are recognized by several attributes in arbitrary order:
   ' - 3 different attributes: relative and absolute cant, track gauge.
   ' - 2 attributes each in 2 different notations:
   '   - Absolute cant is "-iueb*" or "-ueb*"
   '   - Relative cant is "-iu*"   or "-u*"
   iTC_Pattern = iTC_Pattern & "(-i?ueb)" & RegExDecimal_iTC
   iTC_Pattern = iTC_Pattern & "|"
   iTC_Pattern = iTC_Pattern & "(-i?u)"   & RegExDecimal_iTC
   iTC_Pattern = iTC_Pattern & "|"
   iTC_Pattern = iTC_Pattern & "(-sp)"    & RegExDecimal_iTC
   iTC_Pattern = iTC_Pattern & "|"
   
   iTC_Pattern = iTC_Pattern & "(-i?ueb)" & RegExDecimal_iTC & " *(-i?u)"   & RegExDecimal_iTC
   iTC_Pattern = iTC_Pattern & "|"
   iTC_Pattern = iTC_Pattern & "(-i?ueb)" & RegExDecimal_iTC & " *(-sp)"    & RegExDecimal_iTC
   iTC_Pattern = iTC_Pattern & "|"
   iTC_Pattern = iTC_Pattern & "(-i?u)"   & RegExDecimal_iTC & " *(-i?ueb)" & RegExDecimal_iTC
   iTC_Pattern = iTC_Pattern & "|"
   iTC_Pattern = iTC_Pattern & "(-i?u)"   & RegExDecimal_iTC & " *(-sp)"    & RegExDecimal_iTC
   iTC_Pattern = iTC_Pattern & "|"
   iTC_Pattern = iTC_Pattern & "(-sp)"    & RegExDecimal_iTC & " *(-i?u)"   & RegExDecimal_iTC
   iTC_Pattern = iTC_Pattern & "|"
   iTC_Pattern = iTC_Pattern & "(-sp)"    & RegExDecimal_iTC & " *(-i?ueb)" & RegExDecimal_iTC
   iTC_Pattern = iTC_Pattern & "|"
   
   iTC_Pattern = iTC_Pattern & "(-i?ueb)" & RegExDecimal_iTC & " *(-i?u)"   & RegExDecimal_iTC & " *(-sp)"    & RegExDecimal_iTC
   iTC_Pattern = iTC_Pattern & "|"
   iTC_Pattern = iTC_Pattern & "(-i?ueb)" & RegExDecimal_iTC & " *(-sp)"    & RegExDecimal_iTC & " *(-i?u)"   & RegExDecimal_iTC
   iTC_Pattern = iTC_Pattern & "|"
   iTC_Pattern = iTC_Pattern & "(-i?u)"   & RegExDecimal_iTC & " *(-i?ueb)" & RegExDecimal_iTC & " *(-sp)"    & RegExDecimal_iTC
   iTC_Pattern = iTC_Pattern & "|"
   iTC_Pattern = iTC_Pattern & "(-i?u)"   & RegExDecimal_iTC & " *(-sp)"    & RegExDecimal_iTC & " *(-i?ueb)" & RegExDecimal_iTC
   iTC_Pattern = iTC_Pattern & "|"
   iTC_Pattern = iTC_Pattern & "(-sp)"    & RegExDecimal_iTC & " *(-i?ueb)" & RegExDecimal_iTC & " *(-i?u)"   & RegExDecimal_iTC
   iTC_Pattern = iTC_Pattern & "|"
   iTC_Pattern = iTC_Pattern & "(-sp)"    & RegExDecimal_iTC & " *(-i?u)"   & RegExDecimal_iTC & " *(-i?ueb)" & RegExDecimal_iTC
   
   iTC_Pattern = iTC_Pattern & "|"
   iTC_Pattern = iTC_Pattern & "(-i)"             ' Rails.  In fact, in iGeo this is meant generally as *actual position*.
   
   iTC_Pattern = iTC_Pattern & ")?"               ' End attributes
   
   iTC_Pattern = iTC_Pattern & " *((#) ?(.*))?"   ' Comment
   iTC_Pattern = iTC_Pattern & "$"
  '
  
  
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


Property Get Opt_iTrassenCodeAusBemerkung() As Boolean
  'Liefert die Option "iTrassen-Codierung aus Bemerkung extrahieren".
  Opt_iTrassenCodeAusBemerkung = bln_iTrassenCodeAusBemerkung
End Property

Property Let Opt_iTrassenCodeAusBemerkung(inp_iTrassenCodeAusBemerkung As Boolean)
  'Setzt die Option "iTrassen-Codierung aus Bemerkung extrahieren".
  bln_iTrassenCodeAusBemerkung = inp_iTrassenCodeAusBemerkung
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
    dbl_TkBasisUeberhoehung     = ThisWorkbook.Konfig.TkBasisUeberhoehung  'gibt garantiert einen Double-Wert zurück.
  Else
    DebugEcho "CdatDatenpuffer.SetzeStandardEinstellungen(): Keine Konfiguration verfügbar => Setze hart kodierte Standard-Einstellungen."
    bln_VorhWerteUeberschreiben = Std_VorhWerteUeberschreiben
    dbl_TkBasisUeberhoehung     = Std_TkBasisUeberhoehung
  End If
  
  'Anzuwendende Bearbeitungsmodule
  bln_FehlerVerbesserungen = True
  bln_UeberhoehungAusBemerkung = True
  bln_iTrassenCodeAusBemerkung = True
  bln_Transfo_Tk2Gls = True
  
  Exit Sub
  
  Fehler:
  FehlerNachricht "CdatDatenpuffer.SetzeStandardEinstellungen()"
End Sub


Public Sub Daten_Bearbeiten()
  'Alle durch Optionen geforderten Manipulation am Datenpuffer ausführen.
  'On Error GoTo Fehler
  
  If (Me.Opt_iTrassenCodeAusBemerkung Or Me.Opt_UeberhoehungAusBemerkung) Then
      Call Mod_InfoTextAuswerten
  End If
  
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


Public Sub Daten_GetSpaltenInfo()
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
  ErrMessage = "Fehler beim Bestimmen der Fehler und Verbesserungen"
  
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
  '--------------------------------------------------------------------------------------------------------'
  'Manipulation am Datenpuffer: Berechnen von Koordinaten im gedrehten Gleissystem
  '(Zwangspunkt-Reduktion) ==> Berechnung erfolgt für alle WerteStati.
  '--------------------------------------------------------------------------------------------------------'
  
  On Error GoTo Fehler
  
  Dim lb                   As Long
  Dim ub                   As Long
  Dim i                    As Long
  Dim AnzTransfo           As Long
  Dim AbstRed()            As String
  Dim dHRed()              As String
  Dim Radius               As Variant
  Dim Abst                 As Variant
  Dim dH                   As Variant
  Dim u                    As Variant
  Dim ua                   As Variant
  Dim SpN_Abst             As String
  Dim SpN_dH               As String
  Dim SpN_AbstRed          As String
  Dim SpN_dHRed            As String
  Dim WertStatus           As Variant
  Dim Prefix               As String
  Dim ValueExists_R_u      As Boolean
  
  Dim SpExists_Soll_ua     As Boolean
  Dim SpExists_Soll_R_u    As Boolean
  Dim SpExists_Ist_ua      As Boolean
  Dim SpExists_Ist_R_u     As Boolean
  Dim SpExists_success     As Boolean
  
  Dim BasisUeb               As Variant
  Dim SpN_BasisUeb1          As String
  Dim SpN_BasisUeb2          As String
  Dim SpExists_BasisUeb1     As Boolean
  Dim SpExists_BasisUeb2     As Boolean
  
  Dim CountDim_Abst          As Integer
  Dim CountDim_dH            As Integer
  Dim CountDim_BasisUeb1     As Integer
  Dim CountDim_BasisUeb2     As Integer
  
  Echo vbnewline & "Koordinatentransfo' Trasse => Gleis."
  Echo "  Einstellung: Basis für Überhöhung (global)  = " & Format(Me.TkBasisUeberhoehung, "#.000")
  Echo "  Einstellung: Vorhandene Werte überschreiben = " & bool2String(Me.Opt_VorhWerteUeberschreiben)
  Application.StatusBar = "Transformation von Trassenkoo' ins gedrehte Gleissystem (Zwangspunkt-Reduktion)."
  ErrMessage = "Fehler bei Transformation vom Trassen- ins Gleissystem"
  
  ' Vorhandene Spalten für Radius und Überhöhung bestimmen.
  If (oDaten.Exists(SpN_S_Tra_ua)) Then
    SpExists_Soll_ua  = True
  ElseIf (oDaten.Exists(SpN_S_Tra_Radius) And oDaten.Exists(SpN_S_Tra_u)) Then
    SpExists_Soll_R_u = True
  ElseIf (oDaten.Exists(SpN_Tra_ua)) Then
    SpExists_Ist_ua   = True
  ElseIf (oDaten.Exists(SpN_Tra_Radius) And oDaten.Exists(SpN_Tra_u)) Then
    SpExists_Ist_R_u  = True
  End If
  SpExists_success = (SpExists_Soll_ua Or SpExists_Soll_R_u Or SpExists_Ist_ua Or SpExists_Ist_R_u)
  
  If (Not SpExists_success) Then
    Echo "  => Die Tabelle enthält nicht genügend Spalten für diese Berechnung (R+u oder ua / Soll oder Ist)."
  Else
    For Each WertStatus In ThisWorkbook.Konfig.StatusPrefix.Keys
      
      Prefix      = ThisWorkbook.Konfig.StatusPrefix(WertStatus)
      SpN_Abst    = Prefix & SpN_TK_Q
      SpN_dH      = Prefix & SpN_TK_HSOK
      SpN_AbstRed = Prefix & SpN_TK_QG
      SpN_dHRed   = Prefix & SpN_TK_HG
      AnzTransfo  = 0
          
      ' Bestimmung der Überhöhungsbasis: Vorhandene Spalten und Prioritäten.
      If (WertStatus = StatusBez_Ist) Then
        SpN_BasisUeb1 = SpN_Tra_BasisUeb
        SpN_BasisUeb2 = SpN_S_Tra_BasisUeb
      Else
        SpN_BasisUeb1 = SpN_S_Tra_BasisUeb
        SpN_BasisUeb2 = SpN_Tra_BasisUeb
      End If
      SpExists_BasisUeb1 = oDaten.Exists(SpN_BasisUeb1)
      SpExists_BasisUeb2 = oDaten.Exists(SpN_BasisUeb2)
      
      
      If (Not (oDaten.Exists(SpN_Abst) And oDaten.Exists(SpN_dH))) Then
        DebugEcho "  Wertstatus '" & WertStatus & "':  => Die Tabelle enthält nicht genügend Spalten für diese Berechnung (Q + HSOK / Soll oder Ist)."
      Else
        Echo "  Transformation wird durchgeführt für Wertstatus '" & WertStatus & "'."
        lb = LBound(oDaten(SpN_Abst), 1)
        ub = UBound(oDaten(SpN_Abst), 1)
        ReDim AbstRed(lb To ub)
        ReDim dHRed(lb To ub)
        
        CountDim_Abst = CountDim(oDaten(SpN_Abst))
        CountDim_dH   = CountDim(oDaten(SpN_dH))  
        If (SpExists_BasisUeb1) Then CountDim_BasisUeb1 = CountDim(oDaten(SpN_BasisUeb1))
        If (SpExists_BasisUeb2) Then CountDim_BasisUeb2 = CountDim(oDaten(SpN_BasisUeb2))
        
        For i = lb To ub
          ' Zu transformierende Werte bestimmen.
          If (CountDim_Abst = 1) Then Abst = oDaten(SpN_Abst)(i)  Else Abst = oDaten(SpN_Abst)(i, DPidxWert)
          If (CountDim_dH   = 1) Then dH   = oDaten(SpN_dH)(i)    Else dH   = oDaten(SpN_dH)(i, DPidxWert)
          
          ' Überhöhungsbasis bestimmen.
          BasisUeb = Empty
          If (SpExists_BasisUeb1) Then
            If (CountDim_BasisUeb1 = 1) Then BasisUeb = oDaten(SpN_BasisUeb1)(i) Else BasisUeb = oDaten(SpN_BasisUeb1)(i, DPidxWert)
          End If
          If (IsEmpty(BasisUeb) Or (Not IsNumeric(BasisUeb))) Then
            If (SpExists_BasisUeb2) Then
              If (CountDim_BasisUeb2 = 1) Then BasisUeb = oDaten(SpN_BasisUeb2)(i) Else BasisUeb = oDaten(SpN_BasisUeb2)(i, DPidxWert)
            End If
          End If
          If (IsEmpty(BasisUeb) Or (Not IsNumeric(BasisUeb))) Then
            BasisUeb = Me.TkBasisUeberhoehung
          End If
          
          ' Transformation.
          If (IsNumeric(Abst) And IsNumeric(dH)) Then
            
            ' Trassenparameter bestimmen.
            ValueExists_R_u = False
            Radius          = Empty
            u               = Empty
            
            ' 1. Versuch: Soll absolute Überhöhung.
            If (SpExists_Soll_ua) Then
              If (CountDim(oDaten(SpN_S_Tra_ua)) = 1) Then ua = oDaten(SpN_S_Tra_ua)(i)  Else ua = oDaten(SpN_S_Tra_ua)(i, DPidxWert)
              If (IsNumeric(ua)) Then
                Radius = Sgn(ua) * 99
                u      = Abs(ua)
                ValueExists_R_u = True
              End If
            End If
            
            ' 2. Versuch: Soll Radius und relative Überhöhung.
            If (Not ValueExists_R_u) Then
              If (SpExists_Soll_R_u) Then
                If (CountDim(oDaten(SpN_S_Tra_Radius)) = 1) Then Radius = oDaten(SpN_S_Tra_Radius)(i) Else Radius = oDaten(SpN_S_Tra_Radius)(i, DPidxWert)
                If (CountDim(oDaten(SpN_S_Tra_u))      = 1) Then u      = oDaten(SpN_S_Tra_u)(i)      Else u      = oDaten(SpN_S_Tra_u)(i, DPidxWert)
                If (IsNumeric(Radius) And IsNumeric(u)) Then
                  ValueExists_R_u = True
                End If
              End If
            End If
            
            ' 3. Versuch: Ist absolute Überhöhung.
            If (Not ValueExists_R_u) Then
              If (SpExists_Ist_ua) Then
                If (CountDim(oDaten(SpN_Tra_ua)) = 1) Then ua = oDaten(SpN_Tra_ua)(i)  Else ua = oDaten(SpN_Tra_ua)(i, DPidxWert)
                If (IsNumeric(ua)) Then
                  Radius = Sgn(ua) * 99
                  u      = Abs(ua)
                  ValueExists_R_u = True
                End If
              End If
            End If
            
            ' 4. Versuch: Ist Radius und relative Überhöhung.
            If (Not ValueExists_R_u) Then
              If (SpExists_Ist_R_u) Then
                If (CountDim(oDaten(SpN_Tra_Radius)) = 1) Then Radius = oDaten(SpN_Tra_Radius)(i) Else Radius = oDaten(SpN_Tra_Radius)(i, DPidxWert)
                If (CountDim(oDaten(SpN_Tra_u))      = 1) Then u      = oDaten(SpN_Tra_u)(i)      Else u      = oDaten(SpN_Tra_u)(i, DPidxWert)
                If (IsNumeric(Radius) And IsNumeric(u)) Then
                  ValueExists_R_u = True
                End If
              End If
            End If
            
            ' Transformation ausführen.
            If (ValueExists_R_u) Then
              Call Transfo_Tk2Gls(Radius, u, BasisUeb, Abst, dH, AbstRed(i), dHRed(i))
              AnzTransfo = AnzTransfo + 1
            End If
          End If
        Next
        
        Call Vektor2Datenuffer(AbstRed, SpN_AbstRed, Me.Opt_VorhWerteUeberschreiben)
        Call Vektor2Datenuffer(dHRed  , SpN_dHRed  , Me.Opt_VorhWerteUeberschreiben)
        
        Echo "    => " & cstr(AnzTransfo) & " von " & cstr(ub-lb+1) & " Punkte erfolgreich transformiert."
      End If
    Next
    
    ' SpaltenInfo aktualisieren.
    Call Daten_GetSpaltenInfo
  End If
  
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
  
  Dim lb                     As Long
  Dim ub                     As Long
  Dim i                      As Long
  Dim AnzTransfo             As Long
  Dim Abst()                 As String
  Dim dH()                   As String
  Dim Radius                 As Variant
  Dim AbstRed                As Variant
  Dim dHRed                  As Variant
  Dim u                      As Variant
  Dim ua                     As Variant
  Dim SpN_Abst               As String
  Dim SpN_dH                 As String
  Dim SpN_AbstRed            As String
  Dim SpN_dHRed              As String
  Dim WertStatus             As Variant
  Dim Prefix                 As String
  Dim ValueExists_R_u        As Boolean
  
  Dim SpExists_Soll_ua       As Boolean
  Dim SpExists_Soll_R_u      As Boolean
  Dim SpExists_Ist_ua        As Boolean
  Dim SpExists_Ist_R_u       As Boolean
  Dim SpExists_success       As Boolean
  
  Dim BasisUeb               As Variant
  Dim SpN_BasisUeb1          As String
  Dim SpN_BasisUeb2          As String
  Dim SpExists_BasisUeb1     As Boolean
  Dim SpExists_BasisUeb2     As Boolean
  
  Dim CountDim_AbstRed       As Integer
  Dim CountDim_dHRed         As Integer
  Dim CountDim_BasisUeb1     As Integer
  Dim CountDim_BasisUeb2     As Integer
  
  Echo vbnewline & "Koordinatentransfo' Gleis => Trasse."
  Echo "  Einstellung: Basis für Überhöhung (global)  = " & Format(Me.TkBasisUeberhoehung, "#.000")
  Echo "  Einstellung: Vorhandene Werte überschreiben = " & bool2String(Me.Opt_VorhWerteUeberschreiben)
  Application.StatusBar = "Transformation von Koo' des gedrehten Gleissystems in normale Trassenkoo'."
  ErrMessage = "Fehler bei Transformation vom Gleis- ins Trassensystem"
  
  ' Vorhandene Spalten für Radius und Überhöhung bestimmen.
  If (oDaten.Exists(SpN_S_Tra_ua)) Then
    SpExists_Soll_ua   = True
  ElseIf (oDaten.Exists(SpN_S_Tra_Radius) And oDaten.Exists(SpN_S_Tra_u)) Then
    SpExists_Soll_R_u = True
  ElseIf (oDaten.Exists(SpN_Tra_ua)) Then
    SpExists_Ist_ua    = True
  ElseIf (oDaten.Exists(SpN_Tra_Radius) And oDaten.Exists(SpN_Tra_u)) Then
    SpExists_Ist_R_u     = True
  End If
  SpExists_success = (SpExists_Soll_ua Or SpExists_Soll_R_u Or SpExists_Ist_ua Or SpExists_Ist_R_u)
  
  If (Not SpExists_success) Then
    Echo "  => Die Tabelle enthält nicht genügend Spalten für diese Berechnung (R+u oder ua / Soll oder Ist)."
  Else
    For Each WertStatus In ThisWorkbook.Konfig.StatusPrefix.Keys
      
      Prefix      = ThisWorkbook.Konfig.StatusPrefix(WertStatus)
      SpN_Abst    = Prefix & SpN_TK_Q
      SpN_dH      = Prefix & SpN_TK_HSOK
      SpN_AbstRed = Prefix & SpN_TK_QG
      SpN_dHRed   = Prefix & SpN_TK_HG
      AnzTransfo  = 0
          
      ' Bestimmung der Überhöhungsbasis: Vorhandene Spalten und Prioritäten.
      If (WertStatus = StatusBez_Ist) Then
        SpN_BasisUeb1 = SpN_Tra_BasisUeb
        SpN_BasisUeb2 = SpN_S_Tra_BasisUeb
      Else
        SpN_BasisUeb1 = SpN_S_Tra_BasisUeb
        SpN_BasisUeb2 = SpN_Tra_BasisUeb
      End If
      SpExists_BasisUeb1 = oDaten.Exists(SpN_BasisUeb1)
      SpExists_BasisUeb2 = oDaten.Exists(SpN_BasisUeb2)
      
      
      If (Not (oDaten.Exists(SpN_AbstRed) And oDaten.Exists(SpN_dHRed))) Then
        DebugEcho "  Wertstatus '" & WertStatus & "':  => Die Tabelle enthält nicht genügend Spalten für diese Berechnung (QG + HG / Soll oder Ist)."
      Else
        Echo "  Transformation wird durchgeführt für Wertstatus '" & WertStatus & "'."
        lb = LBound(oDaten(SpN_AbstRed), 1)
        ub = UBound(oDaten(SpN_AbstRed), 1)
        ReDim Abst(lb To ub)
        ReDim dH(lb To ub)
        
        CountDim_AbstRed = CountDim(oDaten(SpN_AbstRed))
        CountDim_dHRed   = CountDim(oDaten(SpN_dHRed))  
        If (SpExists_BasisUeb1) Then CountDim_BasisUeb1 = CountDim(oDaten(SpN_BasisUeb1))
        If (SpExists_BasisUeb2) Then CountDim_BasisUeb2 = CountDim(oDaten(SpN_BasisUeb2))
        
        For i = lb To ub
          ' Zu transformierende Werte bestimmen.
          If (CountDim_AbstRed = 1) Then AbstRed = oDaten(SpN_AbstRed)(i) Else AbstRed = oDaten(SpN_AbstRed)(i, DPidxWert)
          If (CountDim_dHRed   = 1) Then dHRed   = oDaten(SpN_dHRed)(i)   Else dHRed   = oDaten(SpN_dHRed)(i, DPidxWert)
          
          ' Überhöhungsbasis bestimmen.
          BasisUeb = Empty
          If (SpExists_BasisUeb1) Then
            If (CountDim_BasisUeb1 = 1) Then BasisUeb = oDaten(SpN_BasisUeb1)(i) Else BasisUeb = oDaten(SpN_BasisUeb1)(i, DPidxWert)
          End If
          If (IsEmpty(BasisUeb) Or (Not IsNumeric(BasisUeb))) Then
            If (SpExists_BasisUeb2) Then
              If (CountDim_BasisUeb2 = 1) Then BasisUeb = oDaten(SpN_BasisUeb2)(i) Else BasisUeb = oDaten(SpN_BasisUeb2)(i, DPidxWert)
            End If
          End If
          If (IsEmpty(BasisUeb) Or (Not IsNumeric(BasisUeb))) Then
            BasisUeb = Me.TkBasisUeberhoehung
          End If
          
          ' Transformation.
          If (IsNumeric(AbstRed) And IsNumeric(dHRed)) Then
            
            ' Trassenparameter bestimmen.
            ValueExists_R_u = False
            Radius          = Empty
            u               = Empty
            
            ' 1. Versuch: Soll absolute Überhöhung.
            If (SpExists_Soll_ua) Then
              If (CountDim(oDaten(SpN_S_Tra_ua)) = 1) Then ua = oDaten(SpN_S_Tra_ua)(i)  Else ua = oDaten(SpN_S_Tra_ua)(i, DPidxWert)
              If (IsNumeric(ua)) Then
                Radius = Sgn(ua) * 99
                u      = Abs(ua)
                ValueExists_R_u = True
              End If
            End If
            
            ' 2. Versuch: Soll Radius und relative Überhöhung.
            If (Not ValueExists_R_u) Then
              If (SpExists_Soll_R_u) Then
                If (CountDim(oDaten(SpN_S_Tra_Radius)) = 1) Then Radius = oDaten(SpN_S_Tra_Radius)(i) Else Radius = oDaten(SpN_S_Tra_Radius)(i, DPidxWert)
                If (CountDim(oDaten(SpN_S_Tra_u))      = 1) Then u      = oDaten(SpN_S_Tra_u)(i)      Else u      = oDaten(SpN_S_Tra_u)(i, DPidxWert)
                If (IsNumeric(Radius) And IsNumeric(u)) Then
                  ValueExists_R_u = True
                End If
              End If
            End If
            
            ' 3. Versuch: Ist absolute Überhöhung.
            If (Not ValueExists_R_u) Then
              If (SpExists_Ist_ua) Then
                If (CountDim(oDaten(SpN_Tra_ua)) = 1) Then ua = oDaten(SpN_Tra_ua)(i)  Else ua = oDaten(SpN_Tra_ua)(i, DPidxWert)
                If (IsNumeric(ua)) Then
                  Radius = Sgn(ua) * 99
                  u      = Abs(ua)
                  ValueExists_R_u = True
                End If
              End If
            End If
            
            ' 4. Versuch: Ist Radius und relative Überhöhung.
            If (Not ValueExists_R_u) Then
              If (SpExists_Ist_R_u) Then
                If (CountDim(oDaten(SpN_Tra_Radius)) = 1) Then Radius = oDaten(SpN_Tra_Radius)(i) Else Radius = oDaten(SpN_Tra_Radius)(i, DPidxWert)
                If (CountDim(oDaten(SpN_Tra_u))      = 1) Then u      = oDaten(SpN_Tra_u)(i)      Else u      = oDaten(SpN_Tra_u)(i, DPidxWert)
                If (IsNumeric(Radius) And IsNumeric(u)) Then
                  ValueExists_R_u = True
                End If
              End If
            End If
            
            ' Transformation ausführen.
            If (ValueExists_R_u) Then
              Call Transfo_Gls2Tk(Radius, u, BasisUeb, AbstRed, dHRed, Abst(i), dH(i))
              AnzTransfo = AnzTransfo + 1
            End If
          End If
        Next
        
        Call Vektor2Datenuffer(Abst, SpN_Abst, Me.Opt_VorhWerteUeberschreiben)
        Call Vektor2Datenuffer(dH  , SpN_dH  , Me.Opt_VorhWerteUeberschreiben)
        
        Echo "    => " & cstr(AnzTransfo) & " von " & cstr(ub-lb+1) & " Punkte erfolgreich transformiert."
      End If
    Next
  End If
  
  ' SpaltenInfo aktualisieren.
  Call Daten_GetSpaltenInfo
  
  Exit Sub
  Fehler:
  FehlerNachricht "CdatDatenpuffer.Mod_Transfo_Gls2Tk()"
End Sub

Sub Mod_InfoTextAuswerten()
  '--------------------------------------------------------------------------------------------------------'
  ' Modifikation am Datenpuffer: Der InfoText (PunktInfo) wird auf diverse Informationen 
  ' untersucht. Gefundene Werte werden aus dem InfoText entfernt und in die
  ' entsprechenden Felder eingetragen:
  '  - Auswertung als iTrassen-Codierung (und Entfernung aus InfoText)
  '  - Überhöhungsangabe der Form "u=xx" (und Entfernung aus InfoText)
  ' Formeln im InfoText gehen verloren!
  '--------------------------------------------------------------------------------------------------------'
  ' - Bei Erfolg werden folgende Teilinformationen extrahiert:
  '   
  '   - Spalte.Pkt.Erl.L     :  Text-Teil der iTrassen-Codierung
  '   - Spalte.Pkt.Art.Bez1  :  Punktart Bezeichnung kurz
  '   - Spalte.Pkt.Art.Bez2  :  Punktart Bezeichnung lang
  '   - Spalte.Pkt.V.ArtKz.2 :  Vermarkungsart für iGeo-Absteckbuch
  '   - Spalte.Pkt.V.ArtKz   :  Vermarkungsart (für Festpunkt oder GVP)
  '   - Spalte.Tra.sp        :  Ist-Spurweite  (für Gleispunkt)
  '   - Spalte.Tra.u         :  Ist-Überhöhung (für Gleispunkt)
  '   - Spalte.Tra.ua        :  Ist-Überhöhung (für Gleispunkt), absolut ("+" => linke Schiene überhöht)
  '    
  ' - Punktarten:  iTrassen-Codierung    Bezeichnung kurz  Bezeichnung lang    Hinweis
  '   -            -i                    Gls               Gleis               ohne Überhöhung
  '   -            -iu*   , -u*          Gls               Gleis               mit  Überhöhung
  '   -            -iueb* , -ueb*        Gls               Gleis               mit  Überhöhung im iGeo-Stil ("+" => rechte Schiene überhöht) 
  '   -            -sp*                  Gls               Gleis               mit  Spurweite
  '   -            -b                    Bstg              Bahnsteig
  '   -            -f*                   PSx               Festpunkt           * = Vermarkungsart
  '   -            -v*                   GVP               GVP                 * = Vermarkungsart
  '   -            -s1 , -s2             SOK1, SOK2        Schiene   1 / 2
  '   -            -m1 , -m2             MP1 , MP2         Messpunkt 1 / 2
  '--------------------------------------------------------------------------------------------------------'
    On Error GoTo Fehler
    
    Dim lb                   As Long
    Dim ub                   As Long
    Dim i                    As Long
    
    Dim u_ist()              As String
    Dim ua_ist()             As String
    Dim sp_ist()             As String
    Dim PArtBez1()           As String
    Dim PArtBez2()           As String
    Dim VArtKz1()            As String
    Dim VArtKz2()            As String
    Dim BemerkNeu()          As String
    
    Dim intPArt              As GeoPointKind
    Dim strPArtText          As String
    Dim strVArtKz1           As String
    Dim strVArtKz2           As String
    Dim strInfoNeu           As String
    
    Dim BemerkAlt            As String
    Dim varUebQuelle         As Variant
    Dim varUebAbsQuelle      As Variant
    Dim varSpurwQuelle       As Variant
    Dim keineFormeln         As Boolean
    Dim FoundITC             As Boolean
    Dim FoundCant            As Boolean
    Dim IsUnknownCant        As Boolean
    Dim ParseCant            As Boolean
    Dim PArtTextExist        As String
    
    Echo vbnewline & "InfoText auswerten (Punktart, Überhöhung usw.)"
    Echo "  Einstellung: iTrassen-Codierung auswerten   = " & bool2String(Me.Opt_iTrassenCodeAusBemerkung)
    Echo "  Einstellung: Überhöhung aus InfoText (u=xx) = " & bool2String(Me.Opt_UeberhoehungAusBemerkung)
    Echo "  Einstellung: Vorhandene Werte überschreiben = " & bool2String(Me.Opt_VorhWerteUeberschreiben)
    Application.StatusBar = "InfoText auswerten."
    ErrMessage = "Fehler bei Auswerten des InfoTextes"
    
    PArtTextExist = oDaten.Exists(SpN_Pkt_Art_Bez1)
    
    If (oDaten.Exists(SpN_Pkt_Erl_L)) Then
        
        lb = LBound(oDaten(SpN_Pkt_Erl_L), 1)
        ub = UBound(oDaten(SpN_Pkt_Erl_L), 1)
        ReDim u_ist(lb To ub)
        ReDim ua_ist(lb To ub)
        ReDim sp_ist(lb To ub)
        ReDim PArtBez1(lb To ub)
        ReDim PArtBez2(lb To ub)
        ReDim VArtKz1(lb To ub)
        ReDim VArtKz2(lb To ub)
        ReDim BemerkNeu(lb To ub)
        keineFormeln = (CountDim(oDaten(SpN_Pkt_Erl_L)) = 1)
        FoundITC = False
        
        For i = lb To ub
            ' Vorbereitungen.
            If (keineFormeln) Then
                BemerkAlt   = oDaten(SpN_Pkt_Erl_L)(i)
                If (PArtTextExist) Then
                    strPArtText = oDaten(SpN_Pkt_Art_Bez1)(i)
                Else
                    strPArtText = ""
                End If
            Else
                BemerkAlt   = oDaten(SpN_Pkt_Erl_L)(i, DPidxWert)
                If (PArtTextExist) Then
                    strPArtText = oDaten(SpN_Pkt_Art_Bez1)(i, DPidxWert)
                Else
                    strPArtText = ""
                End If
            End If
            strInfoNeu      = BemerkAlt
            intPArt         = GeoPointKind.None
            strVArtKz1      = ""
            strVArtKz2      = ""
            varUebQuelle    = ""
            varUebAbsQuelle = ""
            varSpurwQuelle  = ""
            IsUnknownCant   = True
            
            ' iTrassen-Codierung.
            If (Me.Opt_iTrassenCodeAusBemerkung) Then
                FoundITC = Parse_iTC(BemerkAlt, intPArt, strVArtKz1, strVArtKz2, varUebQuelle, varUebAbsQuelle, varSpurwQuelle, strInfoNeu)
                
                ' Flag setzen, wenn Überhöhung erkannt worden ist.
                If (FoundITC) Then
                    If (IsNumeric(varUebQuelle) Or IsNumeric(varUebAbsQuelle)) Then
                      IsUnknownCant = False
                    End If
                Else
                    strInfoNeu = BemerkAlt
                End If
            End If
            
            ' Punktart aus PunktartText ableiten.
            strPArtText = Trim(strPArtText)
            If ((intPArt = GeoPointKind.None) And (strPArtText <> "")) Then
                If (KindText2Kind.Exists(strPArtText)) Then
                    intPArt = KindText2Kind(strPArtText)
                End If
            End If
            
            ' Klassische Überhöhungsangabe aus verbliebenem InfoText.
            ParseCant = ((intPArt = GeoPointKind.None) Or ((intPArt = GeoPointKind.Rails) And IsUnknownCant))
            If (Me.Opt_UeberhoehungAusBemerkung And ParseCant) Then
                FoundCant = ParseInfoForActualCant(strInfoNeu, intPArt, varUebQuelle, varUebAbsQuelle, strInfoNeu)
            End If
            
            ' Parse for point kind.
            If (intPArt = GeoPointKind.None) Then
                intPArt = ParseInfoForPointKind(strInfoNeu)
            End If
            
            ' Felder im Vektor belegen.
            If (FoundITC Or FoundCant) Then
                If (IsNumeric(varUebQuelle)) Then
                  u_ist(i) = varUebQuelle
                End If
                If (IsNumeric(varUebAbsQuelle)) Then
                  ua_ist(i) = varUebAbsQuelle
                End If
                If (IsNumeric(varSpurwQuelle)) Then
                  sp_ist(i) = varSpurwQuelle
                End If
            End If
            BemerkNeu(i) = strInfoNeu
            VArtKz1(i)   = strVArtKz1
            VArtKz2(i)   = strVArtKz2
            PArtBez2(i)  = DefaultKindTextLong(intPArt)
            If (strPArtText <> "") Then
                PArtBez1(i) = strPArtText
            Else
                PArtBez1(i) = DefaultKindTextShort(intPArt)
            End If
        Next
        
        ' Neue Werte in den Datenpuffer schreiben.
        Call Vektor2Datenuffer(u_ist,     SpN_Tra_u,        Me.Opt_VorhWerteUeberschreiben)
        Call Vektor2Datenuffer(ua_ist,    SpN_Tra_ua,       Me.Opt_VorhWerteUeberschreiben)
        Call Vektor2Datenuffer(sp_ist,    SpN_Tra_sp,       Me.Opt_VorhWerteUeberschreiben)
        Call Vektor2Datenuffer(PArtBez1,  SpN_Pkt_Art_Bez1, Me.Opt_VorhWerteUeberschreiben)
        Call Vektor2Datenuffer(PArtBez2,  SpN_Pkt_Art_Bez2, Me.Opt_VorhWerteUeberschreiben)
        Call Vektor2Datenuffer(VArtKz1,   SpN_Pkt_VArt_Kz , False)                          ' Vorh. VArt nicht mit aus InfoText abgeleitetem Wert überschreiben.
        Call Vektor2Datenuffer(VArtKz2,   SpN_Pkt_VArt_Kz2, Me.Opt_VorhWerteUeberschreiben)
        Call Vektor2Datenuffer(BemerkNeu, SpN_Pkt_Erl_L,    Me.Opt_VorhWerteUeberschreiben)
    End If
    
    'SpaltenInfo aktualisieren.
    Call Daten_GetSpaltenInfo
    
    Exit Sub
    
    Fehler:
    FehlerNachricht "CdatDatenpuffer.Mod_InfoTextAuswerten()"
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
  Dim Spalte As Variant
  For Each Spalte In oDaten
    If (isSpalteLeer(oDaten(Spalte))) Then
      oDaten.Remove Spalte
    End If
  Next
  'SpaltenInfo nicht aktualisieren, da Aufruf aus Daten_GetSpaltenInfo heraus erfolgt!
  Exit Sub
  Fehler:
  FehlerNachricht "CdatDatenpuffer.Daten_LoeschenLeereSpalten()"
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
  '         Die Behandlung von Formeln ist demnach auf genau drei Anwendungsfälle abgestimmt:
  '         1. Es gibt keine ;-)
  '         2. blnUeberschreiben = True
  '         3. Einrichten einer Ersatzspalte
  
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
  Dim i As Long
  Dim k As Long
  
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
            
Private Function Parse_iTC(ByVal PointInfoText As String, _
                           ByRef Kind          As GeoPointKind, _
                           ByRef MarkType      As String, _
                           ByRef MarkTypeAB    As String, _
                           ByRef ActualCant    As Variant, _
                           ByRef ActualCantAbs As Variant, _
                           ByRef ActualGauge   As Variant, _
                           ByRef NewInfo       As String _
                          ) As Boolean
  '------------------------------------------------------------------------------    
  ' A given string will be parsed as iGeo "iTrassen-Codierung":
  '
  ' Eingabe:  PointInfoText ... Punktinfo original
  ' Ausgabe:  Kind          ... Punktart
  ' Ausgabe:  MarkType      ... Vermarkungsart (für Festpunkt oder GVP)
  ' Ausgabe:  MarkTypeAB    ... Vermarkungsart für iGeo-Absteckbuch
  ' Ausgabe:  ActualCant    ... Ist-Überhöhung (für Gleispunkt)
  ' Ausgabe:  ActualCantAbs ... Ist-Überhöhung (für Gleispunkt), absolut ("+" => linke Schiene überhöht)
  ' Ausgabe:  ActualGauge   ... Ist-Spurweite  (für Gleispunkt)
  ' Ausgabe:  NewInfo       ... Text-Teil der iTrassen-Codierung
  ' Rückgabe: True, wenn eine iTrassen-Codierung gefunden wurde.
  '  
  ' Erwartete Formate:
  '  -b                        Platform                                 
  '  -i?ueb *=? *<a decimal>  Actual rails with actual absolute cant (sign as iGeo)
  '  -i?u *=? *<a decimal>    Actual rails with actual relative cant
  '  -i                       Actual rails without actual cant
  '  -sp                      Actual rails with actual track gauge
  '  -s1 or -s2               Rail top 1 or 2
  '  -m1 or -m2               Measure point 1 or 2
  '  -v[0-9]+                 Rails fix point [with numeric mark type] 
  '  -f[0-9]+                 Other fix point [with numeric mark type] 
  '  ?-                       If one of the above codes is there, it may be preceeded 
  '                           by one word character which will be treated as MarkTypeAB.
  '------------------------------------------------------------------------------
    Dim FoundITC          As Boolean
    Dim oMatches          As Object
    Dim Key               As String
    Dim i                 As Integer
    Dim iTC_MarkType      As String
    Dim iTC_Kind          As GeoPointKind
    Dim iTC_ActualCant    As Variant
    Dim iTC_ActualCantAbs As Variant
    Dim iTC_ActualGauge  As Variant
    Dim KindFromMark      As GeoPointKind
    
    FoundITC      = False
    Kind          = GeoPointKind.None
    MarkType      = ""
    MarkTypeAB    = ""
    ActualCantAbs = ""
    ActualCant    = ""
    ActualGauge   = ""
    NewInfo       = ""
    
    If (Trim(PointInfoText) <> "") Then
        
        ThisWorkbook.RegExp.IgnoreCase = True
        ThisWorkbook.RegExp.Global     = False
        ThisWorkbook.RegExp.Pattern    = iTC_Pattern
        Set oMatches = ThisWorkbook.RegExp.Execute(PointInfoText)
        
        If (oMatches.Count > 0) Then
            
            FoundITC          = True
            iTC_Kind          = GeoPointKind.None
            iTC_MarkType      = ""
            iTC_ActualCantAbs = Null
            iTC_ActualCant    = Null
            iTC_ActualGauge   = Null
            
            ' Recognize point kinds with attached info.
            For i = 2 To 110
                Key = oMatches(0).SubMatches(i)
                Select Case Key
                    Case "-b":             iTC_Kind = GeoPointKind.Platform
                    Case "-v":             iTC_Kind = GeoPointKind.RailsFixPoint : iTC_MarkType      = oMatches(0).SubMatches(i + 1)
                    Case "-f":             iTC_Kind = GeoPointKind.FixPoint      : iTC_MarkType      = oMatches(0).SubMatches(i + 1)
                    Case "-iueb", "-ueb":  iTC_Kind = GeoPointKind.Rails         : iTC_ActualCantAbs = -1 * CDbl(Replace(oMatches(0).SubMatches(i + 1), " ", "")) / 1000
                    Case "-iu"  , "-u":    iTC_Kind = GeoPointKind.Rails         : iTC_ActualCant    =      CDbl(Replace(oMatches(0).SubMatches(i + 1), " ", "")) / 1000
                    Case "-sp":            iTC_Kind = GeoPointKind.Rails         : iTC_ActualGauge   =      CDbl(Replace(oMatches(0).SubMatches(i + 1), " ", "")) / 1000
                    Case "-i":             iTC_Kind = GeoPointKind.Rails
                    Case "-s1":            iTC_Kind = GeoPointKind.RailTop1
                    Case "-s2":            iTC_Kind = GeoPointKind.RailTop2
                    Case "-m1":            iTC_Kind = GeoPointKind.MeasurePoint1
                    Case "-m2":            iTC_Kind = GeoPointKind.MeasurePoint2
                End Select
                ' Exit For -> not if rails, because "-iueb", "-iu" and "-sp" may be there.
                If ((iTC_Kind <> GeoPointKind.None) And (iTC_Kind <> GeoPointKind.Rails)) Then
                    Exit For
                End If
            Next
                            
            ' Fixpoint: Try to state kind more precisely from mark type.
            If ((iTC_Kind = GeoPointKind.FixPoint) And iTC_MarkType <> "") Then
                If (MarkType2Kind.Exists(iTC_MarkType)) Then
                    KindFromMark = MarkType2Kind(iTC_MarkType)
                    Select Case KindFromMark
                        Case GeoPointKind.FixPoint1D, GeoPointKind.FixPoint2D, GeoPointKind.FixPoint3D :  iTC_Kind = KindFromMark
                    End Select
                End If
            End If
            
            ' Try to apply recognized point kinds with related info.
            If (iTC_Kind <> GeoPointKind.None) Then
                
                Kind     = iTC_Kind
                MarkType = iTC_MarkType
                
                If (Not IsNull(iTC_ActualCantAbs)) Then
                    ActualCantAbs = iTC_ActualCantAbs
                End If
                
                If (Not IsNull(iTC_ActualCant)) Then
                    ActualCant = iTC_ActualCant
                End If
                
                If (Not IsNull(iTC_ActualGauge)) Then
                    ActualGauge = iTC_ActualGauge
                End If
            End If
            
            ' Special mark type.
            MarkTypeAB = oMatches(0).SubMatches(0)
            
            ' Point info text.
            NewInfo = Trim(oMatches(0).SubMatches(113))
        End If
    End If
    Parse_iTC = FoundITC
    Exit Function
  Fehler:
  FehlerNachricht "CdatDatenpuffer.Parse_iTC()"
End Function

Function ParseInfoForActualCant(ByVal PointInfoText As String, _
                                ByRef Kind          As GeoPointKind, _
                                ByRef ActualCant    As Variant, _
                                ByRef ActualCantAbs As Variant, _
                                ByRef NewInfo       As String _
                               ) As Boolean
 '------------------------------------------------------------------------------    
 ' A given string will be parsed for actual cant:
 '
 ' Eingabe:  PointInfoText ... Punktinfo original
 ' Ausgabe:  Kind          ... Punktart. Wird nur auf GeoPointKind.Rails geändert, falls Überhöhung gefunden wurde.
 ' Ausgabe:  ActualCant    ... Null oder Ist-Überhöhung (für Gleispunkt)
 ' Ausgabe:  ActualCantAbs ... Null oder Ist-Überhöhung (für Gleispunkt), absolut ("+" => linke Schiene überhöht)
 ' Ausgabe:  NewInfo       ... PointInfoText ohne gefundene Überhöhungsangabe
 ' Rückgabe: True, wenn eine Überhöhungsangabe gefunden wurde.
 '  
 ' Erwartete Formate:
 '  ueb *= *([+-]? *[0-9]+)  Actual absolute cant (sign as iGeo)
 '  u *= *([+-]? *[0-9]+)    Actual relative cant
 '------------------------------------------------------------------------------    
    Dim FoundCant As Boolean
    Dim Pattern   As Variant
    Dim oMatches  As Object
    
    FoundCant     = False
    ActualCant    = ""
    ActualCantAbs = ""
    
    If (Trim(PointInfoText) <> "") Then
        
        ' Test EVERY pattern against PointInfoText, since both "ueb=" and "u=" may be there.
        For Each Pattern In InfoCantPatterns
            
            ThisWorkbook.RegExp.IgnoreCase = True
            ThisWorkbook.RegExp.Global     = False
            ThisWorkbook.RegExp.Pattern    = Pattern
            Set oMatches = ThisWorkbook.RegExp.Execute(PointInfoText)
            
            If (oMatches.Count > 0) Then
                
                FoundCant = True
                Kind      = GeoPointKind.Rails
                
                ' Get actual cant and remove cant pattern from PointInfoText.
                Select Case LCase(oMatches(0).SubMatches(0))
                    Case "ueb":  ActualCantAbs = -1 * CDbl(Replace(oMatches(0).SubMatches(1), " ", "")) / 1000
                    Case "u":    ActualCant    =      CDbl(Replace(oMatches(0).SubMatches(1), " ", "")) / 1000
                End Select
                'PointInfoText = substitute(oMatches(0).SubMatches(0), "", PointInfoText, False, False)
                PointInfoText = Left(PointInfoText, oMatches(0).FirstIndex) & Mid(PointInfoText, oMatches(0).FirstIndex + oMatches(0).Length + 1)
            End If
        Next
    End If
    NewInfo = PointInfoText
    
    ParseInfoForActualCant = FoundCant
    Exit Function
  Fehler:
  FehlerNachricht "CdatDatenpuffer.ParseInfoForActualCant()"
End Function

Function ParseInfoForPointKind(ByVal PointInfoText As String) As GeoPointKind
 '------------------------------------------------------------------------------    
 ' Guesses the point kind from PointInfoText (somewhat heuristic).
 '
 ' Eingabe:  PointInfoText ... Punktinfo
 ' Rückgabe: Kind          ... erkannte Punktart, sonst GeoPointKind.None
 '  
 ' The following patterns somewhere in PointInfoText lead to guessing the point kind: 
 '  Gls, Gleis        => Actual rails (without actual cant)
 '  Bstg, Bst, Bahnst => Platform
 '  PS4, GVP          => Rails fix point
 '  PS3, HFP, HB, HP  => Fix point 1D
 '  PS2, LFP, PPB     => Fix point 2D
 '  PS1, GPSC, LHFP   => Fix point 3D
 '  PS0, NXO, DBRF    => Fix point 3D
 '  PP, AP, PSx       => General fix point
 '------------------------------------------------------------------------------    
    Dim Pattern   As Variant
    Dim oMatches  As Object
    Dim RetValue  As GeoPointKind
    
    RetValue = GeoPointKind.None
    
    If (Trim(PointInfoText) <> "") Then
            
        For Each Pattern In InfoKindPatterns
            
            ThisWorkbook.RegExp.IgnoreCase = True
            ThisWorkbook.RegExp.Global     = False
            ThisWorkbook.RegExp.Pattern    = Pattern
            Set oMatches = ThisWorkbook.RegExp.Execute(PointInfoText)
            
            If (oMatches.Count > 0) Then
                RetValue = InfoKindPatterns(Pattern)
                Exit For
            End If
        Next
    End If
    
    ParseInfoForPointKind = RetValue
End Function


'für jEdit:  :folding=indent::collapseFolds=1:tabSize=4::indentSize=4:
