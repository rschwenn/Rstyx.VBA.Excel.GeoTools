VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CimpTrassenkoo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2003 - 2023  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'==================================================================================================
'Klasse CimpTrassenkoo
'==================================================================================================
'Import von Trassenkoordinaten bzw. Umformungen
'
'Eingabedatei: Textdatei, die mehrere Umformungen von Gleisachspunkten enthalten kann.
'
'  ==> unterstützte Formate: 
'      - Verm.esn:  - Programmversionen: 6.62 (DOS), 8.40 (Windows)
'                   - klassisch (THW) und 3-Linien (D3L)
'                   - 1-zeilig und 2-zeilig
'                   - Ab Spalte 83/84 kann eine Punktinfo stehen (max. 13-stellig).
' 
'      - iGeo:      - Programmversion: Trassenbibliothek V241
'                   - Ausgabeformate A0 und A1
'                   - Am Zeilenende der A1 kann nach einem "#" ein Kommentar stehen,
'                     der als Punktinfo gewertet wird
' 
'      HINWEISE:    - Die Eingabedatei darf mehrere Ausgaben enthalten, ABER:
'                     => ENTWEDER Verm.esn ODER iGeo!
'                   - Die dazugehörigen Kommentarblöcke müssen komplett erhalten sein!
'==================================================================================================


Option Explicit

'Modulkonstanten
Const strKlasse_Name                    As String = "CimpTrassenkoo"

Const strTitel                          As String = "iGeo-A0/A1 *ODER* VermEsn-Umformungen"
Const strKategorien                     As String = "tabTK_iGeo;Punktbeschreibung;Koordinaten (Trasse);Trassenparameter;Koordinaten (rechtwinklig)"

'Const strQuelle_AsciiDatei_DialogFilter As String = "Trassenkoo' Verm.esn (*.txt),*.txt, Trassenkoo' iGeo A0 (*.a0),*.a0, Trassenkoo' iGeo A1 (*.a1),*.a1, Alle Dateien (*.*),*.*"
Const strQuelle_AsciiDatei_DialogFilter As String = "Trassenkoo' iGeo A0 (*.a0),*.a0, Trassenkoo' iGeo A1 (*.a1),*.a1, Trassenkoo' Verm.esn (*.txt),*.txt, Alle Dateien (*.*),*.*"

Const strZiel_Typ                       As String = io_Typ_Puffer
Const strZiel_XLTabKlasse               As String = ""
Const blnZiel_XLMappeSpeichern          As Boolean = False
Const blnFehlerKonsolidieren            As Boolean = True

'Feldnamen (ZellNamen in Tabelle = Keys des Dictionary)
Const FN_Statistik_Programm             As String = "x.Statistik.Programm"
Const FN_Statistik_Importdatei          As String = "x.Statistik.Importdatei"

'Datenbereich
Dim aktTabZeile                         As Long

'Eigenschaften
Dim strQuelle_AsciiDatei_Name           As String
Dim strMeldung_Titel                    As String
Dim strMeldung_Inhalt                   As String
Dim blnAktionErfolgreich                As Boolean
Dim lngFehlerniveau                     As Long

'math. Konstanten
Dim pi                   As Double
Dim rho                  As Double

Dim ZeilenNr             As Long       'Nr. der zuletzt eingelesenen Zeile der Eingabedatei
Dim NF                   As Long       'Anzahl Felder der aktuellen Zeile
Dim Zeile                As String     'gesamte aktuelle Zeile, die zuletzt eingelesen wurde
Dim TrimZeile            As String     '... getrimmt
Dim Feld()               As String     'einzelne Felder der aktuellen Zeile (1. Index = 1)

'Arrays für Adelt und iGeo/iTrassePC
Dim N_KmLinie()          As String
Dim N_Trasse()           As String
Dim N_Gra()              As String
Dim N_Acfg()             As String
Dim N_UebBand()          As String
Dim N_RegPrf()           As String
Dim N_TunPrf()           As String
Dim N_PrfPkt()           As String
Dim N_GlsPrf()           As String
Dim N_DGM()              As String
Dim Abst()               As String
Dim Bemerkung()          As String
Dim PktKomment()         As String
Dim GradHoehe()          As String
Dim PktNr()              As String
Dim Radius()             As String
Dim Ri()                 As String
Dim dH()                 As String
Dim stat()               As String
Dim Km()                 As String
Dim u()                  As String
Dim X()                  As String
Dim Y()                  As String
Dim Z()                  As String

'Arrays nur für Adelt-Umformung
Dim AbstKm()             As String
'Dim u_ist()              As String

'Arrays nur für TK iGeo/iTrassePC
'Dim AbstRed()            As String
'Dim dHRed()              As String
Dim TK_H()               As String
Dim TK_V()               As String
Dim TK_R()               As String
Dim TK_L()               As String
Dim TK_RG()              As String
Dim TK_LG()              As String
Dim TK_Tm()              As String
Dim TK_QT()              As String
Dim TK_QGT()             As String
Dim TK_HGT()             As String
Dim TK_QGS()             As String
Dim TK_HGS()             As String
Dim TK_KmStatus()        As String
Dim TK_KmText()          As String
Dim DGM_HDGM()           As String
Dim DGM_ZDGM()           As String
Dim Tra_UebLi()          As String
Dim Tra_UebRe()          As String
Dim Tra_Heb()            As String
Dim Tra_G()              As String
Dim Tra_ZLGS()           As String
Dim Tra_RaLGS()          As String
Dim Tra_AbLGS()          As String
Dim Tra_MiniR()          As String
Dim Tra_MiniUeb()        As String
Dim Tra_XA()             As String
Dim Tra_YA()             As String
Dim Tra_ZA()             As String
Dim Tra_STB1()           As String
Dim Tra_STB2()           As String
Dim Tra_STB3()           As String
Dim Tra_STB4()           As String
Dim Tra_STB5()           As String
Dim Tra_STB6()           As String
Dim Tra_STB7()           As String
Dim Tra_STB8()           As String
Dim Tra_STB9()           As String

Dim oDictKopfDat         As Scripting.Dictionary     'Dictionary, das alle Kopfdaten der Importdatei enthält

'Spaltenbeschreibungen:  Key=Spaltenname ohne Prefix und Einheit
Dim oQuellEinheiten      As Scripting.Dictionary       'Item=Name der Einheit
Dim oQuelldatenpuffer    As Scripting.Dictionary       'Item=Datenvektor
'



Private Sub Class_Initialize()
  '
  On Error GoTo Fehler

  DebugEcho "CimpTrassenkoo.Class_Initialize(): Initialisierung beginnt."
  'Mathematische Konstanten
  pi = 4 * Atn(1)
  rho = 200 / pi

  'Startwerte setzen
  Me.Quelle_AsciiDatei_Name = ""
  Set oDictKopfDat = New Scripting.Dictionary
  Set oQuellEinheiten = New Scripting.Dictionary
  Set oQuelldatenpuffer = New Scripting.Dictionary

  'Einheiten der Quelldaten deklarieren.
  Me.Quelle_Einheiten.Add SpN_Pkt_Nr, "ohne"
  Me.Quelle_Einheiten.Add SpN_Pkt_Erl_L, "ohne"
  Me.Quelle_Einheiten.Add SpN_Pkt_Kommentar, "ohne"
  Me.Quelle_Einheiten.Add SpN_GK_X, "m"
  Me.Quelle_Einheiten.Add SpN_GK_Y, "m"
  Me.Quelle_Einheiten.Add SpN_GK_Z, "m"
  Me.Quelle_Einheiten.Add SpN_TK_HSOK, "m"
  Me.Quelle_Einheiten.Add SpN_TK_Km, "m"
  Me.Quelle_Einheiten.Add SpN_TK_Q, "m"
  Me.Quelle_Einheiten.Add SpN_TK_St, "m"
  'Me.Quelle_Einheiten.Add SpN_Tra_u, "mm"
  Me.Quelle_Einheiten.Add SpN_S_Tra_Radius, "m"
  Me.Quelle_Einheiten.Add SpN_S_Tra_Richtung, "gon"
  Me.Quelle_Einheiten.Add SpN_S_Tra_SO, "m"
  Me.Quelle_Einheiten.Add SpN_S_Tra_u, "mm"

  Me.Quelle_Einheiten.Add SpN_TK_QG, "m"
  Me.Quelle_Einheiten.Add SpN_TK_HG, "m"
  Me.Quelle_Einheiten.Add SpN_TK_H, "m"
  Me.Quelle_Einheiten.Add SpN_TK_V, "m"
  Me.Quelle_Einheiten.Add SpN_TK_R, "m"
  Me.Quelle_Einheiten.Add SpN_TK_L, "m"
  Me.Quelle_Einheiten.Add SpN_TK_RG, "m"
  Me.Quelle_Einheiten.Add SpN_TK_LG, "m"
  
  Me.Quelle_Einheiten.Add SpN_TK_Tm, "m"
  Me.Quelle_Einheiten.Add SpN_TK_QT, "m"
  Me.Quelle_Einheiten.Add SpN_TK_QGT, "m"
  Me.Quelle_Einheiten.Add SpN_TK_HGT, "m"
  Me.Quelle_Einheiten.Add SpN_TK_QGS, "m"
  Me.Quelle_Einheiten.Add SpN_TK_HGS, "m"
  
  Me.Quelle_Einheiten.Add SpN_TK_KmStatus, "ohne"
  Me.Quelle_Einheiten.Add SpN_TK_KmText, "ohne"
  
  Me.Quelle_Einheiten.Add SpN_S_Tra_UebLi, "m"
  Me.Quelle_Einheiten.Add SpN_S_Tra_UebRe, "m"
  Me.Quelle_Einheiten.Add SpN_S_Tra_Heb, "m"
  Me.Quelle_Einheiten.Add SpN_S_Tra_G, "rad"
  Me.Quelle_Einheiten.Add SpN_S_Tra_ZLGS, "m"
  Me.Quelle_Einheiten.Add SpN_S_Tra_RaLGS, "m"
  Me.Quelle_Einheiten.Add SpN_S_Tra_AbLGS, "m"
  
  Me.Quelle_Einheiten.Add SpN_S_Tra_MiniR, "m"
  Me.Quelle_Einheiten.Add SpN_S_Tra_MiniUeb, "m"
  Me.Quelle_Einheiten.Add SpN_S_Tra_XA, "m"
  Me.Quelle_Einheiten.Add SpN_S_Tra_YA, "m"
  Me.Quelle_Einheiten.Add SpN_S_Tra_ZA, "m"
  
  Me.Quelle_Einheiten.Add SpN_DGM_HDGM, "m"
  Me.Quelle_Einheiten.Add SpN_S_DGM_ZDGM, "m"
  Me.Quelle_Einheiten.Add SpN_DGM_NameDGM, "ohne"
  
  Me.Quelle_Einheiten.Add SpN_Tra_NameAcfg, "ohne"
  Me.Quelle_Einheiten.Add SpN_Tra_NameGra, "ohne"
  Me.Quelle_Einheiten.Add SpN_Tra_NameKML, "ohne"
  Me.Quelle_Einheiten.Add SpN_Tra_NameTra, "ohne"
  Me.Quelle_Einheiten.Add SpN_Tra_NameUeb, "ohne"
  Me.Quelle_Einheiten.Add SpN_Tra_NameReg, "ohne"
  Me.Quelle_Einheiten.Add SpN_Tra_NameTun, "ohne"
  Me.Quelle_Einheiten.Add SpN_Tra_NamePkt, "ohne"
  Me.Quelle_Einheiten.Add SpN_Tra_NameGls, "ohne"
  
  Me.Quelle_Einheiten.Add SpN_Tra_STB1, "ohne"
  Me.Quelle_Einheiten.Add SpN_Tra_STB2, "ohne"
  Me.Quelle_Einheiten.Add SpN_Tra_STB3, "ohne"
  Me.Quelle_Einheiten.Add SpN_Tra_STB4, "ohne"
  Me.Quelle_Einheiten.Add SpN_Tra_STB5, "ohne"
  Me.Quelle_Einheiten.Add SpN_Tra_STB6, "ohne"
  Me.Quelle_Einheiten.Add SpN_Tra_STB7, "ohne"
  Me.Quelle_Einheiten.Add SpN_Tra_STB8, "ohne"
  Me.Quelle_Einheiten.Add SpN_Tra_STB9, "ohne"

  DebugEcho "CimpTrassenkoo.Class_Initialize(): Initialisierung beendet."
  Exit Sub

Fehler:
  Set oDictKopfDat = Nothing
  Set oQuellEinheiten = Nothing
  Set oQuelldatenpuffer = Nothing
  FehlerNachricht "CimpTrassenkoo.Class_Initialize()"
End Sub


Private Sub Class_Terminate()
  '
  On Error GoTo Fehler
  Set oDictKopfDat = Nothing
  Set oQuellEinheiten = Nothing
  Set oQuelldatenpuffer = Nothing
  DebugEcho "CimpTrassenkoo.Class_Terminate(): fertig."
  Exit Sub

Fehler:
  FehlerNachricht "CimpTrassenkoo.Class_Terminate()"
End Sub




'Eigenschaften (Interface "Importobjekt ASCII Spezial")  *******************************************

Property Get Klasse_Name() As String
  Klasse_Name = strKlasse_Name
End Property

Property Get Kategorien() As String
  'Liefert alle unterschiedlichen Kategorien der Tabelle.
  Kategorien = strKategorien
End Property

Property Get Titel() As String
  'Liefert Titel/Beschreibung der ASCII-Tabelle für den Import/Export-Dialog.
  Titel = strTitel
End Property


Property Get AktionErfolgreich() As Boolean
  'Liefert den Status des erfolgten Aktion (False: Fehler).
  AktionErfolgreich = blnAktionErfolgreich
End Property

Property Let AktionErfolgreich(inpAktionErfolgreich As Boolean)
  'Setzt den Status des erfolgten Aktion (False: Fehler).
  blnAktionErfolgreich = inpAktionErfolgreich
End Property


Property Get Fehlerniveau() As Long
  'Liefert das Fehlerniveau der erfolgten Aktion als Fehlerniveau-Konstante
  Fehlerniveau = lngFehlerniveau
End Property

Property Let Fehlerniveau(inpFehlerniveau As Long)
  'Setzt das Fehlerniveau der erfolgten Aktion als Fehlerniveau-Konstante
  lngFehlerniveau = inpFehlerniveau
End Property


Property Get Meldung_Titel() As String
  'Liefert den Titel der Meldung für CdatExpim.Import
  Meldung_Titel = strMeldung_Titel
End Property

Property Let Meldung_Titel(inpMeldung_Titel As String)
  'Setzt den Titel der Meldung für CdatExpim.Import
  strMeldung_Titel = inpMeldung_Titel
End Property


Property Get Meldung_Inhalt() As String
  'Liefert den Inhalt der Meldung für CdatExpim.Import
  Meldung_Inhalt = strMeldung_Inhalt
End Property

Property Let Meldung_Inhalt(inpMeldung_Inhalt As String)
  'Setzt den Inhalt der Meldung für CdatExpim.Import
  strMeldung_Inhalt = inpMeldung_Inhalt
End Property


Property Get Quelle_AsciiDatei_DialogFilter() As String
  'Liefert den Filter für den Dateidialog der Eingabedatei für Me.Import
  Quelle_AsciiDatei_DialogFilter = strQuelle_AsciiDatei_DialogFilter
End Property


Property Get Quelle_AsciiDatei_Name() As String
  'Liefert den Namen der Eingabedatei für Me.Import
  Quelle_AsciiDatei_Name = strQuelle_AsciiDatei_Name
End Property

Property Let Quelle_AsciiDatei_Name(inpQuelle_AsciiDatei_Name As String)
  'Setzt den Namen der Eingabedatei für Me.Import
  strQuelle_AsciiDatei_Name = inpQuelle_AsciiDatei_Name
End Property


Property Get Quelle_Datenpuffer() As Scripting.Dictionary
  'Bietet Zugriff auf den Quelldatenpuffer.
  Set Quelle_Datenpuffer = oQuelldatenpuffer
End Property


Property Get Quelle_DatenExtra() As Scripting.Dictionary
  'Bietet Zugriff auf Extradaten wie Statistik, Parameter... für Tabellenkopf.
  Set Quelle_DatenExtra = oDictKopfDat
End Property


Property Get Quelle_Einheiten() As Scripting.Dictionary
  'Bietet Zugriff auf die Einheiten der Quelldaten.
  Set Quelle_Einheiten = oQuellEinheiten
End Property


Property Get Ziel_Typ() As String
  'Liefert den Ziel-Typ der importierten Daten.
  Ziel_Typ = strZiel_Typ
End Property


Property Get Ziel_XLMappeSpeichern() As Boolean
  'Liefert den aktuellen Modus für "neu angelegte Mappe speichern".
  Ziel_XLMappeSpeichern = blnZiel_XLMappeSpeichern
End Property


Property Get Ziel_XLTabKlasse() As String
  'Liefert den Namen der XL-Tabellen-Klasse (Kodename) für Me.Import
  Ziel_XLTabKlasse = strZiel_XLTabKlasse
End Property




'Methoden (Interface "Importobjekt ASCII Spezial")  ***********************************************

Sub Import()
  'Import von Trassenkoordinaten in den klasseninternen Datenpuffer.
  'Parameter aus Me.Eigenschaften:  NameEingabeDatei = Pfad\Name der Eingabedatei.
  
  'On Error GoTo Fehler
  
  Dim NameEingabeDatei  As String
  Dim Erfolg            As Boolean
  
  NameEingabeDatei = Me.Quelle_AsciiDatei_Name
  
  If (NameEingabeDatei <> "") Then
    'Existente Eingabedatei ermittelt.
    DebugEcho "CimpTrassenkoo.Import(): Starte Import der Datei: '" & NameEingabeDatei & "'."
    Erfolg = LeseDatei(NameEingabeDatei)
  Else
    'Programmfehler, da der Dateiname bereits von CdatExpim ermittelt wird.
    ErrEcho "CimpTrassenkoo.Import(): Öffnen der folgenden Datei fehlgeschlagen: '" & NameEingabeDatei & "'."
  End If
  
  Exit Sub
  
Fehler:
  FehlerNachricht "CimpTrassenkoo.Import()"
End Sub




'interne Routinen *********************************************************************************

Private Function LeseDatei(ByVal Dateiname As String) As Boolean
  '-----------------------------------------------------------------------------------------------
  'Liest eine Datei mit Verm.esn-Umformungen oder Ausgaben von iGeo/iTrassePC (A0 und A1),
  'und schreibt alle nötigen Daten in die aktive Tabelle
  '-----------------------------------------------------------------------------------------------
  'Eingabe:  "DateiName" = gültiger Pfad\Name der zu lesenden Datei.
  'Rückgabe: - true,  falls mindestens ein Punkt gelesen wurde
  '          - False, falls kein Punkt gelesen wurde
  '-----------------------------------------------------------------------------------------------
  'Analysiert jede gelesene Zeile (ZeilenNr), um den Beginn einer Verm.esn-Umformung
  'oder einer Ausgabe von iGeo/iTrassePC (A0 und A1) festzustellen. Bei Erfolg wird
  'die passende Funktion zum Einlesen der Trassenkoordinaten aufgerufen.
  'Nach deren Ende (d.h. Ende eines Blockes mit TK) geht's hier weiter im Kreis.
  '-----------------------------------------------------------------------------------------------
  
  'On Error GoTo Fehler
  
  'Deklarationen
    Dim Kanal            As Integer
    Dim TkFormat         As String
    Dim TkHerkunft       As String
    Dim TkHerkunftDatei  As String
    Dim Key              As String
    Dim Wert             As String
    Dim ZeilenNr0        As Long
    Dim AnzErr           As Long
    Dim AnzPkt           As Long
    Dim UmfNr            As Long
    Dim ZeileAnalysiert  As Boolean
    Dim Erfolg           As Boolean
    Dim oMatches         As Object
  
  'Import-Datei öffnen
  Kanal = FreeFile()
  Open Dateiname For Input Lock Write As #Kanal
  If (Err) Then GoTo Fehler
  
  'Variableninitialisierung
    UmfNr      = 0                                   'Zähler für Anzahl der Umformungen.
    ZeilenNr   = 0                                   'Zähler für Gesamtzahl gelesener Zeilen
    TkHerkunft = ""                                  'Flag für Quellprogramm.
    TkFormat   = ""                                  'Flag für das Dateiformat. leer = unbekannt.
    Erfolg     = true
    AnzErr     = 0                                   'Zähler für Gesamtzahl gelesener Punkte
    AnzPkt     = 0                                   'Zähler für Gesamtzahl aufgetretener Fehler
    ZeileAnalysiert = true                           'Flag zeigt an, ob $0 bereits analysiert ist.
    
    'Felder initialisieren, die nicht zwangsläufig belegt (und somit redimensioniert) werden.
    ReDim AbstKm(1 To 1)
    ReDim TK_H(1 To 1)
    ReDim TK_V(1 To 1)
    ReDim TK_R(1 To 1)
    ReDim TK_L(1 To 1)
    ReDim TK_RG(1 To 1)
    ReDim TK_LG(1 To 1)
    ReDim TK_Tm(1 To 1)
    ReDim TK_QT(1 To 1)
    ReDim TK_QGT(1 To 1)
    ReDim TK_HGT(1 To 1)
    ReDim TK_QGS(1 To 1)
    ReDim TK_HGS(1 To 1)
    ReDim TK_KmStatus(1 To 1)
    ReDim TK_KmText(1 To 1)
    ReDim DGM_HDGM(1 To 1)
    ReDim DGM_ZDGM(1 To 1)
    ReDim Tra_UebLi(1 To 1)
    ReDim Tra_UebRe(1 To 1)
    ReDim Tra_Heb(1 To 1)
    ReDim Tra_G(1 To 1)
    ReDim Tra_ZLGS(1 To 1)
    ReDim Tra_RaLGS(1 To 1)
    ReDim Tra_AbLGS(1 To 1)
    
    ReDim PktKomment(1 To 1)
    
    ReDim Tra_MiniR(1 To 1)
    ReDim Tra_MiniUeb(1 To 1)
    ReDim Tra_XA(1 To 1)
    ReDim Tra_YA(1 To 1)
    ReDim Tra_ZA(1 To 1)
    
    ReDim Tra_STB1(1 To 1)
    ReDim Tra_STB2(1 To 1)
    ReDim Tra_STB3(1 To 1)
    ReDim Tra_STB4(1 To 1)
    ReDim Tra_STB5(1 To 1)
    ReDim Tra_STB6(1 To 1)
    ReDim Tra_STB7(1 To 1)
    ReDim Tra_STB8(1 To 1)
    ReDim Tra_STB9(1 To 1)
    
    ReDim N_Acfg(1 To 1)
    ReDim N_UebBand(1 To 1)
    ReDim N_RegPrf(1 To 1)
    ReDim N_TunPrf(1 To 1)
    ReDim N_PrfPkt(1 To 1)
    ReDim N_GlsPrf(1 To 1)
    ReDim N_DGM(1 To 1)
    
  
  DebugEcho "LeseDatei(): Lese Trassenkoo' aus Datei '" & Dateiname & "'."
  ProgressbarDateiLesen (Kanal)
  
  Do While true
    'Aus der Textdatei, die mehrere Umformungen (Adelt oder iGeo/iTrassePC) enthalten kann,
    'werden alle umgeformten Punkte in Vektorfelder eingelesen.
    if (ZeileAnalysiert) then
      'Falls die aktuelle Zeile bereits verarbeitet wurde => nächste Zeile lesen.
      'Falls aber Dateiende erreicht ist => Schleife verlassen.
      if (EOF(Kanal)) then
        exit do
      else
        Line Input #Kanal, Zeile
        ProgressbarDateiLesen (Kanal)
        ZeilenNr = ZeilenNr + 1
        NF = splitWords(Zeile, Feld, "")               'awk-like Splitting
      end if
    end if
    
    ZeileAnalysiert = true
    
    'Aktuelle Zeile der Eingabedatei analysieren und ggf. TK-Einleseroutine starten.
    If (NF = 0) Then
      'Leerzeile
      '... ist erlaubt, bedeutet nichts.
      
      'Keine Leerzeile.
    ElseIf (Left(Feld(1), 1) = "#") Then
      'Reine Kommentarzeile: Durchsuchen nach Angaben von iGeo/iTrassePC
      debugecho "LeseDatei(): Kommentar in Zeile " & cStr(ZeilenNr) & " der Datei '" & Dateiname & "'"
      ThisWorkbook.RegExp.IgnoreCase = True
      ThisWorkbook.RegExp.Global = False
      ThisWorkbook.RegExp.Pattern = " *([^ \t]+) *\: *(.+)[ \t]*"
      Set oMatches = ThisWorkbook.RegExp.Execute(Zeile)
      If (oMatches.Count > 0) Then
        Key = oMatches(0).SubMatches(0)
        Wert = oMatches(0).SubMatches(1)
        Select Case Key
          Case "Format":   TkFormat   = trim(Wert)
          Case "Programm": TkHerkunft = trim(Wert)
        End Select
        If ((TkHerkunft = "iTrassePC") Or (TkHerkunft = "iGeo")) Then
          UmfNr = UmfNr + 1
          debugecho "LeseDatei(): Lese " & cStr(UmfNr) & ". Block Trassenkoo' (" & TkHerkunft & ")) in Zeile " & cStr(ZeilenNr) & " der Datei '" & Dateiname & "'."
          Echo vbNewline & "Lese " & UmfNr & ". Block Trassenkoo' (" & TkHerkunft & ")) ab Zeile " & cStr(ZeilenNr) & ":"
          Echo "   - Format:       " & TkFormat
          ZeilenNr0  = ZeilenNr  'Startzeile der Umformung merken
          TkHerkunftDatei = TkHerkunft
          Erfolg     = LeseTKiGeo(Kanal, AnzPkt, AnzErr, TkHerkunft, TkFormat)
          TkHerkunft = ""
          TkFormat   = ""
          if (ZeilenNr0 < ZeilenNr) then ZeileAnalysiert = false
        End If
      End If
      
    Else
      'Zeile mit Daten und eventuell auch einem Kommentar am Zeilenende.
      debugecho "LeseDatei(): Daten in Zeile " & cStr(ZeilenNr) & " der Datei '" & Dateiname & "'."
      if ((Feld(1) = "Trassenumformung") or (Feld(1) = "Umformung")) then
        TkHerkunft = "Verm.esn"
        'LeseTkAdelt() endet mit Lesen der nächsten Überschrift oder am Dateiende.
        UmfNr = UmfNr + 1
        debugecho "LeseDatei(): Lese " & cStr(UmfNr) & ". Block Trassenkoo' (Verm.esn) in Zeile " & cStr(ZeilenNr) & " der Datei '" & Dateiname & "'."
        Echo vbNewline & "Lese " & UmfNr & ". Umformung (Verm.esn) ab Zeile " & cStr(ZeilenNr) & ":"
        ZeilenNr0  = ZeilenNr  'Startzeile der Umformung merken
        Erfolg     = LeseTkAdelt(Kanal, AnzPkt, AnzErr, TkFormat)
        TkHerkunft = ""
        TkFormat   = ""
        TkHerkunftDatei = "Verm.esn"
        if (ZeilenNr0 < ZeilenNr) then ZeileAnalysiert = false
      End If
    End If
  Loop
  Close #Kanal
  
  'Extradaten bereitstellen.
  oDictKopfDat.Add FN_Statistik_Importdatei, Dateiname
  oDictKopfDat.Add FN_Statistik_Programm, TkHerkunftDatei  'Herkunftsprogramm des letzten TK-Blockes
  
  'Schlussmeldungen ...
  Echo vbNewline & "==> Gesamtanzahl Punkte mit Trassenoordinaten = " & cStr(AnzPkt) & "."
  if (AnzErr > 0) then ErrEcho "*** Gesamtanzahl aufgetretener Fehler:        = " & cStr(AnzErr) & "  ***."
  
  if (AnzPkt = 0) then
    Erfolg = false
    if (UmfNr = 0) then ErrEcho "==> Keine Verm.esn-Umformung oder iGeo/iTrassePC-Ausgabe gefunden!"
    ErrEcho "==> Trassenkoordinaten NICHT erfolgreich gelesen!"
    
    Me.Fehlerniveau   = Fehlerniveau_Kritisch
    Me.Meldung_Titel  = "Fehler beim Import von Trassenkoordinaten"
    Me.Meldung_Inhalt = Me.Quelle_AsciiDatei_Name & " enthält keine gültigen Trassenkoordinaten (Umformung Verm.esn oder Ausgabe iGeo/iTrassePC A0 und A1)!" _
                        & vbNewline & "siehe Protokoll!"
  elseif (AnzErr > 0) then
    Echo "==> Trassenkoordinaten MIT FEHLERN gelesen."
    Erfolg = false
    
    Me.Fehlerniveau   = Fehlerniveau_Warnung
    Me.Meldung_Titel  = "Fehler beim Import von Trassenkoordinaten"
    Me.Meldung_Inhalt = cStr(AnzPkt) & " Punkte importiert aus Datei" & vbNewLine & vbNewLine & Me.Quelle_AsciiDatei_Name _
                        & vbNewline & vbNewline & cStr(AnzErr) & " Fehler - siehe Protokoll!"
  else
    Echo "==> Trassenkoordinaten erfolgreich gelesen."
    Erfolg = true
    
    Me.Fehlerniveau   = Fehlerniveau_Kein
    Me.Meldung_Titel  = "Import Trassenkoordinaten"
    Me.Meldung_Inhalt = cStr(AnzPkt) & " Punkte importiert aus Datei" & vbNewLine & vbNewLine & Me.Quelle_AsciiDatei_Name
  end if
  Me.AktionErfolgreich = Erfolg
  
  'Gelesene Daten im Quelldatenpuffer bereitstellen
  If (AnzPkt > 0) Then
    'Vektorfelder in den Quelldatenpuffer schreiben.
      '...add Spaltenname, zugehöriger Datenvektor
      
      Me.Quelle_Datenpuffer.Add SpN_Pkt_Nr,         PktNr
      Me.Quelle_Datenpuffer.Add SpN_Pkt_Erl_L,      Bemerkung
      Me.Quelle_Datenpuffer.Add SpN_Pkt_Kommentar,  PktKomment
      Me.Quelle_Datenpuffer.Add SpN_GK_X,           X
      Me.Quelle_Datenpuffer.Add SpN_GK_Y,           Y
      Me.Quelle_Datenpuffer.Add SpN_GK_Z,           Z
      Me.Quelle_Datenpuffer.Add SpN_TK_HSOK,        dH
      Me.Quelle_Datenpuffer.Add SpN_TK_Km,          Km
      Me.Quelle_Datenpuffer.Add SpN_TK_Q,           Abst
      Me.Quelle_Datenpuffer.Add SpN_TK_St,          stat
      Me.Quelle_Datenpuffer.Add SpN_Tra_NameGra,    N_Gra
      Me.Quelle_Datenpuffer.Add SpN_Tra_NameKML,    N_KmLinie
      Me.Quelle_Datenpuffer.Add SpN_Tra_NameTra,    N_Trasse
      Me.Quelle_Datenpuffer.Add SpN_S_Tra_Radius,   Radius
      Me.Quelle_Datenpuffer.Add SpN_S_Tra_Richtung, Ri
      Me.Quelle_Datenpuffer.Add SpN_S_Tra_SO,       GradHoehe
      Me.Quelle_Datenpuffer.Add SpN_S_Tra_u,        u
      
      'Nur Verm.esn
      Me.Quelle_Datenpuffer.Add SpN_TK_QKm,         AbstKm
      'Me.Quelle_Datenpuffer.Add SpN_Tra_u, u_ist
      
      'Nur iGeo/iTrassePC
      'Me.Quelle_Datenpuffer.Add SpN_TK_QG, AbstRed
      'Me.Quelle_Datenpuffer.Add SpN_TK_HG, dHRed
      Me.Quelle_Datenpuffer.Add SpN_TK_H,          TK_H
      Me.Quelle_Datenpuffer.Add SpN_TK_V,          TK_V
      Me.Quelle_Datenpuffer.Add SpN_TK_R,          TK_R
      Me.Quelle_Datenpuffer.Add SpN_TK_L,          TK_L
      Me.Quelle_Datenpuffer.Add SpN_TK_RG,         TK_RG
      Me.Quelle_Datenpuffer.Add SpN_TK_LG,         TK_LG
      Me.Quelle_Datenpuffer.Add SpN_TK_Tm,         TK_Tm
      Me.Quelle_Datenpuffer.Add SpN_TK_QT,         TK_QT
      Me.Quelle_Datenpuffer.Add SpN_TK_QGT,        TK_QGT
      Me.Quelle_Datenpuffer.Add SpN_TK_HGT,        TK_HGT
      Me.Quelle_Datenpuffer.Add SpN_TK_QGS,        TK_QGS
      Me.Quelle_Datenpuffer.Add SpN_TK_HGS,        TK_HGS
      Me.Quelle_Datenpuffer.Add SpN_TK_KmStatus,   TK_KmStatus
      Me.Quelle_Datenpuffer.Add SpN_TK_KmText,     TK_KmText
      Me.Quelle_Datenpuffer.Add SpN_S_Tra_UebLi,   Tra_UebLi
      Me.Quelle_Datenpuffer.Add SpN_S_Tra_UebRe,   Tra_UebRe
      Me.Quelle_Datenpuffer.Add SpN_S_Tra_Heb,     Tra_Heb
      Me.Quelle_Datenpuffer.Add SpN_S_Tra_G,       Tra_G
      Me.Quelle_Datenpuffer.Add SpN_S_Tra_ZLGS,    Tra_ZLGS
      Me.Quelle_Datenpuffer.Add SpN_S_Tra_RaLGS,   Tra_RaLGS
      Me.Quelle_Datenpuffer.Add SpN_S_Tra_AbLGS,   Tra_AbLGS
      Me.Quelle_Datenpuffer.Add SpN_S_Tra_MiniR,   Tra_MiniR
      Me.Quelle_Datenpuffer.Add SpN_S_Tra_MiniUeb, Tra_MiniUeb
      Me.Quelle_Datenpuffer.Add SpN_S_Tra_XA,      Tra_XA
      Me.Quelle_Datenpuffer.Add SpN_S_Tra_YA,      Tra_YA
      Me.Quelle_Datenpuffer.Add SpN_S_Tra_ZA,      Tra_ZA
      Me.Quelle_Datenpuffer.Add SpN_Tra_STB1,      Tra_STB1
      Me.Quelle_Datenpuffer.Add SpN_Tra_STB2,      Tra_STB2
      Me.Quelle_Datenpuffer.Add SpN_Tra_STB3,      Tra_STB3
      Me.Quelle_Datenpuffer.Add SpN_Tra_STB4,      Tra_STB4
      Me.Quelle_Datenpuffer.Add SpN_Tra_STB5,      Tra_STB5
      Me.Quelle_Datenpuffer.Add SpN_Tra_STB6,      Tra_STB6
      Me.Quelle_Datenpuffer.Add SpN_Tra_STB7,      Tra_STB7
      Me.Quelle_Datenpuffer.Add SpN_Tra_STB8,      Tra_STB8
      Me.Quelle_Datenpuffer.Add SpN_Tra_STB9,      Tra_STB9
      Me.Quelle_Datenpuffer.Add SpN_Tra_NameAcfg,  N_Acfg
      Me.Quelle_Datenpuffer.Add SpN_Tra_NameUeb,   N_UebBand
      Me.Quelle_Datenpuffer.Add SpN_Tra_NameReg,   N_RegPrf
      Me.Quelle_Datenpuffer.Add SpN_Tra_NameTun,   N_TunPrf
      Me.Quelle_Datenpuffer.Add SpN_Tra_NamePkt,   N_PrfPkt
      Me.Quelle_Datenpuffer.Add SpN_Tra_NameGls,   N_GlsPrf
      Me.Quelle_Datenpuffer.Add SpN_DGM_HDGM,      DGM_HDGM
      Me.Quelle_Datenpuffer.Add SpN_S_DGM_ZDGM,    DGM_ZDGM
      Me.Quelle_Datenpuffer.Add SpN_DGM_NameDGM,   N_DGM
      
    'Speicher freigeben (Wirkung?)
      ReDim N_KmLinie(0)
      ReDim N_Trasse(0)
      ReDim N_Gra(0)
      ReDim Abst(0)
      ReDim Bemerkung(0)
      ReDim PktKomment(0)
      ReDim GradHoehe(0)
      ReDim PktNr(0)
      ReDim Radius(0)
      ReDim Ri(0)
      ReDim dH(0)
      ReDim stat(0)
      ReDim Km(0)
      ReDim u(0)
      ReDim X(0)
      ReDim Y(0)
      ReDim Z(0)
      
      'Nur Verm.esn
      ReDim AbstKm(0)
      'ReDim u_ist(0)
        
      'Nur iGeo/iTrassePC
      'ReDim AbstRed(0)
      'ReDim dHRed(0)
      ReDim TK_H(0)
      ReDim TK_V(0)
      ReDim TK_R(0)
      ReDim TK_L(0)
      ReDim TK_RG(0)
      ReDim TK_LG(0)
      ReDim TK_Tm(0)
      ReDim TK_QT(0)
      ReDim TK_QGT(0)
      ReDim TK_HGT(0)
      ReDim TK_QGS(0)
      ReDim TK_HGS(0)
      ReDim TK_KmStatus(0)
      ReDim TK_KmText(0)
      ReDim DGM_HDGM(0)
      ReDim DGM_ZDGM(0)
      ReDim Tra_UebLi(0)
      ReDim Tra_UebRe(0)
      ReDim Tra_Heb(0)
      ReDim Tra_G(0)
      ReDim Tra_ZLGS(0)
      ReDim Tra_RaLGS(0)
      ReDim Tra_AbLGS(0)
      ReDim Tra_MiniR(0)  
      ReDim Tra_MiniUeb(0)
      ReDim Tra_STB1(0)
      ReDim Tra_STB2(0)
      ReDim Tra_STB3(0)
      ReDim Tra_STB4(0)
      ReDim Tra_STB5(0)
      ReDim Tra_STB6(0)
      ReDim Tra_STB7(0)
      ReDim Tra_STB8(0)
      ReDim Tra_STB9(0)
      ReDim Tra_XA(0)      
      ReDim Tra_YA(0)      
      ReDim Tra_ZA(0)      
      ReDim N_Acfg(0)
      ReDim N_UebBand(0)
      ReDim N_RegPrf(0)
      ReDim N_TunPrf(0)
      ReDim N_PrfPkt(0)
      ReDim N_GlsPrf(0)
      ReDim N_DGM(0)
  End If
  
  LeseDatei = Erfolg
  Exit Function
  
Fehler:
  Close #Kanal
  Application.StatusBar = ""
  Me.AktionErfolgreich = False
  Me.Fehlerniveau   = Fehlerniveau_Kritisch
  FehlerNachricht "CimpTrassenkoo.LeseDatei()"
  LeseDatei = Erfolg
End Function



Private Function LeseTKiGeo(ByVal Kanal As Integer, ByRef AnzPkt As Long, ByRef AnzErr As Long, ByVal TkHerkunft As String, ByVal TkFormat As String) As Boolean
  ' -----------------------------------------------------------------------------------------------
  ' Liest eine Ausgabe von iGeo/iTrassePC aus der aktiven Datei.
  ' -----------------------------------------------------------------------------------------------
  ' Eingabe:   Kanal      ... Dateihandle
  '            AnzPkt     ... Anzahl erfolgreich gelesener Punkte, wird weiter gezählt.
  '            AnzErr     ... Anzahl erkannter Fehler, wird weiter gezählt.
  '            TkHerkunft ... "iTrassePC" oder "iGeo"
  '            TkFormat   ... iGeo/iTrassePC-Formatkennung: muss beim Aufruf bereits bekannt sein!
  'Rückgabe: - true,  falls mindestens ein Punkt gelesen wurde
  '          - False, falls kein Punkt gelesen wurde
   '
   ' Status nach Beendigung dieser Funktion:
   '  - Es ist bereits die Zeile "# Format : xxxx" der nächsten iGeo/iTrassePC-Ausgabe derselben Datei gelesen, oder:
   '  - Dateiende ist erreicht.
   '
  ' ==> unterstützte Ausgabeformate: A0 und A1 (iTrasse finale Version 2.0.2, iGeo Trassenbibliothek V227).
  ' ==> Ein am Zeilenende platzierter (mit "#" eingeleiteter) Kommentar wird als Punktinfo gespeichert.
  ' -----------------------------------------------------------------------------------------------
  
  'On Error GoTo Fehler
  
  'Deklarationen
    Dim Meldung         As String
    Dim Meldung2        As String
    
    Dim NAcfg           As String
    Dim NTra            As String
    Dim NKML            As String
    Dim NGra            As String
    Dim NUeb            As String
    Dim NRegel          As String
    Dim NTunnel         As String
    Dim NPrfPkt         As String
    Dim NGlsPrf         As String
    Dim NDGM            As String
    Dim FieldName       As String
    Dim FieldNamesString As String
    
    Dim Kommentar       As String
    Dim strBemerkung    As String
    Dim Datenzeile      As String
    Dim UmfEnde         As Boolean
    Dim Erfolg          As Boolean
    Dim TkFormat_ok     As Boolean
    Dim lfdNr           As Long
    Dim AnzErr0         As Long
    Dim offset          As Integer
    Dim i               As Integer
    Dim idx             As Variant
    
    Dim vGradHoehe      As Variant
    Dim vdH             As Variant
    Dim vY              As Variant
    Dim vX              As Variant
    Dim vZ              As Variant
    Dim Ri_tmp          As Variant
    
    Dim MinAnzWort      As Integer
    Dim F_Ueb           As Integer
    Dim W_PktNr         As Integer
    Dim W_Y             As Integer
    Dim W_X             As Integer
    Dim W_Z             As Integer
    Dim W_St            As Integer
    Dim W_Km            As Integer
    Dim W_Q             As Integer
    Dim W_H             As Integer
    Dim W_HSOK          As Integer
    Dim W_ZSOK          As Integer
    Dim W_ZLGS          As Integer
    Dim W_RaLGS         As Integer
    Dim W_AbLGS         As Integer
    Dim W_MiniR         As Integer
    Dim W_MiniUeb       As Integer
    Dim W_XA            As Integer
    Dim W_YA            As Integer
    Dim W_ZA            As Integer
    Dim W_STB1          As Integer
    Dim W_STB2          As Integer
    Dim W_STB3          As Integer
    Dim W_STB4          As Integer
    Dim W_STB5          As Integer
    Dim W_STB6          As Integer
    Dim W_STB7          As Integer
    Dim W_STB8          As Integer
    Dim W_STB9          As Integer
    Dim W_QG            As Integer
    Dim W_HG            As Integer
    Dim W_RG            As Integer
    Dim W_LG            As Integer
    Dim W_QGT           As Integer
    Dim W_HGT           As Integer
    Dim W_QGS           As Integer
    Dim W_HGS           As Integer
    Dim W_KmStatus      As Integer
    Dim W_KmText        As Integer
    Dim W_Tm            As Integer
    Dim W_QT            As Integer
    Dim W_UebLi         As Integer
    Dim W_UebRe         As Integer
    Dim W_Ueb           As Integer
    Dim W_Heb           As Integer
    Dim W_G             As Integer
    Dim W_Ri            As Integer
    Dim W_Ra            As Integer
    Dim W_V             As Integer
    Dim W_R             As Integer
    Dim W_L             As Integer
    Dim W_HDGM          As Integer
    Dim W_ZDGM          As Integer
    Dim W_Text          As Integer
    Dim W_PktKomment    As Integer
    
    Dim Key             As String
    Dim Wert            As String
    Dim oMatches        As Object
    Dim oFSO            As New Scripting.FileSystemObject
    
  'Unterstützte Formatkennungen
    Const Fmt_A0 = "A0 : Alles EDV"
    Const Fmt_A1 = "A1 : PktNr  Y  X  Z  St/Km"  ' kleinster gemeinsamer Nenner von iTrassePC und iGeo
    
  'Variableninitialisierung
    const TkEntstehung   As String = "Ausgabe"
    const DoublePipeMask As String = "D1o2u3b4l5e6P7i8p9e0"
    
    NAcfg        = ""
    NKML         = ""
    NGra         = ""
    NTra         = ""
    NUeb         = ""
    NRegel       = ""
    NTunnel      = ""
    NPrfPkt      = ""
    NGlsPrf      = ""
    NDGM         = ""
    FieldNamesString = ""
    
    UmfEnde      = false
    Erfolg       = true
    lfdNr        = 0
    AnzErr0      = AnzErr      'Anzahl Fehler zu Beginn dieser Routine
    
    'Wortnummer der Werte in der Datenzeile (0...existiert nicht).
    F_Ueb       = 0  'Faktor für Überhöhung
    W_PktNr     = 0
    W_Y         = 0
    W_X         = 0
    W_Z         = 0
    W_St        = 0
    W_Km        = 0
    W_Q         = 0
    W_H         = 0
    W_HSOK      = 0
    W_ZSOK      = 0
    W_ZLGS      = 0
    W_RaLGS     = 0
    W_AbLGS     = 0
    W_MiniR     = 0
    W_MiniUeb   = 0
    W_XA        = 0
    W_YA        = 0
    W_ZA        = 0
    W_STB1      = 0
    W_STB2      = 0
    W_STB3      = 0
    W_STB4      = 0
    W_STB5      = 0
    W_STB6      = 0
    W_STB7      = 0
    W_STB8      = 0
    W_STB9      = 0
    W_QG        = 0
    W_HG        = 0
    W_RG        = 0
    W_LG        = 0
    W_QGT       = 0
    W_HGT       = 0
    W_QGS       = 0
    W_HGS       = 0
    W_KmStatus  = 0
    W_KmText    = 0
    W_Tm        = 0
    W_QT        = 0
    W_UebLi     = 0
    W_UebRe     = 0
    W_Ueb       = 0
    W_Heb       = 0
    W_G         = 0
    W_Ri        = 0
    W_Ra        = 0
    W_V         = 0
    W_R         = 0
    W_L         = 0
    W_HDGM      = 0
    W_ZDGM      = 0
    W_Text      = 0
    W_PktKomment= 0
    
  'Format-Kennung prüfen!
  If ((TkFormat = Fmt_A0) Or entspricht(Fmt_A1, TkFormat)) Then
    
    TkFormat_ok = true
    
  Else
    TkFormat_ok = false
    Erfolg = false
    AnzErr = AnzErr + 1
    ErrEcho " *** Fehler:       Nicht unterstütztes Ausgabeformat von iGeo/iTrassePC: '" & TkFormat & "'."
    call ThisWorkbook.SysTools.FileErrorsAdd("F", Me.Quelle_AsciiDatei_Name, ZeilenNr - 1, 0, 0, "Nicht unterstütztes Ausgabeformat von iGeo/iTrassePC: '" & TkFormat & "'.", _
                                "Unterstützte Formatkennungen:" & vbNewLine & _
                                Fmt_A0 & vbNewLine & Fmt_A1 & vbNewLine & _
                                "Quelle dieser Meldung: Einlesen einer Datei mit Trassenkoordinaten aus iGeo (CimpTrassenkoo.LeseTKiGeo())", _
                                blnFehlerKonsolidieren)
  End If
  
  ' Gesamte "Umformung" von iGeo/iTrassePC einlesen
  if (TkFormat_ok) then
    
    Do
      'Nächste Zeile lesen.
      'Falls aber Dateiende erreicht ist => Schleife verlassen.
      If (EOF(Kanal)) Then
        Exit Do
      Else
        Line Input #Kanal, Zeile
        ProgressbarDateiLesen (Kanal)
        ZeilenNr  = ZeilenNr + 1
        TrimZeile = Trim(Zeile)
      End If
      
      If (Not (TrimZeile = "")) Then
        'Keine Leerzeile.
        Kommentar  = ""
        Datenzeile = ""
        
        If (Left(TrimZeile, 1) = "#") Then
          'Reine Kommentarzeile: Durchsuchen nach Angaben von iGeo/iTrassePC.
          debugecho "LeseTKiGeo(): Kommentar in Zeile " & cStr(ZeilenNr)
          ThisWorkbook.RegExp.IgnoreCase = True
          ThisWorkbook.RegExp.Global = False
          ThisWorkbook.RegExp.Pattern = "^# *([^\:]+)\: *(.+) *"
          Set oMatches = ThisWorkbook.RegExp.Execute(Zeile)
          If (oMatches.Count > 0) Then
            Key = Trim(oMatches(0).SubMatches(0))
            Wert = oMatches(0).SubMatches(1)
            Select Case Key
              Case "Trassenkonfiguration":  NAcfg   = oFSO.GetBaseName(Wert)
              Case "Achse":                 NTra    = oFSO.GetFileName(Wert)
              Case "Gradiente":             NGra    = oFSO.GetFileName(Wert)
              Case "Ueberhöhungsband":      NUeb    = oFSO.GetFileName(Wert)  ' iTrassePC
              Case "Überhöhungsband":       NUeb    = oFSO.GetFileName(Wert)  ' iGeo
              Case "Stationierungsachse":   NKML    = oFSO.GetFileName(Wert)  ' iTrassePC
              Case "Km-Linie":              NKML    = oFSO.GetFileName(Wert)  ' iGeo
              Case "Regelprofilbereich":    NRegel  = oFSO.GetFileName(Wert)
              Case "Tunnelprofilbereich":   NTunnel = oFSO.GetFileName(Wert)
              Case "Profilpunktbereich":    NPrfPkt = oFSO.GetFileName(Wert)
              Case "Gleisprofilbereich":    NGlsPrf = oFSO.GetFileName(Wert)
              Case "DGM":                   NDGM    = oFSO.GetFileName(Wert)
              Case "Feldnamen":             FieldNamesString = trim(Wert)     ' Format "A0"
              
              Case "Format"
                  UmfEnde = true  'Kommentarblock der nächsten "Umformung" von iGeo/iTrassePC => d.h. die bisherige Umformung ist zu Ende.
                  NF = splitWords(Zeile, Feld, "")               'awk-like Splitting
            End Select
          End If
          
        Else
          'Zeile mit Daten und eventuell auch einem Kommentar am Zeilenende => sollte reguläre Daten enthalten.
          
          'Erster Punkt: Format-"Spezifikation" (erst jetzt, da z.T. abhängig von Verfügbarkeit KML und Üb.Band).
          If (lfdNr = 0) Then
            If (TkFormat = Fmt_A0) Then
              
              If (FieldNamesString = "") Then
                Erfolg = false
                AnzErr = AnzErr + 1
                ErrEcho " *** Fehler:       Ausgabeformat iGeo A0: keine Angabe der FELDNAMEN gefunden!"
                Call ThisWorkbook.SysTools.FileErrorsAdd("f", Me.Quelle_AsciiDatei_Name, ZeilenNr, 0, 0, "Ausgabeformat iGeo A0: keine Angabe der FELDNAMEN gefunden.", _
                                             "Im Kopf der A0-Ausgabe muss die Angabe '# Feldnamen: <Namensliste>' die ausgegebenen Felder auflisten." & vbNewLine & _
                                             "Quelle dieser Meldung: Einlesen einer Datei mit Trassenkoordinaten aus iGeo/iTrassePC (CimpTrassenkoo.LeseTKiGeo())", _
                                             blnFehlerKonsolidieren)
                Exit Do
              Else
                F_Ueb = 1000
                
                NF = splitWords(FieldNamesString, Feld, "[^|]+")    'awk-like Splitting
                Feld(0) = ""                                        'Index 0 wird unten verwendet!
                
                MinAnzWort = NF
                
                For i = 1 To NF
                  
                  FieldName = trim(Feld(i))
                  
                  If (FieldName = "PktNr") Then
                    W_PktNr = i
                  ElseIf (FieldName = "Y") Then
                    W_Y = i
                  ElseIf (FieldName = "X") Then
                    W_X = i
                  ElseIf (FieldName = "Z") Then
                    W_Z = i
                  ElseIf (FieldName = "St") Then
                    W_St = i
                  ElseIf (FieldName = "Km") Then
                    W_Km = i
                  ElseIf (FieldName = "Q") Then
                    W_Q = i
                  ElseIf (FieldName = "H") Then
                    W_H = i
                  ElseIf (FieldName = "HSOK") Then
                    W_HSOK = i
                  ElseIf (FieldName = "ZSOK") Then
                    W_ZSOK = i
                  ElseIf (FieldName = "QG") Then
                    W_QG = i
                  ElseIf (FieldName = "HG") Then
                    W_HG = i
                  ElseIf (FieldName = "QT") Then
                    W_QT = i
                  ElseIf (FieldName = "Tm") Then
                    W_Tm = i
                  ElseIf (FieldName = "UebLi") Then
                    W_UebLi = i
                  ElseIf (FieldName = "UebRe") Then
                    W_UebRe = i
                  ElseIf (FieldName = "Ueb") Then
                    W_Ueb = i
                  ElseIf (FieldName = "Heb") Then
                    W_Heb = i
                  ElseIf (FieldName = "G") Then
                    W_G = i
                  ElseIf (FieldName = "Ri") Then
                    W_Ri = i
                  ElseIf (FieldName = "Ra") Then
                    W_Ra = i
                  ElseIf (FieldName = "V") Then
                    W_V = i
                  ElseIf (FieldName = "R") Then
                    W_R = i
                  ElseIf (FieldName = "L") Then
                    W_L = i
                  ElseIf (FieldName = "HDGM") Then
                    W_HDGM = i
                  ElseIf (FieldName = "ZDGM") Then
                    W_ZDGM = i
                  ElseIf (FieldName = "RG") Then
                    W_RG = i
                  ElseIf (FieldName = "LG") Then
                    W_LG = i
                  ElseIf (FieldName = "QGT") Then
                    W_QGT = i
                  ElseIf (FieldName = "HGT") Then
                    W_HGT = i
                  ElseIf (FieldName = "QGS") Then
                    W_QGS = i
                  ElseIf (FieldName = "HGS") Then
                    W_HGS = i
                  ElseIf (FieldName = "KmStatus") Then
                    W_KmStatus = i
                  ElseIf (FieldName = "KmText") Then
                    W_KmText = i
                  ElseIf (FieldName = "ZLGS") Then
                    W_ZLGS = i
                  ElseIf (FieldName = "RaLGS") Then
                    W_RaLGS = i
                  ElseIf (FieldName = "AbLGS") Then
                    W_AbLGS = i
                  ElseIf (FieldName = "MiniR") Then
                    W_MiniR = i
                  ElseIf (FieldName = "MiniUeb") Then
                    W_MiniUeb = i
                  ElseIf (FieldName = "XA") Then
                    W_XA = i
                  ElseIf (FieldName = "YA") Then
                    W_YA = i
                  ElseIf (FieldName = "ZA") Then
                    W_ZA = i
                  ElseIf (FieldName = "STB1") Then
                    W_STB1 = i
                  ElseIf (FieldName = "STB2") Then
                    W_STB2 = i
                  ElseIf (FieldName = "STB3") Then
                    W_STB3 = i
                  ElseIf (FieldName = "STB4") Then
                    W_STB4 = i
                  ElseIf (FieldName = "STB5") Then
                    W_STB5 = i
                  ElseIf (FieldName = "STB6") Then
                    W_STB6 = i
                  ElseIf (FieldName = "STB7") Then
                    W_STB7 = i
                  ElseIf (FieldName = "STB8") Then
                    W_STB8 = i
                  ElseIf (FieldName = "STB9") Then
                    W_STB9 = i
                  ElseIf (FieldName = "Text") Then
                    W_Text = i
                  ElseIf (FieldName = "Kommentar") Then
                    W_PktKomment = i
                  End If
                Next
              End If
              
            ElseIf (entspricht(Fmt_A1, TkFormat)) Then
              
              MinAnzWort = 7
              W_PktNr  = 1
              W_Y      = 2
              W_X      = 3
              W_Z      = 4
              if (NKML = "") then
                W_St   = 5
                'if (StAlsKm) W_Km = W_St   # Station auch als Km.
              else
                W_Km   = 5
              end if
              W_Q      = 6
              if (NGra <> "") then
                if (NUeb = "") then
                  W_H    = 7
                else
                  W_HSOK = 7
                end if
              end if
              W_Text   = 8
              
            Else
              Erfolg = false
              ErrEcho " *** Programmfehler: Nicht unterstütztes Ausgabeformat von iGeo/iTrassePC: '" & TkFormat & "' (zu spät festgestellt :-( )."
              Call ThisWorkbook.SysTools.FileErrorsAdd("F", Me.Quelle_AsciiDatei_Name, ZeilenNr - 1, 0, 0, "Nicht unterstütztes Ausgabeformat von iGeo/iTrassePC: '" & TkFormat & "' (zu spät festgestellt :-( ).", _
                                           "Unterstützte Formatkennungen:" & vbNewLine & _
                                           Fmt_A0 & vbNewLine & Fmt_A1 & vbNewLine & _
                                           "Quelle dieser Meldung: Einlesen einer Datei mit Trassenkoordinaten aus iGeo/iTrassePC (CimpTrassenkoo.LeseTKiGeo())", _
                                           blnFehlerKonsolidieren)
              exit do
            End If
          End If
          
          ' Zeile teilen in Felder
          If (TkFormat = Fmt_A0) Then
            
            ' Maskierte Trennzeichen ("||") vor splitWords() verstecken
            Datenzeile = replace(Zeile, "||", DoublePipeMask)
            
            NF = splitWords(Datenzeile, Feld, "[^|]+")   'awk-like Splitting
            Feld(0) = ""                                 'Index 0 wird unten verwendet!
            
            ' Falls die Datenzeile mit einem Feldtrenner endet, hat splitWords() das letzte (leere) Wort unterschlagen...
            If (right(Datenzeile, 1) = "|") Then
              NF = NF + 1
              ReDim Preserve Feld(0 To NF)
              Feld(NF) = ""
            End If
            
            ' Maskierte Trennzeichen wiederherstellen
            For i = 1 To NF
              Feld(i) = trim(Feld(i))
              Feld(i) = replace(Feld(i), DoublePipeMask, "|")
            Next
            
            if (NF >= MinAnzWort) then
              ' Kommentar ist bei A0 ein Feld wie alle anderen
              Kommentar = Feld(W_PktKomment)
            end if
            
          Else
            'Zeile teilen in "Datenzeile" und "Kommentar".
            idx = instr(Zeile, "#")
            if (idx > 0) then
              'Kommentarzeichen gefunden (nicht als erstes Zeichen - dieser Fall ist bereits abgefangen).
              Datenzeile = left(Zeile, idx - 1)
              Kommentar  = mid(Zeile, idx + 1)
              Kommentar  = substitute("^ ",  "", Kommentar, false, false)  'Ein einzelnes Leerzeichen am Anfang entfernen
              Kommentar  = substitute(" +$", "", Kommentar, false, false)  'Alle Leerzeichen am Ende entfernen
            else
              Datenzeile = Zeile
              Kommentar  = ""
            end if
            
            'Kommentar vorauswerten.
            'if (Kommentar <> "") then
            'end if
            
            'Datenabschnitt der Zeile splitten.
            NF = splitWords(Datenzeile, Feld, "")  'awk-like Splitting
            Feld(0) = ""                           'Index 0 wird unten verwendet!
          End If
          
          'Felder der Zeile auswerten.
          If (NF < MinAnzWort) Then
            'Fehler: zu wenig Werte.
            Erfolg = false
            AnzErr = AnzErr + 1
            ErrEcho "Fehler in Zeile " & cStr(ZeilenNr) & " der Eingabedatei: zu wenig Datenfelder (erwartet: " & MinAnzWort & ")."
            Meldung  = "Fehler: Zu wenig Datenfelder (erwartet: " & cStr(MinAnzWort) & ")."
            Meldung2 = "Format dieser Ausgabedatei von " & TkHerkunft & ": '" & TkFormat &  "'" & vbNewLine & "Quelle dieser Meldung: 'CimpTrassenkoo.LeseTKiGeo()'"
            Call ThisWorkbook.SysTools.FileErrorsAdd("F", Me.Quelle_AsciiDatei_Name, ZeilenNr, 0, 0, Meldung, Meldung2, blnFehlerKonsolidieren)
          Else
            
            'Zähler
              lfdNr  = lfdNr + 1               'Punktanzahl in dieser Umformung
              AnzPkt = AnzPkt + 1              'Punktanzahl gesamt
            
            'Initialwerte
              vGradHoehe = ""
              vdH = ""
              vY = ""
              vX = ""
              vZ = ""
            
            'Spalten um je 1 Zeile verlängern
              ReDim Preserve N_KmLinie(1 To AnzPkt)
              ReDim Preserve N_Trasse(1 To AnzPkt)
              ReDim Preserve N_Gra(1 To AnzPkt)
              ReDim Preserve N_Acfg(1 To AnzPkt)
              ReDim Preserve N_UebBand(1 To AnzPkt)
              ReDim Preserve N_RegPrf(1 To AnzPkt)
              ReDim Preserve N_TunPrf(1 To AnzPkt)
              ReDim Preserve N_PrfPkt(1 To AnzPkt)
              ReDim Preserve N_GlsPrf(1 To AnzPkt)
              ReDim Preserve N_DGM(1 To AnzPkt)
              ReDim Preserve Abst(1 To AnzPkt)
              ReDim Preserve Bemerkung(1 To AnzPkt)
              ReDim Preserve PktKomment(1 To AnzPkt)
              ReDim Preserve GradHoehe(1 To AnzPkt)
              ReDim Preserve PktNr(1 To AnzPkt)
              ReDim Preserve Radius(1 To AnzPkt)
              ReDim Preserve Ri(1 To AnzPkt)
              ReDim Preserve dH(1 To AnzPkt)
              ReDim Preserve stat(1 To AnzPkt)
              ReDim Preserve Km(1 To AnzPkt)
              ReDim Preserve u(1 To AnzPkt)
              ReDim Preserve X(1 To AnzPkt)
              ReDim Preserve Y(1 To AnzPkt)
              ReDim Preserve Z(1 To AnzPkt)
              
              'ReDim Preserve AbstRed(1 To AnzPkt)
              'ReDim Preserve dHRed(1 To AnzPkt)
              ReDim Preserve TK_H(1 To AnzPkt)
              ReDim Preserve TK_V(1 To AnzPkt)
              ReDim Preserve TK_R(1 To AnzPkt)
              ReDim Preserve TK_L(1 To AnzPkt)
              ReDim Preserve TK_RG(1 To AnzPkt)
              ReDim Preserve TK_LG(1 To AnzPkt)
              ReDim Preserve TK_Tm(1 To AnzPkt)
              ReDim Preserve TK_QT(1 To AnzPkt)
              ReDim Preserve TK_QGT(1 To AnzPkt)
              ReDim Preserve TK_HGT(1 To AnzPkt)
              ReDim Preserve TK_QGS(1 To AnzPkt)
              ReDim Preserve TK_HGS(1 To AnzPkt)
              ReDim Preserve TK_KmStatus(1 To AnzPkt)
              ReDim Preserve TK_KmText(1 To AnzPkt)
              ReDim Preserve DGM_HDGM(1 To AnzPkt)
              ReDim Preserve DGM_ZDGM(1 To AnzPkt)
              ReDim Preserve Tra_UebLi(1 To AnzPkt)
              ReDim Preserve Tra_UebRe(1 To AnzPkt)
              ReDim Preserve Tra_Heb(1 To AnzPkt)
              ReDim Preserve Tra_G(1 To AnzPkt)
              ReDim Preserve Tra_ZLGS(1 To AnzPkt)
              ReDim Preserve Tra_RaLGS(1 To AnzPkt)
              ReDim Preserve Tra_AbLGS(1 To AnzPkt)
              ReDim Preserve Tra_MiniR(1 To AnzPkt)
              ReDim Preserve Tra_MiniUeb(1 To AnzPkt)
              ReDim Preserve Tra_XA(1 To AnzPkt)
              ReDim Preserve Tra_YA(1 To AnzPkt)
              ReDim Preserve Tra_ZA(1 To AnzPkt)
              ReDim Preserve Tra_STB1(1 To AnzPkt)
              ReDim Preserve Tra_STB2(1 To AnzPkt)
              ReDim Preserve Tra_STB3(1 To AnzPkt)
              ReDim Preserve Tra_STB4(1 To AnzPkt)
              ReDim Preserve Tra_STB5(1 To AnzPkt)
              ReDim Preserve Tra_STB6(1 To AnzPkt)
              ReDim Preserve Tra_STB7(1 To AnzPkt)
              ReDim Preserve Tra_STB8(1 To AnzPkt)
              ReDim Preserve Tra_STB9(1 To AnzPkt)
              
            'Doku:  Zuordnungen  Kürzel_iGeo/iTrassePC ... Array
              'PktNr    ... PktNr
              'Y        ... Y
              'X        ... X
              'Z        ... Z
              'St       ... stat
              'Km       ... Km
              'Q        ... Abst
              'HSOK     ... dH
              'ZSOK     ... GradHoehe
              'Ueb      ... u
              'Ri       ... Ri (umrechnen rad => gon)
              'Ra       ... Radius
              'Code     ... Bemerkung
              
              'Spalten, die der Datenpuffer nur durch iGeo/iTrassePC erhält.
              'QG       ... AbstRed
              'HG       ... dHRed
              'H        ... TK_H
              'UebLi    ... Tra_UebLi
              'UebRe    ... Tra_UebRe
              'Heb      ... Tra_Heb
              'G        ... Tra_G
              'ZLGS     ... Tra_ZLGS
              'RaLGS    ... Tra_RaLGS
              'AbLGS    ... Tra_AbLGS
              'MiniR    ... Tra_MiniR
              'MiniUeb  ... Tra_MiniUeb
              'XA       ... Tra_XA
              'YA       ... Tra_YA
              'ZA       ... Tra_ZA
              'STB1     ... Tra_STB1
              'STB2     ... Tra_STB2
              'STB3     ... Tra_STB3
              'STB4     ... Tra_STB4
              'STB5     ... Tra_STB5
              'STB6     ... Tra_STB6
              'STB7     ... Tra_STB7
              'STB8     ... Tra_STB8
              'STB9     ... Tra_STB9
              'V        ... TK_V
              'R        ... TK_R
              'L        ... TK_L
              'RG       ... TK_RG
              'LG       ... TK_LG
              'Tm       ... TK_Tm
              'QT       ... TK_QT
              'QGT      ... TK_QGT
              'HGT      ... TK_HGT
              'QGS      ... TK_QGS
              'HGS      ... TK_HGS
              'KmStatus ... TK_KmStatus
              'KmText   ... TK_KmText
              'HDGM     ... DGM_HDGM
              'ZDGM     ... DGM_ZDGM
            '
            'Fehler tolerieren, da Datenwörter am Zeilenende fehlen dürfen.
            On Error Resume Next
            
            'Großkoordinaten
              vY = GetValue(Feld(W_Y))
              vX = GetValue(Feld(W_X))
              vZ = GetValue(Feld(W_Z))
              
            'Höhenwerte lesen bzw. berechnen
              if (NGra <> "") then vGradHoehe = Feld(W_ZSOK)
              if ((NGra <> "") And (NUeb <> "")) then vdH = Feld(W_HSOK)
            
              'Fehlende Höhenwerte berechnen, wenn möglich.
              If ((vZ = "") And (vdH <> "") And (vGradHoehe <> "")) Then vZ = CDbl(vGradHoehe) + CDbl(vdH)
              If ((vGradHoehe = "") And (vdH <> "") And (vZ <> "")) Then vGradHoehe = CDbl(vZ) - CDbl(vdH)
            
            'Bemerkung: Text und Kommentar verketten.
              'if (trim(Feld(W_Text)) <> "") then
              '  strBemerkung = Feld(W_Text)
              '  
              '  if (trim(Kommentar) <> "") then
              '    strBemerkung = strBemerkung & ", " & trim(Kommentar)
              '  end if
              '  
              'else
              '  strBemerkung = Kommentar
              'end if
            
            'Ist-Überhöhung (nicht hier ermitteln => wird nachträglich gemacht)
              'u_ist = Ueberhoehung(Bemerkung, UebInInfo_Streng)     # ist-Überhöhung in (mm) aus Bemerkung ermitteln
            
            'Krücke: Bei Überhöhung in der Geraden wird R = +-99999.000 m gesetzt, wenn bestimmbar.
              if ((abs(Feld(W_Ueb)) > 0.0) And ((Feld(W_Ra) = "") Or (Feld(W_Ra) = 0.0))) then
                if (isNumeric(Feld(W_UebLi)) And isNumeric(Feld(W_UebRe))) then
                  if (Feld(W_UebLi) > Feld(W_UebRe)) then
                    Feld(W_Ra) = 99999.000
                  elseif (Feld(W_UebLi) < Feld(W_UebRe)) then
                    Feld(W_Ra) = -99999.000
                  end if
                end if
              end if
            
            'Felder belegen
              PktNr(AnzPkt)       = Feld(W_PktNr)
              Y(AnzPkt)           = vY
              X(AnzPkt)           = vX
              Z(AnzPkt)           = vZ
              'Bemerkung(AnzPkt)   = strBemerkung
              Bemerkung(AnzPkt)   = Feld(W_Text)
              PktKomment(AnzPkt)  = Kommentar
              
              Km(AnzPkt)          = GetValue(Feld(W_Km))
              Stat(AnzPkt)        = GetValue(Feld(W_St))
              Abst(AnzPkt)        = GetValue(Feld(W_Q))
              'AbstKm(AnzPkt)     = GetValue(Feld(W_QKm))
              dH(AnzPkt)          = vdH
              TK_H(AnzPkt)        = GetValue(Feld(W_H))
              
              'AbstRed(AnzPkt)    = GetValue(Feld(W_QG))
              'dHRed(AnzPkt)      = GetValue(Feld(W_HG))
              
              GradHoehe(AnzPkt)   = vGradHoehe
              Radius(AnzPkt)      = GetValue(Feld(W_Ra))
              
              TK_V(AnzPkt)        = GetValue(Feld(W_V))
              TK_R(AnzPkt)        = GetValue(Feld(W_R))
              TK_L(AnzPkt)        = GetValue(Feld(W_L))
              
              TK_RG(AnzPkt)       = GetValue(Feld(W_RG))
              TK_LG(AnzPkt)       = GetValue(Feld(W_LG))
              
              TK_Tm(AnzPkt)       = GetValue(Feld(W_Tm))
              TK_QT(AnzPkt)       = GetValue(Feld(W_QT))
              TK_QGT(AnzPkt)      = GetValue(Feld(W_QGT))
              TK_HGT(AnzPkt)      = GetValue(Feld(W_HGT))
              TK_QGS(AnzPkt)      = GetValue(Feld(W_QGS))
              TK_HGS(AnzPkt)      = GetValue(Feld(W_HGS))
              
              Tra_UebLi(AnzPkt)   = GetValue(Feld(W_UebLi))
              Tra_UebRe(AnzPkt)   = GetValue(Feld(W_UebRe))
              Tra_Heb(AnzPkt)     = GetValue(Feld(W_Heb))
              Tra_G(AnzPkt)       = GetValue(Feld(W_G))
              Tra_ZLGS(AnzPkt)    = GetValue(Feld(W_ZLGS))
              Tra_RaLGS(AnzPkt)   = GetValue(Feld(W_RaLGS))
              Tra_AbLGS(AnzPkt)   = GetValue(Feld(W_AbLGS))
              
              Tra_MiniUeb(AnzPkt) = GetValue(Feld(W_MiniUeb))
              Tra_MiniR(AnzPkt)   = GetValue(Feld(W_MiniR))
              Tra_XA(AnzPkt)      = GetValue(Feld(W_XA))
              Tra_YA(AnzPkt)      = GetValue(Feld(W_YA))
              Tra_ZA(AnzPkt)      = GetValue(Feld(W_ZA))
              
              Tra_STB1(AnzPkt)    = GetValue(Feld(W_STB1))
              Tra_STB2(AnzPkt)    = GetValue(Feld(W_STB2))
              Tra_STB3(AnzPkt)    = GetValue(Feld(W_STB3))
              Tra_STB4(AnzPkt)    = GetValue(Feld(W_STB4))
              Tra_STB5(AnzPkt)    = GetValue(Feld(W_STB5))
              Tra_STB6(AnzPkt)    = GetValue(Feld(W_STB6))
              Tra_STB7(AnzPkt)    = GetValue(Feld(W_STB7))
              Tra_STB8(AnzPkt)    = GetValue(Feld(W_STB8))
              Tra_STB9(AnzPkt)    = GetValue(Feld(W_STB9))
              
              TK_KmStatus(AnzPkt) = GetValue(Feld(W_KmStatus))
              TK_KmText(AnzPkt)   = GetValue(Feld(W_KmText))
              
              DGM_HDGM(AnzPkt)    = GetValue(Feld(W_HDGM))
              DGM_ZDGM(AnzPkt)    = GetValue(Feld(W_ZDGM))
              
              if (isNumeric(Feld(W_Ri))) then
                Ri_tmp = Feld(W_Ri) * rho
                if (Ri_tmp < 0) then Ri_tmp = Ri_tmp + 400
                Ri(AnzPkt) = Ri_tmp
              else
                Ri(AnzPkt) = ""
              end if
              
              if (isNumeric(Feld(W_Ueb))) then
                u(AnzPkt) = Feld(W_Ueb) * F_Ueb
              else
                u(AnzPkt) = ""
              end if
              
            On Error GoTo Fehler
            
            'Werte im Gleissystem: nicht hier ermitteln => wird nachträglich gemacht.
            
            'Namen der diversen Bänder
              N_Acfg(AnzPkt)    = NAcfg
              N_KmLinie(AnzPkt) = NKML
              N_Trasse(AnzPkt)  = NTra
              N_Gra(AnzPkt)     = NGra
              N_UebBand(AnzPkt) = NUeb
              N_RegPrf(AnzPkt)  = NRegel
              N_TunPrf(AnzPkt)  = NTunnel
              N_PrfPkt(AnzPkt)  = NPrfPkt
              N_GlsPrf(AnzPkt)  = NGlsPrf
              N_DGM(AnzPkt)     = NDGM
            
            'Variablen zurücksetzen
              vGradHoehe = ""
              vdH = ""
              vY = ""
              vX = ""
              vZ = ""
              Kommentar = ""
              strBemerkung = ""
          End If
        End If
      End If
      
    Loop While (Not UmfEnde)
    
    'Dokumentation der Trassenbandnamen (auch bei Nicht-Erfolg)
    if (NAcfg   <> "") then Echo "   - Trassenkonfig: " & NAcfg
                            Echo "   - Achse:         " & NTra
    if (NKML    <> "") then Echo "   - KM-Linie:      " & NKML
    if (NGra    <> "") then Echo "   - Gradiente:     " & NGra
    if (NUeb    <> "") then Echo "   - Überhöhung:    " & NUeb
    if (NRegel  <> "") then Echo "   - Regelprofile:  " & NRegel
    if (NTunnel <> "") then Echo "   - Tunnelprofile: " & NTunnel
    if (NPrfPkt <> "") then Echo "   - Profilpunkte:  " & NPrfPkt
    if (NGlsPrf <> "") then Echo "   - Gleisprofile:  " & NGlsPrf
    if (NDGM    <> "") then Echo "   - DGM:           " & NDGM
  end if
  
  'Punktanzahl ausgeben, markieren, falls = 0.
  if (lfdNr = 0) then
    Erfolg = false
    WarnEcho " *** Punktanzahl:   0"
  else
    Erfolg = true
    Echo "   - Punktanzahl:   " & lfdNr
  end if
  
  'Fehleranzahl nur ausgeben, falls Fehler aufgetreten sind.
  if ((AnzErr - AnzErr0) > 0) then ErrEcho " *** Fehleranzahl: " & cStr(AnzErr - AnzErr0)
  
  Set oFSO = nothing
  
  LeseTKiGeo = Erfolg
  Exit Function
  
Fehler:
  lfdNr = lfdNr - 0
  AnzPkt = AnzPkt - 1
  LeseTKiGeo = false
  ErrMessage = "Fehlerhaftes Format in Zeile " & ZeilenNr
  FehlerNachricht "CimpTrassenkoo.LeseTKiGeo()"
End Function



Private Function LeseTkAdelt(ByVal Kanal As Integer, ByRef AnzPkt As Long, ByRef AnzErr As Long, ByVal TkFormat As String) As Boolean
  ' -----------------------------------------------------------------------------------------------
  ' Liest eine Verm.esn-Umformung aus der aktiven Datei.
  ' -----------------------------------------------------------------------------------------------
  ' Eingabe:   Kanal    ... Dateihandle
  '            AnzPkt   ... Anzahl erfolgreich gelesener Punkte, wird weiter gezählt.
  '            AnzErr   ... Anzahl erkannter Fehler, wird weiter gezählt.
  '            TkFormat ... wird nicht ausgewertet
  'Rückgabe: - true,  falls mindestens ein Punkt gelesen wurde
  '          - False, falls kein Punkt gelesen wurde
   '
   ' Status nach Beendigung dieser Funktion:
   '  - Es ist bereits die Überschrift der nächsten Adelt-Ausgabe derselben Datei gelesen, oder:
   '  - Dateiende ist erreicht.
   '
  ' ==> unterstützte Adelt-Versionen: 6.62 (DOS), 8.40 (Windows)
  ' ==> Ab Spalte 83/84 kann eine Bemerkung stehen (max. 13-stellig).
  '     (Verm.esn 200x schreibt an diese Stelle das Feld "Pkt. Erl.Lage".)
  ' -----------------------------------------------------------------------------------------------
  
  On Error GoTo Fehler
  
  'Deklarationen
    Dim NKML          As String
    Dim NTra          As String
    Dim NGra          As String
    Dim NUeb          As String
    Dim TkHerkunft    As String
    Dim Format_ok     As Boolean
    Dim UmfEnde       As Boolean
    Dim Erfolg        As Boolean
    Dim PktNr_ok      As Boolean
    Dim is2Zeilig     As Boolean
          
    Dim lfdNr         As Long
    Dim AnzErr0       As Long
    Dim so1           As Variant
    Dim so2           As Variant
    Dim Pnr_Faktor    As Variant
    Dim PktNr_TK      As Variant
    Dim PktNr_GK      As Variant
    
    Dim vGradHoehe    As Variant
    Dim vdH           As Variant
    Dim vY            As Variant
    Dim vX            As Variant
    Dim vZ            As Variant
    Dim vSt           As Variant
    Dim vKm           As Variant
    Dim vQKm          As Variant
    Dim Km_tmp        As Variant
    
    Dim strBemerkung  As String
    Dim ErlLage       As String
    Dim ZeileOhneBemerkung As String
    
    Dim Key           As String
    Dim Wert          As String
    Dim oMatches      As Object
    Dim oFSO          As New Scripting.FileSystemObject
    
  'Variableninitialisierung
    Const VE_CLASSIC   = "Verm.esn/Klassisch"
    Const VE_3LINIEN   = "Verm.esn/3-Linien"
    Const TkEntstehung = "Umformung"
    
    AnzErr0 = AnzErr      'Anzahl Fehler zu Beginn dieser Routine
    lfdNr   = 0
    
    NTra       = ""
    NKML       = ""
    NGra       = ""
    NUeb       = ""
    
    vGradHoehe = ""
    vdH        = ""
    vY         = ""
    vX         = ""
    vZ         = ""
    ErlLage    = ""
    vSt        = ""
    
    TkHerkunft = ""
    Format_ok  = false
    UmfEnde    = false
    Erfolg     = true
    
  'Startzeile der Umformung => Name der Trasse sowie Verm.esn-interne Herkunft ermitteln.
    'Verm.esn-interne Herkunft ermitteln
      if (Feld(1) = "Trassenumformung") then
        TkHerkunft = VE_CLASSIC
      elseif (Feld(1) = "Umformung") then
        TkHerkunft = VE_3LINIEN
      end if
    
    'Achsname ermitteln
      ThisWorkbook.RegExp.IgnoreCase = True
      ThisWorkbook.RegExp.Global = False
      'ThisWorkbook.RegExp.Pattern = "umformung +(.+?)(\.tra)?[ \t]+\("
      'VB.NET: Pattern = "^Trassenumformung\s+(.+?\.tra)\s+\("
      ThisWorkbook.RegExp.Pattern = "umformung +(.+?\.tra)[ \t]+\("
      Set oMatches = ThisWorkbook.RegExp.Execute(Zeile)
      If (oMatches.Count > 0) Then
        NTra = oMatches(0).SubMatches(0)
      else
        'VB.NET: Pattern = "^Trassenumformung\s+(.+?)\s+\("
        ThisWorkbook.RegExp.Pattern = "umformung +(.+?)[ \t]+\("
        Set oMatches = ThisWorkbook.RegExp.Execute(Zeile)
        If (oMatches.Count > 0) Then
          NTra = oMatches(0).SubMatches(0)
        else
          'sicher ist sicher ;-)
          NTra = Feld(2)
        end If
      End If
      debugecho "LeseTkAdelt(): Trassenname komplett = '" & NTra & "'"
      'NTra = oFSO.GetBaseName(NTra)
      NTra = oFSO.GetFileName(NTra)
      debugecho "LeseTkAdelt(): Trassenname kurz = '" & NTra & "'"
          
    'Erfolgskontrolle. Bei Mißerfolg => Abbruch!.
    if ((TkHerkunft = "") or (NTra = "")) then
      'Abbruch
      UmfEnde = true
      Erfolg  = false
      AnzErr  = AnzErr + 1
      ErrEcho " *** Fehler:       in Zeile " & cStr(ZeilenNr) & " der Eingabedatei: Name der Trasse oder Herkunft (Klassisch/D3L) konnte nicht ermittelt werden."
      call ThisWorkbook.SysTools.FileErrorsAdd("F", Me.Quelle_AsciiDatei_Name, ZeilenNr, 0, 0, "Fehler: Name der Trasse oder Herkunft (Klassisch/D3L) konnte nicht ermittelt werden", _
                                   "Es wurde der Beginn einer Verm.esn-Umformung erkannt, aber deren Entstehung oder der Trassenname konnten nicht ermittelt werden." & vbNewLine & _
                                   "Quelle dieser Meldung: Einlesen einer Datei mit Trassenkoordinaten aus Verm.esn (CimpTrassenkoo.LeseTkAdelt())", _
                                   blnFehlerKonsolidieren)
    else
      Echo "   - Herkunft:     " & TkHerkunft & "."
    end if
    
  '
  'Gesamte Umformung ab 2. Zeile einlesen
  Do While (Not UmfEnde)
    'Nächste Zeile lesen.
    'Falls aber Dateiende erreicht ist => Schleife verlassen.
    if (EOF(Kanal)) then
      exit do
    else
      Line Input #Kanal, Zeile
      ProgressbarDateiLesen (Kanal)
      ZeilenNr = ZeilenNr + 1
      NF = splitWords(Zeile, Feld, "")               'awk-like Splitting
    end if
    
    'Weiter, wenn keine Leerzeile.
    If (NF > 0) Then
      'Ab 2. Zeile der Umformung: Name der Trasse und Verm.esn-interne Herkunft sind bereits bekannt.
      
      'Ende der Umformung erkennen (unabhängig davon, ob bereits Punkte gefunden worden sind!).
      Select Case Feld(1)
        Case "Trasse", "Gradiente", "Trassenumformung", "Umformung", "Koordinatenverzeichnis"
          UmfEnde = true
          DebugEcho "LeseTkAdelt(): 1. Wort=" & Feld(1) & ", Ende der Umformung=" & cStr(UmfEnde)
          
        Case Else
          
          If (Not entspricht("^[0-9]", Feld(1))) Then
            'Keine Zeile mit Umformungsdaten, d.h. normalerweise Kopf
            If (Not Format_ok) Then
              'Kopf, da noch kein Punkt gelesen.
              ThisWorkbook.RegExp.IgnoreCase = True
              ThisWorkbook.RegExp.Global = False
              ThisWorkbook.RegExp.Pattern = "\s*(.*?)\s*:\s*(\S*.*\S+)\s*"
              Set oMatches = ThisWorkbook.RegExp.Execute(Zeile)
              If (oMatches.Count > 0) Then
                Key  = oMatches(0).SubMatches(0)
                Wert = oMatches(0).SubMatches(1)
                Select Case Key
                  Case "Gradiente":  NGra = oFSO.GetBaseName(Wert)
                  Case "Reduktion auf Trasse":  NKML = oFSO.GetBaseName(Wert)
                End Select
              End If
              DebugEcho "LeseTkAdelt(): Kopfzeile ausgewertet - aktuelle Namen: Gradiente=" & NGra & ", KML=" & NKML
            End If
            
          Else
            'Zeile mit Umformungsdaten, bei zweizeiliger Umformung eine der beiden Zeilen!
            
            'Erster Punkt der Umformung.
            If (Not Format_ok) Then
              'Erkennung des Pnr-Formates
              If (entspricht("\.[0-9][0-9][0-9][0-9][0-9]$", Feld(1))) Then
                '5-stellige Pnr (ab Adelt v8.30)
                Pnr_Faktor = 100000
                so1 = 1                  'Spalten-Offset
                if (TkHerkunft = VE_CLASSIC) then so2 = 2 else so2 = 6
                TkFormat = "PktNr 7-stellig"
              Else
                '4-stellige Pnr (bis Adelt v7.59)
                Pnr_Faktor = 10000
                so1 = 0                  'Spalten-Offset
                if (TkHerkunft = VE_CLASSIC) then so2 = 0 else so2 = 2
                TkFormat = "PktNr 6-stellig"
              End If
            End If
            
            '"Reine" Datenzeile erzeugen
            ZeileOhneBemerkung = mid(zeile, 1, 80 + so1)
            
            'Großkoordinaten (und Erläuterung/DOS)
            if (((TkHerkunft = VE_CLASSIC) And (InStr(1, ZeileOhneBemerkung, "+", vbTextCompare) > 0)) or _
                ((TkHerkunft = VE_3LINIEN) And (InStr(1, ZeileOhneBemerkung, "+", vbTextCompare) < 1))) then
              'Großkoordinaten
              
              'Erster Punkt einer zweizeiligen Umformung.
              If (not Format_ok) Then
                TkFormat  = TkFormat & ", 2-zeilig"
                is2Zeilig = true
                Format_ok = true
                Echo "   - Format:       " & TkFormat
              end if
              
              'Zeile auswerten.
              if (not is2Zeilig) then
                AnzErr = AnzErr + 1
                Erfolg = false
                ErrEcho " *** Fehler:       in Zeile " & cStr(ZeilenNr) & " der Eingabedatei: Unerwartetes Auftreten einer Koordinatenzeile in einer als 'einzeilig' erkannten Umformung."
                call ThisWorkbook.SysTools.FileErrorsAdd("F", Me.Quelle_AsciiDatei_Name, ZeilenNr, 0, 0, "Fehler: Unerwartetes Auftreten einer Koordinatenzeile in einer als 'einzeilig' erkannten Umformung.", _
                                             "Quelle dieser Meldung: Einlesen einer Datei mit Trassenkoordinaten aus Verm.esn (CimpTrassenkoo.LeseTkAdelt())", _
                                             blnFehlerKonsolidieren)
              else
                'DebugEcho "LeseTkAdelt(): Zeile mit Großkoordinaten."
                PktNr_GK = cStr(cdbl(Trim(Mid(Zeile, 1, 7 + so1))) * Pnr_Faktor)
                vY = Trim$(Mid$(Zeile,  8 +   so1, 13 + so1))
                vX = Trim$(Mid$(Zeile, 21 + 2*so1, 13 + so1))
                vZ = Trim(Mid(Zeile,   52 + 3*so1,  9))
                if (TkHerkunft = VE_CLASSIC) then
                  ErlLage = mid$(zeile, 62 + 3*so1, 13)
                else
                  ErlLage = mid$(zeile, 62 + 2*so1, 13)
                  vSt = mid(zeile, 34 + 3*so1, 16)        'Gleisstation 3/L
                end if
              end if
            Else
              'Zeile mit Trassenkoordinaten.
              
              'DebugEcho "LeseTkAdelt(): Zeile mit Trassenkoordinaten."
              'Erster Punkt  einer einzeiligen Umformung: Erkennung des Datensatz-Formates.
              if (not Format_ok) Then
                TkFormat  = TkFormat & ", 1-zeilig"
                is2Zeilig = false
                Format_ok = true
                Echo "   - Format:       " & TkFormat
              end if
              
              'Punktnummer ggf. prüfen
              PktNr_ok = true
              PktNr_TK = cStr(cdbl(Trim(Mid(Zeile, 1, 7 + so1))) * Pnr_Faktor)
              
              if (is2Zeilig) then
                if (not (PktNr_TK = PktNr_GK)) then
                  AnzErr   = AnzErr + 1
                  Erfolg   = false
                  PktNr_ok = false
                  ErrEcho " *** Fehler:       in Zeile " & cStr(ZeilenNr) & " der Eingabedatei: Die Punktnummern des zweizeiligen Datensatzes stimmen nicht überein (PktNr/TK='" & cStr(PktNr_TK) & "' PktNr/GK='" & cStr(PktNr_GK) & "')."
                  call ThisWorkbook.SysTools.FileErrorsAdd("F", Me.Quelle_AsciiDatei_Name, ZeilenNr, 0, 7 + so1, "Fehler: Die Punktnummern des zweizeiligen Datensatzes stimmen nicht überein (PktNr/TK='" & cStr(PktNr_TK) & "' PktNr/GK='" & cStr(PktNr_GK) & "').", _
                                               "Quelle dieser Meldung: Einlesen einer Datei mit Trassenkoordinaten aus Verm.esn (CimpTrassenkoo.LeseTkAdelt())", _
                                               blnFehlerKonsolidieren)
                end if
              end if
              
              'PktNr-Prüfung erfolgreich
              if (PktNr_ok) then
                'Zähler
                  lfdNr = lfdNr + 1                        'Punktanzahl in dieser Umformung
                  AnzPkt = AnzPkt + 1              'Punktanzahl gesamt
                  
                'Spalten um je 1 Zeile verlängern
                  ReDim Preserve N_KmLinie(1 To AnzPkt)
                  ReDim Preserve N_Trasse(1 To AnzPkt)
                  ReDim Preserve N_Gra(1 To AnzPkt)
                  ReDim Preserve Abst(1 To AnzPkt)
                  ReDim Preserve AbstKm(1 To AnzPkt)
                  'ReDim Preserve AbstRed(1 To AnzPkt)
                  ReDim Preserve Bemerkung(1 To AnzPkt)
                  ReDim Preserve GradHoehe(1 To AnzPkt)
                  ReDim Preserve PktNr(1 To AnzPkt)
                  ReDim Preserve Radius(1 To AnzPkt)
                  ReDim Preserve Ri(1 To AnzPkt)
                  ReDim Preserve dH(1 To AnzPkt)
                  'ReDim Preserve dHRed(1 To AnzPkt)
                  ReDim Preserve stat(1 To AnzPkt)
                  ReDim Preserve Km(1 To AnzPkt)
                  ReDim Preserve u(1 To AnzPkt)
                  'ReDim Preserve u_ist(1 To AnzPkt)
                  ReDim Preserve X(1 To AnzPkt)
                  ReDim Preserve Y(1 To AnzPkt)
                  ReDim Preserve Z(1 To AnzPkt)
                  
                'Höhenwerte lesen bzw. berechnen
                  vGradHoehe = Trim(Mid(Zeile, 56 + so2, 8))
                  vdH        = Trim(Mid(Zeile, 64 + so2, 8))
                  
                  'Fehlende Höhenwerte berechnen, wenn möglich.
                  on error resume next
                  If (vdH <> "") Then vdH = vdH * (-1)
                  if (err.number <> 0) then
                    AnzErr = AnzErr + 1
                    Erfolg = false
                    err.clear
                    ErrEcho " *** Fehler:       in Zeile " & cStr(ZeilenNr) & " der Eingabedatei: dH='" & cStr(vdH) & "'."
                    call ThisWorkbook.SysTools.FileErrorsAdd("f", Me.Quelle_AsciiDatei_Name, ZeilenNr, 0, 0, "Fehler: Der Höhenunterschied (dH='" & cStr(vdH) & "') ist kein numerischer Wert.", _
                                                 "Quelle dieser Meldung: Einlesen einer Datei mit Trassenkoordinaten aus Verm.esn (CimpTrassenkoo.LeseTkAdelt())", blnFehlerKonsolidieren)
                  else
                    If ((vZ = "") And (vdH <> "") And (vGradHoehe <> "")) Then vZ = cdbl(vGradHoehe) + vdH
                    if (err.number <> 0) then
                      AnzErr = AnzErr + 1
                      Erfolg = false
                      err.clear
                      ErrEcho " *** Fehler:       in Zeile " & cStr(ZeilenNr) & " der Eingabedatei: SO-Höhe='" & cStr(vGradHoehe) & "'."
                      call ThisWorkbook.SysTools.FileErrorsAdd("f", Me.Quelle_AsciiDatei_Name, ZeilenNr, 0, 0, "Fehler: Die SO-Höhe (SO='" & cStr(vGradHoehe) & "') ist kein numerischer Wert.", _
                                                   "Quelle dieser Meldung: Einlesen einer Datei mit Trassenkoordinaten aus Verm.esn (CimpTrassenkoo.LeseTkAdelt())", blnFehlerKonsolidieren)
                    end if
                  end if
                  on error goto 0
                  
                'Kilometer und Station
                  vQKm = trim(mid(Zeile, 73 + so2, 7))   'kann auch "außerh." sein. ==> Konsequenzen ???
                  
                  if (TkHerkunft = VE_CLASSIC) then
                    Km_tmp = mid(zeile, 8 + so1, 12)
                  else
                    Km_tmp = GetKm(mid(zeile, 8 + so1, 15))
                  end if
                  if (vQKm = "") then
                    ' Station bezieht sich nicht auf KM-Linie.
                    ' (Verwendung der KML wird nicht immer durch einen Hinweis im Kopf bekannt gemacht).
                    ' => Station wird auch als Kilometer eingetragen, wenn gewünscht.
                    vSt = Km_tmp
                    'if (StAlsKm) then  Km = St  else  Km = ""
                  else
                    vKm = Km_tmp
                  end if
                  
                'Bemerkung
                  strBemerkung = trim(mid(zeile, 81 + so2, 15))
                  if (strBemerkung = "") then strBemerkung = ErlLage
                  strBemerkung = trim(strBemerkung)
                  
                'Ist-Überhöhung in (mm) aus Bemerkung ermitteln => Wird nach dem Import erledigt!
                  'vu_ist = ""
                  'vu_ist = Ueberhoehung(Bemerkung, UebInInfo_Streng)     # ist-Überhöhung in [mm] aus Bemerkung ermitteln
                  
                'Felder belegen
                  PktNr(AnzPkt) = PktNr_TK
                  Y(AnzPkt) = vY
                  X(AnzPkt) = vX
                  Z(AnzPkt) = vZ
                  Bemerkung(AnzPkt) = strBemerkung
                  
                  Km(AnzPkt) = vKm
                  stat(AnzPkt) = vSt
                  AbstKm(AnzPkt) = vQKm
                  dH(AnzPkt) = vdH
                  'u_ist(AnzPkt) = vu_ist
                  
                  GradHoehe(AnzPkt) = vGradHoehe
                  
                  if (TkHerkunft = VE_CLASSIC) then
                    'Tra/Gra klassisch
                    Abst(AnzPkt)   = Trim$(replace(Mid$(Zeile, 20 + so1, 9), "*", ""))
                    Ri(AnzPkt)     = Trim$(Mid$(Zeile, 29 + so1, 10 + so1))
                    Radius(AnzPkt) = Trim$(Mid$(Zeile, 39 + so2, 11))
                    u(AnzPkt)    = val(Trim(Mid(Zeile, 50 + so2, 6)))       'in [mm]
                    
                  else
                    '3 Linien
                    if (Pnr_Faktor = 10000) then
                      'DOS
                      Abst(AnzPkt)   = Trim$(replace(Mid$(zeile, 23, 9), "*", ""))
                      Ri(AnzPkt)     = Trim$(Mid$(zeile, 32,  9))
                      Radius(AnzPkt) = Trim$(Mid$(zeile, 41, 11))
                      u(AnzPkt)    = val(Trim(Mid(zeile, 52,  6)))       'in [mm]
                      
                    else
                      'Windows/200x
                      Abst(AnzPkt)   = Trim$(replace(Mid$(zeile, 24, 9), "*", ""))
                      Ri(AnzPkt)     = Trim$(Mid$(zeile, 33, 11))
                      Radius(AnzPkt) = Trim$(Mid$(zeile, 44, 11))
                      u(AnzPkt)    = val(Trim(Mid(zeile, 55,  7)))       'in [mm]
                    end if
                  end if
                  
                'Koordinatenumformung Trassenkoo' ==> Gleissystem. => Wird nach dem Import erledigt!
                  'Call Transfo_Tk2Gls(CVar(Radius(AnzPkt)), CVar(u(AnzPkt) / 1000), CVar(Abst(AnzPkt)), vdH, AbstRed(AnzPkt), dHRed(AnzPkt))
                  
                'Namen der diversen Bänder
                  N_KmLinie(AnzPkt) = NKML
                  N_Trasse(AnzPkt)  = NTra
                  N_Gra(AnzPkt)     = NGra
                  'N_UebBand(AnzPkt) = NUeb
                  'N_DGM(AnzPkt)     = NDGM
                  
                'Initialwerte für nächsten Punkt
                vGradHoehe = ""
                vdH = ""
                vY = ""
                vX = ""
                vZ = ""
                strBemerkung = ""
                ErlLage = ""
                vSt = ""
              End If
            End If
          End If
      End Select
    End If
  Loop
  
  'Dokumentation der Trassenbandnamen (auch bei Nicht-Erfolg)
                       Echo "   - Achse:        " & NTra
  if (NKML <> "") then Echo "   - KM-Linie:     " & NKML
  if (NGra <> "") then Echo "   - Gradiente:    " & NGra
  if (NUeb <> "") then Echo "   - Überhöhung:   " & NUeb
  
  'Punktanzahl ausgeben, markieren, falls = 0.
  if (lfdNr = 0) then
    Erfolg = false
    WarnEcho " *** Punktanzahl:  0"
  else
    Erfolg = true
    Echo "   - Punktanzahl:  " & lfdNr
  end if
  
  'Fehleranzahl nur ausgeben, falls Fehler aufgetreten sind.
  if ((AnzErr - AnzErr0) > 0) then ErrEcho " *** Fehleranzahl: " & cStr(AnzErr - AnzErr0)
  
  Set oFSO = nothing
  
  LeseTkAdelt = Erfolg
  Exit Function
  
Fehler:
  lfdNr = lfdNr - 0
  AnzPkt = AnzPkt - 1
  LeseTkAdelt = false
  ErrMessage = "Fehlerhaftes Format in Zeile " & ZeilenNr
  FehlerNachricht "CimpTrassenkoo.LeseTkAdelt()"
End Function

Private Function GetValue(InputString As String)
  ' Rückgabe: "", falls InputString = "NaN", sonst InputString
  On Error GoTo Fehler
  Dim RetValue As String
  If (InputString = "NaN") Then
    RetValue = ""
  Else
    RetValue = InputString
  End If
  GetValue = RetValue
  Exit Function
  
Fehler:
  ErrMessage = ""
  FehlerNachricht "CimpTrassenkoo.GetValue()"
End Function


'für jEdit:  :folding=indent::collapseFolds=1:
