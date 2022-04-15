VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CtabCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2009-2022  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'====================================================================================
'Klasse CtabCSV
'====================================================================================
'Verarbeitung einer speziellen CSV-Datei zwecks Import.

Option Explicit

'Modulkonstanten
Const strKlasse_Name                    As String = "CtabCSV"

Const strTitel                          As String = "CSV spezial"
Const strQuelle_AsciiDatei_DialogFilter As String = "Komma-getrennt (*.csv),*.csv, Alle Dateien (*.*),*.*"
Const strZiel_Typ                       As String = io_Typ_Puffer

Const blnFehlerKonsolidieren            As Boolean = True

'Feldnamen (ZellNamen in Tabelle = Keys des Dictionary)
Const FN_Statistik_Importdatei          As String = "x.Statistik.Importdatei"


'Datenbereich
Dim aktTabZeile                         As Long
Dim lfdNrTotal                          As Long

'Eigenschaften
Dim strQuelle_AsciiDatei_Name           As String
Dim strKategorien                       As String

Dim strMeldung_Titel                    As String
Dim strMeldung_Inhalt                   As String
Dim blnAktionErfolgreich                As Boolean
Dim lngFehlerniveau                     As Long

'math. Konstanten
Dim pi                                  As Double
Dim rho                                 As Double

Dim Kanal                               As Integer
Dim NR                                  As Long       'Nr. der zuletzt eingelesenen Zeile der Eingabedatei
Dim NF                                  As Long       'Anzahl Felder der aktuellen Zeile
Dim Zeile                               As String     'gesamte aktuelle Zeile, die zuletzt eingelesen wurde
Dim Feld()                              As String     'einzelne Felder der aktuellen Zeile (1. Index = 1)
Dim datCSV()                            As String
Dim Vektor()                            As String

'lokale Metadaten
Dim oMetadatenExtra                     As Scripting.Dictionary
Dim oMetadatenProjekt                   As Scripting.Dictionary

'Spaltenbeschreibungen:  Key=Spaltenname ohne Prefix und Einheit
Dim oQuellDatenpuffer                   As Scripting.Dictionary       'Item=Datenvektor
Dim oQuellSpaltenPositionen             As Scripting.Dictionary       'lfd Feldnr. der Spalte in der CSV-Datei
Dim oQuellEinheiten                     As Scripting.Dictionary       'Name der Einheit
'Dim oQuellFormate                       As Scripting.Dictionary       'Formatangabe
Dim QuellSpaltenNamen()                 As String                     'Index ab 1 = Position

Dim oZielEinheiten                      As Scripting.Dictionary       'Name der Einheit
Dim oZielFormate                        As Scripting.Dictionary       'Formatangabe

Dim oCsvSettings                        As Scripting.Dictionary
Dim oExcelTableProps                    As Scripting.Dictionary
Dim oDataImportOptions                  As Scripting.Dictionary
Dim oCsvSettingsTyp                     As Scripting.Dictionary
Dim oExcelTablePropsTyp                 As Scripting.Dictionary
Dim oDataImportOptionsTyp               As Scripting.Dictionary
'



Private Sub Class_Initialize()
  On Error GoTo Fehler
  DebugEcho "CtabCSV.Class_Initialize(): Initialisierung beginnt."
  
  Set oMetadatenExtra = New Scripting.Dictionary
  Set oMetadatenProjekt = New Scripting.Dictionary
  
  Set oQuellDatenpuffer = New Scripting.Dictionary
  Set oQuellSpaltenPositionen = New Scripting.Dictionary
  Set oQuellEinheiten = New Scripting.Dictionary
  'Set oQuellFormate = New Scripting.Dictionary
  
  Set oZielEinheiten = New Scripting.Dictionary
  Set oZielFormate = New Scripting.Dictionary
  
  Set oCsvSettings = New Scripting.Dictionary
  Set oExcelTableProps = New Scripting.Dictionary
  Set oDataImportOptions = New Scripting.Dictionary
  
  Set oCsvSettingsTyp = New Scripting.Dictionary
  Set oExcelTablePropsTyp = New Scripting.Dictionary
  Set oDataImportOptionsTyp = New Scripting.Dictionary
  
  
  'Mathematische Konstanten
  pi = 4 * Atn(1)
  rho = 200 / pi
  
  'Startwerte setzen
  oCsvSettings.add "FieldSeparator",   CsvTrenner_Std
  'oCsvSettings.add "DecimalSeparator", CsvDezimalTrenner_Std
  oCsvSettings.add "TextQualifier",    CsvTextQualifier_Std
  oCsvSettings.add "TrimFields",       CsvTrimFields_Std
  
  oExcelTableProps.add "Class",    ""
  oExcelTableProps.add "AutoSave", false
  
  'Standard für Importoptionen ist "nicht gesetzt"
  '=> dann treten die Standardwerte wie bei interaktivem Import in Kraft.
  oDataImportOptions.add "UseBackupColumns", null
  oDataImportOptions.add "EditData", null
  oDataImportOptions.add "EditOpt_OverwriteExistingValues", null
  oDataImportOptions.add "EditModule_SlantFromInfo", null
  oDataImportOptions.add "EditModule_Parse_iTC", null
  oDataImportOptions.add "EditModule_Differences", null
  oDataImportOptions.add "EditModule_RailCooTransfo", null
  
  
  'Typen deklarieren
  oCsvSettingsTyp.add "FieldSeparator",   vbString
 'oCsvSettingsTyp.add "DecimalSeparator", vbString
  oCsvSettingsTyp.add "TextQualifier",    vbString
  oCsvSettingsTyp.add "TrimFields",       vbBoolean
  
  oExcelTablePropsTyp.add "Class",    vbString
  oExcelTablePropsTyp.add "AutoSave", vbBoolean
  
  oDataImportOptionsTyp.add "UseBackupColumns", vbBoolean
  oDataImportOptionsTyp.add "EditData", vbBoolean
  oDataImportOptionsTyp.add "EditOpt_OverwriteExistingValues", vbBoolean
  oDataImportOptionsTyp.add "EditModule_SlantFromInfo", vbBoolean
  oDataImportOptionsTyp.add "EditModule_Parse_iTC", vbBoolean
  oDataImportOptionsTyp.add "EditModule_Differences", vbBoolean
  oDataImportOptionsTyp.add "EditModule_RailCooTransfo", vbBoolean
  
  
  Me.Quelle_AsciiDatei_Name = ""
  NR = 0                            'Zähler für Gesamtzahl gelesener Zeilen
  
  DebugEcho "CtabCSV.Class_Initialize(): Initialisierung beendet."
  Exit Sub
  
Fehler:
  FehlerNachricht "CtabCSV.Class_Initialize()"
  Call CloseCSV()
  Set oCsvSettings = Nothing
  Set oExcelTableProps = Nothing
  Set oDataImportOptions = Nothing
  Set oMetadatenExtra = Nothing
  Set oMetadatenProjekt = Nothing
  Set oQuellDatenpuffer = Nothing
  Set oQuellEinheiten = Nothing
  'Set oQuellFormate = Nothing
  Set oQuellSpaltenPositionen = Nothing
  Set oZielEinheiten = Nothing
  Set oZielFormate = Nothing
End Sub


Private Sub Class_Terminate()
  On Error GoTo Fehler
  Call CloseCSV()
  Set oCsvSettings = Nothing
  Set oExcelTableProps = Nothing
  Set oDataImportOptions = Nothing
  Set oMetadatenExtra = Nothing
  Set oMetadatenProjekt = Nothing
  Set oQuellDatenpuffer = Nothing
  Set oQuellEinheiten = Nothing
  'Set oQuellFormate = Nothing
  Set oQuellSpaltenPositionen = Nothing
  Set oZielEinheiten = Nothing
  Set oZielFormate = Nothing
  DebugEcho "CtabCSV.Class_Terminate(): fertig."
  Exit Sub
  
Fehler:
  FehlerNachricht "CtabCSV.Class_Terminate()"
End Sub




'Eigenschaften *******************************************************************************

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
  'Setzt den Namen der CSV-Datei und syncronisiert einige Eigenschaften mit dieser Datei.
  strQuelle_AsciiDatei_Name = trim(inpQuelle_AsciiDatei_Name)
  if (strQuelle_AsciiDatei_Name <> "") then
    DebugEcho "CtabCSV.Quelle_AsciiDatei_Name [Let]: CSV-Dateinamen gesetzt auf: '" & inpQuelle_AsciiDatei_Name & "'."
    Call Syncronisieren
  end if
End Property


Property Get Quelle_Datenpuffer() As Scripting.Dictionary
  'Bietet Zugriff auf den Quelldatenpuffer.
  Set Quelle_Datenpuffer = oQuellDatenpuffer
End Property


Property Get Quelle_DatenExtra() As Scripting.Dictionary
  'Bietet Zugriff auf Extradaten wie Statistik, Parameter... für Tabellenkopf.
  Set Quelle_DatenExtra = oMetadatenExtra
End Property


Property Get Quelle_Projektdaten() As Scripting.Dictionary
  'Bietet Zugriff auf Projektdaten.
  Set Quelle_Projektdaten = oMetadatenProjekt
End Property


Property Get Quelle_SpaltenPositionen() As Scripting.Dictionary
  'Bietet Zugriff auf die SpaltenPositionen der Quelldaten.
  Set Quelle_SpaltenPositionen = oQuellSpaltenPositionen
End Property


Property Get Quelle_SpaltenNamen() As Variant
  'Liefert das Array der Spaltennamen der Quelldaten.
  'Der erste Index = 1. Die Indizes entsprechen der Spaltenposition in der CSV.
  Quelle_SpaltenNamen = QuellSpaltenNamen
End Property


Property Get Quelle_Einheiten() As Scripting.Dictionary
  'Bietet Zugriff auf die Einheiten der Quelldaten.
  Set Quelle_Einheiten = oQuellEinheiten
End Property


'Property Get Quelle_Formate() As Scripting.Dictionary
'  'Bietet Zugriff auf die Formate der Quelldaten.
'  Set Quelle_Formate = oQuellFormate
'End Property


Property Get Ziel_Einheiten() As Scripting.Dictionary
  'Bietet Zugriff auf die Einheiten für die Exceltabelle.
  Set Ziel_Einheiten = oZielEinheiten
End Property


Property Get Ziel_Formate() As Scripting.Dictionary
  'Bietet Zugriff auf die Formate für die Exceltabelle.
  Set Ziel_Formate = oZielFormate
End Property


Property Get Ziel_Typ() As String
  'Liefert den Ziel-Typ der importierten Daten.
  Ziel_Typ = strZiel_Typ
End Property


Property Get Ziel_XLMappeSpeichern() As Boolean
  'Liefert den aktuellen Modus für "neu angelegte Mappe speichern".
  Ziel_XLMappeSpeichern = oExcelTableProps("AutoSave")
End Property


Property Get Ziel_XLTabKlasse() As String
  'Liefert den Namen der XL-Tabellen-Klasse (Kodename) für Me.Import
  Ziel_XLTabKlasse = oExcelTableProps("Class")
End Property


'Optionen für den Import.
' - Die (fehlertolerante) Konvertierung wird von CdatExpim übernommen.
' - Als 'JA' gelten:   'true', alle Zahlen außer '0'
' - Als 'NEIN' gelten: 'false', Zahl '0'
' - Als 'Nicht festgelegt' gelten: Alle anderen, nicht konvertierbaren Werte.
Property Get Opt_ErsatzZielspaltenVerwenden() As Variant
  'Liefert die entsprechende Import-Option.
  Opt_ErsatzZielspaltenVerwenden = oDataImportOptions("UseBackupColumns")
End Property

Property Get Opt_DatenModifizieren() As Variant
  'Liefert die entsprechende Import-Option.
  Opt_DatenModifizieren = oDataImportOptions("EditData")
End Property

Property Get Opt_VorhWerteUeberschreiben() As Variant
  'Liefert die entsprechende Import-Option.
  Opt_VorhWerteUeberschreiben = oDataImportOptions("EditOpt_OverwriteExistingValues")
End Property

Property Get Opt_FehlerVerbesserungen() As Variant
  'Liefert die entsprechende Import-Option.
  Opt_FehlerVerbesserungen = oDataImportOptions("EditModule_Differences")
End Property

Property Get Opt_UeberhoehungAusBemerkung() As Variant
  'Liefert die entsprechende Import-Option.
  Opt_UeberhoehungAusBemerkung = oDataImportOptions("EditModule_SlantFromInfo")
End Property

Property Get Opt_iTrassenCodeAusBemerkung() As Variant
  'Liefert die entsprechende Import-Option.
  Opt_iTrassenCodeAusBemerkung = oDataImportOptions("EditModule_Parse_iTC")
End Property

Property Get Opt_Transfo_Tk2Gls() As Variant
  'Liefert die entsprechende Import-Option.
  Opt_Transfo_Tk2Gls = oDataImportOptions("EditModule_RailCooTransfo")
End Property



'interne Routinen *********************************************************************************

Private Function LeseDatei() As Integer
  'Liest die eigentlichen CSV-Daten in das Array datCSV.
  'Der Dateikopf muss bereits gelesen sein!
  
  On Error GoTo Fehler
  
  
  'Deklarationen
    Dim Erfolg          As Boolean
    Dim Dateiname       As String
    Dim Meldung         As String
    Dim Meldung2        As String
    Dim CsvDatensatz    As String
    Dim FieldSeparator  As String
    Dim TextQualifier   As String
    Dim TrimFields      As Boolean
    Dim MaxAnzSp
    Dim Ze
    Dim Sp
    
  'Variableninitialisierung
    Erfolg = true
    lfdNrTotal = 0
    FieldSeparator = oCsvSettings("FieldSeparator")
    TextQualifier = oCsvSettings("TextQualifier")
    TrimFields = cBool(oCsvSettings("TrimFields"))  'sicher, da Wert bereits geprüft ist.
  '
  Dateiname = Me.Quelle_AsciiDatei_Name
  
  ProgressbarDateiLesen (Kanal)
  
  'CSV-Dateiname zu Metadaten hinzufügen, falls dieses Feld nicht bereits im CSV-Kopf belegt wurde.
  if (Not oMetadatenExtra.Exists(FN_Statistik_Importdatei)) then
    oMetadatenExtra.Add FN_Statistik_Importdatei, Dateiname
  end if
  
  'Daten einlesen.
  Do While (Not EOF(Kanal))
    'Aus der CSV-Datei werden alle Spalten in Vektorfelder eingelesen.
    Line Input #Kanal, Zeile
    ProgressbarDateiLesen (Kanal)
    NR = NR + 1
    CsvDatensatz = trim(Zeile)
    
    'Weitere Zeilen eines mehrzeiligen Datensatzes lesen.
    'Delimited files can have a odd number of quotes which means the row spans more than 1 line
    do while ((CountChar(CsvDatensatz, TextQualifier) mod 2 <> 0) and Not EOF(Kanal))
      Line Input #Kanal, Zeile
      ProgressbarDateiLesen (Kanal)
      NR = NR + 1
      CsvDatensatz = CsvDatensatz & vbLF & trim(Zeile)
    loop
    
    'Zeile splitten
    'NF = SplitDelim(Zeile, Feld, FieldSeparator)
    NF = SplitCsvLine(CsvDatensatz, Feld, FieldSeparator, TextQualifier, TrimFields)
    
    lfdNrTotal = lfdNrTotal + 1              'Punktanzahl gesamt
    ReDim Preserve datCSV(1 To oQuellSpaltenPositionen.count, 1 To lfdNrTotal)
    
    'Aktuell zu verarbeitende Feldanzahl.
    if (NF <= oQuellSpaltenPositionen.count) then
      MaxAnzSp = NF
    else
      MaxAnzSp = oQuellSpaltenPositionen.count
    end if
    
    'Aktuelle Zeile der Datenmatrix belegen.
    for Sp = 1 to MaxAnzSp
      datCSV(Sp, lfdNrTotal) = Feld(Sp)
    next
    
    'If (LeseZeile() = False) Then
    '  'Fehler beim Verarbeiten der Zeile.
    '  If (Me.Meldung_Inhalt = "") Then Me.Meldung_Inhalt = "Unbekannter Fehler beim Lesen der Zeile " & CStr(NR) & " der Datei " & Dateiname
    '  GoTo Fehler
    'End If
  Loop
  Call CloseCSV()
  
  If (lfdNrTotal = 0) Then
    Me.AktionErfolgreich = False
    Me.Fehlerniveau   = Fehlerniveau_Kritisch
    Me.Meldung_Titel  = "Fehler beim CSV-Import"
    Me.Meldung_Inhalt = Me.Quelle_AsciiDatei_Name & " enthält keine CSV-Daten!"
    
  Else
    'Daten in den Quelldatenpuffer schreiben.
    Me.AktionErfolgreich = True
    Me.Fehlerniveau   = Fehlerniveau_Kein
    Me.Meldung_Titel  = "CSV-Import"
    Me.Meldung_Inhalt = lfdNrTotal & " Punkte importiert aus CSV-Datei" & vbNewLine & vbNewLine & Me.Quelle_AsciiDatei_Name
    
    'CSV-Datenmatrix in Spalten aufteilen und schreiben
    ReDim Vektor(1 To lfdNrTotal)
    for Sp = 1 to oQuellSpaltenPositionen.count
      for Ze = 1 to lfdNrTotal
        Vektor(Ze)= datCSV(Sp, Ze)
      next
      oQuellDatenpuffer.Add QuellSpaltenNamen(Sp), Vektor
    next
    
    ReDim datCSV(0, 0)
  End If
  
  'Kontrolle über Statuszeile an Excel zurückgeben.
  'Application.StatusBar = ""
  LeseDatei = lfdNrTotal
  Exit Function

Fehler:
  FehlerNachricht "CtabCSV.LeseDatei()"
  Call CloseCSV()
  Me.AktionErfolgreich = False
  Me.Fehlerniveau  = Fehlerniveau_Kritisch
  LeseDatei = lfdNrTotal
End Function


Private Function LeseDateiKopf() As Boolean
  '-----------------------------------------------------------------------------------------------
  'Liest den Dateikopf und wertet ihn aus.
  'Aufbau des Kopfes: 1. Optionaler Spezialkopf (muss in der ersten Zeile beginnen!):
  '                      @GEOTOOLS_BEGIN
  '                        ... Kopfdaten ...
  '                      @GEOTOOLS_END
  '                   2. CSV-Zeile mit Spaltennamen
  '-----------------------------------------------------------------------------------------------
  'Rückgabe: - true,  falls der Kopf mit Erfolg gelesen wurde, d.h. mindestens ein Spaltenname.
  '          - False, falls der Kopf ohne Erfolg gelesen wurde
  '-----------------------------------------------------------------------------------------------
  
  On Error GoTo Fehler
  
  'Deklarationen
    Dim Erfolg          As Boolean
    Dim Dateiname       As String
    Dim Meldung         As String
    Dim Meldung2        As String
    Dim Spalte          As Variant
    
  'Variableninitialisierung
    Erfolg = true
    
  Dateiname = Me.Quelle_AsciiDatei_Name
  
  If (not ThisWorkbook.SysTools.isDatei(Dateiname)) Then
    'Programmfehler, da der Dateiname bereits von CdatExpim ermittelt wird.
    ErrEcho "CtabCSV.LeseDateiKopf(): Öffnen der folgenden Datei fehlgeschlagen (existiert nicht): '" & Dateiname & "'."
  Else
    'Eingabedatei existiert.
    DebugEcho "CtabCSV.LeseDateiKopf(): Lese Kopf der Datei: '" & Dateiname & "'."
    
    'Import-Datei öffnen
    Kanal = FreeFile()
    ErrMessage = "Fehler beim Öffnen der CSV-Datei '" & Dateiname & "'."
    Open Dateiname For Input Lock Write As #Kanal
    If (Err) Then GoTo Fehler
    ErrMessage = ""
    
    if (EOF(Kanal)) then
      Erfolg = false
      WarnEcho "Die CSV-Datei '" & Dateiname & "' ist leer!"
      Meldung = "Warnung: Die CSV-Datei ist leer"
      Call ThisWorkbook.SysTools.FileErrorsAdd("W", Dateiname, 1, 0, 0, Meldung, "", blnFehlerKonsolidieren)
      
    else
      '1. Zeile der Datei lesen
      Line Input #Kanal, Zeile
      NR = 1
      
      'Optionalen Spezialkopf lesen. Dieser muss in der ersten Zeile beginnen!
      if (Left(Zeile, len(CsvKopfBeginn)) = CsvKopfBeginn) Then
        'Spezialkopf lesen. Lesen endet mit Lesen der Zeile "@GEOTOOLS_END"
        Erfolg = LeseSpezialKopf()
        
        if (not Erfolg) then
          ErrEcho "Fehler beim Lesen des GeoTools-Spezialkopfes der CSV-Datei '" & Dateiname & "'."
        else
          '1. Zeile nach dem GeoTools-Spezialkopf lesen
          if (EOF(Kanal)) then
            Erfolg = false
            ErrEcho "Fehler: Nach dem GeoTools-Spezialkopf gibt es keine Daten in der CSV-Datei '" & Dateiname & "'."
            Meldung = "Fehler: Nach dem GeoTools-Spezialkopf gibt es keine Daten in der CSV-Datei"
            Call ThisWorkbook.SysTools.FileErrorsAdd("F", Dateiname, NR, 0, 0, Meldung, "", blnFehlerKonsolidieren)
          else
            Line Input #Kanal, Zeile
            NR = NR + 1
          end if
        end if
      end if
      
      'Spaltennamen bestimmen
      if (Erfolg) then
        Call GetQuellSpaltenInfo(Zeile)
        if (oQuellSpaltenPositionen.count < 1) then Erfolg = false
      end if
      
      'Einheitenabgleich: Falls keine Zieleinheit, dann Zieleinheit = Quelleinheit, wenn verfügbar.
      for each Spalte in oQuellEinheiten
        if (Not oZielEinheiten.Exists(Spalte)) then
          if (oQuellEinheiten(Spalte) <> "") then
            oZielEinheiten.add Spalte, oQuellEinheiten(Spalte)
          end if
        end if
      next
    end if
  End If
  
  LeseDateiKopf = Erfolg
  Exit Function
  
Fehler:
  FehlerNachricht "CtabCSV.LeseDateiKopf()"
  Call CloseCSV()
  LeseDateiKopf = false
End Function


Private Function LeseSpezialKopf() As Boolean
  '-----------------------------------------------------------------------------------------------
  'Liest den GeoTools-Spezialkopf und wertet ihn aus.
  'Aufbau des Kopfes:    @GEOTOOLS_BEGIN
  '                        ... Kopfdaten ...
  '                      @GEOTOOLS_END
  '-----------------------------------------------------------------------------------------------
  'Rückgabe: - True,  falls der Kopf mit Erfolg gelesen wurde.
  '          - False, falls ein Fehler aufrat.
  '-----------------------------------------------------------------------------------------------
  'Status zu Beginn dieser Routine: Zeile "@GEOTOOLS_BEGIN" ist bereits gelesen.
  'Status zum Ende  dieser Routine: Zeile "@GEOTOOLS_END"   ist bereits gelesen.
  '-----------------------------------------------------------------------------------------------
  
  'Deklarationen
    Dim Erfolg          As Boolean
    Dim blnKopfEnde     As Boolean
    Dim Dateiname       As String
    Dim Meldung         As String
    Dim Meldung2        As String
    Dim Abschnitt       As String
    Dim Name            As String
    Dim Wert            As String
    Dim WertVariant     As Variant
  '
  DebugEcho "CtabCSV.LeseSpezialKopf(): Lese GeoTools-Spezialkopf."
  Erfolg = true
  Abschnitt = ""
  Dateiname = Me.Quelle_AsciiDatei_Name
  
  
  Do While (Not (blnKopfEnde or EOF(Kanal)))
    '
    Line Input #Kanal, Zeile
    NR = NR + 1
    
    if (Left(Zeile, len(CsvKopfEnde)) = CsvKopfEnde) Then
      'Ende-Kennung
      blnKopfEnde = true
      
    Else
      'Zeile innerhalb des Spezialkopfes => auswerten
      Zeile = trim(Zeile)
      
      if (Zeile = "") then
        'Leerzeile ist ok
        
      elseif (left(Zeile, 1) = "#") then
        'Kommentarzeile ist ok
        
      elseif (left(Zeile, 1) = "@") then
        'Abschnittsname
        Select Case Zeile
          Case "@CSV-Settings", "@Metadata", "@ExcelFieldFormats", "@ExcelFieldUnits", "@ExcelTableProps", "@DataImportOptions"
              Abschnitt = Zeile
          Case Else
              Abschnitt = ""
              Erfolg = false
              ErrEcho "Syntaxfehler im GeoTools-Spezialkopf in Zeile " & cStr(NR) & ": Unbekannter Abschnittsname!"
              Meldung = "Syntaxfehler im GeoTools-Spezialkopf: Unbekannter Abschnittsname!"
              Meldung2 = "Unterstützte Abschnittsnamen: @CSV-Settings, @Metadata, @ExcelFieldFormats, @ExcelFieldUnits, @ExcelTableProps"
              Call ThisWorkbook.SysTools.FileErrorsAdd("F", Dateiname, NR, 0, 0, Meldung, Meldung2, blnFehlerKonsolidieren)
        End Select
        
      else
        'Name/Wert-Paar innerhalb eines Abschnittes
        if (Abschnitt = "") then
          Erfolg = false
          ErrEcho "Syntaxfehler im GeoTools-Spezialkopf in Zeile " & cStr(NR) & ": Informationen außerhalb eines Abschnittes sind nicht erlaubt!"
          Meldung = "Syntaxfehler im GeoTools-Spezialkopf: Informationen außerhalb eines Abschnittes sind nicht erlaubt!"
          Call ThisWorkbook.SysTools.FileErrorsAdd("F", Dateiname, NR, 0, 0, Meldung, "", blnFehlerKonsolidieren)
          
        else
          'Name und Wert extrahieren.
          Name = trim(LeftStr(Zeile, "=", False))
          Wert = trim(MidStr(Zeile, "="))
          
          if ((Name = "") OR (Wert = "")) then
            Erfolg = false
            ErrEcho "Syntaxfehler im GeoTools-Spezialkopf in Zeile " & cStr(NR) & ": Name oder/und Wert sind leer!"
            Meldung = "Syntaxfehler im GeoTools-Spezialkopf: Name oder/und Wert sind leer!"
            Call ThisWorkbook.SysTools.FileErrorsAdd("F", Dateiname, NR, 0, 0, Meldung, "", blnFehlerKonsolidieren)
          else
            'Name und Wert verarbeiten.
            
            Select Case Abschnitt
              Case "@CSV-Settings"
                  if (oCsvSettings.exists(Name)) then
                    if (KonvString(Wert, oCsvSettingsTyp(Name), WertVariant)) then
                      oCsvSettings(Name) = WertVariant
                    else
                      Erfolg = false
                    end if
                  Else
                    Erfolg = false
                    ErrEcho "Syntaxfehler im GeoTools-Spezialkopf in Zeile " & cStr(NR) & ": Unbekannte CSV-Einstellung!"
                    Meldung = "Syntaxfehler im GeoTools-Spezialkopf: Unbekannte CSV-Einstellung!"
                    Meldung2 = "Unterstützte Schlüsselworte: " & join(oCsvSettings.Keys, ", ")
                    Call ThisWorkbook.SysTools.FileErrorsAdd("F", Dateiname, NR, 0, 0, Meldung, Meldung2, blnFehlerKonsolidieren)
                  End If
                  
              Case "@Metadata"
                  if (ThisWorkbook.Metadaten.AlleProjektDaten.exists(Name)) then
                    'Gültiges Projektdatum.
                    oMetadatenProjekt(Name) = Wert
                  Else
                    'Beliebiges Extradatum.
                    if (oMetadatenExtra.Exists(Name)) then
                      oMetadatenExtra(Name) = Wert
                    else
                      oMetadatenExtra.add Name, Wert
                    end if
                  End if
                  
              Case "@DataImportOptions"
                  if (oDataImportOptions.exists(Name)) then
                    if (KonvString(Wert, oDataImportOptionsTyp(Name), WertVariant)) then
                      oDataImportOptions(Name) = WertVariant
                    else
                      Erfolg = false
                    end if
                  Else
                    Erfolg = false
                    ErrEcho "Syntaxfehler im GeoTools-Spezialkopf in Zeile " & cStr(NR) & ": Unbekannte Eigenschaft für Daten-Importoptionen!"
                    Meldung = "Syntaxfehler im GeoTools-Spezialkopf: Unbekannte Eigenschaft für Daten-Importoptionen!"
                    Meldung2 = "Unterstützte Schlüsselworte: " & join(oDataImportOptions.Keys, ", ")
                    Call ThisWorkbook.SysTools.FileErrorsAdd("F", Dateiname, NR, 0, 0, Meldung, Meldung2, blnFehlerKonsolidieren)
                  End If
                  
              Case "@ExcelFieldFormats"
                    if (oZielFormate.Exists(Name)) then
                      oZielFormate(Name) = Wert
                    else
                      oZielFormate.add Name, Wert
                    end if
                  
              Case "@ExcelFieldUnits"
                  if(ThisWorkbook.Konfig.EinheitExists(Wert)) then
                    if (oZielEinheiten.Exists(Name)) then
                      oZielEinheiten(Name) = Wert
                    else
                      oZielEinheiten.add Name, Wert
                    end if
                  else
                    Erfolg = false
                    ErrEcho "Syntaxfehler im GeoTools-Spezialkopf in Zeile " & cStr(NR) & ": Unbekannte Einheit!"
                    Meldung = "Syntaxfehler im GeoTools-Spezialkopf: Unbekannte Einheit!"
                    Meldung2 = "Die unterstützten Einheiten sind am Anfang des Debug-Protokolls nachgewiesen! "
                    Call ThisWorkbook.SysTools.FileErrorsAdd("F", Dateiname, NR, 0, 0, Meldung, Meldung2, blnFehlerKonsolidieren)
                  end if
                  
              Case "@ExcelTableProps"
                  if (oExcelTableProps.exists(Name)) then
                    if (KonvString(Wert, oExcelTablePropsTyp(Name), WertVariant)) then
                      oExcelTableProps(Name) = WertVariant
                    else
                      Erfolg = false
                    end if
                  Else
                    Erfolg = false
                    ErrEcho "Syntaxfehler im GeoTools-Spezialkopf in Zeile " & cStr(NR) & ": Unbekannte Eigenschaft für Exceltabelle!"
                    Meldung = "Syntaxfehler im GeoTools-Spezialkopf: Unbekannte Eigenschaft für Exceltabelle!"
                    Meldung2 = "Unterstützte Schlüsselworte: " & join(oExcelTableProps.Keys, ", ")
                    Call ThisWorkbook.SysTools.FileErrorsAdd("F", Dateiname, NR, 0, 0, Meldung, Meldung2, blnFehlerKonsolidieren)
                  End If
                  
              Case Else
                  ErrEcho "Programmfehler: GeoTools-Spezialkopf enthält Werte für unbekannten Abschnitt in Zeile " & cStr(NR)
            End Select
            
          end if
        end if
      end if
    End If
  Loop
  
  if (Not (blnKopfEnde)) then
    ErrEcho "Endmarke des GeoTools-Spezialkopfes (" & CsvKopfEnde & ") nicht gefunden!"
    Meldung = "Endmarke des GeoTools-Spezialkopfes (" & CsvKopfEnde & ") nicht gefunden"
    Call ThisWorkbook.SysTools.FileErrorsAdd("F", Dateiname, NR, 0, 0, Meldung, "", blnFehlerKonsolidieren)
  end if
  
  LeseSpezialKopf = (Erfolg AND blnKopfEnde)
  Exit Function
  
Fehler:
  FehlerNachricht "CtabCSV.LeseSpezialKopf()"
  Call CloseCSV()
  LeseSpezialKopf = false
End Function


Private Sub GetQuellSpaltenInfo(CsvZeileMitSpaltennamen as String)
  '------------------------------------------------------------------------------------------------
  'Analysiert die übergebene Zeile der CSV-Datei und interpretiert die einzelnen Felder als Spaltennamen.
  'Enthält der Spaltenname auch eine Einheitenangabe, so wird diese extrahiert.
  '
  'Eingabeparameter:
  '  CsvZeileMitSpaltennamen ... Die entsprechende Zeile der CSV-Datei.
  '                              ==> Der Spaltenname darf eine Einheit, aber keinen Präfix ("Spalte.") enthalten.
  'Ergebnis = globale Dictionaries:
  '  oQuellSpaltenPosition ... lfd. CSV-Spaltennummer  gefundenen Spaltennamen.
  '  oQuellEinheiten       ... Die Einheit zu jedem gefundenen Spaltennamen, falls verfügbar; sonst="ohne".
  ' ***** oQuellFormate         ... Formatangabe zu jedem gefundenen Spaltennamen, falls verfügbar; sonst="ohne".
  ' *****                           ==> bisher ungenutzt, also = "ohne".
  '  ==> Der Key der 3 Dictionaries ist der reine Spaltenname ohne Einheitenangabe.
  '------------------------------------------------------------------------------------------------
  
  'On Error Resume Next
  On Error GoTo Fehler
  
  Dim Spaltenname       As String
  Dim SpaltennamePur    As String
  Dim SpaltenEinheit    As String
  Dim i                 As Long
  Dim k                 As Long
  Dim AnzSpalten        As Long
  Dim Spaltennamen()    As String
  Dim NF                As Long
  Dim Feld()            As String
  Dim SpNameTmp         As String
  Dim Meldung           As String
  Dim Meldung2          As String
  
  'CSV-Zeile zerlegen.
  AnzSpalten = SplitDelim(CsvZeileMitSpaltennamen, Spaltennamen, oCsvSettings("FieldSeparator"))
  
  'Spaltennamen analysieren.
  If (AnzSpalten = 0) Then
    'DebugEcho "CtabCSV.GetQuellSpaltenInfo(): Es wurden keine Spaltennamen gefunden => CSV-Datei kann nicht eingelesen werden!"
    ErrEcho "Zeile " & cStr(NR) & ": Es wurden keine Spaltennamen gefunden => CSV-Datei kann nicht eingelesen werden!"
    Meldung = "Es wurden keine Spaltennamen gefunden => CSV-Datei kann nicht eingelesen werden"
    Call ThisWorkbook.SysTools.FileErrorsAdd("F", Me.Quelle_AsciiDatei_Name, NR, 0, 0, Meldung, "", blnFehlerKonsolidieren)
    
  else
    DebugEcho "CtabCSV.GetQuellSpaltenInfo(): Es wurden " & cStr(AnzSpalten) & " Spaltennamen gefunden."
    ReDim QuellSpaltenNamen(1 To AnzSpalten)
    
    'Einheitenangabe aus dem Spaltennamen extrahieren, falls vorhanden.
    For i = 1 To AnzSpalten
      Spaltenname = trim(Spaltennamen(i))
      
      'Ungültige Zeichen aus Spaltenname entfernen
      Spaltenname = BereichsnameBereinigen(Spaltenname)
      
      'Spaltenname verarbeiten
      If (Spaltenname <> "") Then
        NF = SplitDelim(Spaltenname, Feld, TrennerEinheit)
        If (NF > 1) Then
          SpaltennamePur = Feld(1)
          SpaltenEinheit = Feld(NF)
        Else
          SpaltennamePur = Spaltenname
          SpaltenEinheit = "ohne"
        End If
      Else
        SpaltennamePur = "Spalte_" & cStr(i)
        SpaltenEinheit = "ohne"
      End If
      
      'Eindeutigkeit des Spaltennamens herstellen, falls nötig (durch Suffix).
      SpNameTmp = SpaltennamePur
      do while (oQuellSpaltenPositionen.exists(SpaltennamePur))
        k = k +1
        SpaltennamePur = SpNameTmp & "_" & cStr(k)
      loop
      
      'Spalteneigenschaften speichern.
      QuellSpaltenNamen(i) =      SpaltennamePur
      oQuellSpaltenPositionen.Add SpaltennamePur, i
      oQuellEinheiten.Add         SpaltennamePur, SpaltenEinheit
      'oQuellFormate.Add           SpaltennamePur, ""
    Next
  End If
  
  Exit Sub

Fehler:
  FehlerNachricht "CtabCSV.GetQuellSpaltenInfo()"
End Sub

Private Function BereichsnameBereinigen(byVal Name As String) As String
  'Entfernt ungültige Zeichen aus dem Bereichsnamen
  const AlleErsetzen  As Boolean = true
  const SuchMuster    As String  = "[^a-zA-Z0-9_\.]"
  BereichsnameBereinigen = substitute(SuchMuster, "", Name, AlleErsetzen, False)
End Function


Private Function CountChar(SourceLine As String, SearchChar As String)
  'Counts the number of times SearchChar occurs in SourceLine. Used by SplitCsvLine
  'Dim SourceIndex   As Long
  Dim HitCount      As Long
  Dim idxStart      As Long
  Dim idxFound      As Variant
  Dim NothingFound  As Boolean
  
  HitCount = 0
  
  if (Not ((SourceLine = "") or (SearchChar = "") or (len(SearchChar) <> 1))) then
    'for SourceIndex = 1 to len(SourceLine)
    '  if (mid(SourceLine, SourceIndex, 1) = SearchChar) then HitCount = HitCount + 1
    'next
    
    idxFound = 0
    NothingFound = false
    do
      idxStart = idxFound + 1
      idxFound = instr(idxStart, SourceLine, SearchChar)
      if (idxFound > 0) then
        HitCount = HitCount + 1
      else
        NothingFound = true
      end if
    loop until (NothingFound)
  end if
  
  CountChar = HitCount
End Function


Private Function SplitCsvLine(ByVal SourceLine As String, ByRef OutputArray As Variant, ByVal FieldDelimiter As String, ByVal TextQualifier As String, ByVal TrimFields As Boolean) As Long
  'Take the SourceLine string and split it into an array.
  'SourceLine     ... CSV record to split
  'OutputArray    ... Resulting Array with found fields (BaseIndex = 0, but field 0 is always "").
  'FieldDelimiter ... CSV field separator
  'TextQualifier  ... CSV text qualifier (encloses text witch contains CSV field separator)
  'TrimFields     ... Remove leading and traling spaces from every field (only outside of a pair of 'TextQualifier').
  'Return Value    =  number of found fields, -1 on error
  
  On Error GoTo Fehler
  ErrMessage = "Unbekannter Fehler beim Zerlegen des CSV-Datensatzes"
  
  Dim OutputIndex          As Long
  Dim SourceIndex          As Long
  Dim SourceLength         As Long
  Dim InQuotes             As Boolean
  Dim isLastChar           As Boolean
  Dim isDelimiter          As Boolean
  Dim SourceChar           As String
  Dim OutputField          As String
  Dim DoubleTextQualifier  As String
  
  InQuotes     = false
  OutputIndex  = 0
  OutputField  = ""
  SourceLength = len(SourceLine)
  DoubleTextQualifier = TextQualifier & TextQualifier
  ReDim OutputArray(0)
  OutputArray(0) = ""
  
  for SourceIndex = 1 to SourceLength
    SourceChar  = mid(SourceLine, SourceIndex, 1)
    isLastChar  = (SourceIndex = SourceLength)
    isDelimiter = (SourceChar = FieldDelimiter)
    if (SourceChar = TextQualifier) then InQuotes = (Not InQuotes)
    
    if ((isDelimiter and Not InQuotes) or isLastChar) then
      'Finish and save current field.
      if (isLastChar and Not isDelimiter) then OutputField = OutputField & SourceChar
      if (TrimFields) then OutputField = trim(OutputField)
      if (left(OutputField, 1) = TextQualifier) then OutputField = right(OutputField, len(OutputField)-1)
      if (right(OutputField, 1) = TextQualifier) then OutputField = left(OutputField, len(OutputField)-1)
      OutputField = replace(OutputField, DoubleTextQualifier, TextQualifier)
      
      'Save current field...
      OutputIndex = OutputIndex + 1
      redim preserve OutputArray(OutputIndex)
      OutputArray(OutputIndex) = OutputField
      OutputField = ""
    else
      'Collect characters for current field.
      OutputField = OutputField & SourceChar
    end if
  next
  
  SplitCsvLine = OutputIndex
  Exit Function
  
Fehler:
  FehlerNachricht "CtabCSV.SplitCsvLine()"
  SplitCsvLine = -1
End Function

Private Function KonvString(byVal QuellWert As String, byVal ZielTyp As Long, byRef ErgebnisWert As Variant) As Boolean
  '-----------------------------------------------------------------------------
  'Versucht, den Quellwert in den angegebenen Zieltyp zu konvertieren.
  'Bei Mißerfolg wird eine Fehlermeldung erzeugt.
  '
  'Eingabe: QuellWert     ... zu konvertierender Text
  '         ZielTyp       ... eine vbVarType-Konstante
  'Ausgabe: ErgebnisWert  ... das Ergebnis als Variant
  '
  'Funktionswert: true bei Erfolg, sonst false.
  '-----------------------------------------------------------------------------
  On error goto Fehler
  Dim Erfolg         As Boolean
  Dim Meldung        As String
  Dim TypName        As String
  
  Erfolg = true
  
  Select Case ZielTyp
    Case vbBoolean:  ErgebnisWert = cBool(QuellWert)
    Case vbByte:     ErgebnisWert = cByte(QuellWert)
    Case vbCurrency: ErgebnisWert = cCur(QuellWert)
    Case vbDate:     ErgebnisWert = cDate(QuellWert)
    Case vbDouble:   ErgebnisWert = cDbl(QuellWert)
    Case vbInteger:  ErgebnisWert = cInt(QuellWert)
    Case vbLong:     ErgebnisWert = cLng(QuellWert)
    Case vbSingle:   ErgebnisWert = cSng(QuellWert)
    Case vbString:   ErgebnisWert = cStr(QuellWert)
    Case vbVariant:  ErgebnisWert = cVar(QuellWert)
    Case Else
        WarnEcho "CtabCSV.KonvString(): Programmfehler - Nicht unterstützter Zieltyp: " & cStr(ZielTyp)
        Erfolg = false
  End Select
  
  KonvString = Erfolg
  Exit Function
  
Fehler:
  err.clear
  KonvString = false
  
  Select Case ZielTyp
    Case vbBoolean:  TypName = "Boolean"
    Case vbByte:     TypName = "Byte"
    Case vbCurrency: TypName = "Currency"
    Case vbDate:     TypName = "Date"
    Case vbDouble:   TypName = "Double"
    Case vbInteger:  TypName = "Integer"
    Case vbLong:     TypName = "Long"
    Case vbSingle:   TypName = "Single"
    Case vbString:   TypName = "String"
    Case vbVariant:  TypName = "Variant"
  End Select
  
  ErrEcho "Syntaxfehler im GeoTools-Spezialkopf in Zeile " & cStr(NR) & ": Ungültiger Wert! (erforderlich ist '" & TypName & "')"
  Meldung = "Syntaxfehler im GeoTools-Spezialkopf: Ungültiger Wert! (erforderlich ist '" & TypName & "')"
  Call ThisWorkbook.SysTools.FileErrorsAdd("F", Me.Quelle_AsciiDatei_Name, NR, 0, 0, Meldung, "", blnFehlerKonsolidieren)
End Function


Private Function GetKategorien() As String
  'Ermittelt alle unterschiedlichen Spalten-Kategorien der aktiven Datei und
  'den Kodenamen der Tabelle als erste Kategorie.
  'Rückgabe: Liste durch Semikolons getrennt.
  
  On Error GoTo Fehler
  
  Dim DictTmp           As Scripting.Dictionary
  Dim Spalte            As Variant
  Dim Liste             As String
  Dim ListeKeys         As String
  
  Liste = Me.Ziel_XLTabKlasse
  Set DictTmp = New Scripting.Dictionary
  
  If (Not (oQuellSpaltenPositionen Is Nothing)) Then
    On Error Resume Next
    For Each Spalte In oQuellSpaltenPositionen
      If (ThisWorkbook.Konfig.SpaltenKategorie(Spalte) <> "") Then DictTmp.Add ThisWorkbook.Konfig.SpaltenKategorie(Spalte), "*"
    Next
    On Error GoTo 0
    ListeKeys = ListeDerKeys(DictTmp)
    If (ListeKeys <> "") Then
      If (Liste = "") Then
        Liste = ListeKeys
      Else
        Liste = Liste & ";" & ListeKeys
      end if
    end if
  End If
  Set DictTmp = Nothing
  GetKategorien = Liste
  Exit Function
  
Fehler:
  FehlerNachricht "CtabCSV.GetKategorien()"
  GetKategorien = ""
  Set DictTmp = Nothing
End Function


Private Sub CloseCSV()
  On error resume next
  Close #Kanal
  On Error goto 0
End Sub


Private Sub Syncronisieren()
  '=> Diese Routine wird automatisch aufgerufen, wenn die Eigenschaft "Quelle_AsciiDatei_Name" gesetzt wird!
  'Syncronisiert die Daten einiger Eigenschaften mit der aktiven Datei.
  'Arbeitsweise: - Schließen einer eventuell offenen Datei
  '              - Öffnen von Me.Quelle_AsciiDatei_Name
  '              - Lesen des Kopfes bis zur Zeile mit den Spaltennamen
  '              - Bereitstellung der gelesenen Info's in Egenschaften
  '              ==> Der Dateizeiger bleibt in der Zeile mit den Spaltennamen stehen!
  
  On Error GoTo Fehler
  DebugEcho "CtabCSV.Syncronisieren(): Syncronisiere Datei: '" & Me.Quelle_AsciiDatei_Name & "'."
  
  Dim Erfolg  As Boolean
  
  'Vorhandene Zuordnungen löschen.
  oQuellSpaltenPositionen.RemoveAll
  oQuellEinheiten.RemoveAll
  'oQuellFormate.RemoveAll
  ReDim QuellSpaltenNamen(0 To 0)
  
  'Letzte Datei schließen, falls eine offen ist!
  Call CloseCSV()
  
  '1. Gesamten Dateikopf lesen und auswerten.
  Erfolg = LeseDateiKopf() 
  
  if (not Erfolg) then
    Me.Fehlerniveau   = Fehlerniveau_Kritisch
    Me.Meldung_Titel  = "Fehler beim Syncronisieren der CSV-Datei"
    Me.Meldung_Inhalt = Me.Quelle_AsciiDatei_Name & " enthält keinen gültigen Kopf!"
  else
    Me.Fehlerniveau   = Fehlerniveau_Kein
    Me.Meldung_Titel  = "Syncronisieren der CSV-Datei"
    Me.Meldung_Inhalt = Me.Quelle_AsciiDatei_Name & " enthält " & cStr(oQuellSpaltenPositionen.count) & " Datenspalten."
    
    '2. Kategorien konfigurierter Spaltennamen ermitteln.
    strKategorien = GetKategorien()
    
    Call Me.ZeigeSpalten()
  end if
  Me.AktionErfolgreich = Erfolg
  
  Exit Sub
  
Fehler:
  FehlerNachricht "CtabCSV.Syncronisieren()"
End Sub




'Methoden *****************************************************************************************

Sub Import()
  'Import der CSV-Daten in den Datenpuffer.
  'Syncronisation muss bereits erfolgt sein!
  
  On Error GoTo Fehler
  
  Dim NameEingabeDatei  As String
  Dim Anzahl            As Integer
  
  NameEingabeDatei = Me.Quelle_AsciiDatei_Name
  
  If ((NameEingabeDatei <> "") AND (Me.Fehlerniveau = Fehlerniveau_Kein)) Then
    'Eingabedatei gesetzt und erfolgreich syncronisiert.
    
    'Falls nicht (mehr) sichergestellt ist, dass der Dateiname unmittelbar vor Aufruf des Imports
    'gesetzt wird, muss die Datei hier erneut syncronisiert werden.
    'Me.Syncronisieren
    
    Anzahl = LeseDatei()
    
  Else
    'Programmfehler, da der Dateiname bereits von CdatExpim gesetzt wird.
    ErrEcho "Öffnen der folgenden Datei fehlgeschlagen: '" & NameEingabeDatei & "'."
  End If
  
  Exit Sub
  
Fehler:
  FehlerNachricht "CtabCSV.Import()"
End Sub


Public Sub ZeigeSpalten()
  'Schreibt für Kontrollzwecke alle gefundenen Spalten der aktiven Datei ins Debug-Protokoll.
  Dim Spalte As Variant
  For Each Spalte In Me.Quelle_SpaltenPositionen
    DebugEcho "Spalte=" & Spalte & vbTab & vbTab & _
              "Einheit=" & Me.Quelle_Einheiten(Spalte) & vbTab & vbTab & _
              "Position=" & cStr(Me.Quelle_SpaltenPositionen(Spalte))
  Next
  DebugEcho "==> Kategorien = '" & Me.Kategorien & "'" & vbNewLine
End Sub


' Für jEdit:  :collapseFolds=1:
