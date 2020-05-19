VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CdatExpim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2004-2020  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'==================================================================================================
'Klasse CdatExpim
'==================================================================================================
'Stellt Eigenschaften und Methoden für Import- und Export zur Verfügung:
'
' - Lesen von Daten aus bezeichneten Spalten der aktiven Tabelle in den Datenpuffer.
' - Lesen von Daten aus Objekten für den ASCII-Spezialimport in den Datenpuffer.
' - Lesen von Daten aus Objekt für den CSV-Spezialimport in den Datenpuffer.
' - Schreiben der Daten des Datenpuffers in die aktive Tabelle.
'
' - Objekte für den ASCII-Spezialimport müssen folgendes Interface implementieren:
 '
 '  Methoden:
 '  - Import                             ... liest die Daten in den klasseninternen Datenpuffer
 '                                           oder direkt in eine neu angelegte XL-Tabelle
 '  Eigenschaften (G=Get  L=Let):
 '  - G  Klasse_Name                     ... Name der Klasse (bisher nicht verwendet)
 '  - G  Titel                           ... Formatbeschreibung für den Import/Export-Dialog
 '  - G  Kategorien                      ... mögliche Kategorien der zu importierenden Werte
 '  - G
 '  - G  Quelle_AsciiDatei_DialogFilter  ... für Dateidialog..
 '  - GL Quelle_AsciiDatei_Name          ... Name der zu importierenden Datei
 '  - G  Quelle_DatenExtra               ... Dictionary mit Extradaten (Statistik usw.) (Key=Feldname, Item=Inhalt)
 '  - G  Quelle_Datenpuffer              ... Dictionary mit Quelldaten (Key=Spaltenname, Item=Datenvektor)
 '  - G  Quelle_Einheiten                ... Dictionary mit Quelleinheiten (Key=Spaltenname, Item=Einheit)
 '
 '  - G  Ziel_Typ                        ... Angabe, wohin die Import-Methode die gelesenen Daten schreibt.
 '                                           Möglich sind: io_Typ_Puffer oder io_Typ_XlTabNeu
 '  - G  *** wirkungslos ***  Ziel_XLTabKlasse                ... Name der XL-Tabellen-Klasse (Kodename), falls Ziel_Typ = io_Typ_XlTabNeu
 '  - G  Ziel_XLMappeSpeichern           ... Soll die neu angelegte XL_Mappe gespeichert werden?
 '
 '  - GL AktionErfolgreich               ... true oder false. "False" bewirkt Abbruch des Import/Export's
 '  - GL Fehlerniveau                    ... Eine Fehlerniveau-Konstante (zwecks Meldungsmanagement).
 '  - GL Meldung_Inhalt                  ... Schlussmeldung
 '  - GL Meldung_Titel                   ... Schlussmeldung
'
'Historie:
 '02.05.2020  Anpassungen für Excel 2016/365: Application.Visible ersetzt durch SetApplicationVisible()
 '31.12.2013  Anpassungen für Excel 2010:
 '            - Speichern der Arbeitsmappe im aktuellen Standardformat.
 '09.11.2008  - Methoden zum Speichern und Wiederherstellen von Einstellungen (nur Datenquelle)
 '30.05.2004  - Formeln in Zellen werden erhalten beim Import/Export:
 '              - Bezüge auf Zellen im Tabellenkopf bleiben unverändert.
 '              - Bezüge auf Zellen im Datenbereich werden angepaßt entsprechend der 1. Zeile
 '                des Datenbereiches und der Lage der benannten Spalten. Formeln, die sich
 '                auf Spalten beziehen, die es in der Zieltabelle nicht gibt, werden entfernt.
 '29.04.2004  - Konfiguration und Anwendung von Ersatz-ZielSpalten.
 '22.02.2004  - Import/Export anders organisiert zwecks einheitlicher Benutzerschnittstelle:
 '              - Import/Export wird komplett von CdatExpim gesteuert, das alle notwendigen
 '                Steuerparameter per Benutzerdialog abfragt.
 '              - Die im Dialog angebotenen Optionen sind abhängig von der gewählten
 '                Datenquelle, z.B. wird die Dateidialogmaske und das Datenziel durch das
 '                zu einem Spezialformat gehörenden Objekt geliefert.
 '              - CdatExpim übernimmt (auch bei SpezialIm-/Export):
 '                - komplette Benutzerschnittstelle incl. Dateidialoge,
 '                - Existenzprüfung Quell-/Zieldatei, Anlegen der Zieltabelle usw.
 '                - Standard-Lese- und -Scheibvorgänge (nach und vom Datenpuffer)
 '                - Speichern der neuen Xl-Tab (bei Bedarf)
 '                - Wenn Ziel = XL-Tab: Übertragen Formeln und Formate, Fußzeile, Kopfdaten
 '              - Als Zielformate werden nur solche zur Auswahl angeboten, die Größen gleicher
 '                Kategorien wie das Quellformat enthalten. Alternativ können alle Formate
 '                angezeigt werden.
 '            - Lesen aus Excel-Tabelle: - benannte Spalten in den normierten Datenpuffer.
 '              - Änderung der aktiven Projektdaten entsprechend der Werte aus der Tabelle.
 '            - Schreiben des normierten Datenpuffers in eine neue Excel-Tabelle.
'==================================================================================================

Option Explicit

'Konstanten
Const strKlasse_Name                    As String = "CdatExpim"

'Eigenschaften
Dim strQuelle_AsciiDatei_DialogFilter   As String
Dim strQuelle_AsciiDatei_Name           As String
Dim strQuelle_FormatID                  As String
Dim strQuelle_Typ                       As String

Dim strZiel_AsciiDatei_DialogFilter     As String
Dim strZiel_AsciiDatei_Name             As String
Dim strZiel_AsciiDatei_Modus            As String
Dim strZiel_Typ                         As String
Dim strZiel_FormatID                    As String
Dim strZiel_XLTabKlasse                 As String
Dim blnZiel_XLMappeSpeichern            As Boolean

Dim blnMeldung_Ausgeben                 As Boolean
Dim blnDialog_OK                        As Boolean
Dim blnDialog_Anzeigen                  As Boolean

Dim bln_DatenModifizieren               As Boolean
Dim bln_ErsatzZielspaltenVerwenden      As Boolean
Dim bln_FormelnErhalten                 As Boolean

'Objekte
Dim oDatenpuffer          As CdatDatenpuffer
Dim oEinstellungen        As Scripting.Dictionary
'



Private Sub Class_Initialize()
  '
  On Error GoTo Fehler
  
  Dim pi        As Double
  Dim rho       As Double
  
  Set oDatenpuffer = New CdatDatenpuffer
  Set oEinstellungen = New Scripting.Dictionary
  Application.DisplayStatusBar = True
  
  'Mathematische Konstanten
  pi = 4 * Atn(1)
  rho = 200 / pi
  
  'Standard-Eigenschaften
  Call Me.SetzeStandardDatenEinstellungen
  
  blnDialog_Anzeigen = True
  blnDialog_OK = False
  blnMeldung_Ausgeben = True
  
  If (Not ThisWorkbook.Konfig Is Nothing) Then
    blnMeldung_Ausgeben = String2Bool(ThisWorkbook.Konfig.Expim_SchlussMeldung)
  End If
  
  'Cache für Einstellungen vorbereiten
  oEinstellungen.Add "Quelle_Typ", ""
  oEinstellungen.Add "Quelle_FormatID", ""
  oEinstellungen.Add "Quelle_AsciiDatei_Name", ""
  
  DebugEcho "CdatExpim.Class_Initialize() beendet."
  Exit Sub

Fehler:
  FehlerNachricht "CdatExpim.Initialize()"
  Set oDatenpuffer = Nothing
End Sub


Private Sub Class_Terminate()
  Set oDatenpuffer = Nothing
  DebugEcho "CdatExpim.Class_Terminate() beendet."
End Sub



'==== Eigenschaften ===========================================================

Property Get Datenpuffer() As CdatDatenpuffer
  'Bietet Zugriff auf den Datenpuffer.
  Set Datenpuffer = oDatenpuffer
End Property


Property Get Quelle_AsciiDatei_DialogFilter() As String
  'Liefert den Filter für den Dateidialog der Eingabedatei für Me.Import
  Quelle_AsciiDatei_DialogFilter = strQuelle_AsciiDatei_DialogFilter
End Property

Property Let Quelle_AsciiDatei_DialogFilter(inpQuelle_AsciiDatei_DialogFilter As String)
  'Liefert den Filter für den Dateidialog der Eingabedatei für Me.Import
  strQuelle_AsciiDatei_DialogFilter = inpQuelle_AsciiDatei_DialogFilter
End Property


Property Get Quelle_AsciiDatei_Name() As String
  'Liefert Pfad\Namen der Eingabedatei für Me.Import.
  Quelle_AsciiDatei_Name = strQuelle_AsciiDatei_Name
End Property

Property Let Quelle_AsciiDatei_Name(inpQuelle_AsciiDatei_Name As String)
  'Setzt den Namen der Eingabedatei für Me.Import
  strQuelle_AsciiDatei_Name = inpQuelle_AsciiDatei_Name
End Property


Property Get Quelle_FormatID() As String
  'Liefert einen eindeutigen Formatbezeichner für Me.Import.
  Quelle_FormatID = strQuelle_FormatID
End Property

Property Let Quelle_FormatID(inpQuelle_FormatID As String)
  'Setzt einen eindeutigen Formatbezeichner für Me.Import.
  strQuelle_FormatID = inpQuelle_FormatID
End Property


Property Get Quelle_Typ() As String
  'Liefert den Quelle-Typ der importierten Daten.
  Quelle_Typ = strQuelle_Typ
End Property

Property Let Quelle_Typ(inpQuelle_Typ As String)
  'Setzt den Quelle-Typ der importierten Daten.
  strQuelle_Typ = inpQuelle_Typ
End Property


Property Get Ziel_AsciiDatei_DialogFilter() As String
  'Liefert den Filter für den Dateidialog der Ausgabedatei für Me.Import
  Ziel_AsciiDatei_DialogFilter = strZiel_AsciiDatei_DialogFilter
End Property

Property Let Ziel_AsciiDatei_DialogFilter(inpZiel_AsciiDatei_DialogFilter As String)
  'Liefert den Filter für den Dateidialog der Ausgabedatei für Me.Import
  strZiel_AsciiDatei_DialogFilter = inpZiel_AsciiDatei_DialogFilter
End Property


Property Get Ziel_AsciiDatei_Name() As String
  'Liefert Pfad\Namen der Ausgabedatei für Me.Import.
  Ziel_AsciiDatei_Name = strZiel_AsciiDatei_Name
End Property

Property Let Ziel_AsciiDatei_Name(inpZiel_AsciiDatei_Name As String)
  'Setzt den Namen der Ausgabedatei für Me.Import
  strZiel_AsciiDatei_Name = inpZiel_AsciiDatei_Name
End Property


Property Get Ziel_AsciiDatei_Modus() As String
  'Liefert den Modus für die Ausgabedatei für Me.Import.
  Ziel_AsciiDatei_Modus = strZiel_AsciiDatei_Modus
End Property

Property Let Ziel_AsciiDatei_Modus(inpZiel_AsciiDatei_Modus As String)
  'Setzt den Modus für die Ausgabedatei für Me.Import.
  strZiel_AsciiDatei_Modus = inpZiel_AsciiDatei_Modus
End Property


Property Get Ziel_Typ() As String
  'Liefert den Ziel-Typ der importierten Daten.
  Ziel_Typ = strZiel_Typ
End Property

Property Let Ziel_Typ(inpZiel_Typ As String)
  'Setzt den Ziel-Typ der importierten Daten.
  strZiel_Typ = inpZiel_Typ
End Property


Property Get Ziel_XLMappeSpeichern() As Boolean
  'Liefert den aktuellen Modus für "neu angelegte Mappe speichern".
  Ziel_XLMappeSpeichern = blnZiel_XLMappeSpeichern
End Property

Property Let Ziel_XLMappeSpeichern(inpZiel_XLMappeSpeichern As Boolean)
  'Setzt den aktuellen Modus für "neu angelegte Mappe speichern"
  blnZiel_XLMappeSpeichern = inpZiel_XLMappeSpeichern
End Property


Property Get Ziel_FormatID() As String
  'Liefert einen eindeutigen Formatbezeichner für Me.Import.
  Ziel_FormatID = strZiel_FormatID
End Property

Property Let Ziel_FormatID(inpZiel_FormatID As String)
  'Setzt einen eindeutigen Formatbezeichner für Me.Import.
  strZiel_FormatID = inpZiel_FormatID
End Property


Property Get Ziel_XLTabKlasse() As String
  'Liefert den Namen der XL-Tabellen-Klasse (Kodename) für Me.Import
  Ziel_XLTabKlasse = strZiel_XLTabKlasse
End Property

Property Let Ziel_XLTabKlasse(inpZiel_XLTabKlasse As String)
  'Setzt den Namen der XL-Tabellen-Klasse (Kodename) für Me.Import
  strZiel_XLTabKlasse = inpZiel_XLTabKlasse
End Property


Property Get Dialog_OK() As Boolean
  'Liefert den aktuellen Status für "Dialog mit OK beendet".
  Dialog_OK = blnDialog_OK
End Property

Property Let Dialog_OK(inpDialog_OK As Boolean)
  'Setzt den aktuellen Status für "Dialog mit OK beendet".
  blnDialog_OK = inpDialog_OK
End Property


Property Get Dialog_Anzeigen() As Boolean
  'Liefert den aktuellen Status für "Dialog soll angezeigt werden".
  Dialog_Anzeigen = blnDialog_Anzeigen
End Property

Property Let Dialog_Anzeigen(inpDialog_Anzeigen As Boolean)
  'Setzt den aktuellen Status für "Dialog soll angezeigt werden".
  blnDialog_Anzeigen = inpDialog_Anzeigen
End Property


Property Get Meldung_Ausgeben() As Boolean
  'Liefert den aktuellen Modus für "Meldungen ausgeben" (False: keine Meldungen).
  Meldung_Ausgeben = blnMeldung_Ausgeben
End Property

Property Let Meldung_Ausgeben(inpMeldung_Ausgeben As Boolean)
  'Setzt den aktuellen Modus für "Meldungen ausgeben" (False: keine Meldungen).
  blnMeldung_Ausgeben = inpMeldung_Ausgeben
End Property


Property Get Opt_ErsatzZielspaltenVerwenden() As Boolean
  'Liefert die Option "Ersatz-Zielspalten verwenden".
  Opt_ErsatzZielspaltenVerwenden = bln_ErsatzZielspaltenVerwenden
End Property

Property Let Opt_ErsatzZielspaltenVerwenden(inp_ErsatzZielspaltenVerwenden As Boolean)
  'Setzt die Option "Ersatz-Zielspalten verwenden".
  bln_ErsatzZielspaltenVerwenden = inp_ErsatzZielspaltenVerwenden
End Property


Property Get Opt_FormelnErhalten() As Boolean
  'Liefert die Option "Formeln erhalten".
  Opt_FormelnErhalten = bln_FormelnErhalten
End Property

Property Let Opt_FormelnErhalten(inp_FormelnErhalten As Boolean)
  'Setzt die Option "Formeln erhalten".
  bln_FormelnErhalten = inp_FormelnErhalten
End Property


Property Get Opt_DatenModifizieren() As Boolean
  'Liefert die Option "Daten Modifizieren".
  Opt_DatenModifizieren = bln_DatenModifizieren
End Property

Property Let Opt_DatenModifizieren(inp_DatenModifizieren As Boolean)
  'Setzt die Option "Daten Modifizieren".
  bln_DatenModifizieren = inp_DatenModifizieren
End Property




'==== Methoden ================================================================

Sub AktionsManager()
  '1. Der Dialog zum Festlegen der gewünschten Import-/Export-Aktion wird gezeigt,
  '   es sei denn, daß dieser unterdrückt werden soll und ein gültiger Satz
  '   Import-/Export-Einstellungen bereits festgelegt ist.
  '2. Die so festgelegte Import-/Export-Aktion wird ausgeführt.
  
  'On Error GoTo Fehler
  'On Error GoTo 0
  
  Dim NameExcelDatei            As String
  Dim Meldung_Inhalt            As String
  Dim Meldung_Titel             As String
  Dim AktionErfolgreich         As Boolean
  Dim StatusScreen              As Boolean
  Dim DatenImPuffer             As Boolean
  Dim MappeAngelegt             As Boolean
  Dim DialogAnzeigen            As Boolean
  Dim Anzahl                    As Long
  Dim Fehlerniveau              As Long
  
  Dim oImportobjekt             As Object
  Dim oCSV                      As CtabCSV
  
  Dim Dialog    As frmStartExpim
  Set Dialog = New frmStartExpim
  
  DebugEcho "CdatExpim.AktionsManager() gestartet."
  
  'Dialog zum Festlegen der gewünschten Import-/Export-Aktion.
  
  'Prüfen, ob Dialog versteckt bleiben soll oder nicht.
  If (Me.Dialog_Anzeigen) Then
    DialogAnzeigen = true
  else
    'Wenn Dialog nicht alle nötigen Angaben bestimmen konnte => anzeigen!
    'Wenn mehr als ein Zielformat verfügbar ist => anzeigen!
    if ((Dialog.LstZiel_Formate.ListCount > 1) or (Not Dialog.Check_DialogOK)) then
      DialogAnzeigen = true
    else
      'Dialog wirklich nicht anzeigen :-)
      DialogAnzeigen = false
      Me.Dialog_OK   = true    'Bestätigung des Dialoges simulieren.
    end if
  end if
  
  'Dialog anzeigen.
  If (DialogAnzeigen) Then
    'Sicherstellen, dass Excel sichtbar ist.
    '(Bei Fernsteuerung via xlM.vbs wird Excel erst sichtbar geschaltet, wenn das Makro fertig ist...)
    
    SetApplicationVisible(true)
    Application.UserControl    = true
    Application.ScreenUpdating = true
    
    'Dialog anzeigen
    Dialog.Show
  end if
  
  'Dialog wurde mit "OK" bestätigt (oder gar nicht erst angezeigt).
  If (Me.Dialog_OK) Then
    
    'Kontrolle der aktiven Import-/Export-Einstellungen im Protokoll.
    Call ZeigeEinstellungen
    
    ThisWorkbook.AktiveTabelle.FormatDatenOptionenReset
    StatusScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    AktionErfolgreich = False
    Fehlerniveau = Fehlerniveau_Kein
    Meldung_Inhalt = ""
    Meldung_Titel = ""
    NameExcelDatei = ""
    Anzahl = 0
    
    
    '***  Quelle lesen  *****************************************************************************
    'Quelldaten aus der angegebenen Quelle in den normierten Datenpuffer
    'oder (evtl. bei io_Typ_AsciiSpezial) direkt in neue Tabelle lesen.
    
    Select Case Me.Quelle_Typ
      
      Case io_Typ_XlTabAktiv
          
          'Quelldaten aus der angegebenen Quelle in den normierten Datenpuffer lesen.
          Anzahl = GetQuelldaten_XlTabAktiv
          
          'Projekt- und ExtraDaten übernehmen
          ThisWorkbook.Metadaten.Update oPrjLocal:=ThisWorkbook.AktiveTabelle.GetProjektDaten, oExtraLocal:=ThisWorkbook.AktiveTabelle.GetExtraDaten
          
          'Importstatus festhalten.
          Meldung_Titel = "Export aus aktiver Tabelle"
          If (Anzahl > 0) Then
            DatenImPuffer = True
            MappeAngelegt = False
            AktionErfolgreich = True
            Meldung_Inhalt = Anzahl & " Zeilen gelesen aus Tabelle '" & ThisWorkbook.AktiveTabelle.TabName & "'"
          Else
            DatenImPuffer = False
            MappeAngelegt = False
            AktionErfolgreich = False
            Meldung_Inhalt = "Lesen aus Tabelle '" & ThisWorkbook.AktiveTabelle.TabName & "' ist fehlgeschlagen."
          End If
          
          
      Case io_Typ_CsvSpezial
          
          AktionErfolgreich = True
          
          'CSV-Tabellenobjekt instanzieren, dessen Eigenschaften setzen und abfragen.
          DebugEcho "CdatExpim.AktionsManager(): CSV-Importobjekt wird instanziert."
          Set oCSV = New CtabCSV
          
          If (oCSV Is Nothing) Then
            ErrEcho "CdatExpim.AktionsManager(): *** Programmfehler: CSV-Importobjekt konnte nicht instanziert werden!"
          Else
            'Dateiname setzen => Es erfolgt automatisch die Syncronisation von oCSV.
            oCSV.Quelle_AsciiDatei_Name = Me.Quelle_AsciiDatei_Name
            
            If (oCSV.Fehlerniveau = Fehlerniveau_Kein) Then
              'Formelbehandlung abschalten zwecks Performance und Problemvermeidung!
              Me.Opt_FormelnErhalten  = False
              
              'Import starten, Daten innerhalb von oCSV bereitstellen.
              oCSV.Import
              
              'Quelldaten in den normierten Datenpuffer übernehmen
              Call GetQuelldaten_CsvSpezial(oCSV)
              DatenImPuffer = True
              
              'Projekt- und ExtraDaten übernehmen
              ThisWorkbook.Metadaten.Update oPrjLocal:=oCSV.Quelle_Projektdaten, oExtraLocal:=oCSV.Quelle_DatenExtra
              
              NameExcelDatei = Verz(Me.Quelle_AsciiDatei_Name) & "\" & VorName(Me.Quelle_AsciiDatei_Name) '& ".xls"
              Me.Ziel_XLMappeSpeichern = oCSV.Ziel_XLMappeSpeichern
              
              AktionErfolgreich = oCSV.AktionErfolgreich
              Fehlerniveau = oCSV.Fehlerniveau
              Meldung_Inhalt = oCSV.Meldung_Inhalt
              Meldung_Titel = oCSV.Meldung_Titel
            End If
          End If
          
          
      Case io_Typ_AsciiFormatiert
          
          'Quelldaten aus der angegebenen Quelle in den normierten Datenpuffer lesen.
          MsgBox "Expim.AktionsManager:  NOCH  Nicht unterstützter Quell-Datentyp (" & Me.Quelle_Typ & ")"
          Exit Sub
          
          'Anzahl = GetQuelldaten_AsciiFormatiert
          DatenImPuffer = True
          MappeAngelegt = False
          AktionErfolgreich = True
          NameExcelDatei = Verz(Me.Quelle_AsciiDatei_Name) & "\" & VorName(Me.Quelle_AsciiDatei_Name) '& ".xls"
          
          
      Case io_Typ_AsciiSpezial
          
          AktionErfolgreich = True
          
          'Spezial-Importobjekt instanzieren, dessen Zieltyp abfragen und Eigenschaften setzen.
          Select Case Me.Quelle_FormatID
            Case io_Klasse_Trassenkoo
                Set oImportobjekt = New CimpTrassenkoo
            Case Else
                Err.Raise 66666 + vbObjectError, , "Spezial-Import-Klasse '" & Me.Ziel_FormatID & "' ist unbekannt!"
          End Select
          
          oImportobjekt.Quelle_AsciiDatei_Name = Me.Quelle_AsciiDatei_Name
          
          'Wenn Zieltyp nicht der Puffer ist, sondern eine neue XL-Tabelle, die also direkt vom
          'Importobjekt beschrieben wird, so muß die Tabelle an dieser Stelle angelegt werden!
          If (oImportobjekt.Ziel_Typ = io_Typ_XlTabNeu) Then
            MappeAngelegt = NeueMappeAnlegen
            If (Not MappeAngelegt) Then
              AktionErfolgreich = False
              Meldung_Inhalt = "Aktion abgebrochen"
              Meldung_Titel = "Abbruch"
            End If
          End If
          
          If (AktionErfolgreich) Then
            'Import starten, Daten innerhalb von oImportobjekt bereitstellen.
            oImportobjekt.Import
            
            'Quelldaten in den normierten Datenpuffer übernehmen, wenn nötig.
            If (oImportobjekt.Ziel_Typ = io_Typ_Puffer) Then
              'Zieltyp ist der Datenpuffer => Daten normiert in den Datenpuffer übernehmen.
              Call GetQuelldaten_AsciiSpezial(oImportobjekt)
              DatenImPuffer = True
            Else
              'Zieltyp ist nicht der Datenpuffer: Daten wurden bereits direkt ins Ziel geschrieben.
              DatenImPuffer = False
            End If
            
            'Extradaten übernehmen (unabhängig vom Ziel der eigentlichen Daten)
            ThisWorkbook.Metadaten.Update oPrjLocal:=Nothing, oExtraLocal:=oImportobjekt.Quelle_DatenExtra
            
            NameExcelDatei = Verz(Me.Quelle_AsciiDatei_Name) & "\" & VorName(Me.Quelle_AsciiDatei_Name) '& ".xls"
            
            AktionErfolgreich = oImportobjekt.AktionErfolgreich
            Fehlerniveau = oImportobjekt.Fehlerniveau
            Meldung_Inhalt = oImportobjekt.Meldung_Inhalt
            Meldung_Titel = oImportobjekt.Meldung_Titel
          End If
          
      Case Else
          MsgBox "Expim.AktionsManager: Nicht unterstützter Quell-Datentyp (" & Me.Quelle_Typ & ")"
            Exit Sub
    
    End Select
    '***  Quelle gelesen  ***************************************************************************
    
    
    If (Not AktionErfolgreich) Then
      'Quelle lesen war nicht erfolgreich.
      
      'Fehlermeldung ausgeben und eventuell bereits neu erstellte Datei wieder verwerfen.
      ErrEcho Meldung_Inhalt
      MsgBox Meldung_Inhalt, vbExclamation, Meldung_Titel
      If (MappeAngelegt) Then Application.ActiveWorkbook.Close False
      
    Else
      'Quelle prinzipiell erfolgreich gelesen.
      
      'Aber es können trotzdem meldepflichtige Warnungen/Fehler aufgetreten sein.
      If (Fehlerniveau = Fehlerniveau_Kritisch) Then
        ErrEcho Meldung_Inhalt
        MsgBox Meldung_Inhalt, vbCritical, Meldung_Titel
      ElseIf (Fehlerniveau = Fehlerniveau_Warnung) Then
        WarnEcho Meldung_Inhalt
        MsgBox Meldung_Inhalt, vbExclamation, Meldung_Titel
      End If
      
      
      '***  Manipulationen am Datenpuffer  ***********************************************************
      
      If (DatenImPuffer And Me.Opt_DatenModifizieren) Then
        Call Me.Datenpuffer.Daten_Bearbeiten
      End If
      
      '***  Manipulationen am Datenpuffer  ***********************************************************
      
      
      
      '***  Ziel schreiben  ***************************************************************************
      Select Case Me.Ziel_Typ
        
        Case io_Typ_XlTabNeu
            
            'Neue Mappe anlegen, falls nicht bereits erledigt.
            If (Not MappeAngelegt) Then MappeAngelegt = NeueMappeAnlegen
            
            If (Not MappeAngelegt) Then
              Meldung_Inhalt = "Aktion abgebrochen."
            Else
              'Mappe erfolgreich angelegt.
              
              'Spezial nach CSV-Import: Struktur der Zieltabelle evtl. neu herstellen.
              If (Me.Quelle_Typ = io_Typ_CsvSpezial) Then AktionErfolgreich = TabelleStrukturieren(oCSV)
              
              
              'Zieldaten aus dem normierten Datenpuffer in die aktive Tabelle schreiben,
              'falls die Daten im Puffer stehen (andernfalls sollten sie bereits in der Tabelle stehen).
              If (DatenImPuffer) Then
                Anzahl = SchreibeDatenpuffer_XlTabAktiv
                If (Anzahl > 0) Then
                  AktionErfolgreich = True
                  Meldung_Inhalt = Meldung_Inhalt & vbNewLine & vbNewLine & Anzahl & " Spalten geschrieben nach Tabelle '" & ThisWorkbook.AktiveTabelle.TabName & "'"
                Else
                  AktionErfolgreich = False
                  Meldung_Inhalt = Meldung_Inhalt & vbNewLine & vbNewLine & "Schreiben nach Tabelle '" & ThisWorkbook.AktiveTabelle.TabName & "' ist fehlgeschlagen."
                End If
              End If
              
              If (AktionErfolgreich) Then
                'Projekt- und andere Sitzungsdaten in dafür vorgesehene Felder schreiben.
                'ThisWorkbook.Metadaten.Update wird unmittelbar nach Import gemacht (abhängig von Quelle)
                ThisWorkbook.AktiveTabelle.SchreibeMetaDaten Silent:=True
                
                'Formatierung der 1. Datenzeile auf die gesamte Tabelle übertragen
                'Optionen vorher zurücksetzen, damit diese auch vom Importobjekt gesetzt werden können.
                'ThisWorkbook.AktiveTabelle.FormatDatenOptionenReset
                ThisWorkbook.AktiveTabelle.FormatDaten
                ThisWorkbook.AktiveTabelle.UebertragenFormeln
                'Cursor positionieren
                ThisWorkbook.AktiveTabelle.Infotraeger.Cells(1, 1).Select
                ThisWorkbook.AktiveTabelle.SchreibeFusszeile_1 Silent:=True
                
                'Neue Arbeitsmappe speichern (bei Bedarf).
                If (Me.Ziel_XLMappeSpeichern) Then
                  On Error Resume Next
                  ' NOTE: Aktuelles Speicherformat wird verwendet (kann also per Vorlagenmakro gesteuert werden)
                  'Application.ActiveWorkbook.SaveAs FileName:=NameExcelDatei, AddToMru:=True
                  Application.ActiveWorkbook.SaveAs FileFormat:=Application.DefaultSaveFormat, FileName:=NameExcelDatei, AddToMru:=True
                  If (Err) Then
                    ' Zielformat und Extension in jeder Hinsicht unsicher => Speichern dem Nutzer überlassen!
                    'NameExcelDatei = Application.GetSaveAsFilename(InitialFilename:=NameExcelDatei, FileFilter:=DateiFilterXLS, Title:="Arbeitsmappe speichern")
                    'If ((NameExcelDatei <> "False") And (NameExcelDatei <> "Falsch")) Then
                    '  'Application.ActiveWorkbook.SaveAs FileName:=NameExcelDatei, AddToMru:=True
                    '  Application.ActiveWorkbook.SaveAs FileFormat:=Application.DefaultSaveFormat, FileName:=NameExcelDatei, AddToMru:=True
                    'End If
                    MsgBox "Neu erstellte Datei konnte nicht gespeichert werden als '" & NameExcelDatei & ".xls*'." & vbNewLine & vbNewLine & " => Ist eine Datei gleichen Namens vielleicht bereits geöffnet?"
                  End If
                  On Error GoTo Fehler
                End If
              Else
                Application.ActiveWorkbook.Close False
              End If
              
            End If
            
            
        Case io_Typ_AsciiFormatiert
            
            'Zieldaten aus der angegebenen Ziel in den normierten Datenpuffer lesen.
            ErrEcho "Expim.AktionsManager:  NOCH  Nicht unterstützter Ziel-Datentyp (" & Me.Ziel_Typ & ")"
            MsgBox "Expim.AktionsManager:  NOCH  Nicht unterstützter Ziel-Datentyp (" & Me.Ziel_Typ & ")"
            Exit Sub
            
            
        Case io_Typ_AsciiSpezial
            
            'Spezial-Importobjekt instanzieren, dessen Zieltyp abfragen und Eigenschaften setzen.
            'set oImportobjekt = AsciiSpezialObjekt()
            ErrEcho "Expim.AktionsManager:  NOCH  Nicht unterstützter Ziel-Datentyp (" & Me.Ziel_Typ & ")"
            MsgBox "Expim.AktionsManager:  NOCH  Nicht unterstützter Ziel-Datentyp (" & Me.Ziel_Typ & ")"
            Exit Sub
            
            
        Case Else
            ErrEcho "Expim.AktionsManager: FEHLER *** Nicht unterstützter Ziel-Datentyp (" & Me.Ziel_Typ & ")"
            MsgBox "Expim.AktionsManager: FEHLER *** Nicht unterstützter Ziel-Datentyp (" & Me.Ziel_Typ & ")"
            Exit Sub
            
      End Select
      '***  Ziel geschrieben  *************************************************************************
      
      
      'Erfolgsmeldung (vorher Bildschirmaufbau, um Irritationen zu vermeiden).
      Application.ScreenUpdating = StatusScreen
      Echo vbNewLine & Replace(Replace(Meldung_Inhalt, vbNewLine & vbNewLine, vbNewLine), "Datei" & vbNewLine, "Datei ")
      If (Me.Meldung_Ausgeben Or (Not AktionErfolgreich)) Then MsgBox Meldung_Inhalt, vbInformation, Meldung_Titel
      
    End If
    
    
    'Kontrolle über Statuszeile an Excel zurückgeben, Status Bildschirmaufbau und Berechnung wiederherstellen.
    'Application.StatusBar = ""
    Application.ScreenUpdating = StatusScreen
    
  End If
  
  
  Set Dialog = Nothing
  On Error GoTo 0
  DebugEcho "CdatExpim.AktionsManager() beendet."
  If (Not ThisWorkbook.SysTools.FileErrorsShowInJEdit(True)) Then
    Call ShowConsole
  End If
  Exit Sub
  
Fehler:
  Set Dialog = Nothing
  FehlerNachricht "CdatExpim.AktionsManager()"
End Sub


Function GetQuelldaten_XlTabAktiv() As Long
  'Ermittelt die Struktur der aktiven Tabelle, liest alle Datenspalten (Werte und Formeln),
  'belegt damit den Datenpuffer und liest alle in der Tabelle enthaltenen Projektdaten.
  'Rückgabe: Anzahl der gelesenen Zeilen.
  '
  'Diese Routine wird während der Initialisierung des Add-Ins verwendet (ThisWorkbook.Konfig.initialize).
  '=> Die Verwendung der (erst danach initialisierten) globalen Objekte vermeiden!
  
  On Error GoTo Fehler
  
  Dim Anzahl             As Long
  Dim ZeAnz              As Long
  Dim oInfotraeger       As Range
  Dim Spalte             As Variant
  Dim i                  As Variant
  Dim k                  As Variant
  Dim Feld()             As String
  Dim Faktor             As Double
  
  Set oInfotraeger = ThisWorkbook.AktiveTabelle.Infotraeger
  If (oInfotraeger Is Nothing) Then
    ErrMessage = "Es ist kein 'InfoTraeger' vorhanden!"
    GoTo Fehler
  End If
  ZeAnz = ThisWorkbook.AktiveTabelle.AnzahlDatenZeilen
  
  If (isTabellenSchutz) Then
    Err.Raise vbObjectError + ErrNumTabSchutz - vbObjectError, , "Aktion ist nicht möglich, Tabellenschutz ist aktiv."
  End If
  'Eventuelle Mehrfachselection von Tabellen aufheben
  ActiveSheet.Select True
  
  
  'Datenpuffer löschen.
  Me.Datenpuffer.Daten_LoeschenAlleSpalten
  ReDim Feld(1 To ZeAnz, DP2lb To DP2ub)
  
  'Analyse der aktiven Tabelle: Ermitteln von Datenspalten, Einheiten und Faktoren der Quelldaten.
  ThisWorkbook.AktiveTabelle.Syncronisieren
  'ThisWorkbook.AktiveTabelle.ZeigeSpalten
  
  'Benannte Spalten der Tabelle lesen und normiert in den Datenpuffer schreiben.
  If (Not ThisWorkbook.AktiveTabelle.SpaltenErsteZellen Is Nothing) Then
    If (ThisWorkbook.AktiveTabelle.SpaltenErsteZellen.Count > 0) Then
      For Each Spalte In ThisWorkbook.AktiveTabelle.SpaltenErsteZellen
        Application.StatusBar = "Lese Spalte: " & Spalte
        For i = LBound(Feld) To UBound(Feld)
          For k = DP2lb To DP2ub
            Feld(i, k) = ""
          Next
        Next
        '"GetQuelldaten_XlTabAktiv" wird während der Initialisierung des Add-Ins verwendet (ThisWorkbook.Konfig.initialize).
        'Zu diesem Zeitpunkt existiert ThisWorkbook.Konfig noch nicht:
        If (Not ThisWorkbook.Konfig Is Nothing) Then
          Faktor = ThisWorkbook.Konfig.FaktorDerEinheit(ThisWorkbook.AktiveTabelle.SpaltenEinheiten(Spalte))
        Else
          Faktor = 0
        End If
        'MsgBox "ThisWorkbook.AktiveTabelle.SpaltenEinheiten(Spalte)=" & ThisWorkbook.AktiveTabelle.SpaltenEinheiten(Spalte) & ",   Faktor=" & Faktor
        Call Spalte_XL2DP(Feld, ThisWorkbook.AktiveTabelle.SpaltenErsteZellen(Spalte), Faktor)
        Me.Datenpuffer.Daten_AddSpalte Spalte, Feld
      Next
      Anzahl = ZeAnz
    End If
  Else
    Anzahl = 0
  End If
  
  Set oInfotraeger = Nothing
  'Application.StatusBar = ""
  GetQuelldaten_XlTabAktiv = Anzahl
  
  Exit Function
  
Fehler:
  GetQuelldaten_XlTabAktiv = 0
  Set oInfotraeger = Nothing
  FehlerNachricht "CdatExpim.GetQuelldaten_XlTabAktiv()"
End Function


Function SchreibeDatenpuffer_XlTabAktiv() As Long
  'Ermittelt die Struktur der aktiven Tabelle (Zieltabelle) und schreibt alle
  'verfügbaren Daten des Datenpuffers in die passenden Spalten der Tabelle.
  
  'On Error GoTo Fehler
  
  Dim AnzSpalten         As Long
  Dim oInfotraeger       As Range
  Dim Spalte             As Variant
  Dim SpalteErsatz       As String
  Dim SpalteOhneStatus   As String
  Dim SpalteGeschrieben  As Boolean
  Dim StatusScreen       As Boolean
  Dim StatusCalc         As Boolean
  Dim Faktor             As Double
  Dim oSpNameAttr        As Scripting.Dictionary
  
  AnzSpalten = 0
  
  StatusScreen = Application.ScreenUpdating
  StatusCalc = ActiveSheet.EnableCalculation
  Application.ScreenUpdating = False
  ActiveSheet.EnableCalculation = False
  
  Set oInfotraeger = ThisWorkbook.AktiveTabelle.Infotraeger
  If (oInfotraeger Is Nothing) Then
    ErrMessage = "Es ist kein 'InfoTraeger' vorhanden!"
    GoTo Fehler
  End If
  
  If (isTabellenSchutz) Then
    Err.Raise vbObjectError + ErrNumTabSchutz - vbObjectError, , "Aktion ist nicht möglich, Tabellenschutz ist aktiv."
  End If
  'Eventuelle Mehrfachselection von Tabellen aufheben
  ActiveSheet.Select True
  
  
  '1. Analyse der aktiven Tabelle: Ermitteln von Datenspalten, Einheiten und Faktoren der Quelldaten.
  ThisWorkbook.AktiveTabelle.Syncronisieren
  
  
  '2. Bei Notwendigkeit ErsatzZiele einrichten:
  '   a, entweder durch Kopieren von Spalten im Datenpuffer (Name => ErsatzName)
  '   b, oder durch Kopieren der einzelnen Werte, falls die Ersatzspalte im Puffer schon existiert.
  '   ==> funktioniert für alle WertStati.
  If (Me.Opt_ErsatzZielspaltenVerwenden) Then
    For Each Spalte In Me.Datenpuffer.Daten
      If (Not ThisWorkbook.AktiveTabelle.SpaltenErsteZellen.Exists(Spalte)) Then
        'Spalte des Datenpuffers existiert in Zieltabelle nicht.
        Set oSpNameAttr = ThisWorkbook.Konfig.SpNameAttr(Spalte)
        SpalteOhneStatus = oSpNameAttr("NamePur")
        If (SpalteOhneStatus <> SpName_unbekannt) Then
          'Für die Spalte existiert eine Konfiguration.
          If (ThisWorkbook.Konfig.SpaltenErsatzZiel(SpalteOhneStatus) <> "") Then
            'Für "Spalte" ist ein Ersatzziel konfiguriert.
            SpalteErsatz = oSpNameAttr("StatusPrefix") & ThisWorkbook.Konfig.SpaltenErsatzZiel(SpalteOhneStatus)
            If (ThisWorkbook.AktiveTabelle.SpaltenErsteZellen.Exists(SpalteErsatz)) Then
              'Ersatzspalte existiert in Zieldatei => Verwenden!.
              Call Me.Datenpuffer.Daten_SchreibenErsatzSpalte(Spalte)
            End If
          End If
        End If
      End If
    Next
  End If
  
  
  '3. Datenpuffer schreiben.
  For Each Spalte In Me.Datenpuffer.Daten
    If (ThisWorkbook.AktiveTabelle.SpaltenErsteZellen.Exists(Spalte)) Then
      'Spalte des Datenpuffers existiert in Zieltabelle und wird dorthin geschrieben.
      
      'Meldung in Statuszeile ausgeben.
      Set oSpNameAttr = ThisWorkbook.Konfig.SpNameAttr(Spalte)
      If (oSpNameAttr("Titel") <> SpTitel_unbekannt) Then
        Application.StatusBar = "Schreibe Spalte: " & oSpNameAttr("StatusBez") & ": " & oSpNameAttr("Titel")
      Else
        Application.StatusBar = "Schreibe Spalte: " & Spalte
      End If
      
      Faktor = ThisWorkbook.Konfig.FaktorDerEinheit(ThisWorkbook.AktiveTabelle.SpaltenEinheiten(Spalte))
      'MsgBox "ThisWorkbook.AktiveTabelle.SpaltenEinheiten(Spalte)=" & ThisWorkbook.AktiveTabelle.SpaltenEinheiten(Spalte) & ",   Faktor=" & Faktor
      Call Spalte_DP2XL(Me.Datenpuffer.Daten(Spalte), ThisWorkbook.AktiveTabelle.SpaltenErsteZellen(Spalte), Faktor)
      AnzSpalten = AnzSpalten + 1
    End If
  Next
  
  
  Application.ScreenUpdating = StatusScreen
  ActiveSheet.EnableCalculation = StatusCalc
  Set oInfotraeger = Nothing
  Set oSpNameAttr = Nothing
  SchreibeDatenpuffer_XlTabAktiv = AnzSpalten
  
  Exit Function
  
Fehler:
  SchreibeDatenpuffer_XlTabAktiv = AnzSpalten
  Set oInfotraeger = Nothing
  Set oSpNameAttr = Nothing
  Application.ScreenUpdating = StatusScreen
  If (Not ActiveSheet Is Nothing) Then ActiveSheet.EnableCalculation = StatusCalc
  FehlerNachricht "CdatExpim.SchreibeDatenpuffer_XlTabAktiv()"
End Function


Sub SetzeStandardDatenEinstellungen()
  'Alle Einstellungen auf Standardwerte zurücksetzen.
  On Error GoTo Fehler
  
  DebugEcho "CdatExpim.SetzeStandardDatenEinstellungen(): Einstellungen für Datenbearbeitung und -übertragung werden zurückgesetzt."
  
  'Für einige Einstellungen: Konfiguration übernehmen, falls verfügbar
  If (Not ThisWorkbook.Konfig Is Nothing) Then
    bln_DatenModifizieren = String2Bool(ThisWorkbook.Konfig.Expim_Modifizieren)
    bln_ErsatzZielspaltenVerwenden = String2Bool(ThisWorkbook.Konfig.Expim_ErsatzSpalten)
    bln_FormelnErhalten = String2Bool(ThisWorkbook.Konfig.Expim_FormelnErhalten)
  Else
    DebugEcho "CdatExpim.SetzeStandardDatenEinstellungen(): Keine Konfiguration verfügbar => Setze hart kodierte Standard-Einstellungen."
    bln_DatenModifizieren = Std_DatenModifizieren
    bln_ErsatzZielspaltenVerwenden = Std_ErsatzZielspaltenVerwenden
    bln_FormelnErhalten = Std_FormelnErhalten
  End If
  
  'Einstellungen des angehängten Datenpuffers ebenfalls zurücksetzen.
  Me.Datenpuffer.SetzeStandardEinstellungen
  
  Exit Sub
  
Fehler:
  FehlerNachricht "CdatExpim.SetzeStandardDatenEinstellungen()"
End Sub


Sub EinstellungenSpeichern()
  'Speichert einige Eigenschaften im Cache.
  On Error GoTo 0
  
  Dim oKonfig As CdatKonfig
  
  'Datenquelle
  oEinstellungen("Quelle_Typ") = Me.Quelle_Typ
  oEinstellungen("Quelle_FormatID") = Me.Quelle_FormatID
  oEinstellungen("Quelle_AsciiDatei_Name") = Me.Quelle_AsciiDatei_Name
  
  If (ThisWorkbook.Konfig.Cache.Exists("Expim")) Then
    Set oKonfig = ThisWorkbook.Konfig  ' hier in VBA nötig, weil es nur "Getter" gibt !?
    Set oKonfig.Cache("Expim") = oEinstellungen
  Else
    ThisWorkbook.Konfig.Cache.Add "Expim", oEinstellungen
  End If
End Sub


Sub EinstellungenWiederherstellen()
  'Setzt einige der zuletzt im Cache gespeicherten Eigenschaften.
  On Error GoTo 0
  
  'Datenquelle
  If (ThisWorkbook.Konfig.Cache.Exists("Expim")) Then
    Me.Quelle_Typ = ThisWorkbook.Konfig.Cache("Expim")("Quelle_Typ")
    Me.Quelle_FormatID = ThisWorkbook.Konfig.Cache("Expim")("Quelle_FormatID")
    Me.Quelle_AsciiDatei_Name = ThisWorkbook.Konfig.Cache("Expim")("Quelle_AsciiDatei_Name")
  End If
End Sub


Sub ZeigeEinstellungen()
  'Zeigt für Kontrollzwecke alle aktiven Einstellungen für Import/Export an.
  Dim Spalte As Variant
  Dim Message As String
  
  Message = vbNewLine & "--- Starte Import/Export-Vorgang ----------------------------------------------- " & vbNewLine
  
  Message = Message & vbNewLine & "Quelle Typ:      " & Me.Quelle_Typ
  If ((Me.Quelle_Typ <> io_Typ_XlTabAktiv) And (Me.Quelle_Typ <> io_Typ_CsvSpezial)) Then Message = Message & vbNewLine & "Quelle Format:   " & Mid$(Me.Quelle_FormatID, Len(io_Klasse_PrefixImport) + 1)
  If (Me.Quelle_AsciiDatei_Name <> "") Then Message = Message & vbNewLine & "Quelle Datei:    " & Me.Quelle_AsciiDatei_Name
  Message = Message & vbNewLine
  
  Message = Message & vbNewLine & "Ziel Typ:        " & Me.Ziel_Typ
  Message = Message & vbNewLine & "Ziel Format:     " & Me.Ziel_FormatID
  If (Me.Ziel_AsciiDatei_Name <> "") Then Message = Message & vbNewLine & "Ziel Datei:      " & Me.Ziel_AsciiDatei_Name
  If (Me.Ziel_AsciiDatei_Modus <> "") Then Message = Message & vbNewLine & "Ziel Dateimodus: " & Me.Ziel_AsciiDatei_Modus
  Message = Message & vbNewLine
  
  Echo Message
End Sub




'==== interne Routinen ========================================================


Private Function GetQuelldaten_AsciiSpezial(oImportobjekt As Object)
  'Die durch das spezifische Importobjekt gelesenen Quelldaten werden normiert
  'und in den Datenpuffer übernommen.
  'Parameter:       oImportobjekt ... Instanz einer Klasse zum Datenimport
  'Voraussetzungen: Das Importobjekt bietet folgende Eigenschaften:
  '                 - .Quelle_Datenpuffer ... Dictionary mit Quelldaten (Key=Spaltenname, Item=Datenvektor)
  '                 - .Quelle_Einheiten   ... Dictionary mit Quelleinheiten (Key=Spaltenname, Item=Einheit)
  '                 - .Quelle_DatenExtra  ... Dictionary mit Extradaten (Statistik usw.) (Key=Feldname, Item=Inhalt)
  '(noch) nicht realisiert:   Rückgabe: Anzahl der gelesenen Zeilen.
  
  On Error GoTo Fehler
  
  'Dim Anzahl    As Long
  Dim Spalte    As Variant
  Dim Faktor    As Double
  Dim lb        As Long
  Dim ub        As Long
  Dim i         As Long
  Dim ZielWert  As Variant
  Dim QuellWert As Variant
  Dim Vektor()  As String
  
  
  'Datenpuffer löschen.
  Me.Datenpuffer.Daten_LoeschenAlleSpalten
  
  For Each Spalte In oImportobjekt.Quelle_Datenpuffer
    Faktor = ThisWorkbook.Konfig.FaktorDerEinheit(oImportobjekt.Quelle_Einheiten(Spalte))
    'MsgBox Spalte & ": " & oImportobjekt.Quelle_Einheiten(Spalte) & "=" & CStr(Faktor)
    If ((Faktor = 0) Or (Faktor = 1)) Then
      'Kein Umrechnungsfaktor für die Einheit verfügbar oder Faktor=1.
      Me.Datenpuffer.Daten_AddSpalte Spalte, oImportobjekt.Quelle_Datenpuffer(Spalte)
      'Me.Datenpuffer.Daten.Add Spalte, oImportobjekt.Quelle_Datenpuffer(Spalte)
    Else
      'Umrechnung für numerische Werte erforderlich
      lb = LBound(oImportobjekt.Quelle_Datenpuffer(Spalte))
      ub = UBound(oImportobjekt.Quelle_Datenpuffer(Spalte))
      ReDim Vektor(lb To ub)
      For i = lb To ub
        QuellWert = oImportobjekt.Quelle_Datenpuffer(Spalte)(i)
        If (IsNumeric(QuellWert)) Then
          ZielWert = QuellWert / Faktor
        Else
          ZielWert = QuellWert
        End If
        Vektor(i) = ZielWert
        'MsgBox Spalte & ": Quellwert='" & Quellwert & "'  Zielwert='" & Zielwert & "'  (" & me.Datenpuffer.Daten(Spalte)(i) & ")"
        'MsgBox "existiert='" & "me.Datenpuffer.Daten(" & Spalte & ")(" & i & ") ..." & me.Datenpuffer.Daten.Exists(Spalte)
      Next
      Me.Datenpuffer.Daten_AddSpalte Spalte, Vektor
      'Me.Datenpuffer.Daten.Add Spalte, Vektor
    End If
  Next
  
  'GetQuelldaten_AsciiSpezial = Anzahl
  On Error GoTo 0
  Exit Function
  
Fehler:
  'GetQuelldaten_AsciiSpezial = 0
  FehlerNachricht "CdatExpim.GetQuelldaten_AsciiSpezial()"
End Function


Private Function GetQuelldaten_CsvSpezial(oCSV As CtabCSV) As Long
  'Die durch das CSV-Importobjekt gelesenen Quelldaten werden normiert
  'und in den Datenpuffer übernommen.
  'Parameter:       oCSV ... Instanz des CSV-Objektes
  'Voraussetzungen: - oCSV.Quelle_AsciiDatei_Name ist bereits gesetzt und oCSV damit syncronisiert
  '                 - siehe GetQuelldaten_AsciiSpezial()
  '(noch) nicht realisiert:   Rückgabe: Anzahl der gelesenen Zeilen.
  
  On Error GoTo Fehler
  'Dim Anzahl             As Long
  
  'Übernahme von eventuell in der CSV-Datei gesetzten Importoptionen.
  'Diese Optionen wurden zwar bereits angewendet, nachdem die Datei gewählt wurde. Die Datei
  'könnte aber verändert worden sein...
  '=> Nicht konvertierbare Werte werden übergangen
  On Error Resume Next
  Me.Opt_ErsatzZielspaltenVerwenden = CBool(oCSV.Opt_ErsatzZielspaltenVerwenden)
  Me.Opt_DatenModifizieren = CBool(oCSV.Opt_DatenModifizieren)
  Me.Datenpuffer.Opt_VorhWerteUeberschreiben = CBool(oCSV.Opt_VorhWerteUeberschreiben)
  Me.Datenpuffer.Opt_UebAusInfoStreng = CBool(oCSV.Opt_UebAusInfoStreng)
  Me.Datenpuffer.Opt_FehlerVerbesserungen = CBool(oCSV.Opt_FehlerVerbesserungen)
  Me.Datenpuffer.Opt_UeberhoehungAusBemerkung = CBool(oCSV.Opt_UeberhoehungAusBemerkung)
  Me.Datenpuffer.Opt_Transfo_Tk2Gls = CBool(oCSV.Opt_Transfo_Tk2Gls)
  On Error GoTo Fehler
  
  'Daten lesen.
  Call GetQuelldaten_AsciiSpezial(oCSV)
  
  
  'GetQuelldaten_CsvSpezial = Anzahl
  Exit Function
  
Fehler:
  'GetQuelldaten_CsvSpezial = 0
  FehlerNachricht "CdatExpim.GetQuelldaten_CsvSpezial()"
End Function


Private Function TabelleStrukturieren(oCSV As CtabCSV) As Boolean
  'Spezial nach CSV-Import: Struktur der Zieltabelle prüfen:
  'Falls keine Spaltennamen existieren => komplette Struktur neu anlegen.
  'Parameter:       oCSV ... Instanz des CSV-Objektes
  'Rückgabe: false, wenn nicht erfolgreich, sonst true.
  
  On Error GoTo Fehler
  
  Dim Erfolg                  As Boolean
  Dim blnInfotraegerExist     As Boolean
  Dim blnSpaltenNamenExist    As Boolean
  Dim QuellSpaltenNamen       As Variant
  Dim QSp                     As Variant
  Dim ZSp                     As Long
  Dim Zeile                   As Long
  Dim AnzSpalten              As Long
  Dim SpName                  As String
  Dim Einheit                 As String
  Dim Zielformat              As String
  Dim StandardFormatExistiert As Boolean
  Dim StandardFormat          As String
  Dim oRange                  As Range
  
  Erfolg = True
  
  If (Me.Quelle_Typ = io_Typ_CsvSpezial) Then
    'Aktive Tabelle erkunden
    blnInfotraegerExist = (Not (ThisWorkbook.AktiveTabelle.Infotraeger Is Nothing))
    If (blnInfotraegerExist) Then
      If (Not (ThisWorkbook.AktiveTabelle.SpaltenErsteZellen Is Nothing)) Then
        If (ThisWorkbook.AktiveTabelle.SpaltenErsteZellen.Count > 0) Then blnSpaltenNamenExist = True
      End If
    End If
    
    'Wenn kein Infotraeger bzw. kein Spaltenname vorhanden => Struktur neu anlegen.
    If ((Not blnInfotraegerExist) Or (Not blnSpaltenNamenExist)) Then
      DebugEcho "CdatExpim.TabelleStrukturieren: Tabelle wird neu strukturiert..."
      'Vorbereitungen
      QuellSpaltenNamen = oCSV.Quelle_SpaltenNamen
      AnzSpalten = UBound(QuellSpaltenNamen)
      ZSp = ThisWorkbook.AktiveTabelle.ErsteDatenSpalte
      Zeile = ThisWorkbook.AktiveTabelle.ErsteDatenZeile
      If (Zeile < 2) Then Zeile = 2
      'StandardFormat bestimmen
      If (oCSV.Ziel_Formate.Exists(CsvAllOtherColumns)) Then
        StandardFormatExistiert = True
        StandardFormat = oCSV.Ziel_Formate(CsvAllOtherColumns)
      End If
      
      'Infotraeger einfügen bzw. ersetzen
      DebugEcho "CdatExpim.TabelleStrukturieren: Infotraeger wird festgelegt."
      Set oRange = ActiveWorkbook.ActiveSheet.Cells(Zeile, ZSp).Resize(ColumnSize:=AnzSpalten)
      Call ThisWorkbook.AktiveTabelle.SetInfotraeger(oRange)
      
      'Spaltennamen festlegen
      For QSp = 1 To AnzSpalten
        SpName = QuellSpaltenNamen(QSp)
        'Einheit = oCSV.Quelle_Einheiten(SpName)
        Einheit = oCSV.Ziel_Einheiten(SpName)
        DebugEcho "CdatExpim.TabelleStrukturieren: " & CStr(QSp) & ". Spalte wird festgelegt: Name = " & SpName & "   Einheit = " & Einheit
        Set oRange = ActiveWorkbook.ActiveSheet.Cells(Zeile, ZSp)
        Call ThisWorkbook.AktiveTabelle.SetSpaltenname(oRange, SpName, Einheit, blnTitel:=True)
        
        'Spalte mit Zahlenformat formatieren.
        On Error Resume Next
        If (oCSV.Ziel_Formate.Exists(SpName)) Then
          Zielformat = oCSV.Ziel_Formate(SpName)
          oRange.Resize(oRange.End(xlDown).Row - oRange.Row + 1).NumberFormat = Zielformat
          If (Err.Number = 0) Then
            DebugEcho "CdatExpim.TabelleStrukturieren:   ... Zahlenformat = '" & Zielformat & "' => erfolgreich festgelegt."
          Else
            ErrEcho "CdatExpim.TabelleStrukturieren:   ... Zahlenformat = '" & Zielformat & "' => konnte nicht festgelegt werden!"
            Err.Clear
          End If
        ElseIf (StandardFormatExistiert) Then
          Zielformat = StandardFormat
          oRange.Resize(oRange.End(xlDown).Row - oRange.Row + 1).NumberFormat = Zielformat
          If (Err.Number = 0) Then
            DebugEcho "CdatExpim.TabelleStrukturieren:        ... Standard-Zahlenformat = '" & Zielformat & "' => erfolgreich festgelegt."
          Else
            ErrEcho "CdatExpim.TabelleStrukturieren:        ... Standard-Zahlenformat = '" & Zielformat & "' => konnte nicht festgelegt werden!"
            Err.Clear
          End If
        End If
        
        On Error GoTo Fehler
        ZSp = ZSp + 1
      Next
      
      'Geänderte Tabellenstruktur reflektieren
      Call ThisWorkbook.AktiveTabelle.Syncronisieren
    End If
  End If
  
  TabelleStrukturieren = Erfolg
  Exit Function
  
Fehler:
  TabelleStrukturieren = False
  FehlerNachricht "CdatExpim.TabelleStrukturieren()"
End Function


Private Function NeueMappeAnlegen() As Boolean
  'Anlegen einer neuen Arbeitsmappe zum Schreiben der Zieldaten.
  'Tabellenklasse der neuen Tabelle prüfen, falls erforderlich, sowie Tabellenschutz prüfen.
  'Rückgabe: "true" bei Erfolg, sonst false.
  
  On Error GoTo Fehler
  'On Error GoTo 0
  
  'Dim VorlageMappePfadName   As String
  
  ''Vorlagendatei in den Standardpfaden finden
  'VorlageMappePfadName = FindeXLVorlage(Me.)
  'If (VorlageMappePfadName = "") Then
  '  ErrMessage = ErrMsgXLVorlageFehlt
  '  Err.Raise ErrNumXLVorlageFehlt, , "Arbeitsmappen-Vorlage '" & Me. & "' nicht gefunden."
  'End If
  'Application.Workbooks.Add VorlageMappePfadName
  
  Application.Workbooks.Add Me.Ziel_FormatID
  
  ''Tabellenklasse prüfen, falls erforderlich, sowie Tabellenschutz prüfen.
  '*** Me.Ziel_XLTabKlasse wird bisher nie gesetzt!
  '*** <Importobjekt.Ziel_XLTabKlasse> wird bisher nie verwendet!
  'If ((Me.Ziel_XLTabKlasse <> "") And (ThisWorkbook.AktiveTabelle.TabKlasse <> Me.Ziel_XLTabKlasse)) Then
  '  Application.ActiveWorkbook.Close False
  '  ErrMessage = ErrMsgTabKlasseUngueltig
  '  Err.Raise vbObjectError + ErrNumTabKlasseUngueltig - vbObjectError, , "Ungültige Vorlage '" & ThisWorkbook.AktiveTabelle.TabKlasse & "'. Erwartet wird '" & Me.Ziel_XLTabKlasse & "'  (Vorlagenmappe='" & Me.Ziel_FormatID & "')."
  'End If
  If (isTabellenSchutz) Then
    Application.ActiveWorkbook.Close False
    ErrMessage = ""
    Err.Raise vbObjectError + ErrNumTabSchutz - vbObjectError, , "Aktion ist nicht möglich, Tabellenschutz in neu angelegter Tabelle ist aktiv."
  End If
  'Eventuelle Mehrfachselection von Tabellen aufheben
  ActiveSheet.Select True
  NeueMappeAnlegen = True
  
  Exit Function
  
Fehler:
  NeueMappeAnlegen = False
  FehlerNachricht "CdatExpim.NeueMappeAnlegen()"
End Function


Private Sub Spalte_DP2XL(Vektor, oErsteZelle As Range, ByVal Faktor As Double)
  'Schreibt eine Spalte des Datenpuffers als Spalte in die aktive XL-Tabelle,
  'wobei in der Zelle begonnen wird, die durch "oErsteZelle" angegeben ist.
  'Falls "Faktor" angegeben ist (ungleich 0, 1), werden numerische Werte mit
  'diesem multipliziert.
  'Falls "Me.Opt_FormelnErhalten = true" und die Formel der Zelle ist nicht
  'leer, so wird die Formel in die Tabelle geschrieben. In allen anderen Fällen
  'wird der Wert geschrieben.
  
  On Error GoTo Fehler
  
  Dim lb                 As Long
  Dim i                  As Long
  Dim AnzD               As Long
  Dim QuellWert          As Variant
  Dim ZielWert           As Variant
  Dim QuellFormel        As String
  Dim ZielFormel         As String
  Dim FormelGeschrieben  As Boolean
  
  AnzD = CountDim(Vektor)
  
  If (Not oErsteZelle Is Nothing) Then
    lb = LBound(Vektor)
    For i = lb To UBound(Vektor)
      If (AnzD = 1) Then
        QuellWert = Vektor(i)
        QuellFormel = ""
      Else
        QuellWert = Vektor(i, DPidxWert)
        QuellFormel = Vektor(i, DPidxFormel)
      End If
      ZielWert = QuellWert
      ZielFormel = QuellFormel
      
      If ((Faktor <> 0) And (Faktor <> 1)) Then
        If (IsNumeric(QuellWert) And (Not (IsEmpty(QuellWert) Or (IsNull(QuellWert))))) Then
          ZielWert = QuellWert * Faktor
        End If
        If (ZielFormel <> "") Then
          ZielFormel = "=(" & Mid$(QuellFormel, 2) & ")*" & CStr(Faktor)
        End If
      End If
      
      'Wert oder Formel in Tabelle eintragen - dabei Fehler abfangen!
      FormelGeschrieben = false
      If (Me.Opt_FormelnErhalten) Then
        ZielFormel = Formel_DP2XL(ZielFormel)
        If (ZielFormel <> "") Then
          on error resume next
          oErsteZelle.offset(i - lb).Formula = ZielFormel
          if (err.number = 0) then
            FormelGeschrieben = true
          else
            WarnEcho "Die Formel '" & ZielFormel & "' konnte nicht in die Zelle '" & oErsteZelle.offset(i - lb).Address & "' geschrieben werden! => Versuche statt dessen, den Wert zu schreiben."
            err.clear
          end if
          On Error GoTo Fehler
        End If
      End If
      
      'Wert in Tabelle eintragen - dabei Fehler abfangen!
      if (not FormelGeschrieben) then
        on error resume next
        oErsteZelle.offset(i - lb).value = ZielWert
        if (err.number <> 0) then
          '2. Versuch: Notvariante mit führendem Hochkomma.
          WarnEcho "Der Wert '" & ZielWert & "' konnte nicht in die Zelle '" & oErsteZelle.offset(i - lb).Address & "' geschrieben werden! => Versuche, ein Hochkomma davor zu setzen..."
          oErsteZelle.offset(i - lb).value = "'" & ZielWert
          err.clear
        end if
        On Error GoTo Fehler
      End If
      
    Next
  End If
  Exit Sub
  
Fehler:
  FehlerNachricht "CdatExpim.Spalte_DP2XL()"
End Sub


Private Sub Spalte_XL2DP(Vektor, oErsteZelle As Range, ByVal Faktor As Double)
  'Liest die Zellinhalte und Formeln einer Tabellenspalte in die angegebene
  'Spalte des Datenpuffers, wobei in der Zelle begonnen wird, die durch
  '"oErsteZelle" angegeben ist und so viele Zellen gelesen werden, bis der
  'Vektor "voll" ist, d.h. das Vektorfeld muß bereits korrekt dimensioniert sein.
  'Falls "Faktor" angegeben ist, werden numerische Werte durch diesen dividiert
  'und Formeln werden mit dem Faktor als Divisor ergänzt.
  '==> Enthält eine Zelle keine Formel, so ist das entsprechende Feld im "Vektor" = ""
  
  On Error GoTo Fehler
  
  Dim lb          As Long
  Dim i           As Long
  Dim ZielWert    As Variant
  Dim QuellWert   As Variant
  Dim QuellFormel As String
  Dim ZielFormel  As String
  
  If (Not oErsteZelle Is Nothing) Then
    lb = LBound(Vektor)
    For i = lb To UBound(Vektor)
      QuellWert = oErsteZelle.offset(i - lb).value
      ZielWert = QuellWert
      QuellFormel = oErsteZelle.offset(i - lb).Formula
      If ((QuellWert = QuellFormel) Or (Left$(QuellFormel, 1) <> "=") Or (Len(QuellFormel) < 2)) Then
        ZielFormel = ""
      Else
        ZielFormel = QuellFormel
      End If
      If ((Faktor <> 0) And (Faktor <> 1)) Then
        If (IsNumeric(QuellWert) And (Not (IsEmpty(QuellWert) Or (IsNull(QuellWert))))) Then
          ZielWert = QuellWert / Faktor
        End If
        If (ZielFormel <> "") Then
          ZielFormel = "=(" & Mid$(QuellFormel, 2) & ")/" & CStr(Faktor)
        End If
      End If
      Vektor(i, DPidxWert) = ZielWert
      Vektor(i, DPidxFormel) = Formel_XL2DP(ZielFormel)
    Next
  End If
  Exit Sub
  
Fehler:
  FehlerNachricht "CdatExpim.Spalte_XL2DP()"
End Sub


Private Function Formel_XL2DP(ByVal Formel_XL As String)
  'Ersetzt in einer "normalen" Excel-Formel alle Spaltenangaben durch
  'die Add-In-internen Spaltennamen in der Form "@@<Name>@@", und die
  'Zeilenangaben durch relative Angaben im Datenpuffer, d.h. ein Bezug
  'auf die erste Zeile des Datenpuffers ist = 1.
  'Funktionswert: geänderte Formel zur Verwendung im Datenpuffer
  '=> Gibt es in der Formel auch nur einen Bezug auf eine Spalte, für die
  '   kein Add-In-interner Name geführt wird, so ist das Ergebnis eine
  '   leere Zeichenkette!
  'Vorausetzung: Die Zieltabelle muß bereits aktiv sein!
  
  On Error GoTo Fehler
  Dim Zellbezug
  Dim Zellbezuege
  Dim Meldung
  Dim SpNameErsatz   As String
  Dim ZeileErsatz    As String
  Dim BezugErsatz    As String
  Dim BezugSpalte    As String
  Dim ZeilePrefix    As String
  Dim BezugZeile     As String
  Dim Formel_DP      As String
  Dim ZeAnf          As Long
  
  Formel_DP = Formel_XL
  
  If (Formel_XL <> "") Then
    'Festwerte des Datenbereiches ermitteln
    ZeAnf = ThisWorkbook.AktiveTabelle.ErsteDatenZeile
    'If (Err) Then GoTo Fehler
    
    ThisWorkbook.RegExp.Pattern = "([a-zA-Z]+)(\$?)([0-9]+)"
    ThisWorkbook.RegExp.IgnoreCase = True
    ThisWorkbook.RegExp.Global = True
    Set Zellbezuege = ThisWorkbook.RegExp.Execute(Formel_XL)  'Matches-Auflistung.
    
    Meldung = "Entsprechungen gefunden:" & vbCrLf & vbCrLf
    For Each Zellbezug In Zellbezuege
      BezugSpalte = Zellbezug.SubMatches(0)
      ZeilePrefix = Zellbezug.SubMatches(1)
      BezugZeile = Zellbezug.SubMatches(2)
      Meldung = Meldung & "bei Zeichen: " & Zellbezug.FirstIndex & vbTab
      Meldung = Meldung & "Länge=" & Zellbezug.Length & vbTab
      Meldung = Meldung & "String='" & Zellbezug & "'" & vbCrLf
      
      If (ThisWorkbook.AktiveTabelle.SpaltenBuchst2Name.Exists(BezugSpalte)) Then
        'Spalte bezieht sich auf einen Spaltennamen im Sinne des Add-In.
        If (BezugZeile >= ZeAnf) Then
          'Zeile bezieht sich auf den Datenbereich (nicht Tabellenkopf).
          SpNameErsatz = "@@" & ThisWorkbook.AktiveTabelle.SpaltenBuchst2Name(BezugSpalte) & "@@"
          ZeileErsatz = CStr(CLng(BezugZeile) - ZeAnf + 1)
          BezugErsatz = SpNameErsatz & ZeilePrefix & ZeileErsatz
          Formel_DP = Replace(Formel_DP, Zellbezug, BezugErsatz)
        Else
          'Zeile bezieht sich auf den Tabellenkopf => Bezug nicht ändern.
        End If
      Else
        Formel_DP = ""
        Exit For
      End If
    Next
  End If
  
  'DebugEcho Meldung
  Set Zellbezuege = Nothing
  Formel_XL2DP = Formel_DP
  Exit Function

Fehler:
  FehlerNachricht "CdatExpim.Formel_XL2DP()"
End Function


Private Function Formel_DP2XL(ByVal Formel_DP As String)
  'Ersetzt in einer Formel des Datenpuffers alle Add-In-internen Spaltennamen
  'der Form "@@<Name>@@" durch "normale" Excel-Spaltenbezüge und die
  'relativen (im Datenpuffer geltenden) Zeilenangaben durch absolute Angaben.
  'Funktionswert: geänderte Formel zur Verwendung in der Excel-Tabelle.
  'Vorausetzung: Die Zieltabelle muß bereits aktiv sein!
  
  On Error GoTo Fehler
  Dim AddinSpaltenName
  Dim Zellbezug_DP
  Dim Zellbezuege_DP
  Dim Meldung
  Dim SpNameErsatz   As String
  Dim ZeileErsatz    As String
  Dim BezugErsatz    As String
  Dim BezugSpalte    As String
  Dim BezugZeile     As String
  Dim ZeilePrefix    As String
  Dim Formel_XL      As String
  Dim ZeAnf          As Long
  
  Formel_XL = Formel_DP
  
  If (Formel_DP <> "") Then
    'Festwerte des Datenbereiches ermitteln
    ZeAnf = ThisWorkbook.AktiveTabelle.ErsteDatenZeile
    'If (Err) Then GoTo Fehler
    
    'ThisWorkbook.RegExp.Pattern = "(@@([a-zA-Z0-9\.]+)@@(\$?)([0-9]+)){1}?"
    ThisWorkbook.RegExp.Pattern = "@@([a-zA-Z0-9\.]+)@@(\$?)([0-9]+)"
    ThisWorkbook.RegExp.IgnoreCase = True
    ThisWorkbook.RegExp.Global = True
    Set Zellbezuege_DP = ThisWorkbook.RegExp.Execute(Formel_DP)  'Matches-Auflistung.
    
    Meldung = "Entsprechungen gefunden:" & vbCrLf & vbCrLf
    For Each Zellbezug_DP In Zellbezuege_DP
      BezugSpalte = Zellbezug_DP.SubMatches(0)
      ZeilePrefix = Zellbezug_DP.SubMatches(1)
      BezugZeile = Zellbezug_DP.SubMatches(2)
      Meldung = Meldung & "bei Zeichen: " & Zellbezug_DP.FirstIndex & vbTab
      Meldung = Meldung & "Länge=" & Zellbezug_DP.Length & vbTab
      Meldung = Meldung & "String='" & Zellbezug_DP & "'" & vbCrLf
      
      If (ThisWorkbook.AktiveTabelle.SpaltenErsteZellen.Exists(BezugSpalte)) Then
        'Add-In-SpaltenName existiert in der aktiven Tabelle.
        'Zeile bezieht sich auf den Datenbereich (nicht Tabellenkopf).
        SpNameErsatz = UCase(MidStr(ThisWorkbook.AktiveTabelle.SpaltenErsteZellen(BezugSpalte).Address, "$", "$", False))
        ZeileErsatz = CStr(CLng(BezugZeile) + ZeAnf - 1)
        BezugErsatz = SpNameErsatz & ZeilePrefix & ZeileErsatz
        Formel_XL = Replace(Formel_XL, Zellbezug_DP, BezugErsatz)
      Else
        Formel_XL = ""
        Exit For
      End If
    Next
    'DebugEcho Meldung
    Set Zellbezuege_DP = Nothing
  End If
  Formel_DP2XL = Formel_XL
  Exit Function
  
Fehler:
  FehlerNachricht "CdatExpim.Formel_DP2XL()"
End Function


'für jEdit:  :folding=indent::collapseFolds=1:

