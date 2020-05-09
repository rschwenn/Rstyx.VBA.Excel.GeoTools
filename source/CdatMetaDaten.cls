VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CdatMetaDaten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2004-2020  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'==================================================================================================
' Klasse CdatMetaDaten
'==================================================================================================
' Erfassung, temporäre Speicherung und Handhabung von Metadaten.
'
'Beschreibung:
'-------------
'A. Metadaten setzen sich inhaltlich aus 3 Arten zusammen:
 ' 1. Ortsdaten: Adresse, Telefon usw.
 ' 2. Projektdaten:
 '    2.1 Global (PC):   ID bzw. ProjektNr und 3 Bezeichnungen des aktiven Projektes
 '    2.2 Lokal (Datei): bezogen auf die Tabellendaten: Koo'system, Höhensystem, Basis für Überhöhung
 '    2.3 PC:            aktuelles Datum, am PC angemeldeter Benutzer
 ' 3. ExtraDaten:        beliebige andere Informationen (Name/Wert-Paare), deren Bedeutung im Einzelnen
 '                       dem Programm unbekannt ist.
 '
'B. Zeitpunkt der Erfassung und Lebensdauer der Daten:
 ' 1. Während der Initialisierung:  für die Dauer der Excel-Sitzung verfügbare Daten
 '                                  - alle Ortsdaten
 '                                  - Bearbeiter
 ' 2. In Echtzeit (via Me.Update):  - alle Projektdaten außer Bearbeiter
 '                                  - ExtraDaten
'C. Herkunft der Daten:
 ' 1. Ortsdaten:    Ini-Datei (Pfad ist festgelegt in der CFG-Datei, Variable "Ort_Dateiname")
 ' 2. Projektdaten: 
 '    2.1 und 2.2:  - Aktives Projekt des PC (abhängig vom konfigurierten "Anwender"), soweit verfügbar.
 '                  - Einzelne oder alle dieser Daten können von außen (via Me.Update) als
 '                    lokale Daten vorgegeben werden. Dabei handelt es sich um Daten,
 '                    die zu einer vorhandenen Tabelle (XLS oder ASCII) gehören.
 '                    Die globalen Daten (2.1) werden nur im Paket übernommen, d.h. falls mindestens eines
 '                    dieser Felder lokal zur Verfügung steht, werden die global bestimmten Daten gelöscht.
 '    2.3 PC:       Umgebung des PC
 ' 3. ExtraDaten:   werden von außen (via Me.Update) zur Verfügung gestellt
 '
'D. Namen von Metadatenfeldern in Tabellen:
 ' 1. Ortsdaten:    Präfix = "Ort.", Es werden nur hier im Kode vereinbarte Namen verarbeitet.
 ' 2. Projektdaten: Präfix = "Prj.", Es werden nur hier im Kode vereinbarte Namen verarbeitet.
 ' 3. ExtraDaten:   Präfix = "x.",   Alle Namen mit diesem Präfix werden beim Lesen aus der Tabelle
 '                                   als ExtraDaten erkannt. Prinzipiell gibt es aber keinen Zwang
 '                                   zur Verwendung dieses Präfixes.
 '
'E. Aktualisierung:
 'Es gibt keinen Automatismus zur Aktualisierung der Metadaten. Es ist also immer davon auszugehen,
 'dass sie (von den Ortsdaten abgesehen) nicht aktuell sind. Deshalb MUSS unmittelbar vor ihrer 
 'Verwendung immer eine Aktualisierung erfolgen (Me.Update)!!!
 'Vorgehen der Update-Methode:
 ' - Ermitteln des global aktiven Projektes und Übernahme der entsprechenden Projektdaten.
 ' - Setzen des aktuellen Datums.
 ' - Falls Projektdaten als Argument übergeben wurden, werden diese übernommen (siehe C).
 ' - Löschen aller Extradaten und Übernahme der als Argument übergebenen Extradaten.
'==================================================================================================

Option Explicit

'Konstanten
Private Const strKlasse_Name                   As String = "CdatMetaDaten"

'Namen der Projektdatenfelder (ZellNamen in Tabelle = Keys des Dictionary)
  Private Const FN_AuftragNr                   As String = "Prj.AuftragNr"
  Private Const FN_Projektbez1                 As String = "Prj.Projektbez1"
  Private Const FN_Projektbez2                 As String = "Prj.Projektbez2"
  Private Const FN_Projektbez3                 As String = "Prj.Projektbez3"
  Private Const FN_HoehenSystem                As String = "Prj.HoehenSystem"
  Private Const FN_KooSystem                   As String = "Prj.KooSystem"
  Private Const FN_TraBasisUeb                 As String = "Prj.TraBasisUeb"
  Private Const FN_BearbUsername               As String = "Prj.BearbUsername"
  Private Const FN_BearbVollerName             As String = "Prj.BearbVollerName"
  Private Const FN_BearbNachname               As String = "Prj.BearbNachname"
  Private Const FN_Datum                       As String = "Prj.Datum"

'Namen der Ortsdatenfelder (müssen klein geschrieben sein wegen Vergleich mit ini-Keys!)
  Private Const Prefix_FN_Ort                  As String = "Ort."         '"Differenz" zwischen Name in ini-Datei und Feldname
  Private Const FN_Ort_Name                    As String = "Ort.name"
  Private Const FN_Ort_Ort                     As String = "Ort.ort"
  Private Const FN_Ort_Plz                     As String = "Ort.plz"
  Private Const FN_Ort_Strasse                 As String = "Ort.strasse"
  Private Const FN_Ort_Tel                     As String = "Ort.tel"
  Private Const FN_Ort_Fax                     As String = "Ort.fax"
  Private Const FN_Ort_Mail                    As String = "Ort.mail"
  Private Const FN_Ort_Fusszeile_Excel_1       As String = "Ort.fusszeile_excel_1"


'Titel der Projektdatenfelder (für Dialog "Tabellenstruktur verwalten")
  Private Const Titel_AuftragNr                As String = "Auftragsnummer"
  Private Const Titel_Projektbez1              As String = "Projektbezeichnung 1"
  Private Const Titel_Projektbez2              As String = "Projektbezeichnung 2"
  Private Const Titel_Projektbez3              As String = "Projektbezeichnung 3"
  Private Const Titel_HoehenSystem             As String = "Höhensystem"
  Private Const Titel_KooSystem                As String = "Koordinatensystem"
  Private Const Titel_TraBasisUeb              As String = "Basis für Überhöhung [m]"
  Private Const Titel_BearbUsername            As String = "Bearbeiter PC-Benutzername"
  Private Const Titel_BearbVollerName          As String = "Bearbeiter Vor- und Nachname"
  Private Const Titel_BearbNachname            As String = "Bearbeiter Nachname"
  Private Const Titel_Datum                    As String = "Datum"
                                             
'Titel der Ortsdatenfelder
  Private Const Titel_Ort_Name                 As String = "Ort/Adresse: Name (der Niederlassung)"
  Private Const Titel_Ort_Ort                  As String = "Ort/Adresse: Ortsname"
  Private Const Titel_Ort_Plz                  As String = "Ort/Adresse: PLZ"
  Private Const Titel_Ort_Strasse              As String = "Ort/Adresse: Strassenname"
  Private Const Titel_Ort_Tel                  As String = "Ort/Adresse: Telefon"
  Private Const Titel_Ort_Fax                  As String = "Ort/Adresse: Fax"
  Private Const Titel_Ort_Mail                 As String = "Ort/Adresse: Mail-Adresse"
  Private Const Titel_Ort_Fusszeile_Excel_1    As String = "Ort/Adresse: Fusszeile 1 für Excel"


'Dictionaries, die Metadaten enthalten
Private oDictPrjOrt           As Scripting.Dictionary
Private oDictPrjOrtTitel      As Scripting.Dictionary
Private oDictExtra            As Scripting.Dictionary
'



Private Sub Class_Initialize()
  '1. Initialisierung der Dictionaries für Projekt- und Ortsdaten.
  '2. Erhebung der für die Dauer der Excel-Sitzung konstanten Daten (Ortsdaten, Bearbeiter)
  
  On Error GoTo Fehler
  DebugEcho "CdatMetaDaten.Class_Initialize(): Initialisierung beginnt."
  
  'Projekt- und Ortsdaten initialisieren
    Set oDictPrjOrt = New Scripting.Dictionary
    oDictPrjOrt.Add FN_Projektbez1, ""
    oDictPrjOrt.Add FN_Projektbez2, ""
    oDictPrjOrt.Add FN_Projektbez3, ""
    oDictPrjOrt.Add FN_AuftragNr, ""
    oDictPrjOrt.Add FN_BearbUsername, ThisWorkbook.SysTools.Username
    oDictPrjOrt.Add FN_BearbVollerName, oDictPrjOrt(FN_BearbUsername)   ' mangels entsprechender Funktion
    oDictPrjOrt.Add FN_BearbNachname, oDictPrjOrt(FN_BearbUsername)     ' mangels entsprechender Funktion
    oDictPrjOrt.Add FN_HoehenSystem, ""
    oDictPrjOrt.Add FN_KooSystem, ""
    oDictPrjOrt.Add FN_TraBasisUeb, ""
    oDictPrjOrt.Add FN_Datum, Date
    oDictPrjOrt.Add FN_Ort_Name, ""
    oDictPrjOrt.Add FN_Ort_Ort, ""
    oDictPrjOrt.Add FN_Ort_Plz, ""
    oDictPrjOrt.Add FN_Ort_Strasse, ""
    oDictPrjOrt.Add FN_Ort_Tel, ""
    oDictPrjOrt.Add FN_Ort_Fax, ""
    oDictPrjOrt.Add FN_Ort_Mail, ""
    oDictPrjOrt.Add FN_Ort_Fusszeile_Excel_1, ""
  
  'Titel der Projekt- und Ortsdaten (für Dialoge) initialisieren
    Set oDictPrjOrtTitel = New Scripting.Dictionary
    oDictPrjOrtTitel.Add FN_Projektbez1, Titel_Projektbez1
    oDictPrjOrtTitel.Add FN_Projektbez2, Titel_Projektbez2
    oDictPrjOrtTitel.Add FN_Projektbez3, Titel_Projektbez3
    oDictPrjOrtTitel.Add FN_AuftragNr, Titel_AuftragNr
    oDictPrjOrtTitel.Add FN_BearbUsername, Titel_BearbUsername
    oDictPrjOrtTitel.Add FN_BearbVollerName, Titel_BearbVollerName
    oDictPrjOrtTitel.Add FN_BearbNachname, Titel_BearbNachname
    oDictPrjOrtTitel.Add FN_HoehenSystem, Titel_HoehenSystem
    oDictPrjOrtTitel.Add FN_KooSystem, Titel_KooSystem
    oDictPrjOrtTitel.Add FN_TraBasisUeb, Titel_TraBasisUeb
    oDictPrjOrtTitel.Add FN_Datum, Titel_Datum
    oDictPrjOrtTitel.Add FN_Ort_Name, Titel_Ort_Name
    oDictPrjOrtTitel.Add FN_Ort_Ort, Titel_Ort_Ort
    oDictPrjOrtTitel.Add FN_Ort_Plz, Titel_Ort_Plz
    oDictPrjOrtTitel.Add FN_Ort_Strasse, Titel_Ort_Strasse
    oDictPrjOrtTitel.Add FN_Ort_Tel, Titel_Ort_Tel
    oDictPrjOrtTitel.Add FN_Ort_Fax, Titel_Ort_Fax
    oDictPrjOrtTitel.Add FN_Ort_Mail, Titel_Ort_Mail
    oDictPrjOrtTitel.Add FN_Ort_Fusszeile_Excel_1, Titel_Ort_Fusszeile_Excel_1
  
  'Extradaten initialisieren
  Set oDictExtra = New Scripting.Dictionary
  
  Call GetOrtsDaten         'Niederlassungsdaten ermitteln (Anschrift..)
  DebugEcho "CdatMetaDaten.Class_Initialize(): Initialisierung beendet."
  Exit Sub
  
Fehler:
  ErrMessage = "Fehler beim Initialisieren der MetaDaten"
  FehlerNachricht "CdatMetaDaten.Class_Initialize()"
End Sub


Private Sub Class_Terminate()
  Set oDictPrjOrt = Nothing
  Set oDictPrjOrtTitel = Nothing
  Set oDictExtra = Nothing
End Sub



'==== Eigenschaften ===========================================================

Property Get Klasse_Name() As String
  'liefert den Namen der Objektklasse
  Klasse_Name = strKlasse_Name
End Property


Property Get AuftragNr() As String
  'liefert die AuftragNr
  AuftragNr = oDictPrjOrt(FN_AuftragNr)
End Property

Property Let AuftragNr(inpAuftragNr As String)
  'setzt die AuftragNr
  oDictPrjOrt(FN_AuftragNr) = inpAuftragNr
End Property


Property Get Projektbez1() As String
  'liefert Projektbezeichnung Zeile 1
  Projektbez1 = oDictPrjOrt(FN_Projektbez1)
End Property

Property Let Projektbez1(inpProjektbez1 As String)
  'setzt Projektbezeichnung Zeile 1
  oDictPrjOrt(FN_Projektbez1) = inpProjektbez1
End Property


Property Get Projektbez2() As String
  'liefert Projektbezeichnung Zeile 2
  Projektbez2 = oDictPrjOrt(FN_Projektbez2)
End Property

Property Let Projektbez2(inpProjektbez2 As String)
  'setzt Projektbezeichnung Zeile 2
  oDictPrjOrt(FN_Projektbez2) = inpProjektbez2
End Property


Property Get Projektbez3() As String
  'liefert Projektbezeichnung Zeile 3
  Projektbez3 = oDictPrjOrt(FN_Projektbez3)
End Property

Property Let Projektbez3(inpProjektbez3 As String)
  'setzt Projektbezeichnung Zeile 3
  oDictPrjOrt(FN_Projektbez3) = inpProjektbez3
End Property


Property Get HoehenSystem() As String
  'liefert das HoehenSystem
  HoehenSystem = oDictPrjOrt(FN_HoehenSystem)
End Property

Property Let HoehenSystem(inpHoehenSystem As String)
  'setzt das HoehenSystem
  oDictPrjOrt(FN_HoehenSystem) = inpHoehenSystem
End Property


Property Get KooSystem() As String
  'liefert das KooSystem
  KooSystem = oDictPrjOrt(FN_KooSystem)
End Property

Property Let KooSystem(inpKooSystem As String)
  'setzt das KooSystem
  oDictPrjOrt(FN_KooSystem) = inpKooSystem
End Property


Property Get TraBasisUeb() As String
  'liefert die Basis für die Gleisüberhöhung
  TraBasisUeb = oDictPrjOrt(FN_TraBasisUeb)
End Property

Property Let TraBasisUeb(inpTraBasisUeb As String)
  'setzt die Basis für die Gleisüberhöhung
  'Ein leerer String ist zwar ungültig, wird aber kommentarlos durch den Standardwert ersetzt,
  '
  dim dblWert
  if (isNumeric(inpTraBasisUeb)) then
    dblWert = cDbl(inpTraBasisUeb)
  else
    dblWert = ThisWorkbook.Konfig.TkBasisUeberhoehung
    if (not (isEmpty(inpTraBasisUeb) or (inpTraBasisUeb = ""))) then
      ErrEcho vbNewLine & "CdatMetaDaten.TraBasisUeb (Let): Basis für Überhöhung kann nicht auf den Wert '" & inpTraBasisUeb & "' gesetzt werden!"
      msgbox "Basis für Überhöhung kann nicht auf den Wert '" & inpTraBasisUeb & "' gesetzt werden!" & vbNewLine & "Standard-Wert '" & dblWert & "' wird verwendet.", vbOKOnly, "FEHLER"
    else
      WarnEcho vbNewLine & "Basis für Überhöhung ist leer!"
    end if
    Echo "Basis für Überhöhung wird auf Standard-Wert '" & dblWert & "' gesetzt!"
  end if
  oDictPrjOrt(FN_TraBasisUeb) = dblWert
End Property


Property Get BearbUsername() As String
  'liefert den Usernamen des angemeldeten Benutzers
  BearbUsername = oDictPrjOrt(FN_BearbUsername)
End Property

Property Let BearbUsername(inpBearbUsername As String)
  'setzt den Usernamen des angemeldeten Benutzers
  oDictPrjOrt(FN_BearbUsername) = inpBearbUsername
End Property


Property Get BearbVollerName() As String
  'liefert den Vollen Namen des angemeldeten Benutzers
  BearbVollerName = oDictPrjOrt(FN_BearbVollerName)
End Property

Property Let BearbVollerName(inpBearbVollerName As String)
  'setzt den Vollen Namen des angemeldeten Benutzers
  oDictPrjOrt(FN_BearbVollerName) = inpBearbVollerName
End Property


Property Get BearbNachname() As String
  'liefert den Familiennamen des angemeldeten Benutzers
  BearbNachname = oDictPrjOrt(FN_BearbNachname)
End Property

Property Let BearbNachname(inpBearbNachname As String)
  'setzt den Familiennamen des angemeldeten Benutzers
  oDictPrjOrt(FN_BearbNachname) = inpBearbNachname
End Property


Property Get Datum() As Date
  'liefert das Datum
  Datum = oDictPrjOrt(FN_Datum)
End Property

Property Let Datum(inpDatum As Date)
  'setzt das Datum
  oDictPrjOrt(FN_Datum) = inpDatum
End Property



Property Get Ort_Name() As String
  'Liefert den Namen der Niederlassung.
  Ort_Name = oDictPrjOrt(FN_Ort_Name)
End Property

Property Get Ort_Ort() As String
  'Liefert den Ort der Niederlassung.
  Ort_Ort = oDictPrjOrt(FN_Ort_Ort)
End Property

Property Get Ort_PLZ() As String
  'Liefert die PLZ der Niederlassung.
  Ort_PLZ = oDictPrjOrt(FN_Ort_Plz)
End Property

Property Get Ort_Strasse() As String
  'Liefert Straße und Hausnuummer der Niederlassung.
  Ort_Strasse = oDictPrjOrt(FN_Ort_Strasse)
End Property

Property Get Ort_Tel() As String
  'Liefert die Telefonnummer der Niederlassung.
  Ort_Tel = oDictPrjOrt(FN_Ort_Tel)
End Property

Property Get Ort_Fax() As String
  'Liefert die Faxnummer der Niederlassung.
  Ort_Fax = oDictPrjOrt(FN_Ort_Fax)
End Property

Property Get Ort_Mail() As String
  'Liefert die Mail-Adresse der Niederlassung.
  Ort_Mail = oDictPrjOrt(FN_Ort_Mail)
End Property

Property Get Ort_Fusszeile_Excel_1() As String
  'Liefert die für die Niederlassung festgelegte Fusszeile 1 für Excel.
  Ort_Fusszeile_Excel_1 = oDictPrjOrt(FN_Ort_Fusszeile_Excel_1)
End Property


Property Get TitelDerProjektDaten() As Scripting.Dictionary
  'Liefert das Dictionary mit den Beschreibungen aller Projektdaten.
  Set TitelDerProjektDaten = oDictPrjOrtTitel
End Property

Property Get AlleProjektDaten() As Scripting.Dictionary
  'Liefert das Dictionary mit allen Projektdaten.
  Set AlleProjektDaten = oDictPrjOrt
End Property

Property Get AlleExtraDaten() As Scripting.Dictionary
  'Liefert das Dictionary mit den Extradaten.
  Set AlleExtraDaten = oDictExtra
End Property


'==== Methoden ================================================================

Public Sub Update(oPrjLocal As Scripting.Dictionary, oExtraLocal As Scripting.Dictionary)
  '------------------------------------------------------------------------------------------------
  'Alle Werte, die nicht über die gesamte Session konstant bleiben,
  'werden aktualisiert (alle Projektdaten außer Bearbeiter).
  'Eingabe: oPrjLocal   ... Nothing oder Dictionary mit lokalen Projektdaten aus einer Tabelle.
  '                         => Diese haben höhere Priorität als die des aktiven Projektes!
  '         oExtralocal ... Nothing oder Dictionary mit lokalen Extradaten aus einer Tabelle.
  '                         => Diese werden unverändert übernommen.
  '------------------------------------------------------------------------------------------------
  On Error GoTo Fehler
  DebugEcho "CdatMetaDaten.Update(): Aktualisierung der MetaDaten."
  
  Dim key            As Variant
  Dim FeldNichtLeer  As Boolean
  
  'Extradaten ungesehen übernehmen, falls vorhanden
  If (Not (oExtraLocal Is Nothing)) Then
    DebugEcho "CdatMetaDaten.Update():  - Bereitgestellte Extradaten werden übernommen."
    Set oDictExtra = oExtraLocal
  else
    DebugEcho "CdatMetaDaten.Update():  - Keine Extradaten bereitgestellt."
    oDictExtra.RemoveAll
  end if
  
  'Aktives Projekt abfragen => Alle Felder außer Bearbeiter und Datum werden belegt (z.T. leer)
  Call GetAktPrjDat
  
  'Datum: immer aktuell setzen (Session könnte über Mitternacht gehen :-)
  Me.Datum = Date
  
  'Lokale Projektdaten übernehmen.
  If (oPrjLocal Is Nothing) Then
    DebugEcho "CdatMetaDaten.Update():  - Keine lokalen Projektdaten bereitgestellt."
  else
    DebugEcho "CdatMetaDaten.Update():  - Folgende lokale Projektdaten wurden bereitgestellt:"
    
    'Zu ignorierende Daten löschen wegen Abfrage von oPrjLocal.Count weiter unten
      if (oPrjLocal.Exists(FN_Datum))           then oPrjLocal.Remove(FN_Datum)
      if (oPrjLocal.Exists(FN_BearbUsername))   then oPrjLocal.Remove(FN_BearbUsername)
      if (oPrjLocal.Exists(FN_BearbNachname))   then oPrjLocal.Remove(FN_BearbNachname)
      if (oPrjLocal.Exists(FN_BearbVollerName)) then oPrjLocal.Remove(FN_BearbVollerName)
    
    '
    'HoehenSystem immer übernehmen, falls verfügbar.
    if (oPrjLocal.Exists(FN_HoehenSystem)) then
      DebugEcho "CdatMetaDaten.Update():    - HoehenSystem = '" & oPrjLocal(FN_HoehenSystem) & "'"
      Me.HoehenSystem = oPrjLocal(FN_HoehenSystem)
      oPrjLocal.Remove(FN_HoehenSystem)
    end if
    
    'KooSystem immer übernehmen, falls verfügbar.
    if (oPrjLocal.Exists(FN_KooSystem)) then
      DebugEcho "CdatMetaDaten.Update():    - KooSystem =    '" & oPrjLocal(FN_KooSystem) & "'"
      Me.KooSystem = oPrjLocal(FN_KooSystem)
      oPrjLocal.Remove(FN_KooSystem)
    end if
    
    'TraBasisUeb immer übernehmen, falls verfügbar.
    if (oPrjLocal.Exists(FN_TraBasisUeb)) then
      DebugEcho "CdatMetaDaten.Update():    - TraBasisUeb =  '" & oPrjLocal(FN_TraBasisUeb) & "'"
      Me.TraBasisUeb = oPrjLocal(FN_TraBasisUeb)
      oPrjLocal.Remove(FN_TraBasisUeb)
    end if
    
    'Restliche Daten (AuftragNr, ProjektbezX) nur im Paket übernehmen und auch nur dann,
    'wenn mindestens eines dieser Felder nicht leer ist.
    if (oPrjLocal.Count > 0) then
      'Feststellen, ob mindestens ein Feld nicht leer ist
      for each key in oPrjLocal
        if (oPrjLocal(key) <> "") then FeldNichtLeer = true
      next
      
      if (FeldNichtLeer) Then
        'Restliche Felder übernehmen (nicht vorhandene Felder von oPrjLocal werden mit "" initialisiert).
        DebugEcho "CdatMetaDaten.Update():    - AuftragNr =    '" & oPrjLocal(FN_AuftragNr) & "'"
        DebugEcho "CdatMetaDaten.Update():    - Projektbez1 =  '" & oPrjLocal(FN_Projektbez1) & "'"
        DebugEcho "CdatMetaDaten.Update():    - Projektbez2 =  '" & oPrjLocal(FN_Projektbez2) & "'"
        DebugEcho "CdatMetaDaten.Update():    - Projektbez3 =  '" & oPrjLocal(FN_Projektbez3) & "'"
        Me.AuftragNr   = oPrjLocal(FN_AuftragNr)
        Me.Projektbez1 = oPrjLocal(FN_Projektbez1)
        Me.Projektbez2 = oPrjLocal(FN_Projektbez2)
        Me.Projektbez3 = oPrjLocal(FN_Projektbez3)
      end if
    end if
  end if
  
  Call ZeigeMetaDaten()
  Exit Sub

Fehler:
  ErrMessage = "Fehler beim Aktualisieren der MetaDaten"
  FehlerNachricht "CdatMetaDaten.Update()"
End Sub



Public Sub ZeigeMetaDaten()
  'Schreibt für Kontrollzwecke alle aktuellen Metadaten ins Debug-Protokoll.
  DebugEcho "Aktuelle Projektdaten:"
  DebugEcho ListeDictionary(oDictPrjOrt)
  
  DebugEcho "Aktuelle Extradaten:"
  DebugEcho ListeDictionary(oDictExtra)
End Sub




'==== interne Routinen ========================================================

Private Sub GetOrtsDaten()
  'Ermittelt die beschreibenden Daten des Installationsortes (der Niederlassung).
  
  Dim Kanal                As Integer
  Dim NR                   As Long       'Nr. der zuletzt eingelesenen Zeile der Eingabedatei
  Dim NF                   As Long       'Anzahl Felder der aktuellen Zeile
  Dim Zeile                As String     'gesamte aktuelle Zeile, die zuletzt eingelesen wurde
  Dim Feld()               As String     'einzelne Felder der aktuellen Zeile (1. Index = 1)
  Dim Key                  As String
  
  On Error GoTo Fehler
  DebugEcho "CdatMetaDaten.GetOrtsDaten(): Ortsdaten lesen aus: " & ThisWorkbook.Konfig.Ort_Dateiname
  
  Kanal = FreeFile()
  Open ThisWorkbook.Konfig.Ort_Dateiname For Input Lock Write As #Kanal
  If (Err) Then GoTo Fehler
  
  Do While Not EOF(Kanal)
    Line Input #Kanal, Zeile
    NR = NR + 1
    NF = SplitDelim(Zeile, Feld, "=")
    If (NF = 2) Then
      Key = Prefix_FN_Ort & LCase(Trim(Feld(1)))
      If (Not oDictPrjOrt.Exists(Key)) Then oDictPrjOrt.Add Key, ""
      oDictPrjOrt(Key) = Trim(Feld(2))
    End If
  Loop
  
  Close #Kanal
  DebugEcho "CdatMetaDaten.GetOrtsDaten(): Ortsdaten erfolgreich gelesen."
  Exit Sub
  
Fehler:
  Close #Kanal
  DebugEcho "CdatMetaDaten.GetOrtsDaten(): Fehler beim Lesen der Ortsdaten!"
  'FehlerNachricht "CdatMetaDaten.GetOrtsDaten()"
End Sub


Private Sub GetAktPrjDat()
  '=> Diese Routine arbeitet in Abhängigkeit des konfigurierten Anwenders.
  'Ermittelt den Name und die beschreibenden Daten des global aktiven Projektes.
  '=> Bearbeiter und Datum, falls verfügbar, werden ignoriert!
  'Kann kein aktives Projekt ermittelt werden oder tritt ein Fehler auf,
  'so werden alle betroffenen Felder mit "" belegt.
  
  On Error GoTo Fehler
  DebugEcho "CdatMetaDaten.GetAktPrjDat(): Daten des aktiven Projektes ermitteln für Anwender: '" & ThisWorkbook.Konfig.Anwender & "'"
  
  'Alte Daten löschen
  DebugEcho "CdatMetaDaten.GetAktPrjDat(): Alte Projektdaten löschen."
  Me.AuftragNr    = ""
  Me.Projektbez1  = ""
  Me.Projektbez2  = ""
  Me.Projektbez3  = ""
  Me.HoehenSystem = ""
  Me.KooSystem    = ""
  Me.TraBasisUeb  = ThisWorkbook.Konfig.TkBasisUeberhoehung
  
  'Select Case ThisWorkbook.Konfig.Anwender
  '  Case Anw_intermetric
  '       'Call GetAktPrjDat_iGEO    'Projektdaten des aktiven iGEO-Projektes ermitteln
  '  Case Else
  '       'keine Projektdaten
  'End Select
  Exit Sub

Fehler:
  FehlerNachricht "CdatMetaDaten.GetAktPrjDat()"
End Sub


'für jEdit:  :folding=indent::collapseFolds=1:
