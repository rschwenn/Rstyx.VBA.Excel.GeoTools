Attribute VB_Name = "a_Lizenz_History"
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) f�r Geod�ten.
'**************************************************************************************************
'
' The MIT License (MIT)
' 
' Copyright (c) 2003-2021 Robert Schwenn
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'**************************************************************************************************

'==================================================================================================
'Modul Lizenz_History  (Dieses Modul enth�lt keinen Quelltext, sondern nur Kommentare)
'==================================================================================================
'
'N�tige Verweise:   - Microsoft Scripting Runtime
'                   - Windows Script Host Object Model
'                   - Microsoft VBScript Regular Expressions 5.5
'                   - Optional (f�r Protokollierung): Actions.NET-AddIn
'
'
'Versionshistorie:
'=================
'10.04.2022 v3.3.0  - CdatDatenpuffer: 
'                     - Zerlegen der iTrassen-Codierung angepasst an iGeo und Rstyx.Utilities
'                     - Mod_iTrassenCodeAusBemerkung() und Mod_UeberhoehungAusBemerkung()
'                       zusammengefasst zu Mod_InfoTextAuswerten()
'                   - GUI->Berechnung: "InfoText auswerten" statt zwei einzelne Aktionen (s.o.)
'11.07.2021 v3.2.1  - Vorlagencache wird jetzt vorzugsweise nicht mehr im TEMP-Ordner abgelegt,
'                     weil dieser oft geleert wird. Wenn m�glich, wird 
'                     "%LOCALAPPDATA%\GeoTools\GeoTools_xltcache.txt" verwendet, erst bei Mi�erfolg
'                     "%Temp%\GeoTools_xltcache.txt".
'02.03.2021 v3.2.0  - Ribbon: neuer Knopf f�r Handbuch
'                   - Neue Winkeleinheiten: Prozent, Promille, Parts per million
'                   - CimpTrassenkoo: - A0: Unterst�tzung f�r alle restlichen A0-Datenfelder
'                     - A0, A1: Text und Kommentar werden getrennt �bergeben (nicht mehr verkettet)
'                     - Standard-Zieltabelle ist jetzt "tabTK_iGeo".
'                   - Ereignissteuerung an unn�tigen Stellen entfernt (ScreenUpdating, EnableCalculation).
'                   - Bildschirmaktualisierung deaktiviert vor SetRequiredSeparators()
'                     bei Benutzeraktionen ohne Dialog
'                   - CdatDatenpuffer: 
'                     - Unterst�tzung f�r iTrassen-Codierung:  Mod_iTrassenCodeAusBemerkung()
'                     - Transfo' Trassenkoo' <=> Gleissystem arbeitet jetzt mit mehreren
'                       m�glichen S�tzen gegebener Werte in dieser Reihenfolge:
'                       1. Soll: Absolute �berh�hung (S.Tra.ua)
'                       2. Soll: Radius und relative �berh�hung (S.Tra.Radius, S.Tra.u)
'                       3. Ist : Absolute �berh�hung (Tra.ua)
'                       4. Ist : Radius und relative �berh�hung (Tra.Radius, Tra.u)
'12.02.2021 v3.1.2  - CdatKonfig: - Eigenschaft "KonfigDateiGelesen" umbenannt in "KonfigVerfuegbar"
'                                   (signalisiert den Erfolg).
'                                 - neue Eigenschaft "KonfigInitFertig"
'                                   (signalisiert den Abschluss der Init-Phase unabh�ngig vom Erfolg).
'                   - mdlBefehle: Neue Funktion IsGeoToolsInitFinished()
'25.01.2021 v3.1.1  - Das Lesen der Konfiguration erfolgt wieder ohne �nderung der Separatoren.
'                     CdatKonfig.TkBasisUeberhoehung() liefert jetzt einen Double-Wert unabh�ngig
'                     von den Einstellungen in System und Excel.
'05.01.2021 v3.1.0  - Separatoren (Dezimal, Tausend, Liste) werden vor Aktionen mit
'                     Berechnungen wie erfordert gesetzt und danach wieder hergestellt.
'                     => SetRequiredSeparators(), RestoreLastSeparators().
'01.06.2020 v3.0.4  - Workbook_BeforeClose() entfernt, da es beim Entladen der GeoTools
'                     im laufenden Excel st�rte (Excel friert ein).  
'30.05.2020 v3.0.3  - GeoTools-Ribbon-Tab hinter "Ansicht" positioniert.
'27.05.2020 v3.0.2  - Neu: CtabTabelle als Helfer f�r eine nicht aktive Tabelle
'                     - Grundlage: CtabAktiveTabelle
'                     - stark abgespeckt: bislang nur lesende Funktionen
'                     - Anlass: Lesen der Konfigurationsdatei im Hintergrund:
'                   - Die Konfigurationsdatei wird nicht mehr als aktive Datei ge�ffnet, sondern
'                     mit einer zweiten Excel-Instanz (siehe v3.0.0 frmStartExpim). Dadurch startet
'                     Excel ohne zweites Geister-Fenster, wenn es mit einem Datei-Argument aufgerufen wird.
'22.05.2020 v3.0.1  - Tastaturbedienung verbessert (KeyTips)
'                   - Gesch�tzte Ansicht des aktiven Fensters wird erkannt ...
'19.05.2020 v3.0.0  - Umstellung auf Office 2016 / VBA 7:
'                     - Umstellung "Declare"-Anweisungen auf VBA 7: 
'                       - CToolsSystem: GetOpenFileName(), GetSaveFileName(), Type OPENFILENAME
'                       - mdlRibbon:    CopyMemory(), getGeoToolsRibbon()
'                   - Unterst�tzung f�r Office 2016:
'                     - Vorlagensuche (frmStartExpim.frm) auch via Registry-Key:
'                       "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Options\PersonalTemplates"
'                   - CToolsSystem:  Neue Eigenschaft "TausenderTrenner" 
'                   - CdatMetaDaten: Unterst�tzung f�r Alt-iGeo-Projektdaten aufgegeben (GetAktPrjDat_iGEO(), ADODB)
'                   - frmStartExpim: �ffnen zu erkundender Vorlagen erfolgt wieder unsichtbar (zweite Excel-Instanz)
'                   - AddIn "Actions.NET.xll" wird jetzt beim Laden der GeoTools nachgeladen und registriert,
'                     und zwar aus dem GeoTools-Ressourcen-Ordner.
'                   - Fix: Statusleiste blockiert nicht mehr mit "FALSE".
'                     Ausl�ser f�r das Problem ist Ribbon.Invalidate(). Ursache unklar.
'                     Workaround: UpdateGeoToolsRibbon() ruft ClearStatusBarDelayed(1) auf und setzt damit
'                     die Statusleiste nach dem Ribbon-Update zur�ck.
'21.08.2019 v2.14.0 - Transfo_Tk2Gls() und Transfo_Gls2Tk():
'                     Angleichung an iGeo und VermEsn:  Nullpunkt wird nur in der H�he verschoben um u/2.
'25.02.2019 v2.13.0 - CimpTrassenkoo: - A0: Unterst�tzung f�r neue Werte:  KmText, RaLGS, AbLGS.
'                                     - A5: Unterst�tzung des Formates A5 wird aufgegeben!
'                                     - Dateifilter voreingestellt auf "*.A0"
'24.08.2018 v2.12.0 - CtabAktiveTabelle, Fu�zeile eintragen: �nderung der Fu�zeilen-Belegung
'                     (Links = Kontaktdaten, Mitte = leer, Rechts = Seitennummer).
'10.12.2017 v2.11.0 - Neu: BatchPDF
'                   - Neu: mdlToolsExcel.ProgressbarAllgemein()
'                   - Bugfix ErrEcho(): Das Err-Objekt wird jetzt erst nach dessen Auswertung geleert ;-)
'12.07.2016 v2.10.5 - Bugfix FehlerNachricht(): Das Err-Objekt wird jetzt erst nach dessen Auswertung geleert ;-)
'                   - Bugfix frmStartExpim.GetFormatliste_XlVorlagen(): Eine nicht lesbare Vorlage
'                     f�hrt nicht mehr zum Abbruch der Vorlagenerkundung. 
'06.07.2016 v2.10.4 - Beim Lesen der Konfiguration: Anzeige von Warnungen (DisplayAlerts) tempor�r deaktiviert.
'02.06.2016 v2.10.3 - CimpTrassenkoo: - A0, A1, A5: Text und Kommentar werden verkettet (bisher hatte Kommentar Vorrang).
'30.05.2016 v2.10.2 - Bugfix: Die letzte Arbeitsmappe konnte nicht per Fernsteuerung geschlossen werden. 
'                     => CtabAktiveTabelle.App_WorkbookDeactivate(): Verz�gerung nach DebugEcho().  
'21.02.2016 v2.10.1 - Suche nach jedit.jar ge�ndert (siehe Tools_1.vbi 4.2):
'                       Gesucht wird unter: - %JEDIT_HOME%
'                                           - %PROGRAMFILES%\jedit
'                                           - %PROGRAMFILES(X86)%\jedit
'                                           - %PROGRAMW6432%\jedit
'                                           - ThisWorkbook.Konfig.JEDIT_HOME
'                   - Editoraufrufe unterst�tzen jetzt Pfade mit Leerzeichen :-) 
'22.11.2015 v2.10.0 Kompatibilit�t mit Win7 und anderen Installations-Orten:
'                   - Anf�lligkeit gegen�ber "Gesch�tzter Ansicht" (deaktivierte Makros) verringert.
'                   - Vorlagencache wird jetzt als eine einzige Datei in TEMP-Ordner geschrieben.
'                   - Konfiguration und Hilfe werden jetzt nicht mehr im �bergeordneten Verzeichnis
'                     gesucht, sondern im GeoTools-Ressourcen-Ordner (Unterverzeichnis "GeoToolsRes").
'12.06.2015 v2.9.5  - Bugfix CimpTrassenkoo: - A0 darf jetzt mit absolut leerem Kommentar enden.
'                     - Schwerer Performanceeinbruch, wegen GetValue() seit V 2.8.0
'11.06.2015 v2.9.4  - Bugfix: Expim-Dialog deaktivierte die Formelberechnung in aktiver Tabelle.
'07.06.2015 v2.9.3  - Bugfix LeseTKiGeo(): - Nach einem A0-Block wurde kein weiterer Block erkannt.
'                                          - Als Name der Gradiente wurde "Trafo Gradiente" erkannt.
'25.09.2014 v2.9.2  - mdlTabellenFunktionen.bas entfernt (vorher in NivLinien.xltm integriert)
'01.07.2014 v2.9.1  - Bugfix in frmStartExpim: Ereignisbehandlung beim Durchsuchen der Vorlagen
'                     fast komplett ausgeschaltet wegen Interferenzen, die zu Fehlern f�hrten.
'01.06.2014 v2.9.0  - Protokollierung umgestellt:
'                     - VBA-LoggingConsole ausgebaut!
'                     - ***Echo-Methoden benutzen LoggingConsole.NET via Actions.NET-AddIn
'                     - Ist Actions.NET-AddIn nicht geladen, steht kein Protokoll zur Verf�gung.
'                     => Abh�ngigkeit von "Microsoft Windows Common Controls 6.0 (SP6)" entf�llt!
'24.04.2014 v2.8.0  - LeseTKiTrassePC(): - Anpassung f�r A1/A5 iGeo Version 1.2.2 (04/2014)
'                                        - Unterst�tzung f�r alle verf�gbaren Felder
'                                        - Unterst�tzung f�r iGeo-Format A0
'                                        - Umbenamnnt in LeseTKiGeo()
'16.02.2014 v2.7.0  - Umstellung der Oberfl�che von Controls/2003 auf Ribbon/2010
'                   - Umstellung der Konfigurationsdatei auf .xlsx
'                   - Umstellung globaler Objekte auf Eigenschaften von ThisWorkbook, damit
'                     Excel nicht neu gestartet weden muss, wenn AddIn wegen Fehler gestoppt wurde.
'                   - mdlUserInterface umbenannt in mdlBefehle, da es keinen UI-kode mehr enth�lt.
'23.01.2014 v2.6.2  - Bugfix: Nicht alle Kontextmen�-Eintr�ge wurden beim Beenden entfernt.
'                   - Bedingte Formatierung aus Kontextmen� entfernt (evtl. via Hooks.xlam)
'31.12.2013 v2.6.1  Anpassungen f�r Excel 2010:
'                   - CdatExpim          - Dateifilter f�r Vorlagensuche erweitert um .xltx und .xltm.
'                                        - Speichern der Arbeitsmappe im aktuellen Standardformat.
'05.12.2013 v2.6.0  - Lizenz ge�ndert. GeoTools unterliegen jetzt der "The MIT License"
'10.11.2013 v2.5.1  - CimpTrassenkoo     - Anpassung f�r iGeo Version 11/2013 (Gleisprofilbereich)
'15.09.2013 v2.5.0  - CimpTrassenkoo     - Anpassung an iTrassePC 2.0.2 (A1-Kennung ohne Kommas)
'                                        - Unterst�tzung f�r Formate A2, A3 und A4 entfernt
'                                        - Neu: Unterst�tzung f�r iGeo A1 und A5
'                   - LeseTkAdelt():     - Erkennung komplizierter Achsnamen verbessert
'03.03.2012 v2.4.0  - CimpTrassenkoo     - iTrasse-A5: Anpassung an iTrassePC 2.0 (DGM-Unterst�tzung)
'12.03.2011 v2.3.1  - Info-Dialog        - Mail-Adresse ge�ndert: devel@rstyx.de
'                   - LoggingConsole     - Update auf Version 1.3.1
'26.07.2010 v2.3.0  - CimpTrassenkoo     - Umstellung des Windows-Formats auf 8.40 / VE2010 (ge�nderte Koordinatenzeile)
'                                        - Integrit�tspr�fungen f�r zweizeilige Umformungen:
'                                          - PktNr beider Zeilen m�ssen identisch sein
'                                          - Falls die Umformung als "einzeilig" eingestuft wurde, werden
'                                            auftretende Koordinatenzeilen abgewiesen.
'29.06.2010 v2.2.0  - frmInsertLines     - Neu: Dialog zum Einf�gen von Leerzeilen im Intervall.
'                                          Eingebunden in Hauptmen� und Toolbox.
'19.11.2009 v2.1.5  - wbk_GeoTools       - Versuch, den Fokus nach Initialisierung im Excel-Fenster zu behalten
'07.11.2009 v2.1.4  - CToolsSystem       - Editor-Unterst�tzung entsprechend Tools_1.vbi aktualisiert
'                                        - rudiment�re OS-Info (Property "OS")
'                                        - Entscheidung: Windows 98 wird nicht unterst�tzt!!!
'                                        - FindFiles() zeigt Aktivit�t in Statuszeile
'                   - mdlToolsExcel      - WriteStatusBar() als Abstraktion zwecks Portierung
'                   - mdlToolsAllgemein: - Sortierung von Dictionaries sowie ein- und zweidimensionalen Arrays.
'                   - frmStartExpim      - Bugfix: XLT-Liste war nicht mehr sortiert (seit v2.1.3)
'30.10.2009 v2.1.3  - mdlToolsAllgemein  - SplitDelim() kann jetzt Felder trimmen und leere Felder �bergehen.
'                                        - Neu: ListeAuflistung() f�r Nicht-Dictionaries
'                                        - ListeDictionary() und ListeAuflistung() laufzeitoptimiert.
'                   - mdlToolsExcel      - Alle Funktionen f�r Dateisuche entfernt (au�er FindeXLVorlage),
'                                          wegen Unzuverl�ssigkeit des Application.FileSearch-Objektes.
'                   - CToolsSystem       - Funktionen f�r Dateisuche: FindFiles(), FindFile() auf VbScript-Basis.
'                                        - Unterst�tzung f�r Crimson Editor in Editorliste und StartEditor.
'                   - frmStartExpim      - Umstellung der XLT-Suche auf CToolsSystem.FindFiles
'                   - Allgemein          - Lesen von Umgebungsvariablen mit VB-Funktion Environ() statt via vbscript.
'27.09.2009 v2.1.2  - CtabAktiveTabelle  - Bugfix: Beim Einf�gen von Bereichsnamen wird der Name von
'                                          [Mappe]Tabelle immer in Hochkomma eingeschlossen, da dies nicht
'                                          nur f�r Leerzeichen n�tig ist, sondern z.B. auch f�r Minus und Komma...
'31.05.2009 v2.1.1  - CToolsSystem       - Bugfix in GetTmpDateiPfadName()
'19.05.2009 v2.1.0  - LoggingConsole     - Update auf Version 1.2.0 (ListView statt Textbox => schneller)
'                   - Allgemein          - Aufrufe von CreateObject() ersetzt durch direkte Instanzierung.
'                                          =>Entsprechende Bibliotheks-Verweise n�tig.
'25.04.2009 v2.0.2  - frmStartExpim:     - Bugfix: Fehler abgefangen, wenn keine einzige Vorlage existiert.
'                   - LoggingConsole     - Update auf Version 1.1.1 (Fehler, wenn ein modaler Dialog offen war)
'30.03.2009 v2.0.1  - CdatExpim:         - Bugfix: Fehler beim Beschreiben einer Tabellenzelle abgefangen.
'                   - frmStartExpim:     - Bugfix: Beim "Erkunden" der XLT's werden Ereignisse unterdr�ckt
'29.03.2009 v2.0.0  - Release
'29.03.2009 v1.9.92 - Allgemein:         - Bugfixes: Fehlermeldungen...
'28.03.2009 v1.9.91 - CdatExpim:         - Falls Dialog versteckt bleiben soll (i.d.R. bei Fernsteuerung), wird
'                                          er trotzdem angezeigt, wenn mehr als ein Zielformat verf�gbar ist.
'                   - Allgemein:         - Globale Variable oExpim (oExpimGlobal) wird nur noch in Verbindung mit
'                                          dem Aktionsmanager verwendet (n�tig wegen R�ckbezug aus Dialog...)
'                                        - oExpimGlobal wird nur instanziert, wenn es Nothing ist, sonst Fehlermeldung.
'26.03.2009 v1.9.90 - LoggingConsole:    - Aktualisiert auf Version 1.1.
'                   - CtabCSV:           - Bugfix: Wertkonvertierung der Parameter aus CSV-Kopf.
'17.03.2009 v1.9.89 - LoggingConsole:    - Wieder integriert wegen Verweis-Nebenwirkungen auf anderes Add-In.
'                   - Allgemein:         - Zugriff auf VBProject ist nicht mehr n�tig.
'15.03.2009 v1.9.88 - LoggingConsole:    - Entfernt wegen Ausgliederung in eigenes Add-In LoggingConsole.xla.
'15.03.2009 v1.9.87 - ??
'08.03.2009 v1.9.85 - CimpNivLinien:     - entfernt!
'                   - CimpTkElta:        - entfernt!
'                   - CimpPktPaare:      - entfernt!
'08.03.2009 v1.9.82 - TabellenFunktionen:  Berechnung der Niv-Linien-Statistik in mdlTabellenFunktionen,
'                                          um CimpNivLinien entfernen zu k�nnen.
'07.03.2009 v1.9.80 - CtabCSV:           - Optionen f�r den Datenimport (Daten bearbeiten, Ersatzspalten ..)
'                   - frmStartExpim:     - Steuerelemente f�r Ziel-ASCII-Datei entfernt, da ungenutzt.
'                                          Die entsprechenden Verweise im Kode nur auskommentiert.
'27.02.2009 v1.9.61 - Import/Export:     - Neuer Import-Datentyp: CSV-Datei mit speziellem Kopf (CtabCSV).
'                   - Allgemein:         - Verweis auf wshCommondialogs entfernt.
'                                        - diverse Bugfixes bzw. neue ben�tigte Funktionen.
'17.12.2008 v1.9.14 - CToolsSystem:      - Bugfix: Beim Aufbau der Editorliste kam es zu einem Laufzeitfehler,
'                                          wenn UltraEdit nicht gefunden wurde, aber jEdit doch.
'03.12.2008 v1.9.12 - Bugfix:            - oKonfig wird sofort nach Konsole initialisiert...
'16.11.2008 v1.9.9  - CimpTrassenkoo:    - Anzeige von Fehlern der Eingabedatei in jEdit-Fehlerliste.
'                   - frmStartExpim:     - Edit-Knopf f�r Eingabedatei.
'                   - mdlUserInterface:  - Neuer Men�punkt "Datei �ffnen (Name in Zelle)".
'                                        - Verhalten der Symbolleisten normalisiert: Sie werden ohne
'                                          Benutzereingriff nur noch sichtbar geschaltet, wenn sie
'                                          bisher nicht existierten, z.B. nach Erstinstallation
'                   - Refactoring:       - Anpassung einiger Module wegen folgender �nderungen:
'                     CToolsSystem:      - Neue Klasse f�r systemnahe Werkzeuge und Dateihandhabung.
'                                        - Editor und Datei-bezogene Fehlerliste verf�gbar.
'                     mdlToolsExcel:     - Neues Modul f�r Routinen, die auf Excel zur�ckgreifen.
'                     mdlToolsScripting: - Modul entfernt. Inhalt verteilt auf CToolsSystem,
'                                          mdlToolsExcel und mdlToolsAllgemein.
'                     mdlToolsAllgemein: - Teile ausgelagert nach CToolsSystem und mdlToolsExcel.
'                                          Andere Teile �bernommen aus mdlToolsScripting.
'09.11.2008 v1.9.8  - Einstellungen des manuell gestarteten Expim-Dialoges (mdlUserInterface.ExpimManager)
'                     werden gespeichert und beim n�chsten Aufruf wiederhergestellt.
'08.11.2008 v1.9.7  - Import/Export:     - Ersatzspalte: In Formeln wird ein Bezug auf die eigene
'                                          Spalte ge�ndert auf die Ersatzspalte.
'                                          Das Beschreiben der Ersatzspalte erfgolgt jetzt zellenorientiert,
'                                          falls die Ersatzspalte bereits im Datenpuffer existiert. In
'                                          diesem Fall werden keine Formeln �bernommen.
'                                        - Formeln allgemein: Unterst�tzung f�r absolute Zeilenbez�ge.
'03.11.2008 v1.9.6  - CimpTrassenkoo:    - Syncronisierung von Rahmenhandlung und Protokoll mit awk und vbi.
'                   - Cimp***:           - Neue Eigenschaft "Fehlerniveau".
'                   - CdatExpim:         - Meldungen nach Import basierend auf "Fehlerniveau" des Imports.
'01.11.2008 v1.9.5  - Komplette �berarbeitung der Projektdatenverarbeitung - jetzt "Metadaten".
'                     - betrifft einige Module
'                     - oPrjDatGlobal umbenannt nach oMetadaten
'                     - "Basis f�r �berh�hung" wird als Projektdatum verwaltet und bei Berechnungen
'                        ber�cksichtigt, wenn in aktiver Tabelle als Projektdatum vorhanden.
'                     - Feldnamen f�r ExtraDaten sollten jetzt mit "x." beginnen,
'                       damit sie aus der aktiven Tabelle ausgelesen werden k�nnen.
'                   - Belegte Statuszeile wird jetzt verz�gert freigegeben bzw. gel�scht.
'14.04.2008 v1.9.1  - CdatDatenpuffer:   - Transfo' Trasse <=> Gleis: Basis f�r �berh�hung konfigurierbar.
'                                        - Ist-�berh�hung aus Bemerkung: Schalter f�r "streng" (u=xxx) eingef�hrt
'                   - CdatKonfig:        - Verwaltung der o.g. Optionen
'03.04.2008 v1.9.0  - CimpTrassenkoo:    - iTrassePC-Import: - Formatdefinition aktualisiert (wie awk)
'                                                            - Kommentar am Zeilenende erlaubt
'                                                            - mehrere Umformungen in einer Datei
'                                        - Verm.Esn-Import:  - Unterst�tzung f�r 3/L (VE 6.22 + 8.30)
'27.03.2008 v1.8.1  - frmStartExpim:     - Beschleunigung des 1. Starts des Dialoges durch einen Vorlagen-Cache (xlt.cache)
'24.03.2008 v1.7.0  - Protokoll-Konsole eingef�hrt. ErrEcho() und DebugEcho() daf�r umgestellt.
'08.03.2007         - CtabAktiveTabelle: - L�schen des Datenbereiches l�scht nicht mehr die Zellen daneben.
'01.11.2005         - CtabAktiveTabelle: - Die Interpolationsformel kann jetzt sinnvoll seitlich gezogen werden.
'31.10.2005         - CimpTrassenkoo:    - Beim Lesen von iTrassePC-Ausgaben werden die bei der Berechnung
'                                          verwendeten Trassendaten zur Interpretation der einzelnen Werte
'                                          Werte herangezogen (i.d.R um "0.000" als "Leer" anzusehen).
'                                        - iTrassePC-Import: Gradientenh�he wird berechnet, wenn m�glich.
'30.08.2005         - CimpTrassenkoo:    Anpassung an ge�ndertes iTrassePC-Format A1.
'15.04.2005         - CimpTrassenkoo:    Bugfix: Leerzeilen am Ende einer Adelt-Umformung f�hrten zum Fehler.
'20.03.2005         - CimpTrassenkoo:    Unterst�tzung f�r Ausgaben von iTrassePC (Formate A1..A5).
'06.02.2005         - CimpTrassenkoo:    Erkennung Adelt-Pnr-Format nur beim 1. Punkt der Umformung.
'10.11.2004         - mdlToolsAllgemein: Bugfix in StarteDatei().
'23.09.2004         - mdlToolsAllgemein: Bugfix in SetArbeitsverzeichnis.
'21.09.2004         - CimpTkElta:        Neu erstellt zum Import von TK aus S_Trasse/Maintras.
'                   - CimpNivLinien:     �bernahme Soll-dh: Anpassung an NivBrein.awk v1.1.
'16.09.2004         - frmStartExpim:     Vorl�ufig Radio-Buttons entfernt f�r:
'                                        - Quelle-Typ "ASCII formatiert"
'                                        - Ziel-Typ "ASCII formatiert", "ASCII spezial"
'21.08.2004         - CimpTrassenkoo:    Bugfix: Gradientenname enthielt u.U. nachfolgende Leerzeichen.
'                   - CtabAktiveTabelle: keine Nachfrage mehr bei �nderung der Tabellenstruktur, daf�r:
'                   - frmSpaltenVerw:    Hilfebutton f�r Strukturelemente.
'18.08.2004         - CtabAktiveTabelle: Eigenschaft "Silent" zwecks Unterdr�cken von Fehlermeldungen.
'05.06.2004         - CdatExpim:         Start der Manipulationen nur, wenn Daten im Puffer sind.
'                   - frmStartExpim:     Analyse der XL-Vorlagen wird in der Statuszeile protokolliert.
'                   - Spalten benennen:  H�he des Kommentarfeldes erh�ht (im B�ro zu klein).
'31.05.2004         - Lizenzhinweise, Info �ber ...
'                   - Manipulationsroutinen in Toolbox "gtWerkzeuge" eingebunden.
'30.05.2004         - Add-In umbenannt nach "GeoTools".
'12/2003-05/2004    - Import/Export XL <==> XL �ber benannte Spalten incl. Einheitenangabe.
'                   - Trassenkoo' importieren.
'                   - Benutzerdialog f�r Import/Export.
'                   - Dialog zum Erzeugen/Verwalten der Tabellenstruktur und der Spaltennamen
'                     sowie Felder f�r Projektdaten.
'                   - Konfigurationsdatei (XLS) zwecks Spaltenkonfiguration und allgemeine Einstellungen.
'                   - Konfiguration und Anwendung von Ersatz-ZielSpalten (z.B. Station/Km).
'                   - Wertstati f�r automatische Berechnungen: Ist, Soll, Fehler, Verbesserung.
'                   - Datenmanipulationen:
'                     - Berechnung von Fehlern und Verbesserungen
'                     - Ist-�berh�hung aus Bemerkung ermitteln
'                     - Transfo' Trassenkoo' => Gleissystem (Zwangspunktreduktion)
'                     - Transfo' Gleissystem => Trassenkoo' (umgekehrte Zwangspunktreduktion)
'                   
'08.12.2003  v1.1   - Fu�zeile eintragen. => Wird in Importroutinen automatisch ausgef�hrt.
'                   - Doppelte Werte in einer Spalte markieren
'                   
'xx.04.2003  v1.0   - Import NivLinien und Punktpaare
'                   - Projektdaten eintragen
'                   - Interpolationsformel erstellen
'                   - Datenbereich l�schen und formatieren (Streifen, NK-Stellen),
'                     daf�r interaktives Festlegen des "Infotr�gers" und des "Fliesskommabereiches"
'
'==================================================================================================
