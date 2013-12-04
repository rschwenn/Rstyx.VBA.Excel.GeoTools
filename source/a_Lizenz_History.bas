Attribute VB_Name = "a_Lizenz_History"
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2003 - 2013  Robert Schwenn
'
' Dieses Programm ist freie Software. Sie können es unter den Bedingungen der GNU General Public
' License, wie von der Free Software Foundation veröffentlicht, weitergeben und/oder modifizieren,
' entweder gemäß Version 2 der Lizenz oder (nach Ihrer Option) jeder späteren Version.
' Die Veröffentlichung dieses Programms erfolgt in der Hoffnung, daß es Ihnen von Nutzen sein wird,
' aber OHNE IRGENDEINE GARANTIE, sogar ohne die implizite Garantie der MARKTREIFE oder der
' VERWENDBARKEIT FÜR EINEN BESTIMMTEN ZWECK. Details finden Sie in der GNU General Public License.
'
' Sie sollten eine Kopie der GNU General Public License zusammen mit diesem Programm erhalten haben.
' Falls nicht, schreiben Sie an die Free Software Foundation, Inc., 59 Temple Place, Suite 330,
' Boston, MA 02111-1307, USA.
' Siehe auch: http://www.gnu.org/copyleft/gpl.html
'**************************************************************************************************

'==================================================================================================
'Modul Lizenz_History  (Dieses Modul enthält keinen Quelltext, sondern nur Kommentare)
'==================================================================================================
'
'Nötige Verweise:   - Microsoft ActiveX Data Objects 2.1 Library
'                   - Microsoft Scripting Runtime
'                   - Windows Script Host Object Model
'                   - Microsoft Windows Common Controls 6.0 (SP6)
'                   - Microsoft VBScript Regular Expressions 5.5
'
'
'Versionshistorie:
'=================
'10.11.2013 v2.5.1  - CimpTrassenkoo     - Anpassung für iGeo Version 11/2013 (Gleisprofilbereich)
'15.09.2013 v2.5.0  - CimpTrassenkoo     - Anpassung an iTrassePC 2.0.2 (A1-Kennung ohne Kommas)
'                                        - Unterstützung für Formate A2, A3 und A4 entfernt
'                                        - Neu: Unterstützung für iGeo A1 und A5
'                   - LeseTkAdelt():     - Erkennung komplizierter Achsnamen verbessert
'03.03.2012 v2.4.0  - CimpTrassenkoo     - iTrasse-A5: Anpassung an iTrassePC 2.0 (DGM-Unterstützung)
'12.03.2011 v2.3.1  - Info-Dialog        - Mail-Adresse geändert: devel@rstyx.de
'                   - LoggingConsole     - Update auf Version 1.3.1
'26.07.2010 v2.3.0  - CimpTrassenkoo     - Umstellung des Windows-Formats auf 8.40 / VE2010 (geänderte Koordinatenzeile)
'                                        - Integritätsprüfungen für zweizeilige Umformungen:
'                                          - PktNr beider Zeilen müssen identisch sein
'                                          - Falls die Umformung als "einzeilig" eingestuft wurde, werden
'                                            auftretende Koordinatenzeilen abgewiesen.
'29.06.2010 v2.2.0  - frmInsertLines     - Neu: Dialog zum Einfügen von Leerzeilen im Intervall.
'                                          Eingebunden in Hauptmenü und Toolbox.
'19.11.2009 v2.1.5  - wbk_GeoTools       - Versuch, den Fokus nach Initialisierung im Excel-Fenster zu behalten
'07.11.2009 v2.1.4  - CToolsSystem       - Editor-Unterstützung entsprechend Tools_1.vbi aktualisiert
'                                        - rudimentäre OS-Info (Property "OS")
'                                        - Entscheidung: Windows 98 wird nicht unterstützt!!!
'                                        - FindFiles() zeigt Aktivität in Statuszeile
'                   - mdlToolsExcel      - WriteStatusBar() als Abstraktion zwecks Portierung
'                   - mdlToolsAllgemein: - Sortierung von Dictionaries sowie ein- und zweidimensionalen Arrays.
'                   - frmStartExpim      - Bugfix: XLT-Liste war nicht mehr sortiert (seit v2.1.3)
'30.10.2009 v2.1.3  - mdlToolsAllgemein  - SplitDelim() kann jetzt Felder trimmen und leere Felder übergehen.
'                                        - Neu: ListeAuflistung() für Nicht-Dictionaries
'                                        - ListeDictionary() und ListeAuflistung() laufzeitoptimiert.
'                   - mdlToolsExcel      - Alle Funktionen für Dateisuche entfernt (außer FindeXLVorlage),
'                                          wegen Unzuverlässigkeit des Application.FileSearch-Objektes.
'                   - CToolsSystem       - Funktionen für Dateisuche: FindFiles(), FindFile() auf VbScript-Basis.
'                                        - Unterstützung für Crimson Editor in Editorliste und StartEditor.
'                   - frmStartExpim      - Umstellung der XLT-Suche auf CToolsSystem.FindFiles
'                   - Allgemein          - Lesen von Umgebungsvariablen mit VB-Funktion Environ() statt via vbscript.
'27.09.2009 v2.1.2  - CtabAktiveTabelle  - Bugfix: Beim Einfügen von Bereichsnamen wird der Name von
'                                          [Mappe]Tabelle immer in Hochkomma eingeschlossen, da dies nicht
'                                          nur für Leerzeichen nötig ist, sondern z.B. auch für Minus und Komma...
'31.05.2009 v2.1.1  - CToolsSystem       - Bugfix in GetTmpDateiPfadName()
'19.05.2009 v2.1.0  - LoggingConsole     - Update auf Version 1.2.0 (ListView statt Textbox => schneller)
'                   - Allgemein          - Aufrufe von CreateObject() ersetzt durch direkte Instanzierung.
'                                          =>Entsprechende Bibliotheks-Verweise nötig.
'25.04.2009 v2.0.2  - frmStartExpim:     - Bugfix: Fehler abgefangen, wenn keine einzige Vorlage existiert.
'                   - LoggingConsole     - Update auf Version 1.1.1 (Fehler, wenn ein modaler Dialog offen war)
'30.03.2009 v2.0.1  - CdatExpim:         - Bugfix: Fehler beim Beschreiben einer Tabellenzelle abgefangen.
'                   - frmStartExpim:     - Bugfix: Beim "Erkunden" der XLT's werden Ereignisse unterdrückt
'29.03.2009 v2.0.0  - Release
'29.03.2009 v1.9.92 - Allgemein:         - Bugfixes: Fehlermeldungen...
'28.03.2009 v1.9.91 - CdatExpim:         - Falls Dialog versteckt bleiben soll (i.d.R. bei Fernsteuerung), wird
'                                          er trotzdem angezeigt, wenn mehr als ein Zielformat verfügbar ist.
'                   - Allgemein:         - Globale Variable oExpim (oExpimGlobal) wird nur noch in Verbindung mit
'                                          dem Aktionsmanager verwendet (nötig wegen Rückbezug aus Dialog...)
'                                        - oExpimGlobal wird nur instanziert, wenn es Nothing ist, sonst Fehlermeldung.
'26.03.2009 v1.9.90 - LoggingConsole:    - Aktualisiert auf Version 1.1.
'                   - CtabCSV:           - Bugfix: Wertkonvertierung der Parameter aus CSV-Kopf.
'17.03.2009 v1.9.89 - LoggingConsole:    - Wieder integriert wegen Verweis-Nebenwirkungen auf anderes Add-In.
'                   - Allgemein:         - Zugriff auf VBProject ist nicht mehr nötig.
'15.03.2009 v1.9.88 - LoggingConsole:    - Entfernt wegen Ausgliederung in eigenes Add-In LoggingConsole.xla.
'15.03.2009 v1.9.87 - ??
'08.03.2009 v1.9.85 - CimpNivLinien:     - entfernt!
'                   - CimpTkElta:        - entfernt!
'                   - CimpPktPaare:      - entfernt!
'08.03.2009 v1.9.82 - TabellenFunktionen:  Berechnung der Niv-Linien-Statistik in mdlTabellenFunktionen,
'                                          um CimpNivLinien entfernen zu können.
'07.03.2009 v1.9.80 - CtabCSV:           - Optionen für den Datenimport (Daten bearbeiten, Ersatzspalten ..)
'                   - frmStartExpim:     - Steuerelemente für Ziel-ASCII-Datei entfernt, da ungenutzt.
'                                          Die entsprechenden Verweise im Kode nur auskommentiert.
'27.02.2009 v1.9.61 - Import/Export:     - Neuer Import-Datentyp: CSV-Datei mit speziellem Kopf (CtabCSV).
'                   - Allgemein:         - Verweis auf wshCommondialogs entfernt.
'                                        - diverse Bugfixes bzw. neue benötigte Funktionen.
'17.12.2008 v1.9.14 - CToolsSystem:      - Bugfix: Beim Aufbau der Editorliste kam es zu einem Laufzeitfehler,
'                                          wenn UltraEdit nicht gefunden wurde, aber jEdit doch.
'03.12.2008 v1.9.12 - Bugfix:            - oKonfig wird sofort nach Konsole initialisiert...
'16.11.2008 v1.9.9  - CimpTrassenkoo:    - Anzeige von Fehlern der Eingabedatei in jEdit-Fehlerliste.
'                   - frmStartExpim:     - Edit-Knopf für Eingabedatei.
'                   - mdlUserInterface:  - Neuer Menüpunkt "Datei öffnen (Name in Zelle)".
'                                        - Verhalten der Symbolleisten normalisiert: Sie werden ohne
'                                          Benutzereingriff nur noch sichtbar geschaltet, wenn sie
'                                          bisher nicht existierten, z.B. nach Erstinstallation
'                   - Refactoring:       - Anpassung einiger Module wegen folgender Änderungen:
'                     CToolsSystem:      - Neue Klasse für systemnahe Werkzeuge und Dateihandhabung.
'                                        - Editor und Datei-bezogene Fehlerliste verfügbar.
'                     mdlToolsExcel:     - Neues Modul für Routinen, die auf Excel zurückgreifen.
'                     mdlToolsScripting: - Modul entfernt. Inhalt verteilt auf CToolsSystem,
'                                          mdlToolsExcel und mdlToolsAllgemein.
'                     mdlToolsAllgemein: - Teile ausgelagert nach CToolsSystem und mdlToolsExcel.
'                                          Andere Teile übernommen aus mdlToolsScripting.
'09.11.2008 v1.9.8  - Einstellungen des manuell gestarteten Expim-Dialoges (mdlUserInterface.ExpimManager)
'                     werden gespeichert und beim nächsten Aufruf wiederhergestellt.
'08.11.2008 v1.9.7  - Import/Export:     - Ersatzspalte: In Formeln wird ein Bezug auf die eigene
'                                          Spalte geändert auf die Ersatzspalte.
'                                          Das Beschreiben der Ersatzspalte erfgolgt jetzt zellenorientiert,
'                                          falls die Ersatzspalte bereits im Datenpuffer existiert. In
'                                          diesem Fall werden keine Formeln übernommen.
'                                        - Formeln allgemein: Unterstützung für absolute Zeilenbezüge.
'03.11.2008 v1.9.6  - CimpTrassenkoo:    - Syncronisierung von Rahmenhandlung und Protokoll mit awk und vbi.
'                   - Cimp***:           - Neue Eigenschaft "Fehlerniveau".
'                   - CdatExpim:         - Meldungen nach Import basierend auf "Fehlerniveau" des Imports.
'01.11.2008 v1.9.5  - Komplette Überarbeitung der Projektdatenverarbeitung - jetzt "Metadaten".
'                     - betrifft einige Module
'                     - oPrjDatGlobal umbenannt nach oMetadaten
'                     - "Basis für Überhöhung" wird als Projektdatum verwaltet und bei Berechnungen
'                        berücksichtigt, wenn in aktiver Tabelle als Projektdatum vorhanden.
'                     - Feldnamen für ExtraDaten sollten jetzt mit "x." beginnen,
'                       damit sie aus der aktiven Tabelle ausgelesen werden können.
'                   - Belegte Statuszeile wird jetzt verzögert freigegeben bzw. gelöscht.
'14.04.2008 v1.9.1  - CdatDatenpuffer:   - Transfo' Trasse <=> Gleis: Basis für Überhöhung konfigurierbar.
'                                        - Ist-Überhöhung aus Bemerkung: Schalter für "streng" (u=xxx) eingeführt
'                   - CdatKonfig:        - Verwaltung der o.g. Optionen
'03.04.2008 v1.9.0  - CimpTrassenkoo:    - iTrassePC-Import: - Formatdefinition aktualisiert (wie awk)
'                                                            - Kommentar am Zeilenende erlaubt
'                                                            - mehrere Umformungen in einer Datei
'                                        - Verm.Esn-Import:  - Unterstützung für 3/L (VE 6.22 + 8.30)
'27.03.2008 v1.8.1  - frmStartExpim:     - Beschleunigung des 1. Starts des Dialoges durch einen Vorlagen-Cache (xlt.cache)
'24.03.2008 v1.7.0  - Protokoll-Konsole eingeführt. ErrEcho() und DebugEcho() dafür umgestellt.
'08.03.2007         - CtabAktiveTabelle: - Löschen des Datenbereiches löscht nicht mehr die Zellen daneben.
'01.11.2005         - CtabAktiveTabelle: - Die Interpolationsformel kann jetzt sinnvoll seitlich gezogen werden.
'31.10.2005         - CimpTrassenkoo:    - Beim Lesen von iTrassePC-Ausgaben werden die bei der Berechnung
'                                          verwendeten Trassendaten zur Interpretation der einzelnen Werte
'                                          Werte herangezogen (i.d.R um "0.000" als "Leer" anzusehen).
'                                        - iTrassePC-Import: Gradientenhöhe wird berechnet, wenn möglich.
'30.08.2005         - CimpTrassenkoo:    Anpassung an geändertes iTrassePC-Format A1.
'15.04.2005         - CimpTrassenkoo:    Bugfix: Leerzeilen am Ende einer Adelt-Umformung führten zum Fehler.
'20.03.2005         - CimpTrassenkoo:    Unterstützung für Ausgaben von iTrassePC (Formate A1..A5).
'06.02.2005         - CimpTrassenkoo:    Erkennung Adelt-Pnr-Format nur beim 1. Punkt der Umformung.
'10.11.2004         - mdlToolsAllgemein: Bugfix in StarteDatei().
'23.09.2004         - mdlToolsAllgemein: Bugfix in SetArbeitsverzeichnis.
'21.09.2004         - CimpTkElta:        Neu erstellt zum Import von TK aus S_Trasse/Maintras.
'                   - CimpNivLinien:     Übernahme Soll-dh: Anpassung an NivBrein.awk v1.1.
'16.09.2004         - frmStartExpim:     Vorläufig Radio-Buttons entfernt für:
'                                        - Quelle-Typ "ASCII formatiert"
'                                        - Ziel-Typ "ASCII formatiert", "ASCII spezial"
'21.08.2004         - CimpTrassenkoo:    Bugfix: Gradientenname enthielt u.U. nachfolgende Leerzeichen.
'                   - CtabAktiveTabelle: keine Nachfrage mehr bei Änderung der Tabellenstruktur, dafür:
'                   - frmSpaltenVerw:    Hilfebutton für Strukturelemente.
'18.08.2004         - CtabAktiveTabelle: Eigenschaft "Silent" zwecks Unterdrücken von Fehlermeldungen.
'05.06.2004         - CdatExpim:         Start der Manipulationen nur, wenn Daten im Puffer sind.
'                   - frmStartExpim:     Analyse der XL-Vorlagen wird in der Statuszeile protokolliert.
'                   - Spalten benennen:  Höhe des Kommentarfeldes erhöht (im Büro zu klein).
'31.05.2004         - Lizenzhinweise, Info über ...
'                   - Manipulationsroutinen in Toolbox "gtWerkzeuge" eingebunden.
'30.05.2004         - Add-In umbenannt nach "GeoTools".
'12/2003-05/2004    - Import/Export XL <==> XL über benannte Spalten incl. Einheitenangabe.
'                   - Trassenkoo' importieren.
'                   - Benutzerdialog für Import/Export.
'                   - Dialog zum Erzeugen/Verwalten der Tabellenstruktur und der Spaltennamen
'                     sowie Felder für Projektdaten.
'                   - Konfigurationsdatei (XLS) zwecks Spaltenkonfiguration und allgemeine Einstellungen.
'                   - Konfiguration und Anwendung von Ersatz-ZielSpalten (z.B. Station/Km).
'                   - Wertstati für automatische Berechnungen: Ist, Soll, Fehler, Verbesserung.
'                   - Datenmanipulationen:
'                     - Berechnung von Fehlern und Verbesserungen
'                     - Ist-Überhöhung aus Bemerkung ermitteln
'                     - Transfo' Trassenkoo' => Gleissystem (Zwangspunktreduktion)
'                     - Transfo' Gleissystem => Trassenkoo' (umgekehrte Zwangspunktreduktion)
'                   
'08.12.2003  v1.1   - Fußzeile eintragen. => Wird in Importroutinen automatisch ausgeführt.
'                   - Doppelte Werte in einer Spalte markieren
'                   
'xx.04.2003  v1.0   - Import NivLinien und Punktpaare
'                   - Projektdaten eintragen
'                   - Interpolationsformel erstellen
'                   - Datenbereich löschen und formatieren (Streifen, NK-Stellen),
'                     dafür interaktives Festlegen des "Infoträgers" und des "Fliesskommabereiches"
'
'==================================================================================================
'
'
'Wunschliste:      - Import/Export formatierter ASCII-Dateien.
'============      - mehrzeiliger Infoträger
'
'==================================================================================================






'**************************************************************************************************
'         GNU GENERAL PUBLIC LICENSE
'            Version 2, June 1991
'
'  Copyright (C) 1989, 1991 Free Software Foundation, Inc.
'                        59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'  Everyone is permitted to copy and distribute verbatim copies
'  of this license document, but changing it is not allowed.
'
'           Preamble
'
'   The licenses for most software are designed to take away your
' freedom to share and change it.  By contrast, the GNU General Public
' License is intended to guarantee your freedom to share and change free
' software--to make sure the software is free for all its users.  This
' General Public License applies to most of the Free Software
' Foundation's software and to any other program whose authors commit to
' using it.  (Some other Free Software Foundation software is covered by
' the GNU Library General Public License instead.)  You can apply it to
' your programs, too.
'
'   When we speak of free software, we are referring to freedom, not
' price.  Our General Public Licenses are designed to make sure that you
' have the freedom to distribute copies of free software (and charge for
' this service if you wish), that you receive source code or can get it
' if you want it, that you can change the software or use pieces of it
' in new free programs; and that you know you can do these things.
'
'   To protect your rights, we need to make restrictions that forbid
' anyone to deny you these rights or to ask you to surrender the rights.
' These restrictions translate to certain responsibilities for you if you
' distribute copies of the software, or if you modify it.
'
'   For example, if you distribute copies of such a program, whether
' gratis or for a fee, you must give the recipients all the rights that
' you have.  You must make sure that they, too, receive or can get the
' source code.  And you must show them these terms so they know their
' rights.
'
'   We protect your rights with two steps: (1) copyright the software, and
' (2) offer you this license which gives you legal permission to copy,
' distribute and/or modify the software.
'
'   Also, for each author's protection and ours, we want to make certain
' that everyone understands that there is no warranty for this free
' software.  If the software is modified by someone else and passed on, we
' want its recipients to know that what they have is not the original, so
' that any problems introduced by others will not reflect on the original
' authors' reputations.
'
'   Finally, any free program is threatened constantly by software
' patents.  We wish to avoid the danger that redistributors of a free
' program will individually obtain patent licenses, in effect making the
' program proprietary.  To prevent this, we have made it clear that any
' patent must be licensed for everyone's free use or not licensed at all.
'
'   The precise terms and conditions for copying, distribution and
' modification follow.
'
'         GNU GENERAL PUBLIC LICENSE
'    TERMS AND CONDITIONS FOR COPYING, DISTRIBUTION AND MODIFICATION
'
'   0. This License applies to any program or other work which contains
' a notice placed by the copyright holder saying it may be distributed
' under the terms of this General Public License.  The "Program", below,
' refers to any such program or work, and a "work based on the Program"
' means either the Program or any derivative work under copyright law:
' that is to say, a work containing the Program or a portion of it,
' either verbatim or with modifications and/or translated into another
' language.  (Hereinafter, translation is included without limitation in
' the term "modification".)  Each licensee is addressed as "you".
'
' Activities other than copying, distribution and modification are not
' covered by this License; they are outside its scope.  The act of
' running the Program is not restricted, and the output from the Program
' is covered only if its contents constitute a work based on the
' Program (independent of having been made by running the Program).
' Whether that is true depends on what the Program does.
'
'   1. You may copy and distribute verbatim copies of the Program's
' source code as you receive it, in any medium, provided that you
' conspicuously and appropriately publish on each copy an appropriate
' copyright notice and disclaimer of warranty; keep intact all the
' notices that refer to this License and to the absence of any warranty;
' and give any other recipients of the Program a copy of this License
' along with the Program.
'
' You may charge a fee for the physical act of transferring a copy, and
' you may at your option offer warranty protection in exchange for a fee.
'
'   2. You may modify your copy or copies of the Program or any portion
' of it, thus forming a work based on the Program, and copy and
' distribute such modifications or work under the terms of Section 1
' above, provided that you also meet all of these conditions:
'
'     a) You must cause the modified files to carry prominent notices
'     stating that you changed the files and the date of any change.
'
'     b) You must cause any work that you distribute or publish, that in
'     whole or in part contains or is derived from the Program or any
'     part thereof, to be licensed as a whole at no charge to all third
'     parties under the terms of this License.
'
'     c) If the modified program normally reads commands interactively
'     when run, you must cause it, when started running for such
'     interactive use in the most ordinary way, to print or display an
'     announcement including an appropriate copyright notice and a
'     notice that there is no warranty (or else, saying that you provide
'     a warranty) and that users may redistribute the program under
'     these conditions, and telling the user how to view a copy of this
'     License.  (Exception: if the Program itself is interactive but
'     does not normally print such an announcement, your work based on
'     the Program is not required to print an announcement.)
'
' These requirements apply to the modified work as a whole.  If
' identifiable sections of that work are not derived from the Program,
' and can be reasonably considered independent and separate works in
' themselves, then this License, and its terms, do not apply to those
' sections when you distribute them as separate works.  But when you
' distribute the same sections as part of a whole which is a work based
' on the Program, the distribution of the whole must be on the terms of
' this License, whose permissions for other licensees extend to the
' entire whole, and thus to each and every part regardless of who wrote it.
'
' Thus, it is not the intent of this section to claim rights or contest
' your rights to work written entirely by you; rather, the intent is to
' exercise the right to control the distribution of derivative or
' collective works based on the Program.
'
' In addition, mere aggregation of another work not based on the Program
' with the Program (or with a work based on the Program) on a volume of
' a storage or distribution medium does not bring the other work under
' the scope of this License.
'
'   3. You may copy and distribute the Program (or a work based on it,
' under Section 2) in object code or executable form under the terms of
' Sections 1 and 2 above provided that you also do one of the following:
'
'     a) Accompany it with the complete corresponding machine-readable
'     source code, which must be distributed under the terms of Sections
'     1 and 2 above on a medium customarily used for software interchange; or,
'
'     b) Accompany it with a written offer, valid for at least three
'     years, to give any third party, for a charge no more than your
'     cost of physically performing source distribution, a complete
'     machine-readable copy of the corresponding source code, to be
'     distributed under the terms of Sections 1 and 2 above on a medium
'     customarily used for software interchange; or,
'
'     c) Accompany it with the information you received as to the offer
'     to distribute corresponding source code.  (This alternative is
'     allowed only for noncommercial distribution and only if you
'     received the program in object code or executable form with such
'     an offer, in accord with Subsection b above.)
'
' The source code for a work means the preferred form of the work for
' making modifications to it.  For an executable work, complete source
' code means all the source code for all modules it contains, plus any
' associated interface definition files, plus the scripts used to
' control compilation and installation of the executable.  However, as a
' special exception, the source code distributed need not include
' anything that is normally distributed (in either source or binary
' form) with the major components (compiler, kernel, and so on) of the
' operating system on which the executable runs, unless that component
' itself accompanies the executable.
'
' If distribution of executable or object code is made by offering
' access to copy from a designated place, then offering equivalent
' access to copy the source code from the same place counts as
' distribution of the source code, even though third parties are not
' compelled to copy the source along with the object code.
'
'   4. You may not copy, modify, sublicense, or distribute the Program
' except as expressly provided under this License.  Any attempt
' otherwise to copy, modify, sublicense or distribute the Program is
' void, and will automatically terminate your rights under this License.
' However, parties who have received copies, or rights, from you under
' this License will not have their licenses terminated so long as such
' parties remain in full compliance.
'
'   5. You are not required to accept this License, since you have not
' signed it.  However, nothing else grants you permission to modify or
' distribute the Program or its derivative works.  These actions are
' prohibited by law if you do not accept this License.  Therefore, by
' modifying or distributing the Program (or any work based on the
' Program), you indicate your acceptance of this License to do so, and
' all its terms and conditions for copying, distributing or modifying
' the Program or works based on it.
'
'   6. Each time you redistribute the Program (or any work based on the
' Program), the recipient automatically receives a license from the
' original licensor to copy, distribute or modify the Program subject to
' these terms and conditions.  You may not impose any further
' restrictions on the recipients' exercise of the rights granted herein.
' You are not responsible for enforcing compliance by third parties to
' this License.
'
'   7. If, as a consequence of a court judgment or allegation of patent
' infringement or for any other reason (not limited to patent issues),
' conditions are imposed on you (whether by court order, agreement or
' otherwise) that contradict the conditions of this License, they do not
' excuse you from the conditions of this License.  If you cannot
' distribute so as to satisfy simultaneously your obligations under this
' License and any other pertinent obligations, then as a consequence you
' may not distribute the Program at all.  For example, if a patent
' license would not permit royalty-free redistribution of the Program by
' all those who receive copies directly or indirectly through you, then
' the only way you could satisfy both it and this License would be to
' refrain entirely from distribution of the Program.
'
' If any portion of this section is held invalid or unenforceable under
' any particular circumstance, the balance of the section is intended to
' apply and the section as a whole is intended to apply in other
' circumstances.
'
' It is not the purpose of this section to induce you to infringe any
' patents or other property right claims or to contest validity of any
' such claims; this section has the sole purpose of protecting the
' integrity of the free software distribution system, which is
' implemented by public license practices.  Many people have made
' generous contributions to the wide range of software distributed
' through that system in reliance on consistent application of that
' system; it is up to the author/donor to decide if he or she is willing
' to distribute software through any other system and a licensee cannot
' impose that choice.
'
' This section is intended to make thoroughly clear what is believed to
' be a consequence of the rest of this License.
'
'   8. If the distribution and/or use of the Program is restricted in
' certain countries either by patents or by copyrighted interfaces, the
' original copyright holder who places the Program under this License
' may add an explicit geographical distribution limitation excluding
' those countries, so that distribution is permitted only in or among
' countries not thus excluded.  In such case, this License incorporates
' the limitation as if written in the body of this License.
'
'   9. The Free Software Foundation may publish revised and/or new versions
' of the General Public License from time to time.  Such new versions will
' be similar in spirit to the present version, but may differ in detail to
' address new problems or concerns.
'
' Each version is given a distinguishing version number.  If the Program
' specifies a version number of this License which applies to it and "any
' later version", you have the option of following the terms and conditions
' either of that version or of any later version published by the Free
' Software Foundation.  If the Program does not specify a version number of
' this License, you may choose any version ever published by the Free Software
' Foundation.
'
'   10. If you wish to incorporate parts of the Program into other free
' programs whose distribution conditions are different, write to the author
' to ask for permission.  For software which is copyrighted by the Free
' Software Foundation, write to the Free Software Foundation; we sometimes
' make exceptions for this.  Our decision will be guided by the two goals
' of preserving the free status of all derivatives of our free software and
' of promoting the sharing and reuse of software generally.
'
'           NO WARRANTY
'
'   11. BECAUSE THE PROGRAM IS LICENSED FREE OF CHARGE, THERE IS NO WARRANTY
' FOR THE PROGRAM, TO THE EXTENT PERMITTED BY APPLICABLE LAW.  EXCEPT WHEN
' OTHERWISE STATED IN WRITING THE COPYRIGHT HOLDERS AND/OR OTHER PARTIES
' PROVIDE THE PROGRAM "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED
' OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE.  THE ENTIRE RISK AS
' TO THE QUALITY AND PERFORMANCE OF THE PROGRAM IS WITH YOU.  SHOULD THE
' PROGRAM PROVE DEFECTIVE, YOU ASSUME THE COST OF ALL NECESSARY SERVICING,
' REPAIR OR CORRECTION.
'
'   12. IN NO EVENT UNLESS REQUIRED BY APPLICABLE LAW OR AGREED TO IN WRITING
' WILL ANY COPYRIGHT HOLDER, OR ANY OTHER PARTY WHO MAY MODIFY AND/OR
' REDISTRIBUTE THE PROGRAM AS PERMITTED ABOVE, BE LIABLE TO YOU FOR DAMAGES,
' INCLUDING ANY GENERAL, SPECIAL, INCIDENTAL OR CONSEQUENTIAL DAMAGES ARISING
' OUT OF THE USE OR INABILITY TO USE THE PROGRAM (INCLUDING BUT NOT LIMITED
' TO LOSS OF DATA OR DATA BEING RENDERED INACCURATE OR LOSSES SUSTAINED BY
' YOU OR THIRD PARTIES OR A FAILURE OF THE PROGRAM TO OPERATE WITH ANY OTHER
' PROGRAMS), EVEN IF SUCH HOLDER OR OTHER PARTY HAS BEEN ADVISED OF THE
' POSSIBILITY OF SUCH DAMAGES.
'
'          END OF TERMS AND CONDITIONS
'
'       How to Apply These Terms to Your New Programs
'
'   If you develop a new program, and you want it to be of the greatest
' possible use to the public, the best way to achieve this is to make it
' free software which everyone can redistribute and change under these terms.
'
'   To do so, attach the following notices to the program.  It is safest
' to attach them to the start of each source file to most effectively
' convey the exclusion of warranty; and each file should have at least
' the "copyright" line and a pointer to where the full notice is found.
'
'     <one line to give the program's name and a brief idea of what it does.>
'     Copyright (C) <year>  <name of author>
'
'     This program is free software; you can redistribute it and/or modify
'     it under the terms of the GNU General Public License as published by
'     the Free Software Foundation; either version 2 of the License, or
'     (at your option) any later version.
'
'     This program is distributed in the hope that it will be useful,
'     but WITHOUT ANY WARRANTY; without even the implied warranty of
'     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'     GNU General Public License for more details.
'
'     You should have received a copy of the GNU General Public License
'     along with this program; if not, write to the Free Software
'     Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'
' Also add information on how to contact you by electronic and paper mail.
'
' If the program is interactive, make it output a short notice like this
' when it starts in an interactive mode:
'
'     Gnomovision version 69, Copyright (C) year name of author
'     Gnomovision comes with ABSOLUTELY NO WARRANTY; for details type `show w'.
'     This is free software, and you are welcome to redistribute it
'     under certain conditions; type `show c' for details.
'
' The hypothetical commands `show w' and `show c' should show the appropriate
' parts of the General Public License.  Of course, the commands you use may
' be called something other than `show w' and `show c'; they could even be
' mouse-clicks or menu items--whatever suits your program.
'
' You should also get your employer (if you work as a programmer) or your
' school, if any, to sign a "copyright disclaimer" for the program, if
' necessary.  Here is a sample; alter the names:
'
'   Yoyodyne, Inc., hereby disclaims all copyright interest in the program
'   `Gnomovision' (which makes passes at compilers) written by James Hacker.
'
'   <signature of Ty Coon>, 1 April 1989
'   Ty Coon, President of Vice
'
' This General Public License does not permit incorporating your program into
' proprietary programs.  If your program is a subroutine library, you may
' consider it more useful to permit linking proprietary applications with the
' library.  If this is what you want to do, use the GNU Library General
' Public License instead of this License.
'**************************************************************************************************
