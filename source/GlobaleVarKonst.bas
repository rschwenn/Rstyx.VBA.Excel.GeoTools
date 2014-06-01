Attribute VB_Name = "GlobaleVarKonst"
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2003 - 2014  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'====================================================================================
' Modul GlobaleVarKonst
'====================================================================================
'
' Deklaration von Variablen und Konstanten, die für das gesamte Add-In gelten.


Option Explicit

'Programminfo'
Public Const ProgName     As String = "GeoTools"
Public Const VersionNr    As String = "2.9.0"
Public Const VersionDate  As String = "Juni 2014"
Public Const Copyright    As String = "© 2003 - 2014  Robert Schwenn"
Public Const eMail        As String = "devel@rstyx.de"


'Standard-Einstellungen
Public Const Std_TkBasisUeberhoehung         As Double  = 1.500
Public Const Std_UebAusInfoStreng            As Boolean = true
Public Const Std_VorhWerteUeberschreiben     As Boolean = false
Public Const Std_DatenModifizieren           As Boolean = true
Public Const Std_ErsatzZielspaltenVerwenden  As Boolean = true
Public Const Std_FormelnErhalten             As Boolean = true
Public Const Std_ExpimSchlussMeldung         As Boolean = true


'Namen der unterstützten Anwender
Public Const Anw_intermetric                 As String = "intermetric"

'Namen benannter Zellbereiche.
Public Const strInfoTraeger                  As String = "Daten.InfoTraeger"
Public Const strFliesskomma                  As String = "Daten.Fliesskomma"
Public Const strFormel                       As String = "Daten.Formel"
Public Const strErsteZelle                   As String = "Daten.ErsteZelle"

'Syntax des Bereichsnamens für eine Spalte: PrefixSpaltenname<SpaltenName>[TrennerEinheit<Einheit>]
Public Const PrefixSpaltenname               As String = "Spalte."
Public Const TrennerEinheit                  As String = ".."

'Syntax CSV-Datei
Public Const CsvKopfBeginn                   As String = "@GEOTOOLS_BEGIN"
Public Const CsvKopfEnde                     As String = "@GEOTOOLS_END"
Public Const CsvAllOtherColumns              As String = "$AllOtherColumns$"
Public Const CsvTrenner_Std                  As String = ","
'Public Const CsvDezimalTrenner_Std           As String = "."
Public Const CsvTextQualifier_Std            As String = """"
Public Const CsvTrimFields_Std               As String = false  '(entspricht dem Verhalten von Excel beim CSV-Öffnen)

'Wertstatus-Bezeichnungen.
Public Const StatusBez_Ist                   As String = "Ist"
Public Const StatusBez_Soll                  As String = "Soll"
Public Const StatusBez_Fehler                As String = "Fehler"
Public Const StatusBez_Verbesserung          As String = "Verbesserung"

'Spezielle Zeichenketten für Oberfläche.
Public Const SpName_unbekannt                As String = "unbekannt"
Public Const SpTitel_unbekannt               As String = "< unbekannt >"
Public Const Allg_unbekannt                  As String = "unbekannt"

'Dateifilter.
Public Const DateiFilterXLS                  As String = "Exceldateien (*.xlsx), *.xlsx"
'Public Const DateiFilterXLT                  As String = "Exceldateien (*.xlt), *.xlt"
Public Const DateiMaskeXLT                   As String = "*.xlt;*.xltm;*.xltx"

'Allgemeines Verhalten
Public Const StatusBarClearDelay             As Integer = 7  'Verzögerung in Sekunden


'Array-Dimensionen und Indizes.
Public Const DP2lb                           As Long = 1  '2. Dimension des Datenpuffers
Public Const DP2ub                           As Long = 2
Public Const DPidxWert                       As Long = 1  'Wert der Zelle
Public Const DPidxFormel                     As Long = 2  'Formel der Zelle

'Fehlerniveau-Konstante
Public Const Fehlerniveau_Kein               As Long = 0
Public Const Fehlerniveau_Warnung            As Long = 1
Public Const Fehlerniveau_Kritisch           As Long = 2

'Fehler: Nummern und Zusatzmeldungen
Public Const ErrNumTabSchutz                 As Long = 50001
Public Const ErrNumKeineAktiveTabelle        As Long = 50002
Public Const ErrNumTabKlasseUngueltig        As Long = 50003
Public Const ErrNumZellnameProtected         As Long = 50004
Public Const ErrNumXLVorlageFehlt            As Long = 50005
Public Const ErrNumNoRangeSelection          As Long = 50006
Public Const ErrNumFktAufrufUngueltig        As Long = 50007
                                             
Public Const ErrMsgKeineAktiveTabelle        As String = "Es ist keine Tabelle aktiv!"
Public Const ErrMsgTabKlasseUngueltig        As String = "Die Tabellenvorlage und der Programmkode passen nicht zusammen!"
Public Const ErrMsgXLVorlageFehlt            As String = "Die angegebene Tabellenvorlage wird benötigt, kann aber weder in den Office-Vorlagenordnern noch in den Excel-Startordnern gefunden werden."

'Eigenschaften für Export-/Import-Objekte
Public Const io_Typ_XlTabNeu                 As String = "Neue_Excel_Tabelle"
Public Const io_Typ_XlTabAktiv               As String = "Aktive_Excel_Tabelle"
Public Const io_Typ_CsvSpezial               As String = "CSV_Spezial"
Public Const io_Typ_AsciiFormatiert          As String = "ASCII_Formatiert"
Public Const io_Typ_AsciiSpezial             As String = "ASCII_Spezial"
Public Const io_Typ_Puffer                   As String = "Datenpuffer"

Public Const io_Datei_Modus_Neu              As String = "Datei_Neu"
Public Const io_Datei_Modus_Ueberschreiben   As String = "Datei_Ueberschreiben"
Public Const io_Datei_Modus_Anhaengen        As String = "Datei_Anhaengen"

Public Const io_Klasse_PrefixImport          As String = "Cimp"
Public Const io_Klasse_PrefixExport          As String = "Cexp"
Public Const io_Klasse_Trassenkoo            As String = "CimpTrassenkoo"


'Spaltennamen in Tabellen für Export/Import
Public Const SpN_GK_X                        As String = "GK.X"
Public Const SpN_GK_Y                        As String = "GK.Y"
Public Const SpN_GK_Z                        As String = "GK.Z"

Public Const SpN_Pkt_Kz                      As String = "Pkt.Kz"
Public Const SpN_Pkt_Nr                      As String = "Pkt.Nr"
Public Const SpN_Pkt_Erl_H                   As String = "Pkt.Erl.H"
Public Const SpN_Pkt_Erl_L                   As String = "Pkt.Erl.L"

Public Const SpN_S_Tra_Radius                As String = "S.Tra.Radius"
Public Const SpN_S_Tra_Richtung              As String = "S.Tra.Richtung"
Public Const SpN_S_Tra_SO                    As String = "S.Tra.SO"
Public Const SpN_S_Tra_u                     As String = "S.Tra.u"
Public Const SpN_S_Tra_Heb                   As String = "S.Tra.Heb"

Public Const SpN_TK_H                        As String = "TK.H"
Public Const SpN_TK_HG                       As String = "TK.HG"
Public Const SpN_TK_HSOK                     As String = "TK.HSOK"
Public Const SpN_TK_Km                       As String = "TK.Km"
Public Const SpN_TK_L                        As String = "TK.L"
Public Const SpN_TK_Q                        As String = "TK.Q"
Public Const SpN_TK_QKm                      As String = "TK.QKm"
Public Const SpN_TK_QG                       As String = "TK.QG"
Public Const SpN_TK_R                        As String = "TK.R"
Public Const SpN_TK_St                       As String = "TK.St"
Public Const SpN_TK_V                        As String = "TK.V"
Public Const SpN_TK_RG                       As String = "TK.RG"
Public Const SpN_TK_LG                       As String = "TK.LG"
Public Const SpN_TK_QGT                      As String = "TK.QGT"
Public Const SpN_TK_HGT                      As String = "TK.HGT"
Public Const SpN_TK_QGS                      As String = "TK.QGS"
Public Const SpN_TK_HGS                      As String = "TK.HGS"
Public Const SpN_TK_KmStatus                 As String = "TK.KmStatus"

Public Const SpN_Tra_NameGra                 As String = "Tra.NameGra"
Public Const SpN_Tra_NameKML                 As String = "Tra.NameKML"
Public Const SpN_Tra_NameTra                 As String = "Tra.NameTra"
Public Const SpN_Tra_NameUeb                 As String = "Tra.NameUeb"
Public Const SpN_Tra_NameReg                 As String = "Tra.NameReg"
Public Const SpN_Tra_NameTun                 As String = "Tra.NameTun"
Public Const SpN_Tra_NamePkt                 As String = "Tra.NamePkt"
Public Const SpN_Tra_NameGls                 As String = "Tra.NameGls"
Public Const SpN_Tra_u                       As String = "Tra.u"

Public Const SpN_DGM_HDGM                    As String = "DGM.HDGM"
Public Const SpN_S_DGM_ZDGM                  As String = "S.DGM.ZDGM"
Public Const SpN_DGM_NameDGM                 As String = "DGM.NameDGM"

'Public Const SpN_F_GK_Z                      As String = "F.GK.Z"
'Public Const SpN_F_TK_H                      As String = "F.TK.H"
'Public Const SpN_F_TK_HG                     As String = "F.TK.HG"
'Public Const SpN_F_TK_HSOK                   As String = "F.TK.HSOK"
'Public Const SpN_F_TK_Km                     As String = "F.TK.Km"
'Public Const SpN_F_TK_Q                      As String = "F.TK.Q"
'Public Const SpN_F_TK_QG                     As String = "F.TK.QG"
'Public Const SpN_F_TK_R                      As String = "F.TK.R"
'Public Const SpN_F_TK_St                     As String = "F.TK.St"
'Public Const SpN_F_TK_V                      As String = "F.TK.V"
'Public Const SpN_F_Tra_u                     As String = "F.Tra.u"
'                                                                    
'Public Const SpN_V_GK_Z                      As String = "V.GK.Z"
'Public Const SpN_V_TK_H                      As String = "V.TK.H"
'Public Const SpN_V_TK_HG                     As String = "V.TK.HG"
'Public Const SpN_V_TK_HSOK                   As String = "V.TK.HSOK"
'Public Const SpN_V_TK_Km                     As String = "V.TK.Km"
'Public Const SpN_V_TK_Q                      As String = "V.TK.Q"
'Public Const SpN_V_TK_QG                     As String = "V.TK.QG"
'Public Const SpN_V_TK_R                      As String = "V.TK.R"
'Public Const SpN_V_TK_St                     As String = "V.TK.St"
'Public Const SpN_V_TK_V                      As String = "V.TK.V"
'Public Const SpN_V_Tra_u                     As String = "V.Tra.u"


'Objekte
Public oExpimGlobal                          As CdatExpim


Public CfgNichtGelesen                       As Boolean      'zeigt an, ob die Konfig.datei gelesen wurde.
Public ErrMessage                            As String       'zusätzliche Fehlerhinweise


'=== Emulierte Scripting-Konstanten  ==================================================
'WshShell.Run
Public Const WindowStyle_hidden                   As Long = 0
Public Const WindowStyle_normal                   As Long = 1
Public Const WindowStyle_minimized                As Long = 2
Public Const WindowStyle_maximized                As Long = 3
Public Const WaitOnReturn_yes                     As Boolean = true
Public Const WaitOnReturn_no                      As Boolean = false

'Dateioperationen
Public Const NewFileIfNotExist_yes                As Boolean = true
Public Const NewFileIfNotExist_no                 As Boolean = false
Public Const OpenAsASCII                          As Long = -0
Public Const OpenAsUnicode                        As Long = -1
Public Const OpenAsSystemDefault                  As Long = -2

Public Const ForReading                           As Long = 1
Public Const ForWriting                           As Long = 2
Public Const ForAppending                         As Long = 8

Public Const TristateFalse                        As Long = -0
Public Const TristateTrue                         As Long = -1
Public Const TristateUseDefault                   As Long = -2
  
'FileSystemObject.GetSpecialFolder
Public Const WindowsOrdner                        As Long = 0
Public Const SystemOrdner                         As Long = 1
Public Const TempOrdner                           As Long = 2


'für jEdit:  :folding=indent::collapseFolds=1:
