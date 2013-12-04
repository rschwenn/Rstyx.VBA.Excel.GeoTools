Attribute VB_Name = "mdlUserInterface"
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2003 - 2010  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'====================================================================================
'Modul mdlUserInterface
'====================================================================================
'Stellt die Benutzer-Befehle des Add-Ins zur Verfügung und bindet diese
'in ebenfalls hier erzeugte Menüs und Toolboxen ein.
'
'Die Import-Routinen stellen einen Dateiauswahldialog zur Verfügung,
'wenn die zu importierende Datei nicht gefunden wurde.


Option Explicit

Const Name_M1_Tools            As String = "&GeoTools"        '1. Ebene im Hauptmenü
Const Name_M2_Werkzeuge        As String = "&Werkzeuge"       '2. Ebene im Hauptmenü
Const Name_M2_Datenbereich     As String = "&Datenbereich"
Const Name_M2_DatenAendern     As String = "Be&rechnung"
Const Name_TBM_DatenAendern    As String = "Be&rechnung"
Const Name_M2_Projektdaten     As String = "&Tabellenkopf/Layout"
Const Name_TB_Datenbereich     As String = "gtDatenbereich"    'Toolbox
Const Name_TB_Werkzeuge        As String = "gtWerkzeuge"       'Toolbox

Const Name_TB_Icons            As String = "gtDummy_Icons"     'Kopiervorlagen der Icons
'



Sub Erzeuge_InfoKeineKonfig()
  'Erzeugt Button im GeoTools-Hauptmenü "==> Keine Konfiguration verfügbar!"
  'Wird aufgerufen von CdatKonfig\LeseKonfiguration().
  
  Dim cbc_M1_Tools    As CommandBarControl
  Dim cbb             As CommandBarButton
  Dim sTag            As String
  
  'Hauptmenu "GeoTools" finden.
  On Error Resume Next
  sTag = PrefixHauptmenue & TagHauptmenu_GeoTools
  Set cbc_M1_Tools = Application.CommandBars.FindControl(Tag:=sTag, Type:=msoControlPopup)
  
  'Neuen Menüpunkt einrichten.
  If ((Not (cbc_M1_Tools Is Nothing)) And (Not (Err))) Then
    'On Error GoTo 0
    On Error GoTo Fehler
    Set cbb = cbc_M1_Tools.Controls.Add(Type:=msoControlButton, Temporary:=True)
    cbb.Caption = "==> Keine &Konfiguration verfügbar!"
    cbb.OnAction = TagInfoKeineKonfig
    'cbIcons.Controls.Item(TagHilfe_GeoTools).CopyFace
    'cbb.PasteFace
    cbb.Tag = PrefixHauptmenue & TagInfoKeineKonfig
    cbb.BeginGroup = True
  End If
  
  Exit Sub
Fehler:
  FehlerNachricht "mdlUserInterface.Erzeuge_InfoKeineKonfig()"
End Sub


Sub MenuesErzeugen()
  'Symbolleisten und andere Menüeinträge erzeugen.
  'Wird aufgerufen von wbk_GeoTools\Workbook_Open().
  
  On Error GoTo Fehler
  
  Dim cbIcons               As CommandBar
  Dim cb_KM_Zelle           As CommandBar
  Dim cb_TB_Datenbereich    As CommandBar
  Dim cb_TB_Werkzeuge       As CommandBar
  Dim cbp_M1_Tools          As CommandBarPopup
  Dim cbp_M2_Datenbereich   As CommandBarPopup
  Dim cbp_M2_DatenAendern   As CommandBarPopup
  Dim cbp_TBM_DatenAendern  As CommandBarPopup
  Dim cbp_M2_Werkzeuge      As CommandBarPopup
  Dim cbp_M2_Projektdaten   As CommandBarPopup
  Dim cbb                   As CommandBarButton
  Dim cbb2                  As CommandBarButton
  Dim cbc                   As CommandBarControl
  Dim cbcb                  As CommandBarComboBox
  
  Dim i                     As Integer
  Dim TB_Werkzeuge_Neu      As Boolean
  Dim TB_Datenbereich_Neu   As Boolean
  
  'Prefixe für die Tag-Eigenschaft der Controls zwecks späteren Auffindens mit "FindControl"
  Set oMenuTypPrefixe = New Collection
  oMenuTypPrefixe.Add PrefixToolbox
  oMenuTypPrefixe.Add PrefixHauptmenue
  oMenuTypPrefixe.Add PrefixKontextZelle
  
  'alle Basistags für Controls. Diese bilden zusammen mit den Prefixen die Tags.
  Set oBasisTags = New Collection
  oBasisTags.Add TagInfo_GeoTools
  oBasisTags.Add TagHilfe_GeoTools
  oBasisTags.Add TagFormatDaten
  oBasisTags.Add TagFormatDatenMitStreifen
  oBasisTags.Add TagFormatDatenOhneFuellung
  oBasisTags.Add TagLoeschenDaten
  oBasisTags.Add TagFormatDatenNKStellenSetzen
  oBasisTags.Add TagFormatDatenNKStellenAnzahl
  oBasisTags.Add TagSchreibeProjektDaten
  oBasisTags.Add TagSchreibeFusszeile_1
  oBasisTags.Add TagTabellenStruktur
  oBasisTags.Add TagUebertragenFormeln
  oBasisTags.Add TagSelection2Interpolationsformel
  oBasisTags.Add TagSelection2MarkDoppelteWerte
  oBasisTags.Add TagInsertLines
  oBasisTags.Add TagDateiBearbeiten
  oBasisTags.Add TagImport_Pktpaare
  oBasisTags.Add TagImport_NivLinien
  oBasisTags.Add TagModOpt_VorhWerteUeberschreiben
  oBasisTags.Add TagModOpt_FormelnErhalten
  oBasisTags.Add TagMod_Transfo_Tk2Gls
  oBasisTags.Add TagMod_Transfo_Gls2Tk
  oBasisTags.Add TagMod_UeberhoehungAusBemerkung
  oBasisTags.Add TagMod_FehlerVerbesserung
  oBasisTags.Add TagExpimManager
  
  
  'Alle Controls löschen, deren Tags den Konventionen dieses AddIns entsprechen -
  'wo auch immer diese Controls versteckt sind und wie viele Exemplare es davon gibt.
  '==> Dies ist nötig, damit das Setzen der sichtbaren Controls auf den Status
  'aktiv/inaktiv funktionieren kann.
  AlleControlsLoeschen
  
  'Icon-Vorlagen
  Set cbIcons = Application.CommandBars(Name_TB_Icons)
  'Set cbIcons = ThisWorkbook.CommandBars(Name_TB_Icons)
  If (ThisWorkbook.IsAddin) Then
    cbIcons.Enabled = False
  Else
    cbIcons.Enabled = True
  End If
  
  'Kontextmenü für Zellen
  Set cb_KM_Zelle = Application.CommandBars("cell")
  '"Bedingte Formatierung"
  Set cbb = cb_KM_Zelle.Controls.Add(Type:=msoControlButton, Id:=3058, Temporary:=True, before:=cb_KM_Zelle.Controls.Count - 1)
  
  'temporäre Symbolleisten und Menüstruktur erzeugen
  On Error Resume Next
  'Toolboxen erzeugen. Falls bereits vorhanden, sind diese (hoffentlich) leer.
  Set cb_TB_Werkzeuge = Application.CommandBars(Name_TB_Werkzeuge)
  If (cb_TB_Werkzeuge Is Nothing) Then
    'Leere Toolbox erzeugen, da sie noch nicht existiert.
    Set cb_TB_Werkzeuge = Application.CommandBars.Add(Name:=Name_TB_Werkzeuge, Position:=msoBarTop)
    TB_Werkzeuge_Neu    = True
  End If
  Set cb_TB_Datenbereich = Application.CommandBars(Name_TB_Datenbereich)
  If (cb_TB_Datenbereich Is Nothing) Then
    'Leere Toolbox erzeugen, da sie noch nicht existiert.
    Set cb_TB_Datenbereich = Application.CommandBars.Add(Name:=Name_TB_Datenbereich, Position:=msoBarTop)
    TB_Datenbereich_Neu = True
  End If
  On Error GoTo Fehler
  
  'Toolboxen zunächst deaktivieren. Somit bleibt der Aufbauprozess unsichtbar.
  cb_TB_Werkzeuge.Enabled = False
  cb_TB_Datenbereich.Enabled = False
  
  'temporäres Menü "GeoTools" in der Standard-Menüleiste erzeugen
  Set cbp_M1_Tools = Application.CommandBars("Worksheet Menu Bar").Controls.Add(Type:=msoControlPopup, Temporary:=True, before:=Application.CommandBars("Worksheet Menu Bar").Controls.Count - 1)
  cbp_M1_Tools.Caption = Name_M1_Tools
  cbp_M1_Tools.Tag = PrefixHauptmenue & TagHauptmenu_GeoTools
  
  
  'temporäre Submenüs unter "GeoTools" erzeugen
  Set cbp_M2_Projektdaten = cbp_M1_Tools.Controls.Add(Type:=msoControlPopup, Temporary:=True)
  cbp_M2_Projektdaten.Caption = Name_M2_Projektdaten
  Set cbp_M2_Werkzeuge = cbp_M1_Tools.Controls.Add(Type:=msoControlPopup, Temporary:=True)
  cbp_M2_Werkzeuge.Caption = Name_M2_Werkzeuge
  Set cbp_M2_Datenbereich = cbp_M1_Tools.Controls.Add(Type:=msoControlPopup, Temporary:=True)
  cbp_M2_Datenbereich.Caption = Name_M2_Datenbereich
  Set cbp_M2_DatenAendern = cbp_M1_Tools.Controls.Add(Type:=msoControlPopup, Temporary:=True)
  cbp_M2_DatenAendern.Caption = Name_M2_DatenAendern
  
  
  '===================================================================================
  'Erzeugen der Menüeinträge: Nach der Definition eines Buttons in der Toolbox
  'wird dieser sofort in das Hauptmenü und evtl. in das Kontextmenü kopiert.
  'Die "Tags" werden eindeutig gehalten!
  
  
  'Button "ExpimManager"
  Set cbb = cb_TB_Werkzeuge.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.Caption = "Import / E&xport"
  cbb.TooltipText = "Import/Export von und nach ASCII/Excel"
  cbb.OnAction = TagExpimManager
  'cbIcons.Controls.Item(TagExpimManager).CopyFace
  'cbb.PasteFace
  cbb.FaceID = 688
  cbb.Tag = PrefixHauptmenue & TagExpimManager
  cbb.Copy Bar:=cbp_M1_Tools.CommandBar
  cbb.Tag = PrefixToolbox & TagExpimManager
  
  
  'Button "ModOpt_VorhWerteUeberschreiben"
  Set cbb = cb_TB_Werkzeuge.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.BeginGroup = True
  cbb.Caption = "Vorhandene Werte überschreiben/stehenlassen."
  cbb.OnAction = TagModOpt_VorhWerteUeberschreiben
  cbIcons.Controls.Item(TagModOpt_VorhWerteUeberschreiben).CopyFace
  cbb.PasteFace
  cbb.Tag = PrefixHauptmenue & TagModOpt_VorhWerteUeberschreiben
  cbb.Copy Bar:=cbp_M2_DatenAendern.CommandBar
  cbb.Tag = PrefixToolbox & TagModOpt_VorhWerteUeberschreiben
  
  
  'Button "ModOpt_FormelnErhalten"
  Set cbb = cb_TB_Werkzeuge.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.Caption = "Formeln erhalten oder durch Werte ersetzen?."
  cbb.OnAction = TagModOpt_FormelnErhalten
  cbIcons.Controls.Item(TagModOpt_FormelnErhalten).CopyFace
  cbb.PasteFace
  cbb.Tag = PrefixHauptmenue & TagModOpt_FormelnErhalten
  cbb.Copy Bar:=cbp_M2_DatenAendern.CommandBar
  cbb.Tag = PrefixToolbox & TagModOpt_FormelnErhalten
  
  'Popup-Menü in der Toolbox.
  Set cbp_TBM_DatenAendern = cb_TB_Werkzeuge.Controls.Add(Type:=msoControlPopup, Temporary:=True)
  cbp_TBM_DatenAendern.Caption = Name_TBM_DatenAendern
  
  'Button "Mod_FehlerVerbesserung"
  Set cbb = cbp_TBM_DatenAendern.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.Caption = "Fehler und Verbesserungen"
  cbb.OnAction = TagMod_FehlerVerbesserung
  'cbIcons.Controls.Item(TagMod_FehlerVerbesserung).CopyFace
  'cbb.PasteFace
  cbb.Tag = PrefixHauptmenue & TagMod_FehlerVerbesserung
  cbb.Copy Bar:=cbp_M2_DatenAendern.CommandBar
  cbb.Tag = PrefixToolbox & TagMod_FehlerVerbesserung
  cbp_M2_DatenAendern.Controls("Fehler und Verbesserungen").BeginGroup = True
  
  
  'Button "Mod_UeberhoehungAusBemerkung"
  Set cbb = cbp_TBM_DatenAendern.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.Caption = "Ist-Überhöhung aus Bemerkung"
  cbb.OnAction = TagMod_UeberhoehungAusBemerkung
  'cbIcons.Controls.Item(TagMod_UeberhoehungAusBemerkung).CopyFace
  'cbb.PasteFace
  cbb.Tag = PrefixHauptmenue & TagMod_UeberhoehungAusBemerkung
  cbb.Copy Bar:=cbp_M2_DatenAendern.CommandBar
  cbb.Tag = PrefixToolbox & TagMod_UeberhoehungAusBemerkung
  
  
  'Button "Mod_Transfo_Tk2Gls"
  Set cbb = cbp_TBM_DatenAendern.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.Caption = "Trassenkoo' => Gleissystem"
  cbb.OnAction = TagMod_Transfo_Tk2Gls
  'cbIcons.Controls.Item(TagMod_Transfo_Tk2Gls).CopyFace
  'cbb.PasteFace
  cbb.Tag = PrefixHauptmenue & TagMod_Transfo_Tk2Gls
  cbb.Copy Bar:=cbp_M2_DatenAendern.CommandBar
  cbb.Tag = PrefixToolbox & TagMod_Transfo_Tk2Gls
  
  
  'Button "Mod_Transfo_Gls2Tk"
  Set cbb = cbp_TBM_DatenAendern.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.Caption = "Gleissystem => Trassenkoo'"
  cbb.OnAction = TagMod_Transfo_Gls2Tk
  'cbIcons.Controls.Item(TagMod_Transfo_Gls2Tk).CopyFace
  'cbb.PasteFace
  cbb.Tag = PrefixHauptmenue & TagMod_Transfo_Gls2Tk
  cbb.Copy Bar:=cbp_M2_DatenAendern.CommandBar
  cbb.Tag = PrefixToolbox & TagMod_Transfo_Gls2Tk
  
  
  'Button "Selection2Interpolationsformel"
  Set cbb = cb_TB_Werkzeuge.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.BeginGroup = True
  cbb.Caption = "Interpolationsformel erstellen"
  cbb.TooltipText = "Interpolationsformel erstellen (Markierung = 3 Zellen!)"
  cbb.OnAction = TagSelection2Interpolationsformel
  'cbIcons.Controls.Item(TagSelection2Interpolationsformel).CopyFace
  'cbb.PasteFace
  cbb.FaceID = 620
  cbb.Tag = PrefixHauptmenue & TagSelection2Interpolationsformel
  cbb.Copy Bar:=cbp_M2_Werkzeuge.CommandBar
  cbb.Tag = PrefixToolbox & TagSelection2Interpolationsformel
  'cbp_M2_Werkzeuge.Controls("Interpolationsformel erstellen").BeginGroup = True
  
  'Button "Selection2MarkDoppelteWerte"
  Set cbb = cb_TB_Werkzeuge.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.Caption = "Doppelte Werte markieren"
  cbb.TooltipText = "Doppelte Werte markieren (ab der markierten Zelle)"
  cbb.OnAction = TagSelection2MarkDoppelteWerte
  cbIcons.Controls.Item(TagSelection2MarkDoppelteWerte).CopyFace
  cbb.PasteFace
  cbb.Tag = PrefixHauptmenue & TagSelection2MarkDoppelteWerte
  cbb.Copy Bar:=cbp_M2_Werkzeuge.CommandBar
  cbb.Tag = PrefixToolbox & TagSelection2MarkDoppelteWerte
  
  'Button "insertLines"
  Set cbb = cb_TB_Werkzeuge.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.Caption = "Leerzeilen einfügen"
  cbb.TooltipText = "Leerzeilen im Intervall einfügen (mit Dialog)"
  cbb.OnAction = TagInsertLines
  cbb.FaceID = 56
  cbb.Tag = PrefixHauptmenue & TagInsertLines
  cbb.Copy Bar:=cbp_M2_Werkzeuge.CommandBar
  cbb.Tag = PrefixToolbox & TagInsertLines
  
  'Button "DateiBearbeiten"
  Set cbb = cbp_M2_Werkzeuge.Controls.Add(Type:=msoControlButton, Temporary:=True)
  'Set cbb = cb_TB_Werkzeuge.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.Caption = "Datei öffnen (Name in Zelle)"
  cbb.TooltipText = "Die Datei, deren Name die aktive Zelle enthält, wird im Editor geöffnet"
  cbb.OnAction = TagDateiBearbeiten
  cbb.FaceID = 940
  'cbb.Tag = PrefixToolbox & TagDateiBearbeiten
  'cbb.Copy Bar:=cbp_M2_Werkzeuge.CommandBar
  On Error Resume Next
  Set cbb2 = cb_KM_Zelle.Controls("Datei öffnen (Name in Zelle)")
  If (cbb2 Is Nothing) Then
    cbb.Tag = PrefixKontextZelle & TagDateiBearbeiten
    cbb.Copy Bar:=cb_KM_Zelle
    cb_KM_Zelle.Controls("Datei öffnen (Name in Zelle)").BeginGroup = True
  End If
  'On Error GoTo Fehler
  On Error GoTo 0
  cbb.Tag = PrefixHauptmenue & TagDateiBearbeiten
  
  
  'Button "FormatDaten"
  Set cbb = cb_TB_Datenbereich.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.Caption = "Datenbereich formatieren"
  cbb.OnAction = TagFormatDaten
  cbIcons.Controls.Item(TagFormatDaten).CopyFace
  cbb.PasteFace
  On Error Resume Next
  Set cbb2 = cb_KM_Zelle.Controls("Datenbereich formatieren")
  If (cbb2 Is Nothing) Then
    cbb.Tag = PrefixKontextZelle & TagFormatDaten
    cbb.Copy Bar:=cb_KM_Zelle
    cb_KM_Zelle.Controls("Datenbereich formatieren").BeginGroup = False
  End If
  'On Error GoTo Fehler
  On Error GoTo 0
  
  cbb.Tag = PrefixHauptmenue & TagFormatDaten
  cbb.Copy Bar:=cbp_M2_Datenbereich.CommandBar
  cbb.Tag = PrefixToolbox & TagFormatDaten
  
  'Button "FormatDatenMitStreifen"
  Set cbb = cb_TB_Datenbereich.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.BeginGroup = True
  cbb.Caption = "Formatierung mit/ohne Streifen"
  cbb.OnAction = TagFormatDatenMitStreifen
  cbIcons.Controls.Item(TagFormatDatenMitStreifen).CopyFace
  cbb.PasteFace
  cbb.Tag = PrefixHauptmenue & TagFormatDatenMitStreifen
  cbb.Copy Bar:=cbp_M2_Datenbereich.CommandBar
  cbb.Tag = PrefixToolbox & TagFormatDatenMitStreifen
  cbp_M2_Datenbereich.Controls("Formatierung mit/ohne Streifen").BeginGroup = True
  
  'Button "FormatDatenOhneFuellung"
  Set cbb = cb_TB_Datenbereich.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.Caption = "Formatierung mit/ohne Löschen der Füllung"
  cbb.OnAction = TagFormatDatenOhneFuellung
  cbIcons.Controls.Item(TagFormatDatenOhneFuellung).CopyFace
  cbb.PasteFace
  cbb.Tag = PrefixHauptmenue & TagFormatDatenOhneFuellung
  cbb.Copy Bar:=cbp_M2_Datenbereich.CommandBar
  cbb.Tag = PrefixToolbox & TagFormatDatenOhneFuellung
  
  'Button "FormatDatenNKStellenSetzen"
  Set cbb = cb_TB_Datenbereich.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.Caption = "Formatierung mit/ohne Änderung der NK-Stellen"
  cbb.TooltipText = "Formatierung mit/ohne Ändern der Nachkommastellen in den 'Fliesskomma'-Spalten"
  cbb.OnAction = TagFormatDatenNKStellenSetzen
  cbIcons.Controls.Item(TagFormatDatenNKStellenSetzen).CopyFace
  cbb.PasteFace
  cbb.Tag = PrefixHauptmenue & TagFormatDatenNKStellenSetzen
  cbb.Copy Bar:=cbp_M2_Datenbereich.CommandBar
  cbb.Tag = PrefixToolbox & TagFormatDatenNKStellenSetzen
  
  'Listenfeld "FormatDatenNKStellenAnzahl"
  Set cbcb = cb_TB_Datenbereich.Controls.Add(Type:=msoControlDropdown, Temporary:=True)
  cbcb.Tag = PrefixToolbox & TagFormatDatenNKStellenAnzahl
  cbcb.Caption = "Anzahl der Nachkommastellen"
  'cbcb.Caption = TagFormatDatenNKStellenAnzahl
  'cbcb.TooltipText = "Anzahl der Nachkommastellen"
  cbcb.OnAction = TagFormatDatenNKStellenAnzahl
  cbcb.Width = 35
  cbcb.DropDownWidth = -1
  For i = 0 To 9
    cbcb.AddItem i
  Next
  cbcb.ListIndex = 3
  
  'Button "TabellenStruktur"
  Set cbb = cb_TB_Datenbereich.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.BeginGroup = True
  cbb.Caption = "Tabellenstruktur verwalten"
  cbb.TooltipText = "Festlegung der Strukturelemente und Spaltenbezeichnungen"
  cbb.OnAction = TagTabellenStruktur
  'cbIcons.Controls.Item(TagTabellenStruktur).CopyFace
  'cbb.PasteFace
  cbb.FaceID = 583 '333
  cbb.Tag = PrefixHauptmenue & TagTabellenStruktur
  cbb.Copy Bar:=cbp_M2_Datenbereich.CommandBar
  cbb.Tag = PrefixToolbox & TagTabellenStruktur
  cbp_M2_Datenbereich.Controls("Tabellenstruktur verwalten").BeginGroup = True
  
  'Button "UebertragenFormeln"
  Set cbb = cb_TB_Datenbereich.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.BeginGroup = True
  cbb.Caption = "Formeln übertragen"
  cbb.TooltipText = "Übertragung der Formeln des 'Formel'-Bereiches der 1. Datenzeile auf alle weiteren Zeilen."
  cbb.OnAction = TagUebertragenFormeln
  cbIcons.Controls.Item(TagUebertragenFormeln).CopyFace
  cbb.PasteFace
  cbb.Tag = PrefixHauptmenue & TagUebertragenFormeln
  cbb.Copy Bar:=cbp_M2_Datenbereich.CommandBar
  cbb.Tag = PrefixToolbox & TagUebertragenFormeln
  cbp_M2_Datenbereich.Controls("Formeln übertragen").BeginGroup = True
  
  'Button "LoeschenDaten"
  Set cbb = cb_TB_Datenbereich.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.Caption = "Datenbereich löschen"
  cbb.TooltipText = "Datenbereich komplett löschen"
  cbb.OnAction = TagLoeschenDaten
  cbIcons.Controls.Item(TagLoeschenDaten).CopyFace
  cbb.PasteFace
  cbb.Tag = PrefixHauptmenue & TagLoeschenDaten
  cbb.Copy Bar:=cbp_M2_Datenbereich.CommandBar
  cbb.Tag = PrefixToolbox & TagLoeschenDaten
  
  
  'Button "SchreibeProjektDaten"
  Set cbb = cbp_M2_Projektdaten.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.Caption = "Projektdaten eintragen"
  cbb.OnAction = TagSchreibeProjektDaten
  'cbIcons.Controls.Item(TagSchreibeProjektDaten).CopyFace
  'cbb.PasteFace
  cbb.FaceID = 931
  cbb.Tag = PrefixHauptmenue & TagSchreibeProjektDaten
  
  
  'Button "SchreibeFusszeile_1"
  Set cbb = cbp_M2_Projektdaten.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.Caption = "Fusszeile eintragen"
  cbb.OnAction = TagSchreibeFusszeile_1
  'cbIcons.Controls.Item(TagSchreibeFusszeile_1).CopyFace
  'cbb.PasteFace
  cbb.FaceID = 29
  cbb.Tag = PrefixHauptmenue & TagSchreibeFusszeile_1
  
  
  'Button "Protokoll"
  Set cbb = cbp_M1_Tools.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.Caption = "&Protokoll"
  cbb.OnAction = TagProtokoll
  'cbIcons.Controls.Item(TagHilfe_GeoTools).CopyFace
  'cbb.PasteFace
  cbb.FaceID = 521
  cbb.Tag = PrefixHauptmenue & TagHilfe_GeoTools
  cbb.BeginGroup = True
  
  
  'Button "Hilfe_GeoTools"
  Set cbb = cbp_M1_Tools.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.Caption = "&Hilfe"
  cbb.OnAction = TagHilfe_GeoTools
  'cbIcons.Controls.Item(TagHilfe_GeoTools).CopyFace
  'cbb.PasteFace
  cbb.FaceID = 984
  cbb.Tag = PrefixHauptmenue & TagHilfe_GeoTools
  cbb.BeginGroup = false
  
  
  'Button "GeoTools_Info"
  Set cbb = cbp_M1_Tools.Controls.Add(Type:=msoControlButton, Temporary:=True)
  cbb.Caption = "&Info"
  cbb.OnAction = TagInfo_GeoTools
  'cbIcons.Controls.Item(TagInfo_GeoTools).CopyFace
  'cbb.PasteFace
  cbb.Tag = PrefixHauptmenue & TagInfo_GeoTools
  
  
  'Toolboxen verfügbar machen.
  cb_TB_Werkzeuge.Enabled = True
  cb_TB_Datenbereich.Enabled = True
  
  'Toolboxen anzeigen, falls sie bisher nicht existierten
  'Anderenfalls wird die vom Benutzer eingestellte Sichtbarkeit nicht geändert.
  if (TB_Werkzeuge_Neu) then cb_TB_Werkzeuge.Visible = True
  
  
  if (TB_Datenbereich_Neu) then cb_TB_Datenbereich.Visible = True
  
  
  'neues Kontext-Menü vorbereiten
  'Set cb = Application.CommandBars.Add(Name:="NewPopup", Position:=msoBarPopup)
  'With cb
  '   .Controls.Add Type:=msoControlButton, Id:=3  'Speichern
  'End With
  
  ' neuer Eintrag im Extras-Menü
  'Set cbb = Application.CommandBars("Worksheet Menu Bar").Controls("Extras").Controls.Add()
  'cbb.Caption = "Ein neues Kommando"
  'cbb.BeginGroup = True
  'cbb.OnAction = "NewCommand"
  
  'Dummy-Toolbox mit Icons löschen, damit beim nächsten Start des Add-In die
  'aktuelle, an das Add-In gebundene Symbolleiste nach Excel kopiert wird.
  If (ThisWorkbook.IsAddin) Then
    cbIcons.Delete
  End If
  
  Set cbIcons = Nothing
  Set cb_KM_Zelle = Nothing
  Set cb_TB_Datenbereich = Nothing
  Set cb_TB_Werkzeuge = Nothing
  Set cbp_M1_Tools = Nothing
  Set cbp_M2_Datenbereich = Nothing
  Set cbp_M2_Werkzeuge = Nothing
  Set cbp_M2_Projektdaten = Nothing
  Set cbb = Nothing
  Set cbb2 = Nothing
  Set cbc = Nothing
  Set cbcb = Nothing
  
  Exit Sub
  
Fehler:
  If (ThisWorkbook.IsAddin) Then
    cbIcons.Delete
  End If
  ErrMessage = "Die Benutzeroberfläche für " & Err.Source & " konnte nicht erzeugt werden. " & vbNewLine & vbNewLine & _
               "Nach dem nächsten Start von Excel sollte das Problem hoffentlich behoben sein."
  FehlerNachricht "mdlUserInterface.MenuesErzeugen()"
  'Application.Quit  'zu gefährlich (was, wenn Selbstheilung nicht funktioniert?)
End Sub


Sub MenuesEntfernen()
  'Alle erzeugten Symbolleisten und Menüeinträge entfernen.
  'Wird aufgerufen von wbk_GeoTools\Workbook_BeforeClose().
  
  On Error Resume Next
  
  msgbox "MenuesEntfernen"
  'Toolboxen nicht löschen, damit sich Excel die Position merken kann;
  'aber deaktivieren, damit bei Excel-Start ohne dieses Add-In die leere
  'Toolbox für den Benutzer nicht verfügbar ist im "Anpassen"-Dialog.
  Application.CommandBars(Name_TB_Datenbereich).Enabled = False
  Application.CommandBars(Name_TB_Werkzeuge).Enabled = False
  
  'Einträge im Hauptmenü entfernen
  Application.CommandBars("Worksheet Menu Bar").Controls(Name_M1_Tools).Delete
  
  'Einträge im Kontextmenü entfernen
  Application.CommandBars("cell").Controls("Datenbereich formatieren").Delete
  Application.CommandBars("cell").Controls("Bedingte Formatierung...").Delete
  
  'neues Kontextmenü entfernen
  'Application.CommandBars("newpopup").Delete
  
  'Set oMenuTypPrefixe = Nothing
  'Set oBasisTags = Nothing

End Sub


Private Sub AlleControlsLoeschen()
  'Alle Controls löschen, deren Tags den Konventionen dieses AddIns entsprechen -
  'wo auch immer diese Controls versteckt sind und wie viele Exemplare es davon gibt.
  
  On Error Resume Next
  Dim cbc      As CommandBarControl
  Dim Prefix   As Variant
  Dim cTag     As Variant
  Dim sTag     As String
  Dim gefunden As Boolean
  
  Do
    gefunden = False
    For Each cTag In oBasisTags
      For Each Prefix In oMenuTypPrefixe
        sTag = Prefix & cTag
        Set cbc = Application.CommandBars.FindControl(Tag:=sTag)
        If ((Not (cbc Is Nothing)) And (Not (Err))) Then
          gefunden = True
          'MsgBox "Tag=" & sTag & vbNewLine & _
                 "Steuerfeld/Caption=" & cbc.Caption & " (aktiv=" & cbc.Enabled & ")" & vbNewLine & _
                 "enthalten in: " & cbc.Parent.Name & " (sichtbar=" & cbc.Parent.Visible & ")" & vbNewLine & vbNewLine & _
                 "==> wird gelöscht!"
          cbc.Delete
          gefunden = True
        End If
      Next
    Next
  Loop Until gefunden = False
  
  Set cbc = Nothing
End Sub


Sub SetSilent_AktiveTabelle(inpSilent As Boolean)
  'Setzt den aktuellen Modus für "Silent" im Objekt oAktiveTabelle.
  On Error GoTo 0
  oAktiveTabelle.Silent = inpSilent
End Sub


Sub FormatDatenMitStreifen()
  'Reaktion auf Buttonklick "FormatDatenMitStreifen"
  'Änderung von Tooltip und Status des Buttons übernimmt "Property Let FormatDatenMitStreifen"
  oAktiveTabelle.FormatDatenMitStreifen = Not oAktiveTabelle.FormatDatenMitStreifen
End Sub


Sub FormatDatenOhneFuellung()
  'Reaktion auf Buttonklick "FormatDatenOhneFuellung"
  'Änderung von Tooltip und Status des Buttons übernimmt "Property Let FormatDatenOhneFuellung"
  oAktiveTabelle.FormatDatenOhneFuellung = Not oAktiveTabelle.FormatDatenOhneFuellung
End Sub


Sub FormatDatenNKStellenSetzen()
  'Reaktion auf Buttonklick "FormatDatenNKStellenSetzen"
  'Änderung von Tooltip und Status des Buttons übernimmt "Property Let FormatDatenNKStellenSetzen"
  oAktiveTabelle.FormatDatenNKStellenSetzen = Not oAktiveTabelle.FormatDatenNKStellenSetzen
End Sub


Sub FormatDatenNKStellenAnzahl()
  'Reaktion auf Auswahl in der Combobox "FormatDatenNKStellenAnzahl"
  On Error Resume Next
  Dim cbcb As CommandBarComboBox
  Set cbcb = CommandBars.FindControl(Type:=msoControlDropdown, Tag:=PrefixToolbox & TagFormatDatenNKStellenAnzahl)
  If ((Not (cbcb Is Nothing)) And (Not (Err))) Then
    oAktiveTabelle.FormatDatenNKStellenAnzahl = CInt(cbcb.text)
  End If
End Sub


Sub SchreibeProjektDaten()
  'Schreibt von allen verfügbaren Projektdaten diejenigen in die aktive Tabelle,
  'für die entsprechend benannte Zellen existieren.
  On Error GoTo Fehler
  oMetadaten.Update oPrjLocal:=nothing, oExtraLocal:=nothing
  oAktiveTabelle.SchreibeMetaDaten
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlUserInterface.SchreibeProjektDaten()"
End Sub


Sub SchreibeFusszeile_1()
  'Schreibt die Fusszeile_1 in die aktive Tabelle.
  On Error GoTo Fehler
  oAktiveTabelle.SchreibeFusszeile_1
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlUserInterface.SchreibeFusszeile_1()"
End Sub


Sub LoeschenDaten()
  'Alle Datenzeilen der aktiven Tabelle werden gelöscht.
  Dim Titel    As String
  Dim Message  As String
  Dim Buttons  As Integer
  On Error GoTo Fehler
  Message = "Soll der gesamte Datenbereich der Tabelle wirklich gelöscht werden? " & vbNewLine & vbNewLine & _
            "==> Diese Aktion kann NICHT rückgängig gemacht werden!!!"
  Buttons = vbYesNo + vbQuestion + vbDefaultButton2
  Titel = "Datenbereich löschen!"
  If (MsgBox(Message, Buttons, Titel) = vbYes) Then
    oAktiveTabelle.LoeschenDaten
    call ClearStatusBarDelayed(StatusBarClearDelay)
  End If
  Exit Sub
Fehler:
  FehlerNachricht "mdlUserInterface.LoeschenDaten()"
End Sub


Sub FormatDaten()
  'Überträgt das Format des "InfoTraegers" auf alle weiteren Datenzeilen.
  On Error GoTo Fehler
  oAktiveTabelle.FormatDaten
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlUserInterface.FormatDaten()"
End Sub


Sub UebertragenFormeln()
  'Überträgt die Formeln des 'Formel'-Bereiches auf alle weiteren Datenzeilen.
  On Error GoTo Fehler
  oAktiveTabelle.UebertragenFormeln
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlUserInterface.UebertragenFormeln()"
End Sub


Sub ModOpt_VorhWerteUeberschreiben()
  'Reaktion auf Buttonklick "ModOpt_VorhWerteUeberschreiben"
  'Änderung von Caption und Status des Buttons übernimmt "Property Let ModOpt_VorhWerteUeberschreiben"
  oAktiveTabelle.ModOpt_VorhWerteUeberschreiben = Not oAktiveTabelle.ModOpt_VorhWerteUeberschreiben
End Sub


Sub ModOpt_FormelnErhalten()
  'Reaktion auf Buttonklick "ModOpt_FormelnErhalten"
  'Änderung von Caption und Status des Buttons übernimmt "Property Let ModOpt_FormelnErhalten"
  oAktiveTabelle.ModOpt_FormelnErhalten = Not oAktiveTabelle.ModOpt_FormelnErhalten
End Sub


Sub Mod_FehlerVerbesserung()
  'Modifiziert Daten der aktiven Tabelle.
  On Error GoTo Fehler
  oAktiveTabelle.Mod_FehlerVerbesserung
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlUserInterface.Mod_FehlerVerbesserung()"
End Sub


Sub Mod_UeberhoehungAusBemerkung()
  'Modifiziert Daten der aktiven Tabelle.
  On Error GoTo Fehler
  oAktiveTabelle.Mod_UeberhoehungAusBemerkung
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlUserInterface.Mod_UeberhoehungAusBemerkung()"
End Sub


Sub Mod_Transfo_Tk2Gls()
  'Modifiziert Daten der aktiven Tabelle.
  On Error GoTo Fehler
  oAktiveTabelle.Mod_Transfo_Tk2Gls
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlUserInterface.Mod_Transfo_Tk2Gls()"
End Sub


Sub Mod_Transfo_Gls2Tk()
  'Modifiziert Daten der aktiven Tabelle.
  On Error GoTo Fehler
  oAktiveTabelle.Mod_Transfo_Gls2Tk
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlUserInterface.Mod_Transfo_Gls2Tk()"
End Sub


Sub TabellenStruktur()
  'Dialog "Tabellenstruktur und Spaltennamen verwalten" anzeigen.
  Dim Dialog    As frmSpaltenVerw
  Set Dialog = New frmSpaltenVerw
  Dialog.Show vbModeless
  Set Dialog = Nothing
  Exit Sub
Fehler:
  FehlerNachricht "mdlUserInterface.TabellenStruktur()"
End Sub


Sub Selection2Interpolationsformel()
  'Aufgrund der aktuellen Zellauswahl wird eine Interpolationsformel erstellt.
  On Error GoTo Fehler
  oAktiveTabelle.Selection2Interpolationsformel
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlUserInterface.Selection2Interpolationsformel()"
End Sub


Sub Selection2MarkDoppelteWerte()
  'Die markierten (und alle darunter liegenden) Zellen werden mit einer bedingten
  'Formatierung versehen, die alle Zellen mit solchen Werten hervorhebt, die in
  'derselben Spalte mehr als einmal vorkommen.
  On Error GoTo Fehler
  oAktiveTabelle.Selection2MarkDoppelteWerte
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlUserInterface.Selection2MarkDoppelteWerte()"
End Sub


Sub insertLines()
  'Dialog "Leerzeilen einfügen"
  On Error GoTo Fehler
  Dim oDialog    As frmInsertLines
  Set oDialog = New frmInsertLines
  oDialog.Show
  Exit Sub
Fehler:
  Set oDialog = Nothing
  FehlerNachricht "mdlUserInterface.insertLines()"
End Sub


Sub DateiBearbeiten()
  'Die Datei, deren Name dem Inhalt der aktiven Zelle entspricht, wird im
  'konfigurierten Editor geladen. Bei Mißerfolg wird die Windows-Standardanwendung
  'des Dateityps gestartet.
  On Error GoTo Fehler
  dim Datei
  If (ActiveCell Is Nothing) Then
    Application.StatusBar = "Fehler: Es existiert keine aktive Zelle!"
  else
    Datei = trim(ActiveCell.Value)
    if (not oSysTools.isDatei(Datei)) then
      Application.StatusBar = "Der Inhalt der aktiven Zelle ('" & Datei & "') bezeichnet keine existierende Datei !"
    else
      Application.StatusBar = "Datei '" & Datei & "' wird im Editor geöffnet."
      if (not oSysTools.StartEditor(Datei)) then
        Application.StatusBar = "Datei '" & Datei & "' wird mit Standardanwendung geöffnet."
        oSysTools.StarteDatei(Datei)
      end if
    end if
  end if
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlUserInterface.DateiBearbeiten()"
End Sub


Sub Import_Trassenkoo(Optional ByVal ParamDateiName As String = "")
  'Import von Trassenkoordinaten in eine neue Arbeitsmappe.
  'Parameter: ParamDateiName = Pfad\Name der Eingabedatei.
  '=> Diese Routine wird nur zur Fernsteuerng verwendet.
  On Error GoTo Fehler
  if (not oExpimGlobal is Nothing) then
    Application.Visible        = true
    Application.UserControl    = true
    Application.ScreenUpdating = true
    msgbox "Es ist bereits eine Export / Import - Aktion aktiv. => Eine zweite kann nicht gestartet werden!" , vbOKOnly, "Import Trassenkoordinaten"
  else
    Set oExpimGlobal = New CdatExpim
    If (Err) Then GoTo Fehler
    oExpimGlobal.Quelle_Typ = io_Typ_AsciiSpezial
    oExpimGlobal.Quelle_FormatID = "CimpTrassenkoo"
    oExpimGlobal.Quelle_AsciiDatei_Name = ParamDateiName
    oExpimGlobal.Dialog_Anzeigen = False
    oExpimGlobal.AktionsManager
    oExpimGlobal.EinstellungenSpeichern
    call ClearStatusBarDelayed(StatusBarClearDelay)
    Set oExpimGlobal = Nothing
  end if
  Exit Sub
Fehler:
  Set oExpimGlobal = Nothing
  FehlerNachricht "mdlUserInterface.Import_Trassenkoo()"
End Sub


Sub Import_CSV(Optional ByVal ParamDateiName As String = "")
  'Import einer CSV-Datei in eine neue Arbeitsmappe.
  'Parameter: ParamDateiName = Pfad\Name der Eingabedatei.
  '=> Diese Routine wird nur zur Fernsteuerng verwendet.
  On Error GoTo Fehler
  if (not oExpimGlobal is Nothing) then
    Application.Visible        = true
    Application.UserControl    = true
    Application.ScreenUpdating = true
    msgbox "Es ist bereits eine Export / Import - Aktion aktiv. => Eine zweite kann nicht gestartet werden!" , vbOKOnly, "Import CSV-Datei"
  else
    Set oExpimGlobal = New CdatExpim
    If (Err) Then GoTo Fehler
    oExpimGlobal.Quelle_Typ = io_Typ_CsvSpezial
    oExpimGlobal.Quelle_AsciiDatei_Name = ParamDateiName
    oExpimGlobal.Dialog_Anzeigen = False
    oExpimGlobal.AktionsManager
    oExpimGlobal.EinstellungenSpeichern
    call ClearStatusBarDelayed(StatusBarClearDelay)
    Set oExpimGlobal = Nothing
  end if
  Exit Sub
Fehler:
  Set oExpimGlobal = Nothing
  FehlerNachricht "mdlUserInterface.Import_CSV()"
End Sub


Sub ExpimManager(Optional ByVal ParamDateiName As String = "")
  'Aufruf des Import/Export-Managers.
  '=> Diese Routine wird (fast?) nur interaktiv verwendet (Menü GeoTools -> Import / Exports).
  On Error GoTo Fehler
  if (not oExpimGlobal is Nothing) then
    Application.Visible        = true
    Application.UserControl    = true
    Application.ScreenUpdating = true
    msgbox "Es ist bereits eine Export / Import - Aktion aktiv. => Eine zweite kann nicht gestartet werden!" , vbOKOnly, "Export / Import allgemein"
  else
    Set oExpimGlobal = New CdatExpim
    If (Err) Then GoTo Fehler
    oExpimGlobal.EinstellungenWiederherstellen
    if (ParamDateiName <> "") then oExpimGlobal.Quelle_AsciiDatei_Name = ParamDateiName
    oExpimGlobal.AktionsManager
    oExpimGlobal.EinstellungenSpeichern
    call ClearStatusBarDelayed(StatusBarClearDelay)
    Set oExpimGlobal = Nothing
  end if
  Exit Sub
Fehler:
  Set oExpimGlobal = Nothing
  FehlerNachricht "mdlUserInterface.ExpimManager()"
End Sub

Sub Protokoll()
  On Error GoTo Fehler
  ErrMessage = "Protokoll-Konsole existiert nicht!"
  oConsole.Show vbModeless
  ErrMessage = ""
  Exit Sub
Fehler:
  FehlerNachricht "mdlUserInterface.ZeigeProtokoll()"
End Sub


Sub Hilfe_Komplett()
  'Anzeige der Programmdokumentation.
  'Dateiname: <Vorname des AddIn>.pdf
  'Pfad: Eine Verzeichnisebene über der des AddIn
  On Error GoTo Fehler
  Dim hlp As String
  hlp = Verz(ThisWorkbook.Path) & "\" & VorName(ThisWorkbook.Name) & ".chm"
  oSysTools.StarteDatei hlp
  Exit Sub
Fehler:
  FehlerNachricht "mdlUserInterface.Hilfe_Komplett()"
End Sub


Sub GeoTools_Info()
  'Anzeige von Versions- und Lizenzinformationen.
  On Error GoTo Fehler
  Dim Titel       As String
  Dim Meldung     As String
  Titel = "Info über " & ProgName
  Meldung = ProgName & ": Excel-Werkzeuge (nicht nur) für Geodäten." & vbLf & vbLf & _
            "Version"   & vbTab & vbTab & VersionNr & "  (" & VersionDate & ")" & vbLf & vbLf & _
            "Lizenz"    & vbTab & vbTab & "GNU General Public License (GPL)" & vbLf & _
            "Copyright" & vbTab & vbTab & Copyright & "  (" & eMail & ")"
  Call MsgBox(Meldung, vbOKOnly, Titel)
  Exit Sub
Fehler:
  FehlerNachricht "mdlUserInterface.GeoTools_Info()"
End Sub


Sub InfoKeineKonfig()
  'Anzeige von Informationen zum Fehlen der Konfigurationsdatei.
  On Error GoTo Fehler
  Dim Titel       As String
  Dim Meldung     As String
  Dim cfg         As String
  cfg = Verz(ThisWorkbook.Path) & "\" & VorName(ThisWorkbook.Name) & "_cfg.xls"
  Titel = "Keine Konfiguration für " & ProgName & " verfügbar."
  Meldung = "Konfigurationsdatei '" & cfg & "' wurde beim Start nicht gelesen." & vbLf & vbLf & _
            "Mögliche Ursachen:." & vbLf & _
            "  1. Die Datei existiert nicht." & vbLf & _
            "  2. Excel wurde ferngesteuert gestartet." & vbLf & vbLf & _
            "==> Die Funktionalität des Programmes steht dadurch nur eingeschränkt zur Verfügung."
  Call MsgBox(Meldung, vbExclamation, Titel)
  Exit Sub
Fehler:
  FehlerNachricht "mdlUserInterface.InfoKeineKonfig()"
End Sub


'für jEdit:  :folding=indent::collapseFolds=1:
