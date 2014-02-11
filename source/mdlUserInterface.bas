Attribute VB_Name = "mdlUserInterface"
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2003 - 2014  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
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



Sub Erzeuge_InfoKeineKonfig()
 '  'Erzeugt Button im GeoTools-Hauptmenü "==> Keine Konfiguration verfügbar!"
 '  'Wird aufgerufen von CdatKonfig\LeseKonfiguration().
 '  
 '  Dim cbc_M1_Tools    As CommandBarControl
 '  Dim cbb             As CommandBarButton
 '  Dim sTag            As String
 '  
 '  'Hauptmenu "GeoTools" finden.
 '  On Error Resume Next
 '  sTag = PrefixHauptmenue & TagHauptmenu_GeoTools
 '  Set cbc_M1_Tools = Application.CommandBars.FindControl(Tag:=sTag, Type:=msoControlPopup)
 '  
 '  'Neuen Menüpunkt einrichten.
 '  If ((Not (cbc_M1_Tools Is Nothing)) And (Not (Err))) Then
 '    'On Error GoTo 0
 '    On Error GoTo Fehler
 '    Set cbb = cbc_M1_Tools.Controls.Add(Type:=msoControlButton, Temporary:=True)
 '    cbb.Caption = "==> Keine &Konfiguration verfügbar!"
 '    cbb.OnAction = TagInfoKeineKonfig
 '    'cbIcons.Controls.Item(TagHilfe_GeoTools).CopyFace
 '    'cbb.PasteFace
 '    cbb.Tag = PrefixHauptmenue & TagInfoKeineKonfig
 '    cbb.BeginGroup = True
 '  End If
 '  
 '  Exit Sub
 'Fehler:
 ' FehlerNachricht "mdlUserInterface.Erzeuge_InfoKeineKonfig()"
End Sub


Sub MenuesEntfernen()
  ' Altlasten aus Vorgängerversionen entfernen.
  'Wird aufgerufen von wbk_GeoTools\Workbook_BeforeClose().
  On Error Resume Next
  'Einträge im Kontextmenü entfernen
  Application.CommandBars("cell").Controls("Datenbereich formatieren").Delete
  Application.CommandBars("cell").Controls("Bedingte Formatierung...").Delete
  Application.CommandBars("cell").Controls("Datei öffnen (Name in Zelle)").Delete
  Application.CommandBars("gtDummy_Icons").Delete
  On Error Goto 0
End Sub


Sub SetSilent_AktiveTabelle(inpSilent As Boolean)
  'Setzt den aktuellen Modus für "Silent" im Objekt oAktiveTabelle.
  On Error GoTo 0
  oAktiveTabelle.Silent = inpSilent
End Sub

' TODO: entfernen!
Sub FormatDatenMitStreifen()
  'Reaktion auf Buttonklick "FormatDatenMitStreifen"
  'Änderung von Tooltip und Status des Buttons übernimmt "Property Let FormatDatenMitStreifen"
  oAktiveTabelle.FormatDatenMitStreifen = Not oAktiveTabelle.FormatDatenMitStreifen
End Sub


' TODO: entfernen!
Sub FormatDatenOhneFuellung()
  'Reaktion auf Buttonklick "FormatDatenOhneFuellung"
  'Änderung von Tooltip und Status des Buttons übernimmt "Property Let FormatDatenOhneFuellung"
  oAktiveTabelle.FormatDatenOhneFuellung = Not oAktiveTabelle.FormatDatenOhneFuellung
End Sub


' TODO: entfernen!
Sub FormatDatenNKStellenSetzen()
  'Reaktion auf Buttonklick "FormatDatenNKStellenSetzen"
  'Änderung von Tooltip und Status des Buttons übernimmt "Property Let FormatDatenNKStellenSetzen"
  oAktiveTabelle.FormatDatenNKStellenSetzen = Not oAktiveTabelle.FormatDatenNKStellenSetzen
End Sub


' TODO: entfernen!
Sub FormatDatenNKStellenAnzahl()
  ''Reaktion auf Auswahl in der Combobox "FormatDatenNKStellenAnzahl"
  'On Error Resume Next
  'Dim cbcb As CommandBarComboBox
  'Set cbcb = CommandBars.FindControl(Type:=msoControlDropdown, Tag:=PrefixToolbox & TagFormatDatenNKStellenAnzahl)
  'If ((Not (cbcb Is Nothing)) And (Not (Err))) Then
  '  oAktiveTabelle.FormatDatenNKStellenAnzahl = CInt(cbcb.text)
  'End If
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
            "Lizenz"    & vbTab & vbTab & "The MIT License" & vbLf & _
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
