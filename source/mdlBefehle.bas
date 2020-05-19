Attribute VB_Name = "mdlBefehle"
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) f�r Geod�ten.
' Copyright � 2003 - 2020  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'====================================================================================
' Modul mdlBefehle
'====================================================================================
' Stellt die Befehle des Add-Ins zur Verf�gung.
' Diese werden i.d.R. von Ribbon-Callbacks oder per Fernsteuerung aufgerufen.

Option Explicit


Sub SetSilent_AktiveTabelle(inpSilent As Boolean)
  ' Dies wird in VB-Skripten verwendet, um Fehlermeldungen zu unterdr�cken,
  ' die auftreten, wenn z.B. eine GeoTools-Formatierung ausgel�st wird, ohne
  ' vorher zu pr�fen, ob die Tabelle daf�r vorbereitet ist.
  On Error Resume Next
  ThisWorkbook.AktiveTabelle.Silent = inpSilent
  On Error Goto 0
End Sub


Sub SchreibeProjektDaten()
  'Schreibt von allen verf�gbaren Projektdaten diejenigen in die aktive Tabelle,
  'f�r die entsprechend benannte Zellen existieren.
  On Error GoTo Fehler
  ThisWorkbook.Metadaten.Update oPrjLocal:=nothing, oExtraLocal:=nothing
  ThisWorkbook.AktiveTabelle.SchreibeMetaDaten
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlBefehle.SchreibeProjektDaten()"
End Sub


Sub SchreibeFusszeile_1()
  'Schreibt die Fusszeile_1 in die aktive Tabelle.
  On Error GoTo Fehler
  ThisWorkbook.AktiveTabelle.SchreibeFusszeile_1
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlBefehle.SchreibeFusszeile_1()"
End Sub


Sub LoeschenDaten()
  'Alle Datenzeilen der aktiven Tabelle werden gel�scht.
  Dim Titel    As String
  Dim Message  As String
  Dim Buttons  As Integer
  On Error GoTo Fehler
  Message = "Soll der gesamte Datenbereich der Tabelle wirklich gel�scht werden? " & vbNewLine & vbNewLine & _
            "==> Diese Aktion kann NICHT r�ckg�ngig gemacht werden!!!"
  Buttons = vbYesNo + vbQuestion + vbDefaultButton2
  Titel = "Datenbereich l�schen!"
  If (MsgBox(Message, Buttons, Titel) = vbYes) Then
    ThisWorkbook.AktiveTabelle.LoeschenDaten
    call ClearStatusBarDelayed(StatusBarClearDelay)
  End If
  Exit Sub
Fehler:
  FehlerNachricht "mdlBefehle.LoeschenDaten()"
End Sub


Sub FormatDaten()
  '�bertr�gt das Format des "InfoTraegers" auf alle weiteren Datenzeilen.
  On Error GoTo Fehler
  ThisWorkbook.AktiveTabelle.FormatDaten
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlBefehle.FormatDaten()"
End Sub


Sub UebertragenFormeln()
  '�bertr�gt die Formeln des 'Formel'-Bereiches auf alle weiteren Datenzeilen.
  On Error GoTo Fehler
  ThisWorkbook.AktiveTabelle.UebertragenFormeln
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlBefehle.UebertragenFormeln()"
End Sub


Sub Mod_FehlerVerbesserung()
  'Modifiziert Daten der aktiven Tabelle.
  On Error GoTo Fehler
  ThisWorkbook.AktiveTabelle.Mod_FehlerVerbesserung
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlBefehle.Mod_FehlerVerbesserung()"
End Sub


Sub Mod_UeberhoehungAusBemerkung()
  'Modifiziert Daten der aktiven Tabelle.
  On Error GoTo Fehler
  ThisWorkbook.AktiveTabelle.Mod_UeberhoehungAusBemerkung
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlBefehle.Mod_UeberhoehungAusBemerkung()"
End Sub


Sub Mod_Transfo_Tk2Gls()
  'Modifiziert Daten der aktiven Tabelle.
  On Error GoTo Fehler
  ThisWorkbook.AktiveTabelle.Mod_Transfo_Tk2Gls
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlBefehle.Mod_Transfo_Tk2Gls()"
End Sub


Sub Mod_Transfo_Gls2Tk()
  'Modifiziert Daten der aktiven Tabelle.
  On Error GoTo Fehler
  ThisWorkbook.AktiveTabelle.Mod_Transfo_Gls2Tk
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlBefehle.Mod_Transfo_Gls2Tk()"
End Sub


Sub TabellenStruktur()
  'Dialog "Tabellenstruktur und Spaltennamen verwalten" anzeigen.
  Dim Dialog    As frmSpaltenVerw
  Set Dialog = New frmSpaltenVerw
  Dialog.Show vbModeless
  Set Dialog = Nothing
  Exit Sub
Fehler:
  FehlerNachricht "mdlBefehle.TabellenStruktur()"
End Sub


Sub Selection2Interpolationsformel()
  'Aufgrund der aktuellen Zellauswahl wird eine Interpolationsformel erstellt.
  On Error GoTo Fehler
  ThisWorkbook.AktiveTabelle.Selection2Interpolationsformel
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlBefehle.Selection2Interpolationsformel()"
End Sub


Sub Selection2MarkDoppelteWerte()
  'Die markierten (und alle darunter liegenden) Zellen werden mit einer bedingten
  'Formatierung versehen, die alle Zellen mit solchen Werten hervorhebt, die in
  'derselben Spalte mehr als einmal vorkommen.
  On Error GoTo Fehler
  ThisWorkbook.AktiveTabelle.Selection2MarkDoppelteWerte
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlBefehle.Selection2MarkDoppelteWerte()"
End Sub


Sub InsertLines()
  'Dialog "Leerzeilen einf�gen"
  On Error GoTo Fehler
  Dim oDialog    As frmInsertLines
  Set oDialog = New frmInsertLines
  oDialog.Show
  Exit Sub
Fehler:
  Set oDialog = Nothing
  FehlerNachricht "mdlBefehle.insertLines()"
End Sub


Sub DateiBearbeiten()
  'Die Datei, deren Name dem Inhalt der aktiven Zelle entspricht, wird im
  'konfigurierten Editor geladen. Bei Mi�erfolg wird die Windows-Standardanwendung
  'des Dateityps gestartet.
  On Error GoTo Fehler
  dim Datei
  If (ActiveCell Is Nothing) Then
    Application.StatusBar = "Fehler: Es existiert keine aktive Zelle!"
  else
    Datei = trim(ActiveCell.Value)
    if (not ThisWorkbook.SysTools.isDatei(Datei)) then
      Application.StatusBar = "Der Inhalt der aktiven Zelle ('" & Datei & "') bezeichnet keine existierende Datei !"
    else
      Application.StatusBar = "Datei '" & Datei & "' wird im Editor ge�ffnet."
      if (not ThisWorkbook.SysTools.StartEditor("""" & Datei & """")) then
        Application.StatusBar = "Datei '" & Datei & "' wird mit Standardanwendung ge�ffnet."
        ThisWorkbook.SysTools.StarteDatei(Datei)
      end if
    end if
  end if
  call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
Fehler:
  FehlerNachricht "mdlBefehle.DateiBearbeiten()"
End Sub


Sub Import_Trassenkoo(Optional ByVal ParamDateiName As String = "")
  'Import von Trassenkoordinaten in eine neue Arbeitsmappe.
  'Parameter: ParamDateiName = Pfad\Name der Eingabedatei.
  '=> Diese Routine wird nur zur Fernsteuerng verwendet.
  On Error GoTo Fehler
  if (not oExpimGlobal is Nothing) then
    SetApplicationVisible(true)
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
  FehlerNachricht "mdlBefehle.Import_Trassenkoo()"
End Sub


Sub Import_CSV(Optional ByVal ParamDateiName As String = "")
  'Import einer CSV-Datei in eine neue Arbeitsmappe.
  'Parameter: ParamDateiName = Pfad\Name der Eingabedatei.
  '=> Diese Routine wird nur zur Fernsteuerng verwendet.
  On Error GoTo Fehler
  if (not oExpimGlobal is Nothing) then
    SetApplicationVisible(true)
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
  FehlerNachricht "mdlBefehle.Import_CSV()"
End Sub


Sub ExpimManager(Optional ByVal ParamDateiName As String = "")
  'Aufruf des Import/Export-Managers.
  '=> Diese Routine wird (fast?) nur interaktiv verwendet (Men� GeoTools -> Import / Exports).
  On Error GoTo Fehler
  if (not oExpimGlobal is Nothing) then
    SetApplicationVisible(true)
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
  FehlerNachricht "mdlBefehle.ExpimManager()"
End Sub


Sub BatchPDF()
  ' Dialog "BatchPDF" anzeigen.
  Dim Dialog    As frmBatchPDF
  Set Dialog = New frmBatchPDF
  Dialog.Show vbModeless
  Set Dialog = Nothing
  Exit Sub
Fehler:
  FehlerNachricht "mdlBefehle.BatchPDF()"
End Sub


Sub Protokoll()
  Call ShowConsole
End Sub


Sub Hilfe_Komplett()
  'Anzeige der Programmdokumentation.
  'Dateiname: <Vorname des AddIn>.pdf
  'Pfad: Eine Verzeichnisebene �ber der des AddIn
  On Error GoTo Fehler
  Dim hlp As String
  'hlp = Verz(ThisWorkbook.Path) & "\" & VorName(ThisWorkbook.Name) & ".chm"
  hlp = ThisWorkbook.Path & "\" & ResourcesSubFolder & "\" & VorName(ThisWorkbook.Name) & ".chm"
  ThisWorkbook.SysTools.StarteDatei hlp
  Exit Sub
Fehler:
  FehlerNachricht "mdlBefehle.Hilfe_Komplett()"
End Sub


Sub GeoTools_Info()
  'Anzeige von Versions- und Lizenzinformationen.
  'Wiederbelebung der Statusleiste (durch Zur�cksetzen).
  On Error GoTo Fehler
  Dim Titel       As String
  Dim Meldung     As String
  Titel = "Info �ber " & ProgName
  Meldung = ProgName & ": Excel-Werkzeuge (nicht nur) f�r Geod�ten." & vbLf & vbLf & _
            "Version"   & vbTab & vbTab & VersionNr & "  (" & VersionDate & ")" & vbLf & vbLf & _
            "Lizenz"    & vbTab & vbTab & "The MIT License" & vbLf & _
            "Copyright" & vbTab & vbTab & Copyright & "  (" & eMail & ")"
  Application.StatusBar = ProgName & " " & VersionNr
  Call MsgBox(Meldung, vbOKOnly, Titel)
  Application.StatusBar = ""
  
  Exit Sub
Fehler:
  FehlerNachricht "mdlBefehle.GeoTools_Info()"
End Sub


Sub InfoKeineKonfig()
  'Anzeige von Informationen zum Fehlen der Konfigurationsdatei.
  On Error GoTo Fehler
  Dim Titel As String
  Titel = "Keine Konfiguration f�r " & ProgName & " verf�gbar."
  Call MsgBox(ThisWorkbook.Konfig.InfoKeineKonfig, vbExclamation, Titel)
  Exit Sub
Fehler:
  FehlerNachricht "mdlBefehle.InfoKeineKonfig()"
End Sub


'f�r jEdit:  :folding=indent::collapseFolds=1:
