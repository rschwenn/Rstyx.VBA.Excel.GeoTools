VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wbk_GeoTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2003 - 2014  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'====================================================================================
'Arbeitsmappe wbk_GeoTools
'====================================================================================
'Stellt die Ereignisprozeduren des Add-in zur Verfügung.
'Übernimmt die Initialisierung und die Terminierung des Add-in.


Option Explicit

'Benachrichtigung über Zurücksetzen des Projektes durch VBE
'... scheint in Office 2003 an anderer Stelle gespeichert zu sein
'    und trat bisher nicht in Erscheinung.
Dim VBE_Notify_StatusSave   As Long
Const VBE_Notify_RegValue   As String = "HKEY_CURRENT_USER\Software\Microsoft\VBA\Office\NotifyUserBeforeStateLoss"
Const VBE_Notify_RegTyp     As String = "REG_DWORD"
'



Private Sub Workbook_Open()
  'Initialisierung des Add-Ins ("Workbook_AddinInstall()" funktioniert nicht).
  
  On Error GoTo Fehler
  
  '----------------------------------------------------------------------------------
  '1. Protokoll-Konsole starten.
  On Error Resume next
  ErrMessage = "Fehler beim Initialisieren der Protokoll-Konsole"
  Set oConsole = GetNewConsole(ProgName & " Protokoll")
  On Error GoTo Fehler
  
  '2. RegExp-Objekt initialisieren.
  '   ==> Dies ist gleichzeitig die Prüfung, ob Windows Scripting v5.6 korrekt installiert ist.
  ErrMessage = "Windows Scripting v5.6 ist nicht korrekt installiert, wird aber benötigt!" & vbNewLine & vbNewLine & _
               "==> Einige Funktionen werden nicht funktionieren!"
  Set oRegExp = New VBScript_RegExp_55.RegExp
  
  ErrMessage = ""
  
  '----------------------------------------------------------------------------------
  '3. diverse Menüs erzeugen
  '   ==> muß vor Initialisierung von "CtabAktiveTabelle" erfolgen, damit das
  '       Setzen dessen Eigenschaften Auswirkungen auf die Buttons hat.
  'Call MenuesErzeugen
  
  
  '----------------------------------------------------------------------------------
  '4. Initialisieren von TabellenInfo, Konfiguration, Systemwerkzeuge, Metadaten.
  
  'ErrMessage = "Fehler beim Initialisiern der TabellenInfo"
  Set oAktiveTabelle = New CtabAktiveTabelle
  
  ErrMessage = "Fehler beim Initialisiern der Konfiguration"
  Set oKonfig = New CdatKonfig
  
  'Systemwerkzeuge (abhängig von oKonfig!!!)
  ErrMessage = "Fehler beim Initialisiern der Systemwerkzeuge"
  Set oSysTools = New CToolsSystem
  
  ErrMessage = "Fehler beim Initialisiern der Metadaten"
  Set oMetadaten = New CdatMetaDaten
  
  ErrMessage = ""
  
  
  '----------------------------------------------------------------------------------
  '6. Arbeitsverzeichnis setzen und Schlussmeldung
  Call SetArbeitsverzeichnis
  DebugEcho "GeoTools " & VersionNr & ": Initialisierung beendet." & vbnewline & String(80, "=") & vbnewline
  Application.StatusBar = "GeoTools " & VersionNr & " aktiviert"
  If (Not oKonfig.EinstellungenGelesen) Then Application.StatusBar = Application.StatusBar & " (Keine Konfiguration verfügbar!)"
  If (oConsole is Nothing) Then Application.StatusBar = Application.StatusBar & " (Keine Protokollierung verfügbar!)"
  call ClearStatusBarDelayed(StatusBarClearDelay)
  
  
  '----------------------------------------------------------------------------------
  '7. Versuch, den Fokus im Excel-Fenster zu behalten.
  '   (Ansonsten verschwindet dieser - nach der Initialisierung! - im Nirvana)
  oConsole.Show vbModeless
  oConsole.Hide
  
  
  Exit Sub

Fehler:
  FehlerNachricht "Workbook_Open()"
  Set oRegExp = Nothing
  Set oKonfig = Nothing
  Set oMetadaten = Nothing
  Set oAktiveTabelle = Nothing
  Set oSysTools = Nothing
  Set oConsole = Nothing
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
  'MsgBox "Workbook_BeforeClose()"
  Call MenuesEntfernen()
  Set oRegExp = Nothing
  Set oKonfig = Nothing
  Set oMetadaten = Nothing
  Set oAktiveTabelle = Nothing
  Set oSysTools = Nothing
  Set oConsole = Nothing
End Sub


'für jEdit:  :folding=indent::collapseFolds=1:
