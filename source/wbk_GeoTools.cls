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
' Copyright © 2003 - 2020  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'====================================================================================
' Arbeitsmappe wbk_GeoTools
'====================================================================================
' Übernimmt die Initialisierung und die Terminierung des Add-in.
' Stellt globale Objekte als Eigenschaften zur Verfügung.

Option Explicit


Private InitMsg                 As String
Private strVBAHostNameVersion   As String
Private strAddInName            As String
Private strLoggingInfo          As String
Private LoggingAddInPath        As String
Private LoggingAddInLoaded      As Boolean
Private oAktiveTabelle          As CtabAktiveTabelle
Private oKonfig                 As CdatKonfig
Private oMetadaten              As CdatMetaDaten
Private oRegExp                 As VBScript_RegExp_55.RegExp
Private oSysTools               As CToolsSystem


Private Sub Workbook_Open()
    'Initialisierung des Add-Ins ("Workbook_AddinInstall()" funktioniert nicht).
    On Error GoTo Fehler
    
    InitMsg = "GeoTools " & VersionNr & " aktiviert"
    
    ' AddIn mit LoggingConsole laden (Actions.NET).
    LoggingAddInPath   = ThisWorkbook.Path & "\" & ResourcesSubFolder & "\" & LoggingAddInName
    LoggingAddInLoaded = Application.RegisterXLL(LoggingAddInPath)
    If (LoggingAddInLoaded) Then
        DebugEcho "GeoTools-Protokollierung erfolgreich initialisiert (" & LoggingAddInPath & " geladen und registriert)."
        strLoggingInfo = "AddIn '" & LoggingAddInPath & "' wurde erfolgreich geladen." & vbLf & vbLf & _
                         "==> Protokollierung sollte zur Verfügung stehen."
    Else
        ' Protokollierung geht ins Leere, macht also keinen Sinn.
        strLoggingInfo = "AddIn '" & LoggingAddInPath & "' konnte nicht geladen werden." & vbLf & vbLf & _
                         "Mögliche Ursachen:." & vbLf & _
                         "  1. Die Datei existiert nicht." & vbLf & _
                         "  2. Excel wurde ferngesteuert gestartet." & vbLf & vbLf & _
                         "==> Protokollierung steht dadurch nicht zur Verfügung."
        InitMsg = InitMsg & "  (Keine Protokollierung verfügbar)"
    End If
    
    ' Arbeitsverzeichnis setzen und Schlussmeldung
    Call SetArbeitsverzeichnis
    DebugEcho "GeoTools " & VersionNr & ": Initialisierung beendet." & vbnewline & String(80, "=") & vbnewline
    
    
    ' Konfiguration initialisieren (und Mißerfog in Statuszeile dokumentieren)
    If (Not Me.Konfig.KonfigDateiGelesen) Then
        InitMsg = InitMsg & "  (Keine Konfiguration verfügbar!)"
    End If
    Application.StatusBar = InitMsg
    call ClearStatusBarDelayed(StatusBarClearDelay)
    
    Exit Sub
    
Fehler:
    FehlerNachricht "Workbook_Open()"
End Sub

'Private Sub Workbook_BeforeClose(Cancel As Boolean)
    
    ' *******************************************************************************
    ' This routine causes Excel to hang if GeoTools.xlam is closed from living Excel. 
    ' *******************************************************************************
    
    'MsgBox "Workbook_BeforeClose()"
    
    'Set oRegExp = Nothing
    'Set oKonfig = Nothing
    'Set oMetadaten = Nothing
    'Set oAktiveTabelle = Nothing
    'Set oSysTools = Nothing
'End Sub


' Region "Add-In Umgebung"
    
    '(Toleranz gegenüber zwischenzeitlichem Abbruch des Add-Ins)
    
    Public Property Get VBAHostNameVersion() As String
        If (strVBAHostNameVersion = "") Then
            strVBAHostNameVersion = Application.Name & " " & Application.Version
        End If
        VBAHostNameVersion = strVBAHostNameVersion
    End Property
    
    Public Property Get AddInName() As String
        If (strAddInName = "") Then
            strAddInName = ThisWorkbook.Name
        End If
        AddInName = strAddInName
    End Property
    
    Public Property Get LoggingInfo() As String
        LoggingInfo = strLoggingInfo
    End Property
    
'

' Region "Add-In-weit verfügbare Objekte"
    
    '(Toleranz gegenüber zwischenzeitlichem Abbruch des Add-Ins)
    
    Public Property Get AktiveTabelle() As CtabAktiveTabelle
        If (oAktiveTabelle Is Nothing) Then
            Set oAktiveTabelle = New CtabAktiveTabelle
        End If
        Set AktiveTabelle = oAktiveTabelle
    End Function
    
    Public Property Get Konfig() As CdatKonfig
        If (oKonfig Is Nothing) Then
            
            ' Bringt nichts:
            'Dim StatusVisible As Boolean
            'StatusVisible = Application.Visible
            'SetApplicationVisible(False)
            
            Set oKonfig = New CdatKonfig
            oKonfig.LeseKonfiguration
            
            'SetApplicationVisible(StatusVisible)
            If (Not oKonfig.KonfigDateiGelesen) Then
                WarnEcho "Keine Konfiguration verfügbar (Konfigurationsdatei nicht gelesen)."
                Application.StatusBar = Application.StatusBar & " (Keine Konfiguration verfügbar!)"
                Call ClearStatusBarDelayed(StatusBarClearDelay)
            End If
        End If
        Set Konfig = oKonfig
    End Function
    
    Public Property Get Metadaten() As CdatMetaDaten
        If (oMetadaten Is Nothing) Then
            On Error GoTo Fehler
            ErrMessage = "Fehler beim Initialisiern der Metadaten"
            Set oMetadaten = New CdatMetaDaten
        End If
        Set Metadaten = oMetadaten
        
        Exit Sub
        Fehler:
        FehlerNachricht "Metadaten()"
    End Function
    
    Public Property Get RegExp() As VBScript_RegExp_55.RegExp
        If (oRegExp Is Nothing) Then
            On Error GoTo Fehler
            ErrMessage = "Windows Scripting ist nicht korrekt installiert, wird aber benötigt!" & vbNewLine & vbNewLine & _
                         "==> Einige Funktionen werden nicht funktionieren!"
            Set oRegExp = New VBScript_RegExp_55.RegExp
        End If
        Set RegExp = oRegExp
        
        Exit Sub
        Fehler:
        FehlerNachricht "RegExp()"
    End Function
    
    Public Property Get SysTools() As CToolsSystem
        If (oSysTools Is Nothing) Then
            On Error GoTo Fehler
            'Systemwerkzeuge (abhängig von oKonfig!!!)
            ErrMessage = "Fehler beim Initialisiern der Systemwerkzeuge"
            Set oSysTools = New CToolsSystem
        End If
        Set SysTools = oSysTools
        
        Exit Sub
        Fehler:
        FehlerNachricht "SysTools()"
    End Function
    
' End Region


' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
