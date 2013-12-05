VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLoggingConsole 
   Caption         =   "Console"
   ClientHeight    =   7632
   ClientLeft      =   3036
   ClientTop       =   2376
   ClientWidth     =   11928
   OleObjectBlob   =   "frmLoggingConsole.frx":0000
End
Attribute VB_Name = "frmLoggingConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================================================================
' Module:       frmLoggingConsole (VBA Form)
'               
' Copyright:    © 2008-2013  Robert Schwenn, devel@rstyx.de
'               
' License:      The MIT License (see http://opensource.org/licenses/mit-license.html)
'               
' History:      24.03.2008  1.0.0  Initial release (only Excel)
'               23.03.2009  1.1.0  - Dialog can be resized.
'                                  - Settings are stored and restored via registry.
'                                  - VBA host independent
'               25.04.2009  1.1.1  Bugfix: When the console should be shown because of
'                                  logging an error, an error occured when a modeless
'                                  dialog has been shown already.
'               17.05.2009  1.2.0  Performance improved by using ListView instead of textbox.
'               02.10.2010  1.3.0  added properties, which indicate unread mesages:
'                                  - newErrors
'                                  - newWarnings
'                                  - newInfos
'               12.03.2011  1.3.1  Mail address changed
'               05.12.2013  1.5.0  License changed to MIT
'
' Intention:    Simple and straightforeward creation, managing and viewing
'               of small upto medium sized Logs for:
'                 1. presenting program results to the user
'                 2. debugging purposes
'
' Features:    - 4 Log Levels: Error, Warning, Info, Debug
'              - Always 4 Logs at the same time: from "Error" (only contains Errors)
'                                                to   "Debug" (contains all levels)
'              - Every Log: - is managed in a ListView control of the dialog
'                           - can be cleared at every time
'                           - can be saved to a file at every time
'              - Some typical informations can be shown for each log message optionally:
'                date, time, log level, message source
'              - The maximum length of a Log is configurable.
'              - Optionally, the console is shown automatically, when an error or warning is logged.
'              - Most settings and actions can be invoked programatically and interactive,
'                even when the dialog is shown (if it's shown non-modal).
'              - Settings are stored and restored via registry separated for Me.WindowTitle:
'                - restore: when the dialog is loaded and when the Me.WindowTitle was changed
'                - store:   when the dialog is unloaded
'              - GUI languages: English and German
'
' VBA Hosts    - The LoggingConsole (hopefully) should run in every host application.
'                However, it relies on the existence of two properties of the application
'                object: Application.Name and Application.Version.
'                It was tested in: Excel 2003, Word 2003, Microstation V8, XM, V8i
'
' Dependencies: - Microsoft Scripting Runtime
'               - Windows Script Host Object Model
'               - Microsoft Windows Common Controls 6.0 (MSComctlLib.ListView)
'
' Notes:        - Performance is poor when the Logs are growing.
'
' Example:      - Instanciate the object:  Set oConsole = New frmLoggingConsole
'                                          Set oConsole = LoggingConsole.GetNewConsole(title, source)
'               - Change window title:     oConsole.WindowTitle = "My Console"
'               - Change a setting:        oConsole.IncludeDate = true
'               
'               - Log an Error message:    oConsole.logError   "Error text"
'               - Log an Warning message:  oConsole.logWarning "Warning text"
'               - Log an Info message:     oConsole.logInfo    "Info text"
'               - Log an Debug message:    oConsole.logDebug   "Debug text"
'               
'               - Show the Log:            oConsole.Show vbModeless
'               - Save a Log to a file:    oConsole.saveLog oConsole.PAGE_LOG, "C:\testinfo.log"
'               
'               - See LogDemo.bas for further examples.
'               - Look for all properties and public Methods here in this module.
'==================================================================================================
Option Explicit
Private Const INFO_VERSION  As String = "1.5.0"
Private Const INFO_COPYLEFT As String = "© 2008-2013  Robert Schwenn"
Private Const INFO_LICENCE  As String = "The MIT License"
Private Const INFO_MAIL     As String = "devel@rstyx.de"


'***  Default settings  ********************************************************
Private Const DEFAULT_INCLUDECOUNTER As Boolean = False
Private Const DEFAULT_INCLUDEDATE    As Boolean = False
Private Const DEFAULT_INCLUDETIME    As Boolean = True
Private Const DEFAULT_INCLUDELEVEL   As Boolean = True
Private Const DEFAULT_SHOWONERROR    As Boolean = False
Private Const DEFAULT_LOGSOURCE      As String = ""
Private Const DEFAULT_MAXLOGLENGTH   As Long = 5000       'Line Count
Private Const DEFAULT_ACTIVELOGPAGE  As Integer = 2       ' = PAGE_LOG

Private Const DEFAULT_DIALOG_LEFT    As Integer = 150
Private Const DEFAULT_DIALOG_TOP     As Integer = 100
Private Const DEFAULT_DIALOG_WIDTH   As Integer = 600     'default = minimum
Private Const DEFAULT_DIALOG_HEIGHT  As Integer = 400     'default = minimum
Private Const DEFAULT_COLUMN_WIDTH   As Integer = 75
'*******************************************************************************


'***  Other settings  **********************************************************
Private Const DIALOG_MAX_WIDTH      As Integer = 2400
Private Const DIALOG_MAX_HEIGHT     As Integer = 1800

Private Const RESTORE_ACTIVELOGPAGE As Boolean = False   'Should the last active page become active again at init?
'*******************************************************************************


'Declarations for getFileNameFromDialog()
  Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
  Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
  
  Private Type OPENFILENAME
    lStructSize        As Long
    hwndOwner          As Long
    hInstance          As Long
    lpstrFilter        As String
    lpstrCustomFilter  As String
    nMaxCustFilter     As Long
    nFilterIndex       As Long
    lpstrFile          As String
    nMaxFile           As Long
    lpstrFileTitle     As String
    nMaxFileTitle      As Long
    lpstrInitialDir    As String
    lpstrTitle         As String
    Flags              As Long
    nFileOffset        As Integer
    nFileExtension     As Integer
    lpstrDefExt        As String
    lCustData          As Long
    lpfnHook           As Long
    lpTemplateName     As String
  End Type

'Constants
  Private Const LOGLEVEL_ERROR       As Long = 4
  Private Const LOGLEVEL_WARNING     As Long = 3
  Private Const LOGLEVEL_INFO        As Long = 2
  Private Const LOGLEVEL_DEBUG       As Long = 1
  
  Private Const COL_IDX_COUNTER      As Long = 0
  Private Const COL_IDX_DATE         As Long = 1
  Private Const COL_IDX_TIME         As Long = 2
  Private Const COL_IDX_LEVEL        As Long = 3
  Private Const COL_IDX_SOURCE       As Long = 4
  Private Const COL_IDX_MESSAGE      As Long = 5
  
  Private Const COL_KEY_COUNTER      As String = "Counter"
  Private Const COL_KEY_DATE         As String = "Date"
  Private Const COL_KEY_TIME         As String = "Time"
  Private Const COL_KEY_LEVEL        As String = "Level"
  Private Const COL_KEY_SOURCE       As String = "LogSource"
  Private Const COL_KEY_MESSAGE      As String = "Message"
  
  Private Const REGKEY_CLASS_ROOT    As String = "HKCU\Software\VB and VBA Program Settings\LoggingConsole\"    'Registry root key for settings of all implementations of Logging Console
  Private Const VBA_HOST_NAME_EXCEL  As String = "Microsoft Excel"
  Private Const VBA_HOST_NAME_WORD   As String = "Microsoft Word"
  
  
'Declarations
  Public PAGE_ERRORLOG           As Integer
  Public PAGE_WARNINGLOG         As Integer
  Public PAGE_LOG                As Integer
  Public PAGE_DEBUGLOG           As Integer
  
  Private oLogs                  As Scripting.Dictionary
  Private oMsgTypeName           As Scripting.Dictionary
  Private oLabels                As Scripting.Dictionary
  
  Private VBAHostName            As String
  Private VBAHostVersion         As String
  Private VBAHostDisplayName     As String
  Private VBAHostDisplayVersion  As String
  Private LangID                 As Long
  
  Private TypeNameError          As String
  Private TypeNameWarning        As String
  Private TypeNameMessage        As String
  Private TypeNameDebug          As String
  
  Private minWidthPoints         As Integer
  Private minHeightPoints        As Integer
  Private maxWidthPoints         As Integer
  Private maxHeightPoints        As Integer
  Private maxDeltaWidthPoints    As Integer
  Private maxDeltaHeightPoints   As Integer
  Private maxFactor              As Double
  Private SpinFactor             As Double
  
  Private dW_MultiPage           As Integer
  Private dH_MultiPage           As Integer
  Private dW_lvwLog              As Integer
  Private dH_lvwLog              As Integer
  
  Private dW_frmOptions          As Integer
  
  Private dW_btnHide             As Integer
  Private dH_btnHide             As Integer
  Private dH_btnSave             As Integer
  Private dH_btnClear            As Integer
  Private dH_btnClearAll         As Integer
  
  Private REGKEY_ROOT            As String
  Private strLogSource           As String
  Private OS_Win_ScrollbarWidth  As String
  
  Private blnNewErrors          As Boolean
  Private blnNewWarnings        As Boolean
  Private blnNewInfos           As Boolean
'


Private Sub UserForm_Initialize()
  'Initializes this form
  
  On Error GoTo ErrorHandler
  Const RegKey_SysWinMetrics = "HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics\"   'Window system settings
  Dim maxFactorWidth
  Dim maxFactorHeight
  Dim WshShell        As New IWshRuntimeLibrary.WshShell
  
  'Default strings - needed while initializing
    TypeNameError = "[Error]  "
    TypeNameWarning = "[Warning]"
    TypeNameMessage = "[Info]   "
    TypeNameDebug = "[Debug]  "
    
    Set oMsgTypeName = New Scripting.Dictionary
    oMsgTypeName.Add LOGLEVEL_ERROR, TypeNameError
    oMsgTypeName.Add LOGLEVEL_WARNING, TypeNameWarning
    oMsgTypeName.Add LOGLEVEL_INFO, TypeNameMessage
    oMsgTypeName.Add LOGLEVEL_DEBUG, TypeNameDebug
  
  'Public "Constants"
    'Page indexes for MultiPage Control
    PAGE_ERRORLOG = 0
    PAGE_WARNINGLOG = 1
    PAGE_LOG = 2
    PAGE_DEBUGLOG = 3
    
  'Get System Window Metrics
    OS_Win_ScrollbarWidth = 5
    On Error Resume Next
    OS_Win_ScrollbarWidth = CLng(WshShell.RegRead(RegKey_SysWinMetrics & "ScrollWidth")) / -12
    On Error GoTo ErrorHandler
  '
  Call initListViews
  
  logDebug "UserForm_Initialize(): Initializing Logging Console ..."
  
  'Get host related info's.
    Call getHostEnvironment
  
  'Set localized strings
    Call setLocaleStrings
    
  'Set labels for control elements
    Me.Caption = oLabels("WindowTitle")
    Me.frmProgInfo.Caption = oLabels("frmProgInfo") & " " & INFO_VERSION
    Me.lblProgInfo_1.Caption = oLabels("lblProgInfo_1")
    Me.lblProgInfo_2.Caption = VBAHostDisplayName & "  " & VBAHostDisplayVersion
    
    Me.btnHide.Caption = oLabels("btnHide")
    Me.btnClearAll.Caption = oLabels("btnClearAll")
    Me.btnClear.Caption = oLabels("btnClear")
    Me.btnSave.Caption = oLabels("btnSave")
    Me.btnClearAll.ControlTipText = oLabels("btnClearAll_ToolTip")
    Me.btnClear.ControlTipText = oLabels("btnClear_ToolTip")
    Me.btnSave.ControlTipText = oLabels("btnSave_ToolTip")
    
    Me.frmOptions.Caption = oLabels("frmOptions")
    Me.chkCounter.Caption = oLabels("chkCounter")
    Me.chkDate.Caption = oLabels("chkDate")
    Me.chkTime.Caption = oLabels("chkTime")
    Me.chkLevel.Caption = oLabels("chkLevel")
    Me.chkShowOnError.Caption = oLabels("chkShowOnError")
    Me.chkSource.Caption = oLabels("chkSource")
    Me.lblLimit.Caption = oLabels("lblLimit")
    Me.txtLimit.ControlTipText = oLabels("txtLimit_ToolTip")
    Me.lblSize.Caption = oLabels("lblSize")
    Me.spinSize.ControlTipText = oLabels("spinSize_ToolTip")
    
    Me.MultiPageLogOutput.PageErrorlog.Caption = oLabels("PageErrorlog")
    Me.MultiPageLogOutput.PageWarninglog.Caption = oLabels("PageWarninglog")
    Me.MultiPageLogOutput.PageLog.Caption = oLabels("PageLog")
    Me.MultiPageLogOutput.PageDebuglog.Caption = oLabels("PageDebuglog")
    
    Call setListViewColumnLabels
  
  'Set Log level names
    oMsgTypeName(LOGLEVEL_ERROR) = TypeNameError
    oMsgTypeName(LOGLEVEL_WARNING) = TypeNameWarning
    oMsgTypeName(LOGLEVEL_INFO) = TypeNameMessage
    oMsgTypeName(LOGLEVEL_DEBUG) = TypeNameDebug
    
  'Prepare resizing #1: Get some initial geometry values from original design.
    dW_MultiPage = Me.Width - MultiPageLogOutput.Width
    dH_MultiPage = Me.Height - MultiPageLogOutput.Height
    dW_lvwLog = MultiPageLogOutput.Width - lvwLog.Width
    dH_lvwLog = MultiPageLogOutput.Height - lvwLog.Height
    
    dW_frmOptions = Me.Width - frmOptions.Left
    
    dW_btnHide = Me.Width - btnHide.Left
    dH_btnHide = Me.Height - btnHide.Top
    dH_btnSave = Me.Height - btnSave.Top
    dH_btnClear = Me.Height - btnClear.Top
    dH_btnClearAll = Me.Height - btnClearAll.Top
    
  'Init default status: Apply hard coded default settings - That's save.
    Call resetAllSettings
    
  'Prepare resizing #2: SpinButton related.
    maxWidthPoints = DIALOG_MAX_WIDTH
    maxHeightPoints = DIALOG_MAX_HEIGHT
    spinSize.Min = 0
    spinSize.Max = maxWidthPoints / 35
    
    minWidthPoints = Me.Width
    minHeightPoints = Me.Height
    
    maxDeltaWidthPoints = maxWidthPoints - minWidthPoints
    maxDeltaHeightPoints = maxHeightPoints - minHeightPoints
    
    If ((maxDeltaWidthPoints < 10) Or (maxDeltaHeightPoints < 10)) Then
      spinSize.Enabled = False
    Else
      maxFactorWidth = maxDeltaWidthPoints / minWidthPoints
      maxFactorHeight = maxDeltaHeightPoints / minHeightPoints
      maxFactor = IIf(maxFactorWidth < maxFactorHeight, maxFactorWidth, maxFactorHeight)
    End If
    
  'Restore last status
    Call restoreAllSettings    'Errors are ignored silently.
  '
  logDebug "UserForm_Initialize(): Logging Console initialized." & vbNewLine
  Set WshShell = Nothing
  Exit Sub
  
ErrorHandler:
  logDebug "UserForm_Initialize(): Logging Console NOT initialized!"
  MsgBox oLabels("LoggingFailed")
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  'Prevent the dialog from beeing unloaded by the user (via system menu)
  'MsgBox "UserForm_QueryClose():  CloseMode = " & CloseMode
  If (CloseMode = vbFormControlMenu) Then
    call UserForm_Hide()
    Cancel = True
  End If
End Sub

Private Sub UserForm_Resize()
  'Move and resize some controls, after the dialog has been resized.
  'logDebug "UserForm_Resize(): Move and resize some controls."
  
  'Resize log output areas
    MultiPageLogOutput.Width = Me.Width - dW_MultiPage
    MultiPageLogOutput.Height = Me.Height - dH_MultiPage
    lvwLog.Width = MultiPageLogOutput.Width - dW_lvwLog
    lvwLog.Height = MultiPageLogOutput.Height - dH_lvwLog
    lvwDebugLog.Width = MultiPageLogOutput.Width - dW_lvwLog
    lvwDebugLog.Height = MultiPageLogOutput.Height - dH_lvwLog
    lvwErrorLog.Width = MultiPageLogOutput.Width - dW_lvwLog
    lvwErrorLog.Height = MultiPageLogOutput.Height - dH_lvwLog
    lvwWarningLog.Width = MultiPageLogOutput.Width - dW_lvwLog
    lvwWarningLog.Height = MultiPageLogOutput.Height - dH_lvwLog
    
    Call resizeListViewColumns
    
  'Move options frame
    '=> Do not move, so the SpinButton isarray reachable all the time.
    'frmOptions.Left = Me.Width - dW_frmOptions
    
  'Move buttons
    btnHide.Left = Me.Width - dW_btnHide
    btnHide.Top = Me.Height - dH_btnHide
    btnSave.Top = Me.Height - dH_btnSave
    btnClear.Top = Me.Height - dH_btnClear
    btnClearAll.Top = Me.Height - dH_btnClearAll
End Sub

Private Sub UserForm_Terminate()
  Call storeSettings
  Set oLabels = Nothing
  Set oMsgTypeName = Nothing
End Sub

Private Sub UserForm_Hide()
  'keine echte Ereignisroutine!
  blnNewErrors   = false
  blnNewWarnings = false
  blnNewInfos    = false
  Me.Hide
End Sub


' === Event handlers for controls ==============================================

Private Sub chkCounter_Change()
  logDebug "Set IncludeCounter = '" & CStr(chkCounter.Value) & "'"
  Call setListViewColumnVisibility(COL_KEY_COUNTER, chkCounter.Value)
End Sub

Private Sub chkDate_Change()
  logDebug "Set IncludeDate = '" & CStr(chkDate.Value) & "'"
  Call setListViewColumnVisibility(COL_KEY_DATE, chkDate.Value)
End Sub

Private Sub chkTime_Change()
  logDebug "Set IncludeTime = '" & CStr(chkTime.Value) & "'"
  Call setListViewColumnVisibility(COL_KEY_TIME, chkTime.Value)
End Sub

Private Sub chkLevel_Change()
  logDebug "Set IncludeLevel = '" & CStr(chkLevel.Value) & "'"
  Call setListViewColumnVisibility(COL_KEY_LEVEL, chkLevel.Value)
End Sub

Private Sub chkSource_Change()
  logDebug "Set IncludeSource = '" & CStr(chkSource.Value) & "'"
  Call setListViewColumnVisibility(COL_KEY_SOURCE, chkSource.Value)
End Sub

Private Sub chkShowOnError_Click()
  logDebug "Set ShowOnError = '" & CStr(chkShowOnError.Value) & "'"
End Sub

Private Sub txtLimit_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  logDebug "Set MaxLogLength = '" & txtLimit.Value & "' lines."
End Sub

Private Sub spinSize_Change()
  'Change the dialog size.
  Dim currentFactor   As Double
  Dim SpinFactor      As Double
  
  'Note: The spinButton.BeforeUpdate() event will not occur before spinButton
  '      loses focus! This means, that leaving the dialog via ESC when spinSize
  '      has the focus, a former value will be restored...
  
  'Normalizing the Spinner value to a number between 0 and 1 and get an absolut Magnification factor.
  SpinFactor = (spinSize.Value - spinSize.Min) / (spinSize.Max - spinSize.Min)
  currentFactor = maxFactor * SpinFactor
  'logDebug "spinSize_Change(): SpinFactor = " & Format(SpinFactor, "0.0000")
  
  'Set new dialog size
  Me.Width = Int(minWidthPoints + (currentFactor * minWidthPoints))
  Me.Height = Int(minHeightPoints + (currentFactor * minHeightPoints))
End Sub

Private Sub btnHide_Click()
  call UserForm_Hide()
End Sub

Private Sub btnClear_Click()
  Me.clearLog Me.ActiveLogPage
End Sub

Private Sub btnClearAll_Click()
  Me.clearAllLogs
End Sub

Private Sub btnSave_Click()
  Me.saveLog Me.ActiveLogPage
End Sub

Private Sub MultiPageLogOutput_Change()
  'logDebug "Log Page activated: " & Me.MultiPageLogOutput.SelectedItem.Caption
End Sub

Private Sub frmProgInfo_Click()
  Call showAbout
End Sub

Private Sub lblProgInfo_1_Click()
  Call showAbout
End Sub

Private Sub lblProgInfo_2_Click()
  Call showAbout
End Sub


' === Private Routines =========================================================

Private Sub getHostEnvironment()
  'Get host related info's.
  Const Locale_RegValue = "HKEY_CURRENT_USER\Control Panel\International\Locale"
  Dim WshShell  As New IWshRuntimeLibrary.WshShell
  
  logDebug "getHostEnvironment(): Getting host related info's..."
  
  'Application identification
    VBAHostName = Application.Name
    VBAHostVersion = Application.Version
  
  'Application info to display in Console: tuning for special apps ;-)
    VBAHostDisplayName = VBAHostName
    VBAHostDisplayVersion = VBAHostVersion
    
    If (LCase(VBAHostDisplayName) = "ustation") Then
      VBAHostDisplayName = "Microstation"
    End If
    
    If (Left(LCase(VBAHostDisplayVersion), 7) = "version") Then
      VBAHostDisplayVersion = Trim(Right(VBAHostDisplayVersion, Len(VBAHostDisplayVersion) - 7))
    End If
  
  'Get LanguageID
    On Error Resume Next
    LangID = WshShell.RegRead(Locale_RegValue)
    LangID = Val("&H" & LangID)
    If (Err.Number <> 0) Then
      LangID = 0
    End If
    On Error GoTo 0
    If (LangID = 0) Then logWarning "getHostEnvironment(): Couldn't get Language ID!"
  
  'Finish
  logDebug "getHostEnvironment(): VBAHostName = '" & VBAHostName & "'   VBAHostVersion = '" & VBAHostVersion & "'"
  logDebug "getHostEnvironment(): Language ID = " & CStr(LangID)
  logDebug vbNewLine
  Set WshShell = Nothing
End Sub

Private Sub setLocaleStrings()
  'Set localized strings
    Set oLabels = New Scripting.Dictionary
    
    Select Case LangID
      Case 1031, 5127, 4103, 3079, 2055: 'german: de, li, lu, at, ch
           TypeNameError = "[Fehler] "
           TypeNameWarning = "[Warnung]"
           TypeNameMessage = "[Info]   "
           TypeNameDebug = "[Debug]  "
           
           oLabels("Err_Source") = "FEHLER in         : '"
           oLabels("Err_Number") = "Fehlernummer      : "
           oLabels("Err_Description") = "Fehlerbeschreibung: "
           
           oLabels("LoggingFailed") = "Protokollierung ist nicht verfügbar! (Fehler in frmConsole.Initialize)"
           oLabels("saveLogFailed") = "saveLog(): Protokoll konnte aus irgendeinem Grund nicht gespeichert werden."
           oLabels("Filedialog_Title") = "speichern als ..."
           oLabels("Filedialog_Filter") = "Logdatei (*.log), *.log"
           
           oLabels("WindowTitle") = "Protokoll"
           oLabels("about") = "über"
           oLabels("Version") = "Version"
           oLabels("Licence") = "Lizenz"
           oLabels("Copyright") = "Copyright"
           
           oLabels("frmProgInfo") = "Protokoll-Konsole"
           oLabels("lblProgInfo_1") = "Host-Anwendung:"
           
           oLabels("btnHide") = "Konsole Ausblenden"
           oLabels("btnClearAll") = "Alle Protokolle löschen"
           oLabels("btnClearAll_ToolTip") = "Alle Protokolle löschen"
           oLabels("btnClear") = "Protokoll löschen"
           oLabels("btnClear_ToolTip") = "Aktives Protokoll löschen"
           oLabels("btnSave") = "Protokoll speichern"
           oLabels("btnSave_ToolTip") = "Aktives Protokoll als Datei speichern"
           
           oLabels("frmOptions") = "Einstellungen"
           oLabels("chkCounter") = "Laufende Nummer"
           oLabels("chkDate") = "Datum"
           oLabels("chkTime") = "Zeit"
           oLabels("chkLevel") = "Typ (Niveau)"
           oLabels("chkShowOnError") = "Zeige Konsole bei neuen Fehlern"
           oLabels("chkSource") = "Nachrichten-Quelle"
           oLabels("lblLimit") = "Max. Log-Größe [Zeilen]"
           oLabels("txtLimit_ToolTip") = "Ein großes Protokoll verlangsamt das Programm"
           oLabels("lblSize") = "Dialoggröße"
           oLabels("spinSize_ToolTip") = "Dialoggröße ändern"
           
           oLabels("ColumnCounter") = "Nr"
           oLabels("ColumnDate") = "Datum"
           oLabels("ColumnTime") = "Zeit"
           oLabels("ColumnLevel") = "Typ"
           oLabels("ColumnLogSource") = "Quelle"
           oLabels("ColumnMessage") = "Nachricht"
           
           oLabels("PageErrorlog") = "Fehler"
           oLabels("PageWarninglog") = "Warnungen"
           oLabels("PageLog") = "Protokoll"
           oLabels("PageDebuglog") = "Debug"
           
           oLabels("Log") = "Protokoll"
           oLabels("saved_as") = "gespeichert als"
           
           oLabels("Err_clearLog_1") = "Protokoll konnte nicht gelöscht werden. Falscher Seitenindex angegeben"
           oLabels("Err_saveLog_1") = "Protokoll konnte nicht gespeichert werden. Falscher Seitenindex angegeben"
           oLabels("Err_SetPage_1") = "Falscher Seitenindex angegeben"
           
           
      Case Else: 'english
           TypeNameError = "[Error]  "
           TypeNameWarning = "[Warning]"
           TypeNameMessage = "[Info]   "
           TypeNameDebug = "[Debug]  "
           
           oLabels("Err_Source") = "ERROR in         : '"
           oLabels("Err_Number") = "Error Number     : "
           oLabels("Err_Description") = "Error Description: "
           
           oLabels("LoggingFailed") = "Logging is not availlable! (frmConsole.Initialize failed)"
           oLabels("saveLogFailed") = "saveLog(): Log couldn't be saved for any reason"
           oLabels("Filedialog_Title") = "save as ..."
           oLabels("Filedialog_Filter") = "Logfile (*.log), *.log"
           
           oLabels("WindowTitle") = "Log"
           oLabels("about") = "about"
           oLabels("Version") = "Version"
           oLabels("Licence") = "Licence"
           oLabels("Copyright") = "Copyright"
           
           oLabels("frmProgInfo") = "Logging Console"
           oLabels("lblProgInfo_1") = "Host Application:"
           '
           oLabels("btnHide") = "Hide Console"
           oLabels("btnClearAll") = "Clear all Logs"
           oLabels("btnClearAll_ToolTip") = "Clear all Logs"
           oLabels("btnClear") = "Clear Log"
           oLabels("btnClear_ToolTip") = "Clear the active Log"
           oLabels("btnSave") = "Save Log to File"
           oLabels("btnSave_ToolTip") = "Save the active Log to a file"
           
           oLabels("frmOptions") = "Settings"
           oLabels("chkCounter") = "Message Counter"
           oLabels("chkDate") = "Date"
           oLabels("chkTime") = "Time"
           oLabels("chkLevel") = "Type (Level)"
           oLabels("chkShowOnError") = "Show Console on new Errors"
           oLabels("chkSource") = "Message's Source"
           oLabels("lblLimit") = "Max. Log Length [Lines]"
           oLabels("txtLimit_ToolTip") = "A large Log slowes down the program."
           oLabels("lblSize") = "Dialog Size"
           oLabels("spinSize_ToolTip") = "Change the dialog's size"
           
           oLabels("ColumnCounter") = "No"
           oLabels("ColumnDate") = "Date"
           oLabels("ColumnTime") = "Time"
           oLabels("ColumnLevel") = "Type"
           oLabels("ColumnLogSource") = "Source"
           oLabels("ColumnMessage") = "Message"
           
           oLabels("PageErrorlog") = "Errors"
           oLabels("PageWarninglog") = "Warnings"
           oLabels("PageLog") = "Log"
           oLabels("PageDebuglog") = "Debug"
           
           oLabels("Log") = "Log"
           oLabels("saved_as") = "saved as"
           
           oLabels("Err_clearLog_1") = "Log couldn't be cleared. Wrong page index given"
           oLabels("Err_saveLog_1") = "Log couldn't be saved. Wrong page index given"
           oLabels("Err_SetPage_1") = "Wrong page index given"
    End Select
End Sub

Private Sub initListViews()
  'Initializes the ListView cotrols
  Dim Header     As ColumnHeader
  Dim iLogLevel  As Variant
  Dim oLog       As MSComctlLib.ListView
  
  'Container for quick access to the logs
    Set oLogs = New Scripting.Dictionary
    oLogs.Add LOGLEVEL_ERROR, lvwErrorLog
    oLogs.Add LOGLEVEL_WARNING, lvwWarningLog
    oLogs.Add LOGLEVEL_INFO, lvwLog
    oLogs.Add LOGLEVEL_DEBUG, lvwDebugLog
  
  'Init Column Headers
    For Each iLogLevel In oLogs
      Set oLog = oLogs(iLogLevel)
      
      oLog.ColumnHeaders.Add , COL_KEY_COUNTER, "Counter"
      oLog.ColumnHeaders.Add , COL_KEY_DATE, "Date"
      oLog.ColumnHeaders.Add , COL_KEY_TIME, "Time"
      oLog.ColumnHeaders.Add , COL_KEY_LEVEL, "Level"
      oLog.ColumnHeaders.Add , COL_KEY_SOURCE, "LogSource"
      oLog.ColumnHeaders.Add , COL_KEY_MESSAGE, "Message"
      
      'Header position is reset automatically (arbitrary?) => do not change!
      'Set Header = oLog.ColumnHeaders(1)
      'Header.Position = 5
      'Header.Width = oLog.Width * 0.2
    Next
    
  'Ensure to set correct visibility
    Call chkCounter_Change
    Call chkDate_Change
    Call chkTime_Change
    Call chkLevel_Change
    Call chkSource_Change
    
  Set oLog = Nothing
End Sub

Private Sub logMessage(ByVal LogLevel As Long, ByVal Message As String)
  'Adds a new Message with the specified level to the Logs.
  'Input:  LogLevel: see constants above
  '        Message:  vbNewLine's are ok.
  On Error GoTo 0
  Dim i            As Variant
  Dim iLogLevel    As Variant
  Dim MsgLines     As Variant
  Dim Message4Log  As String
  Dim currentDate  As String
  Dim currentTime  As String
  Dim idx          As Long
  Dim maxLogLen    As Long
  Dim NewMsgLen    As Long
  Dim oLog         As MSComctlLib.ListView
  Dim oItem        As MSComctlLib.ListItem
  
  currentDate = Format(Now, "Short Date")
  currentTime = Format(Now, "Long Time")
  
  'Split message into single lines
  If (Message <> "") Then
    MsgLines = Split(Message, vbNewLine)  'the returned array may be empty
  Else
    'to add the empty message.
    MsgLines = Array("")
  End If
  '
  'Create the Message for every appropriate Log
  For iLogLevel = LOGLEVEL_DEBUG To LogLevel
    Set oLog = oLogs(iLogLevel)
    
    'Add each single line of the message as item to the Log
    For i = LBound(MsgLines) To UBound(MsgLines)
      If (oLog.ListItems.Count = 0) Then
        idx = 1
      Else
        idx = Val(oLog.ListItems(oLog.ListItems.Count).Text) + 1
      End If
      Set oItem = oLog.ListItems.Add()
      oItem.Text = CStr(idx)
      oItem.SubItems(COL_IDX_DATE) = currentDate
      oItem.SubItems(COL_IDX_TIME) = currentTime
      oItem.SubItems(COL_IDX_LEVEL) = oMsgTypeName(LogLevel)
      oItem.SubItems(COL_IDX_SOURCE) = Me.LogSource
      oItem.SubItems(COL_IDX_MESSAGE) = MsgLines(i)
    Next
    oItem.EnsureVisible
    oItem.Selected = True
  Next
  '
  'Cut the Logs if any exceed Me.MaxLogLength.
  Call LimitLogs
  
  'Set new message status. Activate Error/Warning Log page.
  If (not trim(Message) = vbNullString) Then
    
    If (LogLevel = LOGLEVEL_INFO) Then blnNewInfos = true
    
    If ((LogLevel = LOGLEVEL_ERROR) Or (LogLevel = LOGLEVEL_WARNING)) Then
      blnNewWarnings = true
      If (LogLevel = LOGLEVEL_ERROR) Then
        blnNewErrors = true
        Me.ActiveLogPage = PAGE_ERRORLOG
      Else
        Me.ActiveLogPage = PAGE_WARNINGLOG
      End If
      If (Me.ShowOnError) Then
        'Show log if possible. An error occurs if a modeless dialog is shown already.
        'Do not show modal, since the calling routing is pausing then!
        On Error Resume Next
        Me.Show vbModeless
        On Error GoTo 0
      End If
    End If
  End If
End Sub

Private Sub setListViewColumnLabels()
  'Initializes the ListView cotrols
  Dim iLogLevel  As Variant
  Dim oLog       As MSComctlLib.ListView
  
  'Init Column Headers
    For Each iLogLevel In oLogs
      Set oLog = oLogs(iLogLevel)
      oLog.ColumnHeaders(COL_KEY_COUNTER).Text = oLabels("ColumnCounter")
      oLog.ColumnHeaders(COL_KEY_DATE).Text = oLabels("ColumnDate")
      oLog.ColumnHeaders(COL_KEY_TIME).Text = oLabels("ColumnTime")
      oLog.ColumnHeaders(COL_KEY_LEVEL).Text = oLabels("ColumnLevel")
      oLog.ColumnHeaders(COL_KEY_SOURCE).Text = oLabels("ColumnLogSource")
      oLog.ColumnHeaders(COL_KEY_MESSAGE).Text = oLabels("ColumnMessage")
    Next
  Set oLog = Nothing
End Sub

Private Sub setListViewColumnVisibility(ByVal ColumnKey As String, ByVal Visible As Boolean)
  'Sets the visibility of the given column in all Logs.
  Dim iLogLevel  As Variant
  Dim Width      As Single
  
  If (Visible) Then
    Width = DEFAULT_COLUMN_WIDTH
  Else
    Width = 0
  End If
  
  For Each iLogLevel In oLogs
    oLogs(iLogLevel).ColumnHeaders(ColumnKey).Width = Width
  Next
  
  Call resizeListViewColumns
End Sub

Private Sub resizeListViewColumns()
  'Resizes the message column to fit into the view.
  Dim iLogLevel  As Variant
  Dim oLog       As Object
  Dim oHeaders   As MSComctlLib.ColumnHeaders
  
  For Each iLogLevel In oLogs
    Set oLog = oLogs(iLogLevel)
    Set oHeaders = oLog.ColumnHeaders
    
    oHeaders(COL_KEY_MESSAGE).Width = oLog.Width - oHeaders(COL_KEY_COUNTER).Width _
                                                 - oHeaders(COL_KEY_DATE).Width _
                                                 - oHeaders(COL_KEY_TIME).Width _
                                                 - oHeaders(COL_KEY_LEVEL).Width _
                                                 - oHeaders(COL_KEY_SOURCE).Width _
                                                 - OS_Win_ScrollbarWidth
  Next
  
  Set oLog = Nothing
  Set oHeaders = Nothing
End Sub

Private Sub LimitLogs()
  'Cuts the content of each Log if it exceeds the maximum number of lines.
  On Error GoTo 0
  Const cutPercent As Long = 20  'part to cut away in percent (ca.)
  
  Dim i            As Variant
  Dim iLogLevel    As Variant
  Dim OldLen       As Long
  Dim MaxLen       As Long
  Dim CutLen       As Long
  Dim oLog         As Object
  
  MaxLen = Me.MaxLogLength
  
  For Each iLogLevel In oLogs
    OldLen = oLogs(iLogLevel).ListItems.Count
    If (OldLen > MaxLen) Then
      Set oLog = oLogs(iLogLevel)
      CutLen = Int((MaxLen * cutPercent / 100) + OldLen - MaxLen)
      
      For i = 1 To CutLen
        oLog.ListItems.Remove (1)
      Next
      
      logDebug "*******  LimitLogs(): Cut Log '" & oLog.Parent.Caption & "'  (Old length = " & Format(OldLen, "#") & " lines, " & _
                                                                 "(New length = " & Format(oLog.ListItems.Count, "#") & " lines)."
    End If
  Next
  Set oLog = Nothing
End Sub

Private Function getFileNameFromDialog( _
  Optional ByVal OpenFile As Boolean = True, _
  Optional ByVal DialogTitle As String, _
  Optional ByVal InitialFilename As String = "", _
  Optional ByVal InitialDirectory As String, _
  Optional ByVal FileFilter As String = "All files (*.*),*.*", _
  Optional ByVal FileFilterIndex As Long = 1, _
  Optional ByVal Flags As Long = 0) As String
  ' ---------------------------------------------------------------------
  ' Get a filename from a common file dialog.
  ' Input Arguments:
  '   OpenFile         ... Witch Dialog should be shown: open = true, save = false
  '                        Default: true = open
  '   DialogTitle      ... Dialog title
  '                        Default: system default
  '   InitialFilename  ... The File Name field is initialized with it.
  '                        If a path is given, the contained directory will be
  '                        the initial directory (Windows 2000 and later)!
  '                        It must not end with a backslash (can't be a directory only)
  '                        Default: empty
  '   InitialDirectory ... This will be the initial directory if it isn't set
  '                        already by InitialFilename.
  '                        Default: current working directory.
  '   FileFilter       ... i.e. "All files (*.*),*.*,Text files (*.txt),*.txt"
  '                        Default: "All files (*.*),*.*"
  '   FileFilterIndex  ... The FileFilter to preselect (First index = 1)
  '                        Default: 1
  '   Flags            ... A combination of OFN_*** constants to initialize the dialog box.
  '                        Default: 0 = no flags
  '
  ' Return value:          full path of the choosen file, or "".
  '
  ' Example call:
  ' path = getFileNameFromDialog (false, "Choose...", "dummy.txt", "C:\shared", "Text Files (*.txt),*.txt", 1)
  ' ---------------------------------------------------------------------
  Const MAX_PATH_LENGTH  As Integer = 255
  
  'Constants for OPENFILENAME structure
    Const OFN_READONLY             As Long = &H1
    Const OFN_OVERWRITEPROMPT      As Long = &H2
    'Const OFN_HideReadOnly         As Long = &H4
    'Const OFN_NOCHANGEDIR          As Long = &H8
    'Const OFN_SHOWHELP             As Long = &H10
    'Const OFN_NOVALIDATE           As Long = &H100
    'Const OFN_ALLOWMULTISELECT     As Long = &H200
    'Const OFN_EXTENSIONDIFFERENT   As Long = &H400
    'Const OFN_PATHMUSTEXIST        As Long = &H800
    Const OFN_FILEMUSTEXIST        As Long = &H1000
    'Const OFN_CREATEPROMPT         As Long = &H2000
    'Const OFN_SHAREAWARE           As Long = &H4000
    'Const OFN_NOREADONLYRETURN     As Long = &H8000
    'Const OFN_NOTESTFILECREATE     As Long = &H10000
    'Const OFN_NONETWORKBUTTON      As Long = &H20000
    'Const OFN_NOLONGNAMES          As Long = &H40000
    'Const OFN_EXPLORER             As Long = &H80000
    'Const OFN_NODEREFERENCELINKS   As Long = &H100000
    'Const OFN_LONGNAMES            As Long = &H200000
  
  'Declarations
    Dim HostApp         As Object
    Dim OFN_Structure   As OPENFILENAME
    Dim success         As Long
    Dim FilePath        As Variant
    Dim DefaultExt      As String
  '
  Select Case VBAHostName
    Case VBA_HOST_NAME_EXCEL
      'Excel: When the general system dialog is shown and the user presses ESCAPE,
        '     then the VBA editor will be shown in debugging mode. That's why using Excel's own dialog.
        'CAUTION: This dialog does not provide a warning related to file overwrite!
        
        Set HostApp = Application  'this is for getting the VBProject compiled under other hosts than Excel
        Const msoFileDialogSaveAs = 2
        If (OpenFile) Then
          FilePath = HostApp.GetOpenFileName(FileFilter, FileFilterIndex, DialogTitle)
        Else
          FilePath = HostApp.getSaveAsFilename(InitialFilename, FileFilter, FileFilterIndex, DialogTitle)
        End If
        
        'Get the result
        If (FilePath = False) Then
          FilePath = ""
          'MsgBox "user cancelled!", vbInformation
        Else
          FilePath = Trim(FilePath)
          'MsgBox "'" & FilePath & "'"
        End If
        
        
    Case Else
        'Validate and edit arguments for convenience and avoiding errors
          If (IsMissing(DialogTitle)) Then DialogTitle = vbNullChar
          If (Right(InitialFilename, 1) = "\") Then InitialFilename = ""
          If (IsMissing(InitialDirectory)) Then InitialDirectory = CurDir()
          FileFilter = Replace(FileFilter, ",", vbNullChar) & vbNullChar
          If (OpenFile) Then
            'open
            Flags = Flags Or OFN_FILEMUSTEXIST Or OFN_READONLY  'resp. HideReadOnly
          Else
            'save
            Flags = Flags Or OFN_OVERWRITEPROMPT
          End If
        
        'Fill OPENFILENAME structure for api call
          OFN_Structure.lStructSize = Len(OFN_Structure)
          With OFN_Structure
            .lpstrTitle = DialogTitle
            .lpstrFile = InitialFilename & String(MAX_PATH_LENGTH - Len(InitialFilename), 0)
            .lpstrInitialDir = InitialDirectory
            .lpstrFilter = FileFilter
            .nFilterIndex = FileFilterIndex
            '.lpstrDefExt     = DefaultExt
            .Flags = Flags
            .nMaxFile = Len(.lpstrFile) - 1
            .lpstrFileTitle = .lpstrFile
            .nMaxFileTitle = .nMaxFile
          End With
        
        'Show dialog
          If (OpenFile) Then
            success = GetOpenFileName(OFN_Structure)
          Else
            success = GetSaveFileName(OFN_Structure)
          End If
        
        'Get the result
          If (success = 0) Then
            FilePath = ""
            'MsgBox "user cancelled or error calling api function!", vbInformation
          Else
            FilePath = Trim(Left(OFN_Structure.lpstrFile, InStr(1, OFN_Structure.lpstrFile, vbNullChar) - 1))
            'MsgBox "'" & FilePath & "'"
          End If
  End Select
  
  logDebug "Logging Console: File dialog returned filename '" & FilePath & "'."
  getFileNameFromDialog = FilePath
End Function


Sub showAbout()
  'About dialog
  Dim Title       As String
  Dim Message     As String
  Title = oLabels("about") & " " & oLabels("frmProgInfo")
  Message = oLabels("Version") & vbTab & vbTab & INFO_VERSION & vbLf & _
            oLabels("Licence") & vbTab & vbTab & INFO_LICENCE & vbLf & _
            oLabels("Copyright") & vbTab & vbTab & INFO_COPYLEFT & " (" & INFO_MAIL & ")" & vbLf & vbLf & _
            "VBAHostName" & vbTab & VBAHostName & vbLf & _
            "VBAHostVersion" & vbTab & VBAHostVersion
  Call MsgBox(Message, vbOKOnly, Title)
  Exit Sub
End Sub




' === Methods ==================================================================

Public Sub logError(ByVal Message As String)
  'Like the others but: if err.number <> 0 the error infos are added.
  Dim ErrInfo
  If (Err.Number <> 0) Then
    ErrInfo = oLabels("Err_Source") & Err.Source & "':" & vbNewLine & _
              oLabels("Err_Number") & CStr(Err.Number) & vbNewLine & _
              oLabels("Err_Description") & Err.Description
    Err.Clear
    'Workaround for unprintable character at the end of some error descriptions
    If (Asc(Right(ErrInfo, 1)) < 32) Then ErrInfo = Left(ErrInfo, Len(ErrInfo) - 1)
    
    logMessage LOGLEVEL_ERROR, ErrInfo
  End If
  logMessage LOGLEVEL_ERROR, Message
End Sub


Public Sub logWarning(ByVal Message As String)
  logMessage LOGLEVEL_WARNING, Message
End Sub


Public Sub logInfo(ByVal Message As String)
  logMessage LOGLEVEL_INFO, Message
End Sub


Public Sub logDebug(ByVal Message As String)
  logMessage LOGLEVEL_DEBUG, Message
End Sub


Public Sub clearAllLogs()
  'Clears all Logs.
  Call clearLog(PAGE_ERRORLOG)
  Call clearLog(PAGE_WARNINGLOG)
  Call clearLog(PAGE_LOG)
  Call clearLog(PAGE_DEBUGLOG)
End Sub


Public Sub clearLog(ByVal LogPageIndex As Long)
  'Clears the specified Log.
  On Error Resume Next
  Me.Controls("MultiPageLogOutput").Pages(LogPageIndex).Controls(0).ListItems.Clear
  If (Err.Number <> 0) Then
    logError "clearLog(): " & oLabels("Err_clearLog_1") & " (" & LogPageIndex & ")?"
  Else
    logDebug "clearLog(): Log cleared: " & Me.Controls("MultiPageLogOutput").Pages(LogPageIndex).Caption
  End If
  Err.Clear
End Sub


Public Function saveLog(ByVal LogPageIndex As Long, Optional ByVal LogFilePath As String) As String
  'Saves the specified Log to a file.
  'Input: LogPageIndex ... a "public" constant: Me.PAGE_ERRORLOG, Me.PAGE_WARNINGLOG, Me.PAGE_LOG, Me.PAGE_DEBUGLOG
  '       LogFilePath  ... if empty, a dialog appears
  'Returns the path of the file that has been written, or "" (if not saved).
  Dim InitialFilename  As String
  Dim InitialDirectory As String
  Dim Title            As String
  Dim LogType          As String
  Dim FileLine         As String
  Dim FileNum          As Integer
  Dim i                As Variant
  Dim iLogLevel        As Variant
  Dim oLog             As MSComctlLib.ListView
  Dim oItem            As MSComctlLib.ListItem
  
  On Error Resume Next
  LogType = Me.Controls("MultiPageLogOutput").Pages(LogPageIndex).Caption
  
  If (Err.Number <> 0) Then
    logError "saveLog(): " & oLabels("Err_saveLog_1") & " (" & LogPageIndex & ")!"
    Err.Clear
  Else
    'Get a filename via dialog
    If (LogFilePath = "") Then
      Title = LogType & " " & oLabels("Filedialog_Title")
      If (Me.LogSource <> "") Then InitialFilename = Me.LogSource & "_"
      InitialFilename = InitialFilename & LogType & ".log"
      InitialDirectory = CurDir()
      LogFilePath = getFileNameFromDialog(False, Title, InitialFilename, InitialDirectory, oLabels("Filedialog_Filter"), 1)
      logDebug "Logging Console: SaveAs dialog returned filename '" & LogFilePath & "'."
    End If
    
    If (LogFilePath = "") Then
      'user canceled
    Else
      'Got an absolute filename in an existing directory. The file may not exist.
      On Error Resume Next
      FileNum = FreeFile
      Open LogFilePath For Output Access Write As #FileNum
      
      'Print every Line of the Log considering the view settings
      Set oLog = Me.Controls("MultiPageLogOutput").Pages(LogPageIndex).Controls(0)
      
      For i = 1 To oLog.ListItems.Count
        Set oItem = oLog.ListItems(i)
        
        FileLine = ""
        If (Me.chkCounter) Then FileLine = FileLine & Format(oItem.Text, "@@@@@@") & vbTab
        If (Me.chkDate) Then FileLine = FileLine & oItem.SubItems(COL_IDX_DATE) & vbTab
        If (Me.chkTime) Then FileLine = FileLine & oItem.SubItems(COL_IDX_TIME) & vbTab
        If (Me.chkLevel) Then FileLine = FileLine & oItem.SubItems(COL_IDX_LEVEL) & vbTab
        If (Me.chkSource) Then FileLine = FileLine & oItem.SubItems(COL_IDX_SOURCE) & ":" & vbTab
        FileLine = FileLine & oItem.SubItems(COL_IDX_MESSAGE)
        
        Print #FileNum, FileLine
      Next
      
      Close #FileNum
      If (Err.Number <> 0) Then
        logError oLabels("saveLogFailed")
        Err.Clear
      Else
        logInfo "'" & Me.Controls("MultiPageLogOutput").Pages(LogPageIndex).Caption & "' " & oLabels("saved_as") & " '" & LogFilePath & "'"
      End If
      
      Set oLog = Nothing
    End If
  End If
  saveLog = LogFilePath
End Function


Public Sub resetAllSettings()
  'Resets all settings to their defaults.
  Call resetLogSettings
  Call resetDialogSettings
End Sub


Public Sub resetLogSettings()
  'Resets Log related settings to their defaults.
  '=> This is intented to use in programs to reset the Log behavior
  '   without changing the dialog position and size.
  
  'Options
    Me.IncludeCounter = DEFAULT_INCLUDECOUNTER
    Me.IncludeDate = DEFAULT_INCLUDEDATE
    Me.IncludeTime = DEFAULT_INCLUDETIME
    Me.IncludeLevel = DEFAULT_INCLUDELEVEL
    Me.ShowOnError = DEFAULT_SHOWONERROR
    Me.LogSource = DEFAULT_LOGSOURCE
    Me.MaxLogLength = DEFAULT_MAXLOGLENGTH
    
  'Active Log Page
    Me.ActiveLogPage = DEFAULT_ACTIVELOGPAGE
End Sub


Public Sub resetDialogSettings()
  'Resets dialog position and size
  Me.Move DEFAULT_DIALOG_LEFT, DEFAULT_DIALOG_TOP, DEFAULT_DIALOG_WIDTH, DEFAULT_DIALOG_HEIGHT
End Sub


Public Sub storeSettings()
  'Stores all settings to the registry
  On Error Resume Next
  
  Const TYP_REGSZ   As String = "REG_SZ"
  Dim WshShell      As New IWshRuntimeLibrary.WshShell
  
  logDebug "Store Settings to registry for '" & Me.WindowTitle & "'"
  
  'Options
    WshShell.RegWrite REGKEY_ROOT & "IncludeCounter", CStr(Me.IncludeCounter), TYP_REGSZ
    WshShell.RegWrite REGKEY_ROOT & "IncludeDate", CStr(Me.IncludeDate), TYP_REGSZ
    WshShell.RegWrite REGKEY_ROOT & "IncludeTime", CStr(Me.IncludeTime), TYP_REGSZ
    WshShell.RegWrite REGKEY_ROOT & "IncludeLevel", CStr(Me.IncludeLevel), TYP_REGSZ
    WshShell.RegWrite REGKEY_ROOT & "IncludeSource", CStr(Me.IncludeSource), TYP_REGSZ
    WshShell.RegWrite REGKEY_ROOT & "ShowOnError", CStr(Me.ShowOnError), TYP_REGSZ
    WshShell.RegWrite REGKEY_ROOT & "LogSource", Me.LogSource, TYP_REGSZ
    WshShell.RegWrite REGKEY_ROOT & "MaxLogLength", CStr(Me.MaxLogLength), TYP_REGSZ
    
  'Active Log Page
    WshShell.RegWrite REGKEY_ROOT & "ActiveLogPage", CStr(Me.ActiveLogPage), TYP_REGSZ
    
  'dialog position and size
    WshShell.RegWrite REGKEY_ROOT & "Dialog_Left", CStr(Me.Left), TYP_REGSZ
    WshShell.RegWrite REGKEY_ROOT & "Dialog_Top", CStr(Me.Top), TYP_REGSZ
    WshShell.RegWrite REGKEY_ROOT & "Dialog_Width", CStr(Me.Width), TYP_REGSZ
    WshShell.RegWrite REGKEY_ROOT & "Dialog_Height", CStr(Me.Height), TYP_REGSZ
    
  On Error GoTo 0
  Set WshShell = Nothing
End Sub


Public Sub restoreAllSettings()
  'Restores all settings from the registry
  On Error Resume Next
  
  Const TYP_REGSZ     As String = "REG_SZ"
  Dim WshShell        As New IWshRuntimeLibrary.WshShell
  Dim DialogLeft      As Integer
  Dim DialogTop       As Integer
  Dim DialogWidth     As Integer
  Dim DialogHeight    As Integer
  Dim currentFactor   As Double
  Dim SpinFactor      As Double
  
  logDebug "Restore Settings from registry for '" & Me.WindowTitle & "'"
  
  'Registry root key for settings of this implementation of Logging Console
    REGKEY_ROOT = REGKEY_CLASS_ROOT & Me.Caption & "\"
  
  'Options
    Me.IncludeCounter = CBool(WshShell.RegRead(REGKEY_ROOT & "IncludeCounter"))
    Me.IncludeDate = CBool(WshShell.RegRead(REGKEY_ROOT & "IncludeDate"))
    Me.IncludeTime = CBool(WshShell.RegRead(REGKEY_ROOT & "IncludeTime"))
    Me.IncludeLevel = CBool(WshShell.RegRead(REGKEY_ROOT & "IncludeLevel"))
    Me.IncludeSource = CBool(WshShell.RegRead(REGKEY_ROOT & "IncludeSource"))
    Me.ShowOnError = CBool(WshShell.RegRead(REGKEY_ROOT & "ShowOnError"))
    Me.LogSource = WshShell.RegRead(REGKEY_ROOT & "LogSource")
    Me.MaxLogLength = CLng(WshShell.RegRead(REGKEY_ROOT & "MaxLogLength"))
    
  'Active Log Page
    If (RESTORE_ACTIVELOGPAGE) Then Me.ActiveLogPage = CInt(WshShell.RegRead(REGKEY_ROOT & "ActiveLogPage"))
    
  'Dialog position
    DialogLeft = CInt(WshShell.RegRead(REGKEY_ROOT & "Dialog_Left"))
    DialogTop = CInt(WshShell.RegRead(REGKEY_ROOT & "Dialog_Top"))
    If (DialogLeft = 0) Then DialogLeft = DEFAULT_DIALOG_LEFT
    If (DialogTop = 0) Then DialogTop = DEFAULT_DIALOG_TOP
    Me.Move DialogLeft, DialogTop
    
  'Dialog size
    DialogWidth = CInt(WshShell.RegRead(REGKEY_ROOT & "Dialog_Width"))
    'DialogHeight = CInt(WshShell.RegRead(REGKEY_ROOT & "Dialog_Height"))
    
    'Min/Max corrections for width and height
    If (DialogWidth < DEFAULT_DIALOG_WIDTH) Then
      DialogWidth = DEFAULT_DIALOG_WIDTH
    ElseIf (DialogWidth > maxWidthPoints) Then
      DialogWidth = maxWidthPoints
    End If
    'If (DialogHeight < DEFAULT_DIALOG_HEIGHT) Then
    '  DialogHeight = DEFAULT_DIALOG_HEIGHT
    'elseif(DialogHeight > maxHeightPoints) Then
    '  DialogHeight = maxHeightPoints
    'end if
    
    'Only width is considered for computing the spinner's factor.
    currentFactor = (DialogWidth - minWidthPoints) / minWidthPoints
    SpinFactor = currentFactor / maxFactor
    
    'Restore last size via changing SpinButton.
    If ((SpinFactor > 0) And (SpinFactor < 1)) Then
      spinSize.Value = spinSize.Min + SpinFactor * (spinSize.Max - spinSize.Min)
    End If
    
  On Error GoTo 0
  Set WshShell = Nothing
End Sub


' === Properties ===============================================================

'Info's to add to the Log message
Public Property Let IncludeCounter(inpIncludeCounter As Boolean)
  chkCounter.Value = inpIncludeCounter
End Property
Public Property Get IncludeCounter() As Boolean
  IncludeCounter = chkCounter.Value
End Property

Public Property Let IncludeDate(inpIncludeDate As Boolean)
  chkDate.Value = inpIncludeDate
End Property
Public Property Get IncludeDate() As Boolean
  IncludeDate = chkDate.Value
End Property

Public Property Let IncludeTime(inpIncludeTime As Boolean)
  chkTime.Value = inpIncludeTime
End Property
Public Property Get IncludeTime() As Boolean
  IncludeTime = chkTime.Value
End Property

Public Property Let IncludeLevel(inpIncludeLevel As Boolean)
  chkLevel.Value = inpIncludeLevel
End Property
Public Property Get IncludeLevel() As Boolean
  IncludeLevel = chkLevel.Value
End Property

Public Property Let IncludeSource(inpIncludeSource As Boolean)
  chkSource.Value = inpIncludeSource
End Property
Public Property Get IncludeSource() As Boolean
  IncludeSource = chkSource.Value
End Property

Public Property Let ShowOnError(inpShowOnError As Boolean)
  chkShowOnError.Value = inpShowOnError
  logDebug "Set ShowOnError = '" & CStr(chkShowOnError.Value) & "'"
End Property
Public Property Get ShowOnError() As Boolean
  ShowOnError = chkShowOnError.Value
End Property

Public Property Let LogSource(inpLogSource As String)
  'Sets the LogSource
  strLogSource = inpLogSource
  logDebug "Set LogSource = '" & strLogSource & "'"
End Property
Public Property Get LogSource() As String
  LogSource = strLogSource
End Property


'Other status properties
Public Property Get newErrors() As Boolean
  'Returns "true", if there are new (resp. unread) errors
  newErrors = blnNewErrors
End Property

Public Property Get newWarnings() As Boolean
  'Returns "true", if there are new (resp. unread) Warnings
  newWarnings = blnNewWarnings
End Property

Public Property Get newInfos() As Boolean
  'Returns "true", if there are new (resp. unread) errors
  newInfos = blnNewInfos
End Property

Public Property Let MaxLogLength(inpMaxLogLength As Long)
  'Sets the max. number of lines of every single Log.
  txtLimit.Value = inpMaxLogLength
  logDebug "Set MaxLogLength = '" & txtLimit.Value & "' lines."
End Property
Public Property Get MaxLogLength() As Long
  'Returns the max. length of every single Log in [kB]
  If (IsNumeric(txtLimit.Value)) Then
    MaxLogLength = CLng(txtLimit.Value)
  Else
    Dim oldWrongLimit
    oldWrongLimit = txtLimit.Value
    txtLimit.Value = CStr(DEFAULT_MAXLOGLENGTH)
    MaxLogLength = DEFAULT_MAXLOGLENGTH
    logDebug "Set MaxLogLength = '" & CStr(DEFAULT_MAXLOGLENGTH) & "' because of wrong old value '" & oldWrongLimit & "'"
  End If
End Property

Public Property Let ActiveLogPage(inpActiveLogPageIndex As Long)
  'Activates a specific LogPage in the dialog
  On Error Resume Next
  Me.MultiPageLogOutput.Value = inpActiveLogPageIndex
  If (Err.Number <> 0) Then
    logError "Set ActiveLogPage(): " & oLabels("Err_SetPage_1") & " (" & inpActiveLogPageIndex & ")"
  End If
  Err.Clear
End Property
Public Property Get ActiveLogPage() As Long
  'Returns the index of the actual Log page.
  ActiveLogPage = Me.MultiPageLogOutput.SelectedItem.Index
End Property

Public Property Let WindowTitle(ByVal inpWindowTitle As String)
  'Set the Console window's title and restores related settings - must not be "".
  inpWindowTitle = Trim(inpWindowTitle)
  If (inpWindowTitle <> "") Then
    Me.Caption = inpWindowTitle
    logDebug "Set Window Title = '" & Me.Caption & "'"
    
    'Restore settings that are stored for this window title
    Call restoreAllSettings
  End If
End Property
Public Property Get WindowTitle() As String
  WindowTitle = Me.Caption
End Property

'for jEdit:  :folding=indent::collapseFolds=1:
