VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBatchPDF 
   Caption         =   "PDF-Export im Stapel"
   ClientHeight    =   4690
   ClientLeft      =   30
   ClientTop       =   370
   ClientWidth     =   7840
   OleObjectBlob   =   "frmBatchPDF.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmBatchPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2004-2017  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'==================================================================================================
'Modul frmBatchPDF
'==================================================================================================
'
'Dialog zum Erfassen der Parameter für die gewünschte PDF-Erzeugung
'
'Als Startwerte für den Dialog werden Eigenschaften des Objektes oExpim verwendet.
'Änderungen der Dialogwerte werden wieder umgesetzt in Eigenschaften des Objektes oExpimGlobal.
'==================================================================================================


Option Explicit

'Deklarationen
Const ForeColor_Red             As Long = &HFF&
Const ForeColor_Black           As Long = &H80000008

Const SheetFilter_Active        As Integer = 1
Const SheetFilter_All           As Integer = 2

Const FileFilter_AllWorkbooks   As String = "*.xlsx;*.xlsm;*.xls"
Const PrefixLockfile            As String = "~$"


Private LastFolder              As String
Private LastFileFilter          As String
Private LastSearchSubFolders    As Boolean
Private FileCount               As Long
Private Folder                  As String
Private FileFilter              As String
Private SheetFilter             As Integer
Private SearchSubFolders        As Boolean
Private IsCancellationRequested As Boolean
Private IsBusy                  As Boolean

Private oSysTools               As CToolsSystem
Private oFileList               As Scripting.Dictionary



Private Sub UserForm_Initialize()
    DebugEcho "frmBatchPDF():  UserForm_Initialize() ..."
    
    Dim StartDir    As String
    
    If (Not Application.ActiveWorkbook Is Nothing) Then
        StartDir = Application.ActiveWorkbook.Path
    End If
    If (StartDir = "") Then
        StartDir = CurDir()
    End If
    
    Me.tbFolder.value = StartDir
    
    Call RefreshUI
End Sub

Private Sub UserForm_Terminate()
    Set oFileList = Nothing
End Sub


Private Sub btnCancel_Click()
    ' Quit dialog or request cancellation of running pdf export
    If (IsBusy) Then
        IsCancellationRequested = True
    Else
        Unload Me
    End If
End Sub

Private Sub btnOK_Click()
    Call ExportPDF
End Sub


Private Sub btnFolder_Click()
    ' Get a folder name (via dummy filename) from file dialog
    
    Dim FileFilterDummy  As Variant
    Dim FolderPath       As Variant
    'Dim InitialDirectory As Variant
    Dim InitialFilename  As Variant
    Dim Title            As Variant
    
    Title = "Ordner wählen ..."
    FileFilterDummy = "dummy (*.___), *.___"
    InitialFilename = "_"
    'InitialDirectory = CurDir()
    'FolderPath       = ThisWorkbook.SysTools.getFileNameFromDialog(False, Title, InitialFilename, InitialDirectory, FileFilterDummy, 1)
    FolderPath = Application.GetSaveAsFilename(InitialFilename, FileFilterDummy, 1, Title)
    DebugEcho "BatchPDF: Folder dialog returned filename '" & FolderPath & "'."
    
    If (FolderPath = False) Then
      'user canceled
    Else
        Me.tbFolder.value = Verz(Trim(FolderPath))
    End If
End Sub


Private Sub optSheetFilter_Active_Change()
    Call RefreshUI
End Sub

Private Sub optSheetFilter_All_Change()
    Call RefreshUI
End Sub


Private Sub optFileFilter_AllWorkbooks_Change()
    Call RefreshUI
End Sub

Private Sub optFileFilter_Glob_Change()
    Call RefreshUI
End Sub

Private Sub tbFileFilter_Glob_Change()
    Call RefreshUI
End Sub

Private Sub chkSearchSubFolders_Change()
    Call RefreshUI
End Sub

Private Sub tbFolder_Change()
    Call RefreshUI
End Sub


Private Sub ExportPDF()
    ' Do the work.
    On Error GoTo Catch
    
    Dim WorkbookOrSheet As Object
    Dim SourceFilePath  As Variant
    Dim TargetFilePath  As String
    Dim Message         As String
    Dim ErrCount        As Integer
    Dim PDFCount        As Integer
    Dim FileCount       As Integer
    Dim MsgBoxFlags     As Integer
    Dim StatusScreen    As Boolean
    Dim StatusEvents    As Boolean
    Dim StatusCalc      As Boolean
    
    StatusScreen = Application.ScreenUpdating
    StatusEvents = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    IsBusy = True
    ErrCount = 0
    PDFCount = 0
    FileCount = 0
    
    For Each SourceFilePath In oFileList
        
        FileCount = FileCount + 1
        
        On Error Resume Next
        Application.Workbooks.Open FileName:=SourceFilePath, ReadOnly:=True, Notify:=False, UpdateLinks:=False
        
        If (Err <> 0) Then
            ErrEcho "BatchPDF:  Fehler beim Öffnen der Datei '" & SourceFilePath & "'"
            ErrCount = ErrCount + 1
        Else
            On Error GoTo Catch
            TargetFilePath = ActiveWorkbook.Path & "\" & VorName(ActiveWorkbook.FullName) & ".pdf"
            
            ' SheetFilter
            If (SheetFilter = SheetFilter_Active) Then
                Set WorkbookOrSheet = Application.ActiveSheet
            Else
                Set WorkbookOrSheet = Application.ActiveWorkbook
            End If
            
            Echo "BatchPDF:  Erzeuge '" & TargetFilePath & "'"
            Call ProgressbarAllgemein(oFileList.Count, FileCount, "Erzeuge " & TargetFilePath)
            
            On Error Resume Next
            WorkbookOrSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=TargetFilePath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
            
            If (Err <> 0) Then
                ErrEcho "BatchPDF:  Fehler beim PDF-Export der Datei '" & SourceFilePath & "'"
                ErrCount = ErrCount + 1
            Else
                PDFCount = PDFCount + 1
            End If
            On Error GoTo Catch
            
            Application.ActiveWorkbook.Close SaveChanges:=False
        End If
        
        DoEvents
        If (IsCancellationRequested) Then Exit For
    Next
    
    On Error GoTo Catch
    
    Application.EnableEvents = StatusEvents
    Application.ScreenUpdating = StatusScreen
    'Call ClearStatusBarDelayed(3)
    
    ' Final Message
    If (oFileList.Count = 1) Then
        Message = oFileList.Count & "  Datei war zu verarbeiten"
    Else
        Message = oFileList.Count & "  Dateien waren zu verarbeiten"
    End If
    If (PDFCount = 1) Then
        Message = Message & vbNewLine & PDFCount & "  Datei erfolgreich als PDF exportiert"
    Else
        Message = Message & vbNewLine & PDFCount & "  Dateien erfolgreich als PDF exportiert"
    End If
    If (ErrCount > 0) Then
        If (ErrCount = 1) Then
            Message = Message & vbNewLine & ErrCount & "  Datei konnte nicht exportiert werden (siehe Protokoll)."
        Else
            Message = Message & vbNewLine & ErrCount & "  Dateien konnten nicht exportiert werden (siehe Protokoll)."
        End If
        MsgBoxFlags = vbOKOnly + vbExclamation
    Else
        MsgBoxFlags = vbOKOnly + vbInformation
    End If
    If (IsCancellationRequested) Then
        Message = "PDF-Export durch Benutzer abgebrochen!" & vbNewLine & vbNewLine & Message
        MsgBoxFlags = vbOKOnly + vbCritical
    End If
    MsgBox Message, MsgBoxFlags, "Ergebnis PDF-Export"
    
    Call UpdateGeoToolsRibbon
    
    IsBusy = False
    IsCancellationRequested = False
    Call ClearStatusBar
    
    
    Exit Sub
Catch:
    IsBusy = False
    IsCancellationRequested = False
    Application.StatusBar = False
    Application.EnableEvents = StatusEvents
    Application.ScreenUpdating = StatusScreen
    ErrMessage = "Fehler beim Erzeugen der PDF-Datei(en) => ABBRUCH"
    FehlerNachricht "frmBatchPDF.ExportPDF()"
End Sub

Private Sub RefreshUI()
    ' Komplette Prüfung/Herstellung der Oberflächenkonsistenz
    On Error GoTo Catch
    
    Dim success         As Boolean
    Dim Key             As Variant
    Dim FileListKeys    As Variant
    Dim idx             As Variant
    Dim FileName        As String
    
    success = True
    Me.btnOK.Enabled = False
    IsBusy = False
    
    ' SheetFilter
    If (optSheetFilter_Active.value) Then
        SheetFilter = SheetFilter_Active
    Else
        SheetFilter = SheetFilter_All
    End If
    
    ' Folder
    Folder = Trim(Me.tbFolder.value)
    If ((Folder = "") Or (Not ThisWorkbook.SysTools.isVerzeichnis(Folder))) Then
        success = False
        Folder = ""
        Me.tbFolder.ForeColor = ForeColor_Red
    Else
        Me.tbFolder.ForeColor = ForeColor_Black
    End If
    
    ' Search Sub Folders
    SearchSubFolders = Me.chkSearchSubFolders.value
    
    ' FileFilter
    If (optFileFilter_AllWorkbooks.value) Then
        FileFilter = FileFilter_AllWorkbooks
    Else
        FileFilter = Trim(tbFileFilter_Glob.value)
        
        If (FileFilter = "") Then
            success = False
        End If
    End If
    
    ' FileList
    If (success) Then
        
        Me.LstFileList.Clear
        
        ' New file search
        If ((Folder <> LastFolder) Or (FileFilter <> LastFileFilter) Or (LastSearchSubFolders Xor SearchSubFolders)) Then
            
            Me.lblFileList.Caption = "Suche nach Dateien ..."
            
            On Error Resume Next
            Set oFileList = ThisWorkbook.SysTools.FindFiles(FileFilter, Folder, SearchSubFolders)
            If (oFileList Is Nothing) Then
                Set oFileList = New Scripting.Dictionary
            End If

            ErrMessage = ""
            On Error GoTo Catch
            
            ' Remove Lockfiles and ThisWorkbook
            FileListKeys = oFileList.Keys
            For idx = 0 To oFileList.Count - 1
                Key = FileListKeys(idx)
                FileName = oFileList.Item(Key)
                If (Left(FileName, 2) = PrefixLockfile) Then
                    DebugEcho "Datei '" & Key & "' wird aus Dateiliste entfernt."
                    oFileList.Remove (Key)
                ElseIf (LCase(FileName) = LCase(ThisWorkbook.Name)) Then
                    DebugEcho "Datei '" & Key & "' wird aus Dateiliste entfernt."
                    oFileList.Remove (Key)
                End If
            Next
            
            Call SortDictionary(oFileList, 1, 1, False)
            
            LastFolder = Folder
            LastFileFilter = FileFilter
            LastSearchSubFolders = SearchSubFolders
        End If
        
        ' Refresh ListBox.
        For Each Key In oFileList
            Me.LstFileList.AddItem oFileList.Item(Key)
        Next
        
        If (oFileList.Count < 1) Then
            success = False
            FileCount = 0
        Else
            FileCount = oFileList.Count
        End If
    Else
        FileCount = 0
        Me.LstFileList.Clear
    End If
    Me.lblFileList.Caption = CStr(FileCount) & " Dateien:"
    
    
    ' OK Button
    Me.btnOK.Enabled = success
    
    Call ClearStatusBarDelayed(StatusBarClearDelay)
    
    Exit Sub
Catch:
    success = False
    FehlerNachricht "frmBatchPDF.RefreshUI()"
End Sub


' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
