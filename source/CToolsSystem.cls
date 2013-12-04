'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2004-2009  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'====================================================================================
'Klasse CToolsSystem
'====================================================================================
'
' ACHTUNG: Windows 98 wird nicht unterstützt!!!
'
' Stellt Werkzeuge im folgenden Kontext zur Verfügung:
'  - Betriebssystem und Registry
'  - Externe Programme
'  - Dateien
'  - Dateibezogene Fehlerliste
' 
' Voraussetzungen: 1. korrekt installiertes Windows Scripting v5.6.
'                  2. oKonfig.JEDIT_HOME liefert jEdit-Programmverz.
' 
' 
' Beschreibung:
 '  
 '  Eigenschaften (G=Get  L=Let):
  '  
  '  G  VBAHostName         ... Name des VBA-Host, z.B. "Microsoft Excel"
  '  G  VBAHostVersion      ... Version des VBA-Host, z.B. "11.0"
  '  
  '  G  OS.Type             ... derzeit wird nur "Windows_NT" unterstützt!
  '  G  OS.SystemDrive      ... %SystemDrive%
  '  G  OS.SystemRoot       ... %SystemRoot%
  '  G  OS.WindowsSys       ... %SystemRoot%\System32
  '  
  '  G  Username            ... Name des Windows-Benutzers
  '  G  Computername        ... Name des PC
  '  G  UserDomain          ... Name der Domäne bzw. des PC
  '  
  '  GL ArbeitsVerz         ... Pfad\Name des Arbeitsverzeichnisses.
  '  G  DezimalTrenner      ... Dezimaltrennzeichen (lt. System-Ländereinstellungen).
  '  G  ListenTrenner       ... Listentrennzeichen  (lt. System-Ländereinstellungen).
  '  G  LangID              ... Language ID (dezimal, lt. System-Ländereinstellungen).
  '                         
  '  G  UEditExe            ... Pfad\Dateiname der exe von UEdit, falls installiert. 
  '  G  CEditExe            ... Pfad\Dateiname der exe von Crimson Editor, falls installiert. 
  '  G  JavaExe             ... Pfad\Dateiname der exe von Java, falls installiert.
  '  G  JAVA_HOME           ... Stammverzeichnis der zu verwendenden Java-Platform (Umgebungsvariable)
  '  G  JEDIT_HOME          ... Verzeichnis, in dem jEdit.jar gefunden wurde.
  '                             Gesucht wird unter %JEDIT_HOME%, dann unter oKonfig.JEDIT_HOME
  '                             (aus Konfigurationsdatei).
  '  G  JEDIT_SETTINGS      ... jEdit-Einstellungsverzeichnis (Umgebungsvariable)
  '  G  jEditJar            ... Pfad\Dateiname von jEdit.jar, falls gefunden.
  '  G  Editoren            ... Liste unterstützter Editoren (2d-Array als Variant).
  '  GL Editor              ... zu verwendender Editor                                     
  '  G  cKeinEditor         ... "Konstante": Wenn "Editor" diesen Wert hat, so ist kein Editor verfügbar.
 '
 '  Methoden
  '  ==> (mit "#" gekennzeichnete Methoden gibt's noch nicht!)
  '  
  '  isDatei                ... Prüft, ob die angegebene Datei existiert.
  '  isVerzeichnis          ... Prüft, ob das angegebene Verzeichnis existiert.
  '  
  '  GetTmpDateiPfadName    ... Ermittelt kompletten Dateinamen für eine temporäre Datei.
  ' #LastBackslashDelete    ... Zeichenkette ohne Backslash am Ende.
  ' #NameExt                ... Gibt aus einem Pfad den Dateinamen ohne Pfad zurück ( mit|ohne Ext.).
  '  Dateiliste             ... Einfache Dateiliste eines Verzeichnisses.
  '  FindFiles              ... Dateisuche mit Wildcards - mehrere Dateimasken, mehrere Verzeichnisse.
  '  FindFile               ... Sucht nach einer einzigen Datei in mehreren Verzeichnissen.
  '  GetDateiInhalt         ... Liest eine Datei vollständig in einen String ein.
  '  OpenTextFile           ... Datei-Öffnen mit Fehlerbehandlung.
  '  getFileNameFromDialog  ... Zeigt einen System-Dateidialog ...
  '                         
  '  StarteDatei            ... Startet Datei mit der verknüpften Anwendung.
  '  StartUEdit             ... Start von UltraEdit mit den angegebenen Argumenten.
  '  StartCEdit             ... Start von Crimson Editor mit den angegebenen Argumenten.
  '  StartJEdit             ... Start von jEdit mit den angegebenen Argumenten.
  '  StartEditor            ... Start des (in Eigenschaft "Editor") eingestellten Editors ...
  '                         
  '  FileErrorsAdd          ... Nimmt einen Fehler in die dateibezogene Fehlerliste auf
  '  FileErrorsShowInJEdit  ... Fehlerliste in jEdit's Fehlerliste anzeigen.
  '  
  '  GetExeAusRegistry      ... Extrahiert Pfad\Dateiname aus einem angegebenen RegistryKey.
  '  RegRead                ... Key oder Wert aus Registry lesen
  '  RegKeyExists           ... Prüft, ob ein Registry-Schlüssel existiert.
  '  RegValueExists         ... Prüft, ob ein Registry-Wert existiert.
  ' #RegGetSubKeys          ... Ermittelt alle direkten Unterschlüssel eines Registry-Keys - nur 1 Ebene.
  ' #RegZweigDelete         ... Löscht einen RegistryKey mit allen Unterschlüsseln.
  ' #RegZweigDelete_2       ... Löscht einen RegistryKey mit allen Unterschlüsseln - 2. Variante.
  '  RegistriereWert        ... Registriert beliebigen Wert in der Registry.
  '
'
' Historie:
 '16.11.2008  - Klasse neu angelegt: enthält einige Routinen aus mdlToolsScripting.bas
 '              und bisher nicht verfügbare Routinen (aus Tools_1.vbi und Global.vbi)
 '03.12.2008  - Bugfix in GetJEditUmgebung(): Absturz, wenn Konfigurationsobjekt nicht existierte.
 '17.03.2009  - Zugriff auf das VBProject vermieden, da er eventuell nicht erlaubt ist.
 '            - Zugriff auf ThisWorkbook vermieden wegen Host-Unabhängigkeit.
 '31.05.2009  - GetTmpDateiPfadName() gibt jetzt korrekten Pfad für temp. Datei zurück.
 '30.10.2009  - Funktionen für Dateisuche: FindFiles(), FindFile() auf VbScript-Basis.
 '            - Unterstützung für Crimson Editor in Editorliste und StartEditor.
 '07.11.2009  - Editor-Unterstützung entsprechend Tools_1.vbi aktualisiert
 '              - Suchpfade für jEdit und Crimson Editor:
 '                - G:\Tools\
 '                - %PROGRAMFILES%\
 '                - %PROGRAMFILES%\Tools\
 '                - Die Unterverzeichnisse heißen: "jEdit" und "Crimson Editor"
 '                - Für jEdit wird außerdem oKonfig.JEDIT_HOME gefragt
 '              - StartJEdit() verwendet: - Kommandozeilenschalter "-Dawt.useSystemAAFontSettings=on"
 '              - Javaw.exe wird gesucht: 1. Umgebungsvariable %JAVA_HOME%
 '                                        2. Windows Systemverzeichnis (system32)
 '                                        3. Registry (Standardanwendung von .jar)
 '            - rudimentäre OS-Info (Property "OS")
 '            - FindFiles() zeigt Aktivität in Statuszeile
'====================================================================================

Option Explicit

'Konstanten
Const   VBAHostNameExcel       as String = "Microsoft Excel"
Const   RegValueEditor         as String = "HKCU\Software\VB and VBA Program Settings\Common\Editor"    'Registry-Value für aktuell zu verwendenden Editor
Public  cKeinEditor            as String   'Quasi-Konstante (Eine Konstante kann nich Public sein...)

'Eigenschaften
Public  VBAHostName            As String
Public  VBAHostVersion         As String
'Public  VBAProjectName         As String
'Public  VBAProjectFileName     As String

Public  Username               As String
Public  Computername           As String
Public  UserDomain             As String

Public  JavaExe                As String
Public  JAVA_HOME              As String
Public  JEDIT_HOME             As String
Public  JEDIT_SETTINGS         As String
Public  jEditJar               As String
Public  UEditExe               As String
Public  CEditExe               As String
Public  ListenTrenner          As String
Public  DezimalTrenner         As String
Public  LangID                 As String
Public  Editoren               As Variant

'Private Vaiablen
Private strEditor              As String

Private Type TOSinfo
  Type                         As String
  SystemDrive                  As String
  SystemRoot                   As String
  WindowsSys                   As String
End Type
Private OSinfo                 As TOSinfo

'Objekte
Private WshShell               As IWshRuntimeLibrary.WshShell
Private fs                     As Scripting.FileSystemObject
Private oFileErrors            As Scripting.Dictionary
'Private oRegExp                As Object  'existiert als globales Objekt


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
  flags              As Long
  nFileOffset        As Integer
  nFileExtension     As Integer
  lpstrDefExt        As String
  lCustData          As Long
  lpfnHook           As Long
  lpTemplateName     As String
End Type
'


Private Sub Class_Initialize()
  '
  'On Error GoTo Fehler
  On Error GoTo 0
  DebugEcho vbnewline & "CToolsSystem.Class_Initialize() beginnt."
  
  set WshShell    = New IWshRuntimeLibrary.WshShell
  set fs          = New Scripting.FileSystemObject
  set oFileErrors = New Scripting.Dictionary
  'Set oRegExp     = New RegExp
  
  '"Konstanten"
  cKeinEditor = "KeinEditor"
  
  'Systemumgebung
  call GetVBAUmgebung()
  call GetSystemUmgebung()
  call GetNetzUmgebung()
  call GetInternational()
  
  'Ermitteln diverser Pfade installierter Anwendungen.
  UEditExe = GetUEditExe()
  CEditExe = GetCEditExe()
  JavaExe  = GetJavaExe()
  call GetJEditUmgebung()
  
  'Liste unterstützter Editoren aufbauen, Standardeditor setzen.
  call InitEditoren()
  
  DebugEcho "*** CToolsSystem - Aktuelle Eigenschaften:"
  DebugEcho "UserDomain = '" & UserDomain & "',  Computername = '" & Computername & "',  Username = '" & Username & "'"
  DebugEcho "VBA-Host:    Name = '" & VBAHostName & "',  Version = '" & VBAHostVersion & "'"
  'DebugEcho "VBA-Projekt: Name = '" & VBAProjectName & "',  Dateipfad = '" & VBAProjectFileName & "'"
  DebugEcho  "Language ID= '" & LangID & "'  DezimalTrenner= '" & DezimalTrenner & "'  ListenTrenner = '" & ListenTrenner & "'."
  DebugEcho "JAVA_HOME = '" & JAVA_HOME & "'."
  DebugEcho "JEDIT_HOME = '" & JEDIT_HOME & "'  JEDIT_SETTINGS = '" & JEDIT_SETTINGS & "'."
  DebugEcho "Exe Java = '" & JavaExe & "'."
  DebugEcho "Jar jEdit = '" & jEditJar & "'."
  DebugEcho "Exe UltraEdit = '" & UEditExe & "'."
  DebugEcho "Exe Crimson Editor = '" & CEditExe & "'."
  DebugEcho "Aktueller Editor = '" & strEditor & "'."
  DebugEcho "CToolsSystem.Class_Initialize() beendet." & vbnewline
  Exit Sub

Fehler:
  FehlerNachricht "CToolsSystem.Initialize()"
  set WshShell    = nothing
  set fs          = nothing
  set oFileErrors = nothing
  'Set oRegExp     = nothing
End Sub


Private Sub Class_Terminate()
  set WshShell    = nothing
  set fs          = nothing
  set oFileErrors = nothing
  'Set oRegExp     = nothing
End Sub



'==== Eigenschaften ===========================================================

property get Editor()
  Editor = strEditor
end property

property let Editor(byVal inpEditor)
  'Setzt den angegebenen Editor aktiv und speichert diesen in der Registry,
  'falls inpEditor eine gültige Kennung für einen verfügbaren Editor ist.
  dim i, Erfolg, vName, vInhalt, vTyp
  Erfolg = false
  for i = 0 to ubound(Editoren, 1)
    if (Editoren(i, 1) = inpEditor) then
      strEditor = inpEditor
      'Editor merken
      vName   = RegValueEditor
      vInhalt = strEditor
      vTyp    = "REG_SZ"
      call RegistriereWert(vName, vInhalt, vTyp)
      DebugEcho "Editor gesetzt auf '" & inpEditor & "'."
      Erfolg = true
      exit for
    end if
  next
  if (not Erfolg) then
    DebugEcho "Editor konnte nicht gesetzt werden auf '" & inpEditor & "'."
  end if
end property

property get ArbeitsVerz()
  ArbeitsVerz = WshShell.CurrentDirectory
end property

property let ArbeitsVerz(inpVerz)
  if (fs.FolderExists(inpVerz)) then
    WshShell.CurrentDirectory = inpVerz
    debugecho "CToolsSystem.ArbeitsVerz(Let): Arbeitsverzeichnis gesetzt auf '" & inpVerz & "'."
  else
    debugecho "CToolsSystem.ArbeitsVerz(Let): Arbeitsverzeichnis setzen: Verzeichnis '" & inpVerz & "' existiert nicht."
  end if
end property


'==== Methoden ================================================================

'***  Abteilung Programmaufrufe  ***************************************************************

Function StarteDatei(ByVal DateiPfadName As String)
  'Startet die mit "DateiPfadName" verknüpfte Anwendung und lädt die Datei.
  On Error GoTo Fehler
  Const bWaitOnReturn = False
  Const intWindowStyle = 5     'Aktiviert das Fenster und zeigt es in der aktuellen Größe und Position an.
  wshShell.Run """" & DateiPfadName & """", intWindowStyle, bWaitOnReturn
  Exit Function
  
Fehler:
  ErrMessage = "Die Datei '" & DateiPfadName & "' oder die mit ihr verknüpfte Anwendung konnte nicht gefunden werden."
  FehlerNachricht "StarteDatei()"
End Function


Function StartEditor(byVal strArgumente as String)
  dim Erfolg
  select case Editor
    case "UltraEdit":       Erfolg = StartUEdit(strArgumente)
    case "Crimson Editor":  Erfolg = StartCEdit(strArgumente)
    case "jEdit":           Erfolg = StartJEdit(strArgumente)
    case cKeinEditor:       ErrEcho "Es ist kein Editor verfügbar"
    case else:              ErrEcho "CToolsSystem.StartEditor() -  Programmfehler: '" & Editor & "' ist keine gültiger Editor!"
                            Erfolg = false
  end select
  StartEditor = Erfolg
End Function


Function StartUEdit(byVal strArgumente as String)
  'Start von UltraEdit, der die angegebene Datei lädt.
  dim Kommando, Erfolg
  Erfolg = false
  if (UEditExe = "") then
    ErrEcho "StartUEdit(): UltraEdit ist nicht installiert."
  else
    Kommando = """" & UEditExe & """ " & strArgumente
    DebugEcho "Kommando wird ausgeführt: '" & Kommando & "'."
    On Error resume next
    call wshShell.Run(Kommando, WindowStyle_normal, WaitOnReturn_no)
    if (err = 0) then
      WshShell.AppActivate "UltraEdit"
      Erfolg = true
    else
      ErrEcho "StartUEdit(): UltraEdit konnte nicht gestartet werden."
    end if
    On Error GoTo 0
  end if
  StartUEdit = Erfolg
end Function

Function StartCEdit(byVal strArgumente as String)
  'Start von Crimson Editor, der die angegebene Datei lädt.
  dim Kommando, Erfolg
  Erfolg = false
  if (CEditExe = "") then
    ErrEcho "StartCEdit(): Crimson Editor ist nicht installiert."
  else
    Kommando = """" & CEditExe & """ " & strArgumente
    DebugEcho "Kommando wird ausgeführt: '" & Kommando & "'."
    On Error resume next
    call wshShell.Run(Kommando, WindowStyle_normal, WaitOnReturn_no)
    if (err = 0) then
      WshShell.AppActivate "Crimson Editor"
      Erfolg = true
    else
      ErrEcho "StartCEdit(): Crimson Editor konnte nicht gestartet werden."
    end if
    On Error GoTo 0
  end if
  StartCEdit = Erfolg
end Function


Function StartJEdit(byVal strArgumente as String) As Boolean
  'Start von jEdit mit den angegebenen Argumenten.
  'Rückgabe: true bei Erfolg, sonst false.
  dim Kommando, ArbVerzSich
  dim Erfolg   As Boolean
  Erfolg = false
  if (jEditJar = "") then
    ErrEcho "StartjEdit(): jEdit.jar wurde nicht gefunden."
  else
    if (JavaExe = "") then
      ErrEcho "StartjEdit(): Die Java-Laufzeitumgebung ist nicht (korrekt) installiert!"
    else
      ArbVerzSich = Me.ArbeitsVerz
      Me.ArbeitsVerz = JEDIT_HOME
      DebugEcho "Arbeitsverzeichnis gesetzt auf " & Me.ArbeitsVerz
      Kommando = """" & JavaExe & """-ms128m -mx1024m -Dawt.useSystemAAFontSettings=on -Dsun.java2d.noddraw=true -jar """ & jEditJar & """ -reuseview -background "
      
      '-Xms<size> set initial Java heap size
      '-Xmx<size> set maximum Java heap size
      '-Xss<size> set java thread stack size
      
      if (JEDIT_SETTINGS <> "") then Kommando = Kommando & """-settings=" & JEDIT_SETTINGS & """ "
      DebugEcho "jEdit-Einstellungsverzeichnis gesetzt auf " & JEDIT_SETTINGS
      
      Kommando = Kommando & strArgumente
      DebugEcho "Kommando wird ausgeführt: '" & Kommando & "'."
      On Error resume next
      call wshShell.Run(Kommando, WindowStyle_minimized, WaitOnReturn_no)
      if (err = 0) then
        WshShell.AppActivate "jEdit"
        Erfolg = true
      else
        ErrEcho "StartjEdit(): jEdit konnte nicht gestartet werden."
      end if
      On Error GoTo 0
      Me.ArbeitsVerz = ArbVerzSich
    end if
  end if
  StartJEdit = Erfolg
end Function



'***  Abteilung Dateien  ***********************************************************************

Function GetTmpDateiPfadName() as String
  ' ermittelt kompletten Dateinamen für eine temp. Datei
  On Error GoTo Fehler
  
  'Dim tmpVerz  as Object
  'Dim tmpDatName, tmpDatPfadName
  Const TempOrdner = 2
  
  'Set tmpVerz = fs.GetSpecialFolder(TempOrdner)
  'tmpDatName  = fs.GetTempName
  GetTmpDateiPfadName = fs.GetSpecialFolder(TempOrdner).Path & "\" & fs.GetTempName
  'GetTmpDateiPfadName = tmpVerz.drive & "\" & tmpVerz.Name & "\" & tmpDatName
  'Set tmpVerz = Nothing
  
  Exit Function

Fehler:
  FehlerNachricht "CToolsSystem.GetTmpDateiPfadName()"
End Function


Function GetDateiInhalt(ByVal DateiName As String) As String
  'Liest die angegebene Datei vollständig ein und gibt den Inhalt als String zurück.
  'Ist kein Pfad angegeben, wird das Arbeitsverzeichnis verwendet.
  On Error goto 0
  DebugEcho "CToolsSystem.GetDateiInhalt(): Datei '" & DateiName & "' einlesen."
  DateiName = fs.GetAbsolutePathName(DateiName)
  if (not fs.FileExists(DateiName)) then
    ErrEcho "Datei '" & DateiName & "' existiert nicht."
  else
    On Error Resume Next
    GetDateiInhalt = fs.OpenTextFile(DateiName).ReadAll
    if (err <> 0) then 
      ErrEcho "Fehler beim Lesen der existierenden Datei '" & DateiName & "'."
    else
      Echo "Datei erfolgreich gelesen:  '" & DateiName & "'."
    end if
  end if
  On Error goto 0
End Function


Function OpenTextFile(ByVal FileName as String, ByVal FileReadWrite, ByVal NewFileIfNotExist, ByVal Format) As Scripting.TextStream
  'Datei-Öffnen mit Fehlerbehandlung.
  'Eingabe:  Parameter für fs.OpenTextFile
  'Rückgabe: Scripting.TextStream-Objekt oder nothing.
  Dim oTS               As Scripting.TextStream
  Dim oModeReadWrite    As Scripting.Dictionary
  Dim Erfolg            As Boolean
  Dim AbsoluteFileName  As String
  
  set oModeReadWrite = New Scripting.Dictionary
  oModeReadWrite.add 1, "Lesen"
  oModeReadWrite.add 2, "Schreiben"
  oModeReadWrite.add 8, "Anhängen"
  
  DebugEcho "CToolsSystem.OpenTextFile(): Datei '" & FileName & "' soll zum " & oModeReadWrite(FileReadWrite) & " geöffnet werden."
  AbsoluteFileName = fs.GetAbsolutePathname(trim(FileName))
  DebugEcho " - absoluter Dateiname = '" & AbsoluteFileName & "'."
  
  if ((FileReadWrite = ForReading) and (not fs.FileExists(AbsoluteFileName))) then
    ErrEcho " ==> Datei '" & AbsoluteFileName & "' existiert nicht!"
    set oTS = nothing
    Erfolg  = False
  else
    DebugEcho " - Datei '" & AbsoluteFileName & "' existiert. Versuche, sie zu öffnen."
    on error resume next
    Set oTS = fs.OpenTextFile(AbsoluteFileName, FileReadWrite, NewFileIfNotExist, Format)
    if (err.number = 0) then
      Erfolg = true
      DebugEcho " - Datei '" & AbsoluteFileName & "' erfolgreich geöffnet."
    else
      ErrEcho " ==> Datei '" & AbsoluteFileName & "' konnte nicht zum " & oModeReadWrite(FileReadWrite) & " geöffnet werden!"
      set oTS = nothing
      Erfolg  = False
    end if
    on error goto 0
  end if
  set oModeReadWrite = nothing
  set OpenTextFile = oTS
End Function


Public Function FindFiles(ByVal FileFilterList As String, ByVal FolderList As String, ByVal blnSubfolders As Boolean) As Scripting.Dictionary
  '-------------------------------------------------------------------------------------------
  'Sucht Dateien, deren Namen mit einer Maske der "FileFilterList" übereinstimmen
  'und die in einem Verzeichnis der "FolderList" gefunden wurden.
  ' Argumente:
  '   "FileFilterList"  = Liste mit Dateimasken, durch Semikolons getrennt (ohne Pfadangabe, mit Wildcards)
  '   "FolderList"      = Verzeichnisliste, durch Semikolons getrennt (relative Angaben erlaubt)
  '   "blnSubfolders"   = Unterverzeichnisse durchsuchen?
  '
  ' Rückgabe            = gefundene Dateien als Keys in einem Dictionary (Item = Dateiname ohne Pfad)
  '-------------------------------------------------------------------------------------------
  'Diese Funktion ist nur ein Wrapper für die rekursive Funktion FindFiles_R(). 
  'Sie bietet in erster Linie einen Punkt, an dem die Gesamtsuche beendet ist...
  '-------------------------------------------------------------------------------------------
  
  On Error GoTo Fehler
  dim StartTime    As Single
  Dim oFoundFiles  As Scripting.Dictionary
  DebugEcho vbNewLine & "FindFiles() startet Dateisuche ..."
  StartTime = Timer()
  
  Set oFoundFiles = FindFiles_R(FileFilterList, FolderList, blnSubfolders)
  
  call WriteStatusBar( "Dateisuche beendet: Anzahl gefundener Dateien = " & cStr(oFoundFiles.Count))
  DebugEcho vbNewLine & "FindFiles(): Dateisuche dauerte insgesamt " & cStr(Timer - StartTime) & " Sekunden."
  set FindFiles = oFoundFiles
  
  Exit Function
Fehler:
  FehlerNachricht "CToolsSystem.FindFiles()"
End Function


Public Function FindFile(ByVal FileFilter As String, ByVal FolderList As String, ByVal blnSubfolders As Boolean) As String
  'Gibt den vollständigen Namen incl. Pfad derjenigen Datei zurück, deren Name mit
  '"FileFilter" übereinstimmt und die zuerst in einem Verzeichnis der "FolderList"
  'gefunden wurde.
  'Existiert keine entsprechende Datei, wird "" zurückgegeben.
  '   "FolderList"    = Verzeichnisliste, durch Semikolons getrennt (relative Angaben erlaubt)
  '   "FileFilter"    = Dateimaske  (ohne Pfadangabe, mit Wildcards)
  '   "blnSubfolders" = Unterverzeichnisse durchsuchen (ja/nein)
  
  On Error goto 0
  Dim FilePath        As String
  Dim oFoundFiles     As Scripting.Dictionary
  
  DebugEcho vbNewLine & "FindFile() sucht nach einer Datei ..."
  
  Set oFoundFiles = FindFiles_R(FileFilter, FolderList, blnSubfolders)
  
  if (oFoundFiles.count > 0) then
    FilePath = oFoundFiles.Keys(0)
    DebugEcho vbNewLine & "FindFile() hat gefunden: " & FilePath
  Else
    FilePath = ""
  End If
  
  On Error GoTo 0
  FindFile = FilePath
End Function


Private Function FindFiles_R(ByVal FileFilterList As String, ByVal FolderList As String, _
                            ByVal blnSubfolders As Boolean, _
                            Optional ByVal blnNewSearch As Boolean = true) As Scripting.Dictionary
  '-------------------------------------------------------------------------------------------
  'Sucht Dateien, deren Namen mit einer Maske der "FileFilterList" übereinstimmen
  'und die in einem Verzeichnis der "FolderList" gefunden wurden.
  ' Argumente:
  '   "FileFilterList"  = Liste mit Dateimasken, durch Semikolons getrennt (ohne Pfadangabe, mit Wildcards)
  '   "FolderList"      = Verzeichnisliste, durch Semikolons getrennt (relative Angaben erlaubt)
  '   "blnSubfolders"   = Unterverzeichnisse durchsuchen?
  '   "blnNewSearch"    = Initialisierung einer neuen Suche?
  '                       Wenn "false", wird die intern vorhandene Dateiliste erweitert!
  '
  ' Rückgabe            = gefundene Dateien als Keys in einem unsortierten Dictionary (Item = Dateiname ohne Pfad)
  '-------------------------------------------------------------------------------------------
  
  On Error GoTo Fehler
  ErrMessage = "unspezifischer Fehler bei Dateisuche ..."
  
  'ACHTUNG: Statische Deklaration:
  Static oFoundFiles          As Scripting.Dictionary
  Static StartTime            As Single
  Static CountFoundFiles      As Long
  
  'Deklarationen
    Dim CountFolders          As Long
    Dim CountFilters          As Long
    Dim i                     As Long
    Dim j                     As Long
    Dim Folders()             As String
    Dim FileFilters()         As String
    Dim FileFilters_RE()      As String
    Dim FileFilterName        As String
    Dim AbsolutePath          As String
    Dim SubFolderList         As String
    Dim oSubFolders           As Scripting.Folders
    Dim oSubFolder            As Scripting.Folder
    Dim oFolder               As Scripting.Folder
    Dim oFiles                As Scripting.Files
    Dim oFile                 As Scripting.File
    Dim oFS                   As New Scripting.FileSystemObject
  
  'Init
  if (blnNewSearch) then
    StartTime = Timer()
    DebugEcho vbNewLine & "FindFiles_R() startet Dateisuche ..."
    set oFoundFiles = New Scripting.Dictionary
    CountFoundFiles = 0
  elseif (oFoundFiles is Nothing) then
    Err.Raise 99991 + vbObjectError, , "Programmfehler: Dateisuche ist noch nicht initialisiert!"
  end if
  const DoEventsInterval = 100
  DebugEcho vbNewLine & "Verzeichnisliste:     '" & FolderList & "'"
  DebugEcho "Dateimaskenliste:     '" & FileFilterList & "'"
  
  'Die übergebenen Listen in Array's wandeln.
  CountFilters = SplitDelim(FileFilterList, FileFilters, ";", true, true)
  CountFolders = SplitDelim(FolderList, Folders, ";", true, true)
  
  if (CountFilters < 1) then
    ErrEcho "FindFiles_R(): Keine Dateimaske angegeben!"
  elseif (CountFolders < 1) then
    ErrEcho "FindFiles_R(): Kein Verzeichnis angegeben!"
  else
    'Reguläre Ausdrücke für Dateimasken vorbereiten
    redim FileFilters_RE(1 to CountFilters)
    For j = 1 To CountFilters
      FileFilters_RE(j) = "^" & FileSpec2RegExp(FileFilters(j)) & "$"
      DebugEcho cStr(j) & ". Dateimaske als Reg. Ausdruck:  '" & FileFilters_RE(j) & "'"
    next
    
    'Jedes Verzeichnis der Liste durchsuchen.
    For i = 1 To CountFolders
      AbsolutePath = oFS.GetAbsolutePathName(Folders(i))
      
      if (not oFS.FolderExists(AbsolutePath)) then
        WarnEcho "FindFiles_R(): Verzeichnis für Dateisuche existiert nicht: '" & AbsolutePath & "'."
      else
        DebugEcho vbNewline & "Suche in Verzeichnis: '" & AbsolutePath & "'"
        set oFolder = oFS.GetFolder(AbsolutePath)
        set oFiles  = oFolder.Files
        
        'Jede Datei des Verzeichnisses verarbeiten.
        for Each oFile in oFiles
          'Bedienoberfläche am Leben halten
          CountFoundFiles = CountFoundFiles + 1
          if ((CountFoundFiles mod DoEventsInterval) = 0) then
            call WriteStatusBar("Dateisuche läuft... (bisher gefunden: " & cStr(CountFoundFiles) & ")")
            DoEvents
          end if
          
          'Dateiname (ohne Pfad) gegen jede Dateimaske prüfen.
          For j = 1 To CountFilters
            if (entspricht(FileFilters_RE(j), oFile.Name)) then
              if (not oFoundFiles.exists(oFile.Path)) then
                oFoundFiles.Add oFile.Path, oFile.Name
                DebugEcho "=> " & cStr(oFoundFiles.count) & ". Datei gefunden:   '" & oFile.Path & "'"
              end if
              exit for
            end if
          next
        next
        
        'Unterverzeichnisse durchsuchen.
        if (blnSubfolders) then
          'Debugecho vbNewLine & "Unterverzeichnisse von " & oFolder.Path & ":"
          set oSubFolders = oFolder.SubFolders
          if (oSubFolders.count > 0) then
            'Debugecho ListeAuflistung(oSubFolders)
            SubFolderList = ""
            for each oSubFolder in oSubFolders
              SubFolderList = SubFolderList & ";" & oSubFolder.path
            next
            SubFolderList = mid(SubFolderList, 2)
            set oFoundFiles = FindFiles_R(FileFilterList, SubFolderList, blnSubfolders, false)
          end if
        end if
      end if
    next
  end if
  
  'Aufräumen
    set oSubFolders = Nothing
    set oSubFolder  = Nothing
    set oFolder     = Nothing
    set oFiles      = Nothing
    set oFile       = Nothing
    set oFS         = Nothing
    'set oFoundFiles = Nothing ... sonst macht "Static" keinen Sinn
  '
  DebugEcho vbNewLine & "FindFiles_R(): Dateisuche dauerte bisher " & cStr(Timer - StartTime) & " Sekunden."
  set FindFiles_R = oFoundFiles
  
  Exit Function
Fehler:
  FehlerNachricht "CToolsSystem.FindFiles_R()"
End Function


Function DateiListe(byVal Verzeichnis, byVal Separator) As String
  'Funktionswert = String, der alle Dateinamen des angegebenen Verzeichnisses
  'enthält - getrennt durch "Separator".
  On Error GoTo Fehler
  Dim Verz, Datei, Dateien, s
  Set Verz = fs.GetFolder(Verzeichnis)
  Set Dateien = Verz.Files
  For Each Datei In Dateien
     s = s & Datei.Name
     s = s & Separator
  Next
  DateiListe = s
  Set Verz = Nothing
  Set Dateien = Nothing
  Exit Function
  
Fehler:
  FehlerNachricht "CToolsSystem.Dateiliste()"
End Function

Public Function isDatei(ByVal Datei As String) As Boolean
  'Gibt "true" zurück, wenn die angegebene Datei existiert.
  '"Datei" kann mit absolutem oder relativen bzw. ohne Pfad angegeben sein.
  isDatei = fs.FileExists(fs.GetAbsolutePathName(Datei))
  Exit Function
Fehler:
  isDatei = False
  'FehlerNachricht "CToolsSystem.isDatei()"
End Function


Public Function isVerzeichnis(ByVal Verzeichnis As String) As Boolean
  'Gibt "true" zurück, wenn das angegebene Verzeichnis existiert.
  '"Verzeichnis" kann mit absolutem oder relativen bzw. ohne Pfad angegeben sein.
  isVerzeichnis = fs.FolderExists(fs.GetAbsolutePathName(Verzeichnis))
  Exit Function
Fehler:
  isVerzeichnis = False
  'FehlerNachricht "CToolsSystem.isVerzeichnis()"
End Function


Public Function getFileNameFromDialog( _
  Optional ByVal OpenFile          As Boolean = true, _
  Optional ByVal DialogTitle       As String, _
  Optional ByVal InitialFilename   As String = "", _
  Optional ByVal InitialDirectory  As String, _
  Optional ByVal FileFilter        As String = "All files (*.*),*.*", _
  Optional ByVal FileFilterIndex   As Long = 1, _
  Optional ByVal Flags             As Long = 0) as String
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
  Const MAX_PATH_LENGTH  as Integer = 255
  
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
    Case VBAHostNameExcel
      'Excel: When the general system dialog is shown and the user presses ESCAPE, 
        '     then the VBA editor will be shown in debugging mode. That's why using Excel's own dialog.
        'CAUTION: This dialog does not provide a warning related to file overwrite!
        
        set HostApp = Application  'this is for getting the VBProject compiled under other hosts than Excel
        const msoFileDialogSaveAs = 2
        If (OpenFile) Then
          FilePath = HostApp.GetOpenFilename(FileFilter, FileFilterIndex, DialogTitle)
        else
          FilePath = HostApp.getSaveAsFilename(InitialFilename, FileFilter, FileFilterIndex, DialogTitle)
        end if
        
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
          if (isMissing(DialogTitle)) then DialogTitle = vbNullChar
          if (Right(InitialFilename, 1) = "\") then InitialFilename = ""
          if (isMissing(InitialDirectory)) then InitialDirectory = CurDir()
          FileFilter = replace(FileFilter, ",", vbNullChar) & vbNullChar
          If (OpenFile) Then
            'open
            Flags = Flags or OFN_FILEMUSTEXIST Or OFN_READONLY  'resp. HideReadOnly
          Else
            'save
            Flags = Flags or OFN_OVERWRITEPROMPT
          End If
        
        'Fill OPENFILENAME structure for api call
          OFN_Structure.lStructSize = Len(OFN_Structure)
          With OFN_Structure
            .lpstrTitle      = DialogTitle
            .lpstrFile       = InitialFilename & String(MAX_PATH_LENGTH - Len(InitialFilename), 0)
            .lpstrInitialDir = InitialDirectory
            .lpstrFilter     = FileFilter
            .nFilterIndex    = FileFilterIndex
            '.lpstrDefExt     = DefaultExt
            .flags           = Flags
            .nMaxFile        = Len(.lpstrFile) - 1
            .lpstrFileTitle  = .lpstrFile
            .nMaxFileTitle   = .nMaxFile
          End With
        
        'Show dialog
          If (OpenFile) Then
            success = GetOpenFileName(OFN_Structure)
          else
            success = GetSaveFileName(OFN_Structure)
          end if
        
        'Get the result
          If (success = 0) Then
            FilePath = ""
            'MsgBox "user cancelled or error calling api function!", vbInformation
          Else
            FilePath = Trim(Left(OFN_Structure.lpstrFile, InStr(1, OFN_Structure.lpstrFile, vbNullChar) - 1))
            'MsgBox "'" & FilePath & "'"
          End If
  End Select
  
  DebugEcho "CToolsSystem.getFileNameFromDialog: dialog returned filename '" & FilePath & "'."
  getFileNameFromDialog = FilePath
End Function



'***  Abteilung Registry  **********************************************************************

Function GetExeAusRegistry(byVal RegKey)
  'Funktionswert = Pfad\Dateiname der in RegKey registrierten und existierenden EXE oder "".
  'RegKey muß mit Backslash abgeschlossen sein, es sei denn, es ist ein RegValue!
  on error goto 0
  dim ExePfadName
  DebugEcho "Pfad\Name einer Programmdatei ermitteln aus RegistryKey '" & RegKey & "'."
  
  if (RegValueExists(RegKey)) then
    ExePfadName = RegRead(RegKey)
    DebugEcho "Inhalt des RegistryKey: '" & ExePfadName & "'"
    ExePfadName = trim(ExePfadName)
    'Platzhalter entfernen
    'ExePfadName = trim(substitute(" \""?%[0-9]\""?", "", ExePfadName, true, true))
    'Pfad\Name extrahieren:
    if (left(ExePfadName,1) = """") then
      'Pfad\Name ist in Anführungszeichen eingeschlossen.
      ExePfadName = MidStr(ExePfadName, 2, """", false)
    else
      'Pfad\Name ist NICHT in Anführungszeichen eingeschlossen
      '=> Ende bei erstem Leerzeichen, falls eines existiert, sonst gesamte Zeichenkette.
      if (instr(1, ExePfadName, " ", vbTextCompare) > 0) then
        ExePfadName = LeftStr(ExePfadName, " ", false)
      end if
    end if
    'Anführungszeichen entfernen.
    'ExePfadName = trim(substitute("""", "", ExePfadName, true, true))
  end if
  DebugEcho "Ermittelter Pfad\Name lautet: '" & ExePfadName & "'"
  if (fs.FileExists(ExePfadName)) then
    DebugEcho "OK - '" & ExePfadName & "' existiert."
  else
    DebugEcho "FEHLER - '" & ExePfadName & "' existiert nicht!"
    ExePfadName = ""
  end if
  GetExeAusRegistry = ExePfadName
End Function


Function RegKeyExists(ByVal Key)
  ' prüft, ob RegistryKey "key" existiert
  ' Rückgabewert: true oder false
  ' Voraussetzung: Objekt WSHShell
  
  ' Lesen eines fehlenden Keys:
  '    err.number      = 0x80070002
  '    err.description = Ungültige Wurzel in Registrierungsschlüssel "HKCU\blah\".
  
  ' Lesen eines vorhandenen Keys ohne Standardwert:
  '    err.number      = 0x80070002 (siehe oben!)
  '    err.description = Registrierungsschlüssel "HKCU\*\" wurde nicht zum Lesen geöffnet
  
  On Error Resume Next
  WSHShell.RegRead (UCase(Key))
  
  If (Err.Number = 0) Then
    RegKeyExists = True
  ElseIf (Left(Err.Description, 43) = "Ungültige Wurzel in Registrierungsschlüssel") Then
    RegKeyExists = False
  Else
    RegKeyExists = True
  End If
  'Message = "Key = " & Key & vbNewLine & "Fehlernr = 0x" & Hex(Err.Number) & vbNewLine & _
  '          "Fehlerbeschreibung = " & Err.Description & vbNewLine & _
  '           vbNewLine & "RegKeyExists = " & RegKeyExists
  'debugecho Message
  On Error GoTo 0
End Function


Function RegValueExists(ByVal value)
  ' prüft, ob RegistryValue "value" mit Standardeintrag existiert
  ' Rückgabewert: true oder false
  ' Voraussetzung: Objekt WSHShell
  ' endet Value mit "\", so ist der Standardeintrag des angegebenen keys gemeint
  
  ' Lesen eines fehlenden Keys:
  '    err.number      = 0x80070002
  '    err.description = Ungültige Wurzel in Registrierungsschlüssel "HKCU\blah\".
  
  ' Lesen eines vorhandenen Keys ohne Standardwert:
  '    err.number      = 0x80070002 (siehe oben!)
  '    err.description = Registrierungsschlüssel "HKCU\*\" wurde nicht zum Lesen geöffnet
  
  On Error Resume Next
  WSHShell.RegRead (UCase(value))
  
  If (Err.Number = 0) Then
    RegValueExists = True
  Else
    RegValueExists = False
  End If
  'Message = "value = " & value & vbNewLine & "Fehlernr = 0x" & Hex(Err.Number) & vbNewLine & _
  '          "Fehlerbeschreibung = " & Err.Description & vbNewLine & _
  '           vbNewLine & "RegValueExists = " & RegValueExists
  'debugecho Message
  On Error GoTo 0
End Function


Function RegRead(ByVal KeyOrValue)
  'Rückgabewert: Wert des RegistryKey "KeyOrValue", im Zweifelsfalle leer
  'Endet "KeyOrValue" mit einem Backslash, so ist es ein Key, dessen Standardwert
  'gelesen wird.
  
  On Error Resume Next
  RegRead = WSHShell.RegRead(UCase(KeyOrValue))
  On Error GoTo 0
End Function


Sub RegistriereWert(ByVal vName, ByVal vInhalt, ByVal vTyp)
  'Registriert beliebigen Wert in der Registry.
  On Error GoTo Fehler
  'msgbox  vName & "=" & vInhalt & "(" & vTyp & ")"
  WSHShell.RegWrite vName, vInhalt, vTyp
  Exit Sub
Fehler:
  FehlerNachricht "RegistriereWert()"
End Sub



'***  Abteilung Fehlermanagement (Fehler sammeln und in jEdit anzeigen)  ***********************

public sub FileErrorsAdd(byVal Typ as String, _
                         byVal Datei as String, _
                         byVal Zeile as Long, _
                         byVal SpAnf as Long, _
                         byVal SpEnd as Long, _
                         byVal Meldung as String, _
                         Optional byVal ExtraZeilen As String = "", _
                         Optional byVal Konsolidieren As Boolean = true)
  '-------------------------------------------------------------------------------
  'Nimmt einen Fehler in die dateibezogene Fehlerliste auf
  'Eingabe:  Typ           ... "F" = Fehler, "W" = Warnung
  '          Datei         ... Pfad\Name der Datei, auf die sich der Fehler bezieht
  '          Zeile         ... Nummer der Zeile, auf die sich der Fehler bezieht
  '          SpAnf         ... Eingrenzung des Fehlers: Startspalte (oder 0)
  '          SpEnd         ... Eingrenzung des Fehlers: Endspalte   (oder 0)
  '          Meldung       ... Fehlermeldung
  '          ExtraZeilen   ... Zusätzliche Meldungszeilen zur Fehlermeldung
  '                            (Mehrere Zeilen durch vbNewLine getrennt)
  '          Konsolidieren ... Mehrere Fehler, die sich auf die gleiche Zeile
  '                            der gleichen Datei beziehen, werden zu einem einzigen Fehler
   '                           zusammengefasst. Die Meldungen ab dem 2. Fehler werden dabei als
   '                           Extrazeilen geführt. 
   '                           - Fehler und Warnungen werden gemischt. Dabei wird immer ein
   '                             Fehler als führende Meldung eingesetzt.
   '                           - Sind die Spaltenangaben nicht identisch, so gilt: 0,0.
   '                           - Extrazeilen werden zum Schluss an die vorh. Extrazeilen gehängt.
  ' 
  ' Struktur der internen Fehlerliste siehe: FileErrorsShowInJEdit().
  '-------------------------------------------------------------------------------
  dim FileErrorsAnz, FehlerLfdNr, j
  dim Meldung_alt      as String
  dim ExtraZeilen_alt  As String
  
  dim oFehler       As Scripting.Dictionary
  dim oFehlerNeu    As Scripting.Dictionary
  set oFehlerNeu = New Scripting.Dictionary
  
  on error goto 0
  
  if (SpAnf > SpEnd) then SpAnf = 0
  
  'Neuen Fehler zusammenstellen
    oFehlerNeu.add "Typ",         ucase(Typ)
    oFehlerNeu.add "Datei",       Datei
    oFehlerNeu.add "Zeile",       Zeile
    oFehlerNeu.add "SpAnf",       SpAnf
    oFehlerNeu.add "SpEnd",       SpEnd
    oFehlerNeu.add "Meldung",     Meldung
    oFehlerNeu.add "ExtraZeilen", ExtraZeilen
  
  'Fehlerliste aktualisieren
    FileErrorsAnz = oFileErrors.count
    FehlerLfdNr   = FileErrorsAnz + 1
    'debugecho(sprintf("  FileErrorsAdd():   Datei='%s',  Zeile = %.0f, Meldung = '%s'", Datei, Zeile, Meldung))
    
    'Fehler suchen, der sich auf die gleiche Zeile der gleichen Datei bezieht.
      if (Konsolidieren) then
        for j = 1 to FileErrorsAnz
          if (oFileErrors(j)("Zeile") = Zeile) then
            if (oFileErrors(j)("Datei") = Datei) then
              'gleiche Zeile, gleiche Datei
              FehlerLfdNr = j
              'debugecho(sprintf("  FileErrorsAdd():   Entsprechender Fehler ist bereits gespeichert als lfd. Nr %.0f => wird ergänzt", FehlerLfdNr))
              exit for
            end if
          end if
        next
      end if
    
    'Fehler in die Liste einfügen / anhängen
      if (FehlerLfdNr = FileErrorsAnz + 1) then
        'Neuer Fehler
        oFileErrors.Add FehlerLfdNr, oFehlerNeu
        
      else
        'Vorhandenen Fehler ergänzen
        
        'Spaltenangaben abgleichen
        if ((oFileErrors(FehlerLfdNr)("SpAnf") <> SpAnf) or (oFileErrors(FehlerLfdNr)("SpEnd") <> SpEnd)) then
          oFileErrors(FehlerLfdNr)("SpAnf") = 0
          oFileErrors(FehlerLfdNr)("SpEnd") = 0
        end if
        
        'Typ und eigentliche Meldung behandeln
        Meldung_alt     = oFileErrors(FehlerLfdNr)("Meldung")
        ExtraZeilen_alt = oFileErrors(FehlerLfdNr)("ExtraZeilen")
        
        if ((oFehlerNeu("Typ") = "F") and (oFileErrors(FehlerLfdNr)("Typ") = "W")) then
          oFileErrors(FehlerLfdNr)("Typ") = oFehlerNeu("Typ")
          oFileErrors(FehlerLfdNr)("Meldung") = Meldung
          oFileErrors(FehlerLfdNr)("ExtraZeilen") = Meldung_alt & vbNewLine & ExtraZeilen_alt
        else
          if (ExtraZeilen_alt = "") then
            oFileErrors(FehlerLfdNr)("ExtraZeilen") = Meldung
          else
            oFileErrors(FehlerLfdNr)("ExtraZeilen") = ExtraZeilen_alt & vbNewLine & Meldung
          end if
        end if
        
        'Extrazeilen in jedem Fall als solche anhängen
        if (ExtraZeilen <> "") then
          if (oFileErrors(FehlerLfdNr)("ExtraZeilen") = "") then
            oFileErrors(FehlerLfdNr)("ExtraZeilen") = ExtraZeilen
          else
            oFileErrors(FehlerLfdNr)("ExtraZeilen") = oFileErrors(FehlerLfdNr)("ExtraZeilen") & vbNewLine & ExtraZeilen
          end if
        end if
      end if
    
  set oFehlerNeu = nothing
End sub


public function FileErrorsShowInJEdit(byVal ListeLeeren as Boolean) as Boolean
  '------------------------------------------------------------------------------------------
  'Gibt alle bisher durch FileErrorsAdd() gesammelten Fehlermeldungen
  'in der Fehlerliste des jEdit aus.
  'Eingabe:  ListeLeeren ... wenn "true", wird die Fehlerliste gelöscht.
  'Rückgabe: false, wenn jEdit offenbar nicht erfolgreich aufgerufen werden konnte, sonst true.
  ' 
  ' Voraussetzung ist die intern als Dictionary geführte Fehlerliste,
  ' wobei jeder einzelne Fehler auch ein Dictionary ist:
  '   
  '   oFileErrors.count      = Anzahl der Fehler = letzte lfdNr
  '   Key                    = lfdNr des Fehlers, beginnend mit 1
  '   Item                   = Dictionary oFehler:
  '   
  '   oFehler("Typ")         = [String] "F" oder "W"
  '   oFehler("Datei")       = [String] Datei (Pfad\Name)
  '   oFehler("Zeile")       = [Long  ] Zeile
  '   oFehler("SpAnf")       = [Long  ] Ort des Fehlers: erste Spalte 
  '   oFehler("SpEnd")       = [Long  ] Ort des Fehlers: letze Spalte
  '   oFehler("Meldung")     = [String] Meldung
  '   oFehler("ExtraZeilen") = [String] zusätzliche Meldungszeilen (vbNewLine gilt als Zeilentrenner)
  '   
  'Dazu wird ein Beanshellskript erzeugt und in jEdit ausgeführt.
  'Hinweis: Unterstreichung des Fehlers erfolgt nicht, wenn Startspalte > 0 und Endspalte = 0 !
  ' => Also entweder exakte Eingrenzung (Start- und Endspalte) oder gar nicht (beide Spalten = 0)!
  '------------------------------------------------------------------------------------------
  dim Extrazeilen, i, k
  
  dim Zeile      As Long
  dim SpAnf      As Long
  dim SpEnd      As Long
  
  dim Meldung    As String
  dim Datei      As String
  dim Typ        As String
  dim BshName    As String
  dim Quelltext  As String
  
  dim Erfolg     As Boolean
  
  dim oFehler    As Scripting.Dictionary
  Dim oBsh       As Scripting.TextStream
  
  Erfolg = true
  call DebugEcho("CToolsSystem.FileErrorsShowInJEdit(): Anzahl Fehler in der Fehlerliste=" & cStr(oFileErrors.count))
  
  if ((jEditJar = "") or (JavaExe = "")) then
    DebugEcho "CToolsSystem.FileErrorsShowInJEdit(): jEdit.jar oder java.exe nicht gefunden => Nichts zu tun."
    Erfolg = false
  else
    
    if (oFileErrors.count > 0) then
      
      'Beanshellskript anlegen
      BshName  = GetTmpDateiPfadName()
      set oBsh = OpenTextFile(BshName, ForWriting, NewFileIfNotExist_yes, OpenAsSystemDefault)
      
      'Beanshellskript beschreiben
      if (not (oBsh is nothing)) then
        
        'Kopf des Beanshell-Skriptes
          oBsh.WriteLine "// :mode=beanshell:"
          oBsh.WriteLine "import errorlist.*;"
          oBsh.WriteLine ""
          oBsh.WriteLine "void addErrorToList() {"
          oBsh.WriteLine "  "
          oBsh.WriteLine "  void run() {"
          oBsh.WriteLine "    //get a valid view at jedit's startup"
          oBsh.WriteLine "    view = jEdit.getLastView();"
          oBsh.WriteLine "    "
          oBsh.WriteLine "    //clear list of errors via the action, that also can be invoked by keyed-in"
          oBsh.WriteLine "    //in the actionbar or by clicking the appropriate button."
          oBsh.WriteLine "    jEdit.getAction(""error-list-clear"").invoke(view);"
          oBsh.WriteLine "    "
          oBsh.WriteLine "    //Create and register DefaultErrorSource"
          oBsh.WriteLine "    DefaultErrorSource errsrc = new DefaultErrorSource(""vbscript.oTools_1.FileErrorsShowInJEdit"");"
          oBsh.WriteLine "    ErrorSource.registerErrorSource(errsrc);"
          oBsh.WriteLine "    "
          oBsh.WriteLine "    "
          oBsh.WriteLine "    //*********************************************************************************************"
        
        'Körper des Beanshell-Skriptes (variabler Teil)
          For i = 1 To oFileErrors.Count
            
            set oFehler = oFileErrors(i)
            
            'Einzelinfo's des aktuellen Fehlers ermitteln
              if (oFehler("Typ") = "W") then Typ = "ErrorSource.WARNING" else Typ = "ErrorSource.ERROR"
              Datei   = oFehler("Datei")
              Zeile   = oFehler("Zeile") - 1
              SpAnf   = oFehler("SpAnf")
              SpEnd   = oFehler("SpEnd")
              Meldung = oFehler("Meldung")
              
              if (SpAnf > 0) then SpAnf = SpAnf - 1
              Datei   = String2Java(Datei)
              Meldung = String2Java(Meldung)
              
            'Neuen Fehler mit allen wichtigen Infos erzeugen
              'Quelltext = vsprintf("    DefaultErrorSource.DefaultError err = new DefaultErrorSource.DefaultError(errsrc, %s, ""%s"", %d, %d, %d, ""%s"");", _
              '                     array(, , , , ))
              Quelltext = "    DefaultErrorSource.DefaultError err = new DefaultErrorSource.DefaultError(errsrc, " & _
                          Typ & ", """ & Datei & """, " & cStr(Zeile) & ", " & cStr(SpAnf) & ", " & cStr(SpEnd) & ", """ & Meldung & """);"
                        
              oBsh.WriteLine Quelltext
            
            'Zusätzliche Meldungszeilen
              'Split message into single lines
              ExtraZeilen = Split(oFehler("ExtraZeilen"), vbNewLine)  'Solange das Argument ein String ist, wird das Ergebnis immer ein Array.
              for k = lbound(ExtraZeilen) to ubound(ExtraZeilen)
                oBsh.WriteLine "    err.addExtraMessage(""" & String2Java(ExtraZeilen(k)) & """);"
              next
            
            'Fehler einloggen
              oBsh.WriteLine "    errsrc.addError(err);" & vbNewLine
          next
          set oFehler = nothing
        
        'Fuß des Beanshell-Skriptes
          oBsh.WriteLine "    //*********************************************************************************************"
          oBsh.WriteLine "    "
          'Hilft nicht:
          'oBsh.WriteLine "    //show error list via the action, that also can be invoked by keyed-in"
          'oBsh.WriteLine "    //in the actionbar or by clicking the appropriate button."
          'oBsh.WriteLine "    jEdit.getAction(""error-list-show"").invoke(view);"
          oBsh.WriteLine "    "
          oBsh.WriteLine "    //Do not unregister - so the errors stay deleteable by the errorlist plugin itself (see above)"
          oBsh.WriteLine "    //ErrorSource.unregisterErrorSource(errsrc);"
          oBsh.WriteLine "    errsrc = null;"
          oBsh.WriteLine "  }"
          oBsh.WriteLine "  "
          oBsh.WriteLine "  //manage startup/nonstartup script"
          oBsh.WriteLine "  if (jEdit.getLastView() == null) {"
          oBsh.WriteLine "    VFSManager.runInAWTThread(this);"
          oBsh.WriteLine "  } else {"
          oBsh.WriteLine "    run();"
          oBsh.WriteLine "  }"
          oBsh.WriteLine "}"
          oBsh.WriteLine ""
          oBsh.WriteLine "addErrorToList();"
          oBsh.Close
        
        'jEdit aufrufen und Fehler anzeigen
          Erfolg = StartJEdit("-run=" & BshName)
          if (Erfolg) then
            call DebugEcho("CToolsSystem.FileErrorsShowInJEdit(): Fehler erfolgreich in die jEdit-Fehlerliste eingestellt.")
          else
            call DebugEcho("CToolsSystem.FileErrorsShowInJEdit(): Fehler beim Aufruf der jEdit-Fehlerliste.")
          end if
      end if
    end if
  end if
  
  'Aktuelle Fehlerliste leeren
  if (ListeLeeren) then oFileErrors.RemoveAll
  
  FileErrorsShowInJEdit = Erfolg
  
End function




'==== interne Routinen ========================================================

private function String2Java(byVal Text As String) As String
  'Rückgabe: Für Java-Quelltext maskierten String
  Text = replace(Text, "\", "\\")    'Backslash
  Text = replace(Text, """", "\""")  'Anführungszeichen
  String2Java = Text
End function


Private Function GetUEditExe()
  'Funktionswert = Pfad\Dateiname der registrierten und existierenden EXE von UEdit oder "".
  dim Key_UEditExe, UEditExe
  on error goto 0
  DebugEcho "Programmdatei von UltraEdit ermitteln."
  
  'Versionen 6 bis 10:
  Key_UEditExe = "HKEY_CLASSES_ROOT\Applications\UEdit32.exe\shell\edit\command\"
  UEditExe = GetExeAusRegistry(Key_UEditExe)
  if (UEditExe = "") then
    Key_UEditExe = "HKEY_CLASSES_ROOT\UltraEdit-32 Document\shell\open\command\"
    UEditExe = GetExeAusRegistry(Key_UEditExe)
  end if
  
  'Version 11
  if (UEditExe = "") then
    'Voraussetzung für die Existenz des folgenden Schlüssels ist leider
    'die von UltraEdit aus erfolgte Dateizuordnung des Typs ".txt"!!!
    Key_UEditExe = "HKEY_CLASSES_ROOT\UltraEdit.txt\shell\open\command\"
    UEditExe = GetExeAusRegistry(Key_UEditExe)
  end if
  GetUEditExe = UEditExe
End Function

Private Function GetCEditExe()
  'Funktionswert = Pfad\Dateiname der registrierten und existierenden EXE von CEdit oder "".
  dim Key_CEditExe, CEditExe, CEditVerz, Erfolg
  on error goto 0
  const ExeName = "cedt.exe"
  Erfolg = false
  DebugEcho "Programmdatei von Crimson Editor ermitteln."
  
  'Von hinten durchs Auge: Was deinstalliert werden kann, ist ja vielleicht vorhanden...
  Key_CEditExe = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Crimson Editor 3.72\DisplayIcon"
  CEditExe = GetExeAusRegistry(Key_CEditExe)
  if (fs.FileExists(CEditExe)) then
    Erfolg = true
  else
    DebugEcho "GetCEditExe(): Crimson Editor nicht gefunden in Registry => suche nach Portable Version"
  end if
  
  'Portable Version an 3 Stellen suchen
  if (not Erfolg) then
    CEditVerz = "G:\Tools\Crimson Editor"
    CEditExe  = CEditVerz & "\" & ExeName
    if (fs.FileExists(CEditExe)) then
      Erfolg = true
    else
      DebugEcho "GetCEditExe(): cedt.exe nicht gefunden im (hart kodierten) Verzeichnis '" & CEditVerz & "'"
    end if
  end if
  
  if (not Erfolg) then
    CEditVerz = Environ("PROGRAMFILES") & "\Crimson Editor"
    CEditExe  = CEditVerz & "\" & ExeName
    if (fs.FileExists(CEditExe)) then
      Erfolg = true
    else
      DebugEcho "GetCEditExe(): cedt.exe nicht gefunden im (hart kodierten) Verzeichnis '" & CEditVerz & "'"
    end if
  end if
  
  if (not Erfolg) then
    CEditVerz = Environ("PROGRAMFILES") & "\Tools\Crimson Editor"
    CEditExe  = CEditVerz & "\" & ExeName
    if (fs.FileExists(CEditExe)) then
      Erfolg = true
    else
      DebugEcho "GetCEditExe(): cedt.exe nicht gefunden im (hart kodierten) Verzeichnis '" & CEditVerz & "'"
    end if
  end if
  
  if (not Erfolg) then CEditExe = ""
  
  GetCEditExe = CEditExe
End Function


Private Function GetJavaExe()
  'Funktionswert = Pfad\Dateiname der registrierten und existierenden Javaw.exe oder "".
  'legt folgende öffentlichen Variablen fest:
  'JAVA_HOME = Home-Verzeichnis von Java
  on error goto 0
  dim Key_JavaExe, Pfad_JavaExe, Erfolg
  const ExeName = "javaw.exe"
  Erfolg = false
  Key_JavaExe = "HKEY_CLASSES_ROOT\jarfile\shell\open\command\"
  DebugEcho "Programmdatei von Java ermitteln."
  
  JAVA_HOME = LastBackslashDelete(Environ("JAVA_HOME"))
  DebugEcho "GetJavaExe(): Umgebungsvariable %JAVA_HOME%='" & JAVA_HOME & "'"
  if (JAVA_HOME <> "") then
    Pfad_JavaExe = JAVA_HOME & "\bin\" & ExeName
    if (fs.FileExists(Pfad_JavaExe)) then
      Erfolg = true
    else
      DebugEcho "GetJavaExe(): javaw.exe nicht gefunden im Verzeichnis der Umgebungsvariable %JAVA_HOME%='" & JAVA_HOME & "'"
    end if
  end if
    
  if (not Erfolg) then
    Pfad_JavaExe = OSinfo.WindowsSys & "\" & ExeName
    if (fs.FileExists(Pfad_JavaExe)) then
      Erfolg = true
    else
      DebugEcho "GetJavaExe(): javaw.exe nicht gefunden im Windows Systemverzeichnis '" & OSinfo.WindowsSys & "'"
    end if
  end if
  
  if (not Erfolg) then Pfad_JavaExe = GetExeAusRegistry(Key_JavaExe)
  
  GetJavaExe = Pfad_JavaExe
End Function


Private sub GetJEditUmgebung()
  'legt folgende öffentlichen Variablen fest:
  'jEditjar   = Pfad\Dateiname der gefundenen Datei jEdit.jar oder "".
  'JEDIT_HOME = Verzeichnis von jEditjar
  on error goto 0
  dim Erfolg
  const JarName = "jEdit.jar"
  Erfolg = false
  DebugEcho "Pfad\Dateiname von " & JarName & " ermitteln."
  
  JEDIT_HOME = LastBackslashDelete(Environ("JEDIT_HOME"))
  if (JEDIT_HOME = "") then
    DebugEcho "GetJEditUmgebung(): Umgebungsvariable %JEDIT_HOME% existiert nicht. => Suche weiter in Konfiguration"
  else
    jEditjar = JEDIT_HOME & "\" & JarName
    if (fs.FileExists(jEditjar)) then
      Erfolg = true
      DebugEcho "GetJEditUmgebung(): jEdit.jar gefunden im Verzeichnis der Umgebungsvariable %JEDIT_HOME%='" & JEDIT_HOME & "'"
    else
      DebugEcho "GetJEditUmgebung(): jEdit.jar nicht gefunden im Verzeichnis der Umgebungsvariable %JEDIT_HOME%='" & JEDIT_HOME & "'"
    end if
  end if
  
  'Spezial für GeoTools:
  if (not Erfolg) then
    if (oKonfig is nothing) then
      DebugEcho "GetJEditUmgebung(): Konfigurationsobjekt existiert nicht!"
    else
      JEDIT_HOME = LastBackslashDelete(oKonfig.JEDIT_HOME)
      jEditjar = JEDIT_HOME & "\" & JarName
      if (fs.FileExists(jEditjar)) then
        Erfolg = true
        DebugEcho "GetJEditUmgebung(): jEdit.jar gefunden im konfigurierten Verzeichnis '" & JEDIT_HOME & "'"
      else
        DebugEcho "GetJEditUmgebung(): jEdit.jar nicht gefunden im konfigurierten Verzeichnis '" & JEDIT_HOME & "'"
      end if
    end if
  end if
    
  if (not Erfolg) then
    JEDIT_HOME = "G:\Tools\jEdit"
    jEditjar = JEDIT_HOME & "\" & JarName
    if (fs.FileExists(jEditjar)) then
      Erfolg = true
    else
      DebugEcho "GetJEditUmgebung(): jEdit.jar nicht gefunden im (hart kodierten) Verzeichnis '" & JEDIT_HOME & "'"
    end if
  end if
  
  if (not Erfolg) then
    JEDIT_HOME = Environ("PROGRAMFILES") & "\jEdit"
    jEditjar = JEDIT_HOME & "\" & JarName
    if (fs.FileExists(jEditjar)) then
      Erfolg = true
    else
      DebugEcho "GetJEditUmgebung(): jEdit.jar nicht gefunden im (hart kodierten) Verzeichnis '" & JEDIT_HOME & "'"
    end if
  end if
  
  if (not Erfolg) then
    JEDIT_HOME = Environ("PROGRAMFILES") & "\Tools\jEdit"
    jEditjar = JEDIT_HOME & "\" & JarName
    if (fs.FileExists(jEditjar)) then
      Erfolg = true
    else
      DebugEcho "GetJEditUmgebung(): jEdit.jar nicht gefunden im (hart kodierten) Verzeichnis '" & JEDIT_HOME & "'"
    end if
  end if
  
  JEDIT_SETTINGS = LastBackslashDelete(Environ("JEDIT_SETTINGS"))
  DebugEcho "GetJEditUmgebung(): Umgebungsvariable %JEDIT_SETTINGS%='" & JEDIT_SETTINGS & "'"
  
  if (not Erfolg) then
    JEDIT_HOME = ""
    jEditjar   = ""
  end if
End sub


Private sub InitEditoren()
  'Füllt die Liste unterstützter Editoren und setzt den Standard-Editor.
  dim i, letzterEditor
  
  'Liste unterstützter und verfügbarer Editoren aufbauen.
  i = -1
  'redim Editoren(1,i)
  if (UEditExe <> "") then
    i = i + 1
    redim Editoren(1,i)
    'redim preserve Editoren(1,i)
    Editoren(0,i) = "UltraEdit"   'sichtbarer Text in Listbox
    Editoren(1,i) = "UltraEdit"   'Wert der Listbox => muss in diesem Fall identisch sein mit dem Text.
  end if
  if (CEditExe <> "") then
    i = i + 1
    if (isEmpty(Editoren)) then
      redim Editoren(1,i)
    else
      redim preserve Editoren(1,i)
    end if
    Editoren(0,i) = "Crimson Editor"   'sichtbarer Text in Listbox
    Editoren(1,i) = "Crimson Editor"   'Wert der Listbox => muss in diesem Fall identisch sein mit dem Text.
  end if
  if ((JEditJar <> "") and (JavaExe <> "")) then
    i = i + 1
    if (isEmpty(Editoren)) then
      redim Editoren(1,i)
    else
      redim preserve Editoren(1,i)
    end if
    Editoren(0,i) = "jEdit"   'sichtbarer Text in Listbox
    Editoren(1,i) = "jEdit"   'Wert der Listbox => muss in diesem Fall identisch sein mit dem Text.
  end if
  'Keinen Editor gefunden.
  if (i = -1) then
    i = 0
    redim Editoren(1,i)
    Editoren(0,i) = "kein Editor verfügbar"
    Editoren(1,i) = cKeinEditor
  end if
  call TransposeArray2d(Editoren)
  
  'Standard-Editor setzen.
  strEditor = cKeinEditor
  'Zuletzt eingestellten Editor aus Registry lesen
  on error resume next
  letzterEditor = RegRead(RegValueEditor)
  on error goto 0
  Editor = letzterEditor
  if (strEditor = cKeinEditor) then Editor = Editoren(0,1)
End sub


Private sub GetInternational() 
  'Ermittelt das in den Ländereinstellungen der Windows-Systemsteuerung festgelegte
  'Listentrennzeichen und Dezimaltrennzeichen. (bei Fehler ist der Rückgabewert = "?")
  '==> in Excel geht's auch so: Application.International(xlListSeparator)
  'On Error Resume Next
  Const Locale_RegValue   = "HKEY_CURRENT_USER\Control Panel\International\Locale"
  Const sList_RegValue    = "HKEY_CURRENT_USER\Control Panel\International\sList"
  Const sDecimal_RegValue = "HKEY_CURRENT_USER\Control Panel\International\sDecimal"
  
  If (RegValueExists(Locale_RegValue)) Then
    LangID = WSHShell.RegRead(Locale_RegValue)
    On Error Resume Next
    LangID = Val("&H" & LangID)
    if (Err.Number <> 0) then
      LangID = 0
      Err.clear
    end if
    On Error GoTo 0
  Else
    LangID = 0
  End If
  
  If (RegValueExists(sList_RegValue)) Then
    ListenTrenner = RegRead(sList_RegValue)
  Else
    ListenTrenner = "?"
  End If
  
  If (RegValueExists(sDecimal_RegValue)) Then
    DezimalTrenner = RegRead(sDecimal_RegValue)
  Else
    DezimalTrenner = "?"
  End If
  
  On Error GoTo 0
End sub


Private sub GetVBAUmgebung()
  'Ermittelt VBA-Umgebung.
  VBAHostName = Application.Name
  VBAHostVersion = Application.Version
  
  'Probleme:
  '1. Zugriff auf das VBProject ist eventuell nicht erlaubt.
  '2. Zugriffskode kann nicht für mehrere Hosts geschrieben werden (?)
  'if (VBAHostName = VBAHostNameExcel) then
  '  VBAProjectName = ThisWorkbook.VBProject.Name
  '  VBAProjectFileName = ThisWorkbook.VBProject.FileName
  'else
  '  'anderer Host
  '  VBAProjectName = "?"
  '  VBAProjectFileName = "?"
  'end if
End sub


private sub GetSystemUmgebung()
  'Ermittelt Umgebung des Betriebssystems:
  
  OSinfo.Type = Environ("OS")
  
  if (OSinfo.Type <> "Windows_NT") then
    ErrEcho "GetSystemUmgebung(): Betriebsystem '" & OSinfo.Type & "' wird nicht unterstützt!!!"
  else
    OSinfo.SystemDrive = Environ("SystemDrive")
    OSinfo.SystemRoot  = Environ("SystemRoot")
    OSinfo.WindowsSys  = OSinfo.SystemRoot & "\System32"
  end if
End sub


Private sub GetNetzUmgebung()
  'Ermittelt die gebräuchlichsten Info's der Netzwerkumgebung
  'und belegt damit globale Variablen.
  'The loop is needed, because W9x machines need a short time
  'to respond to this function during logon
  Const valueDomain    = "HKLM\SYSTEM\CurrentControlSet\Services\MSNP32\NetworkProvider\AuthenticatingAgent"
  Const valueWorkgroup = "HKLM\SYSTEM\CurrentControlSet\Services\VxD\VNETSUP\Workgroup"
  Dim WshNetwork       As New IWshRuntimeLibrary.WshNetwork
  
  'Username, Computername
  Do Until (Username <> "")
    Username = WshNetwork.UserName
  Loop
  Computername = WshNetwork.Computername
  
  'UserDomain
  If (Environ("OS") = "Windows_NT") Then
    UserDomain = WshNetwork.UserDomain
  Else
    UserDomain = RegRead(valueDomain)
    If (UserDomain = "") Then
     UserDomain = RegRead(valueWorkgroup)
    End If
  End If
  Set WshNetwork = Nothing
End sub


'für jEdit:  :folding=indent::collapseFolds=1:
