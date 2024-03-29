Attribute VB_Name = "mdlToolsExcel"
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) f�r Geod�ten.
' Copyright � 2003 - 2024  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'====================================================================================
'Modul mdlToolsExcel
'====================================================================================
'Werkzeuge, die auf Excel zur�ckgreifen.


Option Explicit


Private SavedSystemSeparator_Decimal  As String
Private SavedSystemSeparator_List     As String
Private SavedSystemSeparator_Thousand As String
Private SavedExcelSeparator_Decimal   As String
Private SavedExcelSeparator_Thousand  As String
Private SavedXLUsesSystemSeparators   As Boolean
Private ChangedSystemSeparators       As Boolean
Private ChangedExcelSeparators        As Boolean
Private SavedSeparatorsInitialized    As Boolean



Function IsMacrosExecutable() As Boolean
  'Checks if VBA macros are executable at this point.
  On Error GoTo Fehler
  Application.StatusBar = Application.StatusBar
  IsMacrosExecutable = True
  Exit Function
Fehler:
  Err.Clear
  IsMacrosExecutable = False
End Function


Sub SetApplicationVisible(ByVal Visible As Boolean)
  ' 02.05.2020: Wenn Excel sichtbar geschaltet wird, obwohl es das bereits ist,
  ' dann wird ein extra Fenster ge�ffnet, und Excel l�sst sich nur mit Trick 17 schlie�en!
  On Error GoTo Fehler
  If (Visible) Then
    If (Not Application.Visible) Then Application.Visible = True
  Else
    If (Application.Visible) Then Application.Visible = False
  End If
  Exit Sub
Fehler:
  Err.Clear
End Sub


'***  Abteilung Statuszeile und Meldungen  ***************************************

Public Sub FehlerNachricht(ByVal FehlerQuelle As String)
  'Erzeugt eine Messagebox und einen Protokolleintrag
  'mit folgenden Angaben zum aktuellen Fehler:
  ' - "Number", "Description" und "Source" des err-Objektes
  ' - err.source wird erg�nzt um "\FehlerQuelle"
  ' - "ErrMessage" (globale Variable, darf leer sein)
  'Danach wird der Fehler gel�scht.
  
  Dim Titel     As String
  Dim Message   As String
  Dim ErrNumber As Long
  
  ErrNumber = Err.Number
  
  If (Err.Number <> 0) Then
    Titel = "FEHLER in: '" & Err.Source & "\" & FehlerQuelle & "'"
    'Message = "Fehlernummer      : 0x" & Hex(Err.Number) & vbNewLine &
    Message = "Fehlernummer        : " & Err.Number & vbNewLine & _
              "Fehlerbeschreibung  : " & Err.Description
    Err.Clear
    If (ErrMessage <> "") Then Message = Message & vbNewLine & vbNewLine & "Bemerkung           : " & ErrMessage
    DebugEcho Titel
  Else
    Titel = "FEHLER"
    If (ErrMessage <> "") Then Message = ErrMessage
  End If
  
  ' Nicht eher, da sonst das Err-Objekt geleert wird:
  On Error GoTo Fehler
  
  'Protokolleintrag
  If ((ErrMessage <> "") Or (ErrNumber <> 0)) Then
    'Fehlermeldung f�r Protokoll
    SetApplicationVisible(true)
    Application.UserControl    = true
    Application.ScreenUpdating = true
    ErrEcho replace(Message, vbNewLine & vbNewLine, vbNewLine)
  end if
  
  'Dialog
  If (Message <> "") Then
    SetApplicationVisible(true)
    Application.UserControl    = true
    Application.ScreenUpdating = true
    MsgBox Message, vbExclamation, Titel
  end if
  
  ErrMessage = ""
  call ClearStatusBarDelayed(StatusBarClearDelay)
  
  Exit Sub
Fehler:
  Err.Clear
  ErrMessage = ""
End Sub


Sub ProgressbarAllgemein(ByVal MaxValue As Double, ByVal CurrentValue As Double, ByVal Text As String)
  ' Zeigt in der Statuszeile den angegebenen Fortschritt incl. Text an.
  ' MaxValue    :  Maximum value to process
  ' CurrentValue:  Currently processed value
  ' Text        :  Text to display right of progress bar
  On Error GoTo Fehler
  
  Const SingleBAR = "|"   ' Character for progress bar
  Const MaxBARS   = 50    ' BAR count at 100%
  
  Dim Ratio       As Double
  Dim Percent     As String
  Dim Bar         As String
  Dim BARcount    As Integer
  Dim SpaceCount  As Integer
  
  MaxValue = Abs(MaxValue)
  CurrentValue = Abs(CurrentValue)
  
  If ((CurrentValue <= MaxValue) And (MaxValue > 0)) Then
      
      Ratio      = CurrentValue / MaxValue
      Percent    = CStr(CInt(Ratio * 100)) & "%  "
      BARcount   = CInt(Ratio * MaxBARS)
      SpaceCount = MaxBARS - BARcount - 1
      If (SpaceCount < 0) Then SpaceCount = 0
      Bar        = Percent & String$(BARcount, "|") & String$(SpaceCount, " ") & "|     " & Text
  Else
      Bar = "100%  " & String$(MaxBARS, "|") & "     " & Text
  End If
  
  Application.DisplayStatusBar = True
  Application.StatusBar = Bar
  DoEvents
  
  Exit Sub
  
Fehler:
  FehlerNachricht "mdlToolsExcel.ProgressbarAllgemein()"
End Sub

Sub ProgressbarDateiLesen(ByVal Dateinummer As Integer)
  'zeigt in der Statuszeile den Fortschritt des Einlesens einer Datei an.
  'Parameter:  Dateinummer = logische Nr. der ge�ffneten Datei.
  
  'On Error GoTo Fehler
  Const einBAR = "|"     'Zeichen f�r die Balkenbildung.
  Const maxBARS = 65     'Anzahl "BAR"s bei 100%
  Const F1 = 1.25        'L�ngenverh�ltnis BAR/Leerzeichen
  
  Dim lngLOC  As Long     'aktuelle Leseposition/128, aufgerundet.
  Dim lngLOF  As Long     'L�nge der Datei.
  Dim Faktor  As Double
  Dim Prozent As String
  Dim Bar     As String
  Dim AnzBARs As Integer
  
  lngLOC = Loc(Dateinummer)
  lngLOF = LOF(Dateinummer)
  
  If lngLOC > 0 Then
    Faktor = (lngLOC) * 128 / lngLOF
    if (Faktor > 1) then Faktor = 1  'Falls Dateigr��e kleiner als 128!
    Prozent = CStr(CInt(Faktor * 100)) & "%  "
    AnzBARs = CInt(Faktor * maxBARS)
    Bar = "Lese Datei..  " & Prozent & String$(CInt(Faktor * maxBARS), "|")
    'Bar = "Lese Datei..  " & Prozent & "[" & String$(AnzBARs, "|") & String$(F1 * (maxBARS - AnzBARs), " ") & "]"
  Else
    Bar = "Lese Datei..  " & String$(maxBARS, "|") & " 100%"
  End If
  Application.DisplayStatusBar = True
  Application.StatusBar = Bar
  DoEvents
  
  Exit Sub
  
Fehler:
  FehlerNachricht "mdlToolsExcel.ProgressbarDateiLesen()"
End Sub

Sub ClearStatusBarDelayed(Seconds as integer)
  'Clears the statusbar after a given amount of seconds.
  On Error GoTo Fehler
  Application.OnTime Now + TimeSerial(0,0,Seconds), "ClearStatusBar"
  Exit Sub
Fehler:
  Err.Clear
End Sub

Sub ClearStatusBar()
  'This method is needed for ClearStatusBarDelayed()...
  On Error GoTo Fehler
  Application.StatusBar = ""
  
  ' This statement cancels an active copied status of a range.
  ' Since UpdateGeoToolsRibbon() calls this, It won't be possible
  ' to paste into another sheet!
  'Application.DisplayStatusBar = True
  
  DoEvents
  Exit Sub
Fehler:
  Err.Clear
End Sub

Sub WriteStatusBar(Message as String)
  'This method is needed to catch different status bar handling in different VBA hosts.
  On Error GoTo Fehler
  Application.DisplayStatusBar = True
  Application.StatusBar = Message
  DoEvents
  Exit Sub
Fehler:
  Err.Clear
End Sub



'***  Abteilung Dateien  **************************************************************************

Public Function FindeXLVorlage(ByVal FileName As String) As String
  '
  ' ACHTUNG: 01.05.2020: nicht mehr aktuell, siehe frmStartExpim.frm\GetXLVorlagen()
  '
  'Gibt den vollst�ndigen Namen incl. Pfad der zuerst gefundenen Vorlage zur�ck.
  'Existiert keine entsprechende Datei, wird "" zur�ckgegeben.
  '   "FileName"     = Dateiname ohne absolute oder relative Pfadangabe (keine Wildcards)
  '
  'Sucht die angegebene Datei in den selben Verzeichnissen, die im Normalfall von Excel
  'durchsucht werden, wenn zwecks Anlegen einer neuen Arbeitsmappe alle Vorlagen zur
  'Auswahl angeboten werden. Alle Verzeichnisnamen werden durch Auslesen von Eigenschaften
  'des Excel-VBA-Objektes "Application" ermittelt (Es w�re auch �ber die Registry m�glich).
  'Gesucht wird in folgender, nicht standardkonformer Reihenfolge:
  '    Application.NetworkTemplatesPath  (HKCU\Software\Microsoft\Office\8.0\Common\FileNew\SharedTemplates\@ = R:\OFFICE97\Winword\Vorlagen)
  '    Application.AltStartupPath        (HKCU\Software\Microsoft\Office\8.0\Excel\Microsoft Excel\AltStartup = R:\Office97\Excel\Xlstart)
  '    Application.TemplatesPath         (HKCU\Software\Microsoft\Office\8.0\Common\FileNew\LocalTemplates\@  = R:\OFFICE97\Winword\Vorlagen ?)
  '    Application.StartupPath           ("HKLM\Software\Microsoft\Office\8.0\Excel\InstallRoot\Path" & "\XLStart")
  '==> Es werden auch Unterverzeichnisse dieser 4 Ordner durchsucht, um das normale
  '    Verhalten von Excel f�r die Vorlagenverzeichnisse zu unterst�tzen, deren
  '    Unterverzeichnisse als "Reiter" im Dialog angezeigt werden. Excel durchsucht
  '    zwar keine Unterverzeichnisse der beiden Startordner, allerdings gibt es diese
  '    normalerweise nicht.
  '==> Wurde Excel via Automation gestartet, so werden im Dialog "Datei|Neu" die
  '    Vorlagen der beiden Startordner nicht angeboten. Deshalb sollten Vorlagen
  '    besser in einem (Unterverz. eines) Vorlagenordner abgelegt sein.
  '    ==> z.B. "R:\OFFICE97\Winword\Vorlagen\Tabellen\"
  
  On Error Resume Next
  Dim VerzeichnisListe  As String
  
  VerzeichnisListe = Application.NetworkTemplatesPath & ";" & _
                     Application.AltStartupPath & ";" & _
                     Application.TemplatesPath & ";" & _
                     Application.StartupPath
  'MsgBox VerzeichnisListe
  FindeXLVorlage = ThisWorkbook.SysTools.FindFile(FileName, VerzeichnisListe, True)
  
End Function



'***  Abteilung Excel-Global  *********************************************************************

Public Function SetRequiredSeparators() As Boolean
    ' Setzt die f�r GeoTools erforderlichen Separatoren der Windows-Regionaleinstellungen
    ' und speichert vorher die aktuellen Einstellungen.
    ' R�ckgabe: Erfolg
    On Error Goto 0
    
    Dim success As Boolean
    
    ' Aktuelle Einstellungen merken.
    success = ThisWorkbook.SysTools.GetSeparators(SavedSystemSeparator_Decimal, SavedSystemSeparator_List, SavedSystemSeparator_Thousand)
    SavedExcelSeparator_Decimal  = Application.DecimalSeparator
    SavedExcelSeparator_Thousand = Application.ThousandsSeparator
    SavedXLUsesSystemSeparators  = Application.UseSystemSeparators
    SavedSeparatorsInitialized   = True
    
    ' Feststellen, ob aktuelle Einstellungen von den erforderlichen abweichen.
    ChangedExcelSeparators        = ((Not ((SavedExcelSeparator_Decimal  = RequiredSeparator_Decimal)    And _
                                           (SavedExcelSeparator_Thousand = RequiredSeparator_Thousand))) Or _
                                           (SavedXLUsesSystemSeparators XOR RequiredXLUsesSystemSeparators) )
    ChangedSystemSeparators       = (Not  ((SavedSystemSeparator_Decimal  = RequiredSeparator_Decimal)   And _
                                           (SavedSystemSeparator_List     = RequiredSeparator_List)      And _
                                           (SavedSystemSeparator_Thousand = RequiredSeparator_Thousand) ))
    '
    ' Erforderliche Einstellungen setzen, falls n�tig.
    If (ChangedExcelSeparators) Then
        Application.UseSystemSeparators = RequiredXLUsesSystemSeparators
        Application.DecimalSeparator    = RequiredSeparator_Decimal
        Application.ThousandsSeparator  = RequiredSeparator_Thousand
    End If
    If (ChangedSystemSeparators) Then
        success = ThisWorkbook.SysTools.SetSeparators(RequiredSeparator_Decimal, RequiredSeparator_List, RequiredSeparator_Thousand)
    End If
    
    DebugEcho "SetRequiredSeparators(): Ergebnis effektiv:  DezimalTrenner= '" & Application.International(XLDecimalSeparator) & "'  TausenderTrenner= '" & Application.International(xlThousandsSeparator) & "'  ListenTrenner = '" & Application.International(xlListSeparator) & "'."
    
    SetRequiredSeparators = success
End Function

Public Function RestoreLastSeparators() As Boolean
    ' Stellt die gespeicherten Separatoren der Windows-Regionaleinstellungen wieder her, falls vorher ge�ndert.
    ' R�ckgabe: Erfolg
    On Error Goto 0
    
    Dim success As Boolean
    
    If (SavedSeparatorsInitialized) Then
      If (ChangedExcelSeparators) Then
          Application.UseSystemSeparators = SavedXLUsesSystemSeparators
          Application.DecimalSeparator    = SavedExcelSeparator_Decimal
          Application.ThousandsSeparator  = SavedExcelSeparator_Thousand
      End If
      If (ChangedSystemSeparators) Then
          success = ThisWorkbook.SysTools.SetSeparators(SavedSystemSeparator_Decimal, SavedSystemSeparator_List, SavedSystemSeparator_Thousand)
      End If
    End If
    
    DebugEcho "RestoreLastSeparators(): Ergebnis effektiv:  DezimalTrenner= '" & Application.International(XLDecimalSeparator) & "'  TausenderTrenner= '" & Application.International(xlThousandsSeparator) & "'  ListenTrenner = '" & Application.International(xlListSeparator) & "'."
    
    RestoreLastSeparators = success
End Function


Public Function SetArbeitsverzeichnis(Optional ByVal Verzeichnis As String = "")
  'Funktionswert: Name des eingestellten bzw. beibehaltenen Arbeitsverzeichnisses.
  'Argument:      "Verzeichnis" ... Optional. Zu setzendes Arbeitsverzeichnis.
  'Arbeitsweise:
  'Das angegebene Verzeichnis wird als Arbeitsverzeichnis gesetzt. Falls kein
  'Verzeichnis angegeben wurde, oder das Einstellen dieses Verzeichnisses fehlschlug,
  'wird nur dann ein Verzeichniswechsel durchgef�hrt, wenn das aktuelle Verzeichnis
  'ein Systemverzeichnis ist ("windows" oder "winnt" im Namen). In diesem Fall wird
  'der in den Excel-Optionen angegebene "Standardarbeitsordner" eingestellt. Schl�gt
  'dieser Versuch fehl, so wird das in "ThisWorkbook.Konfig.StdArbeitsverz" festgelegte
  'Verzeichnis verwendet.
  
  On Error GoTo Fehler
  
  Dim strArbeitsverz  As String
  Dim strCurDir       As String
  
  Verzeichnis = LastBackslashDelete(Verzeichnis)
  
  If (ThisWorkbook.SysTools.isVerzeichnis(Verzeichnis)) Then
    'angegebenes Verzeichnis einstellen
    On Error Resume Next
    ChDrive Verzeichnis
    ChDir Verzeichnis
    On Error GoTo 0
  End If
  
  strCurDir = LCase(CurDir())
  If (strCurDir <> LCase(Verzeichnis)) Then
    'eventuell angegebenes Verzeichnis konnte nicht eingestellt werden.
    If ((InStr(1, strCurDir, "windows", vbTextCompare) > 0) Or (InStr(1, strCurDir, "winnt", vbTextCompare) > 0)) Then
      'kein sinnvolles Arbeitsverzeichnis eingestellt (n�mlich Systemverzeichnis)
      strArbeitsverz = LastBackslashDelete(Application.DefaultFilePath)
      If ((Not ThisWorkbook.SysTools.isDatei(strArbeitsverz & "\")) Or (strArbeitsverz = "")) Then
        'Application.DefaultFilePath nicht oder fehlerhaft gesetzt.
        strArbeitsverz = LastBackslashDelete(ThisWorkbook.Konfig.StdArbeitsverz)    'Voreinstellung als Konstante in mdlAllgemein
      End If
      On Error Resume Next
      ChDrive strArbeitsverz
      ChDir strArbeitsverz
      On Error GoTo 0
    End If
  End If
  strCurDir = CurDir()
  Application.StatusBar = "Arbeitsverzeichnis gesetzt auf: " & strCurDir
  DebugEcho "mdlToolsExcel.SetArbeitsverzeichnis(): Arbeitsverzeichnis gesetzt auf: " & strCurDir
  SetArbeitsverzeichnis = strCurDir
  
  Exit Function
  
Fehler:
  Err.Clear
  ErrMessage = ""
End Function



'***  Abteilung Excel-Tabellen  *******************************************************************

Public Function isTabellenSchutz(Optional oTab As Worksheet = Nothing) As Boolean
  'Stellt fest, ob das angegebene Tabellenblatt gesch�tzt ist (vor wem auch immer).
  '
  'Ist "oTab" nicht oder mit Nothing angegeben, dann ist das aktive Arbeitsblatt gemeint.
  On Error GoTo Fehler
  
  If (oTab Is Nothing) Then
    Set oTab = ActiveSheet
  End If
  
  If (Not (oTab Is Nothing)) Then
    With oTab
      If (.ProtectDrawingObjects = True Or _
        .ProtectContents = True Or _
        .ProtectScenarios = True) Then
        isTabellenSchutz = True
      Else
        isTabellenSchutz = False
      End If
    End With
  End If
  Exit Function
  
Fehler:
  FehlerNachricht "mdlToolsExcel.isTabellenSchutz()"
End Function


Public Sub ZellNamensListe()
  'Erstellt im aktiven Tabelenblatt eine Liste mit allen
  'benannten Zellbereichen der aktiven Arbeitsmappe.
  '==> zu Kontrollzwecken
  
  On Error Resume Next
  Dim benannteZellen As Names
  Dim oName          As Range
  Dim Adresse        As String
  Dim Zellname       As String
  Dim Tabelle        As Excel.Worksheet
  Dim i              As Long
  
  Set benannteZellen = ActiveWorkbook.Names
  Set Tabelle = ActiveSheet
  
  Tabelle.Cells(20, 1).value = "Zellname"
  Tabelle.Cells(20, 2).value = "Adresse"
  Tabelle.Cells(20, 3).value = "ActiveWorkbook.Name"
  Tabelle.Cells(20, 4).value = "ActiveSheet.Name"
  Tabelle.Cells(20, 5).value = "lokal ?"
  
  'MsgBox "Anzahl benannter Bereiche=" & benannteZellen.Count
  For i = 1 To benannteZellen.Count
    'MsgBox "Name = " & benannteZellen(i).Name
    Zellname = benannteZellen(i).Name
    Tabelle.Cells(i + 20, 1).value = Zellname
    'Set oName = benannteZellen(i).RefersToRange ==> funktioniert nur mit Komma als aktivem Listentrennzeichen!
    Set oName = Application.Range(benannteZellen(i).RefersTo)
    'Adresse = benannteZellen(i).RefersToRange.Address(External:=True)
    If (oName Is Nothing) Then
      Tabelle.Cells(i + 20, 2).value = "kein g�ltiger Bereich!"
    Else
      Adresse = oName.Address(External:=True)
      Tabelle.Cells(i + 20, 2).value = Adresse
      Tabelle.Cells(i + 20, 3).value = ActiveWorkbook.Name
      Tabelle.Cells(i + 20, 4).value = ActiveSheet.Name
      If (isLokalerZellName(benannteZellen(i))) Then
        Tabelle.Cells(i + 20, 5).value = "lokal"
      End If
    End If
  Next
  
  Set benannteZellen = Nothing
  Set Tabelle = Nothing
  Set oName = Nothing
  On Error GoTo 0
End Sub


Public Function isLokalerZellName(oName As Name, Optional oTab As Worksheet = Nothing) As Boolean
  'Stellt fest, ob der benannte Bereich "oName" im angegebenen Tabellenblatt liegt.
  '
  'Ist "oTab" nicht oder mit Nothing angegeben, dann ist das aktive Arbeitsblatt gemeint.
  
  On Error Resume Next
  Dim Adresse        As String
  Dim TabName        As String
  
  If (oTab Is Nothing) Then
    Set oTab = ActiveSheet
  End If
    
  If (oName Is Nothing) Then
    isLokalerZellName = False
  Else
    Adresse = oTab.Application.Range(oName.RefersTo).Address(External:=True)
    Adresse = substitute("'\[", "\[", Adresse, False, False)
    Adresse = substitute("'!", "!", Adresse, False, False)
    TabName = "[" & oTab.Parent.Name & "]" & oTab.Name & "!"
    If (InStr(Adresse, TabName)) Then
      'in der mit (External:=True) ermittelten Adresse ist immer der "TabName" enthalten.
      isLokalerZellName = True
    Else
      isLokalerZellName = False
    End If
  End If
  On Error GoTo 0
  
End Function


Public Function isSelectionRechteck() As Boolean
  'Liefert "true", wenn die aktive Auswahl aus einem einzigen Rechteck besteht
  'und auch keine ganzen Zeilen oder Spalten markiert sind.
  'Liefert "false", wenn die aktive Auswahl ganzen Zeilen oder Spalten enth�lt oder
  'keine "Range" ist oder gar nicht existiert (z.B. keine Tabelle aktiv).
  
  Dim AnzZeilen       As Long
  Dim AnzSpalten      As Long
  Dim AnzTeilbereiche As Long

  On Error GoTo Fehler
  isSelectionRechteck = False
  If (Not (ActiveCell Is Nothing)) Then
    AnzTeilbereiche = Selection.Areas.Count
    If (AnzTeilbereiche = 1) Then
        AnzZeilen = Selection.Rows.Count
        AnzSpalten = Selection.Columns.Count
        If ((AnzZeilen < ActiveSheet.Rows.Count) And (AnzSpalten < ActiveSheet.Columns.Count)) Then
          isSelectionRechteck = True
        End If
    End If
  End If
  Exit Function

Fehler:
  FehlerNachricht "mdlToolsExcel.isSelectionRechteck()"
End Function


Function GetLokalerZellname(ByVal Name As String, Optional oTab As Worksheet = Nothing) As Range
  'Gibt den in der angegebenen Tabelle liegenden benannten Zellbereich "Name" zur�ck.
  'Bezieht sich der benannte Bereich nicht auf die angegebene Tabelle, oder existiert
  'er gar nicht, so wird "nothing" zur�ckgegeben.
  '
  'Ist "oTab" nicht oder mit Nothing angegeben, dann ist das aktive Arbeitsblatt gemeint.
  
  On Error Resume Next
  
  Dim benannteZellen    As Names
  Dim oRange            As Range
  Dim Zellname          As String
  Dim i                 As Long
  Dim blnNameGefunden   As Boolean
  
  If (oTab Is Nothing) Then
    Set oTab = ActiveSheet
  End If
  
  Set oRange = Nothing
  blnNameGefunden = False
  If (oTab.Type = xlWorksheet) Then
    Set benannteZellen = oTab.Parent.Names
    i = 1
    Do While ((i <= benannteZellen.Count) And (Not blnNameGefunden))
      Zellname = benannteZellen(i).Name
      If (isLokalerZellName(benannteZellen(i), oTab)) Then
        'Bereichsname existiert im angegebenen Tabellenblatt.
        'ZellName kann den Tabellennamen enthalten. Dies steuert Excel automatisch...
        If ((Zellname = Name) Or (entspricht("!" & Name & "$", Zellname))) Then
          blnNameGefunden = True
          Set oRange = oTab.Application.Range(benannteZellen(i).RefersTo)
        End If
      End If
      i = i + 1
    Loop
    Set benannteZellen = Nothing
  End If
  Set GetLokalerZellname = oRange
  Set oRange = Nothing
  Exit Function

Fehler:
  FehlerNachricht "mdlToolsExcel.GetLokalerZellname()"
End Function


Public Function SchreibenFelderInTabelle(oDictionary As Scripting.Dictionary, Optional oTab As Worksheet = Nothing) As Boolean
  'Schreibt von allen verf�gbaren Items des Dictionary diejenigen in die angegebene Tabelle,
  'f�r die entsprechend benannte Zellen existieren, d.h. lokaler Zellname = Key.
  'Parameter:  oDictionary: Typ =Scripting.Dictionary
  '                         Key =Feldname (Zielzelle)
  '                         Item=Feldwert (zu schreibender Wert)
  'R�ckgabe: True, wenn zumindest ein Feld gefunden und beschrieben wurde.
  '
  'Ist "oTab" nicht oder mit Nothing angegeben, dann ist das aktive Arbeitsblatt gemeint.
  'Datumswerte werden formatiert, falls Zielzelle als Text formatiert ist.
  
  On Error GoTo Fehler
  Dim Feld             As Variant
  Dim Value            As Variant
  Dim FeldGeschrieben  As Boolean
  Dim oRangeName       As Range
  
  If (oTab Is Nothing) Then
    Set oTab = ActiveSheet
  End If
  
  FeldGeschrieben = False
  For Each Feld In oDictionary
    Set oRangeName = GetLokalerZellname(Feld, oTab)
    If (Not oRangeName Is Nothing) Then
      Value = oDictionary(Feld)
      
      ' Spezialbehandlung f�r Datumswerte, falls Zielzelle als Text formatiert ist.
      If (VarType(oDictionary(Feld)) = vbDate) Then
        If (Not IsNull(oRangeName.NumberFormat)) Then
          If (oRangeName.NumberFormat = "@") Then
            ' Zielzelle ist als Text formatiert (Einfache Zuweisung w�rde f�hren zu z.B. "1/21/2024").
            Value = Format(oDictionary(Feld), "dd.mm.yyyy")
          End If
        End If
      End If
      
      oRangeName.value = Value
      FeldGeschrieben = True
    End If
  Next
  SchreibenFelderInTabelle = FeldGeschrieben
  Exit Function
  
Fehler:
  FehlerNachricht "mdlToolsExcel.SchreibenFelderInTabelle()"
  SchreibenFelderInTabelle = False
End Function


Public Function LesenFelderAusTabelle(oDictionary As Scripting.Dictionary, Optional oTab As Worksheet = Nothing) As Boolean
  'Belegt alle Items des Dictionary mit Zellinhalten der angegebenen Tabelle,
  'wenn lokaler Zellname = Key des Dictionary.
  'Parameter:  oDictionary: Typ =Scripting.Dictionary
  '                         Key =Feldname (Quellzelle)
  '                         Item=Feldwert (zu belegender Wert)
  'R�ckgabe: True, wenn zumindest ein Feld gefunden wurde.
  '
  'Ist "oTab" nicht oder mit Nothing angegeben, dann ist das aktive Arbeitsblatt gemeint.
  
  On Error GoTo Fehler
  Dim Feld             As Variant
  Dim FeldGelesen      As Boolean
  Dim oRangeName       As Range
  
  If (oTab Is Nothing) Then
    Set oTab = ActiveSheet
  End If
  
  FeldGelesen = False
  For Each Feld In oDictionary
    Set oRangeName = GetLokalerZellname(Feld, oTab)
    If (Not oRangeName Is Nothing) Then
      FeldGelesen = True
      oDictionary(Feld) = oRangeName.value
    End If
  Next
  LesenFelderAusTabelle = FeldGelesen
  Exit Function
  
Fehler:
  FehlerNachricht "mdlToolsExcel.LesenFelderAusTabelle()"
  LesenFelderAusTabelle = False
End Function


Public Function GetFelderAusTabelle(ByVal Prefix As String, Optional oTab As Worksheet = Nothing) As Scripting.Dictionary
  'Findet alle benannten Zellen der Arbeitsmappe, in der sich die angegebene Tabelle befindet
  'und die sich auf die angegebene Tabelle beziehen und deren Namen mit "Prefix" beginnen.
  'Ist "oTab" nicht oder mit Nothing angegeben, dann ist das aktive Arbeitsblatt gemeint.
  '  Prefix   ... Es werden nur Namen ber�cksichtigt, die damit beginnen.
  '  R�ckgabe ... Dictionary (Key=Feldname, Item=Inhalt)
  
  'On Error Resume Next
  On Error GoTo Fehler
  
  Dim benannteZellen        As Names
  Dim Zellname              As String
  Dim ZellnamePur           As String
  Dim RegExPrefix           As String
  Dim RegExZellname         As String
  Dim RegExZellnamePrefix   As String
  Dim i                     As Long
  Dim oRange                As Range
  Dim oFelder               As Scripting.Dictionary
  
  If (oTab Is Nothing) Then
    Set oTab = ActiveSheet
  End If
  
  Set oRange    = Nothing
  Set oFelder   = New Scripting.Dictionary
  
  RegExPrefix   = FileSpec2RegExp(Prefix)
  RegExZellname = "^(.*!)?"                        'reg. Ausdruck f�r einen Zellnamen
  RegExZellnamePrefix = "^(.*!)?" & RegExPrefix    'reg. Ausdruck f�r einen Zellnamen mit Pr�fix
  
  If (oTab.Type = xlWorksheet) Then
    Set benannteZellen = oTab.Parent.Names
    i = 1
    Do While (i <= benannteZellen.Count)
      Zellname = benannteZellen(i).Name
      If (isLokalerZellName(benannteZellen(i), oTab)) Then
        'Bereichsname existiert im angegebenen Tabellenblatt.
        'ZellName kann den Tabellennamen enthalten. Dies steuert Excel automatisch...
        If (entspricht(RegExZellnamePrefix, Zellname)) Then
          'Zellname gefunden, der mit dem gesuchten Pr�fix beginnt.
          ZellnamePur = substitute(RegExZellname, "", Zellname, False, False)  'Ergebnis z.B. = "Prj.KooSystem"
          Set oRange  = oTab.Application.Range(benannteZellen(i).RefersTo)
          oFelder.add ZellnamePur, oRange.Value
        End If
      End If
      i = i + 1
    Loop
    Set benannteZellen = Nothing
  End If
  
  if (oFelder.count = 0) then Set oFelder = Nothing
  
  Set GetFelderAusTabelle = oFelder
  Set oRange = Nothing
  Exit Function
  
Fehler:
  Set oRange = Nothing
  Set GetFelderAusTabelle = Nothing
  FehlerNachricht "mdlToolsExcel.GetFelderAusTabelle()"
End Function


Function FormelOhneVerweis(FormelMitOhneVerweis As String) As String
  'gibt die angegebene Formel ohne einen Verweis auf eine Tabelle zur�ck.
  Dim BeginnFktName  As Integer
  BeginnFktName = InStr(1, FormelMitOhneVerweis, "!", vbTextCompare) + 1
  If (BeginnFktName > 1) Then
    FormelOhneVerweis = "=" & Mid$(FormelMitOhneVerweis, BeginnFktName)
  Else
    FormelOhneVerweis = FormelMitOhneVerweis
  End If
End Function



'f�r jEdit:  :folding=indent::collapseFolds=1:
