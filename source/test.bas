Attribute VB_Name = "test"
'Option Explicit



Sub showColor()
    
    'ActiveCell.Interior.Color = RGB(234, 234, 234)
    'ActiveCell.Interior.ColorIndex = 39

    ActiveCell.value = ActiveCell.Interior.ColorIndex
End Sub







Sub FileNew()
  'Startet direkt den Dialog "Vorlage wählen" im Listenmodus.
  On Error Resume Next
  SendKeys "%2"
  Application.Dialogs(xlDialogNew).Show
End Sub


Sub Editor()
  'oSysTools.StarteDatei ActiveCell.value
  oSysTools.StartEditor ActiveCell.value
End Sub


Sub Test_Systools()
  Dim oSysTools_tmp    As CToolsSystem
  Set oSysTools_tmp = New CToolsSystem
  
  'MsgBox oSysTools_tmp.ArbeitsVerz
  MsgBox oSysTools_tmp.isVerzeichnis(".\GUI\")
  
  Set oSysTools_tmp = Nothing
End Sub


Sub Test_Fehlerliste()
  Dim oSysTools_tmp    As CToolsSystem
  Set oSysTools_tmp = New CToolsSystem
  
  Meldung1 = "Test-Meldung 1"
  Meldung2 = "Test-Meldung 2222"
  Extra1 = "Zeile 2" & vbNewLine & "Zeile 3"
  Extra2 = "Ex 2"
  
  Call oSysTools_tmp.FileErrorsAdd("F", "X:\Quellen\Excel\GeoTools\TkBezug.awk", 2511, 5, 11, Meldung1)
  Call oSysTools_tmp.FileErrorsAdd("F", "X:\Quellen\Excel\GeoTools\TkBezug.awk", 2511, 33, 44, Meldung2)
  
  Call oSysTools_tmp.FileErrorsShowInJEdit(True)
  
    
  
  Set oSysTools_tmp = Nothing
End Sub


Sub test_y()

  'MsgBox Application.Name & vbNewLine & Application.Version & vbNewLine & ThisWorkbook.VBProject.Name
  
  'MsgBox "'" & Format(1234567, "########0") & "'"
  'MsgBox GetStreckenbezeichnung("6258", False)
  'MsgBox GetKm("-3.0 + 5.4589")
  'MsgBox isDatei("C:\temp")
  'tt = "3" & vbNewLine & "rrrr"
  'tt = Null
  'Extrazeilen = Split(tt, vbNewLine)
  'MsgBox IsArray(Extrazeilen)
      
  'Formel_DP = "=@@S.Tra.SO@@+@@TK.HSOK@@4"
  'MsgBox "XL-Formel: '" & Formel_DP2XL(Formel_DP) & "'"
  'MsgBox UCase(Formel_XL)
  'MsgBox xlMoveAndSize & "  " & xlMove & "  " & xlFreeFloating
  'ActiveCell.Comment.Shape.IncrementTop 30
  
  'Selection.ShapeRange.IncrementLeft 2.4
  'Selection.ShapeRange.IncrementTop 16.8
  
  
  'Selection.AddComment "test"
  
  'For Each befehlsLeiste In CommandBars
  '  Debug.Print befehlsLeiste.Name, befehlsLeiste.NameLocal, befehlsLeiste.Visible
  'Next

  'MsgBox ActiveCell.Formula
  
  'MsgBox "'" & Format(" ", "@@@@@@@@@@@@@@@@@@@@") & "'"
  'MsgBox ThisWorkbook.VBProject.VBComponents(4).Name
  
  'MsgBox LeftStr("aaaaaa", ";", False)
  
  'Application.FileSearch.NewSearch
  'Application.FileSearch.FileName = "*.123"
  'MsgBox Application.FileSearch.FileName
  
  'maske = "*.xlt"
  'verzliste = "X:\QUELLEN\Excel\vorlagen\"
  'MsgBox FindeDateien_xMasken_2(maske, verzliste, False)
  'MsgBox Verz(ThisWorkbook.Path) & "\" & VorName(ThisWorkbook.Name) & ".pdf"
  
  'MsgBox ThisWorkbook.Path & vbNewLine & Verz(ThisWorkbook.Path)
  
  'MsgBox ThisWorkbook.Name & vbNewLine & VorName(ThisWorkbook.Name)
  
  'cfg = Verz(ThisWorkbook.Path) & "\" & VorName(ThisWorkbook.Name) & "_cfg.xls"
  'Application.Workbooks.Open FileName:=cfg, ReadOnly:=True
  
  'Dim oKonf As CdatKonfig
  'Set oKonf = New CdatKonfig
  'oKonf.ZeigeSpaltenKonfig

  'Sheets("Einstellungen").Select
  'Sheets("SpaltenGlobal").Activate
   
  'Format = ActiveCell.NumberFormat
  'FormatLocal = ActiveCell.NumberFormatLocal
   
  'MsgBox "'" & Format & "'" & vbNewLine & "'" & FormatLocal & "'"
  
  'MsgBox ActiveWorkbook.BuiltinDocumentProperties("title").value
  
  'echo "xxxxxx"
  'Call AnzeigeMeldungen
  'echo "......"
  'Call AnzeigeMeldungen

  'On Error Resume Next
  'Name = "x:\QUELLEN\Excel\Trassenkoo"
  ''Name = "X:\QUELLEN\Excel\Trassenkoo\Nk_um_gl.txt"
  'ne = NameExt(Name, "mitext")
  'DirDatei = Dir(Name)
  'DirVerz = Dir(Name, vbDirectory)
  'MsgBox isDatei(Name) & vbNewLine & isVerzeichnis(Name)
  'VerzEingabeDatei = Verz(Name)
  'PfadName = CurDir & "\" & NameExt(Name, "mitext")
  ''MsgBox VerzEingabeDatei & vbNewLine & Dirverz & vbNewLine & PfadName
  'On Error GoTo 0
  
  
  'Call AnzeigeMeldungen
  
  'MsgBox ActiveSheet.PageSetup.CenterFooter
  
  'MsgBox oPrjDatGlobal.Ort_Fusszeile_Excel_1
  'oAktiveTabelle.SchreibeFusszeile_1
  'Dim oAktiveTabelle  As CtabAktiveTabelle
  'Set oAktiveTabelle = New CtabAktiveTabelle
  
  'MsgBox oAktiveTabelle.FormatDatenNKStellenAnzahl & " ;Setzen=" & oAktiveTabelle.FormatDatenNKStellenSetzen
  
  'ActiveSheet.Select True
  'MsgBox substitute("[0-9]+$", "", oAktiveTabelle.TabKlasse, False, False)
  
  'MsgBox ActiveSheet.Name
  'MsgBox Application.VBE.VBProjects("Mappe1").VBComponents(4).Name
  'MsgBox ThisWorkbook.VBProject.VBComponents(4).Name
  'MsgBox ActiveWorkbook.VBProject.VBComponents(4).Name
  'MsgBox ActiveSheet.CodeName
  'Application.FindFile
  'VerzListe = "==E:\win32app\ustn_se\mdlsys\asneeded;E:\Programme\Entwicklung\Python\.;E:\WINNT\system32;E:\WINNT;G:\bat\NTFileServer;G:\bat\WinNT;G:\bat;C:\bat;R:\ustn_se\ingr\share;c:\bin\tools\e5;E:\Programme\Leica Geosystems\shared;;U:\Notes\Data;;"
  'VerzListe = ";E:\Programme\Office97\Office\XLStart;;R:\OFFICE97\Excel\xlstart\;;"
  'VerzListe = "R:\OFFICE97\Excel\xlstart\;E:\Programme\Office97\Office\XLStart"
  'DateiName = "_test_user.xlt"
  'DateiName = "cmd.Exe"
  'MsgBox DateiFinden(DateiName, VerzListe)
  'MsgBox "fertig: " & DateiFinden(DateiName, VerzListe, True)
  'MsgBox FindeXLVorlage(DateiName)
  
  'strBezug = "=" & Selection.Address(RowAbsolute:=True, ColumnAbsolute:=True, ReferenceStyle:=xlA1, External:=True)
  'AnzZeilen = Selection.Rows.Count
  'AnzSpalten = Selection.Columns.Count
  'MsgBox strBezug & vbNewLine & AnzZeilen & vbNewLine & AnzSpalten & vbNewLine & isSelectionRechteck
  'AnzRechtecke = Selection.Areas.Count
  'MsgBox isSelectionRechteck
  
  'Set oInfotraeger = GetLokalerZellname(strInfoTraeger)
  'Set ointersect = Intersect(Selection, oInfotraeger)
  'AdrSelection = Selection.Address(RowAbsolute:=True, ColumnAbsolute:=True, ReferenceStyle:=xlA1, External:=True)
  'AdrInfotraeger = oInfotraeger.Address(RowAbsolute:=True, ColumnAbsolute:=True, ReferenceStyle:=xlA1, External:=True)
  'AdrIntersect = ointersect.Address(RowAbsolute:=True, ColumnAbsolute:=True, ReferenceStyle:=xlA1, External:=True)
  
  'AdrSelection = Selection.Address
  'AdrSelection = Union(Selection, Selection).Address
  'AdrInfotraeger = oInfotraeger.Address
  'AdrIntersect = ointersect.Address
  'MsgBox AdrSelection & vbNewLine & AdrInfotraeger & vbNewLine & AdrIntersect
  
  'With Application.CommandBars("Worksheet Menu Bar").Controls("GeoTools").Controls("Format").Controls("Formatierung mit/ohne Streifen")
  '  MsgBox .Caption & ", Tag=" & .Tag & ", Typ=" & .Type
  '  MsgBox Application.CommandBars.FindControl(Type:=msoControlButton, Tag:="FormatDatenMitStreifen").Caption
  '  'MsgBox msoControlButtonPopup
  'End With
  
  'msgbox application.

'AlteStatusLeiste = Application.DisplayStatusBar
'Application.DisplayStatusBar = True
'Application.StatusBar = "Import läuft |||||     |||||     ||||||||||||"
'Workbooks.Open FileName:="GROSS.XLS"
'Application.StatusBar = ""
'Application.DisplayStatusBar = True

'Set oRangeZiel = oInfotraeger.Offset(rowoffset:=1).Resize(RowSize:=ZeAnz - 1)
'ZeEnd = Cells(Rows.Count, Spalte).End(xlUp).Row
'Cells(Rows.Count, 3).End(xlUp).Select

  Dim ur As Range
  Set ur = ActiveWorkbook.ActiveSheet.UsedRange
  LetzteVerwendeteSpalte = ur.Columns(ur.Columns.Count).Column
  MsgBox LetzteVerwendeteSpalte
  
  'MsgBox Application.StartupPath
  'MsgBox Application.TemplatesPath
  'MsgBox Application.NetworkTemplatesPath
  'MsgBox Application.OnWindow
  
End Sub


Private Sub testColl()
  
  Dim oExpim   As Scripting.Dictionary
  Set oExpim = New Scripting.Dictionary
  Dim oLaenge   As Scripting.Dictionary
  Set oLaenge = New Scripting.Dictionary
  Dim oWinkel   As Scripting.Dictionary
  Set oWinkel = New Scripting.Dictionary
  
  oLaenge.Add "m", 1#
  oLaenge.Add "dm", 0.1
  oLaenge.Add "cm", 0.01
  oLaenge.Add "mm", 0.001
  oLaenge.Add "km", 1000#
  
  oWinkel.Add "gon", 1#
  oWinkel.Add "grad", 1 / 0.9
  'oWinkel.add "rad",

  oExpim.Add "Laenge", oLaenge
  oExpim.Add "Winkel", oWinkel

  
  'Kontrolle:
  'For Each Groesse In oExpim
  '  For Each Einheit In oExpim(Groesse)
  '    MsgBox "Einheit=" & Einheit & "  Item=" & oExpim(Groesse)(Einheit) & "  (" & Groesse & ")"
  '  Next
  'Next

End Sub


Private Sub test2()
      
  On Error GoTo 0
  Dim Feld1(2) As String
  
  Dim oDictQuelldaten
  Set oDictQuelldaten = CreateObject("Scripting.Dictionary")
  
  Feld1(0) = "A_0"
  Feld1(1) = "A_1"
  Feld1(2) = "A_2"
  oDictQuelldaten.Add "Feld_A", Feld1
  
  Feld1(0) = "B_0"
  Feld1(1) = "B_1"
  Feld1(2) = "B_2"
  oDictQuelldaten.Add "Feld_B", Feld1
  
  MsgBox "Feld_A(1)=" & oDictQuelldaten("Feld_A")(1) & vbNewLine & "Feld_B(1)=" & oDictQuelldaten("Feld_B")(1)
  
  Set oDictQuelldaten = Nothing
End Sub



Private Sub u()
  On Error GoTo 0
  'dim d as Double
  txt = "Gl 9/D U= -  u= 9"
  'txt = "Gl /D    ="
  MsgBox Ueberhoehung(txt, True)

End Sub



Private Sub testExpim()
  
  Dim oExpim   As CdatExpim
  Set oExpim = New CdatExpim
  
  'oExpim.GetQuelldaten_XlTabAktiv
  oKonfig.ZeigeAlleEinheiten
  
  'MsgBox CStr(vbTab)
  
  Set oExpim = Nothing

End Sub



Sub testExpimTK()
  'On Error GoTo Fehler

  Dim oExpim  As CdatExpim
  Set oExpim = New CdatExpim
  Dim oimpTrassenkoo    As CimpTrassenkoo
  Set oimpTrassenkoo = New CimpTrassenkoo
  
  'DebugMode = True
  
  Dim Dialog    As frmStartExpim
  Set Dialog = New frmStartExpim
    
  Dialog.Show
  'Call AnzeigeMeldungen
  
  'MsgBox "Ergebnis: " & oExpim.Ziel_FormatID

  'If (Err) Then GoTo Fehler
  oExpim.Quelle_AsciiDatei_Name = "x:\QUELLEN\Excel\Trassenkoo\bsp_2000.tx"
  oExpim.Quelle_Typ = io_Typ_AsciiSpezial
  'oExpim.Quelle_FormatID = io_Klasse_Trassenkoo
  oExpim.Quelle_AsciiDatei_DialogFilter = "Logdateien (*.log),*.txt,Alle Dateien (*.*),*.*"
  
  oExpim.Ziel_AsciiDatei_DialogFilter = "Textdateien (*.txt),*.txt,Alle Dateien (*.*),*.*"
  oExpim.Ziel_AsciiDatei_Modus = io_Datei_Modus_Anhaengen
  oExpim.Ziel_AsciiDatei_Name = "ggggg"
  oExpim.Ziel_Typ = io_Typ_XlTabNeu
  'oExpim. = "tK_2"
  
  
  'oExpim.Meldung_Ausgeben = False
  'oExpim.ImportSpezial oimpTrassenkoo
  'oExpim.ZeigeAlleEinheiten
  'MsgBox oExpim.FaktorDerEinheit("")
  'MsgBox oimpTrassenkoo.Quelle_Einheiten(SpN_GK_Y), , "oimpTrassenkoo.Quelle_Einheiten(SpN_GK_Y)"
  'MsgBox oimpTrassenkoo.Quelle_Einheiten("xx"), , "oimpTrassenkoo.Quelle_Einheiten(""xx"")"
  'oExpim.
  'oimpTrassenkoo.
 
  'Call AnzeigeMeldungen
  DebugMode = False
    
  Set oExpim = Nothing
  Set oimpTrassenkoo = Nothing
Exit Sub
  
Fehler:
  FehlerNachricht "xxx.testExpimTK()"
  Set oExpim = Nothing
  Set oimpTrassenkoo = Nothing
End Sub




Private Sub testKat()
  
  'Dim oExpim   As CdatExpim
  
  oAktiveTabelle.Syncronisieren
  MsgBox oAktiveTabelle.Kategorien

End Sub


Private Sub test_vb()
  
  Const PrefixImport  As String = "Cimp"
  Const TypKlasse     As Long = 2

  Dim Modul      As Variant
  Dim AnzModule  As Long
  
  AnzModule = ThisWorkbook.VBProject.VBComponents.Count
  For i = 1 To AnzModule
    With ThisWorkbook.VBProject.VBComponents(i)
      If ((.Type = TypKlasse) And (Left$(.Name, Len(PrefixImport)) = PrefixImport)) Then
        Echo .Name & "    " & .CodeModule.CountOfLines
        '.Activate
      End If
    End With
  Next
  Call ShowConsole

End Sub


Sub test_Mod_XlTabAktiv()
  On Error GoTo Fehler
  Dim oExpim    As CdatExpim
  Set oExpim = New CdatExpim
  oExpim.GetQuelldaten_XlTabAktiv
  
  'oExpim.Datenpuffer.AlleOptionenAus
  oExpim.Datenpuffer.Opt_VorhWerteUeberschreiben = True
  'oExpim.Datenpuffer.Opt_Transfo_Tk2Gls = True
  'oExpim.Datenpuffer.Daten_Bearbeiten
  'oExpim.Datenpuffer.Mod_UeberhoehungAusBemerkung
  oExpim.Datenpuffer.Mod_Transfo_Tk2Gls
  oExpim.Datenpuffer.Mod_FehlerVerbesserung
  
  oExpim.SchreibeDatenpuffer_XlTabAktiv
  Application.StatusBar = ""
  Set oExpim = Nothing
Exit Sub
Fehler:
  Set oExpim = Nothing
  FehlerNachricht "test.test_Mod_XlTabAktiv()"
End Sub


Sub testDialogSpalten()
  
  Dim Dialog    As frmSpaltenVerw
  Set Dialog = New frmSpaltenVerw
    
  Dialog.Show
  
  Set Dialog = Nothing
End Sub
