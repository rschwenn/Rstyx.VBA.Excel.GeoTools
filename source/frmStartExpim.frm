VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStartExpim 
   Caption         =   "Import / Export"
   ClientHeight    =   6132
   ClientLeft      =   40
   ClientTop       =   340
   ClientWidth     =   9800
   OleObjectBlob   =   "frmStartExpim.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmStartExpim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2004-2022  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'==================================================================================================
'Modul frmStartExpim
'==================================================================================================
'
'Dialog zum Erfassen der Parameter für den gewünschten Import/Export,
'der vom Objekt oExpim gestartet wird.
'
'Als Startwerte für den Dialog werden Eigenschaften des Objektes oExpim verwendet.
'Änderungen der Dialogwerte werden wieder umgesetzt in Eigenschaften des Objektes oExpimGlobal.
'==================================================================================================


Option Explicit
'Deklarationen
  Const idxFmtKurzname                          As Integer = 0
  Const idxFmtTitel                             As Integer = 1
  Const idxFmtID                                As Integer = 2
  Const idxFmtKategorien                        As Integer = 3
  Const idxFmtDateifilter                       As Integer = 4
  Const idxFmtIoTyp                             As Integer = 5
  
  Private save_Backcolor_Lst                    As Long
  Private save_Backcolor_TB                     As Long
  Private save_Forecolor_TB                     As Long
  
  Private Ziel_Kategorien                       As String
  Private Ziel_IoTyp                            As String
  Private Ziel_LetztesFormat_XlTabNeu           As String
  Private Ziel_LetztesFormat_AsciiFormatiert    As String
  Private Ziel_LetztesFormat_AsciiSpezial       As String
  Private Ziel_LetzterDateiName                 As String
  Private Ziel_LetzterDateiModusVorhanden       As String
  
  Private Quelle_Kategorien                     As String
  Private Quelle_LetztesFormat_AsciiFormatiert  As String
  Private Quelle_LetztesFormat_AsciiSpezial     As String
  
  Private Liste_XLT_komplett()                  As String
  Private Liste_SpezialImport_komplett()        As String
  Private Liste_Spaltennamen_XlTabAktiv()       As String
  Private Liste_Spaltennamen_CsvSpezial()       As String
  
  Private Sperre_chkMod_Anwenden                  As Boolean
  Private Sperre_chkMod_VorhWerteUeberschreiben   As Boolean
  Private Sperre_chkMod_Ersatzspalten             As Boolean
  
  Private oRecentXLT                            As Scripting.Dictionary
  Private oRecentXLT_Neu                        As Scripting.Dictionary
'



Private Sub UserForm_Initialize()
  'Initialisierung des Dialoges
  
  'Deklarationen
  Dim oCsvSpezial  As CtabCSV
  
  'Dialog-Titel
  Me.Caption = ProgName & " - Import / Export"
  
  
  '***********************************************************************
  'yyy Vorläufig: ungenutzte Optionen ausblenden:
    Me.optQuelle_Typ_AsciiFormatiert.Enabled = False
    Me.optQuelle_Typ_AsciiFormatiert.Visible = False
    Me.optZiel_Typ_AsciiFormatiert.Enabled = False
    Me.optZiel_Typ_AsciiFormatiert.Visible = False
    Me.optZiel_Typ_AsciiSpezial.Enabled = False
    Me.optZiel_Typ_AsciiSpezial.Visible = False
    'Me.optZiel_Typ_XLTabNeu.value = True
  '***********************************************************************
  
  
  oExpimGlobal.Dialog_OK = False
  
  'Listboxen vorformatieren
  Me.LstQuelle_Formate.ColumnCount = 6
  Me.LstQuelle_Formate.ColumnWidths = "110;;0;0;0;0"
  Me.LstZiel_Formate.ColumnCount = 6
  Me.LstZiel_Formate.ColumnWidths = "110;;0;0;0;0"
  
  'Aktive Hintergrundfarben sichern.
  save_Backcolor_Lst = Me.LstQuelle_Formate.BackColor
  save_Backcolor_TB = Me.tbQuelle_AsciiDatei.BackColor
  save_Forecolor_TB = Me.tbQuelle_AsciiDatei.ForeColor
  
  'Dictionaries für Liste der zuletzt vorhandenen XL-Vorlagen
  Set oRecentXLT = New Scripting.Dictionary
  Set oRecentXLT_Neu = New Scripting.Dictionary
  
  'Formatlisten erzeugen und merken (Cache und persistent).
    'Alle XLT-Vorlagen
    If (Not ThisWorkbook.Konfig.Cache.Exists("Liste_XLT_komplett")) Then
      Call GetFormatliste_XlVorlagen
      ThisWorkbook.Konfig.Cache.Add "Liste_XLT_komplett", Liste_XLT_komplett
    End If
    
    'Alle Programmodule für ASCII-Spezialimport
    If (Not ThisWorkbook.Konfig.Cache.Exists("Liste_SpezialImport_komplett")) Then
      Call GetFormatliste_SpezialImport
      ThisWorkbook.Konfig.Cache.Add "Liste_SpezialImport_komplett", Liste_SpezialImport_komplett
    End If
    
    'Alle Spaltennamen der aktiven XL-Tabelle
    Call GetFormatliste_Spaltennamen_XlTabAktiv
  '
  'Dateifilter für CSV-Spezialimport
    If (Not ThisWorkbook.Konfig.Cache.Exists("Quelle_CsvSpezial_DialogFilter")) Then
      Set oCsvSpezial = New CtabCSV
      ThisWorkbook.Konfig.Cache.Add "Quelle_CsvSpezial_DialogFilter", oCsvSpezial.Quelle_AsciiDatei_DialogFilter
      Set oCsvSpezial = Nothing
    End If
  
  'Anfangswerte setzen
  Call SetzeAnfangswerte
  
  Call ClearStatusBarDelayed(StatusBarClearDelay)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If (CloseMode = vbFormControlMenu) Then
    'Dialog wird über das Systemmenü geschlossen
    'zu spät...
  End If
End Sub



'==== Methoden ================================================================

Public Function Check_DialogOK() As Boolean
  'True, wenn alle Felder des Dialoges gültige Werte besitzen,
  'd.h. der Dialog wurde sinnvoll ausgefüllt.
  'Der Status des "OK"-Buttons wird geschaltet.
  'Der Hinweistext wird gesetzt.
  
  Dim blnGueltig  As Boolean
  
  'MsgBox "Check_DialogOK"
  
  If (Not isGueltig_QuellTyp) Then
    blnGueltig = False
    Me.lblHinweis = "Typ der Quelldaten wählen."
  
  ElseIf (Not isGueltig_QuellFormat) Then
    blnGueltig = False
    Me.lblHinweis = "Format der Quelldaten wählen."
  
  ElseIf (Not isGueltig_Quelldatei) Then
    blnGueltig = False
    Me.lblHinweis = "ASCII-Quelldatei wählen."
  
  ElseIf (Not isGueltig_ZielTyp) Then
    blnGueltig = False
    Me.lblHinweis = "Typ der Zieldaten wählen."
  
  ElseIf (Not isGueltig_ZielFormat) Then
    blnGueltig = False
    Me.lblHinweis = "Format der Zieldaten wählen."
  
  ElseIf (Not isGueltig_Zieldatei) Then
    blnGueltig = False
    Me.lblHinweis = "ASCII-Zieldatei wählen (wenn Rot, dann ungültiges Verzeichnis!)."
  
  Else
    blnGueltig = True
  End If
  
  If (blnGueltig) Then
    Me.LstZiel_Formate.SetFocus
    Me.btnOK.Enabled = True
    Me.lblHinweis = "Import/Export kann gestartet werden."
  Else
    'MsgBox Me.lblHinweis, vbOKOnly, "Fehler beim Ausfüllen des Dialoges."
    Me.btnOK.Enabled = False
  End If
  
  Check_DialogOK = blnGueltig
  
End Function



'==== interne Routinen =========================================================

Private Sub btnQuelle_AsciiDatei_Click()
  'Dateiauswahl
  Dim Dateiname   As String
  Dateiname = GetQuellDateinameAusDialog()
  If (Dateiname <> "") Then
    'Pfad\Name einer existierenden Datei ermittelt.
    Me.tbQuelle_AsciiDatei.value = Dateiname
  End If
  Me.tbQuelle_AsciiDatei.SetFocus
End Sub

Private Sub btnQuelle_AsciiDateiEdit_Click()
  'Eingabedatei bearbeiten.
  If (Not ThisWorkbook.SysTools.StartEditor("""" & Me.tbQuelle_AsciiDatei.value & """")) Then
    ThisWorkbook.SysTools.StarteDatei(Me.tbQuelle_AsciiDatei.value)
  End If
End Sub


'Private Sub btnZiel_AsciiDatei_Click()
  'Dim Dateiname   As String
  'Dateiname = GetZielDateinameAusDialog()
  'If (Dateiname <> "") Then
  '  'Pfad\Name einer Datei in einem exist. Verz. ermittelt, die nicht existieren muß.
  '  Me.tbZiel_AsciiDatei.value = Dateiname
  'End If
  'Me.tbZiel_AsciiDatei.SetFocus
'End Sub

Private Sub optQuelle_Typ_XLTabAktiv_Change()
  'Typ der Datenquelle im Dialog geändert
  If (Me.optQuelle_Typ_XLTabAktiv.value) Then
    oExpimGlobal.Quelle_Typ = io_Typ_XlTabAktiv
    Call Changed_Quelle_Typ
    'Call Check_DialogOK
  End If
End Sub

Private Sub optQuelle_Typ_AsciiFormatiert_Change()
  'Typ der Datenquelle im Dialog geändert
  If (Me.optQuelle_Typ_AsciiFormatiert.value) Then
    oExpimGlobal.Quelle_Typ = io_Typ_AsciiFormatiert
    Call Changed_Quelle_Typ
    'Call Check_DialogOK
  End If
End Sub

Private Sub optQuelle_Typ_CsvSpezial_Change()
  'Typ der Datenquelle im Dialog geändert
  If (Me.optQuelle_Typ_CsvSpezial.value) Then
    oExpimGlobal.Quelle_Typ = io_Typ_CsvSpezial
    Call Changed_Quelle_Typ
    'Call Check_DialogOK
  Else
    'Einstellungen für Datenübertragung und -Bearbeitung zurücksetzen
    Call SetzeStandardDatenEinstellungen
  End If
End Sub

Private Sub optQuelle_Typ_AsciiSpezial_Change()
  'Typ der Datenquelle im Dialog geändert
  If (Me.optQuelle_Typ_AsciiSpezial.value) Then
    oExpimGlobal.Quelle_Typ = io_Typ_AsciiSpezial
    Call Changed_Quelle_Typ
    'Call Check_DialogOK
  End If
End Sub

Private Sub optZiel_Typ_XLTabNeu_Change()
  If (Me.optZiel_Typ_XLTabNeu.value) Then
    oExpimGlobal.Ziel_Typ = io_Typ_XlTabNeu
    Call Changed_Ziel_Typ
    'Call Check_DialogOK
  End If
End Sub

Private Sub optZiel_Typ_AsciiFormatiert_Change()
  If (Me.optZiel_Typ_AsciiFormatiert.value) Then
    oExpimGlobal.Ziel_Typ = io_Typ_AsciiFormatiert
    Call Changed_Ziel_Typ
    'Call Check_DialogOK
  End If
End Sub

Private Sub optZiel_Typ_AsciiSpezial_Change()
  If (Me.optZiel_Typ_AsciiSpezial.value) Then
    oExpimGlobal.Ziel_Typ = io_Typ_AsciiSpezial
    Call Changed_Ziel_Typ
    'Call Check_DialogOK
  End If
End Sub

'Private Sub optZiel_AsciiDatei_Neu_Change()
 '  If (Me.optZiel_AsciiDatei_Neu.value) Then
 '    oExpimGlobal.Ziel_AsciiDatei_Modus = io_Datei_Modus_Neu
 '    'Call Check_DialogOK
 '  End If
'End Sub

'Private Sub optZiel_AsciiDatei_Ueberschreiben_Change()
 '  If (Me.optZiel_AsciiDatei_Ueberschreiben.value) Then
 '    oExpimGlobal.Ziel_AsciiDatei_Modus = io_Datei_Modus_Ueberschreiben
 '    Ziel_LetzterDateiModusVorhanden = io_Datei_Modus_Ueberschreiben
 '    'Call Check_DialogOK
 '  End If
'End Sub

'Private Sub optZiel_AsciiDatei_Anhaengen_Change()
 '  If (Me.optZiel_AsciiDatei_Anhaengen.value) Then
 '    oExpimGlobal.Ziel_AsciiDatei_Modus = io_Datei_Modus_Anhaengen
 '    Ziel_LetzterDateiModusVorhanden = io_Datei_Modus_Anhaengen
 '    'Call Check_DialogOK
 '  End If
'End Sub


Private Sub tbQuelle_AsciiDatei_Change()
  'MsgBox "tbQuelle_AsciiDatei_Change"
  If (ThisWorkbook.SysTools.IsDatei(Me.tbQuelle_AsciiDatei.value)) Then
    'If (Me.tbQuelle_AsciiDatei.value = Me.tbZiel_AsciiDatei.value) Then
    '  oExpimGlobal.Quelle_AsciiDatei_Name = ""
    '  Me.tbQuelle_AsciiDatei.ForeColor = &H80FF&
    '  'Me.tbZiel_AsciiDatei.ForeColor = &H80FF&
    'Else
      oExpimGlobal.Quelle_AsciiDatei_Name = Me.tbQuelle_AsciiDatei.value
      Me.tbQuelle_AsciiDatei.ForeColor = save_Forecolor_TB
      Me.btnQuelle_AsciiDateiEdit.Enabled = True
      'MsgBox "tbQuelle_AsciiDatei_Change,  Datei=" & oExpimGlobal.Quelle_AsciiDatei_Name
    'End If
  Else
    oExpimGlobal.Quelle_AsciiDatei_Name = ""
    Me.tbQuelle_AsciiDatei.ForeColor = &HFF&
    Me.btnQuelle_AsciiDateiEdit.Enabled = False
    'MsgBox "tbQuelle_AsciiDatei_Change,  Datei=" & oExpimGlobal.Quelle_AsciiDatei_Name
  End If
  
  'Weitere Aktualisierungen bei CSV-Spezialimport
  If (oExpimGlobal.Quelle_Typ = io_Typ_CsvSpezial) Then Call Changed_CsvDateiname
  
  Call Check_DialogOK
End Sub

Private Sub tbQuelle_AsciiDatei_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  'Wenn das Verzeichnis der eingegebenen Datei nicht existiert bzw.
  'kein Verzeichnis im Namen enthalten ist, so wird das Arbeitsverzeichnis
  'dem Dateinamen vorangestellt.
  'MsgBox "Ereignis  tbQuelle_AsciiDatei_Exit"
  Call AsciiDatei_PfadKorrigieren(Me.tbQuelle_AsciiDatei)
End Sub

Private Sub fraQuelle_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  'Ersetzt "tbQuelle_AsciiDatei_Exit", wenn gleichzeitig der Frame verlassen wird.
  Call AsciiDatei_PfadKorrigieren(Me.tbQuelle_AsciiDatei)
End Sub



Private Sub chkMod_Anwenden_Change()
  oExpimGlobal.Opt_DatenModifizieren = chkMod_Anwenden.value
  If (oExpimGlobal.Opt_DatenModifizieren) Then
    Me.chkMod_VorhWerteUeberschreiben.Enabled = True
  Else
    Me.chkMod_VorhWerteUeberschreiben.Enabled = False
  End If
End Sub

Private Sub chkMod_VorhWerteUeberschreiben_Click()
  oExpimGlobal.Datenpuffer.Opt_VorhWerteUeberschreiben = chkMod_VorhWerteUeberschreiben.value
End Sub

Private Sub chkMod_Ersatzspalten_Click()
  oExpimGlobal.Opt_ErsatzZielspaltenVerwenden = chkMod_Ersatzspalten.value
  'MsgBox oExpimGlobal.Opt_ErsatzZielspaltenVerwenden
End Sub

Private Sub chkMod_FormelnErhalten_Click()
  oExpimGlobal.Opt_FormelnErhalten = chkMod_FormelnErhalten.value
End Sub


'Private Sub tbZiel_AsciiDatei_Change()
  ''MsgBox "Ereignis  tbZiel_AsciiDatei_Change"
  '
  'Dim Verzeichnis  As String
  'Dim NameMitExt   As String
  '
  'Verzeichnis = Verz(Me.tbZiel_AsciiDatei.value)
  'If (Verzeichnis = "") Then Verzeichnis = "NichtVorhanden"
  'NameMitExt = NameExt(Me.tbZiel_AsciiDatei.value, "mitext")
  '
  'If (ThisWorkbook.SysTools.IsDatei(Me.tbZiel_AsciiDatei.value)) Then
  '  'Datei ist vorhanden.
  '  'If (Me.tbQuelle_AsciiDatei.value = Me.tbZiel_AsciiDatei.value) Then
  '  '  oExpimGlobal.Ziel_AsciiDatei_Name = ""
  '  '  Me.tbZiel_AsciiDatei.ForeColor = &H80FF&
  '  '  'Me.tbQuelle_AsciiDatei.ForeColor = &H80FF&
  '  '  Me.optZiel_AsciiDatei_Ueberschreiben.Enabled = False
  '  '  Me.optZiel_AsciiDatei_Anhaengen.Enabled = False
  '  '  Me.optZiel_AsciiDatei_Neu.Enabled = False
  '  'Else
  '    oExpimGlobal.Ziel_AsciiDatei_Name = Me.tbZiel_AsciiDatei.value
  '    Me.tbZiel_AsciiDatei.ForeColor = save_Forecolor_TB
  '    Me.optZiel_AsciiDatei_Ueberschreiben.Enabled = True
  '    Me.optZiel_AsciiDatei_Anhaengen.Enabled = True
  '    Me.optZiel_AsciiDatei_Neu.Enabled = False
  '
  '    'If (oExpimGlobal.Ziel_AsciiDatei_Name = Ziel_LetzterDateiName) Then
  '      If (Ziel_LetzterDateiModusVorhanden = io_Datei_Modus_Anhaengen) Then
  '        Me.optZiel_AsciiDatei_Anhaengen.value = True
  '      Else
  '        Me.optZiel_AsciiDatei_Ueberschreiben.value = True
  '      End If
  '    'Else
  '    '  Me.optZiel_AsciiDatei_Ueberschreiben.value = True
  '    'End If
  '  'End If
  '
  'ElseIf (ThisWorkbook.SysTools.IsVerzeichnis(Verzeichnis) And (Not ThisWorkbook.SysTools.IsVerzeichnis(Me.tbZiel_AsciiDatei.value))) Then
  '  'Neue Datei.
  '  '(Datei ist nicht vorhanden, aber angegeben. Das Verzeichnis existiert.)
  '  If (VorName(Me.tbZiel_AsciiDatei.value) = "") Then
  '    oExpimGlobal.Ziel_AsciiDatei_Name = ""
  '    Me.tbZiel_AsciiDatei.ForeColor = &HFF&
  '    Me.optZiel_AsciiDatei_Ueberschreiben.Enabled = False
  '    Me.optZiel_AsciiDatei_Anhaengen.Enabled = False
  '    Me.optZiel_AsciiDatei_Neu.Enabled = False
  '  Else
  '    oExpimGlobal.Ziel_AsciiDatei_Name = Me.tbZiel_AsciiDatei.value
  '    Me.tbZiel_AsciiDatei.ForeColor = save_Forecolor_TB
  '    Me.optZiel_AsciiDatei_Neu.value = True
  '    Me.optZiel_AsciiDatei_Neu.Enabled = True
  '    Me.optZiel_AsciiDatei_Ueberschreiben.Enabled = False
  '    Me.optZiel_AsciiDatei_Anhaengen.Enabled = False
  '  End If
  '
  'Else
  '  'Verzeichnis ist nicht vorhanden.
  '  oExpimGlobal.Ziel_AsciiDatei_Name = ""
  '  Me.tbZiel_AsciiDatei.ForeColor = &HFF&
  '  Me.optZiel_AsciiDatei_Neu.value = True
  '  Me.optZiel_AsciiDatei_Neu.Enabled = False
  '  Me.optZiel_AsciiDatei_Ueberschreiben.Enabled = False
  '  Me.optZiel_AsciiDatei_Anhaengen.Enabled = False
  '
  'End If
  'Ziel_LetzterDateiName = oExpimGlobal.Ziel_AsciiDatei_Name
  '
  'Call Check_DialogOK
'End Sub


'Private Sub tbZiel_AsciiDatei_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  ''Wenn das Verzeichnis der eingegebenen Datei nicht existiert bzw.
  ''kein Verzeichnis im Namen enthalten ist, so wird das Arbeitsverzeichnis
  ''dem Dateinamen vorangestellt.
  'Call AsciiDatei_PfadKorrigieren(Me.tbZiel_AsciiDatei)
'End Sub

'Private Sub fraziel_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  ''Ersetzt "tbziel_AsciiDatei_Exit", wenn gleichzeitig der Frame verlassen wird.
  'Call AsciiDatei_PfadKorrigieren(Me.tbZiel_AsciiDatei)
'End Sub


Private Sub AsciiDatei_PfadKorrigieren(tbDateiname As MSForms.TextBox)
  'Wenn das Verzeichnis der in der Textbox eingegebenen Datei nicht existiert bzw.
  'kein Verzeichnis im Namen enthalten ist, so wird das Arbeitsverzeichnis
  'dem Dateinamen vorangestellt.
  Dim Verzeichnis  As String
  Dim PfadName     As String
  Dim TB_Wert      As String
  TB_Wert = tbDateiname.value
  If (TB_Wert <> "") Then
    Verzeichnis = Verz(TB_Wert)
    If (Verzeichnis = "") Then Verzeichnis = "NichtVorhanden"
    If (Dir(Verzeichnis & "\", vbDirectory) = "") Then
      tbDateiname.value = CurDir & "\" & NameExt(TB_Wert, "mitext")
    End If
  End If
End Sub


Private Sub btnAbbruch_Click()
  'Me.Hide
  Unload Me
End Sub


Private Sub btnOK_Click()
  oExpimGlobal.Dialog_OK = True
  Me.Hide
  'Unload Me
End Sub


Private Sub Changed_Quelle_Typ()
  'Neuer Typ der Datenquelle wurde gesetzt ==> Liste der Quell-Formate füllen.
  'Inhalt ist abhängig von:
  '  - oExpimGlobal.Quelle_Typ (io_Typ_XlTabAktiv, io_Typ_AsciiFormatiert, io_Typ_AsciiSpezial)
  
  On Error GoTo 0
  
  Dim FormatLetzteWahl     As String
  Dim DateiMaske           As String
  Dim Formatliste          As String
  Dim VorlageStandard      As String
  Dim NF                   As Long
  Dim i                    As Long
  Dim idxVorauswahl        As Integer
  Dim VorlagenVorName      As String
  Dim FormatPfadName()     As String
  Dim TmpList()            As String
  Dim Feld()               As String
  
  DebugEcho vbNewLine & "Changed_Quelle_Typ(): Neuer Wert = " & oExpimGlobal.Quelle_Typ
  
  Select Case oExpimGlobal.Quelle_Typ
    
    Case io_Typ_XlTabAktiv
        
        'Keine Formatauswahl nötig.
        Me.LstQuelle_Formate.Clear
        
        'Kategorien der aktiven Tabelle als Filterkriterium für die Liste der Zielformate.
        'ThisWorkbook.AktiveTabelle.Syncronisieren
        Ziel_Kategorien = ThisWorkbook.AktiveTabelle.Kategorien
        Call GetQuellFormatliste(Liste_Spaltennamen_XlTabAktiv, "#")
        
        'Wenn keine Spalte verfügbar ist, gibt's nichts zu exportieren.
        NF = ThisWorkbook.AktiveTabelle.SpaltenErsteZellen.Count
        If (NF > 0) Then
          Ziel_IoTyp = io_Typ_Puffer
          Me.lblQuelle_Formate.Caption = "Die Tabelle enthält folgende bezeichneten Spalten:"
          Me.LstQuelle_Formate.ControlTipText = "Diese Werte stehen zum Export zur Verfügung"
          Me.fraMod.Enabled = True
          Me.chkMod_Anwenden.Visible = True
          Me.chkMod_VorhWerteUeberschreiben.Visible = True
          Me.chkMod_Ersatzspalten.Visible = True
          Me.optZiel_Typ_XLTabNeu.Enabled = True
          'yyy Me.optZiel_Typ_AsciiFormatiert.Enabled = True
          'yyy Me.optZiel_Typ_AsciiSpezial.Enabled = True
          Me.chkZiel_Formate.Enabled = True
        Else
          Ziel_IoTyp = io_Typ_AsciiSpezial
          Me.lblQuelle_Formate.Caption = "Die Tabelle enthält keine bezeichneten Spalten."
          'Me.LstQuelle_Formate.ControlTipText = "... deshalb kommt nur Spezialexport in Frage."
          Me.LstQuelle_Formate.ControlTipText = "... kein Export möglich"
          Me.fraMod.Enabled = False
          Me.chkMod_Anwenden.Visible = False
          Me.chkMod_VorhWerteUeberschreiben.Visible = False
          Me.chkMod_Ersatzspalten.Visible = False
          Me.optZiel_Typ_XLTabNeu.Enabled = False
          Me.optZiel_Typ_AsciiFormatiert.Enabled = False
          Me.optZiel_Typ_AsciiSpezial.Enabled = False
          'Me.optZiel_Typ_AsciiSpezial.value = True
          Me.chkZiel_Formate.value = False
          Me.chkZiel_Formate.Enabled = False
        End If
        
        'Me.LstQuelle_Formate.Enabled = False
        'Me.LstQuelle_Formate.Locked = True
        Me.LstQuelle_Formate.BackColor = Me.BackColor
        Me.btnQuelle_AsciiDatei.Enabled = False
        Me.btnQuelle_AsciiDateiEdit.Enabled = False
        Me.tbQuelle_AsciiDatei.Enabled = False
        Me.tbQuelle_AsciiDatei.BackColor = Me.BackColor
        
        
    Case io_Typ_CsvSpezial
        
        Ziel_IoTyp = io_Typ_Puffer
        
        'Me.LstQuelle_Formate.Enabled = False
        'Me.LstQuelle_Formate.Locked = True
        Me.LstQuelle_Formate.BackColor = Me.BackColor
        Me.btnQuelle_AsciiDatei.Enabled = True
        Me.tbQuelle_AsciiDatei.Enabled = True
        Me.tbQuelle_AsciiDatei.BackColor = save_Backcolor_TB
        
        'Dateifilter setzen
        oExpimGlobal.Quelle_AsciiDatei_DialogFilter = ThisWorkbook.Konfig.Cache.Item("Quelle_CsvSpezial_DialogFilter")
        
        'Alle weiteren, vom  Dateiinhalt abhängigen GUI-Aktualisierungen.
        Call Changed_CsvDateiname
        
        
    Case io_Typ_AsciiFormatiert
        
        'Import-Formatdateien (Liste = DateiVorname, Formatbeschreibung, Pfad\Name-unsichtbar).
        'Formatliste = AsciiFormatListe()
        'NF = SplitDelim(Formatliste, FormatPfadName, ";")
        '
        'DebugEcho vbNewLine & "Changed_Quelle_Typ: " & NF & " Formate in '" & Formatliste & "'"
        '
        ''3 Spalten (s.o.), PfadName als Wert, aber nicht anzeigen.
        'LstQuelle_Formate.ColumnCount = 3
        'LstQuelle_Formate.ColumnWidths = ";;0"
        'ReDim TmpList(0 To UBound(FormatPfadName) - 1, 0 To LstQuelle_Formate.ColumnCount - 1) As String
        'idxVorauswahl = -1
        'For i = LBound(FormatPfadName) To NF
        '  VorlagenVorName = VorName(FormatPfadName(i))
        '  TmpList(i - 1, 0) = VorlagenVorName
        '  TmpList(i - 1, 1) = FormatPfadName(i)
        '  If (LCase(VorlagenVorName) = LCase(VorlageStandard)) Then idxVorauswahl = i - 1
        'Next
        'LstQuelle_Formate.List = TmpList
        'LstQuelle_Formate.BoundColumn = 2
        'If (idxVorauswahl > -1) Then LstQuelle_Formate.Selected(idxVorauswahl) = True
        Me.LstQuelle_Formate.Clear
        
        Me.optZiel_Typ_XLTabNeu.value = True
        Me.optZiel_Typ_XLTabNeu.Enabled = True
        'yyy Me.optZiel_Typ_AsciiFormatiert.Enabled = True
        'yyy Me.optZiel_Typ_AsciiSpezial.Enabled = True
        
        Me.lblQuelle_Formate.Caption = "Verfügbare ASCII-Formate:"
        Me.lblQuelle_Formate.Visible = True
        Me.LstQuelle_Formate.ControlTipText = "ASCII-Formate für die Quelldatei"
        Me.LstQuelle_Formate.Enabled = True
        Me.LstQuelle_Formate.BackColor = save_Backcolor_Lst
        Me.btnQuelle_AsciiDatei.Enabled = True
        If (ThisWorkbook.SysTools.IsDatei(Me.tbQuelle_AsciiDatei.value)) Then Me.btnQuelle_AsciiDateiEdit.Enabled = True
        Me.tbQuelle_AsciiDatei.Enabled = True
        Me.tbQuelle_AsciiDatei.BackColor = save_Backcolor_TB
        
        Ziel_IoTyp = io_Typ_Puffer
        Ziel_Kategorien = "#"
       
       
    Case io_Typ_AsciiSpezial
        
        'Layout des Dialoges.
        Me.lblQuelle_Formate.Caption = "Verfügbare spezielle ASCII-Formate:"
        Me.lblQuelle_Formate.Visible = True
        
        Me.LstQuelle_Formate.ControlTipText = "spezielle ASCII-Formate für die Quelldatei"
        Me.LstQuelle_Formate.Enabled = True
        Me.LstQuelle_Formate.BackColor = save_Backcolor_Lst
        
        Me.btnQuelle_AsciiDatei.Enabled = True
        If (ThisWorkbook.SysTools.IsDatei(Me.tbQuelle_AsciiDatei.value)) Then Me.btnQuelle_AsciiDateiEdit.Enabled = True
        Me.tbQuelle_AsciiDatei.Enabled = True
        Me.tbQuelle_AsciiDatei.BackColor = save_Backcolor_TB
        
        'Liste ImportKlassen.
        FormatLetzteWahl = Mid$(Quelle_LetztesFormat_AsciiSpezial, Len(io_Klasse_PrefixImport) + 1)
        Ziel_IoTyp = "#"
        Ziel_Kategorien = "#"
        Call GetQuellFormatliste(ThisWorkbook.Konfig.Cache.Item("Liste_SpezialImport_komplett"), FormatLetzteWahl)
        'Call GetQuellFormatliste(Liste_SpezialImport_komplett, FormatLetzteWahl)
        
        
    Case Else
        
        LstQuelle_Formate.Clear
        Ziel_IoTyp = "#"
        Ziel_Kategorien = "#"
        
        Me.lblQuelle_Formate.Caption = ""
        Me.LstQuelle_Formate.Enabled = False
        Me.LstQuelle_Formate.ControlTipText = ""
        Me.LstQuelle_Formate.BackColor = Me.BackColor
        Me.btnQuelle_AsciiDatei.Enabled = False
        Me.btnQuelle_AsciiDateiEdit.Enabled = False
        Me.tbQuelle_AsciiDatei.Enabled = False
        Me.tbQuelle_AsciiDatei.BackColor = Me.BackColor
        
  End Select
  
  'Bei io_Typ_CsvSpezial wird Changed_Ziel_Typ() von Changed_CsvDateiname() aufgerufen...
  If (oExpimGlobal.Quelle_Typ <> io_Typ_CsvSpezial) Then Call Changed_Ziel_Typ
  
End Sub


Private Sub Changed_Ziel_Typ()
  'Liste der Ziel-Formate bzw. Vorlagen füllen.
  'Inhalt ist abhängig von:
  '  - oExpimGlobal.Ziel_Typ (io_Typ_XlTabNeu, io_Typ_AsciiFormatiert, io_Typ_AsciiSpezial)
  
  On Error GoTo 0
  
  Dim FormatLetzteWahl     As String
  
  DebugEcho "Changed_Ziel_Typ(): Neuer Wert = " & oExpimGlobal.Ziel_Typ
  
  Select Case oExpimGlobal.Ziel_Typ
    
    Case io_Typ_XlTabNeu
        
        'Layout des Dialoges.
        Me.lblZiel_Formate.Caption = "Verfügbare Tabellenvorlagen:"
        Me.LstZiel_Formate.Enabled = True
        Me.LstZiel_Formate.ControlTipText = "Tabellenvorlagen für die Zieldatei"
        'Setzen der Hintergrundfarbe macht Auswahl unsichtbar, deshalb vorher setzen.
        Me.LstZiel_Formate.BackColor = save_Backcolor_Lst
        'Me.btnZiel_AsciiDatei.Enabled = False
        'Me.tbZiel_AsciiDatei.Enabled = False
        'Me.tbZiel_AsciiDatei.BackColor = Me.BackColor
        'Me.optZiel_AsciiDatei_Neu.Enabled = False
        'Me.optZiel_AsciiDatei_Ueberschreiben.Enabled = False
        'Me.optZiel_AsciiDatei_Anhaengen.Enabled = False
        
        'Excel-Vorlagenliste.
        FormatLetzteWahl = VorName(Ziel_LetztesFormat_XlTabNeu)
        'Call GetZielFormatliste(Liste_XLT_komplett, FormatLetzteWahl)
        Call GetZielFormatliste(ThisWorkbook.Konfig.Cache.Item("Liste_XLT_komplett"), FormatLetzteWahl)
        
        
    Case io_Typ_AsciiFormatiert
        
        'Export-Formatdateien (Liste = DateiVorname, Formatbeschreibung, Pfad\Name-unsichtbar).
        'Formatliste = AsciiFormatListe()
        'NF = SplitDelim(Formatliste, FormatPfadName, ";")
        '
        'DebugEcho vbNewLine & "Changed_Ziel_Typ: " & NF & " Formate in '" & Formatliste & "'"
        '
        ''3 Spalten (s.o.), PfadName als Wert, aber nicht anzeigen.
        'LstZiel_Formate.ColumnCount = 3
        'LstZiel_Formate.ColumnWidths = ";;0"
        'ReDim TmpList(0 To UBound(FormatPfadName) - 1, 0 To LstZiel_Formate.ColumnCount - 1) As String
        'idxStandard = -1
        'For i = LBound(FormatPfadName) To NF
        '  FormatVorName = VorName(FormatPfadName(i))
        '  TmpList(i - 1, 0) = FormatVorName
        '  TmpList(i - 1, 1) = FormatPfadName(i)
        'Next
        'LstZiel_Formate.List = TmpList
        'LstZiel_Formate.BoundColumn = 2
        'If (idxStandard > -1) Then LstZiel_Formate.Selected(idxStandard) = True
        LstZiel_Formate.Clear
        'LstZiel_Formate.value = ""
        
        Me.lblZiel_Formate.Caption = "Verfügbare ASCII-Formate:"
        Me.LstZiel_Formate.Enabled = True
        Me.LstZiel_Formate.ControlTipText = "ASCII-Formate für die Zieldatei"
        Me.LstZiel_Formate.BackColor = save_Backcolor_Lst
        'Me.btnZiel_AsciiDatei.Enabled = True
        'Me.tbZiel_AsciiDatei.Enabled = True
        'Me.tbZiel_AsciiDatei.BackColor = save_Backcolor_TB
        'Call tbZiel_AsciiDatei_Change
        
        
    Case io_Typ_AsciiSpezial
        
        'Exportobjekte (Liste = Objektname kurz, Beschreibung, ID-Konstante-unsichtbar).
        'Formatliste = AsciiSpezialFormatListe()
        LstZiel_Formate.Clear
        'LstZiel_Formate.value = ""
        
        Me.lblZiel_Formate.Caption = "Verfügbare spezielle ASCII-Formate:"
        Me.LstZiel_Formate.Enabled = True
        Me.LstZiel_Formate.ControlTipText = "spezielle ASCII-Formate für die Zieldatei"
        Me.LstZiel_Formate.BackColor = save_Backcolor_Lst
        'Me.btnZiel_AsciiDatei.Enabled = True
        'Me.tbZiel_AsciiDatei.Enabled = True
        'Me.tbZiel_AsciiDatei.BackColor = save_Backcolor_TB
        'Call tbZiel_AsciiDatei_Change
      
      
    Case Else
      
        LstZiel_Formate.Clear
        'LstZiel_Formate.value = ""
        
        Me.lblZiel_Formate.Caption = ""
        Me.LstZiel_Formate.Enabled = False
        Me.LstZiel_Formate.ControlTipText = ""
        Me.LstZiel_Formate.BackColor = Me.BackColor
        'Me.btnZiel_AsciiDatei.Enabled = False
        'Me.tbZiel_AsciiDatei.Enabled = False
        'Me.tbZiel_AsciiDatei.BackColor = Me.BackColor
        'Me.optZiel_AsciiDatei_Neu.Enabled = False
        'Me.optZiel_AsciiDatei_Ueberschreiben.Enabled = False
        'Me.optZiel_AsciiDatei_Anhaengen.Enabled = False
        
  End Select
  
  If ((oExpimGlobal.Quelle_Typ = io_Typ_XlTabAktiv) And (oExpimGlobal.Ziel_Typ = io_Typ_XlTabNeu)) Then
    Me.chkMod_FormelnErhalten.Visible = True
    Me.chkMod_FormelnErhalten.Enabled = True
  Else
    Me.chkMod_FormelnErhalten.Visible = False
    Me.chkMod_FormelnErhalten.Enabled = False
    'oExpimGlobal.Opt_FormelnErhalten  = False  => wird in CdatExpim.AktionsManager gemacht.
  End If
  
  Call Check_DialogOK
  
End Sub


Private Sub LstQuelle_Formate_Change()
  'Ereignis wird beim Klicken nur ausgelöst, wenn die Liste auch Einträge hat,
  'aber auch durch .BoundColumn und .Selected.
  
  On Error GoTo 0
  
  DebugEcho "LstQuelle_Formate_Change(): Neuer Wert = " & Me.LstQuelle_Formate.value
  
  'Aktuelle Auswahl als Eigenschaft von oExpim speichern.
  If (IsNull(Me.LstQuelle_Formate.value)) Then
    
    oExpimGlobal.Quelle_FormatID = ""
    'oExpimGlobal.Quelle_AsciiDatei_DialogFilter = ""
    
  Else
    
    '***oExpimGlobal.Quelle_FormatID = Me.LstQuelle_Formate.value
    '***oExpimGlobal.Quelle_AsciiDatei_DialogFilter = Me.LstQuelle_Formate.List(Me.LstQuelle_Formate.ListIndex, idxFmtDateifilter)
    
    Select Case oExpimGlobal.Quelle_Typ
      
      Case io_Typ_XlTabAktiv
          
      Case io_Typ_CsvSpezial
          
          
      Case io_Typ_AsciiFormatiert
          
      Case io_Typ_AsciiSpezial
          oExpimGlobal.Quelle_FormatID = Me.LstQuelle_Formate.value
          oExpimGlobal.Quelle_AsciiDatei_DialogFilter = Me.LstQuelle_Formate.List(Me.LstQuelle_Formate.ListIndex, idxFmtDateifilter)
          Quelle_LetztesFormat_AsciiSpezial = oExpimGlobal.Quelle_FormatID
          Ziel_IoTyp = Me.LstQuelle_Formate.List(Me.LstQuelle_Formate.ListIndex, idxFmtIoTyp)
          Ziel_Kategorien = Me.LstQuelle_Formate.List(Me.LstQuelle_Formate.ListIndex, idxFmtKategorien)
         'MsgBox "Ziel_IoTyp=" & Ziel_IoTyp
         
      Case Else
    
    End Select
    
    
    Select Case Ziel_IoTyp
      
        Case io_Typ_XlTabNeu
          'Me.fraMod.Enabled = False
          Me.chkMod_Anwenden.Visible = False
          Me.chkMod_Ersatzspalten.Visible = False
          'Me.chkMod_FormelnErhalten.Visible = False
          Me.chkMod_VorhWerteUeberschreiben.Visible = False
          Me.optZiel_Typ_XLTabNeu.value = True
          Me.optZiel_Typ_XLTabNeu.Enabled = True
          Me.optZiel_Typ_AsciiFormatiert.Enabled = False
          Me.optZiel_Typ_AsciiSpezial.Enabled = False
          Me.chkZiel_Formate.value = False
          Me.chkZiel_Formate.Enabled = False
          
        Case io_Typ_Puffer
          Me.fraMod.Enabled = True
          Me.chkMod_Anwenden.Visible = True
          Me.chkMod_Ersatzspalten.Visible = True
          'Me.chkMod_FormelnErhalten.Visible = True
          Me.chkMod_VorhWerteUeberschreiben.Visible = True
          Me.optZiel_Typ_XLTabNeu.Enabled = True
          Me.optZiel_Typ_XLTabNeu.value = True
          'yyy Me.optZiel_Typ_AsciiFormatiert.Enabled = True
          'yyy Me.optZiel_Typ_AsciiSpezial.Enabled = True
          Me.chkZiel_Formate.Enabled = True
          
        Case Else
          Me.fraMod.Enabled = False
          Me.chkMod_Anwenden.Visible = False
          Me.chkMod_Ersatzspalten.Visible = False
          'Me.chkMod_FormelnErhalten.Visible = False
          Me.chkMod_VorhWerteUeberschreiben.Visible = False
          Me.optZiel_Typ_XLTabNeu.Enabled = False
          Me.optZiel_Typ_AsciiFormatiert.Enabled = False
          Me.optZiel_Typ_AsciiSpezial.Enabled = False
          'MsgBox "Der Zieldatentyp des gewählten Formates wird nicht unterstützt!"
          
    End Select
  End If
  
  Call Changed_Ziel_Typ
  'Call Check_DialogOK
  
End Sub


Private Sub chkZiel_Formate_Click()
  'Checkbox "Alle Zielformate zeigen" geändert.
  'Filterstatus der Ziel-Liste ändern.
  Call Changed_Ziel_Typ
End Sub


Private Sub LstZiel_Formate_Change()
  'Neues Zielformat übernehmen.
  'Inhalt ist abhängig von:
  '  - oExpimGlobal.Ziel_Typ (io_Typ_XlTabNeu, io_Typ_AsciiFormatiert, io_Typ_AsciiSpezial)
  
  On Error GoTo 0
  
  DebugEcho "LstZiel_Formate_Change(): Neuer Wert = " & Me.LstZiel_Formate.value
        
  'Aktuelle Auswahl als Eigenschaft von oExpim speichern.
  If (IsNull(Me.LstZiel_Formate.value)) Then
    oExpimGlobal.Ziel_FormatID = ""
  Else
    oExpimGlobal.Ziel_FormatID = Me.LstZiel_Formate.value
  End If
  
  
  Select Case oExpimGlobal.Ziel_Typ
    
    Case io_Typ_XlTabNeu
        
        'Neue Excel-Vorlage gewählt. Davon hängt nichts weiter ab.
        Ziel_LetztesFormat_XlTabNeu = oExpimGlobal.Ziel_FormatID
        
    Case io_Typ_AsciiFormatiert
        
    Case io_Typ_AsciiSpezial
        
    Case Else
      
  End Select
  
  Call Check_DialogOK
  
End Sub


Private Function GetXLVorlagen(DateiMaske As String) As String
  'Erzeugt eine Liste aller verfügbarer XL-Vorlagen, die der Dateimaske
  'entsprechen und in den üblichen Verzeichnissen zu finden sind.
  '   "DateiMaske"   = DateiMaske ohne Pfadangabe (mit Wildcards, z.B. "*.xltx,*xltm")
  '   Rückgabe       = Dateiliste, durch Semikolons getrennt.
  
  Const PersonalTemplates_RegValue = "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Options\PersonalTemplates"
  
  Dim PersonalTemplates As String
  Dim VerzListe         As String
  Dim VorlagenListe     As String
  Dim oVorlagenListe    As Scripting.Dictionary
  
  VerzListe = ""
  VorlagenListe = ""
  PersonalTemplates = ""
  
  ' Vorlagen-Ordner in Office 365 am 10.05.2020 :
  '----------------------------------------------
  '
  ' Benutzerdefinierte Office-Vorlagen
  '   - Standard:              "Benutzerdefinierte Office-Vorlagen" als (hart kodiertes) Unterverzeichnis
  '                            von C:\Users\<USERNAME>\Documents\
  '   - Einstellung in Word:   Optionen -> Erweitert -> Allgemein -> Dateispeicherorte -> Dokumente
  '   - Einstellung in Excel:  nicht möglich => wird von Word übernommen
  '   - VBA Objektmodell:      ?
  '   - Registry:              HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Personal
  '   - Anmerkungen:
  '     - Dort liegende Vorlagen sind für den Anwender NICHT SICHTBAR !
  '     - Beim Speichern als Vorlage startet dort der Dateidialog in Word und Excel,
  '       falls kein "Standardspeicherort für persönliche Vorlagen" festgelegt ist.
  '
  '
  ' Benutzervorlagen
  '   - Standard:              C:\Users\<USERNAME>\AppData\Roaming\Microsoft\Templates\
  '   - Einstellung in Word:   Optionen -> Erweitert -> Allgemein -> Dateispeicherorte -> Benutzervorlagen
  '   - Einstellung in Excel:  nicht möglich => wird von Word übernommen
  '   - VBA Objektmodell:      Application.TemplatesPath
  '   - Registry:              HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\General\UserTemplates
  '   - Anmerkungen:
  '     - Dort liegende Vorlagen sind für den Anwender NICHT SICHTBAR !
  '     - Dort schreibt Word die Normal.dotm.
  '     - Der Registry-Value existiert nur, wenn der Inhalt nicht dem Standard entspricht.
  '
  '
  ' Arbeitsgruppenvorlagen
  '   - Standard:              <leer>
  '   - Einstellung in Word:   Optionen -> Erweitert -> Allgemein -> Dateispeicherorte -> Arbeitsgruppenvorlagen
  '   - Einstellung in Excel:  nicht möglich => wird von Word übernommen
  '   - VBA Objektmodell:      Application.NetworkTemplatesPath
  '   - Registry:              HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\General\SharedTemplates
  '                            => benutzerabhängig !
  '   - Anmerkungen:
  '     - Dort liegende Vorlagen sind für den Anwender SICHTBAR ***
  '
  '
  ' Standardspeicherort für persönliche Vorlagen
  '   - Standard:              <leer>
  '   - Einstellung in Word:   Optionen -> Speichern -> Standardspeicherort für persönliche Vorlagen
  '   - Einstellung in Excel:  Optionen -> Speichern -> Standardspeicherort für persönliche Vorlagen
  '   - VBA Objektmodell:      ?
  '   - Registry:              HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Word\Options\PersonalTemplates
  '   - Anmerkungen: 
  '     - Dort liegende Vorlagen sind für den Anwender SICHTBAR ***
  '     - Word und Excel verwalten diese Einstellung getrennt voneinander !
  '     - Beim Speichern als Vorlage startet dort der Dateidialog in Word und Excel.
  ' 
  
  If (ThisWorkbook.SysTools.RegValueExists(PersonalTemplates_RegValue)) Then
    PersonalTemplates = ThisWorkbook.SysTools.RegRead(PersonalTemplates_RegValue)
  End If
  
  If (Application.NetworkTemplatesPath <> "") Then VerzListe = VerzListe & ";" & Application.NetworkTemplatesPath
  If (Application.TemplatesPath <> "")        Then VerzListe = VerzListe & ";" & Application.TemplatesPath
  If (PersonalTemplates <> "")                Then VerzListe = VerzListe & ";" & PersonalTemplates
  
  VerzListe = Mid(VerzListe, 2)
  Set oVorlagenListe = ThisWorkbook.SysTools.FindFiles(DateiMaske, VerzListe, True)
  Call SortDictionary(oVorlagenListe, 1, 1, False)
  VorlagenListe = VorlagenListe & ";" & Join(oVorlagenListe.Keys, ";")
  DebugEcho "GetXLVorlagen='" & VorlagenListe & "'"
  
  ''XL-Startverzeichnisse:
  'VerzListe = ""
  'If (Application.AltStartupPath <> "") Then VerzListe = VerzListe & ";" & Application.AltStartupPath
  'If (Application.StartupPath <> "") Then VerzListe = VerzListe & ";" & Application.StartupPath
  'VerzListe = Mid(VerzListe, 2)
  'Set oVorlagenListe = ThisWorkbook.SysTools.FindFiles(DateiMaske, VerzListe, False)
  'VorlagenListe = VorlagenListe & ";" & join(oVorlagenListe.keys, ";")
  'DebugEcho "GetXLVorlagen='" & VorlagenListe & "'"
  
  VorlagenListe = Mid$(VorlagenListe, 2)
  If (Right$(VorlagenListe, 1) = ";") Then VorlagenListe = Left$(VorlagenListe, Len(VorlagenListe) - 1)
  
  DebugEcho vbNewLine & "GetXLVorlagen='" & VorlagenListe & "'"
  GetXLVorlagen = VorlagenListe
End Function


Private Function isGueltig_QuellTyp() As Boolean
  'True, wenn der Typ für die Quelldaten gültig ist.
  
  Select Case oExpimGlobal.Quelle_Typ
    
    Case io_Typ_XlTabAktiv, io_Typ_CsvSpezial, io_Typ_AsciiFormatiert, io_Typ_AsciiSpezial
        isGueltig_QuellTyp = True
        
    Case Else
        isGueltig_QuellTyp = False
        
  End Select

End Function


Private Function isGueltig_QuellFormat() As Boolean
  'True, wenn das Format für die Quelldaten gültig ist.
  Dim blnGueltig As Boolean
  
  blnGueltig = False
  
  Select Case oExpimGlobal.Quelle_Typ
    
    Case io_Typ_XlTabAktiv
        'Keine Formatauswahl nötig.
        blnGueltig = True
    
    Case io_Typ_CsvSpezial
        'Keine Formatauswahl nötig.
        blnGueltig = True
    
    Case io_Typ_AsciiFormatiert, io_Typ_AsciiSpezial
        'Listeneintrag ist gewählt
        If (Not IsNull(Me.LstQuelle_Formate.value)) Then blnGueltig = True
    
    Case Else
    
  End Select
  isGueltig_QuellFormat = blnGueltig
  
End Function


Private Function isGueltig_Quelldatei() As Boolean
  'True, wenn die gewählte Quelldatei existiert.
  Dim blnGueltig As Boolean
  
  Select Case oExpimGlobal.Quelle_Typ
    
    Case io_Typ_XlTabAktiv
        'Keine ASCII-Datei nötig.
        blnGueltig = True
        
    Case io_Typ_CsvSpezial, io_Typ_AsciiFormatiert, io_Typ_AsciiSpezial
        
        If (ThisWorkbook.SysTools.IsDatei(oExpimGlobal.Quelle_AsciiDatei_Name)) Then
          blnGueltig = True
        Else
          blnGueltig = False
        End If
        
    Case Else
        blnGueltig = False
        
  End Select
  isGueltig_Quelldatei = blnGueltig

End Function


Private Function isGueltig_ZielTyp() As Boolean
  'True, wenn der Typ für die Zieldaten gültig ist.
  
  Select Case oExpimGlobal.Ziel_Typ
    
    Case io_Typ_XlTabNeu, io_Typ_AsciiFormatiert, io_Typ_AsciiSpezial
        isGueltig_ZielTyp = True
        
    Case Else
        isGueltig_ZielTyp = False
        
  End Select

End Function


Private Function isGueltig_ZielFormat() As Boolean
  'True, wenn das Format für die Zieldaten gültig ist.
  Dim blnGueltig As Boolean
  
  blnGueltig = False
  
  Select Case oExpimGlobal.Ziel_Typ
    
    Case io_Typ_XlTabNeu, io_Typ_AsciiFormatiert
        If (ThisWorkbook.SysTools.IsDatei(oExpimGlobal.Ziel_FormatID)) Then blnGueltig = True
        
    Case io_Typ_AsciiSpezial
        If (Not IsNull(Me.LstZiel_Formate.value)) Then blnGueltig = True
        
    Case Else
        blnGueltig = False
        
  End Select
  isGueltig_ZielFormat = blnGueltig

End Function


Private Function isGueltig_Zieldatei() As Boolean
  'True, wenn die gewählte Zieldatei existiert bzw. der Name für eine neue Datei gültig ist.
  Dim blnGueltig As Boolean
  
  Select Case oExpimGlobal.Ziel_Typ
    
    Case io_Typ_XlTabNeu
        'Keine ASCII-Datei nötig.
        blnGueltig = True
    
    Case io_Typ_AsciiFormatiert, io_Typ_AsciiSpezial
        
        'If (ThisWorkbook.SysTools.IsDatei(oExpimGlobal.Ziel_AsciiDatei_Name)) Then
        If (oExpimGlobal.Ziel_AsciiDatei_Name <> "") Then
          'Echte Prüfung hat bereits "tbZiel_AsciiDatei_Change" übernommen.
          blnGueltig = True
        Else
          blnGueltig = False
        End If
        
    Case Else
        blnGueltig = False
        
  End Select
  isGueltig_Zieldatei = blnGueltig

End Function


Private Function GetQuellDateinameAusDialog() As String
  'Liefert den gültigen Pfad\Namen einer existierenden Eingabedatei für Me.Import oder aber "".
  'Es wird ein Dateidialog geöffnet. Falls oExpimGlobal.Quelle_AsciiDatei_Name eine Datei in einem
  'existierenden Verzeichnis bezeichnet, so wird dieses zum Startverzeichnis des Dialoges.
  
  On Error Resume Next
  Dim DateiPfadName  As String
  Dim Verzeichnis    As String
  
  'Startverzeichnis des Dialoges festlegen, wenn möglich.
  Call SetArbeitsverzeichnis(Verz(oExpimGlobal.Quelle_AsciiDatei_Name))
  
  'Dateidialog starten.
  Err.Clear
  DateiPfadName = Application.GetOpenFileName(oExpimGlobal.Quelle_AsciiDatei_DialogFilter, , "Quelldatei wählen:")
  If (Err.Number <> 0) Then DateiPfadName = Application.GetOpenFileName("", , "Quelldatei wählen:")
  
  If (ThisWorkbook.SysTools.IsDatei(DateiPfadName)) Then
    'Arbeitsverzeichnis der Eingabedatei setzen (für künftige Öffnen/Speichern-Dialoge)
    Call SetArbeitsverzeichnis(Verz(DateiPfadName))
  Else
    DateiPfadName = ""
  End If
  GetQuellDateinameAusDialog = DateiPfadName
  
  On Error GoTo 0
End Function


Private Function GetZielDateinameAusDialog() As String
  'Liefert den gültigen Pfad\Namen einer Ausgabedatei für Me.Import oder aber "".
  'Der Pfad existiert, die Datei muß nicht existieren.
  'Es wird ein Dateidialog geöffnet. Falls oExpimGlobal.Ziel_AsciiDatei_Name eine Datei in einem
  'existierenden Verzeichnis bezeichnet, so wird dieses zum Startverzeichnis des Dialoges.
  
  On Error Resume Next
  Dim DateiPfadName  As String
  Dim Verzeichnis    As String
  
  'Startverzeichnis des Dialoges festlegen, wenn möglich.
  Verzeichnis = Verz(oExpimGlobal.Ziel_AsciiDatei_Name)
  If (Dir(Verzeichnis & "\") <> "") Then
    ChDrive Verzeichnis
    ChDir Verzeichnis
  End If
  
  'Dateidialog starten.
  Err.Clear
  DateiPfadName = Application.GetSaveAsFilename(NameExt(oExpimGlobal.Ziel_AsciiDatei_Name, "mitext"), oExpimGlobal.Ziel_AsciiDatei_DialogFilter, , "Zieldatei wählen:")
  If (Err.Number <> 0) Then DateiPfadName = Application.GetOpenFileName(, "", , "Zieldatei wählen:")
  
  If (ThisWorkbook.SysTools.IsDatei(DateiPfadName)) Then
    'Arbeitsverzeichnis der Ausgabedatei setzen (für künftige Öffnen/Speichern-Dialoge)
    Verzeichnis = Verz(DateiPfadName)
    ChDrive Verzeichnis
    ChDir Verzeichnis
  ElseIf ((DateiPfadName = "False") Or (DateiPfadName = "Falsch")) Then
    DateiPfadName = ""
  End If
  GetZielDateinameAusDialog = DateiPfadName
  
  On Error GoTo 0
End Function


Private Sub SetzeAnfangswerte()
  'Initialisierung der Oberflächen-Steuerelemente.
  'Bereits gesetzte Eigenschaften von oExpim (CdatExpim) übernehmen bzw. Standardwerte setzen.
  
  'Sperren für Checkboxen deaktivieren
  Sperre_chkMod_Anwenden = False
  Sperre_chkMod_VorhWerteUeberschreiben = False
  Sperre_chkMod_Ersatzspalten = False
  
  'Quelle Typ
  Select Case oExpimGlobal.Quelle_Typ
    Case io_Typ_XlTabAktiv
        Me.optQuelle_Typ_XLTabAktiv.value = True
        'Kein Format sinnvoll.
        
    Case io_Typ_CsvSpezial
        Me.optQuelle_Typ_CsvSpezial.value = True
        
    Case io_Typ_AsciiFormatiert
        Quelle_LetztesFormat_AsciiFormatiert = oExpimGlobal.Quelle_FormatID
        Me.optQuelle_Typ_AsciiFormatiert.value = True
        
    Case io_Typ_AsciiSpezial
        Quelle_LetztesFormat_AsciiSpezial = oExpimGlobal.Quelle_FormatID
        Me.optQuelle_Typ_AsciiSpezial.value = True
        
    Case Else
        'Standard = Aktive XL-Tabelle.
        Quelle_LetztesFormat_AsciiSpezial = ""
        Me.optQuelle_Typ_XLTabAktiv.value = True
  End Select
  
  'Wenn Aktive Tabelle nicht als Quelle dienen kann => Quelle = ASCII_Spezial!
  If (ThisWorkbook.AktiveTabelle.Infotraeger Is Nothing) Then
    Me.optQuelle_Typ_XLTabAktiv.Enabled = False
    'yyy If (Me.optQuelle_Typ_XLTabAktiv) Then Me.optQuelle_Typ_AsciiFormatiert = True
    If (Me.optQuelle_Typ_XLTabAktiv) Then Me.optQuelle_Typ_AsciiSpezial = True
  End If
  
  'Datenbehandlungs-Optionen
  Call ReflektiereDatenEinstellungen
  
  'Ziel Typ
  Select Case oExpimGlobal.Ziel_Typ
    Case io_Typ_XlTabNeu
        Me.optZiel_Typ_XLTabNeu = True
        Ziel_LetztesFormat_XlTabNeu = oExpimGlobal.Ziel_FormatID
    Case io_Typ_AsciiFormatiert
        Me.optZiel_Typ_AsciiFormatiert = True
        'Ziel_LetztesFormat_AsciiFormatiert = oExpim
    Case io_Typ_AsciiSpezial
        Me.optZiel_Typ_AsciiSpezial = True
        'Ziel_LetztesFormat_AsciiSpezial = oExpim
    Case Else
        'Standard = Neue XL-Tabelle.
        Me.optZiel_Typ_XLTabNeu = True
        Ziel_LetztesFormat_XlTabNeu = ""
  End Select
  
  'Ziel-Formatliste.
  Me.chkZiel_Formate.value = False
  
  'Ziel Dateimodus
  Select Case oExpimGlobal.Ziel_AsciiDatei_Modus
    Case io_Datei_Modus_Neu
        'Me.optZiel_AsciiDatei_Neu = True
    Case io_Datei_Modus_Ueberschreiben
        'Me.optZiel_AsciiDatei_Ueberschreiben = True
    Case io_Datei_Modus_Anhaengen
        'Me.optZiel_AsciiDatei_Anhaengen = True
  End Select
  
  'ASCII-Dateinamen.
  Me.tbQuelle_AsciiDatei = oExpimGlobal.Quelle_AsciiDatei_Name
  'Me.tbZiel_AsciiDatei = oExpimGlobal.Ziel_AsciiDatei_Name
  'Call AsciiDatei_PfadKorrigieren(Me.tbZiel_AsciiDatei)
  Call AsciiDatei_PfadKorrigieren(Me.tbQuelle_AsciiDatei)
  
End Sub


Private Sub SetzeStandardDatenEinstellungen()
  'Einstellungen für Datenübertragung und -Bearbeitung zurücksetzen
  'und an der Oberfläche refleklieren
  
  Call oExpimGlobal.SetzeStandardDatenEinstellungen
  
  'Sperren für Checkboxen deaktivieren
  Sperre_chkMod_Anwenden = False
  Sperre_chkMod_VorhWerteUeberschreiben = False
  Sperre_chkMod_Ersatzspalten = False
  
  Call ReflektiereDatenEinstellungen
End Sub


Private Sub ReflektiereDatenEinstellungen()
  'Daten-Einstellungen an der Oberfläche widerspiegeln
  Dim ErsatzKonfiguriert As Boolean
  
  'Checkboxen für Datenbehandlung zunächst aktivieren
  Me.chkMod_Anwenden.Enabled = True
  Me.chkMod_VorhWerteUeberschreiben.Enabled = True
  Me.chkMod_Ersatzspalten.Enabled = True
  
  'Checkbox "Daten modifizieren"
  Me.chkMod_Anwenden.value = oExpimGlobal.Opt_DatenModifizieren
  If (Sperre_chkMod_Anwenden) Then Me.chkMod_Anwenden.Enabled = False
  
  'Checkbox "Vorhandene Werte überschreiben"
  Me.chkMod_VorhWerteUeberschreiben.value = oExpimGlobal.Datenpuffer.Opt_VorhWerteUeberschreiben
  If (Sperre_chkMod_VorhWerteUeberschreiben Or Not Me.chkMod_Anwenden.value) Then
    Me.chkMod_VorhWerteUeberschreiben.Enabled = False
  End If
  
  
  'Option "Ersatzspalten" ist nur verfügbar, wenn solche konfiguriert sind.
  ErsatzKonfiguriert = False
  If (Not ThisWorkbook.Konfig Is Nothing) Then
    If (Not ThisWorkbook.Konfig.SpaltenErsatzZiel Is Nothing) Then
      If (ThisWorkbook.Konfig.SpaltenErsatzZiel.Count > 0) Then
        ErsatzKonfiguriert = True
        Me.chkMod_Ersatzspalten.value = oExpimGlobal.Opt_ErsatzZielspaltenVerwenden
        If (Sperre_chkMod_Ersatzspalten) Then
          Me.chkMod_Ersatzspalten.Enabled = False
        Else
          Me.chkMod_Ersatzspalten.Enabled = True
        End If
      End If
    End If
  End If
  
  'Falls Konfiguration von Ersatz-Zielspalten nicht verfügbar => Option deaktivieren
  If (Not ErsatzKonfiguriert) Then
    Me.chkMod_Ersatzspalten.value = False
    Me.chkMod_Ersatzspalten.Enabled = False
  End If
  
  
  'Checkbox "Formeln Erhalten" wird behandelt in Changed_Ziel_Typ()
    'Me.chkMod_FormelnErhalten.Visible
    'Me.chkMod_FormelnErhalten.Enabled
    
  'Datenpuffer-Optionen, die bisher nicht auf der Oberfläche schaltbar sind.
    'oExpimGlobal.Datenpuffer.Opt_FehlerVerbesserungen
    'oExpimGlobal.Datenpuffer.Opt_UeberhoehungAusBemerkung
    'oExpimGlobal.Datenpuffer.Opt_Transfo_Tk2Gls
End Sub

Private Function GetXltCachePath() As String
    ' Gibt den Pfad für den XLT-Cache mit einem existierenden Verzeichnis zurück.
    ' - Wenn möglich: "%LOCALAPPDATA%\GeoTools\GeoTools_xltcache.txt"
    ' - sonst:        "%Temp%\GeoTools_xltcache.txt" => wird vom System oft gelöscht.
    Dim oFS                 As New Scripting.FileSystemObject
    Dim CacheDir            As String
    Dim LocalAppData        As String
    Dim PfadNameXltCache    As String
    
    Const NameXltCache As String = "GeoTools_xltcache.txt"

    PfadNameXltCache = "?"
    
    ' Versuche, "%LOCALAPPDATA%\GeoTools" zu finden bzw. anzulegen.
    LocalAppData = Environ("LOCALAPPDATA")
    On Error Resume Next
    If (oFS.FolderExists(LocalAppData)) Then
        CacheDir = oFS.GetFolder(LocalAppData).Path & "\GeoTools"
        If (Not oFS.FolderExists(CacheDir)) Then
            oFS.CreateFolder CacheDir
        End If
        If (oFS.FolderExists(CacheDir)) Then
            PfadNameXltCache = CacheDir & "\" & NameXltCache
        End If
    End If
    On Error GoTo 0
    
    ' Fallback: Datei im Temp-Verzeichnis.
    If (PfadNameXltCache = "?") Then
        PfadNameXltCache = oFS.GetSpecialFolder(TempOrdner).Path & "\" & NameXltCache
    End If
    
    GetXltCachePath = PfadNameXltCache

End Function


Private Sub GetFormatliste_XlVorlagen()
  'Erzeugt eine Liste aller verfügbarer XL-Vorlagen für den Import/Export-Dialog.
  'Jede Vorlage wird, wenn nötig, geöffnet und analysiert (Titel, TabName, Kategorien der Spalten).
  'Ergebnis ... Array Liste_XLT_komplett(1 Zeile = DateiVorname, Titel, Pfad\Name, Kategorienliste).
  
  On Error GoTo Fehler
  
  'Deklarationen
    Dim RecentXLT_ok              As Boolean
    Dim DatumIdentisch            As Boolean
    Dim XltInfo_OK                As Boolean
    Dim FormatPfadName()          As String
    Dim tmp()                     As String
    Dim XltPfad                   As Variant
    Dim RecentXLT                 As Variant
    Dim RecentXLT_Neu(0 To 2)     As String
    Dim RecentXLT_Cache(0 To 2)   As String
    Dim PfadNameXltCache          As String
    Dim Titel                     As String
    Dim Kategorien                As String
    Dim Formatliste               As String
    Dim AktFormatPfadName         As String
    Dim DateiAendDatum            As String
    Dim vDateiAendDatum           As Variant
    Dim Zeile                     As String
    Dim Computername              As String
    Dim Username                  As String
    Dim NF                        As Long
    Dim i                         As Long
    Dim iAnz                      As Long
    Dim TemplatesCount            As Long
    Dim lb                        As Long
    Dim ub                        As Long
    Dim CurrentTemplate           As Workbook
    Dim oFS                       As New Scripting.FileSystemObject
    Dim oTS_XltCache              As Scripting.TextStream
    Dim oXlApp2                   As Excel.Application
  '
  'Const NameXltCache As String = "GeoTools_xltcache.txt"
  DebugEcho "GetFormatliste_XlVorlagen(): Liste der verfügbaren XL-Vorlagen zusammenstellen..."
  
  'Vorbereitungen
    Const idxTitel              As Long = 0
    Const idxKategorien         As Long = 1
    Const idxAendDatum          As Long = 2
    
    Const idxNetTemplatesPath   As Long = 0
    Const idxTemplatesPath      As Long = 1
    Const idxAltStartupPath     As Long = 2
    Const idxStartupPath        As Long = 3
    
    On Error Resume Next
    Computername = ThisWorkbook.SysTools.Computername
    Username = ThisWorkbook.SysTools.Username
    On Error GoTo Fehler
  '
  '1. Aktuell auf Festplatte vorhandene Vorlagen suchen und die entsprechenden Verzeichnisse merken.
    '   => Diese Liste wird eine Spalte des Arrays Liste_XLT_komplett
    '      und gibt damit auch die Sortierung vor (Zuerst XLT vom Netzwerk, dann lokale XLT).
    Formatliste = GetXLVorlagen(DateiMaskeXLT)
    NF = SplitDelim(Formatliste, FormatPfadName, ";")
    DebugEcho "Anzahl verfügbarer XL-Vorlagen = " & CStr(NF)
    Call WriteStatusBar("Anzahl verfügbarer XL-Vorlagen = " & CStr(NF))
  '
  '2. Persistenten XLT-Cache lesen
  '3. Array "Liste_XLT_komplett" für Listbox erstellen.
  '4. "Innere" Eigenschaften aller gefundenen Vorlagen zwischenspeichern:
  If (NF > 0) Then
    '2. Eigenschaften der zuletzt vorhandenen Vorlagen lesen (aus persistentem Cache)
    DebugEcho "GetFormatliste_XlVorlagen(): Cache der zuletzt vorhandenen XL-Vorlagen lesen..."
    
    'PfadNameXltCache = oFS.GetSpecialFolder(TempOrdner).Path & "\" & NameXltCache
    PfadNameXltCache = GetXltCachePath()
    
    If (Not ThisWorkbook.SysTools.IsDatei(PfadNameXltCache)) Then
      DebugEcho vbNewLine & "Cache existiert nicht: '" & PfadNameXltCache & "'"
    Else
      DebugEcho vbNewLine & "Cache lesen: '" & PfadNameXltCache & "'"
      Set oTS_XltCache = ThisWorkbook.SysTools.OpenTextFile(PfadNameXltCache, ForReading, NewFileIfNotExist_no, OpenAsSystemDefault)
      If (Not oTS_XltCache Is Nothing) Then
        'Datei erfolgreich geöffnet.
        iAnz = 0
        Do While Not oTS_XltCache.AtEndOfStream
          Zeile = oTS_XltCache.ReadLine
          If (Left(Zeile, 1) <> "#") Then
            'keine Kopfzeile
            NF = SplitDelim(Zeile, tmp, "|")
            If (NF <> 4) Then
              ErrEcho "Fehler: in Zeile " & oTS_XltCache.Line & ": falsche Anzahl Felder: " & CStr(NF)
            Else
              iAnz = iAnz + 1
              RecentXLT_Cache(idxAendDatum) = tmp(1)
              RecentXLT_Cache(idxTitel) = tmp(3)
              RecentXLT_Cache(idxKategorien) = tmp(4)
              RecentXLT = RecentXLT_Cache
              oRecentXLT.Add tmp(2), RecentXLT
            End If
          End If
        Loop
        DebugEcho " - Eigenschaften für " & CStr(iAnz) & " Vorlagen gelesen."
        oTS_XltCache.Close
      End If
    End If
    DebugEcho "==> Eigenschaften für insgesamt " & CStr(oRecentXLT.Count) & " Vorlagen gelesen." & vbNewLine
    
    '3. "Liste_XLT_komplett" erstellen. Dabei:
     '   a, Titel und Kategorien aus oRecentXLT übernehmen, falls das Änderungsdatum dort
     '   mit dem der gefundenen Datei übereinstimmt, ansonsten XLT öffnen ...
     '   b, Liste der zuletzt vorhandenen XLT's neu erstellen (oRecentXLT_Neu)
    DebugEcho "Liste der XL-Vorlagen für Dialog-Listbox erstellen:"
    ub = UBound(FormatPfadName)
    lb = LBound(FormatPfadName)
    ReDim Liste_XLT_komplett(0 To ub - 1, 0 To 5) As String
    iAnz = 0
    TemplatesCount = ub - lb + 1
    
    For i = lb To ub
      
      AktFormatPfadName = FormatPfadName(i)
      DebugEcho "bearbeite Datei '" & AktFormatPfadName & "'"
      XltInfo_OK = False
      
      'Datum der (auf der Festplatte) gefundenen Datei
        ErrMessage = "Datei existiert nicht !?!"  'Kann eigentlich nicht sein.
        vDateiAendDatum = oFS.GetFile(AktFormatPfadName).DateLastModified
        DateiAendDatum  = Format(vDateiAendDatum, "dd/mm/yy hh:mm:ss")
        DebugEcho " - Letzte Änderung der gefundenen Datei = " & DateiAendDatum
      
      'Datum  der zuletzt vorhandenen Datei
        If (oRecentXLT.Exists(AktFormatPfadName)) Then
          RecentXLT_ok = True
          RecentXLT = oRecentXLT(AktFormatPfadName)
          DebugEcho " - Letzte Änderung der Datei laut Cache = " & RecentXLT(idxAendDatum)
        Else
          RecentXLT_ok = False
          DebugEcho " - => neue Vorlage (nicht in " & PfadNameXltCache & " vorhanden)!"
        End If
      
      'Wenn Datum der gefundenen XLT nicht mit dem der zuletzt vorhandenen XLT übereinstimmt => hineinsehen..
      If (RecentXLT_ok) Then DatumIdentisch = (DateiAendDatum = RecentXLT(idxAendDatum)) Else DatumIdentisch = False
      If (RecentXLT_ok And DatumIdentisch) Then
        DebugEcho " - Info's der zuletzt vorhandenen Datei werden verwendet"
        Titel = RecentXLT(idxTitel)
        Kategorien = RecentXLT(idxKategorien)
        XltInfo_OK = True
      Else
        DebugEcho " - Info's der auf Festplatte gefundenen Datei werden ermittelt"
        Call ProgressbarAllgemein(TemplatesCount, i - lb + 1, "GeoTools analysiert Vorlage:     " & AktFormatPfadName)
        
        'XLT öffnen mit neuer Excel-Instanz.
        If (oXlApp2 Is Nothing) Then
          Set oXlApp2 = New Excel.Application
          oXlApp2.EnableEvents = False
          oXlApp2.AutomationSecurity = msoAutomationSecurityForceDisable
          'oXlApp2.Visible = False  (ist von vornherein unsichtbar)
          'oXlApp2.ScreenUpdating = False (ohnehin unsichtbar)
        End If
        On Error Resume Next
        ErrMessage = "Fehler beim Erkunden einer Vorlage"
        Set CurrentTemplate = oXlApp2.Workbooks.Open(FileName:=AktFormatPfadName, ReadOnly:=True, UpdateLinks:=0 , AddToMru:=False) 
        
        If (Err.Number <> 0) Then
            FehlerNachricht "frmStartExpim.GetFormatliste_XlVorlagen()"
        Else
          'XLT: Info's lesen und schließen
          On Error GoTo Fehler
          
          Titel      = CurrentTemplate.BuiltinDocumentProperties("title").value
          Kategorien = GetKategorien(oXlApp2.ActiveSheet)
          
          CurrentTemplate.Close SaveChanges:=False
          
          DebugEcho " - Titel = '" & Titel & "'"
          DebugEcho " - Kategorien = '" & Kategorien & "'"
          
          'Neu ermittelte Info's merken
          RecentXLT_Neu(idxTitel)      = Titel
          RecentXLT_Neu(idxKategorien) = Kategorien
          RecentXLT_Neu(idxAendDatum)  = DateiAendDatum
          RecentXLT = RecentXLT_Neu
          
          XltInfo_OK = True
        End If
      End If
      
      'Vorlage in Array für Listbox eintragen
      If (XltInfo_OK) Then
        iAnz = iAnz + 1
        Liste_XLT_komplett(iAnz - 1, idxFmtKurzname) = VorName(AktFormatPfadName)
        Liste_XLT_komplett(iAnz - 1, idxFmtID) = AktFormatPfadName
        Liste_XLT_komplett(iAnz - 1, idxFmtTitel) = Titel
        Liste_XLT_komplett(iAnz - 1, idxFmtKategorien) = Kategorien
        Liste_XLT_komplett(iAnz - 1, idxFmtDateifilter) = ""
        Liste_XLT_komplett(iAnz - 1, idxFmtIoTyp) = ""
        
        'Vorlage merken...
        oRecentXLT_Neu.Add AktFormatPfadName, RecentXLT
      End If
    Next
    
    If (Not (oXlApp2 Is Nothing)) Then
      oXlApp2.Quit
    End If
    
    ReDim Preserve Liste_XLT_komplett(0 To iAnz - 1, 0 To 5) As String
    
    '4. "Innere" Eigenschaften aller gefundenen Vorlagen zwischenspeichern:
    'a, im internen Cache
    ThisWorkbook.Konfig.Cache.Add "RecentXLT", oRecentXLT_Neu
    
    'b, persistent (Cache-Datei).
    DebugEcho "GetFormatliste_XlVorlagen(): Cache der aktuell verfügbaren XL-Vorlagen schreiben..."
    DebugEcho vbNewLine & "Cache schreiben in Datei: '" & PfadNameXltCache & "'"
    Set oTS_XltCache = ThisWorkbook.SysTools.OpenTextFile(PfadNameXltCache, ForWriting, NewFileIfNotExist_yes, OpenAsSystemDefault)
    If (Not oTS_XltCache Is Nothing) Then
      
      'Kopf schreiben
      oTS_XltCache.WriteLine "# GeoTools.xlam:    Cache für ""innere"" Eigenschaften aller gefundenen XL-Vorlagen"
      oTS_XltCache.WriteLine "# Erstellt:         " & CStr(Now) & "   (Benutzer " & Username & " an PC " & Computername & ")"
      oTS_XltCache.WriteLine "# -------------------------------------------------------------------------------------------------"
      oTS_XltCache.WriteLine "# Datum und Zeit der letzten Änderung | Pfad\Name der Vorlage | Titel | Kategorien der Datenfelder"
      oTS_XltCache.WriteLine "# -------------------------------------------------------------------------------------------------"
      
      For Each XltPfad In oRecentXLT_Neu
        RecentXLT = oRecentXLT_Neu(XltPfad)
        oTS_XltCache.WriteLine RecentXLT(idxAendDatum) & "|" & XltPfad & "|" & RecentXLT(idxTitel) & "|" & RecentXLT(idxKategorien)
      Next
      DebugEcho " - " & CStr(oTS_XltCache.Line - 1) & " Zeilen geschrieben."
      oTS_XltCache.Close
    End If
  End If
  Call ProgressbarAllgemein(TemplatesCount, TemplatesCount, "Anzahl verfügbarer XL-Vorlagen = " & CStr(TemplatesCount))
  
  'Nachbereitung
    Set oFS = Nothing
    Set oTS_XltCache = Nothing
    Set oXlApp2 = Nothing
    Call ClearStatusBarDelayed(StatusBarClearDelay)
    DebugEcho "GetFormatliste_XlVorlagen(): Liste der verfügbaren XL-Vorlagen vollständig."
    Exit Sub

  Fehler:
  Set oFS = Nothing
  Set oTS_XltCache = Nothing
  Application.StatusBar = ""
  FehlerNachricht "frmStartExpim.GetFormatliste_XlVorlagen()"
End Sub


Private Function GetKategorien(Optional oTab As Worksheet = Nothing) As String
  'Ermittelt alle unterschiedlichen Spalten-Kategorien der angegebenen Tabelle
  'und den Kodenamen der Tabelle als erste Kategorie.
  'Ist "oTab" nicht oder mit Nothing angegeben, dann ist das aktive Arbeitsblatt gemeint.
  'Rückgabe: Liste durch Semikolons getrennt.
  
  On Error GoTo Fehler
  
  Dim DictTmp           As Scripting.Dictionary
  Dim oZellname         As Scripting.Dictionary
  Dim SpaltenName       As Variant
  Dim Liste             As String
  Dim ListeKeys         As String
  Dim ZellnamePur       As String
  Dim Feld()            As String
  Dim NF                As Long
  
  If (oTab Is Nothing) Then
    Set oTab = ActiveSheet
  End If
  
  Set DictTmp   = New Scripting.Dictionary
  Set oZellname = New Scripting.Dictionary
  
  ' Kodename der Tabelle.
  Liste = substitute("[0-9]+$", "", oTab.CodeName, False, False)
  
  ' Kategorien aus Spaltennamen.
  Set oZellname = GetFelderAusTabelle(PrefixSpaltenname, oTab)
  If (Not (oZellname Is Nothing)) Then
    On Error Resume Next
    For Each SpaltenName In oZellname
      If (Len(SpaltenName) > Len(PrefixSpaltenname)) Then
        ' Prefix ".Spalte" entfernen.
        ZellnamePur = Mid(SpaltenName, Len(PrefixSpaltenname) + 1)
        ' Einheit entfernen.
        NF = SplitDelim(ZellnamePur, Feld, TrennerEinheit)
        If (NF > 1) Then
          ZellnamePur = Feld(1)
        End If
        If (ThisWorkbook.Konfig.SpaltenKategorie(ZellnamePur) <> "") Then DictTmp.Add ThisWorkbook.Konfig.SpaltenKategorie(ZellnamePur), "*"
      End If
    Next
    On Error GoTo 0
    ListeKeys = ListeDerKeys(DictTmp)
    If (ListeKeys <> "") Then Liste = Liste & ";" & ListeKeys
  End If
  Set DictTmp   = Nothing
  Set oZellname = Nothing
  GetKategorien = Liste
  Exit Function
  
  Fehler:
  GetKategorien = ""
  Set DictTmp   = Nothing
  Set oZellname = Nothing
  FehlerNachricht "frmStartExpim.GetKategorien()"
End Function


Private Function GetKlassennamen(Prefix As String) As String
  'Erzeugt eine Liste aller Klassenmodule des Add-In, deren
  'Namen mit "Prefix" beginnen.
  '   "Prefix"   = s.o.
  '   Rückgabe   = Klassennamensliste, durch Semikolons getrennt.
  
  On Error GoTo Fehler
  
  'Const TypKlasse     As Long = 2  'Typ=Klassenmodul
  'Dim Modul           As Variant
  'Dim i               As Long
  Dim KlassenListe    As String
  
  'KlassenListe = ""
  'For i = 1 To ThisWorkbook.VBProject.VBComponents.Count
  '  With ThisWorkbook.VBProject.VBComponents(i)
  '    If ((.Type = TypKlasse) And (Left$(.Name, Len(Prefix)) = Prefix)) Then
  '      'echo .Name & "    " & .CodeModule.CountOfLines
  '      KlassenListe = KlassenListe & ";" & .Name
  '    End If
  '  End With
  'Next
  ''Call AnzeigeMeldungen
  'KlassenListe = Mid(KlassenListe, 2)
  'If (Right$(KlassenListe, 1) = ";") Then KlassenListe = Left$(KlassenListe, Len(KlassenListe) - 1)
  
  '**************************************************************************************************
  'Liste wird hart kodiert, um den eventuell nicht erlaubten Zugriff auf das VBProject zu vermeiden!
  '=> An anderer Stelle muss der Klasenname ohnehin hart kodiert werden!
  KlassenListe = "CimpTrassenkoo"
  '**************************************************************************************************
  
  DebugEcho "GetKlassennamen='" & KlassenListe & "'"
  GetKlassennamen = KlassenListe
  
  Exit Function
  
  Fehler:
  FehlerNachricht "frmStartExpim.GetKlassennamen()"
End Function


Private Sub Changed_CsvDateiname()
  'Belegt Quell- und Ziel-Formatlisten für den CSV-Spezialimport.
  ' 1. CSV-Spaltennamen ermitteln und in Quellformatliste anzeigen.
  ' 2. Kategorien ermitteln und als Zielformat-Filter anwenden.
  
  On Error GoTo Fehler
  
  Dim oCsvSpezial                    As CtabCSV
  Dim AnzSpalten                     As Long
  Dim optDatenModifizieren           As Variant
  Dim optVorhWerteUeberschreiben     As Variant
  Dim optErsatzZielspaltenVerwenden  As Variant
  
  DebugEcho "frmStartExpim.Changed_CsvDateiname(): Start."
  
  'Liste der in der CSV-Datei gefundenen Spaltennamen löschen
  Me.LstQuelle_Formate.Clear
  
  'Einstellungen für Datenübertragung und -Bearbeitung zurücksetzen
  Call SetzeStandardDatenEinstellungen
  
  
  If (ThisWorkbook.SysTools.IsDatei(Me.tbQuelle_AsciiDatei.value)) Then
    'CSV-Datei existiert => 'rein sehen!
    Set oCsvSpezial = New CtabCSV
    
    'Mit Setzen des Dateinamens wird der CSV-Kopf gelesen und ausgewertet!
    oCsvSpezial.Quelle_AsciiDatei_Name = oExpimGlobal.Quelle_AsciiDatei_Name
    
    'Eventuell aufgetretene Fehler sofort im Editor anzeigen oder die Protokoll-Konsole aufblenden.
    If (Not ThisWorkbook.SysTools.FileErrorsShowInJEdit(True)) Then
      Call ShowConsole
    End If
    
    'Anzahl der gefundenen Spaltennamen (ist bei Fehler = 0)
    AnzSpalten = oCsvSpezial.Quelle_SpaltenPositionen.Count
    
    '
    'Wenn keine Spalte verfügbar ist, gibt's nichts zu exportieren.
    If (AnzSpalten > 1) Then
      '1. CSV-Spaltennamen abfragen und in Quellformatliste anzeigen.
      Call GetFormatliste_Spaltennamen_CsvSpezial(oCsvSpezial)
      Call GetQuellFormatliste(Liste_Spaltennamen_CsvSpezial, "#")
      
      '2. Kategorien der CSV-Datei als Filterkriterium für die Liste der Zielformate.
      Ziel_Kategorien = oCsvSpezial.Kategorien
      
      'GUI aktualisieren
      '1. Unabhängig vom restlichen Inhalt der CSV-Datei
        Me.lblQuelle_Formate.Caption = "Die CSV-Datei enthält folgende bezeichneten Spalten:"
        Me.LstQuelle_Formate.ControlTipText = "Diese Werte stehen zum Import zur Verfügung"
        Me.fraMod.Enabled = True
        Me.chkMod_Anwenden.Visible = True
        Me.chkMod_VorhWerteUeberschreiben.Visible = True
        Me.chkMod_Ersatzspalten.Visible = True
        Me.optZiel_Typ_XLTabNeu.Enabled = True
        'yyy Me.optZiel_Typ_AsciiFormatiert.Enabled = True
        'yyy Me.optZiel_Typ_AsciiSpezial.Enabled = True
        Me.chkZiel_Formate.Enabled = True
        
      '2. Einstellungen aus dem Kopf der CSV-Datei übernehmen und an der Oberfläche widerspiegeln.
        'Übernahme von eventuell in der CSV-Datei gesetzten Importoptionen.
        '=> Nicht konvertierbare Werte werden übergangen
        
        'a, Eigenschaften setzen
          'Eigenschaften mit Steuerelement an der Oberfläche
          'Wird eine Eigenschaft durch die CSV gültig festgelegt, so muss das entsprechende
          'Steuerelement an der Oberfläche gegen Benutzereingriff gesperrt werden!
          optDatenModifizieren = Null
          optVorhWerteUeberschreiben = Null
          optErsatzZielspaltenVerwenden = Null
          On Error Resume Next
          optDatenModifizieren = CBool(oCsvSpezial.Opt_DatenModifizieren)
          optVorhWerteUeberschreiben = CBool(oCsvSpezial.Opt_VorhWerteUeberschreiben)
          optErsatzZielspaltenVerwenden = CBool(oCsvSpezial.Opt_ErsatzZielspaltenVerwenden)
          On Error GoTo Fehler
          
          If (Not IsNull(optDatenModifizieren)) Then
            oExpimGlobal.Opt_DatenModifizieren = optDatenModifizieren
            Sperre_chkMod_Anwenden = True
          End If
          If (Not IsNull(optVorhWerteUeberschreiben)) Then
            oExpimGlobal.Datenpuffer.Opt_VorhWerteUeberschreiben = optVorhWerteUeberschreiben
            Sperre_chkMod_VorhWerteUeberschreiben = True
          End If
          If (Not IsNull(optErsatzZielspaltenVerwenden)) Then
            oExpimGlobal.Opt_ErsatzZielspaltenVerwenden = optErsatzZielspaltenVerwenden
            Sperre_chkMod_Ersatzspalten = True
          End If
          
          'Eigenschaften ohne Steuerelement an der Oberfläche
          On Error Resume Next
          oExpimGlobal.Datenpuffer.Opt_FehlerVerbesserungen = CBool(oCsvSpezial.Opt_FehlerVerbesserungen)
          oExpimGlobal.Datenpuffer.Opt_UeberhoehungAusBemerkung = CBool(oCsvSpezial.Opt_UeberhoehungAusBemerkung)
          oExpimGlobal.Datenpuffer.Opt_iTrassenCodeAusBemerkung = CBool(oCsvSpezial.Opt_iTrassenCodeAusBemerkung)
          oExpimGlobal.Datenpuffer.Opt_Transfo_Tk2Gls = CBool(oCsvSpezial.Opt_Transfo_Tk2Gls)
        
        'b, ... an der Oberfläche widerspiegeln
        Call ReflektiereDatenEinstellungen
      
      
    Else
      Me.lblQuelle_Formate.Caption = "Die CSV-Datei enthält keine (bezeichneten) Spalten."
      Me.LstQuelle_Formate.ControlTipText = "Die erste Zeile der Datei bzw. nach dem Spezialkopf enthält keine Daten => kein Import möglich!"
      
      'keine existierende ASCII-Datei gewählt
      'ReDim Liste_Spaltennamen_CsvSpezial(0 To 0, 0 To 5) As String
      'Call GetQuellFormatliste(Liste_Spaltennamen_CsvSpezial, "#")
      Ziel_Kategorien = ""
      
      Me.fraMod.Enabled = False
      Me.chkMod_Anwenden.Visible = False
      Me.chkMod_VorhWerteUeberschreiben.Visible = False
      Me.chkMod_Ersatzspalten.Visible = False
      Me.optZiel_Typ_XLTabNeu.Enabled = False
      Me.optZiel_Typ_AsciiFormatiert.Enabled = False
      Me.optZiel_Typ_AsciiSpezial.Enabled = False
      'Me.optZiel_Typ_AsciiSpezial.value = True
      Me.chkZiel_Formate.value = False
      Me.chkZiel_Formate.Enabled = False
    End If
    
    Set oCsvSpezial = Nothing
    
  Else
    'keine existierende ASCII-Datei gewählt
    'ReDim Liste_Spaltennamen_CsvSpezial(-1, 0 To 5) As String
    'Liste_Spaltennamen_CsvSpezial = null
    'Call GetQuellFormatliste(Liste_Spaltennamen_CsvSpezial, "#")
    Ziel_Kategorien = ""
    
    Me.lblQuelle_Formate.Caption = ""
    Me.LstQuelle_Formate.ControlTipText = ""
    Me.fraMod.Enabled = False
    Me.chkMod_Anwenden.Visible = False
    Me.chkMod_VorhWerteUeberschreiben.Visible = False
    Me.chkMod_Ersatzspalten.Visible = False
    Me.optZiel_Typ_XLTabNeu.Enabled = False
    Me.optZiel_Typ_AsciiFormatiert.Enabled = False
    Me.optZiel_Typ_AsciiSpezial.Enabled = False
    'Me.optZiel_Typ_AsciiSpezial.value = False
    Me.chkZiel_Formate.value = False
    Me.chkZiel_Formate.Enabled = False
  End If
  
  Call Changed_Ziel_Typ
    
  DebugEcho "frmStartExpim.Changed_CsvDateiname(): Ende."
  
  Exit Sub
  Fehler:
  Set oCsvSpezial = Nothing
  FehlerNachricht "frmStartExpim.Changed_CsvDateiname()"
End Sub


Private Sub GetFormatliste_SpezialImport()
  'Erzeugt eine Liste aller ASCII-Spezialimport-Module für den Import/Export-Dialog.
  'Jedes Klassenmodul wird instanziert und analysiert (Titel, TabName, Kategorien der Spalten, DateidialogFilter, io_Typ).
  'Ergebnis ... Array Liste_SpezialImport_komplett (1 Zeile = Klassenname ohne Prefix, Titel, Klassenname, Kategorienliste, DateidialogFilter, io_Typ).
  
  On Error GoTo Fehler
  
  Dim KlassenName()             As String
  Dim Formatliste               As String
  Dim NF                        As Long
  Dim i                         As Long
  Dim oAsciiSpezial             As Object
  
  Formatliste = GetKlassennamen(io_Klasse_PrefixImport)
  NF = SplitDelim(Formatliste, KlassenName, ";")
  
  If (NF > 0) Then
    ReDim Liste_SpezialImport_komplett(0 To UBound(KlassenName) - 1, 0 To 5) As String
    For i = LBound(KlassenName) To NF
      Liste_SpezialImport_komplett(i - 1, idxFmtKurzname) = Mid$(KlassenName(i), Len(io_Klasse_PrefixImport) + 1)
      Liste_SpezialImport_komplett(i - 1, idxFmtID) = KlassenName(i)
      
      Select Case KlassenName(i)
        Case io_Klasse_Trassenkoo
            Set oAsciiSpezial = New CimpTrassenkoo
        Case Else
            Err.Raise 66666 + vbObjectError, , "Spezial-Expim-Klasse '" & KlassenName(i) & "' konnte nicht instanziert werden."
      End Select
      
      Liste_SpezialImport_komplett(i - 1, idxFmtTitel) = oAsciiSpezial.Titel
      Liste_SpezialImport_komplett(i - 1, idxFmtKategorien) = oAsciiSpezial.Kategorien
      Liste_SpezialImport_komplett(i - 1, idxFmtDateifilter) = oAsciiSpezial.Quelle_AsciiDatei_DialogFilter
      Liste_SpezialImport_komplett(i - 1, idxFmtIoTyp) = oAsciiSpezial.Ziel_Typ
      Set oAsciiSpezial = Nothing
    Next
  End If
  
  Exit Sub
  Fehler:
  FehlerNachricht "frmStartExpim.GetFormatliste_SpezialImport()"
End Sub


Private Sub GetFormatliste_Spaltennamen_XlTabAktiv()
  'Erzeugt eine Liste aller Spaltennamen der aktiven Tabelle zwecks Anzeige im Import/Export-Dialog.
  'Ergebnis ... Array Liste_Spaltennamen_XlTabAktiv (1 Zeile = Spaltenname ohne Prefix, Beschreibung, Größe, Kategorie, alsFilter, SpaltenFormat).
  
  On Error GoTo Fehler
  
  Dim Formatliste          As String
  Dim Spalte               As Variant
  Dim i                    As Long
  Dim TitelPrefix          As String
  Dim oSpNameAttr          As Scripting.Dictionary
  
  If (Not ThisWorkbook.AktiveTabelle.SpaltenErsteZellen Is Nothing) Then
    ThisWorkbook.AktiveTabelle.Syncronisieren
    If (ThisWorkbook.AktiveTabelle.SpaltenErsteZellen.Count > 0) Then
      i = 0
      ReDim Liste_Spaltennamen_XlTabAktiv(0 To ThisWorkbook.AktiveTabelle.SpaltenErsteZellen.Count - 1, 0 To 5) As String
      For Each Spalte In ThisWorkbook.AktiveTabelle.SpaltenErsteZellen
        Set oSpNameAttr = ThisWorkbook.Konfig.SpNameAttr(Spalte)
        Liste_Spaltennamen_XlTabAktiv(i, idxFmtKurzname) = Spalte
        If ((oSpNameAttr("StatusBez") = "") Or (oSpNameAttr("StatusBez") = Allg_unbekannt)) Then
          TitelPrefix = ""
        Else
          TitelPrefix = oSpNameAttr("StatusBez") & ": "
        End If
        Liste_Spaltennamen_XlTabAktiv(i, idxFmtTitel) = TitelPrefix & oSpNameAttr("Titel")
        Liste_Spaltennamen_XlTabAktiv(i, idxFmtID) = ThisWorkbook.Konfig.SpaltenMathGroesse(Spalte)
        Liste_Spaltennamen_XlTabAktiv(i, idxFmtKategorien) = ThisWorkbook.Konfig.SpaltenKategorie(Spalte)
        Liste_Spaltennamen_XlTabAktiv(i, idxFmtDateifilter) = ThisWorkbook.Konfig.KategorieAlsFilter(ThisWorkbook.Konfig.SpaltenKategorie(Spalte))
        Liste_Spaltennamen_XlTabAktiv(i, idxFmtIoTyp) = ThisWorkbook.AktiveTabelle.SpaltenFormate(Spalte)
        i = i + 1
      Next
    End If
  End If
  Set oSpNameAttr = Nothing
  
  Exit Sub

  Fehler:
  Set oSpNameAttr = Nothing
  FehlerNachricht "frmStartExpim.GetFormatliste_Spaltennamen_XlTabAktiv()"
End Sub


Private Sub GetFormatliste_Spaltennamen_CsvSpezial(oCSV As CtabCSV)
  'Erzeugt eine Liste aller Spaltennamen der CSV-Datei zwecks Anzeige im Import/Export-Dialog.
  'Ergebnis: Array Liste_Spaltennamen_CsvSpezial (1 Zeile = Spaltenname ohne Prefix, Beschreibung, Größe, Kategorie, alsFilter, SpaltenFormat).
  'Eingabe:  oCSV ... instanziertes CtabCSV-Objekt mit bereits gesetztem Dateinamen.
  
  On Error GoTo Fehler
  
  Dim Formatliste          As String
  Dim Spalte               As Variant
  Dim i                    As Long
  Dim TitelPrefix          As String
  Dim oSpNameAttr          As Scripting.Dictionary
  
  If (Not oCSV.Quelle_SpaltenPositionen Is Nothing) Then
    'erfolgt automatisch bei Setzen des Dateinamens:  oCSV.Syncronisieren
    If (oCSV.Quelle_SpaltenPositionen.Count > 0) Then
      i = 0
      ReDim Liste_Spaltennamen_CsvSpezial(0 To oCSV.Quelle_SpaltenPositionen.Count - 1, 0 To 5) As String
      For Each Spalte In oCSV.Quelle_SpaltenPositionen
        Set oSpNameAttr = ThisWorkbook.Konfig.SpNameAttr(Spalte)
        Liste_Spaltennamen_CsvSpezial(i, idxFmtKurzname) = Spalte
        If ((oSpNameAttr("StatusBez") = "") Or (oSpNameAttr("StatusBez") = Allg_unbekannt)) Then
          TitelPrefix = ""
        Else
          TitelPrefix = oSpNameAttr("StatusBez") & ": "
        End If
        Liste_Spaltennamen_CsvSpezial(i, idxFmtTitel) = TitelPrefix & oSpNameAttr("Titel")
        Liste_Spaltennamen_CsvSpezial(i, idxFmtID) = ThisWorkbook.Konfig.SpaltenMathGroesse(Spalte)
        Liste_Spaltennamen_CsvSpezial(i, idxFmtKategorien) = ThisWorkbook.Konfig.SpaltenKategorie(Spalte)
        Liste_Spaltennamen_CsvSpezial(i, idxFmtDateifilter) = ThisWorkbook.Konfig.KategorieAlsFilter(ThisWorkbook.Konfig.SpaltenKategorie(Spalte))
        Liste_Spaltennamen_CsvSpezial(i, idxFmtIoTyp) = ""
        'Liste_Spaltennamen_CsvSpezial(i, idxFmtIoTyp) = oCSV.Quelle_Formate(Spalte)
        i = i + 1
      Next
    End If
  End If
  Set oSpNameAttr = Nothing
  
  Exit Sub

  Fehler:
  Set oSpNameAttr = Nothing
  FehlerNachricht "frmStartExpim.GetFormatliste_Spaltennamen_CsvSpezial()"
End Sub


Private Function FilternFormatliste(FormatlisteKomplett As Variant, FormatlisteGefiltert() As String, ByVal KategorienListe As String) As Long
  'Filtert eine Liste aller verfügbarer Format-Vorlagen nach Kategorien für den Import/Export-Dialog.
  'Parameter:  FormatlisteKomplett  ... vollständige Liste (Array).
  '            FormatlisteGefiltert ... nach Kategorien gefilterte Liste (Array).
  '            KategorienListe      ... Durch Semikolon getrennte Auflistung von Kategorien.
  'Rückgabe:   Anzahl der Einträge der gefilterten Liste.
  'Aufbau der beiden Listenfelder:      1 Zeile = 6 Spalten (DateiVorname (Format-Kurzname), Titel, Pfad\Name  (Format-ID), Kategorienliste, DateidialogFilter, io_Typ).
  
  On Error GoTo Fehler
  
  'Deklarationen
    Dim Kategorien()              As String
    Dim Kategorie                 As String
    Dim Indizes()                 As Long
    Dim AnzKat                    As Long
    Dim AnzFormate                As Long
    Dim i                         As Long
    Dim j                         As Long
    Dim k                         As Long
    Dim lb1                       As Long
    Dim lb2                       As Long
    Dim ub1                       As Long
    Dim ub2                       As Long
    Dim FormatUebernommen         As Boolean
    
  'Nullwerte
    AnzFormate = 0
    Erase FormatlisteGefiltert
  '
  'Liste filtern
  If (CountDim(FormatlisteKomplett) > 0) Then
    'FormatlisteKomplett ist ein Array, d.h. es gibt mind. 1 Format.
    'Vorbereitungen
      lb1 = LBound(FormatlisteKomplett, 1)
      ub1 = UBound(FormatlisteKomplett, 1)
      lb2 = LBound(FormatlisteKomplett, 2)
      ub2 = UBound(FormatlisteKomplett, 2)
      j = lb1
      AnzKat = SplitDelim(KategorienListe, Kategorien, ";")
    
    'Zunächst Filtermarkierungen setzen.
    For i = lb1 To ub1
      'Jedes Format der Gesamtliste prüfen
      k = 1
      FormatUebernommen = False
      Do While ((k <= AnzKat) And (Not FormatUebernommen))
        'Jede Kategorie der ZielKategorienListe prüfen.
        Kategorie = Kategorien(k)
        If (Not entspricht("Tabelle", Kategorie)) Then
          'Die erste Kategorie ist i.d.R. der Tabellen-Kodename => Standardname ist hiermit ignoriert.
          If (ThisWorkbook.Konfig.KategorieAlsFilter(Kategorie)) Then
            'Die aktive Kategorie soll als Filter verwendet werden.
            If (InStr(1, FormatlisteKomplett(i, idxFmtKategorien), Kategorie, vbTextCompare) > 0) Then
              'Die aktive Kategorie ist in der Kategorienliste des Formates enthalten => Index der Gesamtliste merken.
              ReDim Preserve Indizes(lb1 To j)
              Indizes(j) = i
              FormatUebernommen = True
              AnzFormate = AnzFormate + 1
              j = j + 1
            End If
          End If
        End If
        k = k + 1
      Loop
    Next
    
    If (AnzFormate > 0) Then
      ReDim FormatlisteGefiltert(lb1 To UBound(Indizes), lb2 To ub2)
      For i = lb1 To UBound(Indizes)
        For k = lb2 To ub2
          FormatlisteGefiltert(i, k) = FormatlisteKomplett(Indizes(i), k)
        Next
      Next
    End If
  End If
  
  FilternFormatliste = AnzFormate
  
  Exit Function
  
  Fehler:
  FehlerNachricht "CdatExpim.FilternFormatliste()"
End Function


Private Sub GetZielFormatliste(Liste_komplett As Variant, FormatLetzteWahl As String)
  'Belegt die Liste der Zielformate und selektiert ein Format.
  '6 Spalten (DateiVorname (Format-Kurzname), Titel, Pfad\Name  (Format-ID), Kategorienliste, DateidialogFilter, io_Typ).
  'Anzeige nur des Vornamens und des Titels, aber PfadName als Wert.
  'Parameter: Liste_komplett ... die komplette, ungefilterte Liste.
  
  On Error GoTo Fehler
  
  Dim Liste_gefiltert()    As String
  Dim Feld()               As String
  Dim SuchMuster           As String
  Dim FilterSetzen         As Boolean
  Dim AnzFormate           As Long
  Dim NF                   As Long
  Dim idxAuswahl           As Integer
  Dim idxLetzteWahl        As Integer
  Dim idxStandard          As Integer
  Dim KategorieStandard    As String
  Dim i                    As Long
  
  FilterSetzen = Not chkZiel_Formate.value
  
  If (FilterSetzen) Then
    AnzFormate = FilternFormatliste(Liste_komplett, Liste_gefiltert, Ziel_Kategorien)
  Else
    AnzFormate = UBound(Liste_komplett, 1) + 1
  End If
  
  If (AnzFormate > 0) Then
    
    If (FilterSetzen) Then
      Me.LstZiel_Formate.List = Liste_gefiltert
    Else
      Me.LstZiel_Formate.List = Liste_komplett
    End If
    'Application.EnableEvents = False
    Me.LstZiel_Formate.BoundColumn = 3
    'Application.EnableEvents = True
    
    'Standardkategorie finden (1. Eintrag in der Kategorienliste).
    NF = SplitDelim(Ziel_Kategorien, Feld, ";")
    If (NF > 0) Then
      KategorieStandard = Feld(1)
    Else
      KategorieStandard = "<keine>"
    End If
    SuchMuster = LCase(KategorieStandard) & ";|" & LCase(KategorieStandard) & "$"
    
    'Vorauswahl eines Eintrages (Reihenfolge: 1. Eintrag mit der Standardkategorie, letzte Wahl, 1. Eintrag).
    idxAuswahl = 0
    idxStandard = -1
    idxLetzteWahl = -1
    For i = LBound(Me.LstZiel_Formate.List, 1) To UBound(Me.LstZiel_Formate.List, 1)
      'If (InStr(1, LCase(Me.LstZiel_Formate.List(i, idxFmtKategorien)), LCase(KategorieStandard), vbTextCompare) > 0) Then idxStandard = i
      If (entspricht(SuchMuster, LCase(Me.LstZiel_Formate.List(i, idxFmtKategorien)))) Then idxStandard = i
      If (LCase(Me.LstZiel_Formate.List(i, idxFmtKurzname)) = LCase(FormatLetzteWahl)) Then idxLetzteWahl = i
    Next
    If (idxStandard > -1) Then
      idxAuswahl = idxStandard
    ElseIf (idxLetzteWahl > -1) Then
      idxAuswahl = idxLetzteWahl
    End If
    'If (idxLetzteWahl > -1) Then
    '  idxAuswahl = idxLetzteWahl
    'ElseIf (idxStandard > -1) Then
    '  idxAuswahl = idxStandard
    'End If
    Me.LstZiel_Formate.Selected(idxAuswahl) = True
    
  Else
    Me.LstZiel_Formate.Clear
  End If
  
  Exit Sub
  
  Fehler:
  Me.LstZiel_Formate.Clear
  Err.Clear
  'FehlerNachricht "frmStartExpim.GetZielFormatliste()"
End Sub


Private Sub GetQuellFormatliste(Liste_komplett As Variant, FormatLetzteWahl As String)
  'Belegt die Liste der Quellformate und selektiert ein Format.
  '6 Spalten (DateiVorname (Format-Kurzname), Titel, Pfad\Name (Format-ID), Kategorienliste, DateidialogFilter, io_Typ).
  'Anzeige nur des Format-Kurznamens und des Titels, aber Format-ID als Wert.
  'Parameter: Liste_komplett   ... die komplette, ungefilterte Liste.
  '           FormatLetzteWahl ... Format-Kurzname des zuletzt gewählten Eintrages zwecks Selektion.
  
  On Error GoTo Fehler
  
  Dim Liste_gefiltert()    As String
  Dim FilterSetzen         As Boolean
  Dim AnzFormate           As Long
  Dim idxAuswahl           As Integer
  Dim FormatVorName        As String
  Dim i                    As Long
  
  'Noch kein Filter vorgesehen:
  'FilterSetzen = Not chkQuelle_Formate.value
  FilterSetzen = False
  
  If (FilterSetzen) Then
    'AnzFormate = FilternFormatliste(Liste_komplett, Liste_gefiltert, Quelle_Kategorien)
  Else
    AnzFormate = UBound(Liste_komplett, 1) + 1
  End If
  
  If (AnzFormate > 0) Then
    
    If (FilterSetzen) Then
      Me.LstQuelle_Formate.List = Liste_gefiltert
    Else
      Me.LstQuelle_Formate.List = Liste_komplett
    End If
    Application.EnableEvents = False
    Me.LstQuelle_Formate.BoundColumn = 3
    Application.EnableEvents = True
    
    'Vorauswahl eines Eintrages (Reihenfolge: letzte Wahl, 1. Eintrag).
    idxAuswahl = 0
    For i = LBound(Me.LstQuelle_Formate.List, 1) To UBound(Me.LstQuelle_Formate.List, 1)
      FormatVorName = Me.LstQuelle_Formate.List(i, idxFmtKurzname)
      If (LCase(FormatVorName) = LCase(FormatLetzteWahl)) Then
        idxAuswahl = i
        Exit For
      End If
    Next
    Me.LstQuelle_Formate.Selected(idxAuswahl) = True
    
  Else
    Me.LstQuelle_Formate.Clear
  End If
  
  Exit Sub
  
  Fehler:
  Me.LstQuelle_Formate.Clear
  Err.Clear
  'FehlerNachricht "frmStartExpim.GetQuellFormatliste()"
End Sub

' Für jEdit:  :collapseFolds=1:mode=vbscript:
