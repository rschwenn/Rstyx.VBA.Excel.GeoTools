VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInsertLines 
   Caption         =   "Leerzeilen einfügen"
   ClientHeight    =   2976
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   4608
   OleObjectBlob   =   "frmInsertLines.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInsertLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2004-2010  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'==================================================================================================
'Modul frmInsertLines
'==================================================================================================
'
'Dialog zum Einfügen von Leerzeilen in die aktive Tabelle.
'Der Dialog fragt die Vorgaben ab: Zeilenbereich, Intervall, Größe des Leerzeilenbereiches.
'==================================================================================================


Option Explicit

'Deklarationen
Private lngRangeFrom       As Long
Private lngRangeUntil      As Long
Private lngBlockInterval   As Long
Private lngBlockLength     As Long
Private insertCount        As Long

Private oUsedRange         As Range


'Ereignis-Routinen (Dialog-Ereignisse) ************************************************************

Private Sub UserForm_Initialize()
  'Verwendeter Tabellenbereich
  Set oUsedRange = ActiveWorkbook.ActiveSheet.UsedRange
  
  'Startwerte für Dialog
  Me.txtBlockInterval.Value = 2
  Me.txtBlockLength.Value = 1
  
  Me.txtRangeFrom.Value = Selection.Row
  if (Selection.Rows.Count > 1) then
    Me.txtRangeUntil.Value = Selection.Rows(Selection.Rows.Count).Row
  else
    Me.txtRangeUntil.Value = oUsedRange.Rows(oUsedRange.Rows.Count).Row
  end if
  
  Me.txtBlockInterval.AutoWordSelect = True
  Me.txtBlockInterval.SetFocus
End Sub

Private Sub UserForm_Terminate()
  Set oUsedRange = Nothing
End Sub


Private Sub txtRangeFrom_Change()
  'msgbox "txtRangeFrom_Change()"
  Call refreshUI
End Sub

Private Sub txtRangeUntil_Change()
  Call refreshUI
End Sub

Private Sub txtBlockInterval_Change()
  Call refreshUI
End Sub

Private Sub txtBlockLength_Change()
  Call refreshUI
End Sub

Private Sub btnCancel_Click()
  Unload Me
End Sub

Private Sub btnOK_Click()
  Call insertLines
  Unload Me
End Sub


'interne Routinen *********************************************************************************

Private Sub refreshUI()
  'msgbox "refreshUI()"
  
  Dim success         As Boolean
  Dim LineCount       As Long
  Dim lastLine        As Long
  Dim lastLineFuture  As Long
  
  On Error GoTo Fehler
  Application.EnableEvents = False
  success = True
  
  'Werte der Textfelder prüfen und übernehmen
  Call checkTextfield(Me.txtRangeFrom, lngRangeFrom, success)
  Call checkTextfield(Me.txtRangeUntil, lngRangeUntil, success)
  Call checkTextfield(Me.txtBlockInterval, lngBlockInterval, success)
  Call checkTextfield(Me.txtBlockLength, lngBlockLength, success)
  
  If (success) Then
    'Vorberechnungen
    LineCount = lngRangeUntil - lngRangeFrom + 1
    insertCount = Int(LineCount / lngBlockInterval)
    If ((LineCount Mod lngBlockInterval) > 0) Then insertCount = insertCount + 1
    
    lastLine = oUsedRange.Rows(oUsedRange.Rows.Count).Row
    lastLineFuture = lastLine + insertCount * lngBlockLength
    
    'Statistikanzeige aktualisieren
    Me.lblInsertCountNumber = CStr(insertCount)
    Me.lblTabEndNumber = CStr(lastLineFuture)
    
    If (insertCount < 1) Then
      success = False
    End If
  Else
    
    'Statistikanzeige aktualisieren
    Me.lblInsertCountNumber = ""
    Me.lblTabEndNumber = ""
  End If
  
  'OK-Knopf
  If (success) Then
    btnOK.Enabled = True
  Else
    btnOK.Enabled = False
  End If
  
  Application.EnableEvents = True
  Exit Sub
  
Fehler:
  Application.EnableEvents = True
  btnOK.Enabled = False
  FehlerNachricht "frmInsertLines.refreshUI()"
End Sub


Private Sub checkTextfield(ByRef oTextField As Object, ByRef lngValue As Long, ByRef success As Boolean)
  'Das angegebene Textfeld wird auf numerischen Inhalt geprüft ...
  'success wird nur bei Mißerfolg geändert.
  If (IsNumeric(oTextField.Value)) Then
    lngValue = CLng(oTextField.Value)
    oTextField.Value = lngValue
    'success = true
  Else
    success = False
    lngValue = -1
    oTextField.Value = ""
  End If
End Sub


Private Sub insertLines()
  'Hier wird die Arbeit erledigt!
  On Error GoTo Fehler
  
  Dim StatusScreen As Boolean
  Dim StatusCalc   As Boolean
  Dim oRange       As Range
  Dim i            As Long
  
  If (ActiveCell Is Nothing) Then
    Err.Raise vbObjectError + ErrNumKeineAktiveTabelle - vbObjectError, , ErrMsgKeineAktiveTabelle
  End If
  If (isTabellenSchutz) Then
    Err.Raise vbObjectError + ErrNumTabSchutz - vbObjectError, , "Aktion ist nicht möglich, Tabellenschutz ist aktiv."
  End If
  
  'Eventuelle Mehrfachselection von Tabellen aufheben
  ActiveSheet.Select True
  StatusScreen = Application.ScreenUpdating
  StatusCalc = ActiveSheet.EnableCalculation
  Application.ScreenUpdating = False
  ActiveSheet.EnableCalculation = False
  
  'Startzeile
  Set oRange = Cells(lngRangeFrom, 1).Resize(RowSize:=lngBlockLength).EntireRow
  
  'Zeilen einfügen
  For i = 1 To insertCount
    oRange.Insert Shift:=xlDown
    Set oRange = oRange.Offset(rowoffset:=lngBlockInterval)
  Next
  
  'Aufräumen
  Application.ScreenUpdating = StatusScreen
  ActiveSheet.EnableCalculation = StatusCalc
  Set oRange = Nothing
  Exit Sub
  
Fehler:
  Application.ScreenUpdating = StatusScreen
  If (Not ActiveSheet Is Nothing) Then ActiveSheet.EnableCalculation = StatusCalc
  Set oRange = Nothing
  FehlerNachricht "frmInsertLines.insertLines()"
End Sub


'für jEdit:  :folding=indent::collapseFolds=1:
