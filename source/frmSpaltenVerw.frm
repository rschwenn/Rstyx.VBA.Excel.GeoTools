VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSpaltenVerw 
   Caption         =   "Tabellenstruktur verwalten"
   ClientHeight    =   3828
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   8436
   OleObjectBlob   =   "frmSpaltenVerw.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmSpaltenVerw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2004-2009  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'==================================================================================================
'Modul frmSpaltenVerw
'==================================================================================================
'
'Dialog zum Erzeugen/Verwalten der Tabellenstruktur und der Spaltennamen
'sowie Felder für Projektdaten.
'==================================================================================================


Option Explicit

'Windows API-Aufrufe für Spezialverhalten (Entmodalisierung :-).
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long

'Zwecks Empfang von Application-Ereignissen.
Private WithEvents App                        As Application
Private WithEvents Sheet                      As Excel.Worksheet
Private strAltesBlatt                         As String
Private strNeuesBlatt                         As String
Private keinAktivesBlatt                      As Boolean

'Identische Indizes der Listenfelder.
Private Const idxSpName                       As Integer = 1
Private Const idxSpTitel                      As Integer = 3
Private Const idxSpKategorie                  As Integer = 5

'Indizes der Liste "verfügbare Spaltennamen (Konfig).
Private Const idxQuelle_Dummy_0               As Integer = 0
Private Const idxQuelle_Groesse               As Integer = 2
Private Const idxQuelle_Format                As Integer = 4

'Indizes der Liste "Zielspalten.
Private Const idxZiel_Zeile                   As Integer = 0
Private Const idxZiel_Einheit                 As Integer = 2
Private Const idxZiel_Adresse                 As Integer = 4

Private Const strJedeKategorie                As String = "JedeKategorie"
Private Const strKategorie_AlleSpalten        As String = "<== alle Spaltennamen ==>"
Private Const strKategorie_TabStruktur        As String = "Tabellenstruktur-Elemente"
Private Const strKategorie_PrjDat             As String = "Projektdaten (einzelne Zellen)"
Private Const strKategorie_OrtDat             As String = "Ortsdaten (einzelne Zellen)"
Private Const strOhneEinheit                  As String = "< ohne >"
Private Const strOhneStatus                   As String = "< ohne >"
Private Const strOhneQuellName                As String = "< ohne >"
Private Const strOhneQuellName_Titel          As String = "< Feld-/Spaltenname löschen >"
Private Const strOhneZielName                 As String = "< ohne >"
Private Const strOhneZielName_Titel           As String = "< Spalte trägt keine Bezeichnung >"
Private Const strUnbekZielName_Titel          As String = "< unbekannter Spaltenname >"

Private Const strKeineAktiveTabelle           As String = "< Es ist keine Tabelle aktiv! >"
Private Const strKeinInfotraeger              As String = "< Es ist kein Infoträger in der Tabelle vorhanden! >"
Private Const strTabelleGeschuetzt            As String = "< Die aktive Tabelle ist schreibgeschützt! >"
Private Const strAktiveAuswahl                As String = "< Aktive Zellauswahl ==> bitte wählen! >"

Private save_Backcolor_cbo                    As Long
Private save_Forecolor_cbo                    As Long

Private strHilfeAktiv                         As String
Private strHilfeInfotraeger                   As String
Private strHilfeFliesskomma                   As String
Private strHilfeFormel                        As String

Private Liste_Namen_Konfig()                  As String
Private Liste_Namen_PrjDat()                  As String
Private Liste_Namen_OrtDat()                  As String
Private Liste_Strukturelemente()              As String
Private Liste_Spalten_XlTab()                 As String
Private Liste_Ziel_Struktur()                 As String


'Aktuelle Dialogwerte
Private Type TStatus
  Kategorie                                   As String
  QuellName                                   As String
  QuellNameOhnePrefix                         As String
  LetzerQuellName                             As String
  QuellEinheit                                As String
  QuellSpalteNr                               As Long
  ZielName                                    As String
  ZielZeile                                   As String
  ZielSpalteNr                                As Long
  ZielSpalteKey                               As String
  LetzerZielSpalteKey                         As String
  WertStatus                                  As String
  WertStatusPrefix                            As String
  LetzerWertStatus                            As String
End Type
Dim Status                                    As TStatus
'


'Ereignis-Routinen (Dialog-Ereignisse) ************************************************************

Private Sub UserForm_Initialize()
  
  'On Error GoTo Fehler
  'Dialog-Titel
  Me.Caption = ProgName & " - Tabellenstruktur verwalten"

  'Hilfetexte.
  strHilfeInfotraeger = "Der 'Informationsträger' der Tabelle legt folgende Dinge für den Datenbereich fest: " & vbNewLine & vbNewLine & _
             "   1. Beginn des Datenbereiches (1. Zeile)" & vbNewLine & _
             "   2. Ausdehnung des Datenbereiches (Spalten)" & vbNewLine & _
             "   3. Formatierung des Datenbereiches" & vbNewLine

  strHilfeFliesskomma = "Der 'Fliesskomma'-Bereich der Tabelle legt für den Datenbereich die Spalten fest, " & _
             "für die menügesteuert die Anzahl der Nachkommastellen geändert werden kann." & vbNewLine & vbNewLine & _
             "==> Dafür müssen alle entsprechenden Zellen markiert sein. Die Markierung kann aus " & _
             "mehreren Teilen bestehen und muß innerhalb des 'Informationsträgers' liegen. "

  strHilfeFormel = "Der 'Formel'-Bereich der Tabelle legt für den Datenbereich die Spalten fest, für die " & _
             "die Formeln der ersten Datenzeile auf alle anderen Zeilen übertragen werden können." & vbNewLine & vbNewLine & _
             "==> Dafür müssen alle entsprechenden Zellen markiert sein. Die Markierung kann aus " & _
             "mehreren Teilen bestehen und muß innerhalb des 'Informationsträgers' liegen. "
  
  'Listenstruktur und -Ansicht.
  Me.LstSpaltennamen.ColumnCount = 6
  Me.LstSpaltennamen.ColumnWidths = "0;80;0;;0;0"
  Me.cboZiel.ColumnCount = 6
  Me.cboZiel.ColumnWidths = ";0;0;0;0;0"
  Me.LstEinheiten.ColumnWidths = "62"
  
  'Aktive Hintergrundfarben sichern.
  save_Backcolor_cbo = Me.cboZiel.BackColor
  save_Forecolor_cbo = Me.cboZiel.ForeColor
  
  'Auswahl-Listen vorbereiten.
  Call GetListe_Namen_Konfig
  Call GetListe_Ziel_Struktur
  Call GetListe_Spalten_XlTab
  
  'Eigenschaften von oAktiveTabelle mit realer Tabelle abgleichen.
  'oAktiveTabelle.Syncronisieren
  
  'Dropdownfeld "Kategorien" belegen => Ereignisse erledigen den Rest.
  Call GetListe_Kategorien

  'Application-Ereignisse empfangen
  ErrMessage = "Fehler beim Initialisiern des Application-Objektes zur Ereignisauswertung"
  Set App = Application
  ErrMessage = "Fehler beim Initialisiern des Worksheet-Objektes zur Ereignisauswertung"
  Set Sheet = ActiveSheet
  ErrMessage = ""
Exit Sub

Fehler:
  'Set App = Nothing
  'Set Sheet = Nothing
  FehlerNachricht "frmSpaltenVerw.Class_Initialize()"
End Sub


'Private Sub UserForm_Activate()
'  'Dialog "Nicht Modal" schalten durch (wieder-)einschalten des Excel-Programmfensters,
'  'das beim Initialisieren des Dialoges automatisch ausgeschaltet wurde.
'  EnableWindow FindWindow("XLMAIN", Application.Caption), 1
'  'Application.CommandBars("Visual Basic").Visible = True
'End Sub


Private Sub UserForm_Terminate()
  Set App = Nothing
  Set Sheet = Nothing
End Sub



Private Sub btnAbbruch_Click()
  Unload Me
End Sub


Private Sub btnAktion_Click()
  'Gewählte Aktion durchführen.
  
  Dim Einheit  As String
  Dim Name     As String
  
  On Error GoTo Fehler
  
  'Call ZeigeEinstellungen
  
  If (Status.Kategorie = strKategorie_TabStruktur) Then
    'Stukturelement einfügen/ändern.
    Select Case Status.QuellName
      Case strInfoTraeger:  Call oAktiveTabelle.Selection2Infotraeger
      Case strFliesskomma:  Call oAktiveTabelle.Selection2Fliesskomma
      Case strFormel:       Call oAktiveTabelle.Selection2Formel
    End Select
  Else
    If (Status.QuellName = strOhneQuellName) Then
      Name = ""
    Else
      Name = Status.QuellName
    End If
    If ((Status.Kategorie = strKategorie_PrjDat) or (Status.Kategorie = strKategorie_OrtDat)) Then
      'Orts- oder Projektdatenfeld
      Call oAktiveTabelle.Selection2Feldname(Name)
    Else
      'Spaltenbezeichnung...
      If (Status.QuellEinheit = strOhneEinheit) Then Einheit = "" Else Einheit = Status.QuellEinheit
      Call oAktiveTabelle.Selection2Spaltenname(Name, Einheit)
    End If
  End If
  Call Syncronisieren
  Exit Sub
  
Fehler:
  FehlerNachricht "frmSpaltenVerw.btnAktion_Click()"
End Sub


Private Sub btnQuelleZeigen_Click()
  Dim Adresse   As String
  'On Error Resume Next
  If (Status.Kategorie = strKategorie_TabStruktur) Then
    Adresse = Me.LblHinweis2.Caption
  ElseIf ((Status.Kategorie = strKategorie_PrjDat) or (Status.Kategorie = strKategorie_OrtDat)) Then
    Adresse = Me.LblHinweis2.Caption
    'Adresse = GetLokalerZellname(Status.QuellName)
  Else
    Adresse = oAktiveTabelle.SpaltenErsteZellen(Status.QuellName).Address
  End If
  Range(Adresse).Select
End Sub


Private Sub cboZiel_Change()
  'MsgBox "cboZiel_Change"
  'If (Not IsNull(Me.cboZiel.value)) Then
  '  MsgBox Me.cboZiel.List(Me.cboZiel.ListIndex, idxZiel_Adresse)
  'End If
  
  If (Me.cboZiel.ListIndex > -1) Then
    
    Status.ZielZeile = Me.cboZiel.List(Me.cboZiel.ListIndex, idxZiel_Zeile)
    Status.ZielName = Me.cboZiel.List(Me.cboZiel.ListIndex, idxSpName)
    Status.ZielSpalteKey = MidStr(Me.cboZiel.List(Me.cboZiel.ListIndex, idxZiel_Adresse), "$", "$", False)
    Status.LetzerZielSpalteKey = Status.ZielSpalteKey
    
    'Steuerung der Anzeige von Steuerelementen im Ziel-Frame.
    If ((Status.ZielZeile = strKeineAktiveTabelle) Or (Status.ZielZeile = strKeinInfotraeger) _
        Or (Status.ZielZeile = strTabelleGeschuetzt)) Then
      Me.cboZiel.BackColor = Me.BackColor
      Me.cboZiel.Locked = True
      Me.fraZiel.Enabled = False
    Else
      If (Status.ZielZeile = strAktiveAuswahl) Then
        Me.cboZiel.BackColor = Me.BackColor
        Me.cboZiel.Locked = True
      Else
        Me.cboZiel.BackColor = save_Backcolor_cbo
        Me.cboZiel.Locked = False
      End If
      Me.fraZiel.Enabled = True
    End If
    
    'Im Spaltenwahlmodus: Gewählte Spalte im Infoträger selektieren.
    Application.EnableEvents = False
    Call ZielZelleMarkieren
    Application.EnableEvents = True
    
  End If
  
  Call Check_DialogAktion
  
End Sub


Private Sub cboKategorie_Change()
  'Kategorienfilter geändert.
  'MsgBox "cboKategorie_Change: "
  If (Me.cboKategorie.ListIndex > -1) Then
    Status.Kategorie = Me.cboKategorie.value
    If ((Status.Kategorie = strKategorie_TabStruktur) Or (Status.Kategorie = strKategorie_PrjDat) or (Status.Kategorie = strKategorie_OrtDat)) Then
      'Me.fraSpaltenBezeichnung.Caption = "einzufügendes Strukturelement"
      Me.LblHinweis1.Caption = "Derzeitiger Bereich:"
      Me.fraZiel.Caption = "Zielbereich"
    Else
      'Me.fraSpaltenBezeichnung.Caption = "zu vergebende Spaltenbezeichnung"
      'Me.LblHinweis1.Caption = "Der gewählte Name ist derzeit vergeben an:"
      Me.LblHinweis1.Caption = "Derzeitiger Bereich von ..."
      Me.fraZiel.Caption = "Zielspalte: Adresse, aktive Belegung (Name, Einheit, Beschreibung)"
    End If
    'Listen neu filtern.
    Call GetListe_ZielBereiche
    Call GetListe_QuellNamen
  End If
End Sub


Private Sub LstSpaltennamen_Change()
  '
  Dim Zelle             As Range
  'MsgBox "LstSpaltennamen_Change"
  
  strHilfeAktiv = ""
  
  If (Me.LstSpaltennamen.ListIndex > -1) Then
    Status.QuellNameOhnePrefix = Me.LstSpaltennamen.List(Me.LstSpaltennamen.ListIndex, idxSpName)
    Status.LetzerQuellName = Status.QuellNameOhnePrefix
    Call GetListe_Einheiten
    Call GetListe_Status
    'MsgBox GetSpaltenNr(Status.QuellName)
    
    If (Status.Kategorie = strKategorie_TabStruktur) Then
      'Wahl des Eintrages der Zielliste, falls Kategorie = strKategorie_TabStruktur.
      If (keinAktivesBlatt Or (ActiveCell Is Nothing)) Then
        Me.cboZiel.ListIndex = 0
        Me.LblHinweis2.Caption = "(keiner)"
        Me.btnQuelleZeigen.Enabled = False
      Else
        
        If (oAktiveTabelle.Infotraeger Is Nothing) Then
          If (Status.QuellName = strInfoTraeger) Then
            Me.cboZiel.ListIndex = 2
          Else
            Me.cboZiel.ListIndex = 1
          End If
        Else
          Me.cboZiel.ListIndex = 2
        End If
        'Anzeige des vom gewählten Strukturelement aktuell belegten Zellbereich.
        If (Status.QuellName = strInfoTraeger) Then
          strHilfeAktiv = strHilfeInfotraeger
          If (oAktiveTabelle.Infotraeger Is Nothing) Then
            Me.LblHinweis2.Caption = "(keiner)"
            Me.btnQuelleZeigen.Enabled = False
          Else
            Me.LblHinweis2.Caption = oAktiveTabelle.Infotraeger.Address(False, False)
            Me.btnQuelleZeigen.Enabled = True
          End If
        ElseIf (Status.QuellName = strFliesskomma) Then
          strHilfeAktiv = strHilfeFliesskomma
          If (oAktiveTabelle.Fliesskomma Is Nothing) Then
            Me.LblHinweis2.Caption = "(keiner)"
            Me.btnQuelleZeigen.Enabled = False
          Else
            Me.LblHinweis2.Caption = oAktiveTabelle.Fliesskomma.Address(False, False)
            Me.btnQuelleZeigen.Enabled = True
          End If
        ElseIf (Status.QuellName = strFormel) Then
          strHilfeAktiv = strHilfeFormel
          If (oAktiveTabelle.Formel Is Nothing) Then
            Me.LblHinweis2.Caption = "(keiner)"
            Me.btnQuelleZeigen.Enabled = False
          Else
            Me.LblHinweis2.Caption = oAktiveTabelle.Formel.Address(False, False)
            Me.btnQuelleZeigen.Enabled = True
          End If
        End If
      End If
    
    ElseIf ((Status.Kategorie = strKategorie_PrjDat) or (Status.Kategorie = strKategorie_OrtDat)) Then
      'Anzeige des vom gewählten Orts-/Projektdaten-Feldnamen aktuell belegten Zellbereich.
      If (Not (keinAktivesBlatt Or (ActiveCell Is Nothing))) Then
        Me.cboZiel.ListIndex = 2
        Set Zelle = GetLokalerZellname(Status.QuellName)
        If (Not Zelle Is Nothing) Then
          Me.LblHinweis2.Caption = Zelle.Address(False, False)
          Me.btnQuelleZeigen.Enabled = True
        Else
          Me.LblHinweis2.Caption = "(keiner)"
          Me.btnQuelleZeigen.Enabled = False
        End If
      Else
        Me.cboZiel.ListIndex = 0
        Me.LblHinweis2.Caption = "(keiner)"
        Me.btnQuelleZeigen.Enabled = False
      End If
    
    Else
      'Anzeige der mit dem gewählten SpaltenNamen aktuell belegten Spalte.
      Call Check_QuellNameExists
    End If

  End If
  
  Set Zelle = Nothing
  Call Check_DialogAktion
  
End Sub



Private Sub Check_QuellNameExists()
  'Anzeige der mit dem gewählten SpaltenNamen aktuell belegten Spalte.
  If (Not (keinAktivesBlatt Or (ActiveCell Is Nothing))) Then
    Me.LblHinweis1.Caption = "Derzeitiger Bereich von '" & Status.QuellName & "':"
    If (oAktiveTabelle.SpaltenErsteZellen.Exists(Status.QuellName)) Then
      Me.LblHinweis2.Caption = "Spalte " & MidStr(oAktiveTabelle.SpaltenErsteZellen(Status.QuellName).Address, "$", "$", False)
      Me.btnQuelleZeigen.Enabled = True
    Else
      Me.LblHinweis2.Caption = "keine Spalte"
      Me.btnQuelleZeigen.Enabled = False
    End If
  Else
    Me.LblHinweis1.Caption = "Derzeitiger Bereich:"
    Me.LblHinweis2.Caption = "keine Spalte"
    Me.btnQuelleZeigen.Enabled = False
  End If
End Sub



Private Sub LstEinheiten_Change()
  'MsgBox "LstEinheiten_Change(): '" & Me.LstEinheiten.value & "'"
  Dim Einheit  As Variant
  'Bug umgehen: Value-Eigenschaft wirf erst nach Mausklick gesetzt!
  Einheit = Me.LstEinheiten.List(Me.LstEinheiten.ListIndex)
  If (IsNull(Einheit)) Then
    Status.QuellEinheit = ""
  Else
    Status.QuellEinheit = Einheit
  End If
End Sub


Private Sub btnHilfe_Click()
  'Aktiven Hilfetext anzeigen.
  Call MsgBox(strHilfeAktiv, vbInformation + vbOKOnly, "Information")
  If (btnAktion.Enabled) Then btnAktion.SetFocus
End Sub



Private Sub cboStatus_Change()
  'MsgBox "cboStatus_Change(): '" & Me.cboStatus.value & "'"
  Dim WertStatus  As Variant
  
  If (Me.cboStatus.ListIndex > -1) Then
    
    'Bug umgehen: Value-Eigenschaft wirf erst nach Mausklick gesetzt!
    WertStatus = Me.cboStatus.List(Me.cboStatus.ListIndex)
    If (IsNull(WertStatus)) Then
      Status.WertStatus = ""
      Status.WertStatusPrefix = ""
    Else
      Status.WertStatus = WertStatus
      If (Status.WertStatus = strOhneStatus) Then
        Status.WertStatusPrefix = ""
      Else
        Status.WertStatusPrefix = oKonfig.StatusPrefix(Status.WertStatus)
      End If
    End If
    Status.LetzerWertStatus = Status.WertStatus
    Status.QuellName = Status.WertStatusPrefix & Status.QuellNameOhnePrefix
    
    If (Status.WertStatus = strOhneStatus) Then
      Me.cboStatus.Enabled = False
      Me.cboStatus.BackColor = Me.BackColor
    Else
      Me.cboStatus.Enabled = True
      Me.cboStatus.BackColor = save_Backcolor_cbo
    End If
    
    If (Status.Kategorie <> strKategorie_TabStruktur) And (Status.Kategorie <> strKategorie_PrjDat And (Status.Kategorie <> strKategorie_OrtDat)) Then
      'Anzeige der mit dem gewählten SpaltenNamen aktuell belegten Spalte.
      Call Check_QuellNameExists
    End If

  End If
  
  Call Check_DialogAktion

End Sub



'Ereignis-Routinen (Application-Ereignisse) *******************************************************


Private Sub App_SheetActivate(ByVal Sh As Object)
  'Wird aufgerufen beim Aktivieren eines Arbeitsblattes in (hoffentlich) jeder Situation.
  strNeuesBlatt = Sh.Parent.Name & "!" & Sh.Name
  strAltesBlatt = strNeuesBlatt
  'MsgBox "neues Arbeitsblatt: " & strNeuesBlatt
  
  'Empfang der Worksheet-Ereignisse vom neu gewählten Blatt.
  Set Sheet = Sh
  
  Call Syncronisieren
End Sub


'Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Excel.Range)
  'Tritt ein, wenn ein beliebiges Tabellenblatt durch den Benutzer oder durch eine
  'externe Verknüpfung geändert wird.
  '==> Keine Reaktion auf das Löschen von Zellen!
  '==> Keine Reaktion auf das Einfügen, Ändern, Löschen von Namen!
'  MsgBox "Geändertes Arbeitsblatt: " & Sh.Parent.Name & "!" & Sh.Name
'  Call Syncronisieren
'End Sub


Private Sub App_WindowActivate(ByVal Wb As Excel.Workbook, ByVal Wn As Excel.Window)
  'Löst "app_SheetActivate" aus, wenn beim Aktivieren eines Fensters auch das Arbeitsblatt wechselt.
  strNeuesBlatt = Wb.Name & "!" & Wb.ActiveSheet.Name
  If (strNeuesBlatt <> strAltesBlatt) Then
    strAltesBlatt = strNeuesBlatt
    Call App_SheetActivate(Wb.ActiveSheet)
  End If
End Sub


Private Sub App_WorkbookDeactivate(ByVal Wb As Excel.Workbook)
  'Reaktion auf das Deaktivieren der (noch) einzigen Arbeitsmappe, d.h. diese
  'wird geschlossen und danach (!) ist also kein Arbeitsblatt mehr aktiv.
  If (Application.Workbooks.Count = 1) Then
    strAltesBlatt = ""
    'Call BefehleAktualisieren(keinAktivesBlatt:=True)
    Call Syncronisieren(blnkeinAktivesBlatt:=True)
    'MsgBox "Gleich gibt's keine aktive Mappe mehr!"
  End If
End Sub


Private Sub Sheet_SelectionChange(ByVal Target As Excel.Range)
  'Wenn gerade eine Spalte gewählt werden soll, so wird sichergestellt,
  'daß sich die aktive Zelle innerhalb des Infoträgers befindet.
  
  'MsgBox "Sheet_SelectionChange()   Adresse=" & CStr(Target.Address(False, False))
  
  Dim AdressesSelection  As String
     
  If ((Status.Kategorie = strKategorie_PrjDat) or (Status.Kategorie = strKategorie_OrtDat)) Then
    Application.EnableEvents = False
    AdressesSelection = Selection_Korrektur()
    Application.EnableEvents = True
    Call Check_DialogAktion
  ElseIf (Status.Kategorie <> strKategorie_TabStruktur) Then
    Application.EnableEvents = False
    AdressesSelection = Selection_Korrektur()
    Application.EnableEvents = True
    Call WaehleListenEintrag(AdressesSelection)
  End If

End Sub





'interne Routinen *********************************************************************************

Private Sub GetListe_Kategorien()
  'Erzeugt eine Liste aller konfigurierten Kategorien.
  'Ergebnis ... Array Liste_Kategorien_Konfig (1 Zeile = Kategorienname ohne Prefix, Beschreibung, Größe, Kategorie, alsFilter, KategorienFormat).
  'Ereignis "cboKategorie_Change" wird ausgelöst...
  
  On Error GoTo Fehler
  
  Dim Kategorien                  As Variant
  Dim Kategorie                   As Variant
  Dim i                           As Long
  Dim AnzExtraEintraege           As Long
  Dim Liste_Kategorien_Konfig()   As String
  
  AnzExtraEintraege = 4
  ReDim Liste_Kategorien_Konfig(0 To AnzExtraEintraege - 1)
  Liste_Kategorien_Konfig(0) = strKategorie_TabStruktur
  Liste_Kategorien_Konfig(1) = strKategorie_PrjDat
  Liste_Kategorien_Konfig(2) = strKategorie_OrtDat
  Liste_Kategorien_Konfig(3) = strKategorie_AlleSpalten
  
  If (Not oKonfig.Kategorien Is Nothing) Then
    If (oKonfig.Kategorien.Count > 0) Then
      ReDim Preserve Liste_Kategorien_Konfig(0 To oKonfig.Kategorien.Count + AnzExtraEintraege - 1)
      i = AnzExtraEintraege
      For Each Kategorie In oKonfig.Kategorien.Keys
        Liste_Kategorien_Konfig(i) = Kategorie
        i = i + 1
      Next
    End If
  End If
  Me.cboKategorie.ColumnCount = 1
  Me.cboKategorie.List = Liste_Kategorien_Konfig
  Me.cboKategorie.ListIndex = 0
  Exit Sub
  
Fehler:
  FehlerNachricht "frmKategorienVerw.GetListe_Kategorien()"
End Sub



Private Sub GetListe_Namen_Konfig()
  'Erzeugt 3 Listen aller konfigurierten Namen von:
  ' - Strukturelementen
  ' - Feldnamen von Projektdaten (nur solche, die mit "Prj." beginnen
  ' - Spalten zwecks Auswahl zum Einfügen in die Tabelle.
  'Ergebnis ... Array Liste_Namen_Konfig (1 Zeile = "", Spaltennname ohne Prefix, Größe, Beschreibung, SpaltenFormat, Kategorie).
  
  'On Error GoTo Fehler
  
  'Deklarationen
    Dim Spalte                   As Variant
    Dim FeldName                 As Variant
    Dim i                        As Long
    Dim AnzExtraEintraege        As Long
    Dim AnzZeilen                As Long
    Dim PrjFeldNamen()           As String
    Dim PrjFeldTitel()           As String
    Dim OrtFeldNamen()           As String
    Dim OrtFeldTitel()           As String
  
  'Strukturelemente
    ReDim Liste_Strukturelemente(0 To 2, 0 To 5)
    Liste_Strukturelemente(0, idxSpName) = strInfoTraeger
    Liste_Strukturelemente(0, idxSpKategorie) = strKategorie_TabStruktur
    Liste_Strukturelemente(0, idxSpTitel) = "Ausdehnung und Format des Datenbereiches."
    Liste_Strukturelemente(1, idxSpName) = strFliesskomma
    Liste_Strukturelemente(1, idxSpKategorie) = strKategorie_TabStruktur
    Liste_Strukturelemente(1, idxSpTitel) = "Änderung Nachkommastellen beim Formatieren."
    Liste_Strukturelemente(2, idxSpName) = strFormel
    Liste_Strukturelemente(2, idxSpKategorie) = strKategorie_TabStruktur
    Liste_Strukturelemente(2, idxSpTitel) = "Formelübertragung beim Formatieren."
  
  'Feldnamen von Projektdaten (nur solche, die mit "Prj." beginnen.
    If (Not oMetadaten.AlleProjektDaten Is Nothing) Then
      i = -1
      For Each FeldName In oMetadaten.AlleProjektDaten
        If (Left(FeldName, 4) = "Prj.") Then
          i = i + 1
          ReDim Preserve PrjFeldNamen(0 To i)
          ReDim Preserve PrjFeldTitel(0 To i)
          PrjFeldNamen(i) = FeldName
          PrjFeldTitel(i) = oMetadaten.TitelDerProjektDaten(FeldName)
        End If
      Next
    End If
    AnzZeilen = i + 1
    AnzExtraEintraege = 1
    ReDim Liste_Namen_PrjDat(0 To AnzZeilen + AnzExtraEintraege - 1, 0 To 5)
    'Eintrag "ohne Feldname"
    Liste_Namen_PrjDat(0, idxSpName) = strOhneQuellName
    Liste_Namen_PrjDat(0, idxSpKategorie) = strKategorie_PrjDat
    Liste_Namen_PrjDat(0, idxSpTitel) = strOhneQuellName_Titel
    'Vorhandene Felder für Projektdaten.
    For i = 1 To AnzZeilen
      Liste_Namen_PrjDat(i, idxSpName) = PrjFeldNamen(i - 1)
      Liste_Namen_PrjDat(i, idxSpKategorie) = strKategorie_PrjDat
      Liste_Namen_PrjDat(i, idxSpTitel) = PrjFeldTitel(i - 1)
    Next
  
  'Feldnamen von Ortsdaten (nur solche, die mit "Ort." beginnen.
    If (Not oMetadaten.AlleProjektDaten Is Nothing) Then
      i = -1
      For Each FeldName In oMetadaten.AlleProjektDaten
        If (Left(FeldName, 4) = "Ort.") Then
          i = i + 1
          ReDim Preserve OrtFeldNamen(0 To i)
          ReDim Preserve OrtFeldTitel(0 To i)
          OrtFeldNamen(i) = FeldName
          OrtFeldTitel(i) = oMetadaten.TitelDerProjektDaten(FeldName)
        End If
      Next
    End If
    AnzZeilen = i + 1
    AnzExtraEintraege = 1
    ReDim Liste_Namen_OrtDat(0 To AnzZeilen + AnzExtraEintraege - 1, 0 To 5)
    'Eintrag "ohne Feldname"
    Liste_Namen_OrtDat(0, idxSpName) = strOhneQuellName
    Liste_Namen_OrtDat(0, idxSpKategorie) = strKategorie_OrtDat
    Liste_Namen_OrtDat(0, idxSpTitel) = strOhneQuellName_Titel
    'Vorhandene Felder für Ortsdaten.
    For i = 1 To AnzZeilen
      Liste_Namen_OrtDat(i, idxSpName) = OrtFeldNamen(i - 1)
      Liste_Namen_OrtDat(i, idxSpKategorie) = strKategorie_OrtDat
      Liste_Namen_OrtDat(i, idxSpTitel) = OrtFeldTitel(i - 1)
    Next
  
  'Konfigurierte Spaltennamen.
    AnzExtraEintraege = 1
    If (Not oKonfig.SpaltenBeschreibung Is Nothing) Then
      If (oKonfig.SpaltenBeschreibung.Count = 0) Then
        ReDim Liste_Namen_Konfig(0 To AnzExtraEintraege - 1, 0 To 5)
      Else
        ReDim Liste_Namen_Konfig(0 To oKonfig.SpaltenBeschreibung.Count + AnzExtraEintraege - 1, 0 To 5)
        i = AnzExtraEintraege
        For Each Spalte In oKonfig.SpaltenBeschreibung.Keys
          Liste_Namen_Konfig(i, idxSpName) = Spalte
          Liste_Namen_Konfig(i, idxSpTitel) = oKonfig.SpaltenBeschreibung(Spalte)
          Liste_Namen_Konfig(i, idxQuelle_Groesse) = oKonfig.SpaltenMathGroesse(Spalte)
          Liste_Namen_Konfig(i, idxSpKategorie) = oKonfig.SpaltenKategorie(Spalte)
          Liste_Namen_Konfig(i, idxQuelle_Dummy_0) = ""
          Liste_Namen_Konfig(i, idxQuelle_Format) = oAktiveTabelle.SpaltenFormate(Spalte)
          i = i + 1
        Next
      End If
      'Eintrag "ohne Spaltenname"
      Liste_Namen_Konfig(0, idxSpName) = strOhneQuellName
      Liste_Namen_Konfig(0, idxSpKategorie) = "#"
      Liste_Namen_Konfig(0, idxSpTitel) = strOhneQuellName_Titel
    End If
  Exit Sub
  
Fehler:
  FehlerNachricht "frmSpaltenVerw.GetListe_Namen_Konfig()"
End Sub



Private Sub GetListe_QuellNamen()
  'Belegt die Listbox der verfügbaren Spalten, u.U. gefiltert.
  '6 Spalten (1 Zeile = "", Spaltennname ohne Prefix, Größe, Beschreibung, SpaltenFormat, Kategorie).
  
  On Error GoTo Fehler
  
  Dim Liste_gefiltert()    As String
  Dim Kategorie            As String
  Dim AnzZeilen            As Long
  Dim idxAuswahl           As Integer
  Dim idxLetzteWahl        As Integer
  Dim i                    As Long
  
  Kategorie = Status.Kategorie
  If (Kategorie = strKategorie_TabStruktur) Then
    AnzZeilen = FilternListe(Liste_Strukturelemente, Liste_gefiltert, Kategorie)
  ElseIf (Kategorie = strKategorie_PrjDat) Then
    AnzZeilen = FilternListe(Liste_Namen_PrjDat, Liste_gefiltert, Kategorie)
  ElseIf (Kategorie = strKategorie_OrtDat) Then
    AnzZeilen = FilternListe(Liste_Namen_OrtDat, Liste_gefiltert, Kategorie)
  Else
    If (Kategorie = strKategorie_AlleSpalten) Then Kategorie = ""
    AnzZeilen = FilternListe(Liste_Namen_Konfig, Liste_gefiltert, Kategorie)
  End If
        
  If (AnzZeilen > 0) Then
    Me.LstSpaltennamen.List = Liste_gefiltert
    Me.LstSpaltennamen.BoundColumn = 1
    
    'Vorauswahl eines Eintrages.
    idxAuswahl = 0
    idxLetzteWahl = -1
    For i = LBound(Me.LstSpaltennamen.List, 1) To UBound(Me.LstSpaltennamen.List, 1)
      If (Me.LstSpaltennamen.List(i, idxSpName) = Status.LetzerQuellName) Then idxLetzteWahl = i
    Next
    If (idxLetzteWahl > -1) Then idxAuswahl = idxLetzteWahl
    
    Me.LstSpaltennamen.Selected(idxAuswahl) = True
  Else
    Me.LstSpaltennamen.Clear
  End If
  
  Exit Sub
  
Fehler:
  Me.LstSpaltennamen.Clear
  Err.Clear
  'FehlerNachricht "frmSpaltenVerw.GetListe_QuellNamen()"
End Sub



Private Function FilternListe(ListeKomplett As Variant, ListeGefiltert() As String, ByVal Kategorie As String) As Long
  'Filtert eine Liste mit Spaltennamen nach der Kategorie.
  'Parameter:  ListeKomplett  ... vollständige Liste (Array).
  '            ListeGefiltert ... nach Kategorien gefilterte Liste (Array).
  '            Kategorie      ... Name einer Kategorie.
  'Rückgabe:   Anzahl der Einträge der gefilterten Liste.
  'Aufbau der beiden Listenfelder:    1 Zeile = 6 Spalten (DateiVorname (Format-Kurzname), Titel, Pfad\Name  (Format-ID), Kategorie, DateidialogFilter, io_Typ).
  
  On Error GoTo Fehler
  
  Dim Indizes()                 As Long
  Dim AnzZeilen                 As Long
  Dim i                         As Long
  Dim j                         As Long
  Dim k                         As Long
  Dim lb1                       As Long
  Dim lb2                       As Long
  Dim ub1                       As Long
  Dim ub2                       As Long
  Dim Kat                       As String
  
  AnzZeilen = 0
  lb1 = LBound(ListeKomplett, 1)
  ub1 = UBound(ListeKomplett, 1)
  lb2 = LBound(ListeKomplett, 2)
  ub2 = UBound(ListeKomplett, 2)
  j = lb1
  Erase ListeGefiltert
  
  For i = lb1 To ub1
    Kat = ListeKomplett(i, idxSpKategorie)
    If ((Kat = Kategorie) Or (InStr(1, Kat, Kategorie, vbTextCompare) > 0) Or (Kat = strJedeKategorie)) Then
      ReDim Preserve Indizes(lb1 To j)
      Indizes(j) = i
      AnzZeilen = AnzZeilen + 1
      j = j + 1
    End If
  Next
  
  If (AnzZeilen > 0) Then
    ReDim ListeGefiltert(lb1 To UBound(Indizes), lb2 To ub2)
    For i = lb1 To UBound(Indizes)
      For k = lb2 To ub2
        ListeGefiltert(i, k) = ListeKomplett(Indizes(i), k)
      Next
    Next
  End If
  
  FilternListe = AnzZeilen
  Exit Function
  
Fehler:
  FehlerNachricht "frmSpaltenVerw.FilternListe()"
End Function



Private Sub GetListe_Einheiten()
  'Erzeugt eine Liste aller für den gewählten Spaltennamen verfügbarer Einheiten.
  'Ergebnis ... Array Liste_Einheiten (1 Zeile = Einheit).
  
  On Error GoTo Fehler
  
  Dim Einheiten           As Variant
  Dim Einheit             As Variant
  Dim i                   As Long
  Dim idxAuswahl          As Long
  Dim Liste_Einheiten()   As String
  Dim AktGroesse          As String
  
  idxAuswahl = 0
  ReDim Liste_Einheiten(0 To 0)
  Liste_Einheiten(0) = strOhneEinheit
  
  AktGroesse = Me.LstSpaltennamen.List(Me.LstSpaltennamen.ListIndex, idxQuelle_Groesse)
  Me.LblEinheiten2.Caption = AktGroesse
  
  If (Not oKonfig.Einheiten Is Nothing) Then
    If (oKonfig.Einheiten.Exists(AktGroesse)) Then
      If (oKonfig.Einheiten(AktGroesse).Count > 0) Then
        ReDim Preserve Liste_Einheiten(0 To oKonfig.Einheiten(AktGroesse).Count)
        i = 1
        For Each Einheit In oKonfig.Einheiten(AktGroesse).Keys
          Liste_Einheiten(i) = Einheit
          If (oKonfig.Einheiten(AktGroesse)(Einheit) = "1") Then idxAuswahl = i
          i = i + 1
        Next
      End If
    End If
  End If
  Me.LstEinheiten.ColumnCount = 1
  Me.LstEinheiten.List = Liste_Einheiten
  'Me.LstEinheiten.BoundColumn = 1
  'Me.LstEinheiten.value = Liste_Einheiten(idxAuswahl)
  'Me.LstEinheiten.Selected(idxAuswahl) = True
  Me.LstEinheiten.ListIndex = idxAuswahl
  
  Exit Sub

Fehler:
  FehlerNachricht "frmSpaltenVerw.GetListe_Einheiten()"
End Sub



Private Sub GetListe_Status()
  'Erzeugt eine Liste aller verfügbarer Stati für numerische Werte.
  'Ergebnis ... Array Liste_Status (1 Zeile = Status).
  
  On Error GoTo Fehler
  
  Dim WertStatus          As Variant
  Dim i                   As Long
  Dim idxAuswahl          As Long
  Dim Liste_Status()      As String
  Dim AktGroesse          As String
  
  idxAuswahl = 0
  If (Status.QuellEinheit = strOhneEinheit) Then
    ReDim Liste_Status(0 To 0)
    Liste_Status(0) = strOhneStatus
  Else
    If (Not oKonfig.StatusPrefix Is Nothing) Then
      If (oKonfig.StatusPrefix.Count > 0) Then
        ReDim Liste_Status(0 To oKonfig.StatusPrefix.Count - 1)
        i = 0
        For Each WertStatus In oKonfig.StatusPrefix.Keys
          Liste_Status(i) = WertStatus
          If (WertStatus = Status.LetzerWertStatus) Then idxAuswahl = i
          i = i + 1
        Next
      End If
    End If
  End If
  Me.cboStatus.ColumnCount = 1
  Me.cboStatus.List = Liste_Status
  Me.cboStatus.ListIndex = idxAuswahl
  
  Exit Sub

Fehler:
  FehlerNachricht "frmSpaltenVerw.GetListe_Status()"
End Sub



Private Sub GetListe_Spalten_XlTab()
  'Erzeugt eine Liste aller Spaltennamen der aktiven Tabelle.
  'Ergebnis ... Array Liste_Spalten_XlTab (1 Zeile = XL-Spaltenbez. (Buchstabe),  Spaltennname ohne Prefix, Einheit, Beschreibung, SpaltenFormat, Kategorie).
  
  On Error GoTo Fehler
  
  Dim Spalte               As Variant
  Dim i                    As Long
  Dim SpAnz                As Long
  Dim ZeAnf                As Long
  Dim SpAnf                As Long
  Dim SpEnd                As Long
  Dim lb                   As Long
  Dim ub                   As Long
  Dim Zelle                As Range
  Dim oSpNameAttr          As Scripting.Dictionary
  
  Erase Liste_Spalten_XlTab
  
  oAktiveTabelle.Syncronisieren
  
  If (keinAktivesBlatt Or (ActiveCell Is Nothing)) Then
    ReDim Liste_Spalten_XlTab(0 To 0, 0 To 5)
    Liste_Spalten_XlTab(0, idxZiel_Zeile) = strKeineAktiveTabelle
    Liste_Spalten_XlTab(0, idxSpKategorie) = strJedeKategorie
  
  ElseIf (oAktiveTabelle.Infotraeger Is Nothing) Then
    ReDim Liste_Spalten_XlTab(0 To 0, 0 To 5)
    Liste_Spalten_XlTab(0, idxZiel_Zeile) = strKeinInfotraeger
    Liste_Spalten_XlTab(0, idxSpKategorie) = strJedeKategorie
  
  Else
    'Festwerte des Datenbereiches ermitteln
    ZeAnf = oAktiveTabelle.ErsteDatenZeile
    'If (Err) Then GoTo Fehler
    SpAnf = oAktiveTabelle.ErsteDatenSpalte
    'If (Err) Then GoTo Fehler
    SpEnd = oAktiveTabelle.LetzteDatenSpalte
    'If (Err) Then GoTo Fehler
    SpAnz = SpEnd - SpAnf + 1
    ReDim Liste_Spalten_XlTab(0 To SpAnz - 1, 0 To 5)
    lb = LBound(Liste_Spalten_XlTab, 1)
    ub = UBound(Liste_Spalten_XlTab, 1)
    
    For i = lb To ub
      Set Zelle = Cells(ZeAnf, SpAnf + i - lb)
      Liste_Spalten_XlTab(i, idxZiel_Adresse) = Zelle.Address
      Liste_Spalten_XlTab(i, idxSpName) = strOhneZielName
      Liste_Spalten_XlTab(i, idxSpTitel) = strOhneZielName_Titel
      Liste_Spalten_XlTab(i, idxZiel_Einheit) = " "
      Liste_Spalten_XlTab(i, idxSpKategorie) = strJedeKategorie
      For Each Spalte In oAktiveTabelle.SpaltenErsteZellen.Keys
        If (Zelle.Column = oAktiveTabelle.SpaltenErsteZellen(Spalte).Column) Then
          Liste_Spalten_XlTab(i, idxSpName) = Spalte
          Liste_Spalten_XlTab(i, idxZiel_Einheit) = oAktiveTabelle.SpaltenEinheiten(Spalte)
          Set oSpNameAttr = oKonfig.SpNameAttr(Spalte)
          If (oSpNameAttr("Titel") <> SpTitel_unbekannt) Then
            'Liste_Spalten_XlTab(i, idxSpTitel) = oSpNameAttr("StatusBez") & ": " & oSpNameAttr("Titel")
            Liste_Spalten_XlTab(i, idxSpTitel) = oSpNameAttr("Titel")
            Liste_Spalten_XlTab(i, idxSpKategorie) = oSpNameAttr("Kategorie")
          Else
            Liste_Spalten_XlTab(i, idxSpTitel) = strUnbekZielName_Titel
            Liste_Spalten_XlTab(i, idxSpKategorie) = strJedeKategorie
          End If
          Exit For
        End If
      Next
      Liste_Spalten_XlTab(i, idxZiel_Zeile) = Format(MidStr(Liste_Spalten_XlTab(i, idxZiel_Adresse), "$", "$", False), "!@@") & "  " & _
                                              Format(Liste_Spalten_XlTab(i, idxSpName), "!@@@@@@@@@@@@@@@@@@@@") & _
                                              Format(Liste_Spalten_XlTab(i, idxZiel_Einheit), "!@@@@@@@@") & _
                                              Liste_Spalten_XlTab(i, idxSpTitel)
    Next
  End If
  
  Set oSpNameAttr = Nothing
  Exit Sub

Fehler:
  Set oSpNameAttr = Nothing
  FehlerNachricht "frmSpaltenVerw.GetListe_Spalten_XlTab()"
End Sub



Private Sub GetListe_Ziel_Struktur()
  'Erzeugt eine Liste der möglichen Zielmeldungen für Strukturelemente.
  'Ergebnis ... Array Liste_Ziel_Struktur (1 Zeile = XL-Spaltenbez. (Buchstabe),  Spaltennname ohne Prefix, Einheit, Beschreibung, SpaltenFormat, Kategorie).
  
  On Error GoTo Fehler
  Erase Liste_Ziel_Struktur
  ReDim Liste_Ziel_Struktur(0 To 2, 0 To 5)
  Liste_Ziel_Struktur(0, idxZiel_Zeile) = strKeineAktiveTabelle
  Liste_Ziel_Struktur(0, idxSpKategorie) = strJedeKategorie
  Liste_Ziel_Struktur(1, idxZiel_Zeile) = strKeinInfotraeger
  Liste_Ziel_Struktur(1, idxSpKategorie) = strJedeKategorie
  Liste_Ziel_Struktur(2, idxZiel_Zeile) = strAktiveAuswahl
  Liste_Ziel_Struktur(2, idxSpKategorie) = strJedeKategorie
  Exit Sub

Fehler:
  FehlerNachricht "frmSpaltenVerw.GetListe_Ziel_Struktur()"
End Sub



Private Sub GetListe_ZielBereiche()
  'Belegt die Ziel-Dropdownliste mit den innerhalb des Infoträgers verfügbaren Spalten
  'oder mit den verfügbaren Meldungen, falls Kategorie = Strukturelemente.
  '6 Spalten (SpaltenName, Beschreibung, Groesse, Kategorie, KategorieAlsFilter, SpaltenFormat).
  
  On Error GoTo Fehler
  
  Dim Liste_gefiltert()    As String
  Dim Kategorie            As String
  Dim ZielSpalteKey        As String
  Dim AnzZeilen            As Long
  Dim idxAuswahl           As Integer
  Dim idxLetzteWahl        As Integer
  Dim i                    As Long
  
  Kategorie = Status.Kategorie
  If ((Kategorie = strKategorie_TabStruktur) Or (Kategorie = strKategorie_PrjDat) or (Status.Kategorie = strKategorie_OrtDat)) Then
    Kategorie = ""
    AnzZeilen = FilternListe(Liste_Ziel_Struktur, Liste_gefiltert, Kategorie)
  Else
    'If (Kategorie = strKategorie_AlleSpalten) Then Kategorie = ""
    Kategorie = ""
    AnzZeilen = FilternListe(Liste_Spalten_XlTab, Liste_gefiltert, Kategorie)
  End If
  
  If (AnzZeilen > 0) Then
    Me.cboZiel.List = Liste_gefiltert
    
    'Vorauswahl eines Eintrages.
    idxAuswahl = 0
    idxLetzteWahl = -1
    For i = LBound(Me.cboZiel.List, 1) To UBound(Me.cboZiel.List, 1)
      ZielSpalteKey = MidStr(Me.cboZiel.List(i, idxZiel_Adresse), "$", "$", False)
      If (ZielSpalteKey = Status.LetzerZielSpalteKey) Then idxLetzteWahl = i
    Next
    If (idxLetzteWahl > -1) Then idxAuswahl = idxLetzteWahl
    
    Me.cboZiel.ListIndex = idxAuswahl
  Else
    Me.cboZiel.Clear
  End If
  
  Exit Sub
  
Fehler:
  Me.cboZiel.Clear
  Err.Clear
  'FehlerNachricht "frmSpaltenVerw.GetListe_ZielBereiche()"
End Sub



Private Function GetSpaltenNr(ByVal SpaltenName As String) As Long
  'Gibt die Nummer der Spalte zurück, die den "Spaltenname" trägt.
  On Error GoTo Fehler
  GetSpaltenNr = 0
  If (oAktiveTabelle.SpaltenErsteZellen.Exists(SpaltenName)) Then
    GetSpaltenNr = oAktiveTabelle.SpaltenErsteZellen(SpaltenName).Column
  End If
  Exit Function
  
Fehler:
  FehlerNachricht "frmSpaltenVerw.GetSpaltenNr()"
End Function



Public Function Check_DialogAktion() As Boolean
  'True, wenn eine Aktion ausgeführt werden kann.
  'Der Status des "Aktion"-Buttons wird geschaltet und dessen Beschreibung gesetzt.
  
  On Error GoTo Fehler
  
  Const strLoeschen   As String = "Löschen"
  Const strAendern    As String = "Ersetzen"
  Const strEinfuegen  As String = "Einfügen"
  Const strFehler     As String = "Fehler"
  
  Dim blnGueltig      As Boolean
  Dim blnEleVorh      As Boolean
  Dim i               As Long
  Dim strDummy        As String
  
  'MsgBox "Check_DialogAktion"
  
  'Status von Me.fraZiel.Enabled wird bereits durch cboZiel_Change() festgelegt.
  If (Not Me.fraZiel.Enabled) Then
    blnGueltig = False
    Me.btnAktion.Caption = strEinfuegen
    Me.btnAktion.Enabled = False
  Else
    If ((Me.cboKategorie.ListIndex < 0) Or (Me.LstSpaltennamen.ListIndex < 0) Or _
        (Me.cboZiel.ListIndex < 0)) Then
      'Sollte nicht vorkommen.
      blnGueltig = False
      Me.btnAktion.Caption = strFehler
      Me.btnAktion.Enabled = False
    Else
      blnGueltig = True
      Me.btnAktion.Enabled = True
      
      If (Status.QuellName = strOhneQuellName) Then
        Me.btnAktion.Caption = strLoeschen
      ElseIf (Status.Kategorie = strKategorie_TabStruktur) Then
        'Strukturelement einfügen/ändern.
        blnEleVorh = False
        If (Status.QuellName = strInfoTraeger) Then
          If (Not oAktiveTabelle.Infotraeger Is Nothing) Then blnEleVorh = True
        ElseIf (Status.QuellName = strFliesskomma) Then
          If (Not oAktiveTabelle.Fliesskomma Is Nothing) Then blnEleVorh = True
        ElseIf (Status.QuellName = strFormel) Then
          If (Not oAktiveTabelle.Formel Is Nothing) Then blnEleVorh = True
        End If
        If (blnEleVorh) Then Me.btnAktion.Caption = strAendern Else Me.btnAktion.Caption = strEinfuegen
      ElseIf ((Status.Kategorie = strKategorie_PrjDat) or (Status.Kategorie = strKategorie_OrtDat)) Then
        'Orts-/Projektdatenfeld einfügen/ändern.
        If (GetLokalerZellname(Status.QuellName) Is Nothing) Then
          On Error Resume Next
          strDummy = ActiveCell.Name.Name
          If (Err.Number = 0) Then
            'Zielzelle hat bereits mindestens einen Namen.
            Me.btnAktion.Caption = strAendern
          Else
            Me.btnAktion.Caption = strEinfuegen
          End If
          Err.Clear
          On Error GoTo Fehler
        Else
          Me.btnAktion.Caption = strAendern
        End If
      Else
        'Spaltenname einfügen/ändern/löschen.
        'MsgBox "QuellName=" & Status.QuellName & "    Zielname=" & Status.ZielName
        Me.btnAktion.Caption = strAendern
        If (Status.ZielName = strOhneQuellName) Then
          blnEleVorh = False
          For i = LBound(Liste_Spalten_XlTab, 1) To UBound(Liste_Spalten_XlTab, 1)
            If (Liste_Spalten_XlTab(i, idxSpName) = Status.QuellName) Then
              blnEleVorh = True
              Exit For
            End If
          Next
          If (Not blnEleVorh) Then Me.btnAktion.Caption = strEinfuegen
        End If
      End If
      
    End If
  End If
  
  'Hilfe-Button: einschalten, wenn Hilfetext verfügbar.
  If (strHilfeAktiv <> "") Then
    btnHilfe.Visible = True
  Else
    btnHilfe.Visible = False
  End If
  
  Check_DialogAktion = blnGueltig
  
  Exit Function
  
Fehler:
  Check_DialogAktion = False
  FehlerNachricht "frmSpaltenVerw.Check_DialogAktion()"
End Function


Private Sub Syncronisieren(Optional ByVal blnkeinAktivesBlatt As Boolean = False)
  'Liest die Spalten der aktiven Tabelle neu ein und leitet alle nötigen Folgemaßnahmen ein.
  'Parameter "keinAktivesBlatt" ist true, wenn gerade die letzte Arbeitsmappe geschlossen wird.
  On Error GoTo Fehler
  keinAktivesBlatt = blnkeinAktivesBlatt
  Call GetListe_Spalten_XlTab
  Call GetListe_ZielBereiche
  Call GetListe_QuellNamen
  Exit Sub
  
Fehler:
  FehlerNachricht "frmSpaltenVerw.Syncronisieren()"
End Sub



Private Function Selection_Korrektur() As String
  'Wenn gerade eine Spalte gewählt werden soll, so wird sichergestellt, daß sich die
  'aktive Zellauswahl innerhalb des Infoträgers befindet und genau eine Spalte breit ist.
  'Funktionswert = absolute Adresse der erzeugten Zellauswahl, sonst "".
  On Error GoTo Fehler
  
  Dim oInfotraeger          As Range
  Dim ointersect            As Range
  
  If ((Status.Kategorie = strKategorie_PrjDat) or (Status.Kategorie = strKategorie_OrtDat)) Then
    ActiveCell.Select
  Else
    If (Not oAktiveTabelle.Infotraeger Is Nothing) Then
      Set oInfotraeger = oAktiveTabelle.Infotraeger
      Set ointersect = Intersect(ActiveCell.EntireColumn, oInfotraeger)
      If (ointersect Is Nothing) Then
        'Die aktive Markierung liegt in einer Spalte außerhalb des Infoträgers.
        If (ActiveCell.Column < oInfotraeger.Column) Then
          'Erste Spalte des Infotraegers markieren.
          Set ointersect = Intersect(oInfotraeger.EntireRow, oInfotraeger.EntireColumn.Resize(columnsize:=1))
        Else
          'Letzte Spalte des Infotraegers markieren.
          Set ointersect = Intersect(oInfotraeger.EntireRow, oInfotraeger.EntireColumn.Offset(columnoffset:=oInfotraeger.Columns.Count - 1).Resize(columnsize:=1))
        End If
      End If
      If (Not ointersect Is Nothing) Then
        ointersect.Select
        Selection_Korrektur = ointersect.Address
      Else
        Selection_Korrektur = ""
      End If
    End If
  End If
  
  Exit Function
  
Fehler:
  Selection_Korrektur = ""
  FehlerNachricht "frmSpaltenVerw.Selection_Korrektur()"
End Function
  


Private Sub WaehleListenEintrag(ByVal AdresseAbsolut As String)
  'Wählt in der Kombobox "Ziel" den Eintrag aus, der mit "Adresse" korrespondiert.
  On Error GoTo Fehler
  
  Dim ZielSpalteKey    As String
  Dim ListSpalteKey    As String
  Dim i                As Long
  
  ZielSpalteKey = MidStr(AdresseAbsolut, "$", "$", False)
  
  For i = 0 To Me.cboZiel.ListCount - 1
    ListSpalteKey = MidStr(Me.cboZiel.List(i, idxZiel_Adresse), "$", "$", False)
    If (ListSpalteKey = ZielSpalteKey) Then
      Me.cboZiel.ListIndex = i
      Exit For
    End If
  Next
  
  Exit Sub
  
Fehler:
  FehlerNachricht "frmSpaltenVerw.WaehleListenEintrag()"
End Sub
  


Private Function ZielZelleMarkieren() As String
  'Gewählte Spalte im Infoträger selektieren.
  On Error GoTo Fehler
  
  Dim oInfotraeger          As Range
  Dim ointersect            As Range
  
  If (Not oAktiveTabelle.Infotraeger Is Nothing) Then
    Set oInfotraeger = oAktiveTabelle.Infotraeger
    Set ointersect = Intersect(Range(Status.ZielSpalteKey & ":" & Status.ZielSpalteKey).EntireColumn, oInfotraeger)
    If (Not ointersect Is Nothing) Then
      ointersect.Select
      ZielZelleMarkieren = ointersect.Address
    Else
      ZielZelleMarkieren = ""
    End If
  End If
  
  Exit Function
  
Fehler:
  ZielZelleMarkieren = ""
  'FehlerNachricht "frmSpaltenVerw.ZielZelleMarkieren()"
  Err.Clear
End Function



Sub ZeigeEinstellungen()
  'Zeigt für Kontrollzwecke alle aktiven Einstellungen des Dialoges an.
  Dim Message As String
    
  Message = Message & vbNewLine & "Kategorie: " & vbTab & Status.Kategorie
  Message = Message & vbNewLine
  Message = Message & vbNewLine & "Quelle Status: " & vbTab & Status.WertStatus
  Message = Message & vbNewLine & "Quelle Name: " & vbTab & Status.QuellName
  Message = Message & vbNewLine & "Quelle SpalteNr: " & vbTab & Status.QuellSpalteNr
  Message = Message & vbNewLine & "Quelle Einheit: " & vbTab & Status.QuellEinheit
  Message = Message & vbNewLine
  Message = Message & vbNewLine & "Ziel Spalte: " & vbTab & Status.ZielSpalteNr & "  (" & Status.ZielSpalteKey & ")"
  Message = Message & vbNewLine & "Ziel Name: " & vbTab & Status.ZielName
  Message = Message & vbNewLine
  
  MsgBox Message
End Sub


' Für jEdit:  :collapseFolds=1:
