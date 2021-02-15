VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CdatKonfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2003-2021  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'====================================================================================
' Klasse CdatKonfig
'====================================================================================
' Stellt die Konfiguration zur Verfügung.
' Wenn möglich, wird GeoTools_cfg.xlsx gelesen,
' wodurch einige Standardwerte überschrieben werden.


Option Explicit

'Konstanten
Const strKlasse_Name             As String = "CdatKonfig"
                                 
'Einstellungen.
Const strTabNameEinstell         As String = "Einstellungen"
Const strSpalteVariable          As String = "Variable"
Const strSpalteWert              As String = "Wert"
                                 
'Spaltenkonfiguration.
Const strTabNameSpalten          As String = "SpaltenKonfig"
Const strSpName                  As String = "SpaltenName"
Const strSpBeschreibung          As String = "Beschreibung"
Const strSpMathGroesse           As String = "Groesse"
Const strSpKategorie             As String = "KategorieName"
Const strSpErsatzZiel            As String = "ErsatzZiel"
Const strSpKatAlsFilter          As String = "KategorieAlsFilter"


'Variablennamen (Name in CFG-Tabelle = Keys des Dictionary)
Const Var_Anwender               As String = "Anwender"
Const Var_Ort_Dateiname          As String = "Ort_Dateiname"
Const Var_StdArbeitsverz         As String = "StdArbeitsverz"
Const Var_JEDIT_HOME             As String = "JEDIT_HOME"
Const Var_Expim_Modifizieren     As String = "Expim_Modifizieren"
Const Var_Expim_Ueberschreiben   As String = "Expim_Ueberschreiben"
Const Var_Expim_FormelnErhalten  As String = "Expim_FormelnErhalten"
Const Var_Expim_ErsatzSpalten    As String = "Expim_ErsatzSpalten"
Const Var_Expim_SchlussMeldung   As String = "Expim_SchlussMeldung"
Const Var_TkBasisUeberhoehung    As String = "TkBasisUeberhoehung"
Const Var_UebAusInfoStreng       As String = "UebAusInfoStreng"

'Fehler.
Dim ErrMsgCfgFehlt               As String
Dim ErrMsgTabEinst               As String
Dim ErrMsgTabSpalten             As String
Dim ErrMsgEinheitUnbekannt       As String
                                
Dim ErrNumEinheitUnbekannt       As Long
Dim ErrNumCfgFehlt               As Long
Dim ErrNumTabEinst               As Long
Dim ErrNumTabSpalten             As Long
                                
'..
Private cfg                     As String
Private strInfoKeineKonfig      As String
Private blnKonfigVerfuegbar     As Boolean
Private blnKonfigInitFertig     As Boolean
                                
Dim oEinstellungen              As Scripting.Dictionary
Dim oSpaltenBeschreibung        As Scripting.Dictionary
Dim oSpaltenMathGroesse         As Scripting.Dictionary
Dim oSpaltenKategorie           As Scripting.Dictionary
Dim oSpaltenErsatzZiel          As Scripting.Dictionary
Dim oKategorieAlsFilter         As Scripting.Dictionary

Dim oStatusPrefix               As Scripting.Dictionary

Dim oEinheiten                  As Scripting.Dictionary
Dim oLaenge                     As Scripting.Dictionary
Dim oWinkel                     As Scripting.Dictionary

Dim oCache                      As Scripting.Dictionary
'


Private Sub Class_Initialize()
  'Lesen der Konfiguration.
  
  Dim pi        As Double
  Dim rho       As Double
  
  'On Error GoTo Fehler
  On Error GoTo 0
  DebugEcho "CdatKonfig.Class_Initialize(): Initialisierung beginnt."
  
  'Name der CFG_Datei = GeoTools_cfg.xlsx (im Ressource-Unterverzeichnis).
  'cfg = Verz(ThisWorkbook.Path) & "\" & VorName(ThisWorkbook.Name) & "_cfg.xlsx"
  cfg = ThisWorkbook.Path & "\" & ResourcesSubFolder & "\" & VorName(ThisWorkbook.Name) & "_cfg.xlsx"
  
  ' Meldung für den Fall, dass die Datei nicht gelesen wurde.
  strInfoKeineKonfig = "Konfigurationsdatei '" & cfg & "' wurde nicht gelesen." & vbLf & vbLf & _
                       "Mögliche Ursachen:." & vbLf & _
                       "  1. Die Datei existiert nicht." & vbLf & _
                       "  2. Excel wurde ferngesteuert gestartet." & vbLf & vbLf & _
                       "==> Die Funktionalität des Programmes steht dadurch nur eingeschränkt zur Verfügung."
  
  'Mathematische Konstanten
  pi = 4 * Atn(1)
  rho = 200 / pi
  
  
  'Fehler
  ErrNumCfgFehlt = 50401 - vbObjectError
  ErrNumTabEinst = 50402 - vbObjectError
  ErrNumTabSpalten = 50403 - vbObjectError
  ErrNumEinheitUnbekannt = 50404 - vbObjectError
  
  ErrMsgEinheitUnbekannt = "Programm- oder Konfigurationsfehler: " & vbNewLine & _
                           "Die verwendete Einheit ist nicht vereinbart." & vbNewLine & _
                           "==> Das Programm wird abgebrochen oder es wird die Standardeinheit verwendet!"
  ErrMsgCfgFehlt = "Die Funktionalität des Programmes steht dadurch nur eingeschränkt zur Verfügung."
  ErrMsgTabEinst = "Die Funktionalität des Programmes steht dadurch nur eingeschränkt zur Verfügung."
  ErrMsgTabSpalten = "Die Funktionalität des Programmes steht dadurch nur eingeschränkt zur Verfügung."
  
  Set oEinstellungen = New Scripting.Dictionary
  Set oSpaltenBeschreibung = New Scripting.Dictionary
  Set oSpaltenMathGroesse = New Scripting.Dictionary
  Set oSpaltenKategorie = New Scripting.Dictionary
  Set oSpaltenErsatzZiel = New Scripting.Dictionary
  Set oKategorieAlsFilter = New Scripting.Dictionary
  Set oStatusPrefix = New Scripting.Dictionary
  Set oEinheiten = New Scripting.Dictionary
  Set oLaenge = New Scripting.Dictionary
  Set oWinkel = New Scripting.Dictionary
  Set oCache = New Scripting.Dictionary
  
  ' Hartkodierte "Standard"-Einstellungen.
  oEinstellungen.Add Var_Anwender, "GeoTools"
  oEinstellungen.Add Var_Ort_Dateiname, ""
  oEinstellungen.Add Var_StdArbeitsverz, ""
  oEinstellungen.Add Var_JEDIT_HOME, ""
  oEinstellungen.Add Var_Expim_Modifizieren, Bool2String(Std_DatenModifizieren)
  oEinstellungen.Add Var_Expim_Ueberschreiben, Bool2String(Std_VorhWerteUeberschreiben)
  oEinstellungen.Add Var_Expim_FormelnErhalten, Bool2String(Std_FormelnErhalten)
  oEinstellungen.Add Var_Expim_ErsatzSpalten, Bool2String(Std_ErsatzZielspaltenVerwenden)
  oEinstellungen.Add Var_Expim_SchlussMeldung, Bool2String(Std_ExpimSchlussMeldung)
  oEinstellungen.Add Var_TkBasisUeberhoehung, cStr(Std_TkBasisUeberhoehung)
  oEinstellungen.Add Var_UebAusInfoStreng, Bool2String(Std_UebAusInfoStreng)
  
  'Wert-Status, gilt für alle Spalten mit konfigurierter Einheit, d.h. numerische Werte.
  oStatusPrefix.Add StatusBez_Ist, ""
  oStatusPrefix.Add StatusBez_Soll, "S."
  oStatusPrefix.Add StatusBez_Fehler, "F."
  oStatusPrefix.Add StatusBez_Verbesserung, "V."
  
  'Unterstützte Einheiten.
  oLaenge.Add "m" , 1#
  oLaenge.Add "dm", 10
  oLaenge.Add "cm", 100
  oLaenge.Add "mm", 1000
  oLaenge.Add "km", 0.001
  
  oWinkel.Add "gon" , 1#
  oWinkel.Add "grad", 0.9
  oWinkel.Add "rad" , 1 / rho
  oWinkel.Add "pph" , tan(1 / rho) * 100      'Prozent
  oWinkel.Add "ppt" , tan(1 / rho) * 1000     'Promille
  oWinkel.Add "ppm" , tan(1 / rho) * 1000000  'Parts per million

  oEinheiten.Add "Laenge", oLaenge
  oEinheiten.Add "Winkel", oWinkel
  
  'Fürs Debug-Protokoll
  Call ZeigeAlleEinheiten

  blnKonfigVerfuegbar = False
  blnKonfigInitFertig = False
  ''Lesen der Konfigurations-Mappe.
  'Call LeseKonfiguration
  'Call ClearStatusBarDelayed(StatusBarClearDelay)
  DebugEcho "CdatKonfig.Class_Initialize(): GeoTools-Konfiguration initialisiert." & vbNewLine
  Exit Sub

Fehler:
  ErrMessage = "Fehler beim Initialisieren der Konfiguration"
  ErrEcho "CdatKonfig.Class_Initialize()"
End Sub


Private Sub Class_Terminate()
  Set oEinstellungen = Nothing
  Set oSpaltenBeschreibung = Nothing
  Set oSpaltenMathGroesse = Nothing
  Set oSpaltenKategorie = Nothing
  Set oSpaltenErsatzZiel = Nothing
  Set oKategorieAlsFilter = Nothing
  Set oStatusPrefix = Nothing
  Set oLaenge = Nothing
  Set oWinkel = Nothing
  Set oEinheiten = Nothing
  Set oCache = Nothing
End Sub




'==== Eigenschaften ===============================================================================

Property Get Klasse_Name() As String
  'liefert den Namen der Objektklasse
  Klasse_Name = strKlasse_Name
End Property

Property Get Anwender() As String
  'Liefert die Kennung des Anwenders zwecks Anpassungen.
  Anwender = oEinstellungen(Var_Anwender)
End Property

Property Get Ort_Dateiname() As String
  'Liefert Pfad\Namen der Ini-Datei mit ortsbezogenen Informationen (einer Niederlassung o.ä.).
  Ort_Dateiname = oEinstellungen(Var_Ort_Dateiname)
End Property

Property Get StdArbeitsverz() As String
  'Liefert Standard-Arbeitsverzeichnis (für Dateidialoge, falls in den Excel-Optionen kein existierender "Standardarbeitsordner" angegeben ist.)
  StdArbeitsverz = oEinstellungen(Var_StdArbeitsverz)
End Property

Property Get JEDIT_HOME() As String
  'Liefert das jEdit-Programmverzeichnis.
  JEDIT_HOME = oEinstellungen(Var_JEDIT_HOME)
End Property

Property Get Expim_Modifizieren() As String
  'Liefert Startwert für Expim-Dialog.
  Expim_Modifizieren = oEinstellungen(Var_Expim_Modifizieren)
End Property

Property Get Expim_Ueberschreiben() As Boolean
  'Liefert Startwert für Expim-Dialog.
  Expim_Ueberschreiben = String2Bool(oEinstellungen(Var_Expim_Ueberschreiben))
End Property

Property Get Expim_FormelnErhalten() As String
  'Liefert Startwert für Expim-Dialog.
  Expim_FormelnErhalten = oEinstellungen(Var_Expim_FormelnErhalten)
End Property

Property Get Expim_ErsatzSpalten() As String
  'Liefert Startwert für Expim-Dialog.
  Expim_ErsatzSpalten = oEinstellungen(Var_Expim_ErsatzSpalten)
End Property

Property Get Expim_SchlussMeldung() As String
  'Liefert Einstellung: "Meldung am Ende einer Import/Export-Operation ausgeben?".
  Expim_SchlussMeldung = oEinstellungen(Var_Expim_SchlussMeldung)
End Property


Property Get TkBasisUeberhoehung() As Double
  'Liefert "Basis für Überhöhung" in [m] als Vorgabe für Transfo' Trassensystem <=> Gleissystem.
  dim strWert
  dim dblWert
  ' IsNumeric() und cDbl() arbeiten mit Komma als Dezimaltrenner immer (unabhängig von Einstellungen in System und Excel):
  strWert = replace(oEinstellungen(Var_TkBasisUeberhoehung), ".", ",")
  if (IsNumeric(strWert)) then dblWert = cDbl(strWert) else dblWert = Std_TkBasisUeberhoehung
  TkBasisUeberhoehung = dblWert
End Property

Property Get UebAusInfoStreng() As Boolean
  'Liefert Einstellung, ob die Angabe der Ist-Überhöhung in der Punktinfo'
  'in der Form u=xxx vorliegen muss oder ob anderenfalls auch die erste Zahl verwendet wird.
  UebAusInfoStreng = String2Bool(oEinstellungen(Var_UebAusInfoStreng))
End Property


Property Get SpaltenBeschreibung() As Scripting.Dictionary
  'Liefert die Beschreibung jeder konfigurierten Spalte.
  Set SpaltenBeschreibung = oSpaltenBeschreibung
End Property

Property Get SpaltenMathGroesse() As Scripting.Dictionary
  'Liefert die mathematische Größe jeder konfigurierten Spalte (Laenge oder Winkel).
  Set SpaltenMathGroesse = oSpaltenMathGroesse
End Property

Property Get SpaltenKategorie() As Scripting.Dictionary
  'Liefert die Kategorie jeder konfigurierten Spalte.
  Set SpaltenKategorie = oSpaltenKategorie
End Property

Property Get SpaltenErsatzZiel() As Scripting.Dictionary
  'Liefert für jede konfigurierte Spalte den Namen der als ErsatzZiel dienenden Spalte.
  Set SpaltenErsatzZiel = oSpaltenErsatzZiel
End Property

Property Get Kategorien() As Scripting.Dictionary
  'Liefert alle konfigurierten Kategorien (key=KategorieName, item=boolean als Filter).
  Set Kategorien = oKategorieAlsFilter
End Property

Property Get KategorieAlsFilter(ByVal inpKategorieName As String) As Boolean
  'Liefert den Schalter, der festlegt, ob die angegebene Kategorie
  'als Filterkriterium für die Formatliste des Dialoges dient.
  If (oKategorieAlsFilter.Exists(inpKategorieName)) Then
    KategorieAlsFilter = oKategorieAlsFilter(inpKategorieName)
  Else
    KategorieAlsFilter = True
  End If
End Property


Property Get StatusPrefix() As Scripting.Dictionary
  'Liefert für jeden unterstützten Werte-Status den Prefix für Spaltenname.
  'Dictionary: Key=Bezeichnung, Item=Prefix für Spaltenname.
  Set StatusPrefix = oStatusPrefix
End Property


Property Get Einheiten() As Scripting.Dictionary
  'Liefert ein Dictionary, das für jede unterstützte Größe ein Dictionary
  'mit allen unterstützten Einheiten enthält.
  Set Einheiten = oEinheiten
End Property


Property Get Cache() As Scripting.Dictionary
  'Liefert den Cache. Dieser ist ein Dictionary, das "alles" enthält,
  'was hinein gestopft wird :-)
  Set Cache = oCache
End Property


Property Get KonfigInitFertig() As Boolean
  'Liefert die Option "Initialisierung der Konfiguration fertig (mit oder ohne Erfolg)".
  KonfigInitFertig = blnKonfigInitFertig
End Property

Property Get KonfigVerfuegbar() As Boolean
  'Liefert die Option "Konfiguration verfügbar".
  KonfigVerfuegbar = blnKonfigVerfuegbar
End Property

Property Get InfoKeineKonfig() As String
  InfoKeineKonfig = strInfoKeineKonfig
End Property



Property Get SpNameAttr(ByVal inpSpName As String) As Scripting.Dictionary
  'Liefert alle Attribute eines gegebenen Spaltennamens.
  'Parameter: inpSpName ... Spaltenname mit/ohne Prefix(e) und Einheit.
  'Rückgabe: Dictionary mit den bei "Standardwerten" festgelegten Keys.
  '=> Ist für den Spaltennamen keine mathematischen Größe konfiguriert,
  '   so wird StatusBez="" gesetzt.
  
  On Error GoTo Fehler
  
  Dim WertStatus         As Variant
  Dim Prefix             As String
  Dim Einheit            As String
  Dim Kategorie          As String
  Dim SpName_1           As String   'Ohne Prefix "Spalte."
  Dim SpName_2           As String   'Ohne Einheit
  Dim SpName_3           As String   'Ohne Wertstatus (Soll, Ist..)
  Dim NF                 As Long
  Dim Feld()             As String
  Dim oAttribute         As Scripting.Dictionary

  Set oAttribute = New Scripting.Dictionary
  
  'Standardwerte.
  oAttribute.Add "StatusPrefix", ""
  oAttribute.Add "StatusBez", "Ist"
  oAttribute.Add "NamePur", SpName_unbekannt
  oAttribute.Add "Titel", SpTitel_unbekannt
  oAttribute.Add "Kategorie", Allg_unbekannt
  oAttribute.Add "MathGroesse", Allg_unbekannt
  oAttribute.Add "Einheit", Allg_unbekannt
  oAttribute.Add "EinheitFaktor", 0
  oAttribute.Add "Liste", "< Fehler >"

  
  'Evtl. vorhandenen Kenner für "Spaltenname" entfernen.
  If (Left$(inpSpName, Len(PrefixSpaltenname)) = PrefixSpaltenname) Then
    SpName_1 = Right(inpSpName, Len(inpSpName) - Len(PrefixSpaltenname))
  Else
    SpName_1 = inpSpName
  End If

  'Einheit extrahieren.
  NF = SplitDelim(SpName_1, Feld, TrennerEinheit)
  If (NF > 1) Then
    SpName_2 = Feld(1)
    Einheit = Feld(NF)
  Else
    SpName_2 = SpName_1
    Einheit = ""
  End If
  oAttribute("Einheit") = Einheit
  oAttribute("EinheitFaktor") = Me.FaktorDerEinheit(Einheit)
  
  'Wertstatus ermitteln.
  SpName_3 = SpName_2
  If (Me.StatusPrefix.Count > 0) Then
    For Each WertStatus In Me.StatusPrefix.Keys
      Prefix = Me.StatusPrefix(WertStatus)
      If (Prefix <> "") Then
        If (Left$(SpName_2, Len(Prefix)) = Prefix) Then
          oAttribute("StatusPrefix") = Prefix
          oAttribute("StatusBez") = WertStatus
          SpName_3 = Right(SpName_2, Len(SpName_2) - Len(Prefix))
        End If
      End If
    Next
  End If

  'Math. Größe der Spalte laut Konfiguration ermitteln.
  If (Me.SpaltenMathGroesse.Count > 0) Then
    If (Me.SpaltenMathGroesse.Exists(SpName_3)) Then
      oAttribute("MathGroesse") = Me.SpaltenMathGroesse(SpName_3)
    End If
  End If
  If (oAttribute("MathGroesse") = "") Then
    'Keine math. Größe => Status zurücksetzen.
    SpName_3 = SpName_2
    oAttribute("StatusPrefix") = ""
    oAttribute("StatusBez") = ""
  End If
  oAttribute("NamePur") = SpName_3

  'Kategorie ermitteln.
  If (Me.SpaltenKategorie.Count > 0) Then
    If (Me.SpaltenKategorie.Exists(SpName_3)) Then
      oAttribute("Kategorie") = Me.SpaltenKategorie(SpName_3)
    End If
  End If

  'SpaltenBeschreibung ermitteln.
  If (Me.SpaltenBeschreibung.Count > 0) Then
    If (Me.SpaltenBeschreibung.Exists(SpName_3)) Then
      oAttribute("Titel") = Me.SpaltenBeschreibung(SpName_3)
    End If
  End If
  
  'Liste zwecks Anzeige/Kontrolle
  oAttribute("Liste") = "übergebener Spaltenname: " & inpSpName & vbNewLine & vbNewLine & _
                        "Kategorie: " & vbTab & oAttribute("Kategorie") & vbNewLine & _
                        "Titel: " & vbTab & vbTab & oAttribute("Titel") & vbNewLine & _
                        "StatusBez: " & vbTab & oAttribute("StatusBez") & vbNewLine & _
                        "StatusPrefix: " & vbTab & oAttribute("StatusPrefix") & vbNewLine & _
                        "NamePur: " & vbTab & oAttribute("NamePur") & vbNewLine & _
                        "Einheit: " & vbTab & vbTab & oAttribute("Einheit") & vbNewLine & _
                        "EinheitFaktor: " & vbTab & oAttribute("EinheitFaktor") & vbNewLine & _
                        "MathGroesse: " & vbTab & oAttribute("MathGroesse")

  Set SpNameAttr = oAttribute
  Set oAttribute = Nothing
  Exit Property
  
Fehler:
  'set SpNameAttr = nothing
  'set oAttribute = nothing
  ErrEcho "CdatKonfig.SpNameAttr()"
End Property



Property Get FaktorDerEinheit(NameDerEinheit As String) As Double
  'Gibt den Umrechnungsfaktor von der Standardeinheit zu der angegebenen Einheit zurück.
  'Eingabeparameter: NameDerEinheit = existierender Key eines Dictionary in me.einheiten.
  '                                   wenn="ohne", "" ==> Wert ohne Einheit
  'Funktionswert:    passender Item aus me.einheiten, falls verfügbar; sonst=0.

  On Error GoTo Fehler

  Dim Groesse As Variant
  Dim Faktor  As Double

  If ((NameDerEinheit = "ohne") Or (NameDerEinheit = "")) Then
    Faktor = 0
  Else
    Faktor = 999999999
    For Each Groesse In Me.Einheiten
      If (Me.Einheiten(Groesse).Exists(NameDerEinheit)) Then
        Faktor = Me.Einheiten(Groesse)(NameDerEinheit)
        Exit For
      End If
    Next
    If (Faktor = 999999999) Then
      ErrMessage = ErrMsgEinheitUnbekannt
      Err.Raise vbObjectError + ErrNumEinheitUnbekannt, , "Unbekannte Einheit: '" & NameDerEinheit & "'"
    End If
  End If

  'MsgBox "NameDerEinheit=" & NameDerEinheit & ",   Faktor=" & Faktor & vbNewLine
  FaktorDerEinheit = Faktor
  
  Exit Property

Fehler:
  FaktorDerEinheit = 0
  ErrEcho "CdatKonfig.FaktorDerEinheit()"
End Property



'==== Methoden ====================================================================================

Function EinheitExists(NameDerEinheit As String) As Boolean
  'Funktionswert:  true, falls Einheit verfügbar, sonst false.
  On Error GoTo Fehler
  
  Dim Groesse    As Variant
  Dim Existiert  As Boolean
  
  If ((NameDerEinheit = "ohne") Or (NameDerEinheit = "")) Then
    Existiert = true
  Else
    For Each Groesse In Me.Einheiten
      If (Me.Einheiten(Groesse).Exists(NameDerEinheit)) Then
        Existiert = true
        Exit For
      End If
    Next
  End If
  EinheitExists = Existiert
  Exit Function
  
Fehler:
  EinheitExists = false
  ErrEcho "CdatKonfig.EinheitExists()"
End Function



'==== interne Routinen ============================================================================

Public Sub LeseKonfiguration()
  'Lesen der Konfigurations-Arbeitsmappe.
  
  Dim StatusScreen     As Boolean
  Dim StatusEvents     As Boolean
  Dim StatusAlerts     As Boolean
  Dim StatusCalc       As Boolean
  Dim Success          As Boolean
  Dim oWbkCfg          As Excel.Workbook
  Dim oXlApp2          As Excel.Application
  
  On Error GoTo Fehler
  DebugEcho "CdatKonfig.LeseKonfiguration(): Lese Konfigurationsdatei '" & cfg & "'."
  
  'StatusScreen = Application.ScreenUpdating
  'StatusEvents = Application.EnableEvents
  'StatusAlerts = Application.DisplayAlerts
  'Application.ScreenUpdating = False
  'Application.DisplayAlerts  = False
  
  ' Konfigurationsdatei öffnen mit zweiter Excel-Instanz.
  On Error GoTo Fehler
  Set oXlApp2 = New Excel.Application
  ErrMessage = "Fehler beim Öffnen der Konfigurationsdatei"
  Set oWbkCfg = oXlApp2.Workbooks.Open(FileName:=cfg, ReadOnly:=True, UpdateLinks:=0 , AddToMru:=False) 
  If (oWbkCfg is nothing) Then
    On Error GoTo Fehler
    Err.Raise ErrNumCfgFehlt + vbObjectError, , ThisWorkbook.Name & ": Konfigurationsdatei '" & cfg & "' nicht gefunden."
  End If
  
    
  ' Konfigurationsdatei lesen (zwei Tabellen).
  Success = LeseEinstellungen(oWbkCfg)
  Success = (Success And LeseSpaltenKonfig(oWbkCfg))
  
  ' Konfigurationsdatei und zweite Excel-Instanz schließen.
  oWbkCfg.Close False
  oXlApp2.Quit
  Set oWbkCfg = nothing
  Set oXlApp2 = nothing
  
  ' Status-Signalisierung.
  blnKonfigInitFertig = True
  blnKonfigVerfuegbar = Success
  DebugEcho "CdatKonfig.LeseKonfiguration(): GeoTools-Konfiguration gelesen. (" & cfg & ")."
  
  'Application.EnableEvents   = StatusEvents
  'Application.ScreenUpdating = StatusScreen
  'Application.DisplayAlerts  = StatusAlerts
  
  Application.StatusBar = "GeoTools-Konfiguration gelesen. (" & cfg & ")."
  Call ClearStatusBarDelayed(StatusBarClearDelay)
  Exit Sub
  
Fehler:
  Application.ScreenUpdating = StatusScreen
  'ErrEcho "CdatKonfig.LeseKonfiguration()"
  ErrEcho "CdatKonfig.LeseKonfiguration(): *** mit Fehler beendet ***."
End Sub



Private Function LeseEinstellungen(oWbkCfg As Excel.Workbook) As Boolean
  ' Liest die allgemeinen Einstellungen aus der Tabelle "Einstellungen".
  
  Dim lb              As Long
  Dim ub              As Long
  Dim i               As Long
  Dim Anzahl          As Long
  Dim Spalte          As Variant
  Dim oExpim          As CdatExpim
  Dim oSheetSettings  As Excel.Worksheet
  
  On Error Resume Next
  DebugEcho "CdatKonfig.LeseEinstellungen(): Lese allgemeine Einstellungen aus Tabelle '" & strTabNameEinstell & "'."
  
  ' Tabelle finden.
  Set oSheetSettings = oWbkCfg.Sheets(strTabNameEinstell)
  If (Err.Number <> 0) Then
    On Error GoTo Fehler
    ErrMessage = ErrMsgTabEinst
    Err.Raise ErrNumTabEinst + vbObjectError, , "Tabelle '" & strTabNameEinstell & "' in Konfigurationsdatei '" & cfg & "' nicht gefunden."
  End If
  
  ' Tabelle lesen und Daten abspeichern.
  Set oExpim = New CdatExpim
  Anzahl = oExpim.GetQuelldaten_XlTab(oSheetSettings)
  lb = LBound(oExpim.Datenpuffer.Daten(strSpalteVariable), 1)
  ub = UBound(oExpim.Datenpuffer.Daten(strSpalteVariable), 1)
  For i = lb To ub
    oEinstellungen(oExpim.Datenpuffer.Daten(strSpalteVariable)(i, DPidxWert)) = oExpim.Datenpuffer.Daten(strSpalteWert)(i, DPidxWert)
  Next
  Set oExpim = Nothing
  Set oSheetSettings = nothing
  
  DebugEcho "CdatKonfig.LeseEinstellungen(): Folgende allgemeine Einstellungen wurden gelesen:"
  DebugEcho ListeDictionary(oEinstellungen)
  
  LeseEinstellungen = True
  On Error GoTo 0
  Exit Function

Fehler:
  Set oSheetSettings = nothing
  ErrEcho "CdatKonfig.LeseEinstellungen()"
End Function



Private Function LeseSpaltenKonfig(oWbkCfg As Excel.Workbook)
  'Liest die Spalten-Konfiguration aus der Tabelle "SpaltenKonfig".
  
  Dim lb                    As Long
  Dim ub                    As Long
  Dim i                     As Long
  Dim Anzahl                As Long
  Dim Spalte                As String
  Dim Kategorie             As String
  Dim ErsatzZiel            As String
  Dim KategorieAlsFilter    As String
  Dim bool                  As Boolean
  Dim oExpim                As CdatExpim
  Dim oSheetColumnConfig    As Excel.Worksheet
  
  DebugEcho "CdatKonfig.LeseSpaltenKonfig(): Lese Spalten-Konfiguration aus Tabelle '" & strTabNameSpalten & "'."
  On Error Resume Next
  
  ' Tabelle finden.
  Err.Clear
  Set oSheetColumnConfig = oWbkCfg.Sheets(strTabNameSpalten)
  If (Err.Number <> 0) Then
    On Error GoTo Fehler
    ErrMessage = ErrMsgTabSpalten
    Err.Raise ErrNumTabSpalten + vbObjectError, , "Tabelle '" & strTabNameSpalten & "' in Konfigurationsdatei '" & cfg & "' nicht gefunden."
  End If
    
  ' Tabelle lesen und Daten abspeichern.
  Set oExpim = New CdatExpim
  Anzahl = oExpim.GetQuelldaten_XlTab(oSheetColumnConfig)
  lb = LBound(oExpim.Datenpuffer.Daten(strSpName), 1)
  ub = UBound(oExpim.Datenpuffer.Daten(strSpName), 1)
  For i = lb To ub
    Spalte = oExpim.Datenpuffer.Daten(strSpName)(i, DPidxWert)
    Kategorie = oExpim.Datenpuffer.Daten(strSpKategorie)(i, DPidxWert)
    ErsatzZiel = oExpim.Datenpuffer.Daten(strSpErsatzZiel)(i, DPidxWert)
    KategorieAlsFilter = oExpim.Datenpuffer.Daten(strSpKatAlsFilter)(i, DPidxWert)
    oSpaltenBeschreibung.Add Spalte, oExpim.Datenpuffer.Daten(strSpBeschreibung)(i, DPidxWert)
    oSpaltenMathGroesse.Add Spalte, oExpim.Datenpuffer.Daten(strSpMathGroesse)(i, DPidxWert)
    oSpaltenKategorie.Add Spalte, Kategorie
    oSpaltenErsatzZiel.Add Spalte, ErsatzZiel
    If (Kategorie <> "") Then
      If (Not oKategorieAlsFilter.Exists(Kategorie)) Then
        If (KategorieAlsFilter <> "") Then bool = True Else bool = False
        oKategorieAlsFilter.Add Kategorie, bool
      End If
    End If
  Next
  Set oExpim = Nothing
  Set oSheetColumnConfig = nothing
  
  DebugEcho "CdatKonfig.LeseSpaltenKonfig(): Gelesene Spaltenkonfiguration:"
  Call ZeigeSpaltenKonfig
  
  On Error GoTo 0
  LeseSpaltenKonfig = True
  Exit Function
  
Fehler:
  Set oSheetColumnConfig = nothing
  ErrEcho "CdatKonfig.LeseSpaltenKonfig()"
End Function



Public Sub ZeigeSpaltenKonfig()
  'Schreibt für Kontrollzwecke die gefundene Spaltenkonfiguration ins Debug-Protokoll.
  Dim Spalte As Variant
  DebugEcho "----------------------------------------------------------------------------------------"
  DebugEcho "Spalte" & vbTab & vbTab & "Beschreibung" & vbTab & vbTab & "Math. Größe" & vbTab & vbTab & "Kategorie"
  DebugEcho "----------------------------------------------------------------------------------------"
  For Each Spalte In Me.SpaltenBeschreibung
    DebugEcho Spalte & vbTab & vbTab & _
              Me.SpaltenBeschreibung(Spalte) & vbTab & vbTab & _
              Me.SpaltenMathGroesse(Spalte) & vbTab & vbTab & _
              Me.SpaltenKategorie(Spalte)
  Next
  DebugEcho "----------------------------------------------------------------------------------------"
End Sub



Public Sub ZeigeAlleEinheiten()
  'Schreibt für Kontrollzwecke alle verfügbaren Einheiten ins Debug-Protokoll.
  Dim Groesse As Variant
  Dim Einheit As Variant
  Dim Message As String
  DebugEcho vbNewLine & "CdatKonfig.ZeigeAlleEinheiten(): Konfiguration der Einheiten:"
  DebugEcho "-----------------------------------------------------------"
  For Each Groesse In Me.Einheiten
    For Each Einheit In Me.Einheiten(Groesse)
      Message = Message & "Einheit=" & Einheit & "  Faktor=" & Me.Einheiten(Groesse)(Einheit) & "  (" & Groesse & ")" & vbNewLine
    Next
  Next
  DebugEcho Message
End Sub

'für jEdit:  :folding=indent::collapseFolds=1:
