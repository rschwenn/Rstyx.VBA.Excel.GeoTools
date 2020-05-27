VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CtabTabelle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2003-2020  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'====================================================================================
' Klasse CtabTabelle
'====================================================================================
'
' Stellt Eigenschaften und Methoden für eine Tabelle zur Verfügung,
' die nicht aktiv sein muss:
 ' - Tabellenname
 ' - Tabellenklasse (Vorlagentyp)
 ' - Größe des aktuell genutzten Datenbereiches
 ' - Eigenschaften, die als Optionen das Verhalten von Methoden beeinflussen
 ' - Ermitteln und Vorhalten der Spaltenwidmung innerhalb des Datenbereiches
 '   (Größen und Einheiten)
 ' - Formatierung des Datenbereiches
 ' - Übertragen der Formel in der 1. Datenzeile auf alle anderen Zeilen.
 ' - Loeschen des Datenbereiches
 ' - Schreiben der Projektdaten in die vorgesehenen Felder
 ' - Schreiben einer Fußzeile
 ' - Bearbeiten der Tabellendaten mit Hilfe von Methoden des Datenpuffers bzw.
 '   des Expim-Objektes.
 ' - Vorbereitung der Tabelle für die eben genannten Operationen durch Benennung von
 '   Bereichen (Tabellenstruktur, Spaltenwidmung, Felder für Projektdaten)
 '
 ' - Interpolationsformel erstellen.
 ' - Doppelte Werte einer Spalte markieren.
 '
 ' ==> Die Eigenschaften bzgl. Datenbereich werden via Me.Syncronisieren ermittelt.
'
' Kern der Tabellenstruktur ist ein benannter Bereich mit dem in der Konstanten
' strInfoTraeger festgelegten Namen. Dieser Bereich muß identisch sein mit der ersten
' Datenzeile und legt somit Spaltenanfang und -ende sowie erste Zeile fest.
' Die letzte Zeile wird über "UsedRange" ermittelt.
' 
' Position und Größe des Datenbereiches werden folgendermaßen bestimmt:
'  - Zunächst werden via "UsedRange" alle Parameter bestimmt.
'  - Folgende Faktoren, wenn vorhanden, schränken diesen Bereich anschließend ein:
'    - Bereich "Daten.ErsteZelle" überschreibt Anfangszeile und -spalte
'    - Bereich "Daten.InfoTraeger" überschreibt Anfangszeile, -spalte und Endspalte.
'  - Die Endzeile des Datenbereiches wird immer in Echtzeit bestimmt.
'  - Bei Mißerfolg (keine Tabelle) liefern die entsprechenden Eigenschaften die Werte "-1".
'====================================================================================


Option Explicit

'Konstanten
Const strKlasse_Name               As String = "CtabTabelle"


'Variablen
 'Dim strAltesBlatt                  As String
 'Dim strNeuesBlatt                  As String
 
 'Spaltenanalyse für oExpim.
 Dim oSpaltenErsteZellen            As Scripting.Dictionary       '1. Zelle als Range
 Dim oSpaltenEinheiten              As Scripting.Dictionary       'Name der Einheit
 Dim oSpaltenFormate                As Scripting.Dictionary       'Formatangabe
 Dim oSpaltenBuchst2Name            As Scripting.Dictionary       'Key=Buchst., item=Spaltenname Add-In
 Dim strKategorien                  As String                     'Kategorien der Tabellenspalten
 
 'Namen benannter Zellbereiche.
 'Dim strInfoTraeger                 As String
 'Dim strFliesskomma                 As String
 'Dim strFormel                      As String
 
 Dim Listentrennzeichen             As String
 
 'Eigenschaften
 Dim lngErsteDatenZeile             As Long
 Dim lngErsteDatenSpalte            As Long
 Dim lngLetzteDatenSpalte           As Long
 Dim oInfotraeger                   As Range
 
 
 Dim ErrMsgDaten                    As String
 Dim ErrMsgInfoTraeger              As String
 Dim ErrMsgFliesskomma              As String
 Dim ErrMsgFormel                   As String
                                   
 Dim ErrNumErsteDatenZeile          As Long
 Dim ErrNumLetzteDatenZeile         As Long
 Dim ErrNumAnzahlDatenZeilen        As Long
 Dim ErrNumErsteDatenSpalte         As Long
 Dim ErrNumLetzteDatenSpalte        As Long
 Dim ErrNumTabKlasse                As Long
 Dim ErrNumTabName                  As Long
 Dim ErrNumLoeschenZellname         As Long
 Dim ErrNumAuswahlZellname          As Long
 Dim ErrNumInfoTraeger              As Long
 Dim ErrNumFliesskomma              As Long
 Dim ErrNumFormel                   As Long
 Dim ErrNumListentrenner            As Long
 Dim ErrNumInterpol                 As Long
 Dim ErrNumBereiche                 As Long
 Dim ErrNumPrjdatNichtGeschr        As Long
 Dim ErrNumFuss1NichtGeschr         As Long
 
 Dim intFormatDatenNKStellenAnzahl  As Integer
 Dim blnFormatDatenNKStellenSetzen  As Boolean
 Dim blnFormatDatenMitStreifen      As Boolean
 Dim blnFormatDatenOhneFuellung     As Boolean
 Dim bln_VorhWerteUeberschreiben    As Boolean
 Dim bln_FormelnErhalten            As Boolean
 Dim blnSilent                      As Boolean
 
 Dim oZielTabelle                   As Excel.Worksheet
'


Private Sub Class_Initialize()
    'On Error GoTo Fehler
    
    DebugEcho "CtabTabelle.Class_Initialize() startet ..."
    
    'Fehlernummern
    ErrNumInfoTraeger = 50100 - vbObjectError
    ErrNumErsteDatenZeile = 50101 - vbObjectError
    ErrNumLetzteDatenZeile = 50102 - vbObjectError
    ErrNumAnzahlDatenZeilen = 50103 - vbObjectError
    ErrNumErsteDatenSpalte = 50104 - vbObjectError
    ErrNumLetzteDatenSpalte = 50105 - vbObjectError
    ErrNumFliesskomma = 50106 - vbObjectError
    ErrNumFormel = 50107 - vbObjectError
    ErrNumListentrenner = 50108 - vbObjectError
    ErrNumInterpol = 50109 - vbObjectError
    ErrNumBereiche = 50110 - vbObjectError
    ErrNumTabKlasse = 50121 - vbObjectError
    ErrNumTabName = 50122 - vbObjectError
    ErrNumLoeschenZellname = 50131 - vbObjectError
    ErrNumAuswahlZellname = 50132 - vbObjectError
    ErrNumPrjdatNichtGeschr = 50141 - vbObjectError
    ErrNumFuss1NichtGeschr = 50142 - vbObjectError
    
    ErrMsgDaten = "Zwei mögliche Fehlerursachen:" & vbNewLine & vbNewLine & _
                  "1. Es ist keine Tabelle angegeben." & vbNewLine & vbNewLine & _
                  "2. Die angegebene Tabelle enthält keinen Zellbereich, der als erste Datenzeile " & vbNewLine & _
                  "   festgelegt ist und gleichzeitig als Träger der Formatierung dient." & vbNewLine & _
                  "   (Dieser Zellbereich muß den Namen '" & strInfoTraeger & "' tragen.)"
    
    ErrMsgInfoTraeger = "Die angegebene Tabelle enthält keinen Zellbereich, der als erste Datenzeile " & _
                        "festgelegt ist und gleichzeitig als Träger der Formatierung dient." & vbNewLine & _
                        "(Dieser Zellbereich muß den Namen '" & strInfoTraeger & "' tragen.)"
    
    ErrMsgFliesskomma = "Die angegebene Tabelle enthält keinen Zellbereich, der innerhalb der ersten Datenzeile " & _
                        "liegt und als 'Fliesskomma'-Bereich festgelegt ist." & vbNewLine & _
                        "(Dieser Zellbereich muß den Namen '" & strFliesskomma & "' tragen.)"
    
    ErrMsgFormel = "Die angegebene Tabelle enthält keinen Zellbereich, der innerhalb der ersten Datenzeile " & _
                   "liegt und als 'Formelträger' festgelegt ist." & vbNewLine & _
                   "(Dieser Zellbereich muß den Namen '" & strFormel & "' tragen.)"
    
    
    'Standardwerte
    Call FormatDatenOptionenReset
    Me.ModOpt_VorhWerteUeberschreiben = False
    Me.ModOpt_FormelnErhalten = True
    
    ErrMessage = ""
    
    Set oSpaltenErsteZellen = New Scripting.Dictionary
    Set oSpaltenEinheiten = New Scripting.Dictionary
    Set oSpaltenBuchst2Name = New Scripting.Dictionary
    Set oSpaltenFormate = New Scripting.Dictionary
    
    DebugEcho "CtabTabelle.Class_Initialize() fertig."
    Exit Sub
    
    Fehler:
    Set oSpaltenErsteZellen = Nothing
    Set oSpaltenEinheiten = Nothing
    Set oSpaltenBuchst2Name = Nothing
    Set oSpaltenFormate = Nothing
    FehlerNachricht "CtabTabelle.Class_Initialize()"
End Sub

Private Sub Class_Terminate()
    Set oSpaltenErsteZellen = Nothing
    Set oSpaltenEinheiten = Nothing
    Set oSpaltenBuchst2Name = Nothing
    Set oSpaltenFormate = Nothing
    DebugEcho "CtabTabelle.Class_Terminate() fertig."
End Sub

Private Sub InitZielTabelle()
    ' Setzt die Zieltabelle.
    On Error GoTo 0
    
    DebugEcho "CtabTabelle.InitZielTabelle() startet ..."
    
    If (oZielTabelle Is Nothing) Then
        ErrMessage = ErrMsgKeineZielTabelle
        Err.Raise vbObjectError + ErrNumTabName, , "Fehler beim Initialisieren der Zieltabelle."
    End If
    
    If (isTabellenSchutz(oZielTabelle)) Then
        ErrMessage = ""
        Err.Raise vbObjectError + ErrNumTabSchutz - vbObjectError, , "Aktion ist nicht möglich, Tabellenschutz ist aktiv."
    End If
    
    Call Me.Syncronisieren
    
    DebugEcho "CtabTabelle.InitZielTabelle() fertig."
    'Exit Sub
    '
    'Fehler:
    'FehlerNachricht "CtabTabelle.InitZielTabelle()"
End Sub



Sub FormatDatenOptionenReset()
    ' Setzt alle Optionen zur Formatierung des Datenbereiches auf die Standardwerte zurück.
    ' Diese bewirken beim Formatieren eine unveränderte Übernahme des "Infoträgers".
    
    On Error GoTo Fehler
    
    Me.FormatDatenNKStellenAnzahl = 3
    Me.FormatDatenNKStellenSetzen = False
    Me.FormatDatenMitStreifen = False
    Me.FormatDatenOhneFuellung = False
    
    Exit Sub
    
    Fehler:
    FehlerNachricht "CtabTabelle.FormatDatenOptionenReset()"
End Sub


Property Set ZielTabelle(inpZielTabelle As Excel.Worksheet)
    ' Setzt die Zieltabelle.
    Set oZielTabelle = inpZielTabelle
    Call InitZielTabelle()
End Property

Property Get ZielTabelle() As Excel.Worksheet
    ' Liefert Zieltabelle als Worksheet-Objekt.
    Set ZielTabelle = oZielTabelle
End Property


Property Get Klasse_Name() As String
    ' Liefert den Namen des Klassenmoduls
    Klasse_Name = strKlasse_Name
End Property


Property Get TabName() As String
    ' Liefert den Namen der Tabelle
    If (Not (oZielTabelle Is Nothing)) Then
        TabName = oZielTabelle.Name
    Else
        TabName = ""
        On Error Resume Next
        ErrMessage = ErrMsgKeineZielTabelle
        Err.Raise vbObjectError + ErrNumTabName, , "Fehler beim Ermitteln des Namens der Tabelle."
    End If
End Property


Property Get TabKlasse() As String
    '  Liefert den Namen der TabellenKlasse. Dieser entspricht dem Klassennamen des
    ' VBProjektes ohne Ziffern am Ende, die i.d.R. automatisch angehängt werden.
    ' => D.h. der gelieferte Name entspricht dem bereits in der Vorlagendatei
    '    festgelegten Klassennamen. Dieser darf also nicht mit einer Ziffer enden!
    ' => Damit kann festgestellt werden, ob bestimmte Aktionen auf die
    '    angegebene Tabelle sinnvoll angewendet werden könen.
    If (Not (oZielTabelle Is Nothing)) Then
        TabKlasse = substitute("[0-9]+$", "", oZielTabelle.CodeName, False, False)
    Else
        TabKlasse = ""
        On Error Resume Next
        ErrMessage = ErrMsgKeineZielTabelle
        Err.Raise vbObjectError + ErrNumTabKlasse, , "Fehler beim Ermitteln der Klasse (Vorlagentyp) der Tabelle."
    End If
End Property


Property Get Silent() As Boolean
    ' Liefert den aktuellen Modus für "Silent" (False: keine Meldungen).
    Silent = blnSilent
End Property

Property Let Silent(inpSilent As Boolean)
    ' Setzt den aktuellen Modus für "Silent" (False: keine Meldungen).
    blnSilent = inpSilent
End Property


Property Get Infotraeger() As Range
    ' Liefert den "Infotraeger" als Range-Objekt (kann 'Nothing' sein).
    Set Infotraeger = oInfotraeger
End Property


Property Get Fliesskomma() As Range
    ' Liefert den "Fliesskomma"-Bereich als bereinigtes Range-Objekt
    Dim oFliesskomma   As Range
    Set oFliesskomma = GetLokalerZellname(strFliesskomma, oZielTabelle)
    If (Not oFliesskomma Is Nothing) Then
        Set Fliesskomma = Union(oFliesskomma, oFliesskomma)
    Else
        Set Fliesskomma = Nothing
        On Error Resume Next
        ErrMessage = ErrMsgFliesskomma
        Err.Raise vbObjectError + ErrNumFliesskomma, , "Fehler beim Ermitteln des 'Fliesskomma'-Bereiches der Tabelle."
    End If
    Set oFliesskomma = Nothing
End Property


Property Get Formel() As Range
    ' Liefert den "Formel"-Bereich als bereinigtes Range-Objekt
    Dim oFormel   As Range
    Set oFormel = GetLokalerZellname(strFormel, oZielTabelle)
    If (Not oFormel Is Nothing) Then
        Set Formel = Union(oFormel, oFormel)
    Else
        Set Formel = Nothing
        On Error Resume Next
        ErrMessage = ErrMsgFormel
        Err.Raise vbObjectError + ErrNumFormel, , "Fehler beim Ermitteln des 'Formel'-Bereiches der Tabelle."
    End If
    Set oFormel = Nothing
End Property


Property Get ErsteDatenZeile() As Long
    ' Liefert die Nummer der ersten Zeile des Datenbereiches der Tabelle.
    ErsteDatenZeile = lngErsteDatenZeile
End Property


Property Get LetzteDatenZeile() As Long
    ' Liefert die Nummer der letzten Zeile des Datenbereiches der Tabelle.
    Dim ZeEnd       As Long
    Dim oUsedRange  As Range
    
    
    If (Not (oZielTabelle.Type = xlWorksheet)) Then
        DebugEcho "CtabTabelle.LetzteDatenZeile [Get]: Keine Excel-Tabelle angegeben!"
        ZeEnd = -1
    else
        Set oUsedRange = oZielTabelle.UsedRange
        ZeEnd = oUsedRange.Rows(oUsedRange.Rows.Count).Row
        'Sicherheit: Wichtig, wenn Infotraeger vorhanden. Kann sonst auch nicht schaden.
        If (ZeEnd < Me.ErsteDatenZeile) Then ZeEnd = Me.ErsteDatenZeile
        Set oUsedRange = Nothing
    end if
    LetzteDatenZeile = ZeEnd
End Property


Property Get AnzahlDatenZeilen() As Long
    ' Liefert die Anzahl der Datenzeilen der Tabelle.
    AnzahlDatenZeilen = Me.LetzteDatenZeile - Me.ErsteDatenZeile + 1
End Property


Property Get ErsteDatenSpalte() As Long
    ' Liefert die Nummer der ersten Spalte des Datenbereiches der Tabelle.
    ErsteDatenSpalte = lngErsteDatenSpalte
End Property


Property Get LetzteDatenSpalte() As Long
    ' Liefert die Nummer der letzten Spalte des Datenbereiches der Tabelle.
    LetzteDatenSpalte = lngLetzteDatenSpalte
End Property


Property Get FormatDatenMitStreifen() As Boolean
    ' Liefert den aktuellen Modus für FormatDaten (True: Formatierung mit Streifen).
    FormatDatenMitStreifen = blnFormatDatenMitStreifen
End Property

Property Let FormatDatenMitStreifen(inpFormatDatenMitStreifen As Boolean)
    ' Setzt den aktuellen Modus für FormatDaten (True: Formatierung mit Streifen).
    'On Error Resume Next
    On Error GoTo 0
    
    blnFormatDatenMitStreifen = inpFormatDatenMitStreifen
    
    'Mit Streifen heißt auch: ohne Füllung.
    If (blnFormatDatenMitStreifen) Then Me.FormatDatenOhneFuellung = True
    
End Property


Property Get FormatDatenOhneFuellung() As Boolean
    ' Liefert den aktuellen Füll-Modus für FormatDaten
    ' True: Füllung der Zellen des Formatträgers wird vor der Formatübertragung gelöscht.
    FormatDatenOhneFuellung = blnFormatDatenOhneFuellung
End Property

Property Let FormatDatenOhneFuellung(inpFormatDatenOhneFuellung As Boolean)
    ' Setzt den aktuellen Füll-Modus für FormatDaten.
    ' True: Füllung der Zellen des Formatträgers wird vor der Formatübertragung gelöscht.
    blnFormatDatenOhneFuellung = inpFormatDatenOhneFuellung
    
    ' Mit Füllung heißt auch: ohne Streifen.
    If (Not blnFormatDatenOhneFuellung) Then Me.FormatDatenMitStreifen = False
    
End Property


Property Get FormatDatenNKStellenAnzahl() As Integer
    ' Liefert die Anzahl der NachkommaStellenAnzahl für die 'Fliesskomma'-Spalten des Datenbereiches.
    FormatDatenNKStellenAnzahl = intFormatDatenNKStellenAnzahl
End Property

Property Let FormatDatenNKStellenAnzahl(inpFormatDatenNKStellenAnzahl As Integer)
    ' Setzt die Anzahl der NachkommaStellenAnzahl für die 'Fliesskomma'-Spalten des Datenbereiches.
    If ((inpFormatDatenNKStellenAnzahl >= 0) And (inpFormatDatenNKStellenAnzahl < 9)) Then
        intFormatDatenNKStellenAnzahl = inpFormatDatenNKStellenAnzahl
        
        'NK-Anzahl einstellen heißt auch: "Setzen der NachkommaStellen" einschalten!
        Me.FormatDatenNKStellenSetzen = True
    End If
End Property


Property Get FormatDatenNKStellenSetzen() As Boolean
    ' Liefert die Schalterstellung "Setzen der NachkommaStellen"
    FormatDatenNKStellenSetzen = blnFormatDatenNKStellenSetzen
End Property

Property Let FormatDatenNKStellenSetzen(inpFormatDatenNKStellenSetzen As Boolean)
    ' Setzt die Schalterstellung "Setzen der NachkommaStellen"
    blnFormatDatenNKStellenSetzen = inpFormatDatenNKStellenSetzen
End Property


Property Get SpaltenErsteZellen() As Scripting.Dictionary
    ' Liefert die ersten Zellen jeder Listenspalte als Range-Objekt in einem Dictionary.
    Set SpaltenErsteZellen = oSpaltenErsteZellen
End Property

Property Get SpaltenBuchst2Name() As Scripting.Dictionary
    ' Liefert die Zuordnung: XL-Spaltenbuchstabe => Add-In-Spaltenname.
    Set SpaltenBuchst2Name = oSpaltenBuchst2Name
End Property

Property Get SpaltenEinheiten() As Scripting.Dictionary
    ' Liefert die Namen der Einheiten jeder Listenspalte als String in einem Dictionary.
    Set SpaltenEinheiten = oSpaltenEinheiten
End Property

Property Get SpaltenFormate() As Scripting.Dictionary
    ' Liefert die Formate jeder Listenspalte als ???????? in einem Dictionary.
    Set SpaltenFormate = oSpaltenFormate
End Property

Property Get Kategorien() As String
    ' Liefert alle unterschiedlichen Kategorien der Tabelle.
    Kategorien = strKategorien
End Property


Property Get ModOpt_VorhWerteUeberschreiben() As Boolean
    ' Liefert die Option "Vorhandene Werte überschreiben".
    ModOpt_VorhWerteUeberschreiben = bln_VorhWerteUeberschreiben
End Property

Property Let ModOpt_VorhWerteUeberschreiben(inp_VorhWerteUeberschreiben As Boolean)
    ' Setzt die Option "Vorhandene Werte überschreiben".
    bln_VorhWerteUeberschreiben = inp_VorhWerteUeberschreiben
End Property


Property Get ModOpt_FormelnErhalten() As Boolean
    ' Liefert die Option "Formeln erhalten".
    ModOpt_FormelnErhalten = bln_FormelnErhalten
End Property

Property Let ModOpt_FormelnErhalten(inp_FormelnErhalten As Boolean)
    ' Setzt die Option "Formeln erhalten".
    bln_FormelnErhalten = inp_FormelnErhalten
End Property


Property Get TabTitel() As String
    ' Liefert den Titel der Tabelle anhand der Dokumenteigenschaften.
    TabTitel = ""
    If (Not (oZielTabelle Is Nothing)) Then
        TabTitel = oZielTabelle.Parent.BuiltinDocumentProperties("title").value
    End If
End Property



Public Function ExistsLokalerZellname(ByVal Name As String) As Boolean
    ' Überprüft, ob der benannte Zellbereich "Name" existiert und sich
    ' auf Zellen der Tabelle bezieht.
    On Error GoTo Fehler
    'On Error Resume Next
    Dim oNameRange As Range
    Set oNameRange = GetLokalerZellname(Name, oZielTabelle)
    If (oNameRange Is Nothing) Then
        ExistsLokalerZellname = False
    Else
        ExistsLokalerZellname = True
    End If
    Exit Function
    
    Fehler:
    FehlerNachricht "CtabTabelle.ExistsLokalerZellname()"
End Function



Public Sub FormatDaten()
    ' Überträgt das Format des "InfoTraegers" auf alle weiteren Datenzeilen
    ' und berücksichtigt dabei alle aktiven Format-Optionen.
    
    On Error GoTo Fehler
    
    Dim ZeAnz            As Long
    Dim i                As Integer
    Dim FormatString     As String
    Dim StatusScreen     As Boolean
    Dim StatusCalc       As Boolean
    Dim oRangeQuelle     As Range
    Dim oRangeZiel       As Range
    'Dim oStatusAktZelle  As Range
    'Dim oStatusAuswahl   As Range
    Dim Eigenschaft      As Variant
    
    if (Me.Infotraeger is Nothing) then
        ErrMessage = "Es ist kein 'InfoTraeger' vorhanden!"
        GoTo Fehler
    end if
    ZeAnz = Me.AnzahlDatenZeilen
    
    StatusScreen = oZielTabelle.Application.ScreenUpdating
    StatusCalc = oZielTabelle.EnableCalculation
    oZielTabelle.Application.ScreenUpdating = False
    oZielTabelle.EnableCalculation = False
    'Set oStatusAktZelle = ActiveCell
    'Set oStatusAuswahl = Selection
    
    If (Me.FormatDatenNKStellenSetzen) Then
        Set oRangeZiel = Me.Fliesskomma
        If (Not oRangeZiel Is Nothing) Then
            ' Benannten 'Fliesskomma'-Bereich formatieren
            If (Err) Then GoTo Fehler
            FormatString = "0"
            If (Me.FormatDatenNKStellenAnzahl > 0) Then FormatString = FormatString & "." & String(CLng(Me.FormatDatenNKStellenAnzahl), "0")
            oRangeZiel.NumberFormat = FormatString
        End If
    End If
    
    If (ZeAnz > 1) Then
        ' Formatvorlagenzeile (1. Datenzeile) kopieren und bei Bedarf Füllung (Farbe) löschen
        Set oRangeQuelle = oInfotraeger
        If (Me.FormatDatenOhneFuellung) Then oRangeQuelle.Interior.ColorIndex = xlNone
        
        If (Me.FormatDatenMitStreifen) Then
            ' Excel verwendet nicht den mit RGB() gelieferten Farbwert, sondern
            ' die in der Farbpalette der Arbeitsmappe verfügbare, ähnlichste Farbe!
            ' DESHALB: Sicherstellen, daß die richtige Farbpalette zur Verfügung steht.
            ' Nicht sehr rücksichtsvoll, aber mit Wirkung (und Nebenwirkungen):
            ' oZielTabelle.Parent.Colors = ThisWorkbook.Colors
              
            ' 2. Datenzeile: Format der 1. Datenzeile übertragen und grau einfärben
            ' Set oRangeZiel = oRangeQuelle.Resize(RowSize:=2)
            ' oRangeQuelle.AutoFill Destination:=oRangeZiel, Type:=xlFillFormats
            Set oRangeZiel = oInfotraeger.Offset(rowoffset:=1)
            oRangeQuelle.Copy
            oRangeZiel.PasteSpecial Paste:=xlPasteFormats
            
            'Set oRangeZiel = oRangeQuelle.Offset(rowoffset:=1)
            With oRangeZiel.Interior
                .Color = RGB(234, 234, 234)
                '.ColorIndex = 39
                .Pattern = xlSolid
            End With
            
            'Format der 1.+2. Datenzeile auf alle weiteren Zeilen übertragen
            Set oRangeQuelle = oInfotraeger.Resize(RowSize:=2)
            'Set oRangeZiel = oRangeQuelle.Resize(RowSize:=ZeAnz)
            'oRangeQuelle.AutoFill Destination:=oRangeZiel, Type:=xlFillFormats
            Set oRangeZiel = oInfotraeger.Offset(rowoffset:=2).Resize(RowSize:=ZeAnz - 2)
            oRangeQuelle.Copy
            oRangeZiel.PasteSpecial Paste:=xlPasteFormats
        Else
            ' Format ohne Streifen
            ' alle weiteren Zeilen formatieren
              
            ' Kopiert auch Grafikelemente der Quellzellen!!!
            ' Set oRangeZiel = oRangeQuelle.Resize(RowSize:=ZeAnz)
            ' oRangeQuelle.AutoFill Destination:=oRangeZiel, Type:=xlFillFormats
            
            Set oRangeZiel = oInfotraeger.Offset(rowoffset:=1).Resize(RowSize:=ZeAnz - 1)
            oRangeQuelle.Copy
            
            '!!! Wenn das aktive Fenster nicht das Fenster 1 ist, dann bewirkt PasteSpecial
            'die Aktivierung der Tabelle auch im Fenster Nr. 1 und die Markierung
            'des Zielbereiches darin!
            oRangeZiel.PasteSpecial Paste:=xlPasteFormats
        End If
        
        oZielTabelle.Application.CutCopyMode = False
        
        ' Zeilenhöhe einstellen
        Set oRangeZiel = oInfotraeger.Offset(rowoffset:=1).Resize(RowSize:=ZeAnz - 1)
        oRangeZiel.Rows.RowHeight = oInfotraeger.EntireRow.RowHeight
        'oRangeZiel.Locked = False        'Feldschutz nicht aufheben, sondern übertragen.
    End If
    
    'Alten Zustand wiederherstellen
    'oStatusAuswahl.Select
    'oStatusAktZelle.Activate
    
    oZielTabelle.Application.ScreenUpdating = StatusScreen
    oZielTabelle.EnableCalculation = StatusCalc
    'Set oStatusAuswahl = Nothing
    'Set oStatusAktZelle = Nothing
    Set oRangeQuelle = Nothing
    Set oRangeZiel = Nothing
    Exit Sub
    
    Fehler:
    oZielTabelle.Application.ScreenUpdating = StatusScreen
    If (Not oZielTabelle Is Nothing) Then oZielTabelle.EnableCalculation = StatusCalc
    'Set oStatusAuswahl = Nothing
    'Set oStatusAktZelle = Nothing
    Set oRangeQuelle = Nothing
    Set oRangeZiel = Nothing
    If (Me.Silent) Then
        ErrMessage = ""
        Err.Clear
    Else
        FehlerNachricht "CtabTabelle.FormatDaten()"
    End If
End Sub



Public Sub LoeschenDaten()
    ' Standardlöschen.
    ' Alle Datenzeilen der Tabelle werden gelöscht.
    
    On Error GoTo Fehler
    
    Dim StatusScreen  As Boolean
    Dim StatusCalc    As Boolean
    Dim oRangeZiel    As Range
    Dim ur            As Range
    Dim LetzteSpalte  As Long
    Dim ZeAnz         As Long
    Dim SpEnd         As Long
    
    if (Me.Infotraeger is Nothing) then
        ErrMessage = "Es ist kein 'InfoTraeger' vorhanden!"
        GoTo Fehler
    end if
    ZeAnz = Me.AnzahlDatenZeilen
    SpEnd = Me.LetzteDatenSpalte
    
    StatusScreen = oZielTabelle.Application.ScreenUpdating
    StatusCalc = oZielTabelle.EnableCalculation
    oZielTabelle.Application.ScreenUpdating = False
    oZielTabelle.EnableCalculation = False
    
    ' Formatierung der 1. Datenzeile erhalten:
    oInfotraeger.ClearContents
    
    ' Restliche Zeilen komplett entfernen.
    If (ZeAnz > 1) Then
        ' Ermitteln der am weitesten rechts liegenden, verwendeten Spalte.
        Set ur = oZielTabelle.UsedRange
        LetzteSpalte = ur.Columns(ur.Columns.Count).Column
        
        ' Löschen der Zellen oder der ganzen Zeilen.
        Set oRangeZiel = oInfotraeger.Offset(rowoffset:=1).Resize(RowSize:=ZeAnz - 1)
        if (LetzteSpalte > SpEnd) then
            ' Rechts neben dem Datenbereich werden Zellen genutzt => nur den Datenbereich löschen!
            oRangeZiel.Delete (xlShiftUp)
        else
            ' Es wird nur der Datenbereich genutzt => ganze Zeilen löschen.
            oRangeZiel.EntireRow.Delete
        end if
    End If
    
    ' Cursor positionieren
    'oInfotraeger.Cells(1, 1).Select
    'oZielTabelle.Application.GoTo Reference:=oInfotraeger.Cells(1, 1), Scroll:=True
    oZielTabelle.Application.ScreenUpdating = StatusScreen
    oZielTabelle.EnableCalculation = StatusCalc
    Set oRangeZiel = Nothing
    Exit Sub
    
    Fehler:
    oZielTabelle.Application.ScreenUpdating = StatusScreen
    If (Not oZielTabelle Is Nothing) Then oZielTabelle.EnableCalculation = StatusCalc
    Set oRangeZiel = Nothing
    FehlerNachricht "CtabTabelle.LoeschenDaten()"
End Sub



Public Sub UebertragenFormeln()
    ' Die Formeln des benannten Bereiches "Daten.Formel" werden in alle darunter
    ' liegenden Zellen bis Ende des Datenbereiches kopiert.
    
    On Error GoTo Fehler
    
    Dim ZeAnz         As Long
    Dim StatusScreen  As Boolean
    Dim StatusCalc    As Boolean
    Dim oFormeln      As Range
    Dim oRangeQuelle  As Range
    
    ZeAnz = Me.AnzahlDatenZeilen
    'If (Err) Then GoTo Fehler
    
    StatusScreen = oZielTabelle.Application.ScreenUpdating
    StatusCalc = oZielTabelle.EnableCalculation
    oZielTabelle.Application.ScreenUpdating = False
    oZielTabelle.EnableCalculation = False
    
    If (ZeAnz > 1) Then
        ' Formeln übertragen
        Set oFormeln = Me.Formel
        'If (Err) Then GoTo Fehler
        If (Not (oFormeln Is Nothing)) Then
            For Each oRangeQuelle In oFormeln.Areas
                oRangeQuelle.Resize(ZeAnz).FormulaLocal = oRangeQuelle.FormulaLocal
            Next
        End If
    End If
    
    oZielTabelle.Application.ScreenUpdating = StatusScreen
    oZielTabelle.EnableCalculation = StatusCalc
    Set oFormeln = Nothing
    Set oRangeQuelle = Nothing
    
    Exit Sub
    
    Fehler:
    oZielTabelle.Application.ScreenUpdating = StatusScreen
    If (Not oZielTabelle Is Nothing) Then oZielTabelle.EnableCalculation = StatusCalc
    FehlerNachricht "CtabTabelle.UebertragenFormeln()"
End Sub



Public Sub SchreibeFusszeile_1(Optional Silent As Boolean = False)
    ' Schreibt die Fusszeile_1 in die angegebene Tabelle.
    ' Parameter:  Silent ... wenn "ja", wird im Fehlerfall die Meldung unterdrückt.
    
    On Error GoTo Fehler
    
    Dim FeldGefunden  As Boolean
    Dim StatusScreen  As Boolean
    Dim StatusCalc    As Boolean
    Dim Tx_FussLi     As String
    Dim Tx_FussMi     As String
    Dim Tx_FussRe     As String
    Dim Fundstellen   As Object
    
    'StatusScreen = oZielTabelle.Application.ScreenUpdating
    'StatusCalc = oZielTabelle.EnableCalculation
    'oZielTabelle.Application.ScreenUpdating = False
    'oZielTabelle.EnableCalculation = False
    
    'FeldGefunden = SchreibenFelderInTabelle(oDictPrjDat)
    
    If (ThisWorkbook.Metadaten.Ort_Fusszeile_Excel_1 = "") Then
        ErrMessage = "Keine Fusszeile verfügbar. Mögliche Fehlerursachen:" & vbNewLine & vbNewLine & _
                     "1. Die Datei mit Standortinformationen (" & ThisWorkbook.Konfig.Ort_Dateiname & ") ist nicht vorhanden." & vbNewLine & vbNewLine & _
                     "2. In dieser Datei existiert kein Eintrag 'Fusszeile_Excel_1'."
        Err.Raise vbObjectError + ErrNumFuss1NichtGeschr, , "Fusszeile konnten nicht geschrieben werden."
    Else
        With oZielTabelle.PageSetup
            'Schriftart und -größe merken (&"Schriftart, Stil", &nn)
            ThisWorkbook.RegExp.IgnoreCase = True
            ThisWorkbook.RegExp.Pattern = "^(&""[0-9A-Za-z, _]*"")?(&[0-9]{1,2})?"
            'ThisWorkbook.RegExp.Pattern = "^(&[0-9]{1,2})?(&""[A-Za-z, _]*"")?(&[0-9]{1,2})?"
            ThisWorkbook.RegExp.Global = False
            Set Fundstellen = ThisWorkbook.RegExp.Execute(Trim$(.LeftFooter))
            If (Fundstellen.Count > 0) Then Tx_FussLi = Fundstellen(0).value
            Set Fundstellen = ThisWorkbook.RegExp.Execute(Trim$(.CenterFooter))
            If (Fundstellen.Count > 0) Then Tx_FussMi = Fundstellen(0).value
            Set Fundstellen = ThisWorkbook.RegExp.Execute(Trim$(.RightFooter))
            If (Fundstellen.Count > 0) Then Tx_FussRe = Fundstellen(0).value
            
            'MsgBox "links: " & .LeftFooter & "  (" & Tx_FussLi & ")" & vbNewLine & _
            '"mitte: " & .CenterFooter & "  (" & Tx_FussMi & ")" & vbNewLine & _
            '"rechts: " & .RightFooter & "  (" & Tx_FussRe & ")"
            
            'Neue Texte mit alter Schriftgröße setzen
            '.LeftFooter = Tx_FussLi & "&F (&A)"
            '.CenterFooter = Tx_FussMi & ThisWorkbook.Metadaten.Ort_Fusszeile_Excel_1
            
            .LeftFooter = Tx_FussLi & ThisWorkbook.Metadaten.Ort_Fusszeile_Excel_1
            .CenterFooter = ""
            .RightFooter = Tx_FussRe & "Seite &P von &N"
        End With
        oZielTabelle.Application.StatusBar = "Fußzeile eingerichtet."
    End If
    'oZielTabelle.Application.ScreenUpdating = StatusScreen
    'oZielTabelle.EnableCalculation = StatusCalc
    Exit Sub
    
    Fehler:
    'oZielTabelle.Application.ScreenUpdating = StatusScreen
    'oZielTabelle.EnableCalculation = StatusCalc
    If (Silent or Me.Silent) Then
        ErrMessage = ""
        Err.Clear
    Else
        FehlerNachricht "CtabTabelle.SchreibeFusszeile_1()"
    End If
End Sub



Public Sub SchreibeMetaDaten(Optional Silent As Boolean = False)
    ' Schreibt von allen verfügbaren Metadaten diejenigen in die Tabelle,
    ' für die entsprechend benannte Zellen existieren.
    ' => Dieser Routine sollte i.d.R. ein Metadaten-Update vorausgehen.
    
    On Error GoTo Fehler
    
    Dim ErfogProjekt  As Boolean
    Dim ErfogExtra    As Boolean
    Dim StatusScreen  As Boolean
    Dim StatusCalc    As Boolean
    
    StatusScreen = oZielTabelle.Application.ScreenUpdating
    StatusCalc = oZielTabelle.EnableCalculation
    oZielTabelle.Application.ScreenUpdating = False
    oZielTabelle.EnableCalculation = False
    
    ErfogProjekt = SchreibenFelderInTabelle(ThisWorkbook.Metadaten.AlleProjektDaten, oZielTabelle)
    ErfogExtra   = SchreibenFelderInTabelle(ThisWorkbook.Metadaten.AlleExtraDaten, oZielTabelle)
    
    If (Not (ErfogProjekt or ErfogExtra)) Then
        ErrMessage = "Mögliche Fehlerursachen:" & vbNewLine & vbNewLine & _
                     "1. Es ist keine Tabelle angegeben." & vbNewLine & vbNewLine & _
                     "2. Alle vorhandenen Projektdatenfelder sind schreibgeschützt." & vbNewLine & vbNewLine & _
                     "3. Die angegebene Tabelle enthält keine benannten Zellen für die Aufnahme " & _
                     "von Metadaten. Folgende Namen sind dafür vereinbart:" & vbNewLine & vbNewLine & _
                     "   Prj.ProjektbezX (mit X=1,2,3)" & vbNewLine & _
                     "   Prj.AuftragNr" & vbNewLine & _
                     "   Prj.BearbUsername" & vbNewLine & _
                     "   Prj.BearbVollerName" & vbNewLine & _
                     "   Prj.BearbNachname" & vbNewLine & _
                     "   Prj.Datum" & vbNewLine & _
                     "   Prj.HoehenSystem" & vbNewLine & _
                     "   Prj.KooSystem" & vbNewLine & _
                     "   Prj.TraBasisUeb" & vbNewLine & _
                     "   x.*"
        Err.Raise vbObjectError + ErrNumPrjdatNichtGeschr, , "Metadaten konnten nicht geschrieben werden."
    End If
    Echo "Projektdaten wurden in die Tabelle eingetragen."
    oZielTabelle.Application.StatusBar = "Projektdaten wurden in die Tabelle eingetragen."
    oZielTabelle.Application.ScreenUpdating = StatusScreen
    oZielTabelle.EnableCalculation = StatusCalc
    Exit Sub
    
    Fehler:
    oZielTabelle.Application.ScreenUpdating = StatusScreen
    oZielTabelle.EnableCalculation = StatusCalc
    If (Silent or Me.Silent) Then
        ErrMessage = ""
        Err.Clear
    Else
        FehlerNachricht "CtabTabelle.SchreibeMetaDaten()"
    End If
End Sub



Public Function GetProjektDaten() As Scripting.Dictionary
    ' Liest alle verfügbaren Projektdaten aus der Tabelle und
    ' gibt diese als Dictionary (oder nothing) zurück.
    
    On Error GoTo Fehler
    set GetProjektDaten = GetFelderAusTabelle("Prj.", oZielTabelle)
    Exit Function
    
    Fehler:
    FehlerNachricht "CtabTabelle.GetProjektDaten()"
End Function



Public Function GetExtraDaten() As Scripting.Dictionary
    ' Liest alle verfügbaren Extradaten aus der Tabelle und
    ' gibt diese als Dictionary (oder nothing) zurück.
    
    On Error GoTo Fehler
    set GetExtraDaten = GetFelderAusTabelle("x.", oZielTabelle)
    Exit Function
    
    Fehler:
    FehlerNachricht "CtabTabelle.GetExtraDaten()"
End Function


Sub Selection2Infotraeger()
    
    ' *** Selection gibt es nur in der AKTIVEN Tabelle ***
    
    '' Die aktuelle Selection wird als neuer Infoträger festgelegt.
    'On Error GoTo Fehler
    'If (Not isSelectionRechteck) Then
    '    Err.Raise vbObjectError + ErrNumNoRangeSelection - vbObjectError, , "Der markierte Bereich muß ein (einziges) Rechteck sein!"
    'End If
    'If (Selection.Rows.Count <> 1) Then
    '    Err.Raise vbObjectError + ErrNumNoRangeSelection - vbObjectError, , "Der markierte Bereich darf sich nicht über mehrere Zeilen erstrecken!"
    'End If
    'On Error Resume Next
    '
    'Call SetInfotraeger(Selection)
    'Exit Sub
    '
    'Fehler:
    'FehlerNachricht "CtabTabelle.Selection2Infotraeger()"
End Sub


Sub SetInfotraeger(Bereich As Range)
    ' Ein neuer Infoträger wird festgelegt.
    ' Eingabe:  Bereich ... Zellbereich für den Infotraeger.
    
    Dim benannteZellen        As Names
    Dim strBezug              As String
    Dim Feldname              As String
    
    On Error GoTo Fehler
    
    ' Bestehenden Infotraeger und enthaltene Bereiche löschen.
    If (ExistsLokalerZellname(strInfoTraeger)) Then LoeschenZellname (strInfoTraeger)
    If (ExistsLokalerZellname(strFliesskomma)) Then LoeschenZellname (strFliesskomma)
    If (ExistsLokalerZellname(strFormel)) Then LoeschenZellname (strFormel)
    If (Err) Then GoTo Fehler
    Set benannteZellen = oZielTabelle.Parent.Names
    'If ((InStr(oZielTabelle.Parent.Name, " ") > 0) Or (InStr(oZielTabelle.Name, " ") > 0) _
    '                                          Or (InStr(oZielTabelle.Name, "-") > 0) _
    '                                          Or (InStr(oZielTabelle.Name, ",") > 0)) Then
    Feldname = "'" & oZielTabelle.Name & "'!" & strInfoTraeger
    'Else
    '  Feldname = oZielTabelle.Name & "!" & strInfoTraeger
    'End If
    strBezug = "=" & Bereich.Address(RowAbsolute:=True, ColumnAbsolute:=True, ReferenceStyle:=xlA1, External:=True)
    benannteZellen.Add Name:=Feldname, RefersTo:=strBezug
    Set benannteZellen = Nothing
    Exit Sub
    
    Fehler:
    FehlerNachricht "CtabTabelle.SetInfotraeger()"
End Sub


Sub Selection2Fliesskomma()
    
    ' *** Selection gibt es nur in der AKTIVEN Tabelle ***
    
    '' Die aktuelle Selection wird als neuer Bereich festgelegt,
    '' für den programm- und menügesteuert die Nachkommastellen eingestellt werden können.
    '
    'Dim benannteZellen        As Names
    'Dim ointersect            As Range
    'Dim Message1              As String
    'Dim strBezug              As String
    'Dim Feldname              As String
    '
    'Message1 = "Der 'Fliesskomma'-Bereich der Tabelle legt für den Datenbereich die Spalten fest, " & _
    '           "für die menügesteuert die Anzahl der Nachkommastellen geändert werden kann." & vbNewLine & _
    '           "==> Dafür müssen alle entsprechenden Zellen markiert sein. Die Markierung kann aus " & _
    '           "mehreren Teilen bestehen und muß innerhalb des 'Informationsträgers' liegen. "
    '
    'On Error GoTo Fehler
    '
    'If (Not (oZielTabelle.Type = xlWorksheet)) Then
    '    Err.Raise vbObjectError + ErrNumKeineZielTabelle - vbObjectError, , ErrMsgKeineZielTabelle
    'End If
    '
    'If (oInfotraeger Is Nothing) Then
    '    ErrMessage = ErrMsgInfoTraeger
    '    Err.Raise vbObjectError + ErrNumInfoTraeger, , "Zuerst die erste Datenzeile ('Infoträger') festlegen!"
    'End If
    'Set ointersect = Intersect(Selection, oInfotraeger)
    'If (ointersect Is Nothing) Then
    '    ErrMessage = Message1
    '    Err.Raise vbObjectError + ErrNumFliesskomma, , "Die aktive Markierung liegt nicht innerhalb der ersten Datenzeile!"
    'End If
    'If (Union(Selection, Selection).Address <> ointersect.Address) Then
    '    ErrMessage = Message1
    '    Err.Raise vbObjectError + ErrNumFliesskomma, , "Die aktive Markierung liegt nicht innerhalb der ersten Datenzeile!"
    'End If
    'On Error Resume Next
    '
    'If (ExistsLokalerZellname(strFliesskomma)) Then LoeschenZellname (strFliesskomma)
    'If (Err) Then GoTo Fehler
    'Set benannteZellen = oZielTabelle.Parent.Names
    ''If ((InStr(oZielTabelle.Parent.Name, " ") > 0) Or (InStr(oZielTabelle.Name, " ") > 0)) Then
    'Feldname = "'" & oZielTabelle.Name & "'!" & strFliesskomma
    ''Else
    ''  Feldname = oZielTabelle.Name & "!" & strFliesskomma
    ''End If
    'strBezug = "=" & Union(Selection, Selection).Address(RowAbsolute:=True, ColumnAbsolute:=True, ReferenceStyle:=xlA1, External:=True)
    'benannteZellen.Add Name:=Feldname, RefersTo:=strBezug
    'Set benannteZellen = Nothing
    'Exit Sub
    '
    'Fehler:
    'FehlerNachricht "CtabTabelle.Selection2Fliesskomma()"
End Sub


Sub Selection2Spaltenname(ByVal SpName As String, Optional ByVal Einheit As String = "")
    
    ' *** Selection gibt es nur in der AKTIVEN Tabelle ***
    
    '' Es wird ein benannter Bereich erzeugt, der sich auf die aktive Zellauswahl bezieht.
    '' Ist der Name bereits vorhanden (unabhängig von der Einheit), wird dieser zunächst gelöscht.
    '' Trägt die aktive Zellauswahl bereits einen Namen, wird dieser ebenfalls gelöscht.
    '' Parameter: SpName   ... Spaltenname ohne Prefix und ohne Einheit (="", wenn Name nur gelöscht werden soll).
    ''                         => Prefix für WertStatus (Ist, Soll ..) ist bereits enthalten.
    ''            Einheit  ... Name einer Einheit (="", wenn keine Einheit angegeben werden soll).
    'On Error GoTo Fehler
    'If (Not isSelectionRechteck) Then
    '  Err.Raise vbObjectError + ErrNumNoRangeSelection - vbObjectError, , "Der markierte Bereich muß ein (einziges) Rechteck sein!"
    'End If
    'If (Selection.Columns.Count <> 1) Then
    '  Err.Raise vbObjectError + ErrNumNoRangeSelection - vbObjectError, , "Der markierte Bereich darf sich nicht über mehrere Spalten erstrecken!"
    'End If
    'On Error Resume Next
    '
    'Call SetSpaltenname(Selection, SpName, Einheit)
    '
    '' Eigenschaften dieser Klasse mit der geänderten Tabelle abgleichen.
    'Call Me.Syncronisieren
    'Exit Sub
    '
    'Fehler:
    'FehlerNachricht "CtabTabelle.Selection2Spaltenname()"
End Sub


Sub SetSpaltenname(Bereich As Range, ByVal SpName As String, ByVal Einheit As String, Optional byVal blnTitel As Boolean = false)
    ' Es wird ein Spaltenname mit dem angegebenen Zellbezug erzeugt.
    ' Die Zelle in der darüber gelegenen Zeile erhält den Kommentar und bei Bedarf als Wert den Spaltennamen.
    ' Ist der Name bereits vorhanden (unabhängig von der Einheit), wird dieser zunächst gelöscht.
    ' Trägt der angegebene Zellbezug bereits einen Namen, wird dieser ebenfalls gelöscht.
    ' Parameter: Bereich  ... Zellbereich, der den Spaltenname erhalten soll (d.h. die erste Zeile der Spalte)
    '                         => muss sich auf genau eine Zelle beziehen.
    '            SpName   ... Spaltenname ohne Prefix und ohne Einheit (="", wenn Name nur gelöscht werden soll).
    '                         => Prefix für WertStatus (Ist, Soll ..) ist bereits enthalten.
    '            Einheit  ... Name einer Einheit ("" oder "ohne", wenn keine Einheit angegeben werden soll).
    Dim benannteZellen                  As Names
    Dim Kommentar                       As Excel.Comment
    Dim oRange                          As Range
    Dim AnzZeilen                       As Long
    Dim strBezug                        As String
    Dim Feldname                        As String
    Dim strAlterNameKomplett            As String
    Dim strAlteEinheit                  As String
    Dim strNeuerNameKomplett            As String
    Dim strKommentar                    As String
    Dim oSpNameAttr                     As Scripting.Dictionary
    
    On Error GoTo Fehler
    
    ' Bereich auf genau eine Zelle einschränken.
    Set oRange = Bereich.Cells(1,1)
    if (Einheit = "ohne") Then Einheit = ""
    
    ' Alle Namen und den zugehörigen Kommentar des Bereiches löschen.
    On Error Resume Next
    Do
        oRange.Name.Delete
    Loop Until (Err.Number <> 0)
    Err.Clear
    On Error GoTo Fehler
    If (oRange.Row > 1) Then
        If (Not oRange.Offset(rowoffset:=-1).Comment Is Nothing) Then oRange.Offset(rowoffset:=-1).Comment.Delete
    End If
    
    ' Bestehende Spaltennamen incl. Kommentare außerhalb des Bereiches löschen.
    If (Me.SpaltenEinheiten.Exists(SpName)) Then
        strAlterNameKomplett = PrefixSpaltenname & SpName
        strAlteEinheit = Me.SpaltenEinheiten(SpName)
        If (strAlteEinheit <> "ohne") Then strAlterNameKomplett = strAlterNameKomplett & TrennerEinheit & strAlteEinheit
        If (ExistsLokalerZellname(strAlterNameKomplett)) Then
            Call LoeschenZellname(strAlterNameKomplett)
            If (Me.SpaltenErsteZellen(SpName).Row > 1) Then
                Set Kommentar = Me.SpaltenErsteZellen(SpName).Offset(rowoffset:=-1).Comment
                If (Not Kommentar Is Nothing) Then Kommentar.Delete
            End If
        End If
        If (Err) Then GoTo Fehler
    End If
    
    If (SpName <> "") Then
        ' Neuen Spaltennamen einfügen.
        strNeuerNameKomplett = PrefixSpaltenname & SpName
        If (Einheit <> "") Then strNeuerNameKomplett = strNeuerNameKomplett & TrennerEinheit & Einheit
        Set benannteZellen = oZielTabelle.Parent.Names
        'If ((InStr(oZielTabelle.Parent.Name, " ") > 0) Or (InStr(oZielTabelle.Name, " ") > 0)) Then
        Feldname = "'" & oZielTabelle.Name & "'!" & strNeuerNameKomplett
        'Else
        '  Feldname = oZielTabelle.Name & "!" & strNeuerNameKomplett
        'End If
        strBezug = "=" & oRange.Address(RowAbsolute:=True, ColumnAbsolute:=True, ReferenceStyle:=xlA1, External:=True)
        benannteZellen.Add Name:=Feldname, RefersTo:=strBezug
        Set benannteZellen = Nothing
        
        'Titel und Kommentar einfügen/ändern.
        If (oRange.Row > 1) Then
            ' Titel
            if (blnTitel) then oRange.Offset(rowoffset:=-1).Value = SpName
            
            ' Kommentar
            Set Kommentar = oRange.Offset(rowoffset:=-1).Comment
            If (Not Kommentar Is Nothing) Then Kommentar.Delete
            Set oSpNameAttr = ThisWorkbook.Konfig.SpNameAttr(SpName)
            AnzZeilen = 3
            strKommentar = vbLf & "Spalte:       " & SpName & vbLf & _
                                  "Kategorie:  " & oSpNameAttr("Kategorie") & vbLf & _
                                  "Titel:          " & oSpNameAttr("Titel")
            If (oSpNameAttr("StatusBez") <> "") Then
                AnzZeilen = AnzZeilen + 1
                strKommentar = strKommentar & vbLf & "Status:       " & oSpNameAttr("StatusBez")
            End If
            If ((oSpNameAttr("MathGroesse") <> "") or (Einheit <> "")) Then
                AnzZeilen = AnzZeilen + 1
                If (Einheit <> "") Then
                    strKommentar = strKommentar & vbLf & "Einheit:       " & Einheit
                Else
                    strKommentar = strKommentar & vbLf & "Einheit:       " & "(ohne)"
                End If
            End If
            Call oRange.Offset(rowoffset:=-1).AddComment(strKommentar)
            Set Kommentar = oRange.Offset(rowoffset:=-1).Comment
            If (Not Kommentar Is Nothing) Then
                Kommentar.Shape.Width = 200
                Kommentar.Shape.Height = 20 + AnzZeilen * 12
            End If
        End If
    End If
    
    Set oRange = Nothing
    Set Kommentar = Nothing
    Set oSpNameAttr = Nothing
    Exit Sub
    
    Fehler:
    FehlerNachricht "CtabTabelle.SetSpaltenname()"
End Sub


Sub Selection2Feldname(ByVal Feldname As String)
    
    ' *** Selection gibt es nur in der AKTIVEN Tabelle ***
    
    '' Es wird ein benannter Bereich erzeugt, der sich auf die aktive Zelle bezieht.
    '' (Bei verbundenen Zellen ist dies die linke obere Ecke).
    '' Ist der Name bereits vorhanden, wird dieser zunächst gelöscht.
    '' Trägt die aktive Zellauswahl bereits einen Namen, wird dieser ebenfalls gelöscht.
    '' Parameter: Feldname  ... kompletter Bereichsname (="", wenn Name nur gelöscht werden soll).
    '
    'Dim benannteZellen                  As Names
    'Dim Kommentar                       As Excel.Comment
    'Dim BereichQuelle                   As Range
    'Dim BereichZiel                     As Range
    ''Dim BereichZielKommentar            As Range
    'Dim strBezug                        As String
    'Dim FeldNameAbsolut                 As String
    'Dim strAlterNameKomplett            As String
    'Dim strAlteEinheit                  As String
    'Dim strNeuerNameKomplett            As String
    'Dim strKommentar                    As String
    'Dim strEinheitImKommentar           As String
    '
    'On Error GoTo Fehler
    'If (Not isSelectionRechteck) Then
    '  Err.Raise vbObjectError + ErrNumNoRangeSelection - vbObjectError, , "Der markierte Bereich muß ein (einziges) Rechteck sein!"
    'End If
    '' Verbundene Zellen sind erlaubt:
    ''If ((Selection.Columns.Count <> 1) Or (Selection.Rows.Count <> 1)) Then
    ''  Err.Raise vbObjectError + ErrNumNoRangeSelection - vbObjectError, , "Der markierte Bereich darf sich nicht über mehrere Zeilen oder Spalten erstrecken!"
    ''End If
    '
    'Set BereichZiel = ActiveCell
    '
    '' Alle Namen und den Kommentar des Zielbereiches löschen.
    'On Error Resume Next
    'Do
    '    BereichZiel.Name.Delete
    'Loop Until (Err.Number <> 0)
    'Err.Clear
    'On Error GoTo Fehler
    'If (Not BereichZiel.Comment Is Nothing) Then BereichZiel.Comment.Delete
    '
    '
    'If (Feldname <> "") Then
    '    ' Bestehenden Feldnamen incl. Kommentar außerhalb der aktiven Zellauswahl löschen.
    '    If (ExistsLokalerZellname(Feldname)) Then
    '        Set BereichQuelle = GetLokalerZellname(Feldname, oZielTabelle)
    '        Set Kommentar = Cells(BereichQuelle.Row, BereichQuelle.Columns(1).Column).Comment
    '        If (Not Kommentar Is Nothing) Then Kommentar.Delete
    '        Call LoeschenZellname(Feldname)
    '    End If
    '    If (Err) Then GoTo Fehler
    '    
    '    ' Neuen Feldnamen einfügen.
    '    'strNeuerNameKomplett = PrefixFeldname & Feldname
    '    strNeuerNameKomplett = Feldname
    '    Set benannteZellen = oZielTabelle.Parent.Names
    '    'If ((InStr(oZielTabelle.Parent.Name, " ") > 0) Or (InStr(oZielTabelle.Name, " ") > 0)) Then
    '    FeldNameAbsolut = "'" & oZielTabelle.Name & "'!" & strNeuerNameKomplett
    '    'Else
    '    '  FeldNameAbsolut = oZielTabelle.Name & "!" & strNeuerNameKomplett
    '    'End If
    '    strBezug = "=" & BereichZiel.Address(RowAbsolute:=True, ColumnAbsolute:=True, ReferenceStyle:=xlA1, External:=True)
    '    benannteZellen.Add Name:=FeldNameAbsolut, RefersTo:=strBezug
    '    Set benannteZellen = Nothing
    '    
    '    ' Kommentar einfügen/ändern.
    '    strKommentar = "Feld:    " & Feldname & vbLf & _
    '                   "Inhalt:  " & ThisWorkbook.Metadaten.TitelDerProjektDaten(Feldname)
    '    Call BereichZiel.AddComment(strKommentar)
    '    Set Kommentar = BereichZiel.Comment
    '    If (Not Kommentar Is Nothing) Then
    '        Kommentar.Shape.Width = 155
    '        Kommentar.Shape.Height = 25
    '    End If
    'End If
    '
    'Set Kommentar = Nothing
    'Set BereichQuelle = Nothing
    'Set BereichZiel = Nothing
    '
    '
    'Exit Sub
    '
    'Fehler:
    'FehlerNachricht "CtabTabelle.Selection2Feldname()"
End Sub



Sub Selection2Formel()
    
    ' *** Selection gibt es nur in der AKTIVEN Tabelle ***
    
    '' Die aktuelle Selection wird als neuer Bereich festgelegt,
    '' für den die Formeln der ersten Datenzeile auf alle anderen Zeilen
    '' übertragen werden können.
    '
    'Dim benannteZellen        As Names
    'Dim ointersect            As Range
    '
    'Dim Titel                 As String
    'Dim Message               As String
    'Dim Message1              As String
    'Dim Buttons               As Integer
    'Dim strBezug              As String
    'Dim Feldname              As String
    '
    'Message1 = "Der 'Formel'-Bereich der Tabelle legt für den Datenbereich die Spalten fest, für die " & _
    '           "die Formeln der ersten Datenzeile auf alle anderen Zeilen übertragen werden können." & vbNewLine & _
    '           "==> Dafür müssen alle entsprechenden Zellen markiert sein. Die Markierung kann aus " & _
    '           "mehreren Teilen bestehen und muß innerhalb des 'Informationsträgers' liegen. "
    '
    'On Error GoTo Fehler
    '
    'If (Not (oZielTabelle.Type = xlWorksheet)) Then
    '  Err.Raise vbObjectError + ErrNumKeineZielTabelle - vbObjectError, , ErrMsgKeineZielTabelle
    'End If
    'If (oInfotraeger Is Nothing) Then
    '  ErrMessage = ErrMsgInfoTraeger
    '  Err.Raise vbObjectError + ErrNumInfoTraeger, , "Zuerst die erste Datenzeile ('Infoträger') festlegen!"
    'End If
    'Set ointersect = Intersect(Selection, oInfotraeger)
    'If (ointersect Is Nothing) Then
    '  ErrMessage = Message1
    '  Err.Raise vbObjectError + ErrNumFormel, , "Die aktive Markierung liegt nicht innerhalb der ersten Datenzeile!"
    'End If
    'If (Union(Selection, Selection).Address <> ointersect.Address) Then
    '  ErrMessage = Message1
    '  Err.Raise vbObjectError + ErrNumFormel, , "Die aktive Markierung liegt nicht innerhalb der ersten Datenzeile!"
    'End If
    'On Error Resume Next
    '
    'If (ExistsLokalerZellname(strFormel)) Then LoeschenZellname(strFormel)
    'If (Err) Then GoTo Fehler
    'Set benannteZellen = oZielTabelle.Parent.Names
    ''If ((InStr(oZielTabelle.Parent.Name, " ") > 0) Or (InStr(oZielTabelle.Name, " ") > 0)) Then
    'Feldname = "'" & oZielTabelle.Name & "'!" & strFormel
    ''Else
    ''  Feldname = oZielTabelle.Name & "!" & strFormel
    ''End If
    'strBezug = "=" & Union(Selection, Selection).Address(RowAbsolute:=True, ColumnAbsolute:=True, ReferenceStyle:=xlA1, External:=True)
    'benannteZellen.Add Name:=Feldname, RefersTo:=strBezug
    'Set benannteZellen = Nothing
    'Exit Sub
    '
    'Fehler:
    'FehlerNachricht "CtabTabelle.Selection2Formel()"
End Sub



Public Sub Mod_FehlerVerbesserung()
    
    ' ******************************************************************    
    ' *** CdatExpim kann bislang nur in die AKTIVE Tabelle schreiben ***
    ' ******************************************************************    
    
    '' Manipulation der Tabellendaten: Berechnen von Fehlern und Verbesserungen.
    'On Error GoTo Fehler
    'Dim oExpim    As CdatExpim
    'Set oExpim = New CdatExpim
    'oExpim.GetQuelldaten_XlTab(oZielTabelle)
    'oExpim.Opt_FormelnErhalten = Me.ModOpt_FormelnErhalten
    'oExpim.Datenpuffer.Opt_VorhWerteUeberschreiben = Me.ModOpt_VorhWerteUeberschreiben
    'oExpim.Datenpuffer.Mod_FehlerVerbesserung
    'oExpim.SchreibeDatenpuffer_XlTabAktiv
    'call ClearStatusBarDelayed(StatusBarClearDelay)
    'Set oExpim = Nothing
    'Exit Sub
    '
    'Fehler:
    'Set oExpim = Nothing
    'FehlerNachricht "CtabTabelle.Mod_FehlerVerbesserung()"
End Sub



Public Sub Mod_UeberhoehungAusBemerkung()
    
    ' ******************************************************************    
    ' *** CdatExpim kann bislang nur in die AKTIVE Tabelle schreiben ***
    ' ******************************************************************    
    
    '' Manipulation der Tabellendaten: Berechnen von Fehlern und Verbesserungen.
    'On Error GoTo Fehler
    'Dim oExpim    As CdatExpim
    'Set oExpim = New CdatExpim
    'oExpim.GetQuelldaten_XlTab(oZielTabelle)
    'oExpim.Opt_FormelnErhalten = Me.ModOpt_FormelnErhalten
    'oExpim.Datenpuffer.Opt_VorhWerteUeberschreiben = Me.ModOpt_VorhWerteUeberschreiben
    'oExpim.Datenpuffer.Mod_UeberhoehungAusBemerkung
    'oExpim.SchreibeDatenpuffer_XlTabAktiv
    'call ClearStatusBarDelayed(StatusBarClearDelay)
    'Set oExpim = Nothing
    'Exit Sub
    '
    'Fehler:
    'Set oExpim = Nothing
    'FehlerNachricht "CtabTabelle.Mod_UeberhoehungAusBemerkung()"
End Sub



Public Sub Mod_Transfo_Tk2Gls()
    
    ' ******************************************************************    
    ' *** CdatExpim kann bislang nur in die AKTIVE Tabelle schreiben ***
    ' ******************************************************************    
    
    '' Manipulation der Tabellendaten: Transfo' Trassenkoo' => Gleissystem.
    'On Error GoTo Fehler
    'Dim oExpim    As CdatExpim
    'Set oExpim = New CdatExpim
    'oExpim.GetQuelldaten_XlTab(oZielTabelle)
    'oExpim.Opt_FormelnErhalten = Me.ModOpt_FormelnErhalten
    'oExpim.Datenpuffer.Opt_VorhWerteUeberschreiben = Me.ModOpt_VorhWerteUeberschreiben
    '
    '' Basis für Überhöhung: Auf Umwegen ermitteln...
    'ThisWorkbook.Metadaten.Update oPrjLocal:=Me.GetProjektDaten, oExtraLocal:=Me.GetExtraDaten
    'oExpim.Datenpuffer.TkBasisUeberhoehung = ThisWorkbook.Metadaten.TraBasisUeb
    '
    'oExpim.Datenpuffer.Mod_Transfo_Tk2Gls
    'oExpim.SchreibeDatenpuffer_XlTabAktiv
    '
    '' Basis für Überhöhung: in die Tabelle schreiben (falls geändert wegen Ungültigkeit).
    'Me.SchreibeMetaDaten
    '
    'call ClearStatusBarDelayed(StatusBarClearDelay)
    'Set oExpim = Nothing
    'Exit Sub
    '
    'Fehler:
    'Set oExpim = Nothing
    'FehlerNachricht "CtabTabelle.Mod_Transfo_Tk2Gls()"
End Sub



Public Sub Mod_Transfo_Gls2Tk()
    
    ' ******************************************************************    
    ' *** CdatExpim kann bislang nur in die AKTIVE Tabelle schreiben ***
    ' ******************************************************************    
    
    '' Manipulation der Tabellendaten: Transfo' Gleissystem => Trassenkoo'.
    'On Error GoTo Fehler
    'Dim oExpim    As CdatExpim
    'Set oExpim = New CdatExpim
    'oExpim.GetQuelldaten_XlTab(oZielTabelle)
    'oExpim.Opt_FormelnErhalten = Me.ModOpt_FormelnErhalten
    'oExpim.Datenpuffer.Opt_VorhWerteUeberschreiben = Me.ModOpt_VorhWerteUeberschreiben
    '
    '' Basis für Überhöhung: Auf Umwegen ermitteln...
    'ThisWorkbook.Metadaten.Update oPrjLocal:=Me.GetProjektDaten, oExtraLocal:=Me.GetExtraDaten
    'oExpim.Datenpuffer.TkBasisUeberhoehung = ThisWorkbook.Metadaten.TraBasisUeb
    '
    'oExpim.Datenpuffer.Mod_Transfo_Gls2Tk
    'oExpim.SchreibeDatenpuffer_XlTabAktiv
    '
    '' Basis für Überhöhung: in die Tabelle schreiben (falls geändert wegen Ungültigkeit).
    'Me.SchreibeMetaDaten
    '
    'call ClearStatusBarDelayed(StatusBarClearDelay)
    'Set oExpim = Nothing
    'Exit Sub
    '
    'Fehler:
    'Set oExpim = Nothing
    'FehlerNachricht "CtabTabelle.Mod_Transfo_Gls2Tk()"
End Sub



Sub Selection2Interpolationsformel()
    
    ' *** Selection gibt es nur in der AKTIVEN Tabelle ***
    
    '' Aufgrund der aktuellen Zellauswahl wird eine Interpolationsformel erstellt.
    '' Die Zellauswahl muß genau 3 Teilbereiche mit jeweils genau 1 Zelle enthalten.
    '' Zwei der gewählten Zellen müssen zur gleichen Spalte gehören und markieren
    '' Anfangs- und Endwert der bekannten Größe (z.B. Station).
    '' Die dritte Zelle muß in der Spalte der zu interpolierenden Größe liegen.
    '' Die dazugehörigen Anfangs- und Endwerte werden den Zellen dieser Spalte
    '' entnommen, die in den gleichen Zeilen stehen wie die entsprechenden Werte
    '' der bekannten Größe.
    '' ==> In die dritte Zelle wird die Formel eingetragen.
    '
    'Dim oAuswahl              As Range
    'Dim oZellen               As Areas
    '
    'Dim Selection_OK          As Boolean
    'Dim k                     As Integer
    'Dim i                     As Integer
    'Dim ii                    As Integer
    'Dim idxInt                As Integer
    'Dim idxAnf                As Integer
    'Dim idxEnd                As Integer
    '
    'Dim AdrBekAnf             As String
    'Dim AdrBekEnd             As String
    'Dim AdrBekZwi             As String
    'Dim AdrIntAnf             As String
    'Dim AdrIntEnd             As String
    'Dim Formel                As String
    '
    'Dim Message               As String
    '
    'Message = "Die Zellauswahl muß genau 3 einzelne Zellen enthalten." & vbNewLine & vbNewLine & _
    '          "Zwei der gewählten Zellen müssen in der gleichen Spalte liegen und markieren " & _
    '          "Anfangs- und Endwert der bekannten Größe (z.B. Station). " & _
    '          "Die dritte Zelle muß in der Spalte der zu interpolierenden Größe liegen. " & _
    '          "Die dazugehörigen Anfangs- und Endwerte werden den Zellen dieser Spalte " & _
    '          "entnommen, die in den gleichen Zeilen stehen wie die entsprechenden Werte " & _
    '          "der bekannten Größe. " & vbNewLine & vbNewLine & _
    '          "==> In die dritte Zelle wird die Formel eingetragen."
    '
    'On Error GoTo Fehler
    '
    'If (Not (oZielTabelle.Type = xlWorksheet)) Then
    '    Err.Raise vbObjectError + ErrNumKeineZielTabelle - vbObjectError, , ErrMsgKeineZielTabelle
    'End If
    '
    'Selection_OK = True
    'idxInt = 0
    'Set oAuswahl = Union(Selection, Selection)
    'If (oAuswahl.Areas.Count <> 3) Then
    '    Selection_OK = False
    'Else
    '    Set oZellen = oAuswahl.Areas
    '    For i = 1 To 3
    '        If (oZellen(i).Cells.Count <> 1) Then Selection_OK = False
    '    Next
    '    If (Selection_OK) Then
    '        ' Genau 3 einzelne Zellen markiert.
    '        ' Zwei Zellen in der gleichen Spalte suchen.
    '        For i = 1 To 3
    '            If (i > 1) Then ii = i - 1 Else ii = 3
    '            If (ii > 1) Then k = ii - 1 Else k = 3
    '            If (oZellen(i).Column = oZellen(ii).Column) Then
    '                ' Zwei Zellen in der gleichen Spalte gefunden.
    '                If (oZellen(i).Row > oZellen(ii).Row) Then
    '                    idxAnf = ii
    '                    idxEnd = i
    '                Else
    '                    idxAnf = i
    '                    idxEnd = ii
    '                End If
    '                idxInt = k
    '            End If
    '        Next
    '        If (idxInt <> 0) Then
    '            ' Prüfen, ob Position der 3. Zelle i.O. ist.
    '            If ((oZellen(idxInt).Row <= oZellen(idxAnf).Row) Or (oZellen(idxInt).Row >= oZellen(idxEnd).Row)) Then
    '                Selection_OK = False
    '            End If
    '        Else
    '            Selection_OK = False
    '        End If
    '    End If
    'End If
    '
    'If (Not Selection_OK) Then
    '    ErrMessage = Message
    '    Err.Raise vbObjectError + ErrNumInterpol, , "Fehlerhafte Zellauswahl zum Erstellen einer Interpolationsformel."
    'End If
    '
    '
    '' Auswahl ist in Ordnung, jetzt kann's tatsächlich losgehen...
    'AdrBekAnf = oZellen(idxAnf).Address(RowAbsolute:=True, ColumnAbsolute:=True)
    'AdrBekEnd = oZellen(idxEnd).Address(RowAbsolute:=True, ColumnAbsolute:=True)
    'AdrBekZwi = Cells(oZellen(idxInt).Row, oZellen(idxAnf).Column).Address(RowAbsolute:=False, ColumnAbsolute:=True)
    'AdrIntAnf = Cells(oZellen(idxAnf).Row, oZellen(idxInt).Column).Address(RowAbsolute:=True, ColumnAbsolute:=False)
    'AdrIntEnd = Cells(oZellen(idxEnd).Row, oZellen(idxInt).Column).Address(RowAbsolute:=True, ColumnAbsolute:=False)
    '
    'Formel = "=" & AdrIntAnf & "+(" & AdrIntEnd & "-" & AdrIntAnf & ")*((" & AdrBekZwi & "-" & AdrBekAnf & ")/(" & AdrBekEnd & "-" & AdrBekAnf & "))"
    ''MsgBox "Formel='" & Formel & "'"
    'oZellen(idxInt).Formula = Formel
    '
    'Set oZellen = Nothing
    'Set oAuswahl = Nothing
    'Exit Sub
    '
    'Fehler:
    'FehlerNachricht "CtabTabelle.Selection2Interpolationsformel()"
End Sub



Sub Selection2MarkDoppelteWerte()
    
    ' *** Selection gibt es nur in der AKTIVEN Tabelle ***
    
    '' Die markierten (und alle darunter liegenden) Zellen werden mit einer bedingten
    '' Formatierung versehen, die alle Zellen mit solchen Werten hervorhebt, die in
    '' derselben Spalte mehr als einmal vorkommen.
    '
    'Dim oAuswahl               As Range
    'Dim oBereich               As Range
    'Dim oRangeQuelle           As Range
    'Dim oRangeZiel             As Range
    'Dim oBedFormate            As FormatConditions
    '
    'Dim StatusScreen           As Boolean
    'Dim StatusCalc             As Boolean
    ''Dim oStatusAktZelle        As Range
    ''Dim oStatusAuswahl         As Range
    '
    'Dim Titel                  As String
    'Dim Message                As String
    'Dim Message1               As String
    'Dim Buttons                As Integer
    'Dim Formel                 As String
    'Dim strBezugQuelleRel      As String
    'Dim strBezugVorgAbs        As String
    '
    'Dim i                      As Long
    'Dim Spalte                 As Long
    'Dim ZeAnf                  As Long
    'Dim ZeEnd                  As Long
    'Dim ZeAnz                  As Long
    '
    'Message1 = "Die markierte und alle in derselben Spalte darunter (!) liegenden Zellen werden mit einer " & _
    '           "bedingten Formatierung versehen, die mehrfach in dieser Spalte auftretende Werte hervorhebt. " & _
    '           "Normalerweise sollte dafür die 1. Datenzelle einer Spalte markiert sein. " & _
    '           "Die Markierung darf frühestens in Zeile 2 beginnen!" & vbNewLine & vbNewLine & _
    '           "==> Vorhandene Bedingte Formate werden gelöscht! "
    '
    'On Error GoTo Fehler
    '
    'If (Not (oZielTabelle.Type = xlWorksheet)) Then
    '    Err.Raise vbObjectError + ErrNumKeineZielTabelle - vbObjectError, , ErrMsgKeineZielTabelle
    'End If
    'Set oAuswahl = Union(Selection, Selection)
    '' Unergründlichen Formel-Fehler abfangen, der auftritt, wenn verschiedene Bereiche nicht
    '' in der gleichen Zeile beginnen (Microsoft'sche Eigenintelligenz?)
    'If (oAuswahl.Areas.Count > 1) Then
    '    Err.Raise vbObjectError + ErrNumBereiche, , "Aktion ist nicht möglich. Es darf nur ein einzelner Bereich markiert sein."
    'End If
    '' Status merken.
    'StatusScreen = oZielTabelle.Application.ScreenUpdating
    'StatusCalc = oZielTabelle.EnableCalculation
    'oZielTabelle.Application.ScreenUpdating = False
    'oZielTabelle.EnableCalculation = False
    ''Set oStatusAktZelle = ActiveCell
    ''Set oStatusAuswahl = Selection
    '
    'On Error Resume Next
    '
    'Message = Message1 & vbNewLine & vbNewLine & "Sollen in der gewählten Spalte doppelte Werte markiert werden?"
    'Buttons = vbYesNo + vbQuestion + vbDefaultButton1
    'Titel = "Doppelte Werte markieren"
    'If (MsgBox(Message, Buttons, Titel) = vbYes) Then
    '    ' Aktion wirklich durchführen ;-)
    '    If (Err) Then GoTo Fehler
    '    For Each oBereich In oAuswahl.Areas
    '        For Spalte = oBereich.Columns.Column To oBereich.Columns(oBereich.Columns.Count).Column
    '            ' Für 1. bis letzte Spalte des Einzelbereiches
    '            ZeAnf = oBereich.Rows.Row
    '            If (ZeAnf < 2) Then
    '                Message = "Spalte " & Spalte & " wird ignoriert, da Markierung in Zeile 1 beginnt!"
    '                Buttons = vbOK
    '                Titel = "Nicht unterstützte Markierung"
    '                MsgBox Message, Buttons, Titel
    '            Else
    '                ZeEnd = Cells(Rows.Count, Spalte).End(xlUp).Row
    '                ZeAnz = ZeEnd - ZeAnf + 1
    '                ' Auch markierte Zelle einer leeren Spalte formatieren.
    '                If (ZeAnz < 1) Then ZeAnz = 1
    '                
    '                Set oRangeQuelle = Cells(ZeAnf, Spalte)
    '                Set oRangeZiel = oRangeQuelle.Resize(ZeAnz)
    '                Set oBedFormate = oRangeZiel.FormatConditions
    '                
    '                ' Vorhandene Bedingte Formate löschen.
    '                For i = oBedFormate.Count To 1 Step -1
    '                    oBedFormate.Item(i).Delete
    '                Next
    '                i = 0
    '                strBezugQuelleRel = oRangeQuelle.Address(RowAbsolute:=False, ColumnAbsolute:=True, ReferenceStyle:=xlA1, External:=False)
    '                
    '                ' Vergleich mit darüberliegenden Zellen (geht nicht in Zeile 1).
    '                strBezugVorgAbs = oRangeQuelle.Offset(rowoffset:=-1, columnoffset:=0).Address(RowAbsolute:=True, ColumnAbsolute:=True, ReferenceStyle:=xlA1, External:=False)
    '                'MsgBox "Spalte=" & Spalte & "   Adresse=" & strBezugQuelleRel
    '                Formel = "=ISTZAHL(VERGLEICH(" & strBezugQuelleRel & ",BEREICH.VERSCHIEBEN(" & strBezugVorgAbs & ",0,0,Zeile(" & strBezugQuelleRel & ")-" & CStr(ZeAnf - 1) & ",1),0))"
    '                oBedFormate.Add Type:=xlExpression, Formula1:=Formel
    '                i = i + 1
    '                With oBedFormate.Item(i)
    '                  .Interior.ColorIndex = 24
    '                End With
    '                    
    '                ' Vergleich mit darunterliegenden Zellen.
    '                Formel = "=ISTZAHL(VERGLEICH(" & strBezugQuelleRel & ",BEREICH.VERSCHIEBEN(" & strBezugQuelleRel & ",1,0," & Rows.Count & "-ZEILE(" & strBezugQuelleRel & "),1),0))"
    '                oBedFormate.Add Type:=xlExpression, Formula1:=Formel
    '                i = i + 1
    '                With oBedFormate.Item(i)
    '                    .Font.ColorIndex = 2
    '                    .Interior.ColorIndex = 32
    '                End With
    '            End If
    '        Next
    '    Next
    'End If
    '
    '' Alten Zustand wiederherstellen
    ''oStatusAuswahl.Select
    ''oStatusAktZelle.Activate
    '
    'oZielTabelle.Application.ScreenUpdating = StatusScreen
    'oZielTabelle.EnableCalculation = StatusCalc
    ''Set oStatusAuswahl = Nothing
    ''Set oStatusAktZelle = Nothing
    'Set oAuswahl = Nothing
    'Set oBereich = Nothing
    'Set oRangeQuelle = Nothing
    'Set oRangeZiel = Nothing
    'Set oBedFormate = Nothing
    'Exit Sub
    '
    'Fehler:
    'oZielTabelle.Application.ScreenUpdating = StatusScreen
    'If (Not oZielTabelle Is Nothing) Then oZielTabelle.EnableCalculation = StatusCalc
    ''Set oStatusAuswahl = Nothing
    ''Set oStatusAktZelle = Nothing
    'Set oAuswahl = Nothing
    'Set oBereich = Nothing
    'Set oRangeQuelle = Nothing
    'Set oRangeZiel = Nothing
    'Set oBedFormate = Nothing
    'FehlerNachricht "CtabTabelle.Selection2MarkDoppelteWerte()"
End Sub



Private Sub LoeschenZellname(ByVal Name As String)
    ' Löscht den benannten Zellbereich "Name".
    On Error Resume Next
    Dim oNameRange As Range
    Set oNameRange = GetLokalerZellname(Name, oZielTabelle)
    If (Not oNameRange Is Nothing) Then
        oNameRange.Name.Delete
        If (Err) Then
            Err.Raise vbObjectError + ErrNumLoeschenZellname, , "Fehler beim Löschen des benannten Bereiches '" & Name & "'."
        End If
    End If
    Exit Sub
    
    Fehler:
    FehlerNachricht "CtabTabelle.LoeschenZellname()"
End Sub



Private Sub AuswahlZellname(ByVal Name As String)
    
    ' *** Selection gibt es nur in der AKTIVEN Tabelle ***
    
    '' Wählt den benannten Zellbereich "Name" aus (Selection).
    ''oZielTabelle.Application.Goto (benannteZellen(i).RefersToR1C1)
    'On Error Resume Next
    'Dim oNameRange As Range
    'Set oNameRange = GetLokalerZellname(Name, oZielTabelle)
    'If (Not oNameRange Is Nothing) Then
    '    oNameRange.Select
    '    If (Err) Then
    '        Err.Raise vbObjectError + ErrNumAuswahlZellname, , "Fehler bei der Auswahl des benannten Bereiches '" & Name & "'."
    '    End If
    'End If
    'Exit Sub
    '
    'Fehler:
    'FehlerNachricht "CtabTabelle.AuswahlZellname()"
End Sub



Private Sub BestimmeDatenbereich()
    '------------------------------------------------------------------------------------------------
    ' Position und Ausdehnung des Datenbereiches der Tabelle werden bestimmt
    ' und in folgenden globalen Variablen zwecks Abruf per Eigenschaften abgelegt:
    ' 
    '    oInfoTraeger        : eben dieser als Range-Objekt     oder 'Nothing'
    '    lngErsteDatenZeile  : Anfangszeile  des Datenbereiches oder -1
    '    lngErsteDatenSpalte : Anfangsspalte des Datenbereiches oder -1
    '    lngLetzteDatenSpalte: Endspalte     des Datenbereiches oder -1
    ' 
    ' => Die Endzeile des Datenbereiches wird immer in Echtzeit bestimmt 
    '    bei Verwendung von Me.LetzteDatenZeile
    ' 
    ' Hinweise: Der Datenbereich wird zunächst via "UsedRange" bestimmt.
    '           Folgende Faktoren, wenn vorhanden, schränken diesen Bereich anschließend ein:
    '           - Bereich "Daten.ErsteZelle" überschreibt Anfangszeile und -spalte
    '           - Bereich "Daten.InfoTraeger" überschreibt Anfangszeile, -spalte und Endspalte.
    '------------------------------------------------------------------------------------------------
    Dim oUsedRange    As Range
    Dim oErsteZelle   As Range
    
    DebugEcho "CtabTabelle.BestimmeDatenbereich(): Versuche, Informationen zur Tabelle zu bestimmen."
    
    If (Not (oZielTabelle.Type = xlWorksheet)) Then
        DebugEcho "CtabTabelle.BestimmeDatenbereich(): Keine Excel-Tabelle angegeben!"
        Set oInfotraeger = Nothing
        lngErsteDatenZeile   = -1
        lngErsteDatenSpalte  = -1
        lngLetzteDatenSpalte = -1
    else
        ' 1. Infotraeger suchen.
        DebugEcho "CtabTabelle.BestimmeDatenbereich(): Suche InfoTraeger der Tabelle '" & oZielTabelle.Name & "'"
        Set oInfotraeger = GetLokalerZellname(strInfoTraeger, oZielTabelle)
        If (oInfotraeger Is Nothing) Then
            DebugEcho "CtabTabelle.BestimmeDatenbereich(): InfoTraeger nicht gefunden in Tabelle '" & oZielTabelle.Name & "'"
        else
            DebugEcho "CtabTabelle.BestimmeDatenbereich(): Bereich des InfoTraegers = " & oInfotraeger.Address
        End If
        
        ' 2. Allgemeingültigen Bereich bestimmen.
        set oUsedRange       = oZielTabelle.UsedRange
        lngErsteDatenZeile   = oUsedRange.Rows(1).Row
        lngErsteDatenSpalte  = oUsedRange.Columns(1).Column
        lngLetzteDatenSpalte = oUsedRange.Columns(oUsedRange.Columns.Count).Column
        
        ' 3. Falls Infotraeger oder Markierung für "ErsteZelle" vorhanden.
        if (not oInfotraeger is nothing) then
            lngErsteDatenZeile   = oInfotraeger.Rows(1).Row
            lngErsteDatenSpalte  = oInfotraeger.Columns(1).Column
            lngLetzteDatenSpalte = oInfotraeger.Columns(oInfotraeger.Columns.Count).Column
        else
            set oErsteZelle = GetLokalerZellname(strErsteZelle, oZielTabelle)
            If (oErsteZelle Is Nothing) Then
                DebugEcho "CtabTabelle.BestimmeDatenbereich(): Bereich '" & strErsteZelle & "' auch nicht gefunden in Tabelle '" & oZielTabelle.Name & "'"
            else
                DebugEcho "CtabTabelle.BestimmeDatenbereich(): Bereich '" & strErsteZelle & "' = " & oErsteZelle.Address
                lngErsteDatenZeile  = oErsteZelle.Rows(1).Row
                lngErsteDatenSpalte = oErsteZelle.Columns(1).Column
            end if
        end if
          
        ' Sicherheit: Wichtig, wenn Infotraeger vorhanden. Kann sonst auch nicht schaden.
        If (lngLetzteDatenSpalte < lngErsteDatenSpalte) Then lngLetzteDatenSpalte = lngErsteDatenSpalte
        
        DebugEcho "CtabTabelle.BestimmeDatenbereich(): Datenbereich: " & cStr(lngErsteDatenSpalte) & ". - " & cStr(lngLetzteDatenSpalte) & ". Spalte.  " _
                   & cStr(lngErsteDatenZeile) & ". - " & cStr(Me.LetzteDatenZeile) & ". Zeile"
        ' Aufräumen
        Set oUsedRange = Nothing
        Set oErsteZelle = Nothing
    end if
End Sub


Public Sub Syncronisieren()
    ' Syncronisiert die Daten einiger Eigenschaften mit der Tabelle.
    '  - Datenbereich
    '  - Spaltennamen und -Einheiten
    '  - Kategorien
    '  - Verfügbarkeit von Methoden signalisieren
    On Error GoTo Fehler
    DebugEcho "CtabTabelle.Syncronisieren(): Syncronisiere Tabelle: '" & Me.TabName & "'."
    
    ' 1. Position und Ausdehnung des Datenbereiches bestimmen.
    Call BestimmeDatenbereich()
    
    ' 2. Benannte Spalten der Tabelle erfassen.
    Call GetSpaltenInfo(PrefixSpaltenname, oSpaltenErsteZellen, oSpaltenEinheiten, _
                        oSpaltenFormate, oSpaltenBuchst2Name)
      
    ' 3. Kategorien konfigurierter Spaltennamen ermitteln.
    strKategorien = GetKategorien()
    
    Call Me.ZeigeSpalten()
    Exit Sub
    
    Fehler:
    FehlerNachricht "CtabTabelle.Syncronisieren()"
End Sub


Private Sub GetSpaltenInfo(ByVal Prefix As String, oZellen As Scripting.Dictionary, _
                           oZellEinheiten As Scripting.Dictionary, _
                           oZellFormate As Scripting.Dictionary, _
                           oBuchst2Name As Scripting.Dictionary)
    ' Findet alle benannten Zellen der Arbeitsmappe, die sich auf die angegebene Tabelle
    ' beziehen und deren Namen mit "Prefix" beginnen. Enthält der Spaltenname auch
    ' eine Einheitenangabe, so wird diese extrahiert.
    ' Eingabeparameter:
    '   Prefix         ... Es werden nur Namen berücksichtigt, die damit beginnen.
    ' Ausgabeparameter:
    '   oZellen        ... Alle gefundenen Zellen als Range.
    '   oZellEinheiten ... Die Einheit  zu jeder gefundenen Zelle, falls verfügbar; sonst="ohne".
    '   oZellFormate   ... Formatangabe zu jeder gefundenen Zelle, falls verfügbar; sonst="ohne".
    '   ==> Der Key der 3 Dictionaries ist der reine Zellname ohne Prefix und Einheitenangabe.
    '   ==> Alle Dictionaries werden zunächst geleert!
    
    
    'On Error Resume Next
    On Error GoTo Fehler
    
    Dim benannteZellen    As Names
    Dim oRange            As Range
    Dim Zellname          As String
    Dim ZellnamePur       As String
    Dim ZellEinheit       As String
    Dim ZellBuchstabe     As String
    Dim RegExPrefix       As String
    Dim RegExZellname     As String
    Dim i                 As Long
    Dim NF                As Long
    Dim Feld()            As String
    
    oZellen.RemoveAll
    oZellEinheiten.RemoveAll
    oZellFormate.RemoveAll
    oBuchst2Name.RemoveAll
    
    RegExPrefix = FileSpec2RegExp(Prefix)
    RegExZellname = "^(.*!)?" & RegExPrefix    'regulärer Ausdruck für einen Zellnamen
    
    Set oRange = Nothing
    If (oZielTabelle.Type = xlWorksheet) Then
        Set benannteZellen = oZielTabelle.Parent.Names
        i = 1
        Do While (i <= benannteZellen.Count)
            Zellname = benannteZellen(i).Name
            If (isLokalerZellName(benannteZellen(i), oZielTabelle)) Then
                ' Bereichsname existiert im Tabellenblatt.
                ' ZellName kann trotzdem den Tabellennamen enthalten. Dies steuert Excel nach eigenem Willen...
                If (entspricht(RegExZellname, Zellname)) Then
                    ZellnamePur = substitute(RegExZellname, "", Zellname, False, False)
                    NF = SplitDelim(ZellnamePur, Feld, TrennerEinheit)
                    If (NF > 1) Then
                        ZellnamePur = Feld(1)
                        ZellEinheit = Feld(NF)
                    Else
                        ZellEinheit = "ohne"
                    End If
                    Set oRange = oZielTabelle.Application.Range(benannteZellen(i).RefersTo)
                    oZellen.Add ZellnamePur, oRange
                    oZellEinheiten.Add ZellnamePur, ZellEinheit
                    oZellFormate.Add ZellnamePur, oRange.NumberFormatLocal
                    'oZellFormate.Add ZellnamePur, oRange.NumberFormat
                    ZellBuchstabe = UCase(MidStr(oRange.Address, "$", "$", False))
                    oBuchst2Name.Add ZellBuchstabe, ZellnamePur
                End If
            End If
            i = i + 1
        Loop
        Set benannteZellen = Nothing
    End If
    Set oRange = Nothing
    Exit Sub
    
    Fehler:
    Set oRange = Nothing
    FehlerNachricht "CtabTabelle.GetSpaltenInfo()"
End Sub


Private Function GetKategorien() As String
    ' Ermittelt alle unterschiedlichen Spalten-Kategorien der angegebene Tabelle und
    ' den Kodenamen der Tabelle als erste Kategorie.
    ' Rückgabe: Liste durch Semikolons getrennt.
    
    On Error GoTo Fehler
    
    Dim DictTmp           As Scripting.Dictionary
    Dim Spalte            As Variant
    Dim Liste             As String
    Dim ListeKeys         As String
    
    Liste = Me.TabKlasse
    Set DictTmp = New Scripting.Dictionary
    
    If (Not (oSpaltenErsteZellen Is Nothing)) Then
        On Error Resume Next
        For Each Spalte In oSpaltenErsteZellen
            If (ThisWorkbook.Konfig.SpaltenKategorie(Spalte) <> "") Then DictTmp.Add ThisWorkbook.Konfig.SpaltenKategorie(Spalte), "*"
        Next
        On Error GoTo 0
        ListeKeys = ListeDerKeys(DictTmp)
        If (ListeKeys <> "") Then Liste = Liste & ";" & ListeKeys
    End If
    Set DictTmp = Nothing
    GetKategorien = Liste
    Exit Function
    
    Fehler:
    GetKategorien = ""
    Set DictTmp = Nothing
    FehlerNachricht "CtabTabelle.GetKategorien()"
End Function


Public Sub ZeigeSpalten()
    ' Schreibt für Kontrollzwecke alle gefundenen Spalten der Tabelle ins Debug-Protokoll.
    Dim Spalte As Variant
    For Each Spalte In Me.SpaltenErsteZellen
        DebugEcho "Spalte=" & Spalte & vbTab & vbTab & _
                  "Einheit=" & Me.SpaltenEinheiten(Spalte) & vbTab & vbTab & _
                  "Zelle=" & Me.SpaltenErsteZellen(Spalte).Address
    Next
    DebugEcho "==> Kategorien = '" & Me.Kategorien & "'" & vbNewLine
End Sub



' for jEdit:  :folding=indent::collapseFolds=1::tabSize=4::indentSize=4:
