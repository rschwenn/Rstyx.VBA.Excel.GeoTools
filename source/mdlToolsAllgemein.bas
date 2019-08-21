Attribute VB_Name = "mdlToolsAllgemein"
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2003 - 2019  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'====================================================================================
'Modul mdlToolsAllgemein
'====================================================================================
'Stellt allgemeine Werkzeuge in erster Linie zur Programmierung zur Verfügung.
'Einige Funktionen sind auch für den Einsatz in Tabellenblättern sinnvoll einsetzbar.


Option Explicit


'Konstanten zur Steuerung der Array-Sortierung
const SORTTYPE_NUMERIC As Byte = 0  'Numerisch, wenn möglich, sonst Alphanumerisch
const SORTTYPE_STRING  As Byte = 1  'Alphanumerisch



'***  Abteilung Allgemeines  **********************************************************************

Public Sub Wait(X)
  'wartet x Sekunden
  Dim Start, i
  Start = CInt(Second(Time))
  i = 0
  Do
    i = i + 1
  Loop While (CInt(Second(Time)) < Start + CInt(X))
End Sub



'***  Abteilung Dateien  **************************************************************************

Public Function NameExt(ByVal Pfad As String, ByVal MitOhneExt As String)
  'Gibt den Dateinamen ohne Pfad zurück. Ob die Extension mit zurückgegeben
  'wird, hängt vom zweiten Parameter ab
  'Parameter "Pfad":       voll qualifizierter Dateiname
  'Parameter "MitOhneExt": wenn "mitext", dann enthält der Funktionswert
  '                        auch die in "Pfad" angegebene Extension, sonst nicht
  On Error GoTo Fehler
  
  Dim lastBackslashAt  As String
  Dim lastPointAt      As String
  
  If (Len(Pfad) > 0) Then
    lastBackslashAt = InStrRev(Pfad, "\", -1, vbTextCompare)
    lastPointAt = InStrRev(Pfad, ".", -1, vbTextCompare)
    'if lastBackslashAt=0 then lastBackslashAt=1-1
    If (lastPointAt = 0 Or MitOhneExt = "mitext") Then lastPointAt = Len(Pfad) + 1
    NameExt = Mid(Pfad, lastBackslashAt + 1, lastPointAt - lastBackslashAt - 1)
  End If
  
  Exit Function
  
Fehler:
  FehlerNachricht "mdlToolsAllgemein.NameExt()"
End Function


Public Function Verz(ByVal Pfad As String)
  'Funktionswert:    Verzeichnisname ohne Backslash am Ende
  'Parameter "Pfad": voll qualifizierter Dateiname
  
  On Error GoTo Fehler
  
  Dim lastBackslashAt  As String
  Dim lastPointAt      As String
  
  Verz = ""
  If (Len(Pfad) > 1) Then
    lastBackslashAt = InStrRev(Pfad, "\", -1, vbTextCompare)
    If (lastBackslashAt > 0) Then Verz = Mid(Pfad, 1, lastBackslashAt - 1)
  End If
  
  Exit Function
  
Fehler:
  FehlerNachricht "mdlToolsAllgemein.Verz()"
End Function


Public Function VorName(ByVal Pfad As String)
  'Rückgabe: DateiName ohne Extension, d.h. ohne den letzten Punkt nebst folgender Zeichen,
  '                                    wenn diese kein Leerzeichen enthalten.
  'Eingabe:  DateiName ... mit/ohne Verzeichnis
  On Error GoTo Fehler
  
  Dim fs    As New Scripting.FileSystemObject
  VorName = fs.GetBaseName(Pfad)
  set fs = Nothing
  Exit Function
  
Fehler:
  set fs = Nothing
  FehlerNachricht "mdlToolsAllgemein.VorName()"
End Function


Public Function LastBackslashDelete(ByVal Verzeichnis As String)
  'Funktionswert:           Verzeichnisname ohne Backslash am Ende
  'Parameter "Verzeichnis": Verzeichnisname mit oder ohne abschließenden Backslash
  On Error GoTo Fehler
  
  Dim L As Long
  L = Len(Verzeichnis)
  If (L > 0) Then
    If (Mid(Verzeichnis, L) = "\") Then Verzeichnis = Mid(Verzeichnis, 1, L - 1)
  End If
  LastBackslashDelete = Verzeichnis
  
  Exit Function
  
Fehler:
  FehlerNachricht "mdlToolsAllgemein.LastBackslashDelete()"
End Function




'***  Abteilung Strings und Reguläre Ausdrücke  ***************************************************

Function Bool2String(ByVal bool As Boolean) As String
  'Liefert die der Variable entsprechende textliche Wahrheitsausage.
  If (bool) Then
    Bool2String = "ja"
  Else
    Bool2String = "nein"
  End If
End Function


Function String2Bool(ByVal text As String) As Boolean
  'Liefert den der textlichen Wahrheitsaussage entsprechenden boolschen Wert.
  'If (entspricht("ja", text)) Then
  If ((entspricht("ja", text)) Or (entspricht("true", text)) Or (entspricht("wahr", text)) _
      Or (entspricht("1", text))) Then
    String2Bool = True
  Else
    String2Bool = False
  End If
End Function


Public Function SplitDelim(ByVal sString As String, ByRef Feld, ByVal sDelim As String, _
                           Optional ByVal blnTrimField = false, Optional ByVal blnNoEmptyFields = false)
  ' ==================================================================
  ' Splittet sString in Teilstrings, die durch sDelim getrennt sind.
  ' Leerzeichen an Anfang und Ende der Teilstrings werden nicht entfernt.
  ' Wird kein Separator in sString gefunden, so wird Feld(1) = sString gesetzt.
  '   sString          : zu splittender String
  '   sDelim           : Delimiter
  '   blnTrimField     : Felder trimmen?
  '   blnNoEmptyFields : leere Felder übergehen?
  '
  '   Feld             : Rückgabefeld, enthält gefundene Wörter (Index = 1..max)
  '   Funktionswert    : Anzahl Wörter in Feld()
  ' ==================================================================
  Dim iPos      As Integer
  Dim iNextPos  As Integer
  Dim iDelimLen As Integer
  Dim iCount    As Integer
  Dim tmpField  As String
  
  iCount = 0
  If (Trim(sString) <> "") Then
    Erase Feld
    iDelimLen = Len(sDelim)
    iPos = 1
    iNextPos = InStr(sString, sDelim)
    
    Do While iNextPos > 0
      tmpField = Mid$(sString, iPos, (iNextPos - iPos))
      if (blnTrimField) then tmpField = trim(tmpField)
      
      if ((not blnNoEmptyFields) or (tmpField <> "")) then
        iCount = iCount + 1
        ReDim Preserve Feld(1 To iCount) As String
        Feld(iCount) = tmpField
      end if
      
      iPos = iNextPos + iDelimLen
      iNextPos = InStr(iPos, sString, sDelim)
    Loop
    
    tmpField = Mid$(sString, iPos)
    if (blnTrimField) then tmpField = trim(tmpField)
    
    if ((not blnNoEmptyFields) or (tmpField <> "")) then
      iCount = iCount + 1
      ReDim Preserve Feld(1 To iCount) As String
      Feld(iCount) = tmpField
    end if
  End If
  
  SplitDelim = iCount
End Function


Public Function LeftStr(sString As Variant, ByVal vDelimiter As Variant, Optional ByVal bDelimiter As Boolean = False) As String
  ' LeftStr ersetzt die Left-Funktion
  '
  ' vDelimiter gibt entweder die Position als Zahl an
  ' oder aber das gesuchte Zeichen, bis zu dem
  ' der Teilstring zurückgegeben werden soll.
  '
  ' bDelimiter legt fest, ob der Teilstring bis einschl.
  ' dem ersten Vorkommen des gesuchten Zeichens
  ' zurückgegeben werden soll (True) oder um ein Zeichen
  ' gekürzt (False = Standard)
  
  Dim lPos As Long
  
  If VarType(vDelimiter) = vbString Then
    lPos = InStr(sString, vDelimiter)
    If Not bDelimiter Then lPos = lPos - 1
    If lPos < 1 Then lPos = 0
  Else
    lPos = val(vDelimiter)
  End If
  
  LeftStr = Left(sString, lPos)
End Function


Public Function RightStr(sString As Variant, ByVal vDelimiter As Variant, Optional ByVal bDelimiter As Boolean = False) As String
  ' RightStr ersetzt die Right-Funktion
  '
  ' vDelimiter gibt entweder die Position als Zahl an
  ' oder aber das gesuchte Zeichen, ab dessen letztes
  ' Vorkommen der Teilstring zurückgegeben werden soll.
  '
  ' bDelimiter legt fest, ob der Teilstring einschl.
  ' dem gesuchten Zeichen zurückgegeben werden soll (True)
  ' oder um ein Zeichen nach rechts versetzt
  ' (False = Standard)
  
  Dim lPos As Long
  
  If VarType(vDelimiter) = vbString Then
    lPos = InStrRev(sString, vDelimiter)
    If lPos > 0 Then
      If Not bDelimiter Then lPos = lPos + 1
      RightStr = Mid(sString, lPos)
    End If
  Else
    lPos = val(vDelimiter)
    RightStr = Right(sString, lPos)
  End If
End Function


Public Function MidStr(sString As Variant, ByVal vDelimiter1 As Variant, Optional ByVal vDelimiter2 As Variant, Optional ByVal bDelimiter As Boolean = False) As String
  ' MidStr ersetzt die Mid-Funktion
  '
  ' vDelimiter1 gibt entweder die Position als Zahl an
  ' oder aber das gesuchte Zeichen, ab der der Teilstring
  ' zurückgegeben werden soll.
  '
  ' vDelimiter2 ist optional und bestimmt die Länge des
  ' Teilstrings, der zurückgegeben werden soll.
  ' Wird hier ein Zeichen angegeben, so wird der
  ' Teilstring ab vDelimiter1 und bis zum nächsten
  ' Vorkommen des in vDelimiter2 angegebenen Zeichens
  ' ermittelt.
  '
  ' bDelimiter legt fest, ob der Teilstring einschl.
  ' dem der in vDelimiter1 und vDelimiter2 angegebenen
  ' Zeichen zurückgegeben werden soll (True) oder um
  ' je ein Zeichen links und rechts gekürzt (False).
  
  Dim lPos As Long
  
  If VarType(vDelimiter1) = vbString Then
    lPos = InStr(sString, vDelimiter1)
    If lPos > 0 And Not bDelimiter Then lPos = lPos + 1
  Else
    lPos = val(vDelimiter1)
  End If
  
  If lPos > 0 Then
    If IsMissing(vDelimiter2) Then
      MidStr = Mid(sString, lPos)
    Else
      MidStr = LeftStr(Mid(sString, lPos), vDelimiter2, bDelimiter)
    End If
  End If
End Function


Function Konv_OEM2ANSI(ByVal text)
  'Konversion der deutschen Umlaute
  text = Replace(text, Chr(148), Chr(246))  'ö
  text = Replace(text, Chr(132), Chr(228))  'ä
  text = Replace(text, Chr(129), Chr(252))  'ü
  text = Replace(text, Chr(153), Chr(214))  'Ö
  text = Replace(text, Chr(142), Chr(196))  'Ä
  text = Replace(text, Chr(154), Chr(220))  'Ü
  text = Replace(text, Chr(225), Chr(223))  'ß
  Konv_OEM2ANSI = text
End Function


Function FileSpec2RegExp(ByVal Spec As String) As String
  'Steve Fulton
  'convert a filespec to a pattern used for Regular expressions
  Dim Pattern As String
  With ThisWorkbook.RegExp
    .Global = True
    .Pattern = "\\"
    Pattern = .Replace(Spec, "\\")
    .Pattern = "[.(){}[\]$^]"
    Pattern = .Replace(Pattern, "\$&")
    .Pattern = "\?"
    Pattern = .Replace(Pattern, ".")
    .Pattern = "\*"
    Pattern = .Replace(Pattern, ".*")
  End With
  'FileSpec2RegExp = "^" & Pattern & "$"
  FileSpec2RegExp = Pattern
End Function


Function splitWords(ByVal text As String, ByRef Feld As Variant, ByVal WordRegEx As String)
  'splittet einen String auf Grundlage des für die Wortsuche (!) angegebenen
  'regulären Ausdruckes und gibt die gefundenen Wörter in einem Array zurück
  'text      ... zu splittender String
  'Feld      ... Feld mit Ergebnissen (Wörtern)
  'WordRegEx ... reg. Ausdruck für die zu suchenden Wörter (nicht den Separator!),
  '              wenn WordRegEx="" dann wird verwendet "\S+"
  'Funktionswert = NF (Anzahl der gefundenen Felder)
  '              = -1 bei Fehler
  
  On Error GoTo Fehler
  
  Dim cWords  As Object
  'Dim ThisWorkbook.RegExp As Object
  Dim NF      As Long
  Dim i       As Long
  
  With ThisWorkbook.RegExp
    If (Trim(WordRegEx) = "") Then
      .Pattern = "\S+"
    Else
      .Pattern = WordRegEx
    End If
    .Global = True
    .IgnoreCase = False
    Set cWords = .Execute(text)
    NF = cWords.Count
  End With
  ReDim Feld(0 To NF)
  Feld(0) = ""
  For i = 1 To NF
    Feld(i) = cWords(i - 1)
  Next
  splitWords = NF
  
  Set cWords = Nothing
  'Set ThisWorkbook.RegExp = Nothing
  Exit Function
  
Fehler:
  FehlerNachricht "splitWords()"
  splitWords = -1
End Function


Function substitute(ByVal SuchMuster, ByVal Ersatzstring, ByVal Zeichenfolge, _
                    ByVal AlleErsetzen As Boolean, ByVal AbbruchBeiFehler As Boolean)
  'Funktionswert: Ergebniszeichenfolge
  'ersetzt das "Suchmuster" durch den "Ersatzstring" in der "Zeichenfolge"
  'ist "AbbruchBeiFehler"=true so erfolgt eine Meldung mit Abbruch der
  'Skriptverarbeitung, wenn das Suchmuster nicht gefunden wird.
  On Error GoTo Fehler
      
  'Dim ThisWorkbook.RegExp As Object
  'Set ThisWorkbook.RegExp = CreateObject("VBScript.RegExp")
  
  ThisWorkbook.RegExp.Pattern = SuchMuster       ' Setzt das Muster.
  ThisWorkbook.RegExp.IgnoreCase = True          ' Ignoriert die Schreibweise. (Namen in Excel sind nicht case sensitive!)
  ThisWorkbook.RegExp.Global = AlleErsetzen      ' Legt globales Anwenden fest.
  If ThisWorkbook.RegExp.test(Zeichenfolge) Then
    substitute = ThisWorkbook.RegExp.Replace(Zeichenfolge, Ersatzstring)   ' Führt die Ersetzung durch.
  Else
    If AbbruchBeiFehler Then
      ErrMessage = "Fehler beim Ersetzen:" & vbNewLine & vbNewLine & _
                   "Suchmuster '" & SuchMuster & "' nicht gefunden."
      GoTo Fehler
      'MsgBox ErrMessage, vbExclamation, "Fehler"
      'wscript.quit
    Else
      substitute = Zeichenfolge     ' keine Änderung
    End If
  End If
  'Set ThisWorkbook.RegExp = Nothing
  Exit Function

Fehler:
  FehlerNachricht "substitute()"
End Function


Function RegExpTest(SuchMuster, Zeichenfolge)
  'Gibt in einem String eine Erfolgsmeldung über das in Zeichenfolge
  'gefundene "Suchmuster" zurück.
  'Demonstriert die Anwendung von regulären Ausdrücken.
  On Error GoTo Fehler
  Dim Uebereinstimmung, Uebereinstimmungen, Rueckgabe   ' Erstellt Variablen.
  'Dim ThisWorkbook.RegExp As Object
  Rueckgabe = ""
  'Set ThisWorkbook.RegExp = CreateObject("VBScript.RegExp")
  ThisWorkbook.RegExp.Pattern = SuchMuster       ' Setzt das Muster.
  ThisWorkbook.RegExp.IgnoreCase = False         ' Ignoriert die Schreibweise.
  ThisWorkbook.RegExp.Global = False             ' Legt globales Anwenden fest.
  Set Uebereinstimmungen = ThisWorkbook.RegExp.Execute(Zeichenfolge)   ' Führt die Suche aus.
  For Each Uebereinstimmung In Uebereinstimmungen   ' Durchläuft die Auflistung der Übereinstimmungen.
     Rueckgabe = Rueckgabe & "Entsprechung gefunden bei "
     Rueckgabe = Rueckgabe & Uebereinstimmung.FirstIndex & ". Wert: '"
     Rueckgabe = Rueckgabe & Uebereinstimmung.value & "'." & vbCrLf
  Next
  'Set ThisWorkbook.RegExp = Nothing
  Set Uebereinstimmungen = Nothing
  RegExpTest = Rueckgabe
  Exit Function

Fehler:
  FehlerNachricht "RegExpTest()"
End Function


Function entspricht(SuchMuster, Zeichenfolge) as Boolean
  'Liefert "true", wenn Zeichenfolge dem Suchmuster entspricht
  On Error GoTo Fehler
  
  'Dim ThisWorkbook.RegExp As Object
  'Set ThisWorkbook.RegExp = CreateObject("VBScript.RegExp")
  ThisWorkbook.RegExp.Pattern = SuchMuster       ' Setzt das Muster.
  ThisWorkbook.RegExp.IgnoreCase = True          ' Ignoriert die Schreibweise.
  ThisWorkbook.RegExp.Global = False             ' Legt globales Anwenden fest.
  If ThisWorkbook.RegExp.test(Zeichenfolge) Then
    entspricht = True
  Else
    entspricht = False
  End If
  'Set ThisWorkbook.RegExp = Nothing
  Exit Function

Fehler:
  FehlerNachricht "entspricht()"
End Function



'***  Abteilung Berechnung und Vermessung *********************************************************

Function iGEO_aktPrjName() As String
  'Liefert Name des Projektes, der in %GEO_HOST%.SET eingetragen ist
  'Liefert "?", wenn kein aktives Projekt ermittelt werden kann.
  On Error Resume Next
  DebugEcho "mdlToolsAllgemein.iGEO_aktPrjName(): Name des aktiven iGEO-Projektes ermitteln."
  
  Dim geo_home    As String
  Dim geo_host    As String
  Dim VerzDatSet  As String
  Dim Zeile       As String
  Dim Kanal       As Integer
  Dim NF          As Long
  Dim Feld()      As String
  Dim PrjName     As String
  
  PrjName = "?"
  
  geo_home = Environ("GEO_HOME")
  geo_host = Environ("GEO_HOST")
  VerzDatSet = geo_home & "\hosts\" & geo_host & ".set"
  DebugEcho "mdlToolsAllgemein.iGEO_aktPrjName(): Lese Datei '" & VerzDatSet & "'"
  
  Kanal = FreeFile()
  Open VerzDatSet For Input Lock Write As #Kanal
  If (Err) Then GoTo Fehler
  Do While Not EOF(Kanal)
    Line Input #Kanal, Zeile
    NF = splitWords(Zeile, Feld, "")              'awk-like Splitting
    If (NF = 3) Then
      If (Feld(1) = "Name") Then
        PrjName = Feld(3)
        Exit Do
      End If
    End If
  Loop
  Close #Kanal
  DebugEcho "mdlToolsAllgemein.iGEO_aktPrjName(): Aktives iGeo-Projekt: '" & PrjName & "'"
  
  iGEO_aktPrjName = PrjName
  Exit Function

Fehler:
  Close #Kanal
End Function


Sub Transfo_Tk2Gls(ByVal Radius, ByVal u, ByVal BasisUeb, ByVal Abst, ByVal dH, ByRef AbstRed, ByRef dHRed)
  'Reduziert Abst und dH, falls der Punkt in einem Bogen mit Überhöhung liegt.
  'Eingabeparameter: Radius   [m] (Tra.Radius) wichtig ist nur das Vorzeichen
  '                  u        [m] (Tra.u)      Überhöhung
  '                  BasisUeb [m]              Basis für Überhöhung = ca. Abstand der Schienenmitten
  '                  Abst     [m] (TK.Q)       Waagerechter Abstand zur Achse
  '                  dH       [m] (TK.HSOK)    vertikaler Abstand zu SO
  'Ausgabeparameter: AbstRed  [m] (TK.QG)      Abstand zur Achse im gedrehten Gleissystem
  '                  dHRed    [m] (TK.HG)      Höhenunterschied zu SO im gedrehten Gleissystem
  
  Dim sf        As Integer
  Dim X0        As Double
  Dim Y0        As Double
  Dim phi       As Double
  Dim X         As Double
  Dim CosPhi    As Double
  Dim SinPhi    As Double
    
  'Rückgabewerte, falls nicht berechenbar
  AbstRed = ""
  dHRed   = ""
  
  'Berechnung
  If (IsNumeric(Radius) And IsNumeric(u) And IsNumeric(Abst) And IsNumeric(dH)) Then
    'Kein Eingabewert ist leer.
    
    'sf  = Sgn(Radius)
    'phi = sf * Atn(u / BasisUeb) * (-1)
    'X0  = abs(BasisUeb/2 * sin(phi))
    'Y0  = sf * (BasisUeb/2 - (BasisUeb/2 * cos(phi)))
    
    ' 21.08.2019 (Angleichung an iGeo und VermEsn:  Nullpunkt wird nur in der Höhe verschoben um u/2)
    sf     = Sgn(Radius) * Sgn(u)
    X      = Abs(u) / BasisUeb
    phi    = sf  *  Atn(X / Sqr(-X * X + 1))  *  (-1)    ' Arcsin(X) = Atn(X / Sqr(-X * X + 1))
    CosPhi = Cos(phi)
    SinPhi = Sin(phi)
    X0     = Abs(u) / 2
    Y0     = 0.0
    
    'Reduktion (bzw. Koordinatenumformung)
    If (u = 0) Then
      AbstRed = Abst
      dHRed   = dH
    ElseIf (sf <> 0) Then
      'AbstRed = (Abst - Y0) * Cos(phi) + (dH - X0) * Sin(phi)
      'dHRed   = (dH - X0) * Cos(phi) - (Abst - Y0) * Sin(phi)
      
      AbstRed = (Abst - Y0) * CosPhi + (dH - X0) * SinPhi
      dHRed   = (dH - X0) * CosPhi - (Abst - Y0) * SinPhi
    End If
  End If
End Sub


Sub Transfo_Gls2Tk(ByVal Radius, ByVal u, ByVal BasisUeb, ByVal AbstRed, ByVal dHRed, ByRef Abst, ByRef dH)
  'Transformation von Koo' im gedrehten Gleissystem in normale Trassenkoordinaten (y=waagerecht).
  'Sind Radius oder Überhöhung gleich Null, so ist das Ergebnis identisch mit den Eingangswerten.
  'Eingabeparameter: Radius   [m] (Tra.Radius) wichtig ist nur das Vorzeichen
  '                  u        [m] (Tra.u)      Überhöhung
  '                  BasisUeb [m]              Basis für Überhöhung = ca. Abstand der Schienenmitten
  '                  AbstRed  [m] (TK.Q)       Abstand zur Achse im gedrehten Gleissystem
  '                  dHRed    [m] (TK.HSOK)    Höhenunterschied zu SO im gedrehten Gleissystem
  'Ausgabeparameter: Abst     [m] (TK.QG)      Waagerechter Abstand zur Achse
  '                  dH       [m] (TK.HG)      vertikaler Abstand zu SO
  
  Dim sf        As Integer
  Dim X0        As Double
  Dim Y0        As Double
  Dim phi       As Double
  Dim X         As Double
  Dim CosPhi    As Double
  Dim SinPhi    As Double
  
  'Rückgabewerte, falls nicht berechenbar
  Abst = ""
  dH   = ""
  
  'Berechnung
  If (IsNumeric(Radius) And IsNumeric(u) And IsNumeric(AbstRed) And IsNumeric(dHRed)) Then
    'Kein Eingabewert ist leer.
    
    'sf  = Sgn(Radius)
    'phi = sf * Atn(u / BasisUeb) * (-1)
    'X0  = abs(BasisUeb/2 * sin(phi))
    'Y0  = sf * (BasisUeb/2 - (BasisUeb/2 * cos(phi)))
    'cp  = Cos(phi)
    'sp  = Sin(phi)
    
    ' 21.08.2019 (Angleichung an iGeo und VermEsn:  Nullpunkt wird nur in der Höhe verschoben um u/2)
    sf     = Sgn(Radius) * Sgn(u)
    X      = Abs(u) / BasisUeb
    phi    = sf  *  Atn(X / Sqr(-X * X + 1))  *  (-1)    ' Arcsin(X) = Atn(X / Sqr(-X * X + 1))
    CosPhi = Cos(phi)
    SinPhi = Sin(phi)
    X0     = Abs(u) / 2
    Y0     = 0.0
    
    'Reduktion (bzw. Koordinatenumformung)
    If (u = 0) Then
      Abst = AbstRed
      dH = dHRed
    ElseIf (sf <> 0) Then
      'dH = X0 + (AbstRed / cp + dHRed / sp) / (cp / sp + sp / cp)
      'Abst = Y0 + AbstRed / cp - (dH - X0) * sp
      
      dH   = X0 + (AbstRed / CosPhi + dHRed / SinPhi) / (CosPhi / SinPhi + SinPhi / CosPhi)
      Abst = Y0 + (AbstRed - (dH - X0) * SinPhi) / CosPhi
    End If
  End If

End Sub


function GetStreckenbezeichnung(byVal StrNr, byVal BezLang)
  'Rückgabe: DB-Streckenbezeichnung zur Streckennummer
  'Eingabe:
  ' StrNr:   Streckennummer
  ' BezLang: (true|false). Wenn false, werden Streckenbeginn und -ende ab Komma abgeschnitten
  
  'Deklarationen:
  'Objekte
    Dim fs        As New Scripting.FileSystemObject
    Dim oTS       As Scripting.TextStream
  
  'Strings   
    Dim i_Uebersichten, StreckenDatei
    Dim StrNrnLine
    Dim StrNrnBeschreibung
    dim StreckentextLang, StreckentextKurz, Streckentext
    dim VonNach, i, tmp
    dim strFirstString
    dim strWorkLine
    dim blnStreckeGefunden
    dim UEBERSICHTEN_STANDARD
    dim STRECKEN_DATEI
    
  'Variableninitialisierungen
    UEBERSICHTEN_STANDARD = "P:\Uebersichten"
    STRECKEN_DATEI        = "\Bahn\Strecken_Daten\Strecken.txt"
  
  'Pfad für Datei mit Streckenverzeichnis
    i_Uebersichten = Environ("I_UEBERSICHTEN")
    if (i_Uebersichten = "") then i_Uebersichten = UEBERSICHTEN_STANDARD
    StreckenDatei = i_Uebersichten & STRECKEN_DATEI
    
  'Streckendatei verarbeiten
    if (not fs.fileexists(StreckenDatei)) then
      ErrEcho "StreckenDatei '" & StreckenDatei & "' nicht gefunden!"
      Streckentext  = "StreckenDatei '" & StreckenDatei & "' nicht gefunden!"
      
    else
      'Datei öffnen und zeilenweise lesen, bis Streckennummer gefunden ist.
        set oTS = fs.OpenTextFile(StreckenDatei, 1)
        blnStreckeGefunden = false
        do while (not (oTS.AtEndOfStream or blnStreckeGefunden))
          strWorkLine = oTS.ReadLine
          'Wenn Zeile nicht leer, dann String vergleichen
          if (strWorkLine <> "") then
             strFirstString = left(trim(strWorkLine), 4)
             if (StrComp(StrNr, strFirstString, 1) = 0) then blnStreckeGefunden = true
          end If
        loop
        if (blnStreckeGefunden) then StrNrnLine = trim(Konv_OEM2ANSI(strWorkLine)) else StrNrnLine = ""
        oTS.Close
      
      'auszugebenden Text bestimmen
        If (StrNrnLine = "") Then
          Streckentext = "Strecke " & StrNr & " nicht gefunden! "
        else
          StrNrnBeschreibung = right(StrNrnLine,len(StrNrnLine)-4)
          StreckentextLang   = "Strecke " & StrNr & "  " & StrNrnBeschreibung
          
          'kurze Streckenbezeichnung bestimmen
          VonNach = Split(StrNrnBeschreibung, " - ", -1, 1)
          tmp     = Split(VonNach(0), ",", -1, 1)
          StreckentextKurz = "Strecke " & StrNr & "  " & tmp(0)
          
          if (ubound(VonNach) > 0) then
            for i = 1 to ubound(VonNach)
              tmp = Split(VonNach(i), ",", -1, 1)
              StreckentextKurz = StreckentextKurz & " - " & tmp(0)
            next
          end if
          
          'auszugebende Streckenbezeichnung: Standard ist "kurz"
          if (BezLang) then Streckentext = StreckentextLang else Streckentext = StreckentextKurz
        end if
      
      'Streckenbezeichnung bzw. Fehlermeldung zurückgeben:
      GetStreckenbezeichnung = Streckentext
    end if
    
  'Aufräumen
    set oTS = nothing
end function


Function Ueberhoehung(ByVal text As String, UebInInfo_Streng As Boolean) As String
  '--------------------------------------------------------------------------------------------------------'
  ' Rückgabewert = ist-Überhöhung in [mm], aus Bemerkung ermittelt (leer, wenn unbekannt).
  ' Eingabe: text             ... Punktinfo'
  '          UebInInfo_Streng ...(1=ja, 0=nein)... Wenn = 1, dann wird nur "u=xxx" erkannt,
  '                                                sonst wird auch die erste Zahl als Überhöhung verwendet.
  '--------------------------------------------------------------------------------------------------------'
  ' ==> Es wird versucht, der Punktinfo die gemessene Ist-Überhöhung
  '     nach folgenden Regeln zu entnehmen:
  '     1. Falls die Zeichenkette "u= xxx" (an irgendeiner Stelle) enthalten
  '        ist, so wird "xxx" als Ist-Überhöhung angesehen.
  '     2. Falls Variante 1 nicht zum Erfolg führt und in den Einstellungen
  '        nicht nur die strenge Variante erlaubt ist, wird:
  '        => die erste Zahl als Ist-Überhöhung verwendet.
  '--------------------------------------------------------------------------------------------------------'
  
  Dim ui           As String
  Dim Fundstellen  As Object
  'Dim ThisWorkbook.RegExp As Object
  'Set ThisWorkbook.RegExp = CreateObject("VBScript.RegExp")
  
  text = Trim$(text)
  ui = ""
  
  '1. nach "u=..." suchen
  ThisWorkbook.RegExp.IgnoreCase = True
  ThisWorkbook.RegExp.Global = False
  ThisWorkbook.RegExp.Pattern = "u *= *[-|+]? *[0-9]+"
  Set Fundstellen = ThisWorkbook.RegExp.Execute(text)
  If (Fundstellen.Count > 0) Then
    ui = Fundstellen(0).value
    ThisWorkbook.RegExp.Global = True
    ThisWorkbook.RegExp.Pattern = "u *= *"
    ui = ThisWorkbook.RegExp.Replace(ui, "")
    ThisWorkbook.RegExp.Pattern = " +"
    ui = ThisWorkbook.RegExp.Replace(ui, "")
  end if
  
  '2. "u=..." nicht gefunden => erste Zahl nehmen, falls erlaubt.'
  If ((ui = "") and not UebInInfo_Streng) Then
    ThisWorkbook.RegExp.Global = False
    ThisWorkbook.RegExp.Pattern = "[-|+]?[0-9]+"
    Set Fundstellen = ThisWorkbook.RegExp.Execute(text)
    If (Fundstellen.Count > 0) Then
      ui = Fundstellen(0).value
      ThisWorkbook.RegExp.Global = True
      ThisWorkbook.RegExp.Pattern = " "
      ui = ThisWorkbook.RegExp.Replace(ui, "")
    End If
  End If
  
  Set Fundstellen = Nothing
  'Set ThisWorkbook.RegExp = Nothing
  Ueberhoehung = ui
End Function


Function GetKm(byVal KmAngabe)
  'Gibt die der KmAngabe entsprechende reele Zahl in [m] oder NULL zurück.
  'Eingabe:  KmAngabe ... reele Zahl in [m] oder übliche Km-Angabe (Vorzeichen des Km-Anteils ist allein entscheidend)
  dim Km, Hm, m, T, KmReal, VorzKm, VorzM, Vorz
  Dim oMatches
  'Dim ThisWorkbook.RegExp As Object
  'Set ThisWorkbook.RegExp = CreateObject("VBScript.RegExp")
  KmReal = null
  ThisWorkbook.RegExp.IgnoreCase = True
  ThisWorkbook.RegExp.Global = False
  ThisWorkbook.RegExp.Pattern = "^ *([+\-]? *[0-9]*[.]*[0-9]+)([-+ ]+)([0-9]*[.]*[0-9]+) *$"
  Set oMatches = ThisWorkbook.RegExp.Execute(KmAngabe)
  if (oMatches.Count = 0) then
    'Keine gültige Km-Schreibweise => eventuell eine reele Zahl.
    if (isNumeric(KmAngabe)) then KmReal = cdbl(KmAngabe)
  else
    Km = oMatches(0).SubMatches(0)
    T  = oMatches(0).SubMatches(1)
    m  = oMatches(0).SubMatches(2)  'ohne Vorzeichen
    Km = replace(Km, " ", "")
    VorzKm = sgn(Km)
    if (instr(T, "-") > 0) then VorzM = -1 else VorzM = 1
    if ((VorzM = -1) or (VorzKm = -1)) then Vorz = -1 else Vorz = 1
    'msgbox "km=" & km & vbnewline & "m=" & m & vbnewline & "Vorz=" & Vorz
    KmReal = Vorz * (abs(Km) * 1000 + m)
  end if
  Set oMatches = Nothing
  'Set ThisWorkbook.RegExp = Nothing
  GetKm = KmReal
end function



'***  Abteilung Arrays und Dictionaries ***********************************************************


Function isVektorEmpty(byRef Vektor As Variant) As Boolean
  'Parameter: Vektor ... eindimensionales Array
  'Rückgabe:  false, wenn nicht alle Werte des Vektors leer sind bzw. "",
  '           true in allen anderen Fällen (auch bei Fehler).
  
  On Error GoTo Fehler
  
  Dim lb        As Long
  Dim ub        As Long
  Dim i         As Long
  
  isVektorEmpty = True
  
  lb = LBound(Vektor)
  ub = UBound(Vektor)
  For i = lb To ub
    If (Vektor(i) <> "") Then
      isVektorEmpty = False
      Exit For
    End If
  Next
  Exit Function
  
Fehler:
  FehlerNachricht "mdlToolsAllgemein.isVektorEmpty()"
End Function


Function CountDim(Feld) As Long
  'Parameter: Feld ... ein- oder mehrdimensionales Array
  'Rückgabe:  Anzahl der Dimensionen von Feld, kann auch "0" sein.
  On Error Resume Next
  Dim idx  As Long
  Dim i    As Long
  i = 0
  Do
    i = i + 1
    idx = UBound(Feld, i)
  Loop Until (Err <> 0)
  CountDim = i - 1
  On Error GoTo 0
  Exit Function
Fehler:
  FehlerNachricht "mdlToolsAllgemein.CountDim()"
End Function


sub TransposeArray2d(byRef Matrix As Variant)
  'Rückgabe: Transponiertes zweidimensionales Array "Matrix".
  'Eingabe:  Matrix        ... zu transponierendes Array
  dim TM          As Variant
  dim lbZeilen    as Long
  dim lbSpalten   as Long
  dim ubZeilen    as Long
  dim ubSpalten   as Long
  dim ze          as Long
  dim sp          as Long
  
  if (CountDim(Matrix) = 2) then
    lbZeilen  = LBound(Matrix, 1)
    lbSpalten = LBound(Matrix, 2)
    ubZeilen  = UBound(Matrix, 1)
    ubSpalten = UBound(Matrix, 2)
    redim TM(ubSpalten, ubZeilen)
    for ze = lbZeilen to ubZeilen
      for sp = lbSpalten to ubSpalten
        TM(sp, ze) = Matrix(ze, sp)
      next
    next
  end if
  
  Matrix = TM
end sub


sub SortArray1d(byRef Vector, byVal SortType As Byte, byVal Descending As Boolean)
  'Sortiert das eindimensionale Array "Vector" durch Start von QuickSort1d() für das gesamte Array.
  'Beschreibung der Parameter: siehe QuickSort1d()
  Dim idxU      As Long
  Dim idxO      As Long
  
  idxU = lbound(Vector)
  idxO = ubound(Vector)
  DebugEcho "SortArray1d(): Initialisierung QuickSort1d(Vector," & "," & SortType & "," & cStr(Descending) & "," & cStr(idxU) & "," & cStr(idxO) & ")" 
  call QuickSort1d(Vector, SortType, Descending, idxU, idxO)
end sub


Sub QuickSort1d(byRef Vector, byVal SortType As Byte, byRef Descending As Boolean, byRef idxU As Long, byRef idxO As Long)
  '-------------------------------------------------------------------------------------------------
  'Sortiert das eindimensionale Array "Vector" im Indexbereich idxU - idxO.
  'Zum Sortieren des gesamten Arrays kann diese Routine indirekt gestartet werden über SortArray1d().
  '(ACHTUNG: Eingabeparameter werden wegen Performance als Referenz übergeben...)
  '
  'Eingabe:  Vector     ... zu sortierendes Array
  '          SortType   ... eine der SORTTYPE_*-Konstanten - siehe isLesser()
  '          Descending ... wenn true, dann wird in absteigender Reihenfolge sortiert.
  '          idxU       ... kleinster Array-Index (für Partitionierung).
  '          idxO       ... größter Array-Index   (für Partitionierung).
  '
  'Grundlage war VB-Kode eines Tutorials (QuickSort eines 1d-Arrays) von Klaus Neumann:
  'siehe http://www.activevb.de/tutorials/tut_sortalgo/sortalgo.html
  '-------------------------------------------------------------------------------------------------
  Dim sek_Dim     As Long
  Dim i           As Long
  Dim k           As Long
  Dim idxM        As Long
  Dim idxSpalte   As Long
  Dim Wert_idxM   As Variant
  Dim temp        As Variant
  
  'DebugEcho "Starte QuickSort1d(Vector," & SortType & "," & Descending & "," & idxU & "," & idxO & ")" 
  if ((idxU < 0) or (idxO < 0)) then
    DebugEcho "QuickSort1d(): Sortieren unmöglich, da mindestens ein Index < 0!"
  else
    idxM = int((idxU + idxO) / 2)
    i    = idxU
    k    = idxO
    
    'Pivotelement: Wert ca. in der Mitte der Schlüsselspalte.
    Wert_idxM = Vector(idxM)
    
    Do
      Do While isLesser(Vector(i), Wert_idxM, SortType, Descending)
        i = i + 1
        'Schleifenausgang spätestens mit i = idxM + 1
      Loop
      
      Do While isLesser(Wert_idxM, Vector(k), SortType, Descending)
        k = k - 1
        'Schleifenausgang spätestens mit k = idxM - 1
      Loop
      
      If (i <= k) Then
        'Arraywerte der Indizes i und k tauschen.
        temp = Vector(k)
        Vector(k) = Vector(i)
        Vector(i) = temp
        
        i = i + 1
        k = k - 1
      End If
      
    Loop Until (i > k)
    
    If (idxU < k) Then call QuickSort1d(Vector, SortType, Descending, idxU, k)
    If (i < idxO) Then call QuickSort1d(Vector, SortType, Descending, i, idxO)
  end if
End Sub


sub SortArray2d(byRef Matrix, byVal Key_Dim As Long, byVal Key_Idx As Long, byVal SortType As Byte, byVal Descending As Boolean)
  'Sortiert das zweidimensionale Array "Matrix" durch Start von QuickSort2d() für das gesamte Array.
  'Beschreibung der Parameter: siehe QuickSort2d()
  Dim idxU      As Long
  Dim idxO      As Long
  Dim sek_Dim   As Long
  
  if (Key_Dim = 1) then sek_Dim = 2 else sek_Dim = 1
  idxU = lbound(Matrix, sek_Dim)
  idxO = ubound(Matrix, sek_Dim)
  DebugEcho "SortArray2d(): Initialisierung QuickSort2d(Matrix," & cStr(Key_Dim) & "," & cStr(Key_Idx) & "," & SortType & "," & cStr(Descending) & "," & cStr(idxU) & "," & cStr(idxO) & ")" 
  call QuickSort2d(Matrix, Key_Dim, Key_Idx, SortType, Descending, idxU, idxO)
end sub


Sub QuickSort2d(byRef Matrix, byRef Key_Dim As Long, byRef Key_Idx As Long, _ 
                byVal SortType As Byte, byRef Descending As Boolean, _
                byRef idxU As Long, byRef idxO As Long)
  '-------------------------------------------------------------------------------------------------
  'Sortiert das zweidimensionale Array "Matrix" im Indexbereich idxU - idxO.
  'Zum Sortieren des gesamten Arrays kann diese Routine indirekt gestartet werden über SortArray2d().
  '(ACHTUNG: Eingabeparameter werden wegen Performance als Referenz übergeben...)
  '
  'Eingabe:  Matrix     ... zu sortierendes Array
  '          Key_Dim    ... (1|2)   Nr. der Dimension, die die Schlüsselspalte enthält.
  '          Key_Idx    ... (0,1,.) Nr. des Indexes (von Key_Dim), der der Schlüsselspalte entspricht.
  '                         => Die Schlüsselspalte ist diejenige, nach der sortiert werden soll.
  '          SortType   ... eine der SORTTYPE_*-Konstanten - siehe isLesser()
  '          Descending ... wenn true, dann wird in absteigender Reihenfolge sortiert.
  '          idxU       ... kleinster Array-Index (für Partitionierung).
  '          idxO       ... größter Array-Index   (für Partitionierung).
  '
  'Grundlage war VB-Kode eines Tutorials (QuickSort eines 1d-Arrays) von Klaus Neumann:
  'siehe http://www.activevb.de/tutorials/tut_sortalgo/sortalgo.html
  '-------------------------------------------------------------------------------------------------
  Dim sek_Dim     As Long
  Dim i           As Long
  Dim k           As Long
  Dim idxM        As Long
  Dim idxSpalte   As Long
  Dim Wert_idxM   As Variant
  Dim temp        As Variant
  
  'DebugEcho "Starte QuickSort2d(Matrix," & Key_Dim & "," & Key_Idx & "," & SortType & "," & Descending & "," & idxU & "," & idxO & ")" 
  if ((idxU < 0) or (idxO < 0)) then
    DebugEcho "QuickSort2d(): Sortieren unmöglich, da mindestens ein Index < 0!"
  else
    if (Key_Dim = 1) then sek_Dim = 2 else sek_Dim = 1
    
    idxM = int((idxU + idxO) / 2)
    i    = idxU
    k    = idxO
    
    'Pivotelement: Wert ca. in der Mitte der Schlüsselspalte.
    Wert_idxM = ArrayValue(Matrix, Key_Dim, Key_Idx, idxM)
    
    Do
      Do While isLesser(ArrayValue(Matrix, Key_Dim, Key_Idx, i), Wert_idxM, SortType, Descending)
        i = i + 1
        'Schleifenausgang spätestens mit i = idxM + 1
      Loop
      
      Do While isLesser(Wert_idxM, ArrayValue(Matrix, Key_Dim, Key_Idx, k), SortType, Descending)
        k = k - 1
        'Schleifenausgang spätestens mit k = idxM - 1
      Loop
      
      If (i <= k) Then
        'Arraywerte der Indizes i und k tauschen.
        For idxSpalte = lbound(Matrix, Key_Dim) To ubound(Matrix, Key_Dim)
          if (Key_Dim = 1) then
            temp = Matrix(idxSpalte, k)
            Matrix(idxSpalte, k) = Matrix(idxSpalte, i)
            Matrix(idxSpalte, i) = temp
          else
            temp = Matrix(k, idxSpalte)
            Matrix(k, idxSpalte) = Matrix(i, idxSpalte)
            Matrix(i, idxSpalte) = temp
          end if
        Next
        
        i = i + 1
        k = k - 1
      End If
      
    Loop Until (i > k)
    
    If (idxU < k) Then call QuickSort2d(Matrix, Key_Dim, Key_Idx, SortType, Descending, idxU, k)
    If (i < idxO) Then call QuickSort2d(Matrix, Key_Dim, Key_Idx, SortType, Descending, i, idxO)
  end if
End Sub


Function ArrayValue(byRef Matrix, byRef Key_Dim As Long, byRef Key_Idx As Long, byRef Sek_Idx As Long) As Variant
  'Rückgabe: angegebener Wert des zweidimensionalen Arrays "Matrix".
  '          ==> Hilfsfunktion für QuickSort2d!
  'Eingabe:  Matrix        ... zu sortierendes Array
  '          Key_Dim       ... (1|2)   Nr. der Dimension, die die Schlüsselspalte enthält.
  '          Key_Idx       ... (0,1,.) Nr. des Indexes (von Key_Dim), der der Schlüsselspalte entspricht.
  '          Sek_Idx       ... (0,1,.) Nr. des anderen Indexes.
  if (Key_Dim = 1) then
    ArrayValue = Matrix(Key_Idx, Sek_Idx)
  else
    ArrayValue = Matrix(Sek_Idx, Key_Idx)
  end if
end function


Function isLesser(Wert1 As Variant, Wert2 As Variant, ByVal SortType As Byte, ByVal Reverse As Boolean) As Boolean
  '-------------------------------------------------------------------------------------------------
  'Vergleicht Wert1 und Wert2 auf die in "SortType" angegebene Weise.
  'Eingabe:  Wert1+2  ... zu vergleichende Werte
  '          SortType ... eine der SORTTYPE_*-Konstanten:
  '                       - SORTTYPE_NUMERIC = 0  'Numerisch, wenn möglich, sonst Alphanumerisch
  '                       - SORTTYPE_STRING  = 1  'Alphanumerisch
  '          Reverse  ... true | false (siehe Rückgabe).
  '
  'Rückgabe: - true, wenn Wert1 < Wert2 und Reverse = false
  '          - true, wenn Wert1 > Wert2 und Reverse = true,
  '          - false in allen anderen Fällen.
  '          => Ein NULL-Wert ist immer kleiner als alle anderen Werte (außer NULL :-).
  '
  '==> Diese Funktion ist als Hilfe für Sortierungs-Routinen gedacht. Also Vorsicht bei Änderungen!!!
  '-------------------------------------------------------------------------------------------------
  
  dim blnVgl            As Boolean
  dim intRev            As Integer
  dim internalSortType  As Byte
  
  blnVgl = false
  internalSortType = SortType
  if (Reverse) then intRev = -1 else intRev = 1
  
  if (isNull(Wert1) or isNull(Wert2)) then
    'Ein NULL-Wert ist immer kleiner als alle anderen Werte
    if (isNull(Wert1) and isNull(Wert2)) then
      blnVgl = false
    elseif (isNull(Wert1)) then
      if (not Reverse) then blnVgl = true else blnVgl = false
    else
      if (Reverse) then blnVgl = true else blnVgl = false
    end if
    
  else
    if (internalSortType = 0) then
      'Wenn numerisch nicht möglich, dann Alphanumerisch => Texte werden ans Ende sortiert.
      if (not (isNumeric(Wert1) and isNumeric(Wert2))) then internalSortType = 1
    end if
    
    select case internalSortType
      
      case 0
          'Numerische Sortierung.
          Wert1 = cDbl(Wert1)
          Wert2 = cDbl(Wert2)
          if (not Reverse) then blnVgl = (Wert1 < Wert2) else blnVgl = (Wert2 < Wert1)
          
      case else
          'Alphanumerische Sortierung.
          Wert1 = cStr(Wert1)
          Wert2 = cStr(Wert2)
          if (strComp(Wert1, Wert2, vbTextCompare) * intRev = -1) then blnvgl = true
    end select
  end if
  
  isLesser = blnVgl
end function



Public Function SortDictionary(byRef oDictionary As Scripting.Dictionary, byVal sortBy As Long, _
                               byVal SortType As Byte, byRef Descending As Boolean) As Scripting.Dictionary
  '-------------------------------------------------------------------------------------------------
  'Sortiert das Dictionary nach Keys oder Items
  '
  'Eingabe:  oDictionary ... zu sortierendes Dictionary
  '          sortBy      ... 0 = Key, 1 = Item
  '          SortType    ... eine der SORTTYPE_*-Konstanten - siehe isLesser()
  '          Descending  ... wenn true, dann wird in absteigender Reihenfolge sortiert.
  '
  'Ausgabe:  oDictionary ... sortiertes Dictionary
  '
  'Grundlage war eine Funktion mit integriertem ShellSort-Algorithmus:
  'siehe http://support.microsoft.com/?scid=kb%3Ben-us%3B246067&x=12&y=10
  '-------------------------------------------------------------------------------------------------
  'Declarations
  Dim KeysItems()  As Variant
  Dim Key          As Variant
  dim i            As Long
  dim ItemCount    As Long
  
  Const idxKey  = 0
  Const idxItem = 1
  
  ItemCount = oDictionary.Count
  
  'Store dictionary information into an array
  If (ItemCount > 1) Then
    ReDim KeysItems(0 to ItemCount-1, 0 to 1)
    i = 0
    For Each Key In oDictionary
      KeysItems(i, idxKey)  = Key
      KeysItems(i, idxItem) = oDictionary.Item(Key)
      i = i + 1
    Next
    
    'Perform a QuickSort of the array
    call SortArray2d(KeysItems, 2, sortBy, SortType, Descending)
    
    'Erase the contents of the dictionary object
    oDictionary.RemoveAll
    
    'Repopulate the dictionary in sorted order
    For i = lbound(KeysItems, 1) to ubound(KeysItems, 1)
      oDictionary.Add KeysItems(i, idxKey), KeysItems(i, idxItem)
    Next
  End If
End Function


Function ListeDerKeys(oDictionary As Scripting.Dictionary) As String
  'Listet alle Keys eines Dictionary in einem String durch Semikolons getrennt.
  Dim DictionaryItems, DictionaryKeys, Liste, i
  Liste = ""
  If (Not oDictionary Is Nothing) Then
    DictionaryKeys = oDictionary.Keys
    For i = 0 To oDictionary.Count - 1
      If (i = 0) Then
        Liste = DictionaryKeys(i)
      Else
        Liste = Liste & ";" & DictionaryKeys(i)
      End If
    Next
  End If
  ListeDerKeys = Liste
End Function


Function ListeDictionary(ByRef oDictionary As Scripting.Dictionary) As String
  'Listet ein Dictionary in einem String zwecks Anzeige...
  Dim DictionaryItems  As Variant
  Dim DictionaryKeys   As Variant
  Dim i                As Long
  Dim idx              As Long
  Dim lb               As Long
  Dim StringCount      As Long
  Dim StringArray()    As String
  
  DictionaryItems = oDictionary.Items
  DictionaryKeys = oDictionary.Keys
  lb = LBound(DictionaryKeys)
  
  StringCount = 3 + oDictionary.Count
  ReDim StringArray(1 To StringCount)
  
  StringArray(1) = "Das Dictionary hat " & CStr(oDictionary.Count) & " Einträge"
  StringArray(2) = "Inhalt des Dictonary: " & vbNewLine & "i" & vbTab & "Key" & vbTab & "Eintrag" & vbNewLine & "-------------------------------------------------------------------"
  
  For i = lb To UBound(DictionaryKeys)
    'Hinweise zur Performance:
    '1. Klammern der kurzen Strings bedeutet: Schleife braucht nur noch 17% der Zeit ohne Klammern
    '2. Ersetzen der Stringverkettung durch Speichern der einzelnen Zeilen im Array und Join() reduziert die Laufzeit auf 1%!!! 
    'Liste = Liste & (CStr(i) & vbTab & DictionaryKeys(i) & vbTab & DictionaryItems(i) & vbNewLine)
    
    StringArray(3 + i - lb) = (CStr(i) & vbTab & DictionaryKeys(i) & vbTab & DictionaryItems(i))
  Next
  StringArray(StringCount) = vbNewLine
  
  ListeDictionary = Join(StringArray, vbNewLine)
End Function


Function ListeAuflistung(oAuflistung) As String
  'Rückgabe: String mit Auflistung aller Einträge einer Auflistung/Collection zwecks Anzeige für Debug-Zwecke.
  dim i                As Variant
  dim Eintrag          As Variant
  Dim StringCount      As Long
  Dim StringArray()    As String
  
  StringCount = 3 + oAuflistung.Count
  ReDim StringArray(1 To StringCount)
  
  StringArray(1) = "Die Auflistung/Collection hat " & CStr(oAuflistung.Count) & " Einträge"
  StringArray(2) = "Inhalt der Auflistung/Collection: " & vbNewLine & "i" & vbTab & "Eintrag" & vbNewLine & "-------------------------------------------------------------------"
  
  i = 2
  for each Eintrag in oAuflistung
    i = i + 1
    StringArray(i) = (CStr(i) & vbTab & Eintrag)
  next
  StringArray(StringCount) = vbNewLine
  
  ListeAuflistung = Join(StringArray, vbNewLine)
End Function


'für jEdit:  :folding=indent::collapseFolds=1:
