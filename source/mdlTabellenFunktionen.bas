Attribute VB_Name = "mdlTabellenFunktionen"
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2003 - 2014  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'====================================================================================
'Modul mdlTabellenFunktionen
'====================================================================================
'Stellt Funktionen zur Verfügung, die in Tabellenformeln verwendet werden können

'==> Parameter "Range" wird nicht verarbeitet. Excel berechnet die Formel nur dann
'    automatisch, wenn sich Zellen ändern, die als Funktionsargument dienen.


Option Explicit


'Struktur der NivLinien-Tabelle
Private Const Spalte_s          As Integer = 11    'Spalte "Strecke/Mittel"
Private Const Spalte_d          As Integer = 12    'Spalte "Differenz der dh's zw. Hin- und Rückweg"
Private Const Spalte_ds         As Integer = 14    'Spalte "Streckendifferenz zw. Hin- und Rückweg"

'Kenner der Zielgrößen für getNivStatistik()
Private Const NivStat_Summe_dds As Integer = 1     'Summe(dd/s)
Private Const NivStat_n         As Integer = 2     'Anzahl dh in Hin- und Rückgang
Private Const NivStat_m         As Integer = 3     'Mittlerer Kilometerfehler


'=== NivLinien ======================================================================

Function Niv_AnzDoppelDH(Optional oRange As Range) As Integer
  'Liefert die aktuell in der Tabelle enthaltene Anzahl an Nivellementlinien,
  'die im Hin- und Rückweg gemessen wurden.
  'Parameter "oRange" wird nicht verarbeitet. Er soll Excel veranlassen, die
  'Funktion neu zu berechnen, wenn eine Zelle dieses Bereiches geändert wird.
  'Liefert "-999" bei Fehler.
  
  'wenn Volatile="true": wilkürliche Neuberechnung, auch wenn Tabelle nicht aktiv ist!...
  Application.Volatile False
  On Error GoTo Fehler
  
  Dim oCallerRange As Range
  Dim CallerBlatt  As String
  Dim CallerMappe  As String
  Dim AktivesBlatt As String
  Dim AktiveMappe  As String
  
  Set oCallerRange = Application.Caller
  CallerBlatt = oCallerRange.Parent.Name
  CallerMappe = oCallerRange.Parent.Parent.Name
  AktivesBlatt = ActiveSheet.Name
  AktiveMappe = ActiveSheet.Parent.Name
  
  'MsgBox "Aufruf Niv_AnzDoppelDH() durch: Mappe=" & CallerMappe & "   Blatt=" & CallerBlatt & vbNewLine & _
         "Aktive Tabelle: Mappe=" & AktiveMappe & "  Blatt=" & AktivesBlatt & vbNewLine & _
         "Ereignisse aktiviert: " & Application.EnableEvents
  
  If ((AktiveMappe <> CallerMappe) Or (AktivesBlatt <> CallerBlatt)) Then
    'Funktionsaufruf kommt nicht aus der aktiven Tabelle
    '==> funktioniert nicht, deshalb Wert der Zelle nicht verändern.
    'MsgBox "Aufruf Niv_AnzDoppelDH() durch: Mappe=" & CallerMappe & "   Blatt=" & CallerBlatt & vbNewLine & _
           "Aktive Tabelle: Mappe=" & AktiveMappe & "  Blatt=" & AktivesBlatt & vbNewLine & _
           "Ereignisse aktiviert: " & Application.EnableEvents
    Err.Raise vbObjectError + ErrNumFktAufrufUngueltig - vbObjectError, , "Funktionsaufruf kommt nicht aus der aktiven Tabelle!"
    'Niv_AnzDoppelDH = oCallerRange.value
  Else
    Niv_AnzDoppelDH = getNivStatistik(NivStat_n, Spalte_s, Spalte_d, Spalte_ds)
  End If
  Set oCallerRange = Nothing
Exit Function
Fehler:
  'FehlerNachricht "mdlTabellenFunktionen.Niv_AnzDoppelDH()"
  Set oCallerRange = Nothing
  Niv_AnzDoppelDH = -999
End Function



Function Niv_SummeDDS(Optional oRange As Range) As Double
  'Liefert das aktuelle Zwischenergebnis Summe(dd/s)
  'Parameter "oRange" wird nicht verarbeitet. Er soll Excel veranlassen, die
  'Funktion neu zu berechnen, wenn eine Zelle dieses Bereiches geändert wird.
  'Liefert "-999" bei Fehler.
  
  'wenn Volatile="true": wilkürliche Neuberechnung, auch wenn Tabelle nicht aktiv ist!...
  Application.Volatile False
  On Error GoTo Fehler
  
  Dim oCallerRange As Range
  Dim CallerBlatt  As String
  Dim CallerMappe  As String
  Dim AktivesBlatt As String
  Dim AktiveMappe  As String
  
  Set oCallerRange = Application.Caller
  CallerBlatt = oCallerRange.Parent.Name
  CallerMappe = oCallerRange.Parent.Parent.Name
  AktivesBlatt = ActiveSheet.Name
  AktiveMappe = ActiveSheet.Parent.Name
  
  If ((AktiveMappe <> CallerMappe) Or (AktivesBlatt <> CallerBlatt)) Then
    'Funktionsaufruf kommt nicht aus der aktiven Tabelle
    '==> funktioniert nicht, deshalb Wert der Zelle nicht verändern.
    'MsgBox "Aufruf Niv_AnzDoppelDH() durch: Mappe=" & CallerMappe & "   Blatt=" & CallerBlatt & vbNewLine & _
           "Aktive Tabelle: Mappe=" & AktiveMappe & "  Blatt=" & AktivesBlatt & vbNewLine & _
           "Ereignisse aktiviert: " & Application.EnableEvents
    Err.Raise vbObjectError + ErrNumFktAufrufUngueltig - vbObjectError, , "Funktionsaufruf kommt nicht aus der aktiven Tabelle!"
    'Niv_SummeDDS = oCallerRange.value
  Else
    Niv_SummeDDS = getNivStatistik(NivStat_Summe_dds, Spalte_s, Spalte_d, Spalte_ds)
  End If
  Set oCallerRange = Nothing
  Exit Function
  
Fehler:
  'FehlerNachricht "mdlTabellenFunktionen.Niv_SummeDDS()"
  Set oCallerRange = Nothing
  Niv_SummeDDS = -999
End Function



Function Niv_KmFehler(Optional oRange As Range) As Integer
  'Liefert den aktuellen "Mittleren Kilometerfehler"
  'Parameter "oRange" wird nicht verarbeitet. Er soll Excel veranlassen, die
  'Funktion neu zu berechnen, wenn eine Zelle dieses Bereiches geändert wird.
  'Liefert "-999" bei Fehler.
  
  'wenn Volatile="true": wilkürliche Neuberechnung, auch wenn Tabelle nicht aktiv ist!...
  Application.Volatile False
  On Error GoTo Fehler
  
  Dim oCallerRange As Range
  Dim CallerBlatt  As String
  Dim CallerMappe  As String
  Dim AktivesBlatt As String
  Dim AktiveMappe  As String
  
  Set oCallerRange = Application.Caller
  CallerBlatt = oCallerRange.Parent.Name
  CallerMappe = oCallerRange.Parent.Parent.Name
  AktivesBlatt = ActiveSheet.Name
  AktiveMappe = ActiveSheet.Parent.Name
  
  If ((AktiveMappe <> CallerMappe) Or (AktivesBlatt <> CallerBlatt)) Then
    'Funktionsaufruf kommt nicht aus der aktiven Tabelle
    '==> funktioniert nicht, deshalb Wert der Zelle nicht verändern.
    'MsgBox "Aufruf Niv_KmFehler() durch: Mappe=" & CallerMappe & "   Blatt=" & CallerBlatt & vbNewLine & _
           "Aktive Tabelle: Mappe=" & AktiveMappe & "  Blatt=" & AktivesBlatt & vbNewLine & _
           "Ereignisse aktiviert: " & Application.EnableEvents
    Err.Raise vbObjectError + ErrNumFktAufrufUngueltig - vbObjectError, , "Funktionsaufruf kommt nicht aus der aktiven Tabelle!"
    'Niv_KmFehler = oCallerRange.value
  Else
    Niv_KmFehler = getNivStatistik(NivStat_m, Spalte_s, Spalte_d, Spalte_ds)
  End If
  Set oCallerRange = Nothing
Exit Function
Fehler:
  'FehlerNachricht "mdlTabellenFunktionen.Niv_KmFehler()"
  Set oCallerRange = Nothing
  Niv_KmFehler = -999
End Function


Private Function getNivStatistik(byVal ZielGroesse As Integer, byVal Sp_s As Integer, byVal Sp_d As Integer, byVal Sp_ds As Integer) As Variant
  '----------------------------------------------------------------------------------------------
  'In der aktuellen Tabelle wird die Nivellement-Statistik berechnet.
  '
  'Rückgabe: Die mit "ZielGroesse" angegebene statistische Größe.
  '
  'Eingabe:  ZielGroesse ... Bestimmt, welche der 3 möglichen Groeßen zurückgegeben wird.
  '                          Eine der NivStat-Konstanten:
  '                           - NivStat_Summe_dds = Summe(dd/s)
  '                           - NivStat_n         = Anzahl dh in Hin- und Rückgang
  '                           - NivStat_m         = Mittlerer Kilometerfehler
  '          Sp_s        ... Nummer der Tabellenspalte "Strecke/Mittel"
  '          Sp_d        ... Nummer der Tabellenspalte "Differenz der dh's zw. Hin- und Rückweg"
  '          Sp_ds       ... Nummer der Tabellenspalte "Streckendifferenz zw. Hin- und Rückweg"
  '----------------------------------------------------------------------------------------------
  On Error GoTo Fehler
  
  Dim Summe_dds  As Variant       'Summe dd/s
  Dim n          As Variant       'Anzahl der im Hin- und Rückweg gemessenen Höhenunterschiede
  Dim m          As Variant       'Mittlerer Kilometerfehler
  Dim d          As Variant
  Dim s          As Variant
  Dim ZeAnf      As Long
  Dim ZeEnd      As Long
  Dim i          As Long
  
  'Festwerte des Datenbereiches ermitteln
  ZeAnf = ThisWorkbook.AktiveTabelle.ErsteDatenZeile
  ZeEnd = ThisWorkbook.AktiveTabelle.LetzteDatenZeile
  
  n = 0
  m = 0
  Summe_dds = 0
  For i = ZeAnf To ZeEnd
    If ((Not (IsEmpty(Cells(i, Sp_s).value))) And (Not (IsEmpty(Cells(i, Sp_d).value))) _
      And (Not (Cells(i, Sp_s).value = "")) And (Not (Cells(i, Sp_d).value = ""))) Then
      'Sowohl Strecke als auch "d" sind nicht leer.
      n = n + 1
      d = Cells(i, Sp_d).value
      s = Cells(i, Sp_s).value
      Summe_dds = Summe_dds + (d * d) / s
        If (n <> 0) Then
          m = Sqr(Summe_dds / (4 * n))
        Else
          m = 0
        End If
    End If
  Next
  
  Select Case ZielGroesse
    Case NivStat_Summe_dds: getNivStatistik = Summe_dds
    Case NivStat_n:         getNivStatistik = n
    Case NivStat_m:         getNivStatistik = m
    Case else:              getNivStatistik = -1
  End Select
  
  Exit Function
  
Fehler:
  'FehlerNachricht "mdlTabellenFunktionen.getNivStatistik()"
  getNivStatistik = -999
End Function
