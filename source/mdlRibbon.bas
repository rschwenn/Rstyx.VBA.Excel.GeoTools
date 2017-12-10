Attribute VB_Name = "mdlRibbon"
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2014-2017  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'===============================================================================
' Modul mdlRibbon                                                                   
'===============================================================================
' Stellt Zugriff auf das Menüband sowie Ribbon-Calllbacks Verfügung.

Option Explicit


Private oRibbon As IRibbonUI

' Region "Ribbon-Objekt (Referenz, Update)"
    
    ' Siehe http://social.msdn.microsoft.com/Forums/office/en-US/99a3f3af-678f-4338-b5a1-b79d3463fb0b/how-to-get-the-reference-to-the-iribbonui-in-vba
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)
    
    ' Bei Initialisierung der RibbonUI: Speichern einer Referenz auf das Ribbon-Objekt
    ' und als Backup eines entsprechenden Integer-Zeigers in die Add-In-interne Tabelle.
    Public Sub OnGeoToolsRibbonLoad(ribbon As IRibbonUI)
        Set oRibbon = ribbon
        tabGeoTools.Range("A1").Value = ObjPtr(ribbon)
    End Sub
    
    ' Beziehen einer Referenz auf das Ribbon-Objekt (Sollte auch nach Fehler im Add-In funktionieren).
    Function getGeoToolsRibbon() As IRibbonUI
        ' "oRibbon" ist normalerweise nur dann "Nothing", wenn das AddIn wegen eines Fehlers gestoppt wurde.
        ' Dann kann der vorher gespeicherte Zeiger verwendet werden.
        ' ABER: Wenn das AddIn nicht schreibgeschützt ist, kann der Zeiger auch veraltet sein.
        '       => Dann stürzt Excel ab und nichts geht mehr.
        If (oRibbon Is Nothing) Then
            If (ThisWorkbook.ReadOnly) Then
                Dim ribbonPointer As Long
                ribbonPointer = tabGeoTools.Range("A1").value
                If (ribbonPointer > 0) Then
                    On Error Resume Next  ' Nützt nix!
                    Call CopyMemory(oRibbon, ribbonPointer, 4)
                    On Error GoTo 0
                End If
            End If
        End If
        
        Set getGeoToolsRibbon = oRibbon
    End Function
    
    ' Status-Aktualisierung aller Ribbon-Steuerelemente erzwingen.
    ' Falls das AddIn gestoopt wurde, impliziert diese Routine auch dessen Neustart
    ' durch die Verwendung der Eigenschaft ThisWorkbook.AktiveTabelle...
    Public Sub UpdateGeoToolsRibbon(Optional ByVal keinAktivesBlatt As Boolean = False)
        On Error Resume Next
        getGeoToolsRibbon().Invalidate
        On Error Goto 0
    End Sub
    
' End Region


' Region "No Config Button"
    
    Sub GetVisibleNoConfigButton(control As IRibbonControl, ByRef returnedVal)
        returnedVal = True
        On Error Resume Next
        returnedVal = (Not ThisWorkbook.Konfig.KonfigDateiGelesen)
        On Error Goto 0
    End Sub
    
    Sub GetSupertipNoConfigButton(control As IRibbonControl, ByRef returnedVal)
        returnedVal = ThisWorkbook.Konfig.InfoKeineKonfig
    End Sub
    
    Sub NoConfigButtonAction(ByVal control As IRibbonControl)
        Call InfoKeineKonfig
        Call UpdateGeoToolsRibbon
    End Sub
    
' End Region

' Region "ComboBox"
    
    ' Anzahl NK-Stellen wurde via GUI geändert.
    Sub FmtOptPrecisionNumberChange(control As IRibbonControl, text As String)
        On Error Resume Next
        ThisWorkbook.AktiveTabelle.FormatDatenNKStellenAnzahl = cInt(text)
        On Error Goto 0
    End Sub
    
    'Callback for FmtOptPrecisionNumber getLabel
    Sub GetTextFmtOptPrecisionNumber(control As IRibbonControl, ByRef returnedVal)
        On Error Resume Next
        returnedVal = ThisWorkbook.AktiveTabelle.FormatDatenNKStellenAnzahl
        On Error Goto 0
    End Sub
    
' End Region

' Region "Action Buttons"
    
    Sub GeoToolsButtonAction(ByVal control As IRibbonControl)
        'On Error Resume Next
        Select Case control.ID
            Case "InfoButton"                   : Call GeoTools_Info
            Case "HelpButton"                   : Call Hilfe_Komplett
            Case "LogButton"                    : Call Protokoll
            Case "ImportExportButton"           : Call ExpimManager
            Case "TableStructureButton"         : Call TabellenStruktur
            Case "FormatButton"                 : Call FormatDaten
            Case "CalcDiffsButton"              : Call Mod_FehlerVerbesserung
            Case "CalcCantButton"               : Call Mod_UeberhoehungAusBemerkung
            Case "CalcHorizontalToCantedButton" : Call Mod_Transfo_Tk2Gls
            Case "CalcCantedToHorizontalButton" : Call Mod_Transfo_Gls2Tk
            Case "FormulaButton"                : Call UebertragenFormeln
            Case "DeleteButton"                 : Call LoeschenDaten
            Case "InterpolButton"               : Call Selection2Interpolationsformel
            Case "DuplicatesButton"             : Call Selection2MarkDoppelteWerte
            Case "BlankLinesButton"             : Call InsertLines
            Case "EditFileButton"               : Call DateiBearbeiten
            Case "SetFooterButton"              : Call SchreibeFusszeile_1
            Case "BatchPDFButton"               : Call BatchPDF
            Case "FormatContextMenuButton"      : Call FormatDaten
            Case Else                           : WarnEcho "mdlRibbon.GeoToolsButtonAction(): Unbekannte Control.ID = " & control.ID
        End select
        Call UpdateGeoToolsRibbon
        'On Error Goto 0
    End Sub
    
' End Region

' Region "Toggle Buttons"
    
    ' Response to a click on a toggle button.
    Sub GeoToolsToggleButtonAction(control As IRibbonControl, pressed As Boolean)
        On Error Resume Next
        Select Case control.ID
            Case "FmtOptStripesButton":         ThisWorkbook.AktiveTabelle.FormatDatenMitStreifen = pressed
            Case "FmtOptBackgroundButton":      ThisWorkbook.AktiveTabelle.FormatDatenOhneFuellung = pressed
            Case "FmtOptPrecisionButton":       ThisWorkbook.AktiveTabelle.FormatDatenNKStellenSetzen = pressed
            Case "CalcOptOverrideButton":       ThisWorkbook.AktiveTabelle.ModOpt_VorhWerteUeberschreiben = pressed
            Case "CalcOptKeepFormulasButton":   ThisWorkbook.AktiveTabelle.ModOpt_FormelnErhalten = pressed
            Case Else:                          WarnEcho "mdlRibbon.GeoToolsToggleButtonAction(): Unbekannte Control.ID = " & control.ID
        End select
        Call UpdateGeoToolsRibbon
        On Error Goto 0
    End Sub
    
    ' Get status of a toggle button.
    Sub GeoToolsToggleButtonGetPressed(control As IRibbonControl, ByRef returnedVal)
        On Error Resume Next
        Select Case control.ID
            Case "FmtOptStripesButton":         returnedVal = ThisWorkbook.AktiveTabelle.FormatDatenMitStreifen
            Case "FmtOptBackgroundButton":      returnedVal = ThisWorkbook.AktiveTabelle.FormatDatenOhneFuellung
            Case "FmtOptPrecisionButton":       returnedVal = ThisWorkbook.AktiveTabelle.FormatDatenNKStellenSetzen
            Case "CalcOptOverrideButton":       returnedVal = ThisWorkbook.AktiveTabelle.ModOpt_VorhWerteUeberschreiben
            Case "CalcOptKeepFormulasButton":   returnedVal = ThisWorkbook.AktiveTabelle.ModOpt_FormelnErhalten
            Case Else:                          WarnEcho "mdlRibbon.GeoToolsToggleButtonGetPressed(): Unbekannte Control.ID = " & control.ID
        End select
        On Error Goto 0
    End Sub
    
' End Region

' Region "GetEnabled"
    
    ' Verfügbar, wenn Makros nicht deaktiviert sind.
    Sub GetEnabledMacrosExecutable(control As IRibbonControl, ByRef returnedVal)
        returnedVal = IsMacrosExecutable
    End Sub
    
    ' Verfügbar, wenn das aktive Blatt eine Tabelle ist und Makros nicht deaktiviert sind.
    Sub GetEnabledTable(control As IRibbonControl, ByRef returnedVal)
        
        returnedVal = False
        
        If (Not (ActiveCell Is Nothing)) Then
            If (Not (Selection Is Nothing)) Then
                If (TypeOf Selection Is Range) Then
                    returnedVal = True
                End If
            End If
        End If
        
        returnedVal = (returnedVal And IsMacrosExecutable)
    End Sub
    
    ' Verfügbar, wenn der Infotraeger definiert ist ("GeoTools-Tabelle").
    Sub GetEnabledGeoToolsTable(control As IRibbonControl, ByRef returnedVal)
        
        Call GetEnabledTable(control, returnedVal)
        
        If (returnedVal) Then
            returnedVal = False
            Dim oTable As CtabAktiveTabelle
            Set oTable = ThisWorkbook.AktiveTabelle
            If (Not oTable Is Nothing) Then
                If (oTable.ExistsLokalerZellname(strInfoTraeger)) Then
                    returnedVal = True
                End If
            End If
        End If
    End Sub
    
    ' Verfügbar, wenn der Fliesskommabereich definiert ist.
    Sub GetEnabledFloatingPoint(control As IRibbonControl, ByRef returnedVal)
        
        Call GetEnabledGeoToolsTable(control, returnedVal)
        
        If (returnedVal) Then
            If (Not ThisWorkbook.AktiveTabelle.ExistsLokalerZellname(strFliesskomma)) Then
                returnedVal = False
            End If
        End If
    End Sub
    
    ' Verfügbar, wenn der Formelbereich definiert ist.
    Sub GetEnabledFormula(control As IRibbonControl, ByRef returnedVal)
        
        Call GetEnabledGeoToolsTable(control, returnedVal)
        
        If (returnedVal) Then
            If (Not ThisWorkbook.AktiveTabelle.ExistsLokalerZellname(strFormel)) Then
                returnedVal = False
            End If
        End If
    End Sub
    
' End Region


' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
