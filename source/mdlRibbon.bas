Attribute VB_Name = "mdlRibbon"
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2014-2025  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'===============================================================================
' Modul mdlRibbon                                                                   
'===============================================================================
' Stellt Zugriff auf das Menüband sowie Ribbon-Callbacks Verfügung.

Option Explicit


Private oRibbon As IRibbonUI

' Region "Ribbon-Objekt (Referenz, Update)"
    
    ' Siehe http://social.msdn.microsoft.com/Forums/office/en-US/99a3f3af-678f-4338-b5a1-b79d3463fb0b/how-to-get-the-reference-to-the-iribbonui-in-vba
    'Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByRef lpSource As Any, ByVal cBytes As Long)
    
    ' Bei Initialisierung der RibbonUI: Speichern einer Referenz auf das Ribbon-Objekt
    ' und als Backup eines entsprechenden Integer-Zeigers in die Add-In-interne Tabelle.
    Public Sub OnGeoToolsRibbonLoad(ribbon As IRibbonUI)
        Set oRibbon = ribbon
        tabGeoTools.Range("A1").Value = ObjPtr(ribbon)
    End Sub
    
    ' x64: Beziehen einer Referenz auf das Ribbon-Objekt (Sollte auch nach Fehler im Add-In funktionieren).
    Function getGeoToolsRibbon() As IRibbonUI
        ' "oRibbon" ist normalerweise nur dann "Nothing", wenn das AddIn wegen eines Fehlers gestoppt wurde.
        ' Dann kann der vorher gespeicherte Zeiger verwendet werden.
        ' ABER: Wenn das AddIn nicht schreibgeschützt ist, kann der Zeiger auch veraltet sein.
        '       => Dann stürzt Excel ab und nichts geht mehr.
        If (oRibbon Is Nothing) Then
            If (ThisWorkbook.ReadOnly) Then
                If (Not tabGeoTools Is Nothing) Then
                    Dim ribbonPointer As LongPtr
                    ribbonPointer = tabGeoTools.Range("A1").value
                    If (ribbonPointer > 0) Then
                        On Error Resume Next  ' Nützt nix!
                        Call CopyMemory(oRibbon, ribbonPointer, LenB(ribbonPointer))
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
        
        Set getGeoToolsRibbon = oRibbon
    End Function
    
    ' x32: Beziehen einer Referenz auf das Ribbon-Objekt (Sollte auch nach Fehler im Add-In funktionieren).
    'Function getGeoToolsRibbon() As IRibbonUI
        '' "oRibbon" ist normalerweise nur dann "Nothing", wenn das AddIn wegen eines Fehlers gestoppt wurde.
        '' Dann kann der vorher gespeicherte Zeiger verwendet werden.
        '' ABER: Wenn das AddIn nicht schreibgeschützt ist, kann der Zeiger auch veraltet sein.
        ''       => Dann stürzt Excel ab und nichts geht mehr.
        'If (oRibbon Is Nothing) Then
        '    If (ThisWorkbook.ReadOnly) Then
        '        Dim ribbonPointer As Long
        '        ribbonPointer = tabGeoTools.Range("A1").value
        '        If (ribbonPointer > 0) Then
        '            On Error Resume Next  ' Nützt nix!
        '            Call CopyMemory(oRibbon, ribbonPointer, 4)
        '            On Error GoTo 0
        '        End If
        '    End If
        'End If
        '
        'Set getGeoToolsRibbon = oRibbon
    'End Function
    
    ' Status-Aktualisierung aller Ribbon-Steuerelemente erzwingen.
    Public Sub UpdateGeoToolsRibbon(Optional ByVal keinAktivesBlatt As Boolean = False)
        On Error Resume Next
        getGeoToolsRibbon().Invalidate
        call ClearStatusBarDelayed(3)
        On Error Goto 0
    End Sub
    
' End Region


' Region "No Config Button"
    
    Sub GetVisibleNoConfigButton(control As IRibbonControl, ByRef returnedVal)
        returnedVal = True
        On Error Resume Next
        returnedVal = (Not ThisWorkbook.Konfig.KonfigVerfuegbar)
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

' Region "Precision Dropdown"
    
    ' siehe https://www.contextures.com/excelribbonmacrostab.html#download
    
    ' Select appropriate item.
    Sub PrecisionDropdownGetSelectedItemIndex(ByVal control As IRibbonControl, ByRef ItemIndex)
        On Error Resume Next
        ItemIndex = ThisWorkbook.AktiveTabelle.FormatDatenNKStellenAnzahl
        On Error Goto 0
    End Sub
    
    ' Anzahl NK-Stellen wurde via GUI geändert.
    Sub PrecisionDropdownAction(ByVal control As IRibbonControl, selectedID As String, selectedIndex As Integer)
        On Error Resume Next
        ThisWorkbook.AktiveTabelle.FormatDatenNKStellenAnzahl = selectedIndex
        On Error Goto 0
    End Sub
    
    
' End Region

' Region "Editor Dropdown"
    
    ' siehe https://www.contextures.com/excelribbonmacrostab.html#download
    
    ' Init EditorDropdown (part 1).
    Sub EditorDropdownGetItemCount(ByVal control As IRibbonControl, ByRef count)
        count = ubound(ThisWorkbook.SysTools.Editoren, 1) + 1
    End Sub
    
    ' Init EditorDropdown (part 2).
    Sub EditorDropdownGetItemID(ByVal control As IRibbonControl, Index As Integer, ByRef ItemID)
        Dim Editors as Variant
        Editors = ThisWorkbook.SysTools.Editoren
        ItemID  = Editors(Index, 1)
    End Sub
    
    ' Init EditorDropdown (part 3).
    Sub EditorDropdownGetItemLabel(ByVal control As IRibbonControl, Index As Integer, ByRef ItemLabel)
        Dim Editors as Variant
        Editors   = ThisWorkbook.SysTools.Editoren
        ItemLabel = Editors(Index, 1)
    End Sub
    
    ' Select appropriate item.
    Sub EditorDropdownGetSelectedItemID(ByVal control As IRibbonControl, ByRef ItemID)
        ItemID = ThisWorkbook.SysTools.Editor
    End Sub
    
    ' Editor wurde via GUI geändert.
    Sub EditorDropdownAction(ByVal control As IRibbonControl, selectedID As String, selectedIndex As Integer)
        On Error Resume Next
        ThisWorkbook.SysTools.Editor = selectedID
        On Error Goto 0
    End Sub
    
    ' Verfügbar, wenn mindestens ein unterstützter Editor gefunden worden ist.
    Sub EditorDropdownGetEnabled(control As IRibbonControl, ByRef returnedVal)
        returnedVal = (not (ThisWorkbook.SysTools.Editor = ThisWorkbook.SysTools.cKeinEditor))
    End Sub
    
    
' End Region

' Region "Action Buttons"
    
    Sub GeoToolsButtonAction(ByVal control As IRibbonControl)
        'On Error Resume Next
        Select Case control.ID
            Case "InfoButton"                   : Call GeoTools_Info
            Case "HelpButton"                   : Call Hilfe
            Case "ManualButton"                 : Call Handbuch
            Case "LogButton"                    : Call Protokoll
            Case "ImportExportButton"           : Call ExpimManager
            Case "TableStructureButton"         : Call TabellenStruktur
            Case "FormatButton"                 : Call FormatDaten
            Case "CalcDiffsButton"              : Call Mod_FehlerVerbesserung
            Case "CalcParseInfoTextButton"      : Call Mod_InfoTextAuswerten
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
    
    ' Verfügbar, wenn Makros im aktiven Fenster ausführbar sind.
    Sub GetEnabledMacrosExecutable(control As IRibbonControl, ByRef returnedVal)
        returnedVal = (IsMacrosExecutable() And (Application.ActiveProtectedViewWindow Is Nothing))
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
        
        returnedVal = (returnedVal And IsMacrosExecutable())
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
