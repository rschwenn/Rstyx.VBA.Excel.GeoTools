Attribute VB_Name = "mdlRibbon"
'===============================================================================
'Modul mdlRibbon                                                                  
'===============================================================================

Option Explicit


Public oRibbon As IRibbonUI

' Region "Referenz auf das Ribbon-Objekt"
    
    ' Siehe http://social.msdn.microsoft.com/Forums/office/en-US/99a3f3af-678f-4338-b5a1-b79d3463fb0b/how-to-get-the-reference-to-the-iribbonui-in-vba
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)
    
    ' Initialisierung der RibbonUI: Speichern einer Referenz auf das Ribbon-Objekt
    ' und als Backup eines entsprechenden Integer-Zeigers in die Add-In-interne Tabelle.
    Public Sub OnGeoToolsRibbonLoad(ribbon As IRibbonUI)
        Set oRibbon = ribbon
        tabGeoTools.Range("A1").Value = ObjPtr(ribbon)
    End Sub
    
    ' Beziehen einer Referenz auf das Ribbon-Objekt (Sollte auch nach Fehler im Add-In funktionieren).
    Function getRibbon() As IRibbonUI
        ' "oRibbon" ist normalerweise nur dann "Nothing", wenn das AddIn wegen eines Fehlers gestoppt wurde.
        ' Dann kann der vorher gespeicherte Zeigert verwendet werden.
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
        
        Set getRibbon = oRibbon
    End Function
    
    ' Status-Aktualisierung aller Ribbon-Steuerelemente erzwingen.
    Public Sub UpdateRibbon(Optional ByVal keinAktivesBlatt As Boolean = False)
        On Error Resume Next
        getRibbon().Invalidate
        On Error Goto 0
    End Sub
    
' End Region


' Region "###  Verschieben  ###"
    
    Function getAktiveTabelle() As CtabAktiveTabelle
        If (oAktiveTabelle Is Nothing) Then
            Set oAktiveTabelle = New CtabAktiveTabelle
        End If
        Set getAktiveTabelle = oAktiveTabelle
    End Function
    
    Function getKonfig() As CdatKonfig
        If (oKonfig Is Nothing) Then
            Set oKonfig = New CdatKonfig
        End If
        Set getKonfig = oKonfig
    End Function
    
' End Region


' Region "No Config Button"
    
    Sub GetVisibleNoConfigButton(control As IRibbonControl, ByRef returnedVal)
        returnedVal = True
        On Error Resume Next
        returnedVal = (Not getKonfig().EinstellungenGelesen)
        On Error Goto 0
    End Sub
    
    Sub GetSupertipNoConfigButton(control As IRibbonControl, ByRef returnedVal)
        Dim cfg As String
        cfg = Verz(ThisWorkbook.Path) & "\" & VorName(ThisWorkbook.Name) & "_cfg.xlsx"
        returnedVal = GetInfoKeineKonfig()
    End Sub
    
    Sub NoConfigButtonAction(ByVal control As IRibbonControl)
        Call InfoKeineKonfig
    End Sub
    
' End Region



' Region "Action Buttons"
    
    Sub LogButtonAction(ByVal control As IRibbonControl)
        Call Protokoll
    End Sub
    
    Sub HelpButtonAction(ByVal control As IRibbonControl)
        Call Hilfe_Komplett
    End Sub
    
    Sub InfoButtonAction(ByVal control As IRibbonControl)
        Call GeoTools_Info
    End Sub
    
    Sub ImportExportButtonAction(ByVal control As IRibbonControl)
        Call ExpimManager
    End Sub
    
    Sub TableStructureButtonAction(ByVal control As IRibbonControl)
        Call TabellenStruktur
    End Sub
    
    Sub FormatButtonAction(ByVal control As IRibbonControl)
        Call FormatDaten
    End Sub
    
    Sub FormulaButtonAction(ByVal control As IRibbonControl)
        Call UebertragenFormeln
    End Sub
    
    Sub DeleteButtonAction(ByVal control As IRibbonControl)
        Call LoeschenDaten
    End Sub
    
    Sub CalcDiffsButtonAction(ByVal control As IRibbonControl)
        Call Mod_FehlerVerbesserung
    End Sub
    
    Sub CalcCantButtonAction(ByVal control As IRibbonControl)
        Call Mod_UeberhoehungAusBemerkung
    End Sub
    
    Sub CalcHorizontalToCantedButtonAction(ByVal control As IRibbonControl)
        Call Mod_Transfo_Tk2Gls
    End Sub
    
    Sub CalcCantedToHorizontalButtonAction(ByVal control As IRibbonControl)
        Call Mod_Transfo_Gls2Tk
    End Sub
    
    Sub InterpolButtonAction(ByVal control As IRibbonControl)
        Call Selection2Interpolationsformel
    End Sub
    
    Sub DuplicatesButtonAction(ByVal control As IRibbonControl)
        Call Selection2MarkDoppelteWerte
    End Sub
    
    Sub BlankLinesButtonAction(ByVal control As IRibbonControl)
        Call insertLines
    End Sub
    
    Sub EditFileButtonAction(ByVal control As IRibbonControl)
        Call DateiBearbeiten
    End Sub
    
    Sub SetFooterButtonAction(ByVal control As IRibbonControl)
        Call SchreibeFusszeile_1
    End Sub
    
' End Region

' Region "Toggle Buttons"
    
    Sub FmtOptStripesButtonAction(control As IRibbonControl, pressed As Boolean)
        On Error Resume Next
        getAktiveTabelle().FormatDatenMitStreifen = pressed
        On Error Goto 0
    End Sub
    
    Sub FmtOptBackgroundButtonAction(control As IRibbonControl, pressed As Boolean)
        On Error Resume Next
        getAktiveTabelle().FormatDatenOhneFuellung = pressed
        On Error Goto 0
    End Sub
    
    Sub FmtOptPrecisionButtonAction(control As IRibbonControl, pressed As Boolean)
        On Error Resume Next
        getAktiveTabelle().FormatDatenNKStellenSetzen = pressed
        On Error Goto 0
    End Sub
    
    Sub CalcOptOverrideButtonAction(control As IRibbonControl, pressed As Boolean)
        On Error Resume Next
        getAktiveTabelle().ModOpt_VorhWerteUeberschreiben = pressed
        On Error Goto 0
    End Sub
    
    Sub CalcOptKeepFormulasButtonAction(control As IRibbonControl, pressed As Boolean)
        On Error Resume Next
        getAktiveTabelle().ModOpt_FormelnErhalten = pressed
        On Error Goto 0
    End Sub
    
    
    Sub FmtOptStripesButtonGetPressed(control As IRibbonControl, ByRef returnedVal)
        On Error Resume Next
        returnedVal = getAktiveTabelle().FormatDatenMitStreifen
        On Error Goto 0
    End Sub
    
    Sub FmtOptBackgroundButtonGetPressed(control As IRibbonControl, ByRef returnedVal)
        On Error Resume Next
        returnedVal = getAktiveTabelle().FormatDatenOhneFuellung
        On Error Goto 0
    End Sub
    
    Sub FmtOptPrecisionButtonGetPressed(control As IRibbonControl, ByRef returnedVal)
        On Error Resume Next
        returnedVal = getAktiveTabelle().FormatDatenNKStellenSetzen
        On Error Goto 0
    End Sub
    
    Sub CalcOptOverrideButtonGetPressed(control As IRibbonControl, ByRef returnedVal)
        On Error Resume Next
        returnedVal = getAktiveTabelle().ModOpt_VorhWerteUeberschreiben
        On Error Goto 0
    End Sub
    
    Sub CalcOptKeepFormulasButtonGetPressed(control As IRibbonControl, ByRef returnedVal)
        On Error Resume Next
        returnedVal = getAktiveTabelle().ModOpt_FormelnErhalten
        On Error Goto 0
    End Sub
    
' End Region


' Region "Controls Enabled / Disabled"
    
    ' Verfügbar, wenn das aktive Blatt eine Tabelle ist.
    Sub GetEnabledTable(control As IRibbonControl, ByRef returnedVal)
        
        returnedVal = False
        
        If (Not (ActiveCell Is Nothing)) Then
            If (Not (Selection Is Nothing)) Then
                If (TypeOf Selection Is Range) Then
                    returnedVal = True
                End If
            End If
        End If
    End Sub
    
    ' Verfügbar, wenn der Infotraeger definiert ist ("GeoTools-Tabelle").
    Sub GetEnabledGeoToolsTable(control As IRibbonControl, ByRef returnedVal)
        
        Call GetEnabledTable(control, returnedVal)
        
        If (returnedVal) Then
            returnedVal = False
            Dim oTable As CtabAktiveTabelle
            Set oTable = getAktiveTabelle()
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
            If (Not getAktiveTabelle().ExistsLokalerZellname(strFliesskomma)) Then
                returnedVal = False
            End If
        End If
    End Sub
    
    ' Verfügbar, wenn der Formelbereich definiert ist.
    Sub GetEnabledFormula(control As IRibbonControl, ByRef returnedVal)
        
        Call GetEnabledGeoToolsTable(control, returnedVal)
        
        If (returnedVal) Then
            If (Not getAktiveTabelle().ExistsLokalerZellname(strFormel)) Then
                returnedVal = False
            End If
        End If
    End Sub
    
' End Region

' Anzahl NK-Stellen wurde via GUI geändert.
Sub FmtOptPrecisionNumberChange(control As IRibbonControl, text As String)
    On Error Resume Next
    getAktiveTabelle().FormatDatenNKStellenAnzahl = cInt(text)
    On Error Goto 0
End Sub

'Callback for FmtOptPrecisionNumber getLabel
Sub GetTextFmtOptPrecisionNumber(control As IRibbonControl, ByRef returnedVal)
    On Error Resume Next
    returnedVal = getAktiveTabelle().FormatDatenNKStellenAnzahl
    On Error Goto 0
End Sub


' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
