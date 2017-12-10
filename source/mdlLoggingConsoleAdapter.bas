Attribute VB_Name = "mdlLoggingConsoleAdapter"
'**************************************************************************************************
' GeoTools: Excel-Werkzeuge (nicht nur) für Geodäten.
' Copyright © 2003 - 2017  Robert Schwenn  (Lizenzbestimmungen siehe Modul "Lizenz_History")
'**************************************************************************************************

'====================================================================================
'Modul mdlLoggingConsoleAdapter
'====================================================================================
' Stellt eine Schnittstelle zu LoggingConsole.NET zur Verfügung:
' 
' - Die hier bereitgestellten Methoden lösen nie einen Fehler aus.
' - Ist das Actions.NET-AddIn geladen, werden die Log-Meldungen darüber
'   an die LoggingConsole.NET weitergeleitet
' - Nach einer festgelegten Anzahl an Fehlversuchen wird das Actions.NET-AddIn
'   als nicht geladen betrachtet und keine weiteren Versuche unternommen (Performace).
'====================================================================================
Option Explicit

Const MAX_TRIALS  As Integer = 99

' IsLoggingNotAvailable wird auf True gesetzt, nachdem MAX_TRIALS Fehler aufgetreten sind.
Dim IsLoggingNotAvailable As Boolean
Dim FailedTrialsCount     As Integer


Sub Echo(ByVal Message As String)
    On Error GoTo Catch
    If (Not IsLoggingNotAvailable) Then
        Call Application.Run("LoggingConsoleLogInfo", Message)
    End If
    
    Exit Sub
    Catch:
    Call TrackError
End Sub

Sub ErrEcho(ByVal Message As String)
    Dim ErrInfo As String
    
    If (Err.Number <> 0) Then
        ErrInfo = "FEHLER in         : '" & Err.Source & "':" & vbNewLine & _
                  "Fehlernummer      : "  & CStr(Err.Number) & vbNewLine & _
                  "Fehlerbeschreibung: "  & Err.Description
    End If
    
    On Error GoTo Catch
    
    If (Not IsLoggingNotAvailable) Then
        If (ErrInfo <> "") Then
            'Workaround for unprintable character at the end of some error descriptions
            If (Asc(Right(ErrInfo, 1)) < 32) Then ErrInfo = Left(ErrInfo, Len(ErrInfo) - 1)
            
            Call Application.Run("LoggingConsoleLogError", ErrInfo)
        End If
        Call Application.Run("LoggingConsoleLogError", Message)
    End If
    
    Exit Sub
    Catch:
    Call TrackError
End Sub

Sub WarnEcho(ByVal Message As String)
    On Error GoTo Catch
    If (Not IsLoggingNotAvailable) Then
        Call Application.Run("LoggingConsoleLogWarning", Message)
    End If
    
    Exit Sub
    Catch:
    Call TrackError
End Sub

Sub DebugEcho(ByVal Message As String)
    On Error GoTo Catch
    If (Not IsLoggingNotAvailable) Then
        Call Application.Run("LoggingConsoleLogDebug", Message)
    End If
    
    Exit Sub
    Catch:
    Call TrackError
End Sub

Sub ShowConsole()
    'Shows the Logging Console Dock Panel.
    On Error GoTo Catch
    Call Application.Run("LoggingConsoleShow")
    
    Exit Sub
    Catch:
    MsgBox "Protokollierung steht nicht zur Verfügung!" & vbNewLine & "(Eventuell wurde Excel per Automation gestartet)"
End Sub

Sub TrackError()
    FailedTrialsCount = FailedTrialsCount + 1
    If (FailedTrialsCount > MAX_TRIALS) Then
        IsLoggingNotAvailable = True
    End If
End Sub


' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
