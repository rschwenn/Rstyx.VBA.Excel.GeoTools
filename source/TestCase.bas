Attribute VB_Name = "TestCase"

Dim oConsole As LoggingConsole


'Testcase for Logging Console
Sub TestLogging()
  If (oConsole Is Nothing) Then Set oConsole = New LoggingConsole
  
  
  'queue messages
  For i = 1 To 20
    oConsole.logInfo "infoline " & CStr(i)
  Next
  
  
  'show Console
  oConsole.Show vbModeless
  
  'queue messages
  For i = 1 To 20
    oConsole.logDebug "debug message " & CStr(i)
  Next
  
  
  'Change settings
  oConsole.LogSource = "TestLogging()"
  oConsole.IncludeDate = True
  
  
  'queue messages
  For i = 1 To 20
    oConsole.logError "error message " & CStr(i)
  Next
  
  For i = 1 To 20
    oConsole.logWarning "Warning message " & CStr(i)
  Next
  
  
  'save normal Log (info) as specified file
  oConsole.saveLog oConsole.PAGE_LOG, "C:\testinfo.log"
  
  'save error Log (filename via dialog)
  oConsole.saveLog oConsole.PAGE_ERRORLOG
  
  'Hide Console
  ''oConsole.hide
End Sub



'Methods for controlling the console

Sub showConsole()
  If (oConsole Is Nothing) Then Set oConsole = New LoggingConsole
  oConsole.Show vbModeless
End Sub

Sub clearErrors()
  If (Not oConsole Is Nothing) Then
    oConsole.clearLog oConsole.PAGE_ERRORLOG
  End If
End Sub

Sub ShowWarnings()
  If (oConsole Is Nothing) Then Set oConsole = New LoggingConsole
  oConsole.ActiveLogPage = oConsole.PAGE_WARNINGLOG
End Sub

Sub StopLogging()
  If (Not oConsole Is Nothing) Then
    oConsole.Hide
    Set oConsole = Nothing
  End If
End Sub

