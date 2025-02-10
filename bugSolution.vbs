Function MyFunction(param1, param2)
  On Error GoTo ErrorHandler
  ' ... some code ...
  If param1 = "" Then
    Err.Raise 13, , "Parameter 1 cannot be empty"
  End If
  ' ... more code ...
  Exit Function
ErrorHandler:
  MsgBox "Error: " & Err.Description, vbCritical
  ' Clean up or take appropriate action
End Function