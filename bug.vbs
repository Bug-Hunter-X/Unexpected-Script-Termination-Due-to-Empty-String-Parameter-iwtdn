Function MyFunction(param1, param2)
  ' ... some code ...
  If param1 = "" Then
    Err.Raise 13, , "Parameter 1 cannot be empty"
  End If
  ' ... more code ...
End Function