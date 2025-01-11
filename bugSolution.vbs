Function f(a, b)
  If VarType(a) = VarType(b) And a = b Then
    MsgBox "a and b are equal"
  Else
    MsgBox "a and b are not equal"
  End If
End Function