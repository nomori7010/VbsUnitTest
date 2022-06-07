Function Sum(n, m)
    Sum = Val(n) + Val(m)
End Function
Function Diff(n, m)
    Diff = Val(n) - Val(m)
End Function
Function IsOdd(n)
    IsOdd = Val(n) mod 2 = 1
End Function
Function IsEven(n)
    IsEven = Val(n) mod 2 = 0
End Function
Function Val(value)
    If IsNull(value) Or IsEmpty(value) Then
        Val = 0
        Exit Function
    End If
    If IsNumeric(value) Then
        Val = CDbl(value)
        Exit Function
    End If
    Val = 0
End Function