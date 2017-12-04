Attribute VB_Name = "Module1"
Option Explicit

Public Function Factorial(ByVal N As Integer) As Long
    If N > 1 Then
        Factorial = N * Factorial(N - 1)
    Else
        Factorial = 1
    End If
End Function

Public Function Power(ByVal B As Integer, ByVal E As Integer) As Single
    If E > 1 Then
        Power = B * Power(B, E - 1)
    Else
        Power = B
    End If
End Function

Public Function Fibonacci(ByVal N As Integer) As Long
    If N > 1 Then
        Fibonacci = Fibonacci(N - 1) + Fibonacci(N - 2)
    ElseIf N = 0 Then
        Fibonacci = 0
    Else
        Fibonacci = 1
    End If
End Function

Public Sub Reverse(N() As Integer, ByVal K As Integer)
    If K > 1 Then
        frmMain.Print N(K)
        Reverse N(), K - 1
    Else
        frmMain.Print N(1)
    End If
End Sub
