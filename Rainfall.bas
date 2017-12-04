Attribute VB_Name = "Module1"
Option Explicit
Global Const MAXCITY = 1000
Global Const MONTH = 12
Global NumCity As Integer
Global LeastMonth As String
Global Least As Integer
Global LeastCity As String
Global City(1 To MAXCITY) As String
Global Temp(1 To MAXCITY, 1 To MONTH) As Integer

Public Sub Initialize(A As Single, MA() As Single)
    Dim K As Integer
    Dim X As Integer
    
    NumCity = 0
    A = 0
    LeastMonth = ""
    Least = 0
    LeastCity = ""
    For K = 1 To MAXCITY
        City(K) = ""
        MA(K) = 0
        For X = 1 To MONTH
            Temp(K, X) = 0
        Next X
    Next K
End Sub

Public Function ReadFile(ByVal FName As String) As Integer
    Dim K As Integer
    Dim X As Integer
    
    K = 0
    
    Open FName For Input As #1
    Do While Not EOF(1)
        K = K + 1
        Input #1, City(K)
        For X = 1 To MONTH
            Input #1, Temp(K, X)
        Next X
    Loop
    Close #1
    
    ReadFile = K
End Function

Public Sub MonthlyAverage(MA() As Single)
    Dim K As Integer
    Dim X As Integer
    Dim Sum As Integer
    
    Sum = 0
    
    For K = 1 To NumCity
        For X = 1 To MONTH
            Sum = Sum + Temp(K, X)
        Next X
        MA(K) = Sum / MONTH
    Next K
End Sub

Public Sub Display(Pic As Variant, MA() As Single, ByVal A As Single)
    Const MTH = "JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC"
    Dim K As Integer
    Dim X As Integer
    
    Pic.Cls
    
    Pic.Print "City"; Tab(20);
    For K = 1 To MONTH
        Pic.Print Right$(Left$(MTH, K * 3), 3); Spc(5);
    Next K
    Pic.Print " "; "Average"
    
    For K = 1 To NumCity
        Pic.Print City(K); Tab(20);
        For X = 1 To MONTH
            Pic.Print Temp(K, X); Spc(4);
        Next X
        Pic.Print MA(K)
    Next K
    
    Pic.Print
    Pic.Print "The overall monthly average rainfall of all cities is "; A
    Pic.Print "The city with the lowest rainfall is "; LeastCity; " in "; LeastMonth;
    Pic.Print " with"; Least; "cm of rain"
End Sub

Public Function OverallAverage() As Single
    Dim K As Integer
    Dim X As Integer
    Dim Sum As Integer
    
    Sum = 0
    
    For K = 1 To NumCity
        For X = 1 To MONTH
            Sum = Sum + Temp(K, X)
        Next X
    Next K
    
    OverallAverage = Sum / (NumCity * MONTH)
End Function

Public Sub LeastRain()
    Const MTH = "JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC"
    Dim K As Integer
    Dim X As Integer
    Dim LCity As Integer
    Dim LMonth As Integer
    
    LMonth = 1
    LCity = 1
    Least = Temp(1, 1)
    
    For K = 1 To NumCity
        For X = 2 To MONTH
            If Temp(K, X) < Least Then
                Least = Temp(K, X)
                LMonth = X
                If K <> LCity Then
                    LCity = K
                End If
            End If
        Next X
    Next K
    
    LeastMonth = Right$(Left$(MTH, LMonth * 3), 3)
    LeastCity = City(LCity)
End Sub
