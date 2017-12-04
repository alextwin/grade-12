Attribute VB_Name = "Module1"
Option Explicit

Public Function GetDay(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer) As String
    Dim Century_Year As Integer
    Dim Century As Integer
    Dim Weekday As Integer
    Dim W As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    
    If Month = 1 Or Month = 2 Then
        Month = Month + 10
        Year = Year - 1
    Else
        Month = Month - 2
    End If
    
    Century_Year = Val(Right$(Trim$(Str$(Year)), 2))
    Century = Val(Left$(Trim$(Str$(Year)), 2))
    
    W = (13 * Month - 1) \ 5
    X = Century_Year \ 4
    Y = Century \ 4
    Z = W + X + Y + Day + Century_Year - 2 * Century + 7777
    Weekday = Z Mod 7
    
'    Select Case Weekday
'        Case 0
'            GetDay = "Sunday"
'        Case 1
'            GetDay = "Monday"
'        Case 2
'            GetDay = "Tuesday"
'        Case 3
'            GetDay = "Wednesday"
'        Case 4
'            GetDay = "Thursday"
'        Case 5
'            GetDay = "Friday"
'        Case 6
'            GetDay = "Saturday"
'    End Select
    GetDay = WeekdayName(Weekday + 1)
End Function

Public Sub CheckDays(cDay As ComboBox, ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer, ByVal M_DAY As Integer)
    Const LEAPYEAR = 29
    Dim K As Integer
    Dim CurrentlySelected As Integer
    
    CurrentlySelected = Val(cDay.List(cDay.ListIndex))
    cDay.Clear
    
    Select Case Month
        Case 1, 3, 5, 7, 8, 10, 12
            For K = 0 To M_DAY - 1
                cDay.AddItem Trim$(Str$(K + 1)), K
            Next K
        Case 4, 6, 9, 11
            For K = 0 To M_DAY - 2
                cDay.AddItem Trim$(Str$(K + 1)), K
            Next K
            If CurrentlySelected > 30 Then
                CurrentlySelected = 30
            End If
        Case Else
            If Year Mod 4 = 0 And (Year Mod 100 <> 0 Or Year Mod 400 = 0) Then
                For K = 0 To LEAPYEAR - 1
                    cDay.AddItem Trim$(Str$(K + 1)), K
                Next K
                If CurrentlySelected > 29 Then
                    CurrentlySelected = 29
                End If
            Else
                For K = 0 To LEAPYEAR - 2
                    cDay.AddItem Trim$(Str$(K + 1)), K
                Next K
                If CurrentlySelected > 28 Then
                    CurrentlySelected = 28
                End If
            End If
    End Select
    
    cDay.ListIndex = CurrentlySelected - 1
End Sub
