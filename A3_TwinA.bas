Attribute VB_Name = "Module1"
Option Explicit

'Declaring record type
Type StudentRec
    LastName As String * 20
    FirstName As String * 15
    HF As String * 3
    Mark As Integer
End Type

Public Sub Initialize(S() As StudentRec, N As Integer, FName As String, ByVal M As Integer)
    Dim K As Integer
    
    For K = 1 To M
        'Initialize each field of record
        S(K).FirstName = ""
        S(K).HF = ""
        S(K).LastName = ""
        S(K).Mark = 0
    Next K
    
    'Initialize NumStudents to 0
    N = 0
End Sub

Public Function GetFile(Dialog As Control) As String
    Dialog.FileName = ""
    Dialog.InitDir = App.Path
    Dialog.Filter = "Record Files|*.rec|Text Files|*.txt"
    Dialog.ShowOpen
    
    'Return the file name
    GetFile = Dialog.FileName
End Function

Public Function GetSave(Dialog As Control) As String
    Dialog.FileName = ""
    Dialog.InitDir = App.Path
    Dialog.Filter = "Record Files|*.rec"
    Dialog.ShowSave
    
    'Return the name of the save file
    GetSave = Dialog.FileName
End Function

'Reading data from a record file
Public Sub ReadRecordFile(S() As StudentRec, N As Integer, ByVal FName As String, ByVal RecLen As Integer)
    Dim K As Integer
    
    K = 0
    Open FName For Random As #1 Len = RecLen
    Do While Not EOF(1)
        K = K + 1
        'Put data into record array
        Get #1, K, S(K)
    Loop
    Close #1
    'Return the number of students
    N = K - 1
End Sub

'Reading data from a text file
Public Sub ReadTextFile(S() As StudentRec, N As Integer, ByVal FName As String)
    Dim K As Integer
    
    K = 0
    
    Open FName For Input As #1
    Do While Not EOF(1)
        K = K + 1
        'Put data into record array
        Input #1, S(K).LastName
        Input #1, S(K).FirstName
        Input #1, S(K).HF
        Input #1, S(K).Mark
    Loop
    Close #1
    'Return the number of students
    N = K
End Sub

'Display information
Public Sub Display(lblTotal As Label, PicBox As PictureBox, S() As StudentRec, N As Integer)
    Dim K As Integer
    
    'Clear picture box
    PicBox.Cls
    
    'Print headings
    PicBox.Print Tab(5); "First Name"; Tab(30); "Last Name"; Tab(61); "HF"; Tab(72); "Mark"
    PicBox.Print
    
    For K = 1 To N
        'Display data
        PicBox.Print Format$(Trim$(Str$(K)), "@@"); ". "; S(K).FirstName; Spc(10); S(K).LastName; Spc(10); Format$(Trim$(S(K).HF), "@@@"); Spc(10); Format$(Trim$(Str$(S(K).Mark)), "@@@")
    Next K
    
    'Display number of students
    lblTotal.Caption = Trim$(Str$(N))
End Sub

'Get the name of the save file
Public Sub SaveFile(S() As StudentRec, ByVal N As Integer, ByVal FName As String, ByVal RecLen As Integer)
    Dim K As Integer
    
    'Go to ErrorHandler when an error has occurred
    On Error GoTo ErrorHandler
    'Delete file
    'If the file does not exist, it goes to ErrorHandler
    Kill FName
    
    'Create record file
    Open FName For Random As #1 Len = RecLen
    For K = 1 To N
        'Inputting information to record file
        Put #1, K, S(K)
    Next K
    Close #1
    'Ends procedure so the program doesn't go to ErrorHandler
    Exit Sub
    
'An error has occurred
ErrorHandler:
    'If an error has occurred, go to the next line
    Resume Next
End Sub

Public Function SeqSearch(S() As StudentRec, ByVal N As Integer, ByVal Target As String) As Integer
    Dim K As Integer
    Dim Position As Integer
    
    Position = -1
    K = 1
    Do While (Position = -1) And (K <= N)
        If UCase$(Target) = UCase$(RTrim$(S(K).LastName)) Then
            Position = K
        End If
        K = K + 1
    Loop
    SeqSearch = Position
End Function

Public Function Binary_Search(S() As StudentRec, ByVal N As Integer, ByVal Target As String) As Integer
    Dim First As Integer
    Dim Last As Integer
    Dim Middle As Integer
    Dim Found As Boolean
    
    First = 1
    Last = N
    Binary_Search = -1
    Found = False
    
    Do While (Not Found) And (First <= Last)
        Middle = (First + Last) \ 2
        If Target < Trim$(S(Middle).LastName) Then
            Last = Middle - 1
        ElseIf Target > Trim$(S(Middle).LastName) Then
            First = Middle + 1
        Else
            Found = True
            Binary_Search = Middle
        End If
    Loop
End Function

Public Sub BubbleSort(S() As StudentRec, ByVal N As Integer)
    Dim K As Integer
    Dim X As Integer
    
    For K = 1 To N - 1
        For X = 1 To N - K
            If Trim$(S(X).LastName) > Trim$(S(X + 1).LastName) Then
                Swap S(X), S(X + 1)
            End If
        Next X
    Next K
End Sub

Public Sub Swap(K As StudentRec, X As StudentRec)
    Dim Temp As StudentRec
    
    Temp = K
    K = X
    X = Temp
End Sub



