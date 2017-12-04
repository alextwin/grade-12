Attribute VB_Name = "Module1"
Option Explicit

'Declaring record type
Type StudentRec
    LastName As String * 20
    FirstName As String * 15
    HF As String * 3
    Mark As Integer
End Type

Global Const MAX = 100
Global Student(1 To MAX) As StudentRec
Global NumStudents As Integer
Global CurrentRecord As Integer

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
Public Sub Display(lblTotal As Label, lstBox As ListBox, S() As StudentRec, ByVal N As Integer)
    Dim K As Integer
    
    'Clear picture box
    lstBox.Clear
    
    For K = 1 To N
        'Display data
        lstBox.AddItem VBA.Format$(VBA.Trim$(VBA.Str$(K)), "@@") & ". " & S(K).LastName & "               " & S(K).FirstName & "               " & VBA.Format$(VBA.Trim$(S(K).HF), "@@@") & "              " & VBA.Format$(VBA.Trim$(VBA.Str$(S(K).Mark)), "@@@")
    Next K
    
    'Display number of students
    lblTotal.Caption = VBA.Trim$(VBA.Str$(N))
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

'Add record to list box
Public Sub AddRecord(lblTotal As Label, lstBox As ListBox, S() As StudentRec, ByVal N As Integer)
    lstBox.AddItem VBA.Format$(VBA.Trim$(VBA.Str$(N)), "@@") & ". " & S(N).LastName & "               " & S(N).FirstName & "               " & VBA.Format$(VBA.Trim$(S(N).HF), "@@@") & "              " & VBA.Format$(VBA.Trim$(VBA.Str$(S(N).Mark)), "@@@"), N - 1
    'Update total
    lblTotal.Caption = Val(N)
End Sub

'Load modify form
Public Sub LoadModify(fModify As Form, lstBox As ListBox, S() As StudentRec)
    Load fModify
    With fModify
        .Caption = "Modify Record"
        .cmdAdd.Visible = False
        .cmdModify.Visible = True
        'Put selected record's information in the text boxes
        .txtLastName.Text = VBA.RTrim$(S(CurrentRecord).LastName)
        .txtFirstName.Text = VBA.RTrim$(S(CurrentRecord).FirstName)
        .txtHF.Text = VBA.RTrim$(S(CurrentRecord).HF)
        .txtMark.Text = S(CurrentRecord).Mark
        .Show vbModal
    End With
End Sub

'Modify the record
Public Sub Modify(lstBox As ListBox, S() As StudentRec, ByVal LN As String, ByVal FN As String, ByVal HF As String, ByVal M As Integer)
    'Replace record with new the new information
    S(CurrentRecord).LastName = LN
    S(CurrentRecord).FirstName = FN
    S(CurrentRecord).HF = HF
    S(CurrentRecord).Mark = M
    'Remove the old record information on list box
    lstBox.RemoveItem CurrentRecord - 1
    'Display to list box
    lstBox.AddItem VBA.Format$(VBA.Trim$(VBA.Str$(CurrentRecord)), "@@") & ". " & S(CurrentRecord).LastName & "               " & S(CurrentRecord).FirstName & "               " & VBA.Format$(VBA.Trim$(S(CurrentRecord).HF), "@@@") & "              " & VBA.Format$(VBA.Trim$(VBA.Str$(S(CurrentRecord).Mark)), "@@@"), CurrentRecord - 1
End Sub

'Deleting a record
Public Sub Delete(lstBox As ListBox, S() As StudentRec, N As Integer, ByVal C As Integer)
    Dim K As Integer
    
    'Remove selected record on list box
    lstBox.RemoveItem C - 1
    'Update number of students
    N = N - 1
    
    'Reorder and display
    Reorder frmMain.lblNumRecords, lstBox, S(), N, C
End Sub

Public Sub Reorder(lblTotal As Label, lstBox As ListBox, S() As StudentRec, ByVal N As Integer, ByVal C As Integer)
    Dim K As Integer
    
    'Move student records one spot up
    For K = C To N
        S(K) = S(K + 1)
    Next K
    
    For K = C To N
        lstBox.RemoveItem K - 1
        'Display data
        lstBox.AddItem VBA.Format$(VBA.Trim$(VBA.Str$(K)), "@@") & ". " & S(K).LastName & "               " & S(K).FirstName & "               " & VBA.Format$(VBA.Trim$(S(K).HF), "@@@") & "              " & VBA.Format$(VBA.Trim$(VBA.Str$(S(K).Mark)), "@@@"), K - 1
    Next K
    
    'Display number of students
    lblTotal.Caption = VBA.Trim$(VBA.Str$(N))
End Sub
