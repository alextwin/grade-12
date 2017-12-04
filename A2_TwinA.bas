Attribute VB_Name = "Module1"
Option Explicit
Global Const NSTORE = 4 'Max number of stores

Public Function GetFile(Dialog As CommonDialog) As String
    'Initialize file name to null string
    Dialog.FileName = ""
    'Initial directory to location where project is saved
    Dialog.InitDir = App.Path
    'Filters out text files
    Dialog.Filter = "Text Files|*.txt"
    'Open common dialog
    Dialog.ShowOpen
    'Reutrn file name
    GetFile = Dialog.FileName
End Function

Public Function ReadData(TN() As String, TP() As Single, S() As Integer, ByVal FName As String) As Integer
    Dim K As Integer
    Dim X As Integer
    
    'Initialize K to 0
    K = 0
    
    'Open file
    Open FName For Input As #1
    Do While Not EOF(1)
        K = K + 1
        'Grab data from sequential file
        Input #1, TN(K), TP(K)
        For X = 1 To NSTORE
            Input #1, S(K, X)
        Next X
    Loop
    Close #1
    
    'Returns the number of toys
    ReadData = K
End Function

'All calculations done here
Public Sub Calculate(TP() As Single, S() As Integer, TotS() As Integer, ToyS() As Single, Tot As Single, ByVal NT As Integer, Ind() As String)
    Dim Num As Integer 'Number of indicators
    Dim K As Integer
    Dim X As Integer
    
    For K = 1 To NT
        For X = 1 To NSTORE
            'Calculate total sold
            TotS(K) = TotS(K) + S(K, X)
        Next X
        
        'If total sold is divisble by 5 then,
        If TotS(K) Mod 5 = 0 Then
            'Number of indicators is total sold divided by 5
            Num = TotS(K) / 5
        'If total sold is equal to 0
        ElseIf TotS(K) = 0 Then
            Num = 0
        'If total sold is not divisble by 5 and not zero
        Else
            'Number of indicators is the integer of total sold divided by 5 plus 1
            Num = (TotS(K) \ 5) + 1
        End If
        
        'Add indicators to array
        For X = 1 To Num
            Ind(K) = Ind(K) & "0"
        Next X
        
        ToyS(K) = TotS(K) * TP(K)
        Tot = Tot + ToyS(K)
    Next K
End Sub

'Output to picData
Public Sub Display(lTotal As Label, PicBox As PictureBox, TN() As String, TP() As Single, S() As Integer, TotS() As Integer, ToyS() As Single, ByVal Tot As Single, ByVal NT As Integer)
    Dim K As Integer
    Dim X As Integer
    
    'Clear picture box
    PicBox.Cls
    
    'Display headings
    PicBox.Print "Toy"; Tab(37);
    PicBox.Print Format$("EAST", "@@@@@"); Tab(44);
    PicBox.Print Format$("NORTH", "@@@@@"); Tab(51);
    PicBox.Print Format$("SOUTH", "@@@@@"); Tab(58);
    PicBox.Print Format$("WEST", "@@@@@"); Tab(70);
    PicBox.Print Format$("Total", "@@@@@"); Tab(88);
    PicBox.Print "Toy"
    PicBox.Print "Description"; Tab(24);
    PicBox.Print "Price"; Tab(37);
    
    For K = 1 To NSTORE
        PicBox.Print Format$("Store", "@@@@@"); Spc(2);
    Next K
    
    PicBox.Print Tab(70); Format$("Sold", "@@@@@");
    PicBox.Print Tab(87); "Sales"
    PicBox.Print
    
    For K = 1 To NT
        'Display toy names
        If Len(TN(K)) > 15 Then
            'Add dots if name is greater than 15
            PicBox.Print Left$(TN(K), 15); "...";
        Else
            PicBox.Print TN(K);
        End If
        'Display toy price
        PicBox.Print Tab(22); "$"; Format$(Format$(TP(K), "0.00"), "@@@@@@");
        PicBox.Print Tab(37);
        'Display number of toys sold at store
        For X = 1 To NSTORE
            PicBox.Print Format$(Str$(S(K, X)), "@@@@@"); Spc(2);
        Next X
        'Display total sold
        PicBox.Print Tab(70); Format$(Str$(TotS(K)), "@@@@@");
        'Display toy sales
        PicBox.Print Tab(82); "$"; Format$(Format$(ToyS(K), "#,##0.00"), "@@@@@@@@@")
    Next K
    
    'Output total sales to label
    lTotal = Format$(Tot, "currency")
End Sub

'Output to picChart
Public Sub DisplayChart(PicBox As PictureBox, TN() As String, Ind() As String, ByVal NT As Integer)
    Dim K As Integer
    Dim X As Integer
    
    'Initialize X to 31
    X = 31
    PicBox.Cls
    
    'Display headings
    PicBox.Print "Toy Name";
    
    For K = 0 To 250 Step 50
        PicBox.Print Tab(X); K;
        'Place next heading 10 positions down
        X = X + 10
    Next K
    
    PicBox.Print
    PicBox.Print
    
    'Display toy name and indicators
    For K = 1 To NT
        PicBox.Print TN(K); Tab(33);
        PicBox.Print Ind(K)
    Next K
End Sub
