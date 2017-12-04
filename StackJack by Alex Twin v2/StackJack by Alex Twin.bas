Attribute VB_Name = "Module1"
Option Explicit

Type HighScoreRec
    Name As String * 28
    Score As Integer
    Time As String * 5
End Type

Global Const HIGHSCORE_MAX = 5
Global Const FNAME = "highscore.rec"
Global HighScore(1 To HIGHSCORE_MAX) As HighScoreRec
Global RecLen As Integer
Global GameStarted As Boolean
Dim nWidth As Long
Dim nHeight As Long
Dim ExtraPoint As Boolean 'Determines if user gets an extra 10 points

Public Sub Initialize(Points As Integer, Seconds As Integer, Minutes As Integer, Consecutive As Integer, BustCounter As Integer, Cards() As Long, Counter As Integer, VDisp() As Long, Total() As Integer, NumAce() As Integer, ByVal MINCARD As Integer, ByVal MAXCARD As Integer, ByVal MAXCOLUMN As Integer)
    Dim K As Long
    Dim X As Long
    
    'Initialize variables to 0
    Counter = 0
    Points = 0
    BustCounter = 0
    Consecutive = 0
    Seconds = 0
    Minutes = 0
    
    'Initialize cards
    X = cdtInit(nWidth, nHeight)
    
    For K = MINCARD To MAXCARD
        Cards(K) = K
    Next K
    
    For K = 1 To MAXCOLUMN
        VDisp(K) = 0
        Total(K) = 0
        NumAce(K) = 0
    Next K
    
    'Put cards in random order
    Shuffle Cards(), MINCARD, MAXCARD
End Sub

Public Sub Shuffle(Cards() As Long, ByVal MINCARD As Integer, ByVal MAXCARD As Integer)
    Dim K As Integer
    Dim Temp As Integer
    Dim X As Integer
    Dim Y As Integer
    
    For K = 1 To 10000
        'Random number from 0 to 51
        X = Int(Rnd() * (MAXCARD - MINCARD + 1) + MINCARD)
        Y = Int(Rnd() * (MAXCARD - MINCARD + 1) + MINCARD)
        'Swap the card numbers
        Temp = Cards(X)
        Cards(X) = Cards(Y)
        Cards(Y) = Temp
    Next K
End Sub

'Show card
Public Sub ShowCard(ByVal Cards As Long, PicBox As PictureBox, ByVal VDisp As Long, ByVal SIDE As Long)
    Dim Ret As Long
    
    Ret = cdtDraw(PicBox.hDC, 0, VDisp, Cards, SIDE, 0)
    PicBox.Refresh
End Sub

'Update count
Public Sub UpCount(lblCount As Label, ByVal Cards As Long, Total As Integer, NumAce As Integer)
    Dim K As Integer
    Dim Value As Integer
    
    'Take card number, integer division by 4, plus 1 will give you the value of a card
    '1 is ace
    '2 is 2 etc...
    Value = Cards \ 4 + 1
    'Assume user won't get extra points
    ExtraPoint = False
    
    'If it is a jack, queen, king, or ace
    If Value > 10 Or Value = 1 Then
        'User gets extra points
        ExtraPoint = True
        
        'If it is a jack, queen, or king
        If Value > 10 Then
            'Value is 10
            Value = 10
        'If it is an ace
        Else
            'Increment number of aces by 1
            NumAce = NumAce + 1
            'Give it a value of 11
            Value = 11
        End If
    End If
    
    'Update the total count
    Total = Total + Value
    
    'If the total is over 21, and there is at least one ace
    If Total > 21 And NumAce > 0 Then
        'Decrement total by 10 (ace now has a value of 1)
        Total = Total - 10
        'Decrement number of aces with a value of 11 by 1
        NumAce = NumAce - 1
    End If
    
    'Display new total count for the column
    lblCount.Caption = VBA.Trim$(VBA.Str$(Total))
End Sub

'Update points
Public Sub PointCounter(Points As Integer, Consecutive As Integer, BustCounter As Integer, PicBox As PictureBox, Total As Integer, lblCount As Label, lblPoints As Label, VDisp As Long, NumAce As Integer)
    'If total count is less than 21
    If Total < 21 Then
        'Increment points by 40
        Points = Points + 40
        Consecutive = 0
        
        'Give user extra 10 points for jack, queen, king, and ace
        If ExtraPoint = True Then
            Points = Points + 10
        End If
    'If the total is exactly 21
    ElseIf Total = 21 Then
        Consecutive = Consecutive + 1
        'Increment points by 500
        Points = Points + 500 * Consecutive
        'Clear column
        PicBox.Cls
        'Set count label to 0
        lblCount.Caption = "0"
        'Set variables back to 0
        VDisp = 0
        Total = 0
        NumAce = 0
    'If the column goes bust
    Else
        'Increment bust counter by 1
        BustCounter = BustCounter + 1
        Consecutive = 0
        'Disable column
        PicBox.Enabled = False
        'Change font colour of the count to red
        lblCount.ForeColor = RGB(255, 0, 0)
        
        'If all columns are bust or there are 700 or less points
        If BustCounter = 5 Or Points <= 700 Then
            'Points will be 0
            Points = 0
        ElseIf Points > 700 Then
            'Decrement points by 700
            Points = Points - 700
        End If
    End If
    
    'Display points to label
    lblPoints.Caption = VBA.Trim$(VBA.Str$(Points))
End Sub

Public Sub NewGame(lblTime As Label, cmdDiscard As CommandButton, picCard As PictureBox, lblPoints As Label, picColumn As Variant, lblCount As Variant, ByVal MAXCOLUMN As Integer)
    Dim K As Integer
    
    'Clear picCard
    picCard.Cls
    'Set points label to 0
    lblPoints.Caption = "0"
    lblTime.Caption = "00:00"
    'Enable discard button
    cmdDiscard.Enabled = True
    
    For K = 1 To MAXCOLUMN
        'Clear column picture boxes
        picColumn(K).Cls
        'Enable column picture boxes
        picColumn(K).Enabled = True
        'Set count labels to 0
        lblCount(K).Caption = "0"
        'Change the colour back to black
        lblCount(K).ForeColor = RGB(0, 0, 0)
    Next K
End Sub

'When the user has run out of cards
Public Sub EndGame(Points As Integer, PicBox As Variant, ByVal MAXCOLUMN As Integer)
    Dim DMsg As String
    Dim DType As Integer
    Dim DTitle As String
    Dim K As Integer
    
    DMsg = "Game Over!  You have used up all the playing cards!" & vbCrLf & "Your final score was " & VBA.Trim$(VBA.Str$(Points)) & "."
    DType = vbInformation + vbOKOnly
    DTitle = "Game Over"
    
    MsgBox DMsg, DType, DTitle
    
    For K = 1 To MAXCOLUMN
        'Disable the column picture boxes
        PicBox(K).Enabled = False
    Next K
End Sub

'When all 5 columns are bust
Public Sub AllBust()
    Dim DMsg As String
    Dim DType As Integer
    Dim DTitle As String
    
    DMsg = "Game Over!  All columns have gone bust!" & vbCrLf & "Your final score was 0."
    DType = vbInformation + vbOKOnly
    DTitle = "All Bust"
    
    MsgBox DMsg, DType, DTitle
End Sub

Public Sub CheckFile()
    Dim K As Integer
    
    If Dir$(FNAME) = "" Then
        Open App.Path & "\" & FNAME For Random As #1 Len = RecLen
        For K = 1 To HIGHSCORE_MAX
            HighScore(K).Score = 0
            HighScore(K).Name = "Anonymous"
            HighScore(K).Time = "00:00"
            Put #1, K, HighScore(K)
        Next K
        Close #1
    End If
End Sub

Public Sub DisplayHighScore(Grid As MSFlexGrid)
    Dim K As Integer
    Dim R As Integer
    
    R = 1
    
    Grid.Rows = HIGHSCORE_MAX + 1
    Grid.Height = Grid.RowHeight(0) * (Grid.Rows) + 95
    
    For K = 1 To HIGHSCORE_MAX
        Grid.Row = K
        Grid.Col = 0
        Grid.Text = HighScore(K).Name
        Grid.Col = 1
        Grid.CellAlignment = flexAlignRightCenter '7
        Grid.Text = VBA.Str$(HighScore(K).Score)
        Grid.Col = 2
        Grid.Text = HighScore(K).Time
    Next K
        
End Sub

Public Sub CheckHS(ByVal Points As Integer, ByVal Minutes As Integer, ByVal Seconds As Integer)
    Dim K As Integer
    Dim X As Integer
    Dim TempRec As HighScoreRec
    Dim MoveRec As HighScoreRec
    Dim HSMsg As String
    Dim HSType As Integer
    Dim HSTitle As String
    Dim TimeFormat As String
    
    TimeFormat = VBA.Format$(VBA.Str$(Minutes), "00") & ":" & VBA.Format$(VBA.Str$(Seconds), "00")
    
    Open App.Path & "\" & FNAME For Random As #1 Len = RecLen
    For K = 1 To HIGHSCORE_MAX
        'Put data into record array
        Get #1, K, HighScore(K)
    Next K
    Close #1
    
    For K = 1 To HIGHSCORE_MAX
        If Points > HighScore(K).Score Or (Points = HighScore(K).Score And TimeFormat < HighScore(K).Time) Then
            X = K
            TempRec = HighScore(X)
            Do While X <= 4
                X = X + 1
                MoveRec = HighScore(X)
                HighScore(X) = TempRec
                TempRec = MoveRec
            Loop
            
            HSMsg = "Congratulations, a new high score has been achieved!"
            HSType = vbOKOnly + vbInformation
            HSTitle = "High Score Achieved"
            MsgBox HSMsg, HSType, HSTitle
            
            HighScore(K).Name = InputBox$("Please eneter your name.", HSTitle)
            If VBA.Trim$(HighScore(K).Name) = "" Then
                HighScore(K).Name = "Anonymous"
            ElseIf Len(VBA.Trim$(HighScore(K).Name)) > 25 Then
                HighScore(K).Name = VBA.Left$(HighScore(K).Name, 25) & "..."
            End If
            HighScore(K).Score = Points
            HighScore(K).Time = TimeFormat
            
            Open App.Path & "\" & FNAME For Random As #1 Len = RecLen
                For X = 1 To HIGHSCORE_MAX
                    Put #1, X, HighScore(X)
                Next X
            Close #1
            Exit For
        End If
    Next K
End Sub

Public Sub CentreForm(CentredForm As Form)
    CentredForm.Move (Screen.Width - CentredForm.Width) / 2, (Screen.Height - CentredForm.Height) / 2
End Sub

