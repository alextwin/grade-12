Attribute VB_Name = "Module1"
Option Explicit

Public Sub Initialize(Cards() As Long, Count() As Integer, Counter As Integer, ByVal nWidth As Long, ByVal nheight As Long, ByVal MINCARD As Integer, ByVal MAXCARD As Integer, ByVal MCOUNT As Integer)
    Dim K As Long
    Dim X As Long
    
    Counter = 0
    
    X = cdtInit(nWidth, nheight)
    
    For K = MINCARD To MAXCARD
        Cards(K) = K
    Next K
    
    For K = 1 To MCOUNT
        Count(K) = 0
    Next K
    
    Shuffle Cards(), MINCARD, MAXCARD
End Sub

Public Sub Shuffle(Cards() As Long, ByVal MINCARD As Integer, ByVal MAXCARD As Integer)
    Dim K As Integer
    Dim Temp As Integer
    Dim X As Integer
    Dim Y As Integer
    
    For K = 1 To 10000
        X = Int(Rnd() * (MAXCARD - MINCARD + 1) + MINCARD)
        Y = Int(Rnd() * (MAXCARD - MINCARD + 1) + MINCARD)
        Temp = Cards(X)
        Cards(X) = Cards(Y)
        Cards(Y) = Temp
    Next K
End Sub

Public Sub ShowCard(Cards() As Long, PicBox As PictureBox, Counter)
    Dim Ret As Long
    
    Ret = cdtDraw(PicBox.hDC, 0, 0, Cards(Counter), C_FACES, 0)
    PicBox.Refresh
End Sub
