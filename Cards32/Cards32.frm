VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cards Demo (32-bit)"
   ClientHeight    =   5985
   ClientLeft      =   1845
   ClientTop       =   1845
   ClientWidth     =   6750
   Icon            =   "Cards32.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   6750
   Begin VB.CommandButton cmdDeck 
      Caption         =   "Show Deck"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
   End
   Begin VB.PictureBox picCards 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   5400
      ScaleHeight     =   1695
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCard 
      Caption         =   "Show Card"
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nWidth As Long
Dim nHeight As Long

Private Sub cmdCard_Click()
    Dim X As Long, Ret As Long

    X = Int(Rnd * 52)   ' Random num between 0 and 51
    Ret = cdtDraw(picCards.hDC, 0, 0, X, C_FACES, 0)
End Sub

Private Sub cmdDeck_Click()
    Dim X As Long, Ret As Long

    For X = 0 To 51 Step 4
        Ret = cdtDraw(frmMain.hDC, X * 4, 0, X, C_FACES, 0)
    Next X
    For X = 1 To 51 Step 4
        Ret = cdtDraw(frmMain.hDC, (X - 1) * 4, 100, X, C_FACES, 0)
    Next X
    For X = 2 To 51 Step 4
        Ret = cdtDraw(frmMain.hDC, (X - 2) * 4, 200, X, C_FACES, 0)
    Next X
    For X = 3 To 51 Step 4
        Ret = cdtDraw(frmMain.hDC, (X - 3) * 4, 300, X, C_FACES, 0)
    Next X
    
    For X = 54 To 68
        Ret = cdtDraw(frmMain.hDC, 275, (X - 54) * 20, X, C_BACKS, 1)
    Next X
End Sub

Private Sub Form_Load()
    Dim X As Long

    Randomize   ' Changing the seed

    X = cdtInit(nWidth, nHeight)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Ret As Long

    Ret = cdtTerm()
End Sub
