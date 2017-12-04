VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "QuickSort vs BubbleSort"
   ClientHeight    =   9765
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   9090
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9765
   ScaleWidth      =   9090
   Begin VB.CommandButton Command3 
      Caption         =   "QuickSort"
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   9120
      Width           =   2775
   End
   Begin VB.CommandButton cmdBubbleSort 
      Caption         =   "BubbleSort"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   9120
      Width           =   2775
   End
   Begin VB.CommandButton cmdOriginal 
      Caption         =   "Original Numbers"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   9120
      Width           =   2775
   End
   Begin VB.PictureBox picQuickSort 
      Height          =   8895
      Left            =   6120
      ScaleHeight     =   8835
      ScaleWidth      =   2715
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.PictureBox picBubbleSort 
      Height          =   8895
      Left            =   3120
      ScaleHeight     =   8835
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.PictureBox picOriginal 
      Height          =   8895
      Left            =   120
      ScaleHeight     =   8835
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const LOW = 1
Const MAX = 2000
Dim Num(1 To MAX) As Integer

Private Sub cmdOriginal_Click()
    Dim K As Integer
    
    picOriginal.Cls
    
    For K = 1 To MAX
        Num(K) = Int(Rnd * (MAX - LOW + 1) + 1)
        picOriginal.Print Num(K)
    Next K
End Sub

Private Sub Form_Load()
    Randomize
End Sub
