VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   Caption         =   "StackJack by Alex Twin"
   ClientHeight    =   6165
   ClientLeft      =   3675
   ClientTop       =   4665
   ClientWidth     =   11460
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "StackJack by Alex Twin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   11460
   Begin VB.PictureBox picColumn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   5415
      Index           =   5
      Left            =   10080
      ScaleHeight     =   5385
      ScaleWidth      =   1065
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.PictureBox picColumn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   5415
      Index           =   4
      Left            =   8400
      ScaleHeight     =   5385
      ScaleWidth      =   1065
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.PictureBox picColumn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   5415
      Index           =   3
      Left            =   6720
      ScaleHeight     =   5385
      ScaleWidth      =   1065
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.PictureBox picColumn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   5415
      Index           =   2
      Left            =   5040
      ScaleHeight     =   5385
      ScaleWidth      =   1065
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.PictureBox picColumn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   5415
      Index           =   1
      Left            =   3360
      ScaleHeight     =   5385
      ScaleWidth      =   1065
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdDiscard 
      Caption         =   "Discard Card"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.PictureBox picCard 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   600
      ScaleHeight     =   1425
      ScaleWidth      =   1065
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "0:00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   21
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblPoints 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   20
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   10920
      TabIndex        =   19
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   9240
      TabIndex        =   18
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   7560
      TabIndex        =   17
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5880
      TabIndex        =   16
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   15
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Count:"
      Height          =   255
      Left            =   10080
      TabIndex        =   14
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Count:"
      Height          =   255
      Left            =   8400
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Count:"
      Height          =   255
      Left            =   6720
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Count:"
      Height          =   255
      Left            =   5040
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Count:"
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Elapsed:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Points:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Card"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New Game"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuHighScore 
         Caption         =   "High Score"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MINCARD = 0
Const MAXCARD = 51
Const MAXCOUNT = 5

Dim Cards(MINCARD To MAXCARD) As Long
Dim nWidth As Long
Dim nheight As Long
Dim PointCount(1 To MAXCOUNT) As Integer
Dim Counter As Integer

Private Sub cmdDiscard_Click()
    Counter = Counter + 1
    ShowCard Cards(), picCard, Counter
End Sub

Private Sub Form_Load()
    Randomize
    
    Initialize Cards(), PointCount(), Counter, nWidth, nheight, MINCARD, MAXCARD, MAXCOUNT
    ShowCard Cards(), picCard, Counter
End Sub

Private Sub picColumn_Click(Index As Integer)
    ShowCard Cards(), picColumn(Index), Counter
    cmdDiscard_Click
End Sub
