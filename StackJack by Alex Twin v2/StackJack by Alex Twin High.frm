VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHighScore 
   BackColor       =   &H00FF80FF&
   Caption         =   "High Score"
   ClientHeight    =   3165
   ClientLeft      =   5460
   ClientTop       =   5280
   ClientWidth     =   8070
   Icon            =   "StackJack by Alex Twin High.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   8070
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   1
      Top             =   2400
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid grdHighScore 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   3413
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmHighScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload frmHighScore
End Sub

Private Sub Form_Load()
    Dim K As Integer
    Dim WName As Integer
    Dim W As Integer
    
    WName = 5430
    W = 1150
    
    grdHighScore.Cols = 3
    
    grdHighScore.ColWidth(0) = WName
    grdHighScore.ColWidth(1) = W
    grdHighScore.ColWidth(2) = W
    
    grdHighScore.Col = 0
    grdHighScore.Row = 0
    grdHighScore.Text = "Name"
    
    grdHighScore.Col = 1
    grdHighScore.CellAlignment = flexAlignRightCenter '7
    grdHighScore.Text = "Score"
    
    grdHighScore.Col = 2
    grdHighScore.CellAlignment = flexAlignRightCenter '7
    grdHighScore.Text = "Time"
    
    Open App.Path & "\" & FNAME For Random As #1 Len = RecLen
    For K = 1 To HIGHSCORE_MAX
        'Put data into record array
        Get #1, K, HighScore(K)
    Next K
    Close #1
    
    DisplayHighScore grdHighScore
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If GameStarted = True Then
        frmMain.tmrTimer.Enabled = True
    End If
End Sub
