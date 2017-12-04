VERSION 5.00
Begin VB.Form frmRules 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rules"
   ClientHeight    =   9360
   ClientLeft      =   8580
   ClientTop       =   3420
   ClientWidth     =   8625
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "StackJack by Alex Twin Rules.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   8625
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   6840
      TabIndex        =   19
      Top             =   8760
      Width           =   1575
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "-150"
      Height          =   375
      Left            =   7680
      TabIndex        =   21
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Discard Card"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   7080
      Width           =   2895
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "-700"
      Height          =   375
      Left            =   7680
      TabIndex        =   18
      Top             =   8280
      Width           =   735
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Over 21 (Bust)"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   8280
      Width           =   2895
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "500 x Number of Consecutive Clears"
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   7920
      Width           =   4215
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Value"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Of 21"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      Height          =   375
      Left            =   7680
      TabIndex        =   13
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Jack, Queen, King, Ace"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   6720
      Width           =   2895
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "40"
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "2-10"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Points"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   9
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Card"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Scoring:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5640
      Width           =   8295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "-There is no negative scoring (the player will not have a score less than 0)."
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   8295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "-If a column adds to a value of exactly 21, the column will be cleared and the count will be reset to 0."
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   8295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "-If all five columns go over, then the game is over and the player's score is zero."
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   8295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "-If a column adds to a value over 21, it goes ""bust"" and it can no longer be used for the remainder of the game."
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   8295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Rules:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   8295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "How to Play:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"StackJack by Alex Twin Rules.frx":030A
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8295
   End
End
Attribute VB_Name = "frmRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    'Unload Rules form
    Unload frmRules
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If GameStarted = True Then
        frmMain.tmrTimer.Enabled = True
    End If
End Sub
