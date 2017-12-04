VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3360
   ClientLeft      =   4635
   ClientTop       =   6075
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrTimer 
      Interval        =   3000
      Left            =   4320
      Top             =   1800
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "By Alex Twin"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   7560
      TabIndex        =   1
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "StackJack"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   3720
      TabIndex        =   0
      Top             =   -240
      Width           =   7215
   End
   Begin VB.Image Image1 
      Height          =   3060
      Left            =   120
      Picture         =   "StackJack by Alex Twin Splash.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3315
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    CentreForm frmSplash
End Sub

Private Sub tmrTimer_Timer()
    'After 3 seconds unload splash form
    Unload frmSplash
    'Load and show main form
    Load frmMain
    CentreForm frmMain
    frmMain.Show
End Sub
