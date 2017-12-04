VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Trigonometric Functions"
   ClientHeight    =   7440
   ClientLeft      =   1890
   ClientTop       =   1935
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   9810
   Begin VB.CommandButton cmdSin 
      Caption         =   "Sine"
      Height          =   1215
      Left            =   6960
      TabIndex        =   0
      Top             =   5760
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdSin_Click()
    Const PI = 3.141592654
    Dim Sine As Single
    Dim K As Integer
    Dim Cosine As Single
    
    For K = 0 To 360
        Sine = -1000 * Sin(K * PI / 180) + 2000
        Cosine = -1000 * Cos(K * PI / 180) + 2000
        PSet (K, Sine)
        PSet (K, Cosine), RGB(255, 0, 0)
    Next K
End Sub
