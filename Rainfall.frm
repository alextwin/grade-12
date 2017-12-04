VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "CP24"
   ClientHeight    =   5955
   ClientLeft      =   1890
   ClientTop       =   2205
   ClientWidth     =   13890
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
   ScaleHeight     =   5955
   ScaleWidth      =   13890
   Begin VB.PictureBox picData 
      Height          =   5655
      Left            =   120
      ScaleHeight     =   5595
      ScaleWidth      =   13395
      TabIndex        =   0
      Top             =   120
      Width           =   13455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuRead 
         Caption         =   "Read File"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Average As Single
Dim MAverage(1 To MAXCITY) As Single

Private Sub Form_Load()
    Initialize Average, MAverage()
End Sub

Private Sub mnuRead_Click()
    Initialize Average, MAverage()
    NumCity = ReadFile(App.Path & "\RAINFALL.TXT")
    MonthlyAverage MAverage()
    Average = OverallAverage()
    LeastRain
    Display picData, MAverage(), Average
End Sub
