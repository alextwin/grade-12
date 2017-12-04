VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Day of the Week"
   ClientHeight    =   3000
   ClientLeft      =   3450
   ClientTop       =   2760
   ClientWidth     =   6195
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   6195
   Begin VB.ComboBox cboYear 
      Height          =   345
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox cboDay 
      Height          =   345
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox cboMonth 
      Height          =   345
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton cmdGetDay 
      Caption         =   "Get the Day of the Week"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   5895
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   5895
   End
   Begin VB.Label Label3 
      Caption         =   "Year:"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Day:"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Month:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MAX_YEAR = 2100
Const LOW_YEAR = 1900
Const MAX_MONTH = 12
Const MAX_DAY = 31
Dim CurrentYear As Integer
Dim CurrentMonth As Integer
Dim CurrentDay As Integer

Private Sub cboDay_Click()
    CurrentDay = Val(cboDay.List(cboDay.ListIndex))
End Sub

Private Sub cboMonth_Click()
    CurrentMonth = cboMonth.ListIndex + 1
    CheckDays cboDay, CurrentYear, CurrentMonth, CurrentDay, MAX_DAY
End Sub

Private Sub cboYear_Click()
    CurrentYear = Val(cboYear.List(cboYear.ListIndex))
    CheckDays cboDay, CurrentYear, CurrentMonth, CurrentDay, MAX_DAY
End Sub

Private Sub cmdGetDay_Click()
    lblDay.Caption = GetDay(CurrentYear, CurrentMonth, CurrentDay)
End Sub

Private Sub Form_Load()
    Dim K As Integer
    
    For K = LOW_YEAR To MAX_YEAR
        cboYear.AddItem Trim$(Str$(K)), K - LOW_YEAR
    Next K
    For K = 0 To MAX_MONTH - 1
        cboMonth.AddItem MonthName(K + 1), K
    Next K
    For K = 0 To MAX_DAY - 1
        cboDay.AddItem Trim$(Str$(K + 1)), K
    Next K
    cboYear.ListIndex = 0
    cboMonth.ListIndex = 0
    cboDay.ListIndex = 0

    CurrentYear = Val(cboYear.List(cboYear.ListIndex))
    CurrentMonth = cboMonth.ListIndex + 1
    CurrentDay = Val(cboDay.List(cboDay.ListIndex))
    cmdGetDay_Click
End Sub
