VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Calendar"
   ClientHeight    =   4485
   ClientLeft      =   8190
   ClientTop       =   5490
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LAB1_TwinA.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   8775
   Begin VB.ComboBox cboWeekday 
      Height          =   390
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.ComboBox cboMonth 
      Height          =   390
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid grdCalendar 
      Height          =   4215
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7435
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   8421504
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name: Alex Twin
'Date: February 24, 2017
'Purpose: Create a calendar that allows the user to pick a month
'and a starting weekday
Option Explicit
Dim Weekday As Integer
Dim NumDays As Integer

Private Sub cboMonth_Click()
    Dim Month As Integer
    
    'Number value of month (1 is January, 2 is February etc...)
    Month = cboMonth.ListIndex + 1
    'Find number of days in a month
    NumDays = CheckMonth(Month)
    'Change calendar
    ChangeCalendar NumDays, Weekday, grdCalendar
End Sub

Private Sub cboWeekday_Click()
    'Weekday number (0 is Sunday, 1 is Monday etc...)
    Weekday = cboWeekday.ListIndex
    'Change calendar
    ChangeCalendar NumDays, Weekday, grdCalendar
End Sub

Private Sub Form_Load()
    'Initialize grid, combo boxes and Weekday variable
    Intialize grdCalendar, cboWeekday, cboMonth, Weekday
End Sub
