VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Toys Cubed"
   ClientHeight    =   5130
   ClientLeft      =   2025
   ClientTop       =   2400
   ClientWidth     =   9825
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "A2_TwinA.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   9825
   Begin VB.PictureBox picData 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4305
      ScaleWidth      =   9585
      TabIndex        =   7
      Top             =   120
      Width           =   9615
      Begin VB.Line linData 
         Visible         =   0   'False
         X1              =   0
         X2              =   9600
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      Height          =   495
      Left            =   7680
      TabIndex        =   6
      Top             =   4560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdShowChart 
      Caption         =   "Show Chart"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   1815
   End
   Begin VB.PictureBox picChart 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4305
      ScaleWidth      =   9585
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   9615
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9600
         Y1              =   360
         Y2              =   360
      End
   End
   Begin MSComDlg.CommonDialog cdlDialog 
      Left            =   960
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   5
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label lblTotalSales 
      Caption         =   "Total Sales:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   4680
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer: Alex Twin
'Date: October 25, 2016
'Purpose: The purpose of this assignment is to use 2D arrays to read and display data
Option Explicit
Const MAX = 1000                                 'Max number of toys
Dim ToyName(1 To MAX) As String
Dim ToyPrice(1 To MAX) As Single
Dim StoreSold(1 To MAX, 1 To NSTORE) As Integer
Dim TotalSold(1 To MAX) As Integer               'Total number of toys sold
Dim ToySales(1 To MAX) As Single
Dim Indicator(1 To MAX) As String                'Indicator for bar graph
Dim TotalSales As Single
Dim NumToys As Integer

Private Sub cmdExit_Click()
    Dim DMsg As String
    Dim DType As Integer
    Dim DTitle As String
    Dim Response As Integer
    
    'Make sure the user actually wants to leave
    DMsg = "Are you sure you want to exit?"
    DType = vbYesNo + vbExclamation + vbDefaultButton2
    DTitle = "Exit"
    
    Response = MsgBox(DMsg, DType, DTitle)
    
    If Response = vbYes Then
        End
    End If
End Sub

Private Sub cmdOpen_Click()
    Dim FileName As String
    Dim K As Integer
    Dim X As Integer
    
    'Get the name of the file
    FileName = GetFile(cdlDialog)
    
    If FileName <> "" Then
        
        'Initialize every time a new file is opened
        For K = 1 To MAX
            ToyName(K) = ""
            ToyPrice(K) = 0
            TotalSold(K) = 0
            ToySales(K) = 0
            Indicator(K) = ""
            For X = 1 To NSTORE
                StoreSold(K, X) = 0
            Next X
        Next K
        TotalSales = 0
        NumToys = 0
        
        'Enable the show chart button
        cmdShowChart.Enabled = True
        'Read data, and find number of toys
        NumToys = ReadData(ToyName(), ToyPrice(), StoreSold(), FileName)
        'Calculations
        Calculate ToyPrice(), StoreSold(), TotalSold(), ToySales(), TotalSales, NumToys, Indicator()
        'Make the line visible
        linData.Visible = True
        'Display to picData
        Display lblTotal, picData, ToyName(), ToyPrice(), StoreSold(), TotalSold(), ToySales(), TotalSales, NumToys
        'Display to picChart
        DisplayChart picChart, ToyName(), Indicator(), NumToys
    End If
End Sub

'Returns back to show picData
Private Sub cmdReturn_Click()
    frmMain.Caption = "Toys Cubed"
    picData.Visible = True
    cmdOpen.Visible = True
    cmdShowChart.Visible = True
    cmdExit.Visible = True
    lblTotalSales.Visible = True
    lblTotal.Visible = True
    cmdReturn.Visible = False
    picChart.Visible = False
End Sub

'Displays chart
Private Sub cmdShowChart_Click()
    frmMain.Caption = "Toys Cubed (Sales Chart)"
    picData.Visible = False
    cmdOpen.Visible = False
    cmdShowChart.Visible = False
    cmdExit.Visible = False
    lblTotalSales.Visible = False
    lblTotal.Visible = False
    cmdReturn.Visible = True
    picChart.Visible = True
End Sub
