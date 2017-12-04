VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Grid Control"
   ClientHeight    =   8820
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   10335
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7680
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdTrig 
      Caption         =   "Trig"
      Height          =   975
      Left            =   8520
      TabIndex        =   4
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmdDemo2 
      Caption         =   "Demo 2"
      Height          =   975
      Left            =   8280
      TabIndex        =   3
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdDemo 
      Caption         =   "Demo"
      Height          =   1095
      Left            =   8280
      TabIndex        =   2
      Top             =   4320
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdData 
      Height          =   7695
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   13573
      _Version        =   393216
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   735
      Left            =   8400
      TabIndex        =   0
      Top             =   3120
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDemo_Click()
    Dim C As Integer
    Dim W As Integer
    Dim R As Integer
    
    grdData.Clear
    
    'Set the dimensions of the grid
    
    grdData.Rows = 25
    grdData.Cols = 10
    
    'Set the column widths
    
    W = TextWidth("M")
    For C = 1 To grdData.Cols - 1
        grdData.ColWidth(C) = W * C
    Next C
    
    'Set the column headings
    
    grdData.Row = 0
    grdData.Col = 3
    grdData.Text = "Moe"
    grdData.Col = 4
    grdData.Text = "Curly"
    grdData.Col = 5
    grdData.Text = "Larry"
    
    'Fill the third column with random numbers
    
    grdData.Col = 3
    For R = 1 To grdData.Rows - 1
        grdData.Row = R
        grdData.Text = Int(Rnd * 100) + 1
    Next R
End Sub

Private Sub cmdDemo2_Click()
    Dim C As Integer
    Dim W As Integer
    Dim R As Integer
    
    grdData.Clear
    
    'Set the dimensions of the grid
    
    grdData.Rows = 5
    grdData.Cols = 6
    grdData.FixedRows = 0
    grdData.FixedCols = 0
    
    'Set the row heights and column widths
    
    For R = 0 To grdData.Rows - 1
        grdData.RowHeight(R) = 500
    Next R
    
    For C = 0 To grdData.Cols - 1
        grdData.ColWidth(C) = 1500
    Next C
    
    'Manually clear the contents of the grid
    
    For R = 0 To grdData.Rows - 1
        For C = 0 To grdData.Cols - 1
            grdData.Row = R
            grdData.Col = C
            grdData.Text = ""
        Next C
    Next R
End Sub

Private Sub cmdTest_Click()
    Dim C As Integer
    Dim W As Integer
    Dim R As Integer
    
    grdData.Clear
    
    'Set the dimensions of the grid
    
    grdData.Rows = 25
    grdData.Cols = 10
    
    'Set the column widths
    
    W = TextWidth("XXXX")
    grdData.ColWidth(0) = W
    For C = 1 To grdData.Cols - 1
        grdData.ColWidth(C) = W * 2
    Next C
    
    'Set the column headings
    
    grdData.Row = 0
    grdData.Col = 1
    grdData.Text = "One"
    grdData.Col = 2
    grdData.Text = "Two"
    grdData.Col = 3
    grdData.Text = "Three"
    
    'Fill the third column with random numbers
    
    grdData.Col = 3
    For R = 1 To grdData.Rows - 1
        grdData.Row = R
        grdData.Text = Int(Rnd * 100) + 1
    Next R
End Sub

Private Sub cmdTrig_Click()
    Const PI = 3.141592654
    Const WIDTH = 2000
    Dim K As Integer
    Dim Radians As Single
    
    grdData.Rows = 360 / 5 + 2
    grdData.Cols = 4
    
    For K = 1 To grdData.Cols - 1
        grdData.ColWidth(K) = WIDTH
    Next K
    
    grdData.Row = 0
    grdData.Col = 1
    grdData.Text = "Sine"
    grdData.Col = 2
    grdData.Text = "Cosine"
    grdData.Col = 3
    grdData.Text = "Tangent"
    
    For K = 0 To 360 Step 5
        Radians = K * PI / 180
        grdData.Row = K / 5 + 1
        
        grdData.Col = 0
        grdData.Text = K
        
        grdData.Col = 1
        grdData.Text = Sin(Radians)
        
        grdData.Col = 2
        grdData.Text = Cos(Radians)
        
        grdData.Col = 3
        grdData.Text = Tan(Radians)
    Next K
End Sub

Private Sub grdData_Click()
    MsgBox "Row=" & VBA.Str$(grdData.Row) & " Col=" & VBA.Str$(grdData.Col), vbInformation, "Grid Co-ordinates"
End Sub
