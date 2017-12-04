VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Credit Card Test Analyzer"
   ClientHeight    =   6825
   ClientLeft      =   1890
   ClientTop       =   1950
   ClientWidth     =   10230
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Reader_TwinA.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   10230
   Begin VB.PictureBox picCardNumbers 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3465
      ScaleWidth      =   9945
      TabIndex        =   12
      Top             =   1440
      Width           =   9975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Result"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   9975
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Unknown Card:"
         Height          =   375
         Left            =   7200
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblUnknown 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   8760
         TabIndex        =   13
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Invalid Card:"
         Height          =   375
         Left            =   7080
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblInvalid 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   8760
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblAMEX 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5880
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.Image Image3 
         Height          =   840
         Left            =   4680
         Picture         =   "Reader_TwinA.frx":030A
         Stretch         =   -1  'True
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblMasterCard 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.Image Image2 
         Height          =   810
         Left            =   2400
         Picture         =   "Reader_TwinA.frx":5041C
         Stretch         =   -1  'True
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblVisa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   840
         Left            =   240
         Picture         =   "Reader_TwinA.frx":53183
         Stretch         =   -1  'True
         Top             =   480
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read File && Analyze"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lblFile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   840
      Width           =   5295
   End
   Begin VB.Label lblNumCards 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Current File:"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Cards:"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Programmer: Alex Twin
'Date: September 27, 2016
'Purpose: The purpose of this program is to take and read a list of credit card numbers
'and to check to see whether or not the credit card numbers are valid or not.  If they are valid,
'it checks whether the credit card number is a Visa, MasterCard, or American Express.

Private Sub cmdExit_Click()
    'End program
    End
End Sub

Private Sub cmdRead_Click()
    Const MAX = 20
    Dim CardNum(1 To MAX) As String
    Dim NumCards As Integer
    Dim NumVisa As Integer
    Dim NumMasterCard As Integer
    Dim NumAMEX As Integer
    Dim NumInvalid As Integer
    Dim NumUnknown As Integer
    Dim FileName As String
    Dim K As Integer
    
    'Initialize
    NumVisa = 0
    NumMasterCard = 0
    NumAMEX = 0
    NumInvalid = 0
    NumUnknown = 0
    'Path to file
    FileName = App.Path & "\CC_TEST.TXT"
    'Read text file
    NumCards = ReadFile(FileName, CardNum())
    'Clears picture box
    picCardNumbers.Cls
    
    'Go through all credit card numbers
    For K = 1 To NumCards
        'Print credit card numbers to picture box
        picCardNumbers.Print CardNum(K),
        'Puts credit card numbers into 3 columns
        If K Mod 3 = 0 Then
            picCardNumbers.Print
        End If
        'Verify if credit card number is valid or not
        If InvalidCard(CardNum(), K) = True Then
            'If it is, add to invalid counter
            NumInvalid = NumInvalid + 1
        'If it passes validity test, checks for credit card type
        Else
            'If the credit card number starts with 4, it's a Visa
            If Left$(CardNum(K), 1) = "4" Then
                NumVisa = NumVisa + 1
            'If it starts with 34 or 37, it's American Express
            ElseIf Val(Left$(CardNum(K), 2)) = 34 Or Val(Left$(CardNum(K), 2)) = 37 Then
                NumAMEX = NumAMEX + 1
            'If it starts with 51, 52, 53, 54 and 55, it is a MasterCard
            ElseIf Val(Left$(CardNum(K), 2)) >= 51 And Val(Left$(CardNum(K), 2)) <= 55 Then
                NumMasterCard = NumMasterCard + 1
            'If credit card number somehow passes validity test, but does not start with any
            'of the numbers above, add to unknown counter
            Else
                NumUnknown = NumUnknown + 1
            End If
        End If
    Next K
    
    'Output
    lblNumCards.Caption = Trim$(Str$(NumCards))
    'Add "..." if the length of the filename is greater than 50
    If Len(FileName) > 50 Then
        lblFile.Caption = Left$(FileName, 47) & "..."
    Else
        lblFile.Caption = FileName
    End If
    lblVisa.Caption = Trim$(Str$(NumVisa))
    lblMasterCard.Caption = Trim$(Str$(NumMasterCard))
    lblAMEX.Caption = Trim$(Str$(NumAMEX))
    lblInvalid.Caption = Trim$(Str$(NumInvalid))
    lblUnknown.Caption = Trim$(Str$(NumUnknown))
End Sub

