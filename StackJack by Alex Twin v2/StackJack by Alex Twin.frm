VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   Caption         =   "StackJack by Alex Twin"
   ClientHeight    =   5895
   ClientLeft      =   3810
   ClientTop       =   4200
   ClientWidth     =   11385
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "StackJack by Alex Twin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   11385
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   4800
   End
   Begin VB.PictureBox picColumn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   5175
      Index           =   5
      Left            =   10080
      ScaleHeight     =   5145
      ScaleWidth      =   1065
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.PictureBox picColumn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   5175
      Index           =   4
      Left            =   8400
      ScaleHeight     =   5145
      ScaleWidth      =   1065
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.PictureBox picColumn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   5175
      Index           =   3
      Left            =   6720
      ScaleHeight     =   5145
      ScaleWidth      =   1065
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.PictureBox picColumn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   5175
      Index           =   2
      Left            =   5040
      ScaleHeight     =   5145
      ScaleWidth      =   1065
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.PictureBox picColumn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   5175
      Index           =   1
      Left            =   3360
      ScaleHeight     =   5145
      ScaleWidth      =   1065
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdDiscard 
      Caption         =   "Discard Card"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.PictureBox picCard 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   600
      ScaleHeight     =   1425
      ScaleWidth      =   1065
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   21
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label lblPoints 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   20
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   10920
      TabIndex        =   19
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   9240
      TabIndex        =   18
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   7560
      TabIndex        =   17
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   5880
      TabIndex        =   16
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   15
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Count:"
      Height          =   255
      Left            =   10080
      TabIndex        =   14
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Count:"
      Height          =   255
      Left            =   8400
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Count:"
      Height          =   255
      Left            =   6720
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Count:"
      Height          =   255
      Left            =   5040
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Count:"
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Elapsed:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Points:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Card"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New Game"
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuHighScore 
         Caption         =   "High Score..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuRules 
         Caption         =   "Rules..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name: Alex Twin
'Date: June 5, 2017
'Purpose: Create a electronic and variation of StackJack v2
Option Explicit
Const MINCARD = 0                         'Lowest card number
Const MAXCARD = 51                        'Highest card number
Const MAXCOLUMN = 5                       'Max number of columns
Const ZERODISP = 0                        'Constant for position of card

Dim Cards(MINCARD To MAXCARD) As Long     'Array for cards
Dim Counter As Integer                    'Keep track of cards left
Dim VDisplacement(1 To MAXCOLUMN) As Long 'Vertical position of card
Dim Total(1 To MAXCOLUMN) As Integer      'Array for the count of each column
Dim NumAce(1 To MAXCOLUMN) As Integer     'Number of aces (with value of 11) each column has
Dim BustCounter As Integer                'Number of busts
Dim Points As Integer
Dim Consecutive As Integer
Dim Minutes As Integer
Dim Seconds As Integer

Private Sub cmdDiscard_Click()
    If tmrTimer.Enabled = False Then
        tmrTimer.Enabled = True
    End If
    
    'If points is greater than 150, subtract 150
    If Points > 150 Then
        Points = Points - 150
    'Otherwise points = 0
    Else
        Points = 0
    End If
    
    'Display points
    lblPoints.Caption = VBA.Trim$(VBA.Str$(Points))
    Consecutive = 0
    
    'Increment counter
    Counter = Counter + 1
    'If there are still cards left
    If Counter < 52 Then
        'Show card next card
        ShowCard Cards(Counter), picCard, ZERODISP, C_FACES
    Else
        GameStarted = False
        tmrTimer.Enabled = False
        'Enabled new game menu button
        mnuNew.Enabled = True
        'Disable discard button
        cmdDiscard.Enabled = False
        'Displays the back of a card with island image
        ShowCard Island, picCard, ZERODISP, C_BACKS
        'End game
        EndGame Points, picColumn(), MAXCOLUMN
        CheckHS Points, Minutes, Seconds
    End If
End Sub

Private Sub Form_Load()
    Randomize
    RecLen = Len(HighScore(1))
    GameStarted = False
    
    CheckFile
    
    'Initialize variables
    Initialize Points, Seconds, Minutes, Consecutive, BustCounter, Cards(), Counter, VDisplacement(), Total(), NumAce(), MINCARD, MAXCARD, MAXCOLUMN
    'Show the first card
    ShowCard Cards(Counter), picCard, ZERODISP, C_FACES
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Ret As Long
    
    'Release memory back to windows
    Ret = cdtTerm()
End Sub

Private Sub mnuAbout_Click()
    tmrTimer.Enabled = False
    'Load about form
    Load frmAbout
    CentreForm frmAbout
    'Show about form
    frmAbout.Show vbModal
End Sub

Private Sub mnuExit_Click()
    Dim DMsg As String
    Dim DType As Integer
    Dim DTitle As String
    Dim Response As Integer
    Dim Ret As Long
    
    DMsg = "Are you sure you would like to exit?"
    DType = vbYesNo + vbQuestion
    DTitle = "Exit StackJack"
    Response = MsgBox(DMsg, DType, DTitle)
    
    'Check to see if user wants to exit
    If Response = vbYes Then
        End
    End If
End Sub

Private Sub mnuHighScore_Click()
    tmrTimer.Enabled = False
    Load frmHighScore
    CentreForm frmHighScore
    frmHighScore.Show vbModal
End Sub

Private Sub mnuNew_Click()
    'Disable new game menu
    mnuNew.Enabled = False
    'Start new game
    NewGame lblTime, cmdDiscard, picCard, lblPoints, picColumn(), lblCount(), MAXCOLUMN
    'Initialize variables
    Initialize Points, Seconds, Minutes, Consecutive, BustCounter, Cards(), Counter, VDisplacement(), Total(), NumAce(), MINCARD, MAXCARD, MAXCOLUMN
    'Show card
    ShowCard Cards(Counter), picCard, ZERODISP, C_FACES
End Sub

Private Sub mnuRules_Click()
    tmrTimer.Enabled = False
    'Load rules form
    Load frmRules
    CentreForm frmRules
    'Show rules form
    frmRules.Show vbModal
End Sub

'User clicks on one of the columns
Private Sub picColumn_Click(Index As Integer)
    If tmrTimer.Enabled = False Then
        tmrTimer.Enabled = True
    End If
    
    'Show next card
    ShowCard Cards(Counter), picColumn(Index), VDisplacement(Index), C_FACES
    
    'Increment the vertical displacement by 25
    VDisplacement(Index) = VDisplacement(Index) + 25
    
    'Update the count
    UpCount lblCount(Index), Cards(Counter), Total(Index), NumAce(Index)
    'Update the points
    PointCounter Points, Consecutive, BustCounter, picColumn(Index), Total(Index), lblCount(Index), lblPoints, VDisplacement(Index), NumAce(Index)
    
    'Increment counter
    Counter = Counter + 1
    'If all columns are bust or there are no more cards
    If BustCounter = 5 Or Counter > 51 Then
        GameStarted = False
        tmrTimer.Enabled = False
        'Enable new game menu button
        mnuNew.Enabled = True
        'Show the back of a card with the island image
        ShowCard Island, picCard, ZERODISP, C_BACKS
        'Disable discard button
        cmdDiscard.Enabled = False
        
        'If all columns are bust
        If BustCounter = 5 Then
            AllBust
        'If there are no more cards
        Else
            EndGame Points, picColumn(), MAXCOLUMN
            CheckHS Points, Minutes, Seconds
        End If
    'If there are more cards
    Else
        'Show the next card
        ShowCard Cards(Counter), picCard, ZERODISP, C_FACES
    End If
End Sub

Private Sub tmrTimer_Timer()
    Seconds = Seconds + 1
    
    If Seconds = 60 Then
        Minutes = Minutes + 1
        Seconds = 0
    End If
    
    lblTime.Caption = VBA.Format$(VBA.Str$(Minutes), "00") & ":" & VBA.Format$(VBA.Str$(Seconds), "00")
End Sub
