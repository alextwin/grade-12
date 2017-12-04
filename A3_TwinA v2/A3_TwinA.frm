VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Random Access Files"
   ClientHeight    =   6960
   ClientLeft      =   7290
   ClientTop       =   4350
   ClientWidth     =   8430
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "A3_TwinA.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   8430
   Begin VB.PictureBox picHeader 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   8115
      TabIndex        =   3
      Top             =   120
      Width           =   8175
   End
   Begin VB.ListBox lstData 
      Height          =   5715
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   8175
   End
   Begin MSComDlg.CommonDialog cdlDialog 
      Left            =   120
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Number of Students:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label lblNumRecords 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   0
      Top             =   6480
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save As..."
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add Record..."
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete Record..."
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuModify 
         Caption         =   "&Modify Record..."
         Enabled         =   0   'False
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSort 
         Caption         =   "Sort"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGetInfo 
         Caption         =   "Get Info"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
         Visible         =   0   'False
         Begin VB.Menu mnuSequential 
            Caption         =   "Sequential"
         End
         Begin VB.Menu mnuBinary 
            Caption         =   "Binary"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer: Alex Twin
'Date: January 27, 2017
'Purpose: Update the program to add extra features such as add record, modify record and delete record
Option Explicit
Dim FileName As String
Dim RecLength As Integer             'Length of a record

Private Sub Form_Load()
    'Finding the length of a record
    RecLength = Len(Student(1))
    'Initialize
    Initialize Student(), NumStudents, FileName, MAX
    'Displays headings
    picHeader.Print Tab(5); "Last Name"; Tab(40); "First Name"; Tab(71); "HF"; Tab(87); "Mark"
End Sub

Private Sub lstData_Click()
    'When an item is selected, enable modify and delete
    mnuModify.Enabled = True
    mnuDelete.Enabled = True
    'Keep track of the current record
    CurrentRecord = lstData.ListIndex + 1
End Sub

Private Sub lstData_DblClick()
    'Load Modify
    LoadModify frmDetailedInfo, lstData, Student()
End Sub

Private Sub mnuAbout_Click()
    'Loading about form
    Load frmAbout
    'Displaying about form
    frmAbout.Show vbModal
End Sub

Private Sub mnuAdd_Click()
    'If there are too many records, don't display the add form
    If NumStudents = MAX Then
        MsgBox "You have reached the maximum amount of records", vbOKOnly + vbCritical, "Maximum Records Reached"
    Else
        'Load the form to add records
        Load frmDetailedInfo
        'Change form caption
        frmDetailedInfo.Caption = "Add Record"
        'Show the add command button
        frmDetailedInfo.cmdAdd.Visible = True
        'Hide the modify command button
        frmDetailedInfo.cmdModify.Visible = False
        frmDetailedInfo.Show vbModal
    End If
End Sub

Private Sub mnuModify_Click()
    'Load modify
    LoadModify frmDetailedInfo, lstData, Student()
End Sub

Private Sub mnuDelete_Click()
    Dim DMsg As String
    Dim DTitle As String
    Dim DType As Integer
    Dim Response As Integer
    
    'Verify the user's request to delete
    DMsg = "Are you sure you want to delete this record?"
    DTitle = "Delete Record"
    DType = vbYesNo + vbQuestion
    Response = MsgBox(DMsg, DType, DTitle)
    
    If Response = vbYes Then
        'Go to delete procdure
        Delete lstData, Student(), NumStudents, CurrentRecord
    End If
    
    'Disable modify and delete menu buttons
    mnuModify.Enabled = False
    mnuDelete.Enabled = False
End Sub


Private Sub mnuExit_Click()
    'Close
    End
End Sub

Private Sub mnuOpen_Click()
    'Get file name
    FileName = GetFile(cdlDialog)
    
    'If the user actually entered a file name
    If FileName <> "" Then
        'Initialize variables
        Initialize Student(), NumStudents, FileName, MAX
        'If it is a text file
        If VBA.Right$(VBA.LCase$(FileName), 3) = "txt" Then
            'Go to the read text file procedure
            ReadTextFile Student(), NumStudents, FileName
        'If it is a record file
        ElseIf VBA.Right$(VBA.LCase$(FileName), 3) = "rec" Then
            'Go to the read record file procedure
            ReadRecordFile Student(), NumStudents, FileName, RecLength
        End If
        'If the file actually had data
        If NumStudents > 0 Then
            'Display information
            Display lblNumRecords, lstData, Student(), NumStudents
        End If
    End If
End Sub

Private Sub mnuSave_Click()
    'Get the name of the save file
    FileName = GetSave(cdlDialog)
    
    'If the user enetered a file name
    If FileName <> "" Then
        'Go to the save file procedure
        SaveFile Student(), NumStudents, FileName, RecLength
    End If
End Sub
