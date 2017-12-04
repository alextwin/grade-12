VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Random Access Files"
   ClientHeight    =   6960
   ClientLeft      =   7155
   ClientTop       =   3885
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
   Begin VB.PictureBox picData 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   120
      ScaleHeight     =   6195
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.Line linData 
         Visible         =   0   'False
         X1              =   0
         X2              =   8160
         Y1              =   360
         Y2              =   360
      End
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
      TabIndex        =   2
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
      TabIndex        =   1
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
      End
      Begin VB.Menu mnuModity 
         Caption         =   "&Modify Record..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSort 
         Caption         =   "Sort"
      End
      Begin VB.Menu mnuGetInfo 
         Caption         =   "Get Info"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
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
'Date: November 14, 2016
'Purpose: To read and display data from either a text file or record file and save that
'information to a record file
Option Explicit
Const MAX = 100
Dim Student(1 To MAX) As StudentRec
Dim NumStudents As Integer
Dim FileName As String
Dim RecLength As Integer             'Length of a record

Private Sub Form_Load()
    'Finding the length of a record
    RecLength = Len(Student(1))
End Sub

Private Sub mnuAbout_Click()
    'Loading about form
    Load frmAbout
    'Displaying about form
    frmAbout.Show vbModal
End Sub

Private Sub mnuAdd_Click()
    'Add Record function will be added in next version
    MsgBox "This function is currently not available and will be implemented in the next version.", vbOKOnly & vbInformation, "Not Available"
End Sub

Private Sub mnuBinary_Click()
    Dim SearchName As String
    Dim Position As Integer
    
    SearchName = InputBox$("Enter the mans last name:", "Search Name")
    Position = Binary_Search(Student(), NumStudents, SearchName)
    If Position <> -1 Then
        MsgBox SearchName & " is found at position " & Str$(Position) & ".", vbOKOnly + vbInformation, "Search Name"
    Else
        MsgBox SearchName & " cannot be found.", vbOKOnly + vbCritical, "Not Found"
    End If
End Sub

Private Sub mnuExit_Click()
    'Close
    End
End Sub

Private Sub mnuGetInfo_Click()
    Dim RecNum As Integer
    Dim TotalRecords As Integer
    Dim Client As StudentRec
    
    RecNum = Val(InputBox$("Enter a record number:", "Record Number"))
    TotalRecords = FileLen(FileName) \ RecLength
    
    If RecNum < 1 Or RecNum > TotalRecords Then
        MsgBox "Oops! Record #" & Trim$(Str$(RecNum)) & " does not seem to exist!", vbInformation, "Invalid Record Number"
    Else
        Open FileName For Random As #1 Len = RecLength
        Get #1, RecNum, Client
        MsgBox RTrim$(Client.FirstName) & " " & Client.LastName & vbCrLf & Client.HF & vbCrLf & Client.Mark, vbInformation, "Record Details"
        Close #1
    End If
End Sub

Private Sub mnuOpen_Click()
    'Initialize variables
    Initialize Student(), NumStudents, FileName, MAX
    
    'Get file name
    FileName = GetFile(cdlDialog)
    
    'If the user actually entered a file name
    If FileName <> "" Then
        'If it is a text file
        If Right$(LCase$(FileName), 3) = "txt" Then
            'Go to the read text file procedure
            ReadTextFile Student(), NumStudents, FileName
        'If it is a record file
        ElseIf Right$(LCase$(FileName), 3) = "rec" Then
            'Go to the read record file procedure
            ReadRecordFile Student(), NumStudents, FileName, RecLength
        End If
        'If the file actually had data
        If NumStudents > 0 Then
            'Make line visible
            linData.Visible = True
            'Display information
            Display lblNumRecords, picData, Student(), NumStudents
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

Private Sub mnuSequential_Click()
    Dim SearchName As String
    Dim Position As Integer
    
    SearchName = InputBox$("Enter the mans last name:", "Search Name")
    Position = SeqSearch(Student(), NumStudents, SearchName)
    If Position <> -1 Then
        MsgBox SearchName & " is found at position " & Str$(Position) & ".", vbOKOnly + vbInformation, "Search Name"
    Else
        MsgBox SearchName & " cannot be found.", vbOKOnly + vbCritical, "Not Found"
    End If
End Sub

Private Sub mnuSort_Click()
    BubbleSort Student(), NumStudents
End Sub
