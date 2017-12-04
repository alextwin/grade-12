VERSION 5.00
Begin VB.Form frmDetailedInfo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Record"
   ClientHeight    =   2280
   ClientLeft      =   7755
   ClientTop       =   6510
   ClientWidth     =   7245
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "A3DetailedInfo_TwinA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   7245
   Begin VB.CommandButton cmdModify 
      Caption         =   "Modify"
      Height          =   735
      Left            =   6120
      TabIndex        =   9
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   735
      Left            =   6120
      TabIndex        =   10
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   735
      Left            =   6120
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtMark 
      Height          =   375
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtHF 
      Height          =   375
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtFirstName 
      Height          =   375
      Left            =   1920
      MaxLength       =   15
      TabIndex        =   5
      Top             =   720
      Width           =   4095
   End
   Begin VB.TextBox txtLastName 
      Height          =   375
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   4
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label4 
      Caption         =   "Mark:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "HF:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "First Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Last Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmDetailedInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    Dim LastName As String
    Dim HF As String
    Dim Mark As Integer
    Dim FirstName As String
    
    FirstName = txtFirstName.Text
    HF = txtHF.Text
    LastName = txtLastName.Text
    
    'Validating user's input
    'Prevents user from adding too many records
    If NumStudents = MAX Then
        MsgBox "You have reached the maximum amount of records", vbOKOnly + vbCritical, "Maximum Records Reached"
    'No last name entered
    ElseIf VBA.Trim$(LastName) = "" Then
        MsgBox "Please enter a last name", vbOKOnly + vbCritical, "Missing Last Name"
        txtLastName.Text = ""
        txtLastName.SetFocus
    'No first name entered
    ElseIf VBA.Trim$(FirstName) = "" Then
        MsgBox "Please enter a first name", vbOKOnly + vbCritical, "Missing First Name"
        txtFirstName.Text = ""
        txtFirstName.SetFocus
    'No Home form entered or invalid home form
    ElseIf (Val(HF) < 9 Or Val(HF) > 12) Or (VBA.Right$(HF, 1) < "A" Or VBA.Right$(HF, 1) > "Z") Then
        MsgBox "Please enter a valid home form", vbOKOnly + vbCritical, "Invalid Home Form"
        txtHF.Text = ""
        txtHF.SetFocus
    'Checks for null string input
    ElseIf Not IsNumeric(txtMark.Text) Then
        MsgBox "Please enter a mark", vbOKOnly + vbCritical, "No Mark"
        txtMark.Text = ""
        txtHF.SetFocus
    ElseIf IsNumeric(txtMark.Text) Then
        Mark = Val(txtMark.Text)
        'Checks for valid range
        If Val(Mark) < 0 Or Val(Mark) > 100 Then
            MsgBox "Please enter a mark between 0-100", vbOKOnly + vbCritical, "Invalid Mark"
            txtMark.Text = ""
            txtMark.SetFocus
        'If everything is valid, add record
        Else
            'Update total
            NumStudents = NumStudents + 1
            'Store new record into record array
            Student(NumStudents).LastName = LastName
            Student(NumStudents).FirstName = FirstName
            Student(NumStudents).HF = HF
            Student(NumStudents).Mark = Mark
            'Display record
            AddRecord frmMain.lblNumRecords, frmMain.lstData, Student(), NumStudents
            'Empty fields
            txtLastName.Text = ""
            txtFirstName.Text = ""
            txtHF.Text = ""
            txtMark.Text = ""
            txtLastName.SetFocus
        End If
    End If
End Sub

Private Sub cmdDone_Click()
    Unload frmDetailedInfo
End Sub

Private Sub cmdModify_Click()
    Dim LastName As String
    Dim Mark As Integer
    Dim HF As String
    Dim FirstName As String
    
    FirstName = txtFirstName.Text
    HF = txtHF.Text
    LastName = txtLastName.Text
    
    'Validating user's input
    'No last name entered
    If VBA.Trim$(LastName) = "" Then
        MsgBox "Please enter a last name", vbOKOnly + vbCritical, "Missing Last Name"
        txtLastName.Text = ""
        txtLastName.SetFocus
    'No first name entered
    ElseIf VBA.Trim$(FirstName) = "" Then
        MsgBox "Please enter a first name", vbOKOnly + vbCritical, "Missing First Name"
        txtFirstName.Text = ""
        txtFirstName.SetFocus
    'No Home form entered or invalid home form
    ElseIf (Val(HF) < 9 Or Val(HF) > 12) Or (VBA.Right$(HF, 1) < "A" Or VBA.Right$(HF, 1) > "Z") Then
        MsgBox "Please enter a valid home form", vbOKOnly + vbCritical, "Invalid Home Form"
        txtHF.Text = ""
        txtHF.SetFocus
    'Checks for null string input
    ElseIf Not IsNumeric(txtMark.Text) Then
        MsgBox "Please enter a mark", vbOKOnly + vbCritical, "No Mark"
        txtMark.Text = ""
        txtHF.SetFocus
    ElseIf IsNumeric(txtMark.Text) Then
        Mark = Val(txtMark.Text)
        'Checks for valid range
        If Val(Mark) < 0 Or Val(Mark) > 100 Then
            MsgBox "Please enter a mark between 0-100", vbOKOnly + vbCritical, "Invalid Mark"
            txtMark.Text = ""
            txtMark.SetFocus
        'If everything is valid, modify
        Else
            'Display modified record
            Modify frmMain.lstData, Student(), LastName, FirstName, HF, Mark
            txtLastName.SetFocus
        End If
    End If
End Sub

Private Sub txtHF_KeyPress(KeyAscii As Integer)
    Dim Ch As Integer
    
    Ch = KeyAscii
    
    'Uppercase letters automatically
    If VBA.Chr$(Ch) >= "a" And VBA.Chr$(Ch) <= "z" Then
        KeyAscii = Asc(VBA.UCase$(VBA.Chr$(KeyAscii)))
    'Prevent user from putting anything other than numbers and letters
    ElseIf Ch <> vbKeyBack And (VBA.Chr$(Ch) < "0" Or VBA.Chr$(Ch) > "9") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtMark_KeyPress(KeyAscii As Integer)
    Dim Ch As Integer
    
    Ch = KeyAscii
    
    'Prevent user from inputting anything but numbers
    If Ch <> vbKeyBack And (VBA.Chr$(Ch) < "0" Or VBA.Chr$(Ch) > "9") Then
        KeyAscii = 0
    End If
End Sub
