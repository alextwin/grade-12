VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Recursion"
   ClientHeight    =   9360
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   11430
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   11430
   Begin VB.CommandButton cmdReverse 
      Caption         =   "Reverse"
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton cmdFibonacci 
      Caption         =   "Fibonacci"
      Height          =   615
      Left            =   8400
      TabIndex        =   2
      Top             =   8040
      Width           =   2295
   End
   Begin VB.CommandButton cmdPower 
      Caption         =   "Power"
      Height          =   735
      Left            =   5280
      TabIndex        =   1
      Top             =   7920
      Width           =   2295
   End
   Begin VB.CommandButton cmdFactorial 
      Caption         =   "Factorial"
      Height          =   855
      Left            =   1560
      TabIndex        =   0
      Top             =   8040
      Width           =   3135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFactorial_Click()
    Dim Num As Integer
    
    Cls
    Num = Val(InputBox$("Enter a number:", "Factorial Number"))
    Print Factorial(Num)
End Sub

Private Sub cmdFibonacci_Click()
    Dim Num As Integer
    
    Cls
    
    Num = Val(InputBox$("Enter a number:", "Fibonacci Number:"))
    Print Fibonacci(Num)
End Sub

Private Sub cmdPower_Click()
    Dim Base As Integer
    Dim Exponent As Integer
    Dim Answer As Single
    
    Cls
    
    Base = Val(InputBox$("Enter the base:", "Base"))
    Exponent = Val(InputBox$("Enter the exponent:", "Exponent"))
    If Exponent = 0 Then
        Answer = 1
    ElseIf Exponent > 0 Then
        Answer = Power(Base, Exponent)
    Else
        Answer = 1 / Power(Base, Abs(Exponent))
    End If
    
    Print Answer
End Sub

Private Sub cmdReverse_Click()
    Const MAX = 25
    Dim Num(1 To MAX) As Integer
    Dim K As Integer
    
    For K = 1 To MAX
        Num(K) = K
    Next K
    Reverse Num(), MAX
End Sub
