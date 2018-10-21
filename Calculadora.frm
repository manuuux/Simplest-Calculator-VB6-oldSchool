VERSION 5.00
Begin VB.Form Calculadora 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calc"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   1950
   Icon            =   "Calculadora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   1950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "C"
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Results 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton CommandMult 
      Caption         =   "*"
      Height          =   495
      Left            =   960
      TabIndex        =   14
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton CommandDivide 
      Caption         =   "/"
      Height          =   495
      Left            =   1440
      TabIndex        =   13
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton CommandMinus 
      Caption         =   "-"
      Height          =   495
      Left            =   1440
      TabIndex        =   12
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton CommandSum 
      Caption         =   "+"
      Height          =   975
      Left            =   1440
      TabIndex        =   11
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton CommandEqual 
      Caption         =   "="
      Height          =   495
      Left            =   960
      TabIndex        =   10
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command0 
      Caption         =   "0"
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   495
      Left            =   960
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   495
   End
End
Attribute VB_Name = "Calculadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PreviusResult As Long
Public operation As String
Private Sub Command0_Click()
Results.Text = Results.Text + "0"
End Sub
Private Sub Command1_Click(Index As Integer)
Results.Text = Results.Text + "1"
End Sub

Private Sub Command10_Click()
Results.Text = ""
PreviusResult = 0
End Sub

Private Sub Command2_Click()
Results.Text = Results.Text + "2"
End Sub

Private Sub Command3_Click()
Results.Text = Results.Text + "3"
End Sub

Private Sub Command4_Click()
Results.Text = Results.Text + "4"
End Sub

Private Sub Command5_Click()
Results.Text = Results.Text + "5"
End Sub

Private Sub Command6_Click()
Results.Text = Results.Text + "6"
End Sub

Private Sub Command7_Click()
Results.Text = Results.Text + "7"
End Sub

Private Sub Command8_Click()
Results.Text = Results.Text + "8"
End Sub

Private Sub Command9_Click()
Results.Text = Results.Text + "9"
End Sub

Private Sub CommandDivide_Click()
PreviusResult = CLng(Results.Text)
Results.Text = ""
operation = "div"
End Sub

Private Sub CommandDot_Click()
Results.Text = Results.Text + "."
End Sub

Private Sub CommandEqual_Click()
total
End Sub

Private Sub CommandMinus_Click()
PreviusResult = CLng(Results.Text)
Results.Text = ""
operation = "minus"
End Sub

Private Sub CommandMult_Click()
PreviusResult = CLng(Results.Text)
Results.Text = ""
operation = "mult"
End Sub

Private Sub CommandSum_Click()
PreviusResult = CLng(Val(Results.Text))
Results.Text = ""
operation = "sum"
End Sub

Private Function total()
    If operation = "sum" Then
        PreviusResult = PreviusResult + CLng(Results.Text)
    ElseIf operation = "minus" Then
        PreviusResult = PreviusResult - CLng(Results.Text)
    ElseIf operation = "div" Then
        PreviusResult = PreviusResult / CLng(Results.Text)
    ElseIf operation = "mult" Then
        PreviusResult = PreviusResult * CLng(Results.Text)
    End If
    Results.Text = CStr(PreviusResult)
    operation = "none"
End Function


Private Sub Form_Load()
PreviusResult = 0
operation = "none"
End Sub
