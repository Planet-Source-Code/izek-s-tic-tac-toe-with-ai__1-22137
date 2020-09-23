VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   Icon            =   "tic tac toe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   3630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   " Plyer 1"
      Height          =   975
      Left            =   1200
      TabIndex        =   15
      Top             =   3600
      Width           =   1095
      Begin VB.OptionButton Option1 
         Caption         =   "Human"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "CPU"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Player 2"
      Height          =   975
      Left            =   2400
      TabIndex        =   12
      Top             =   3600
      Width           =   1095
      Begin VB.OptionButton Option3 
         Caption         =   "Human"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "CPU"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   1065
      Index           =   2
      Left            =   600
      Picture         =   "tic tac toe.frx":000C
      ScaleHeight     =   1005
      ScaleWidth      =   1005
      TabIndex        =   10
      Top             =   5640
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   1065
      Index           =   1
      Left            =   1680
      Picture         =   "tic tac toe.frx":1EEF
      ScaleHeight     =   1005
      ScaleWidth      =   1005
      TabIndex        =   9
      Top             =   5640
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.PictureBox Picture1 
      Height          =   1000
      Index           =   9
      Left            =   2520
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   8
      Top             =   2520
      Width           =   1000
   End
   Begin VB.PictureBox Picture1 
      Height          =   1000
      Index           =   8
      Left            =   1320
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   7
      Top             =   2520
      Width           =   1000
   End
   Begin VB.PictureBox Picture1 
      Height          =   1000
      Index           =   7
      Left            =   120
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   6
      Top             =   2520
      Width           =   1000
   End
   Begin VB.PictureBox Picture1 
      Height          =   1000
      Index           =   6
      Left            =   2520
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   5
      Top             =   1320
      Width           =   1000
   End
   Begin VB.PictureBox Picture1 
      Height          =   1000
      Index           =   5
      Left            =   1320
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   4
      Top             =   1320
      Width           =   1000
   End
   Begin VB.PictureBox Picture1 
      Height          =   1000
      Index           =   4
      Left            =   120
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   3
      Top             =   1320
      Width           =   1000
   End
   Begin VB.PictureBox Picture1 
      Height          =   1000
      Index           =   3
      Left            =   2520
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   2
      Top             =   120
      Width           =   1000
   End
   Begin VB.PictureBox Picture1 
      Height          =   1000
      Index           =   2
      Left            =   1320
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   1
      Top             =   120
      Width           =   1000
   End
   Begin VB.PictureBox Picture1 
      Height          =   1000
      Index           =   1
      Left            =   120
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   0
      Top             =   120
      Width           =   1000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim leftx&, topy&
Private Sub Command1_Click()
'restart the game if settings change
Call GameOver
End Sub
Private Sub Form_Load()
'restart the game if settings change
Call GameOver
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.WindowState = vbMinimized
End Sub
Private Sub Label2_Click()
End
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
leftx = x
topy = Y
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    Me.Move Me.Left + x - leftx, Me.Top + Y - topy
End If
End Sub
Private Sub Option1_Click()
'restart the game if settings change
Call GameOver
End Sub
Private Sub Option2_Click()
'restart the game if settings change
Call GameOver
End Sub
Private Sub Option3_Click()
'restart the game if settings change
Call GameOver
End Sub
Private Sub Option4_Click()
'restart the game if settings change
Call GameOver
End Sub
Private Sub Picture1_Click(index As Integer)
'update players move
Call PlayerMove(index)
End Sub

