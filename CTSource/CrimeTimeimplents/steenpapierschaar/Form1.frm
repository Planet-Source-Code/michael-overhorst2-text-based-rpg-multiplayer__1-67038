VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "paper scissors stone by Dutchbull"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Stone"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Scissors"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Paper"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Dont have to copy the money labels... they are allready in Crime Time"
      Height          =   1455
      Left            =   5520
      TabIndex        =   8
      Top             =   960
      Width           =   2535
   End
   Begin VB.Line Line1 
      X1              =   5040
      X2              =   5040
      Y1              =   0
      Y2              =   3120
   End
   Begin VB.Label Label2 
      Caption         =   "Money:"
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblmoney 
      Caption         =   "500"
      Height          =   255
      Left            =   6480
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblcomphand 
      Caption         =   "Choose one First"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Computer has:"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblpss 
      Alignment       =   2  'Center
      Caption         =   "Win Or Lose"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   2760
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    ComputerGen
    Paper

End Sub

Private Sub Command2_Click()

    ComputerGen
    Scissors

End Sub

Private Sub Command3_Click()

    ComputerGen
    Stone

End Sub

