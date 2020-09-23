VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Hoger/Lager Game Made By Dutchbull"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Lager"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hoger"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   $"Form1.frx":0000
      Height          =   1095
      Left            =   5880
      TabIndex        =   9
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Line Line1 
      X1              =   5520
      X2              =   5520
      Y1              =   0
      Y2              =   2640
   End
   Begin VB.Label Label3 
      Caption         =   "Money"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblhl 
      Caption         =   "win or lose"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblmoney 
      Caption         =   "500"
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "nieuwe nummer"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "oude nummer"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblnewnumber 
      Caption         =   "Label1"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label lbloldnumber 
      Caption         =   "Label1"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Hoger
End Sub

Private Sub Command2_Click()
Lager
End Sub

Private Sub Form_Load()
GenNumbers
End Sub
