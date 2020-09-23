VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Head Or Tail Game Made By Dutchbull"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Tail"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Head"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblmoney 
      Caption         =   "500"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Money:"
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblHOT2 
      Alignment       =   2  'Center
      Caption         =   "Bet on one first"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Head Or Tail Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label lblHOT 
      Alignment       =   2  'Center
      Caption         =   "Head Or Tail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Dont Copy The Money Labels they are allready in the game!! online the stuff on the left side!"
      Height          =   855
      Left            =   4920
      TabIndex        =   0
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Line Line1 
      X1              =   4320
      X2              =   4320
      Y1              =   0
      Y2              =   3000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
HOTHead
End Sub

Private Sub Command2_Click()
HOTTail
End Sub

