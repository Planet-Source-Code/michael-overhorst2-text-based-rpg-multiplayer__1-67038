VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Raad Het Getal  By Dutchbull"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtgetal 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Text            =   "0  to 10"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton txtgetalknop 
      Caption         =   "kijk of het goed is"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Dont Copy the money labels!! they are allready in Crime Time only copy the left part and the module!"
      Height          =   1215
      Left            =   4800
      TabIndex        =   7
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lblinfoo 
      Alignment       =   2  'Center
      Caption         =   "Typ een Getal in"
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
      Left            =   600
      TabIndex        =   6
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label lblgetal 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lbllalala 
      Caption         =   "Getal:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   4440
      X2              =   4440
      Y1              =   0
      Y2              =   2400
   End
   Begin VB.Label lblmoney 
      Caption         =   "500"
      Height          =   255
      Left            =   6000
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "money:"
      Height          =   255
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtgetalknop_Click()
RaadGetal
End Sub
