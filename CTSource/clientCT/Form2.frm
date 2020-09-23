VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Admin And Moderator Panel"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7440
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Unban"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send Report"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   2520
      Width           =   4455
   End
   Begin VB.TextBox report 
      Height          =   2055
      Left            =   2880
      TabIndex        =   3
      Text            =   "Report here"
      Top             =   360
      Width           =   4455
   End
   Begin VB.TextBox Ban 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ban"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   $"Form2.frx":0CCA
      Height          =   1095
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Line Line2 
      X1              =   2640
      X2              =   2640
      Y1              =   0
      Y2              =   3000
   End
   Begin VB.Line Line1 
      X1              =   2760
      X2              =   2760
      Y1              =   0
      Y2              =   3000
   End
   Begin VB.Label Label3 
      Caption         =   "Send report of ban to Dutchbull "
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Winsock1.SendData "BAN" & " " & Ban.Text
End Sub

Private Sub Command2_Click()
Form1.Winsock1.SendData "UNBAN" & " " & Ban.Text
End Sub

Private Sub Command3_Click()
Form1.Winsock1.SendData "AdminAlert" & " Report From: " & Form1.lblnick.Caption & "  ::::::  " & report.Text
MsgBox "Report Sended"
report.Text = ""
End Sub
