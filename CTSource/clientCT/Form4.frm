VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form Form4 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Link Partners"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4200
   DrawMode        =   2  'Blackness
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   1320
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lbllink7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   3600
      Width           =   4215
   End
   Begin VB.Label lbllink6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   3120
      Width           =   4215
   End
   Begin VB.Label lbllink5 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2640
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Dutchbull@darksoft3d.com to place a link"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   4095
   End
   Begin VB.Label lbllink4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   2160
      Width           =   4215
   End
   Begin VB.Label lbllink3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label lbllink2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Label lbllink1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Link Partners"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = True
End Sub



Private Sub lbllink1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
    Dim link
    link = ShellExecute(hWnd, "Open", MidWord(Inet1.OpenURL, 2, 1), &O0, &O0, SW_NORMAL)
End If
End Sub


Private Sub lbllink2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
    Dim link
    link = ShellExecute(hWnd, "Open", MidWord(Inet1.OpenURL, 4, 1), &O0, &O0, SW_NORMAL)
End If
End Sub


Private Sub lbllink3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
    Dim link
    link = ShellExecute(hWnd, "Open", MidWord(Inet1.OpenURL, 6, 1), &O0, &O0, SW_NORMAL)
End If
End Sub




Private Sub lbllink4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
    Dim link
    link = ShellExecute(hWnd, "Open", MidWord(Inet1.OpenURL, 8, 1), &O0, &O0, SW_NORMAL)
End If
End Sub

Private Sub lbllink5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
    Dim link
    link = ShellExecute(hWnd, "Open", MidWord(Inet1.OpenURL, 10, 1), &O0, &O0, SW_NORMAL)
End If
End Sub

Private Sub lbllink6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
    Dim link
    link = ShellExecute(hWnd, "Open", MidWord(Inet1.OpenURL, 12, 1), &O0, &O0, SW_NORMAL)
End If
End Sub

Private Sub lbllink7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
    Dim link
    link = ShellExecute(hWnd, "Open", MidWord(Inet1.OpenURL, 14, 1), &O0, &O0, SW_NORMAL)
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Inet1.OpenURL "http://www.darksoft3d.com/rpg/links.txt"
lbllink1.Caption = MidWord(Inet1.OpenURL, 1, 1)
lbllink2.Caption = MidWord(Inet1.OpenURL, 3, 1)
lbllink3.Caption = MidWord(Inet1.OpenURL, 5, 1)
lbllink4.Caption = MidWord(Inet1.OpenURL, 7, 1)
lbllink5.Caption = MidWord(Inet1.OpenURL, 9, 1)
lbllink6.Caption = MidWord(Inet1.OpenURL, 11, 1)
lbllink7.Caption = MidWord(Inet1.OpenURL, 13, 1)
Timer1.Enabled = False
End Sub
