VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Night Walkers"
   ClientHeight    =   4020
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6150
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2774.676
   ScaleMode       =   0  'User
   ScaleWidth      =   5775.168
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4680
      TabIndex        =   0
      Top             =   3600
      Width           =   1260
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      Caption         =   "Special thanks to my crew!"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   5895
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      Caption         =   $"frmAbout.frx":0CCA
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   5895
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "©2006-2010 Night Walkers"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label lblurll 
      BackColor       =   &H80000009&
      Caption         =   "http://www.Night-Walkers.org"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Please Contact Dutchbull@darksoft3d.com if you find a error or a bug."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   5895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5634.31
      Y1              =   2319.132
      Y2              =   2319.132
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H80000009&
      Caption         =   "Night Walkers is a multiplayer Text-Based rpg and even the first multiplayer text-based rpg application based with this systems!!"
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   120
      TabIndex        =   1
      Top             =   1005
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H80000009&
      Caption         =   "Night Walkers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   90
      TabIndex        =   2
      Top             =   120
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H80000009&
      Caption         =   "Version 0.0 Public Beta"
      Height          =   225
      Left            =   90
      TabIndex        =   3
      Top             =   660
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
lblVersion.Caption = "Version" & " " & " " & App.Major & "." & App.Minor
End Sub

Private Sub lblurll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Dim link
    link = ShellExecute(hWnd, "Open", "http://www.night-walkers.org", &O0, &O0, SW_NORMAL)
End If
End Sub
