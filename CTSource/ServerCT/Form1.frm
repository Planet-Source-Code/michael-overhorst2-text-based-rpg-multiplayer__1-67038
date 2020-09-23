VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game RPG Server"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8625
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   8625
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0CCA
      Left            =   2160
      List            =   "Form1.frx":0CD1
      TabIndex        =   15
      Text            =   "http://musicvbrpg.no-ip.info:8000"
      Top             =   3120
      Width           =   3135
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Discon all clients"
      Height          =   375
      Left            =   6840
      TabIndex        =   14
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stream On"
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   975
      Left            =   6840
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "Form1.frx":0D10
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "alert message to all"
      Height          =   375
      Left            =   6840
      TabIndex        =   11
      Top             =   3000
      Width           =   1695
   End
   Begin VB.ListBox alertlist 
      Height          =   255
      Left            =   7680
      TabIndex        =   10
      Top             =   4560
      Width           =   135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reset Usercount"
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Kick"
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6840
      TabIndex        =   7
      Text            =   "connection index"
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   3075
      Width           =   1290
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1950
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Initializing Ports 0 - 9999"
      Top             =   1260
      Width           =   2850
   End
   Begin MSWinsockLib.Winsock host 
      Index           =   0
      Left            =   45
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   21
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   150
      Left            =   1950
      TabIndex        =   4
      Top             =   1545
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   265
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   9999
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1350
      Top             =   45
   End
   Begin VB.ListBox List2 
      Height          =   840
      ItemData        =   "Form1.frx":0D1A
      Left            =   5745
      List            =   "Form1.frx":0D1C
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   165
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00004000&
      Height          =   2955
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   6510
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "Kick User"
      Height          =   255
      Left            =   7200
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   4320
      TabIndex        =   2
      Top             =   3120
      Width           =   2280
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   3015
      Left            =   60
      Top             =   60
      Width           =   6570
   End
   Begin VB.Menu tray 
      Caption         =   "To Tray"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim databuffer As String
Dim connection As Long
Dim maxconnections As Long

Private Sub Command1_Click()
On Error Resume Next
    End
End Sub

Private Sub Command2_Click()
On Error Resume Next
host(Text2.Text).SendData "KICKED"
End Sub

Private Sub Command3_Click()
On Error Resume Next
SetINI "Reggedusers", "Usercount", "0"
End Sub

Private Sub Command4_Click()
On Error Resume Next
ALERTMSG (Text3.Text)
End Sub

Private Sub Command5_Click()
Dim I As Integer
 On Error Resume Next
    
    For I = 1 To host.Count
        
        If host(I).State = 7 Then
        
            host(I).SendData "STREAMON " & Combo1.Text
            DoEvents
        Else
        
        End If
        
    Next I
End Sub

Private Sub Command6_Click()
    Dim I As Integer
 On Error Resume Next
    
    For I = 1 To host.Count
        
        If host(I).State = 7 Then
        
            host(I).SendData "OFFERROR"
            DoEvents
        Else
        
        End If
        
    Next I
End Sub

Private Sub Form_Load()
On Error Resume Next

    
    Form1.Show
    
    maxconnections = 1000
    
    host(0).LocalPort = 6666
    
    
    Text1.Text = "Enumerating Sockets 1 to " & maxconnections
    ProgressBar1.Max = maxconnections
    For x = 0 To maxconnections: DoEvents
        List2.AddItem Format(x, "0000")
        ProgressBar1.Value = x:
    Next x
    ProgressBar1.Visible = False: Text1.Visible = False: List1.Visible = True
    host(0).Listen
    Label1.Caption = "Hosting on port: " & host(0).LocalPort
    If FileExist(App.Path & "/Alerts.txt") = True Then
    LoadList App.Path & "/Alerts.txt", alertlist
    End If
End Sub

Private Sub host_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
Dim usercount As String
usercount = GetINI("Reggedusers", "Usercount")
    List2.Selected(0) = True
    connection = List2.Text
    connection = connection + 1
    If host(0).State = 2 Then
        Load host(connection)
        host(connection).Accept requestID
        host(connection).SendData "ONLINE " & host.Count - 1 & " " & usercount
        List2.RemoveItem (0)
        List1.AddItem "Connected: " & Format(connection, "0000")
        If Timer1.Enabled = False Then Timer1.Enabled = True
    End If

End Sub

Private Sub host_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
    Dim StrData As String
    Dim hstname As String
    Dim usercount As String
    usercount = GetINI("Reggedusers", "Usercount") + 1
    
    host(Index).GetData StrData
    hstname = host(Index).RemoteHostIP
    List1.AddItem Format(Index, "0000") & " DATA: " & StrData
    
If StrData = "DISCON" Then
host(Index).Close
End If

If MidWord(StrData, 1, 1) = "AdminAlert" Then
StrData = DelWord(StrData, 1, 1)
alertlist.AddItem "------------------------------"
alertlist.AddItem StrData
alertlist.AddItem "------------------------------"
SaveList App.Path & "/Alerts.txt", alertlist
End If

If MidWord(StrData, 1, 1) = "REG" Then
Dim reffs As String
If Len(GetINI("Users", MidWord(StrData, 2, 1))) Or Len(GetINI("Users", hstname)) Then
host(Index).SendData "ACC EXIST"
Else
reffs = GetINI("Users", "Reffs" & MidWord(StrData, 5, 1)) + 1

If reffs >= "100" Then
If GetINI("Users", "What" & MidWord(StrData, 5, 1)) = "Member" Then
SetINI "Users", "What" & MidWord(StrData, 5, 1), "VIP"
End If
End If

SetINI "Users", "-------------------------=" & MidWord(StrData, 2, 1), "-------------------------"
SetINI "Users", MidWord(StrData, 2, 1), MidWord(StrData, 3, 1)
SetINI "Users", "Cash" & MidWord(StrData, 2, 1), "2000"
SetINI "Users", "Power" & MidWord(StrData, 2, 1), "100"
SetINI "Users", "Ras" & MidWord(StrData, 2, 1), MidWord(StrData, 4, 1)
SetINI "Users", "Ban" & MidWord(StrData, 2, 1), "0"
SetINI "Users", "What" & MidWord(StrData, 2, 1), "Member"
SetINI "Users", "Reffs" & MidWord(StrData, 2, 1), "0"
SetINI "Users", "Reffs" & MidWord(StrData, 5, 1), reffs
SetINI "Reggedusers", "Usercount", usercount
SetINI "Users", hstname, "1"
host(Index).SendData "ACC CREATED"
End If
End If
If MidWord(StrData, 1, 1) = "SAV" Then
SetINI "Users", "Cash" & MidWord(StrData, 2, 1), MidWord(StrData, 3, 1)
SetINI "Users", "Power" & MidWord(StrData, 2, 1), MidWord(StrData, 4, 1)
End If

If MidWord(StrData, 1, 1) = "BAN" Then
If Len(GetINI("Users", MidWord(StrData, 2, 1))) Then
SetINI "Users", "Ban" & MidWord(StrData, 2, 1), "1"
Else
host(Index).SendData "NO ACC"
End If
End If

If MidWord(StrData, 1, 1) = "UNBAN" Then
If Len(GetINI("Users", MidWord(StrData, 2, 1))) Then
SetINI "Users", "Ban" & MidWord(StrData, 2, 1), "0"
Else
host(Index).SendData "NO ACC"
End If
End If

If MidWord(StrData, 1, 1) = "RESPOW" Then
If Len(GetINI("Users", MidWord(StrData, 2, 1))) Then
SetINI "Users", "Power" & MidWord(StrData, 2, 1), "0"
Else
host(Index).SendData "NO ACC"
End If
End If

If MidWord(StrData, 1, 1) = "CHAT" Then
StrData = DelWord(StrData, 1, 1)
SendChat StrData
End If


If MidWord(StrData, 1, 1) = "LGN" Then
If GetINI("Users", "Ban" & MidWord(StrData, 2, 1)) = "1" Then
host(Index).SendData "BANNED"
Else
If GetINI("Users", MidWord(StrData, 2, 1)) = MidWord(StrData, 3, 1) Then
host(Index).SendData "LOGIN OK" & " " & GetINI("Users", "Cash" & MidWord(StrData, 2, 1)) & " " & GetINI("Users", "Ras" & MidWord(StrData, 2, 1)) & " " & GetINI("Users", "Power" & MidWord(StrData, 2, 1)) & " " & GetINI("Users", "What" & MidWord(StrData, 2, 1)) & " " & GetINI("Users", "Reffs" & MidWord(StrData, 2, 1))
Else
host(Index).SendData "LOGIN WRONG"
End If
End If
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    DoEvents
    On Error GoTo errhandler
    For x = 1 To host.Count
    DoEvents
    On Error GoTo errhandler
        If host(x).State = 8 Then
            host(x).Close
            List2.AddItem Format(x - 1, "0000"): Unload host(x)
            List1.AddItem "Disconnected: " & Format(x, "0000")
        End If
next1:
    Next x
Exit Sub
errhandler:
If Err.Number = 340 Then
    On Error GoTo errhandler
    GoTo next1
Else: MsgBox Err.Number & ": " & Err.Description, vbOKOnly, "Error"
End If
End Sub

Private Sub SendChat(ByVal chattext As String)

Dim I      As Long
Dim WinCnt As Long

    WinCnt = host.Count - 1
    For I = 0 To WinCnt
        If Not host(I) Is Nothing Then
            If host(I).State = 7 Then
                host(I).SendData chattext
            End If
        End If
    Next I

Exit Sub

  
End Sub


Public Sub ALERTMSG(Text As String)
 Dim I As Integer
 On Error Resume Next
    
    For I = 1 To host.Count
        
        If host(I).State = 7 Then
        
            host(I).SendData "ALERTMSG " & Text
            DoEvents
        Else
        
        End If
        
    Next I
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    Static lngMsg As Long
    Dim blnflag As Boolean, lngResult As Long
    
    lngMsg = x / Screen.TwipsPerPixelX
    If blnflag = False Then
        blnflag = True
        Select Case lngMsg
        Case WM_RBUTTONUP
            Call SetForegroundWindow(Me.hWnd)
        Case WM_LBUTTONDBLCLK
            Call SystrayOff(Form1)
            Call SetForegroundWindow(Me.hWnd)
            Form1.Show
            FormOnTop Form1
        End Select
        blnflag = False
    End If
End Sub

Private Sub tray_Click()
On Error Resume Next
 Call SystrayOn(Form1, "Crime Time Server")
    Form1.Hide
End Sub
