VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Night Walkers"
   ClientHeight    =   6240
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7905
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   6240
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6375
      Left            =   -360
      TabIndex        =   0
      Top             =   -120
      Width           =   8415
      Begin VB.Timer Timer8 
         Interval        =   60000
         Left            =   5040
         Top             =   3360
      End
      Begin VB.Timer Timer7 
         Interval        =   500
         Left            =   7080
         Top             =   4320
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000009&
         Caption         =   "Save my information."
         Height          =   375
         Left            =   6480
         TabIndex        =   11
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin InetCtlsObjects.Inet Inet2 
         Left            =   3120
         Top             =   2760
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         Protocol        =   4
         URL             =   "http://"
      End
      Begin VB.Timer newstimer 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3360
         Top             =   1800
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lycan"
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Top             =   3240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vampire"
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000000&
         ForeColor       =   &H80000006&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   480
         MaxLength       =   9
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000000&
         ForeColor       =   &H80000006&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   480
         MaxLength       =   9
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000000&
         ForeColor       =   &H80000006&
         Height          =   285
         Left            =   480
         MaxLength       =   9
         TabIndex        =   1
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H80000000&
         ForeColor       =   &H80000006&
         Height          =   285
         Left            =   480
         TabIndex        =   6
         Top             =   3960
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Register"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000000&
         ForeColor       =   &H80000006&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   6360
         MaxLength       =   9
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000000&
         ForeColor       =   &H80000006&
         Height          =   285
         Left            =   6360
         MaxLength       =   9
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Login"
         Height          =   255
         Left            =   6480
         TabIndex        =   10
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Timer Timer3 
         Interval        =   1000
         Left            =   4200
         Top             =   2520
      End
      Begin VB.Timer Timerexit 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   3960
         Top             =   3240
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   3480
         Top             =   3240
      End
      Begin VB.Label lblversionn 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   97
         Top             =   6000
         Width           =   735
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Your Running Version:"
         Height          =   255
         Left            =   480
         TabIndex        =   96
         Top             =   6000
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "What do you want to be?"
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password:"
         Height          =   255
         Left            =   720
         TabIndex        =   21
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Register"
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
         TabIndex        =   17
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblreff 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Refferal: (leave empty for none)"
         Height          =   255
         Left            =   480
         TabIndex        =   92
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label lblusercount 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7560
         TabIndex        =   47
         Top             =   5040
         Width           =   615
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Registered Users :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6120
         TabIndex        =   46
         Top             =   5040
         Width           =   1335
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Members Online:"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6240
         TabIndex        =   45
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label lblonline 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7560
         TabIndex        =   44
         Top             =   5400
         Width           =   615
      End
      Begin VB.Label lblstatuss 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Please Register Or Login"
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
         Height          =   855
         Left            =   480
         TabIndex        =   18
         Top             =   4680
         Width           =   7695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
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
         Index           =   0
         Left            =   6840
         TabIndex        =   14
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   255
         Left            =   6840
         TabIndex        =   13
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         Height          =   255
         Index           =   0
         Left            =   6840
         TabIndex        =   12
         Top             =   840
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   6195
         Left            =   2280
         Picture         =   "Form1.frx":0CCA
         Top             =   120
         Width           =   4080
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   0
      TabIndex        =   19
      Top             =   720
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      Tabs            =   10
      TabsPerRow      =   10
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "News"
      TabPicture(0)   =   "Form1.frx":CEC5
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblnews"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label31"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Inet1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Crimes"
      TabPicture(1)   =   "Form1.frx":CEE1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text8"
      Tab(1).Control(1)=   "timertijd"
      Tab(1).Control(2)=   "Command3"
      Tab(1).Control(3)=   "crime2"
      Tab(1).Control(4)=   "crime3"
      Tab(1).Control(5)=   "crime1"
      Tab(1).Control(6)=   "lblcode1"
      Tab(1).Control(7)=   "Label34"
      Tab(1).Control(8)=   "lbltijd"
      Tab(1).Control(9)=   "Label13"
      Tab(1).Control(10)=   "lblkans"
      Tab(1).Control(11)=   "Label12"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Stats"
      TabPicture(2)   =   "Form1.frx":CEFD
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblnick"
      Tab(2).Control(1)=   "lblmoney"
      Tab(2).Control(2)=   "Label8"
      Tab(2).Control(3)=   "Label10"
      Tab(2).Control(4)=   "lblras"
      Tab(2).Control(5)=   "Label11"
      Tab(2).Control(6)=   "lblpower"
      Tab(2).Control(7)=   "lblra"
      Tab(2).Control(8)=   "lblrank"
      Tab(2).Control(9)=   "lblwhat"
      Tab(2).Control(10)=   "lblreffers"
      Tab(2).Control(11)=   "Label22"
      Tab(2).Control(12)=   "Label32"
      Tab(2).Control(13)=   "Timer2"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Shop"
      TabPicture(3)   =   "Form1.frx":CF19
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command7"
      Tab(3).Control(1)=   "Command6"
      Tab(3).Control(2)=   "Command5"
      Tab(3).Control(3)=   "Command4"
      Tab(3).Control(4)=   "Label15"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "VIP Crimes"
      TabPicture(4)   =   "Form1.frx":CF35
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Text10"
      Tab(4).Control(1)=   "Timertijd2"
      Tab(4).Control(2)=   "Option5"
      Tab(4).Control(3)=   "Option4"
      Tab(4).Control(4)=   "Option3"
      Tab(4).Control(5)=   "Command8"
      Tab(4).Control(6)=   "lblcode2"
      Tab(4).Control(7)=   "Label35"
      Tab(4).Control(8)=   "Label19"
      Tab(4).Control(9)=   "lbltime"
      Tab(4).Control(10)=   "Label20"
      Tab(4).Control(11)=   "lblkans2"
      Tab(4).Control(12)=   "Label18"
      Tab(4).ControlCount=   13
      TabCaption(5)   =   "VIP Shop"
      TabPicture(5)   =   "Form1.frx":CF51
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Command12"
      Tab(5).Control(1)=   "Command11"
      Tab(5).Control(2)=   "Command10"
      Tab(5).Control(3)=   "Command9"
      Tab(5).Control(4)=   "Label21"
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "Chat"
      TabPicture(6)   =   "Form1.frx":CF6D
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Text6"
      Tab(6).Control(1)=   "Text7"
      Tab(6).Control(2)=   "Command13"
      Tab(6).Control(3)=   "Command16"
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "Get VIP/Power"
      TabPicture(7)   =   "Form1.frx":CF89
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label23"
      Tab(7).Control(1)=   "Label24"
      Tab(7).Control(2)=   "Label25"
      Tab(7).Control(3)=   "Line1(0)"
      Tab(7).Control(4)=   "Line2(0)"
      Tab(7).Control(5)=   "Line3(0)"
      Tab(7).Control(6)=   "Line4(0)"
      Tab(7).Control(7)=   "Label33"
      Tab(7).Control(8)=   "Label36"
      Tab(7).ControlCount=   9
      TabCaption(8)   =   "Casino"
      TabPicture(8)   =   "Form1.frx":CFA5
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Label26"
      Tab(8).Control(1)=   "Label27"
      Tab(8).Control(2)=   "lbloldnumber"
      Tab(8).Control(3)=   "Label29"
      Tab(8).Control(4)=   "lblnewnumber"
      Tab(8).Control(5)=   "lblhl"
      Tab(8).Control(6)=   "Label30"
      Tab(8).Control(7)=   "Command14"
      Tab(8).Control(8)=   "Command15"
      Tab(8).Control(9)=   "Timer5"
      Tab(8).ControlCount=   10
      TabCaption(9)   =   "VIP Casino"
      TabPicture(9)   =   "Form1.frx":CFC1
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Label14"
      Tab(9).Control(1)=   "Line9"
      Tab(9).Control(2)=   "lblHOT2"
      Tab(9).Control(3)=   "Labelblabla"
      Tab(9).Control(4)=   "lblHOT"
      Tab(9).Control(5)=   "lblinfoo"
      Tab(9).Control(6)=   "lblgetal"
      Tab(9).Control(7)=   "lbllalala"
      Tab(9).Control(8)=   "Label3333(1)"
      Tab(9).Control(9)=   "Command222"
      Tab(9).Control(10)=   "Command111"
      Tab(9).Control(11)=   "txtgetal"
      Tab(9).Control(12)=   "txtgetalknop"
      Tab(9).Control(13)=   "Timer6"
      Tab(9).ControlCount=   14
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   -71040
         TabIndex        =   103
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   -70920
         TabIndex        =   100
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Enter"
         Height          =   375
         Left            =   -68280
         TabIndex        =   99
         Top             =   5040
         Width           =   855
      End
      Begin VB.Timer Timer6 
         Interval        =   5000
         Left            =   -71880
         Top             =   4200
      End
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   -69240
         Top             =   3240
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   3480
         Top             =   1800
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.CommandButton txtgetalknop 
         Caption         =   "Bet"
         Height          =   495
         Left            =   -70080
         TabIndex        =   85
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox txtgetal 
         Height          =   285
         Left            =   -70080
         TabIndex        =   84
         Text            =   "0  to 10"
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CommandButton Command111 
         Caption         =   "Head"
         Height          =   495
         Left            =   -74160
         TabIndex        =   80
         Top             =   2520
         Width           =   2055
      End
      Begin VB.CommandButton Command222 
         Caption         =   "Tail"
         Height          =   495
         Left            =   -74160
         TabIndex        =   79
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Lower"
         Height          =   495
         Left            =   -72240
         TabIndex        =   76
         Top             =   3540
         Width           =   2295
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Higher"
         Height          =   495
         Left            =   -72240
         TabIndex        =   75
         Top             =   2940
         Width           =   2295
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Send"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -68280
         TabIndex        =   66
         Top             =   5040
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         MaxLength       =   60
         TabIndex        =   65
         Top             =   5040
         Width           =   6375
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   4095
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   64
         Text            =   "Form1.frx":CFDD
         Top             =   840
         Width           =   7455
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Buy Apache (gives 8000 power costs 8000)"
         Height          =   495
         Left            =   -72960
         TabIndex        =   62
         Top             =   4140
         Width           =   3495
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Buy Gaurd (gives 4000 power costs 4000)"
         Height          =   495
         Left            =   -73080
         TabIndex        =   61
         Top             =   3540
         Width           =   3735
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Buy Grenade (gives 2000 power costs 2000)"
         Height          =   495
         Left            =   -73080
         TabIndex        =   60
         Top             =   2940
         Width           =   3735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Buy Uzi (gives 1000 power costs 1000)"
         Height          =   495
         Left            =   -73080
         TabIndex        =   59
         Top             =   2340
         Width           =   3735
      End
      Begin VB.Timer Timertijd2 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   -73800
         Top             =   3420
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Bite one of your race"
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
         Left            =   -72000
         TabIndex        =   53
         Top             =   3000
         Width           =   2775
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Steal blood from your group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72000
         TabIndex        =   52
         Top             =   2640
         Width           =   2895
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Bite a master"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72000
         TabIndex        =   51
         Top             =   2280
         Width           =   2415
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Do It VIP Style"
         Height          =   375
         Left            =   -72120
         TabIndex        =   49
         Top             =   4080
         Width           =   2175
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Buy Super Armor(gives 5000 power costs 50000)"
         Height          =   495
         Left            =   -72960
         TabIndex        =   43
         Top             =   4140
         Width           =   3735
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Buy Bomb (gives 500 power costs 5000)"
         Height          =   495
         Left            =   -73020
         TabIndex        =   42
         Top             =   3540
         Width           =   3855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Buy Stungun (gives 400 power costs 4000)"
         Height          =   495
         Left            =   -73020
         TabIndex        =   41
         Top             =   2940
         Width           =   3855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Buy Shotgun (gives 300 power costs 3000)"
         Height          =   495
         Left            =   -73020
         TabIndex        =   40
         Top             =   2340
         Width           =   3855
      End
      Begin VB.Timer timertijd 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   -73560
         Top             =   4080
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Do it"
         Height          =   375
         Left            =   -71880
         TabIndex        =   35
         Top             =   4080
         Width           =   2175
      End
      Begin VB.OptionButton crime2 
         Caption         =   "Rob the blood bank"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71880
         TabIndex        =   34
         Top             =   2640
         Width           =   2895
      End
      Begin VB.OptionButton crime3 
         Caption         =   "Steal blood of a slayer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71880
         TabIndex        =   33
         Top             =   3000
         Width           =   2535
      End
      Begin VB.OptionButton crime1 
         Caption         =   "Bite a human"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71880
         TabIndex        =   31
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   -68880
         Top             =   4560
      End
      Begin VB.Label Label36 
         Caption         =   "Note: if you call and you are allready VIP then you will get $5000 and 5000 power."
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   -72720
         TabIndex        =   106
         Top             =   3000
         Width           =   3495
      End
      Begin VB.Label lblcode2 
         Alignment       =   2  'Center
         Caption         =   "code here"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -71040
         TabIndex        =   105
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label35 
         Caption         =   "Code Here:"
         Height          =   255
         Left            =   -72120
         TabIndex        =   104
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label lblcode1 
         Alignment       =   2  'Center
         Caption         =   "code here"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -70920
         TabIndex        =   102
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label34 
         Caption         =   "Code here:"
         Height          =   255
         Left            =   -71760
         TabIndex        =   101
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label33 
         Caption         =   "Click Here To Go To The Payment Site"
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   -71760
         TabIndex        =   98
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Caption         =   "Stats"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74880
         TabIndex        =   95
         Top             =   720
         Width           =   7695
      End
      Begin VB.Label Label22 
         Caption         =   "Refferals :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72000
         TabIndex        =   94
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label lblreffers 
         Caption         =   "0"
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
         Left            =   -70680
         TabIndex        =   93
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Line Line10 
         X1              =   0
         X2              =   7920
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Caption         =   "NEWS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   91
         Top             =   480
         Width           =   7695
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Caption         =   "Higher / Lower"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74400
         TabIndex        =   90
         Top             =   960
         Width           =   6615
      End
      Begin VB.Label Label3333 
         Alignment       =   2  'Center
         Caption         =   "Guess the number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   -70800
         TabIndex        =   89
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Label lbllalala 
         Caption         =   "Number:"
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
         Left            =   -69840
         TabIndex        =   88
         Top             =   2280
         Width           =   855
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
         Left            =   -68880
         TabIndex        =   87
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblinfoo 
         Alignment       =   2  'Center
         Caption         =   "Guess the number First"
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
         Left            =   -70800
         TabIndex        =   86
         Top             =   3720
         Width           =   3375
      End
      Begin VB.Label lblHOT 
         Alignment       =   2  'Center
         Caption         =   "Head Or Tail"
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
         Left            =   -74160
         TabIndex        =   83
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Labelblabla 
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
         Left            =   -74400
         TabIndex        =   82
         Top             =   1680
         Width           =   2535
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
         Left            =   -74880
         TabIndex        =   81
         Top             =   3720
         Width           =   3615
      End
      Begin VB.Line Line9 
         X1              =   -71055
         X2              =   -71055
         Y1              =   1320
         Y2              =   5400
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   -69000
         X2              =   -69000
         Y1              =   2760
         Y2              =   1200
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   -72960
         X2              =   -69000
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   -72960
         X2              =   -72960
         Y1              =   2760
         Y2              =   1200
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   -72960
         X2              =   -69000
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "VIP Casino"
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
         Left            =   -74880
         TabIndex        =   78
         Top             =   840
         Width           =   7695
      End
      Begin VB.Label lblhl 
         Alignment       =   2  'Center
         Caption         =   "Bet Costs $50"
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
         Left            =   -72720
         TabIndex        =   77
         Top             =   4140
         Width           =   3255
      End
      Begin VB.Label lblnewnumber 
         Caption         =   "Choose First (higher or lower)"
         Height          =   255
         Left            =   -71040
         TabIndex        =   74
         Top             =   1980
         Width           =   2415
      End
      Begin VB.Label Label29 
         Caption         =   "New Number:"
         Height          =   255
         Left            =   -72120
         TabIndex        =   73
         Top             =   1980
         Width           =   1095
      End
      Begin VB.Label lbloldnumber 
         Caption         =   "5"
         Height          =   255
         Left            =   -71040
         TabIndex        =   72
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label Label27 
         Caption         =   "Number        :"
         Height          =   255
         Left            =   -72120
         TabIndex        =   71
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "Casino"
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
         Left            =   -74880
         TabIndex        =   70
         Top             =   480
         Width           =   7575
      End
      Begin VB.Label Label25 
         Caption         =   "Note: if this way dont fit in your country just contact me bij msn or email (Dutchbull@darksoft3d.com)"
         Height          =   615
         Left            =   -72960
         TabIndex        =   69
         Top             =   4680
         Width           =   3975
      End
      Begin VB.Label Label24 
         Caption         =   "Call and follow the instructions in the page you see after you called etc."
         Height          =   495
         Left            =   -72480
         TabIndex        =   68
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "To Start goto this url (click in it)"
         Height          =   375
         Left            =   -74640
         TabIndex        =   67
         Top             =   1320
         Width           =   7215
      End
      Begin VB.Label lblnews 
         Alignment       =   2  'Center
         Caption         =   "Loading News"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   120
         TabIndex        =   63
         Top             =   1380
         Width           =   7455
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "VIP Shop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74880
         TabIndex        =   58
         Top             =   780
         Width           =   7455
      End
      Begin VB.Label Label19 
         Caption         =   "Note: You can only do 1 crime, VIP or normal crimes at the time!"
         Height          =   375
         Left            =   -69720
         TabIndex        =   57
         Top             =   4920
         Width           =   2415
      End
      Begin VB.Label lbltime 
         Caption         =   "0"
         Height          =   255
         Left            =   -72720
         TabIndex        =   56
         Top             =   5100
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "Minutes To Wait Till Next Crime:"
         Height          =   255
         Left            =   -75000
         TabIndex        =   55
         Top             =   5100
         Width           =   2415
      End
      Begin VB.Label lblkans2 
         Alignment       =   2  'Center
         Caption         =   "Choose what you want to do."
         Height          =   375
         Left            =   -72840
         TabIndex        =   54
         Top             =   4560
         Width           =   3735
      End
      Begin VB.Label Label18 
         Caption         =   "VIP Crimes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -72360
         TabIndex        =   50
         Top             =   780
         Width           =   2535
      End
      Begin VB.Label lblwhat 
         Alignment       =   2  'Center
         Caption         =   "Member"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   48
         Top             =   1500
         Width           =   7575
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Shop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74880
         TabIndex        =   39
         Top             =   780
         Width           =   7575
      End
      Begin VB.Label lbltijd 
         Caption         =   "0"
         Height          =   255
         Left            =   -72720
         TabIndex        =   38
         Top             =   5100
         Width           =   255
      End
      Begin VB.Label Label13 
         Caption         =   "Minutes To Wait Till Next Crime"
         Height          =   375
         Left            =   -75000
         TabIndex        =   37
         Top             =   5100
         Width           =   2295
      End
      Begin VB.Label lblkans 
         Alignment       =   2  'Center
         Caption         =   "Choose what you want to do."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -73800
         TabIndex        =   36
         Top             =   4560
         Width           =   6135
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Crimes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74880
         TabIndex        =   32
         Top             =   780
         Width           =   7455
      End
      Begin VB.Label lblrank 
         Caption         =   "Loading Rank"
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
         Left            =   -70680
         TabIndex        =   30
         Top             =   3600
         Width           =   3135
      End
      Begin VB.Label lblra 
         Caption         =   "Rank        :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72000
         TabIndex        =   29
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label lblpower 
         Caption         =   "0"
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
         Left            =   -70680
         TabIndex        =   28
         Top             =   3240
         Width           =   3495
      End
      Begin VB.Label Label11 
         Caption         =   "Power      :"
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
         Left            =   -72000
         TabIndex        =   27
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblras 
         Caption         =   "ras"
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
         Left            =   -70680
         TabIndex        =   26
         Top             =   2880
         Width           =   3015
      End
      Begin VB.Label Label10 
         Caption         =   "Race        :"
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
         Left            =   -72000
         TabIndex        =   25
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Money      :"
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
         Left            =   -72000
         TabIndex        =   23
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblmoney 
         Caption         =   "0"
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
         Left            =   -70680
         TabIndex        =   22
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Label lblnick 
         Alignment       =   2  'Center
         Caption         =   "NickName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74880
         TabIndex        =   20
         Top             =   1920
         Width           =   7575
      End
   End
   Begin VB.Line Line4 
      Index           =   1
      X1              =   2040
      X2              =   2040
      Y1              =   5160
      Y2              =   3600
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   0
      X2              =   3960
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   0
      Picture         =   "Form1.frx":D007
      Top             =   0
      Width           =   7965
   End
   Begin VB.Menu SAE 
      Caption         =   "Save And Exit"
      Visible         =   0   'False
   End
   Begin VB.Menu music 
      Caption         =   "Music Player"
      Visible         =   0   'False
   End
   Begin VB.Menu adminpanel 
      Caption         =   "Admin Panel"
      Visible         =   0   'False
   End
   Begin VB.Menu Links 
      Caption         =   "Links"
      Visible         =   0   'False
   End
   Begin VB.Menu CashGameButton 
      Caption         =   "CashGame"
      Visible         =   0   'False
   End
   Begin VB.Menu about 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub about_Click()
frmAbout.Visible = True
End Sub

Private Sub adminpanel_Click()
Form2.Visible = True
End Sub


Private Sub CashGameButton_Click()
Form5.Show
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim RAS As String

If Option1.Value = True Then
RAS = "Vampire"
GoTo E
End If

If Option2.Value = True Then
RAS = "Lycan"
GoTo E
End If

If Option1.Value = False And Option2.Value = False Then
MsgBox "Please choose what you want to be."
GoTo a
End If

E:
If Text3.Text = Text9.Text Then
MsgBox "You cant put your self in your refferal"
Else
If Text4.Text = Text5.Text Then
Winsock1.SendData "REG " & Text3.Text & " " & Text4.Text & " " & RAS & " " & Text9.Text
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text9.Text = ""
Else
lblstatuss.Caption = "Passwords Dont Match"
End If
End If
a:
End Sub

Private Sub Command10_Click()
On Error Resume Next
If lblmoney.Caption <= 1999 Then
MsgBox "Not enough money."
Else
lblmoney.Caption = lblmoney.Caption - 2000
lblpower.Caption = lblpower.Caption + 2000
MsgBox "Item Bought"
End If
End Sub

Private Sub Command11_Click()
On Error Resume Next
If lblmoney.Caption <= 3999 Then
MsgBox "Not enough money."
Else
lblmoney.Caption = lblmoney.Caption - 4000
lblpower.Caption = lblpower.Caption + 4000
MsgBox "Item Bought"
End If
End Sub

Private Sub Command12_Click()
On Error Resume Next
If lblmoney.Caption <= 7999 Then
MsgBox "Not enough money."
Else
lblmoney.Caption = lblmoney.Caption - 8000
lblpower.Caption = lblpower.Caption + 8000
MsgBox "Item Bought"
End If
End Sub

Private Sub Command13_Click()
On Error Resume Next
If Text7.Text = "" Then
MsgBox "no empty lines please"
Else
Winsock1.SendData "CHAT " & "[" & lblwhat.Caption & "]" & lblnick & " : " & Text7.Text
Text7.Text = ""
End If
End Sub

Private Sub Command14_Click()
Dim Newnumber As String
Dim oldnumber As String
If lblmoney.Caption <= 50 Then
MsgBox "Not enough money,You need $50."
Else
lblmoney.Caption = lblmoney.Caption - 50
Newnumber = Int(Rnd * 50)
oldnumber = Int(Rnd * 50)
lblnewnumber.Caption = Newnumber

If Newnumber >= lbloldnumber.Caption Then
lblhl.Caption = "You Won"
lblmoney.Caption = lblmoney.Caption + 100
Else
lblhl.Caption = "You Lost"
End If
lbloldnumber.Caption = oldnumber
End If

Timer5.Enabled = True
Command14.Enabled = False
Command15.Enabled = False
End Sub

Private Sub Command15_Click()
Dim Newnumber2 As String
Dim oldnumber2 As String
If lblmoney.Caption <= 50 Then
MsgBox "Not enough money,You need $50."
Else
lblmoney.Caption = lblmoney.Caption - 50
Newnumber2 = Int(Rnd * 50)
oldnumber2 = Int(Rnd * 50)
lblnewnumber.Caption = Newnumber2
If Newnumber2 <= lbloldnumber.Caption Then
lblhl.Caption = "You Won"
lblmoney.Caption = lblmoney.Caption + 100
Else
lblhl.Caption = "You Lost"
End If
lbloldnumber.Caption = oldnumber2
End If

Timer5.Enabled = True
Command14.Enabled = False
Command15.Enabled = False
End Sub

Private Sub Command16_Click()
Text6.Text = "Welcome to the Night Walkers Chat Room" & vbNewLine
Text6.Enabled = True
Text7.Enabled = True
Command13.Enabled = True
Command16.Visible = False
Command16.Enabled = False
Winsock1.SendData "CHAT " & "[" & lblwhat.Caption & "]" & lblnick.Caption & " Entered The ChatRoom"
End Sub

Private Sub Command2_Click()
On Error Resume Next
If Text2.Text = " " Then
Text2.Text = "nahnah"
End If
If Text2.Text = "" Then
Text2.Text = "nahnah"
End If
If Text1.Text = "" Then
Text1.Text = "Name Please"
End If
If Text1.Text = " " Then
Text1.Text = "Name Please"
End If

If Check1.Value = "1" Then
SetINI "User", "Username", Text1.Text
SetINI "User", "password", Text2.Text
End If

If Check1.Value = "0" Then
SetINI "User", "Username", ""
SetINI "User", "password", ""
End If

Winsock1.SendData "LGN " & Text1.Text & " " & Text2.Text
lblnick.Caption = Text1.Text
Text1.Text = ""
Text2.Text = ""
End Sub


Private Sub Command3_Click()
On Error Resume Next
Dim kans As Long
Dim Geld As Long
If Text8.Text = lblcode1.Caption Then
If lbltijd.Caption = 0 Then
If crime1.Value = False And crime2.Value = False And crime3.Value = False Then
MsgBox "Please choose an option!"
Else

Geld = Int(Rnd * 500)
kans = Int(Rnd * 100)

lbltijd.Caption = "2"
timertijd.Enabled = True
Command3.Enabled = False
Command8.Enabled = False
If kans <= 50 Then
lblkans.Caption = "Cant you even do that? DaMn"
lblcode1.Caption = Int(Rnd * 5000)
End If

If kans >= 50 Then
lblkans.Caption = ("Good job.. you made it! you got $" & Geld)
lblmoney.Caption = (lblmoney.Caption + Geld)
lblpower.Caption = lblpower.Caption + 50
lblcode1.Caption = Int(Rnd * 5000)
End If
End If
Else
MsgBox "dont try to cheat! or risk a ban"
End If
Else
MsgBox "wrong number entered"
End If

End Sub

Private Sub Command4_Click()
On Error Resume Next
If lblmoney.Caption <= 2999 Then
MsgBox "Not enough money."
Else
lblmoney.Caption = lblmoney.Caption - 3000
lblpower.Caption = lblpower.Caption + 300
MsgBox "Item Bought"
End If
End Sub

Private Sub Command5_Click()
On Error Resume Next
If lblmoney.Caption <= 3999 Then
MsgBox "Not enough money."
Else
lblmoney.Caption = lblmoney.Caption - 4000
lblpower.Caption = lblpower.Caption + 400
MsgBox "Item Bought"
End If
End Sub

Private Sub Command6_Click()
On Error Resume Next
If lblmoney.Caption <= 4999 Then
MsgBox "Not enough money."
Else
lblmoney.Caption = lblmoney.Caption - 5000
lblpower.Caption = lblpower.Caption + 500
MsgBox "Item Bought"
End If
End Sub

Private Sub Command7_Click()
On Error Resume Next
If lblmoney.Caption <= 49999 Then
MsgBox "Not enough money."
Else
lblmoney.Caption = lblmoney.Caption - 50000
lblpower.Caption = lblpower.Caption + 5000
MsgBox "Item Bought"
End If
End Sub



Private Sub Command8_Click()
On Error Resume Next
Dim kans2 As Long
Dim Geld As Long
If Text10.Text = lblcode2.Caption Then
If lbltime.Caption = 0 Then
If Option3.Value = False And Option4.Value = False And Option5.Value = False Then
MsgBox "Please choose an option!"
Else

Geld = Int(Rnd * 1000)
kans2 = Int(Rnd * 100)

lbltime.Caption = "2"
Timertijd2.Enabled = True
Command8.Enabled = False
Command3.Enabled = False
If kans2 <= 50 Then
lblkans2.Caption = "Cant you even do that? DaMn"
lblcode2.Caption = Int(Rnd * 5000)
End If

If kans2 >= 50 Then
lblkans2.Caption = ("Good job.. you made it! you got $" & Geld)
lblmoney.Caption = (lblmoney.Caption + Geld)
lblpower.Caption = lblpower.Caption + 50
lblcode2.Caption = Int(Rnd * 5000)
End If
End If
Else
MsgBox "dont try to cheat! or risk a ban"
End If
Else
MsgBox "wrong number entered"
End If
End Sub

Private Sub Command9_Click()
On Error Resume Next
If lblmoney.Caption <= 999 Then
MsgBox "Not enough money."
Else
lblmoney.Caption = lblmoney.Caption - 1000
lblpower.Caption = lblpower.Caption + 1000
MsgBox "Item Bought"
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
With Winsock1
.Close
.RemoteHost = "127.0.0.1"
.RemotePort = "6666"
.Connect
End With

With SSTab1
.TabEnabled(4) = False
.TabEnabled(5) = False
.TabEnabled(9) = False
End With

music.Visible = False
'newstimer.Enabled = True
CashGameButton.Visible = False
lblversionn.Caption = App.Major & "." & App.Minor
Text1.Text = GetINI("User", "Username")
Text2.Text = GetINI("User", "Password")
lblcode1.Caption = Int(Rnd * 5000)
lblcode2.Caption = Int(Rnd * 5000)
End Sub



Private Sub Form_Unload(Cancel As Integer)
End
End Sub



Private Sub Image10_Click()

End Sub

Private Sub Image4_Click()

End Sub

Private Sub Image2_Click()

End Sub

Private Sub Label33_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Dim link
    link = ShellExecute(hWnd, "Open", "http://www.eurobellen.nl/bel/?pid=12876", &O0, &O0, SW_NORMAL)
End If
End Sub

Private Sub Links_Click()
Form4.Visible = True
End Sub

Private Sub music_Click()
Form3.Visible = True
End Sub

Private Sub newstimer_Timer()
Inet1.OpenURL "http://www.darksoft3d.com/rpg/news.txt"
lblnews.Caption = Inet1.OpenURL
Inet2.OpenURL "http://www.darksoft3d.com/rpg/version.txt"

If lblversionn.Caption + 1 <= Inet2.OpenURL Then
MsgBox "Theres a new version go to www.darksoft3d.com to download Version " & Inet2.OpenURL
End
End If

newstimer.Enabled = False
End Sub



Private Sub Option1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Command1_Click
End If
End Sub

Private Sub Option2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Command1_Click
End If
End Sub

Private Sub SAE_Click()
On Error Resume Next
Winsock1.SendData "SAV" & " " & lblnick.Caption & " " & lblmoney.Caption & " " & lblpower.Caption
Timerexit.Enabled = True
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Command2_Click
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
If Text7.Text = "" Then
MsgBox "no empty lines please"
Else
Winsock1.SendData "CHAT " & "[" & lblwhat & "]" & lblnick & " : " & Text7.Text
Text7.Text = ""
End If
End If
End Sub

Private Sub Text9_Change()
If KeyAscii = vbKeyReturn Then
Command1_Click
End If
End Sub

Private Sub Timer1_Timer()
Text1.Text = Replace(Text1.Text, " ", "-")
Text3.Text = Replace(Text3.Text, " ", "-")
End Sub

Private Sub Timer2_Timer()
On Error Resume Next

If lblpower.Caption <= 5000 Then
lblrank.Caption = "Human"
End If

If lblpower.Caption >= 5000 Then
lblrank.Caption = "Bitten"
End If

If lblpower.Caption >= 10000 Then
lblrank.Caption = "Blood Taster"
End If

If lblpower.Caption >= 150000 Then
lblrank.Caption = "Half Creature"
End If

If lblpower.Caption >= 200000 Then
lblrank.Caption = "Night Walker"
End If

If lblpower.Caption >= 500000 Then
lblrank.Caption = "Day Walker"
End If

If lblpower.Caption >= 1000000 Then
lblrank.Caption = "Ancient One"
End If

End Sub



Private Sub Timer3_Timer()
  If App.PrevInstance = True Then MsgBox "Already running": End
  Timer3.Enabled = False
End Sub



Private Sub Timer5_Timer()
Command15.Enabled = True
Command14.Enabled = True
End Sub

Private Sub Timer6_Timer()
Command111.Enabled = True
Command222.Enabled = True
Timer6.Enabled = False
End Sub

Private Sub Timer7_Timer()
If lblusercount.Caption <= lblonline.Caption Then
lblonline.Caption = lblusercount.Caption
End If
Timer7.Enabled = False
End Sub

Private Sub Timer8_Timer()
If Winsock1.State = 8 Then
MsgBox "Connection Error.. Server offline or crashed (its still beta version so can happen!)"
End If
End Sub

Private Sub Timerexit_Timer()
End
Timerexit.Enabled = False
End Sub



Private Sub timertijd_Timer()
On Error Resume Next
If lbltijd.Caption = "1" Then
lbltijd.Caption = "0"
Command3.Enabled = True
Command8.Enabled = True
timertijd.Enabled = False
lblkans.Caption = "Choose what you want to do."
Alert "You can commit a crime again."
Else
lbltijd.Caption = lbltijd.Caption - 1
End If
End Sub

Private Sub Timertijd2_Timer()
On Error Resume Next
If lbltime.Caption = "1" Then
lbltime.Caption = "0"
Command3.Enabled = True
Command8.Enabled = True
Timertijd2.Enabled = False
lblkans2.Caption = "Choose what you want to do."
Alert "You can commit a crime again."
Else
lbltime.Caption = lbltime.Caption - 1
End If
End Sub

Private Sub txtgetalknop_Click()
RaadGetal
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim sData As String
On Error Resume Next
Winsock1.GetData sData

If sData = "KICKED" Then
MsgBox "You are kicked by admin"
End
End If

If sData = "ACC EXIST" Then
lblstatuss.Caption = "Account allready exists or you allready have an account on your ip."
End If

If sData = "BANNED" Then
lblstatuss.Caption = "You Are Banned.."
End If


If sData = "ACC CREATED" Then
lblstatuss.Caption = "Account Created Ready To Login"
End If

If sData = "NO ACC" Then
MsgBox "No Such Account Admin"
End If

If MidWord(sData, 1, 1) = "ONLINE" Then
lblonline.Caption = MidWord(sData, 2, 1)
lblusercount.Caption = MidWord(sData, 3, 1)
End If

If MidWord(sData, 1, 1) = "STREAMON" Then
Form3.WindowsMediaPlayer1.URL = MidWord(sData, 2, 1)
MsgBox "Music Stream Turned On By Admin Go to the Music Player To Controll It"
End If


If MidWord(sData, 1, 1) = "CHAT" Then
sData = DelWord(sData, 1, 1)
ChatText (sData)
End If

If MidWord(sData, 1, 1) = "OFFERROR" Then
MsgBox "We are back online as soon as possible"
End
End If

If MidWord(sData, 1, 1) = "ALERTMSG" Then
sData = DelWord(sData, 1, 1)
Alert (sData)
End If


If MidWord(sData, 1, 2) = "LOGIN OK" Then
lblstatuss.Caption = "Logged in"
Frame1.Visible = False
Timer1.Enabled = False
lblmoney.Caption = MidWord(sData, 3, 1)
lblras.Caption = MidWord(sData, 4, 1)
lblname.Caption = Text1.Text
lblpower.Caption = MidWord(sData, 5, 1)
lblwhat.Caption = MidWord(sData, 6, 1)
lblreffers.Caption = MidWord(sData, 7, 1)
Timer2.Enabled = True
Timersave.Enabled = True
music.Visible = True
SAE.Visible = True
Links.Visible = True
CashGameButton.Visible = True
If lblwhat.Caption = "Admin" Then

With SSTab1
.TabEnabled(4) = True
.TabEnabled(5) = True
.TabEnabled(9) = True
End With

adminpanel.Visible = True
End If
End If

If lblwhat.Caption = "GM" Then

With SSTab1
.TabEnabled(4) = True
.TabEnabled(5) = True
.TabEnabled(9) = True
End With

adminpanel.Visible = True
End If


If lblwhat.Caption = "VIP" Then

With SSTab1
.TabEnabled(4) = True
.TabEnabled(5) = True
.TabEnabled(9) = True
End With

End If




If sData = "LOGIN WRONG" Then
lblstatuss.Caption = "Wrong username or password"
End If


End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
lblstatuss.Caption = "Server Offline Try Again Later"
End Sub


Private Sub Alert(Text As String)
On Error Resume Next
    Dim AlertBox As frmAlert
    Set AlertBox = New frmAlert
    AlertBox.DisplayAlert Text, 1000
    Me.SetFocus
End Sub


Public Sub ChatText(Text As String)
    On Error Resume Next
    With Form1
   Text = Replace(Text, "CHAT", "")
        .Text6.Text = .Text6.Text & Text & vbNewLine
        .Text6.SelStart = Len(.Text6) - 2
    End With
    
End Sub

Private Sub Command111_Click()
HOTHead

Timer6.Enabled = True
Command111.Enabled = False
Command222.Enabled = False
End Sub

Private Sub Command222_Click()
HOTTail

Command111.Enabled = False
Command222.Enabled = False
Timer6.Enabled = True
End Sub
