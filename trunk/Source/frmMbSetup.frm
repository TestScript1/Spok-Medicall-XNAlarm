VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMbSetup 
   Caption         =   "Xnalarm Metrabyte Setup"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   5640
      Width           =   975
   End
   Begin VB.Frame framePortInfo 
      Caption         =   "Frame1"
      Height          =   4095
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   7815
      Begin VB.TextBox txtSupMsgOff 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   39
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox txtSupMsgOn 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   38
         Top             =   3228
         Width           =   2055
      End
      Begin VB.TextBox txtPageTimeout 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   37
         Top             =   1074
         Width           =   855
      End
      Begin VB.TextBox txtPageMinutes 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   36
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtAlertWidth 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   35
         Top             =   1752
         Width           =   855
      End
      Begin VB.TextBox txtOutputPort 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   34
         Top             =   1413
         Width           =   855
      End
      Begin VB.TextBox txtPageRetry 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   33
         Top             =   2121
         Width           =   855
      End
      Begin VB.TextBox txtSource 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   32
         Top             =   2490
         Width           =   2055
      End
      Begin VB.TextBox txtDestination 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   31
         Top             =   2859
         Width           =   2055
      End
      Begin VB.TextBox txtExtension 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5640
         MaxLength       =   10
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtXnMessage 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4800
         MaxLength       =   40
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "frmMbSetup.frx":0000
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtPageMessage 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4800
         MaxLength       =   40
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   2520
         Width           =   2895
      End
      Begin VB.ComboBox cmbAction 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtPageNumeric 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1920
         Width           =   2895
      End
      Begin VB.ComboBox cmbUsage 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbState 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Sup Msg Off"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Sup Msg On"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   3345
         Width           =   1215
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         Caption         =   "Page Timeout"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1095
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Page Interval"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Output Port"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1470
         Width           =   1215
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Alert With"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1845
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Retry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   2220
         Width           =   855
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Sorce"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2595
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Destination"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2970
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Extension"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Xn Message"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3480
         TabIndex        =   20
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Action Reminder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3000
         TabIndex        =   19
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Page Alpha Message"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3600
         TabIndex        =   18
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Page Numeric Message"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3360
         TabIndex        =   17
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Usage"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tabs            =   25
      TabsPerRow      =   9
      TabHeight       =   582
      TabCaption(0)   =   "Board 1 Options"
      TabPicture(0)   =   "frmMbSetup.frx":0002
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmbControlWord"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtBaseAddress"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Port 1"
      TabPicture(1)   =   "frmMbSetup.frx":001E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Port 2"
      TabPicture(2)   =   "frmMbSetup.frx":003A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Port 3"
      TabPicture(3)   =   "frmMbSetup.frx":0056
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Port 4"
      TabPicture(4)   =   "frmMbSetup.frx":0072
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Port 5"
      TabPicture(5)   =   "frmMbSetup.frx":008E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Port 6"
      TabPicture(6)   =   "frmMbSetup.frx":00AA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Port 7"
      TabPicture(7)   =   "frmMbSetup.frx":00C6
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "Port 8"
      TabPicture(8)   =   "frmMbSetup.frx":00E2
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
      TabCaption(9)   =   "Port 9"
      TabPicture(9)   =   "frmMbSetup.frx":00FE
      Tab(9).ControlEnabled=   0   'False
      Tab(9).ControlCount=   0
      TabCaption(10)  =   "Port 10"
      TabPicture(10)  =   "frmMbSetup.frx":011A
      Tab(10).ControlEnabled=   0   'False
      Tab(10).ControlCount=   0
      TabCaption(11)  =   "Port 11"
      TabPicture(11)  =   "frmMbSetup.frx":0136
      Tab(11).ControlEnabled=   0   'False
      Tab(11).ControlCount=   0
      TabCaption(12)  =   "Port 12"
      TabPicture(12)  =   "frmMbSetup.frx":0152
      Tab(12).ControlEnabled=   0   'False
      Tab(12).ControlCount=   0
      TabCaption(13)  =   "Port 13"
      TabPicture(13)  =   "frmMbSetup.frx":016E
      Tab(13).ControlEnabled=   0   'False
      Tab(13).ControlCount=   0
      TabCaption(14)  =   "Port 14"
      TabPicture(14)  =   "frmMbSetup.frx":018A
      Tab(14).ControlEnabled=   0   'False
      Tab(14).ControlCount=   0
      TabCaption(15)  =   "Port 15"
      TabPicture(15)  =   "frmMbSetup.frx":01A6
      Tab(15).ControlEnabled=   0   'False
      Tab(15).ControlCount=   0
      TabCaption(16)  =   "Port 16"
      TabPicture(16)  =   "frmMbSetup.frx":01C2
      Tab(16).ControlEnabled=   0   'False
      Tab(16).ControlCount=   0
      TabCaption(17)  =   "Port 17"
      TabPicture(17)  =   "frmMbSetup.frx":01DE
      Tab(17).ControlEnabled=   0   'False
      Tab(17).ControlCount=   0
      TabCaption(18)  =   "Port 18"
      TabPicture(18)  =   "frmMbSetup.frx":01FA
      Tab(18).ControlEnabled=   0   'False
      Tab(18).ControlCount=   0
      TabCaption(19)  =   "Port 19"
      TabPicture(19)  =   "frmMbSetup.frx":0216
      Tab(19).ControlEnabled=   0   'False
      Tab(19).ControlCount=   0
      TabCaption(20)  =   "Port 20"
      TabPicture(20)  =   "frmMbSetup.frx":0232
      Tab(20).ControlEnabled=   0   'False
      Tab(20).ControlCount=   0
      TabCaption(21)  =   "Port 21"
      TabPicture(21)  =   "frmMbSetup.frx":024E
      Tab(21).ControlEnabled=   0   'False
      Tab(21).ControlCount=   0
      TabCaption(22)  =   "Port 22"
      TabPicture(22)  =   "frmMbSetup.frx":026A
      Tab(22).ControlEnabled=   0   'False
      Tab(22).ControlCount=   0
      TabCaption(23)  =   "Port 23"
      TabPicture(23)  =   "frmMbSetup.frx":0286
      Tab(23).ControlEnabled=   0   'False
      Tab(23).ControlCount=   0
      TabCaption(24)  =   "Port 24"
      TabPicture(24)  =   "frmMbSetup.frx":02A2
      Tab(24).ControlEnabled=   0   'False
      Tab(24).ControlCount=   0
      Begin VB.TextBox txtBaseAddress 
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Text            =   "Hex"
         Top             =   1920
         Width           =   855
      End
      Begin VB.ComboBox cmbControlWord 
         Height          =   315
         ItemData        =   "frmMbSetup.frx":02BE
         Left            =   2400
         List            =   "frmMbSetup.frx":02F2
         TabIndex        =   3
         Text            =   "Control Word"
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Control Word"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Base Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   1
         Top             =   1920
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmMbSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mbControl(15) As metrabyteControlType
Private firstTime As Boolean

Sub SaveMetrabyteParams()
    Dim i As Integer
    Dim j As Integer
    Dim Temp As String
    Dim temp2 As String
    
    metrabyte(metrabyteBoardNumber).address = Val("&H" + Trim$(txtBaseAddress.Text))
    Temp = "Board_" + Trim$(Str$(metrabyteBoardNumber)) + "_Address"
    temp2 = "&H" + Trim$(txtBaseAddress.Text)
    j = WriteIniString("METRABYTE", Temp, temp2, Parameter.pPath)
    
    Temp = "Board_" + Trim$(Str$(metrabyteBoardNumber)) + "_ControlWord"
    temp2 = Str$(metrabyte(metrabyteBoardNumber).controlByte)
    j = WriteIniString("METRABYTE", Temp, temp2, Parameter.pPath)
    
    For i = 0 To 23
        Temp = "Board_" + Trim$(Str$(metrabyteBoardNumber)) + "_Port_" + Trim$(Str$(i + 1)) + "_Type"
        temp2 = Str$(metrabyte(metrabyteBoardNumber).port(i).Type)
        j = WriteIniString("METRABYTE", Temp, temp2, Parameter.pPath)
    
        Temp = "Board_" + Trim$(Str$(metrabyteBoardNumber)) + "_Port_" + Trim$(Str$(i + 1)) + "_Usage"
        temp2 = Str$(metrabyte(metrabyteBoardNumber).port(i).usage)
        j = WriteIniString("METRABYTE", Temp, temp2, Parameter.pPath)
    
        Temp = "Board_" + Trim$(Str$(metrabyteBoardNumber)) + "_Port_" + Trim$(Str$(i + 1)) + "_State"
        temp2 = Str$(metrabyte(metrabyteBoardNumber).port(i).state)
        j = WriteIniString("METRABYTE", Temp, temp2, Parameter.pPath)
    
        Temp = "Board_" + Trim$(Str$(metrabyteBoardNumber)) + "_Port_" + Trim$(Str$(i + 1)) + "_Page_Interval"
        temp2 = Str$(metrabyte(metrabyteBoardNumber).port(i).pageInterval)
        j = WriteIniString("METRABYTE", Temp, temp2, Parameter.pPath)
    
        Temp = "Board_" + Trim$(Str$(metrabyteBoardNumber)) + "_Port_" + Trim$(Str$(i + 1)) + "_Page_Timeout"
        temp2 = Str$(metrabyte(metrabyteBoardNumber).port(i).pageTimeout)
        j = WriteIniString("METRABYTE", Temp, temp2, Parameter.pPath)
    
        Temp = "Board_" + Trim$(Str$(metrabyteBoardNumber)) + "_Port_" + Trim$(Str$(i + 1)) + "_Page_Out_Port"
        temp2 = Str$(metrabyte(metrabyteBoardNumber).port(i).outputPort)
        j = WriteIniString("METRABYTE", Temp, temp2, Parameter.pPath)
    
        Temp = "Board_" + Trim$(Str$(metrabyteBoardNumber)) + "_Port_" + Trim$(Str$(i + 1)) + "_Page_Width"
        temp2 = Str$(metrabyte(metrabyteBoardNumber).port(i).pageAlertWidth)
        j = WriteIniString("METRABYTE", Temp, temp2, Parameter.pPath)
    
        Temp = "Board_" + Trim$(Str$(metrabyteBoardNumber)) + "_Port_" + Trim$(Str$(i + 1)) + "_Page_Retry"
        temp2 = Str$(metrabyte(metrabyteBoardNumber).port(i).pageRetry)
        j = WriteIniString("METRABYTE", Temp, temp2, Parameter.pPath)
        
        Temp = "Board_" + Trim$(Str$(metrabyteBoardNumber)) + "_Port_" + Trim$(Str$(i + 1)) + "_Source"
        temp2 = metrabyte(metrabyteBoardNumber).port(i).sourceFile
        j = WriteIniString("METRABYTE", Temp, temp2, Parameter.pPath)
    
        Temp = "Board_" + Trim$(Str$(metrabyteBoardNumber)) + "_Port_" + Trim$(Str$(i + 1)) + "_Destination"
        temp2 = metrabyte(metrabyteBoardNumber).port(i).destinationFile
        j = WriteIniString("METRABYTE", Temp, temp2, Parameter.pPath)
        
        Temp = "Board_" + Trim$(Str$(metrabyteBoardNumber)) + "_Port_" + Trim$(Str$(i + 1)) + "_Extension"
        temp2 = metrabyte(metrabyteBoardNumber).port(i).Extension
        j = WriteIniString("METRABYTE", Temp, temp2, Parameter.pPath)
    
        Temp = "Board_" + Trim$(Str$(metrabyteBoardNumber)) + "_Port_" + Trim$(Str$(i + 1)) + "_Message"
        temp2 = metrabyte(metrabyteBoardNumber).port(i).xnMessage
        j = WriteIniString("METRABYTE", Temp, temp2, Parameter.pPath)
    
        Temp = "Board_" + Trim$(Str$(metrabyteBoardNumber)) + "_Port_" + Trim$(Str$(i + 1)) + "_ActionReminder"
        temp2 = Str$(metrabyte(metrabyteBoardNumber).port(i).ActionReminder)
        j = WriteIniString("METRABYTE", Temp, temp2, Parameter.pPath)
    
        Temp = "Board_" + Trim$(Str$(metrabyteBoardNumber)) + "_Port_" + Trim$(Str$(i + 1)) + "_Page_Numeric"
        temp2 = metrabyte(metrabyteBoardNumber).port(i).pageNumeric
        j = WriteIniString("METRABYTE", Temp, temp2, Parameter.pPath)
    
        Temp = "Board_" + Trim$(Str$(metrabyteBoardNumber)) + "_Port_" + Trim$(Str$(i + 1)) + "_Page_Alpha"
        temp2 = metrabyte(metrabyteBoardNumber).port(i).pageAlpha
        j = WriteIniString("METRABYTE", Temp, temp2, Parameter.pPath)
        
        Temp = "Board_" + Trim$(Str$(metrabyteBoardNumber)) + "_Port_" + Trim$(Str$(i + 1)) + "_Sup_Msg_On"
        temp2 = metrabyte(metrabyteBoardNumber).port(i).supMsgOn
        j = WriteIniString("METRABYTE", Temp, temp2, Parameter.pPath)
    
        Temp = "Board_" + Trim$(Str$(metrabyteBoardNumber)) + "_Port_" + Trim$(Str$(i + 1)) + "_Sup_Msg_Off"
        temp2 = metrabyte(metrabyteBoardNumber).port(i).supMsgOff
        j = WriteIniString("METRABYTE", Temp, temp2, Parameter.pPath)
    Next
End Sub

Sub SetPortEditOptions(portNumber As Integer)
    If metrabyte(metrabyteBoardNumber).port(portNumber).Type = False Then
        ' Output port
        Label6.Visible = False
        txtPageMinutes.Visible = False
        Label12.Visible = False
        txtPageTimeout.Visible = False
        Label5.Visible = False
        txtOutputPort.Visible = False
        Label13.Visible = False
        txtAlertWidth.Visible = False
        Label14.Visible = False
        txtPageRetry.Visible = False
        Label15.Visible = False
        txtSource.Visible = False
        Label16.Visible = False
        txtDestination.Visible = False
        Label7.Visible = False
        txtExtension.Visible = False
        Label8.Visible = False
        txtXnMessage.Visible = False
        Label9.Visible = False
        cmbAction.Visible = False
        Label11.Visible = False
        txtPageNumeric.Visible = False
        Label10.Visible = False
        txtPageMessage.Visible = False
        txtSupMsgOn.Visible = False
        txtSupMsgOff.Visible = False
    Else
        ' Input port
        If cmbUsage.ListIndex = 1 Then
            ' Input Page port
            Label6.Visible = True
            txtPageMinutes.Visible = True
            Label12.Visible = True
            txtPageTimeout.Visible = True
            Label5.Visible = True
            txtOutputPort.Visible = True
            Label13.Visible = True
            txtAlertWidth.Visible = True
            Label14.Visible = True
            txtPageRetry.Visible = True
            Label15.Visible = True
            txtSource.Visible = True
            Label16.Visible = True
            txtDestination.Visible = True
        Else
            Label6.Visible = False
            txtPageMinutes.Visible = False
            Label12.Visible = False
            txtPageTimeout.Visible = False
            Label5.Visible = False
            txtOutputPort.Visible = False
            Label13.Visible = False
            txtAlertWidth.Visible = False
            Label14.Visible = False
            txtPageRetry.Visible = False
            Label15.Visible = False
            txtSource.Visible = False
            Label16.Visible = False
            txtDestination.Visible = False
        End If
        Label7.Visible = True
        txtExtension.Visible = True
        Label8.Visible = True
        txtXnMessage.Visible = True
        Label9.Visible = True
        cmbAction.Visible = True
        Label11.Visible = True
        txtPageNumeric.Visible = True
        Label10.Visible = True
        txtPageMessage.Visible = True
        txtSupMsgOn.Visible = True
        txtSupMsgOff.Visible = True
    End If
End Sub


Private Sub cmbControlWord_Click()
    Dim i As Integer
    
    metrabyte(metrabyteBoardNumber).controlByte = Val("&H" + cmbControlWord.Text)
    For i = 0 To 15
        SSTab1.Tab = i + 1
        
        metrabyte(metrabyteBoardNumber).port(i).Type = mbControl(cmbControlWord.ListIndex).PType(Int(i / 8))
        If metrabyte(metrabyteBoardNumber).port(i).Type = False Then
            SSTab1.Caption = "Output " & Trim$(Str$(i + 1))
        Else
            SSTab1.Caption = "Input " & Trim$(Str$(i + 1))
        End If
    Next
    For i = 16 To 23
        SSTab1.Tab = i + 1
        metrabyte(metrabyteBoardNumber).port(i).Type = mbControl(cmbControlWord.ListIndex).PType(Int(i / 4) - 2)
        If metrabyte(metrabyteBoardNumber).port(i).Type = False Then
            SSTab1.Caption = "Output " & Trim$(Str$(i + 1))
        Else
            SSTab1.Caption = "Input " & Trim$(Str$(i + 1))
        End If
    Next
    SSTab1.Tab = 0

End Sub

Private Sub cmbUsage_Click()
    metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).usage = cmbUsage.ListIndex
    Call SetPortEditOptions(SSTab1.Tab - 1)
End Sub

Private Sub cmdCancel_Click()
    'Call frmMetrabyte.GetMetrabyteParams
    Unload frmMbSetup
    
End Sub

Private Sub cmdOk_Click()
    Dim j As Integer
    
    If SSTab1.Tab <> 0 Then
        metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).usage = cmbUsage.ListIndex
        metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).state = cmbState.ListIndex
        metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).pageInterval = Val(txtPageMinutes.Text)
        metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).pageTimeout = Val(txtPageTimeout.Text)
        metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).outputPort = Val(txtOutputPort.Text)
        metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).pageAlertWidth = Val(txtAlertWidth.Text)
        metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).pageRetry = Val(txtPageRetry.Text)
        metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).sourceFile = txtSource.Text
        metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).destinationFile = txtDestination.Text
        metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).Extension = txtExtension.Text
        metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).xnMessage = txtXnMessage.Text
        metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).ActionReminder = cmbAction.ListIndex
        metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).pageNumeric = txtPageNumeric.Text
        metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).pageAlpha = txtPageMessage.Text
        metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).supMsgOn = txtSupMsgOn.Text
        metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).supMsgOff = txtSupMsgOff.Text
    End If
    Call SaveMetrabyteParams
            
    For j = 0 To 23
        If metrabyte(metrabyteBoardNumber).port(j).Type = False Then
            If metrabyte(metrabyteBoardNumber).port(j).usage = 0 Then
                Call frmMetrabyte.MetrabyteSetPort(False, metrabyteBoardNumber, j)
            Else
                Call frmMetrabyte.MetrabyteSetPort(True, metrabyteBoardNumber, j)
            End If
        End If
    Next
    
    Unload frmMbSetup
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    
    cmbUsage.AddItem "Alarm"
    cmbUsage.AddItem "Page"

    cmbState.AddItem "Positive"
    cmbState.AddItem "Negative"
    
    cmbAction.AddItem "Yes"
    cmbAction.AddItem "No"
    
    mbControl(0).Control = &H80
    mbControl(0).PType(0) = False
    mbControl(0).PType(1) = False
    mbControl(0).PType(2) = False
    mbControl(0).PType(3) = False
    mbControl(1).Control = &H81
    mbControl(1).PType(0) = False
    mbControl(1).PType(1) = False
    mbControl(1).PType(2) = True
    mbControl(1).PType(3) = False
    mbControl(2).Control = &H82
    mbControl(2).PType(0) = False
    mbControl(2).PType(1) = True
    mbControl(2).PType(2) = False
    mbControl(2).PType(3) = False
    mbControl(3).Control = &H83
    mbControl(3).PType(0) = False
    mbControl(3).PType(1) = True
    mbControl(3).PType(2) = True
    mbControl(3).PType(3) = False
    mbControl(4).Control = &H88
    mbControl(4).PType(0) = False
    mbControl(4).PType(1) = False
    mbControl(4).PType(2) = False
    mbControl(4).PType(3) = True
    mbControl(5).Control = &H89
    mbControl(5).PType(0) = False
    mbControl(5).PType(1) = False
    mbControl(5).PType(2) = True
    mbControl(5).PType(3) = True
    mbControl(6).Control = &H8A
    mbControl(6).PType(0) = False
    mbControl(6).PType(1) = True
    mbControl(6).PType(2) = False
    mbControl(6).PType(3) = True
    mbControl(7).Control = &H8B
    mbControl(7).PType(0) = False
    mbControl(7).PType(1) = True
    mbControl(7).PType(2) = True
    mbControl(7).PType(3) = True
    mbControl(8).Control = &H90
    mbControl(8).PType(0) = True
    mbControl(8).PType(1) = False
    mbControl(8).PType(2) = False
    mbControl(8).PType(3) = False
    mbControl(9).Control = &H91
    mbControl(9).PType(0) = True
    mbControl(9).PType(1) = False
    mbControl(9).PType(2) = True
    mbControl(9).PType(3) = False
    mbControl(10).Control = &H92
    mbControl(10).PType(0) = True
    mbControl(10).PType(1) = True
    mbControl(10).PType(2) = False
    mbControl(10).PType(3) = False
    mbControl(11).Control = &H93
    mbControl(11).PType(0) = True
    mbControl(11).PType(1) = True
    mbControl(11).PType(2) = True
    mbControl(11).PType(3) = False
    mbControl(12).Control = &H98
    mbControl(12).PType(0) = True
    mbControl(12).PType(1) = False
    mbControl(12).PType(2) = False
    mbControl(12).PType(3) = True
    mbControl(13).Control = &H99
    mbControl(13).PType(0) = True
    mbControl(13).PType(1) = False
    mbControl(13).PType(2) = True
    mbControl(13).PType(3) = True
    mbControl(14).Control = &H9A
    mbControl(14).PType(0) = True
    mbControl(14).PType(1) = True
    mbControl(14).PType(2) = False
    mbControl(14).PType(3) = True
    mbControl(15).Control = &H9B
    mbControl(15).PType(0) = True
    mbControl(15).PType(1) = True
    mbControl(15).PType(2) = True
    mbControl(15).PType(3) = True
        
    SSTab1.Tab = 0
    SSTab1.Caption = "Board " + Trim$(Str$(metrabyteBoardNumber + 1)) + " Options"
    txtBaseAddress.Text = Trim$(Hex$(metrabyte(metrabyteBoardNumber).address))
    For i = 0 To 15
        If metrabyte(metrabyteBoardNumber).controlByte = mbControl(i).Control Then
            cmbControlWord.ListIndex = i
        End If
    Next
    For i = 0 To 23
        SSTab1.Tab = i + 1
        If metrabyte(metrabyteBoardNumber).port(i).Type = False Then
            SSTab1.Caption = "Output " & Trim$(Str$(i + 1))
        Else
            SSTab1.Caption = "Input " & Trim$(Str$(i + 1))
        End If
    Next
    framePortInfo.ZOrder 1
    SSTab1.Tab = 0
    firstTime = True
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
    If PreviousTab <> 0 And firstTime = False Then
        metrabyte(metrabyteBoardNumber).port(PreviousTab - 1).usage = cmbUsage.ListIndex
        metrabyte(metrabyteBoardNumber).port(PreviousTab - 1).state = cmbState.ListIndex
        metrabyte(metrabyteBoardNumber).port(PreviousTab - 1).pageInterval = Val(txtPageMinutes.Text)
        metrabyte(metrabyteBoardNumber).port(PreviousTab - 1).pageTimeout = Val(txtPageTimeout.Text)
        metrabyte(metrabyteBoardNumber).port(PreviousTab - 1).outputPort = Val(txtOutputPort.Text)
        metrabyte(metrabyteBoardNumber).port(PreviousTab - 1).pageAlertWidth = Val(txtAlertWidth.Text)
        metrabyte(metrabyteBoardNumber).port(PreviousTab - 1).pageRetry = Val(txtPageRetry.Text)
        metrabyte(metrabyteBoardNumber).port(PreviousTab - 1).sourceFile = txtSource.Text
        metrabyte(metrabyteBoardNumber).port(PreviousTab - 1).destinationFile = txtDestination.Text
        metrabyte(metrabyteBoardNumber).port(PreviousTab - 1).Extension = txtExtension.Text
        metrabyte(metrabyteBoardNumber).port(PreviousTab - 1).xnMessage = txtXnMessage.Text
        metrabyte(metrabyteBoardNumber).port(PreviousTab - 1).ActionReminder = cmbAction.ListIndex
        metrabyte(metrabyteBoardNumber).port(PreviousTab - 1).pageNumeric = txtPageNumeric.Text
        metrabyte(metrabyteBoardNumber).port(PreviousTab - 1).pageAlpha = txtPageMessage.Text
        metrabyte(metrabyteBoardNumber).port(PreviousTab - 1).supMsgOn = txtSupMsgOn.Text
        metrabyte(metrabyteBoardNumber).port(PreviousTab - 1).supMsgOff = txtSupMsgOff.Text
    End If
    If SSTab1.Tab <> 0 Then
        cmbUsage.ListIndex = metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).usage
        cmbState.ListIndex = metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).state
        txtPageMinutes.Text = Trim$(Str$(metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).pageInterval))
        txtPageTimeout.Text = Trim$(Str$(metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).pageTimeout))
        txtOutputPort.Text = Trim$(Str$(metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).outputPort))
        txtAlertWidth.Text = Trim$(Str$(metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).pageAlertWidth))
        txtPageRetry.Text = Trim$(Str$(metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).pageRetry))
        txtSource.Text = metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).sourceFile
        txtDestination.Text = metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).destinationFile
        txtExtension.Text = metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).Extension
        txtXnMessage.Text = metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).xnMessage
        cmbAction.ListIndex = metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).ActionReminder
        txtPageNumeric.Text = metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).pageNumeric
        txtPageMessage.Text = metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).pageAlpha
        txtSupMsgOn.Text = metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).supMsgOn
        txtSupMsgOff.Text = metrabyte(metrabyteBoardNumber).port(SSTab1.Tab - 1).supMsgOff
        Call SetPortEditOptions(SSTab1.Tab - 1)
    End If
    If SSTab1.Tab <> 0 Then
        framePortInfo.ZOrder 0
    Else
        framePortInfo.ZOrder 1
    End If
    firstTime = False
End Sub

