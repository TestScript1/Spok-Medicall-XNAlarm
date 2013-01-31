VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form DpSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dataprobe Setup"
   ClientHeight    =   4005
   ClientLeft      =   990
   ClientTop       =   2535
   ClientWidth     =   10350
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
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4005
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbSendBanner 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9195
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   1200
      Width           =   780
   End
   Begin RichTextLib.RichTextBox txtXnMessage 
      Height          =   615
      Left            =   6810
      TabIndex        =   33
      Top             =   1725
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ScrollBars      =   1
      TextRTF         =   $"DPSETUP.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Probe Ports:"
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   9975
      Begin VB.OptionButton optDP 
         Caption         =   "16"
         Height          =   315
         Index           =   16
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   360
         Width           =   550
      End
      Begin VB.OptionButton optDP 
         Caption         =   "15"
         Height          =   315
         Index           =   15
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   360
         Width           =   550
      End
      Begin VB.OptionButton optDP 
         Caption         =   "14"
         Height          =   315
         Index           =   14
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   360
         Width           =   550
      End
      Begin VB.OptionButton optDP 
         Caption         =   "13"
         Height          =   315
         Index           =   13
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   360
         Width           =   550
      End
      Begin VB.OptionButton optDP 
         Caption         =   "12"
         Height          =   315
         Index           =   12
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   360
         Width           =   550
      End
      Begin VB.OptionButton optDP 
         Caption         =   "11"
         Height          =   315
         Index           =   11
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   360
         Width           =   550
      End
      Begin VB.OptionButton optDP 
         Caption         =   "10"
         Height          =   315
         Index           =   10
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   360
         Width           =   550
      End
      Begin VB.OptionButton optDP 
         Caption         =   "9"
         Height          =   315
         Index           =   9
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Width           =   550
      End
      Begin VB.OptionButton optDP 
         Caption         =   "8"
         Height          =   315
         Index           =   8
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   360
         Width           =   550
      End
      Begin VB.OptionButton optDP 
         Caption         =   "7"
         Height          =   315
         Index           =   7
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   550
      End
      Begin VB.OptionButton optDP 
         Caption         =   "6"
         Height          =   315
         Index           =   6
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   550
      End
      Begin VB.OptionButton optDP 
         Caption         =   "5"
         Height          =   315
         Index           =   5
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Width           =   550
      End
      Begin VB.OptionButton optDP 
         Caption         =   "4"
         Height          =   315
         Index           =   4
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   550
      End
      Begin VB.OptionButton optDP 
         Caption         =   "3"
         Height          =   315
         Index           =   3
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   550
      End
      Begin VB.OptionButton optDP 
         Caption         =   "2"
         Height          =   315
         Index           =   2
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   550
      End
      Begin VB.OptionButton optDP 
         Caption         =   "1"
         Height          =   315
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   360
         Width           =   550
      End
   End
   Begin VB.TextBox txtPageNumeric 
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
      Left            =   6840
      MaxLength       =   20
      TabIndex        =   15
      Text            =   "txtpagenumber"
      Top             =   2520
      Width           =   3135
   End
   Begin VB.ComboBox cmbAction 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6840
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1200
      Width           =   840
   End
   Begin VB.TextBox txtExtension 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   12
      Text            =   "txtExtensi"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.ComboBox cmbClearAlarm 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3480
      Width           =   1335
   End
   Begin VB.ComboBox cmbOutputState 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2880
      Width           =   1695
   End
   Begin VB.ComboBox cmbInputState 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2280
      Width           =   1695
   End
   Begin VB.ComboBox cmbType 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1680
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox txtPageMessage 
      Height          =   615
      Left            =   6825
      TabIndex        =   34
      Top             =   3225
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ScrollBars      =   1
      TextRTF         =   $"DPSETUP.frx":008B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Send ticker"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6930
      TabIndex        =   35
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Page Numeric Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Page Alpha Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Action Reminder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Xn Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Extension"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Clear Alarm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Output State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Input State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1665
      Width           =   855
   End
End
Attribute VB_Name = "DpSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dataProbePort As Integer

Dim thisChannel As Integer
Dim prevChannel As Integer

Function currentChannel() As Integer
Dim i As Integer
For i = 1 To 16
    If optDP(i).Value = True Then
        currentChannel = i
        Exit For
    End If
Next

End Function

Private Sub DpChangePortData(dpData As Integer)

    dataProbePort = dpData
    
    cmbType.ListIndex = dpPort(dataProbePort).Type
    
    cmbInputState.ListIndex = dpPort(dataProbePort).inputState
    cmbOutputState.ListIndex = dpPort(dataProbePort).outputState
    cmbClearAlarm.ListIndex = dpPort(dataProbePort).clearAlarm
    cmbAction.ListIndex = dpPort(dataProbePort).ActionReminder
    
    txtExtension.Text = dpPort(dataProbePort).Extension
    txtXnMessage.Text = dpPort(dataProbePort).xnMessage
    txtPageNumeric.Text = dpPort(dataProbePort).pageNumeric
    txtPageMessage.Text = dpPort(dataProbePort).pageMessage
End Sub

Private Sub dpSetPortData(dpData As Integer)
    dpPort(dataProbePort).Type = cmbType.ListIndex
    dpPort(dataProbePort).inputState = cmbInputState.ListIndex
    dpPort(dataProbePort).outputState = cmbOutputState.ListIndex
    dpPort(dataProbePort).clearAlarm = cmbClearAlarm.ListIndex
    'dpPort(dataProbePort).pageMin = txtPageMinutes.Text
    'dpPort(dataProbePort).pageTimeout = txtPageTimeout.Text
    dpPort(dataProbePort).ActionReminder = cmbAction.ListIndex
    dpPort(dataProbePort).Extension = txtExtension.Text
    dpPort(dataProbePort).xnMessage = txtXnMessage.Text
    dpPort(dataProbePort).pageMessage = txtPageMessage.Text
    dpPort(dataProbePort).pageNumeric = txtPageNumeric.Text
End Sub

Private Sub Form_Load()
    optDP(1).Value = True
    dataProbePort = 0
    
    cmbType.AddItem "Alarm"
    cmbType.AddItem "Page"

    cmbInputState.AddItem "Positive"
    cmbInputState.AddItem "Negative"
    
    cmbOutputState.AddItem "Positive"
    cmbOutputState.AddItem "Negative"
    
    cmbClearAlarm.AddItem "Yes"
    cmbClearAlarm.AddItem "No"

    cmbAction.AddItem "Yes"
    cmbAction.AddItem "No"
    
    cmbSendBanner.AddItem "Yes"
    cmbSendBanner.AddItem "No"
    
    'Call DpChangePortData(dataProbePort)
    'Dim i As Integer
    thisChannel = 1
    DPChannels.ChannelNum = 1
    ShowDPChanInfo thisChannel
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DPChannels.Item(thisChannel).DPType = cmbType.ListIndex
    DPChannels.Item(thisChannel).DPInputState = cmbInputState.ListIndex
    DPChannels.Item(thisChannel).DPOutputState = cmbOutputState.ListIndex
    DPChannels.Item(thisChannel).DPClearAlarm = cmbClearAlarm.ListIndex
    DPChannels.Item(thisChannel).DPActionReminder = cmbAction.ListIndex
    DPChannels.Item(thisChannel).DPSendBanner = IIf(cmbSendBanner.Text = "Yes", True, False)
   Call SaveDPChannels
End Sub



Private Sub Label6_Click()

End Sub

Private Sub optDP_Click(index As Integer)

If Me.Visible Then
    
    prevChannel = DPChannels.ChannelNum
    thisChannel = index
    ' remember previous data
    DPChannels.Item(prevChannel).DPType = cmbType.ListIndex
    DPChannels.Item(prevChannel).DPInputState = cmbInputState.ListIndex
    DPChannels.Item(prevChannel).DPOutputState = cmbOutputState.ListIndex
    DPChannels.Item(prevChannel).DPClearAlarm = cmbClearAlarm.ListIndex
    DPChannels.Item(prevChannel).DPActionReminder = cmbAction.ListIndex
    DPChannels.Item(prevChannel).DPSendBanner = IIf(cmbSendBanner.Text = "Yes", True, False)
    'Show new data
    ShowDPChanInfo index
    DPChannels.ChannelNum = thisChannel
    txtExtension.SetFocus
End If
End Sub






Private Sub txtExtension_Change()

DPChannels.Item(thisChannel).DPExtension = txtExtension

End Sub




Private Sub txtPageMessage_Change()
DPChannels.Item(thisChannel).DPPageMessage = txtPageMessage.Text

End Sub

Private Sub txtPageNumeric_Change()
DPChannels.Item(thisChannel).DPPageNumeric = txtPageNumeric

End Sub

Private Sub txtPageNumeric_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 59 Then
    KeyAscii = 0
End If
End Sub


Private Sub txtXnMessage_Change()
DPChannels.Item(thisChannel).DPxnMessage = txtXnMessage.Text
End Sub

