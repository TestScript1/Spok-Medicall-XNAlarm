VERSION 5.00
Begin VB.Form ComSetup 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Port Configuration"
   ClientHeight    =   2235
   ClientLeft      =   2430
   ClientTop       =   2670
   ClientWidth     =   4155
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2235
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox baud 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1920
      TabIndex        =   2
      Text            =   "1200"
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox setup 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1920
      TabIndex        =   3
      Text            =   "N,8,1"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox port 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1920
      TabIndex        =   1
      Text            =   "1"
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Baud:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Setup:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Com Port:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "ComSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim i As Integer

    port.AddItem "1"
    port.AddItem "2"
    port.AddItem "3"
    port.AddItem "4"
    port.AddItem "5"
    port.AddItem "6"
    port.AddItem "7"
    port.AddItem "8"
    port.AddItem "9"
    port.AddItem "10"

    baud.AddItem "300"
    baud.AddItem "1200"
    baud.AddItem "2400"
    baud.AddItem "4800"
    baud.AddItem "9600"

    setup.AddItem "N,8,1"
    setup.AddItem "E,7,1"

    port.Text = Str$(cport(setupPort).port)
    i = InStr(cport(setupPort).setup, ",")
    baud.Text = Mid$(cport(setupPort).setup, 1, i - 1)
    setup.Text = Mid$(cport(setupPort).setup, i + 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next  'Added CKO 8/31/00
  
    Dim Temp As String

    Parameter.RectSetup.Top = ComSetup.Top
    Parameter.RectSetup.Left = ComSetup.Left
    Parameter.RectSetup.Height = ComSetup.Height
    Parameter.RectSetup.Width = ComSetup.Width
    
    cport(setupPort).port = Val(port.Text)
    cport(setupPort).setup = Trim$(baud.Text) + "," + Trim$(setup.Text)
    
    If UCase(Parameter.alarmtype) = "DATAPROBE" Then
        ComMain!ComPort(setupPort).PortOpen = False
        ComMain!ComPort(setupPort).CommPort = cport(setupPort).port
        ComMain!ComPort(setupPort).Settings = cport(setupPort).setup
        ComMain!ComPort(setupPort).PortOpen = True
        Temp = ComMain!ComPort(setupPort).Input
    ElseIf UCase(Parameter.alarmtype) = "SIMPLEX" Then
        frmSimplex!MSComm1(setupPort).PortOpen = False
        frmSimplex!MSComm1(setupPort).CommPort = cport(setupPort).port
        frmSimplex!MSComm1(setupPort).Settings = cport(setupPort).setup
        frmSimplex!MSComm1(setupPort).PortOpen = True
        Temp = frmSimplex!MSComm1(setupPort).Input
    End If
End Sub

