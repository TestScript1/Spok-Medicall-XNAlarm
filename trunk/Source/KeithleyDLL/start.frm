VERSION 5.00
Object = "{4DE9E2A3-150F-11CF-8FBF-444553540000}#4.0#0"; "DlxOCX32.ocx"
Begin VB.Form frmDriverLINX 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin DlsrLib.DriverLINXSR SR3 
      Left            =   1080
      Top             =   2400
      _Version        =   262144
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   64
   End
   Begin DlsrLib.DriverLINXSR SR2 
      Left            =   2760
      Top             =   1320
      _Version        =   262144
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   64
   End
   Begin DlsrLib.DriverLINXLDD LDD1 
      Left            =   2760
      Top             =   960
      _Version        =   262144
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   64
      _Version        =   262144
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   64
   End
   Begin DlsrLib.DriverLINXSR SR1 
      Left            =   2760
      Top             =   480
      _Version        =   262144
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   64
   End
   Begin VB.Label lblModelName 
      Caption         =   "ModelName"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "frmDriverLINX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub ConfigureChannel(ByVal channel As Long, ByVal parMode As String)


With SR2
    If parMode = "out" Then
    .Req_subsystem = DL_DO
    Else  ' in
    .Req_subsystem = DL_DI
    End If
    .Evt_Tim_dioChannel = channel
    .Refresh
End With

End Sub

Public Sub DIConfig(ByVal channel As Long)
Dim resultCode As Integer

resultCode = InitializeDriverLINXDevice(SR1, 0)
If resultCode <> DL_NoErr Then

    LogMessage ErrorMessage, "Can't configure the board, result code=" & CStr(resultCode)
End If
Call InitializeDriverLINXDevice(SR2, 0)

lblModelName.Caption = GetModelName(SR1, LDD1)

SetupDriverLINXSingleValueIO SR1, LDD1, 0, DL_DI, 0, 0, 0
SetupDriverLINXInitDIOPort SR2, 0, DL_DI, 0



End Sub

Public Function DIRead(ByVal channel As Long) As Long
Dim msgStatus As String
Dim resultCode As Integer

SetupDriverLINXSingleValueIO SR1, LDD1, 0, DL_DI, channel, 0, 0
'SetupDriverLINXInitDIOPort SR2, 0, DL_DI, Channel
With SR1
    .Sel_chan_start = channel
    .Sel_chan_stop = channel
    .Refresh
    resultCode = GetDriverLINXStatus(SR1, msgStatus)
    If resultCode = DL_NoErr Then
        DIRead = GetDriverLINXDISingleValue(SR1)
    Else
        DIRead = -1
    End If

    
End With

End Function
Public Sub DOWrite(ByVal WriteValue As Single, ByVal channel As Long)
Dim retCode As Integer
Dim statusMsg As String

SetupDriverLINXSingleValueIO SR1, LDD1, 0, DL_DO, channel, 0, 0
'SetupDriverLINXInitDIOPort SR2, 0, DL_DO, Channel
With SR1
   
    .Sel_chan_start = channel
    If ISInDriverLINXExtendedDigitalRange(SR1, LDD1, DL_DO, 0, WriteValue) Then
        .Res_Sta_ioValue = WriteValue
    End If
    .Refresh
    If .Res_result <> 0 Then
        statusMsg = .Message
        LogMessage ErrorMessage, "Can't write: " & statusMsg
    End If

End With

End Sub

Public Function GetDriverStatus() As String

Dim DLMessage As String

GetDriverLINXStatus SR1, DLMessage
GetDriverStatus = DLMessage

End Function

Public Function InitializeDevice(Driver As String, device As Long) As String

InitializeDevice = OpenDriverLINXDriver(SR1, Driver, True)
SR2.Req_DLL_name = SR1.Req_DLL_name

End Function

Private Sub Form_Unload(Cancel As Integer)
StopDriverLINXIO SR1
StopDriverLINXIO SR2

End Sub


