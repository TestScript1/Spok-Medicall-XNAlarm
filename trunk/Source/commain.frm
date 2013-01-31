VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form ComMain 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  Xtend - Dataprobe ANN Interface"
   ClientHeight    =   2025
   ClientLeft      =   1815
   ClientTop       =   3645
   ClientWidth     =   8565
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
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
   Icon            =   "COMMAIN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   8565
   Begin VB.Timer timerCheckFile 
      Interval        =   5000
      Left            =   150
      Top             =   1425
   End
   Begin VB.CommandButton comActivate 
      Caption         =   "&Activate now"
      Height          =   615
      Left            =   3840
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer TimerStart 
      Interval        =   500
      Left            =   240
      Top             =   120
   End
   Begin VB.Timer DataprobeTimer 
      Interval        =   500
      Left            =   7440
      Top             =   1320
   End
   Begin MSCommLib.MSComm ComPort 
      Index           =   0
      Left            =   8040
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
   Begin VB.Label lblTime 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblLed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape GreenLed 
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   840
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   255
   End
   Begin VB.Shape RedLed 
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   840
      Shape           =   3  'Circle
      Top             =   840
      Width           =   255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSaveParams 
         Caption         =   "&Save Settings"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuVI 
      Caption         =   "View Info"
      Begin VB.Menu mnuFileDebug 
         Caption         =   "&Debug"
      End
      Begin VB.Menu mComPortInfo 
         Caption         =   "Com Port Info"
      End
      Begin VB.Menu mnuSystemMessages 
         Caption         =   "System &Messages"
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "Setup"
      Begin VB.Menu mnuSetupCom 
         Caption         =   "Setup Com &1"
         Index           =   0
      End
      Begin VB.Menu mnuDpSetup 
         Caption         =   "Data &Probe Setup"
      End
   End
   Begin VB.Menu mnuSysAlarm 
      Caption         =   "Server Alarm"
   End
End
Attribute VB_Name = "ComMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const DARK_GREEN = &H547C58
Const DARK_ORANGE = &H4675A4

Private Sub AniPushButton1_Click()

End Sub




Function FirstResponse() As String
Dim i As Integer
Const Default_CMD = "x6+01+x2+S+x3"
Dim sTmp As String
Dim arrCmd() As String
sTmp = GetIniString("DATA PROBE", "ACK_INIT_COMMAND", Default_CMD, Parameter.pPath, True)
arrCmd() = Split(sTmp, "+")
sTmp = ""
For i = 0 To UBound(arrCmd)

    If LCase(Left(arrCmd(i), 1)) = "x" Then
        sTmp = sTmp & Chr(Right(arrCmd(i), 1))
    Else
        sTmp = sTmp & arrCmd(i)
    End If
    
Next

Call dpCheckSum(sTmp)
FirstResponse = sTmp

End Function

Sub SendACK()
Dim strACK As String
strACK = Chr(6) & "01" & Chr(2) & "W" & Chr(3)
Call dpCheckSum(strACK)
ComPort(0).Output = strACK
Call ShowData(frmDPInfo.txtTerm, strACK, vbRed)

End Sub

Function VerifyCheckSum(parPacket As String) As Boolean
Dim origPacket As String

' here we verify of check sum is correct, if not we need to send negative acknoledge signal
origPacket = parPacket

Call dpCheckSum(origPacket)
If origPacket = parPacket Then
    VerifyCheckSum = True
Else
    VerifyCheckSum = False
End If

End Function

Private Sub comActivate_Click()
    TimerStart.Enabled = False
    StartingProc = True
    lblTime.Caption = ""
    comActivate.Visible = False
    
End Sub

Private Sub ComPort_OnComm(index As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim L As Integer
    Dim tempData As String
    Dim in1 As String
    Dim cl As String
    Dim Temp As String
    Dim msgOK As Boolean
    Static bNextTime As Boolean
    
    On Error GoTo porterror

    cl = Chr$(1) + "01" + Chr$(2) + "A" + Chr$(3) + "b"
    
    Select Case ComPort(index).CommEvent
    ' Errors
        Case comEventBreak    ' A Break was received.
                                ' Code to handle a BREAK goes here.
        Case comEventCDTO     ' CD (RLSD) Timeout.
        Case comEventCTSTO    ' CTS Timeout.
        Case comEventDSRTO    ' DSR Timeout.
        Case comEventFrame    ' Framing Error
        Case comEventOverrun  ' Data Lost.
        Case comEventRxOver   ' Receive buffer overflow.
        Case comEventRxParity ' Parity Error.
        Case comEventTxFull   ' Transmit buffer full.
    ' Events
        Case comEvCD       ' Change in the CD line.
        Case comEvCTS      ' Change in the CTS line.
        Case comEvDSR      ' Change in the DSR line.
        Case comEvRing     ' Change in the Ring Indicator.

        Case comEvReceive  ' Received RThreshold # of chars.
        
            If (ComPort(index).InBufferCount) > 0 Then
                tempData = ComPort(index).Input
                
                Call Scope.ScopeInput(index, tempData)
                
                For i = 1 To Len(tempData)
                    in1 = Mid$(tempData, i, 1)
                    Select Case comState
                    Case 0
                        If Asc(in1) = 1 Or Asc(in1) = 6 Then    ' SOH
                            portInBuffer = in1
                            comState = 1
                        End If
                    Case 1
                        portInBuffer = portInBuffer + in1
                        If Asc(in1) = 3 Then    ' ETX
                            comState = 2
                        End If
                    Case 2
                        portInBuffer = portInBuffer + in1
                        Temp = Mid(portInBuffer, 1, 5) & Chr(3)   ' eliminate data from the ACK message
                        If VerifyCheckSum(portInBuffer) = False Then ' we send negative ACK signal NAK
                            msgOK = False
                            Mid(Temp, 1, 1) = Chr(21)  'NAK
                            Call dpCheckSum(Temp)
                        Else
                            msgOK = True
                            Mid$(Temp, 1, 1) = Chr$(6)  'ACK
                            Call dpCheckSum(Temp)
                        End If
                        If bNextTime = False Then Temp = FirstResponse(): bNextTime = True
                        ComPort(index).Output = Temp
                        Call Scope.ScopeOutPut(index, Temp)
                        Call dpProcessPacket(portInBuffer)
                        Call ShowData(frmDPInfo.txtTerm, portInBuffer, vbBlue)
                        If msgOK Then Call ShowData(frmDPInfo.txtTerm, Temp, DARK_ORANGE)
                        If Not msgOK Then Call ShowData(frmDPInfo.txtTerm, Temp, vbRed)
                        comState = 0
                    Case Else
                        comState = 0
                    End Select
                Next i
            End If

        Case comEvSend ' There are SThreshold number of
            ' characters in the transmit buffer.
        Case Else
    End Select

exitport:
Exit Sub

porterror:

LogMessage ErrorMessage, "Unexpected error in OnComm " & Error$
Resume exitport

End Sub


Private Sub DataprobeTimer_Timer()
    Dim i As Integer
    Dim Temp As String
    Dim temp2 As String
    Dim LogPATH As String
    Dim timenow As Integer

    On Error GoTo timererror
    
    Exit Sub  ' TK  disable timer
    
    
    timenow = Hour(Now)
    If startStat <> timenow Then
       Call WriteStats(i, "XNAL", "", maxport, alarmStat())
       If i Then startStat = timenow
    End If
    timerState = timerState + 1
    If timerState > 30 Then timerState = 0

    Select Case timerState
        Case 0, 5, 15, 20
            Temp = Chr$(1) + "01" + Chr$(2) + "S" + Chr$(3)
            Call dpCheckSum(Temp)
            temp2 = Chr$(1) + "01" + Chr$(2) + "P" + Chr$(3)
            Call dpCheckSum(temp2)
            For i = 0 To Parameter.totalPorts - 1
                Select Case cport(i).checkPointer
                    Case 0
                        ComMain!ComPort(i).Output = Temp
                        Call Scope.ScopeOutPut(i, Temp)
                        cport(i).checkPointer = 1
                    Case 1
                        ComMain!ComPort(i).Output = Temp ' temp2
                        'Call Scope.ScopeOutPut(i, temp2)
                        Call Scope.ScopeOutPut(i, Temp)
                        cport(i).checkPointer = 0
                    Case Else
                        cport(i).checkPointer = 0
                End Select
            Next
        Case 30
            For i = 0 To 15
                If dpPort(i).Type = 1 Then
                    Select Case dpPort(i).pageState
                        Case 0      ' Wait for timer to expire
                                    ' and send a test page
                            If DateDiff("n", dpPort(i).timer, Now) >= dpPort(i).pageMin Then
                                LogMessage SysMessage, "Transmitter test: port " & i + 1
                                Temp = dpPort(i).pageMessage
                                temp2 = dpPort(i).pageNumeric
                                Call SendPage("XNALARM", dpPort(i).Extension, Temp, temp2, "", Parameter.CheckStatus)
                                startTimeP(i) = Now
                                dpPort(i).timer = Now
                                dpPort(i).pageState = 1
                                alarmStat(i).recType = "P" & Format(i + 1, "00") 'AD
                                alarmStat(i).bucket(0) = alarmStat(i).bucket(0) + 1      'Send  AD
                            End If
                        Case 1      ' Wait for test page response
                            If DateDiff("n", dpPort(i).timer, Now) >= dpPort(i).pageTimeout Then
                                LogMessage SysMessage, "NO Response, Port" & i + 1
                                Temp = "Failure, No Response"
                                temp2 = ""
                                Call XPutMessage(0, "AL", dpPort(i).Extension, dpPort(i).xnMessage, Temp, temp2, , , , , , , , , , gMessageDelivered)
                                dpPort(i).timer = Now
                                dpPort(i).pageState = 0
                                alarmStat(i).recType = "P" & Format(i + 1, "00") 'AD
                                alarmStat(i).bucket(1) = alarmStat(i).bucket(1) + 1   'No response
                                startTimeP(i) = 0
                            End If
                        Case 2
                            LogMessage SysMessage, "Transmitter Resepnse: port " & i + 1
                            Temp = "Pager Responded"
                            temp2 = ""
                            alarmStat(i).recType = "P" & Format(i + 1, "00") 'AD
                            alarmStat(i).bucket(2) = alarmStat(i).bucket(2) + 1
                            alarmStat(i).bucket(3) = alarmStat(i).bucket(3) + DateDiff("s", dpPort(i).timer, Now)
                            Call XPutMessage(dpPort(i).ActionReminder, "AL", dpPort(i).Extension, dpPort(i).xnMessage, Temp, temp2, , , , , , , , , , gMessageDelivered)
                            dpPort(i).timer = Now
                            dpPort(i).relayState = 1
                            dpPort(i).pageState = 3
                            Call dpSetRelayOn
                            startTimeP(i) = 0
                        Case 3
                            If DateDiff("s", dpPort(i).timer, Now) >= 5 Then
                                LogMessage SysMessage, "Shutting off relay, Port " & i + 1
                                dpPort(i).timer = Now
                                dpPort(i).pageState = 4
                                dpPort(i).relayState = 0
                                Call dpSetRelayOff
                            End If
                        Case 4
                            If DateDiff("s", dpPort(i).timer, Now) >= 5 Then
                                LogMessage SysMessage, "Relay On: port " & i + 1
                                dpPort(i).timer = Now
                                dpPort(i).relayState = 1
                                dpPort(i).pageState = 5
                                Call dpSetRelayOn
                            End If
                        Case 5
                            If DateDiff("s", dpPort(i).timer, Now) >= 5 Then
                                LogMessage SysMessage, "Relay off, Port " & i + 1
                                dpPort(i).timer = Now
                                dpPort(i).pageState = 0
                                dpPort(i).relayState = 0
                                Call dpSetRelayOff
                            End If
                        Case Else
                            dpPort(i).pageState = 0
                    End Select
                End If
            Next
    End Select
    '
    '       Set the color of the LEDs
    '
    For i = 0 To 15
        If dpPort(i).pageState = 1 And dpPort(i).alarmStatus = 0 Then dpPort(i).pageState = 2
        If dpPort(i).alarmStatus <> -1 Then
            If timerState / 2 = Int(timerState / 2) Then  ' only for timerState = even number
                If dpPort(i).alarmStatus = 0 And dpPort(i).alarmAcked <> 0 Then
                    ComMain!RedLed(i).FillColor = &HFF&
                End If
                If dpPort(i).alarmStatus <> 0 And dpPort(i).alarmClear <> 0 Then
                    ComMain!GreenLed(i).FillColor = &HFF00&
                End If
            Else
                If dpPort(i).alarmStatus = 0 And dpPort(i).alarmAcked <> 0 Then
                    ComMain!RedLed(i).FillColor = &HC0C0FF
                End If
                If dpPort(i).alarmStatus <> 0 And dpPort(i).alarmClear <> 0 Then
                    ComMain!GreenLed(i).FillColor = &HC0FFC0
                End If
            End If
        End If
    Next

exittimer:
Exit Sub

timererror:

LogMessage ErrorMessage, "Unexpected error in Timer " & Error$
Resume exittimer

End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim Temp As String
    
    DataprobeTimer.Enabled = False
    
    On Error GoTo OOPS
    If appTitle <> "" Then
        Me.Caption = appTitle
    End If
    mnuSysAlarm.Enabled = AlarmOn
    For i = 1 To 15
        Load RedLed(i)
        RedLed(i).Left = RedLed(i - 1).Left + 450
        RedLed(i).Visible = True
        Load GreenLed(i)
        GreenLed(i).Left = GreenLed(i - 1).Left + 450
        GreenLed(i).Visible = True
        Load lblLed(i)
        lblLed(i).Left = lblLed(i - 1).Left + 450
        lblLed(i).Caption = Str$(i + 1)
        lblLed(i).Visible = True
    Next
    For i = 0 To 15
        dpPort(i).alarmPreviousStatus = -1
        dpPort(i).alarmStatus = -1
    Next
    If ComPort(0).PortOpen <> False Then ComPort(0).PortOpen = False
    ComPort(0).CommPort = cport(0).port
    ComPort(0).Settings = cport(0).setup
    ComPort(0).PortOpen = True
    Temp = ComPort(0).Input

    portInBuffer = ""
    comState = 0
    timerState = 31
    'DataprobeTimer.Enabled = True
    
    Load frmDPInfo
    
ExitHere:
    Exit Sub
OOPS:
LogMessage ErrorMessage, "Error in Form.Commain_Load, " & Error$
Resume ExitHere
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then Cancel = True

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    On Error Resume Next
    
    Parameter.RectMain.Top = ComMain.Top
    Parameter.RectMain.Left = ComMain.Left
    Parameter.RectMain.Height = ComMain.Height
    Parameter.RectMain.Width = ComMain.Width
    
    ComPort(0).PortOpen = False
    
    Unload ComSetup
    Unload Scope
    Unload frmMetrabyte
    Unload frmDPInfo
    XnFrqTable.Close
    
    CloseAllTables
    '''CloseScheduleBTR
    Call XnCloseDataBase
    Call SaveParameters
    Call SaveDPChannels
    For i = 0 To 15
        If dpStartTime(i) > 0 Then
            alarmStat(i).recType = "A" & Format(i + 1, "00")
            alarmStat(i).bucket(1) = alarmStat(i).bucket(1) + DateDiff("s", dpStartTime(i), Now)
        End If
        If startTimeP(i) > 0 Then
           alarmStat(i).recType = "P" & Format(i + 1, "00")
           alarmStat(i).bucket(1) = alarmStat(i).bucket(1) + 1
        End If
    Next
    Call WriteStats(i, "XNAL", "", maxport, alarmStat()) 'AD
       
    Set objCdoEmail = Nothing
    Set objMapiOutlook = Nothing
    
    End
End Sub

Private Sub Grid1_RowColChange()

End Sub

Private Sub mComPortInfo_Click()
frmDPInfo.Show

End Sub

Private Sub mnuDpSetup_Click()
 '   Dim i As Integer
    Unload DpSetup
    Load DpSetup
'
    DpSetup.Show 0
End Sub

Private Sub mnuFileDebug_Click()
    Call Scope.ScopeStart
End Sub

Private Sub mnuFileExit_Click()

    Dim szResult As String
    If FileForExiting <> "" Then
        If Dir(FileForExiting) <> "" Then
            LogMessage SysMessage, "File " & FileForExiting & " is found ! Delete this file and exit XNALARM."
            Kill FileForExiting
        End If
    End If
    If AlarmOn Then Unload SysAlarm
    Unload Me
End Sub

Private Sub mnuSaveParams_Click()
    Call SaveParameters
End Sub

Private Sub mnuSetupCom_Click(index As Integer)
    Unload ComSetup
    setupPort = index
    Load ComSetup
    ComSetup.Caption = "Port " + Trim$(Str$(index + 1)) + " Configuration"
    ComSetup.Top = Parameter.RectSetup.Top
    ComSetup.Left = Parameter.RectSetup.Left
    ComSetup.Height = Parameter.RectSetup.Height
    ComSetup.Width = Parameter.RectSetup.Width
    ComSetup.Show 0
End Sub

Private Sub mnuSysAlarm_Click()
    SysAlarm.Visible = True
End Sub

Private Sub mnuSystemMessages_Click()
    If Errorform.Visible = True Then
        mnuSystemMessages.Checked = False
        Errorform.Visible = False
    Else
        Errorform.Visible = True
        mnuSystemMessages.Checked = True
    End If
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub timerCheckFile_Timer()
Dim bExit As Boolean

If FileForExiting = "" Then timerCheckFile.Enabled = False: Exit Sub

CheckForTheFile bExit

If bExit Then
    mnuFileExit_Click
    
End If

End Sub

Private Sub TimerStart_Timer()
Static oldTime As Date
Static tic As Integer
Dim timeDiff As Integer

On Error GoTo OOPS

If tic = 0 Then
    tic = 1
    oldTime = Now
    lblTime.Caption = "Time to activate: " & DelayToStart & " sec"
    lblTime.Refresh
    comActivate.Visible = True
    Exit Sub
End If


timeDiff = DateDiff("s", oldTime, Now)
lblTime.Caption = "Time to activate: " & CStr(DelayToStart - timeDiff) & " sec"
If timeDiff > DelayToStart Then ' 2.5 minutes

    TimerStart.Enabled = False
    
    StartingProc = True
    lblTime.Caption = ""
   comActivate.Visible = False
End If
lblTime.Refresh
ExitHere:
Exit Sub

OOPS:
LogMessage ErrorMessage, "Error in TimerStart, " & Error$
Resume ExitHere
End Sub


