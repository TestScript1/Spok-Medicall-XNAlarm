VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmSimplex 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Xtend Alarm Gateway"
   ClientHeight    =   2265
   ClientLeft      =   150
   ClientTop       =   750
   ClientWidth     =   5175
   FillColor       =   &H00C0C0C0&
   Icon            =   "Frmsimpl.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMultiLineAlarm 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   720
      Top             =   360
   End
   Begin VB.Timer timerCheckFile 
      Interval        =   5000
      Left            =   90
      Top             =   390
   End
   Begin VB.CommandButton comAlarms 
      Caption         =   "Show &Alarms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3930
      TabIndex        =   8
      Top             =   825
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show &Buffer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2640
      TabIndex        =   7
      Top             =   825
      Width           =   1185
   End
   Begin VB.CommandButton btnSystemMessages 
      Caption         =   "&Show Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1365
      TabIndex        =   1
      Top             =   825
      Width           =   1185
   End
   Begin VB.CommandButton btnOptionsDebug 
      Caption         =   "&View Scope"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   60
      TabIndex        =   0
      Top             =   825
      Width           =   1185
   End
   Begin VB.Timer SimplexTimer 
      Interval        =   500
      Left            =   4710
      Top             =   345
   End
   Begin MSCommLib.MSComm MSComm1 
      Index           =   0
      Left            =   4575
      Top             =   -135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   4096
      NullDiscard     =   -1  'True
      RThreshold      =   1
      BaudRate        =   1200
   End
   Begin VB.Label lblEventcount 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   4260
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblWarning 
      Caption         =   "Database is not responding !!! Please check."
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   270
      TabIndex        =   9
      Top             =   1980
      Width           =   4725
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   2910
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   2910
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Alarms"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Current Hour  -    Data Packets:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   150
      TabIndex        =   2
      Top             =   1455
      Width           =   4800
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuComSetup 
         Caption         =   "&Com Port Setup"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSysAlarm 
      Caption         =   "Server alarm"
   End
End
Attribute VB_Name = "frmSimplex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub HandleInput(ByVal tmpBuffer As String)
Dim i As Integer
Dim j As Integer
Dim tmp As String
Static tempData As String
Dim in1 As String
Dim totalChars As Integer
Static inHere As Boolean  ' to prevent re-currency
    
On Error GoTo OOPS
100                tmp = tmpBuffer
110                tmp = tempData + tmp
115
120                tempData = ""  ' new added 03/22/01
130                totalChars = Len(tmp)
135                For i = 1 To totalChars  ' parse wrong data
140                     in1 = Mid$(tmp, i, 1)
150                     If Asc(in1) <> 0 Then
155                          If (Asc(in1) < 32 Or Asc(in1) > 126) And Asc(in1) <> 13 And Asc(in1) <> 10 _
                                     And Asc(in1) <> 27 And Asc(in1) <> 160 Then
                                     in1 = ""   ' ignore character
175                          ElseIf Asc(in1) = 27 And i < totalChars Then ' skip next char too if escape
176                                 in1 = ""  ' eliminate 27
177                                 i = i + 1  ' skip character following 27
179                          End If
190                          tempData = tempData + in1
200                     End If
210                 Next i
201                 Do
202                     LogMessage SysMessage, "Enter Do LOOP.------------------------------"
220                     Do While Left(tempData, 1) = vbCr Or Left(tempData, 1) = vbLf
230                          tempData = Mid(tempData, 2)
240                     Loop
245                     LogMessage SysMessage, "tempData = " & tempData & " Length=" & Len(tempData)
250                     j = InStr(tempData, vbCr)
255                     LogMessage SysMessage, "vbCr IS FOUND IN POSITION:" & j
260                     If j > 0 Then
270                         portInBuffer = Left(tempData, j + 1)
280                         tempData = Mid(tempData, j + 1)
285                         frmSimplexList.AddToList portInBuffer
                              If inHere = False Then '///////////// in here if ------------
                                    inHere = True
                                    
290                               Call SimplexProccessPacket(portInBuffer)  '----> need to move to a timer for processing
292                               inHere = False
320                               portInBuffer = ""
323                         Else
324                               lblEventcount = Val(lblEventcount) + 1: lblEventcount.Refresh
325                         End If
335                     End If
                          'DoEvents
337                  Loop While j > 0
                       LogMessage SysMessage, "Finished Do Loop.----------------"



ExitHere:


Exit Sub


OOPS:
LogMessage ErrorMessage, "Unexpected error: " & CStr(Err) & " in HandleInput " & Error$ & ", line#" & CStr(Erl)

Resume ExitHere
    
End Sub

Sub PositionTheForm()

With Me
    .Top = GetIniString("SCREEN", "frmSimplex_Top", .Top, Parameter.pPath)
    .Left = GetIniString("SCREEN", "frmSimplex_Left", .Left, Parameter.pPath)
    .Width = GetIniString("SCREEN", "frmSimplex_Width", .Width, Parameter.pPath)
    .Height = GetIniString("SCREEN", "frmSimplex_Height", .Height, Parameter.pPath)
End With
End Sub

Private Sub btnOptionsDebug_Click()
    MousePointer = 11
    Call Scope.ScopeStart
    MousePointer = 0
End Sub

Private Sub btnOptionsDebug_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblHelp.Caption = "Allows you to see all data as they come through the Com Ports."
End Sub


Private Sub btnSystemMessages_Click()
    MousePointer = 11
    If Errorform.Visible = True Then
        btnSystemMessages.Caption = "Show Status"
        Errorform.Visible = False
    Else
        Errorform.Visible = True
        'btnSystemMessages.Caption = "Hide Status"
    End If
    MousePointer = 0
End Sub

Private Sub btnSystemMessages_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblHelp.Caption = "You can view the activites, events, errors, and any system messages."
End Sub


Private Sub comAlarms_Click()
frmAlarms.Show
End Sub

Private Sub comAlarms_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblHelp.Caption = "You can view the alarms list, parsed from the port buffer."
End Sub


Private Sub Command1_Click()
frmSimplexList.Show

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
 lblHelp.Caption = "You can view buffer information, coming from com port."
End Sub


Private Sub Form_Activate()
    lblHelp.Caption = "Move your mouse over any button to see a quick description."
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strTemp As String
    
    centerform frmSimplex
    lblWarning.Visible = False
    PositionTheForm
    mnuSysAlarm.Enabled = AlarmOn
    SimplexTimer.Enabled = True
    Me.Caption = appTitle & " (ver. " & App.Major & "." & App.Minor & "." & App.Revision & ")"
    LogMessage SysMessage, "Simplex AlarmTypes: " & Parameter.maxAlarmTypes
    For i = 0 To Parameter.maxAlarmTypes - 1
        LogMessage SysMessage, "Type: " & simplex(i).alarm & " Position: " & simplex(i).alarmPosition
    Next
    
    If MSComm1(0).PortOpen = True Then
        MSComm1(0).PortOpen = False
    End If
    MSComm1(0).CommPort = cport(0).port
    MSComm1(0).Settings = cport(0).setup
    MSComm1(0).PortOpen = True
    strTemp = MSComm1(0).Input
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblHelp.Caption = "Move your mouse over any button to see a quick description."
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then Cancel = True
If UnloadMode = vbFormCode Then LogMessage SysMessage, "=========== Application was closed from the menu. ==============="
If UnloadMode = vbAppWindows Then LogMessage SysMessage, "=========== Application was shut down by Windows. ==============="
If UnloadMode = vbAppTaskManager Then LogMessage SysMessage, "=========== Application was closed using Task Manager. ==============="

End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next  'Added CKO 8/31/00
  
    Dim x As Integer
    
    '*** CKO 8/31/00
    If ActionReminder Then
      Set ActReminder = Nothing
    End If
    '***
    
    SimplexTimer.Enabled = False
    Unload ComSetup
    Unload Scope
    
    XnFrqTable.Close
    
    CloseAllTables
    Call XnCloseDataBase
    Call SaveParameters
    
    MSComm1(0).PortOpen = False
    
    x = False              'AD
    Call WriteStats(x, "XNAL", "", Parameter.maxAlarmTypes, simplexStat())
       
    Set objCdoEmail = Nothing
    Set objMapiOutlook = Nothing
    
    End
End Sub





Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblHelp.Caption = "Move your mouse over any button to see a quick description."
End Sub


Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuComSetup_Click()
    Dim szResult As String
    
    szResult = InputBox("Please, enter a password to access Com Port settings.", "Enter Password")
    If UCase(szResult) <> "XTEND" Then Exit Sub
    centerform ComSetup
    MousePointer = 11
    Load ComSetup
    ComSetup.Show
    MousePointer = 0
End Sub

Private Sub mnuExit_Click()
    Dim szResult As String
    
    If FileForExiting <> "" Then
        If Dir(FileForExiting) <> "" Then
            LogMessage SysMessage, "File " & FileForExiting & " is found ! Delete this file and exit XNALARM."
            Kill FileForExiting
        End If
    End If

    If AlarmOn Then Unload SysAlarm
    Unload frmSimplex

End Sub

Private Sub mnuSysAlarm_Click()
    SysAlarm.Visible = True
End Sub

Private Sub MSComm1_OnComm(index As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim tmp As String
    Static tempData As String
    Dim in1 As String
    Dim totalChars As Integer
    Static inHere As Boolean  ' to prevent re-currency
    
    On Error GoTo porterror

    Select Case MSComm1(index).CommEvent
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
       ' DoEvents
                
100            If (MSComm1(index).InBufferCount) > 0 Then
105                tmp = MSComm1(index).Input

106                Call Scope.ScopeInput(index, tmp)
107                Call LogMessage(CommPortMsg, tmp)  ' save to file xnALARM.Port
                     Call HandleInput(tmp)
110                'tmp = tempData + tmp
115
120                'tempData = ""  ' new added 03/22/01
130                'totalChars = Len(tmp)
135                'For i = 1 To totalChars  ' parse wrong data
140                '     in1 = Mid$(tmp, i, 1)
150                '     If Asc(in1) <> 0 Then
155                '          If (Asc(in1) < 32 Or Asc(in1) > 126) And Asc(in1) <> 13 And Asc(in1) <> 10 _
                     '                And Asc(in1) <> 27 And Asc(in1) <> 160 Then
                     '                in1 = ""   ' ignore character
175                '          ElseIf Asc(in1) = 27 And i < totalChars Then ' skip next char too if escape
176                '                 in1 = ""  ' eliminate 27
177                '                 i = i + 1  ' skip character following 27
179                '          End If
190                '          tempData = tempData + in1
200                '     End If
210                ' Next i
201                ' Do
202                 '    LogMessage SysMessage, "Enter Do LOOP.------------------------------"
220                 '    Do While Left(tempData, 1) = vbCr Or Left(tempData, 1) = vbLf
230                 '         tempData = Mid(tempData, 2)
240                 '    Loop
245                 '    LogMessage SysMessage, "tempData = " & tempData & " Length=" & Len(tempData)
250                 '    j = InStr(tempData, vbCr)
255                 '    LogMessage SysMessage, "vbCr IS FOUND IN POSITION:" & j
260                 '    If j > 0 Then
270                 '        portInBuffer = Left(tempData, j + 1)
280                 '        tempData = Mid(tempData, j + 1)
285                 '        frmSimplexList.AddToList portInBuffer
                      '        If inHere = False Then '///////////// in here if ------------
                      '              inHere = True
'
'290                               Call SimplexProccessPacket(portInBuffer)  '----> need to move to a timer for processing
'292                               inHere = False
'320                               portInBuffer = ""
'323                         Else
'324                               lblEventcount = Val(lblEventcount) + 1: lblEventcount.Refresh
'325                         End If
'335                     End If
'                          'DoEvents
'337                  Loop While j > 0
'                       LogMessage SysMessage, "Finished Do Loop.----------------"
338            End If

        Case comEvSend ' There are SThreshold number of
            ' characters in the transmit buffer.
        Case Else
    End Select

exitport:
Exit Sub

porterror:

LogMessage ErrorMessage, "Unexpected error: " & CStr(Err) & " in OnComm " & Error$ & ", line#" & CStr(Erl)
Resume exitport
Resume
End Sub




Private Sub SimplexTimer_Timer()
    Dim timenow As Integer
    Dim i As Integer
    Dim Temp As String

    On Error GoTo timererror
    ' This timer clear the counter boxes every hour.
    ' And write statistics
    
    timenow = Hour(Now)
    If startStat <> timenow Then
        Call WriteStats(i, "XNAL", "", Parameter.maxAlarmTypes, simplexStat())
        If i = True Then
            startStat = timenow
            simplexPackets = 0
            frmSimplex.Label2.Caption = 0
            frmSimplex.Label4.Caption = 0
        End If
    End If
      
exittimer:
Exit Sub

timererror:

LogMessage ErrorMessage, "Unexpected error in Timer " & Error$
Resume exittimer

End Sub


Private Sub timerCheckFile_Timer()
Dim bExit As Boolean

If FileForExiting = "" Then timerCheckFile.Enabled = False: Exit Sub

CheckForTheFile bExit

If bExit Then
    mnuExit_Click
    
End If


End Sub


Private Sub tmrMultiLineAlarm_Timer()
' this timer waits a sertain amount of time for the Com port
' to send second line of alarm message.

End Sub


