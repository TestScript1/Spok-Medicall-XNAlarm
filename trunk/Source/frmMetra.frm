VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmMetrabyte 
   Caption         =   "XnAlarm - Metrabyte PIO-12"
   ClientHeight    =   3660
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   4845
   Icon            =   "frmMetra.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3480
      Top             =   2400
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3285
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   661
      SimpleText      =   ""
      ShowTips        =   0   'False
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6704
            MinWidth        =   6704
            Text            =   "Version 3.0 Copyright 1998 Xtend Communications"
            TextSave        =   "Version 3.0 Copyright 1998 Xtend Communications"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Object.Width           =   2469
            MinWidth        =   2469
            TextSave        =   "9:23 AM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5530
      _Version        =   393216
      Cols            =   5
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "Port  |Type          |^State |Status                     |Counter"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSetup 
         Caption         =   "Setup Board &1"
         Index           =   0
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuExti 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "&System"
      Begin VB.Menu mnuMessages 
         Caption         =   "&Messages"
      End
      Begin VB.Menu mnuServerAlarm 
         Caption         =   "&Server Alarm"
      End
   End
End
Attribute VB_Name = "frmMetrabyte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub GetMetrabyteParams()
    Dim i As Integer
    Dim boardNum As Integer
    Dim Temp As String
    
    For boardNum = 0 To Parameter.numMetrabyteBoards - 1
        Temp = "Board_" + Trim$(Str$(boardNum)) + "_Address"
        metrabyte(boardNum).address = Val(GetIniString("METRABYTE", Temp, "0", Parameter.pPath))
        Temp = "Board_" + Trim$(Str$(boardNum)) + "_ControlWord"
        metrabyte(boardNum).controlByte = GetIniVal("METRABYTE", Temp, 0, Parameter.pPath)
        For i = 0 To 23
            Temp = "Board_" + Trim$(Str$(boardNum)) + "_Port_" + Trim$(Str$(i + 1)) + "_Type"
            metrabyte(boardNum).Port(i).Type = GetIniVal("METRABYTE", Temp, 0, Parameter.pPath)
            
            Temp = "Board_" + Trim$(Str$(boardNum)) + "_Port_" + Trim$(Str$(i + 1)) + "_Usage"
            metrabyte(boardNum).Port(i).usage = GetIniVal("METRABYTE", Temp, 0, Parameter.pPath)
            
            Temp = "Board_" + Trim$(Str$(boardNum)) + "_Port_" + Trim$(Str$(i + 1)) + "_State"
            metrabyte(boardNum).Port(i).state = GetIniVal("METRABYTE", Temp, 0, Parameter.pPath)
        
            Temp = "Board_" + Trim$(Str$(boardNum)) + "_Port_" + Trim$(Str$(i + 1)) + "_Page_Interval"
            metrabyte(boardNum).Port(i).pageInterval = GetIniVal("METRABYTE", Temp, 0, Parameter.pPath)
        
            Temp = "Board_" + Trim$(Str$(boardNum)) + "_Port_" + Trim$(Str$(i + 1)) + "_Page_Timeout"
            metrabyte(boardNum).Port(i).pageTimeout = GetIniVal("METRABYTE", Temp, 0, Parameter.pPath)
        
            Temp = "Board_" + Trim$(Str$(boardNum)) + "_Port_" + Trim$(Str$(i + 1)) + "_Page_Out_Port"
            metrabyte(boardNum).Port(i).outputPort = GetIniVal("METRABYTE", Temp, 0, Parameter.pPath)
        
            Temp = "Board_" + Trim$(Str$(boardNum)) + "_Port_" + Trim$(Str$(i + 1)) + "_Page_Width"
            metrabyte(boardNum).Port(i).pageAlertWidth = GetIniVal("METRABYTE", Temp, 0, Parameter.pPath)
        
            Temp = "Board_" + Trim$(Str$(boardNum)) + "_Port_" + Trim$(Str$(i + 1)) + "_Page_Retry"
            metrabyte(boardNum).Port(i).pageRetry = GetIniVal("METRABYTE", Temp, 0, Parameter.pPath)
        
            Temp = "Board_" + Trim$(Str$(boardNum)) + "_Port_" + Trim$(Str$(i + 1)) + "_Source"
            metrabyte(boardNum).Port(i).sourceFile = GetIniString("METRABYTE", Temp, "", Parameter.pPath)
        
            Temp = "Board_" + Trim$(Str$(boardNum)) + "_Port_" + Trim$(Str$(i + 1)) + "_Destination"
            metrabyte(boardNum).Port(i).destinationFile = GetIniString("METRABYTE", Temp, "", Parameter.pPath)
        
            Temp = "Board_" + Trim$(Str$(boardNum)) + "_Port_" + Trim$(Str$(i + 1)) + "_Extension"
            metrabyte(boardNum).Port(i).Extension = GetIniString("METRABYTE", Temp, "", Parameter.pPath)
        
            Temp = "Board_" + Trim$(Str$(boardNum)) + "_Port_" + Trim$(Str$(i + 1)) + "_Message"
            metrabyte(boardNum).Port(i).xnMessage = GetIniString("METRABYTE", Temp, "", Parameter.pPath)
        
            Temp = "Board_" + Trim$(Str$(boardNum)) + "_Port_" + Trim$(Str$(i + 1)) + "_ActionReminder"
            metrabyte(boardNum).Port(i).ActionReminder = GetIniVal("METRABYTE", Temp, 0, Parameter.pPath)
        
            Temp = "Board_" + Trim$(Str$(boardNum)) + "_Port_" + Trim$(Str$(i + 1)) + "_Page_Numeric"
            metrabyte(boardNum).Port(i).pageNumeric = GetIniString("METRABYTE", Temp, "", Parameter.pPath)
        
            Temp = "Board_" + Trim$(Str$(boardNum)) + "_Port_" + Trim$(Str$(i + 1)) + "_Page_Alpha"
            metrabyte(boardNum).Port(i).pageAlpha = GetIniString("METRABYTE", Temp, "", Parameter.pPath)
        
            Temp = "Board_" + Trim$(Str$(boardNum)) + "_Port_" + Trim$(Str$(i + 1)) + "_Sup_Msg_On"
            metrabyte(boardNum).Port(i).supMsgOn = GetIniString("METRABYTE", Temp, "", Parameter.pPath)
        
            Temp = "Board_" + Trim$(Str$(boardNum)) + "_Port_" + Trim$(Str$(i + 1)) + "_Sup_Msg_Off"
            metrabyte(boardNum).Port(i).supMsgOff = GetIniString("METRABYTE", Temp, "", Parameter.pPath)
        Next
    Next
End Sub

Sub MetraByteCheckAlarmChange(board As Integer, Port As Integer)
    Dim i As Integer
    Dim alarmFlag As Boolean
    Dim LogPATH  As String
    Dim msg1 As String
    Dim msg2 As String
    Dim msg3 As String
    'Dim startTime() As String * 4
    'Dim stopTime() As String * 4
    'Dim duration As String * 4

    On Error GoTo mbcheckalarmerror
  
  
    If metrabyte(board).Port(Port).usage <> 0 Then
        Exit Sub
    End If
    
    i = (board * 24) + Port '  - board ( commented by TK 07/01/1998)
    If metrabyte(board).Port(Port).alarmStatus <> -1 And metrabyte(board).Port(Port).alarmPreviousStatus <> -1 And metrabyte(board).Port(Port).alarmStatus <> metrabyte(board).Port(Port).alarmPreviousStatus And metrabyte(board).Port(Port).Extension <> "" Then
        msg1 = ""
        msg2 = ""
        msg3 = ""
        LogPATH = GetLogPath(metrabyte(board).Port(Port).Extension, "LOG", Parameter.LogPATH)
                
        Select Case metrabyte(board).Port(Port).alarmStatus
        Case 0
            If metrabyte(board).Port(Port).state = 0 Then
                msg1 = "Alarm " + Str$(i + 1) + " On"
                mbStatStart(i) = Now
                mbStat(i).recType = "I" & Format(i + 1, "00")
                mbStat(i).bucket(0) = mbStat(i).bucket(0) + 1
                alarmFlag = True
            Else
                msg1 = "Alarm " + Str$(i + 1) + " Off"
                mbStatStop(i) = Now
                alarmFlag = False
            End If
        Case Else
            If metrabyte(board).Port(Port).state = 0 Then
                msg1 = "Alarm " + Str$(i + 1) + " Off"
                mbStatStop(i) = Now
                alarmFlag = False
            Else
                msg1 = "Alarm " + Str$(i + 1) + " On"
                mbStatStart(i) = Now
                mbStat(i).recType = "I" & Format(i + 1, "00")
                mbStat(i).bucket(0) = mbStat(i).bucket(0) + 1
                alarmFlag = True
            End If
        End Select

        If mbStatStart(i) > 0 Then
            If mbStatStop(i) > 0 Then
                If mbStatStop(i) > mbStatStart(i) Then
                  mbStat(i).recType = "I" & Format(i + 1, "00")
                  mbStat(i).bucket(1) = mbStat(i).bucket(1) + DateDiff("s", mbStatStart(i), mbStatStop(i))
                  mbStatStart(i) = 0
                  mbStatStop(i) = 0
                End If
            End If
        End If
        
        Call XPutMessage(metrabyte(board).Port(Port).ActionReminder, "AL", metrabyte(board).Port(Port).Extension, msg1, metrabyte(board).Port(Port).xnMessage, msg3, , , , , , , , , , gMessageDelivered)
        LogMessage SysMessage, msg1
        If metrabyte(board).Port(Port).pageAlpha <> "" Then
            msg1 = metrabyte(board).Port(Port).pageAlpha
            If alarmFlag = True Then
                msg1 = msg1 + " On"
            Else
                msg1 = msg1 + " Off"
            End If
            msg2 = metrabyte(board).Port(Port).pageNumeric
            If Parameter.BannerAvailable Then gWriteBanner.AddToFile msg1 & " " & msg2
            Call SendPage("XNALARM", metrabyte(board).Port(Port).Extension, msg1, msg2, "", Parameter.CheckStatus)
        End If
        If Parameter.supervisorExtension <> "" Then
            If alarmFlag = True Then
                If metrabyte(board).Port(Port).supMsgOn <> "" Then
                    msg1 = metrabyte(board).Port(Port).supMsgOn
                    If Parameter.BannerAvailable Then gWriteBanner.AddToFile msg1 & " " & msg2
                    Call SendPage("XNALARM", Parameter.supervisorExtension, msg1, msg1, "", Parameter.CheckStatus)
                End If
            Else
                If metrabyte(board).Port(Port).supMsgOff <> "" Then
                    msg1 = metrabyte(board).Port(Port).supMsgOff
                    If Parameter.BannerAvailable Then gWriteBanner.AddToFile msg1 & " " & msg2
                    Call SendPage("XNALARM", Parameter.supervisorExtension, msg1, msg1, "", Parameter.CheckStatus)
                End If
            End If
        End If
    End If
    
    If metrabyte(board).Port(Port).alarmStatus <> -1 And metrabyte(board).Port(Port).alarmStatus <> metrabyte(board).Port(Port).alarmPreviousStatus Then
        msg1 = ""
        Select Case metrabyte(board).Port(Port).alarmStatus
            Case 0
                If metrabyte(board).Port(Port).state = 0 Then
                    msg1 = "Alarm On"
                Else
                    msg1 = "Alarm Off"
                End If
            Case Else
                If metrabyte(board).Port(Port).state = 0 Then
                    msg1 = "Alarm Off"
                Else
                    msg1 = "Alarm On"
                End If
        End Select
        frmMetrabyte.MSFlexGrid1.TextMatrix(i + 1, 3) = msg1
    End If
    metrabyte(board).Port(Port).alarmPreviousStatus = metrabyte(board).Port(Port).alarmStatus

exitmbcheckalarm:
Exit Sub

mbcheckalarmerror:

LogMessage ErrorMessage, "Error in mbCheckAlarmChange " & Error$
Resume exitmbcheckalarm

End Sub

Sub MetrabyteCheckPageState(board As Integer, Port As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim Temp As String
    Dim temp2 As String
    Dim LogPATH As String
    
    On Error GoTo mbcheckpagestateerror
    
    If metrabyte(board).Port(Port).usage = 1 Then
        j = (board * 24) + Port - board
    
        If metrabyte(board).Port(Port).pageState = 1 And metrabyte(board).Port(Port).alarmStatus = 0 And metrabyte(board).Port(Port).state = 0 Then metrabyte(board).Port(Port).pageState = 2
        If metrabyte(board).Port(Port).pageState = 1 And metrabyte(board).Port(Port).alarmStatus = 1 And metrabyte(board).Port(Port).state = 1 Then metrabyte(board).Port(Port).pageState = 2
        
        Select Case metrabyte(board).Port(Port).pageState
            Case 0      ' Wait for timer to expire
                        ' and send a test page
                If DateDiff("s", metrabyte(board).Port(Port).timer, Now) >= metrabyte(board).Port(Port).pageInterval * 60 Then
                    LogMessage SysMessage, "Mb page test: board: " & board + 1 & " port " & Port + 1
                    Temp = metrabyte(board).Port(Port).pageAlpha
                    temp2 = metrabyte(board).Port(Port).pageNumeric
                    'If Parameter.BannerAvailable Then gWriteBanner.AddToFile Temp & " " & temp2
                    Call SendPage("XNALARM", metrabyte(board).Port(Port).Extension, Temp, temp2, "", Parameter.CheckStatus)
                    metrabyte(board).Port(Port).pageTimer = Now
                    metrabyte(board).Port(Port).pageState = 1
                    metrabyte(board).Port(Port).pageCount = 1
                    frmMetrabyte.MSFlexGrid1.TextMatrix(j + 1, 3) = "Page Listen"
                    mbStatStart(i) = Now
                    mbStat(j).recType = "T" & Format(j + 1, "00")
                    mbStat(j).bucket(0) = mbStat(j).bucket(0) + 1
                End If
            Case 1      ' Wait for test page response
                If DateDiff("s", metrabyte(board).Port(Port).pageTimer, Now) >= metrabyte(board).Port(Port).pageTimeout * 60 Then
                    LogMessage SysMessage, "Mb Page Test: NO Response, board " & board + 1 & " port" & Port + 1
                    LogMessage StatMessage, ", " & board & ", " & Port & "," & Format$(metrabyte(board).Port(Port).pageTimer, "MM/DD HH:MM:SS") & "," & Format$(Now, "MM/DD HH:MM:SS") & ", FAILURE"
                    metrabyte(board).Port(Port).timer = Now
                    If metrabyte(board).Port(Port).pageCount < metrabyte(board).Port(Port).pageRetry Then
                        metrabyte(board).Port(Port).pageCount = metrabyte(board).Port(Port).pageCount + 1
                        LogMessage SysMessage, "Mb page test: board: " & board + 1 & " port " & Port + 1 & " Retry: " & metrabyte(board).Port(Port).pageCount
                        Temp = metrabyte(board).Port(Port).pageAlpha
                        temp2 = metrabyte(board).Port(Port).pageNumeric
                        If Parameter.BannerAvailable Then gWriteBanner.AddToFile Temp & " " & temp2
                        Call SendPage("XNALARM", metrabyte(board).Port(Port).Extension, Temp, temp2, "", Parameter.CheckStatus)
                        metrabyte(board).Port(Port).pageTimer = Now
                        frmMetrabyte.MSFlexGrid1.TextMatrix(j + 1, 3) = "Page Listen " & metrabyte(board).Port(Port).pageCount
                        mbStatStart(j) = Now
                        mbStat(j).recType = "T" & Format(j + 1, "00")
                        mbStat(j).bucket(0) = mbStat(j).bucket(0) + 1      'Send
                    Else
                        Temp = "Failure, No Response"
                        temp2 = ""
                        Call XPutMessage(metrabyte(board).Port(Port).ActionReminder, "AL", metrabyte(board).Port(Port).Extension, metrabyte(board).Port(Port).xnMessage, Temp, temp2, , , , , , , , , , 0)    'gMessageDelivered)
                        If Parameter.BannerAvailable Then gWriteBanner.AddToFile metrabyte(board).Port(Port).xnMessage & " " & Temp
                        metrabyte(board).Port(Port).pageState = 0
                        metrabyte(board).Port(Port).pageAlarm = 1    ' Alarm Condition
                        frmMetrabyte.MSFlexGrid1.TextMatrix(j + 1, 3) = "Page Idle"
                    End If
                    mbStat(j).recType = "T" & Format(j + 1, "00")
                    mbStat(j).bucket(1) = mbStat(j).bucket(1) + 1   'No response
                    mbStatStart(j) = Now
                End If
            Case 2
                LogMessage StatMessage, ", " & board & ", " & Port & "," & Format$(metrabyte(board).Port(Port).pageTimer, "MM/DD HH:MM:SS") & "," & Format$(Now, "MM/DD HH:MM:SS") & ", RECEIVED"
                LogMessage SysMessage, "Pager response: board " & board + 1 & " port " & Port + 1
                Temp = "Pager Responded"
                temp2 = ""
                Call XPutMessage(1, "AL", metrabyte(board).Port(Port).Extension, metrabyte(board).Port(Port).xnMessage, Temp, temp2, , , , , , , , , , gMessageDelivered)
                metrabyte(board).Port(Port).timer = Now
                metrabyte(board).Port(Port).pageState = 3
                frmMetrabyte.MSFlexGrid1.TextMatrix(j + 1, 3) = "Page Received"
            Case 3
                i = False
                k = False
                '
                ' if edge sensitive just check to see if alarmstatus changes
                '
                If metrabyte(board).Port(Port).pageAlertWidth = 0 Then
                    If metrabyte(board).Port(Port).alarmStatus = 0 And metrabyte(board).Port(Port).state = 1 Then i = True
                    If metrabyte(board).Port(Port).alarmStatus = 1 And metrabyte(board).Port(Port).state = 0 Then i = True
                Else
                    If metrabyte(board).Port(Port).alarmStatus = 0 And metrabyte(board).Port(Port).state = 1 Then
                        i = True
                        k = True
                    End If
                    If metrabyte(board).Port(Port).alarmStatus = 1 And metrabyte(board).Port(Port).state = 0 Then
                        i = True
                        k = True
                    End If
                    If i And DateDiff("s", metrabyte(board).Port(Port).timer, Now) < metrabyte(board).Port(Port).pageAlertWidth Then
                        i = False
                    End If
                    If i = False And k = True Then
                        metrabyte(board).Port(Port).pageState = 1
                        metrabyte(board).Port(Port).timer = Now
                        frmMetrabyte.MSFlexGrid1.TextMatrix(j + 1, 3) = "Page Listen"
                    End If
                End If
                If i Or DateDiff("s", metrabyte(board).Port(Port).timer, Now) >= 60 Then
                    LogMessage SysMessage, "Shutting off pager, Board " & board + 1 & " Port " & Port + 1
                    metrabyte(board).Port(Port).timer = Now
                    metrabyte(board).Port(Port).pageState = 4
                    metrabyte(board).Port(Port).pageAlarm = 0     ' No Alarm
                    frmMetrabyte.MSFlexGrid1.TextMatrix(j + 1, 3) = "Pager Off"
                    Call MetrabyteSetPort(False, board, metrabyte(board).Port(Port).outputPort - 1)
                    mbStatStop(j) = Now
                    If mbStatStart(j) > 0 Then
                        If mbStatStop(j) > 0 Then
                            If mbStatStop(j) > mbStatStart(j) Then
                                LogMessage SysMessage, "Page " & mbStat(j).bucket(2) & " recieved seconds " & mbStat(j).bucket(3)
                                mbStat(j).recType = "T" & Format(j + 1, "00") 'AD
                                mbStat(j).bucket(2) = mbStat(j).bucket(2) + 1
                                mbStat(j).bucket(3) = mbStat(j).bucket(3) + DateDiff("s", mbStatStart(j), mbStatStop(j))
                                mbStatStart(j) = 0
                                mbStatStop(j) = 0
                            End If
                        End If
                    End If
                    LogMessage SysMessage, "Debug 4"
                End If
            Case 4
                If DateDiff("s", metrabyte(board).Port(Port).timer, Now) >= 5 Then
                    LogMessage SysMessage, "Pager on, board " & board + 1 & " port " & Port + 1
                    metrabyte(board).Port(Port).timer = Now
                    metrabyte(board).Port(Port).pageState = 0
                    frmMetrabyte.MSFlexGrid1.TextMatrix(j + 1, 3) = "Page Idle"
                    Call MetrabyteSetPort(True, board, metrabyte(board).Port(Port).outputPort - 1)
                End If
            Case Else
                metrabyte(board).Port(Port).pageState = 0
        End Select
        '==========================================================================
        
        
        If metrabyte(board).Port(Port).pageAlarm <> -1 And metrabyte(board).Port(Port).pageAlarm <> metrabyte(board).Port(Port).alarmPreviousStatus Then
            If metrabyte(board).Port(Port).pageAlarm = 0 Then
                Call RemoveFile(board, Port)
                If Parameter.supervisorExtension <> "" And metrabyte(board).Port(Port).supMsgOff <> "" Then
                    Temp = metrabyte(board).Port(Port).supMsgOff
                    'If Parameter.BannerAvailable Then gWriteBanner.AddToFile Temp
                    Call SendPage("XNALARM", Parameter.supervisorExtension, Temp, Temp, "", Parameter.CheckStatus)
                End If
            Else
                mbStat(j).recType = "T" & Format(j + 1, "00") 'AD
                mbStat(j).bucket(4) = mbStat(j).bucket(4) + 1
                Call MoveFile(board, Port)
                If Parameter.supervisorExtension <> "" And metrabyte(board).Port(Port).supMsgOn <> "" Then
                    Temp = metrabyte(board).Port(Port).supMsgOn
                    'If Parameter.BannerAvailable Then gWriteBanner.AddToFile Temp
                    Call SendPage("XNALARM", Parameter.supervisorExtension, Temp, Temp, "", Parameter.CheckStatus)
                End If
            End If
        End If
    End If

exitmbcheckpage:
Exit Sub

mbcheckpagestateerror:

LogMessage ErrorMessage, "Error in mbCheckPageState " & Error$
Resume exitmbcheckpage

End Sub

Sub MetrabyteSetPort(state As Boolean, board As Integer, Port As Integer)
Dim i As Integer
Dim j As Integer
Dim Value As Byte

On Error GoTo mbsetporterror

If board < 0 Or Port < 0 Then Exit Sub
LogMessage SysMessage, "Setting Board: " & board + 1 & " port: " & Port + 1 & " state: " & state
i = Chan(Port)
'Select Case port
'    Case 0 To 7
'        i = 0
'    Case 8 To 15
'        i = 1
'    Case 16 To 23
'        i = 2
'    Case Else
'        i = 0
'End Select
j = board * 24 + Port + 1
If UCase(Parameter.alarmtype) = "METRABYTE" Then
    Value = DlPortReadPortUchar(metrabyte(board).address + i)
End If
If UCase(Parameter.alarmtype) = "KEITHLEY" Then
    Value = gDriverLINX.ReadChannel(i)
End If
If state = True Then
    If metrabyte(board).Port(Port).state = 0 Then
        Value = Value Or (2 ^ (Port - i * 8))
        frmMetrabyte.MSFlexGrid1.TextMatrix(j, 2) = "1"
    Else
        Value = Value And Not (2 ^ (Port - i * 8))
        frmMetrabyte.MSFlexGrid1.TextMatrix(j, 2) = "0"
    End If
Else
    If metrabyte(board).Port(Port).state = 0 Then
        Value = Value And Not (2 ^ (Port - i * 8))
        frmMetrabyte.MSFlexGrid1.TextMatrix(j, 2) = "0"
    Else
        Value = Value Or (2 ^ (Port - i * 8))
        frmMetrabyte.MSFlexGrid1.TextMatrix(j, 2) = "1"
    End If
End If
If UCase(Parameter.alarmtype) = "METRABYTE" Then
    DlPortWritePortUchar metrabyte(board).address + i, Value
End If
If UCase(Parameter.alarmtype) = "KEITHLEY" Then
    Call gDriverLINX.WriteChannel(i, Value)  ' error
    
    
End If
exitmbsetport:
    Exit Sub
mbsetporterror:
    LogMessage ErrorMessage, "Error: " & Err & " in mbSetPort " & Error$
    Resume exitmbsetport
End Sub

Sub RemoveFile(board As Integer, Port As Integer)
    On Error GoTo removefileerror
    
    If metrabyte(board).Port(Port).destinationFile <> "" And metrabyte(board).Port(Port).alarmPreviousStatus <> -1 Then
        LogMessage SysMessage, "Deleteing " & metrabyte(board).Port(Port).destinationFile
        Kill metrabyte(board).Port(Port).destinationFile
    End If
    metrabyte(board).Port(Port).alarmPreviousStatus = metrabyte(board).Port(Port).pageAlarm

exitremovefile:
    Exit Sub
removefileerror:
    If Err = 53 Then
        metrabyte(board).Port(Port).alarmPreviousStatus = metrabyte(board).Port(Port).pageAlarm
    End If
    LogMessage ErrorMessage, "Error: " & Err & " in RemoveFile " & Error$
    Resume exitremovefile
End Sub

Sub MoveFile(board As Integer, Port As Integer)
    On Error GoTo movefileerror
    
    If metrabyte(board).Port(Port).destinationFile <> "" And metrabyte(board).Port(Port).sourceFile <> "" And metrabyte(board).Port(Port).alarmPreviousStatus <> -1 Then
        LogMessage SysMessage, "Renaming " & metrabyte(board).Port(Port).sourceFile & " as " & metrabyte(board).Port(Port).destinationFile
        FileCopy metrabyte(board).Port(Port).sourceFile, metrabyte(board).Port(Port).destinationFile
    End If
    metrabyte(board).Port(Port).alarmPreviousStatus = metrabyte(board).Port(Port).pageAlarm

exitmovefile:
    Exit Sub
movefileerror:
    LogMessage ErrorMessage, "Error: " & Err & " in MoveFile " & Error$
    Resume exitmovefile
End Sub

Private Sub Form_Load()
    Call GetMetrabyteParams
    StatusBar1.Panels(1).Text = "Version " & App.Major & "." & App.Minor & "." & App.Revision & "  Copyright Xtend Communications"
    Me.Caption = appTitle & " - " & Parameter.alarmtype
    If UCase(Parameter.alarmtype) = "KEITHLEY" Then
        Me.Caption = Me.Caption & " " & gDriverLINX.DLModelName
    End If
    
    
    mnuServerAlarm.Enabled = AlarmOn
        
End Sub

Private Sub Form_Resize()
On Error Resume Next
MSFlexGrid1.Width = Me.Width - 250
MSFlexGrid1.Height = Me.Height - StatusBar1.Height - 700

'ExitHere:
'Exit Sub

'OOPS:
'LogMessage ErrorMessage, "Error in frmMetrabyte.Form_Resize " & Error$
'Resume ExitHere
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    On Error Resume Next

    If AlarmOn Then
        Unload SysAlarm
    End If

    Call WriteStats(i, "XNAL", "", mbPorts, mbStat())
    XnFrqTable.Close
    
    CloseAllTables
    Call XnCloseDataBase
    CloseLog SysMessage
    CloseLog ErrorMessage
    CloseLog StatMessage
    Set gDriverLINX = Nothing
    
    Set objCdoEmail = Nothing
    Set objMapiOutlook = Nothing
    
    End

exitmetrabyteunload:
    End
metrabyteunloaderror:
    LogMessage ErrorMessage, "Error: " & Err & " in Metrabyte Unload " & Error$
    Resume exitmetrabyteunload
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuExti_Click()
Unload frmMetrabyte
End Sub

Private Sub mnuMessages_Click()
    If Errorform.Visible = True Then
        mnuMessages.Checked = False
        Errorform.Visible = False
    Else
        mnuMessages.Checked = True
        Errorform.Visible = True
    End If
End Sub

Private Sub mnuServerAlarm_Click()
SysAlarm.Visible = True
End Sub

Private Sub mnuSetup_Click(index As Integer)
    metrabyteBoardNumber = index
    frmMbSetup.Show
End Sub

Private Sub mnuSystem_Click()
    If Errorform.Visible = True Then
        mnuMessages.Checked = True
    Else
        mnuMessages.Checked = False
    End If
End Sub


Private Sub Timer1_Timer()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim L As Integer
    Dim m As Integer
    Dim result As Integer
    Dim Value As Byte
    Dim timenow As Integer
    
    On Error GoTo mbtimererror
    
    Timer1.Enabled = False
    
    timenow = Hour(Now)
    If mbStartStat <> timenow Then
       Call WriteStats(i, "XNAL", "", mbPorts, mbStat())
       If i Then mbStartStat = timenow
    End If
    
    '
    '       Check for changes in Metrabyte ports
    '
    L = 1
    For i = 0 To Parameter.numMetrabyteBoards - 1
        m = 0
        For j = 0 To 2
            If UCase(Parameter.alarmtype) = "METRABYTE" Then
                Value = DlPortReadPortUchar(metrabyte(i).address + j)
            End If
            If UCase(Parameter.alarmtype) = "KEITHLEY" Then
                Value = gDriverLINX.ReadChannel(j)
            End If
            For k = 0 To 7
                If metrabyte(i).Port(m).Type = True Then
                    result = Value And 2 ^ k
                    If result > 0 Then result = 1
                    frmMetrabyte.MSFlexGrid1.TextMatrix(L, 2) = result
                    metrabyte(i).Port(m).alarmStatus = result
                    Call MetraByteCheckAlarmChange(i, m)
                    Call MetrabyteCheckPageState(i, m)
                End If
                L = L + 1
                m = m + 1
            Next
        Next
    Next
    Timer1.Enabled = True
    
exitmbtimer:
Exit Sub

mbtimererror:

LogMessage ErrorMessage, "Error in Metrabyte timer " & Error$
Resume exitmbtimer
    
End Sub
