Attribute VB_Name = "START"
Option Explicit

Public AlarmOn  As Boolean ' TK added 06/18/2001
'---------- for DB connection -----
Public username As String
Public XnPassword As String
'-----------------------------------
Private Const MODULENAME = " @Start." 'Added CKO 8/31/00
Public appTitle As String
Public gDriverLINX As Object  'DIOComm.clsKeithleyDriver   'Object

Public gWriteBanner As clsScreenNotify

Type portType
    checkPointer As Integer
    port As Integer
    setup As String
End Type

Type DPType
    Type As Integer
    inputState As Integer
    outputState As Integer
    clearAlarm As Integer
    pageMin As Integer
    pageTimeout As Integer
    pageState As Integer
    ActionReminder As Integer
    Extension As String
    xnMessage As String
    pageMessage As String
    pageNumeric As String
    alarmPreviousStatus As Integer
    alarmStatus As Integer
    alarmAcked As Integer
    alarmClear As Integer
    relayState As Integer      ' 0 Off, 1 On
    timer As Variant
End Type
 
Type SimplexType
    alarmPosition As Integer
    conditionPosition As Integer
    conditionLenght As Integer
    Page As String  '   Boolean TK changed 01/22/2004 from Boolean to String
    ActionReminder As Boolean
    alarm As String
    conditions As String
    AlarmLines As Integer  ' TK added 03/25/04 to accept more then 1 buffer for the alarm
    ThisIsAlarm() As Boolean  ' TRUE-Take this line as  an alarm, false - skip this line
    SendBanner As Boolean
    
End Type

Type metrabytePortType
    Type As Integer
    usage As Integer
    state As Integer
    pageInterval As Integer
    pageTimeout As Integer
    outputPort As Integer
    pageAlertWidth As Integer
    pageRetry As Integer
    ActionReminder As Integer
    alarmPreviousStatus As Integer
    alarmStatus As Integer              ' Status of Alarm port: 0 - port on, 1 port off
    pageState As Integer
    pageAlarm As Integer                ' 0 - No Alarm, 1 Alarm State
    pageCount As Integer
    sourceFile As String
    destinationFile As String
    Extension As String
    xnMessage As String
    pageNumeric As String
    pageAlpha As String
    timer As Variant
    pageTimer As Variant
    supMsgOn As String
    supMsgOff As String
End Type

Type metrabyteBoardType
    address As Integer
    controlByte As Integer
    port(23) As metrabytePortType
End Type

Type metrabyteControlType
    Control As Integer
    PType(3) As Boolean
End Type

Type ParameterType
    pPath As String
    LogPATH As String
    DiagMsgFile As String
    XnDataBase As String
    totalPorts As Integer
    RectMain As FrameType
    RectSetup As FrameType
    RectDpSetup As FrameType
    alarmtype As String
    maxAlarmTypes As Integer
    pointPosition As Integer
    pointLength As Integer
    pointExclude As String
    logPrefix As String
    alphaprefix As String
    numericprefix As String
    alphamessagelength As Integer
    alphamessageposition As Integer
    numericmessagelength As Integer
    numericmessageposition As Integer
    CheckStatus As Integer
    username As String
    logalphamessageoffset As Integer
    logalphamessagelength As Integer
    lognumericmessageoffset As Integer
    lognumericmessagelength As Integer
    numMetrabyteBoards As Integer
    supervisorExtension As String
    BannerAvailable As Boolean  ' if set to true - save paged msg to a special file that will appear on a operator's screen
    EmailAvailable As Boolean
End Type

Public gMessageDelivered As Boolean
Global cport() As portType
Global Parameter As ParameterType
Global dpPort(16) As DPType

Global simplex() As SimplexType
Global simplexStartStat  As Integer                 ' simplex statistics timer
Global simplexPackets As Integer                    ' Total Packets recieved within hour
Global simplexStat() As StatType                    ' simplex statistic records

Global metrabyte() As metrabyteBoardType
Global metrabyteBoardNumber As Integer          ' board number for setup program
Global mbPorts As Integer                       ' total number of metrabyte ports
Global mbStartStat  As Integer                  ' metrabyte statistics timer
Global mbStat() As StatType                     ' metrabyte statistic records
Global mbStatStart() As Variant                 ' metrabyte statistics timer record
Global mbStatStop() As Variant                  ' metrabyte statistics timer record

Global portInBuffer As String
Global captureFileNum As Integer
Global monitorPort As Integer
Global setupPort As Integer

Global maxport As Integer               ' dataprobe number statistic records
Global startStat  As Integer            ' dataprobe statistics timer
Global alarmStat() As StatType          ' dataprobe statistic records
Global dpStartTime() As Variant           ' dataprobe alarm statistics timer record
Global stopTime() As Variant            ' dataprobe alarm statistics timer record
Global startTimeP() As Variant          ' dataprobe page statistics timer record
Global stopTimeP() As Variant           ' dataprobe page statistics timer record

Global comState As Integer
Global timerState As Integer

Global ErrorMessage As MessageType
Global SysMessage As MessageType
Global StatMessage As MessageType
Global CommPortMsg As MessageType

Global alarmindex As Integer

' Set ActionReminder = 1 in INI file to use the Calendar object
Public ActionReminder As Integer  'Added CKO 8/31/00

Public gAppVersion As String  'Added CKO 9/1/00
Public FileForExiting ' file and path in that must be in order to xnALARM exit itself

Public objCdoEmail As Object  ' send e-mail in case of DB disconnect or network problems
Public objMapiOutlook As Object  ' send e-mail using EmailSrv.dll

Public gEmailServer As String
Public gEmailAddress As String

Public Function Chan(ByVal parPort As Integer) As Integer

' returns channel # if accepted port #

Select Case parPort

    Case 0 To 7
        Chan = 0
    Case 8 To 15
        Chan = 1
    Case 16 To 23
        Chan = 2
    Case Else
        Chan = 0
End Select
End Function

Public Sub CheckForTheFile(parResult As Boolean)
' this routine will check for the special file that appears on a special sub-dir
' if the file is found, kill this file and exit the application XNALARM
' This is done in order to make backups and updates at night
On Error GoTo TRAP

If Dir(FileForExiting) <> "" Then
    LogMessage SysMessage, "File " & FileForExiting & " is found ! Delete this file and exit XNALARM."
    Kill FileForExiting
    parResult = True
Else
    parResult = False
End If

ExitHere:
Exit Sub

TRAP:
parResult = False
Resume ExitHere
End Sub

Sub GetParameters()
  On Error GoTo ErrHandler  'Added CKO 8/31/00
  
  Const MYNAME = "GetParameters()"  'Added CKO 8/31/00
    Dim AlarmTimerCounter As Integer, AlarmTimeSchedule    As Integer
    Dim iRepeatInterval As Integer
    Dim i As Integer
    Dim j As Single
    Dim Temp As String
    Dim strAlarmMsg1 As String, strAlarmMsg2 As String
    Dim dsn As String
    Dim Value As String
    Dim strEmailMode As String
        
    'dsn = GetIniString("GENERAL", "DSN", "Smart Answer", Parameter.pPath)
    username = GetIniString("GENERAL", "UserID", "", Parameter.pPath)
    XnPassword = GetIniString("GENERAL", "Password", "", Parameter.pPath)
    
    Parameter.alarmtype = GetIniString("GENERAL", "AlarmType", ".", Parameter.pPath)
    Parameter.username = GetIniString("GENERAL", "UserName", "", Parameter.pPath)
    Parameter.supervisorExtension = GetIniString("GENERAL", "SupervisorExtension", "", Parameter.pPath)
      
    Parameter.numMetrabyteBoards = GetIniVal("METRABYTE", "numBoards", 0, Parameter.pPath)
    
    Parameter.totalPorts = GetIniVal("PORTS", "TotalPorts", 1, Parameter.pPath)
    Parameter.DiagMsgFile = GetIniString("XN", "DIAGNOSTIC_MSG_FILE", "XNALR.MSG", Parameter.pPath)
    Parameter.XnDataBase = GetIniString("General", "DSN", "Smart Answer", Parameter.pPath)
    Parameter.LogPATH = GetIniString("XN", "LOG_PATH", ".", Parameter.pPath)
    
    '*** Added CKO 8/31/00
    Temp = GetIniString("XN", "ActionReminder", "", Parameter.pPath)
    If Len(Temp) = 0 Then
      Temp = "0"
      WriteIniString "XN", "ActionReminder", Temp, Parameter.pPath
    End If
    ActionReminder = CInt(Temp)
    If ActionReminder Then
      On Error Resume Next
      Set ActReminder = CreateObject("CalendarVB5.CalendarCol")
      DoEvents
      If Err.Number <> 0 Then
        Temp = "[E] " & Err.Number & ": " & Err.Description & MODULENAME & MYNAME _
          & "Problem to Set ActReminder = CreateObject(""CalendarVB5.CalendarCol"")"
        MsgBox Temp, vbCritical
        End
      End If
      On Error GoTo ErrHandler
    End If
    '***
    
    Parameter.maxAlarmTypes = GetIniVal("SIMPLEX", "MAXALARMTYPES", 0, Parameter.pPath)
    Parameter.pointExclude = GetIniString("SIMPLEX", "PointExclude", "-", Parameter.pPath)
    Parameter.pointLength = GetIniVal("SIMPLEX", "PointLength", 6, Parameter.pPath)
    Parameter.pointPosition = GetIniVal("SIMPLEX", "PointPosition", 12, Parameter.pPath)
    Parameter.logPrefix = GetIniString("SIMPLEX", "LogPrefix", "", Parameter.pPath)

    Parameter.logalphamessageoffset = GetIniVal("SIMPLEX", "LogAlpahMessageOffset", 0, Parameter.pPath)
    Parameter.logalphamessagelength = GetIniVal("SIMPLEX", "LogAlpahMessageLength", 0, Parameter.pPath)
    Parameter.lognumericmessageoffset = GetIniVal("SIMPLEX", "LogNumericMessageOffset", 0, Parameter.pPath)
    Parameter.lognumericmessagelength = GetIniVal("SIMPLEX", "LogNumericMessageLength", 0, Parameter.pPath)
    
    Parameter.CheckStatus = GetIniVal("SIMPLEX", "CheckStatus", 1, Parameter.pPath)
    Parameter.alphamessagelength = GetIniVal("SIMPLEX", "AlphaMessageLength", 0, Parameter.pPath)
    Parameter.alphamessageposition = GetIniVal("SIMPLEX", "AlphaMessagePosition", 0, Parameter.pPath)
    Parameter.numericmessagelength = GetIniVal("SIMPLEX", "numericmessagelength", 0, Parameter.pPath)
    Parameter.numericmessageposition = GetIniVal("SIMPLEX", "numericmessageposition", 0, Parameter.pPath)
    Parameter.numericprefix = GetIniString("SIMPLEX", "numericprefix", "", Parameter.pPath)
    Parameter.alphaprefix = GetIniString("SIMPLEX", "alphaprefix", "", Parameter.pPath)
       
    ReDim cport(Parameter.totalPorts)
    For i = 0 To Parameter.totalPorts - 1
        Temp = "ComPort" + Trim$(Str$(i + 1))
        cport(i).port = GetIniVal("PORTS", Temp, i + 1, Parameter.pPath)
        Temp = "ComSetup" + Trim$(Str$(i + 1))
        cport(i).setup = GetIniString("PORTS", Temp, "1200,N,8,1", Parameter.pPath)
        cport(i).checkPointer = 1
    Next
    
    ReDim simplex(Parameter.maxAlarmTypes) As SimplexType
    
    For i = 0 To Parameter.maxAlarmTypes - 1
        Temp = "AlarmType(" + Trim$(Str$(i + 1)) + ")"
        simplex(i).alarm = GetIniString("SIMPLEX", Temp, "", Parameter.pPath)
        Temp = "AlarmCondition(" + Trim$(Str$(i + 1)) + ")"
        simplex(i).conditions = GetIniString("SIMPLEX", Temp, "", Parameter.pPath)
        Temp = "AlarmPosition(" + Trim$(Str$(i + 1)) + ")"
        simplex(i).alarmPosition = GetIniVal("SIMPLEX", Temp, 1, Parameter.pPath)
        Temp = "ConditionPosition(" + Trim$(Str$(i + 1)) + ")"
        simplex(i).conditionPosition = GetIniVal("SIMPLEX", Temp, 18, Parameter.pPath)
        Temp = "ConditionLenght(" + Trim$(Str$(i + 1)) + ")"
        simplex(i).conditionLenght = GetIniVal("SIMPLEX", Temp, 3, Parameter.pPath)
        Temp = "Page(" + Trim$(Str$(i + 1)) + ")"
        simplex(i).Page = UCase(GetIniString("SIMPLEX", Temp, True, Parameter.pPath))
        Temp = "ActionReminder(" + Trim$(Str$(i + 1)) + ")"
        simplex(i).ActionReminder = GetIniBool("SIMPLEX", Temp, True, Parameter.pPath)
        Temp = "TotalAlarmLines(" & CStr(i + 1) & ")"
        simplex(i).AlarmLines = GetIniVal("SIMPLEX", Temp, 1, Parameter.pPath)
        If simplex(i).AlarmLines = 0 Then simplex(i).AlarmLines = 1 ' cannot be 0
        
        
        If simplex(i).AlarmLines > 1 Then  ' for multi-line alarms - determine if we skip a line
            For j = 1 To simplex(i).AlarmLines
                Temp = "ALARM_LINE" & CStr(j) & "(" & CStr(i + 1) & ")"
                Value = GetIniBool("SIMPLEX", Temp, True, Parameter.pPath)
                ReDim Preserve simplex(i).ThisIsAlarm(j)
                simplex(i).ThisIsAlarm(j) = Value
            Next
        End If
        Temp = "SEND_TICKER(" & CStr(i + 1) & ")"
        simplex(i).SendBanner = GetIniBool("SIMPLEX", Temp, False, Parameter.pPath, True)
    Next

    For i = 0 To 15
        Temp = Trim$(Str$(i + 1)) + "_Type"
        dpPort(i).Type = GetIniVal("DATA PROBE", Temp, 0, Parameter.pPath)
        Temp = Trim$(Str$(i + 1)) + "_InputState"
        dpPort(i).inputState = GetIniVal("DATA PROBE", Temp, 0, Parameter.pPath)
        Temp = Trim$(Str$(i + 1)) + "_OutPutState"
        dpPort(i).outputState = GetIniVal("DATA PROBE", Temp, 0, Parameter.pPath)
        Temp = Trim$(Str$(i + 1)) + "_ClearAlarm"
        dpPort(i).clearAlarm = GetIniVal("DATA PROBE", Temp, 0, Parameter.pPath)
        Temp = Trim$(Str$(i + 1)) + "_PageMin"
        dpPort(i).pageMin = GetIniVal("DATA PROBE", Temp, 0, Parameter.pPath)
        Temp = Trim$(Str$(i + 1)) + "_PageTimeout"
        dpPort(i).pageTimeout = GetIniVal("DATA PROBE", Temp, 0, Parameter.pPath)
        Temp = Trim$(Str$(i + 1)) + "_ActionReminder"
        dpPort(i).ActionReminder = GetIniVal("DATA PROBE", Temp, 0, Parameter.pPath)
        Temp = Trim$(Str$(i + 1)) + "_Extension"
        dpPort(i).Extension = GetIniString("DATA PROBE", Temp, "", Parameter.pPath)
        Temp = Trim$(Str$(i + 1)) + "_XnMessage"
        dpPort(i).xnMessage = GetIniString("DATA PROBE", Temp, "", Parameter.pPath)
        Temp = Trim$(Str$(i + 1)) + "_PageMEssage"
        dpPort(i).pageMessage = GetIniString("DATA PROBE", Temp, "", Parameter.pPath)
        Temp = Trim$(Str$(i + 1)) + "_PageNumeric"
        dpPort(i).pageNumeric = GetIniString("DATA PROBE", Temp, "", Parameter.pPath)
    Next
    DelayToStart = GetIniVal("DATA PROBE", "DELAY_TO_START", 0, Parameter.pPath)
    
    Parameter.RectMain.Top = GetIniSingle("SCREEN", "ComMain_Top", 4920, Parameter.pPath)
    Parameter.RectMain.Left = GetIniSingle("SCREEN", "ComMain_Left", 600, Parameter.pPath)
    Parameter.RectMain.Height = GetIniSingle("SCREEN", "ComMain_Height", 2055, Parameter.pPath)
    Parameter.RectMain.Width = GetIniSingle("SCREEN", "ComMain_Width", 4320, Parameter.pPath)
                                                  
    Parameter.RectSetup.Top = GetIniSingle("SCREEN", "ComSetup_Top", 3765, Parameter.pPath)
    Parameter.RectSetup.Left = GetIniSingle("SCREEN", "ComSetup_Left", 1515, Parameter.pPath)
    Parameter.RectSetup.Height = GetIniSingle("SCREEN", "ComSetup_Height", 2640, Parameter.pPath)
    Parameter.RectSetup.Width = GetIniSingle("SCREEN", "ComSetup_Width", 4275, Parameter.pPath)
    appTitle = GetIniString("GENERAL", "TITLE", "Xtend Alarm Gateway", Parameter.pPath)
    
    ' added by TK for system alarm on  06/18/2001
    AlarmOn = GetIniBool("XTEND ALARM", "AlarmOn", False, Parameter.pPath)
    AlarmTimerCounter = GetIniVal("XTEND ALARM", "AlarmTimerCounter", 60, Parameter.pPath)
    AlarmTimeSchedule = GetIniVal("XTEND ALARM", "AlarmTimeSchedule", 100, Parameter.pPath)
    strAlarmMsg1 = GetIniString("XTEND ALARM", "ABNORMAL MSG ID", "", Parameter.pPath)
    strAlarmMsg2 = GetIniString("XTEND ALARM", "CLOSE MSG ID", "", Parameter.pPath)
    iRepeatInterval = GetIniVal("XTEND ALARM", "RepeatInterval", 70, Parameter.pPath)
    
    If AlarmOn Then
        SysAlarm.SetObjStatPara AlarmTimerCounter, AlarmTimeSchedule, strAlarmMsg1, strAlarmMsg2, iRepeatInterval, Parameter.XnDataBase, username, XnPassword      ' 9,10 - needs to be changed
        Load SysAlarm
    End If
    Parameter.BannerAvailable = GetIniBool("BANNER", "AVAILABLE", False, Parameter.pPath, True)
    
    If Parameter.BannerAvailable Then
        Set gWriteBanner = New clsScreenNotify
    End If
    
    Parameter.EmailAvailable = GetIniBool("ErrorSystem", "SendEmail", False, Parameter.pPath, True)
    
    If Parameter.EmailAvailable Then
       
        gEmailAddress = GetIniString("ErrorSystem", "EmailAddress", "", Parameter.pPath, True)
        strEmailMode = GetIniString("ErrorSystem", "EmailMode", "", Parameter.pPath, True)
        
        Select Case UCase(strEmailMode)
            Case "SMTP"
                Set objCdoEmail = CreateObject("CDOSendEmail.NewMail")
                
            Case "MAPI", "OUTLOOK"
                Set objMapiOutlook = CreateObject("EMailSrv.OLEMailing")
                objMapiOutlook.Mode = strEmailMode
                If objMapiOutlook.Mode <> "OUTLOOK" Then
                    objMapiOutlook.username = GetIniString("ErrorSystem", "EmailProfile", "", Parameter.pPath, True)
                    objMapiOutlook.PassWord = GetIniString("ErrorSystem", "EmailPassword", "", Parameter.pPath, True)
                End If

            Case Else
                Parameter.EmailAvailable = False
            
        End Select
        
    End If
    
    
GetOut:
  Exit Sub
  
ErrHandler:
  MsgBox "[E] " & Err.Number & ": " & Err.Description & MODULENAME & MYNAME
  Resume Next
End Sub



Public Sub SendEmailNote(parMessage As String)
Dim strEmailFROM As String
Dim strEmailSUBJECT As String
Dim obj As Object ' for compatability

On Error GoTo OOPS

strEmailFROM = "XNALARM"
strEmailSUBJECT = "DB Connection or Network is down."
If Parameter.EmailAvailable Then
    If Not (objCdoEmail Is Nothing) Then
        objCdoEmail.SendEMail strEmailFROM, gEmailAddress, strEmailSUBJECT, parMessage
    End If
    If Not (objMapiOutlook Is Nothing) Then
        ' send using mapi or outlook
        objMapiOutlook.EMailAddr = gEmailAddress
        objMapiOutlook.Subject = strEmailSUBJECT
        objMapiOutlook.JobType = "E"
        objMapiOutlook.MessageToDeliver = parMessage
        objMapiOutlook.SendEMail obj
           
    End If
End If
ExitHere:
Exit Sub

OOPS:
LogMessage ErrorMessage, "Err " & Err.Number & " " & Err.Description & " at SendEmailNote()"
Resume ExitHere
End Sub


Public Sub SetAllTables()

    OpenXKMTable (Parameter.pPath)
    SetPagersTable (Parameter.pPath)
    SetMsgTable (Parameter.pPath)
    SetCountTable (Parameter.pPath)


End Sub

Public Sub SplitIntoLines(InLine As String, LineLength As Integer, Prefix As String, Suffix As String, Lines() As String)

' Devide BIG message to be sent to a pager by portions of xxx characters

Dim iSub As Integer, pos As Integer, Temp As String, j As Integer

    Temp = RTrim$(InLine)
    iSub = -1
    pos = LineLength - Len(Suffix)
    If Len(Temp) > LineLength Then
        Do While Len(Temp) > 0
            If iSub >= 0 Then Lines(iSub) = Lines(iSub) & Suffix
            If Len(Temp) > pos Then
                For j = pos + 1 To 1 Step -1
                    If Mid$(Temp, j, 1) = " " Then
                        pos = j
                        Exit For
                    End If
                Next
            End If
            iSub = iSub + 1
            ReDim Preserve Lines(iSub)
            Lines(iSub) = RTrim$(Mid$(Temp, 1, pos))
            Temp = LTrim$(Mid$(Temp, pos + 1))
            pos = LineLength - Len(Prefix)
            If iSub > 0 Then Lines(iSub) = Prefix & Lines(iSub)
            If Len(Temp) > pos Then pos = pos - Len(Suffix)
        Loop
    End If
    If iSub = -1 Then
        ReDim Lines(0)
        Lines(0) = InLine
    End If
End Sub
Sub Main()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim Temp As String
    Dim formname As Form
    Dim bFRQOpen As Boolean
    On Error GoTo MainError

    'Debug.Print "Time is  " & startStat; " o'clock"
    Call getpath
    Parameter.pPath = pth + "XNALARM.INI"
    
    '*** Added CKO 9/1/00
    gAppVersion = App.Major & "." & Format$(App.Minor, "00") & "." & Format$(App.Revision, "0000")
    '***
    
    Call GetParameters
    Call InitLog(ErrorMessage, Parameter.pPath, "XNALARM.ERR", "Error")
    Call InitLog(SysMessage, Parameter.pPath, "XNALARM.LOG", "System")
    Call InitLog(StatMessage, Parameter.pPath, "XNALSTAT.DAT", "Stat")
    Call InitLog(CommPortMsg, Parameter.pPath, "XNALARM.PORT", "Commport")
    
    StatMessage.fileCircular = False
    Errorform.Caption = "Xtend Alarm Monitoring System Status"
    LogMessage SysMessage, "Initializing XNALARM"
    Temp = App.EXEName
    
    FileForExiting = GetIniString("GENERAL", "FILENAME_FOR_EXIT", "", Parameter.pPath)
    If FileForExiting <> "" Then CheckForTheFile False
    gMessageDelivered = GetIniBool("XN", "MESSAGE_DELIVERED", False, Parameter.pPath, True)
    If UCase(Parameter.alarmtype) = "DATAPROBE" Then
        '
        '       Initialize statistic records
        '
        startStat = Hour(Now)
        maxport = 16
        ReDim alarmStat(maxport)
        ReDim dpStartTime(maxport)
        ReDim stopTime(maxport)
        ReDim startTimeP(maxport)
        ReDim stopTimeP(maxport)
        For i = 0 To 15
            dpStartTime(i) = 0
            stopTime(i) = 0
        Next
        Call InitializeStats(maxport, alarmStat())
    
       Load ComMain
       If Parameter.totalPorts > 1 Then
            For i = 1 To Parameter.totalPorts - 1
                Load ComMain!ComPort(i)
                ComMain!ComPort(i).PortOpen = False
                ComMain!ComPort(i).CommPort = cport(i).port
                ComMain!ComPort(i).Settings = cport(i).setup
                ComMain!ComPort(i).PortOpen = True
                Temp = ComMain!ComPort(i).Input

                Load ComMain!mnuSetupCom(i)
                ComMain!mnuSetupCom(i).Caption = "Setup Com &" + Trim$(Str$(i + 1))
            Next
        End If
    
        For i = 0 To 15
            dpPort(i).timer = Now
            dpPort(i).pageState = 0
            dpPort(i).relayState = 0
        Next

        ComMain.Top = Parameter.RectMain.Top
        ComMain.Left = Parameter.RectMain.Left
        'ComMain.Height = Parameter.RectMain.Height
        'ComMain.Width = Parameter.RectMain.Width

        Set formname = ComMain
        ComMain.Show 0
        '''Call dpSetRelayOn
        '''Call dpSetRelayOff
        Call InitDataProbe
        ComMain.Refresh
        Call Scope.ScopeInitialize(Parameter.totalPorts, "XNALARM.CAP", Parameter.pPath)
        Call ReadDPChannels
    ElseIf UCase(Parameter.alarmtype) = "SIMPLEX" Then
        ReDim simplexStat(Parameter.maxAlarmTypes)
        simplexPackets = 0
        Call InitializeStats(Parameter.maxAlarmTypes, simplexStat())
        simplexStartStat = Hour(Now)
        
        Load frmSimplex
        frmSimplex!SimplexTimer.Enabled = True
        Set formname = frmSimplex
        frmSimplex.Show
        Call Scope.ScopeInitialize(Parameter.totalPorts, "XNALARM.CAP", Parameter.pPath)
    End If
      
       
    
    
    If UCase(Parameter.alarmtype) = "KEITHLEY" And Parameter.numMetrabyteBoards > 0 Then
        ' call object (initialize board)
        Set gDriverLINX = CreateObject("DIOComm.clsKeithleyDriver") 'New DIOComm.clsKeithleyDriver
        Call InitializeBoard
        
    End If
    If UCase(Parameter.alarmtype) = "METRABYTE" And Parameter.numMetrabyteBoards > 0 Then
        '
        '       Initialize statistic records
        '
  '      mbStartStat = Hour(Now)
  '      mbPorts = Parameter.numMetrabyteBoards * 24 - 1
  '      ReDim mbStat(mbPorts)
  '      ReDim mbStatStart(mbPorts)
  '      ReDim mbStatStop(mbPorts)
  '      For i = 0 To mbPorts
  '          mbStatStart(i) = 0
  '          mbStatStop(i) = 0
  '      Next
  '      Call InitializeStats(mbPorts, mbStat())
        '
        '
        '
  '      ReDim metrabyte(Parameter.numMetrabyteBoards)
  '      Load frmMetrabyte
  '      If Parameter.numMetrabyteBoards > 1 Then
  '          For i = 1 To Parameter.numMetrabyteBoards - 1
  '             Load frmMetrabyte.mnuSetup(i)
  '              frmMetrabyte.mnuSetup(i).Caption = "Setup Board &" + Trim$(Str$(i + 1))
  '          Next
  '      End If
  '      frmMetrabyte.MSFlexGrid1.Rows = Parameter.numMetrabyteBoards * 24 + 1
  '      k = 1
  '      For i = 0 To Parameter.numMetrabyteBoards - 1
  '          For j = 0 To 23
  '              frmMetrabyte.MSFlexGrid1.TextMatrix(k, 0) = Trim$(Str$(i + 1)) + "-" + Trim$(Str$(j + 1))
  '              If metrabyte(i).port(j).Type = False Then
  '                  frmMetrabyte.MSFlexGrid1.TextMatrix(k, 1) = "Out "
  '              Else
  '                  frmMetrabyte.MSFlexGrid1.TextMatrix(k, 1) = "In "
  '              End If
  '              If metrabyte(i).port(j).usage = 0 Then
  '                  frmMetrabyte.MSFlexGrid1.TextMatrix(k, 1) = frmMetrabyte.MSFlexGrid1.TextMatrix(k, 1) + "Alarm"
  '              Else
  '                  frmMetrabyte.MSFlexGrid1.TextMatrix(k, 1) = frmMetrabyte.MSFlexGrid1.TextMatrix(k, 1) + "Page"
  '                  frmMetrabyte.MSFlexGrid1.TextMatrix(k, 3) = "Page Idle"
  '              End If
  '              k = k + 1
  '              metrabyte(i).port(j).alarmPreviousStatus = -1
  '              metrabyte(i).port(j).alarmStatus = -1
  '              metrabyte(i).port(j).pageAlarm = -1
  '              metrabyte(i).port(j).pageState = 0
  '              metrabyte(i).port(j).timer = Now
  '          Next
            Call InitializeBoard
            
'            If metrabyte(i).address > 0 Then
'                DlPortWritePortUchar metrabyte(i).address + 3, metrabyte(i).controlByte
'            End If
'            For j = 0 To 23
'                If metrabyte(i).port(j).Type = False Then
'                    If metrabyte(i).port(j).usage = 0 Then
'                        Call frmMetrabyte.MetrabyteSetPort(False, i, j)
'                    Else
'                        Call frmMetrabyte.MetrabyteSetPort(True, i, j)
'                    End If
'                End If
'            Next
'        Next
'        frmMetrabyte.Show 0
    End If
    
  Call GetPagingStatus(Parameter.pPath)
  
  gOperator = GetIniString("XN", "OperName", "Unknown", Parameter.pPath)
  Call XnOpenDataBase(Parameter.XnDataBase, username, XnPassword)
  gDBName = Parameter.XnDataBase
  gMsSQLUserID = username
  gMsSQLPassword = XnPassword
  
  For i = 1 To 3
    bFRQOpen = InitFrqTable(Parameter.pPath)
    If bFRQOpen Then Exit For
    Delay 300&
  Next
  
  If Not bFRQOpen Then
     If frmSimplex.Visible Then
        frmSimplex.lblWarning = "Cannot open frq table"
        frmSimplex.lblWarning.Visible = True
     End If
  Else

    
    SetAllTables
  End If
  Call GetNopaging(Parameter.pPath)
  ''OpenScheduleBtr Parameter.pPath
  
  Exit Sub
ExitMain:
    Unload ComSetup
    Unload Scope
    
    
    If XnFrqTable.state = adStateOpen Then
        XnFrqTable.Close
    End If
    
    CloseAllTables
    Call XnCloseDataBase
    End
    Exit Sub

MainError:
    MsgBox "Error initializing" & Error$, vbOKOnly
    Resume ExitMain
Resume
End Sub






Sub SaveParameters()
    Dim i As Integer
    Dim j As Integer
    Dim Temp As String
    Dim temp2 As String
    
    On Error GoTo OOPS
    
    If UCase(Parameter.alarmtype) = "DATAPROBE" Then
        Temp = Str$(ComMain.Top)
        i = WriteIniString("SCREEN", "ComMain_Top", Temp, Parameter.pPath)
        Temp = Str$(ComMain.Left)
        i = WriteIniString("SCREEN", "ComMain_Left", Temp, Parameter.pPath)
        Temp = Str$(ComMain.Height)
        i = WriteIniString("SCREEN", "ComMain_Height", Temp, Parameter.pPath)
        Temp = Str$(ComMain.Width)
        i = WriteIniString("SCREEN", "ComMain_Width", Temp, Parameter.pPath)
    End If
    
    Temp = Str$(Parameter.RectSetup.Top)
    i = WriteIniString("SCREEN", "ComSetup_Top", Temp, Parameter.pPath)
    Temp = Str$(Parameter.RectSetup.Left)
    i = WriteIniString("SCREEN", "ComSetup_Left", Temp, Parameter.pPath)
    Temp = Str$(Parameter.RectSetup.Height)
    i = WriteIniString("SCREEN", "ComSetup_Height", Temp, Parameter.pPath)
    Temp = Str$(Parameter.RectSetup.Width)
    i = WriteIniString("SCREEN", "ComSetup_Width", Temp, Parameter.pPath)
        
    '-----------frmSimplex ------------------------
    If UCase(Parameter.alarmtype) = "SIMPLEX" Then
        i = WriteIniString("SCREEN", "frmSimplex_Left", frmSimplex.Left, Parameter.pPath)
        i = WriteIniString("SCREEN", "frmSimplex_Top", frmSimplex.Top, Parameter.pPath)
        i = WriteIniString("SCREEN", "frmSimplex_Height", frmSimplex.Height, Parameter.pPath)
        i = WriteIniString("SCREEN", "frmSimplex_Width", frmSimplex.Width, Parameter.pPath)
    End If
    '---------frmAlarm -----------------
    i = WriteIniString("SCREEN", "frmAlarms_Left", frmAlarms.Left, Parameter.pPath)
    i = WriteIniString("SCREEN", "frmAlarms_Top", frmAlarms.Top, Parameter.pPath)
    i = WriteIniString("SCREEN", "frmAlarms_Height", frmAlarms.Height, Parameter.pPath)
    i = WriteIniString("SCREEN", "frmAlarms_Width", frmAlarms.Width, Parameter.pPath)
    
    '--------------frmSimplexList ---------------------
    i = WriteIniString("SCREEN", "frmSimplexList_Left", frmSimplexList.Left, Parameter.pPath)
    i = WriteIniString("SCREEN", "frmSimplexList_Top", frmSimplexList.Top, Parameter.pPath)
    i = WriteIniString("SCREEN", "frmSimplexList_Height", frmSimplexList.Height, Parameter.pPath)
    i = WriteIniString("SCREEN", "frmSimplexList_Width", frmSimplexList.Width, Parameter.pPath)
    '---------------------------------------------------------------
    Call Scope.ScopeSaveParams(Parameter.pPath)
    
    For i = 0 To Parameter.totalPorts - 1
        Temp = "ComPort" + Trim$(Str$(i + 1))
        temp2 = Trim(CStr(cport(i).port))
        j = WriteIniString("PORTS", Temp, temp2, Parameter.pPath)
        Temp = "ComSetup" + Trim$(Str$(i + 1))
        j = WriteIniString("PORTS", Temp, cport(i).setup, Parameter.pPath)
    Next
ExitHere:
Exit Sub

OOPS:
LogMessage ErrorMessage, "Error in SaveParameters, " & Error$
Resume ExitHere
End Sub



Sub InitializeBoard()
Const MYNAME = "InitializeBoard"
Dim i As Integer
Dim k As Integer
Dim j As Integer
On Error GoTo OOPS

mbStartStat = Hour(Now)
mbPorts = Parameter.numMetrabyteBoards * 24 - 1
ReDim mbStat(mbPorts)
ReDim mbStatStart(mbPorts)
ReDim mbStatStop(mbPorts)
For i = 0 To mbPorts
    mbStatStart(i) = 0
    mbStatStop(i) = 0
Next
Call InitializeStats(mbPorts, mbStat())
'
'
'
ReDim metrabyte(Parameter.numMetrabyteBoards)
Load frmMetrabyte
If Parameter.numMetrabyteBoards > 1 Then
    For i = 1 To Parameter.numMetrabyteBoards - 1
        Load frmMetrabyte.mnuSetup(i)
        frmMetrabyte.mnuSetup(i).Caption = "Setup Board &" + Trim$(Str$(i + 1))
    Next
End If
frmMetrabyte.MSFlexGrid1.Rows = Parameter.numMetrabyteBoards * 24 + 1
k = 1
For i = 0 To Parameter.numMetrabyteBoards - 1
    For j = 0 To 23
        frmMetrabyte.MSFlexGrid1.TextMatrix(k, 0) = Trim$(Str$(i + 1)) + "-" + Trim$(Str$(j + 1))
        If metrabyte(i).port(j).Type = False Then
            frmMetrabyte.MSFlexGrid1.TextMatrix(k, 1) = "Out "
        Else
            frmMetrabyte.MSFlexGrid1.TextMatrix(k, 1) = "In "
        End If
        If UCase(Parameter.alarmtype) = "KEITHLEY" Then
            Call gDriverLINX.SetupChannel(Chan(j), Trim(frmMetrabyte.MSFlexGrid1.TextMatrix(k, 1)))
        End If
        If metrabyte(i).port(j).usage = 0 Then
            frmMetrabyte.MSFlexGrid1.TextMatrix(k, 1) = frmMetrabyte.MSFlexGrid1.TextMatrix(k, 1) + "Alarm"
        Else
            frmMetrabyte.MSFlexGrid1.TextMatrix(k, 1) = frmMetrabyte.MSFlexGrid1.TextMatrix(k, 1) + "Page"
            frmMetrabyte.MSFlexGrid1.TextMatrix(k, 3) = "Page Idle"
        End If
        k = k + 1
        metrabyte(i).port(j).alarmPreviousStatus = -1
        metrabyte(i).port(j).alarmStatus = -1
        metrabyte(i).port(j).pageAlarm = -1
        metrabyte(i).port(j).pageState = 0
        metrabyte(i).port(j).timer = Now
    Next
    If metrabyte(i).address > 0 Then
        If UCase(Parameter.alarmtype) = "METRABYTE" Then
            DlPortWritePortUchar metrabyte(i).address + 3, metrabyte(i).controlByte
        End If
        If UCase(Parameter.alarmtype) = "KEITHLEY" Then
            Call gDriverLINX.WriteChannel(1, metrabyte(i).controlByte)  '??????
        End If
    End If
    For j = 0 To 23
        If metrabyte(i).port(j).Type = False Then
            If metrabyte(i).port(j).usage = 0 Then
                Call frmMetrabyte.MetrabyteSetPort(False, i, j)
            Else
                Call frmMetrabyte.MetrabyteSetPort(True, i, j)
            End If
        End If
    Next
Next
frmMetrabyte.Show 0


ExitHere:
Exit Sub

OOPS:
  MsgBox "[E] " & Err.Number & ": " & Err.Description & MODULENAME & MYNAME

Resume ExitHere
End Sub

