Attribute VB_Name = "DATAPRB"
Option Explicit

Public DPChannels As New DPChannels

Public StartingProc As Boolean

Public DelayToStart As Integer
Sub InitDataProbe()
Dim Temp As String
Dim arrCmd() As String
Dim i As Integer

Const Default_CMD = "x1+01+x2+S+x3"  'Chr(1) & "01" & Chr(2) & "S" & Chr(3)

Temp = GetIniString("Data PROBE", "INIT_COMMAND", Default_CMD, Parameter.pPath, True)

arrCmd = Split(Temp, "+")
Temp = ""

For i = 0 To UBound(arrCmd)
    If LCase(Left(arrCmd(i), 1)) = "x" Then
        Temp = Temp & Chr(Right(arrCmd(i), 1))
    Else
        Temp = Temp & arrCmd(i)
    End If
Next

Call dpCheckSum(Temp)

ComMain!ComPort(0).Output = Temp

Call Scope.ScopeOutPut(0, Temp)
Call ShowData(frmDPInfo.txtTerm, Temp, vbBlack)
    
End Sub

Sub ReadDPChannels()
' Parameter.pPath   - ini file path and name
Dim i As Integer
Dim strPageNum As String
Dim strPageMsg As String
Dim strXnMsg As String
Dim strExt As String
Dim intAC As Integer
Dim intClearAL As Integer
Dim intOUTState As Integer
Dim intINPState As Integer
Dim intType As Integer
Dim bSendBanner As Boolean
Dim strKey As String


For i = 1 To 16
    strKey = CStr(i) & "_Type"
    intType = GetIniVal("DATA PROBE", strKey, 0, Parameter.pPath, True)
    
    strKey = CStr(i) & "_InputState"
    intINPState = GetIniVal("DATA PROBE", strKey, 0, Parameter.pPath, True)
    
    strKey = CStr(i) & "_OutputState"
    intOUTState = GetIniVal("DATA PROBE", strKey, 0, Parameter.pPath, True)
    
    strKey = CStr(i) & "_ClearAlarm"
    intClearAL = GetIniVal("DATA PROBE", strKey, 0, Parameter.pPath, True)
    
    strKey = CStr(i) & "_ActionReminder"
    intAC = GetIniVal("DATA PROBE", strKey, 0, Parameter.pPath, True)
    
    strKey = CStr(i) & "_Extension"
    strExt = GetIniString("DATA PROBE", strKey, "", Parameter.pPath, True)
    
    strKey = CStr(i) & "_XnMessage"
    strXnMsg = GetIniString("DATA PROBE", strKey, "", Parameter.pPath, True)
    
    strKey = CStr(i) & "_PageMessage"
    strPageMsg = GetIniString("DATA PROBE", strKey, "", Parameter.pPath, True)
    
    strKey = CStr(i) & "_PageNumeric"
    strPageNum = GetIniString("DATA PROBE", strKey, "", Parameter.pPath, True)
    
    strKey = CStr(i) & "_SendTicker"
    bSendBanner = GetIniBool("DATA PROBE", strKey, False, Parameter.pPath, True)
    DPChannels.Add strPageNum, strPageMsg, strXnMsg, strExt, intAC, intClearAL, intOUTState, intINPState, intType, bSendBanner
    
Next

End Sub

Public Sub SaveDPChannels()
Dim strPageNum As String
Dim strPageMsg As String
Dim strXnMsg As String
Dim strExt As String
Dim intAC As Integer
Dim intClearAL As Integer
Dim intOUTState As Integer
Dim intINPState As Integer
Dim intType As Integer
Dim x As Integer
Dim i As Integer
Dim strKey As String

On Error GoTo OOPS

For i = 1 To 16
    strKey = CStr(i) & "_Type"
    x = WriteIniString("DATA PROBE", strKey, DPChannels.Item(i).DPType, Parameter.pPath)
    
    strKey = CStr(i) & "_InputState"
    x = WriteIniString("DATA PROBE", strKey, DPChannels.Item(i).DPInputState, Parameter.pPath)
    
    strKey = CStr(i) & "_OutputState"
    x = WriteIniString("DATA PROBE", strKey, DPChannels.Item(i).DPOutputState, Parameter.pPath)
    
    strKey = CStr(i) & "_ClearAlarm"
    x = WriteIniString("DATA PROBE", strKey, DPChannels.Item(i).DPClearAlarm, Parameter.pPath)
    
    strKey = CStr(i) & "_ActionReminder"
    x = WriteIniString("DATA PROBE", strKey, DPChannels.Item(i).DPActionReminder, Parameter.pPath)
    
    strKey = CStr(i) & "_Extension"
    x = WriteIniString("DATA PROBE", strKey, DPChannels.Item(i).DPExtension, Parameter.pPath)
    
    strKey = CStr(i) & "_XnMessage"
    x = WriteIniString("DATA PROBE", strKey, DPChannels.Item(i).DPxnMessage, Parameter.pPath)
    
    strKey = CStr(i) & "_PageMessage"
    x = WriteIniString("DATA PROBE", strKey, DPChannels.Item(i).DPPageMessage, Parameter.pPath)
    
    strKey = CStr(i) & "_PageNumeric"
    x = WriteIniString("DATA PROBE", strKey, DPChannels.Item(i).DPPageNumeric, Parameter.pPath)
           
    strKey = CStr(i) & "_SendTicker"
    x = WriteIniString("DATA PROBE", strKey, DPChannels.Item(i).dpSendBanner, Parameter.pPath)
        
Next

ExitHere:
Exit Sub

OOPS:
LogMessage ErrorMessage, "Error in SaveDPChannels, " & Error$
Resume ExitHere
End Sub

Public Static Sub ShowData(Term As Control, Data As String, ByVal parColor As Long)
    On Error GoTo Handler
    Const MAXTERMSIZE = 16000
    Dim TermSize As Long, i As Integer
    Dim x As String
    
    ' Make sure the existing text doesn't get too large.
    If Len(Term.Text) > 850 Then Term.Text = ""
    
    TermSize = Len(Term.SelText)
    Term.SelColor = parColor
  
    Term.SelStart = Len(Term.Text)

    ' Add the filtered data to the SelText property.
    For i = 1 To Len(Data)
        x = Mid(Data, i, 1)
        If Asc(x) < 32 Then
            Term.SelText = "[0x" & Hex(Asc(x)) & "]"
        Else
            Term.SelText = x
        End If
    Next
    '''''' Term.SelText = Data  ' see the loop above to replace this line
  
    Term.SelStart = Len(Term.Text)
Exit Sub

Handler:
    MsgBox Error$
    Resume Next
End Sub

Function ConvertComData(parInputData As String, arrBinary() As String) As String
Dim iDec As Integer
Dim i As Integer
Dim sBinary As String
Dim binLine As String

For i = 1 To 4
    iDec = Asc(Mid(parInputData, i, 1))
    sBinary = GetBinary(iDec)
    Select Case i
        Case 1  ' 5,6,7,8
            arrBinary(5 - 1) = Mid(sBinary, 4, 1)
            arrBinary(6 - 1) = Mid(sBinary, 3, 1)
            arrBinary(7 - 1) = Mid(sBinary, 2, 1)
            arrBinary(8 - 1) = Mid(sBinary, 1, 1)
        Case 2  ' 1,2,3,4
            arrBinary(1 - 1) = Mid(sBinary, 4, 1)
            arrBinary(2 - 1) = Mid(sBinary, 3, 1)
            arrBinary(3 - 1) = Mid(sBinary, 2, 1)
            arrBinary(4 - 1) = Mid(sBinary, 1, 1)
        Case 3 ' 13,14,15,16
            arrBinary(13 - 1) = Mid(sBinary, 4, 1)
            arrBinary(14 - 1) = Mid(sBinary, 3, 1)
            arrBinary(15 - 1) = Mid(sBinary, 2, 1)
            arrBinary(16 - 1) = Mid(sBinary, 1, 1)
        Case 4  '9,10,11,12
            arrBinary(9 - 1) = Mid(sBinary, 4, 1)
            arrBinary(10 - 1) = Mid(sBinary, 3, 1)
            arrBinary(11 - 1) = Mid(sBinary, 2, 1)
            arrBinary(12 - 1) = Mid(sBinary, 1, 1)
    End Select
Next
For i = 0 To 15
    binLine = binLine & arrBinary(i)
Next
ConvertComData = binLine
End Function

Sub dpCheckAlarmChange()
    Dim i As Integer
    Dim LogPATH  As String
    Dim msg1 As String
    Dim msg2 As String
    Dim msg3 As String
    'Dim startTime() As String * 4
    'Dim stopTime() As String * 4
    'Dim duration As String * 4

    On Error GoTo checkalarmerror
   ' ReDim startTime(maxport)
    'ReDim stopTime(maxport)
    
    'If Not StartingProc Then Exit Sub ' time is not up yet...
    
    For i = 0 To 15
        'if And dpPort(i).alarmPreviousStatus <> -1
        If dpPort(i).alarmStatus <> -1 _
            And dpPort(i).alarmStatus <> dpPort(i).alarmPreviousStatus And _
            DPChannels.Item(i + 1).DPExtension <> "" And DPChannels.Item(i + 1).DPType = 0 And dpPort(i).alarmStatus = 0 Then
            
            msg1 = ""
            msg2 = ""
            msg3 = ""
            LogPATH = GetLogPath(dpPort(i).Extension, "LOG", Parameter.LogPATH)
            Select Case dpPort(i).alarmStatus
            
            Case 0
                If DPChannels.Item(i + 1).DPInputState = 0 Then
                    msg1 = "Alarm " + Str$(i + 1) + " On"
                    dpStartTime(i) = Now
                    alarmStat(i).recType = "A" & Format(i + 1, "00")
                    alarmStat(i).bucket(0) = alarmStat(i).bucket(0) + 1   'AD
                Else
                    msg1 = "Alarm " + Str$(i + 1) + " Off"
                    stopTime(i) = Now
                End If
            Case Else
                If DPChannels.Item(i + 1).DPInputState = 0 Then
                    msg1 = "Alarm " + Str$(i + 1) + " Off"
                    stopTime(i) = Now
                Else
                    msg1 = "Alarm " + Str$(i + 1) + " On"
                    dpStartTime(i) = Now
                    alarmStat(i).recType = "A" & Format(i + 1, "00")
                    alarmStat(i).bucket(0) = alarmStat(i).bucket(0) + 1    'AD
                End If
            End Select
            If dpStartTime(i) > 0 Then
                If stopTime(i) > 0 Then
                   If stopTime(i) > dpStartTime(i) Then
                      alarmStat(i).recType = "A" & Format(i + 1, "00")
                      alarmStat(i).bucket(1) = alarmStat(i).bucket(1) + DateDiff("s", dpStartTime(i), stopTime(i))
                      dpStartTime(i) = 0
                      stopTime(i) = 0
                   End If
                End If
            End If
            
            If StartingProc Then
                Call XPutMessage(DPChannels.Item(i + 1).DPActionReminder, "AL", DPChannels.Item(i + 1).DPExtension, msg1, DPChannels.Item(i + 1).DPxnMessage, msg3, , , , , , , , , , gMessageDelivered)
                LogMessage SysMessage, msg1
                If DPChannels.Item(i + 1).DPPageMessage <> "" Then
                    msg1 = DPChannels.Item(i + 1).DPPageMessage
                    msg2 = DPChannels.Item(i + 1).DPPageNumeric
                    If Parameter.BannerAvailable And DPChannels.Item(i + 1).dpSendBanner Then gWriteBanner.AddToFile msg1 & " " & msg2
                    Call SendPage("XNALARM", DPChannels.Item(i + 1).DPExtension, msg1, msg2, "", Parameter.CheckStatus)
                End If
            End If
        End If
    Next

exitcheckalarm:
Exit Sub

checkalarmerror:

LogMessage ErrorMessage, "Unexpected error in DpCheckAlarmChange " & Error$
Resume exitcheckalarm
Resume
End Sub

Sub dpParsePacket(dpCommandType As Integer, dpCommand As String, dpData As String)
    Dim i As Integer
    Dim j As Integer
    Dim iDec As Integer
    Dim sBinary As String
    Dim strACK As String
    Dim bACKSent As Boolean
    ReDim binstatus(16) As String
    
    Static bNextTime As Boolean
    
    On Error GoTo parsepacketerror
    
    'If bNextTime = False Then
    '    bNextTime = True
    '    Exit Sub
    'End If
    'If Len(dpData) <= 4 Then Exit Sub
    'If dpCommandType = 0 Then
        If dpCommand = "W" Or dpCommand = "S" Then
        
        
            For j = 0 To 15
                dpPort(j).alarmPreviousStatus = dpPort(j).alarmStatus
            Next
            sBinary = ConvertComData(dpData, binstatus())
            For i = 0 To 15: dpPort(i).alarmStatus = binstatus(i): Next
            Call dpCheckAlarmChange  ' send page and save message to Console Database
    
    
    
            bACKSent = False
    

            For i = 0 To 15
                If dpPort(i).alarmStatus <> -1 Then
                    If dpPort(i).alarmStatus = 0 Then
                        ComMain!RedLed(i).FillColor = &HFF&
                        ComMain!GreenLed(i).FillColor = &HC0FFC0
                    Else
                        ComMain!RedLed(i).FillColor = &HC0C0FF
                        ComMain!GreenLed(i).FillColor = &HFF00&
                    End If
                End If
            Next
    End If
'End If

exitparsepacket:
Exit Sub

parsepacketerror:

LogMessage ErrorMessage, "Unexpected error in DpParsePacket " & Error$
Resume exitparsepacket

End Sub

Public Function GetBinary(parValue As Integer) As String
Dim lCount As Long, lConvert As Long, sBinary As String
Dim iMaxpower As Integer
iMaxpower = 6

sBinary = ""
lConvert = parValue

sBinary = sBinary + "0"

For lCount = iMaxpower To 0 Step -1
    If lConvert And (2 ^ lCount) Then 'Use the logical "AND" operator...
        sBinary = sBinary + "1"
    Else
        sBinary = sBinary + "0"
    End If
Next

GetBinary = Right(sBinary, 4) 'The bin string contains the binary number...

End Function

Sub dpProcessPacket(strparam As String)
    Dim i As Integer
    Dim j As Integer
    Dim dpCommandType As Integer    ' 0 message, 1 ACK,
                                    ' 2 NACK
    Dim dpCommand As String
    Dim dpData As String
    Dim Temp As String

    On Error GoTo processpacketerror

    j = 0
    dpData = ""
    For i = 1 To Len(strparam)
        Temp = Mid$(strparam, i, 1)
        Select Case j
            Case 0
                Select Case Asc(Temp)
                    Case 1
                        dpCommandType = 0  ' message
                    Case 2
                        j = 1
                    Case 6
                        dpCommandType = 1   'ACK
                    Case 21
                        dpCommandType = 2  ' NAC
                End Select
            Case 1
                dpCommand = Temp
                j = 2
            Case 2
                If Asc(Temp) <> 3 Then
                    dpData = dpData + Temp
                Else
                    j = 3
                End If
            Case 3
                Call dpParsePacket(dpCommandType, dpCommand, dpData)
                'Call PressACK(dpCommand, dpData)
                j = 0
            Case Else
                j = 0
        End Select
    Next

exitprocesspacket:
Exit Sub

processpacketerror:

LogMessage ErrorMessage, "Unexpected error in DpProcessPacket " & Error$
Resume exitprocesspacket

End Sub

Sub dpSetRelayOff()
    Dim i As Integer
    Dim dpData As String * 4
    Dim Temp As String
    
    On Error GoTo setrelayofferror

    i = INVERTBOOL(dpPort(0).relayState)
    i = i + INVERTBOOL(dpPort(1).relayState) * 2
    i = i + INVERTBOOL(dpPort(2).relayState) * 4
    i = i + INVERTBOOL(dpPort(3).relayState) * 8
    Mid$(dpData, 2, 1) = Chr$(i + 48)

    i = INVERTBOOL(dpPort(4).relayState)
    i = i + INVERTBOOL(dpPort(5).relayState) * 2
    i = i + INVERTBOOL(dpPort(6).relayState) * 4
    i = i + INVERTBOOL(dpPort(7).relayState) * 8
    Mid$(dpData, 1, 1) = Chr$(i + 48)

    i = INVERTBOOL(dpPort(8).relayState)
    i = i + INVERTBOOL(dpPort(9).relayState) * 2
    i = i + INVERTBOOL(dpPort(10).relayState) * 4
    i = i + INVERTBOOL(dpPort(11).relayState) * 8
    Mid$(dpData, 4, 1) = Chr$(i + 48)

    i = INVERTBOOL(dpPort(12).relayState)
    i = i + INVERTBOOL(dpPort(13).relayState) * 2
    i = i + INVERTBOOL(dpPort(14).relayState) * 4
    i = i + INVERTBOOL(dpPort(15).relayState) * 8
    Mid$(dpData, 3, 1) = Chr$(i + 48)

    Temp = Chr$(1) + "01" + Chr$(2) + "C" + dpData + Chr$(3)
    Call dpCheckSum(Temp)
    ComMain!ComPort(0).Output = Temp
    Call Scope.ScopeOutPut(0, Temp)
    Call ShowData(frmDPInfo.txtTerm, Temp, vbBlack)

exitsetrelayoff:
Exit Sub

setrelayofferror:

LogMessage ErrorMessage, "Unexpected error in DpSetRelayOff " & Error$
Resume exitsetrelayoff

End Sub

Sub dpSetRelayOn()
    Dim i As Integer
    Dim dpData As String * 4
    Dim Temp As String
    
    On Error GoTo setrelayerror

    i = dpPort(0).relayState
    i = i + dpPort(1).relayState * 2
    i = i + dpPort(2).relayState * 4
    i = i + dpPort(3).relayState * 8
    Mid$(dpData, 2, 1) = Chr$(i + 48)

    i = dpPort(4).relayState
    i = i + dpPort(5).relayState * 2
    i = i + dpPort(6).relayState * 4
    i = i + dpPort(7).relayState * 8
    Mid$(dpData, 1, 1) = Chr$(i + 48)

    i = dpPort(8).relayState
    i = i + dpPort(9).relayState * 2
    i = i + dpPort(10).relayState * 4
    i = i + dpPort(11).relayState * 8
    Mid$(dpData, 4, 1) = Chr$(i + 48)

    i = dpPort(12).relayState
    i = i + dpPort(13).relayState * 2
    i = i + dpPort(14).relayState * 4
    i = i + dpPort(15).relayState * 8
    Mid$(dpData, 3, 1) = Chr$(i + 48)

    Temp = Chr$(1) + "01" + Chr$(2) + "W" + Chr$(3)
    Call dpCheckSum(Temp)
    ComMain!ComPort(0).Output = Temp
    Call Scope.ScopeOutPut(0, Temp)
    Call ShowData(frmDPInfo.txtTerm, Temp, vbBlack)
exitsetrelay:
Exit Sub

setrelayerror:

LogMessage ErrorMessage, "Unexpected error in DpSetRelayOn " & Error$
Resume exitsetrelay

End Sub

Function INVERTBOOL(param1 As Integer) As Integer
    If param1 = 0 Then
        INVERTBOOL = 1
    Else
        INVERTBOOL = 0
    End If
End Function

Sub dpCheckSum(ByRef strparam As String)
    Dim total As Integer
    Dim i As Integer
    Dim chrctr As Integer
    Dim NewString As String

    On Error GoTo checksumerror

    NewString = ""
    i = 1
    total = 0
    Do
        NewString = NewString + Mid$(strparam, i, 1)
        chrctr = Asc(Mid$(strparam, i, 1))
        If chrctr >= 32 Then
            total = total + chrctr - 32
        Else
            total = total + chrctr
        End If
        If chrctr = 3 Then Exit Do
        
        i = i + 1
    Loop
    NewString = NewString + Chr$(((total + 184) Mod 95) + 32)
    strparam = NewString

exitchecksum:
Exit Sub

checksumerror:

LogMessage ErrorMessage, "Unexpected error in DpCheckSum " & Error$
Resume exitchecksum

End Sub



Sub ShowDPChanInfo(parChannel As Integer)
With DpSetup

    .cmbType.ListIndex = DPChannels.Item(parChannel).DPType
    .cmbAction.ListIndex = DPChannels.Item(parChannel).DPActionReminder
    .cmbClearAlarm.ListIndex = DPChannels.Item(parChannel).DPClearAlarm
    .cmbInputState.ListIndex = DPChannels.Item(parChannel).DPInputState
    .cmbOutputState.ListIndex = DPChannels.Item(parChannel).DPOutputState
    .cmbSendBanner.Text = IIf(DPChannels.Item(parChannel).dpSendBanner, "Yes", "No")
    .txtExtension = DPChannels.Item(parChannel).DPExtension
    .txtPageMessage.Text = DPChannels.Item(parChannel).DPPageMessage
    .txtPageMessage.Refresh
    .txtPageNumeric = DPChannels.Item(parChannel).DPPageNumeric
    .txtXnMessage.Text = DPChannels.Item(parChannel).DPxnMessage
    .txtXnMessage.Refresh
End With

End Sub


