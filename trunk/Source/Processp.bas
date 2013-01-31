Attribute VB_Name = "SimplexProcessPacket"
Option Explicit


Private Sub SimplexAlarmNotify(parProfile As String, ByVal idx As Integer, parMsg As String, parNumMsg As String, ByVal s As String)
Dim iLoop As Integer
Dim temp2 As String
Dim p As Integer
Dim strPageLines() As String

On Error GoTo OOPS

    LogMessage SysMessage, parProfile & " Alpha : " & parMsg & "     Numeric : " & parNumMsg
    
    frmAlarms.AddToList "Condition:" & s & " alarm:" & simplex(idx).alarm & " Msg:" & parMsg
    
    If simplex(idx).ActionReminder Then
        Call XPutMessage(0, "AL", parProfile, parMsg, "Alarm: " & simplex(idx).alarm, "Condition: " & s, , , , , , , , , , gMessageDelivered)
    End If
    
    If simplex(idx).Page <> "FALSE" And simplex(idx).Page <> "" Then  ' added by TK 01/22/2004
        If simplex(idx).Page <> "TRUE" Then  ' if there is a profile id parameter in INI file - then page to this pager
            parProfile = simplex(idx).Page
        End If
        temp2 = parProfile
        If Parameter.BannerAvailable And simplex(idx).SendBanner Then gWriteBanner.AddToFile parMsg & " " & parNumMsg
        Call SplitIntoLines(parMsg, 120, "<-inued>", "<Cont->", strPageLines())
        For p = 0 To UBound(strPageLines)
            'If Parameter.BannerAvailable And simplex(idx).SendBanner Then gWriteBanner.AddToFile strPageLines(p) & " " & parNumMsg
            Call SendPage("XNALARM", parProfile, strPageLines(p), parNumMsg, "", Parameter.CheckStatus)
        Next
        If temp2 <> parProfile Then ' when status is covered by....
            iLoop = 0
            While iLoop < 10 And temp2 <> parProfile
                Call SplitIntoLines(parMsg, 120, "<-inued>", "<Cont->", strPageLines())
                For p = 0 To UBound(strPageLines)
                    'If Parameter.BannerAvailable And simplex(idx).SendBanner Then gWriteBanner.AddToFile strPageLines(p) & " " & parNumMsg
                    Call SendPage("XNALARM", parProfile, strPageLines(p), parNumMsg, "", Parameter.CheckStatus)
                Next
                iLoop = iLoop + 1
            Wend
        End If
    End If



ExitHere:
Exit Sub

OOPS:
LogMessage ErrorMessage, "Err: " & Err.Number & " " & Err.Description & " in SimplexAlarmNotify()"
Resume ExitHere

End Sub

Public Sub SimplexProccessPacket(tmpbuffer As String)
    Dim i As Integer
    Dim j As Integer
    Dim iLoop As Integer
    Dim searchresults As String
    Dim searchresults2 As String
    Dim x As String
    Dim s As String
    Dim Temp As String
    Dim temp2 As String
    Dim temp3 As String
    Dim logfilename As String
    Dim findpoint As String
        
    Static alarmMsg() As String
    Static countAlarm() As Integer
    Static thisProfile() As String
    Static alarmNumMsg() As String
    Static sCondition() As String
    
    On Error GoTo OOPS
    
    LogMessage SysMessage, "DATA: " & tmpbuffer
    simplexPackets = simplexPackets + 1
    frmSimplex.Label2.Caption = simplexPackets
    
    
    ReDim Preserve alarmMsg(Parameter.maxAlarmTypes)
    ReDim Preserve alarmNumMsg(Parameter.maxAlarmTypes)
    ReDim Preserve thisProfile(Parameter.maxAlarmTypes)
    ReDim Preserve countAlarm(Parameter.maxAlarmTypes)
    ReDim Preserve sCondition(Parameter.maxAlarmTypes)
        
    tmpbuffer = RemoveInvisiblechars(tmpbuffer)
    
    For i = 0 To Parameter.maxAlarmTypes - 1
        searchresults = InStr(tmpbuffer, simplex(i).alarm)   ' Find AlarmType
        If Mid$(tmpbuffer, simplex(i).alarmPosition, Len(simplex(i).alarm)) = simplex(i).alarm Then
            If simplex(i).conditionPosition > 0 Then
                s = Mid$(tmpbuffer, simplex(i).conditionPosition, simplex(i).conditionLenght)   ' Conditions
                searchresults2 = InStr(simplex(i).conditions, s)  ' Find AlarmConditions
                If searchresults2 Then
                    j = True
                Else
                    j = False
                   
                End If
            Else
                j = True
                s = "None"
            End If
            
            If j Then
                simplexStat(i).recType = "S" & Format(i + 1, "00")
                simplexStat(i).bucket(0) = simplexStat(i).bucket(0) + 1
                frmSimplex.Label4.Caption = simplexStat(i).bucket(0)
                x = Mid$(tmpbuffer, Parameter.pointPosition, Parameter.pointLength)
                LogMessage SysMessage, "Valid Alarm: " & simplex(i).alarm & "   " & simplex(i).conditions ' Point
                logfilename = Trim(Parameter.logPrefix)
                For iLoop = 1 To Len(x)
                    findpoint = Mid$(x, iLoop, 1)
                    If InStr(Parameter.pointExclude, findpoint) = 0 Then
                        logfilename = Trim(logfilename & findpoint)
                    End If
                Next iLoop
                
                Temp = Trim$(Parameter.alphaprefix) & " " & Trim$(Mid$(tmpbuffer, Parameter.alphamessageposition, Parameter.alphamessagelength))
                LogMessage SysMessage, "Alpha Message: " & Temp
                
                temp2 = Mid$(tmpbuffer, Parameter.numericmessageposition, Parameter.numericmessagelength)
                temp2 = Trim(Parameter.numericprefix & temp2) ' Only numeric message
                temp3 = ""
                For iLoop = 1 To Len(temp2)
                    If IsNumeric(Mid$(temp2, iLoop, 1)) Then
                        temp3 = temp3 + Mid$(temp2, iLoop, 1)
                    End If
                Next iLoop
                
                If simplex(i).AlarmLines = 1 Then
                    Call SimplexAlarmNotify(logfilename, i, Temp, temp3, s)
                Else
                    If simplex(i).ThisIsAlarm(countAlarm(i) + 1) Then
                        alarmMsg(i) = alarmMsg(i) & GetRidOfJunkChar(Temp)
                        alarmNumMsg(i) = alarmNumMsg(i) & temp3
                    End If
                    countAlarm(i) = countAlarm(i) + 1
                    thisProfile(i) = logfilename
                    sCondition(i) = s
                End If
                '============================================================
            End If      ' Alarm Condition
        Else
            ' Alarm not found ,.But it may belong to the previous alarm (see code below)
             If countAlarm(i) > 0 Then ' this is ALARM continues......
                countAlarm(i) = countAlarm(i) + 1
                If countAlarm(i) <= simplex(i).AlarmLines Then
                    If simplex(i).ThisIsAlarm(countAlarm(i)) Then alarmMsg(i) = alarmMsg(i) & " " & GetRidOfJunkChar(tmpbuffer) ' still collection data
                    If countAlarm(i) = simplex(i).AlarmLines Then
                        Call SimplexAlarmNotify(thisProfile(i), i, alarmMsg(i), alarmNumMsg(i), sCondition(i))
                        alarmMsg(i) = ""
                        countAlarm(i) = 0
                        thisProfile(i) = ""
                        alarmNumMsg(i) = ""
                        sCondition(i) = ""
                   End If
                End If
            End If
        End If          ' Alarm Found
    Next i              ' Alarm Type
ExitHere:
Exit Sub

OOPS:
LogMessage ErrorMessage, "Err: " & Err.Number & " " & Err.Description & " in SimplexProcessPacket()"
Resume ExitHere
Resume
End Sub

