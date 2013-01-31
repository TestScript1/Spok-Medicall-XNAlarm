Attribute VB_Name = "XnQue"
Option Explicit

Const FromLock = 11
Const ToLock = 20
Const PageQuePrefix = "NOPG"
Const PageQueSuffix = ".QUE"
'Const Pagetype = "54"
Const WaitForLock = 2

Dim QueRecordLen As Long
Dim PointErrecLen  As Long
Dim ListPage As String  ' %xx - indicates : this is list page (page Numeric+alpha)

Type Pt
    GetPointer As Integer
    Temp1 As Integer
    PError1 As Integer
    PError2 As Integer
    PType As String * 2
    PutPointer As Integer
    temp2 As Integer
    Junk As String * 6
End Type
Dim PointErrec As Pt

Type PageQueRecord
    PType As String * 2
    PStatus As String * 2
    PDatein As String * 10
    PTimein As String * 8
    PDateout As String * 10
    PTimeout As String * 8
    PExtension As String * 8
    PExtid As String * 7
    PIdin As String * 10
    Packed As String * 1
    Packtime As Integer
    PPointer As Integer
    PPrinted As String * 1
    PVoice As String * 5
    PVoicef As Integer
    PInfo As String * 148
End Type
Dim QueRecord As PageQueRecord

Type PageQueBtrRecord
    interfaceType As String * 10    '   Used to identify interface used to send page
    priority As String * 2          '   0 highest, 99 lowest
    initiatedDate As String * 10
    initiatedTime As String * 8
    initiatedId As String * 10
    ProfileID As String * 10
    PagerId As String * 10
    PageType As String * 10         '   Identifies type of pager
    voiceFlag As String * 1         '   Y indicates voice page
    voiceFormat As String * 15      '   Encoding format such as wav, dialogic, ect. Blank idicates default
    filename As String * 120        '   filename of text to page or voice file to page
    pageInfo As String * 300        '   page information + message. As used in queue file inteface
    reserved As String * 194        '   reserved for future expansion
End Type
Dim QueBtrRecord As PageQueBtrRecord
    
Type NoPgStat                       'Ana Del Campo  05/21/96
    Identifier     As Integer       'Structure to Read the NOPGSTAT File
    Description    As String * 30
    PagingAllowed  As Boolean
    Page           As String
End Type
Public statuslist()     As NoPgStat

Type NoPagingType
    MsgType As Integer      ' 0 = Numeric  1 = Alpha
    Source(3) As Integer    ' 0 Unused, 1 Script,
                            ' 2 Frq Data, 3 Input
    FieldNum As Integer
    PageType As String
    Alias As String
    QueDirect As String
    Prefix As String
    Sufix As String
    Number As String
    Pin As String
    Description As String       'Ana Del Campo   05/21/96
    CallBackNumber As String
End Type
Global PagingType(40)   As NoPagingType


Private FrqTypes As String ' big line of ALL frq type devided by _ : %04_%02_%41,....
Private NoPagingFilePath As String
Private pageQueueBtrvTable As String
Private PagStatFile As String
Global NoPagingStatus As Integer
Private QuePath As String
Private ActionReminder As String  '''' "Y"   or , "N"
Private ActionReminderFlagOffset As Integer
Private PageFlagOffset  As Integer  ' number of the flags from long list of YNFlagX ( X-125)  PageFlagOffset= from 1 to 25

'
'   PageSetup:      Parses paging information into a final paging string
'
'
'
'
'
Sub PageSetup(Message As String, mLine1 As String, mLine2 As String)
    
    Dim i As Integer
    Dim j As Integer
    Dim Temp As String
    Static Part(3) As String
    Dim PageNumber As String
    Dim PagePin As String
    Dim pageMessage As String
    Dim NewMessage As String

    On Error GoTo PageSetupError

    NewMessage = ""
    mLine1 = ""
    mLine2 = ""
    Debug.Print "FRQ Type", XnFrqTable!FrqType
    For i = 0 To 39
        If XnFrqTable!FrqType = PagingType(i).PageType Then
            Part(0) = "": Part(1) = "": Part(2) = ""
            Temp = Trim$(XnFrqTable!frqnumber)
            j = InStr(Temp, "+")
            If j <> 0 Then
                Part(0) = Left$(Temp, j - 1)
                Temp = Right$(Temp, Len(Temp) - j)
                j = InStr(Temp, "+")
                If j <> 0 Then
                    Part(1) = Left$(Temp, j - 1)
                    Part(2) = Right$(Temp, Len(Temp) - j)
                Else
                    Part(1) = Temp
                End If
            Else
                Part(0) = Temp
            End If
            j = 0
            mLine2 = Message
            Select Case PagingType(i).Source(0)
                Case 1
                    NewMessage = PagingType(i).Prefix + PagingType(i).Number + PagingType(i).Sufix
                    mLine1 = NewMessage
                Case 2
                    NewMessage = PagingType(i).Prefix + Part(j) + PagingType(i).Sufix
                    j = j + 1
                    mLine1 = NewMessage
                Case 3
                    NewMessage = Message
            End Select
            Select Case PagingType(i).Source(1)
                Case 1
                    If NewMessage <> "" Then NewMessage = NewMessage + "+"
                    NewMessage = NewMessage + PagingType(i).Pin
                    mLine1 = NewMessage
                Case 2
                    If NewMessage <> "" Then NewMessage = NewMessage + "+"
                    NewMessage = NewMessage + Part(j)
                    j = j + 1
                    mLine1 = NewMessage
                Case 3
                    If NewMessage <> "" Then NewMessage = NewMessage + "+"
                    NewMessage = NewMessage + Message
            End Select
            Select Case PagingType(i).Source(2)
                Case 1
                    If NewMessage <> "" Then NewMessage = NewMessage + "+"
                    NewMessage = NewMessage + PagingType(i).CallBackNumber
                    mLine2 = PagingType(i).CallBackNumber
                Case 2
                    If NewMessage <> "" Then NewMessage = NewMessage + "+"
                    NewMessage = NewMessage + Part(j)
                    j = j + 1
                    mLine2 = Part(j)
                Case 3
                    If NewMessage <> "" Then NewMessage = NewMessage + "+"
                    NewMessage = NewMessage + Message
            End Select
            Exit For
        End If
    Next
    Message = NewMessage
ExitPageSetup:
Exit Sub

PageSetupError:

LogMessage ErrorMessage, "Unexpected error in PageSetup " & Error$
Resume ExitPageSetup

End Sub

Sub WritePageQue(Initiator As String, Extension As String, logBuffer As LogInfoType, Message As String)

    'If pageQueueBtrvTable = "" Then
        Call WritePageQueFile(Initiator, Extension, logBuffer, Message)
    'Else
    '    Call WritePageQueBtr(Initiator, Extension, logBuffer, Message)
    'End If
End Sub

Sub WritePageQueFile(Initiator As String, Extension As String, logBuffer As LogInfoType, Message As String)

    Dim filename As String
    Dim fileNum As Integer
    Dim waittime As Variant
    Dim scrollflag As Integer
    Dim PutPageQuePointer As Long
    Dim i As Integer

5:  QueRecordLen = Len(QueRecord)
    PointErrecLen = Len(PointErrec)

    On Error GoTo WritePageQueError

    fileNum = FreeFile
    
    filename = ""
    For i = 0 To 39
        If logBuffer.PageType = PagingType(i).PageType Then
            filename = QuePath & "\" & PageQuePrefix & Right$(PagingType(i).QueDirect, 2) & PageQueSuffix
            Exit For
        End If
    Next

    If filename = "" Then
        filename = QuePath & "\" & PageQuePrefix & Right$(logBuffer.PageType, 2) & PageQueSuffix
    End If
    LogMessage SysMessage, "Writing page queue file: " & filename

    waittime = DateAdd("s", WaitForLock, Now)

10: Open filename For Binary Shared As fileNum

    If LOF(fileNum) < 20 Then
        PointErrec.GetPointer = 0
        PointErrec.PError1 = 0
        PointErrec.PError2 = 0
        PointErrec.PType = "1 "
'        pointerrec.putpointer = 1
        PointErrec.PutPointer = 0
        PointErrec.Junk = String$(Len(PointErrec.Junk), " ")
        Put #fileNum, 1, PointErrec
    End If

15: QueRecord.PType = Right$(logBuffer.PageType, 2)
    QueRecord.PStatus = "P "
    QueRecord.PDatein = Date$
    QueRecord.PTimein = Time$
    QueRecord.PExtension = Trim$(Extension)
    QueRecord.PIdin = Initiator
    QueRecord.PInfo = Message

    waittime = DateAdd("s", WaitForLock, Now)

25: Lock #fileNum, FromLock To ToLock
26: Get #fileNum, 1, PointErrec

' JS   putpagequepointer = pointerreclen + 1 + ((pointerrec.putpointer - 1) * querecordlen)
    LogMessage SysMessage, "Pointer Rec Len: " & PointErrecLen & " Put Pointer: " & PointErrec.PutPointer & " Que Record Len: " & QueRecordLen
27: PutPageQuePointer = PointErrecLen + 1 + ((PointErrec.PutPointer) * QueRecordLen)

30: Put #fileNum, PutPageQuePointer, QueRecord
    PointErrec.PutPointer = PointErrec.PutPointer + 1
    Put #fileNum, 1, PointErrec
ex2:
    Unlock #fileNum, FromLock To ToLock
ex1:
    Close fileNum
ex:
Exit Sub

WritePageQueError:

If Err = 70 And Erl = 10 Then
    If waittime > Now Then
        Resume 10
    Else
        LogMessage ErrorMessage, "Can't access page queue file: " & filename & " in " & WaitForLock & " seconds."
        SendEmailNote "Can't access page queue file: " & filename & " in " & WaitForLock & " seconds."
        Resume ex
    End If
End If

If Err = 70 And Erl = 25 Then
    If waittime > Now Then
        Resume 25
    Else
        LogMessage ErrorMessage, "Can't lock page queue file: " & filename & " in " & WaitForLock & "seconds."
        SendEmailNote "Can't lock page queue file: " & filename & " in " & WaitForLock & "seconds."
        Resume ex1
    End If
End If

LogMessage ErrorMessage, "Unexpected error in WRITEPAGEQUE " & Erl & ":" & Error$
SendEmailNote "Unexpected error in WRITEPAGEQUE " & Erl & ":" & Error$
If Erl <= 25 Then
    Resume ex1
ElseIf Erl > 25 Then
    Resume ex2
Else
    Resume ex
End If

End Sub



Sub SendPage(xnModual As String, Extension As String, pageMessage As String, pageNumeric As String, PagerType As String, CheckStatus As Integer)

    Dim pagemsgType As Integer          ' 0 = Numeric, 1=alpha , 2=listpage
    Dim strPageNumber As String
    Dim LogPATH As String
    Dim msgline1 As String
    Dim msgLine2 As String
    Dim extensionTmp As String
    Dim logBuffer As LogInfoType
    Dim Temp As String
    Dim i As Integer

    On Error GoTo sendpageerror
    
    extensionTmp = Extension
    
    LogMessage SysMessage, "Paging extension: " & extensionTmp + "-" + PagerType
    'LogPATH = GetLogPath(Extension, "LOG", XN_LogPATH)
    
    Call GetLogInfo(LogPATH, extensionTmp, logBuffer, i, "1")
    
    LogPATH = "PG" ' message type  ' added by TK
    If CheckStatus > 0 Then
        Temp = GetRidOfJunkChar(logBuffer.coverExtension)  ' remove junk chars
        If Temp <> "" Then
            Temp = "Covered By: " & Trim(logBuffer.coverExtension)
            Call XPutMessage(1, LogPATH, extensionTmp, msgline1, msgLine2, Temp, , , , , , , , , , gMessageDelivered)
            extensionTmp = Trim$(logBuffer.coverExtension)
            Exit Sub
        End If
    End If
    i = Asc(logBuffer.pageStatus) - 48
    If Not statuslist(i).PagingAllowed And CheckStatus > 0 Then
        Select Case statuslist(i).Page
            Case "CVR"
                'finds the covering extension
                Temp = "Covered By: " & Trim(logBuffer.coverExtension)
                Call XPutMessage(1, LogPATH, extensionTmp, msgline1, msgLine2, Temp, , , , , , , , , , gMessageDelivered)
                extensionTmp = Trim$(logBuffer.coverExtension)
                Exit Sub
            Case "FRD"
                Temp = ""
                Call PageSetup(pageMessage, msgline1, msgLine2)
                Call XPutMessage(1, LogPATH, extensionTmp, msgline1, msgLine2, Temp, , , , , , , , , , gMessageDelivered)
                Exit Sub
            Case "FDP"
                            
            Case "MSG"
                Temp = ""
                Call PageSetup(pageMessage, msgline1, msgLine2)
                Call XPutMessage(1, LogPATH, extensionTmp, msgline1, msgLine2, Temp, , , , , , , , , , gMessageDelivered)
                Exit Sub
            Case "VFD"
                Temp = ""
                Call PageSetup(pageMessage, msgline1, msgLine2)
                Call XPutMessage(1, LogPATH, extensionTmp, msgline1, msgLine2, Temp, , , , , , , , , , gMessageDelivered)
                Exit Sub
            Case "0FD"
                Temp = ""
                Call PageSetup(pageMessage, msgline1, msgLine2)
                Call XPutMessage(1, LogPATH, extensionTmp, msgline1, msgLine2, Temp, , , , , , , , , , gMessageDelivered)
                Exit Sub
            Case "0F0"
                Temp = ""
                Call PageSetup(pageMessage, msgline1, msgLine2)
                Call XPutMessage(1, LogPATH, extensionTmp, msgline1, msgLine2, Temp, , , , , , , , , , gMessageDelivered)
                Exit Sub
            Case "#FD"
                Temp = ""
                Call PageSetup(pageMessage, msgline1, msgLine2)
                Call XPutMessage(1, LogPATH, extensionTmp, msgline1, msgLine2, Temp, , , , , , , , , , gMessageDelivered)
                Exit Sub
        End Select
    End If
    
    If PagerType <> "" Then logBuffer.PageType = PagerType   ' Ana Del Campo   05/20/96
    If GetFrq(extensionTmp, logBuffer.PageType) = True Then
        If logBuffer.PageType <> "" Then
            Call GetPageType(pagemsgType, logBuffer.PageType)
            Temp = ""
            i = 1
            If ActionReminder = "Y" Then
              i = 0
            End If
            
            If ActionReminderFlagOffset > 0 Then
              If logBuffer.ynFlags(ActionReminderFlagOffset - 1) = "Y" Then
                i = 0
              End If
            End If
            
            Select Case pagemsgType
                Case 0 ' ==========NUMERIC PAGE ===============
                    Call PageSetup(pageNumeric, msgline1, msgLine2)
                    If PageFlagOffset = 0 Then
                        Call WritePageQue(xnModual, extensionTmp, logBuffer, pageNumeric)
                    ElseIf PageFlagOffset > 0 Then
                        If logBuffer.ynFlags(PageFlagOffset - 1) = "Y" Then
                            Call WritePageQue(xnModual, extensionTmp, logBuffer, pageNumeric)
                        End If
                    End If
                    Call XPutMessage(i, LogPATH, extensionTmp, msgline1, msgLine2, Temp, , , , , , , , , , gMessageDelivered)
                    
                Case 1 '========== ALPHA PAGE ===============
                    Call PageSetup(pageMessage, msgline1, msgLine2)
                    If PageFlagOffset = 0 Then
                        Call WritePageQue(xnModual, extensionTmp, logBuffer, pageMessage)
                    ElseIf PageFlagOffset > 0 Then
                        If logBuffer.ynFlags(PageFlagOffset - 1) = "Y" Then
                            Call WritePageQue(xnModual, extensionTmp, logBuffer, pageMessage)
                        End If
                    End If
                    Call XPutMessage(i, LogPATH, extensionTmp, msgline1, msgLine2, Temp, , , , , , , , , , gMessageDelivered)
                    
                Case 2   ' ======== LIST PAGE ==================
                    pageMessage = pageNumeric & "+" & pageMessage
                    If PageFlagOffset = 0 Then
                        Call WritePageQue(xnModual, extensionTmp, logBuffer, pageMessage)
                    ElseIf PageFlagOffset > 0 Then
                        If logBuffer.ynFlags(PageFlagOffset - 1) = "Y" Then
                            Call WritePageQue(xnModual, extensionTmp, logBuffer, pageMessage)
                        End If
                    End If
                    msgline1 = pageMessage
                    Call XPutMessage(i, LogPATH, extensionTmp, msgline1, msgLine2, Temp, , , , , , , , , , gMessageDelivered)
                    
            End Select
        Else
            LogMessage ErrorMessage, "Page Info not found for extension: " & extensionTmp
        End If
    Else
        LogMessage ErrorMessage, "FRQ Info " & logBuffer.PageType & " not found for extension: " & extensionTmp
    End If
exitsendpage:
Exit Sub

sendpageerror:

LogMessage ErrorMessage, "Unexpected error in SENDPAGE " & Error$
Resume exitsendpage
Resume
End Sub

Sub SendWakeup(xnModual As String, Extension As String, pageMessage As String, filename As String, WakeType As String)

    Dim pagemsgType As Integer          ' 0 = Numeric, 1=alpha
    Dim strPageNumber As String
    Dim LogPATH As String
    Dim msgline1 As String
    Dim msgLine2 As String
    Dim extensionTmp As String
    Dim logBuffer As LogInfoType
    Dim Temp As String
    Dim i As Integer

    On Error GoTo sendpageerror
    
    extensionTmp = Extension
    
    LogMessage SysMessage, "Wakeup extension: " & extensionTmp + "-" + WakeType
    
    LogPATH = "W" ' message type  ' added by AD
    
    logBuffer.PageType = WakeType
    msgline1 = pageMessage
    Call WriteWakeQue(xnModual, extensionTmp, filename, logBuffer, pageMessage)
    Call XPutMessage(1, LogPATH, extensionTmp, msgline1, msgLine2, Temp)
exitsendpage:
Exit Sub

sendpageerror:

LogMessage ErrorMessage, "Unexpected error in SENDPAGE " & Error$
Resume exitsendpage
Resume

End Sub

Sub GetPageType(resultType, PageType)

    Dim i As Integer
    On Error GoTo OOPS
        resultType = -1
        For i = 0 To 39
        If PageType = PagingType(i).PageType Then
            resultType = PagingType(i).MsgType
            Exit For
        End If
    Next

ExitHere:
Exit Sub
OOPS:

resultType = -1
Resume ExitHere
End Sub



Sub GetNopaging(iniFileName As String)
    
    Dim fileNum%
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Dim Buffer1 As String
    Dim Buffer2 As String
    Dim Buffer3 As String
    Dim filename As String
    Dim part1 As String
    Dim part2 As String
    Dim Part3 As String
    Dim Temp As String
    Dim TotalRecords As Integer
    Dim TotalTypes As Integer
    Dim dime As String

    On Local Error GoTo XPutError
    
    ListPage = GetIniString("XN", "LIST_PAGE", "", iniFileName, True)
    NoPagingFilePath = GetIniString("XN", "NOPAGING_PATH", ".", iniFileName)
    FrqTypes = GetIniString("XN", "FRQ_TYPES", "", iniFileName)
    QuePath = GetIniString("XN", "QUE_PATH", ".", iniFileName)
    pageQueueBtrvTable = GetIniString("XN", "PAGE_QUEUE_BTRV_TABLE", "", iniFileName)
    ActionReminderFlagOffset = GetIniVal("XN", "ActionReminderFlagOffset", 0, iniFileName)
    ActionReminder = GetIniString("XN", "ActionReminder", "", iniFileName)
    PageFlagOffset = GetIniVal("XN", "pageflagoffset", 0, iniFileName)
  
    'Call XnOpenQueueBtrv
  
    fileNum% = FreeFile
Xretry:
    Close #fileNum%
    filename = NoPagingFilePath + "\NOPAGING"
    If Dir$(filename) = "" Then GoTo FileNotFound
    Open filename For Input Access Read As #fileNum%
    Line Input #fileNum%, Buffer1
    TotalRecords = Val(Buffer1)
    Line Input #fileNum%, Buffer1
    Line Input #fileNum%, Buffer1
    Line Input #fileNum%, Buffer1
    TotalTypes = 0
    For i = 1 To TotalRecords
        Buffer1 = "": Buffer2 = "": Buffer3 = ""
        Line Input #fileNum%, Buffer1
        Line Input #fileNum%, Buffer2
        Line Input #fileNum%, Buffer3
        Temp = Left$(Buffer1, 3)
        j = InStr(1, FrqTypes, Temp)
        If j <> 0 Then
'            Debug.Print "Parse: "; buffer1, buffer3
            PagingType(TotalTypes).QueDirect = ""
            PagingType(TotalTypes).Prefix = ""
            PagingType(TotalTypes).Sufix = ""
            PagingType(TotalTypes).Pin = ""
            PagingType(TotalTypes).Number = ""
            PagingType(TotalTypes).CallBackNumber = ""
            PagingType(TotalTypes).Description = ""
            PagingType(TotalTypes).MsgType = 0
            For k = 0 To 2
                PagingType(TotalTypes).Source(k) = 0
            Next
            PagingType(TotalTypes).Alias = Temp
            PagingType(TotalTypes).PageType = Temp
            If Left$(Buffer3, 1) = "%" And Mid$(Buffer3, 2, 1) >= "0" And Mid$(Buffer3, 2, 1) <= "9" And Mid$(Buffer3, 3, 1) >= "0" And Mid$(Buffer3, 3, 1) <= "9" Then
                PagingType(TotalTypes).Alias = Left$(Buffer3, 3)
            End If
            PagingType(TotalTypes).QueDirect = Temp
            j = InStr(1, Buffer3, "%Q")
            If j <> 0 Then
                PagingType(TotalTypes).QueDirect = "%" + Mid$(Buffer3, j + 2, 2)
            End If
            Temp = LTrim(RTrim(Buffer2))    'Ana Del Campo  05/21/96
            PagingType(TotalTypes).Description = Temp   'Ana Del Campo  05/21/96
            Temp = Buffer3
            Call TrimScript(Temp)
            part1 = "": part2 = "": Part3 = ""
            j = InStr(Temp, "+")
            If j <> 0 Then
                part1 = Left$(Temp, j - 1)
                Temp = Right$(Temp, Len(Temp) - j)
                j = InStr(Temp, "+")
                If j <> 0 Then
                    part2 = Left$(Temp, j - 1)
                    Part3 = Right$(Temp, Len(Temp) - j)
                Else
                    part2 = Temp
                End If
            Else
                part1 = Temp
            End If
            PagingType(TotalTypes).FieldNum = 1
'            Debug.Print part1, part2, part3
            j = InStr(part1, "%")
            If j <> 0 Then
                If j > 1 Then
                    PagingType(TotalTypes).Prefix = Left$(part1, j - 1)
                End If
                If j + 2 < Len(part1) Then
                    PagingType(TotalTypes).Sufix = Right$(part1, Len(part1) - j - 2)
                End If
                If InStr(part1, "%N") <> 0 Then
                    PagingType(TotalTypes).Source(0) = 2
                Else
                    PagingType(TotalTypes).Source(0) = 3
                End If
            Else
                PagingType(TotalTypes).Source(0) = 1
                PagingType(TotalTypes).Number = part1
            End If
            If Part3 = "" Then
                If part2 <> "" Then PagingType(TotalTypes).FieldNum = 2
                If Mid$(part2, 1, 1) <> "%" Then
                    PagingType(TotalTypes).CallBackNumber = part2
                    PagingType(TotalTypes).Source(2) = 1
                Else
                    If InStr(part2, "%N") <> 0 Then
                        PagingType(TotalTypes).Source(2) = 2
                    Else
                        PagingType(TotalTypes).Source(2) = 3
                        If InStr(part2, "%I1") <> 0 Then
                            PagingType(TotalTypes).MsgType = 0
                        Else
                            PagingType(TotalTypes).MsgType = 1
                        End If
                        If PagingType(TotalTypes).PageType = ListPage Then PagingType(TotalTypes).MsgType = 2  ' List Page
                    End If
                End If
            Else
                PagingType(TotalTypes).FieldNum = 3
                If Mid$(part2, 1, 1) <> "%" Then
                    PagingType(TotalTypes).Pin = part2
                    PagingType(TotalTypes).Source(1) = 1
                Else
                    If InStr(part2, "%N") <> 0 Then
                        PagingType(TotalTypes).Source(1) = 2
                    Else
                        PagingType(TotalTypes).Source(1) = 3
                        If InStr(part2, "%I1") <> 0 Then
                            PagingType(TotalTypes).MsgType = 0
                        Else
                            PagingType(TotalTypes).MsgType = 1
                        End If
                        If PagingType(TotalTypes).PageType = ListPage Then PagingType(TotalTypes).MsgType = 2  ' List Page
                    End If
                End If
                If Mid$(Part3, 1, 1) <> "%" Then
                    PagingType(TotalTypes).CallBackNumber = Part3
                    PagingType(TotalTypes).Source(2) = 1
                Else
                    If InStr(Part3, "%N") <> 0 Then
                        PagingType(TotalTypes).Source(2) = 2
                    Else
                        PagingType(TotalTypes).Source(2) = 3
                        If InStr(Part3, "%I1") <> 0 Then
                            PagingType(TotalTypes).MsgType = 0
                        Else
                            PagingType(TotalTypes).MsgType = 1
                        End If
                        If PagingType(TotalTypes).PageType = ListPage Then PagingType(TotalTypes).MsgType = 2  ' List Page
                    End If
                End If
            End If
            LogMessage SysMessage, "Page Type: " & PagingType(TotalTypes).PageType & " Direct: " & PagingType(TotalTypes).QueDirect & " Alias: " & PagingType(TotalTypes).Alias & " Prefix: " & PagingType(TotalTypes).Prefix & " Sufix: " & PagingType(TotalTypes).Sufix & " Pin: " & PagingType(TotalTypes).Pin & " Number: " & PagingType(TotalTypes).Number & " Callback: " & PagingType(TotalTypes).CallBackNumber & " Mesg Type: " & PagingType(TotalTypes).MsgType
            TotalTypes = TotalTypes + 1
        End If
    Next
    Close #fileNum%
    
PxExit:
    Exit Sub
    
FileNotFound:
 LogMessage ErrorMessage, "File Not Found -> " & filename
 'Beep
 Close #fileNum%
GoTo PxExit
    
    

XPutError:
    If Err = 70 Then Resume Xretry
'                       possible error may retry foreever
    
    LogMessage ErrorMessage, "Unexpected error in GetNopaging " & Error$
    'Beep
    Close #fileNum%
    Resume PxExit
Resume
End Sub

Sub TrimScript(ScriptStr As String)
    
    On Error GoTo OOPS
    
    Dim i As Integer
    Dim j As Integer
    Dim IString As String
    Dim NewString As String
    Dim FinalString As String
    IString = ""
    FinalString = ""
    NewString = ScriptStr

    i = InStr(NewString, "i=")
    If i <> 0 Then
        j = InStr(i + 2, NewString, "i")
        If j > i Then
            IString = Mid$(NewString, i + 2, j - i - 2)
        End If
        NewString = Left$(NewString, i - 1)
        i = InStr(NewString, "i")
        If i <> 0 Then
            FinalString = Left$(NewString, i - 1) + IString + Right$(NewString, Len(NewString) - i)
            ScriptStr = FinalString
        '    Debug.Print scriptStr
        End If
    End If
    i = InStr(ScriptStr, "%Z")
    If i <> 0 Then
        ScriptStr = Right$(ScriptStr, Len(ScriptStr) - i - 1)
     '   Debug.Print scriptStr
    End If
    i = InStr(ScriptStr, "%Z")
    If i <> 0 Then
        ScriptStr = Left$(ScriptStr, i - 1)
      '  Debug.Print scriptStr
    End If
    
ExitHere:
    Exit Sub
    
OOPS:
    LogMessage ErrorMessage, "Unexpected error in TrimScript " & Error$
    Resume ExitHere
    
End Sub


Sub GetPagingStatus(iniFileName As String)

'Reads the NOPGSTAT file to Determine if a Pager Number Can Be paged or Not
' Ana del Campo   05/21/96

On Error GoTo OOPS

    Dim Temp        As String
    Dim NewTemp     As String
    Dim i           As Integer
    Dim j           As Integer
    Dim Wait        As Long
    Dim Count       As Integer
    Dim StartTime   As Long
    Dim HowMuch     As Integer
    Dim fileNum%
    Dim Tell        As String
    Dim Flag        As Boolean

    PagStatFile = GetIniString("XN", "PageStatFile", "NOPGSTAT", iniFileName)
    NoPagingStatus = GetIniVal("XN", "NoPageStatus", 30, iniFileName)
    
   If Dir$(PagStatFile) = "" Then Error 53
    fileNum% = FreeFile

    ReDim Preserve statuslist(NoPagingStatus)
TryAgain:
    
    Close fileNum%
    Open PagStatFile For Binary Access Read As #fileNum%
    Temp = Input(3, fileNum%)
    
    j = 1
    For i = 0 To NoPagingStatus - 1
        With statuslist(i)
            .Description = Input(30, fileNum%)
            Temp = Input(2, fileNum%)
            .Page = Input(3, fileNum%)
            If Mid(.Page, 1, 1) = "Y" Then
               .PagingAllowed = True
            Else
              .PagingAllowed = False
            End If
            .Identifier = i
        End With
    Next i
    Close #fileNum%

ExitHere:
    Exit Sub
    
OOPS:

    Select Case Err
        Case 53
            LogMessage ErrorMessage, "The Paging Status File : " & PagStatFile & " Wasn't Found"
            Resume ExitHere
        Case 55
            Wait = timer
            Do Until (timer - Wait) > 5
              DoEvents
            Loop
            HowMuch = HowMuch + 1
            If HowMuch = 3 Then
                LogMessage ErrorMessage, "Couldn't Open " & PagStatFile
                Resume ExitHere
            End If
            Resume TryAgain
        Case Else
             LogMessage ErrorMessage, Trim(Str(Err)) & " Unexpected error in GetPagingStatus " & Error$
             Resume ExitHere
    End Select

End Sub

Sub WriteWakeQue(Initiator As String, Extension As String, filename As String, logBuffer As LogInfoType, Message As String)
    'If pageQueueBtrvTable = "" Then
        Call WriteWakeQueFile(Initiator, Extension, filename, logBuffer, Message)
    'Else
    '    Call WriteWakeQueBtr(Initiator, Extension, filename, logBuffer, Message)
    'End If
End Sub

Sub WriteWakeQueFile(Initiator As String, Extension As String, filename As String, logBuffer As LogInfoType, Message As String)
    'Dim filename As String
    Dim fileNum As Integer
    Dim waittime As Variant
    Dim scrollflag As Integer
    Dim PutPageQuePointer As Long
    Dim i As Integer

5:  QueRecordLen = Len(QueRecord)
    PointErrecLen = Len(PointErrec)

    On Error GoTo WritePageQueError

    fileNum = FreeFile
    
    'filename = ""
    
    'filename = GetIniString("XN", "WAKEQUEPATHANDLOCATION", "", Parameter.pPath) 'QuePath & "\" & PageQuePrefix & Right$(logBuffer.PageType, 2) & PageQueSuffix
  
    LogMessage SysMessage, "Writing page queue file: " & filename

    waittime = DateAdd("s", WaitForLock, Now)

10: Open filename For Binary Shared As fileNum

    If LOF(fileNum) < 20 Then
        PointErrec.GetPointer = 0
        PointErrec.PError1 = 0
        PointErrec.PError2 = 0
        PointErrec.PType = "1 "
'        pointerrec.putpointer = 1
        PointErrec.PutPointer = 0
        PointErrec.Junk = String$(Len(PointErrec.Junk), " ")
        Put #fileNum, 1, PointErrec
    End If

15: QueRecord.PType = Right$(logBuffer.PageType, 2)
    QueRecord.PStatus = "P "
    QueRecord.PDatein = Date$
    QueRecord.PTimein = Time$
    QueRecord.PExtension = Trim$(Extension)
    QueRecord.PIdin = Initiator
    QueRecord.PInfo = Message

    waittime = DateAdd("s", WaitForLock, Now)

25: Lock #fileNum, FromLock To ToLock
26: Get #fileNum, 1, PointErrec

' JS   putpagequepointer = pointerreclen + 1 + ((pointerrec.putpointer - 1) * querecordlen)
    LogMessage SysMessage, "Pointer Rec Len: " & PointErrecLen & " Put Pointer: " & PointErrec.PutPointer & " Que Record Len: " & QueRecordLen
27: PutPageQuePointer = PointErrecLen + 1 + ((PointErrec.PutPointer) * QueRecordLen)

30: Put #fileNum, PutPageQuePointer, QueRecord
    PointErrec.PutPointer = PointErrec.PutPointer + 1
    Put #fileNum, 1, PointErrec
ex2:
    Unlock #fileNum, FromLock To ToLock
ex1:
    Close fileNum
ex:
Exit Sub

WritePageQueError:

If Err = 70 And Erl = 10 Then
    If waittime > Now Then
        Resume 10
    Else
        LogMessage ErrorMessage, "Can't access page queue file: " & filename & " in " & WaitForLock & " seconds."
        Resume ex
    End If
End If

If Err = 70 And Erl = 25 Then
    If waittime > Now Then
        Resume 25
    Else
        LogMessage ErrorMessage, "Can't lock page queue file: " & filename & " in " & WaitForLock & "seconds."
        Resume ex1
    End If
End If

LogMessage ErrorMessage, "WriteWakeQueFile " & Erl & ":" & Error$
If Erl <= 25 Then
    Resume ex1
ElseIf Erl > 25 Then
    Resume ex2
Else
    Resume ex
End If

End Sub





