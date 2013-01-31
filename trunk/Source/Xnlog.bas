Attribute VB_Name = "Xnlog"
Option Explicit

'''''  ADDED BY TK """""""""""""""""""""""
Public Const SAVED_NEW = 1

Public gLogFileNum             As Integer
Public gOperator               As String

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public statuslist As New clsNoPgStats
'Public PagingType As New clsNoPagings
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public gstrMessageType As String

#If ApplicationType <> "DLL" Then
    Public gErrors As Object
#End If

Public ActReminder As Object
''''''''''''''
Type InstructionType
    Greeting As String * 70
    Update As String * 7
    Opid As String * 2
    Dum0 As String * 1
End Type

Type NameType
    Names(2) As String * 25
    Zdum0 As String * 5
End Type

Type LogInfoType  ' new XKM--------------------------------------------------
    ProfileID As String
    Duplicated As String
    duplicatedCounter As Integer
    Name As String              ' Offset 1 XKM!Name of XKM!FirstName + XKM!LastName
    firstName As String
    outdialno As String  ' XKM!dialNumber
    PagerId As String
    WhereAbouts As String
    WhereaboutsDT As String
    WhereaboutsOPID As String
    Department As String
    DepartmentOPID As String
    DepartmentDT As String
    PatientID As String
    deptNum As String * 11          ' Offset 869   ' XKM!DepartmentNumber
    Location As String
    groupCode As String
    instruction(9) As InstructionType  ' XKM!Instruction1-XKM!Instruction10
    Format As Integer
    extensionType As String
    PassWord As String 'XKM!Password
    address1 As String
    address2 As String
    address3 As String
    address4 As String
    address5 As String
    groupExt As String           ' Offset 479  'XKM!MainGroupExtension
    registryFlag As String
    printBills As String
    extLight As String           ' Offset 727  ' XKM!Lights
    lamp As String               ' Offset 0  ' XKM!Lamps
    history As String
    ynFlags(24) As String   'XKMFlag1-XKM!Flag25
    Color As Integer
    roomNo As String * 8            ' Offset 66  ' XKM!RoomNumber
    Company As String
    VaXfer As String
    CardNumber As String
    PasswordDT As String
    PasswordOPID As String
    RecordDT As String
    RecordOPID As String
    VaXferNumber As String
    GroupMessage As String
    WebXchange As String
    ' From Msg Counter
    MessageTOTAL As String
    messageViewed As String
    messagePrinted As String
    ' from Pagers tbl
    pageStatus As String     ' Pagers!Status
    statTime As String        'Pagers!StatusDateTime
    statOpId As String        '  Pagers!StatusOperator
    PageType As String        ' Pagers!DefaultPager
    pageTypeOPID As String   'Pagers!DefaultOperator
    pageTypeDT As String         'Pagers!DefaultDateTime
    coverExtension As String * 10 'Pagers!CoveringExtension
    coverName As String * 40      ' using coveringextension find in XKM coverName
    coverId As String * 10          ' Offset 1050 Lenght 10 and covering ID  '????????
    Free1 As String
    Free2 As String
    Free3 As String
    Free4 As String
End Type
'------------------------------------------------------------------------------------
'Type LogInfoType
'    cflag As String * 1             ' Offset 36
'    primaryExtension As String * 4  ' Offset 62
'    Greeting As String * 75         ' Offset 82
'    greetingOpInitls As String * 2  ' Offset 157
'    printerNo As String * 12        ' Offset 440
'    patchNo As String * 22          ' Offset 452
'    SharedName(2) As NameType       ' Offset 480
'    nextN(8) As String * 5          ' Offset 732
'    sharedExtFlag As String * 1     ' Offset 777
'    voiceMainNo As String * 22      ' Offset 778
'    primaryExt As String * 5        ' Offset 800
'    statExtra As String * 30        ' Offset 805
'    roomNoExtra As String * 22      ' Offset 835
'    recNum As String * 12           ' Offset 857
'
'
'    received As String * 4          ' Offset 960
'    Delivered As String * 4         ' Offset 964
'    printed As String * 4           ' Offset 968
'    vSmile As String * 14           ' Offset 972
'
'
'
'
'    nameExtra As String * 15        ' Offset 1060
'    ownp As String * 3              ' Offset 1075
'    recNumExtra As String * 3       ' Offset 1078
'
'    zenextn(8) As String * 5        ' Offset 1106
'    zphil As String * 49            ' Offset 1151
'
'    emp1 As String * 45             ' Offset 1235
'
'    emp2 As String * 45             ' Offset 1315
'
'    emp3 As String * 45             ' Offset 1395
'
'    emp4 As String * 45             ' Offset 1475
'    endFiller As String * 480       ' Offset 1520
'End Type                            ' Total size 2000
'
Type LogBufferType
    Data As String * 2000
End Type

Type ShortRec
        Ff   As String * 1
        Nm   As String * 25
        Dum  As String * 23
        Mcnt As String * 4
        Mvce As String * 4
        Mprt As String * 4
        Dum1 As String * 19
End Type

Type ShortOneRecord
    Zln As String * 25
    Zpw As String * 10
    Zcf As String * 1
    Zcn As String * 12
    Zmcnt As String * 4
    Zmvce As String * 4
    Zmprt As String * 4
    Zprsc As String * 1
    Zpext As String * 4
    Zroom As String * 8
    Zuno As String * 6
End Type

Type ShortTwoRecord
    Ztme As String * 11
    Zmsg As String * 69
End Type

Type ShortThreeRecord
    Zmext As String * 4
    Zcomma1 As String * 1
    Zstamp As String * 11
    Zmopid As String * 2
    Zcomma2 As String * 1
    Zmcatagory As String * 15
    Zcomma3 As String * 1
    Zmamt As String * 10
    Zcomma4 As String * 1
    Zmdescription As String * 30
    Zdum0 As String * 4
End Type

Type Message
        Tme As String * 11
        Msg As String * 69
End Type

Type CountRec
    M1 As String * 4
    M2 As String * 4
    M3 As String * 4
    M4 As String * 68
End Type

Public Type Reminder
    fileName     As String * 10
    MsgLine      As String * 4
    Date         As String * 10
    Time         As String * 5
    TimeOut      As Integer
End Type


' -------   FROM XNXKMDB.BAS ---------------------
Type XkmrecInfo
    Extension     As String * 10
    Name          As String * 15
    DefPageMacro  As String * 10
    Department    As String * 10
    PagerId       As String * 10
    Type          As String * 1
    DialNum       As Variant
    Format As Integer
    CoveringExtn  As String * 10
    CoveringName  As String * 11
    WhereAbouts   As String * 65
    Status As String * 1
    PatientID As String
    registryFlag As String
    MessageTOTAL As Long
    Duplicated As String
    duplicatedCounter As Integer
    PassWord As String
End Type

Global Xkm      As XkmrecInfo  ' This info is now in Pagers tbl


Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'  TK added 06/11/1998
Public CurrentSysMessage As String
Public CurrentErrMessage As String



Sub BreakApartByDelimiter(OneBigLine As String, outLines() As String, ByVal parDelimiter As String, Optional parExtraStr)

Dim intCount As Integer
Dim i As Integer, p As Integer
Dim pos As Integer
Dim strPartLine As String

On Error GoTo OOPS

intCount = UBound(outLines)
If intCount = 0 Then Exit Sub
p = 1

For i = 0 To UBound(outLines) - 1
    pos = InStr(OneBigLine, parDelimiter)
    If pos = 0 Then outLines(i) = OneBigLine: Exit For
    outLines(i) = Mid(OneBigLine, p, pos - 1)
    OneBigLine = Mid(OneBigLine, pos + Len(parDelimiter))  'vbCrLf))
    
Next
If IsMissing(parExtraStr) = False Then

    parExtraStr = OneBigLine
End If
ExitHere:
Exit Sub
OOPS:

Resume ExitHere
End Sub

Sub BreakApart(OneBigLine As String, strLInes() As String, ByVal intMaxPiece As Integer, ByVal intTotalLines As Integer)
Dim intAllLength As Integer
Dim intManySmallLines As Integer
Dim i As Integer
intAllLength = Len(OneBigLine)

If intAllLength > intTotalLines * intMaxPiece Then intAllLength = intTotalLines * intMaxPiece

intManySmallLines = intAllLength \ intMaxPiece

ReDim strLInes(intManySmallLines)

For i = 0 To UBound(strLInes)
    strLInes(i) = Mid(OneBigLine, i * intMaxPiece + 1, intMaxPiece)

Next
End Sub



Function CollectMsgLines(strLInes() As String) As String
On Error GoTo OOPS

Dim i As Integer
Dim strOneBigLine As String

20 For i = 0 To UBound(strLInes)
30    strOneBigLine = strOneBigLine & strLInes(i)
40 Next

50 CollectMsgLines = strOneBigLine
ExitHere:
Exit Function

OOPS:
CollectMsgLines = ""
Resume ExitHere
End Function



Sub RefreshDBConnection()
Dim localMCountRs As ADODB.Recordset
Dim iTry As Integer

On Error GoTo OOPS
' Just try any sql statement.....
TryAgain:
Set localMCountRs = New ADODB.Recordset
localMCountRs.Open _
    "Select LastMessageNumber From MessageCounters", DB, adOpenForwardOnly
localMCountRs.Close
Set localMCountRs = Nothing
LogMessage SysMessage, "Refreshed DB connection OK."

ExitHere:
Exit Sub

OOPS:
' Error !!!  DB is not working
SendEmailNote Error$

If iTry <= 1 Then
    DB.Close
    Set DB = Nothing

    Delay 100
' Re-open DB
    Set DB = New ADODB.Connection
    XnOpenDataBase gDBName, username, XnPassword
    iTry = iTry + 1
    LogMessage ErrorMessage, "Error:" & Error & " in RefreshDBconnection(). Re-try again."
    GoTo TryAgain:
Else
    LogMessage ErrorMessage, "RefreshDBConnection can't be done !!!   Re-starting computer is necessary."
End If
Resume ExitHere

End Sub

Private Function DeleteCrLfChar(ByVal parTempStr As String) As String

' This func will remove vbCrLf chars inside the string

Dim i As Integer
Dim normalStr As String

For i = 1 To Len(parTempStr)
    Select Case Asc(Mid(parTempStr, i, 1))
        Case 10, 13
            normalStr = normalStr & " "
        Case Else
            normalStr = normalStr & Mid(parTempStr, i, 1)
    End Select
        
Next
DeleteCrLfChar = normalStr

End Function

Function RegularDateTime(parDateTime) As String
    Dim strDate As String
    Dim strTime As String
    
    strDate = Mid$(parDateTime, 5, 2) & "/" & Mid$(parDateTime, 7, 2) & "/" & Mid$(parDateTime, 1, 4)
    strTime = Format$(Mid$(parDateTime, 9, 2) & ":" & Mid$(parDateTime, 11, 2), "hh:nna/p")
    RegularDateTime = strDate & " " & strTime
    
End Function
Function HeaderDateTime(parDateTime) As String
    '------------
    Dim strDate As String
    Dim strTime As String
    
    strDate = Mid$(parDateTime, 7, 4) & Mid$(parDateTime, 1, 2) & Mid$(parDateTime, 4, 2)
    strTime = Mid(parDateTime, 12)
    strTime = Format(strTime, "Short Time")
    strTime = Mid$(strTime, 1, 2) & Mid(strTime, 4, 2)
    HeaderDateTime = strDate & strTime
    
End Function

Function ConvertNullStr(parData As Variant, Optional parType As Variant)
Dim i As Integer
Dim strTemp As String
'    Convert String from Null to empty
If IsNull(parData) Then

    If Not IsMissing(parType) Then
        If parType = 1 Then
            ConvertNullStr = 0 ' for numeric data
        Else  ' boolean
            ConvertNullStr = False
        End If
    Else
        ConvertNullStr = ""
    End If
Else
   '  Strip|||| in the field OPERATOR ID
   If InStr(parData, Chr$(0)) Then
   
    strTemp = ""
    For i = 1 To Len(parData)
        If Mid$(parData, i, 1) <> Chr$(0) Then strTemp = strTemp & Mid$(parData, i, 1)
    Next
    parData = "": parData = Trim$(strTemp)
    
   End If
   ConvertNullStr = parData
End If
End Function
Function Read_Xkm(lgBuffer As LogInfoType, parInfo As String, parInfoDT As String, parInfoOper As String) As Integer
On Error GoTo OOPS

With Xkm
    .CoveringExtn = lgBuffer.coverExtension
    .CoveringName = lgBuffer.coverName
    .DefPageMacro = lgBuffer.PageType
    .Department = lgBuffer.Department
    .DialNum = lgBuffer.outdialno
    .Extension = lgBuffer.ProfileID
    .Name = lgBuffer.Name
    .PagerId = lgBuffer.PagerId
    .Type = lgBuffer.extensionType
    .WhereAbouts = parInfo  'lgBuffer.WhereAbouts
    .Status = lgBuffer.pageStatus
    .registryFlag = lgBuffer.registryFlag
    .MessageTOTAL = Val(lgBuffer.MessageTOTAL)
    .duplicatedCounter = lgBuffer.duplicatedCounter
    .Duplicated = lgBuffer.Duplicated
    .PassWord = lgBuffer.PassWord
End With
ExitHere:
'Errorform.Show
Exit Function

Resume ExitHere

OOPS:
LogMessage ErrorMessage, "Error in Read_XKM " & Error$
Resume ExitHere
Resume
End Function

Public Function RemoveInvisiblechars(ByVal parString As String) As String

Dim i As Integer
Dim myLine As String
Dim okChar As String

For i = 1 To Len(parString)

    okChar = Mid(parString, i, 1)
    If Asc(okChar) <= 127 Then
        myLine = myLine & okChar
    End If
Next

RemoveInvisiblechars = myLine
End Function

Function UpdateXkmTable(ThisPerson As XkmrecInfo) As Boolean
' Update Pagers Info -(This info was kept in XKM tbl for the Old system ----------------------------------------
 On Error GoTo OOPS

    Dim Temp As Integer
    UpdatePagersInfo ThisPerson.Extension, ThisPerson.PagerId, ThisPerson.Status, ThisPerson.DefPageMacro
    
ExitHere:
    Exit Function
    
OOPS:
    'ErrorMessages Err, Error$, "On UpdateXkmTable"
    LogMessage ErrorMessage, "Unexpected error in UpdateXkmTable " & Error$
    UpdateXkmTable = False
    Resume ExitHere

End Function



Public Sub XPutMessage(DoActionRemind As Integer, parMsgType As String, parProfileID As String, _
    parMessage As String, parMessage2 As String, parMessage3 As String, Optional parMessage4, _
    Optional parMessage5, _
    Optional parMessage6, _
    Optional parMessage7, _
    Optional parMessage8, _
    Optional parMsgFrom, _
    Optional parMsgTo, _
    Optional parARAlarmType, _
    Optional parActReminderEx, Optional markDelivered)
    
'***
'CKO 8/31/2000: Added an optional argument parActReminderEx.
' The function SetReminderEx() is called when set parActReminderEx to 1, the
' function SetReminder() is called if parActReminderEx is missing.
' The difference between SetReminderEx() and SetReminder() is that the latter creates and destroys the
' CalendarVB5 object every time when SetReminder() is called. While the SetReminderEx() only calls
' CalendarVB5's function, CatchNewRecord().
'***
Dim tempId As String * 10
Dim LastMsgNumber As Long
Dim res            As Boolean
Dim strFrom     As String
Dim strTo         As String
Dim strMsg     As String
Dim iCounter As Integer
Dim strInformation As String
Dim bMarkDelivered As Boolean  ' we will specify delivery flag

Const RECORD_TYPE = 6
Const ALARM_TYPE = "1 "
On Error GoTo OOPS

RSet tempId = parProfileID

'Check of Connection is still available  DB

RefreshDBConnection
strMsg = parMessage
If parMessage2 <> "" Then strMsg = strMsg & vbCrLf & parMessage2
If parMessage3 <> "" Then strMsg = strMsg & vbCrLf & parMessage3
If IsMissing(parMessage4) = False Then strMsg = strMsg & vbCrLf & parMessage4
If IsMissing(parMessage5) = False Then strMsg = strMsg & vbCrLf & parMessage5
If IsMissing(parMessage6) = False Then strMsg = strMsg & vbCrLf & parMessage6
If IsMissing(parMessage7) = False Then strMsg = strMsg & vbCrLf & parMessage7
If IsMissing(parMessage8) = False Then strMsg = strMsg & vbCrLf & parMessage8

If IsMissing(markDelivered) = False Then bMarkDelivered = markDelivered

LogMessage SysMessage, "Writing message profile: " & tempId
If IsMissing(parMsgFrom) = False Then strFrom = parMsgFrom
If IsMissing(parMsgTo) = False Then strTo = parMsgTo

TryAgain:
LastMsgNumber = SaveNewMsg(tempId, parMsgType, gOperator, strMsg, bMarkDelivered, strFrom, strTo)
If UCase(Parameter.alarmtype) = "SIMPLEX" Then
    If frmSimplex.Visible Then
        frmSimplex.lblWarning.Visible = False
        If frmSimplex.MSComm1(0).Break = True Then
            frmSimplex.MSComm1(0).Break = False
        End If
    End If
End If
If LastMsgNumber = -1 Then  ' error
    If frmSimplex.Visible Then frmSimplex.lblWarning.Visible = True
    frmSimplex.lblWarning.Caption = "Error when saving Message to the Messages table"
    frmSimplex.Refresh
    LogMessage ErrorMessage, ">>>>>> Error when saving Message to the Messages table >>>>>>>>>>"
    Exit Sub
End If
res = UpdateTotalForNewMsg(tempId, parMsgType, CStr(LastMsgNumber))
If DoActionRemind = 0 Then
    '   Setup Action reminder for the message
'        SetReminder tempId, LastMsgNumber, Format(Now, "yyyymmddhhnn"), parARAlarmType     // changed JS 6/22/00
  'If IsMissing(parActReminderEx) Then
    ''''''''''SetReminder tempId, Str(LastMsgNumber), Format(Now, "yyyymmddhhnn")
    strInformation = ALARM_TYPE & Space(10 - Len(Trim(parProfileID))) & Trim(parProfileID) & CStr(LastMsgNumber)
    
    CatchNewRecord RECORD_TYPE, strInformation, Format(Now, "yyyymmddhhnn"), parProfileID
  'ElseIf parActReminderEx = 1 Then
    'SetReminderEx tempId, Str(LastMsgNumber), Format(Now, "yyyymmddhhnn")
    'CatchNewRecord RECORD_TYPE, strInformation, DateTime, parProfileID
  'Else 'for future options
  '  LogMessage ErrorMessage, "[E] invalid function argument '" & parActReminderEx & "' for parActionReminderEx"
  'End If
End If
ExitHere:
Exit Sub

OOPS:
LogMessage ErrorMessage, "Unexpected Error " & CStr(Erl) & " in XPutMessage " & Error$ & " for ProfileID: " & tempId
Resume ExitHere
Resume
End Sub

Public Function DeliverDay(Tstr) As String
    Dim sWeek As String
    Dim sMonth As String, iMonth As Integer
    Dim sDay As String, iDay As Integer
    Dim sYear As String, iYear As Integer
    Dim sHour As String
    Dim sMinute As String
    Dim sAmPm As String
    Dim MsgDay As Long
    Dim today As Date, YearNow As Integer

    'If Not DbBUG Then On Error GoTo OOPS

    sMonth = Left$(Tstr, 2)
    sDay = Mid$(Tstr, 3, 2)
    sHour = Mid$(Tstr, 5, 2)
    sMinute = Mid$(Tstr, 7, 2)
    sWeek = Mid$(Tstr, 9, 1)
    sYear = Right$(Tstr, 1)
    iMonth = Val(sMonth)
    iDay = Val(sDay)
    If sHour <> "" Then
       If CInt(sHour) >= 12 Then
           sAmPm = "PM"
       Else
           sAmPm = "AM"
       End If
    End If
    'calulate the year number
    today = Date: YearNow = Year(today)
    If Right$(YearNow, 1) >= sYear Then
        iYear = (Val(Left$(YearNow, 3))) * 10 + Val(sYear)
    Else
        iYear = (Val(Left$(YearNow, 3)) - 1) * 10 + Val(sYear)
    End If
    
    MsgDay = DateSerial(iYear, iMonth, iDay)
    
    DeliverDay = Format$(MsgDay, "long date") & _
        "  AT " & sHour & ":" & sMinute & " " & sAmPm

ExitHere:
Exit Function

OOPS:
'Select Case ErrorTrap("DeliverDay", "Global")
  '  Case 1: Resume
  '  Case 2: Resume Next
  '  Case Else: Resume ExitHere
'End Select
        
End Function

' This code was changed by TK  06/10/1998
' Updates XKM table , Pagers Table

Function UpdateLogInfo(Extension As String, LogInformation As LogInfoType) As Boolean

On Error GoTo OOPS
    
    Dim ThisFile As String
    Dim fileNumber As Integer
    Dim Wait As Long
    Dim HowMuch As Integer
    
    LogMessage SysMessage, "Update Log Info " & Extension
    UpdateLogInfo = True
    
    UpdateXkmInfo Extension, LogInformation
    
    UpdatePagersInfo Extension, _
        LogInformation.PagerId, LogInformation.pageStatus, LogInformation.PageType
    
ExitHere:
    Exit Function
    
OOPS:

        LogMessage ErrorMessage, "Error UpdateLogInfo " & Error$
        UpdateLogInfo = False
        Resume ExitHere
  
End Function

'
'   UpdateLogBuffer:   Retieves the first 2000 bytes of a log file and stores the information
'                      in a string
'
'   Parameters:     exttension  - log extension to write to
'                   logPath     - path to log file
'                   lgBuffer    - string store info
'
Function UpdateLogBuffer(Extension As String, LogPATH As String, LogInformation As LogBufferType) As Boolean

On Error GoTo OOPS

    Dim ThisFile As String
    Dim fileNumber As Integer
    Dim Wait As Long
    Dim HowMuch As Integer
    
    UpdateLogBuffer = True
    
    ThisFile = GetLogPath(Extension, "LOG", LogPATH)
    fileNumber = FreeFile
    
Retry:
    
    Open ThisFile For Random Access Read Write Lock Read As fileNumber Len = Len(LogInformation)
    Put #fileNumber, 1, LogInformation
    Close #fileNumber
    
ExitHere:
    Exit Function
    
OOPS:
    Select Case Err
        Case 53
            LogMessage ErrorMessage, "UpdateLogBuffer, Log File for Extension : " & Extension & " not found"
            UpdateLogBuffer = False
            Resume ExitHere
        Case 55, 70
            Wait = timer
            Do Until (timer - Wait) > 5
                DoEvents
            Loop
            HowMuch = HowMuch + 1
            If HowMuch = 3 Then
                LogMessage ErrorMessage, "Couldn't open Log File for Extension : " & Extension
                UpdateLogBuffer = False
                Resume ExitHere
            End If
            Resume Retry
        Case Else
            LogMessage ErrorMessage, "Error UpdateLogBuffer " & Error$
            UpdateLogBuffer = False
            Resume ExitHere
    End Select
End Function



Public Function GetRidOfJunkChar(ByVal RawData As String) As String
    Dim TotalChar As Integer
    Dim tmpstr As String
    Dim Pt As Integer
   
    Dim iCount As Integer
    
    tmpstr = DeleteCrLfChar(RawData)
    TotalChar = Len(RawData)
    For Pt = TotalChar To 1 Step -1
        Select Case Asc(Mid$(RawData, Pt, 1))
        Case 0, 9, 10, 13, 32
            tmpstr = Left$(RawData, Pt - 1)
        Case Else
            Exit For
        End Select
    Next
    ' Remove leading junk characters
        
    For Pt = 1 To Len(tmpstr)
        If Asc(Mid(tmpstr, Pt, 1)) = 32 Or Asc(Mid(tmpstr, Pt, 1)) = 160 Then
            iCount = iCount + 1
        Else
            Exit For
        End If
    Next
    
    GetRidOfJunkChar = Mid(tmpstr, iCount + 1)
    
End Function




'   GetLogPath:     Builds the directory path to find a LOG, FRQ, or SUP file
'
'   Return:         String containing directory and file name
'
'   Parameters:     extn    - extension to find i.e. C100
'                   type    - type of file "LOG", "FRQ", "SUP"
'
'   Comments:       paramter.logPath global variable containing XN directory path required
'
'   Original:       Joseph Slawinski        12/26/95
'   Changed by  Tatyana Kharakh        06/10/1998
Public Function GetLogPath(strProfile As String, ExtType As String, LogPATH As String) As String
    
Dim strSubDir1 As String
Dim strSubDir2 As String
Dim strSubDir3 As String

Dim strSubDir As String
Dim strDiskLetter As String

On Error GoTo OOPS

    strDiskLetter = Trim(LogPATH)
    If Right(strDiskLetter, 1) <> "\" Then strDiskLetter = strDiskLetter & "\"
    
 
    If IsNumeric(Mid$(Right$(strProfile, 2), 1, 1)) Then
        strSubDir1 = "N" & Mid$(Right$(strProfile, 2), 1, 1)
    Else
        strSubDir1 = "N" & CStr((Asc(Mid$(Right$(strProfile, 2), 1, 1)) - 65) Mod 10)
    End If
    If Dir(strDiskLetter & strSubDir1, vbDirectory) = "" Then MkDir strDiskLetter & strSubDir1
    
    If IsNumeric(Right$(strProfile, 1)) Then
        strSubDir2 = "N" & Right$(strProfile, 1)
    Else
        strSubDir2 = "N" & CStr((Asc(Right$(strProfile, 1)) - 65) Mod 10)
    End If
    
    If Dir(strDiskLetter & strSubDir1 & "\" & strSubDir2, vbDirectory) = "" Then MkDir strDiskLetter & strSubDir1 & "\" & strSubDir2
    strSubDir = strDiskLetter & strSubDir1 & "\" & strSubDir2 & "\" & Trim(strProfile)
    If Dir(strSubDir, vbDirectory) = "" Then MkDir strSubDir
    
    GetLogPath = strSubDir
    
ExitHere:
    Exit Function
    
OOPS:
    GetLogPath = ""
    Resume ExitHere
End Function

'
'   Get Info From XKM , Pagers tables
Sub GetLogInfo(fileName$, Ext$, lgBuffer As LogInfoType, errorValue As Integer, Optional MakeSearch)
    
    Dim FlName$
    Dim fileNum%
    Dim FileIsFree As Boolean
    Dim Wait As Long
    Dim StartTime As Long
    Dim HowLong As Integer
    Dim IsSearch As Boolean
    Dim strInfo As String  ' formerly WhereAbouts
    Dim strInfoDT As String '''''' WA date time
    Dim strInfoOper As String ''' WA Operator
    ' This code was created by TK
    ' 06/11/1998
    
    ' BY DEAFAULT I MAKE SEARCH ( IF YOU DO NOT PASS THE OPTIONAL PARAM. I MAKE SEARCH )
    
    errorValue = 0
    If IsMissing(MakeSearch) Then
        ' ---- MAKE SEARCH --------
        IsSearch = True
    Else
        If IsNumeric(MakeSearch) Then
            IsSearch = MakeSearch
        Else
            IsSearch = False
        End If
    End If
        
    Call TestDBConnection ' allow to restore db connection if it was dropped TK 11/21/2011
    
    GetXKMInfo Ext$, lgBuffer, IsSearch
    If errorValue = 0 Then
        GetPagersInfo lgBuffer.PagerId, lgBuffer.pageStatus, lgBuffer.statTime, _
            lgBuffer.statOpId, lgBuffer.PageType, lgBuffer.pageTypeOPID, lgBuffer.pageTypeDT, _
            lgBuffer.coverExtension, lgBuffer.coverName, strInfo, strInfoDT, strInfoOper
    
        Read_Xkm lgBuffer, strInfo, strInfoDT, strInfoOper
    End If

PxExit2:
    Exit Sub
    
XGetError:
    errorValue = Error
    LogMessage ErrorMessage, "GetLogInfo error " & Error$
    Resume PxExit2

End Sub


Private Sub TestDBConnection()
Dim tmp As New ADODB.Recordset

On Error GoTo OOPS

Start1:

tmp.ActiveConnection = DB

tmp.Open "SELECT * FROM XKM Where ProfileId ='     XTEND'"

tmp.Close

ExitHere:

Exit Sub


OOPS:
XnCloseDataBase
LogMessage ErrorMessage, "Err " & Err.Number & " " & Err.Description & " in TestDBConnection()"
If XnOpenDataBase(gDBName, username, XnPassword) Then
    Resume Start1
Else
    LogMessage ErrorMessage, "CANNOT OPEN DATABASE !"
    Resume ExitHere
End If


End Sub

Sub InitXnLog(iniFileName As String)
    gOperator = GetIniString("XN", "OperName", "Unknown", iniFileName)
    
    OpenAllTables (iniFileName)
End Sub


Public Function GetFilePath(Extn As String, ExtType As String, LogPATH As String, GetlogPathMode As Integer) As String
    
    On Error GoTo OOPS     ' 09/10/97 this function was added by Igor Kovac
                           
    If GetlogPathMode Then
        GetFilePath = GetLogPath(Extn, ExtType, LogPATH)
    Else
        GetFilePath = LogPATH + "\" + Trim(Extn) + "." + ExtType
    End If
    
ExitHere:
   Exit Function
   
OOPS:
   MsgBox "Error : " & Trim(Str(Err)) & " at GetFilePath " & Error$
   
   Resume ExitHere

End Function


'
'   GetLogBuffer:   Retieves the first 2000 bytes of a log file and stores the information
'                   in a string
'
'   Parameters:     filename$  - file name and path of the log file i.e. \N3\N0\XB00.LOG
'                   ext$       - log extension to write to
'                   lgBuffer   - string store info
'                   errorValue - Optional value of any error encountered
'
Sub GetLogBuffer(fileName$, Ext$, lgBuffer As LogBufferType, errorValue As Integer)
    
    Dim FlName$
    Dim fileNum%
    Dim FileIsFree As Boolean
    Dim Wait As Long
    Dim StartTime As Long
    Dim HowLong As Integer
    
    ' This Code Wait for the LOG file to be Free and to read it
    ' Ana Del Campo         05/20/96
    
    errorValue = 0
    FileIsFree = True

    On Local Error GoTo XGetError
    fileNum% = FreeFile
XRetry3:
    If Not FileIsFree Then
        HowLong = HowLong + 1
        Do
            Wait = timer
            Wait = Wait - StartTime
            If Wait > 10 Then Exit Do
        Loop
    End If
    Close #fileNum%
    If Dir$(fileName$) = "" Then Error 30057
    FlName$ = fileName$
    Open FlName$ For Random Access Read Lock Read Write As #fileNum% Len = Len(lgBuffer)
    Get #fileNum%, , lgBuffer
    Close #fileNum%
PxExit2:
    Exit Sub
    
XGetError:
    errorValue = Err
    If Err = 70 Then
        FileIsFree = False
        StartTime = timer
        If HowLong = 3 Then
           LogMessage ErrorMessage, "GetLogBuffer error " & Error$
           'Beep
           Resume PxExit2
        End If
        Resume XRetry3
    End If
    If Err = 30057 Then
        LogMessage ErrorMessage, "File Not Found ->" & fileName$
        'Beep
        Close #fileNum%
        Resume PxExit2
    End If
    LogMessage ErrorMessage, "GetLogBuffer error " & Error$
    Close #fileNum%
    Resume PxExit2

End Sub




