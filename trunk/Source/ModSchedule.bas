Attribute VB_Name = "ModSchedule"
Option Explicit
Public Schedule As New BtrieveRecordSet
Public NameOfScheduleTable As String




Public Sub CloseScheduleBTR()
On Error Resume Next
    Schedule.CloseRecordset
    

End Sub

Public Sub OpenScheduleBtr(strIniPath As String)
    Dim NameOfScheduleTable As String
    On Error GoTo OOPS
    NameOfScheduleTable = GetIniString("XN", "ScheduleTable", "Schedule", strIniPath)
    Set Schedule = DB.OpenRecordSet(NameOfScheduleTable)
exiterror:
Exit Sub
OOPS:
    LogMessage ErrorMessage, "OpenScheduleBtr :" & Err.Description
    Resume exiterror
End Sub

Function CatchNewRecord(intRecordType As Integer, strInfo As String, strHeader As String, strProfileID As String) As Boolean
        Dim tempId As String * 10
        Dim intCounter As Integer
        Dim i As Integer
        Dim j As Integer
        On Error GoTo OOPS
        LogMessage SysMessage, "Local CatchNewrecord Enter :" & strInfo
        RSet tempId = Trim(strProfileID)
StartAgain:
        With Schedule
             .index = 5
             .Search ">=", intRecordType
             If .NoMatch = True Then
                intCounter = 1
             Else
                i = 0
                If Not .EOF Then
                    Do Until (.Fields("RecordType") <> intRecordType) Or .EOF
                        If (i + 1) <> .Fields("RecordCount") Then
                           intCounter = i + 1
                           i = i + 1
                           Exit Do
                        Else
                           i = .Fields("RecordCount")
                           intCounter = i + 1
                        End If
                        .MoveNext
                        If .EOF Then Exit Do
                    Loop
                End If
                
             End If
             If intCounter = 0 Then intCounter = 1
    
        .AddNew
        .Fields("Header") = strHeader
        .Fields("RecordType") = intRecordType
        .Fields("TimeType") = 0
        .Fields("generalinformation") = strInfo
        .Fields("Frequency") = 0
        .Fields("ScheduleTime") = ""
        .Fields("Extension") = tempId
        .Fields("RecordCount") = intCounter
        .Update
        LogMessage SysMessage, "Local CatchNewrecord Exit :" & strInfo
        End With
        CatchNewRecord = True
exiterror:
        Exit Function
OOPS:
        LogMessage ErrorMessage, "CatchNewRecord :" & Err.Description & " " & Err.Number & ", Status " & CStr(Schedule.Status)
        'Resume exiterror
        j = j + 1
        If j < 2 Then
         Select Case Schedule.Status
             Case 84
              DelayTime 1
              Resume StartAgain
            Case 3, 95, 3006
                   'CloseAllTables
                   CloseScheduleBTR
                   'XnCloseDataBase
                   DelayTime 5
                   'XnOpenDataBase (Parameter.XnDataBase)
                   'InitXnLog (Parameter.pPath)
                   'InitFrqTable (Parameter.pPath)
                   OpenScheduleBtr (Parameter.pPath)
                   Resume StartAgain
            
            Case Else
                LogMessage ErrorMessage, "CatchNewRecord :" & Err.Description & " " & Err.Number
                Resume exiterror
            End Select
        End If
        CatchNewRecord = False
        Resume exiterror

End Function

