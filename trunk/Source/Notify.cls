VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScreenNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private fileName As String
Private MAXRECORDNUMBER As Integer
Private MESSAGELEN As Integer

Public Sub AddToFile(parMsg As String)
'

Dim pointer As Long
Dim mess As String
Dim messageRec As String ' * MESSAGELEN           ' JS
Dim MsgStart As Integer, SegLen As Integer
Dim fileNum As Integer
Dim iTry As Integer
Dim iTotalRecLen As Integer
Const NOTEDATETIME = "MM/DD/YYYY HH:NN:SS"

'must move all the below code

On Error GoTo multierrorwrite
  '    record has:         Date/time           space     Message   vbCrLf
    iTotalRecLen = Len(NOTEDATETIME) + 1 + MESSAGELEN + 2 ''( for vbCrLf)
    
    messageRec = String(iTotalRecLen, Chr(0))
       
    
    fileNum = FreeFile
    Open fileName For Random Access Read Write Lock Read Write As fileNum Len = iTotalRecLen + 2
    
            
    Get fileNum, 1, messageRec
    If Len(messageRec) <> iTotalRecLen Then messageRec = String(iTotalRecLen, Chr(0))
    
    If EOF(fileNum) Then
        pointer = 1
    Else
        pointer = Val(messageRec)
    End If

    mess = Format$(Now, NOTEDATETIME) & " "
    SegLen = Len(messageRec) - 2 - Len(mess)
    MsgStart = 1
    Do While MsgStart <= Len(parMsg)
        pointer = pointer + 1
        
        If pointer > MAXRECORDNUMBER Then pointer = 2
        
        mess = mess & Mid$(parMsg, MsgStart, SegLen)
        MsgStart = MsgStart + SegLen
        
        messageRec = Left(mess & Space$(Len(messageRec)), Len(messageRec) - 2) & vbCrLf
        Put fileNum, pointer, messageRec
        LogMessage SysMessage, ">> Ticker message added:" & messageRec
        mess = Space$(Len(NOTEDATETIME) + 1)
    Loop


    mess = Str$(pointer) + Space$(Len(messageRec))
    mess = Left(mess, Len(messageRec) - 2) & vbCrLf
        
    messageRec = mess
    Put fileNum, 1, messageRec
    
        
WriteExit:
Close fileNum
Exit Sub

multierrorwrite:
iTry = iTry + 1
If iTry < 3 Then
   Resume
    
End If

LogMessage ErrorMessage, "Unexpected Error " & CStr(Erl) & " in clsScreenNotify.AddtoFile " & Error$
Resume WriteExit
Resume
End Sub

Private Sub Class_Initialize()
fileName = GetIniString("BANNER", "FILENAME", "", Parameter.pPath, True)
MAXRECORDNUMBER = GetIniString("BANNER", "MAX_MESSAGES", 30, Parameter.pPath, True)

MESSAGELEN = GetIniString("BANNER", "MAX_MESSAGE_LENGTH", 80, Parameter.pPath, True)

End Sub


