VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKeithleyDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Public Property Get DLModelName() As String
DLModelName = frmDriverLINX.lblModelName.Caption

End Property


Public Property Get DLStatus() As String

    DLStatus = frmDriverLINX.GetDriverStatus()
End Property



Public Function ReadChannel(ByVal parChannel As Long) As Long

    ReadChannel = frmDriverLINX.DIRead(parChannel)
End Function

Public Sub SetupChannel(ByVal parChannel As Long, ByVal commMode As String)

    frmDriverLINX.ConfigureChannel parChannel, LCase(commMode)
End Sub


Public Sub WriteChannel(ByVal parChannel As Long, ByVal parValue As Long)

    frmDriverLINX.DOWrite parValue, parChannel
End Sub


Private Sub Class_Initialize()
    InitLog ErrorMessage, "", App.Path & "\" & App.EXEName & ".err", "Error"
    If frmDriverLINX.InitializeDevice("KPCIPIO", 0) <> "" Then
        
        frmDriverLINX.DIConfig 0
    Else
        LogMessage ErrorMessage, "Can't initialize driver"
    End If
End Sub

Private Sub Class_Terminate()
Unload frmDriverLINX
End Sub


