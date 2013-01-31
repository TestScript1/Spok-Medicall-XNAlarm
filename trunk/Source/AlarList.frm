VERSION 5.00
Begin VB.Form frmAlarms 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Alarms Info"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox alarmList 
      Appearance      =   0  'Flat
      Height          =   3540
      ItemData        =   "AlarList.frx":0000
      Left            =   45
      List            =   "AlarList.frx":0002
      TabIndex        =   0
      Top             =   60
      Width           =   6435
   End
End
Attribute VB_Name = "frmAlarms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub AddToList(ByVal parInfo As String)
Dim i As Integer

On Error GoTo OOPS
With alarmList

    If .ListCount > 199 Then .RemoveItem .ListCount - 1
    .AddItem parInfo, 0
    .ListIndex = 0
End With

ExitHere:
Exit Sub

OOPS:
LogMessage ErrorMessage, "Error: " & CStr(Err) & " in AddtoList, " & Error$
Resume ExitHere

End Sub

Private Sub Form_Load()
With Me
    .Top = GetIniString("SCREEN", "frmAlarms_Top", 0, Parameter.pPath)
    .Left = GetIniString("SCREEN", "frmAlarms_Left", 0, Parameter.pPath)
    .Width = GetIniString("SCREEN", "frmAlarms_Width", 3000, Parameter.pPath)
    .Height = GetIniString("SCREEN", "frmAlarms_Height", 3000, Parameter.pPath)
    .Caption = .Caption & " " & appTitle
End With
End Sub

Private Sub Form_Resize()
alarmList.Width = Me.Width - 200
alarmList.Height = Me.Height - 200

End Sub

Private Sub Form_Unload(Cancel As Integer)

Me.Hide
Cancel = 1

End Sub

