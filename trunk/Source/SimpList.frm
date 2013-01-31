VERSION 5.00
Begin VB.Form frmSimplexList 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Buffer Information"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox bufferList 
      Appearance      =   0  'Flat
      Height          =   3540
      ItemData        =   "SimpList.frx":0000
      Left            =   60
      List            =   "SimpList.frx":0002
      TabIndex        =   0
      Top             =   15
      Width           =   6435
   End
End
Attribute VB_Name = "frmSimplexList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub AddToList(ByVal parListInfo As String)
Dim i As Integer

On Error GoTo OOPS
With bufferList

    If .ListCount > 199 Then .RemoveItem .ListCount - 1
    .AddItem parListInfo, 0
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
    .Top = GetIniString("SCREEN", "frmSimplexList_Top", 0, Parameter.pPath)
    .Left = GetIniString("SCREEN", "frmSimplexList_Left", 0, Parameter.pPath)
    .Width = GetIniString("SCREEN", "frmSimplexList_Width", 3000, Parameter.pPath)
    .Height = GetIniString("SCREEN", "frmSimplexList_Height", 3000, Parameter.pPath)
    .Caption = .Caption & " " & appTitle
End With
End Sub

Private Sub Form_Resize()
bufferList.Width = Me.Width - 200
bufferList.Height = Me.Height - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
Cancel = 1
End Sub


