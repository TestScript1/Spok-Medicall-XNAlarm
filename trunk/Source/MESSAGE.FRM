VERSION 5.00
Begin VB.Form Message 
   Caption         =   "Message"
   ClientHeight    =   690
   ClientLeft      =   1485
   ClientTop       =   4170
   ClientWidth     =   6420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   690
   ScaleWidth      =   6420
   Visible         =   0   'False
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Ok 
      Caption         =   "O&k"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox AlphaMessage 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.TextBox NumericMessage 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label FinalMessage 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Problem 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "The Pager Type is not Valid, Check Configuration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   4575
   End
End
Attribute VB_Name = "Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()
   
    Unload Me

End Sub


Private Sub Form_Activate()

    On Error GoTo OOPS

    If NumericMessage.Visible = True Or Problem.Visible Then
        Me.Height = 1050
        Me.Left = 2640
        Me.Top = 3720
        Me.Width = 6540
    End If
    If AlphaMessage.Visible = True Then
        Me.Height = 3510
        Me.Left = 2880
        Me.Top = 2820
        Me.Width = 6540
    End If
    ResizeForm Me
    CenterThisForm Me

ExitHere:
    Exit Sub
    
OOPS:
    ErrorMessages Err, Error$, "On Activate Message"
    Resume ExitHere

End Sub

Private Sub Ok_Click()

    On Error GoTo OOPS

    Dim i          As Integer
    Dim ThisChar   As String
    Dim Numeric    As Boolean
    Dim Temp       As Integer

    If NumericMessage.Visible = True Then
        NumericMessage = LTrim(RTrim(NumericMessage))
        If Len(NumericMessage) < 1 Then
            MsgBox "I Don't Have A Message"
            NumericMessage.SetFocus
            Exit Sub
        End If
        Numeric = True
        For i = 1 To Len(NumericMessage)
            ThisChar = Mid(NumericMessage, i, 1)
            If Not ThisChar Like "[1,2,3,4,5,6,7,8,9,0]" Then
               Numeric = False
               Exit For
            End If
        Next i
        If Not Numeric Then
            MsgBox "The Message Has to Be Numeric"
            NumericMessage.SetFocus
            Exit Sub
        End If
        FinalMessage = NumericMessage
    End If
    
    If AlphaMessage.Visible = True Then
        AlphaMessage = LTrim(RTrim(AlphaMessage))
        If Len(AlphaMessage) < 1 Then
            MsgBox "I Don't have a message"
            AlphaMessage.SetFocus
            Exit Sub
        End If
        FinalMessage = AlphaMessage
    End If
    
    Me.Hide
    
ExitHere:
    Exit Sub
    
OOPS:
    ErrorMessages Err, Error$, "On Ok At Message"
    Resume ExitHere

End Sub


