VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DPChannels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"DPChannel"
Attribute VB_Ext_KEY = "Member0" ,"DPChannel"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public ChannelNum As Integer
Public Function Add(DPPageNumeric As String, _
    DPPageMessage As String, DPxnMessage As String, DPExtension As String, _
    DPActionReminder As Integer, DPClearAlarm As Integer, DPOutputState As Integer, _
    DPInputState As Integer, DPType As Integer, dpSendBanner As Boolean) As DPChannel
    
    'create a new object
    Dim objNewMember As DPChannel
    On Error GoTo OOPS
    Set objNewMember = New DPChannel

    
    'set the properties passed into the method
    objNewMember.DPPageNumeric = DPPageNumeric
    objNewMember.DPPageMessage = DPPageMessage
    objNewMember.DPxnMessage = DPxnMessage
    objNewMember.DPExtension = DPExtension
    objNewMember.DPActionReminder = DPActionReminder
    objNewMember.DPClearAlarm = DPClearAlarm
    objNewMember.DPOutputState = DPOutputState
    objNewMember.DPInputState = DPInputState
    objNewMember.DPType = DPType
    objNewMember.dpSendBanner = dpSendBanner
    
    mCol.Add objNewMember
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

ExitHere:
Exit Function

OOPS:
LogMessage ErrorMessage, "Error in DPChannels_Add, " & Error$
Resume ExitHere
End Function

Public Property Get Item(vntIndexKey As Variant) As DPChannel
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property






Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
 
    Set mCol = Nothing
End Sub

