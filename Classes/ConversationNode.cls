VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConversationNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Fields

Private Const rootClass As String = "ConversationNode"


Private strName As String
Private strText As String
Private iLevel As Integer
Private cvNode As ConversationNode
Private oContext As Object
Private oTag As Object
Private strTopic As String
Private strConvId As String
Private strIndex As String
Private iLength As Integer
Private iClass As Integer

' Events


' Properties

' Node Name
Public Property Let Name(ByVal nme As String)
    strName = Name
End Property
Public Property Get Name() As String
    Name = strName
End Property

' Node Text
Public Property Let Text(ByVal txt As String)
    strText = txt
End Property
Public Property Get Text() As String
    Text = strText
End Property

' Level
Public Property Let Level(ByVal ilvl As Integer)
    iLevel = ilvl
End Property
Public Property Get Level() As Integer
    Level = iLevel
End Property

' Parent ConversationNode
Public Property Set Parent(ByVal nd As ConversationNode)
    Set cvNode = nd
End Property
Public Property Get Parent() As ConversationNode
    Set Parent = cvNode
End Property

' Node Context, e.g. Outlook Object
Public Property Set Context(ByVal o As Object)
    Set oContext = o
End Property
Public Property Get Context() As Object
    Set Context = oContext
End Property

' Tag
Public Property Set Tag(ByVal o As Object)
    Set oTag = o
End Property
Public Property Get Tag() As Object
    Set Tag = oTag
End Property

' Conversation Topic
Public Property Let Topic(ByVal tp As String)
    strTopic = tp
End Property
Public Property Get Topic() As String
    Topic = strTopic
End Property

' Conversation ID
Public Property Let ConvId(ByVal id As String)
    strConvId = id
End Property
Public Property Get ConvId() As String
    ConvId = strConvId
End Property

' Index
Public Property Let index(ByVal idxString As String)
    strIndex = idxString
End Property
Public Property Get index() As String
    index = strIndex
End Property

'  Length
Public Property Let Length(ByVal ilen As Integer)
    iLength = ilen
End Property
Public Property Get Length() As Integer
    Length = iLength
End Property

' Context Type = Outlook OlClass
Public Property Let ContextType(ByVal icls As Integer)
    iClass = icls
End Property
Public Property Get ContextType() As Integer
    ContextType = iClass
End Property

' Constructor

Private Sub Class_Initialize()

End Sub

' Methods

''' Fill object instance from an Outlook Item
Public Sub fill(ByVal o As Object)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":FillFromOutlookObject"
    
    Me.ContextType = -1
    
    If TypeOf o Is Outlook.MailItem Then Me.ContextType = Outlook.OlObjectClass.olMail
    If TypeOf o Is Outlook.ContactItem Then Me.ContextType = Outlook.OlObjectClass.olContact
    If TypeOf o Is Outlook.TaskItem Then Me.ContextType = Outlook.OlObjectClass.olTask
    If TypeOf o Is Outlook.AppointmentItem Then Me.ContextType = Outlook.OlObjectClass.olAppointment
    If TypeOf o Is Outlook.MeetingItem Then Me.ContextType = Outlook.OlObjectClass.olMeetingRequest
    
    If Me.ContextType < 0 Then
        strTrace = "WARNING: Unrecognized object type."
        GoTo ThrowException
    End If
    
    FillFromOutlookObject o
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Public Sub FillVariant(ByVal v As Variant)
    Me.Name = v
End Sub

' Supporting Methods

Private Sub FillFromOutlookObject(ByVal obj As Object)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":FillFromOutlookObject"

    On Error GoTo ThrowException
    
    Dim o As Outlook.MailItem

    strText = obj.Subject
    strName = obj.SentOn & " | " & obj.SenderName
    Me.Topic = obj.ConversationTopic
    Me.ConvId = obj.ConversationID
    Me.index = obj.ConversationIndex
    Me.Length = Len(obj.ConversationIndex)
    
    Set Me.Context = obj
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

