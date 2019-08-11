VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fmeMailItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' - - Fields

Private Const rootClass As String = "fmeMail"

Dim stgs As Settings

Private WithEvents m_mail As Outlook.MailItem
Attribute m_mail.VB_VarHelpID = -1

Dim f_id As String
Dim f_lastTouched As Date

' - - Properties

''' Unique Identifier for this MailItem
Public Property Let Id(ByVal myId As String)
    f_id = myId
End Property
Public Property Get Id() As String
    Id = f_id
End Property

''' Most recent date and time the mailItem had an event occur
Public Property Get Touched() As Date
    Touched = f_lastTouched
End Property

' - - Event Handlers

Private Sub m_mail_Open(Cancel As Boolean)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_mail_Open " & f_id
    
    f_lastTouched = Now
    
    strTrace = "Mail Item opened."
    LogMessage strTrace, strRoutine
    
End Sub

Private Sub m_mail_Close(Cancel As Boolean)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_mail_Close " & f_id
        
    f_lastTouched = Now

    strTrace = "Mail Item closed."
    LogMessage strTrace, strRoutine
    
End Sub

Private Sub m_mail_PropertyChange(ByVal Name As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_mail_PropertyChange " & f_id
        
    f_lastTouched = Now
    
    strTrace = "Mail Item property (" & Name & ") changed."
    LogMessage strTrace, strRoutine
    
End Sub

Private Sub m_mail_Read()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_mail_Read " & f_id
        
    f_lastTouched = Now
    
    strTrace = "Mail Item Read."
    LogMessage strTrace, strRoutine

End Sub

Private Sub m_mail_Reply(ByVal Response As Object, Cancel As Boolean)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_mail_Reply " & f_id
    
    Dim bFile As Boolean
    bFile = stgs.MoveOnReply
    If bFile Then
        Dim ut As New Utilities
        If ut.MoveToArchive(m_mail) Then
            strTrace = "Successfully filed after Forward message created."
        Else
            strTrace = "Failed to file after Forward message created."
        End If
        LogMessage strTrace, strRoutine
    Else
        strTrace = "Mail Forward request - MoveOnReply flag not set."
        LogMessage strTrace, strRoutine
    End If
           
    strTrace = "Mail Item reply started."
    LogMessage strTrace, strRoutine
    
    Exit Sub
           
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Private Sub m_mail_Forward(ByVal Forward As Object, Cancel As Boolean)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_mail_Forward " & f_id
    
    Dim bFile As Boolean
    bFile = stgs.MoveOnReply
    If bFile Then
        Dim ut As New Utilities
        If ut.MoveToArchive(m_mail) Then
            strTrace = "Successfully filed after Forward message created."
        Else
            strTrace = "Failed to file after Forward message created."
        End If
        LogMessage strTrace, strRoutine
    Else
        strTrace = "Mail Forward request - MoveOnReply flag not set."
        LogMessage strTrace, strRoutine
    End If
           
    strTrace = "Mail Item forward started."
    LogMessage strTrace, strRoutine
    
    Exit Sub
           
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Private Sub m_mail_Send(Cancel As Boolean)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_mail_Send " & f_id
        
    On Error GoTo ThrowException
        
    f_lastTouched = Now
           
    strTrace = "Mail Item Send."
    LogMessage strTrace, strRoutine
    
    Exit Sub
           
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Private Sub m_mail_Unload()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_mail_Unload " & f_id
        
    f_lastTouched = Now
    
    strTrace = "Mail Item unloaded from memory."
    LogMessage strTrace, strRoutine

End Sub

' - - Constructor

Private Sub Class_Initialize()
    Set stgs = New Settings
End Sub

' Methods

Public Sub Load(ByVal oMail As Outlook.MailItem)
    Set m_mail = oMail
    f_id = Common.GenerateUniqueID(3)
End Sub

