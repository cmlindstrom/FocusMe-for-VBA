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

Dim f_Id As String
Dim f_lastTouched As Date

' - - Properties

''' Unique Identifier for this MailItem
Public Property Let id(ByVal myId As String)
    f_Id = myId
End Property
Public Property Get id() As String
    id = f_Id
End Property

''' Most recent date and time the mailItem had an event occur
Public Property Get Touched() As Date
    Touched = f_lastTouched
End Property

' - - Event Handlers

Private Sub m_mail_Open(Cancel As Boolean)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_mail_Open " & f_Id
    
    f_lastTouched = Now
    
    strTrace = "Mail Item opened."
    LogMessage strTrace, strRoutine
    
End Sub

Private Sub m_mail_Close(Cancel As Boolean)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_mail_Close " & f_Id
        
    f_lastTouched = Now

    strTrace = "Mail Item closed."
    LogMessage strTrace, strRoutine
    
End Sub

Private Sub m_mail_PropertyChange(ByVal Name As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_mail_PropertyChange " & f_Id
        
    f_lastTouched = Now
    
    strTrace = "Mail Item property (" & Name & ") changed."
    LogMessage strTrace, strRoutine
    
End Sub

Private Sub m_mail_Read()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_mail_Read " & f_Id
        
    f_lastTouched = Now
    
    strTrace = "Mail Item Read."
    LogMessage strTrace, strRoutine

End Sub

Private Sub m_mail_Reply(ByVal Response As Object, Cancel As Boolean)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_mail_Reply " & f_Id
    
    On Error GoTo ThrowException
    
    strTrace = "Mail Item Reply started."
    LogMessage strTrace, strRoutine
    
    Dim bFile As Boolean
    bFile = stgs.MoveOnReply
    If bFile Then
        Dim ut As Utilities
        Set ut = New Utilities
        If stgs.IgnoreSentMailMove Then
            If ut.IsItemParent(m_mail, "Sent Items") Then
                strTrace = "Ignoring the move request."
                LogMessage strTrace, strRoutine
            Else
                If ut.MoveToArchive(m_mail) Then
                    strTrace = "Successfully filed after Reply message created."
                Else
                    strTrace = "Failed to file after Reply message created."
                End If
                LogMessage strTrace, strRoutine
            End If
        Else
            If ut.MoveToArchive(m_mail) Then
                strTrace = "Successfully filed after Reply message created."
            Else
                strTrace = "Failed to file after Reply message created."
            End If
            LogMessage strTrace, strRoutine
        End If
    Else
        strTrace = "Mail Reply request - MoveOnReply flag not set."
        LogMessage strTrace, strRoutine
    End If
             
    Exit Sub
           
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Private Sub m_mail_ReplyAll(ByVal Response As Object, Cancel As Boolean)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_mail_ReplyAll " & f_Id
    
    On Error GoTo ThrowException
    
    strTrace = "Mail Item ReplyAll started."
    LogMessage strTrace, strRoutine
    
    Dim bFile As Boolean
    bFile = stgs.MoveOnReply
    If bFile Then
        Dim ut As New Utilities
        If stgs.IgnoreSentMailMove Then
            If ut.IsItemParent(m_mail, "Sent Items") Then
                strTrace = "Ignoring the move request."
                LogMessage strTrace, strRoutine
            Else
                If ut.MoveToArchive(m_mail) Then
                    strTrace = "Successfully filed after Reply message created."
                Else
                    strTrace = "Failed to file after Reply message created."
                End If
                LogMessage strTrace, strRoutine
            End If
        Else
            If ut.MoveToArchive(m_mail) Then
                strTrace = "Successfully filed after Reply message created."
            Else
                strTrace = "Failed to file after Reply message created."
            End If
            LogMessage strTrace, strRoutine
        End If
    Else
        strTrace = "Mail ReplyAll request - MoveOnReply flag not set."
        LogMessage strTrace, strRoutine
    End If
           
    Exit Sub
           
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Private Sub m_mail_Forward(ByVal Forward As Object, Cancel As Boolean)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_mail_Forward " & f_Id
    
    On Error GoTo ThrowException
    
    strTrace = "Mail Item forward started."
    LogMessage strTrace, strRoutine
    
    Dim bFile As Boolean
    bFile = stgs.MoveOnReply
    If bFile Then
        Dim ut As New Utilities
        If stgs.IgnoreSentMailMove Then
            If ut.IsItemParent(m_mail, "Sent Items") Then
                strTrace = "Ignoring the move request."
                LogMessage strTrace, strRoutine
            Else
                If ut.MoveToArchive(m_mail) Then
                    strTrace = "Successfully filed after Reply message created."
                Else
                    strTrace = "Failed to file after Reply message created."
                End If
                LogMessage strTrace, strRoutine
            End If
        Else
            If ut.MoveToArchive(m_mail) Then
                strTrace = "Successfully filed after Reply message created."
            Else
                strTrace = "Failed to file after Reply message created."
            End If
            LogMessage strTrace, strRoutine
        End If
    Else
        strTrace = "Mail Forward request - MoveOnReply flag not set."
        LogMessage strTrace, strRoutine
    End If
    
    Exit Sub
           
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Private Sub m_mail_Send(Cancel As Boolean)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_mail_Send " & f_Id
        
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
    strRoutine = rootClass & ":m_mail_Unload " & f_Id
        
    f_lastTouched = Now
    
    strTrace = "Mail Item unloaded from memory."
    LogMessage strTrace, strRoutine

End Sub

' - - Constructor

Private Sub Class_Initialize()
    Set stgs = New Settings
End Sub

Private Sub Class_Terminate()
    Set stgs = Nothing
End Sub

' Methods

Public Sub Load(ByVal oMail As Outlook.MailItem)
    Set m_mail = oMail
    f_Id = Common.GenerateUniqueID(3)
End Sub

