VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fmeTaskItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' - - Fields

Private Const rootClass As String = "fmeTaskItem"

Dim stgs As Settings

Private WithEvents m_task As Outlook.TaskItem
Attribute m_task.VB_VarHelpID = -1

' - - Properties

Dim f_Id As String
Dim f_lastTouched As Date

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

Private Sub m_task_Open(Cancel As Boolean)
    
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_task_Open " & f_Id
    
    f_lastTouched = Now
    
    strTrace = "Task Item opened."
    LogMessage strTrace, strRoutine
    
End Sub

Private Sub m_task_Close(Cancel As Boolean)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_task_Close " & f_Id
        
    f_lastTouched = Now

    strTrace = "Task Item closed."
    LogMessage strTrace, strRoutine

End Sub

Private Sub m_task_PropertyChange(ByVal Name As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_task_PropertyChange " & f_Id
        
    f_lastTouched = Now
    
    strTrace = "Task Item property (" & Name & ") changed."
    LogMessage strTrace, strRoutine

End Sub

Private Sub m_task_CustomAction(ByVal Action As Object, ByVal Response As Object, Cancel As Boolean)
    
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_task_CustomAction"
    
    f_lastTouched = Now
        
    strTrace = "Task Item had a custom action request."
    LogMessage strTrace, strRoutine

End Sub

Private Sub m_task_Forward(ByVal Forward As Object, Cancel As Boolean)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_task_Forward " & f_Id
    
    On Error GoTo ThrowException
    
    Dim ut As Utilities
    Set ut = New Utilities
    
    Dim m As Outlook.MailItem
    Set m = ut.GetMessageForTask(m_task)
    
    ' If TaskAttachment required do it here
    Dim bl As Boolean
    bl = stgs.AttachTaskToEmail
    If bl Then
        ' Implement later
    End If
    
    m.Display

    ' Ignoring Outlook's Forward method
    Cancel = True
   
    GoTo Finally
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

Finally:
    Set ut = Nothing

End Sub

Private Sub m_task_Unload()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":m_task_Unload " & f_Id
        
    f_lastTouched = Now
    
    strTrace = "Task Item unloaded from memory."
    LogMessage strTrace, strRoutine

End Sub

' - - Constructor

Private Sub Class_Initialize()
    Set stgs = New Settings
End Sub

Private Sub Class_Terminate()
    Set stgs = Nothing
    Set m_task = Nothing
End Sub

' Methods

Public Sub Load(ByVal oTask As Outlook.TaskItem)
    Set m_task = oTask
    f_Id = Common.GenerateUniqueID(3)
End Sub


