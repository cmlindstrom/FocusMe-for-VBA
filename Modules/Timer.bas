Attribute VB_Name = "Timer"
Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerfunc As LongPtr) As Long
Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private Const rootClass As String = "Timer"

' ADD: a tracking queue for the timers so they can be cleaned up.


'Need a timer ID to eventually turn off the timer. If the timer ID <> 0 then the timer is running
Public timerId As Long

Public Function IsRunning() As Boolean
    IsRunning = timerId <> 0
End Function

''' Deactivates the specified timer
''' - if timerId not specified will deactivate the last timer that
'''   was activated
Public Sub DeactivateTimer(Optional ByVal tId As Long = -1)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":DeactivateTimer"
    
    On Error GoTo ThrowException
    
    If tId = -1 Then tId = timerId
    
    Dim lSuccess As Long
    lSuccess = KillTimer(o, tId)
    If lSuccess = 0 Then
        strTrace = "Timer (" & tId & ") failed to be deactivated."
        GoTo ThrowException
    End If
    
    ' Log the timer deactivation
    strTrace = "Timer (" & tId & ") has been successfully deactivated."
    LogMessage strTrace, strRoutine
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Sets up a new timer that triggers every lSeconds.
''' - Returns the id for the timer
Public Function ActivateTimer(ByVal lSeconds As Long) As Long

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ActivateTimer"
    
    On Error GoTo ThrowException
    
    Dim mSec As Long
    mSec = lSeconds * 1000 ' milliseconds
    
    Dim tId As Long
    tId = SetTimer(0, 0, mSec, AddressOf TriggerTimer)
    If tId = 0 Then
        strTrace = "Timer activation failed; interval request = " & lSeconds & " seconds."
        GoTo ThrowException
    End If
    
    ' Capture the latest timer request
    timerId = tId
    
    ' Log the timer activation
    strTrace = "A Timer (" & tId & ") has been activated; interval request = " & lSeconds & " seconds."
    LogMessage strTrace, strRoutine
    
    ' Pass back the timer identifier
    ActivateTimer = tId
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    ActivateTimer = -1

End Function



Public Sub TriggerTimer(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idevent As Long, ByVal Systime As Long)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TriggerTimer"

    Dim myQueue As MemoryLogger
    Set myQueue = ThisOutlookSession.MessageQueue
    If Not myQueue Is Nothing Then
        strTrace = "Timer (" & idevent & ") has been triggered."
        Call myQueue.Post(idevent, strTrace)
    Else
        strTrace = "The Addin Message Queue returned as null."
        LogMessageEx strTrace, Nothing, strRoutine
    End If
  
'    If idevent = TimerId Then
'        MsgBox "The TriggerTimer function has been automatically called!"
'    Else
'        MsgBox "Another TriggerTimer function automatically called."
'    End If
  
End Sub


' - Legacy

Private Sub ActivateTimerOld(ByVal nMinutes As Long)

    Dim mSec As Long
    mSec = nMinutes * 1000 * 60 'The SetTimer call accepts milliseconds, so convert from minutes
    If timerId <> 0 Then Call DeactivateTimer 'Check to see if timer is running before call to SetTimer
    timerId = SetTimer(0, 0, mSec, AddressOf TriggerTimer)
    If timerId = 0 Then
        MsgBox "The timer failed to activate."
    End If
    
End Sub
Private Sub DeactivateTimerOld()
    Dim lSuccess As Long
    lSuccess = KillTimer(0, timerId)
    If lSuccess = 0 Then
        MsgBox "The timer failed to deactivate."
    Else
        timerId = 0
    End If
End Sub
