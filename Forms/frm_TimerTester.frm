VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_TimerTester 
   Caption         =   "UserForm1"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frm_TimerTester.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_TimerTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' - Fields

Private Const rootClass As String = "frm_TimerTester"

Private Const Period As Double = 0.5 ' seconds

Dim WithEvents myQueue As MemoryLogger
Attribute myQueue.VB_VarHelpID = -1

Dim alerts As Integer
Dim myTimerId As Long

Dim lastClicked As Date

' - Event Handlers

Private Sub btn_Start_Click()

    Dim strTrace As String

'    If Timer.IsRunning Then Exit Sub
'    Timer.ActivateTimer 1
'    Status "Started the timer, running every minute."

    If myTimerId > 0 Then
        strTrace = "A timer (" & myTimerId & ") is already running..."
        Status strTrace
        Exit Sub
    End If

    myTimerId = Timer.ActivateTimer(Period)
    If myTimerId <= 0 Then
        strTrace = "Failed to start the timer..."
        Status strTrace
    End If
    
    Status "Started the timer (" & myTimerId & "), running every " & Period & " seconds."

'    Dim rId As String
'    rId = Common.GenerateUniqueID(3)
'    strTrace = "New message " & rId
'    Call myQueue.Post(10, strTrace)

    lastClicked = Now
    
End Sub

Private Sub btn_Stop_Click()

    If myTimerId <= 0 Then
        strTrace = "No timer is currently running..."
        Status strTrace
        Exit Sub
    End If
    
    Timer.DeactivateTimer myTimerId
    myTimerId = 0
    Status "Stopped the timer."
    alerts = 0

End Sub

''' Monitor the Addin Message Queue and respond to Timer messages
Private Sub myQueue_NewMessage(ByVal id As Long, ByVal msg As String)
    alerts = alerts + 1
    Status "Message posted for channel (" & id & "), alert: " & alerts & " msg: " & msg
    
    ' Key Tracker
    If IsKeyPressed(WinContextMenu.VK_LBUTTON) Then
        Status "Left mouse button has been pressed - " & alerts
    End If
    
End Sub

' - Constructor

Private Sub UserForm_Initialize()

    ' Grab the global queue
    Set myQueue = ThisOutlookSession.MessageQueue

    Me.Caption = "Timer Tester"
    lbl_Description.Caption = "Run this form to test if Outlook VBA can use a timer."
    Call Status
    
    myTimerId = 0
    
End Sub

Private Sub UserForm_Terminate()
    If myTimerId > 0 Then Timer.DeactivateTimer myTimerId
End Sub

' - Methods

' - Supporting Methods

Private Sub Status(Optional ByVal msg As String = "")
    sb_Status.SimpleText = msg
    
    AppendLog msg
End Sub

Private Sub AppendLog(ByVal msg As String)

    If Len(txtbx_Log.text) = 0 Then
        txtbx_Log.text = msg
    Else
        txtbx_Log.text = txtbx_Log.text & vbCrLf & msg
    End If

End Sub

