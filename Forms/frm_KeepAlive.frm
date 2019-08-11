VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_KeepAlive 
   Caption         =   "Session Keeper"
   ClientHeight    =   2280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3240
   OleObjectBlob   =   "frm_KeepAlive.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_KeepAlive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' - Fields

Private Const rootClass As String = "frm_KeepAlive"

Private Const keyWatch As Double = 0.5 ' seconds

Private WithEvents tmr_userActive As fmeTimer
Attribute tmr_userActive.VB_VarHelpID = -1

' - Properties
Private started As Date ' the last time the Start button was pressed
Private inActiveTime As Long ' in seconds
Private lastActive As Date ' last time the user engaged in the UI

' - Event Handlers

Private Sub btn_Start_Click()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":btn_Start_Click"
    
    On Error GoTo ThrowException
    
    If tmr_userActive.IsRunning Then
        strTrace = "Keep alive timer already running."
        GoTo ThrowException
    End If
    
    strTrace = "Start the key watch timer."
    tmr_userActive.StartTimer keyWatch
    
    strTrace = "Populate UI"
    started = Now
    Me.lbl_StartedAt.Caption = Format(started, "dd mmm yyyy hh:nn")
    UpdateInactiveTime 0
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Private Sub btn_Start_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    LogMessage "Control lost focus.", rootClass & ":btn_Start_Exit"
End Sub

Private Sub btn_Stop_Click()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":btn_Stop_Click"
    
    On Error GoTo ThrowException

    If Not tmr_userActive.IsRunning Then
        strTrace = "Keep alive timer is not currently running."
        GoTo ThrowException
    End If

    tmr_userActive.StopTimer
    UpdateInactiveTime -1
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Private Sub btn_Stop_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    LogMessage "Control lost focus.", rootClass & ":btn_Stop_Exit"
End Sub

Private Sub btn_Test_Click()

    ' Send useless keystroke
    SendKey VK_Tab

End Sub

Private Sub tmr_userActive_TimerTick()

    ' Key Tracker
    If IsKeyPressed(WinContextMenu.VK_LBUTTON) Then
        ' Reset Inactive time
        inActiveTime = 0
        lastActive = Now
    Else
        ' Update Inactive time
        inActiveTime = DateDiff(GetDatePartFormat(Second), lastActive, Now)
        
        ' Call KeepAlive to check threshold and refresh sess
        KeepAlive inActiveTime
        
    End If
    UpdateInactiveTime inActiveTime
        
End Sub

Private Sub txtbx_KeepAliveTime_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    LogMessage "Control lost focus.", rootClass & ":txtbx_KeepAliveTime_Exit"
End Sub


' - Constructor

Private Sub UserForm_Initialize()
    Set tmr_userActive = New fmeTimer
    Status "Idle"
    UpdateInactiveTime -1
End Sub

Private Sub UserForm_Terminate()
    If tmr_userActive.IsRunning Then tmr_userActive.StopTimer
    Set tmr_userActive = Nothing
End Sub

' - Methods

' - Supporting Methods

Private Sub UpdateInactiveTime(ByVal s As Long)

    Dim el As Integer ' elapsed time in seconds

    If s < 0 Then
        Me.lbl_InactiveTimeValue.Caption = "Not tracking.."
        Status "Idle"
    Else
        Me.lbl_InactiveTimeValue.Caption = FormatElapsedTime(s)
        
        el = DateDiff(GetDatePartFormat(Second), started, Now)
        Status "Running: " & FormatElapsedTime(el)
    End If
End Sub

Private Sub Status(Optional ByVal msg As String = "")
    Me.sb_Status.SimpleText = msg
End Sub

Private Sub KeepAlive(ByVal inActive As Variant)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":KeepAlive"
    
    On Error GoTo ThrowException

    ' Check for Threshold time - move cursor to keep session alive
    Dim t As Long ' threshold in seconds
    t = CDbl(Me.txtbx_KeepAliveTime.text) * 60
    
    Dim x As Integer
    Dim y As Integer
    
    If inActive > t Then
    
        ' Move the mouse to keep the session alive
        Call GetCursorPosition(x, y)
        Call SetCursorPosition(x + 100, y)
        Sleep 150
        Call SetCursorPosition(x, y)
        
        ' Send useless keystroke
        ' SendKey VK_Tab
        
        ' Reset Inactive time
        inActiveTime = 0
        lastActive = Now
        
        ' Announce to user
        UpdateInactiveTime inActiveTime
            
    End If
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
End Sub
