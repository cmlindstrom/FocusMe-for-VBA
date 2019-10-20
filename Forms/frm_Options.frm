VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Options 
   Caption         =   "Settings"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4935
   OleObjectBlob   =   "frm_Options.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' - - Fields

Private Const rootClass As String = "frm_Options"

Private bDirty As Boolean


' - - Properties

' - - Event Handlers


' - Buttons

Private Sub btn_About_Click()
    Dim f As New frm_About
    f.Show
End Sub

Private Sub btn_IndexFolders_Click()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":btn_IndexFolders_Click"
    
    On Error GoTo ThrowException
    
    Call Status("Working...")
    
    Dim st As New Setup
    st.IndexOutlookFolders
    
    GoTo Finally
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
Finally:
    Call Status

End Sub

Private Sub btn_ImportCategories_Click()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":btn_ImportCategories_Click"
    
    On Error GoTo ThrowException
    
    Call Status("Working...")
    
    Dim pCnt As Integer
    Dim st As New Setup
    pCnt = st.ImportProjectsFromCategories
          
SkipOut:
    MsgBox "Imported " & pCnt & " projects to the datastore."
    
    GoTo Finally
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

Finally:
    Call Status

End Sub

Private Sub btn_Cancel_Click()

    If bDirty Then
      Dim dResult As DialogResult
      dResult = MsgBox("Do you want to save your setting changes?", _
                        vbQuestion Or vbYesNoCancel, "Update Settings")
      If dResult = DialogResult_Yes Then
        Call SaveSettings
      End If
      If dResult = DialogResult_Cancel Then Exit Sub
      
    End If
    
    bDirty = False
    Unload Me

End Sub

Private Sub btn_OK_Click()

    Call SaveSettings
    bDirty = False
    Unload Me

End Sub

Private Sub btn_ResetPane_Click()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":btn_ResetPane_Click"

    Dim f As FME_Pane
    
    Dim UFHWnd As Long
    UFHWnd = WinForms.GetUserFormHandle("Task List")
    If UFHWnd <= 0 Then
        strTrace = "ERROR: Failed to find the Task List form."
        LogMessage strTrace, strRoutine
        Exit Sub
    End If
    
    Call SetWindowPosition(UFHWnd, 10, 10)
    
    ' - Logic didn't work
    ' Set f = ThisOutlookSession.FMEPane
    ' If Not f Is Nothing Then Call SetFormPosition(f, 10, 10)
    
End Sub

Private Sub btn_SelectFolder_Click()

    Dim ut As New Utilities
    Dim destFldr As Outlook.Folder
    Set destFldr = ut.SelectOutlookFolder
    If Not IsNothing(destFldr) Then
        lbl_DestinationFolderPath.Caption = destFldr.FolderPath
        bDirty = True
    End If

End Sub

Private Sub chkbx_EnableApp_Change()
    bDirty = True
End Sub

Private Sub chkbx_ShowTaskPaneOnStartup_Change()
    bDirty = True
End Sub

Private Sub rb_DoNotMoveMail_Change()
    bDirty = True
    If rb_DoNotMoveMail.value Then
        EnableEvents False
    Else
        EnableEvents True
    End If
End Sub

Private Sub rb_MoveMailAll_Change()
    bDirty = True
    If rb_MoveMailAll.value Then
        EnableEvents False
    Else
        EnableEvents True
    End If
End Sub

Private Sub rb_MoveMailOnSpecific_Change()
    bDirty = True
    If rb_MoveMailOnSpecific.value Then
        EnableEvents True
    End If
End Sub

' - - Constructor

Private Sub UserForm_Initialize()

    Me.Caption = "Options - " & Commands.AppName
   
    Call GetUI
    
    Call Status
        
    bDirty = False
    
End Sub

Private Sub UserForm_Terminate()
    If bDirty Then
        MsgBox "Do you want to save your setting changes?"
    End If
End Sub

' - - Methods

' - - Supporting Methods

Private Sub GetUI()

    Dim stgs As New Settings
    
    ' General
    chkbx_ShowTaskPaneOnStartup.value = stgs.ShowOnStartup
    chkbx_EnableApp.value = stgs.EnableAppEvents
    
    ' Events
    Dim bMove As Boolean
    bMove = stgs.AutoMove
    If Not bMove Then
        rb_DoNotMoveMail.value = True
        EnableEvents False
    Else
        Dim bSpecific As Boolean
        bSpecific = stgs.MoveOnSpecificEvents
        If Not bSpecific Then
            rb_MoveMailAll.value = True
        Else
            rb_MoveMailOnSpecific.value = True
        End If
    End If
    chkbx_DeferToAppt.value = stgs.MoveOnDeferToAppt
    chkbx_DeferToTask.value = stgs.MoveOnDeferToTask
    chkbx_Delegate.value = stgs.MoveOnDelegate
    chkbx_FileInDrawer.value = stgs.MoveOnFileInDrawer
    chkbx_OnReply.value = stgs.MoveOnReply
    
    ' Location
    lbl_DestinationFolderPath.Caption = stgs.DestinationFolder
    
    chkbx_IgnoreSentItemsMove.value = stgs.IgnoreSentMailMove

End Sub

Private Sub SaveSettings()

        Dim stgs As New Settings
        
        ' General
        stgs.ShowOnStartup = chkbx_ShowTaskPaneOnStartup.value
        stgs.EnableAppEvents = chkbx_EnableApp.value
        
        ' Events
        If rb_DoNotMoveMail.value Then
            stgs.AutoMove = False
        Else
            stgs.AutoMove = True
            If rb_MoveMailAll.value Then
                stgs.MoveOnSpecificEvents = False
                ' Set 5Ds flags to True
                stgs.MoveOnDeferToAppt = True
                stgs.MoveOnDeferToTask = True
                stgs.MoveOnDelegate = True
                stgs.MoveOnFileInDrawer = True
                stgs.MoveOnReply = True
            Else
                stgs.MoveOnSpecificEvents = True
                ' Gather 5D flags
                stgs.MoveOnDeferToAppt = chkbx_DeferToAppt.value
                stgs.MoveOnDeferToTask = chkbx_DeferToTask.value
                stgs.MoveOnDelegate = chkbx_Delegate.value
                stgs.MoveOnFileInDrawer = chkbx_FileInDrawer.value
                stgs.MoveOnReply = chkbx_OnReply.value
            End If
        End If
        
        ' Location
        stgs.DestinationFolder = lbl_DestinationFolderPath.Caption
        stgs.IgnoreSentMailMove = chkbx_IgnoreSentItemsMove.value
        
        ' Commit Settings
        stgs.Save

End Sub

Private Sub EnableEvents(ByVal bl As Boolean)

    chkbx_DeferToTask.Enabled = bl
    chkbx_DeferToAppt.Enabled = bl
    chkbx_Delegate.Enabled = bl
    chkbx_FileInDrawer.Enabled = bl
    chkbx_OnReply.Enabled = bl

End Sub

Private Sub Status(Optional ByVal msg As String = "")
    If Len(msg) = 0 Then
        Me.lbl_Status.Visible = False
    Else
        Me.lbl_Status.Visible = True
        Me.lbl_Status.Caption = msg
    End If
    DoEvents
End Sub

