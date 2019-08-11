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

Private Sub btn_IndexFolders_Click()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":btn_IndexFolders_Click"
    
    On Error GoTo ThrowException
    
    Call Status("Working...")
    
    strTrace = "Get Outlook folders."
    Dim ut As New Utilities
    Dim fldrs As Folders
    Set fldrs = ut.IndexFolders()
    If fldrs Is Nothing Then
        strTrace = "Failed to query the Outlook dataStore folder structure."
        GoTo ThrowException
    End If
    If fldrs.Count = 0 Then
        strTrace = "User has no folders in their folder tree."
        MsgBox strTrace
        Exit Sub
    End If
    
    Dim ldb As New dsDataStore
    ldb.Connect
    
    strTrace = "Clear the current folder set."
    ldb.ClearEntireCollection "Folder"
    
    strTrace = "Update the local datastore with the indexed folders from Outlook."
    Dim fCnt As Integer
    fCnt = 0
    Dim f As fmeFolder
    For Each f In fldrs.Items
    
        If Not Len(f.EntryId) = 0 Then

            ' Add folder to local datastore
            If ldb.Insert(f, "Folder") Then
                strTrace = "Indexed a new folder: " & f.Path
                fCnt = fCnt + 1
            Else
                strTrace = "Failed to insert a new folder (" & f.Path & ") into the datastore."
                LogMessage strTrace, strRoutine
            End If
                
        Else
            strTrace = "Ignoring Outlook folder - empty path, id: " & f.id
            LogMessage strTrace, strRoutine
        End If
                
    Next
    
    strTrace = "Indexed " & fCnt & " folders from Outlook."
    LogMessage strTrace, strRoutine
    
    ldb.Disconnect
    
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
    
    Dim pm As New ProjectManager
    Dim myList As Projects
    Set myList = pm.ImportFromMCL
    If myList.Count = 0 Then
        strTrace = "No projects found in the Outlook Master Category List."
        MsgBox strTrace
        Exit Sub
    End If
    
    Dim ldb As New dsDataStore
    ldb.Connect
    
    Dim arProj As ArrayList
    Set arProj = ldb.GetEntireCollection("Project")
    If arProj Is Nothing Then
        Exit Sub
    End If
    
    Dim pCnt As Integer
    pCnt = 0
    
    Dim p As fmeProject
    Dim ip As fmeProject
    Dim bFnd As Boolean
    
    For Each ip In myList.Items
        bFnd = False
        For Each p In arProj
           If LCase(ip.Subject) = LCase(p.Subject) Then
                bFnd = True
                Exit For
            End If
        Next
        
        If Not bFnd Then
            ldb.Insert ip, "Project"
            pCnt = pCnt + 1
        End If
        
    Next
        
    ldb.Disconnect
       
SkipOut:
    MsgBox "Imported " & myList.Items.Count & " potential projects, saved " & pCnt & " to the datastore."
    
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
    If rb_DoNotMoveMail.Value Then
        EnableEvents False
    Else
        EnableEvents True
    End If
End Sub

Private Sub rb_MoveMailAll_Change()
    bDirty = True
    If rb_MoveMailAll.Value Then
        EnableEvents False
    Else
        EnableEvents True
    End If
End Sub

Private Sub rb_MoveMailOnSpecific_Change()
    bDirty = True
    If rb_MoveMailOnSpecific.Value Then
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
    chkbx_ShowTaskPaneOnStartup.Value = stgs.ShowOnStartup
    chkbx_EnableApp.Value = stgs.EnableAppEvents
    
    ' Events
    Dim bMove As Boolean
    bMove = stgs.AutoMove
    If Not bMove Then
        rb_DoNotMoveMail.Value = True
        EnableEvents False
    Else
        Dim bSpecific As Boolean
        bSpecific = stgs.MoveOnSpecificEvents
        If Not bSpecific Then
            rb_MoveMailAll.Value = True
        Else
            rb_MoveMailOnSpecific.Value = True
        End If
    End If
    chkbx_DeferToAppt.Value = stgs.MoveOnDeferToAppt
    chkbx_DeferToTask.Value = stgs.MoveOnDeferToTask
    chkbx_Delegate.Value = stgs.MoveOnDelegate
    chkbx_FileInDrawer.Value = stgs.MoveOnFileInDrawer
    chkbx_OnReply.Value = stgs.MoveOnReply
    
    ' Location
    lbl_DestinationFolderPath.Caption = stgs.DestinationFolder

End Sub

Private Sub SaveSettings()

        Dim stgs As New Settings
        
        ' General
        stgs.ShowOnStartup = chkbx_ShowTaskPaneOnStartup.Value
        stgs.EnableAppEvents = chkbx_EnableApp.Value
        
        ' Events
        If rb_DoNotMoveMail.Value Then
            stgs.AutoMove = False
        Else
            stgs.AutoMove = True
            If rb_MoveMailAll.Value Then
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
                stgs.MoveOnDeferToAppt = chkbx_DeferToAppt.Value
                stgs.MoveOnDeferToTask = chkbx_DeferToTask.Value
                stgs.MoveOnDelegate = chkbx_Delegate.Value
                stgs.MoveOnFileInDrawer = chkbx_FileInDrawer.Value
                stgs.MoveOnReply = chkbx_OnReply.Value
            End If
        End If
        
        ' Location
        stgs.DestinationFolder = lbl_DestinationFolderPath.Caption
        
        
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

