VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' no public withevents

' --- Fields

Private Const rootClass As String = "ThisOutlookSession"

' --- Class Internal Variables

Private WithEvents frm As FME_Pane
Attribute frm.VB_VarHelpID = -1

Private frmProjects As frm_ProjectPicker
Attribute frmProjects.VB_VarHelpID = -1

Private frmHandle As Long
Private frmParentHandle As Long

Public f_queue As MemoryLogger

' Private WithEvents ReplyButton As Office.CommandBarButton

''' Collection containing the Outlook.MailItems that have been touched by the User
Private MailItems As ArrayList

' --- Properties

''' Intra App Messaging Queue
Public Property Get MessageQueue() As MemoryLogger
    If f_queue Is Nothing Then Set f_queue = New MemoryLogger
    Set MessageQueue = f_queue
End Property

' currently selected OutlookItem
Public Property Get CurrentItem() As Object

    Dim myItem As Object

    ' Use method to get current explorer selection
    Dim xSelection As Outlook.Selection
    Set xSelection = Selection
    If Not xSelection Is Nothing Then
        If xSelection.Count > 0 Then
            Set myItem = xSelection.Item(1)
        End If
    Else
        Debug.Print ("Selection was null.")
    End If
    
    Set CurrentItem = myItem
    
End Property

''' Pointer to FME Task Pane
'Public Property Get FMEPane() As FME_Pane
'    Set FMEPane = frm
'End Property

' --- Event Handlers

Private Sub Application_ItemLoad(ByVal Item As Object)
    ' Captures the mailItem for later event processing
    Call ActOnItemLoad(Item, "AppStart")
    ' Clears the mail flag from the FME_Task Form if
    '   the user goes directly to the Outlook Mail list
    If Not frm Is Nothing Then Call frm.ClearMail
End Sub

Private Sub LogItem(ByVal mi As Outlook.MailItem)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":LogItem"

    strTrace = "Identify mail item: " & mi.EntryId
    LogMessage strTrace, strRoutine
    
End Sub

Private Sub Application_NewMail()
    ' If Not IsNothing(frm) Then Call frm.NewMail
End Sub

Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Application_NewMailEx"
    
    On Error GoTo ThrowException
    
    Dim stgs As New Settings
    
    If Not stgs.EnableAppEvents Then Exit Sub

    If Not IsNothing(frm) Then
        strTrace = "Processing mail Collection: " & vbCrLf & EntryIDCollection
        Call frm.NewMail
        LogMessage strTrace, strRoutine
    Else
        strTrace = "Received new mail signal - unable to pass on, " & EntryIDCollection
        LogMessage strTrace, strRoutine
    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Private Sub frm_FormClosing()
    LogMessage "FME Pane closing.", "frm_FormClosing"
End Sub

Private Sub frm_FormClosed()
    Set frm = Nothing
    LogMessage "FME Pane closed.", "frm_FormClosed"
End Sub

Private Sub Application_ItemSend(ByVal Item As Object, cancel As Boolean)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Application_ItemSend"
    
    On Error GoTo ThrowException
    
    Dim ut As Utilities
    Dim tm As TaskManager
    Dim stgs As New Settings
    
    If Not stgs.EnableAppEvents Then Exit Sub
    
    If Not cancel Then
    
        If TypeOf Item Is Outlook.MailItem Then
            ' Cast the Item
            Dim oMail As Outlook.MailItem
            Set oMail = Item
            
            If Len(oMail.FlagRequest) > 0 Then
                strTrace = "Sending a marked (" & oMail.FlagRequest & ") email."
                
                ' Set up Managers
                Set ut = New Utilities
                Set tm = New TaskManager
                
                ' Create the new tracking task
                '   need to set categories and subject to match
                Dim t As Outlook.TaskItem
                Set t = ut.CreateRelatedTask(oMail, True)
                
                If LCase(oMail.FlagRequest) = "taskrequest" Then
                    ' Link task to contact
                    ' Link task to mail
                    tm.LinkMailToTask oMail, t
                    ' Show the task
                    t.Display
                End If
                If LCase(oMail.FlagRequest) = "followupflag" Then
                    ' Link task to mail
                    tm.LinkMailToTask oMail, t
                    ' Show the task
                    t.Display
                End If
            Else
                strTrace = "Sending an unmarked email."
            End If
    
        Else
            strTrace = "Not a MailItem in the ItemSend event, hmmmm!"
            LogMessage strTrace, strRoutine
        End If
        
        LogMessage strTrace, strRoutine
        
    Else
        strTrace = "Cancel a send."
        LogMessage strTrace, strRoutine
    End If
    
    GoTo Finally
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
Finally:
    Set ut = Nothing
    Set tm = Nothing
    
End Sub

' --- Session Handlers

Private Sub Application_Startup()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Application_Startup"
    
    strTrace = "Starting FocusMe for VBA macro code..."
    LogMessage strTrace, strRoutine

   ' Set ReplyButton = Application.ActiveExplorer.CommandBars.FindControl(, 354)
   
   ' Inter App messaging service
   Set f_queue = New MemoryLogger
   
   ' Initialize Collections
   Set MailItems = New ArrayList
   
    ' Get Settings
    Dim stgs As New Settings
    Dim bShow As Boolean
    bShow = stgs.ShowOnStartup
   
    If bShow Then
        Set frm = New FME_Pane
        frmHandle = HWndOfUserForm(frm)
        frmParentHandle = GetFormParent(frm)
        frm.Show
        
        strTrace = "FME Pane Handle: " & frmHandle & ", Parent Handle: " & frmParentHandle
        LogMessage strTrace, strRoutine
    Else
        strTrace = "Show on Startup flag is False."
        LogMessage strTrace, strRoutine
    End If
   
End Sub

Private Sub Application_Quit()
    Set frm = Nothing
    Set frmProjects = Nothing
End Sub

' --- Methods

Public Sub StartTaskList()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":StartTaskList"
    
    On Error GoTo ThrowException

    If IsNothing(frm) Then
        strTrace = "Starting a new FME pane."
    
        Dim hwnd As Long
        hwnd = FindExplorerWindow(ActiveExplorer.Caption)
        
        Set frm = New FME_Pane
        Dim b As Boolean
        b = SetFormParent(frm, frmParentHandle)
        If b Then
            frm.Show
        Else
            strTrace = "WARNING: Failed to set the form's parent - show any way."
            LogMessage strTrace, ""
            frm.Show
        End If
    Else
        strTrace = "Showing existing window."
        Set frm = New FME_Pane
        frm.Show
    End If
    
    LogMessage strTrace, strRoutine
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Shows the Outlook Explorer Window at the Mail Pane
Public Sub GoToMailPane()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GoToMailPane"
    
    On Error GoTo ThrowException

    strTrace = "Navigating to Mail Module View."
    Dim oPane As Outlook.NavigationPane
    Dim oModule As Outlook.NavigationModule

    ' Get Current Explorer's Navigation Pane
    Set oPane = Me.ActiveExplorer.NavigationPane

    ' Find MailModule pointer
    Set oModule = oPane.Modules.GetNavigationModule(Outlook.OlNavigationModuleType.olModuleMail)
    ' Set navigation pane to the mail module
    Set oPane.CurrentModule = oModule
    
    ' Restore the Outlook Explorer if needed
    If Me.ActiveExplorer.WindowState = olMinimized Then
        Me.ActiveExplorer.Activate
    End If

    'If Not LocalDataStore.Explorers.IsAnExplorerActive Then
    '    olApp.ActiveExplorer.Activate()
    'End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Public Sub StartAllProjects()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":StartAllProjects"
    
    On Error GoTo ThrowException
    
    If IsNothing(frmProjects) Then
        ' Start a new window
        strTrace = "Starting a new Projects Window."
        Set frmProjects = New frm_ProjectPicker
        frmProjects.Show
    Else
        ' Restore current window
        strTrace = "Restoring the current Projects Window."
        Set frmProjects = New frm_ProjectPicker
        frmProjects.Show
    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Starts the Timecard Window
Public Sub StartTimecard(Optional ByVal startDate As Date = #1/1/1970#, _
                         Optional ByVal endDate As Date = #1/1/1970#)
                         
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":StartTimecard"
    
    On Error GoTo ThrowException
    
    Dim objView As CalendarView
    Dim dates As Variant
    
    If YearPart(startDate) = 1970 And YearPart(endDate) = 1970 Then
        ' Check if the current view is a calendar view.
        If Me.ActiveExplorer.CurrentView.ViewType = olCalendarView Then
            strTrace = "Getting dates from Calendar View."
            
            Set objView = Me.ActiveExplorer.CurrentView
            dates = objView.DisplayedDates
            startDate = dates(LBound(dates))
            endDate = dates(UBound(dates))
            
        Else
            strTrace = "View not a CalendarView, and no dates specified."
            ' Else use Today and 1 week prior
            endDate = GetDateOnly(Now)
            startDate = DateAdd(GetDatePartFormat(DateInterval.Day), -7, endDate)
        End If
    End If
    If YearPart(startDate) = 1970 Or YearPart(endDate) = 1970 Then
        If Not YearPart(startDate) = 1970 Then
            strTrace = "StartDate specified, EndDate not specified."
            ' Use StartDate and 1 week prior
            endDate = startDate
            startDate = DateAdd(GetDatePartFormat(DateInterval.Day), -7, endDate)
        Else
            strTrace = "StartDate not specified, ignore EndDate."
            ' Use Today and 1 week prior
            endDate = GetDateOnly(Now)
            startDate = DateAdd(GetDatePartFormat(DateInterval.Day), -7, endDate)
        End If
    End If
    
    ' Start the Timecard Dialog HERE
    Dim frmTC As New frm_TimeCard
    frmTC.startDate = startDate
    frmTC.endDate = endDate
    frmTC.Refresh
    frmTC.Show
    
'    Dim tr As New TimeRecords
'    tr.Load startDate, endDate
    
    'strTrace = "Found " & tr.Items.Count & " appointments in window: " & _
                GetDateTimeStamp(startDate) & " - " & GetDateTimeStamp(endDate)
    'LogMessage strTrace, strRoutine
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
                         
End Sub

'
' Returns the Active Explorer's Selection
'
Public Function Selection() As Outlook.Selection

    Dim app As Outlook.Application
    Dim oNS As Outlook.NameSpace
    Set oNS = GetNamespace("MAPI")

    Dim myExplorer As Outlook.Explorer
    Set app = oNS.Application
    Set myExplorer = app.ActiveExplorer
    If Not myExplorer.Selection Is Nothing Then
        Set Selection = myExplorer.Selection
    Else
        Debug.Print ("Selection was null.")
    End If

End Function

Private Function GetAPICall() As Object

Dim ie As New InternetExplorer
ie.Visible = False
ie.navigate ("API Get Call")
Do
    DoEvents
Loop Until ie.readyState = READYSTATE_COMPLETE

Dim payload As Object
Set payload = ie.document

Set GetAPICall = payload

' ie.navigate2 allows headers and postdata

End Function


' - - Supporting Methods

Private Sub ActOnItemLoad(ByVal Item As Object, _
                 Optional ByVal caller As String = "")

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ActOnItemLoad"
    
    On Error GoTo ThrowException
    
    Dim stgs As New Settings
    
    If Not stgs.EnableAppEvents Then Exit Sub
    
    If IsNothing(MailItems) Then Set MailItems = New ArrayList
    
    If Len(caller) > 0 Then strRoutine = strRoutine & "-" & caller
    
    If TypeOf Item Is Outlook.MailItem Then
        
        Dim cMail As New fmeMailItem
        cMail.Load Item
        MailItems.Add cMail
        
        strTrace = "Loaded mail item: " & cMail.id
        LogMessage strTrace, strRoutine
    Else
        strTrace = "Unrecognized item loaded into memory."
    End If

    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub



