VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TaskManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' - Fields

Private Const rootClass As String = "TaskManager"

Dim app As Outlook.Application
Dim oNS As Outlook.NameSpace

Private WithEvents singleInspector As Outlook.Inspector
Attribute singleInspector.VB_VarHelpID = -1
Private ut As Utilities
Attribute ut.VB_VarHelpID = -1
Private WithEvents pm As ProjectManager
Attribute pm.VB_VarHelpID = -1

' - -  Settings
Private stgs As Settings

Private WithEvents myListView As MSComctlLib.ListView
Attribute myListView.VB_VarHelpID = -1

Dim sortingColumn As enuSortOn
Dim sortingDirection As enuSortDirection

' - Events

''' Thrown when a Task is changed by this class
Public Event ItemUpdated()
Public Event ItemAdded()
Public Event collectionUpdated()

' - Properties
Dim myItems As Tasks
Dim selItem As Outlook.TaskItem
Dim selItems As Tasks

Dim lvCheckbox As Boolean
Dim lvMultiSelect As Boolean
Dim f_SuspendUI As Boolean
Dim f_lastQuery As String
Dim f_listFiltered As Boolean


''' Latest collection of Tasks
Public Property Set Items(ByVal coll As Tasks)
    Set myItems = coll
End Property
Public Property Get Items() As Tasks
    Set Items = myItems
End Property

''' Current Task being changed
Public Property Get SelectedItem() As Outlook.TaskItem
    Set SelectedItem = selItem
End Property

''' Currently selected Collection of Tasks
Public Property Get SelectedItems() As Tasks
    Set SelectedItems = selItems
End Property

''' Managed ListView
Public Property Set ListView(ByVal lv As MSComctlLib.ListView)
    Set myListView = lv
End Property
Public Property Get ListView() As MSComctlLib.ListView
    Set ListView = myListView
End Property

''' Sets / Gets the 'show checkboxes' flag for the ListView
Public Property Let ListViewCheckBox(ByVal b As Boolean)
    lvCheckbox = b
End Property
Public Property Get ListViewCheckBox() As Boolean
    ListViewCheckBox = lvCheckbox
End Property

''' Sets / Gets the ability to select more than one Project
Public Property Let ListViewMultiSelect(ByVal b As Boolean)
    lvMultiSelect = b
End Property
Public Property Get ListViewMultiSelect() As Boolean
    ListViewMultiSelect = lvMultiSelect
End Property

''' Sets / Gets flag for ignoring UI event calls
''' - Useful when more than one controller managing
''      a shared UI control, e.g. ListView
Public Property Let SuspendUIEvents(ByVal b As Boolean)
    f_SuspendUI = b
End Property
Public Property Get SuspendUIEvents() As Boolean
    SuspendUIEvents = f_SuspendUI
End Property

' - Event Handlers

''' - Outlook Inspector Handlers
Private Sub singleInspector_Activate()

End Sub

Private Sub singleInspector_Close()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":singleInspector_Close"
    
    On Error GoTo ThrowException
    
    ' Check last mod date - this will error if task
    '   has been deleted
    strTrace = selItem.LastModificationTime
    
    Dim dteSaved As Date
    dteSaved = selItem.LastModificationTime
    
    If IsDateNone(dteSaved) Then
        ' Ignore - item opened and never saved
    Else
        ' Look in the collection
        Dim idx As Integer
        idx = IsItemInCollection(selItem)
        If idx >= 0 Then
            ' Update collection & UI
            UpdateItem selItem, idx
            RaiseEvent ItemUpdated
            
        Else
            ' Update collection
            myItems.AddItem selItem
            myItems.Sort DueDate, Ascending
            RaiseEvent ItemAdded
            
            ' Update Managed UI
            AddListViewItem selItem
        End If

    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err
    
End Sub

' - - ListView Event Handlers

Private Sub myListView_AfterLabelEdit(Cancel As Integer, NewString As String)

    If f_SuspendUI Then Exit Sub

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":myListView_AfterLabelEdit"
   
    Dim li As ListItem
    Set li = myListView.SelectedItem
    selItem.Subject = NewString
    selItem.Save
    
    strTrace = "Updated subject for: " & selItem.Subject
    LogMessage strTrace, strRoutine

End Sub

Private Sub myListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":myListView_ColumnClick"
    
    If f_SuspendUI Then Exit Sub
    
    Dim selectSortOn As enuSortOn
    If InStr(1, LCase(ColumnHeader.key), "subject") Then selectSortOn = Subject
    If InStr(1, LCase(ColumnHeader.key), "duedate") Then selectSortOn = DueDate
    
    If Not sortingColumn = selectSortOn Then
        ' Sorting Column changed
        sortingColumn = selectSortOn
    Else
        ' Sorting Direction changed
        If sortingDirection = Ascending Then
            sortingDirection = Descending
        Else
            sortingDirection = Ascending
        End If
    End If
    RefreshListView sortingColumn, sortingDirection

End Sub

Private Sub myListView_DblClick()

    Dim strTrace As String
    strTrace = ""
    Dim strRoutine As String
    strRoutine = rootClass & ":myListView_DblClick"
    
    If f_SuspendUI Then Exit Sub
    
    Dim ut As Utilities
    Set ut = New Utilities
    
    ' Get selected Item
    Dim li As ListItem
    Set li = myListView.SelectedItem
    
    Dim eid As String
    eid = li.key
    Dim t As Outlook.TaskItem
    Set t = ut.GetOutlookItemFromID(eid)
    
    OpenTask t

End Sub

Private Sub myListView_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    Dim strTrace As String
    strTrace = ""
    Dim strRoutine As String
    strRoutine = rootClass & ":myListView_ItemCheck"
    
    If f_SuspendUI Then Exit Sub

    On Error GoTo ThrowException

    Dim eid As String
    eid = Item.key
    If Len(eid) = 0 Then
        strTrace = "Failed to get the Outlook item's EntryId."
        GoTo ThrowException
    End If
    
    ' Get Outlook Utilities
    Dim ut As Utilities
    Set ut = New Utilities
        
    Dim t As Outlook.TaskItem
    Set t = ut.GetOutlookItemFromID(eid)
    If Not IsNothing(t) Then
        Dim bItmChanged As Boolean
        bItmChanged = False
        If Item.checked Then
            If Not t.Complete Then
                ' User marked item completed
                MarkComplete t
                bItmChanged = True
            End If
        Else
            If t.Complete Then
                ' User marked item as incomplete
                MarkIncomplete t
                bItmChanged = True
            End If
        End If
        If bItmChanged Then UpdateListViewItem t
    End If
    
    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Sub myListView_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Dim strTrace As String
    strTrace = ""
    Dim strRoutine As String
    strRoutine = rootClass & ":myListView_ItemClick"
    
    If f_SuspendUI Then Exit Sub

    Dim eid As String
    eid = Item.key
    Set selItem = ut.GetOutlookItemFromID(eid)
    If IsNothing(selItem) Then
        strTrace = "Failed to find an Outlook Task for a known good EntryID: '" & eid & "'."
        LogMessage "WARNING: " & strTrace, strRoutine
    Else
        strTrace = "Selected a taskItem."
        LogMessage "INFORMATION: " & strTrace, strRoutine
    End If
    
End Sub

Private Sub myListView_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)

    Dim strTrace As String
    strTrace = ""
    Dim strRoutine As String
    strRoutine = rootClass & ":myListView_MouseDown"
    
    If f_SuspendUI Then Exit Sub
        
    On Error GoTo ThrowException

    If Button = 1 Then

    End If
    If Button = 2 Then
       
        ' Get the pointed to listview item
        Dim li As ListItem
        Set li = LV_GetItemAt(X, y)
        If Not IsNothing(li) Then
        
            ' Select Item if not already selected
            Dim sLi As ListItem
            Set sLi = myListView.SelectedItem
            If Not sLi.key = li.key Then
                ' DIDN"T LOOK RIGHT
                ' Select the right clicked item
                ' Set lv_Tasks.selectedItem = li
                ' DoEvents
            End If
            
            ' Get selected Task
            Dim t As Outlook.TaskItem
            Set t = ut.GetOutlookItemFromID(li.key)
            If IsNothing(t) Then
                strTrace = "WARNING: Failed to find a selected task, id: " & li.key
                GoTo ThrowException
            End If
            strTrace = "Right clicked on task: " & t.Subject
    
            strTrace = "Create the context menu."
            Dim cMenu As ContextMenu
            Dim cMenu2 As ContextMenu
            
            strTrace = "Prepare context menu location."
            Dim xSgl As Single
            xSgl = X
            Dim ySgl As Single
            ySgl = y
            
            strTrace = "Present context menu."
            Dim bUseOld As Boolean
            bUseOld = stgs.UseContextMenuWindows
            
            Dim id As Long
                        
            If bUseOld Then
                Set cMenu = GetContextMenu(t)
                id = ShowPopup(myListView, cMenu, xSgl, ySgl)
            Else
                Set cMenu2 = GetContextMenuTwoTier(t)
                id = ShowMenu(cMenu2)
            End If
       
            strTrace = "Selected menu item: " & id
            LogMessage strTrace, strRoutine
            
            Call ProcessContextMenuSelection(id, t)
        
        Else
            strTrace = "WARNING: failed to find a ListItem at the loc: (" & X & "," & y & ")."
            LogMessage strTrace, strRoutine
        End If
              
    End If

    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Sub myListView_OLECompleteDrag(Effect As Long)

    If f_SuspendUI Then Exit Sub

    ' Effect = 5
End Sub

Private Sub myListView_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
    
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":myListView_OLEDragDrop"
    
    On Error GoTo ThrowException
    
    If f_SuspendUI Then Exit Sub

    'The Format numbers used in the OLE DragDrop data structure, are:
    '   Text = 1 (vbCFText)
    '   Bitmap = 2 (vbCFBitmap)
    '   Metafile = 3
    '   Emetafile = 14
    '   DIB = 8
    '   palette = 9
    '   Files = 15 (vbCFFiles)
    '   RTF = -16639
    
    ' Get the target listview item
    Dim t As Outlook.TaskItem
    Set t = Nothing
    Dim li As ListItem
    Set li = LV_GetItemAt(X, y, 20)
    If Not IsNothing(li) Then
        ' Get selected Task
        Dim eid As String
        eid = li.key
        Set t = ut.GetOutlookItemFromID(eid)
    Else
        strTrace = "Failed to identify target task."
        GoTo ThrowException
    End If
        
    ' TEXT
    '   Outlook.MailItems - tab delimited with headers {Starts with From}
    '   Outlook.ContactItems - tab delimited with headers {starts with 'Full Name'}
    '   Outlook.AppointmentItems = appointment title {no distinuquishing factors}
    '   Outlook.TaskItems = tab delimited table with headers {Starts with 'Task Subject'}
    If Data.GetFormat(1) Then
        ' Process the text drop
        strTrace = Data.getData(1)
        ProcessDroppedText strTrace, t
    End If
    If Data.GetFormat(2) Then
        strTrace = "Bitmap dropped."
        LogMessage strTrace, strRoutine
    End If
    If Data.GetFormat(3) Then
        strTrace = "Metafile dropped."
        LogMessage strTrace, strRoutine
    End If
    If Data.GetFormat(8) Then
        strTrace = "DIB dropped."
        LogMessage strTrace, strRoutine
    End If
    If Data.GetFormat(9) Then
        strTrace = "Palette dropped."
        LogMessage strTrace, strRoutine
    End If
    If Data.GetFormat(14) Then
        strTrace = "Emetafile dropped."
        LogMessage strTrace, strRoutine
    End If
        
    ' FILES - Array of Names from WinExplorer
    If Data.GetFormat(15) Then
        strTrace = "File collection dropped, file count: " & Data.Files.count
        LogMessage strTrace, strRoutine
    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
      
End Sub

Private Sub myListView_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single, State As Integer)
    
    If f_SuspendUI Then Exit Sub
    
    On Error Resume Next
    Effect = 1
    Set myListView.DropHighlight = lv_Tasks.HitTest(X * 15, y * 20)
End Sub

' - Constructor

Private Sub Class_Initialize()
    
    ' Initialize variables
    Set myItems = New Tasks
    
    Set oNS = GetNamespace("MAPI")
    Set app = oNS.Application
    Set ut = New Utilities
    Set stgs = New Settings
    
    Set pm = New ProjectManager
    
End Sub

Private Sub Class_Terminate()
    Set myItems = Nothing
    Set selItem = Nothing
    Set selItems = Nothing
    Set myListView = Nothing
    
    Set oNS = Nothing
    Set app = Nothing
    Set ut = Nothing
    Set stgs = Nothing
End Sub


' - Methods

''' Load all Outlook Tasks
Public Sub Load()

    'Set myItems.Items = GetOutlookTasks()
    'Call Refresh
    
    LoadDASL ""
    
    ' Inform listeners
    RaiseEvent collectionUpdated
    
End Sub

Public Sub LoadByProject(ByVal Name As String, _
                Optional ByVal includeCompleted As Boolean = False)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":LoadByProject"
    
    If Len(Name) > 0 Then
        Dim strQuery As String
        
        If Not includeCompleted Then
            ' Not complete and matching the category name
            strQuery = "@SQL=""" & PR_TASK_COMPLETE & """ = 0 AND """ & PR_CATEGORIES & """ like '%" & Name & "%'"
        Else
            ' Any task matching the category name
            strQuery = "@SQL=""" & PR_CATEGORIES & """ like '%" & Name & "%'"
        End If
        
        LoadDASL strQuery
        
    Else
        strTrace = "An empty project name encountered."
        LogMessage "WARNING: " & strTrace, strRoutine
    End If
    
End Sub

''' Load tasks according to the special filter
Public Sub LoadSpecial(ByVal fltr As enuTaskFilters)

' Daily = 1
' Master = 2
' NoCategory = 3
' HighPriority = 4
' PastDue = 5
' InProgress = 6
' Waiting = 7
' Deferred = 8

    Dim strQuery As String
    Select Case fltr
        Case enuTaskFilters.Daily 'Not Complete with a DueDate
            strQuery = "@SQL=(""" & PR_TASK_COMPLETE & """ = 0 AND NOT(""" & PR_TASK_DUEDATE & """ IS NULL))"
        Case enuTaskFilters.Master 'Not Complete without a DueDate
            strQuery = "@SQL=(""" & PR_TASK_COMPLETE & """ = 0 AND """ & PR_TASK_DUEDATE & """ IS NULL)"
        Case enuTaskFilters.NoCategory 'Not Complete with empty Categories
            strQuery = "@SQL=""" & PR_TASK_COMPLETE & """ = 0 AND """ & PR_CATEGORIES & """ IS NULL"
        Case enuTaskFilters.HighPriority 'Not Complete and High Priority
            strQuery = "@SQL=""" & PR_TASK_COMPLETE & """ = 0 AND """ & PR_TASK_PRIORITY & """ = 2"
        Case enuTaskFilters.PastDue 'Not Complete and Past Due
            strQuery = "@SQL=(""" & PR_TASK_COMPLETE & """ = 0 AND """ & PR_TASK_DUEDATE & """ < 'Today')"
        Case enuTaskFilters.InProgress 'Not Complete and Status is In Progress
            strQuery = "@SQL=""" & PR_TASK_COMPLETE & """ = 0 AND """ & PR_TASK_STATUS & """ = 1"
        Case enuTaskFilters.Waiting 'Not Complete and Status is Waiting
            strQuery = "@SQL=""" & PR_TASK_COMPLETE & """ = 0 AND """ & PR_TASK_STATUS & """ = 3"
        Case enuTaskFilters.Deferred 'Not Complete and Status is Deferred
            strQuery = "@SQL=""" & PR_TASK_COMPLETE & """ = 0 AND """ & PR_TASK_STATUS & """ = 4"
           
    End Select
    
    LoadDASL strQuery
    
End Sub

''' Load tasks using the specified DASL Query
Public Sub LoadDASL(ByVal daslFilter As String)

    Dim arList As ArrayList
    Set arList = GetOutlookTasks(daslFilter, True)
    myItems.Items = arList
    myItems.Sort DueDate, Ascending
    
    ' Capture query
    f_lastQuery = daslFilter
    
    ' Update the Managed UI
    Call Refresh
    
    ' Inform listeners
    RaiseEvent collectionUpdated
    
End Sub

' Sorts the current Outlook TaskItem collection
Public Sub Sort(ByVal srt As enuSortOn, ByVal dir As enuSortDirection)
    myItems.Sort srt, dir
    
    ' Update UI
    Call Refresh
    
End Sub

' Filters current list using the Subject property
Public Sub Filter(ByVal strFilter As String, Optional resetFilter As Boolean = True)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Filter"
    
    On Error GoTo ThrowException

    If Len(strFilter) = 0 Then
        ' Do nothing
        strTrace = "WARNING: Empty filter encountered, ignoring request."
        LogMessage strTrace, strRoutine
        Exit Sub
    End If
    
    ' If specified clears a previous filter
    Dim arList As ArrayList
    If resetFilter And f_listFiltered Then
        strTrace = "Clearing previous filter."
        Set arList = GetOutlookTasks(f_lastQuery, True)
        myItems.Items = arList
        myItems.Sort DueDate, Ascending
    End If
    
    strTrace = "Setting up filter: " & strFilter
    Dim fltrList As New ArrayList
    Dim t As Outlook.TaskItem
    Dim searchIn As String
    Dim i As Integer
    For Each t In myItems.Items
        searchIn = UCase(t.Subject)
        If Contains(UCase(strFilter), searchIn) Then
            fltrList.Add t
        End If
    Next
    
    ' Reset internal collection
    strTrace = "Resetting to a filtered list."
    myItems.Items = fltrList
    
    ' Update the Managed UI
    Call Refresh
    
    ' Inform listeners
    RaiseEvent collectionUpdated
    
    f_listFiltered = True

    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

' Clears a filtered list to the last DASL query and updates the managed UI
Public Sub ClearFilter()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ClearFilter"
    
    On Error GoTo ThrowException

    If Not f_listFiltered Then
        strTrace = "Attempting to clear an unfiltered list."
        GoTo ThrowException
    End If
    
    LoadDASL f_lastQuery

    f_listFiltered = False
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

' - - TaskItem Commands

Public Sub OpenTask(Optional ByVal t As Outlook.TaskItem = Nothing)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":OpenTask"
    
    On Error GoTo ThrowException

    If t Is Nothing Then
        Set t = selItem
    Else
        Set selItem = t
    End If
    
    If IsNothing(t) Then
        strTrace = "Incoming TaskItem was null."
        GoTo ThrowException
    End If

   ' Prepare to present Outlook Task UI
    Set singleInspector = t.GetInspector
    
    ' Show the Task using the Outlook interface
    t.Display
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Open Task's associated Project
Public Sub OpenProject(Optional ByVal t As Outlook.TaskItem = Nothing)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":DeleteTask"
    
    If t Is Nothing Then
        Set t = selItem
    Else
        Set selItem = t
    End If
    
    If IsNothing(t) Then
        strTrace = "Incoming TaskItem was null."
        GoTo ThrowException
    End If
    
    Dim myProject As fmeProject
    pm.GetProject t, myProject
    If Not myProject Is Nothing Then
        pm.OpenProject myProject
    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

''' Creates a new Task to Edit
Public Sub NewTask(Optional p As fmeProject = Nothing)

    ' Create a new Task
    Dim t As Outlook.TaskItem
    Set t = ut.CreateTask
    
    ' If a project provided, assign task to the Project
    If Not p Is Nothing Then t.Categories = p.Subject
    
    ' Prepare to present Outlook Task UI
    Set selItem = t
    Set singleInspector = t.GetInspector
    
    ' Show the Task using the Outlook interface
    t.Display

End Sub

Public Sub DeleteTask(Optional ByVal t As Outlook.TaskItem = Nothing)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":DeleteTask"
    
    If t Is Nothing Then
        Set t = selItem
    Else
        Set selItem = t
    End If
    
    If IsNothing(t) Then
        strTrace = "Incoming TaskItem was null."
        GoTo ThrowException
    End If
    
    strTrace = "Are you sure you want to delete: " & vbCrLf & vbCrLf & _
                t.Subject
    Dim dResult As VbMsgBoxResult
    dResult = MsgBox(strTrace, vbQuestion Or vbYesNo, "Title")
    If dResult = vbYes Then
            
        strTrace = "Removed task: " & t.Subject & " from the Outlook datastore."
        LogMessage strTrace, strRoutine
    
        ' Delete the Task from Outlook
        t.Delete
    
        ' Update UI
        DeleteListViewItem t

    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Prints the task to the user's default printer
Public Sub PrintTask(ByVal t As Outlook.TaskItem)
    t.PrintOut
End Sub

Public Sub Delegate(ByVal t As Outlook.TaskItem)

End Sub

Public Sub DisplayRelatedMail(Optional ByVal t As Outlook.TaskItem = Nothing)

    If t Is Nothing Then
        Set t = selItem
    Else
        Set selItem = t
    End If

    If Not IsNothing(t) Then
    
        Dim frm As New frm_MailViewer
        frm.title = "Mail Viewer - Related Mail"
        frm.Header = t.Subject
        Dim refId As String
        refId = ut.GetReferenceID(t)
        frm.Show
        frm.Load t

    End If

End Sub

''' Creates an Email about the specified Task
Public Sub MessageTask(ByVal t As Outlook.TaskItem)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":MessageTask"

    Dim oMail As Outlook.MailItem
    Set oMail = ut.GetMessageForTask(t)
    If Not IsNothing(oMail) Then
        oMail.Display
    Else
        strTrace = "Failed to create a new message for task: " & t.Subject
        LogMessageEx strTrace, Nothing, strRoutine
    End If

End Sub

''' Marks the specified task as complete
Public Sub MarkComplete(ByRef t As Outlook.TaskItem)

    t.MarkComplete
    
    Set selItem = t
    
    ' Update collection & UI
    UpdateItem t
    RaiseEvent ItemUpdated
    
End Sub

''' Marks the specified task as incomplete
Public Sub MarkIncomplete(ByRef t As Outlook.TaskItem)

    t.Complete = False
    t.DueDate = #1/1/4501#
    t.Status = olTaskNotStarted
    t.Save
    
    Set selItem = t
    
    ' Update collection & UI
    UpdateItem t
    RaiseEvent ItemUpdated

End Sub

''' Changes the specified tasks's due date to Today
Public Sub MakeDueToday(ByRef t As Outlook.TaskItem)

    Dim dte As Date
    dte = Date
    t.DueDate = dte
    t.Save
    
    Set selItem = t
    
    ' Update collection and UI
    UpdateItem t
    RaiseEvent ItemUpdated
    
End Sub

''' Changes the specified task's due date to Tomorrow
Public Sub MakeDueTomorrow(ByRef t As Outlook.TaskItem)

    Dim dte As Date
    dte = Date
    t.DueDate = DateAdd("d", 1, dte)
    t.Save
    
    Set selItem = t
    
    ' Update collection and UI
    UpdateItem t
    RaiseEvent ItemUpdated

End Sub

''' Moves the DueDate forward (plus values) or backward (minus values) for the
''' specified task
Public Sub MoveDueDate(ByRef t As Outlook.TaskItem, ByVal dys As Integer)

    ' Add specified days
    Dim dte As Date
    Dim dteOld As Date
    dteOld = t.DueDate
    If IsDateNone(t.DueDate) Then dteOld = Date
    
    dte = DateAdd("d", dys, dteOld)
    
    ' Check to make sure a week day is selected
    Dim dayName As String
    dayName = LCase(Format(dte, "ddd"))
    If dayName = "sun" Then
        dte = DateAdd("d", 1, dte)
    End If
    If dayName = "sat" Then
        dte = DateAdd("d", 2, dte)
    End If
    
    ' Update the DueDate
    t.DueDate = dte
    t.Save
    
    Set selItem = t
    
    ' Update collection and UI
    UpdateItem t
    RaiseEvent ItemUpdated
    
End Sub

''' Moves the specified Task to the Master Task list
Public Sub MoveToMasterTasks(ByRef t As Outlook.TaskItem)
   
    t.startDate = #1/1/4501#
    t.DueDate = #1/1/4501#
    t.Save

    Set selItem = t
    
    ' Update collection and UI
    UpdateItem t
    RaiseEvent ItemUpdated
    
End Sub

''' Sets the Task Reference information into the mail item's body
Public Sub LinkMailToTask(ByVal m As Outlook.MailItem, ByVal t As Outlook.TaskItem)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":LinkMailToTask"
    
    On Error GoTo ThrowException
        
    strTrace = "Get Reference Id from the Task."
    Dim refId As String
    refId = ut.GetReferenceID(t)
    If Len(refId) = 0 Then
        ' Need to set a reference Id
        ut.SetReferenceID t
        refId = ut.GetReferenceID(t)
    End If
    
    strTrace = "Update MailItem's body with FSRefId message."
    ut.AppendMarketingSignature m, refId
    
    ' Not appropriate for an embedded object before the parent item (e.g. Task) has been saved.
    ' m.Save
    
    ' Organized and powered by FocusMe for Outlook, Ceptara Corp © 2019
    ' Bringing excellence to managing chaos (email) and staying goal focused!
    ' [FSRefID: GV1V72]
   
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

''' Returns the text used to tag an email to a task
Public Function GetTaskTag(ByVal t As Outlook.TaskItem) As String

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":LinkMailToTask"
    
    On Error GoTo ThrowException

    strTrace = "Get Reference Id from the Task."
    Dim refId As String
    refId = ut.GetReferenceID(t)
    If Len(refId) = 0 Then
        ' Need to set a reference Id
        ut.SetReferenceID t
        refId = ut.GetReferenceID(t)
    End If
    
    strTrace = ut.GetReferenceMessage(refId)
    GetTaskTag = strTrace
    
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    GetTaskTag = "No Tag Found, see error log."

End Function

''' Displays the TaskItem's properties
Public Sub TaskProperties(ByVal t As Outlook.TaskItem)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TaskProperties"
    
    On Error GoTo ThrowException
    
    Dim frm As New frm_ItemProperties
    
    frm.title = "Outlook TaskItem"
    frm.Description = t.Subject
    Set frm.Item = t
    frm.Show
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Processes a Text stream dropped onto a Task
Public Sub ProcessDroppedText(ByVal strMsg As String, ByVal t As Outlook.TaskItem)

' Expected dropped text
   ' TEXT
    ' (1)   Outlook.MailItems - tab delimited with headers {Starts with From}
    ' (2)   Outlook.ContactItems - tab delimited with headers {starts with 'Full Name'}
    ' (3)   Outlook.AppointmentItems = appointment title {no distinuquishing factors}
    ' (4)   Outlook.TaskItems = tab delimited table with headers {Starts with 'Task Subject'}

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ProcessDroppedText"
    
    If Len(strMsg) = 0 Then
        strTrace = "Dropped text is empty."
        GoTo ThrowException
    End If
    
    ' Initialize Outlook Interface
    Dim ut As New Utilities
    
    ' Inititalize Importer
    Dim importer As csvImporter
    Set importer = New csvImporter
    
    Dim dt As DataTable
    
    ' Figure out what got dropped
    Dim txtStarter As String
    txtStarter = Left(strMsg, 12)
    
    Dim strHeader As String
    
    Dim typ As Integer
    If InStr(1, LCase(txtStarter), "from" & vbTab) Then
        typ = 1
        strHeader = "From"
    End If
    If InStr(1, LCase(txtStarter), "to" & vbTab) Then
        typ = 1
        strHeader = "To"
    End If
    If InStr(1, LCase(txtStarter), "full name") Then typ = 2
    If InStr(1, LCase(txtStarter), "task subject") Then typ = 4
    
    Select Case typ
        Case 1 ' One or more MailItems dropped
            Set dt = importer.Import(strMsg, 1, strHeader, vbTab)
            strTrace = "Mail: " & dt.Name & " found " & dt.rows.count & " items."
            ' Find Mail Item
            Dim Mail As ArrayList
            Set Mail = GatherDroppedMail(dt)
            If IsNothing(Mail) Then
                strTrace = "Outlook mail query failed."
                GoTo ThrowException
            End If
            If Mail.count > 0 Then
                Dim m As Outlook.MailItem
                For Each m In Mail
                    ' Link MailItem to TaskItem
                    LinkMailToTask m, t
                Next
            Else
                strTrace = "Failed to find mail matching the dropped " & _
                            dt.rows.count & " mailItems."
                GoTo ThrowException
            End If
            
        Case 2 ' One or more ContactItems dropped
            Set dt = importer.Import(strMsg, 1, "Full Name", vbTab)
            strTrace = "Contacts: " & dt.Name & " found " & dt.rows.count & " items."

        Case 4 ' One or more Tasks dropped
            Set dt = importer.Import(strMsg, 1, "Task Subject", vbTab)
            strTrace = "Tasks: " & dt.Name & " found " & dt.rows.count & " items."
            
        Case Else
            strTrace = "Unrecognized dropped text: " & strMsg
            GoTo ThrowException
    End Select
        
    LogMessage strTrace, strRoutine
        
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub
Private Function GatherDroppedMail(ByVal dt As DataTable) As ArrayList

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GatherDroppedMail"
    
    Dim retList As New ArrayList
    Dim tmpList As ArrayList
    
    Dim r As DataRow
    For Each r In dt.rows.Items
        Dim subj As String
        subj = r.GetItem("Subject")
        Set tmpList = FindMailBySubject(subj)
        If tmpList.count > 0 Then
            If tmpList.count > 10 Then
                strTrace = "WARNING: Seems like a lot of mail to mark, items: " & tmpList.count
                LogMessage strTrace, strRoutine
            End If
        Else
            strTrace = "Mail query on the 'subject' failed."
            GoTo ThrowException
        End If
        
        ' Report the results
        Dim m As Outlook.MailItem
        For Each m In tmpList
            retList.Add m
        Next
        
    Next
    
    Set GatherDroppedMail = retList
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GatherDroppedMail = Nothing

End Function
Private Function FindMailBySubject(ByVal subj As String) As ArrayList

' By Subject
' "urn:schemas:httpmail:subject" LIKE '%123%'
' By Conversation Topic
' "urn:schemas:httpmail:thread-topic" LIKE '%123%'

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":FindMailBySubject"
    
    If Len(subj) = 0 Then
        strTrace = "An empty Subject encountered."
        GoTo ThrowException
    End If

    Dim ut As New Utilities
    
    Dim retList As ArrayList

    ' trim the subject
    Dim snippet As String
    snippet = subj
    ' If Len(subj) > 10 Then snippet = Left(subj, 10)
    Dim strQuery As String
    strQuery = "@SQL=""" & PR_MAILITEM_SUBJECT & """ LIKE '%" & snippet & "%'"
    Set retList = ut.GetOutlookMail(strQuery)
    If IsNothing(retList) Then
        strTrace = "WARNING: Outlook interface failed - return empty list."
        LogMessage strTrace, strRoutine
        Set retList = New ArrayList
    End If
    
    Set FindMailBySubject = retList
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Function

' - - Managed UI

Public Sub Refresh()
        
    If Not myListView Is Nothing Then RefreshListView sortingColumn, sortingDirection
    
End Sub

' - - Context Menu

''' Returns a collection of ContextMenuItems
Public Function GetContextMenu(ByVal t As Outlook.TaskItem) As ContextMenu

' +-----------------
' | Open            (1001) <- Starts the Outlook Task Inspector
' | Open Project    (1019)
' | Print...        (1002)
' +-----------------
' | Delegate...     (1003)
' | Related Mail... (1004)
' | Forward...      (1016)
' | Schedule Time...(1024)
' +-----------------
' | Copy Tag...     (1017)
' +-----------------
' | Mark Incomplete (1005)
' | Mark Complete   (1006)
' | Due Today       (1007)
' | Tomorrow        (1008)
' | 3 Days          (1009)
' | Next Week       (1010)
' | Next Month      (1011)
' | Master Task     (1012)
' +-----------------
' | Delete          (1014)
' | Properties      (1015)
' +-----------------

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetContextMenu"
    
    Dim cm As ContextMenu
    Set cm = New ContextMenu

    Dim mni As ContextMenuItem
    
    ' Open Task
    Set mni = New ContextMenuItem
    mni.Caption = "Open..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1001
    cm.AddItem mni
    
    ' Open Project
    Set mni = New ContextMenuItem
    mni.Caption = "Open Project..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1019
    cm.AddItem mni
    
    ' Print Task
    Set mni = New ContextMenuItem
    mni.Caption = "Print..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1002
    cm.AddItem mni
    
    ' Separator
    Set mni = New ContextMenuItem
    mni.Caption = "-"
    mni.Enabled = True
    mni.itemType = Separator
    mni.UID = 10022
    cm.AddItem mni
        
    ' Delegate
    Set mni = New ContextMenuItem
    mni.Caption = "Delegate..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1003
    cm.AddItem mni

    ' Related Mail
    Set mni = New ContextMenuItem
    mni.Caption = "Related Mail..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1004
    cm.AddItem mni
    
    ' Message Task
    Set mni = New ContextMenuItem
    mni.Caption = "Forward..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1016
    cm.AddItem mni
    
    ' Schedule Time - Appointment
    Set mni = New ContextMenuItem
    mni.Caption = "Schedule Time..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1024
    cm.AddItem mni
    
        ' Separator
    Set mni = New ContextMenuItem
    mni.Caption = "-"
    mni.Enabled = True
    mni.itemType = Separator
    mni.UID = 10057
    cm.AddItem mni
    
    ' Copy Tag
    Set mni = New ContextMenuItem
    mni.Caption = "Copy Tag..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1017
    cm.AddItem mni
    
    ' Separator
    Set mni = New ContextMenuItem
    mni.Caption = "-"
    mni.Enabled = True
    mni.itemType = Separator
    mni.UID = 10055
    cm.AddItem mni
    
    If t.Complete Then
        ' Mark Incomplete
        Set mni = New ContextMenuItem
        mni.Caption = "Mark Incomplete"
        mni.Enabled = True
        mni.itemType = Button
        mni.UID = 1005
        cm.AddItem mni
    Else
        ' Mark Complete
        Set mni = New ContextMenuItem
        mni.Caption = "Mark Complete"
        mni.Enabled = True
        mni.itemType = Button
        mni.UID = 1006
        cm.AddItem mni
    End If
    
    ' Due Today
    Set mni = New ContextMenuItem
    mni.Caption = "Due Today"
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1007
    cm.AddItem mni
    
    ' Tomorrow
    Set mni = New ContextMenuItem
    mni.Caption = "Tomorrow"
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1008
    cm.AddItem mni
    
    ' 3 Days
    Set mni = New ContextMenuItem
    mni.Caption = "Move 3 Days"
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1009
    cm.AddItem mni
    
    ' Next Week
    Set mni = New ContextMenuItem
    mni.Caption = "Next Week"
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1010
    cm.AddItem mni
    
    ' Next Month
    Set mni = New ContextMenuItem
    mni.Caption = "Next Month"
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1011
    cm.AddItem mni
    
    ' Master Task
    Set mni = New ContextMenuItem
    mni.Caption = "Master Task"
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1012
    cm.AddItem mni
    
    ' Separator
    Set mni = New ContextMenuItem
    mni.Caption = "-"
    mni.Enabled = True
    mni.itemType = Separator
    mni.UID = 1013
    cm.AddItem mni
    
    ' Delete
    Set mni = New ContextMenuItem
    mni.Caption = "Delete"
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1014
    cm.AddItem mni
    
    ' Properties
    Set mni = New ContextMenuItem
    mni.Caption = "Properties..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1015
    cm.AddItem mni
    
    Set GetContextMenu = cm

End Function

''' Returns a collection of ContextMenuItems
Public Function GetContextMenuTwoTier(ByVal t As Outlook.TaskItem) As ContextMenu

' +-----------------
' | Open            (1001) <- Starts the Outlook Task Inspector
' | Open Project  >>(1019)  +-----------------
'                           | Open Project      (1020)
'                           | Windows Folder    (1021)
'                           | Outlook Folder    (1022)
'                           | OneNote Location  (1023)
'                           +------------------
' | Print...        (1002)
' +-----------------
' | Delegate...     (1003)
' | Related Mail... (1004)
' | Forward...      (1016)
' | Schedule Time...(1024)
' +-----------------
' | Copy Tag...     (1017)
' +-----------------
' | Start Task...   (1025)
' | Mark Incomplete (1005)
' | Mark Complete   (1006)
' | Due Today       (1007)
' | Reschedule    >>(1018)  +-----------------
'                           | Tomorrow        (1008)
'                           | 3 Days          (1009)
'                           | Next Week       (1010)
'                           | Next Month      (1011)
'                           | Master Task     (1012)
'                           +-----------------
' +-----------------
' | Delete          (1014)
' | Properties      (1015)
' +-----------------

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetContextMenuTwoTier"
    
    On Error GoTo ThrowException
    
    Dim cm As ContextMenu
    Set cm = New ContextMenu

    Dim mni As ContextMenuItem
    
    Dim cmSub As ContextMenu
    
    ' Open Task
    Set mni = New ContextMenuItem
    mni.Caption = "Open..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1001
    cm.AddItem mni

    ' Reload Project List from DB
    pm.Reload
    
    ' Check for an associated Project
    Dim myProject As fmeProject
    pm.GetProject t, myProject
    
    If Not myProject Is Nothing Then
    
        ' - - Open Project SubMenu
        Set cmSub = New ContextMenu
        
        ' Open Associated Project
        Set mni = New ContextMenuItem
        mni.Caption = "Open Project"
        mni.Enabled = True
        mni.itemType = Button
        mni.UID = 1020
        cmSub.AddItem mni
    
        If Len(myProject.WindowsFolder) > 0 Then
            Set mni = New ContextMenuItem
            mni.Caption = "Windows Folder"
            mni.Enabled = True
            mni.itemType = Button
            mni.UID = 1021
            cmSub.AddItem mni
        End If
    
        If Len(myProject.OutlookFolder) > 0 Then
            Set mni = New ContextMenuItem
            mni.Caption = "Outlook Folder"
            mni.Enabled = True
            mni.itemType = Button
            mni.UID = 1022
            cmSub.AddItem mni
        End If
        
        If cmSub.count > 1 Then
            ' - - Open Project >>
            Set mni = New ContextMenuItem
            mni.Caption = "Open Project"
            mni.Enabled = True
            mni.itemType = Menu
            mni.UID = 10190
            Set mni.SubMenu = cmSub
            cm.AddItem mni
        Else
        ' - - Open Project...
            Set mni = New ContextMenuItem
            mni.Caption = "Open Project..."
            mni.Enabled = True
            mni.itemType = Menu
            mni.UID = 1020
            ' Set mni.SubMenu = cmSub
            cm.AddItem mni
        End If
    End If
    
    ' Print Task
    Set mni = New ContextMenuItem
    mni.Caption = "Print..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1002
    cm.AddItem mni
    
    ' Separator
    Set mni = New ContextMenuItem
    mni.Caption = "-"
    mni.Enabled = True
    mni.itemType = Separator
    mni.UID = 10022
    cm.AddItem mni
        
    ' Delegate
    Set mni = New ContextMenuItem
    mni.Caption = "Delegate..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1003
    cm.AddItem mni

    ' Related Mail
    Set mni = New ContextMenuItem
    mni.Caption = "Related Mail..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1004
    cm.AddItem mni
    
    ' Message Task
    Set mni = New ContextMenuItem
    mni.Caption = "Forward..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1016
    cm.AddItem mni
    
    ' Schedule Time - Appointment
    Set mni = New ContextMenuItem
    mni.Caption = "Schedule Time..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1024
    cm.AddItem mni
    
        ' Separator
    Set mni = New ContextMenuItem
    mni.Caption = "-"
    mni.Enabled = True
    mni.itemType = Separator
    mni.UID = 10057
    cm.AddItem mni
    
    ' Copy Tag
    Set mni = New ContextMenuItem
    mni.Caption = "Copy Tag..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1017
    cm.AddItem mni
    
    ' Separator
    Set mni = New ContextMenuItem
    mni.Caption = "-"
    mni.Enabled = True
    mni.itemType = Separator
    mni.UID = 10055
    cm.AddItem mni
    
    If t.Complete Then
        ' Mark Incomplete
        Set mni = New ContextMenuItem
        mni.Caption = "Mark Incomplete"
        mni.Enabled = True
        mni.itemType = Button
        mni.UID = 1005
        cm.AddItem mni
    Else
        ' Mark Complete
        Set mni = New ContextMenuItem
        mni.Caption = "Mark Complete"
        mni.Enabled = True
        mni.itemType = Button
        mni.UID = 1006
        cm.AddItem mni
    End If
    
    ' Due Today
    Set mni = New ContextMenuItem
    mni.Caption = "Due Today"
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1007
    cm.AddItem mni
    
    ' - - Reschedule SubMenu
    Set cmSub = New ContextMenu
    
    ' Tomorrow
    Set mni = New ContextMenuItem
    mni.Caption = "Tomorrow"
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1008
    cmSub.AddItem mni
    
    ' 3 Days
    Set mni = New ContextMenuItem
    mni.Caption = "Move 3 Days"
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1009
    cmSub.AddItem mni
    
    ' Next Week
    Set mni = New ContextMenuItem
    mni.Caption = "Next Week"
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1010
    cmSub.AddItem mni
    
    ' Next Month
    Set mni = New ContextMenuItem
    mni.Caption = "Next Month"
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1011
    cmSub.AddItem mni
    
    ' Master Task
    Set mni = New ContextMenuItem
    mni.Caption = "Master Task"
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1012
    cmSub.AddItem mni
    
    ' Reschedule
    Set mni = New ContextMenuItem
    mni.Caption = "Reschedule"
    mni.Enabled = True
    mni.itemType = Menu
    mni.UID = 1018
    Set mni.SubMenu = cmSub
    cm.AddItem mni
    
    ' Separator
    Set mni = New ContextMenuItem
    mni.Caption = "-"
    mni.Enabled = True
    mni.itemType = Separator
    mni.UID = 1013
    cm.AddItem mni
    
    ' Delete
    Set mni = New ContextMenuItem
    mni.Caption = "Delete"
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1014
    cm.AddItem mni
    
    ' Properties
    Set mni = New ContextMenuItem
    mni.Caption = "Properties..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1015
    cm.AddItem mni
    
    Set GetContextMenuTwoTier = cm
    
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetContextMenuTwoTier = Nothing

End Function

Public Sub ProcessContextMenuSelection(ByVal id As String, ByRef t As Outlook.TaskItem)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":ProcessContextMenuSelection"

    Dim iDiff As Integer

    Dim num As Integer
    num = CInt(id)
    Select Case num
        Case 1001   ' Open
            OpenTask t
        Case 1002   ' Print
            PrintTask t
        Case 1003   ' Delegate
        Case 1004   ' Get related mail
            DisplayRelatedMail t
        Case 1005   ' Mark Incomplete
            MarkIncomplete t
        Case 1006   ' Mark Complete
            MarkComplete t
        Case 1007   ' Due Today
            MakeDueToday t
        Case 1008   ' Tomorrow
            MakeDueTomorrow t
        Case 1009   ' 3 Days
            iDiff = DateDiff("d", t.DueDate, DateAdd("d", 3, Date))
            MoveDueDate t, iDiff
        Case 1010   ' Next week
            iDiff = DateDiff("d", t.DueDate, DateAdd("d", 7, Date))
            MoveDueDate t, iDiff
        Case 1011   ' Next month
            iDiff = DateDiff("d", t.DueDate, DateAdd("d", 30, Date))
            MoveDueDate t, iDiff
        Case 1012   ' Master Task
            MoveToMasterTasks t
        Case 1014   ' Delete
        Case 1015   ' Properties
            TaskProperties t
        Case 1016   ' Message Task
            MessageTask t
        Case 1017   ' Copy Task Tag
            strTrace = GetTaskTag(t)
            CopyToClipboard strTrace
        Case 1018   ' Present Reschedule SubMenu
            ' Handled in the ContextMenu form
            ' Ignore it here
        Case 1019 ' Present Open Project SubMenu
              ' Handled in the ContextMenu form
            ' Ignore it here
        Case 1020 ' Open associated Project Properties
            OpenProject t
        Case 1021 ' Open associated Project's Windows Folder
        
        Case 1022 ' Open associated Project's Outlook Folder
        
        Case 1024 ' Schedule Time
        
        Case Else
            strTrace = "Unrecognised menu code: " & id & " - ignoring request."
            LogMessage strTrace, strRoutine
    End Select

End Sub

' - Supporting Methods

''' Updates the specified Item in the collection
Private Sub UpdateItem(ByVal t As Outlook.TaskItem, Optional idx As Integer = -1)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":UpdateItem"

    If idx < 0 Then
        ' Look for the item
        idx = IsItemInCollection(t)
        If idx < 0 Then
            strTrace = "Cannot update an item that is not in the collection, failed to find: " & t.Subject
            GoTo ThrowException
        End If
    End If

    ' Replace item at the specified idx
    myItems.RemoveAt idx
    myItems.AddItem t
    
    ' Update Managed UI
    UpdateListViewItem t

    Exit Sub

ThrowException:
    LogMessage strTrace, strRoutine

End Sub

''' Loops thru the internal collection of Tasks and locates the
''' specified item returning its index
''' Returns -1 if not found
Private Function IsItemInCollection(ByVal t As Outlook.TaskItem) As Integer

    Dim iReturn As Integer
    iReturn = -1

    Dim bFnd As Boolean
    bFnd = False
    For i = 0 To myItems.Items.count - 1
        Dim tmp As Outlook.TaskItem
        Set tmp = myItems.Items(i)
        If tmp.EntryId = t.EntryId Then
            bFnd = True
            iReturn = i
            Exit For
        End If
    Next
    
    IsItemInCollection = iReturn

End Function

''' <summary>
''' Gets a list of Outlook.TaskItems for a given folder set and query.
''' </summary>
''' <param name="DASLQuery">String: Optional DASL query - if empty returns all items.</param>
''' <param name="AllFolders">Boolean (Optional, defaults to False): if True queries
''' all folders with a default ItemType = Outlook.TaskItem
''' <returns>List(Of Outlook.TaskItem)</returns>
''' <remarks>If an error occurs returns Nothing.</remarks>
Private Function GetOutlookTasks(Optional ByVal DASLQuery As String = "", _
                                 Optional ByVal AllFolders As Boolean = False) As ArrayList

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":GetOutlookTasks"
    
    On Error GoTo ThrowException
                                                
    Dim retItems As ArrayList
    Dim tmpItems As ArrayList
    
    Dim fldr As Outlook.Folder
    Dim fldrList As New ArrayList
    
    If AllFolders Then
        strTrace = "Retrieve items from all Task folders."
        Set fldrList = ut.GetFoldersByType(olTaskItem)
    Else
        strTrace = "Retrieve items from the default folder."
        Set fldr = oNS.GetDefaultFolder(olFolderTasks)
        fldrList.Add fldr
    End If
    
    ' Get List from Outlook
    Set tmpItems = GetOutlookTasksFromFolders(DASLQuery, fldrList)
    
    ' When using SharePoint lists, its possible to get duplicate EntryIDs
    Set retItems = DeDuplicateTaskList(tmpItems)
    
    Set GetOutlookTasks = retItems
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetOutlookTasks = Nothing

End Function

''' <summary>
''' Gets a list of Outlook.TaskItems for a given folder set and query.
''' </summary>
''' <param name="DASLQuery">String: Optional DASL query - if empty returns all items.</param>
''' <param name="Folders">List(Of Outlook.MapyFolder) - if nothing returns default task folder.</param>
''' <returns>List(Of Outlook.TaskItem)</returns>
''' <remarks>If an error occurs returns Nothing.</remarks>
Private Function GetOutlookTasksFromFolders(ByVal DASLQuery As String, _
                                            ByVal Folders As ArrayList) As ArrayList
                                     
    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":GetOutlookTasksFromFolders"
    
    On Error GoTo ThrowException
    
    If Folders Is Nothing Then
        strTrace = "A source list of task folders is required."
        GoTo ThrowException
    End If
    If Folders.count = 0 Then
        strTrace = "Source folder list was empty - no folders to process."
        GoTo ThrowException
    End If
                                                
    Dim retItems As ArrayList
    Set retItems = New ArrayList
                                                
    Dim fldr As Outlook.Folder
    Dim oItems As Outlook.Items
    
    Dim r As Outlook.Recipient
    Dim oMe As Outlook.ContactItem
    Set oMe = ut.WhoAmI
    
    Dim bJustMe As BookmarkEnum
    bJustMe = True
    Dim bInclude As Boolean
    Dim recEmailAddr As String
    
    For Each fldr In Folders
    
        If Len(DASLQuery) = 0 Then
            Set oItems = fldr.Items
        Else
            Set oItems = fldr.Items.Restrict(DASLQuery)
        End If
    
        strTrace = "Add the Outlook Items collection to the return list."
        For Each o In oItems
            If TypeOf o Is Outlook.TaskItem Then
                ' Capture TaskItem
                Dim t As Outlook.TaskItem
                Set t = o
                
                ' Load questions
                If t.Recipients.count > 0 Then
                    strTrace = "Found a shared task."
                    If Not oMe Is Nothing Then
                    
                        ' Catch only 'my tasks'?
                        If bJustMe Then
                            Set r = t.Recipients.Item(1)
                            recEmailAddr = r.AddressEntry.Address
                            
                            bInclude = False
                            If LCase(recEmailAddr) = LCase(oMe.Email1Address) Or _
                                LCase(recEmailAddr) = LCase(oMe.Email2Address) Then
                                ' Found me
                                bInclude = True
                            End If
                            
                            If bInclude Then retItems.Add o
                                                   
                        Else
                            ' Capture all assigned tasks.
                            retItems.Add o
                        End If
                    
                    End If
                Else
                    strTrace = "Unassigned task - capture it."
                    retItems.Add o
                End If
                
            End If
        Next
    
    Next
                                                
    Set GetOutlookTasksFromFolders = retItems
    Exit Function
        
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetOutlookTasksFromFolders = Nothing
        
End Function

''' Searches thru the list and removes any duplicate EntryIDs
Private Function DeDuplicateTaskList(ByVal arList As ArrayList) As ArrayList

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":DeDuplicateTaskList"
    
    On Error GoTo ThrowException
    
    If arList Is Nothing Then
        strTrace = "A null list was encountered."
        GoTo ThrowException
    End If
    
    Dim retItems As ArrayList
    Set retItems = New ArrayList
    
    Dim t As Outlook.TaskItem
    Dim tmpT As Outlook.TaskItem
    Dim bFnd As Boolean
    For Each t In arList
        bFnd = False
        For Each tmpT In retItems
            If tmpT.EntryId = t.EntryId Then
                bFnd = True
                Exit For
            End If
        Next
        
        If Not bFnd Then retItems.Add t
        
    Next

    Set DeDuplicateTaskList = retItems
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set DeDuplicateTaskList = Nothing

End Function
                                                
' - - ListView Supporting Methods

Private Sub AddListViewItem(ByVal t As Outlook.TaskItem, Optional ByVal idx As Integer = -1)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":AddListViewItem"
    
    If myListView Is Nothing Then
        Exit Sub
    End If

    ' Check the index
    If idx < 0 Then idx = myListView.ListItems.count + 1

    ' Add Item to ListView
    Dim ky As String
    ' ky = "k" & Common.GenerateUniqueID(3) & "_" & t.EntryId
    ky = t.EntryId
    Dim li As ListItem
    Set li = myListView.ListItems.Add(idx, ky, t.Subject)
    If IsDateNone(t.DueDate) Then
        li.SubItems(1) = "None"
    Else
        li.SubItems(1) = Format(t.DueDate, "mm/dd/yyyy")
    End If
           
    ' Format the row
    FormatLVRow li, t
    
    strTrace = "Added task to ListView: " & t.Subject & ", key: " & li.key
    LogMessage strTrace, strRoutine

End Sub
Private Sub UpdateListViewItem(ByVal t As Outlook.TaskItem)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":UpdateListViewItem"
    
    If myListView Is Nothing Then
        Exit Sub
    End If
    
    On Error GoTo ThrowException
    
    Dim li As ListItem
    Set li = FindListViewItem(t)
    If Not IsNothing(li) Then
        ' Update list view here
        li.text = t.Subject
            
        If IsDateNone(t.DueDate) Then
            li.SubItems(1) = "None"
        Else
            li.SubItems(1) = Format(t.DueDate, "mm/dd/yyyy")
        End If
           
        ' Format the row
        FormatLVRow li, t
        
        strTrace = "Updated ListView for task: " & t.Subject
    Else
        strTrace = "Add new task to the ListView."
        AddListViewItem t
    End If
    
    LogMessage strTrace, strRoutine
    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub
Private Sub DeleteListViewItem(ByVal t As Outlook.TaskItem)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":DeleteListViewItem"
    
    If myListView Is Nothing Then
        Exit Sub
    End If
    
    On Error GoTo ThrowException
    
    If IsNothing(t) Then
        strTrace = "A null taskItem encountered."
        GoTo ThrowException
    End If

    Dim li As ListItem
    Set li = FindListViewItem(t)
    If Not IsNothing(li) Then
        myListView.ListItems.Remove li.Index
    Else
        strTrace = "WARNING: unable to find task: " & t.Subject & " in the listview."
        GoTo ThrowException
    End If
    
    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub
Private Function FindListViewItem(ByVal t As Outlook.TaskItem) As ListItem

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":FindListViewItem"

    If myListView Is Nothing Then
        Exit Function
    End If
    
    Dim bFnd As Boolean
    bFnd = False
    
    Dim retItem As ListItem
    Set retItem = Nothing
    
    For i = 1 To myListView.ListItems.count
        Dim li As ListItem
        Set li = myListView.ListItems(i)
        If li.key = t.EntryId Then
            Set retItem = li
            Exit For
        End If
    Next
    
    Set FindListViewItem = retItem
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Function

Private Sub RefreshListView(ByVal sortOn As enuSortOn, ByVal sortDir As enuSortDirection)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":RefreshListView"
    
    If myListView Is Nothing Then
        Exit Sub
    End If
    
    On Error GoTo ThrowException
    
    ' Reset these...
    myListView.ColumnHeaders.Clear
    
    ' Setup ListView Columns and Configuration
    If myListView.ColumnHeaders.count = 0 Then
        'Initialize the View
        Dim ch1 As ColumnHeader
        Dim ch2 As ColumnHeader
        With myListView
            Set ch1 = .ColumnHeaders.Add(1, "Subject", "Description")
            Set ch2 = .ColumnHeaders.Add(2, "DueDate", "Date Due")
            
                ' Reset these (in case the ListView control
                '   was modified by another Controller/Manager
            .Checkboxes = True
            .View = lvwReport
        End With
    End If
        
    ' Clear current LV collection
    myListView.ListItems.Clear
    
    ' Add current class collection
    If myItems.count > 0 Then
    
        ' Sort the internal list
        Dim strSort As String
        strSort = "Subject"
        If sortingColumn = DueDate Then strSort = "DueDate"
        
        Dim collSort As New SortCollection
        collSort.Sort strSort, myItems.Items, sortingDirection
          
        ' Load the ListView
        For i = 0 To myItems.count - 1
            Dim t As Outlook.TaskItem
            Set t = myItems.Item(i)
                                 
            ' Add Item to ListView
            AddListViewItem t, i + 1
            
        Next
        
        Call ResizeLVColumns
        
    End If
    
    strTrace = "Items: " & myItems.count & " tasks..."
    
ThrowException:
    LogMessage strTrace, strRoutine

End Sub

Private Sub FormatLVRow(ByVal li As ListItem, ByVal t As Outlook.TaskItem)

    ' Format the checkbox
    li.checked = t.Complete
                   
    ' Color the Task
    Dim today As Date
    today = Date
    If t.DueDate < Date Then
        li.ForeColor = &HFF& ' Red
    Else
        li.ForeColor = &H80000007 ' Black
    End If
    If t.DueDate = today Then li.ForeColor = &HFF0000 ' Blue
    If t.Importance = olImportanceHigh Then
        li.ForeColor = &H80& ' Magenta
    End If
    
    ' Show as completed if appropriate
    If t.Complete Then
        li.ForeColor = &HC0C0C0 ' Light Gray
    End If

End Sub

Private Sub ResizeLVColumns()

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":ResizeLVColumns"
    
    On Error GoTo ThrowException
    
    strColumnWidths = "75;25"
    
    Dim totWidth As Integer
    totWidth = myListView.Width
    
    ' if scrollbar present, make space
    Dim bScrollbar As Boolean
    With myListView
        bScrollbar = (.font.SIZE + 4 + 1) * .ListItems.count > .Height
    End With
    
    If bScrollbar Then totWidth = totWidth - 15
    
    Dim widths() As String
    widths = Split(strColumnWidths, ";")
    
    For i = LBound(widths) To UBound(widths)
        Dim colWidth As Integer
        colWidth = CInt((widths(i) / 100) * totWidth) - 1
        myListView.ColumnHeaders(i + 1).Width = colWidth
    Next
    
    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Private Function LV_GetItemAt(ByVal X As stdole.OLE_XPOS_PIXELS, _
                              ByVal y As stdole.OLE_YPOS_PIXELS, _
                     Optional ByVal factor As Integer = 15) As ListItem


    ' Convert Pixels to TWIPS
    ' - .net uses Pixels, VBA uses TWIPS for ListView and TreeView (OLE_PIXELS?)
    ' - "on most computers 1 pixel = 15 TWIPS"
    '    https://stackoverflow.com/questions/36442535/vba-drag-drop-from-treeview-to-listview-listview-to-treeview-activex-controls
    Dim xInt As Single
    xInt = X
    Dim yInt As Single
    yInt = y
           
    Dim li As ListItem
    Set li = myListView.HitTest(X * factor, y * factor)
        
    Set LV_GetItemAt = li
        
End Function


