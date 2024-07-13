VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjectManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' - Fields

Private Const rootClass As String = "ProjectManager"

Private ut As Utilities
Private oNS As Outlook.NameSpace

' - -  Settings
Private stgs As Settings

Private WithEvents myListView As MSComctlLib.ListView
Attribute myListView.VB_VarHelpID = -1

Dim WithEvents frmProject As frm_ProjectProperties
Attribute frmProject.VB_VarHelpID = -1

Dim sortingColumn As enuSortOn
Dim sortingDirection As enuSortDirection

' - Events

''' Thrown when a Project is changed by this class
Public Event ItemUpdated()
Public Event ItemAdded()
Public Event collectionUpdated()

' - Properties
Dim myItems As Projects
Dim selItem As fmeProject
Dim selItems As Projects

Dim lvCheckbox As Boolean
Dim lvMultiSelect As Boolean
Dim f_SuspendUI As Boolean

''' Latest collection of Projects
Public Property Set Items(ByVal coll As Projects)
    Set myItems = coll
    RaiseEvent collectionUpdated
End Property
Public Property Get Items() As Projects
    Set Items = myItems
End Property

''' Currently selected Item
Public Property Get SelectedItem() As fmeProject
    Set SelectedItem = selItem
End Property

''' Currently selected Collection of Projects
Public Property Get SelectedItems() As Projects
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

' - - ListView Event Handlers

Private Sub myListView_AfterLabelEdit(Cancel As Integer, NewString As String)

    If f_SuspendUI Then Exit Sub

    ' Update the Project Name
End Sub

Private Sub myListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If f_SuspendUI Then Exit Sub

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":myListView_ColumnClick"
    
    Dim selectSortOn As enuSortOn
    If InStr(1, LCase(ColumnHeader.key), "name") Then selectSortOn = Name
    If InStr(1, LCase(ColumnHeader.key), "code") Then selectSortOn = Code
    
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

    If f_SuspendUI Then Exit Sub

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":myListView_DblClick"
    
    On Error GoTo ThrowException
    
    Dim li As ListItem
    Set li = myListView.SelectedItem

    Dim eid As String
    eid = DecodeProjectID(li.key)
    Dim p As fmeProject
    Set p = GetItemById(eid)
    If IsNothing(p) Then
        strTrace = "Failed select the project, id: " & eid
        GoTo ThrowException
    End If
    
    Set selItem = p
    
    OpenProject p
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Allows the ability to track multi-selection
Private Sub myListView_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    If f_SuspendUI Then Exit Sub

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":myListView_ItemCheck"
    
    On Error GoTo ThrowException

    Dim eid As String
    eid = DecodeProjectID(Item.key)
    Dim p As fmeProject
    Set p = GetItemById(eid)
    If IsNothing(p) Then
        strTrace = "Failed check/uncheck the project, id: " & eid
        GoTo ThrowException
    End If

    If Item.checked Then
        ' Add Project to selected items collection
        TrackSelectedItems p, "add"
    Else
        ' Remove Project from selected items collection
        TrackSelectedItems p, "remove"
    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Sub myListView_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If f_SuspendUI Then Exit Sub

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":myListView_ItemClick"
    
    On Error GoTo ThrowException

    Dim eid As String
    eid = DecodeProjectID(Item.key)
    Dim p As fmeProject
    Set p = GetItemById(eid)
    If IsNothing(p) Then
        strTrace = "Failed select the project, id: " & eid
        GoTo ThrowException
    End If
    
    Set selItem = p
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Private Sub myListView_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)

    If f_SuspendUI Then Exit Sub

    Dim strTrace As String
    strTrace = ""
    Dim strRoutine As String
    strRoutine = rootClass & ":myListView_MouseDown"
    
    ' Catch a context menu call
    
    If Button = 2 Then
        ' Render the context menu
        
        ' Get the pointed to listview item
        Dim li As ListItem
        Set li = LV_GetItemAt(X, Y)
        If Not IsNothing(li) Then
        
            ' Select Item if not already selected
            Dim sLi As ListItem
            Set sLi = myListView.SelectedItem
        
           ' Get selected Project
            Dim p As fmeProject
            Set p = GetItemById(DecodeProjectID(li.key))
            If IsNothing(p) Then
                strTrace = "WARNING: Failed to find a selected project, id: " & li.key
                GoTo ThrowException
            End If
            strTrace = "Right clicked on project: " & p.Subject

            ' Create context menu
            strTrace = "Create the context menu."
            Dim cMenu As ContextMenu
            Dim cMenu2 As ContextMenu
            
            strTrace = "Prepare context menu location."
            Dim xSgl As Single
            xSgl = X
            Dim ySgl As Single
            ySgl = Y
            
            strTrace = "Present context menu."
            Dim bUseWindowsMenus As Boolean
            bUseWindowsMenus = stgs.UseContextMenuWindows
            
            Dim id As Long
                        
            If bUseWindowsMenus Then
                Set cMenu = GetContextMenu(p)
                id = ShowPopup(myListView, cMenu, xSgl, ySgl)
            Else
                Set cMenu2 = GetContextMenu(p)
                id = ShowMenu(cMenu2)
            End If
       
            strTrace = "Selected menu item: " & id
            LogMessage strTrace, strRoutine
            
            Call ProcessContextMenuSelection(id, p)
        
        End If
        
    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Private Sub myListView_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":myListView_OLEDragDrop"
    
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

End Sub

Private Sub myListView_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Effect = 1
    Set myListView.DropHighlight = myListView.HitTest(X * 15, Y * 20)
End Sub

' Project Properties form

Private Sub frmProject_Closing(ByVal sender As Object, ByVal saved As Boolean)

    Dim fp As frm_ProjectProperties
    Set fp = sender
    
    If saved Then
        ' Update UI
        UpdateListViewUI fp.Project
    End If
    
    ' Unload the from memory
    Unload fp

End Sub

' - Constructor

Private Sub Class_Initialize()

    ' Initialize Variables
    Set myItems = New Projects
    Set oNS = GetNamespace("MAPI")
    Set ut = New Utilities
    Set stgs = New Settings
    
    Set selItems = New Projects
    
    ' Default Values
    lvCheckbox = False
    lvMultiSelect = False
    f_SuspendUI = False
    
    sortingColumn = Name
    sortingDirection = Ascending
       
End Sub

Private Sub Class_Terminate()

    ' Clean Up Memory
    Set myItems = Nothing
    Set selItems = Nothing
    Set oNS = Nothing
    Set ut = Nothing
    
End Sub

' - Methods

''' Load all Projects from the DataStore
Public Sub Load()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Load"
    
    On Error GoTo ThrowException
    
    ' Retrieve Projects from DataStore
    Set myItems.Items = GetProjectsFromDB
    
    ' Inform Listeners
    RaiseEvent collectionUpdated
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

''' Refreshes the internal Projects collection from the DB
Public Sub Reload()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Load"
    
    On Error GoTo ThrowException
    
    ' Retrieve Projects from DataStore
    Set myItems.Items = GetProjectsFromDB
       
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Reviews the Outlook Master Categories List
''' Returns a collection of imported projects from the MCL
Public Function ImportFromMCL() As Projects

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ImportFromMCL"
    
    On Error GoTo ThrowException
    
    Dim newList As New Projects
    
    Dim olCategory As Outlook.Category
    Dim olCategories As Outlook.Categories
    Dim p As fmeProject
    
    Set olCategories = oNS.Categories
    If olCategories.count > 0 Then
    
        strTrace = "Importing " & olCategories.count & " Outlook categories."
        LogMessage strTrace, strRoutine
    
        For Each olCategory In olCategories
            ' Create a new Project
            Set p = New fmeProject
            ' Evaluate the Category
            Dim strCatName As String
            strCatName = olCategory.Name
            Dim hook As Integer
            hook = InStr(1, strCatName, " - ")
            If hook > 0 Then
                ' Code included in the category name
                p.Code = Left(strCatName, hook - 1)
                p.Name = Right(strCatName, Len(strCatName) - (hook + 2))
                
                p.CombineTitleCode = True
            Else
                ' Just a Name, ignore?
                p.Name = strCatName
                p.CombineTitleCode = False
            End If
            
            ' Capture MCL Id
            p.mclId = olCategory.CategoryID
            p.Color = olCategory.Color
            p.Active = True
            
            ' Capture the created Project
            newList.AddItem p
            
            ' Temporary
            ' If newList.count > 20 Then Exit For
            
        Next
    Else
        strTrace = "No categories exist - MCL was empty."
        LogMessage strTrace, strRoutine
    End If
        
    Set ImportFromMCL = newList
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set ImportFromMCL = Nothing

End Function

''' Returns the Project in the managed collection matching
''' the specified unique identifier
Public Function GetItemById(ByVal id As String) As fmeProject

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetItemById"
    
    On Error GoTo ThrowException
    
    Dim retProject As fmeProject
    Set retProject = myItems.GetItemById(id)
    
    If IsNothing(retProject) Then
        strTrace = "Failed to find a Project (" & id & ") in collection, items count: " & myItems.count
        LogMessage "WARNING: " & strTrace, strRoutine
    End If

    Set GetItemById = retProject
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetItemById = Nothing

End Function

''' Finds a project or projects referenced in a Category
Public Function ParseProjectsFromString(ByVal strCat As String) As ArrayList

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ParseProjectsFromString"
    
    On Error GoTo ThrowException

    If Len(strCat) = 0 Then
        strTrace = "Encountered an empty Categories string."
        GoTo ThrowException
    End If
    If myItems.Items.count = 0 Then
        strTrace = "Zero projects found try loading" & _
                    " the ProjectManager prior to calling this method."
        GoTo ThrowException
    End If
    
    Dim retList As ArrayList
    Set retList = New ArrayList
    
    Dim vCats() As String
    vCats = Split(strCat, ut.GetLocalizedDelimiter)
    
    Dim i As Integer
    Dim myCat As String
    Dim p As fmeProject
    For i = LBound(vCats) To UBound(vCats)
        myCat = vCats(i)
        If Not Len(myCat) = 0 Then
            For Each p In myItems.Items
                Dim pName As String
                pName = LCase(p.Name)
                If InStr(LCase(myCat), pName) > 0 Then
                    retList.Add p
                    Exit For
                End If
            Next
        End If
    Next
    
    Set ParseProjectsFromString = retList
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set ParseProjectsFromString = Nothing

End Function

' Sorts the current Outlook TaskItem collection
Public Sub Sort(ByVal srt As enuSortOn, ByVal dir As enuSortDirection)
   
    ' Capture Sort Variables
    sortingColumn = srt
    sortingDirection = dir
    
    ' Update UI
    Call Refresh
    
End Sub

''' Refreshes the view for managed UI controls
Public Sub Refresh()
    If Not myListView Is Nothing Then RefreshListView sortingColumn, sortingDirection
End Sub

' - - Project Commands

''' Opens an existing project
Public Sub OpenProject(Optional ByVal p As fmeProject = Nothing)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":OpenProject"

    If p Is Nothing Then
        Set p = selItem
    Else
        Set selItem = p
    End If
    
    If IsNothing(p) Then
        strTrace = "Incoming Project was null."
        GoTo ThrowException
    End If
    
    Dim frmP As New frm_ProjectProperties
    frmP.Load p
    frmP.Show False
    ' frmP.Show ' Dialog

   '' Prepare to present Outlook Task UI
   ' Set singleInspector = t.GetInspector
    
    '' Show the Task using the Outlook interface
    't.Display
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Creates a new Project to Edit
Public Sub NewProject()

    ' Show as a Dialog - to create a new Project
    Dim frmP As New frm_ProjectProperties
    frmP.Show

End Sub

Public Sub DeleteProject(Optional ByVal p As fmeProject = Nothing)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":DeleteProject"
    
    On Error GoTo ThrowException
    
    If p Is Nothing Then
        Set p = selItem
    Else
        Set selItem = p
    End If
    
    If IsNothing(p) Then
        strTrace = "Incoming Project was null."
        GoTo ThrowException
    End If
    
    Dim ds As New dsDataStore
    
    strTrace = "Are you sure you want to delete: " & vbCrLf & vbCrLf & _
                p.Subject
    Dim dResult As VbMsgBoxResult
    dResult = MsgBox(strTrace, vbQuestion Or vbYesNo, "Title")
    If dResult = vbYes Then
    
        ' Open the datastore
        ds.Connect
        
        ' Delete the Project from DataStore
        ds.Delete p, "Project"
            
        strTrace = "Removed Project: " & p.Subject & " from the datastore."
        LogMessage strTrace, strRoutine
        
        '  Delete the Project from the MCL
        Dim catName As String
        catName = p.Subject
        If Not p.CombineTitleCode Then catName = p.Name
        ut.RemoveFromMCL catName
        
        strTrace = "Removed Project: " & catName & " from the Master Category List."
        LogMessage strTrace, strRoutine
        
        ' Update UI
        DeleteListViewItem p

    End If
    
    GoTo Finally
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

Finally:
    ' Commit changes to the XML file if a 'delete' has occured
    If ds.IsConnected Then ds.Disconnect
    Set ds = Nothing

End Sub

''' Returns a Project if one has been assigned to the
''' specified Outlook Item
Public Sub GetProject(ByVal oItem As Object, ByRef p As fmeProject)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetProject"
    
    On Error GoTo ThrowException
    
    If oItem Is Nothing Then
        strTrace = "A null item was encountered."
        GoTo ThrowException
    End If
    
    Dim retProject As fmeProject
    Set retProject = Nothing
    
    Dim strCat As String
    strCat = oItem.Categories
    
    If Len(strCat) = 0 Then
        strTrace = "No projects found in Categories property - string is empty."
        GoTo ThrowException
    End If
    
    Dim arList As ArrayList
    Set arList = ParseProjectsFromString(strCat)
    If arList Is Nothing Then
        strTrace = "Fatal error while parsing the item's Categories."
        GoTo ThrowException
    End If
    If arList.count > 0 Then
        If arList.count = 1 Then
            Set retProject = arList(0)
        Else
            strTrace = "Multiple categories or projects found - unhandled exception."
            GoTo ThrowException
        End If
    Else
        strTrace = "Failed to parse categories: " & strCat
        GoTo ThrowException
    End If
    
    Set p = retProject
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set p = Nothing
    
End Sub

''' Creates a new task assigned to the current or specified project
Public Sub NewTask(Optional ByVal p As fmeProject = Nothing)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":NewTask"

    If p Is Nothing Then
        Set p = selItem
    Else
        Set selItem = p
    End If
    
    If IsNothing(p) Then
        strTrace = "Incoming Project was null."
        GoTo ThrowException
    End If
    
    ' Create a new Task
    Dim t As Outlook.TaskItem
    Set t = ut.CreateTask
    
    ' If a project provided, assign task to the Project
    t.Categories = p.Subject
    
    ' Prepare to present Outlook Task UI
    ' Set selItem = t
    ' Set singleInspector = t.GetInspector
    
    ' Show the Task using the Outlook interface
    t.Display
        
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

''' Creates a new email message assigned to the current or specified project
Public Sub NewMessage(Optional ByVal p As fmeProject = Nothing)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":NewMessage"

    If p Is Nothing Then
        Set p = selItem
    Else
        Set selItem = p
    End If
    
    If IsNothing(p) Then
        strTrace = "Incoming Project was null."
        GoTo ThrowException
    End If
    
    ' Create a new email message
    Dim m As Outlook.MailItem
    Set m = ut.CreateMailItem
    
    ' If a project provided, assign email to the Project
    m.Categories = p.Subject
    
    ' Prepare to present Outlook Task UI
    ' Set selItem = t
    ' Set singleInspector = t.GetInspector
    
    ' Show the MailItem using the Outlook interface
    m.Display

    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

''' Creates a new Appointment assigned to the current or specified project
Public Sub NewAppointment(Optional ByVal p As fmeProject = Nothing)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":NewAppointment"

    If p Is Nothing Then
        Set p = selItem
    Else
        Set selItem = p
    End If
    
    If IsNothing(p) Then
        strTrace = "Incoming Project was null."
        GoTo ThrowException
    End If
    
    ' Create a new appointmentItem
    Dim a As Outlook.AppointmentItem
    Set a = ut.CreateAppointment
    
    ' If a project provided, assign item to the Project
    a.Categories = p.Subject
    
    ' Prepare to present Outlook Task UI
    ' Set selItem = t
    ' Set singleInspector = t.GetInspector
    
    ' Show the Appointment using the Outlook interface
    a.Display

    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Opens the current or specified Project's Windows folder
Public Sub OpenWindowsFolder(Optional ByVal p As fmeProject = Nothing)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":OpenWindowsFolder"

    If p Is Nothing Then
        Set p = selItem
    Else
        Set selItem = p
    End If
    
    If IsNothing(p) Then
        strTrace = "Incoming Project was null."
        GoTo ThrowException
    End If
    
    If Len(p.WindowsFolder) = 0 Then
        strTrace = "No Windows folder specified for project: " & p.Subject
        GoTo ThrowException
    End If
    
    Dim strCmd As String
    strCmd = "explorer.exe " & p.WindowsFolder
    
    Dim retCode As Double
    retCode = Shell(strCmd, vbNormalFocus)
    If retCode = 0 Then
        strTrace = "Failed to run Shell command: " & strCmd
        GoTo ThrowException
    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Opens the current or specified Project's Outlook folder
Public Sub OpenOutlookFolder(Optional ByVal p As fmeProject = Nothing)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":OpenOutlookFolder"
    
    ut.SelectFolderFromPath p.OutlookFolder
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

' - Supporting Methods

Private Function GetProjectsFromDB() As ArrayList

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":GetProjectsFromDB"
    
    On Error GoTo ThrowException

    Dim ldb As New dsDataStore
    ldb.Connect ' Note no need to .Disconnect, only reading the data
    
    strTrace = "Get local Folder collection."
    Dim arItems As ArrayList
    Set arItems = ldb.GetCollection("Project")
    If arItems Is Nothing Then
        strTrace = "Local datastore query failed."
        GoTo ThrowException
    End If
    
    Set GetProjectsFromDB = arItems
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Function

''' Returns a collection of ContextMenuItems
Public Function GetContextMenu(ByVal p As fmeProject) As ContextMenu

    ' +-----------------
    ' | Open            (1001)
    ' | New          >> (1002)  +----------------+
    '                           | Task              (10021)
    '                           | Email             (10022)
    '                           | Appointment       (10023)
    '                           +----------------+
    ' | Open Folder  >> (1003)  +----------------+
    '                           | Windows Folder    (10031)
    '                           | Outlook Folder    (10032)
    '                           +----------------+
    ' +-----------------
    ' | Delete          (1020)
    ' +-----------------

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetContextMenu"
    
    On Error GoTo ThrowException
    
    Dim cm As ContextMenu
    Set cm = New ContextMenu

    Dim mni As ContextMenuItem
    
    Dim cmSub As ContextMenu
    
    ' Open
    Set mni = New ContextMenuItem
    mni.Caption = "Open..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1001
    cm.AddItem mni
    
    ' New Outlook Item Submenu
    Set cmSub = New ContextMenu
        
    ' - New Task
    Set mni = New ContextMenuItem
    mni.Caption = "Task..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 10021
    cmSub.AddItem mni
    
    ' - New Email
    Set mni = New ContextMenuItem
    mni.Caption = "Email..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 10022
    cmSub.AddItem mni
    
    ' - New Appointment
    Set mni = New ContextMenuItem
    mni.Caption = "Appointment..."
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 10023
    cmSub.AddItem mni
    
    ' Add New Outlook Item Submenu
    Set mni = New ContextMenuItem
    mni.Caption = "New"
    mni.Enabled = True
    mni.itemType = Menu
    mni.UID = 1002
    Set mni.SubMenu = cmSub
    cm.AddItem mni
    
    ' Open Folder
    If Len(p.OutlookFolder) > 0 Or Len(p.WindowsFolder) > 0 Then
        If Len(p.OutlookFolder) > 0 And Len(p.WindowsFolder) > 0 Then
            ' Prepare cascading menu
            
            ' New Open Folder SubMenu
            Set cmSub = New ContextMenu
            
            ' Open Windows Folder
            Set mni = New ContextMenuItem
            mni.Caption = "Windows Folder"
            mni.Enabled = True
            mni.itemType = Button
            mni.UID = 10031
            cmSub.AddItem mni
            
            ' Open Outlook Folder
            Set mni = New ContextMenuItem
            mni.Caption = "Outlook Folder"
            mni.Enabled = True
            mni.itemType = Button
            mni.UID = 10032
            cmSub.AddItem mni
        
            ' Add SubMenu to Main Menu
            Set mni = New ContextMenuItem
            mni.Caption = "Open Folder"
            mni.Enabled = True
            mni.itemType = Menu
            mni.UID = 1003
            Set mni.SubMenu = cmSub
            cm.AddItem mni
            
        Else
            ' Prepare one menuItem
            If Len(p.OutlookFolder) > 0 Then
                Set mni = New ContextMenuItem
                mni.Caption = "Outlook Folder"
                mni.Enabled = True
                mni.itemType = Button
                mni.UID = 10032
                cm.AddItem mni
            End If
            If Len(p.WindowsFolder) > 0 Then
                Set mni = New ContextMenuItem
                mni.Caption = "Windows Folder"
                mni.Enabled = True
                mni.itemType = Button
                mni.UID = 10031
                cm.AddItem mni
            End If
        End If
    Else
        strTrace = "No folders specified for this project: " & p.Subject
        LogMessage strTrace, strRoutine
    End If
    
    ' Separator
    Set mni = New ContextMenuItem
    mni.Caption = "-"
    mni.Enabled = True
    mni.itemType = Separator
    mni.UID = 1099
    cm.AddItem mni
    
    ' Delete
    Set mni = New ContextMenuItem
    mni.Caption = "Delete"
    mni.Enabled = True
    mni.itemType = Button
    mni.UID = 1020
    cm.AddItem mni
    
    Set GetContextMenu = cm
    
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetContextMenu = Nothing

End Function

' - - Menus Supporting Methods

Public Sub ProcessContextMenuSelection(ByVal id As String, ByRef p As fmeProject)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":ProcessContextMenuSelection"
    
    If Len(id) = 0 Then
        strTrace = "No menu id provided."
        GoTo ThrowException
    End If
    If IsNothing(p) Then
        strTrace = "A null project encountered."
        GoTo ThrowException
    End If
    
    Dim num As Integer
    num = CInt(id)
    Select Case num
        Case 1001 ' Open Project
            OpenProject p
        Case 1002 ' New Outlook Item Menu
            ' Ignored, handled in the ContextMenu
        Case 10021 ' New Task
            NewTask p
        Case 10022 ' New Email
            NewMessage p
        Case 10023 ' New Appointment
            NewAppointment p
        Case 1003 ' Open Folder subMenu
            ' Ignored, handled by ContextMenu
        Case 10031 ' Open Windows Folder
            OpenWindowsFolder p
        Case 10032 ' Open Outlook Folder
            OpenOutlookFolder p
        Case 1099 ' Seperator
            strTrace = "Ignoring call."
        Case Else
            strTrace = "Unrecognised menu code: " & id & " - ignoring request."
            LogMessage strTrace, strRoutine
    End Select
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

' - - Collection Supporting Methods

Private Sub TrackSelectedItems(ByVal p As fmeProject, ByVal Action As String)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":TrackSelectedItems"
    
    On Error GoTo ThrowException
    
    If IsNothing(p) Then
        strTrace = "A null Project encountered."
        GoTo ThrowException
    End If
    
    Dim exProject As fmeProject
    Set exProject = selItems.GetItemById(p.id)
    
    If InStr(1, LCase(Action), "add") > 0 Then
        If Not exProject Is Nothing Then
            strTrace = "Cannot add a Project more than once to the selected Items, id: " & p.id
            LogMessage "WARNING: " & strTrace, strRoutine
        Else
            selItems.AddItem p
        End If
    End If
    If InStr(1, LCase(Action), "rem") > 0 Then
        If exProject Is Nothing Then
            strTrace = "Cannot remove a project that is not currently in the selected items, id: " & p.id
            LogMessage "WARNING: " & strTrace, strRoutine
        Else
            selItems.RemoveItem p
        End If
    End If
      
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

' - - ListView Supporting Methods

Private Sub UpdateListViewUI(ByVal p As fmeProject)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":UpdateListViewUI"
    
    If p Is Nothing Then
        strTrace = "A null Project encountered."
        GoTo ThrowException
    End If
    
    On Error GoTo ThrowException
    
    Dim li As ListItem
    Set li = FindListViewItem(p)
    If Not IsNothing(li) Then
        ' UpdateListViewItem
        UpdateListViewItem p
    Else
        ' AddListViewItem
        AddListViewItem p
    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Sub AddListViewItem(ByVal p As fmeProject, Optional ByVal idx As Integer = -1)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":AddListViewItem"
    
    On Error GoTo ThrowException

    If IsNothing(p) Then
        strTrace = "A null Item encountered."
        GoTo ThrowException
    End If
    
    strTrace = "Processing project: '" & p.Name & "'."
    
    Dim subs(1) As String
    
    
    ' Check the index
    If idx < 0 Then idx = myListView.ListItems.count + 1
    
    ' Set the key
    Dim key As String
    key = CodeProjectID(p.id) ' ListView keys do not accept numeric values

    strTrace = "Adding Item to ListView (Index: " & idx & ", Key: " & key & ", Text: " & p.Name & ")."
    LogMessage strTrace, strRoutine
    
    Dim li As ListItem
    Set li = myListView.ListItems.Add(idx, key, p.Name)
    
    strTrace = "Added listItem: " & p.Name
    LogMessage strTrace, strRoutine
    
    ' Dim lisb As ListSubItem
    If Len(p.Code) > 0 Then
        ' Set lisb = li.ListSubItems.Add(, , p.Code, "Project Code")
        
        strTrace = "Adding a subItem to the listItem: " & p.Code
        LogMessage strTrace, strRoutine
        
        li.SubItems(1) = p.Code
        
        strTrace = "Added a subItem to the listItem: " & p.Code
        LogMessage strTrace, strRoutine
              
    End If
           
    ' Format the row
    FormatLVRow li, p
    
    strTrace = "Added item to ListView: " & p.Subject
    LogMessage strTrace, strRoutine
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub
Private Sub UpdateListViewItem(ByVal p As fmeProject)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":UpdateListViewItem"
    
    On Error GoTo ThrowException
    
    If IsNothing(p) Then
        strTrace = "A null Item encountered."
        GoTo ThrowException
    End If
    
    Dim li As ListItem
    Set li = FindListViewItem(p)
    If Not IsNothing(li) Then
        ' Update list view here
        li.text = p.Name
        li.SubItems(1) = p.Code
                      
        ' Format the row
        FormatLVRow li, p
        
        strTrace = "Updated ListView for item: " & p.Subject
    Else
        strTrace = "Add new item to the ListView."
        AddListViewItem p
    End If
    
    LogMessage strTrace, strRoutine
    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub
Private Sub DeleteListViewItem(ByVal p As fmeProject)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":DeleteListViewItem"
    
    On Error GoTo ThrowException
    
    If IsNothing(p) Then
        strTrace = "A null Item encountered."
        GoTo ThrowException
    End If

    Dim li As ListItem
    Set li = FindListViewItem(p)
    If Not IsNothing(li) Then
        myListView.ListItems.Remove li.Index
    Else
        strTrace = "WARNING: unable to find item: " & p.Subject & " in the listview."
        GoTo ThrowException
    End If
    
    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub
Private Function FindListViewItem(ByVal p As fmeProject) As ListItem

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":FindListViewItem"
    
    On Error GoTo ThrowException
    
    If IsNothing(p) Then
        strTrace = "A null Item encountered."
        GoTo ThrowException
    End If
    
    Dim bFnd As Boolean
    bFnd = False
    
    Dim retItem As ListItem
    Set retItem = Nothing
    
    Dim searchKey As String
    searchKey = CodeProjectID(p.id) ' See AddListViewItem method
    
    Dim i As Integer
    For i = 1 To myListView.ListItems.count
        Dim li As ListItem
        Set li = myListView.ListItems(i)
        If li.key = searchKey Then
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
    
    On Error GoTo ThrowException
    
    ' Reset - multiple uses for the LV control (projects, tasks, ...)
    myListView.ColumnHeaders.Clear
    
    ' Setup ListView Columns and Configuration
    If myListView.ColumnHeaders.count = 0 Then
        'Initialize the View
        Dim ch1 As ColumnHeader
        Dim ch2 As ColumnHeader
        With myListView
            Set ch1 = .ColumnHeaders.Add(1, "Name", "Name")
            Set ch2 = .ColumnHeaders.Add(2, "Code", "Code")
            
            .Checkboxes = lvCheckbox
            .MultiSelect = lvMultiSelect
            .View = lvwReport
            
        End With
    End If
    
    ' Clear current LV collection
    myListView.ListItems.Clear
    
    ' Add current class collection
    If myItems.count > 0 Then
    
        ' Sort the internal list
        Dim strSort As String
        strSort = "Name"
        If sortOn = Code Then strSort = "Code"
        
        Dim collSort As New SortCollection
        collSort.Sort strSort, myItems.Items, sortDir
          
        ' Load the ListView
        Dim i As Integer
        For i = 0 To myItems.Items.count - 1
            Dim p As fmeProject
            Set p = myItems.Item(i)
                                 
            ' Add Item to ListView
            AddListViewItem p, i + 1
            
        Next
        
        Call ResizeLVColumns
        
    End If
    
    strTrace = "Items: " & myItems.count & " items..."
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Sub FormatLVRow(ByVal li As ListItem, ByVal p As fmeProject)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":FormatLVRow"
    
    On Error GoTo ThrowException

    ' Format the checkbox
    ' li.checked = p.Complete
    
    ' Color the Item
    If p.Priority = olImportanceHigh Then
        li.ForeColor = &H80& ' Magenta
    Else
        li.ForeColor = &H80000007 ' Black
    End If
                       
    ' Show as completed if appropriate
    If p.Complete Then
        li.ForeColor = &HC0C0C0 ' Light Gray
    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Sub ResizeLVColumns()

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":ResizeLVColumns"
    
    Dim strColumnWidths As String
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
    
    Dim i As Integer
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
                              ByVal Y As stdole.OLE_YPOS_PIXELS, _
                     Optional ByVal factor As Integer = 15) As ListItem


    ' Convert Pixels to TWIPS
    ' - .net uses Pixels, VBA uses TWIPS for ListView and TreeView (OLE_PIXELS?)
    ' - "on most computers 1 pixel = 15 TWIPS"
    '    https://stackoverflow.com/questions/36442535/vba-drag-drop-from-treeview-to-listview-listview-to-treeview-activex-controls
    Dim xInt As Single
    xInt = X
    Dim yInt As Single
    yInt = Y
           
    Dim li As ListItem
    Set li = myListView.HitTest(X * factor, Y * factor)
        
    Set LV_GetItemAt = li
        
End Function

Private Function CodeProjectID(ByVal pid As Variant) As Variant

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":CodeProjectID"
    
    If Len(pid) = 0 Then
        strTrace = "Incoming project ID was empty or null."
        GoTo ThrowException
    End If
    
    Dim retKey As Variant
    retKey = "A" & pid
    
    CodeProjectID = retKey
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Function

Private Function DecodeProjectID(ByVal key As Variant) As Variant

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":DecodeProjectID"
    
    If Len(key) = 0 Then
        strTrace = "Incoming Key was empty or null."
        GoTo ThrowException
    End If
    
    Dim retPID As Variant
    
    Dim hook As Integer
    hook = InStr(1, key, "A")
    If hook > 0 Then
        retPID = Right(key, Len(key) - hook)
    Else
        strTrace = "Failed to parse the Key: '" & key & "'."
        GoTo ThrowException
    End If
    
    DecodeProjectID = retPID
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Function

