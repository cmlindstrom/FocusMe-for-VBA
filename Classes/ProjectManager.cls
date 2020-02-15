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
    eid = li.key
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
    eid = Item.key
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
    eid = Item.key
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

Private Sub myListView_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)

    If f_SuspendUI Then Exit Sub

    Dim strTrace As String
    strTrace = ""
    Dim strRoutine As String
    strRoutine = rootClass & ":myListView_MouseDown"
    
    ' Catch a context menu call
    
    If Button = 2 Then
        ' Render the context menu
    End If
    
End Sub

Private Sub myListView_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)

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

Private Sub myListView_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single, State As Integer)
    Effect = 1
    Set myListView.DropHighlight = myListView.HitTest(X * 15, y * 20)
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
            End If
            
            ' Capture MCL Id
            p.mclId = olCategory.CategoryID
            p.Color = olCategory.Color
            p.Active = True
            
            ' Capture the created Project
            newList.AddItem p
            
            ' Temporary
            If newList.count > 20 Then Exit For
            
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
    
    ' Check the index
    If idx < 0 Then idx = myListView.ListItems.count + 1

    strTrace = "Adding Item to ListView (" & idx & "," & p.id & "," & p.Name & ")."
    Dim li As ListItem
    Set li = myListView.ListItems.Add(idx, p.id, p.Name)
    li.SubItems(1) = p.Code
           
    ' Format the row
    FormatLVRow li, p
    
    strTrace = "Added item to ListView: " & p.Subject
    
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
    
    Dim i As Integer
    For i = 1 To myListView.ListItems.count
        Dim li As ListItem
        Set li = myListView.ListItems(i)
        If li.key = p.id Then
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

