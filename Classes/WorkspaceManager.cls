VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkspaceManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' - Fields

Private Const rootClass As String = "WorkspaceManager"

Private WithEvents myListView As MSComctlLib.ListView
Attribute myListView.VB_VarHelpID = -1

Dim sortingColumn As enuSortOn
Dim sortingDirection As enuSortDirection

' - Events

Public Event ItemUpdated()
Public Event ItemAdded()
Public Event collectionUpdated()

' - Properties

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

Private Sub myListView_AfterLabelEdit(Cancel As Integer, NewString As String)

    If f_SuspendUI Then Exit Sub

    ' Update the Workspace Name
End Sub


' - Constructor

Private Sub Class_Initialize()

    ' Initialize Variables
    Set myItems = New Workspaces
    Set selItems = New Workspaces
    
    ' Default Values
    lvCheckbox = False
    lvMultiSelect = False
    f_SuspendUI = False
    
    sortingColumn = name
    sortingDirection = Ascending
       
End Sub

Private Sub Class_Terminate()

    ' Clean Up Memory
    Set myItems = Nothing
    Set selItems = Nothing
    
End Sub

' - Methods

''' Load all Projects from the DataStore
Public Sub Load()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Load"
    
    On Error GoTo ThrowException
    
    ' Retrieve Projects from DataStore
    Set myItems.Items = GetItemsFromDB
    
    ' Inform Listeners
    RaiseEvent collectionUpdated
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

''' Returns the Workspace in the managed collection matching
''' the specified unique identifier
Public Function GetItemById(ByVal id As String) As fmeWorkspace

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetItemById"
    
    On Error GoTo ThrowException
    
    Dim retItem As fmeWorkspace
    Set retItem = myItems.GetItemById(id)
    
    If IsNothing(retItem) Then
        strTrace = "Failed to find a Workspace (" & id & ") in collection, items count: " & myItems.Count
        LogMessage "WARNING: " & strTrace, strRoutine
    End If

    Set GetItemById = retItem
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetItemById = Nothing

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

' - - Workspace Commands

' - Supporting Methods

Private Function GetItemsFromDB() As ArrayList

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":GetItemsFromDB"
    
    On Error GoTo ThrowException

    Dim ldb As New dsDataStore
    ldb.Connect ' Note no need to .Disconnect, only reading the data
    
    strTrace = "Get local Folder collection."
    Dim arItems As ArrayList
    Set arItems = ldb.GetEntireCollection("Workspace")
    If arItems Is Nothing Then
        strTrace = "Local datastore query failed."
        GoTo ThrowException
    End If
    
    Set GetItemsFromDB = arItems
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetItemsFromDB = New ArrayList
    
End Function

' - - ListView Supporting Methods

Private Sub UpdateListViewUI(ByVal w As fmeWorkspace)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":UpdateListViewUI"
    
    If w Is Nothing Then
        strTrace = "A null Workspace encountered."
        GoTo ThrowException
    End If
    
    On Error GoTo ThrowException
    
    Dim li As ListItem
    Set li = FindListViewItem(w)
    If Not IsNothing(li) Then
        ' UpdateListViewItem
        UpdateListViewItem w
    Else
        ' AddListViewItem
        AddListViewItem w
    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Sub AddListViewItem(ByVal w As fmeWorkspace, Optional ByVal idx As Integer = -1)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":AddListViewItem"
    
    On Error GoTo ThrowException

    If IsNothing(w) Then
        strTrace = "A null Item encountered."
        GoTo ThrowException
    End If
    
    ' Check the index
    If idx < 0 Then idx = myListView.ListItems.Count + 1

    strTrace = "Adding Item to ListView (" & idx & "," & w.id & "," & w.name & ")."
    Dim li As ListItem
    Set li = myListView.ListItems.Add(idx, w.id, w.name)
    li.SubItems(1) = w.Code
           
    ' Format the row
    FormatLVRow li, w
    
    strTrace = "Added item to ListView: " & w.name
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub
Private Sub UpdateListViewItem(ByVal w As fmeWorkspace)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":UpdateListViewItem"
    
    On Error GoTo ThrowException
    
    If IsNothing(w) Then
        strTrace = "A null Item encountered."
        GoTo ThrowException
    End If
    
    Dim li As ListItem
    Set li = FindListViewItem(w)
    If Not IsNothing(li) Then
        ' Update list view here
        li.text = w.name
        li.SubItems(1) = w.Code
                      
        ' Format the row
        FormatLVRow li, w
        
        strTrace = "Updated ListView for item: " & w.name
    Else
        strTrace = "Add new item to the ListView."
        AddListViewItem w
    End If
    
    LogMessage strTrace, strRoutine
    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub
Private Sub DeleteListViewItem(ByVal w As fmeWorkspace)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":DeleteListViewItem"
    
    On Error GoTo ThrowException
    
    If IsNothing(w) Then
        strTrace = "A null Item encountered."
        GoTo ThrowException
    End If

    Dim li As ListItem
    Set li = FindListViewItem(w)
    If Not IsNothing(li) Then
        myListView.ListItems.Remove li.Index
    Else
        strTrace = "WARNING: unable to find item: " & w.name & " in the listview."
        GoTo ThrowException
    End If
    
    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub
Private Function FindListViewItem(ByVal w As fmeWorkspace) As ListItem

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":FindListViewItem"
    
    On Error GoTo ThrowException
    
    If IsNothing(w) Then
        strTrace = "A null Item encountered."
        GoTo ThrowException
    End If
    
    Dim bFnd As Boolean
    bFnd = False
    
    Dim retItem As ListItem
    Set retItem = Nothing
    
    Dim i As Integer
    For i = 1 To myListView.ListItems.Count
        Dim li As ListItem
        Set li = myListView.ListItems(i)
        If li.key = w.id Then
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
    If myListView.ColumnHeaders.Count = 0 Then
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
    If myItems.Count > 0 Then
    
        ' Sort the internal list
        Dim strSort As String
        strSort = "Name"
        If sortOn = Code Then strSort = "Code"
        
        Dim collSort As New SortCollection
        collSort.Sort strSort, myItems.Items, sortDir
          
        ' Load the ListView
        Dim i As Integer
        For i = 0 To myItems.Items.Count - 1
            Dim w As fmeWorkspace
            Set w = myItems.Item(i)
                                 
            ' Add Item to ListView
            AddListViewItem w, i + 1
            
        Next
        
        Call ResizeLVColumns
        
    End If
    
    strTrace = "Items: " & myItems.Count & " items..."
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Sub FormatLVRow(ByVal li As ListItem, ByVal w As fmeWorkspace)

    ' Format the checkbox
    ' li.checked = p.Complete
    
    ' Color the Item
    'If p.Priority = olImportanceHigh Then
    '    li.ForeColor = &H80& ' Magenta
    'Else
    '    li.ForeColor = &H80000007 ' Black
    'End If
                       
    ' Show as completed if appropriate
    'If p.Complete Then
    '    li.ForeColor = &HC0C0C0 ' Light Gray
    'End If

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
        bScrollbar = (.font.SIZE + 4 + 1) * .ListItems.Count > .Height
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

Private Function LV_GetItemAt(ByVal x As stdole.OLE_XPOS_PIXELS, _
                              ByVal y As stdole.OLE_YPOS_PIXELS, _
                     Optional ByVal factor As Integer = 15) As ListItem


    ' Convert Pixels to TWIPS
    ' - .net uses Pixels, VBA uses TWIPS for ListView and TreeView (OLE_PIXELS?)
    ' - "on most computers 1 pixel = 15 TWIPS"
    '    https://stackoverflow.com/questions/36442535/vba-drag-drop-from-treeview-to-listview-listview-to-treeview-activex-controls
    Dim xInt As Single
    xInt = x
    Dim yInt As Single
    yInt = y
           
    Dim li As ListItem
    Set li = myListView.HitTest(x * factor, y * factor)
        
    Set LV_GetItemAt = li
        
End Function


