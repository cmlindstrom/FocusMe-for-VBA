VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FME_Pane 
   Caption         =   "Task List"
   ClientHeight    =   8235.001
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   4710
   OleObjectBlob   =   "FME_Pane.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "FME_Pane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True



' - - Fields

Private Const rootClass As String = "FME_Pane"
Private Const FormTitle As String = "Task List - " & Commands.AppName

Dim leftMouseDown As Boolean
Dim rightMouseDown As Boolean

Private Enum enuController
    None = 0
    TaskManager = 1
    ProjectManager = 2
End Enum
Dim controller As enuController

Dim tabView As String

' Form
Dim frmX As Integer
Dim frmY As Integer

' ListBox
Dim lstX As Integer
Dim lstY As Integer

' ListView
Dim lvX As Integer
Dim lvY As Integer

Dim WithEvents tm As TaskManager
Attribute tm.VB_VarHelpID = -1
Dim WithEvents pm As ProjectManager
Attribute pm.VB_VarHelpID = -1

Dim myList As ArrayList

Dim strColumnWidths As String
Dim numColumns As Integer

Dim sortingColumn As enuSortOn
Dim sortingDirection As enuSortDirection

' - - Resizing
Private m_clsAnchors As CAnchors

' - -  Settings
Private stgs As Settings

' - -
Public Event FormClosed()
Public Event FormClosing()

' - - Properties
Dim selItem As Object
Dim f_origin As Point
Dim f_searchMode As Boolean

Public Property Get CurrentItem() As Object
    If controller = TaskManager Then Set selItem = tm.SelectedItem
    If controller = ProjectManager Then Set selItem = pm.SelectedItem
    Set CurrentItem = selItem
End Property


Public Property Get SearchMode() As Boolean
    SearchMode = f_searchMode
End Property

Public Property Get CurrentView() As String
    CurrentView = tabView
End Property



' - - Event Handlers

' - - - ProjectManager Events

Private Sub pm_collectionUpdated()

    Dim strTrace As String
    strTrace = "Items: " & pm.Items.count & " Projects..."
    Status strTrace

End Sub

' - - - TaskManager Events

Private Sub tm_collectionUpdated()

    Dim strTrace As String
    strTrace = "Items: " & tm.Items.count & " Tasks..."
    Status strTrace
    
End Sub

' - - Form Handlers

' - - - Buttons

Private Sub btn_Add_Click()
    If controller = TaskManager Then tm.NewTask
    If controller = ProjectManager Then pm.NewProject
End Sub

Private Sub btn_Edit_Click()
    If controller = TaskManager Then tm.OpenTask
    If controller = ProjectManager Then pm.OpenProject
End Sub

Private Sub btn_Delete_Click()
    If controller = TaskManager Then tm.DeleteTask
    If controller = ProjectManager Then pm.DeleteProject
End Sub

Private Sub btn_NewMail_Click()

    ' Present Conversation List and Activate Inbox
    ThisOutlookSession.GoToMailPane

    ' Clear count
    btn_NewMail.Caption = "0"
    
    ' Turn off Button
    btn_NewMail.Visible = False
    
End Sub

Private Sub btn_Options_Click()

    Dim frm As New frm_Options
    frm.Show

End Sub

Private Sub btn_Related_Click()
    If controller = TaskManager Then tm.DisplayRelatedMail

End Sub

Private Sub btn_Refresh_Click()
    If controller = TaskManager Then tm.Refresh
    If controller = ProjectManager Then pm.Refresh
End Sub

Private Sub btn_Projects_Click()
   ' ThisOutlookSession.StartAllProjects
   
   Dim strTrace As String
   
   If Contains("Project", tabView) Then
        strTrace = "Search not supported for the Projects view."
        MsgBox strTrace, vbOKOnly Or vbInformation, Commands.AppName
        Exit Sub
   End If
   
   If f_searchMode Then
        ' Process Search
        RefreshView tabView, txtbx_Search.text
   Else
        ' Establish Search Mode
        SetSearchMode True
   End If
   
End Sub

Private Sub btn_searchOff_Click()
    SetSearchMode False
    RefreshView tabView
End Sub

Private Sub txtbx_Search_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    Dim strTrace As String

    If KeyCode = 13 Then
        If f_searchMode Then RefreshView tabView, txtbx_Search.text
    End If
    
End Sub

Private Sub txtbx_Status_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        anchorX = X
        anchorY = Y
    End If
End Sub

Private Sub txtbx_Status_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Dim strTrace As String
    
    If Button = 1 Then
        ' Move the window
        strTrace = "Left button down: " & leftMouseDown & " X: " & X
        ' LogMessage strTrace, rootClass & ":txtbx_Status_MouseMove"
        MoveFMEWindow X, Y
        
        
    End If
    
End Sub

Private Sub MoveFMEWindow(ByVal X As Single, ByVal Y As Single)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":MoveWindow"
    
    On Error GoTo ThrowException
    
    Dim newX As Long
    Dim newY As Long
    
    Dim diffX As Long
    Dim diffY As Long
    diffX = X - anchorX
    diffY = anchorY - Y

    newX = Me.Left + diffX
    newY = Me.Top + diffY
    
'    If SetFormPosition(Me, newX, newY) Then
'
'    Else
'
'    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Dim strTrace As String

    If Button = 1 Then
        strTrace = "Left Mouse Up."
        leftMouseDown = False
    End If
    If Button = 2 Then
        strTrace = "Right Mouse Up."
        rightMouseDown = False
    End If

End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Dim strTrace As String

    If Button = 1 Then
        strTrace = "Left Mouse Down."
        leftMouseDown = True
    End If
    If Button = 2 Then
        strTrace = "Right Mouse Down."
        rightMouseDown = True
    End If

End Sub

' - Form Event Handlers
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Dim Effect As Integer

    If leftMouseDown Then
       ' Dim myDO As DataObject
       ' myDO = New DataObject
        
       ' Effect = myDO.StartDrag
        
        ' Track location of mouse pointer
        frmX = X
        frmY = Y
    End If

End Sub


' - TabStrip Event Handlers

Private Sub ts_Tasks_Change()

    Dim tb As Object
    Set tb = ts_Tasks.SelectedItem
    tabView = tb.Caption
    
    ' Refresh the item list
    RefreshView tabView
    
    ' Control buttons to present
    Call SetButtons
        
End Sub

' - - Constructor

Private Sub UserForm_Initialize()

    Set m_clsAnchors = New CAnchors
    
    Set m_clsAnchors.Parent = Me
    
    ' restrict minimum size of userform
    m_clsAnchors.MinimumWidth = 247.5
    m_clsAnchors.MinimumHeight = 441
    
    ' Set Anchors
    With m_clsAnchors
    
        ' Tab Strip
        .Anchor("ts_Tasks").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleRight Or enumAnchorStyleTop
        
        ' ListView
        With .Anchor("lv_tasks")
            .AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleRight _
                            Or enumAnchorStyleTop Or enumAnchorStyleBottom
            .MinimumHeight = 348
        End With
        
        ' Buttons
        .Anchor("btn_Add").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleBottom
        .Anchor("btn_Edit").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleBottom
        .Anchor("btn_Delete").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleBottom
        
        .Anchor("btn_Options").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleBottom
        .Anchor("btn_Related").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleBottom
        
        .Anchor("btn_NewMail").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleBottom
        .Anchor("btn_Projects").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleBottom
        .Anchor("btn_searchOff").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleBottom
        
        ' TextBoxes
        .Anchor("txtbx_Search").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleRight Or enumAnchorStyleBottom
        
        ' Status Bar
        .Anchor("txtbx_Status").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleRight Or enumAnchorStyleBottom
    
    End With
    
    ' Configure Controls
    ' - Set form title
    Me.Caption = FormTitle
    
    ' - Create TabStrip tabs
    Me.ts_Tasks.Tabs.Clear
    Me.ts_Tasks.Tabs.Add "daily", "Daily", 0
    Me.ts_Tasks.Tabs.Add "due", "Due", 1
    Me.ts_Tasks.Tabs.Add "highpriority", "Important", 2
    Me.ts_Tasks.Tabs.Add "waiting", "Waiting", 3
    Me.ts_Tasks.Tabs.Add "master", "Master", 4
    Me.ts_Tasks.Tabs.Add "projects", "Projects", 5
    
    ' - Select the default start tab & controller
    Me.ts_Tasks.value = 0 ' Daily
    tabView = "Daily"
    controller = TaskManager
    
    Call ClearMail
    
    ' Initialize variables
    Set stgs = New Settings
    Set f_origin = New Point
    
    Set myList = New ArrayList
    
    Set tm = New TaskManager
    Set tm.ListView = lv_Tasks
    
    Set pm = New ProjectManager
    Set pm.ListView = lv_Tasks
    
    strColumnWidths = "15;120;40;0"
    numColumns = 4
    
    sortingColumn = DueDate
    sortingDirection = Ascending
    
    f_searchMode = False
    SetSearchMode f_searchMode
    
    RefreshView tabView
   
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    RaiseEvent FormClosing
    
    Set f_origin = Nothing
    Set selItem = Nothing
    Set myList = Nothing
    Set m_clsAnchors = Nothing
    Set tm = Nothing
    Set pm = Nothing
    
    RaiseEvent FormClosed
        
End Sub

Private Sub UserForm_Terminate()

    Set f_origin = Nothing
    Set selItem = Nothing
    Set myList = Nothing
    Set m_clsAnchors = Nothing
    Set tm = Nothing
    Set pm = Nothing
    
End Sub

Private Sub UserForm_Activate()

    ' Set to saved position
    Call RepositionForm
    
End Sub

Private Sub UserForm_Resize()
    ' record position and size
    Call RecordPosition
End Sub

' - - Methods

Public Sub RecordPosition()

    Dim X As Long
    Dim Y As Long
    Call WinForms.GetFormPosition(Me, X, Y)
    
    Dim pt As New Point
    pt.X = X
    pt.Y = Y
    Set stgs.TaskPaneLocation = pt
    stgs.Save

End Sub

Public Sub RepositionForm(Optional ByVal iX As Integer = -1, Optional ByVal iY As Integer = -1)

    ' Get current screen configuration
    Dim w As Integer    ' screen width
    Dim h As Integer    ' screen height
    Dim n As Integer    ' # of monitors
    GetScreenResolution h, w, n

    Dim pt As Point
    Set pt = stgs.TaskPaneLocation
    Dim X As Integer
    X = pt.X
    Dim Y As Integer
    Y = pt.Y
    
    If iX = -1 Then
        ' no incoming position specified, check for off-screen
        If X + Me.Width > w Then
            ' last saved position, puts form off of the screen
            If n = 1 Then
                ' monitor may have changed since last saved position, adjust position
                X = w - Me.Width - 10
            Else
                ' most likely coordinates are ok, since > 1 monitor
            End If
        End If
    Else
        ' Set position to specific arguments
        X = iX
        Y = iY
    End If
    
    f_origin.X = X
    f_origin.Y = Y
    
    Call SetFormPosition(Me, f_origin.X, f_origin.Y)

End Sub

''' Report New Mail
Public Sub NewMail(Optional ByVal numMail As Integer = 1)

    If Not btn_NewMail.Visible Then btn_NewMail.Visible = True
    Dim numCurrent As Integer
    numCurrent = CInt(btn_NewMail.Caption)
    numCurrent = numCurrent + numMail
    btn_NewMail.Caption = CStr(numCurrent)
    
    Call RecordPosition
    
End Sub

''' Clears the Mail Flag
Public Sub ClearMail()
    btn_NewMail.Visible = False
    btn_NewMail.Caption = "0"
End Sub

''' Updates the status bar of the form
Public Sub Status(Optional msg As String = "")
    Me.txtbx_Status.text = msg
End Sub

' - -  Supporting Methods

Private Sub SetButtons()

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":SetButtons"

    If controller = TaskManager Then
        btn_Related.Visible = True
        btn_Projects.Visible = True
    End If
    If controller = ProjectManager Then
        btn_Related.Visible = False
        btn_Projects.Visible = False
    End If
    
End Sub

Private Sub RefreshView(ByVal tabName As String, _
                Optional ByVal strFilter As String = "")

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":RefreshView"
    
    On Error GoTo ThrowException

    Dim strTab As String
    strTab = LCase(tabName)
    
    ' Default to TaskManager
    controller = TaskManager
    ' Turn on Tasks UI
    tm.SuspendUIEvents = False
    
    ' Shut off other Manager listeners
    pm.SuspendUIEvents = True
    
    ' Get the selected view
    Select Case strTab
        Case "daily"
            tm.LoadSpecial enuTaskFilters.Daily
            Set myList = tm.Items.Items
            
            If Len(strFilter) > 0 Then tm.Filter strFilter
            
        Case "due"
            tm.LoadSpecial enuTaskFilters.PastDue
            Set myList = tm.Items.Items
            
            If Len(strFilter) > 0 Then tm.Filter strFilter
            
        Case "important"
            tm.LoadSpecial enuTaskFilters.HighPriority
            Set myList = tm.Items.Items

            If Len(strFilter) > 0 Then tm.Filter strFilter

        Case "master"
            tm.LoadSpecial enuTaskFilters.Master
            Set myList = tm.Items.Items
            
            If Len(strFilter) > 0 Then tm.Filter strFilter
            
        Case "unassigned"
            tm.LoadSpecial enuTaskFilters.NoCategory
            Set myList = tm.Items.Items
            
            If Len(strFilter) > 0 Then tm.Filter strFilter
            
        Case "high"
            tm.LoadSpecial enuTaskFilters.HighPriority
            Set myList = tm.Items.Items
            
            If Len(strFilter) > 0 Then tm.Filter strFilter
            
        Case "waiting"
            tm.LoadSpecial enuTaskFilters.Waiting
            Set myList = tm.Items.Items
        
            If Len(strFilter) > 0 Then tm.Filter strFilter
        
        Case "projects"
            ' Turn on Projects UI
            pm.SuspendUIEvents = False
            ' Turn off Tasks UI
            tm.SuspendUIEvents = True
            ' Get the latest collection from the dataStore
            pm.Load
            ' Refresh the UI
            pm.Refresh
            ' Capture the collection (ArrayList)
            Set myList = pm.Items.Items
            ' Specify the Controller
            controller = ProjectManager
            
        Case Else
            strTrace = "Unsupported tab: " & strTab
            GoTo ThrowException
        
    End Select

    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Sub AddListViewItem(ByVal t As Outlook.TaskItem, Optional ByVal idx As Integer = -1)

    ' Check the index
    If idx < 0 Then idx = lv_Tasks.ListItems.count + 1

    ' Add Item to ListView
    Dim li As ListItem
    Set li = lv_Tasks.ListItems.Add(idx, t.EntryId, t.Subject)
    If IsDateNone(t.DueDate) Then
        li.SubItems(1) = "None"
    Else
        li.SubItems(1) = Format(t.DueDate, "mm/dd/yyyy")
    End If
           
    ' Format the row
    FormatLVRow li, t
    
    strTrace = "Added task to ListView: " & t.Subject

End Sub
Private Sub UpdateListViewItem(ByVal t As Outlook.TaskItem)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":UpdateListViewItem"
    
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
    
    On Error GoTo ThrowException
    
    If IsNothing(t) Then
        strTrace = "A null taskItem encountered."
        GoTo ThrowException
    End If

    Dim li As ListItem
    Set li = FindListViewItem(t)
    If Not IsNothing(li) Then
        Me.lv_Tasks.ListItems.Remove li.Index
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
    
    Dim bFnd As Boolean
    bFnd = False
    
    Dim retItem As ListItem
    Set retItem = Nothing
    
    For i = 1 To lv_Tasks.ListItems.count
        Dim li As ListItem
        Set li = lv_Tasks.ListItems(i)
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
    
    ' Setup ListView Columns and Configuration
    If lv_Tasks.ColumnHeaders.count = 0 Then
        'Initialize the View
        Dim ch1 As ColumnHeader
        Dim ch2 As ColumnHeader
        With lv_Tasks
            Set ch1 = .ColumnHeaders.Add(1, "Subject", "Description")
            Set ch2 = .ColumnHeaders.Add(2, "DueDate", "Date Due")
            
            .Checkboxes = True
            .View = lvwReport
            
        End With
    End If
    
    ' Clear current LV collection
    lv_Tasks.ListItems.Clear
    
    ' Add current class collection
    If myList.count > 0 Then
    
        ' Sort the internal list
        Dim strSort As String
        strSort = "Subject"
        If sortingColumn = DueDate Then strSort = "DueDate"
        
        Dim collSort As New SortCollection
        collSort.Sort strSort, myList, sortingDirection
          
        ' Load the ListView
        For i = 0 To myList.count - 1
            Dim t As Outlook.TaskItem
            Set t = myList(i)
                                 
            ' Add Item to ListView
            AddListViewItem t, i + 1
            
        Next
        
        Call ResizeLVColumns
        
    End If
    
    strTrace = "Items: " & myList.count & " tasks..."
    Status (strTrace)
    
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
    
    strColumnWidths = "75;25"
    
    Dim totWidth As Integer
    totWidth = lv_Tasks.Width
    
    ' if scrollbar present, make space
    Dim bScrollbar As Boolean
    With lv_Tasks
        bScrollbar = (.font.SIZE + 4 + 1) * .ListItems.count > .Height
    End With
    
    If bScrollbar Then totWidth = totWidth - 15
    
    Dim widths() As String
    widths = Split(strColumnWidths, ";")
    
    For i = LBound(widths) To UBound(widths)
        Dim colWidth As Integer
        colWidth = CInt((widths(i) / 100) * totWidth) - 1
        lv_Tasks.ColumnHeaders(i + 1).Width = colWidth
    Next

ThrowException:
    LogMessage strTrace, strRoutine
    
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
    Set li = lv_Tasks.HitTest(X * factor, Y * factor)
        
    Set LV_GetItemAt = li
        
End Function

''' Search Supporting Methods

Private Sub SetSearchMode(ByVal b As Boolean)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SetSearchMode"
    
    On Error GoTo ThrowException
    
    Dim btnSearchLeft As Integer
    Dim btnNewMailLeft As Integer
    Dim txtbxSearchLeft As Integer
    Dim txtbxSearchWidth As Integer
    Dim btnSearchOffLeft As Integer
    
    f_searchMode = b
    
    If f_searchMode Then
    
        ' Set UI for search mode
        
        ' Hide unused buttons
        btn_Add.Visible = False
        btn_Edit.Visible = False
        btn_Delete.Visible = False
        btn_Options.Visible = False
        btn_Related.Visible = False
        
        ' Move used buttons
        btnSearchLeft = lv_Tasks.Left
        btnSearchOffLeft = lv_Tasks.Left + lv_Tasks.Width - btn_searchOff.Width
        btnNewMailLeft = lv_Tasks.Left + lv_Tasks.Width - btn_NewMail.Width
        
        btn_Projects.Left = btnSearchLeft
        btn_NewMail.Left = btnNewMailLeft
        btn_searchOff.Left = btnSearchOffLeft
        
        ' ReAnchor search button
        m_clsAnchors.Anchor("btn_Projects").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleBottom
        
        ' ReAnchor search cancel button
        m_clsAnchors.Anchor("btn_searchOff").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleBottom
                
        ' Resize search textBox
        txtbxSearchLeft = btnSearchLeft + btn_Projects.Width + 2
        txtbx_Search.Left = txtbxSearchLeft
        
        txtbxSearchWidth = btnSearchOffLeft - btnSearchLeft + btn_Projects.Width + 2
        ' txtbx_Search.Width = txtbxSearchWidth
        
        ' Clear current search contents
        txtbx_Search.text = ""
        
        ' Show search components
        txtbx_Search.Visible = True
        btn_searchOff.Visible = True
        
        ' Set focus to the search textbox
        txtbx_Search.SetFocus
        
    Else
        ' Remove search mode UI
        
        ' Hide search textBox
        txtbx_Search.Visible = False
        
        ' Hide search cancel button
        btn_searchOff.Visible = False
        
        ' Move used buttons back
        btnSearchLeft = lv_Tasks.Left + lv_Tasks.Width - btn_Projects.Width
        btnNewMailLeft = btnSearchLeft - btn_NewMail.Width - 5
        
        btn_Projects.Left = btnSearchLeft
        btn_NewMail.Left = btnNewMailLeft
        
        ' ReAnchor search button
        m_clsAnchors.Anchor("btn_Projects").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleBottom
        
        ' Unhide non-Search buttons
        btn_Add.Visible = True
        btn_Edit.Visible = True
        btn_Delete.Visible = True
        btn_Options.Visible = True
        btn_Related.Visible = True
   
    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

'''''' LISTBOX CODE

''' ListBox Handlers

Private Sub lstbx_Tasks_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    Cancel = True
    Effect = fmDropEffectCopy
End Sub

Private Sub lstbx_Tasks_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    Effect = fmDropEffectNone
    
    Dim strTrace As String
    strTrace = Data.GetText
    
    ' Process Dropped Object
    strTrace = ""
    
End Sub

Private Sub lstbx_Tasks_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Dim strTrace As String
    strTrace = ""
    Dim strRoutine As String
    strRoutine = rootClass & ":lstbx_Tasks_DblClick"
    
    Dim ut As Utilities
    Set ut = New Utilities
    
    Dim i As Integer
    i = lstbx_Tasks.ListIndex
    
    Dim eid As String
    eid = lstbx_Tasks.List(i, 3)
    Dim t As Outlook.TaskItem
    Set t = ut.GetOutlookItemFromID(eid)
    
    ' Prepare to present Outlook Task UI
    Set SelectedItem = t
    Set singleInspector = t.GetInspector
    
    ' Show the Task using the Outlook interface
    t.Display
    
End Sub

Private Sub lstbx_Tasks_AfterUpdate()

    Dim strTrace As String
    strTrace = ""
    Dim strRoutine As String
    strRoutine = rootClass & ":lstbx_Tasks_AfterUpdate"
    
    strTrace = "Item updated."
    LogMessage strTrace, strRoutine
    
End Sub

Private Sub lstbx_Tasks_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":lstbx_Tasks_MouseDown"
    
    Dim ut As Utilities
    Set ut = New Utilities

    If Button = 1 Then
        ' Selected an Item
        Dim i As Integer
        i = lstbx_Tasks.ListIndex
        
        ' Check to see if marking complete Column 0 width = 15
        If lstX < 15 Then
            Dim eid As String
            eid = lstbx_Tasks.List(i, 3)
            Dim t As Outlook.TaskItem
            Set t = ut.GetOutlookItemFromID(eid)
            
            ' Mark as complete / incomplete
            i = 1
        End If

    End If
    If Button = 2 Then
        ' Present ContextMenu
        Dim cMenu As ContextMenu
        ' Set cMenu = tm.GetContextMenu
        
        Dim id As Long
        id = ShowPopup(Me, cMenu, X, Y)
       
        strTrace = "Selected menu item: " & id
        LogMessage strTrace, strRoutine
        

        
    End If

End Sub

Private Sub lstbx_Tasks_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lstX = X
    lstY = Y
End Sub


Private Sub RefreshListBox()
    
    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":RefreshListBox"
    
    With lstbx_Tasks
        .Clear
        
        .ColumnCount = numColumns
        .ColumnWidths = strColumnWidths
        .ColumnHeads = True
        
    End With
    
    Dim chkbx As Long
    chkbx = &H2610
    Dim chkbx_checked As Long
    chkbx_checked = &H2611
    Dim chkbx_x As Long
    chkbx_x = &H2612
    
        
    If myList.count > 0 Then
       
        For i = 0 To myList.count - 1
            Dim t As Outlook.TaskItem
            Set t = myList(i)
            
            Dim checked As Long
            If t.Complete Then
                checked = chkbx_x
            Else
                checked = chkbx
            End If
            Dim strSubj As String
            strSubj = t.Subject
            Dim strDue As String
            strDue = t.DueDate
            If IsDateNone(t.DueDate) Then strDue = "None"
            
            With lstbx_Tasks
                .AddItem
                .List(i, 0) = ChrW(checked)
                .List(i, 1) = strSubj
                .List(i, 2) = strDue
                .List(i, 3) = t.EntryId
            End With

        Next
        
        strTrace = "Items: " & myList.count & " tasks..."
        Status (strTrace)

    Else
        
    
    End If
    
    
    Exit Sub
    
ThrowException:
    LogMessage strTrace, strRoutine
    
End Sub





