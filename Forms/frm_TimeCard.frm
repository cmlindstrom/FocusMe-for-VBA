VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_TimeCard 
   Caption         =   "FME VBA - Timecard"
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8370
   OleObjectBlob   =   "frm_TimeCard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_TimeCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Fields

Private Const rootClass As String = "frm_Timecard"

Dim WithEvents frm_PickStartDate As frm_DatePicker
Attribute frm_PickStartDate.VB_VarHelpID = -1
Dim WithEvents frm_PickEndDate As frm_DatePicker
Attribute frm_PickEndDate.VB_VarHelpID = -1

Dim stgs As Settings

Dim TimecardItems As TimeRecords

Dim bFirstRendered As Boolean

' Properties

Dim f_startDate As Date
Dim f_endDate As Date

''' Form Title
Public Property Let Title(ByVal strTitle As String)
    Me.Caption = strTitle
End Property
Public Property Get Title() As String
    Title = Me.Caption
End Property

''' Start date
Public Property Let startDate(ByVal dte As Date)
    f_startDate = dte
End Property
Public Property Get startDate() As Date
    startDate = f_startDate
End Property

''' End date
Public Property Let endDate(ByVal dte As Date)
    f_endDate = dte
End Property
Public Property Get endDate() As Date
    endDate = f_endDate
End Property

' Event Handlers

Private Sub btn_Refresh_Click()
    Refresh
End Sub

Private Sub btn_Done_Click()
    Unload Me
End Sub

Private Sub dtp_EndDate_Click()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":dtp_EndDate_Click"
    
    On Error GoTo ThrowException
    
    Dim myDate As Date
    
    Dim X As Integer
    Dim Y As Integer
    
    'If frm_PickEndDate Is Nothing Then
    'End If
    
    Set frm_PickEndDate = New frm_DatePicker
    With frm_PickEndDate
        .Caption = "Select End Date - Timecard"
        .SelectedDate = f_endDate
        .MultiSelect = False
            
        TryGetRelativePosition Me.txtbx_EndDate, X, Y
        .Top = Y
        .Left = X
            
    End With
    
    frm_PickEndDate.Show
    myDate = frm_PickEndDate.SelectedDate
    If myDate >= f_startDate Then
        f_endDate = myDate
        SetEndDate f_endDate
    Else
        ' Error - ignore selection
        strTrace = "The end date must be greater that the start date."
        MsgBox strTrace, vbOKOnly Or vbExclamation, Commands.AppName
    End If
   
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Private Sub dtp_StartDate_Click()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":dtp_StartDate_Click"
    
    On Error GoTo ThrowException

'    If frm_PickStartDate Is Nothing Then
'    End If
    
    Dim X As Integer
    Dim Y As Integer

    Set frm_PickStartDate = New frm_DatePicker
    With frm_PickStartDate
        .Caption = "Select Start Date - Timecard"
        .SelectedDate = f_startDate
        .MultiSelect = False
            
        TryGetRelativePosition Me.txtbx_StartDate, X, Y
        .Top = Y
        .Left = X
            
    End With
        
    frm_PickStartDate.Show
    f_startDate = frm_PickStartDate.SelectedDate
    SetStartDate f_startDate
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Private Sub chkbx_Include24Hr_Change()

    Dim strTrace As String
    
    If bFirstRendered Then
        strTrace = "Update setting..."
        stgs.TimecardInclude24hEvents = Me.chkbx_Include24Hr.value
        stgs.Save
    End If
    
End Sub

Private Sub chkbx_OnlyBusy_Change()

    Dim strTrace As String
    
    If bFirstRendered Then
        strTrace = "Update setting..."
        stgs.TimecardIncludeBusyOnly = Me.chkbx_OnlyBusy.value
        stgs.Save
    End If
    
End Sub


' Constructor

Private Sub UserForm_Initialize()

    f_startDate = DateAdd(GetDatePartFormat(DateInterval.Day), -7, Now)
    f_endDate = Now
    
    Set stgs = New Settings
    Set TimecardItems = New TimeRecords
    
    bFirstRendered = False
    Status
    
End Sub

Private Sub UserForm_Terminate()
    Set frm_PickStartDate = Nothing
    Set frm_PickEndDate = Nothing
    Set stgs = Nothing
End Sub

Private Sub UserForm_Activate()
    SetUI
End Sub

' Methods

''' Use the dates to refresh the timecard information
Public Sub Refresh()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Refresh"
    
    On Error GoTo ThrowException
    
    Status "Refreshing timecard records..."
    
    Call TimecardItems.Load(f_startDate, f_endDate, _
                                Me.chkbx_Include24Hr.value, _
                                Me.chkbx_OnlyBusy.value)
    
    strTrace = "Found " & TimecardItems.Count & " time records..."
    Status strTrace
    
    Dim dt As DataTable
    Set dt = TimecardItems.Analyze("Timecard", True, , True, True, False)
    
    RefreshListView dt
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

' Supporting Methods

Private Sub SetUI()

    SetStartDate f_startDate
    SetEndDate f_endDate
    
    Me.chkbx_Include24Hr.value = stgs.TimecardInclude24hEvents
    Me.chkbx_OnlyBusy.value = stgs.TimecardIncludeBusyOnly

    ' Helps to ignore first time UI setting
    bFirstRendered = True

End Sub

Private Sub SetStartDate(ByVal dte As Date)
    Me.txtbx_StartDate.text = Format(dte, "mm/dd/yyyy")
End Sub

Private Sub SetEndDate(ByVal dte As Date)
    Me.txtbx_EndDate.text = Format(dte, "mm/dd/yyyy")
End Sub

Private Sub Status(Optional ByVal msg As String = "")
    Me.sb_Status.SimpleText = msg
End Sub

Private Sub RefreshListView(ByVal dt As DataTable)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":RefreshListView"
    
    On Error GoTo ThrowException
    
    If dt.Columns.Count = 0 Then
        strTrace = "No columns found in the dataTable."
        GoTo ThrowException
    End If
    If dt.Rows.Count = 0 Then
        strTrace = "An empty table encountered."
        GoTo ThrowException
    End If
    
    Dim dc As DataColumn
    Dim i As Integer
    i = 1
    Dim j As Integer
    
    With lv_TimeCard
    
        ' Clear current Columns
        .ColumnHeaders.Clear
        
        ' Create Header Row - odd behavior, need to do this
        '   in reverse order to get listview columns to match
        '   table columns order
        For j = dt.Columns.Count - 1 To 0 Step -1
            Set dc = dt.Columns.Items(j)
            .ColumnHeaders.Add i, dc.Name, dc.Name
        Next
            
        ' Configure ListView
        .View = lvwReport
        .FullRowSelect = True
        .LabelEdit = lvwManual
        
    End With
    
    ' Clear current LV collection
    lv_TimeCard.ListItems.Clear
    
    ' Load ListView
    Dim dr As DataRow
    For i = 1 To dt.Rows.Count - 1
        Set dr = dt.Rows.Items(i)
        AddListViewItem dr
    Next
    
    ' Adjust columns size
    ResizeLVColumns
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Sub AddListViewItem(ByVal dr As DataRow, Optional ByVal idx As Integer = -1)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":AddListViewItem"
    
    If dr Is Nothing Then
        strTrace = "A null dataRow encountered."
        GoTo ThrowException
    End If
    
    On Error GoTo ThrowException

    ' Check the index
    If idx < 0 Then idx = lv_TimeCard.ListItems.Count + 1

    ' Add Item to ListView
    Dim id As String
    id = "LV_" & Common.GenerateUniqueID(4)
    Dim li As ListItem
    Dim rowName As String
    rowName = dr.GetItem("Name")
    Set li = lv_TimeCard.ListItems.Add(idx, id, rowName)
    
    Dim dc As DataColumn
    Dim i As Integer
    For i = 1 To dr.Parent.Table.Columns.Count - 1
        Set dc = dr.Parent.Table.Columns.Items(i)
        li.SubItems(i) = dr.GetItem(dc.Name)
    Next
           
    ' Format the row
    ' FormatLVRow li, t
    
    strTrace = "Added dataRow to ListView: " & rowName
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Sub ResizeLVColumns()

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":ResizeLVColumns"
    
    On Error GoTo ThrowException
    
    ' Get first column width
    Dim colWidth As Integer
    colWidth = GetMaxNameSize()
    lv_TimeCard.ColumnHeaders(1).Width = colWidth
    
    ' Adjust remaining columns to minimum widths
    Dim i As Integer
    With lv_TimeCard
        If .ColumnHeaders.Count > 1 Then
            For i = 2 To .ColumnHeaders.Count
                .ColumnHeaders(i).Width = MeasureString(.ColumnHeaders(i).text, .font)
            Next
        End If
    End With
    
    ' if vertical scrollbar present, make space
    Dim bScrollbar As Boolean
    With lv_TimeCard
        bScrollbar = (.font.SIZE + 4 + 1) * .ListItems.Count > .Height
    End With
    
    Exit Sub

ThrowException:
    LogMessage strTrace, strRoutine
    
End Sub
Private Function GetMaxNameSize() As Integer

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":GetMaxNameSize"
    
    Dim li As ListItem
    
    Dim lLen As Integer
    Dim iMax As Integer
    iMax = 0
    
    For i = 1 To lv_TimeCard.ListItems.Count
        Set li = lv_TimeCard.ListItems(i)
        ilen = MeasureString(li.text)
        If ilen > iMax Then iMax = ilen
    Next
    
    GetMaxNameSize = iMax
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    GetMaxNameSize = 0

End Function



Private Function TryGetRelativePosition(ByVal ctrl As control, _
                                         ByRef X As Integer, ByRef Y As Integer, _
                                    Optional ByVal sp As Integer = 0) As Boolean
                                         
    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":TryGetRelativePosition"
        
    On Error GoTo ThrowException
    
    Dim tX As Integer
    Dim tY As Integer
    Dim tH As Integer
    Dim tW As Integer
    
    ' Get center of application screen
    Dim appX As Integer
    Dim appY As Integer
    
    Dim titleBarWidth As Integer
    titleBarHeight = 23
    
'    With ThisOutlookSession.ActiveExplorer
'        tX = .Left
'        tY = .Top
'        tH = .Height
'        tW = .Width
'    End With
'
'    appX = tX + CInt(tW / 2)
'    appY = tY + CInt(tH / 2)
    
    ' frm_Timecard screen position
    tX = Me.Left
    tY = Me.Top
    
    ' Return position aligned to the left and under the specified control
    X = tX + ctrl.Left  '(Me.Width / 2)
    Y = tY + ctrl.Top + titleBarHeight + ctrl.Height ' (Me.Height / 2)
    
    '  Assume starts in center of application screen
    TryGetRelativePosition = True
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    TryGetRelativePosition = False
End Function

