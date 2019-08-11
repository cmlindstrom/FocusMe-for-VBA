VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_TimeCard 
   Caption         =   "FME VBA - Timecard"
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8370.001
   OleObjectBlob   =   "frm_TimeCardOld.frx":0000
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

    If frm_PickEndDate Is Nothing Then
        Set frm_PickEndDate = New frm_DatePicker
        With frm_PickEndDate
            .SelectedDate = f_endDate
            .MultiSelect = False
        End With
    End If
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

    If frm_PickStartDate Is Nothing Then
        Set frm_PickStartDate = New frm_DatePicker
        With frm_PickStartDate
            .SelectedDate = f_startDate
            .MultiSelect = False
        End With
    End If
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

