VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_DatePicker 
   Caption         =   "UserForm1"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3360
   OleObjectBlob   =   "frm_DatePickerNew.frx":0000
End
Attribute VB_Name = "frm_DatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Fields

Private Const rootClass As String = "frm_DatePicker"

Dim dirty As Boolean
Private Const BackColorNormal As Long = &H8000000F
Private Const BackColorHighlight As Long = &HC0FFC0     '&H8000&

' Events

Public Event DateSeleted(ByVal control As Object, ByVal args As fmeEventArgs)

' Properties

Dim f_navDate As Date
Dim f_selectedDate As Date
Dim f_selection As ArrayList
Dim f_multiSelect As Boolean

''' Sets/Gets Selected Date
''' - can be used to set the initial view
Public Property Let SelectedDate(ByVal dte As Date)
    f_selectedDate = dte
End Property
Public Property Get SelectedDate() As Date
    SelectedDate = f_selectedDate
End Property

''' Gets selected dates collection
Public Property Get Selection() As Date()
    Selection = f_selection
End Property

''' Sets/Gets flag for selecting multiple dates
Public Property Let MultiSelect(ByVal b As Boolean)
    f_multiSelect = b
End Property
Public Property Get MultiSelect() As Boolean
    MultiSelect = f_multiSelect
End Property


' Event Handlers

Private Sub btn_Next_Click()
    f_navDate = DateAdd("m", 1, f_navDate)
    PaintCalendar f_navDate
End Sub

Private Sub btn_Previous_Click()
    f_navDate = DateAdd("m", -1, f_navDate)
    PaintCalendar f_navDate
End Sub

Private Sub btn_Set_Click()
    ' Returns the highlighted buttons in the Selection property
    Me.Hide
End Sub

Private Sub btn_Today_Click()
    f_navDate = Now
    PaintCalendar f_navDate
End Sub

Private Sub btn_None_Click()
    f_navDate = #1/1/1970#
End Sub

Private Sub btn_01_Click()
    ProcessSelection btn_01
End Sub

Private Sub btn_02_Click()
    ProcessSelection btn_02
End Sub

Private Sub btn_03_Click()
    ProcessSelection btn_03
End Sub

Private Sub btn_04_Click()
    ProcessSelection btn_04
End Sub

Private Sub btn_05_Click()
    ProcessSelection btn_05
End Sub

Private Sub btn_06_Click()
    ProcessSelection btn_06
End Sub

Private Sub btn_07_Click()
    ProcessSelection btn_07
End Sub

Private Sub btn_08_Click()
    ProcessSelection btn_08
End Sub

Private Sub btn_09_Click()
    ProcessSelection btn_09
End Sub

Private Sub btn_10_Click()
    ProcessSelection btn_10
End Sub

Private Sub btn_11_Click()
    ProcessSelection btn_11
End Sub

Private Sub btn_12_Click()
    ProcessSelection btn_12
End Sub

Private Sub btn_13_Click()
    ProcessSelection btn_13
End Sub

Private Sub btn_14_Click()
    ProcessSelection btn_14
End Sub

Private Sub btn_15_Click()
    ProcessSelection btn_15
End Sub

Private Sub btn_16_Click()
    ProcessSelection btn_16
End Sub

Private Sub btn_17_Click()
    ProcessSelection btn_17
End Sub

Private Sub btn_18_Click()
    ProcessSelection btn_18
End Sub

Private Sub btn_19_Click()
    ProcessSelection btn_19
End Sub

Private Sub btn_20_Click()
    ProcessSelection btn_20
End Sub

Private Sub btn_21_Click()
    ProcessSelection btn_21
End Sub

Private Sub btn_22_Click()
    ProcessSelection btn_22
End Sub

Private Sub btn_23_Click()
    ProcessSelection btn_23
End Sub

Private Sub btn_24_Click()
    ProcessSelection btn_24
End Sub

Private Sub btn_25_Click()
    ProcessSelection btn_25
End Sub

Private Sub btn_26_Click()
    ProcessSelection btn_26
End Sub

Private Sub btn_27_Click()
    ProcessSelection btn_27
End Sub

Private Sub btn_28_Click()
    ProcessSelection btn_28
End Sub

Private Sub btn_29_Click()
    ProcessSelection btn_29
End Sub

Private Sub btn_30_Click()
    ProcessSelection btn_30
End Sub

Private Sub btn_31_Click()
    ProcessSelection btn_31
End Sub

Private Sub btn_32_Click()
    ProcessSelection btn_32
End Sub

Private Sub btn_33_Click()
    ProcessSelection btn_33
End Sub

Private Sub btn_34_Click()
    ProcessSelection btn_34
End Sub

Private Sub btn_35_Click()
    ProcessSelection btn_35
End Sub

Private Sub btn_36_Click()
    ProcessSelection btn_36
End Sub

Private Sub btn_37_Click()
    ProcessSelection btn_37
End Sub

' Constructor

Private Sub UserForm_Initialize()
    f_multiSelect = True
    f_selectedDate = #1/1/1970#
    Set f_selection = New ArrayList
    dirty = False
End Sub

Private Sub UserForm_Terminate()
    Set f_selection = Nothing
End Sub

Private Sub UserForm_Activate()
    SetUI
End Sub

Private Sub UserForm_Deactivate()
    GetUI
End Sub

' Methods

' Supporting Methods

Private Sub SetUI()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SetUI"
    
    On Error GoTo ThrowException
    
    ' Configure for starter date
    If IsDateNone(f_selectedDate) Then
        f_navDate = DateSerial(2019, 5, 1)
    Else
        f_navDate = f_selectedDate
    End If
    
    ' Configure for multi selection
    If f_multiSelect Then
        btn_Set.Visible = True
    Else
        btn_Set.Visible = False
    End If
    
    ' Configure the calendar
    Call PaintCalendar(f_navDate)
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Sub GetUI()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SetUI"
    
    On Error GoTo ThrowException
    
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Sub ProcessSelection(ByVal btn As control)

    dirty = True
    
    Dim dte As Date
    dte = btn.Tag
    
    Dim ev As fmeEventArgs

    If f_multiSelect Then
        ' Decorate the button
        If btn.BackColor = BackColorHighlight Then
            btn.BackColor = BackColorNormal
            RemoveDate dte
        Else
            btn.BackColor = BackColorHighlight
            AddDate dte
            f_selectedDate = dte
        End If
        Set ev = New fmeEventArgs
        ev.Create f_selection, False
        RaiseEvent DateSeleted(Me, ev)
    Else
        ' Set selection properties
        f_selectedDate = dte
        Set ev = New fmeEventArgs
        ev.CreateVariant f_selectedDate, False
        RaiseEvent DateSeleted(Me, ev)
        ' Close the dialog, i.e. Hide
        Me.Hide
    End If
    
   ' Me.Repaint

End Sub

Private Sub PaintCalendar(ByVal dte As Date)

    ' Set the calendar title
    Me.lbl_Title.Caption = Format(dte, "mmmm - yyyy")

    ' Configure the buttons
    
    ' Reset the button view and tags
    Dim vars As Variant
    Dim c As control
    Dim i As Integer
    For Each c In Me.Controls
        Call ConfigureButton(c, "", "", False)
    Next
    
    ' Turn on the buttons for the specified month
     Call ConfigureButtons(dte)

    Me.Repaint

End Sub

Private Sub ConfigureButtons(ByVal startDate As Date)
                             
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ConfigureButtons"
    
    On Error GoTo ThrowException
    
    Dim i As Integer
    
    Dim dteFirst As Date
    dteFirst = DateSerial(YearPart(startDate), MonthPart(startDate), 1)
    Dim offset As Integer
    offset = DatePart(GetDatePartFormat(DayOfWeek), dteFirst)
    Dim monthDays As Integer
    monthDays = DaysInMonth(dteFirst)
    
    Dim btnNum As Integer
    Dim btnDate As Date
    Dim dy As Integer
    dy = 1
    Dim c As control
    For Each c In Me.Controls
        btnNum = IsDateButton(c)
        If btnNum > 0 Then
            ' On a Button for a date
            If btnNum >= offset And btnNum < offset + monthDays Then
                dy = btnNum - offset + 1
                btnDate = DateAdd(GetDatePartFormat(DateInterval.Day), dy - 1, dteFirst)
                Call ConfigureButton(c, dy, btnDate, True)
            End If
        End If
    Next
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Sub ConfigureButton(ByVal ctl As control, _
                            ByVal lbl As Variant, _
                            ByVal tg As Variant, _
                            ByVal vis As Boolean)
                
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ConfigureButton"
    
    If ctl Is Nothing Then
        strTrace = "A null control encountered."
        GoTo ThrowException
    End If
    
    On Error GoTo ThrowException
    
    If IsDateButton(ctl) > 0 Then
        ctl.Visible = vis
        ctl.Tag = tg
        ctl.Caption = lbl
    Else
        ' not a developer labeled button
    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Function IsDateButton(ByVal ctl As control) As Integer

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":IsDateButton"
    
    If ctl Is Nothing Then
        strTrace = "A null control encountered."
        GoTo ThrowException
    End If
    
    On Error GoTo ThrowException
    
    Dim iReturn As Integer
    iReturn = 0
    
    Dim vars As Variant
    Dim i As Integer
    If Contains("btn_", ctl.Name) Then
        vars = Split(ctl.Name, "_")
        i = CInt(vars(1))
        If i > 0 Then
            iReturn = i
        End If
    End If
    
    IsDateButton = iReturn
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    IsDateButton = 0

End Function

Private Sub AddDate(ByVal dte As Date)
    f_selection.Add dte
    
    ' LogMessage "records: " & f_selection.Count, "AddDate"
End Sub

Private Sub RemoveDate(ByVal dte As Date)
    f_selection.Remove dte
    
    ' LogMessage "records: " & f_selection.Count, "RemoveDate"
End Sub


