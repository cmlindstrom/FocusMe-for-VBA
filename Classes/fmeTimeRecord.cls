VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fmeTimeRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' - Fields

Private Const rootClass As String = "TimeRecord"

Private Const OneDayHours As Integer = 9

Private Enum enuTimeRecordType
    Appointment = 1
    Journal = 2
    Task = 3
End Enum

    Private f_Project As fmeProject

' - Events

' - Properties

    Private f_EntryID As String
    Private f_Categories As String
    Private f_RecordType As String
    Private f_RecordDate As Date
    Private f_Subject As String
    Private f_Duration As Integer ' in minutes
    Private f_CalendarFolder As Outlook.Folder
    
    Private Const fields As String = "EntryId,RecordType,RecordDate,Subject,Duration," & _
                                    "Categories,CalendarFolder"
    

''' Outlook Entry Id for the referenced Outlook Item
Public Property Let EntryId(ByVal id As String)
    f_EntryID = id
End Property
Public Property Get EntryId() As String
    EntryId = f_EntryID
End Property

''' Category assigned to the TimeRecord
Public Property Let Categories(ByVal cat As String)
    f_Categories = cat
End Property
Public Property Get Categories() As String
    Categories = f_Categories
End Property

''' RecordType - String
Public Property Let RecordType(ByVal typ As String)
    f_RecordType = typ
End Property
Public Property Get RecordType() As String
    RecordType = f_RecordType
End Property

''' RecordDate - Date
Public Property Let RecordDate(ByVal dte As Date)
    f_RecordDate = dte
End Property
Public Property Get RecordDate() As Date
    RecordDate = f_RecordDate
End Property

''' Subject - String
Public Property Let Subject(ByVal title As String)
    f_Subject = title
End Property
Public Property Get Subject() As String
    Subject = f_Subject
End Property

''' Duration (in minutes) - Integer
Public Property Let Duration(ByVal dur As Integer)
    f_Duration = dur
End Property
Public Property Get Duration() As Integer
    Duration = f_Duration
End Property

''' Calendar Folder for the TimeRecord
Public Property Set CalendarFolder(ByVal fldr As Outlook.Folder)
    Set f_CalendarFolder = fldr
End Property
Public Property Get CalendarFolder() As Outlook.Folder
    Set CalendarFolder = f_CalendarFolder
End Property

' - Constructor

Private Sub Class_Initialize()
    Set f_CalendarFolder = Nothing
    Set f_Project = Nothing
End Sub

Private Sub Class_Terminate()
    Set f_CalendarFolder = Nothing
    Set f_Project = Nothing
End Sub

' - Methods

''' Returns the first project, if exists
Public Function Project() As fmeProject
    Set Project = f_Project
End Function

''' Returns a collection of projects, if
''' more than one assigned to the TimeRecord
Public Function ProjectList() As ArrayList

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ProjectList"
    
    On Error GoTo ThrowException

    Dim pm As New ProjectManager
    pm.Load
       
    ' Capture list of assigned Projects
    Dim projList As New ArrayList
    Dim pList As ArrayList
    Set pList = pm.ParseProjectsFromString(f_Categories)
    If Not pList Is Nothing Then
        If pList.Count > 0 Then
                  
            ' Capture all projects
            Dim i As Integer
            Dim myP As fmeProject
            For i = 0 To pList.Count - 1
                Set myP = pList(i)
                projList.Add myP
                
                ' Grab the first project
                If i = 0 Then Set f_Project = myP
                
            Next
            
        End If
    End If
    
    Set ProjectList = projList
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set ProjectList = New ArrayList

End Function

''' Returns a list of assigned Categories to the TimeRecord
Public Function CategoryList() As ArrayList

    Dim myList As ArrayList
    Set myList = ParseCategories(f_Categories)
    If myList Is Nothing Then Set myList = New ArrayList

    Set CategoryList = myList
    
End Function

''' Returns the recorded Duration in Hours (vs Minutes)
Public Function DurationHours() As Double

    Dim dblReturn As Double
    dblReturn = 0
    
    dblReturn = f_Duration / 60
    
    DurationHours = dblReturn
    
End Function

''' Returns the recorded Duration accomodating the user
''' set Timecard Interval, e.g. report time in quarter hours
Public Function ReportedDurationHours() As Double

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ReportedDurationHours"
    
    On Error GoTo ThrowException

    Dim dblReturn As Double
    dblReturn = 0#
    Dim dblInc As Double
    dblInc = 0.25
    
    dblReturn = Math.RoundHours(DurationHours, dblInc)
    
    ReportedDurationHours = dblReturn
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    ReportedDurationHours = -1#

End Function

''' Returns the date associated with the end of the week
''' for the TimeRecord Date.
Public Function WeekEndingDate() As Date

    Dim dteReturn As Date
    dteReturn = GetDateOnly(LastDayOfTheWeek(f_RecordDate))

    WeekEndingDate = dteReturn

End Function

''' Fills the TimeRecord from an Outlook Appointment
Public Sub FillAppointment(ByVal oAppt As Outlook.AppointmentItem)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":FillAppointment"
    
    On Error GoTo ThrowException
    
    f_EntryID = oAppt.EntryId
    f_RecordDate = oAppt.Start
    f_Duration = oAppt.Duration
    f_RecordType = "Appointment"
    f_Subject = oAppt.Subject
    f_Categories = oAppt.Categories
    
    Set f_CalendarFolder = oAppt.Parent
    
    ' f_Mileage = CDbl(oAppt.Mileage)
    
    ' Make AllDay appointments one shift
    If oAppt.AllDayEvent Then f_Duration = oAppt.Duration / 3
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Fills the TimeRecord from an Outlook Recurrence Pattern
Public Sub FillRecurrencePattern(ByVal oAppt As Outlook.AppointmentItem, ByVal strtDate As Date)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":FillRecurrencePattern"
    
    On Error GoTo ThrowException

    f_EntryID = oAppt.EntryId
    f_RecordDate = strtDate
    f_Duration = oAppt.Duration
    f_RecordType = "Appointment"
    f_Subject = oAppt.Subject
    f_Categories = oAppt.Categories
    
    Set f_CalendarFolder = oAppt.Parent
    
    ' f_Mileage = CDbl(oAppt.Mileage)
    
    ' Make AllDay appointments one shift
    If oAppt.AllDayEvent Then f_Duration = oAppt.Duration / 3

    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Makes a copy of the current object instance
Public Function Clone() As fmeTimeRecord

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetListOfProperties"
    
    On Error GoTo ThrowException
    
    Dim retObj As New fmeTimeRecord
    
    Dim s As Variant
    Dim props() As String
    props = GetListOfProperties
    For Each s In props
        ' Ignore the CalendarFolder using this method
        If Not Contains("Calendar", s) Then
            Call SetProperty(retObj, s, GetProperty(Me, s))
        End If
    Next
    
    Set retObj.CalendarFolder = Me.CalendarFolder
    
    Set Clone = retObj
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set Clone = Nothing

End Function

' - Common Methods

''' Returns a string array of the object's property names
Public Function GetListOfProperties() As String()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetListOfProperties"
    
    On Error GoTo ThrowException

    Dim arReturn() As String
    
    arReturn = Split(fields, ",")
    
    GetListOfProperties = arReturn
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    GetListOfProperties = arReturn

End Function

' - Supporting Methods

''' Parse each Category from the Categories property
Private Function ParseCategories(ByVal strCat As String) As ArrayList

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ParseCategories"
    
    On Error GoTo ThrowException
   
    Dim retList As ArrayList
    Set retList = New ArrayList
    
    If Len(strCat) = 0 Then
        strTrace = "Encountered an empty Categories string."
        retList.Add "No Category"
        GoTo SkipOut
    End If
    
    Dim vCats() As String
    vCats = Split(strCat, ";")
    
    Dim i As Integer
    Dim myCat As String
    For i = LBound(vCats) To UBound(vCats)
        myCat = vCats(i)
        If Not Len(myCat) = 0 Then
            retList.Add myCat
        End If
    Next
    
SkipOut:
    Set ParseCategories = retList
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set ParseCategories = Nothing

End Function


