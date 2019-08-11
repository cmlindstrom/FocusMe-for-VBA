VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' - Fields

Private Const rootClass As String = "TimeRecords"

Dim f_sortOn As enuSortOn ' StartDate, Subject, Type??, Calendar
Dim sortDir As enuSortDirection
Dim myItems As ArrayList

Dim ut As Utilities
Dim stgs As Settings


' - Events

Public Event OnProgress()

' - Properties

''' Latest collection of Items
Public Property Set Items(ByVal ar As ArrayList)
    Set myItems = ar
End Property
Public Property Get Items() As ArrayList
    Set Items = myItems
End Property

''' Sort Property
Property Let SortBy(ByVal srt As enuSortOn)
    f_sortOn = srt
End Property
Property Get SortBy() As enuSortOn
    SortBy = f_sortOn
End Property

' - Constructor

Private Sub Class_Initialize()

    ' Initialize class variables
    Set myItems = New ArrayList
    f_sortOn = Subject
    sortDir = Ascending
    
    ' Connect to Outlook Session
    Set ut = New Utilities
    
    ' Allow user customizations
    Set stgs = New Settings

End Sub

Private Sub Class_Terminate()

    Set myItems = Nothing

    Set ut = Nothing
    Set stgs = Nothing

End Sub

' - Methods

Public Sub AddItem(ByVal tr As fmeTimeRecord)
    myItems.Add tr
    InternalSort f_sortOn, sortDir
End Sub

Public Sub RemoveItem(ByVal tr As fmeTimeRecord)
    myItems.Remove tr
End Sub

Public Sub RemoveAt(ByVal idx As Integer)
    myItems.RemoveAt idx
End Sub

Public Function Item(ByVal idx As Integer) As fmeTimeRecord
    Set Item = myItems(idx)
End Function

Public Function GetItemById(ByVal id As String) As fmeProject

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetItemById"

    Dim retItem As fmeTimeRecord
    Set retItem = Nothing

    Dim bFnd As Boolean
    bFnd = False
    Dim tr As fmeTimeRecord
    For Each tr In myItems
        If LCase(tr.EntryId) = LCase(id) Then
            Set retItem = tr
            bFnd = True
            Exit For
        End If
    Next
    
    Set GetItemById = retItem
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetItemById = Nothing

End Function

''' Sorts the current collection of TimeRecords
Public Sub Sort(ByVal srt As enuSortOn, Optional ByVal dir As enuSortDirection)
    
    If Not IsMissing(dir) Then sortDir = dir
    InternalSort srt, sortDir
   
End Sub

''' Returns the size of the current collection
Public Function Count() As Integer
    Count = myItems.Count
End Function

' - Collection

Public Sub Load(ByVal startDate As Date, _
                ByVal endDate As Date, _
        Optional ByVal include24hAppointments As Boolean = False, _
        Optional ByVal includeBusyStatusOnly As Boolean = True)
        
' OlBusyStatus
'   olBusy = 2
'   olFree = 0
'   olOutOfOffice = 3
'   olTentative = 1
'   olWorkingElsewhere = 4
        
' ("urn:schemas:calendar:dtstart" >= '3/16/2019 12:00 AM' AND
'   "urn:schemas:calendar:dtend" = '3/23/2019 12:00 AM' AND
'   "urn:schemas:calendar:alldayevent" = 0 AND
' "http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/82310003" = 0) ' Recurrence

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Load"
    
    On Error GoTo ThrowException
    
    Dim bAllFolders As Boolean
    bAllFolders = True

    ' Get the starting day
    Dim strStart As String
    strStart = GetDASLFormattedDate(startDate)
    
    ' To get the day at the end of the window
    Dim strEnd As String
    Dim dteEnd As Date
    dteEnd = DateAdd(GetDatePartFormat(DateInterval.day), 1, endDate)
    strEnd = GetDASLFormattedDate(dteEnd)
    
    ' Create a filter for the requested time window
    Dim strFilter As String
    strFilter = "@SQL=""" & PR_APPOINTMENT_STARTDATE & """ >= '" & strStart & "' AND " & _
                """" & PR_APPOINTMENT_ENDDATE & """ <= '" & strEnd & "'"
    
    strTrace = "Querying the time records using filter: " & strFilter
    LogMessage strTrace, strRoutine
    
    Dim fldrSet As ArrayList
    Set fldrSet = Nothing
    If bAllFolders Then Set fldrSet = ut.GetCalendarFolders
    
    Dim myList As ArrayList
    Set myList = ut.GetOutlookItems(strFilter, olAppointmentItem, fldrSet)
    If myList Is Nothing Then
        strTrace = "Outlook query failed."
        GoTo ThrowException
    End If
    If myList.Count = 0 Then
        strTrace = "No records returned for filter: " & strFilter
        GoTo ThrowException
    End If
    
    ' Clear current collection
    myItems.Clear
    
    Dim tr As fmeTimeRecord
    
    Dim oRecPattern As Outlook.RecurrencePattern
    
    Dim o As Object
    For Each o In myList
        On Error GoTo ThrowException
        If TypeOf o Is Outlook.AppointmentItem Then
            strTrace = "Cast the ApppointmentItem."
            Dim oAppt As Outlook.AppointmentItem
            Set oAppt = o
            
            Dim bIsException As Boolean
            Dim bIsRecurring As Boolean
            bIsRecurring = oAppt.IsRecurring
            strTrace = "Processing appointment: '" & oAppt.Subject & "', starts: " & oAppt.Start & _
                            ", Recurring: " & bIsRecurring
            LogMessage strTrace, strRoutine
            
            ' Check if need to include busy status only, and if so,
            ' check to make sure appt is marked as busy
            If includeBusyStatusOnly Then
                If oAppt.BusyStatus <> Outlook.OlBusyStatus.olBusy Then GoTo SkipItem
            End If
            
            ' Check to see if need to remove 24 hr appointments from query result
            If Not include24hAppointments Then
                If oAppt.AllDayEvent Then GoTo SkipItem
            End If
            
            strTrace = "Creating new time record for appointment: '" & oAppt.Subject & "'."
            If bIsRecurring Then
                On Error GoTo SkipItem
                ' Check to see if can retrieve the recurrence pattern, may fail.
                Set oRecPattern = oAppt.GetRecurrencePattern
                ' Okay to proceed, all other properties will return properly
                strTrace = "Creating new time record: '" & oAppt.Subject & _
                                "', starts: " & oRecPattern.StartTime
                LogMessage strTrace, strRoutine
                
                ' Check for recurrence
                Dim myRec As fmeRecurrencePattern
                Set myRec = New fmeRecurrencePattern
                myRec.FillFromOutlook oRecPattern
                Set myRec.Parent = oAppt
                
                Dim dteTemp As Date
                Dim arList As ArrayList
                Set arList = myRec.GetAllRecurrences(startDate, endDate)
                If arList.Count > 0 Then
                    Dim dte As Variant
                    For Each dte In arList
                    
                        ' Check for exceptions
                        bIsException = False
                        If oRecPattern.Exceptions.Count > 0 Then
                            Dim exp As Outlook.Exception
                            For Each exp In oRecPattern.Exceptions
                                If IsDateEqual(dte, exp.OriginalDate) Then
                                    bIsException = True
                                    Exit For
                                End If
                            Next
                        End If
                        
                        If Not bIsException Then
                            ' Add Recurrence
                            dteTemp = dte
                            Set tr = New fmeTimeRecord
                            tr.FillRecurrencePattern oAppt, dteTemp
                        
                            ' Update the internal collection
                            myItems.Add tr
                        End If
                        
                    Next
                End If
                
'                dteTemp = myRec.GetNextRecurrence(startDate, endDate)
'                If Not IsDateNone(dteTemp) Then
'                    Set tr = New fmeTimeRecord
'                    tr.FillRecurrencePattern oAppt, dteTemp
'                Else
'                    ' Recurrence not in period window
'                    GoTo SkipItem
'                End If
        
            Else
                strTrace = "Creating new time record: '" & oAppt.Subject & _
                                "', starts: " & oAppt.Start
                LogMessage strTrace, strRoutine
                
                Set tr = New fmeTimeRecord
                tr.FillAppointment oAppt
                
                ' Update the internal collection
                myItems.Add tr
                
            End If
                                    
        Else
            strTrace = "Encountered an invalid appointment item, filter: '" & strFilter & "'."
            LogMessage strTrace, strRoutine
        End If
SkipItem:
    Next
    
    strTrace = "Loaded " & myItems.Count & " records..."
    LogMessage strTrace, strRoutine
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Exports the current collection to a CSV file
Public Sub Export(ByVal strFullFilePath As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Export"
    
    On Error GoTo ThrowException
    
    strTrace = "Error checking."
    If Len(strFullFilePath) = 0 Then
        strTrace = "Export filename not specified."
        GoTo ThrowException
    End If
    If myItems.Count = 0 Then
        strTrace = "Internal collection is empty."
        GoTo ThrowException
    End If
    
    ' Method Variables
    Dim i As Integer
    Dim tr As fmeTimeRecord
    Dim s As Variant
    
    strTrace = "Creating .csv formatted string for export."
    Dim strCSV As String
    Dim strLine As String
    
    ' Create header row
    Set tr = New fmeTimeRecord
    Dim props() As String
    props = tr.GetListOfProperties()
    For Each s In props
        strLine = strLine & s & ","
    Next
    ' Add pivotTable field headers
    strLine = strLine & "Reported_Duration,Date,"
    
    ' Remove trailing comma
    strLine = Left(strLine, Len(strLine) - 1)
    
    ' Add Header
    strCSV = strCSV & strLine & vbCrLf
    
    ' Create table body
    For i = 0 To myItems.Count - 1
        ' Reset Line
        strLine = ""
        ' Iterate thru the TimeRecord
        Set tr = myItems(i)
        For Each s In props
            strTrace = "Process field: " & s & "."
            strLine = strLine & CStr(GetProperty(tr, s)) & ","
        Next
        ' Add pivotTable fields
        strLine = strLine & CStr(tr.DurationHours) & "," & GetDateOnly(tr.RecordDate) & ","
        
        ' Remove trailing comma
        strLine = Left(strLine, Len(strLine) - 1)
    
        ' Add Row
        strCSV = strCSV & strLine & vbCrLf
        
    Next
    
    ' Save to Text File
    Call Common.WriteTextFile(strFullFilePath, strCSV)

    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Pivots the current timeRecord collection according to Row/Column specifiers
Public Function Analyze(Optional ByVal tblName As String = "Default", _
                        Optional ByVal includeRowName As Boolean = False, _
                        Optional ByVal RowNameHeaderLabel As String = "Name", _
                        Optional ByVal includeHeaderRow As Boolean = False, _
                        Optional ByVal includeTotalsRow As Boolean = False, _
                        Optional ByVal includeTotalsColumn As Boolean = False, _
                        Optional ByVal TotalsColumnHeaderLabel As String = "Totals") As DataTable

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":Analyze"
    
    On Error GoTo ThrowException
    
    ' Return Table
    Dim dt As DataTable
    Set dt = Nothing
    
    ' Common variables
    Dim tr As fmeTimeRecord
    Dim s As Variant
    
    strTrace = "Flatten the TimeRecords, 1 per Category."
    Dim rawItems As New ArrayList
    
    For Each tr In myItems
        For Each s In tr.CategoryList
            ' Add item to raw collection
            Dim dup As fmeTimeRecord
            Set dup = tr.Clone
            If Not dup Is Nothing Then
                dup.Categories = s
                rawItems.Add dup
            End If
        Next
    Next
    
    strTrace = "Create a collection of Series, 1 per category."
    Dim SeriesCollection As ArrayList
    Set SeriesCollection = SerializeByCategory()
    If SeriesCollection Is Nothing Then
        strTrace = "Failed to serialize the records by Category."
        GoTo ThrowException
    End If
 
    strTrace = "GroupBy Day."
    Dim ser As Series
    Dim tSer As Series
    Dim TransformedSeriesCollection As New ArrayList
    For Each ser In SeriesCollection
        Set tSer = ser.Transform(enuGroupBy.day, enuOperation.Sum)
        TransformedSeriesCollection.Add tSer
    Next
    
    strTrace = "Create the Table."
    Set dt = ConvertSeriesCollectionToTable(tblName, _
                                            TransformedSeriesCollection, _
                                            includeRowName, _
                                            RowNameHeaderLabel, _
                                            includeHeaderRow, _
                                            includeTotalsRow, _
                                            includeTotalsColumn, _
                                            TotalsColumnHeaderLabel)
    
    ' Return Table
    Set Analyze = dt
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set Analyze = Nothing

End Function

' - Supporting Methods

Private Sub InternalSort(ByVal srt As enuSortOn, ByVal dir As enuSortDirection)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":InternalSort"
    
    On Error GoTo ThrowException
    
    Dim p As fmeTimeRecord
    Dim thisValue As String
    Dim thisItem As fmeTimeRecord
    Dim nextValue As String
    Dim nextItem As fmeTimeRecord
    Dim bPassDirty As Boolean
    bPassDirty = True
    
    Do While bPassDirty
    
        ' Assume clean pass
        bPassDirty = False
    
        ' Evalute items
        Dim i As Integer
        For i = 0 To myItems.Count - 1
            ' Get current value
            Set p = myItems(i)
            thisValue = GetIndexValue(p)
            Set thisItem = p
            ' Get Next Value
            If Not i = myItems.Count - 1 Then
                Set p = myItems(i + 1)
                nextValue = GetIndexValue(p)
                Set nextItem = p
            Else
                ' Done with the pass
                GoTo PassComplete
            End If
        
            ' Compare
            Dim bSwap As Boolean
            bSwap = False
            If sortDir = Ascending Then bSwap = nextValue < thisValue
            If sortDir = Descending Then bSwap = nextValue > thisValue
        
            ' Swap necessary?
            If bSwap Then
                ' Code Check
                Dim iCurrent As Integer
                iCurrent = myItems.Count
        
                ' Swapped two items, mark the pass as dirty
                ' - Remove nextItem
                myItems.Remove nextItem
                ' - Insert nextItem before thisItem
                myItems.Insert i, nextItem
                
                ' Mark the pass as dirty
                bPassDirty = True
                
                ' Code Check
                If Not iCurrent = myItems.Count Then
                    strTrace = "Elongating the array."
                End If
            End If
        
        Next
PassComplete:
    
    Loop
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub
Private Function GetIndexValue(ByVal tr As fmeTimeRecord) As String

' StartDate, Subject, Code, Type??, Calendar

    Dim strReturn As String
    
    Select Case f_sortOn
        Case enuSortOn.Subject
            strReturn = tr.Subject
        Case enuSortOn.Calendar
            strReturn = tr.CalendarFolder.name
        Case enuSortOn.startDate
            strReturn = Format(tr.RecordDate, "yyyymmdd")
    End Select
    
    GetIndexValue = strReturn
    
End Function

Private Function GetDASLFormattedDate(ByVal dteIn As Date) As String

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":GetDASLFormattedDate"
    
    On Error GoTo ThrowException
    
    Dim strReturn As String

    ' Month
    Dim m As Integer
    m = DatePart("m", dteIn)
    
    ' Day
    Dim d As Integer
    d = DatePart("d", dteIn)
    
    ' Year
    Dim y As Integer
    y = DatePart("yyyy", dteIn)
    
    ' Hour
    Dim h As Integer
    h = DatePart("h", dteIn)
    
    ' Minute
    Dim n As Integer
    n = DatePart("n", dteIn)
    Dim strMin As String
    If n < 10 Then
        strMin = "0" & CStr(n)
    Else
        strMin = CStr(n)
    End If

    strReturn = Format(dteIn, "m/d/yyyy")
    If h > 12 Then
        h = h - 12
        strReturn = strReturn & " " & h & ":" & strMin & " PM"
    Else
        strReturn = strReturn & " " & h & ":" & strMin & " AM"
    End If
    
    strTrace = "Date Format: " & strReturn
    LogMessage strTrace, strRoutine

    GetDASLFormattedDate = strReturn
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    GetDASLFormattedDate = ""
    
End Function

''' Creates a SeriesCollection for the TimeRecords by Category
Private Function SerializeByCategory() As ArrayList

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":SerializeByCategory"
    
    On Error GoTo ThrowException
    
    Dim SeriesCollection As New ArrayList
    
    ' Common variables
    Dim tr As fmeTimeRecord
    Dim s As Variant
            
    ' Create a series for every category
    Dim recCategories As ArrayList
    Dim bFnd As Boolean
    Dim ser As Series
    Dim focusSeries As Series
    Dim sp As SeriesPoint
    For Each tr In myItems
        Set recCategories = tr.CategoryList
        For Each s In recCategories
            
            ' Check to see if Series exists for this category
            bFnd = False
            For Each ser In SeriesCollection
                If ser.name = s Then
                    Set focusSeries = ser
                    bFnd = True
                    Exit For
                End If
            Next
            
            If Not bFnd Then
                ' Create a new series
                Set focusSeries = New Series
                focusSeries.name = s
                
                ' Add it to the collection
                SeriesCollection.Add focusSeries
            End If
            
            ' Add this timeRecord to the series
            Set sp = New SeriesPoint
            Call sp.Create(GetDateOnly(tr.RecordDate), tr.DurationHours)
            focusSeries.Add sp
            
        Next
    Next
    
    ' Sort collection by category name
    Dim sColl As New SortCollection
    sColl.Sort "Name", SeriesCollection
       
    Set SerializeByCategory = sColl.List ' SeriesCollection
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set SerializeByCategory = Nothing

End Function

Private Function ConvertSeriesCollectionToTable(ByVal tblName As String, _
                                                ByVal sColl As ArrayList, _
                                       Optional ByVal includeRowName As Boolean = False, _
                                       Optional ByVal RowNameHeaderLabel As String = "Name", _
                                       Optional ByVal includeHeaderRow As Boolean = False, _
                                       Optional ByVal includeTotalsRow As Boolean = False, _
                                       Optional ByVal includeTotalsColumn As Boolean = False, _
                                       Optional ByVal TotalsColumnHeaderLabel As String = "Totals") As DataTable
                                                
    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":ConvertSeriesCollectionToTable"
    
    On Error GoTo ThrowException

    strTrace = "Create table to return."
    Dim dt As New DataTable
    dt.name = tblName
    
    ' Common Variables
    Dim dr As DataRow
    Dim cl As DataColumn
    Dim ser As Series
    Dim sp As SeriesPoint
    Dim i As Integer
    
    Dim tGroupBy As enuGroupBy
    tGroupBy = enuGroupBy.None
    Dim tSummBy As enuOperation
    tSummBy = enuOperation.None
    
    ' Find Min and Max dates from the data & error check
    Dim sDate As Date
    sDate = #1/1/4501#
    Dim eDate As Date
    eDate = #1/1/1970#
    Dim tDate As Date
    
    i = 0
    For Each ser In sColl
        i = i + 1
        ' Make sure every series is grouped and summarized the same
        If i = 1 Then
            If ser.LastGroupedBy = enuGroupBy.None Then
                strTrace = "FATAL ERROR: Unable to create columns in the table, no GroupBy property set."
                GoTo ThrowException
            End If
            tGroupBy = ser.LastGroupedBy
        Else
           If Not ser.LastGroupedBy = tGroupBy Then
                ' Mismatched series
                strTrace = "Series collection are not all grouped by the same setting."
                GoTo ThrowException
            End If
        End If

'        If Not ser.LastSummarizedBy = enuOperation.None Then
'            If Not ser.LastSummarizedBy = tSummBy Then
'                ' Mismatched series
'                strTrace = "Series collection are not all summarized by the same setting."
'                GoTo ThrowException
'            End If
'        Else
'            If i = 1 Then tSummBy = ser.LastSummarizedBy
'        End If
    
        ' Find the min and max dates
        For Each sp In ser.Items
            tDate = CDate(sp.label)
            If tDate < sDate Then sDate = tDate
            If tDate > eDate Then eDate = tDate
        Next
    Next
    strTrace = "Data traverses from " & sDate & " to " & eDate
    LogMessage strTrace, strRoutine
    
    ' Create and sort an array of contiguous periods
    Dim colLabels As New ArrayList
    tDate = sDate
    Do While tDate <= eDate
    
        strTrace = "Adding a column for date: " & tDate
        
        ' Add the date as a column
        colLabels.Add tDate
        
        ' Set the next date
        Select Case tGroupBy
            Case enuGroupBy.day
                tDate = DateAdd(GetDatePartFormat(DateInterval.day), 1, tDate)
            Case enuGroupBy.Week
                tDate = DateAdd(GetDatePartFormat(DateInterval.day), 7, tDate)
            Case enuGroupBy.Month
                tDate = DateAdd(GetDatePartFormat(DateInterval.Month), 1, tDate)
            Case Else
                 tDate = DateAdd(GetDatePartFormat(DateInterval.day), 1, tDate)
        End Select
        
    Loop
    
    ' Add a RowName column if flag set
    Dim col As DataColumn
    If includeRowName Then
        Set col = dt.Columns.Add(RowNameHeaderLabel)
    End If
    
    ' Add a column for each date to the table
    For i = 0 To colLabels.Count - 1
        Set col = dt.Columns.Add(CStr(colLabels(i)))
    Next
    
    ' Add a Totals column if flag is set
    If includeTotalsColumn Then
        Set col = dt.Columns.Add(TotalsColumnHeaderLabel)
    End If
    
    ' Create a Header Row if required
    '   - can use to format each column header while
    '     using the date as the column 'name'
    Dim hdrRow As DataRow
    If includeHeaderRow Then
        Set hdrRow = New DataRow
        For Each col In dt.Columns.Items
            hdrRow.Add col.name, col.name
        Next
        dt.rows.AddRow hdrRow
    End If
    
    strTrace = "Datatable columns created: " & dt.Columns.Count & "."
    
    ' Create a Totals Row - initialized to zero
    Dim totRow As New DataRow
    Dim total As Double
    total = 0#
    Dim grandTotal As Double
    grandTotal = 0#
    If includeTotalsRow Then
        If includeRowName Then
            totRow.Add RowNameHeaderLabel, "------ Totals"
        End If
        For Each col In dt.Columns.Items
            If col.name <> RowNameHeaderLabel Then
                totRow.Add col.name, 0
            End If
        Next
    End If
    
    strTrace = "Totals row has been initialized."
    
    ' Populate the body of the table
    Dim colName As String
    For Each ser In sColl
        strTrace = "Creating a row for series: " & ser.name
        ' Create a Row for each series in the collection
        '   with totals for each date
        Set dr = New DataRow
        
        Dim iColStart As Integer
        iColStart = 0
        If includeRowName Then
            dr.Add RowNameHeaderLabel, ser.name
            iColStart = 1
        End If
        
        ' Accomodate Totals Column
        Dim cDiff As Integer
        cDiff = 1
        If includeTotalsColumn Then cDiff = 2
        Dim rowTotal As Double
        rowTotal = 0#
        
        ' Iterate thru the columns adding matching points, and
        '   adding zero points where no values are found
        Dim dte As Date
        For i = iColStart To dt.Columns.Count - cDiff
            colName = dt.Columns.Item(i).name
            dte = CDate(colName)
            If ser.TryGetPoint(dte, sp) Then
                ' Populate the Point value
                dr.Add colName, sp.value
                
                ' Update row total for TotalsColumn
                rowTotal = rowTotal + sp.value
            
                ' Update the totals row as you go
                total = totRow.GetItem(colName)
                total = total + sp.value
                totRow.SetItem colName, total
            
            Else
                ' Populate a zero value
                dr.Add colName, 0
            End If
        Next
        
        ' Sum Row values if set
        If includeTotalsColumn Then
            dr.Add TotalsColumnHeaderLabel, rowTotal
        End If
        
        ' Track a grand total for all values
        grandTotal = grandTotal + rowTotal
               
        ' Add the Row to the Table
        dt.rows.AddRow dr
    
    Next
    
    strTrace = "Table body has been created, rows: " & dt.rows.Count & "."
    
    ' Add the totals row
    If includeTotalsRow Then
        If includeTotalsColumn Then
            totRow.SetItem TotalsColumnHeaderLabel, grandTotal
        End If
        
        strTrace = "Totals row added to the table, returning the DataTable."
        dt.rows.AddRow totRow
    End If
        
    Set ConvertSeriesCollectionToTable = dt
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set ConvertSeriesCollectionToTable = Nothing

End Function
