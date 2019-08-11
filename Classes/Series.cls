VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Series"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Fields

Private Const rootClass As String = "Series"

Private Const startMaxValue As Double = -10 ^ 16

Public Enum enuSortOrder
    None = 0
    LabelAscending = 1
    LabelDescending = 2
    ValueAscending = 3
    ValueDescending = 4
End Enum

Dim f_sortOn As enuSortOrder

Public Enum enuFilterType
    None = 0
    Top10 = 1
    Bottom10 = 2
    Before = 3
    After = 4
    Range = 5
End Enum

Public Enum enuFilterFocus
    label = 1
    value = 2
End Enum

Public Enum enuOperation
    None = 0
    Count = 1
    Sum = 2
    Product = 3
    Minimum = 4
    Maximum = 5
    Average = 6
End Enum

Public Enum enuGroupBy
    None = 0
    Hour = 5
    Day = 1
    Week = 2
    Month = 3
    Year = 4
    Auto = 99
End Enum

' Properties

Dim f_name As String
Dim f_groupBy As enuGroupBy
Dim f_summBy As enuOperation

Private innerList As ArrayList

''' <summary>
''' Name of the Series.
''' </summary>
''' <value>String</value>
''' <returns>String</returns>
''' <remarks></remarks>
Public Property Let Name(ByVal nme As String)
    f_name = nme
End Property
Public Property Get Name() As String
    Name = f_name
End Property

''' Gets / Sets the GroupBy setting - if the series has been transformed
''' from another Series
Public Property Let LastGroupedBy(ByVal gb As enuGroupBy)
    f_groupBy = gb
End Property
Public Property Get LastGroupedBy() As enuGroupBy
    LastGroupedBy = f_groupBy
End Property

''' Gets / Sets the SummarizeBy setting - if the series has been transformed
''' from another Series
Public Property Let LastSummarizedBy(ByVal op As enuOperation)
    f_summBy = op
End Property
Public Property Get LastSummarizedBy() As enuOperation
    LastSummarizedBy = f_summBy
End Property

''' Returns the series collection
Public Property Get Items() As ArrayList
    Set Items = innerList
End Property

' Constructor

Private Sub Class_Initialize()
    f_name = ""
    Set innerList = New ArrayList
    f_sortOn = LabelAscending
    f_groupBy = enuGroupBy.None
    f_summBy = enuOperation.None
End Sub

' Method

''' Adds a new point to the series
Public Sub Add(ByVal sp As SeriesPoint)
    innerList.Add sp
End Sub

''' Adds a collection points to the series
Public Sub AddRange(ByVal rng As ArrayList)
    If Not rng Is Nothing Then
        If rng.Count > 0 Then
            Dim i As Integer
            For i = 0 To rng.Count - 1
                Add rng(i)
            Next
        End If
    End If
End Sub

''' Returns the SeriesPoint at the specified index
Public Function Point(ByVal idx As Integer) As SeriesPoint
    Set Point = innerList(idx)
End Function

''' Looks for a SeriesPoint matching the specified label, returns
''' the point via 'sp' as a reference parameter
Public Function TryGetPoint(ByVal label As Variant, ByRef sp As SeriesPoint) As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TryGetPoint"
    
    On Error GoTo ThrowException
    
    Dim bReturn As Boolean
    bReturn = False
    
    Dim spt As SeriesPoint
    For Each spt In innerList
        If spt.label = label Then
            bReturn = True
            Set sp = spt
        End If
    Next
    
    TryGetPoint = bReturn
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    TryGetPoint = False

End Function

''' Removes the specified Point from the collection
Public Sub Remove(ByVal sp As SeriesPoint)
    innerList.Remove sp
End Sub

''' Removes the SeriesPoint at the specified index
Public Sub RemoveAt(ByVal idx As Integer)
    innerList.RemoveAt idx
End Sub

''' Creates a copy of the Series
Public Function Clone() As Series

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Clone"
    
    On Error GoTo ThrowException
    
    Dim newSeries As New Series

    Dim i As Integer
    For i = 0 To innerList.Count - 1
        newSeries.Add innerList(i)
    Next
    
    Set Clone = newSeries
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set Clone = Nothing

End Function

''' Sorts the current Series collection
Public Sub Sort(ByVal order As enuSortOrder)
    InternalSort order
End Sub

''' <summary>
''' Returns the maximum value of the pt.value property in a SeriesPoint
''' </summary>
''' <returns>Double:</returns>
''' <remarks></remarks>
Public Function GetMaxValue() As Double

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetMaxValue"
    
    On Error GoTo ThrowException
    
    Dim maxValue As Double
    maxValue = startMaxValue
    
    Dim pt As SeriesPoint
    Dim i As Integer
    For i = 0 To innerList.Count - 1
        Set pt = innerList(i)
        If Not IsNothing(pt) Then
            If pt.value > maxValue Then maxValue = pt.value
        End If
    Next
    
    GetMaxValue = maxValue
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    GetMaxValue = -1

End Function

''' <summary>
''' Creates a new Series using the specified filters
''' </summary>
''' <param name="filterType">enuFilterType:</param>
''' <param name="filterFocus">enuFilterFocus:</param>
''' <param name="Before">Variant (Optional, only used for Before filter type):</param>
''' <param name="After">Variant (Optional, only used for the After filter type):</param>
''' <returns>Series:</returns>
Public Function Filter(ByVal filterType As enuFilterType, _
                       ByVal filterFocus As enuFilterFocus, _
              Optional ByVal Before As Variant, _
              Optional ByVal After As Variant) As Series
              
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Filter"
    
    On Error GoTo ThrowException
    
    ' Error checking
    If (filterType = enuFilterType.After Or filterType = enuFilterType.Before) And IsMissing(Before) Then
        strTrace = "Parameter not specified."
        GoTo ThrowException
    End If

    If (filterType = enuFilterType.Range) And (IsMissing(Before) Or IsMissing(After)) Then
        strTrace = "Parameters not specified."
        GoTo ThrowException
    End If
    
    strTrace = "Create an empty Series."
    Dim newSeries As New Series
    newSeries.Name = Me.Name
    
    strTrace = "Set up list in the right order."
    If filterType = Bottom10 Then
        If filterFocus = label Then
            Me.Sort LabelAscending
        Else
            Me.Sort ValueAscending
        End If
    Else
        If filterFocus = label Then
            Me.Sort LabelDescending
        Else
            Me.Sort ValueDescending
        End If
    End If
    
    strTrace = "Apply the conditions."
    Dim focus As Variant
    Dim condition As Boolean
    Dim pt As SeriesPoint
    Dim i As Integer
    For i = 0 To innerList.Count - 1
        Set pt = innerList(i)
        
        Select Case filterFocus
            Case enuFilterFocus.label
                focus = pt.label
            Case enuFilterFocus.value
                focus = pt.value
        End Select
        
        condition = False
        Select Case filterType
            Case enuFilterType.After
                condition = (focus >= After)
            Case enuFilterType.Before
                condition = (focus <= Before)
            Case enuFilterType.Bottom10
                condition = (i < 10)
            Case enuFilterType.Top10
                condition = (i < 10)
            Case enuFilterType.Range
                condition = (focus >= After And focus <= Before)
        End Select
        
        If condition Then
            newSeries.Add pt.Clone
            ' Important to add a 'copy' of the point, not the point itself
        End If
        
    Next
    
    Set Filter = newSeries
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set Filter = Nothing

End Function

''' <summary>
''' Returns the SeriesPoints that fall between the MinX and MaxX values.
''' </summary>
''' <param name="MinX">Object</param>
''' <param name="MaxX">Object</param>
''' <returns>Series</returns>
''' <remarks></remarks>
Public Function Trim(ByVal MinX As Variant, ByVal MaxX As Variant) As Series
        
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Trim"
    
    On Error GoTo ThrowException
    
    strTrace = "Create an empty Series."
    Dim newSeries As New Series
    newSeries.Name = Me.Name
    
    Dim pt As SeriesPoint
    Dim i As Integer
    For i = 0 To innerList.Count - 1
        Set pt = innerList(i)
        
        If pt.label >= MinX And pt.label <= MaxX Then
            newSeries.Add pt.Clone
        End If
    Next
        
    Set Trim = newSeries
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set Trim = Nothing
End Function

''' <summary>
''' Transforms the current collection of SeriesPoints and pivots according to the GroupBy and SummarizeBy parameters.
''' </summary>
''' <param name="GroupBy">enuGroupBy</param>
''' <param name="SummarizeBy">enuOperation</param>
''' <returns>Series</returns>
''' <remarks>"Pivots" the data, grouping by 'GroupBy' setting (e.g. Month, Week), and combines using
''' 'SummarizeBy' setting such as  adding or multiplying the values, or finding the min or max, etc.
''' </remarks>
Public Function Transform(ByVal groupBy As enuGroupBy, ByVal summarizeBy As enuOperation) As Series

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Transform"
    
    On Error GoTo ThrowException
    
    ' Error Checking
    If innerList.Count = 0 Then
        strTrace = "Current series is empty."
        GoTo ThrowException
    End If
    
    ' Common Variables
    Dim i As Integer
    Dim j As Integer
    Dim pt As SeriesPoint
    
    strTrace = "Create new series to return."
    Dim transformSeries As New Series
    transformSeries.Name = Me.Name
    
    ' Prepare current collection
    Me.Sort LabelAscending
    
    ' Set the defaults if not specified
    If groupBy = enuGroupBy.None Then groupBy = Auto
    If summarizeBy = enuOperation.None Then summarizeBy = Average
    
    ' Set the Properties of the transformed series
    transformSeries.LastGroupedBy = groupBy
    transformSeries.LastSummarizedBy = summarizeBy
    
    Dim startDate As Date
    startDate = #1/1/4501#
    Dim endDate As Date
    endDate = #1/1/1970#
    Dim currDate As Date
    Dim iDays As Long
    
    If groupBy = Auto Then
    
        ' Evalutate the range of dates to establish the start and stop dates
        For i = 0 To innerList.Count - 1
            Set pt = innerList(i)
            If Not pt Is Nothing Then
                currDate = CDate(pt.label)
                If currDate > endDate Then endDate = currDate
                If currDate < startDate Then startDate = currDate
            End If
        Next
        
        ' Set up appropriate GroupBy interval
        iDays = DateDiff("d", startDate, endDate)
        If iDays <= 10 Then groupBy = enuGroupBy.Day
        If iDays > 10 And iDays <= 70 Then groupBy = enuGroupBy.Week
        If iDays > 70 And iDays <= 365 Then groupBy = enuGroupBy.Month
        If iDays > 365 Then groupBy = enuGroupBy.Year
        
    End If
    
    ' Transform the Series collection
    Dim bFnd As Boolean
    Dim cpt As SeriesPoint
    For i = 0 To innerList.Count - 1
        ' Process each point
        Set pt = innerList(i)
        If Not pt Is Nothing Then
            bFnd = False
            For j = 0 To transformSeries.Items.Count - 1
                Set cpt = transformSeries.Items(j)
                If Not cpt Is Nothing Then
                    If IsConditionMet(pt, cpt, groupBy) Then
                        cpt.Samples = cpt.Samples + 1
                        cpt.value = ProcessValue(cpt.value, pt.value, summarizeBy)
                        transformSeries.Items(j) = cpt
                        bFnd = True
                        Exit For
                    End If
                Else
                    strTrace = "Transform SeriesPoint was null, idx = " & j & "."
                    LogMessage strTrace, strRoutine
                End If
            Next
            
            ' If point not 'bucketed' - create a new bucket
            If Not bFnd Then
                Dim newPoint As SeriesPoint
                Set newPoint = CreateNewBucket(pt, groupBy, summarizeBy)
                If Not newPoint Is Nothing Then transformSeries.Add newPoint
            End If
        
        Else
            strTrace = "Internal SeriesPoint was null, idx = " & i & "."
            LogMessage strTrace, strRoutine
        End If
        
    Next
    
    ' Sort the transformed series
    transformSeries.Sort LabelAscending
           
    Set Transform = transformSeries
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set Transform = Me

End Function
Private Function IsConditionMet(ByVal pt1 As SeriesPoint, ByVal pt2 As SeriesPoint, ByVal groupBy As enuGroupBy) As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":IsConditionMet"
    
    On Error GoTo ThrowException
    
    Dim bReturn As Boolean
    bReturn = False
    
    Dim currDate As Date
    currDate = CDate(pt1.label)
    Dim compDate As Date
    compDate = CDate(pt2.label)
    
    Select Case groupBy
        Case enuGroupBy.Hour
            bReturn = IsSameHour(currDate, compDate)
        Case enuGroupBy.Day
            bReturn = IsDateEqual(currDate, compDate)
        Case enuGroupBy.Week
            bReturn = IsSameWeek(currDate, compDate)
        Case enuGroupBy.Month
            bReturn = IsSameMonth(currDate, compDate)
        Case enuGroupBy.Year
            bReturn = (YearPart(currDate) = YearPart(compDate))
        Case enuGroupBy.None
            ' add every point
    End Select
        
    IsConditionMet = bReturn
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    IsConditionMet = False

End Function
Private Function CreateNewBucket(ByVal pt As SeriesPoint, _
                                 ByVal groupBy As enuGroupBy, _
                                 ByVal summBy As enuOperation) As SeriesPoint
    
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":IsConditionMet"
    
    On Error GoTo ThrowException
    
    Dim retPt As SeriesPoint
    Set retPt = Nothing
    
    '
    Dim dateOnly As Date
    Dim AddDate As Date
    
    ' Capture incoming Points Label - which should be a date
    Dim nDate As Date
    nDate = CDate(pt.label)

    ' Set Label
    Select Case groupBy
        Case enuGroupBy.Hour
            dateOnly = GetDateOnly(nDate)
            Dim hr As Integer
            hr = DatePart(GetDatePartFormat(Hour), nDate)

            AddDate = DateAdd(GetDatePartFormat(Hour), hr, dateOnly)
                       
        Case enuGroupBy.Day
            AddDate = GetDateOnly(nDate)

        Case enuGroupBy.Week
            AddDate = LastDayOfTheWeek(nDate)

        Case enuGroupBy.Month
            AddDate = DateSerial(YearPart(nDate), MonthPart(nDate), DaysInMonth(nDate))
            
        Case enuGroupBy.Year
            AddDate = DateSerial(YearPart(nDate), 1, 1)
    
        Case enuGroupBy.None
            AddDate = nDate
            
        Case Else
            strTrace = "Unrecognized GroupBy."
            GoTo ThrowException
    End Select
    
    ' Set Value
    Dim dbl As Double
    Select Case summBy
        Case enuOperation.Count
            dbl = 1#
        Case Else ' All else start with first point's value
            dbl = pt.value
    End Select
    
    Set retPt = New SeriesPoint
    retPt.Create AddDate, dbl

    Set CreateNewBucket = retPt
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set CreateNewBucket = Nothing
    
End Function

' Supporting Methods

''' <summary>
''' Performs the specified Operation given an incoming value and a focusValue.
''' </summary>
''' <param name="incomingValue">Double</param>
''' <param name="focusValue">Double</param>
''' <param name="Operation">enuOperation</param>
''' <returns>Double:</returns>
''' <remarks>e.g. if the operation is Sum, then if incomingValue = 10 and focusValue = 5
''' the the return value would be 10+5</remarks>
Private Function ProcessValue(ByVal incomingValue As Double, _
                              ByVal focusValue As Double, _
                              ByVal Operation As enuOperation, _
                     Optional ByVal Sample As Integer = 1) As Double
                            
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ProcessValue"
    
    On Error GoTo ThrowException
    
    Dim retValue As Double
    retValue = incomingValue
    
    Select Case Operation
        Case enuOperation.Average
            retValue = AverageBayes(incomingValue, focusValue, Sample)
        Case enuOperation.Count
            retValue = retValue + 1
        Case enuOperation.Sum
            retValue = retValue + focusValue
        Case enuOperation.Product
            retValue = retValue * focusValue
        Case enuOperation.Maximum
            If focusValue > retValue Then retValue = focusValue
        Case enuOperation.Minimum
            If focusValue < retValue Then retValue = focusValue
        Case Else
            retValue = incomingValue
    End Select
    
    ProcessValue = retValue
    Exit Function
                            
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    ProcessValue = incomingValue

End Function

Private Sub InternalSort(ByVal srt As enuSortOrder)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":InternalSort"
    
    Dim p As SeriesPoint
    Dim thisValue As String
    Dim thisItem As SeriesPoint
    Dim nextValue As String
    Dim nextItem As SeriesPoint
    Dim bPassDirty As Boolean
    bPassDirty = True
    
    Do While bPassDirty
    
        ' Assume clean pass
        bPassDirty = False
    
        ' Evalute items
        Dim i As Integer
        For i = 0 To innerList.Count - 1
            ' Get current value
            Set p = innerList(i)
            thisValue = GetIndexValue(p, srt)
            Set thisItem = p
            ' Get Next Value
            If Not i = innerList.Count - 1 Then
                Set p = innerList(i + 1)
                nextValue = GetIndexValue(p, srt)
                Set nextItem = p
            Else
                ' Done with the pass
                GoTo PassComplete
            End If
        
            ' Compare
            Dim bSwap As Boolean
            bSwap = False
            If srt = LabelAscending Or srt = ValueAscending Then bSwap = nextValue < thisValue
            If srt = LabelDescending Or srt = ValueDescending Then bSwap = nextValue > thisValue
            
            ' If sortDir = Ascending Then bSwap = nextValue < thisValue
            ' If sortDir = Descending Then bSwap = nextValue > thisValue
        
            ' Swap necessary?
            If bSwap Then
                ' Code Check
                Dim iCurrent As Integer
                iCurrent = innerList.Count
        
                ' Swapped two items, mark the pass as dirty
                ' - Remove nextItem
                innerList.Remove nextItem
                ' - Insert nextItem before thisItem
                innerList.Insert i, nextItem
                
                ' Mark the pass as dirty
                bPassDirty = True
                
                ' Code Check
                If Not iCurrent = innerList.Count Then
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
Private Function GetIndexValue(ByVal tr As SeriesPoint, ByVal sortOn As enuSortOrder) As String

' StartDate, Subject, Code, Type??, Calendar

    Dim strReturn As String
    
    Select Case sortOn
        Case enuSortOrder.LabelAscending
            If VarType(tr.label) = vbDate Then
                strReturn = Format(tr.label, "yyyymmdd")
            Else
                strReturn = CStr(tr.label)
            End If
            strReturn = tr.label
        Case enuSortOrder.LabelDescending
            If VarType(tr.label) = vbDate Then
                strReturn = Format(tr.label, "yyyymmdd")
            Else
                strReturn = CStr(tr.label)
            End If
            strReturn = tr.label
        Case enuSortOrder.ValueAscending
            strReturn = CStr(tr.value)
        Case enuSortOrder.ValueDescending
            strReturn = CStr(tr.value)
    End Select
    
    GetIndexValue = strReturn
    
End Function
