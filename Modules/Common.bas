Attribute VB_Name = "Common"
Private Const rootClass As String = "Common"

Private Const traceLogFileName As String = "TraceLog.txt"
Private Const errorLogFileName As String = "SystemLog.txt"

''' OLE Colors
Public Enum enuColors
    Black = 0
    Blue = 1
    Green = 2
    Cyan = 3
    Red = 4
    Magenta = 5
    Yellow = 6
    White = 7
End Enum

'''  Used to add resizing to forms
Public Enum enumAnchorStyles
    enumAnchorStyleNone = 0
    enumAnchorStyleTop = 1
    enumAnchorStyleBottom = 2
    enumAnchorStyleLeft = 4
    enumAnchorStyleRight = 8
End Enum

''' Shadows VBA DateInteral Strings
Public Enum DateInterval
    None = 0
    Year = 1        ' yyyy
    Quarter = 2     ' q
    Month = 3       ' m
    DayOfYear = 4   ' y
    day = 5         ' d
    DayOfWeek = 6   ' w
    Week = 7        ' ww
    Hour = 8        ' h
    Minute = 9      ' n
    Second = 10     ' s
End Enum

''' Works with Common.WeekDay
Public Enum enuDayOfWeek
    None = -1
    Sunday = 1
    Monday = 2
    Tuesday = 3
    Wednesday = 4
    Thursday = 5
    Friday = 6
    Saturday = 7
End Enum

Public Type Point
   X As Long
   y As Long
End Type

''' Function returns True if obj is null or nothing
Public Function IsNothing(obj As Object) As Boolean

    Dim bReturn As Boolean
    bReturn = False
    
    If obj Is Nothing Then
        bReturn = True
    Else
        bReturn = False
    End If
    
    IsNothing = bReturn

End Function

' NUMERIC FUNCTIONS

''' Returns the portion of a real number that is on the left
''' of the decimal place
Public Function WholeNumber(ByVal dbl As Double) As Integer

    Dim iReturn As Integer
    iReturn = -1
    
    Dim strNum As String
    strNum = Format(dbl, "Standard")
    
    Dim parts() As String
    parts = Split(strNum, ".")
    
    iReturn = CInt(parts(LBound(parts)))

    WholeNumber = iReturn

End Function

''' Returns the portion of a real number that is on the right
''' of the decimal place
Public Function Fraction(ByVal dbl As Double) As Double

    Dim dblReturn As Double
    dblReturn = -1
    
    Dim strNum As String
    strNum = Format(dbl, "Standard")
    
    Dim parts() As String
    parts = Split(strNum, ".")
    
    Dim prefix As String
    prefix = "0."
    
    Dim rgt As String
    rgt = parts(UBound(parts))
    If Left(rgt, 1) = "0" Then prefix = "0.0"
    
    Dim i As Integer
    i = CInt(parts(UBound(parts)))
    dblReturn = CDbl(prefix & i)

    Fraction = dblReturn

End Function

' STRING FUNCTIONS

''' Function returns True if strSearch found in strTarget
Public Function Contains(ByVal searchFor As String, ByVal inString As String) As Boolean

    Dim bReturn As Boolean
    bReturn = False

    If Len(searchFor) = 0 Then
        bReturn = True
    Else
      If InStr(1, inString, searchFor) > 0 Then
        bReturn = True
      End If
    End If
    
    Contains = bReturn

End Function

''' <summary>
''' Generates a unique ID with a given length.
''' </summary>
''' <param name="iLengthOfString">Integer: # of characters in the unique ID</param>
''' <returns>String</returns>
''' <remarks></remarks>
Public Function GenerateUniqueID(ByVal iLengthOfString As Integer) As String
    ' GenerateUniqueID = RandomString(iLengthOfString)
    GenerateUniqueID = RandomGuid(iLengthOfString) ' More reliably random
End Function

''' <summary>
''' Creates a random string that has a length as specified in the Length parameter.  Possible
''' ASCII characters range from 65 to 101
''' </summary>
''' <param name="Length">Integer</param>
''' <returns>String</returns>
''' <remarks></remarks>
Public Function RandomString(ByVal Length As Integer) As String

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":RandomString"

    Dim strReturn As String
    
    ' by making Generator static, we preserve the same instance
    ' (i.e., do not create new instances with the same seed over and over
    ' since the 'time based' nature of the random number generator

    ' Static Generator As New Random
    Randomize
    
    Dim upperbound As Integer
    upperbound = 99
    Dim lowerbound As Integer
    lowerbound = 65

    Dim charOutput() As Char
    ReDim charOutput(Length - 1)
    
    Dim selector As Integer

    For i = 0 To Length - 1

        selector = Int((upperbound - lowerbound + 1) * rnd + lowerbound)
        If selector > 90 Then
            selector = selector - 42
        End If

        strReturn = strReturn & Chr(selector)
    Next
   
    RandomString = strReturn

End Function

''' <summary>
''' Creates a random string that has a length as specified in the Length parameter.
''' </summary>
''' <param name="Length">Integer</param>
''' <returns>String</returns>
''' <remarks>https://stackoverflow.com/questions/45332357/ms-access-vba-error-run-time-error-70-permission-denied</remarks>
Public Function RandomGuid(ByVal iLength As Integer) As String

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":RandomGuid"
    
    Dim guid As String
    guid = CreateGuidString ' see OLE32 libary functions in WinGuid module
    
    guid = Replace(guid, "{", "")
    guid = Replace(guid, "}", "")
    guid = Replace(guid, "-", "")
    
    Dim strReturn As String
    strReturn = Left(guid, iLength)
    
    RandomGuid = strReturn
    

End Function

''' Pads the input string with specified character on the left of the input string
''' args: strIn = Input String, count = number of characters to pad (defaults to 1),
''' alpha = the character used to pad (defaults to a space)
Public Function PadLeft(ByVal strIn As String, _
            Optional ByVal count As Integer = 1, _
            Optional ByVal alpha As String = " ") As String
    
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":PadLeft"
    
    Dim strReturn As String
    Dim strPad As String
    
    Dim i As Integer
    For i = 1 To count
        strPad = alpha & strPad
    Next
    
    strReturn = strPad & strIn
    PadLeft = strReturn
    
    Exit Function
    
ThrowException:
    PadLeft = ""
    
End Function

''' Pads the input string with the specied character on the right of the input string
''' args: strIn = Input String, count = number of characters to pad (defaults to 1),
''' alpha = the character used to pad (defaults to a space)
Public Function PadRight(ByVal strIn As String, _
            Optional ByVal count As Integer = 1, _
            Optional ByVal alpha As String = " ") As String
            
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":PadRight"
    
    Dim strReturn As String
    Dim strPad As String
    
    Dim i As Integer
    For i = 1 To count
        strPad = strPad & alpha
    Next
            
    strReturn = strIn & strPad
    PadRight = strReturn
    
    Exit Function
    
ThrowException:
    PadRight = ""

End Function


' DATE FUNCTIONS

''' Returns the Year for the specified date
Public Function YearPart(ByVal dte As Date) As Integer

    Dim y1 As Integer
    y1 = DatePart(GetDatePartFormat(Year), dte)
    
    YearPart = y1

End Function

''' Returns the Month for the specified date
Public Function MonthPart(ByVal dte As Date) As Integer

    Dim m As Integer
    m = DatePart(GetDatePartFormat(Month), dte)
    MonthPart = m

End Function

''' Returns the Day for the specified date
Public Function DayPart(ByVal dte As Date) As Integer
    Dim d As Integer
    d = DatePart(GetDatePartFormat(DateInterval.day), dte)
    DayPart = d
End Function

''' Returns the Hour for the specified date
Public Function HourPart(ByVal dte As Date) As Integer
    Dim hr As Integer
    hr = DatePart(GetDatePartFormat(DateInterval.Hour), dte)
    HourPart = hr
End Function

''' Returns the Minute for the specified date
Public Function MinutePart(ByVal dte As Date) As Integer
    Dim m As Integer
    m = DatePart(GetDatePartFormat(Minute), dte)
    MinutePart = m
End Function

''' Returns the Second for the specified date
Public Function SecondPart(ByVal dte As Date) As Integer
    Dim s As Integer
    s = DatePart(GetDatePartFormat(Second), dte)
    SecondPart = s
End Function

''' Returns True if an Outlook or FME default date
''' is encountered.
Public Function IsDateNone(ByVal dte As Date) As Boolean
    
    Dim bReturn As Boolean
        
    Dim yr As Integer
    yr = DatePart("yyyy", dte)
    If yr > 4500 Or yr <= 1970 Then
            bReturn = True
    Else
            bReturn = False
    End If
        
    IsDateNone = bReturn
    
End Function

''' <summary>
''' Function used to check if two dates are identical. Ignores the time component.
''' </summary>
''' <param name="dteDate1">Date1</param>
''' <param name="dteDate2">Date2</param>
''' <returns>Boolean: True if identical; False if not or error</returns>
''' <remarks></remarks>
Public Function IsDateEqual(ByVal Date1 As Date, ByVal Date2 As Date) As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":IsDateEqual"
    
    On Error GoTo ThrowException

    Dim bReturn As Boolean
    bReturn = False
    
    Dim y1 As Integer
    y1 = DatePart(GetDatePartFormat(Year), Date1)
    Dim y2 As Integer
    y2 = DatePart(GetDatePartFormat(Year), Date2)
    If y1 = y2 Then
        Dim d1 As Integer
        d1 = DatePart(GetDatePartFormat(DayOfYear), Date1)
        Dim d2 As Integer
        d2 = DatePart(GetDatePartFormat(DayOfYear), Date2)
        If d1 = d2 Then bReturn = True
    End If
    
    IsDateEqual = bReturn
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    IsDateEqual = False
    
End Function

''' <summary>
''' Function returns true if Date1 is within the same month as Date2.
''' </summary>
''' <param name="Date1"></param>
''' <param name="Date2"></param>
''' <returns>Boolean: True if same month, otherwise False if not or error occurs.</returns>
''' <remarks></remarks>
Public Function IsSameMonth(ByVal Date1 As Date, ByVal Date2 As Date) As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":IsSameMonth"
    
    On Error GoTo ThrowException

    Dim bReturn As Boolean
    bReturn = False
    
    If YearPart(Date1) = YearPart(Date2) Then
        If MonthPart(Date1) = MonthPart(Date2) Then
            bReturn = True
        End If
    End If
    
    IsSameMonth = bReturn
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    IsSameMonth = False

End Function

''' <summary>
''' Function returns true if Date1 is within the same week as Date2.
''' </summary>
''' <param name="Date1">Input Date</param>
''' <param name="Date2">Date within the week of Interest</param>
''' <returns>Boolean: True if same week, otherwise False if not or an error occurs.</returns>
''' <remarks></remarks>
Public Function IsSameWeek(ByVal Date1 As Date, ByVal Date2 As Date) As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":IsSameWeek"
    
    On Error GoTo ThrowException
    
    Dim bReturn As Boolean
    bReturn = False
    
    Dim y1 As Integer
    y1 = DatePart(GetDatePartFormat(Year), Date1)
    Dim y2 As Integer
    y2 = DatePart(GetDatePartFormat(Year), Date2)
    
    If y1 = y2 Then
        Dim w1 As Integer
        w1 = DatePart(GetDatePartFormat(Week), Date1)
        Dim w2 As Integer
        w2 = DatePart(GetDatePartFormat(Week), Date2)
    Else
        ' DatePart method doesn't work when two dates cross
        '   the year mark
        Dim lastDay1 As Date
        lastDay1 = GetDateOnly(LastDayOfTheWeek(Date1))
        Dim lastDay2 As Date
        lastDay2 = GetDateOnly(LastDayOfTheWeek(Date2))
        
        If lastDay1 = lastDay2 Then bReturn = True
        
    End If
    
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    IsSameWeek = False
        
End Function

''' Returns True if both dates fall within the same hour on the same day
Public Function IsSameHour(ByVal dteDate1 As Date, ByVal dteDate2 As Date) As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":IsSameHour"
    
    On Error GoTo ThrowException
    
    Dim bReturn As Boolean
    bReturn = False

    If IsDateEqual(dteDate1, dteDate2) Then
        Dim h1 As Integer
        h1 = DatePart(GetDatePartFormat(Hour), dteDate1)
        Dim h2 As Integer
        h2 = DatePart(GetDatePartFormat(Hour), dteDate2)
        If h1 = h2 Then bReturn = True
    End If

    IsSameHour = bReturn
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    IsSameHour = False
   
End Function

''' Returns the VBA string used in the VBA DatePart function
Public Function GetDatePartFormat(ByVal iv As DateInterval) As String

'   Year = 1        ' yyyy
'   Quarter = 2     ' q
'   Month = 3       ' m
'   DayOfYear = 4   ' y
'   Day = 5         ' d
'   Weekday = 6     ' w
'   Week = 7        ' ww
'   Hour = 8        ' h
'   Minute = 9      ' n
'   Second = 10     ' s

    Dim s As String
    s = ""

    Select Case iv
        Case DateInterval.day
            s = "d"
        Case DateInterval.DayOfYear
            s = "y"
        Case DateInterval.Hour
            s = "h"
        Case DateInterval.Minute
            s = "n"
        Case DateInterval.Month
            s = "m"
        Case DateInterval.Quarter
            s = "q"
        Case DateInterval.Second
            s = "s"
        Case DateInterval.Week
            s = "ww"
        Case DateInterval.DayOfWeek
            s = "w"
        Case DateInterval.Year
            s = "yyyy"
        Case Else
            
    End Select
    
    GetDatePartFormat = s

End Function
    
''' Returns a string of yyyymmdd.hhnnss, e.g. 20190304.140345
Public Function GetDateTimeStamp(ByVal dte As Date) As String
    
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetDateTimeStamp"
          
    Dim strDate As String
    strDate = Format(dte, "yyyymmdd.hhnnss")
        
    GetDateTimeStamp = strDate
    Exit Function
                                    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    GetDateTimeStamp = ""
    
End Function

    ''' <summary>
    ''' Returns a date that is nearest the specified date for the given interval. If increment
    ''' is specified will return the date closest to the interval.
    ''' </summary>
    ''' <param name="dt">Date:</param>
    ''' <param name="interval">DateInterval:</param>
    ''' <param name="increment">Integer:</param>
    ''' <returns>Date:</returns>
    ''' <remarks>e.g. GetDateTimeNearest(Now, DateInterval.Minute, 30) will return the closest date and
    ''' time within a 30 minute window.</remarks>
    Public Function GetDateTimeNearest(ByVal dt As Date, Interval As DateInterval, Optional increment As Integer = 1) As Date
    
     '   Year = 1        ' yyyy
     '   Quarter = 2     ' q
     '   Month = 3       ' m
     '   DayOfYear = 4   ' y
     '   Day = 5         ' d
     '   Weekday = 6     ' w
     '   Week = 7        ' ww
     '   Hour = 8        ' h
     '   Minute = 9      ' n
     '   Second = 10     ' s
    
        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":MakeTaskFromItem"
    
        If increment <= 0 Then
            strTrace = "Invalid Increment, must be 1 or greater."
            GoTo ThrowException
        End If
        
        Dim min As Integer
        Dim hr As Integer
        Dim dy As Integer
        Dim mth As Integer
        Dim qtr As Integer
        
        Dim retDate As Date
        retDate = dt
        
        If increment = 1 Then
        
            Select Case Interval
                Case DateInterval.Year
                    mth = DatePart("m", dt)
                    If mth <= 6 Then
                        retDate = DateSerial(DatePart("yyyy", dt), 1, 1)
                    Else
                        retDate = DateSerial(DatePart("yyyy", dt) + 1, 1, 1)
                    End If
                
                Case DateInterval.Quarter
                    qtr = CInt(DatePart("m", dt) / 4) + 1
                    mth = qtr * 3
                    retDate = DateSerial(DatePart("yyyy", dt), mth, 1)
            
                Case DateInterval.Month
                    Dim dys As Integer
                    dys = 30 ' need a DaysInMonth calculator
                    Dim dyThreshold As Integer
                    dyThreshold = dys / 2
                    If DatePart("d", dt) <= dyThreshold Then
                        retDate = DateSerial(DatePart("yyyy", dt), DatePart("m", dt), 1)
                    Else
                        retDate = DateSerial(DatePart("yyyy", dt), DatePart("m", dt), dys)
                    End If
                
                Case DateInterval.day
                    retDate = DateSerial(DatePart("yyyy", dt), DatePart("m", dt), DatePart("d", dt))
            
                Case DateInterval.Hour
                    min = DatePart("n", dt)
                    If min <= 30 Then
                        retDate = DateSerial(DatePart("yyyy", dt), DatePart("m", dt), DatePart("d", dt)) + _
                                    TimeSerial(DatePart("h", dt), 0, 0)
                    Else
                        retDate = DateSerial(DatePart("yyyy", dt), DatePart("m", dt), DatePart("d", dt)) + _
                                    TimeSerial(DatePart("h", dt) + 1, 0, 0)
                    End If
            
                Case DateInterval.Minute
                    retDate = DateSerial(DatePart("yyyy", dt), DatePart("m", dt), DatePart("d", dt)) + _
                                    TimeSerial(DatePart("h", dt), DatePart("n", dt), 0)
            
                Case Else
                    strTrace = "Unsupported DateInterval: " & Interval & " - 1st Select."
                    LogMessage strTrace, strRoutine
            End Select
        
        Else
            Select Case Interval
                Case DateInterval.Month
                    mth = DatePart("m", dt)
                    Dim newMth As Integer
                    newMth = (WholeNumber(mth / increment) + 1) * increment
                
                    If newMth > increment Then
                         retDate = DateSerial(DatePart("yyyy", dt), increment, 1)
                    Else
                         retDate = DateSerial(DatePart("yyyy", dt) + 1, increment, 1)
                    End If
                
                Case DateInterval.day
                    dy = DatePart("d", dt)
                    Dim newDy As Integer
                    newDy = (WholeNumber(dy / increment) + 1) * increment
                    
                    retDate = DateSerial(DatePart("yyyy", dt), DatePart("m", dt), increment)
                    If newDy > increment Then retDate = DateAdd("m", 1, retDate)
                
                Case DateInterval.Hour
                    hr = DatePart("h", dt)
                    Dim newHr As Integer
                    newHr = (WholeNumber(hr / increment) + 1) * increment
                    
                    Dim dtTemp As Date
                    dtTemp = DateSerial(DatePart("yyyy", dt), DatePart("m", dt), DatePart("d", dt))
                    retDate = DateAdd("h", newHr, dtTemp)
                    
                Case DateInterval.Minute
                    hr = DatePart("h", dt)
                    min = DatePart("n", dt)
                    Dim newMin As Integer
                    newMin = (WholeNumber(min / increment) + 1) * increment
                    
                    dtTemp = DateSerial(DatePart("yyyy", dt), DatePart("m", dt), DatePart("d", dt)) + TimeSerial(hr, 0, 0)
                    retDate = DateAdd("n", newMin, dtTemp)
                
                Case Else
                    strTrace = "Unsupported DateInterval: " & Interval & " - 2nd Select."
                    LogMessage strTrace, strRoutine

            End Select
        
        End If
        

    
        GetDateTimeNearest = retDate
        Exit Function
                                    
ThrowException:
        LogMessageEx strTrace, err, strRoutine
    
    End Function
        
    ''' Returns an integer representing the day of the week
    ''' specified by the incoming date
    Public Function WeekDay(ByVal dteIn As Date) As Integer
    
        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":WeekDay"
    
        '   Sunday = 1
        '   Monday = 2
        '   Tuesday = 3
        '   Wednesday = 4
        '   Thursday = 5
        '   Friday = 6
        '   Saturday = 7
    
        Dim strDay As String
        strDay = Format(dteIn, "w")
        
        Dim iReturn As Integer
        iReturn = CInt(strDay)
        
        strTrace = Format(dteIn, "yyyymmdd - w")
    
        WeekDay = iReturn
    
    End Function

    ''' Returns the date that ends the week (i.e. Saturday) for
    ''' the week the specified date falls within
    Public Function LastDayOfTheWeek(ByVal dteIn As Date) As Date
    
        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":LastDayOfTheWeek"
        
        On Error GoTo ThrowException
        
        Dim dteReturn As Date
        dteReturn = #1/1/1970#
        
        Dim j As Integer
        j = WeekDay(dteIn)

        Dim intDelta As Integer
        intDelta = 7 - j
        
        dteReturn = DateAdd("d", intDelta, dteIn)
    
        LastDayOfTheWeek = dteReturn
        Exit Function
    
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        LastDayOfTheWeek = DateSerial(1970, 1, 1)
        
    End Function
    
    ''' Returns the number of days in a month
    Public Function DaysInMonth(ByVal dteIn As Date) As Integer
    
        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":DaysInMonth"
        
        On Error GoTo ThrowException
    
        Dim iDays As Integer
        
        Dim m As Integer
        m = MonthPart(dteIn)
        
        Select Case m
            Case 1
                iDays = 31
            Case 2
                iDays = 28
                If YearPart(dteIn) Mod 4 = 0 Then iDays = 29
            Case 3
                iDays = 31
            Case 4
                iDays = 30
            Case 5
                iDays = 31
            Case 6
                iDays = 30
            Case 7
                iDays = 31
            Case 8
                iDays = 31
            Case 9
                iDays = 30
            Case 10
                iDays = 31
            Case 11
                iDays = 30
            Case 12
                iDays = 31
        End Select
        
        DaysInMonth = iDays
        Exit Function
    
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        DaysInMonth = -1

    End Function
    
    ''' Returns just the Date part of the incoming date, i.e.
    ''' the Year, Month and Day (not Hour, Min, Second, etc..)
    Public Function GetDateOnly(ByVal dteIn As Date) As Date
    
        Dim dteReturn As Date
        
        dtReturn = DateSerial(DatePart("yyyy", dteIn), _
                              DatePart("m", dteIn), _
                              DatePart("d", dteIn))
                              
        GetDateOnly = dtReturn
    
    End Function

''' Returns a friendly format for showing elapsed time based on 's' seconds
Public Function FormatElapsedTime(ByVal s As Variant) As String

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":FormatElapsedTime"
        
    On Error GoTo ThrowException
        
    Dim strReturn As String
    strReturn = ""
        
    Dim i As Integer
        
    Dim dbl As Double
    dbl = s
        
        ' Days
        If dbl > 86400 Then
            ' > than 1 day
            i = WholeNumber(dbl / 86400)
            strReturn = strReturn & i & " days "
            dbl = Fraction(dbl / 86400) * 86400
        End If
        
        ' Hours
        If dbl > 3600 Then
            ' > than 1 hour
            i = WholeNumber(dbl / 3600)
            strReturn = strReturn & i & " hours "
            dbl = Fraction(dbl / 3600) * 3600
        End If
        
        ' Minutes
        If dbl > 60 Then
            ' > 1 minute
            i = WholeNumber(dbl / 60)
            strReturn = strReturn & i & " minutes "
            dbl = Fraction(dbl / 60) * 60
        End If
        
        strReturn = strReturn & CInt(dbl) & " seconds"

    FormatElapsedTime = strReturn
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    FormatElapsedTime = s & " seconds"

End Function

''' Returns the size of the specified file in bytes
Public Function FileLength(ByVal fullFilePath As String) As Long

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":FormatElapsedTime"
        
    On Error GoTo ThrowException
    
    If Len(fullFilePath) = 0 Then
        strTrace = "File path was empty."
        GoTo ThrowException
    End If
    
    Dim l As Long
    l = FileLen(fullFilePath)
    
    FileLength = l
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    FileLength = -1
    
End Function

''' Error Logger
Public Sub LogMessage(errMsg As String, method As String)

    Dim dte As Date
    dte = Now
    
    Dim strTrace As String
    strTrace = dte & "|" & method & "|" & errMsg
    
    ' Write to Console
    Debug.Print strTrace
    
    ' Write to TraceLog
    WriteToTraceLog strTrace

End Sub

Public Sub LogMessageEx(ByVal trace As String, _
                        ByVal err As ErrObject, _
               Optional ByVal method As String = "")

    Dim strTrace As String
    If Not err.Number = 0 Then
        strTrace = err.Number & "|" & trace & " " & err.Description & " - DLL Error: " & err.LastDllError & _
                    "Source: " & err.Source
        If Not Len(err.Source) = 0 Then method = err.Source & ":" & method
    Else
        strTrace = trace
    End If
    
    Dim dte As Date
    dte = Now
    
    Dim strMsg As String
    strMsg = dte & "|" & method & "|" & strTrace
    
    ' Write to SystemLog
    WriteToErrorLog strMsg
    
    ' Write to Console
    Debug.Print strMsg
   
End Sub

Public Sub TrimLogs(Optional ByVal rowsToKeep As Integer = 500)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TrimLogs"
    
    On Error GoTo ThrowException
    
    Call ClearTraceLog(rowsToKeep)
    Call ClearErrorLog(rowsToKeep)
    
    strTrace = "Trimmed system logs to " & rowsToKeep & " lines."
    LogMessage strTrace, strRoutine
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Public Sub ClearTraceLog(Optional ByVal rowsToKeep As Integer = 200)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ClearTraceLog"
    
    On Error GoTo ThrowException
    
    Dim logPath As String
    
    Dim errPath As String
    errPath = GetAppSystemPath
    logPath = errPath & "\" & traceLogFileName
    
    ClearLog logPath, True, rowsToKeep
    
    Exit Sub
    
ThrowException:
    Debug.Print "Failed to clear the trace log. " & err.Description

End Sub

Public Sub ClearErrorLog(Optional ByVal rowsToKeep As Integer = 200)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ClearErrorLog"
    
    On Error GoTo ThrowException
    
    Dim logPath As String
    
    Dim errPath As String
    errPath = GetAppSystemPath
    logPath = errPath & "\" & errorLogFileName
    
    ClearLog logPath, True, rowsToKeep
    
    Exit Sub
    
ThrowException:
    Debug.Print "Failed to clear the error/system log. " & err.Description

End Sub


''' Writes the specified message to the Error log file
Private Sub WriteToErrorLog(ByVal msg As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":WriteToErrorLog"
    
    On Error GoTo ThrowException
    
    Dim errPath As String
    errPath = GetAppSystemPath
    logPath = errPath & "\" & errorLogFileName
    AppendTextFile logPath, msg
    
    Exit Sub
    
ThrowException:
    Debug.Print "Failed to write to the error log. " & err.Description

End Sub

''' Writes the specified message to the Tracer log file
Private Sub WriteToTraceLog(ByVal msg As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":WriteToTraceLog"
    
    On Error GoTo ThrowException
    
    Dim logPath As String
    
    Dim errPath As String
    errPath = GetAppSystemPath
    logPath = errPath & "\" & traceLogFileName
    AppendTextFile logPath, msg
    
    Exit Sub
    
ThrowException:
    Debug.Print "Failed to write to the trace log. " & err.Description

End Sub

''' Clears the specified log file
Private Sub ClearLog(ByVal logFilePath As String, _
            Optional ByVal keepLastRows As Boolean = True, _
            Optional ByVal rowsToKeep As Integer = 500)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ClearLog"
    
    On Error GoTo ThrowException

    If Len(logFilePath) = 0 Then
        strTrace = "No log file path specified."
        GoTo ThrowException
    End If
    If Not FileExists(logFilePath) Then
        strTrace = "The specified log file '" & logFilePath & "'  does not exist."
        GoTo ThrowException
    End If
    
    If keepLastRows Then
    
        Dim strLine As String
        Dim strKeep As String
        Dim strIn As String
        
        ' read the current log
        strIn = ReadTextFile(logFilePath)
        If Len(strIn) = 0 Then
            strTrace = "Log file was empty."
            GoTo ThrowException
        End If
        
        ' parse the rows
        Dim rows() As String
        rows = Split(strIn, vbCr)
        
        ' Capture the last N rows
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        j = UBound(rows) - rowsToKeep
        If j < 0 Then j = 0
        k = UBound(rows) - 1
        For i = j To k
            strLine = Trim(rows(i))
            If Len(strLine) > 0 Then strKeep = strKeep & strLine
            
            strTrace = "Line length (" & i & ") = " & Len(strLine)
            Debug.Print strTrace
        Next
        
        ' Write the cleaned log file
        WriteTextFile logFilePath, strKeep
    
    Else
        If FileDelete(logFilePath) Then
            strTrace = "Successfully deleted the file: " & logFilePath & "."
            LogMessage strTrace, strRoutine
        Else
            strTrace = "Failed to delete the file: " & logFilePath & "'."
            GoTo ThrowException
        End If
    End If

    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub


' REFLECTION FUNCTIONS

Public Function GetProperty(ByVal o As Object, ByVal propName As String) As Variant

    Dim retObj As Object
    
    GetProperty = CallByName(o, propName, VbGet)
    
End Function

Public Sub SetProperty(ByVal o As Object, ByVal propName As String, ByVal val As Variant)

    CallByName o, propName, VbLet, val

End Sub

' SYSTEM FUNCTIONS

''' Returns the name of the current PC
Public Function GetComputerName() As String
    GetComputerName = VBA.Environ$("COMPUTERNAME")
End Function

''' Returns the path to the user's Roaming AppData path
Public Function GetUserAppDataPath() As String
    GetUserAppDataPath = VBA.Environ$("APPDATA")
End Function

''' Returns the path to the Ceptara Application Data Folder
Public Function GetCeptaraRootPath() As String

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetCeptaraRootPath"
    
    On Error GoTo ThrowException
    
    Dim strPath As String
    strPath = GetUserAppDataPath & "\Ceptara"
    
    If VBA.dir(strPath, vbDirectory) = "" Then
        MkDir strPath
    End If
    
    GetCeptaraRootPath = strPath
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    GetCeptaraRootPath = VBA.Environ$("TEMP")

End Function

''' Returns the file path to the Application's Root Storage folder
Public Function GetAppRootPath() As String

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetAppRootPath"
    
    On Error GoTo ThrowException
    
    Dim strPath As String
    strPath = GetCeptaraRootPath & "\" & Commands.AppName
    
    If VBA.dir(strPath, vbDirectory) = "" Then
        MkDir strPath
    End If
    
    GetAppRootPath = strPath
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    GetAppRootPath = VBA.Environ$("TEMP")

End Function

''' Returns the file path to the Application's Root Storage folder
Public Function GetAppDataPath() As String

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetAppDataPath"
    
    On Error GoTo ThrowException
    
    Dim strPath As String
    strPath = GetAppRootPath & "\Data"
    
    If VBA.dir(strPath, vbDirectory) = "" Then
        MkDir strPath
    End If
    
    GetAppDataPath = strPath
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    GetAppDataPath = VBA.Environ$("TEMP")

End Function

''' Returns the file path to the Application's Root Storage folder
Public Function GetAppSystemPath() As String

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetAppSystemPath"
    
    On Error GoTo ThrowException
    
    Dim strPath As String
    strPath = GetAppRootPath & "\System"
    
    If VBA.dir(strPath, vbDirectory) = "" Then
        MkDir strPath
    End If
    
    GetAppSystemPath = strPath
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    GetAppSystemPath = VBA.Environ$("TEMP")

End Function

''' Creates/Appends a Text file at the specified filePath with
''' the specified text (fileContent)
Public Sub AppendTextFile(ByVal filePath As String, ByVal fileContent As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":AppendTextFile"
    
    On Error GoTo ThrowException
    
    If FileLength(filePath) > 500000 Then
        TrimLogs 1000
    End If

    Dim TextFile As Integer

    'Determine the next file number available for use by the FileOpen function
    TextFile = FreeFile

    'Open the text file
    Open filePath For Append As TextFile

    'Write the text
    Print #TextFile, fileContent
  
    'Save & Close Text File
    Close TextFile
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
  
End Sub

''' Creates/Overwrites a Text file at the specified filePath with
''' the specified text (fileContent)
Public Sub WriteTextFile(ByVal filePath As String, ByVal fileContent As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":WriteTextFile"
    
    On Error GoTo ThrowException

    Dim TextFile As Integer

    'Determine the next file number available for use by the FileOpen function
    TextFile = FreeFile

    'Open the text file
    strTrace = "Opening file path: '" & filePath & "'..."
    Open filePath For Output As TextFile

    'Write the text
    strTrace = "Writing content (" & Len(fileContent) & " chars) to path: " & filePath
    Print #TextFile, fileContent
  
    'Save & Close Text File
    Close TextFile
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
  
End Sub

'''  Reads the text file at the specified filePath
Public Function ReadTextFile(ByVal filePath As String) As String

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ReadTextFile"
    
    On Error GoTo ThrowException

    Dim FileNum As Integer
    Dim DataLine As String
    Dim fileContent As String

    ' Determine next file # available
    FileNum = FreeFile()
    
    ' Open the Text file
    Open filePath For Input As #FileNum

    ' Get the contents of the file
    fileContent = Input(LOF(FileNum), FileNum)

    ' Close the file
    Close FileNum
    
'    ' Or line by line
'    While Not EOF(FileNum)
'        Line Input #FileNum, DataLine ' read in data 1 line at a time
'        ' decide what to do with dataline,
'        ' depending on what processing you need to do for each case
'    Wend
    
    ReadTextFile = fileContent
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    ReadTextFile = ""
    
End Function

''' Get SaveAs File path
Public Function GetWindowsFolder() As String

 

End Function

''' Deletes the specified file if it exists
Public Function FileDelete(ByVal strFullFilePath As String) As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":FileDelete"
    
    On Error GoTo ThrowException
    
    If FileExists(strFullFilePath) Then
        SetAttr strFullFilePath, vbNormal
        Kill strFullFilePath
    Else
        strTrace = "Failed to find file: " & strFullFilePath & " for deletion - does not exists."
        GoTo ThrowException
    End If
    
    FileDelete = True
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    FileDelete = False

End Function

''' Returns True if the specified file exists
Public Function FileExists(ByVal strFullFilePath As String) As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":FileExists"
    
    On Error GoTo ThrowException

    Dim bExists As Boolean
    
    If Len(strFullFilePath) = 0 Then
        strTrace = "No filename specified, string was empty."
        GoTo ThrowException
    End If
    If Right(strFullFilePath, 1) = "\" Then
        strTrace = "Invalid file path."
        GoTo ThrowException
    End If
    
    bExists = VBA.dir(strFullFilePath) <> ""
    
    FileExists = bExists
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    FileExists = False

End Function

''' <summary>
''' Used to start an application by passing a full file name to process.start
''' </summary>
''' <param name="strFileFullPath">Fully qualified file name</param>
''' <remarks></remarks>
Public Sub RunProgramFromFile(ByVal strFileFullPath As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":RunProgramFromFile"
    
    On Error GoTo ThrowException
    
    Dim shex As Object
    Set shex = CreateObject("Shell.Application")
    shex.ShellExecute strFileFullPath

    GoTo Finally

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
Finally:
    Set shex = Nothing
    
End Sub


' Environmental Variable Names
' ALLUSERSPROFILE=C:\ProgramData
' APPDATA=C:\Users\chris\AppData\Roaming
' asl.Log = Destination = File
' CommonProgramFiles=C:\Program Files\Common Files
' CommonProgramFiles(x86)=C:\Program Files (x86)\Common Files
' CommonProgramW6432=C:\Program Files\Common Files
' COMPUTERNAME = D5450WIN10
' ComSpec=C:\WINDOWS\system32\cmd.exe
' DriverData=C:\Windows\System32\Drivers\DriverData
' FPS_BROWSER_APP_PROFILE_STRING=Internet Explorer
' FPS_BROWSER_USER_PROFILE_STRING = Default
' HOMEDRIVE = c:
' HOMEPATH=\Users\chris
' LOCALAPPDATA=C:\Users\chris\AppData\Local
' LOGONSERVER=\\D5450WIN10
' NUMBER_OF_PROCESSORS = 4
' OneDrive=C:\Users\chris\OneDrive - Ceptara Corp
' OneDriveCommercial=C:\Users\chris\OneDrive - Ceptara Corp
' OneDriveConsumer=C:\Users\chris\OneDrive
' OS = Windows_NT
' Path=C:\Program Files\Microsoft Office\Root\Office16\;C:\Program Files (x86)\Intel\iCLS Client\;C:\Program Files\Intel\iCLS Client\;C:\WINDOWS\system32;C:\WINDOWS;C:\WINDOWS\System32\Wbem;C:\WINDOWS\System32\WindowsPowerShell\v1.0\;C:\Program Files\Microsoft SQL Server\130\Tools\Binn\;C:\Program Files\Git\cmd;C:\Program Files\dotnet\;C:\Program Files (x86)\Intel\Intel(R) Management Engine Components\DAL;C:\Program Files\Intel\Intel(R) Management Engine Components\DAL;C:\Program Files (x86)\Intel\Intel(R) Management Engine Components\IPT;C:\Program Files\Intel\Intel(R) Management Engine Components\IPT;C:\WINDOWS\System32\OpenSSH\;C:\Users\chris\AppData\Local\Microsoft\WindowsApps;;C:\Program Files\Microsoft Office\root\Client
' PATHEXT=.COM;.EXE;.BAT;.CMD;.VBS;.VBE;.JS;.JSE;.WSF;.WSH;.MSC
' PROCESSOR_ARCHITECTURE = AMD64
' PROCESSOR_IDENTIFIER=Intel64 Family 6 Model 61 Stepping 4, GenuineIntel
' PROCESSOR_LEVEL = 6
' PROCESSOR_REVISION = 30000#
' ProgramData=C:\ProgramData
' ProgramFiles=C:\Program Files
' ProgramFiles(x86)=C:\Program Files (x86)
' ProgramW6432=C:\Program Files
' PSModulePath=C:\Program Files\WindowsPowerShell\Modules;C:\WINDOWS\system32\WindowsPowerShell\v1.0\Modules
' PUBLIC=C:\Users\Public
' SESSIONNAME = Console
' SystemDrive = c:
' SystemRoot=C:\WINDOWS
' TEMP=C:\Users\chris\AppData\Local\Temp
' TMP=C:\Users\chris\AppData\Local\Temp
' USERDOMAIN = D5450WIN10
' USERDOMAIN_ROAMINGPROFILE = D5450WIN10
' UserName = chris
' USERPROFILE=C:\Users\chris
' windir=C:\WINDOWS

Sub EnumSEVars()
    Dim strVar As String
    Dim i As Long
    For i = 1 To 255
        strVar = VBA.Environ$(i)
        If LenB(strVar) = 0& Then Exit For
        Debug.Print strVar
    Next
End Sub


''' DRAWING Methods

''' Returns a value consistent with the ActiveX controls definition, see this URL
'''  https://supportline.microfocus.com/Documentation/AcucorpProducts/docs/v6_online_doc/gtman2/gt2913.htm
Public Function GetOleColor(ByVal c As enuColors) As Long

''' OLE Colors
'Private Const OLE_BLACK As Long = 0
'Private Const OLE_BLUE As Long = 255 * 65536
'Public Const OLE_GREEN As Long = 255 * 256
'Private Const OLE_CYAN = 255 * 256 + 255 * 65536
'Private Const OLE_RED = 255
'Private Const OLE_MAGENTA = 255 + 255 * 65536
'Private Const OLE_YELLOW = 255 + 255 * 256
'Private Const OLE_WHITE = 255 + 255 * 256 + 255 * 65536

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetOleColor"
    
    On Error GoTo ThrowException

    Dim retValue As Long

    Select Case c
        Case enuColors.Black
            retValue = 0
        Case enuColors.Blue
            retValue = 255 * 65536
        Case enuColors.Green
            retValue = 255 * 256 ' overflows on this number, not sure why
        Case enuColors.Cyan
            retValue = (255 * 256) + (255 * 65536)
        Case enuColors.Red
            retValue = 255
        Case enuColors.Magenta
            retValue = 255 + (255 * 65536)
        Case enuColors.Yellow
            retValue = 255 + (255 * 256)
        Case enuColors.White
            retValue = 255 + (255 * 256) + (255 * 65536)
        Case Else
    End Select
    
    GetOleColor = retValue
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    GetOleColor = 0

End Function
