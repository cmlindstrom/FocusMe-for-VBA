Attribute VB_Name = "Tests"
''' UnitTest Methods - - -

Private Const rootClass As String = "Tests"

Public Sub Test()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Test"

    ' TestFolderRecursion
    
    ' TestDeferToTask
    
    ' ThisOutlookSession.StartTaskList

    ' Call TestRandom
    
    ' Call TestEnvironVariables
    
    ' Call TestGetCeptaraRootPath
    
    ' Call TestReadSettings
    
    ' Call TestProjectProperties
    
    ' Call TestCommandBars
    
    ' TestDataSet
    
    Call TestSettings
   
   
    ' Dim strStart As String
    ' strStart = Format(startDate, "mm/dd/yyyy hh:mm")
    
    ' ThisOutlookSession.StartTimecard
    
End Sub

Sub GetScreenSize()

    Dim w As Integer
    Dim h As Integer
    Dim n As Integer
    
    GetScreenResolution h, w, n

End Sub


Sub GetMonitorInfo()

    'WMI query
    Dim objWmiInterface As Object
    Dim objWmiQuery As Object
    Dim objWmiQueryItem As Object
    Dim strWQL As String
    'outputs
    Dim strDeviceId As String
    Dim strScreenName As String
    Dim varScreenHeight As Variant
    Dim varScreenWidth As Variant

    'run query
    strWQL = "Select * From Win32_DesktopMonitor"
    Set objWmiInterface = GetObject("winmgmts:root/CIMV2")
    Set objWmiQuery = objWmiInterface.ExecQuery(strWQL)
    'iterate output
    For Each objWmiQueryItem In objWmiQuery
        strDeviceId = objWmiQueryItem.DeviceId
        strScreenName = objWmiQueryItem.Name
        varScreenHeight = objWmiQueryItem.ScreenHeight
        varScreenWidth = objWmiQueryItem.ScreenWidth
        Debug.Print strDeviceId
        Debug.Print strScreenName
        Debug.Print varScreenHeight
        Debug.Print varScreenWidth
        Debug.Print ""
    Next

End Sub

Private Sub TestGetMoveCursor()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TestBrowseFolder"
    
    Dim bTest As Boolean
    bTest = False
    Dim bPass As Boolean
    bPass = True
    
    On Error GoTo ThrowException
    
    Dim X As Integer
    Dim y As Integer
    
    If GetCursorPosition(X, y) Then
        strTrace = "Current position = x: " & X & " y: " & y
        bTest = True
    Else
        strTrace = "An error occurred calling the Windows API (GetCursorPos)."
        bTest = False
    End If
    LogMessage strTrace, strRoutine
    bPass = bPass And bTest
    
    Sleep 1000
    
    strTrace = "Move the cursor 100 pixels to the right."
    If SetCursorPosition(X + 100, y) Then
        strTrace = "Moved the cursor to the new coordinates."
        bTest = True
    Else
        strTrace = "An error occurred calling the Windows API (SetCursorPos)."
        bTest = False
    End If
    LogMessage strTrace, strRoutine
    bPass = bPass And bTest
    
    Sleep 1000
    
    strTrace = "Move the cursor back to the original position."
    If SetCursorPosition(X, y) Then
        strTrace = "Moved the cursor to the new coordinates."
        bTest = True
    Else
        strTrace = "An error occurred calling the Windows API (SetCursorPos)."
        bTest = False
    End If
    LogMessage strTrace, strRoutine
    bPass = bPass And bTest
    
    GoTo Finally
       
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    bPass = False
        
Finally:
    strTrace = "Test disposition: " & bPass & "."
    LogMessage strTrace, strRoutine

End Sub

Private Sub TestElapsedTimeFormatter()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TestBrowseFolder"
    
    Dim bTest As Boolean
    bTest = False
    Dim bPass As Boolean
    bPass = True
    
    On Error GoTo ThrowException

    Dim s As Long
    
    s = 61
    strTrace = s & " seconds translates to " & FormatElapsedTime(s)
    MsgBox strTrace
    
    s = 7210
    strTrace = s & " seconds translates to " & FormatElapsedTime(s)
    MsgBox strTrace
    
    s = 100000
    strTrace = s & " seconds translates to " & FormatElapsedTime(s)
    MsgBox strTrace

    GoTo Finally
       
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    bPass = False
        
Finally:
    strTrace = "Test disposition: " & bPass & "."
    LogMessage strTrace, strRoutine
    
End Sub

Private Sub TestBrowseFolder()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TestBrowseFolder"
    
    Dim bTest As Boolean
    bTest = False
    Dim bPass As Boolean
    bPass = True
    
    On Error GoTo ThrowException
    
    strTrace = BrowseFolderEx("Test Title")
    If Len(strTrace) > 0 Then
        strTrace = "Selected folder: " & vbCrLf & strTrace
    Else
        strTrace = "No folder selected."
    End If
    
    MsgBox strTrace, vbOKOnly, strRoutine
    
    GoTo Finally
       
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    bPass = False
        
Finally:
    strTrace = "Test disposition: " & bPass & "."
    LogMessage strTrace, strRoutine

End Sub

Private Sub TestItemProperties()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TestRecurrencePattern"
    
    Dim bTest As Boolean
    bTest = False
    Dim bPass As Boolean
    bPass = True
    
    On Error GoTo ThrowException
    
    Dim ut As New Utilities
    Dim oContact As Outlook.ContactItem
    
    Set oContact = ut.WhoAmI
    
    Dim frm As New frm_ItemProperties
    frm.title = "Outlook Contact"
    frm.Description = oContact.FullName
    Set frm.Item = oContact
    frm.Show
    
    GoTo Finally
       
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    bPass = False
        
Finally:
    strTrace = "Test disposition: " & bPass & "."
    LogMessage strTrace, strRoutine

End Sub

Private Sub TestRecurrencePattern()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TestRecurrencePattern"
    
    Dim bTest As Boolean
    bTest = False
    Dim bPass As Boolean
    bPass = True
    
    On Error GoTo ThrowException
    
    Dim dteTest As Date
    dteTest = #6/3/2019# ' Monday
    
    Dim iDOW As Integer
    
    iDOW = WeekDay(dteTest)
    If iDOW = 2 Then
        strTrace = "June 3rd, 2019 is a Monday..."
        bTest = True
    Else
        strTrace = "June 3rd is a Monday, not a " & iDOW
        bTest = False
    End If
    bPass = bPass And bTest
    LogMessage strTrace, strRoutine
    
    strTrace = "Adding Monday to the Recurrence Pattern."
    LogMessage strTrace, strRoutine
    Dim fr As New fmeRecurrencePattern
    fr.AddToDayOfWeekPattern enuDayOfWeek.Monday
    
    bTest = fr.IsDayOfWeekChecked(enuDayOfWeek.Monday)
    If bTest Then
        strTrace = "Monday is in the mask values."
    Else
        strTrace = "Monday is in the mask values, but was not found."
    End If
    bPass = bPass And bTest
    LogMessage strTrace, strRoutine
    
    Dim iDay As Integer
'    iDay = fr.GetWeekDayFromMask(enuDayOfWeekMask.Monday)
'    If iDay = 2 Then
'        strTrace = "Returned Monday - PASS."
'        bTest = True
'    Else
'        strTrace = "Returned an incorrect WeekDayMask value: " & iDay
'        bTest = False
'    End If
'    bPass = bPass And bTest
'    LogMessage strTrace, strRoutine

    Dim dte As Date
    Set fr = New fmeRecurrencePattern
    fr.SetYearlyInstance 3, 15
    fr.Interval = 2 ' every 2 years
    fr.startDate = #4/15/2011#
    
    
    
    GoTo Finally
       
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    bPass = False
        
Finally:
    strTrace = "Test disposition: " & bPass & "."
    LogMessage strTrace, strRoutine

End Sub

Private Sub TestWhoAmI()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TestWhoAmI"
    
    Dim bTest As Boolean
    bTest = True
    
    On Error GoTo ThrowException

    strTrace = "Test Label"
    
    Dim ut As New Utilities
    Dim oContact As Outlook.ContactItem
    
    Set oContact = ut.WhoAmI

    GoTo Finally
       
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    bTest = False
        
Finally:
    strTrace = "Test disposition: " & bTest & "."
    LogMessage strTrace, strRoutine
    
End Sub

Private Sub TestFraction()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TestFraction"
    
    Dim bTest As Boolean
    bTest = True
    
    On Error GoTo ThrowException

    strTrace = "Test Label"
    
    Dim dbl As Double
    Dim frac As Double
    
    dbl = 3.46
    frac = Math.Fraction(dbl)
    If frac = 0.46 Then
        strTrace = "PASS: Converted: " & FormatNumber(dbl, 3) & " into " & FormatNumber(frac, 3)
    Else
        strTrace = "FAIL: Converted: " & FormatNumber(dbl, 3) & " into " & FormatNumber(frac, 3)
    End If
    LogMessage strTrace, strRoutine
    
    dbl = 3.65
    frac = Math.Fraction(dbl)
    If frac = 0.65 Then
        strTrace = "PASS: Converted: " & FormatNumber(dbl, 3) & " into " & FormatNumber(frac, 3)
    Else
        strTrace = "FAIL: Converted: " & FormatNumber(dbl, 3) & " into " & FormatNumber(frac, 3)
    End If
    LogMessage strTrace, strRoutine
    
    GoTo Finally
       
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    bTest = False
        
Finally:
   
    strTrace = "Test disposition: " & bTest & "."
    LogMessage strTrace, strRoutine

End Sub

Private Sub TestWholeNumber()

   Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TestWholeNumber"
    
    Dim bTest As Boolean
    bTest = True

    strTrace = "Test Label"
    Dim dbl As Double
    Dim l As Long
    
    dbl = 3.46
    l = Math.WholeNumber(dbl)
    If l = 3 Then
        strTrace = "PASS: Converted: " & FormatNumber(dbl, 3) & " into " & l
    Else
        strTrace = "FAIL: Converted: " & FormatNumber(dbl, 3) & " into " & l
    End If
    LogMessage strTrace, strRoutine
    
    dbl = 3.5
    l = Math.WholeNumber(dbl)
    If l = 3 Then
        strTrace = "PASS: Converted: " & FormatNumber(dbl, 3) & " into " & l
    Else
        strTrace = "FAIL: Converted: " & FormatNumber(dbl, 3) & " into " & l
    End If
    LogMessage strTrace, strRoutine
    
    dbl = 3.65
    l = Math.WholeNumber(dbl)
    If l = 3 Then
        strTrace = "PASS: Converted: " & FormatNumber(dbl, 3) & " into " & l
    Else
        strTrace = "FAIL: Converted: " & FormatNumber(dbl, 3) & " into " & l
    End If
    LogMessage strTrace, strRoutine

    GoTo Finally
       
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    bTest = False
        
Finally:
   
    strTrace = "Test disposition: " & bTest & "."
    LogMessage strTrace, strRoutine

End Sub

Private Sub TestGraphics()

   Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TestGraphics"
    
    Dim bTest As Boolean
    bTest = True

    strTrace = "Test Label"
    
    Dim ilen As Integer
    ilen = MeasureString(strTrace)
    
    If ilen > 0 Then
        strTrace = "PASS: The length of the string: " & strtace & " = " & ilen
    Else
        strTrace = "FAIL: gdi library call failed."
    End If
    LogMessage strTrace, strRoutine

    GoTo Finally
       
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    bTest = False
        
Finally:
   
    strTrace = "Test disposition: " & bTest & "."
    LogMessage strTrace, strRoutine

End Sub

Private Sub TestRunFromFile()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TestRunFromFile"
    
    Dim bTest As Boolean
    bTest = True
    
    Dim strFilePath As String
    strFilePath = "C:\Users\chris.m.lindstrom\Desktop\Test.txt"
    
    Dim bExists As Boolean
    bExists = VBA.dir(strFilePath) <> ""
    
    If Not bExists Then
        strTrace = "File: " & strFilePath & " does not exist - aborting test."
        GoTo ThrowException
    End If
    
    RunProgramFromFile strFilePath
       
    GoTo Finally
       
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    bTest = False
        
Finally:
   
    strTrace = "Test disposition: " & bTest & "."
    LogMessage strTrace, strRoutine

End Sub

Private Sub TestGetCalendars()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TestDateFormatting"
    
    Dim bTest As Boolean
    bTest = True

    Dim ut As New Utilities
    Dim myList As ArrayList
    Set myList = ut.GetCalendarFolders

    If myList Is Nothing Then
        strTrace = "Folder query failed."
        bTest = False
    Else
        If myList.count = 0 Then
            strTrace = "Didn't find any Calendar folders - probably wrong."
            bTest = False
        Else
            strTrace = "Found " & myList.count & " calendar folders."
        End If
    End If
    LogMessage strTrace, strRoutine

    GoTo Finally
       
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    bTest = False
        
Finally:
   
    strTrace = "Test disposition: " & bTest & "."
    LogMessage strTrace, strRoutine

End Sub

Private Sub TestTimeRecords()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TestTimeRecords"
    
    Dim bTest As Boolean
    bTest = True
    
    Dim strFilePath As String
  
    Dim strStart As String
    strStart = Format(Now, "mm/dd/yyyy hh:mm")
    
    Dim dteEnd As Date
    dteEnd = DateSerial(2019, 4, 6)
    
    Dim dteStart As Date
    dteStart = DateAdd("d", -7, Now)
    
    strTrace = "Date Format: " & strStart
    LogMessage strTrace, strRoutine
    
    Dim tr As New TimeRecords
    tr.Load dteStart, dteEnd
    
    Dim t As fmeTimeRecord
    For Each t In tr.Items
        strTrace = t.Subject & ", " & t.RecordDate & " | Dur: " & _
                    t.Duration & ", " & t.ReportedDurationHours & _
                    " | Categories: " & t.Categories
        LogMessage strTrace, strRoutine
    Next
    
'    strFilePath = "C:\Users\chris.m.lindstrom\Desktop\Test.csv"
'    tr.Export strFilePath
    
    Dim dt As DataTable
    Set dt = tr.Analyze("Test", True, True)
    
    If dt Is Nothing Then
        strTrace = "Analysis failed."
        GoTo ThrowException
    End If
    
    Dim exp As New csvExporter
    strFilePath = "C:\Users\chris.m.lindstrom\Desktop\Timecard.csv"
    strTrace = exp.ExportTable(dt, strFilePath)
    If Len(strTrace) = 0 Then
        strTrace = "Export table failed."
        GoTo ThrowException
    End If
    strTrace = "Table Dump:" & vbCrLf & strTrace
    LogMessage strTrace, strRoutine
          
    GoTo Finally
       
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    bTest = False
        
Finally:
   
    strTrace = "Test disposition: " & bTest & "."
    LogMessage strTrace, strRoutine

End Sub

Private Sub TestDisableAddin()

    Dim c As COMAddIn
    For Each c In ThisOutlookSession.COMAddIns
        LogMessage "Addin: " & c.Description, "TestDisableAddin"
        If InStr(1, LCase(c.Description), "classific") > 0 Then
            c.Connect = True
        End If
    Next


End Sub

Private Sub TestGetMailFolders()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TestGetMailFolders"
    
    Dim bTest As Boolean
    bTest = True
    
    Dim ut As New Utilities
    
    Dim mFolders As ArrayList
    Set mFolders = ut.GetMailFolders
    If mFolders Is Nothing Then
        strTrace = "Method failed."
        GoTo ThrowException
    End If
    If mFolders.count = 0 Then
        strTrace = "FAIL: found zero folders."
        LogMessage strTrace, strRoutine
        bTest = False
        GoTo Finally
    End If
    
    Dim f As Outlook.Folder
    For Each f In mFolders
        strTrace = "Mail folder: " & f.FolderPath
        LogMessage strTrace, strRoutine
    Next
    
    strTrace = "Found " & mFolders.count & " mail folders."
    LogMessage strTrace, strRoutine
       
    GoTo Finally
       
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    bTest = False
        
Finally:
    Set ut = Nothing
    
    strTrace = "Test disposition: " & bTest & "."
    LogMessage strTrace, strRoutine

End Sub

Private Sub TestDataStoreRemove()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TestDataStoreRemove"
    
    Dim bTest As Boolean
    bTest = False
    
    Dim ldb As New dsDataStore
    ldb.Connect "TestDB"

    Dim id As String
    id = "C3E369A7"
    
    Dim p As fmeProject
    Set p = ldb.GetItemById(id, "Project")
    If p Is Nothing Then
        LogMessage "Failed to retrieve a Project from the 'Project' collection.", strRoutine
        Exit Sub
    End If
    
    bTest = ldb.Remove(id, "Project")
    If bTest Then
        strTrace = "PASS: Remove completed successfully."
    Else
        strTrace = "FAIL: failed to remove the test project."
    End If
    LogMessage strTrace, strRoutine
    
    ldb.Disconnect
    
End Sub

Private Sub TestDataStoreUpdate()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TestDataStoreUpdate"
    
    Dim bTest As Boolean
    bTest = False

    Dim ldb As New dsDataStore
    ldb.Connect "TestDB"

    Dim id As String
    id = "A95691AD"
    
    Dim p As fmeProject
    Set p = ldb.GetItemById(id, "Project")
    If p Is Nothing Then
        LogMessage "Failed to retrieve a Project from the 'Project' collection.", strRoutine
        Exit Sub
    End If
    
    Dim c As Integer
    c = p.Color
    
    p.Color = c + 1
    bTest = ldb.Update(p, "Project")
    If bTest Then
        strTrace = "PASS: update completed successfully."
    Else
        strTrace = "FAIL: failed to update the test project."
    End If
    LogMessage strTrace, strRoutine
    
    ldb.Disconnect
    
End Sub

Private Sub TestDataStoreInsert()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TestDataStoreInsert"
    
    Dim ldb As New dsDataStore
    ldb.Connect "TestDB"
    
    Dim p As New fmeProject
    p.Active = True
    p.Code = GenerateUniqueID(4)
    p.Name = "Number " & GenerateUniqueID(2) & " Project"
    p.Color = 8
    
    ldb.Insert p, "Project"
    
    ldb.Disconnect

End Sub

Private Sub TestDataSet()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TestDataSet"

    Dim ds As New DataSet
    ds.Name = "TestSet"
    
    ' Table 1
    Dim dt As New DataTable
    dt.Name = "Profile"
    dt.Columns.Add "Name"
    dt.Columns.Add "Description"
    dt.Columns.Add "DateCreated"
    
    Dim dr As New DataRow
    dr.Add "Name", "Chris Lindstrom"
    dr.Add "Description", "Tall and Willowy"
    dr.Add "DateCreated", #1/1/1970#
    
    dt.rows.AddRow dr
    
    ds.Tables.Add dt
    
    ' Table 2
    Set dt = New DataTable
    dt.Name = "Address"
    dt.Columns.Add "Name"
    dt.Columns.Add "Addr1"
    dt.Columns.Add "City"
    dt.Columns.Add "State"
    
    Set dr = New DataRow
    dr.Add "Name", "Chris Lindstrom"
    dr.Add "Addr1", "1601 R St."
    dr.Add "City", "Denison"
    dr.Add "State", "IA"
    
    dt.rows.AddRow dr
    
    ds.Tables.Add dt
    
    Dim strXml As String
    strXml = ds.GetXml

    LogMessage strXml, strRoutine
    
    Dim fn As String
    fn = GetAppDataPath & "\testDSFile.xml"
    ds.WriteXmlFile fn
    
    Sleep 1000

    ds.ReadXmlFile fn
    
    fn = GetAppDataPath & "\testDSFileBackup.xml"
    ds.WriteXmlFile fn

End Sub

Private Sub TestShowPopup()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":TestShowPopup"

    Dim myExplorer As Outlook.Explorer
    Set myExplorer = ThisOutlookSession.ActiveExplorer
    
    Dim objCommandBars As Office.CommandBars
    Set objCommandBars = myExplorer.CommandBars
    
    Dim strMnuName As String
    strMnuName = "TestPopup"
    
    On Error Resume Next
    objCommandBars(strMnuName).Delete
    On Error GoTo 0
    
    With objCommandBars.Add(Name:=strMnuName, Position:=msoBarPopup)
        With .Controls.Add(msoControlButton)
            .OnAction = "Tests.ShowForm"
            .FaceId = 264
            .Caption = "Test Button"
        End With
    End With
    
    objCommandBars(strMnuName).Visible = True
    objCommandBars(strMnuName).ShowPopup

End Sub

Public Sub ShowForm()
    MsgBox "Showed form."
End Sub

Private Sub TestProjectProperties()

    Dim p As New fmeProject
    p.Name = "Test 123"
    p.Code = "100110"
    ' p.SetStatusFromName ("Completed")
    p.SetStatus olTaskInProgress
    p.SetPriority olImportanceHigh
    
    Dim frm As New frm_ProjectProperties
    frm.Load p
    frm.Show

End Sub

Private Sub TestReadSettings()

    Dim stgs As New Settings
    Dim b As Boolean
    b = stgs.AutoMove

End Sub

Private Sub TestWriteSettings()

    Dim stgs As New Settings
    
    stgs.AutoMove = True
    
    stgs.Save
    

End Sub

Private Sub TestGetAppRootPath()
    Dim s As String
    s = Common.GetAppRootPath
    LogMessage "Application Root Path: " & s, "Tests"
End Sub

Private Sub TestGetCeptaraRootPath()
    Dim s As String
    s = Common.GetCeptaraRootPath
    LogMessage "Ceptara Root Path: " & s, "Tests"
End Sub

Private Sub TestGetAppDataPath()
    Dim s As String
    s = Common.GetUserAppDataPath
    LogMessage "APPDATA Path: " & s, "Tests"
End Sub

Private Sub TestEnvironVariables()
    EnumSEVars
End Sub

Private Sub TestRandom()

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":TestRandom"

    strTrace = RandomGuid(8)
    LogMessage strTrace, strRoutine

End Sub

Private Sub TestFolderRecursion()

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":TestFolderRecursion"
    
    LogMessage "-- Starting Test-- TestFolderRecursion ", strRoutine

    Dim strRootFolderName As String
    strRootFolderName = "Archive"
    
    Dim ut As New Utilities
    
    Dim myColl As Collection
    Set myColl = ut.GetFoldersWithName(strRootFolderName)
    
    If IsNothing(myColl) Then
        strTrace = "FAIL: Failed to find '" & strRootFolderName & " in the folder tree."
        LogMessage strTrace, strRoutine
    Else
        strTrace = "PASS: Found " & myColl.count & " folders in the folder tree."
        LogMessage strTrace, strRoutine
        
        For i = 1 To myColl.count
            Dim f As Outlook.Folder
            Set f = myColl(i)
            
            strTrace = "- - Found folder: " & f.FolderPath
            LogMessage strTrace, strRoutine
        Next
        
    End If

End Sub

Private Sub TestDeferToTask()

    Set myItem = ThisOutlookSession.CurrentItem
    
    Dim ut As New Utilities
    
    Dim oTask As Outlook.TaskItem
    Set oTask = ut.MakeTaskFromItem(myItem, embed, True, True)
    
    oTask.Display

End Sub

Private Sub TestItemSelect()
  ' Dim frm As New FME_Pane
  ' frm.Show
  
    Set myItem = ThisOutlookSession.CurrentItem
    If myItem Is Nothing Then
        MsgBox "No selection set."
    Else
        If TypeOf myItem Is Outlook.MailItem Then
            Dim oItem As Outlook.MailItem
            Set oItem = myItem
            MsgBox "Selection: " & oItem.Subject
        Else
            MsgBox "Selection is set - unknown type."
        End If
    End If
  
End Sub

Private Sub TestCommandBars()

    Dim myExplorer As Outlook.Explorer
    Set myExplorer = ThisOutlookSession.ActiveExplorer
    
    Dim objCommandBars As Object
    Set objCommandBars = myExplorer.CommandBars
    If IsNothing(objCommandBars) Then
        LogMessage "CommandBars object was null.", rootClass
        Exit Sub
    End If
    
    LogMessage "---- TestCommandBars", rootClass
    
    For Each objCommandBar In objCommandBars
        LogMessage objCommandBar.Name, rootClass
        
        If objCommandBar.Controls.count > 0 Then
            LogMessage " - Controls are:", rootClass
            For Each ctl In objCommandBar.Controls
                LogMessage objCommandBar.Name & ", control: " & ctl.Caption, rootClass
            Next ctl
        Else
            LogMessage " - no buttons found...", rootClass
        End If
    
    Next objCommandBar

End Sub

Private Sub TestLogger()
    Common.LogMessage "Error message", "calling method"
End Sub

Private Sub FirstTest()
    Dim dResult As VbMsgBoxResult
    dResult = MsgBox("Passed", vbCritical Or vbOKOnly, "Test: FirstTest")
End Sub

Private Sub TestSettings()
    Dim stgs As New Settings
    stgs.UpdateSetting_Test
End Sub

