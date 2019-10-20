Attribute VB_Name = "Commands"
''' Fields

    Private Const rootClass As String = "FME_Commands"
    
    Public Const AppName As String = "FocusMe for VBA"
    
    Public Const AppVersion As String = "1.0.24"

    Public Const DefaultLongStringLength As Integer = 3990

    Public Const DefaultDateOutlook As Date = #1/1/4501#

    Public Const ReferenceIDLength As Integer = 6
    Public Const ReferenceIDString As String = "FSRefID:"
    Public Const ReferenceIDUserProperty As String = "FMEReferenceID" ' Across device and service unique identifier

    Public Enum enuLinkStrategy
        None = 0
        text = 1        ' Copies ObjIn's text into ObjOut
        embed = 2       ' Embeds ObjIn into ObjOut
        multiple = 3    ' Copies ObjIn's attachments only
        referenced = 4  ' Creates a reference Id, linking the object's together
    End Enum
    
    Public Enum enuSortOn
        None = 0
        DueDate = 1
        Subject = 2
        Code = 3
        Priority = 4
        Name = 5
        startDate = 6
        CreatedDate = 6
        FolderPath = 7
        DeletedDate = 8
        ModifiedDate = 9
        itemType = 10
        Class = 11
        Calendar = 12
        LastName = 13
    End Enum

    Public Enum enuSortDirection
        None = 0
        Ascending = 1
        Descending = 2
    End Enum
  
    Public Const TeaserLength As Integer = 255
    
    Public Const PR_CATEGORIES As String = "urn:schemas-microsoft-com:office:office#Keywords"
    
    Public Const PR_EXCHANGE_ENTRYID As String = "http://schemas.microsoft.com/mapi/proptag/0x0FFF0102"
    Public Const PR_EXCHANGE_LONGTERM_ENTRYID As String = "http://schemas.microsoft.com/mapi/proptag/0x66700102"
    Public Const PR_SEARCH_KEY As String = "http://schemas.microsoft.com/mapi/proptag/0x300B0102"

    ' AppointmentItem Properties
    Public Const PR_APPOINTMENT_RECURRENCE As String = "http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/82310003"
    Public Const PR_APPOINTMENT_STARTDATE As String = "urn:schemas:calendar:dtstart"
    Public Const PR_APPOINTMENT_ENDDATE As String = "urn:schemas:calendar:dtend"
    Public Const PR_APPOINTMENT_ALLDAY As String = "urn:schemas:calendar:alldayevent"

    ' MailItem Properties
    Public Const PR_MAILITEM_SUBJECT As String = "urn:schemas:httpmail:subject"
    Public Const PR_MAILITEM_TOPIC As String = "urn:schemas:httpmail:thread-topic"
    Public Const PR_MAILITEM_FROM1ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x0065001f"
    Public Const PR_MAILITEM_FROM2ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x0042001f"
    Public Const PR_MAILITEM_TO1ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x0e04001f"
    Public Const PR_MAILITEM_TO2ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x0e03001f"

    ' ContactItem User Properties

    Public Const PR_POSCONTACT_PROPERTIES As String = "http://schemas.microsoft.com/mapi/string/" & _
            "{00020329-0000-0000-C000-000000000046}/POSContactProperties/0x0000001f"

    Public Const PR_IsMarkedAsTask_Flag As String = "http://schemas.microsoft.com/mapi/proptag/0x0E2B0003"

    Public Const PR_CONTACT_EMAIL1ADDRESS As String = "http://schemas.microsoft.com/mapi/id/{00062004-0000-0000-C000-000000000046}/8084001f"
    Public Const PR_CONTACT_EMAIL2ADDRESS As String = "http://schemas.microsoft.com/mapi/id/{00062004-0000-0000-C000-000000000046}/8094001f"
    Public Const PR_CONTACT_EMAIL3ADDRESS As String = "http://schemas.microsoft.com/mapi/id/{00062004-0000-0000-C000-000000000046}/80a4001f"

    ' TaskItem User Properties
    
    Public Const PR_TASK_COMPLETE As String = "http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/811c000b" ' {0 = No, 1 = Yes}
    Public Const PR_TASK_DUEDATE As String = "http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81050040"   ' Exists is not Null
    Public Const PR_TASK_PRIORITY As String = "urn:schemas:httpmail:importance" ' 0=Low, 1=Normal, 2=High
    Public Const PR_TASK_STATUS As String = "http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/81010003"
        ' 0=Not Started 1=In Progress 2=Complete 3=Waiting, 4=Deferred

    Public Const PR_POSTASK_PROPERTIES As String = "http://schemas.microsoft.com/mapi/string/" & _
        "{00020329-0000-0000-C000-000000000046}/POSTaskProperties/0x0000001f"

    Public Const PR_POSTASKCOLLABORATION_PROPERTIES As String = "http://schemas.microsoft.com/mapi/string/" & _
        "{00020329-0000-0000-C000-000000000046}/POSCollaborationProperties/0x0000001f"

    Public Const PR_POSTASKPARENT_PROPERTIES As String = "http://schemas.microsoft.com/mapi/string/" & _
        "{00020329-0000-0000-C000-000000000046}/POSParentProperties/0x0000001f"
        

    ''' Special filters for Tasks
    Public Enum enuTaskFilters
        None = 0
        Daily = 1
        Master = 2
        NoCategory = 3
        HighPriority = 4
        PastDue = 5
        InProgress = 6
        Waiting = 7
        Deferred = 8
    End Enum
        
   ''' Settings (Hard coded for now)
    ' Public Const AutoMove As Boolean = True
    'Public Const DestinationFolder As String = "\\christopher.m.lindstrom.civ@mail.mil\Archive"
    ' Public Const DestinationFolder As String = "\\chris@ceptara.net\Archive"
    ' Public Const ShowTaskWindowOnStartup As Boolean = True
    
' VBA: Sub OnAction(control As IRibbonControl, byRef CancelDefault)
    
''' Ribbon Callbacks

    Public Sub MyRibbonStartup(ByVal ribbonUI As Office.IRibbonUI)
        LogMessage "Ribbon callback worked.", "Commands:MyRibbonStartup"
    End Sub

    ''' Returns True if one item in the selected collection is an Outlook.MailItem
    Public Function IsMailItemSelected(ByVal control As Office.IRibbonControl) As Boolean
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":IsMailItemSelected"
        
        Dim bReturn As Boolean
        bReturn = False
        
        Dim myCollection As Outlook.Selection
        Set myCollection = ThisOutlookSession.Selection
             
        If myCollection.Count = 0 Then
            strTrace = "No incoming item found."
            GoTo ThrowException
        ElseIf myCollection.Count = 1 Then
            Set myItem = ThisOutlookSession.CurrentItem
            If TypeOf myItem Is Outlook.MailItem Then
                bReturn = True
            End If
        End If
        
        IsMailItemSelected = bReturn
        Exit Function
          
ThrowException:
        LogMessage strTrace, strRoutine
        IsMailItemSelected = False
          
    End Function

''' Menu Commands - - -

    Public Sub ReplyWithTracker()
        
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":ReplyWithTracker"
        
        On Error GoTo ThrowException
        
        strTrace = "Set up a Reply message with a tracking task."
               
        Dim myCollection As Outlook.Selection
        Set myCollection = ThisOutlookSession.Selection
        
        Dim ut As New Utilities
        Dim stgs As New Settings
    
        Dim oTask As Outlook.TaskItem
    
        If myCollection.Count = 0 Then
            strTrace = "No incoming item found."
            GoTo ThrowException
            
        ElseIf myCollection.Count = 1 Then
            Set myItem = ThisOutlookSession.CurrentItem
            If TypeOf myItem Is Outlook.MailItem Then
                Dim oMail As Outlook.MailItem
                Set oMail = myItem
               
                Dim rMail As Outlook.MailItem
                Set rMail = oMail.Reply
                rMail.FlagRequest = "FollowUpFlag"
                rMail.Display
                
            Else
                strTrace = "Unsupported Outlook Item encountered."
                LogMessage "WARNING: " & strTrace, strRoutine
            End If
        Else
            strTrace = "Replying to a collection of mailItems is not supported, " & _
                        "please select one outlook item and try again."
            MsgBox strTrace, vbInformation Or vbOKOnly, AppName
            
            strTrace = "Attempted to process multiple Outlook Items."
            GoTo ThrowException
            
        End If
        
        GoTo Finally
          
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        
Finally:
        Set ut = Nothing
        Set tm = Nothing
        
    End Sub
    
    Public Sub ForwardWithTracker()
        MsgBox "Set up a Forward message with a tracking task."
    End Sub
    


    ''' Defers the selected item to a task and moves
    ''' the originating item to the Project or Destination Folder
    Public Sub DeferToTask()
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":DeferToTask"
        
        On Error GoTo ThrowException
    
        Dim myCollection As Outlook.Selection
        Set myCollection = ThisOutlookSession.Selection
        
        Dim ut As New Utilities
        Dim stgs As New Settings
    
        Dim oTask As Outlook.TaskItem
    
        If myCollection.Count = 0 Then
            strTrace = "No incoming item found."
            GoTo ThrowException
            
        ElseIf myCollection.Count = 1 Then
        
            ' Create a task from the incoming Item
            Set myItem = ThisOutlookSession.CurrentItem
            Set oTask = ut.MakeTaskFromItem(myItem, embed, True, True)
            
            ' Link incoming MailItem to Task
            If TypeOf myItem Is Outlook.MailItem Then
                Dim tm As New TaskManager
                tm.LinkMailToTask myItem, oTask
            End If

            ' Show user the task
            oTask.Display
            
            ' Move the incoming Item if setting True
            If stgs.AutoMove Then
                If stgs.MoveOnDeferToTask Then
                    If stgs.IgnoreSentMailMove Then
                        If ut.IsItemParent(myItem, "Sent Items") Then
                            strTrace = "Ignoring the move request."
                            LogMessage strTrace, strRoutine
                        Else
                            ut.MoveToArchive myItem
                        End If
                    Else
                        ut.MoveToArchive myItem
                    End If
                End If
            End If
                 
        Else
            strTrace = "Converting a collection of incoming Outlook Items to a Task is not supported, " & _
                        "please select one outlook item and try again."
            MsgBox strTrace, vbInformation Or vbOKOnly, AppName
            
            strTrace = "Attempted to process multiple Outlook Items."
            GoTo ThrowException
            
        End If
          
        Exit Sub
          
ThrowException:
        LogMessageEx strTrace, err, strRoutine
          
    End Sub

    ''' Defers the selected item to a task and moves
    ''' the origination item to the Project or Destination Folder
    Public Sub DeferToAppt()
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":DeferToAppt"
    
        Dim myCollection As Outlook.Selection
        Set myCollection = ThisOutlookSession.Selection
        
        Dim ut As New Utilities
        Dim stgs As New Settings
    
        Dim oAppt As Outlook.AppointmentItem
    
        If myCollection.Count = 0 Then
            strTrace = "No incoming item found."
            GoTo ThrowException
        ElseIf myCollection.Count = 1 Then
            Set myItem = ThisOutlookSession.CurrentItem
            Set oAppt = ut.MakeAppointmentFromItem(myItem, embed, True, True)
            oAppt.Display
            
            If stgs.AutoMove Then
                If stgs.MoveOnDeferToAppt Then
                    If stgs.IgnoreSentMailMove Then
                        If ut.IsItemParent(myItem, "Sent Items") Then
                            strTrace = "Ignoring the move request."
                            LogMessage strTrace, strRoutine
                        Else
                            ut.MoveToArchive myItem
                        End If
                    Else
                        ut.MoveToArchive myItem
                    End If
                End If
            End If
            
        Else
            strTrace = "Converting a collection of incoming Outlook Items to an appointment is not supported, " & _
                        "please select one outlook item and try again."
            MsgBox strTrace, vbInformation Or vbOKOnly, AppName
            
            strTrace = "Attempted to process multiple Outlook Items."
            GoTo ThrowException
        End If
       
        Exit Sub
          
ThrowException:
        LogMessage strTrace, strRoutine
    
    End Sub
    
    ''' Creates a tracking task with a status of 'Waiting' - creates a stamped reply message
    ''' that connects the sent mail message with the tracking task
    Public Sub DelegateMailItem()
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":DelegateMailItem"
        
        NotImplemented (strRoutine)
        
        Exit Sub
          
ThrowException:
        LogMessage strTrace, strRoutine
    
    End Sub
    
    ''' Creates a linked mailItem from a selected
    ''' task.
    Public Sub MessageTask()
       
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":MessageTask"
        
        On Error GoTo ThrowException
        
        Dim myCollection As Outlook.Selection
        Set myCollection = ThisOutlookSession.Selection
        
        Dim ut As New Utilities
      
        If myCollection.Count = 0 Then
            strTrace = "No incoming item found."
            GoTo ThrowException
        ElseIf myCollection.Count = 1 Then
            Set myItem = ThisOutlookSession.CurrentItem
            Dim oMail As Outlook.MailItem
            Set oMail = ut.GetMessageForTask(myItem)
            If Not oMail Is Nothing Then oMail.Display
        Else
            For i = 1 To myCollection.Count
                Set myItem = myCollection(i)
                ut.MoveToArchive myItem
            Next i
        End If
        
        GoTo Finally
          
ThrowException:
        LogMessageEx strTrace, err, strRoutine
    
Finally:
    
    End Sub
    
    ''' Files the selected mail item into the item's project folder
    ''' or destination folder
    Public Sub FileInFolder()
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":FileInFolder"
        
        On Error GoTo ThrowException
        
        Dim myCollection As Outlook.Selection
        Set myCollection = ThisOutlookSession.Selection
        
        Dim ut As New Utilities
      
        If myCollection.Count = 0 Then
            strTrace = "No incoming item found."
            GoTo ThrowException
        ElseIf myCollection.Count = 1 Then
            Set myItem = ThisOutlookSession.CurrentItem
            ut.MoveToArchive myItem
        Else
            For i = 1 To myCollection.Count
                Set myItem = myCollection(i)
                ut.MoveToArchive myItem
            Next i
        End If
        
        GoTo Finally
          
ThrowException:
        LogMessage strTrace, strRoutine
        
Finally:
        Set ut = Nothing
    
    End Sub
    
    ''' Files the selected mail item as an attachment to a JournalItem, then
    ''' Moves the mail item to the item's project folder or destination folder
    Public Sub FileInJournal()
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":FileInJournal"
        
        Dim myCollection As Outlook.Selection
        Set myCollection = ThisOutlookSession.Selection
        
        Dim ut As New Utilities
        Dim stgs As New Settings
    
        Dim oJournal As Outlook.JournalItem
    
        If myCollection.Count = 0 Then
            strTrace = "No incoming item found."
            GoTo ThrowException
        ElseIf myCollection.Count = 1 Then
            Set myItem = ThisOutlookSession.CurrentItem
            Set oJournal = ut.MakeJournalEntryFromItem(myItem, embed, True, True)
            oJournal.Display
            
            If stgs.AutoMove Then
               If stgs.MoveOnFileInDrawer Then ut.MoveToArchive (myItem)
            End If
                 
        Else
            strTrace = "Converting a collection of incoming Outlook Items to a journal entry is not supported, " & _
                        "please select one outlook item and try again."
            MsgBox strTrace, vbInformation Or vbOKOnly, AppName
            
            strTrace = "Attempted to process multiple Outlook Items."
            GoTo ThrowException
        End If
        
        Exit Sub
          
ThrowException:
        LogMessage strTrace, strRoutine
    
    End Sub
    
    ''' Saves the selected mail item as a .msg file to a selected
    ''' Windows folder
    Public Sub FileInOSFolder()
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":FileInOSFolder"
        
        NotImplemented (strRoutine)
        
        Exit Sub
          
ThrowException:
        LogMessage strTrace, strRoutine
    
    End Sub
    
    ''' Presents the Mail Options dialog
    Public Sub PresentMailOptions()
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":PresentMailOptions"
        
        Dim frm As New frm_Options
        frm.Show
        
        Exit Sub
          
ThrowException:
        LogMessage strTrace, strRoutine
    
    End Sub
    
    ''' Presents timecard Form
    Public Sub PresentTimecard()
        ThisOutlookSession.StartTimecard
    End Sub
    
    
''' Text File Manipulation - - -

''' Creates/Appends a Text file at the specified filePath with
''' the specified text (fileContent)
Public Sub AppendTextFile(ByVal filePath As String, ByVal fileContent As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":AppendTextFile"
    
    On Error GoTo ThrowException

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
    Open filePath For Output As TextFile

    'Write the text
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


    
    Private Sub NotImplemented(Optional ByVal Name As String = "")
        If Len(Name) = 0 Then
            strTrace = "Selected function not implemented."
        Else
            strTrace = "Selected function (" & Name & ") not implemented."
        End If
        MsgBox strTrace, vbInformation Or vbOKOnly, AppName
    End Sub

