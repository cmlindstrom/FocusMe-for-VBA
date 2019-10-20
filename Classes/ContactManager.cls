VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContactManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' - Fields

Private Const rootClass As String = "ContactManager"

Dim app As Outlook.Application
Dim oNS As Outlook.NameSpace

Private WithEvents singleInspector As Outlook.Inspector
Attribute singleInspector.VB_VarHelpID = -1
Private ut As Utilities

Private WithEvents myListView As MSComctlLib.ListView
Attribute myListView.VB_VarHelpID = -1

Dim sortingColumn As enuSortOn
Dim sortingDirection As enuSortDirection

' - Events

''' Thrown when a contact is changed by this class
Public Event ItemUpdated()
Public Event ItemAdded()
Public Event collectionUpdated()

' - Properties

Dim myItems As Contacts
Dim selItem As Outlook.ContactItem
Dim selItems As Contacts

Dim lvCheckbox As Boolean
Dim lvMultiSelect As Boolean
Dim f_SuspendUI As Boolean
Dim f_lastQuery As String
Dim f_listFiltered As Boolean

''' Latest collection of Contacts
Public Property Set Items(ByVal coll As Contacts)
    Set myItems = coll
End Property
Public Property Get Items() As Contacts
    Set Items = myItems
End Property

''' Current Task being changed
Public Property Get SelectedItem() As Outlook.ContactItem
    Set SelectedItem = selItem
End Property

''' Currently selected Collection of Contacts
Public Property Get SelectedItems() As Contacts
    Set SelectedItems = selItems
End Property

''' Number of items that are checked
Public Property Get ItemsChecked() As Integer
    Dim i As Integer
    i = selItems.Count
    ItemsChecked = i
End Property


''' Managed ListView
Public Property Set ListView(ByVal lv As MSComctlLib.ListView)
    Set myListView = lv
End Property
Public Property Get ListView() As MSComctlLib.ListView
    Set ListView = myListView
End Property

''' Sets / Gets the 'show checkboxes' flag for the ListView
Public Property Let ListViewCheckBox(ByVal b As Boolean)
    lvCheckbox = b
End Property
Public Property Get ListViewCheckBox() As Boolean
    ListViewCheckBox = lvCheckbox
End Property

''' Sets / Gets the ability to select more than one Project
Public Property Let ListViewMultiSelect(ByVal b As Boolean)
    lvMultiSelect = b
End Property
Public Property Get ListViewMultiSelect() As Boolean
    ListViewMultiSelect = lvMultiSelect
End Property

''' Sets / Gets flag for ignoring UI event calls
''' - Useful when more than one controller managing
''      a shared UI control, e.g. ListView
Public Property Let SuspendUIEvents(ByVal b As Boolean)
    f_SuspendUI = b
End Property
Public Property Get SuspendUIEvents() As Boolean
    SuspendUIEvents = f_SuspendUI
End Property


' - Event Handlers
Private Sub singleInspector_Activate()

End Sub

Private Sub singleInspector_Close()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":singleInspector_Close"
    
    On Error GoTo ThrowException
    
    ' Check last mod date - this will error if contact
    '   has been deleted
    strTrace = selItem.LastModificationTime
    
    Dim dteSaved As Date
    dteSaved = selItem.LastModificationTime
    
    If IsDateNone(dteSaved) Then
        ' Ignore - item opened and never saved
    Else
        ' Look in the collection
        Dim idx As Integer
        idx = IsItemInCollection(selItem)
        If idx >= 0 Then
            ' Update collection & UI
            UpdateItem selItem, idx
            RaiseEvent ItemUpdated
            
        Else
            ' Update collection
            myItems.AddItem selItem
            myItems.Sort Name, Ascending
            RaiseEvent ItemAdded
            
            ' Update Managed UI
            AddListViewItem selItem
        End If

    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err

End Sub

' - - ListView Event Handlers

Private Sub myListView_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Dim strTrace As String
    strTrace = ""
    Dim strRoutine As String
    strRoutine = rootClass & ":myListView_ItemClick"
    
    If f_SuspendUI Then Exit Sub

    Dim eid As String
    eid = Item.key
    Set selItem = ut.GetOutlookItemFromID(eid)
    If IsNothing(selItem) Then
        strTrace = "Failed to find an Outlook Item for a known good EntryID: '" & eid & "'."
        LogMessage "WARNING: " & strTrace, strRoutine
    Else
        strTrace = "Selected a ContactItem."
        LogMessage "INFORMATION: " & strTrace, strRoutine
    End If
    
End Sub

Private Sub myListView_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":myListView_ItemCheck"
    
    If f_SuspendUI Then Exit Sub
    
    Dim eid As String
    eid = Item.key
    Dim c As Outlook.ContactItem
    Set c = ut.GetOutlookItemFromID(eid)
    If Not IsNothing(c) Then
        If Item.checked Then
            ' Add the item to the selected items
            If Not selItems.IsItemPresent(eid) Then
                selItems.AddItem c
            Else
                strTrace = "Attempted to add a duplicate contact to the selected items."
                LogMessage "WARNING: " & strTrace, strRoutine
            End If
        Else
            ' Remove the item from the selected items
            Dim ct As Outlook.ContactItem
            Set ct = selItems.GetItemById(eid)
            If Not IsNothing(ct) Then
                strTrace = "Remove contact from selected items; eid: " & eid & "."
                selItems.RemoveItem ct
            End If
        End If
    Else
        strTrace = "Failed to find an Outlook Item for a known good EntryID: '" & eid & "'."
        LogMessage "WARNING: " & strTrace, strRoutine
    End If
    

End Sub

' - Constructor

Private Sub Class_Initialize()
    Set myItems = New Contacts
    Set selItems = New Contacts
    
    Set oNS = GetNamespace("MAPI")
    Set app = oNS.Application
    Set ut = New Utilities
    
    sortingDirection = Ascending
    
End Sub

Private Sub Class_Terminate()
    Set myItems = Nothing
    Set selItem = Nothing
    Set selItems = Nothing
    Set myListView = Nothing
    
    Set oNS = Nothing
    Set app = Nothing
    Set ut = Nothing
End Sub

' - Methods

''' Load all Outlook Contacts
Public Sub Load()

    'Set myItems.Items = GetOutlookTasks()
    'Call Refresh
    
    LoadDASL ""
    
    ' Inform listeners
    RaiseEvent collectionUpdated
    
End Sub

Public Sub LoadByProject(ByVal Name As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":LoadByProject"
    
    If Len(Name) > 0 Then
        Dim strQuery As String
        
        ' Any contact matching the category name
        strQuery = "@SQL=""" & PR_CATEGORIES & """ like '%" & Name & "%'"
        LoadDASL strQuery
        
    Else
        strTrace = "An empty project name encountered."
        LogMessage "WARNING: " & strTrace, strRoutine
    End If
    
End Sub

''' Load tasks using the specified DASL Query
Public Sub LoadDASL(ByVal daslFilter As String)

    Dim arList As ArrayList
    Set arList = GetOutlookContacts(daslFilter, True)
    myItems.Items = arList
    myItems.Sort Name, Ascending
    
    ' Capture query
    f_lastQuery = daslFilter
    
    ' Update the Managed UI
    Call Refresh
    
    ' Inform listeners
    RaiseEvent collectionUpdated
    
End Sub

' Sorts the current Outlook Item collection
Public Sub Sort(ByVal srt As enuSortOn, ByVal dir As enuSortDirection)
    myItems.Sort srt, dir
    
    ' Update UI
    Call Refresh
    
End Sub

' - Contact Item Commands

''' Opens an existing contact
Public Sub OpenContact(Optional ByVal c As Outlook.ContactItem = Nothing)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":OpenContact"
    
    On Error GoTo ThrowException

    If c Is Nothing Then
        Set c = selItem
    Else
        Set selItem = c
    End If
    
    If IsNothing(c) Then
        strTrace = "Incoming OutlookItem was null."
        GoTo ThrowException
    End If

   ' Prepare to present Outlook Task UI
    Set singleInspector = c.GetInspector
    
    ' Show the Item using the Outlook interface
    c.Display
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Creates a new Contact to Edit
Public Sub NewContact(Optional p As fmeProject = Nothing)

    ' Create a new Contact
    Dim c As Outlook.ContactItem
    Set c = ut.CreateContact
    
    ' If a project provided, assign task to the Project
    If Not p Is Nothing Then c.Categories = p.Subject
    
    ' Prepare to present Outlook Contact UI
    Set selItem = c
    Set singleInspector = c.GetInspector
    
    ' Show the Contact using the Outlook interface
    c.Display

End Sub

''' Remove from Project
Public Sub RemoveContactProject(ByVal p As fmeProject, _
                       Optional ByVal c As Outlook.ContactItem = Nothing)
                       
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":RemoveContactProject"

    If p Is Nothing Then
        Exit Sub
    End If

    If c Is Nothing Then c = selItem

    ' Get the list of projects
    Dim pm As New ProjectManager
    pm.Load
    
    ' Get the contact's projects
    Dim currProjs As ArrayList
    Set currProjs = pm.ParseProjectsFromString(c.Categories)
    
    ' Remove specified project from contact's project list
    Dim tp As fmeProject
    For Each tp In currProjs
        If LCase(p.Subject) = LCase(tp.Subject) Then
            ' remove project from contact's list
            currProjs.Remove tp
            ' Log
            strTrace = "INFO: Removed project: " & tp.Subject & " from Contact."
            LogMessage strTrace, strRoutine
            Exit For
        End If
    Next
    
    ' Rewrite contact's categories
    
    
    
End Sub

''' Send Email to contact
Public Sub NewEmail(Optional p As fmeProject = Nothing, _
                    Optional ByVal c As Outlook.ContactItem = Nothing, _
                    Optional useBCC As Boolean = False)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":NewEmail"
    
    On Error GoTo ThrowException

    If c Is Nothing Then Set c = selItem
    
    If IsNothing(c) Then
        strTrace = "Incoming OutlookItem was null."
        GoTo ThrowException
    End If
    
    strTrace = "Create a new email."
    Dim m As Outlook.MailItem
    Set m = ut.CreateMailItem
    If IsNothing(m) Then
        strTrace = "Failed to create a new mailItem."
        GoTo ThrowException
    End If
    
    strTrace = "Get the contact's email address(es)"
    Dim eList As ArrayList
    Set eList = GetEmailAddresses(c)
    If eList.Count = 0 Then
        strTrace = "No email addresses found."
        GoTo ThrowException
    End If
    
    Dim delim As String
    delim = ut.GetLocalizedDelimiter
    
    strTrace = "Assemble the email to or BCC string."
    Dim strCat As String
    Dim s As Variant
    For Each s In eList
        strCat = strCat & s & delim
    Next
    strCat = Left(strCat, Len(strCat) - 1)
    
    If useBCC Then
        m.BCC = strCat
    Else
        m.To = strCat
    End If
    
    strTrace = "Assign to a project, if specified."
    If Not IsNothing(p) Then m.Categories = p.Subject
    
    strTrace = "Display the new email."
    m.Display
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Send Email to collection of contacts
Public Sub NewGroupEmail(Optional ByVal p As fmeProject = Nothing, _
                         Optional ByVal useBCC As Boolean = False)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":NewGroupEmail"
    
    On Error GoTo ThrowException
    
    If selItems.Count = 0 Then
        strTrace = "No contacts were selected."
        GoTo ThrowException
    End If
    If selItems.Count = 1 Then
        Call NewEmail(p, selItems.Items(0), useBCC)
        Exit Sub
    End If
    
    strTrace = "Create a new email."
    Dim m As Outlook.MailItem
    Set m = ut.CreateMailItem
    If IsNothing(m) Then
        strTrace = "Failed to create a new mailItem."
        GoTo ThrowException
    End If
    
    strTrace = "Gather email addresses from each contact."
    Dim s As Variant
    Dim aList As ArrayList
    Set aList = New ArrayList
    
    Dim eList As ArrayList
    Dim tc As Outlook.ContactItem
    For Each tc In selItems.Items
        strTrace = "Get the contact's email address(es)"
        Set eList = GetEmailAddresses(tc)
        If eList.Count > 0 Then
            For Each s In eList
                aList.Add s
            Next
        Else
            strTrace = "No email addresses found for '" & tc.FullName & "'."
            LogMessage strTrace, strRoutine
        End If
    Next
           
    Dim delim As String
    delim = ut.GetLocalizedDelimiter
    
    strTrace = "Assemble the email to or BCC string."
    Dim strCat As String
    For Each s In aList
        strCat = strCat & s & delim
    Next
    strCat = Left(strCat, Len(strCat) - 1)
    
    If useBCC Then
        m.BCC = strCat
    Else
        m.To = strCat
    End If
    
    strTrace = "Assign to a project, if specified."
    If Not IsNothing(p) Then m.Categories = p.Subject
    
    strTrace = "Display the new email."
    m.Display
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Schedule a new calendar event with the Contact
Public Sub ScheduleMeeting(Optional p As fmeProject = Nothing, _
                           Optional ByVal c As Outlook.ContactItem = Nothing)
                    
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ScheduleMeeting"
    
    On Error GoTo ThrowException

    If c Is Nothing Then Set c = selItem
    
    If IsNothing(c) Then
        strTrace = "Incoming OutlookItem was null."
        GoTo ThrowException
    End If
    
    strTrace = "Get the contact's email address(es)"
    Dim eList As ArrayList
    Set eList = GetEmailAddresses(c)
    If eList.Count = 0 Then
        strTrace = "No email addresses found."
        GoTo ThrowException
    End If
    
    strTrace = "Create a new appointment."
    Dim a As Outlook.AppointmentItem
    Set a = ut.CreateAppointment
    If IsNothing(a) Then
        strTrace = "Failed to create a new appointmentItem."
        GoTo ThrowException
    End If
    
    strTrace = "Set the appointment as a meeting."
    a.MeetingStatus = olMeeting
    
    strTrace = "Add recipients using their email addresses."
    Dim s As Variant
    For Each s In eList
        a.Recipients.Add s
    Next
    ' Resolve the email addresses to contacts
    a.Recipients.ResolveAll
       
    strTrace = "Assign to a project, if specified."
    If Not IsNothing(p) Then a.Categories = p.Subject
    
    strTrace = "Display the new meeting request."
    a.Display
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

''' Schedule a meeting with a collection of contacts
Public Sub ScheduleGroupMeeting(Optional ByVal p As fmeProject = Nothing)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ScheduleGroupMeeting"
    
    On Error GoTo ThrowException
    
    If selItems.Count = 0 Then
        strTrace = "No contacts were selected."
        GoTo ThrowException
    End If
    If selItems.Count = 1 Then
        Call ScheduleMeeting(p, selItems(0))
        Exit Sub
    End If
    
    strTrace = "Gather email addresses from each contact."
    Dim s As Variant
    Dim aList As ArrayList
    Set aList = New ArrayList
    
    Dim eList As ArrayList
    Dim tc As Outlook.ContactItem
    For Each tc In selItems.Items
        strTrace = "Get the contact's email address(es)"
        Set eList = GetEmailAddresses(tc)
        If eList.Count > 0 Then
            For Each s In eList
                aList.Add s
            Next
        Else
            strTrace = "No email addresses found for '" & tc.FullName & "'."
            LogMessage strTrace, strRoutine
        End If
    Next
           
    strTrace = "Create a new appointment."
    Dim a As Outlook.AppointmentItem
    Set a = ut.CreateAppointment
    If IsNothing(a) Then
        strTrace = "Failed to create a new appointmentItem."
        GoTo ThrowException
    End If
    
    strTrace = "Set the appointment as a meeting."
    a.MeetingStatus = olMeeting
    
    strTrace = "Add recipients using their email addresses."
    For Each s In aList
        a.Recipients.Add s
    Next
    ' Resolve the email addresses to contacts
    a.Recipients.ResolveAll
       
    strTrace = "Assign to a project, if specified."
    If Not IsNothing(p) Then a.Categories = p.Subject
    
    strTrace = "Display the new meeting request."
    a.Display
       
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

                         
End Sub

''' Gets a list of email addresses for the Contact
Public Function GetEmailAddresses(Optional ByVal c As Outlook.ContactItem) As ArrayList
                                  
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetEmailAddresses"
    
    On Error GoTo ThrowException

    If c Is Nothing Then Set c = selItem
    
    If IsNothing(c) Then
        strTrace = "Incoming OutlookItem was null."
        GoTo ThrowException
    End If

    Dim retList As ArrayList
    Set retList = New ArrayList
    
    Dim tmpList As ArrayList
    Set tmpList = New ArrayList
    
    Dim s As Variant
    Dim ts As Variant
    Dim strAddress As String
    
    ' Gather email addresses
    If Len(c.Email1Address) > 0 Then
        strAddress = ut.GetSMTPEmailAddressFromContactEmailEntry( _
                            c.Email1DisplayName, c.Email1Address, c.Email1AddressType)
        If Len(strAddress) > 0 Then tmpList.Add strAddress
    End If
    If Len(c.Email2Address) > 0 Then
        strAddress = ut.GetSMTPEmailAddressFromContactEmailEntry( _
                            c.Email2DisplayName, c.Email2Address, c.Email2AddressType)
        If Len(strAddress) > 0 Then tmpList.Add strAddress
    End If
    If Len(c.Email3Address) > 0 Then
        strAddress = ut.GetSMTPEmailAddressFromContactEmailEntry( _
                            c.Email3DisplayName, c.Email3Address, c.Email3AddressType)
        If Len(strAddress) > 0 Then tmpList.Add strAddress
    End If
    
    ' Deduplicate if needed
    Dim bFnd As Boolean
    If tmpList.Count > 1 Then
        For Each s In tmpList
            bFnd = False
            For Each ts In retList
                If LCase(ts) = LCase(s) Then
                    bFnd = True
                    Exit For
                End If
            Next
            If Not bFnd Then retList.Add s
        Next
    Else
        Set retList = tmpList
    End If

    Set GetEmailAddresses = retList
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetEmailAddresses = New ArrayList

End Function

' - - Managed UI

Public Sub Refresh()
        
    If Not myListView Is Nothing Then RefreshListView sortingColumn, sortingDirection
    
End Sub

' - Supporting Methods

''' Updates the specified Item in the collection
Private Sub UpdateItem(ByVal c As Outlook.ContactItem, Optional idx As Integer = -1)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":UpdateItem"

    If idx < 0 Then
        ' Look for the item
        idx = IsItemInCollection(c)
        If idx < 0 Then
            strTrace = "Cannot update an item that is not in the collection, failed to find: " & c.Subject
            GoTo ThrowException
        End If
    End If

    ' Replace item at the specified idx
    myItems.RemoveAt idx
    myItems.AddItem c
    
    ' Update Managed UI
    UpdateListViewItem c

    Exit Sub

ThrowException:
    LogMessage strTrace, strRoutine

End Sub

''' Loops thru the internal collection of Items and locates the
''' specified item returning its index
''' Returns -1 if not found
Private Function IsItemInCollection(ByVal c As Outlook.ContactItem) As Integer

    Dim iReturn As Integer
    iReturn = -1

    Dim i As Integer
    Dim bFnd As Boolean
    bFnd = False
    For i = 0 To myItems.Items.Count - 1
        Dim tmp As Outlook.ContactItem
        Set tmp = myItems.Items(i)
        If tmp.EntryId = c.EntryId Then
            bFnd = True
            iReturn = i
            Exit For
        End If
    Next
    
    IsItemInCollection = iReturn

End Function

''' <summary>
''' Gets a list of Outlook.ContactItems for a given folder set and query.
''' </summary>
''' <param name="DASLQuery">String: Optional DASL query - if empty returns all items.</param>
''' <param name="AllFolders">Boolean (Optional, defaults to False): if True queries
''' all folders with a default ItemType = Outlook.ContactItem
''' <returns>ArrayList:</returns>
''' <remarks>If an error occurs returns Nothing.</remarks>
Private Function GetOutlookContacts(Optional ByVal DASLQuery As String = "", _
                                    Optional ByVal AllFolders As Boolean = False) As ArrayList

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":GetOutlookContacts"
    
    On Error GoTo ThrowException
                                                
    Dim retItems As ArrayList
    Dim tmpItems As ArrayList
    
    Dim fldr As Outlook.Folder
    Dim fldrList As New ArrayList
    
    If AllFolders Then
        strTrace = "Retrieve items from all Contact folders."
        Set fldrList = ut.GetFoldersByType(olContactItem)
    Else
        strTrace = "Retrieve items from the default folder."
        Set fldr = oNS.GetDefaultFolder(olFolderContacts)
        fldrList.Add fldr
    End If
    
    ' Get List from Outlook
    Set tmpItems = GetOutlookContactsFromFolders(DASLQuery, fldrList)
    
    ' When using SharePoint lists, its possible to get duplicate EntryIDs
    Set retItems = DeDuplicateContactList(tmpItems)
    
    Set GetOutlookContacts = retItems
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetOutlookContacts = Nothing

End Function

''' <summary>
''' Gets a list of Outlook.TaskItems for a given folder set and query.
''' </summary>
''' <param name="DASLQuery">String: Optional DASL query - if empty returns all items.</param>
''' <param name="Folders">List(Of Outlook.MapyFolder) - if nothing returns default task folder.</param>
''' <returns>List(Of Outlook.TaskItem)</returns>
''' <remarks>If an error occurs returns Nothing.</remarks>
Private Function GetOutlookContactsFromFolders(ByVal DASLQuery As String, _
                                            ByVal Folders As ArrayList) As ArrayList
                                     
    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":GetOutlookContactsFromFolders"
    
    On Error GoTo ThrowException
    
    If Folders Is Nothing Then
        strTrace = "A source list of contact folders is required."
        GoTo ThrowException
    End If
    If Folders.Count = 0 Then
        strTrace = "Source folder list was empty - no folders to process."
        GoTo ThrowException
    End If
                                                
    Dim retItems As ArrayList
    Set retItems = New ArrayList
                                                
    Dim fldr As Outlook.Folder
    Dim oItems As Outlook.Items
    
    Dim r As Outlook.Recipient
    Dim oMe As Outlook.ContactItem
    Set oMe = ut.WhoAmI
    
    Dim bJustMe As BookmarkEnum
    bJustMe = True
    Dim bInclude As Boolean
    Dim recEmailAddr As String
    
    For Each fldr In Folders
    
        If Len(DASLQuery) = 0 Then
            Set oItems = fldr.Items
        Else
            Set oItems = fldr.Items.Restrict(DASLQuery)
        End If
    
        strTrace = "Add the Outlook Items collection to the return list."
        Dim o As Object
        For Each o In oItems
            If TypeOf o Is Outlook.ContactItem Then
                ' Capture ContactItem
                Dim c As Outlook.ContactItem
                Set c = o
                
                retItems.Add c
            End If
        Next
    
    Next
                                                
    Set GetOutlookContactsFromFolders = retItems
    Exit Function
        
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetOutlookContactsFromFolders = Nothing
        
End Function

''' Searches thru the list and removes any duplicate EntryIDs
Private Function DeDuplicateContactList(ByVal arList As ArrayList) As ArrayList

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":DeDuplicateContactList"
    
    On Error GoTo ThrowException
    
    If arList Is Nothing Then
        strTrace = "A null list was encountered."
        GoTo ThrowException
    End If
    
    Dim retItems As ArrayList
    Set retItems = New ArrayList
    
    Dim c As Outlook.ContactItem
    Dim tmpC As Outlook.ContactItem
    Dim bFnd As Boolean
    For Each c In arList
        bFnd = False
        For Each tmpC In retItems
            If LCase(tmpC.Subject) = LCase(c.Subject) Then
                bFnd = True
                Exit For
            End If
        Next
        
        If Not bFnd Then retItems.Add c
        
    Next

    Set DeDuplicateContactList = retItems
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set DeDuplicateContactList = Nothing

End Function

' - - ListView Supporting Methods

Private Sub AddListViewItem(ByVal c As Outlook.ContactItem, Optional ByVal idx As Integer = -1)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":AddListViewItem"
    
    If myListView Is Nothing Then
        Exit Sub
    End If
    
    On Error GoTo ThrowException

    ' Check the index
    If idx < 0 Then idx = myListView.ListItems.Count + 1

    ' Add Item to ListView
    Dim ky As String
    ' ky = "k" & Common.GenerateUniqueID(3) & "_" & t.EntryId
    ky = c.EntryId
    Dim li As ListItem
    Set li = myListView.ListItems.Add(idx, ky, c.Subject)
    
        If Len(c.BusinessTelephoneNumber) > 0 Then
            li.SubItems(1) = c.BusinessTelephoneNumber
        Else
            li.SubItems(1) = "Missing"
        End If
        
'    If IsDateNone(t.DueDate) Then
'        li.SubItems(1) = "None"
'    Else
'        li.SubItems(1) = Format(t.DueDate, "mm/dd/yyyy")
'    End If
           
    ' Format the row
    ' FormatLVRow li, c
    
    strTrace = "Added Contact to ListView: " & c.Subject & ", key: " & li.key
    LogMessage strTrace, strRoutine
    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub
Private Sub UpdateListViewItem(ByVal c As Outlook.ContactItem)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":UpdateListViewItem"
    
    If myListView Is Nothing Then
        Exit Sub
    End If
    
    On Error GoTo ThrowException
    
    If IsNothing(c) Then
        strTrace = "A null Item encountered."
        GoTo ThrowException
    End If
    
    Dim li As ListItem
    Set li = FindListViewItem(c)
    If Not IsNothing(li) Then
        ' Update list view here
        li.text = c.Subject
        
        If Len(c.BusinessTelephoneNumber) > 0 Then
            li.SubItems(1) = c.BusinessTelephoneNumber
        Else
            li.SubItems(1) = "Missing"
        End If
            
'        If IsDateNone(t.DueDate) Then
'            li.SubItems(1) = "None"
'        Else
'            li.SubItems(1) = Format(t.DueDate, "mm/dd/yyyy")
'        End If
           
        ' Format the row
        FormatLVRow li, c
        
        strTrace = "Updated ListView for contact: " & c.Subject
    Else
        strTrace = "Add new contact to the ListView."
        AddListViewItem c
    End If
    
    LogMessage strTrace, strRoutine
    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub
Private Sub DeleteListViewItem(ByVal c As Outlook.ContactItem)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":DeleteListViewItem"
    
    If myListView Is Nothing Then
        Exit Sub
    End If
    
    On Error GoTo ThrowException
    
    If IsNothing(c) Then
        strTrace = "A null Item encountered."
        GoTo ThrowException
    End If

    Dim li As ListItem
    Set li = FindListViewItem(c)
    If Not IsNothing(li) Then
        myListView.ListItems.Remove li.Index
    Else
        strTrace = "WARNING: unable to find item: " & c.Subject & " in the listview."
        GoTo ThrowException
    End If
    
    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub
Private Function FindListViewItem(ByVal c As Outlook.ContactItem) As ListItem

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":FindListViewItem"

    If myListView Is Nothing Then
        Exit Function
    End If
    
    On Error GoTo ThrowException
    
    If IsNothing(c) Then
        strTrace = "A null Item encountered."
        GoTo ThrowException
    End If
    
    Dim bFnd As Boolean
    bFnd = False
    
    Dim retItem As ListItem
    Set retItem = Nothing
    
    Dim i As Integer
    For i = 1 To myListView.ListItems.Count
        Dim li As ListItem
        Set li = myListView.ListItems(i)
        If li.key = c.EntryId Then
            Set retItem = li
            Exit For
        End If
    Next
    
    Set FindListViewItem = retItem
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set FindListViewItem = Nothing

End Function

Private Sub RefreshListView(ByVal sortOn As enuSortOn, ByVal sortDir As enuSortDirection)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":RefreshListView"
    
    If myListView Is Nothing Then
        Exit Sub
    End If
    
    On Error GoTo ThrowException
    
    ' Reset these...
    myListView.ColumnHeaders.Clear
    
    selItems.Items.Clear
    Set selItem = Nothing
    
    ' Setup ListView Columns and Configuration
    If myListView.ColumnHeaders.Count = 0 Then
        'Initialize the View
        Dim ch1 As ColumnHeader
        Dim ch2 As ColumnHeader
        With myListView
            Set ch1 = .ColumnHeaders.Add(1, "Name", "Name")
            Set ch2 = .ColumnHeaders.Add(2, "Phone", "Phone")
            
                ' Reset these (in case the ListView control
                '   was modified by another Controller/Manager
            .Checkboxes = True
            .View = lvwReport
        End With
    End If
        
    ' Clear current LV collection
    myListView.ListItems.Clear
    
    ' Add current class collection
    If myItems.Count > 0 Then
    
        ' Sort the internal list
        Dim strSort As String
        strSort = "Subject"
        If sortingColumn = DueDate Then strSort = "Subject"
        
        Dim collSort As New SortCollection
        collSort.Sort strSort, myItems.Items, sortingDirection
          
        ' Load the ListView
        Dim i As Integer
        For i = 0 To myItems.Count - 1
            Dim c As Outlook.ContactItem
            Set c = myItems.Item(i)
                                 
            ' Add Item to ListView
            AddListViewItem c, i + 1
            
        Next
        
        Call ResizeLVColumns
        
    End If
    
    strTrace = "Items: " & myItems.Count & " contacts..."
    Exit Sub
    
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
    
    On Error GoTo ThrowException
    
    Dim strColumnWidths As String
    strColumnWidths = "60;40"
    
    Dim totWidth As Integer
    totWidth = myListView.Width
    
    ' if scrollbar present, make space
    Dim bScrollbar As Boolean
    With myListView
        bScrollbar = (.font.SIZE + 4 + 1) * .ListItems.Count > .Height
    End With
    
    If bScrollbar Then totWidth = totWidth - 15
    
    Dim widths() As String
    widths = Split(strColumnWidths, ";")
    
    Dim i As Integer
    For i = LBound(widths) To UBound(widths)
        Dim colWidth As Integer
        colWidth = CInt((widths(i) / 100) * totWidth) - 1
        myListView.ColumnHeaders(i + 1).Width = colWidth
    Next
    
    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Private Function LV_GetItemAt(ByVal x As stdole.OLE_XPOS_PIXELS, _
                              ByVal y As stdole.OLE_YPOS_PIXELS, _
                     Optional ByVal factor As Integer = 15) As ListItem


    ' Convert Pixels to TWIPS
    ' - .net uses Pixels, VBA uses TWIPS for ListView and TreeView (OLE_PIXELS?)
    ' - "on most computers 1 pixel = 15 TWIPS"
    '    https://stackoverflow.com/questions/36442535/vba-drag-drop-from-treeview-to-listview-listview-to-treeview-activex-controls
    Dim xInt As Single
    xInt = x
    Dim yInt As Single
    yInt = y
           
    Dim li As ListItem
    Set li = myListView.HitTest(x * factor, y * factor)
        
    Set LV_GetItemAt = li
        
End Function









