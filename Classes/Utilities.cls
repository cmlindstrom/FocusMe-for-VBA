VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Utilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' Fields

Private Const rootClass As String = "Utilities"

Dim app As Outlook.Application
Dim oNS As Outlook.NameSpace
Dim stgs As Settings

Private Const ReferenceIDLength As Integer = 8
Private Const ReferenceIDString As String = "FSRefID:"
Private Const ReferenceIDUserProperty As String = "FMEReferenceID" ' Across device and service unique identifier
    ' Public Const ParentReferenceIDUserProperty As String = "FMEParentReferenceID"
Private Const MessageIDString As String = "FSMsgID:"
    '  Public Const SharePointIDProperty As String = "FMESPID"

Private Const LinkedProjectsUserProperty As String = "FMELinkedProjects"

Private Const LinkedContactsUserProperty As String = "FMELinkedContacts"
    
''' Properties

Public Property Get Application() As Outlook.Application
    Set Application = app
End Property

''' Constructor

Private Sub Class_Initialize()

    ' Set app = oNS.Application
    Set app = ThisOutlookSession.Application
    Set oNS = app.GetNamespace("MAPI") 'GetNamespace("MAPI")
    
    Set stgs = New Settings

End Sub

''' Methods

' App Level ---

    ' Returns True if the user has an Exchange Account loaded within their profile
    Public Function IsUsingExchange() As Boolean
    
        Dim bReturn As Boolean
        bReturn = False

        Dim myAccounts As Outlook.Account
        Set myAccounts = oNS.Accounts
            
        If Not IsNothing(myAccounts) Then
            Dim a As Outlook.Account
            For Each a In myAccounts
                If a.AccountType = OlAccountType.olExchange Then
                    bReturn = True
                    Exit For
                End If
            Next
        Else
            strTrace = "Failed to gather the Outlook Accounts."
        End If

        IsUsingExchange = bReturn
    
    End Function
    
    ''' Returns the logged in user's contact information
    Public Function WhoAmI() As Outlook.ContactItem
    
        Dim oMe As Outlook.ContactItem
        Dim oExUser As Outlook.ExchangeUser
    
        Set oMe = oNS.Session.CurrentUser.AddressEntry.GetContact
        Set oExUser = oNS.Session.CurrentUser.AddressEntry.GetExchangeUser
        
        If oMe Is Nothing And Not oExUser Is Nothing Then
            Set oMe = app.CreateItem(olContactItem)
            With oExUser
                oMe.FileAs = .Name
                oMe.FirstName = .FirstName
                oMe.LastName = .LastName
                oMe.FullName = .FirstName & " " & .LastName
                oMe.Email1Address = .PrimarySmtpAddress
                oMe.Email1AddressType = "smtp"
                oMe.Email2Address = .Address
                oMe.Email2AddressType = .Type
                oMe.CompanyName = .CompanyName
                oMe.Department = .Department
                oMe.OfficeLocation = .OfficeLocation
                oMe.BusinessAddress = .StreetAddress
                oMe.BusinessAddressCity = .City
                oMe.BusinessAddressState = .StateOrProvince
                oMe.BusinessAddressPostalCode = .PostalCode
                oMe.BusinessTelephoneNumber = .BusinessTelephoneNumber
                oMe.MobileTelephoneNumber = .MobileTelephoneNumber
            End With
        End If
        
        Set WhoAmI = oMe
        
    End Function

    ''' <summary>
    ''' Returns the current Outlook profile name
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Function GetCurrentProfileName() As String
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":GetCurrentProfileName"
        
        On Error GoTo ThrowException
               
        strTrace = "Interrogating Namespace."
        GetCurrentProfileName = oNS.CurrentProfileName
        Exit Function
        
ThrowException:
        GetCurrentProfileName = "Error"
        LogMessageEx strTrace, err
    
    End Function
    
' Master Categories Management

    ''' <summary>
    ''' Retrieves the ID from the Master Category List for the specified CategoryName
    ''' </summary>
    ''' <param name="CategoryName">Name of the category to find</param>
    ''' <returns>String: Identifer if found, otherwise an empty string</returns>
    ''' <remarks></remarks>
    Public Function GetMCLID(ByVal CategoryName As String) As String
    
        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":GetMCLID"
        
        On Error GoTo ThrowException
        
        If Len(CategoryName) = 0 Then
            strTrace = "No category name specified."
            GoTo ThrowException
        End If
        
        Dim olCategory As Outlook.Category
        Dim olCategories As Outlook.Categories

        Dim vRetVal As String

        Set olCategories = oNS.Categories
        If olCategories.count > 0 Then
            For Each olCategory In olCategories
                If LCase(olCategory.Name) = LCase(CategoryName) Then
                    vRetVal = CStr(olCategory.CategoryID)
                    Exit For
                End If
            Next
        Else
            strTrace = "No categories exist - MCL was empty."
            LogMessage strTrace, strRoutine
        End If
        
        If Len(vRetVal) = 0 Then
            strTrace = "No MCL category found for '" & CategoryName & "'."
            LogMessage strTrace, strRoutine
        End If
        
        GetMCLID = vRetVal
        Exit Function

ThrowException:
        LogMessageEx strTrace, err, strRoutine

    End Function
      
    ''' <summary>
    ''' Add an entry into the Master Category List
    ''' </summary>
    ''' <param name="CategoryName">Name of the Category</param>
    ''' <param name="CategoryColor">Category Color</param>
    ''' <returns>String: MCL ID</returns>
    ''' <remarks></remarks>
    Public Function AddtoMCL(ByVal CategoryName As String, _
                             ByVal CategoryColor As Integer) As String

        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":AddtoMCL"
        
        On Error GoTo ThrowException

        Dim olCategory As Outlook.Category
        Set olCategory = Nothing
        Dim bretVal As String

        ' Check to see if category already exists
        bretVal = GetMCLID(CategoryName)
        If Len(bretVal) = 0 Then
            ' Didn't find an existing category with the specified name
            Set olCategory = oNS.Categories.Add(CategoryName, CategoryColor)
            bretVal = olCategory.CategoryID
            
            strTrace = "Added a new category for '" & CategoryName & "'."
            LogMessage strTrace, strRoutine
        End If

        AddtoMCL = bretVal
        Exit Function
        
ThrowException:
        LogMessageEx strTrace, err, strRoutine

    End Function

    ''' <summary>
    ''' Updates or adds the project's subject to the Master Category List.
    ''' </summary>
    ''' <param name="prProject">Project</param>
    ''' <returns>String: Category ID in the MCL</returns>
    ''' <remarks></remarks>
    Public Function UpdateMCL(ByVal mclId As String, _
                            ByVal Name As String, _
                            ByVal Color As Integer) As String
    
        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":UpdateMCL"
        
        On Error GoTo ThrowException
        
        If Len(mclId) = 0 Then
            strTrace = "Missing Category unique id."
            GoTo ThrowException
        End If

        Dim olCategory As Outlook.Category
        Dim bretVal As String
        Dim bFound As Boolean
        bFound = False

        ' There seems to be some issue with updating the name of an MCL name, therefore
        ' need to remove and add versus simply updating the specified category
        
        strTrace = "Processing category '" & Name & "' for id: " & mclId & ", Color: " & Color & "."
        LogMessage strTrace, strRoutine

        Dim i As Integer
        If oNS.Categories.count > 0 Then
            For i = 1 To oNS.Categories
                Set olCategory = oNS.Categories(i)
                If olCategory.CategoryID = mclId Then
                    ' Check for substantive Change
                    If (olCategory.Name = Name) And _
                            (olCategory.Color = Color) Then
                        ' Nothing changed
                        strTrace = "No change detected."
                        LogMessage strTrace, strRoutine
                        
                        bretVal = olCategory.CategoryID
                        GoTo Success
                    End If
                    
                    strTrace = "Removing category, mclId: " & mclId
                    LogMessage strTrace, strRoutine
                    
                    oNS.Categories.Remove i
                    
                    strTrace = "Removed category, mclId: " & mclId
                    LogMessage strTrace, strRoutine
                    
                    bFound = True
                    Exit For
                End If
            Next
                
            Set olCategory = oNS.Categories.Add(Name, Color)
            bretVal = olCategory.CategoryID
            If bFound Then
                strTrace = "Updated the current category for '" & Name & "'."
                LogMessage strTrace, strRoutine
            Else
                strTrace = "Added a new category for '" & Name & "'."
                LogMessage strTrace, strRoutine
            End If
        
        Else
            strTrace = "No categories currently exist in the MCL."
            GoTo ThrowException
        End If

Success:

        UpdateMCL = bretVal
        Exit Function
            
ThrowException:
        LogMessageEx strTrace, err, strRoutine

    End Function
    
    ''' <summary>
    ''' Removes a specified Category from the Master Category List.
    ''' </summary>
    ''' <param name="CategoryName">Name of the Category</param>
    ''' <returns>Boolean:</returns>
    ''' <remarks></remarks>
    Public Function RemoveFromMCL(ByVal CategoryName As String) As Boolean
    
        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":RemoveFromMCL"
        
        On Error GoTo ThrowException
        
        Dim olCategory As Outlook.Category

        Dim bretVal As Boolean
        bretVal = True

        If oNS.Categories.count > 0 Then
            For Each olCategory In oNS.Categories
                If LCase(CategoryName) = LCase(olCategory.Name) Then
                    oNS.Categories.Remove olCategory.CategoryID
                    Exit For
                End If
            Next
        Else
            strTrace = "No Outlook categories found matching '" & CategoryName & "'."
            LogMessage strTrace, strRoutine
            bretVal = False
        End If
        
        RemoveFromMCL = bretVal
        Exit Function
        
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        
    End Function
    
    ''' Returns the Global Delimiter that separates the
    ''' Categories within an Outlook Categories Property value
    Public Function GetLocalizedDelimiter() As String
        
        GetLocalizedDelimiter = ";"
        
    End Function

    ''' Returns the name of the color matching the Outlook Colors
    Public Function GetCategoryColorName(ByVal iColorNumber As Integer) As String
    
        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":GetCategoryColorName"
        
        On Error GoTo ThrowException
            
        Dim strReturn As String
        strReturn = ""
                
        Select Case iColorNumber
            Case 0 ' No Color
                strReturn = "None"
            Case 1 ' Red
                strReturn = "Red"
            Case 2 ' Orange
                strReturn = "Orange"
            Case 3 ' Peach
                strReturn = "Peach"
            Case 4 ' Yellow
                strReturn = "Yellow"
            Case 5 ' Green
                strReturn = "Green"
            Case 6 ' Teal
                strReturn = "Teal"
            Case 7 ' Olive
                strReturn = "Olive"
            Case 8 ' Blue
                strReturn = "Blue"
            Case 9 ' Purple
                strReturn = "Purple"
            Case 10 ' Maroon
                strReturn = "Maroon"
            Case 11 ' Steel
                strReturn = "Steel"
            Case 12 ' Dark Steel
                strReturn = "Dark Steel"
            Case 13 ' Gray
                strReturn = "Gray"
            Case 14 ' Dark Gray
                strReturn = "Dark Gray"
            Case 15 ' Black
                strReturn = "Black"
            Case 16 ' Dark Red
                strReturn = "Dark Red"
            Case 17 ' Dark Orange
                strReturn = "Dark Orange"
            Case 18 ' Dark Peach
                strReturn = "Dark Peach"
            Case 19 ' Dark Yellow
                strReturn = "Dark Yellow"
            Case 20 ' Dark Green
                strReturn = "Dark Green"
            Case 21 ' Dark Teal
                strReturn = "Dark Teal"
            Case 22 ' Dark Olive
                strReturn = "Dark Olive"
            Case 23 ' Dark Blue
                strReturn = "Dark Blue"
            Case 24 ' Dark Purple
                strReturn = "Dark Purple"
            Case 25 ' Dark Maroon
                strReturn = "Dark Maroon"
            Case Else
                strReturn = "Unknown"
        End Select

        GetCategoryColorName = strReturn
        Exit Function
        
ThrowException:
        LogMessageEx strTrace, err, strRoutine

    End Function

' Folders ---

    Public Function SelectOutlookFolder() As Outlook.Folder
    
      Dim myFolder As Outlook.Folder
      Set myFolder = oNS.PickFolder
      Set SelectOutlookFolder = myFolder
          
    End Function
    
    ''' Returns True if the designated item is in a Folder whose name
    ''' matches the string specified in folderName
    Public Function IsItemParent(ByVal myItem As Object, ByVal folderName As String) As Boolean
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":IsItemParent"
        
        On Error GoTo ThrowException
        
        Dim bReturn As Boolean
        bReturn = False
    
        Dim srcFolder As String
        Dim parentFolder As Outlook.Folder
        Set parentFolder = myItem.Parent
        If Not parentFolder Is Nothing Then
            srcFolder = parentFolder.FolderPath
            If Len(srcFolder) > 0 Then
                If Contains(LCase(folderName), LCase(srcFolder)) Then
                    strTrace = "Found '" & folderName & "' in the item's parent folder: " & srcFolder & "."
                    bReturn = True
                Else
                    strTrace = "Item's folder path did not contain the search string: '" & folderName & "'."
                    GoTo ThrowException
                End If
            Else
                strTrace = "Item's folder path was not specified."
                GoTo ThrowException
            End If
        Else
            strTrace = "No folder path was returned for the specified item."
            GoTo ThrowException
        End If
            
        IsItemParent = bReturn
        Exit Function
        
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        IsItemParent = False
    
    End Function

    Public Function MoveToArchive(ByVal myItem As Object) As Boolean
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":MoveToArchive"
        
        On Error GoTo ThrowException
        
        Dim oMail As Outlook.MailItem
    
        Dim fldrPath As String
        fldrPath = stgs.DestinationFolder
    
        Dim destFldr As Outlook.Folder
        ' Check for project folder (not implemented), if not, send to destination folder
        Set destFldr = GetFolderByPath(fldrPath)
        If Not IsNothing(destFldr) Then
            ' Check to see if item already in the destination folder
            Dim parentFolder As Outlook.Folder
            Set parentFolder = myItem.Parent
            If Not parentFolder.FolderPath = destFldr.FolderPath Then
                If myItem.UnRead Then
                    myItem.UnRead = False
                    myItem.Save
                End If
                myItem.Move destFldr
            Else
                strTrace = "Item already filed."
                LogMessage strTrace, strRoutine
            End If
        Else
            strTrace = "Failed to find destination folder: " & DestinationFolder
            LogMessage strTrace, strRoutine
        End If
        
        MoveToArchive = True
        Exit Function
        
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        MoveToArchive = False
    
    End Function
    
    '''  Selects an Outlook folder for the ActiveExplorer
    Public Sub SelectFolderFromPath(ByVal oFolderPath As String)
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":SelectFolderFromPath"
        
        On Error GoTo ThrowException
        
        If Len(oFolderPath) = 0 Then
            strTrace = "No folder path was specified."
            GoTo ThrowException
        End If
        
        Dim f As Outlook.Folder
        Set f = GetFolderByPath(oFolderPath)
        If IsNothing(f) Then
            strTrace = "Failed to retrieve the Outlook folder form path: " & oFolderPath
            GoTo ThrowException
        End If
        
        SelectFolder f
        
        Exit Sub
        
ThrowException:
        LogMessageEx strTrace, err, strRoutine
    
    End Sub
    
    '''  Selects an Outlook folder for the ActiveExplorer
    Public Sub SelectFolder(ByVal f As Outlook.Folder)
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":SelectFolder"
        
        On Error GoTo ThrowException
        
        If IsNothing(f) Then
            strTrace = "Folder object was null."
            GoTo ThrowException
        End If
        
        ' ThisOutlookSession.ActiveExplorer.CurrentFolder = f
        
        ' app.ActiveExplorer.CurrentFolder = oNS.GetDefaultFolder(olFolderContacts)

        ' Creates a new Outlook Explorer Window with the specified Folder displayed
        f.Display

        Exit Sub
        
ThrowException:
        LogMessageEx strTrace, err, strRoutine
    
    End Sub

    ''' Returns an Outlook Folder using the folder's path
    Public Function GetFolderByPath(ByVal strPath As String) As Outlook.Folder
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":GetFolderByPath"
        
        On Error GoTo ThrowException
        
        Dim retList As Collection
        Set retList = GetFoldersWithName(strPath)
        If IsNothing(retList) Then
            strTrace = "Folder query failed."
            GoTo ThrowException
        End If
        
        Dim retFolder As Outlook.Folder
        
        If retList.count > 0 Then
            Set retFolder = retList.Item(1)
            
            If retList.count > 1 Then
                strTrace = "WARNING: found multiple folders for query: " & strPath
                LogMessage strTrace, strRoutine
            End If
        Else
            strTrace = "Found zero matches for: " & strPath
            GoTo ThrowException
        End If
        
        Set GetFolderByPath = retFolder
        Exit Function
        
ThrowException:
        LogMessage strTrace, strRoutine
        Set GetFolderByPath = Nothing
    
    End Function

    ''' <summary>
    ''' Returns a list of folders that contain the specified name in the folder name.
    ''' </summary>
    ''' <param name="strFolderName">String:</param>
    ''' <returns>Collection of(Outlook.Folder):</returns>
    ''' <remarks>Case insensitive - peruses all Outlook Accounts.</remarks>
    Public Function GetFoldersWithName(ByVal strFolderName As String) As Collection

        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":GetFoldersWithName"
        
        On Error GoTo ThrowException
        
        Set retList = New Collection
        
        Dim oStores As Outlook.Stores
        Set oStores = app.Session.Stores
        
        For Each oStore In oStores
            strTrace = "Evaluating store: " & oStore.DisplayName
            Dim myList As Collection
            Set myList = GetFoldersWithNameFromStore(strFolderName, oStore)
            If Not IsNothing(myList) And myList.count > 0 Then
                For Each f In myList
                    retList.Add (f)
                Next
            End If
        Next
        
        Set GetFoldersWithName = retList
        Exit Function
        
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        Set GetFoldersWithName = New Collection

    End Function
    
    ''' <summary>
    ''' Returns a list of folders from the specific Outlook Store that contains
    ''' the specified name in the name of the folder.
    ''' </summary>
    ''' <param name="Name">String:</param>
    ''' <param name="oStore">Outlook.Store:</param>
    ''' <returns>List(Of Redemption.RDOFolder:</returns>
    ''' <remarks>Case Insensitive</remarks>
    Public Function GetFoldersWithNameFromStore(ByVal strFolderName As String, ByVal oStore As Outlook.Store) As Collection
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":GetFoldersWithNameFromStore"
    
        Dim retFldrs As New Collection

        Dim oInbox As Outlook.Folder
        Set inbox = oStore.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
        Dim oRootFolder As Outlook.Folder
        Set oRootFolder = inbox.Parent
        If Not IsNothing(oRootFolder) Then
        
            If Len(oRootFolder.Name) > 0 Then
                strTrace = "Evaluating folder: " & oRootFolder.Name & " in store: " & oStore.DisplayName
                LogMessage strTrace, strRoutine
        
                Set retFldrs = SearchFolderByName(oRootFolder, strFolderName)
            Else
            
                strTrace = "Evaluating folder: " & oRootFolder.Name & " in store: " & oStore.DisplayName
                LogMessage strTrace, strRoutine
        
                Set retFldrs = SearchFolderByName(oRootFolder, strFolderName)
                
               ' strTrace = "Ignoring store: " & oStore.DisplayName
               ' LogMessage strTrace, strRoutine
            End If
            
        End If

        Set GetFoldersWithNameFromStore = retFldrs
    
    End Function
    
    ''' Recurses the Folder tree, starting with the specified Outlook.Folder
    ''' collecting any folders with a name that matches the specified name
    Public Function SearchFolderByName(ByVal myFolder As Outlook.Folder, _
                                       ByVal Name As String) As Collection
                                       
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":SearchFolderByName"
        
        On Error GoTo ThrowException
        
        If IsNothing(myFolder) Then
            strTrace = "A null Outlook.Folder encountered."
            GoTo ThrowException
        End If
        
        ' Dim retFldrs As New Collection
        Set retFldrs = New Collection
        
        Dim srcName As String
        Dim tgtName As String

 ' Dim fldr As Outlook.Folder
 ' fldr.DefaultItemType

        If myFolder.Folders.count > 0 Then
            ' Start search at the top of the hierarchy
            srcName = UCase(Name)
            For Each f In myFolder.Folders
            
                ' Check for 'special' folders
                If Contains("COMMON_VIEWS", UCase(f.Name)) Then
                    strTrace = "Reserved Folder, ignoring: " & f.Name & "."
                    LogMessage strTrace, strRoutine
                    GoTo SkipFolder1
                End If
            
                ' Check for problem children
                If IsNothing(f.Parent) Then
                    strTrace = "System folder encountered, no PARENT."
                    LogMessage strTrace, strRoutine
                    GoTo SkipFolder1
                End If
                If Len(f.Name) = 0 Then
                    strTrace = "Invalid folder encountered, no NAME."
                    LogMessage strTrace, strRoutine
                    GoTo SkipFolder1
                End If
                                                
                ' Make case insensitive
                strTrace = "Evaluating folder: " & f.Name
                LogMessage strTrace, strRoutine
                
                tgtName = UCase(f.Name)
                
                ' Check the folder path
                Dim pathName As String
                pathName = UCase(f.FolderPath)
                If Contains(srcName, pathName) Then
                    retFldrs.Add (f)
                End If
                
                strTrace = "Look to see if current folder has sub-folders."
                If f.Folders.count > 0 Then
                    Dim fldrs As Collection
                    Set fldrs = SearchFolderByName(f, Name)
                    If Not IsNothing(fldrs) And fldrs.count > 0 Then
                        For Each fd In fldrs
                            retFldrs.Add (fd)
                        Next
                    End If
                Else
                    strTrace = "Folder has no sub-folders."
                End If
SkipFolder1:
            Next
        Else
            strTrace = "Evaluating folder: " & myFolder.Name
            ' Make case insensitive
            strTrace = "Evaluating folder: " & myFolder.Name
            tgtName = UCase(myFolder.Name)
            
            strTrace = "Looking for: " & srcName & " in folder: " & myFolder.FolderPath
            LogMessage strTrace, strRoutine

            ' Check the path
            pathName = UCase(myFolder.FolderPath)
            If Contains(srcName, pathName) Then
                retFldrs.Add (f)
            End If
        End If

SkipOut:
        Set SearchFolderByName = retFldrs
        Exit Function
        
ThrowException:
        Set SearchFolderByName = Nothing
                                       
    End Function
       
    ''' Returns a collection of fmeFolders based on the current
    ''' user's folder tree
    Public Function IndexFolders() As fmeFolders
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":IndexFolders"
        
        On Error GoTo ThrowException
        
        Dim retColl As fmeFolders
        Dim myColl As fmeFolders
        
        Set retColl = New fmeFolders
        
        Dim oStores As Outlook.Stores
        Set oStores = app.Session.Stores
        
        Dim f As fmeFolder
        For Each oStore In oStores
            strTrace = "Evaluating store: " & oStore.DisplayName
            LogMessage strTrace, strRoutine
            
            Set myColl = IndexFoldersInStore(oStore)
            If Not IsNothing(myColl) And myColl.count > 0 Then
                For Each f In myColl.Items
                    retColl.AddItem f
                Next
            End If
        Next
        
        strTrace = "Found " & retColl.count & " folders while indexing..."
        LogMessage strTrace, strRoutine
        
        Set IndexFolders = retColl
        Exit Function
        
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        Set IndexFolders = Nothing
    
    End Function
    
    ''' Returns a collection of Folders (of fmeFolder) for the specified
    ''' Outlook Store.
    Public Function IndexFoldersInStore(ByVal oStore As Outlook.Store) As fmeFolders
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":IndexFoldersInStore"
        
        On Error GoTo ThrowException
        
        Dim retColl As fmeFolders
        
        Dim oInbox As Outlook.Folder
        Set inbox = oStore.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
        Dim oRootFolder As Outlook.Folder
        Set oRootFolder = inbox.Parent
        If Not IsNothing(oRootFolder) Then
        
            strTrace = "Evaluating folder: " & oRootFolder.Name & " in store: " & oStore.DisplayName
            LogMessage strTrace, strRoutine
        
            Set retColl = IndexFoldersWithParent(oRootFolder)
            If retColl Is Nothing Then
                strTrace = "Failed to interrogate Store: " & oStore.DisplayName & _
                    ", returning an empty collection."
                LogMessage "WARNING: " & strTrace, strRoutine
                Set retColl = New fmeFolders
            Else
                strTrace = "Found " & retColl.count & " folders in " & oRootFolder.Name
                LogMessage strTrace, strRoutine
            End If
        
        Else
            strTrace = "Failed to find the default Inbox in Store: " & oStore.DisplayName
            GoTo ThrowException
        End If
        
        Set IndexFoldersInStore = retColl
        Exit Function
        
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        Set IndexFoldersInStore = Nothing
    
    End Function
    
    ''' Recurses a root folder, indexing all folder children
    Public Function IndexFoldersWithParent(ByVal oParentFolder As Outlook.Folder) As fmeFolders
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":IndexFoldersWithParent"
        
        If oParentFolder Is Nothing Then
            strTrace = "Encountered a null Outlook Folder."
            GoTo ThrowException
        End If
        
        On Error GoTo ThrowException
        
        Dim retColl As fmeFolders
        Set retColl = New fmeFolders
        
        Dim fldr As fmeFolder
        Dim tFldr As fmeFolder
        Dim fldrs As fmeFolders
        
        If oParentFolder.Folders.count > 0 Then
        
            Dim oTF As Outlook.Folder
            For Each oTF In oParentFolder.Folders
            
                ' Check for 'special' folders
                If Contains("COMMON_VIEWS", UCase(oTF.Name)) Then
                    strTrace = "Reserved Folder, ignoring: " & oTF.Name & "."
                    LogMessage strTrace, strRoutine
                    GoTo SkipFolder1
                End If
            
                ' Check for problem children
                If IsNothing(oTF.Parent) Then
                    strTrace = "System folder encountered, no PARENT."
                    LogMessage strTrace, strRoutine
                    GoTo SkipFolder1
                End If
                If Len(oTF.Name) = 0 Then
                    strTrace = "Invalid folder encountered, no NAME."
                    LogMessage strTrace, strRoutine
                    GoTo SkipFolder1
                End If
                              
                ' Add folder to the collection
                Set fldr = New fmeFolder
                fldr.Fill oTF
                retColl.AddItem fldr
                
                strTrace = "Added folder to collection: " & oTF.Name & _
                    " folder count: " & retColl.count
                LogMessage strTrace, strRoutine
              
                ' Recurse if folder has sub-folders
                If oTF.Folders.count > 0 Then
                    strTrace = "Recursing folder: " & oTF.Name
                    LogMessage strTrace, strRoutine
                
                    Set fldrs = IndexFoldersWithParent(oTF)
                    If Not IsNothing(fldrs) Then
                        If fldrs.count > 0 Then
                            For Each tFldr In fldrs.Items
                                retColl.AddItem tFldr
                            Next
                        End If
                    End If
                
                Else
                    strTrace = "Folder (" & oTF.Name & ") has no children."
                End If
SkipFolder1:
            Next
        
        Else
            strTrace = "Adding parent folder with no children to collection: " & oTF.Name
            LogMessage strTrace, strRoutine
                
            Set fldr = New fmeFolder
            fldr.Fill parentFolder
            retColl.AddItem fldr
        End If
        
        strTrace = "Returning a collection of " & retColl.count & " folders."
        LogMessage strTrace, strRoutine
        
        Set IndexFoldersWithParent = retColl
        Exit Function
    
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        Set IndexFoldersWithParent = Nothing
    
    End Function
    
    ''' Retrieves a list of Outlook.Folders that contain
    ''' items of the specified type (typ) from the Index folder set
    Public Function GetFoldersByType(ByVal typ As Outlook.OlItemType) As ArrayList
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":GetFoldersByType"
        
        On Error GoTo ThrowException
        
        Dim retList As New ArrayList
        
        Dim ldb As New dsDataStore
        ldb.Connect
        
        Dim arList As ArrayList
        Set arList = ldb.GetEntireCollection("Folder")
        If arList Is Nothing Then
            strTrace = "Failed to query the datastore."
            GoTo ThrowException
        End If
        If arList.count = 0 Then
            strTrace = "Found an empty Folder collection - try indexing via the Options dialog."
            GoTo ThrowException
        End If
        
        Dim oFldr As Outlook.Folder
        Dim fmeFldr As fmeFolder
        
        For Each fmeFldr In arList
            strTrace = "Check default itemType for fme Folder: " & fmeFldr.Name
            If fmeFldr.itemType = typ Then
                On Error Resume Next
                strTrace = "Retrieving Outlook Folder: " & fmeFldr.path
                Set oFldr = oNS.GetFolderFromID(fmeFldr.EntryId)
                If oFldr Is Nothing Then
                    strTrace = "Outlook GetFolder call failed for " & fmeFldr.path
                    LogMessage strTrace, strRoutine
                    GoTo SkipItem
                End If
                strTrace = "Add Outlook Folder to the return list."
                retList.Add oFldr
            End If
SkipItem:
        Next
        
        Set GetFoldersByType = retList
        Exit Function
        
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        Set GetFoldersByType = Nothing
    
    End Function
    
    ''' Retrieves a list of Outlook.Folders that contain
    ''' MailItems - via the Indexed Folder set.
    Public Function GetMailFolders() As ArrayList

        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":GetMailFolders"
        
        On Error GoTo ThrowException
               
        Set GetMailFolders = GetFoldersByType(olMailItem)
        Exit Function
        
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        Set GetMailFolders = Nothing
    
    End Function
    
    ''' Retrieves a list of Outlook.Folders that contain
    ''' AppointmentItems - via the Indexed Folder set.
    Public Function GetCalendarFolders() As ArrayList

        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":GetCalendarFolders"
        
        On Error GoTo ThrowException
               
        Set GetCalendarFolders = GetFoldersByType(olAppointmentItem)
        Exit Function
        
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        Set GetCalendarFolders = Nothing
    
    End Function
    
    ''' Retrieves a list of Outlook.Folders that contain MailItems
    ''' - temporary uses Inbox and DestinationFolder
    Public Function GetMailFoldersOld() As ArrayList
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":GetMailFolders"
        
        On Error GoTo ThrowException
        
        Dim retList As New ArrayList
    
        ' Get default inbox
        Dim fldr As Outlook.Folder
        Set fldr = oNS.GetDefaultFolder(olFolderInbox)
        If Not IsNothing(fldr) Then retList.Add fldr
        
        ' Get SentItems
        Dim sentFldr As Outlook.Folder
        Set sentFldr = oNS.GetDefaultFolder(olFolderSentMail)
        If Not IsNothing(sentFldr) Then retList.Add sentFldr
    
        ' Get Archive Folder
        Dim destFldr As Outlook.Folder
        Set destFldr = GetFolderByPath(stgs.DestinationFolder)
        If Not IsNothing(destFldr) Then retList.Add destFldr
        
        Set GetMailFoldersOld = retList
        
        Exit Function
        
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        Set GetMailFoldersOld = Nothing
        
    End Function
                                               


' Item ---

    ''' <summary>
    ''' Retrieves an Outlook item from the datastore using its EntryID.
    ''' </summary>
    ''' <param name="ID">String: OutlookObject.EntryID</param>
    ''' <returns>Object: e.g. Outlook.TaskItem, Outlook.AppointmentItem, etc.</returns>
    ''' <remarks></remarks>
    Public Function GetOutlookItemFromID(ByVal id As String) As Object
        
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":LinkObjects"
        
        On Error GoTo ThrowException
        
        Dim retObject As Object
        Set retObject = Nothing
        
        Set retObject = oNS.GetItemFromId(id)
        
        ' Return Updated Object
        Set GetOutlookItemFromID = retObject
        Exit Function
        
ThrowException:
        Set GetOutlookItemFromID = Nothing
        LogMessage strTrace & " " & err.Description, strRoutine
        
    End Function

    ''' <summary>
    ''' Copies the attachments from the SourceItem to the TargetItem
    ''' </summary>
    ''' <param name="objSourceItem">Outlook Object</param>
    ''' <param name="objTargetItem">Outlook Object</param>
    ''' <remarks></remarks>
    Public Function CopyAttachments(objSourceItem As Object, objTargetItem As Object) As Object

        Dim fso As Object
        Dim fldTemp As Object

        Dim strPath As String, strFile As String

        Set fso = CreateObject("Scripting.FileSystemObject")

        fldTemp = fso.GetSpecialFolder(2)
        strPath = fldTemp.path & "\"
        For Each objAtt In objSourceItem.Attachments
            If objAtt.Type <> Interop.Outlook.OlAttachmentType.olOLE Then
                strFile = strPath & objAtt.FileName
                objAtt.SaveAsFile (strFile)
                objTargetItem.Attachments.Add strFile, , , objAtt.DisplayName
                fso.DeleteFile (strFile)
            End If
        Next

        fldTemp = Nothing
        fso = Nothing

        Set CopyAttachments = objTargetItem
    
    End Function

    ''' <summary>
    ''' Links two Outlook objects together using the specified LinkStrategy
    ''' </summary>
    ''' <param name="LinkBy">enuLinkStrategy:</param>
    ''' <param name="objIn">Outlook Object: Incoming object</param>
    ''' <param name="subject">String:</param>
    ''' <param name="body">String:</param>
    ''' <param name="objOut">Outlook Object Reference: Destination object</param>
    Public Function LinkObjects(ByVal linkBy As enuLinkStrategy, _
                                ByVal objIn As Object, _
                                ByVal Subject As String, _
                                ByVal body As String, _
                                ByVal objOut As Object) As Object
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":LinkObjects"
        
        ' Error Checking
        If IsNothing(objIn) Then
            strTrace = "A null source item was encountered."
            Common.LogMessage strTrace, strRoutine
            GoTo ThrowException
        End If
        If IsNothing(objOut) Then
            strTrace = "A null target item was encountered."
            Common.LogMessage strTrace, strRoutine
            GoTo ThrowException
        End If
        
        Dim retObject As Object
        Set retObject = objOut
        
        objOut.Subject = Subject
        objOut.body = body

        Select Case linkBy
            Case enuLinkStrategy.text
                Set objOut = CopyAttachments(objIn, objOut)
            Case enuLinkStrategy.embed
                objOut.Attachments.Add objIn, Outlook.OlAttachmentType.olEmbeddeditem, , Subject
            Case enuLinkStrategy.multiple
                objOut.body = Subject & vbCrLf & vbCrLf
                ' CopyAttachments(objIn, objOut)
            Case enuLinkStrategy.referenced
                Set objOut = ReferenceObjects(objIn, objOut)
            Case Else
                Set objOut = CopyAttachments(objIn, objOut)
        End Select
        
      '  Dim oT As Outlook.TaskItem
      '  oT.Attachments.Add objIn, Outlook.OlAttachmentType.olEmbeddeditem, 1, ""
    
        ' Return Updated Object
        Set LinkObjects = retObject
        Exit Function
        
ThrowException:
        Set LinkObjects = Nothing
        
    End Function
    
    
    Public Function ReferenceObjects(ByRef objIn As Object, _
                                     ByRef objOut As Object) As Object
    
    '   In      Out     Comment
    '   Mail    Task    Task's reference Id documented in Mail's body
    '   Task    Mail    Task's reference Id documented in Mail's body
    '   Mail    Appt
    '   Appt    Mail
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":ReferenceObjects"
        
        On Error GoTo ThrowException
        
        Dim tm As New TaskManager
        
        Dim olClassIn As Outlook.OlObjectClass
        Dim olClassOut As Outlook.OlObjectClass
        olClassIn = objIn.Class
        olClassOut = objOut.Class
        
        Select Case olClassIn
            Case Outlook.OlObjectClass.olMail
                Select Case olClassOut
                    Case Outlook.OlObjectClass.olTask
                        Call tm.LinkMailToTask(objIn, objOut)
                    Case Else
                        strTrace = "Unrecognized output object class: " & olClassIn
                        GoTo ThrowException
                End Select
            Case Outlook.OlObjectClass.olTask
                    Select Case olClassOut
                        Case Outlook.OlObjectClass.olMail
                            Call tm.LinkMailToTask(objOut, objIn)
                        Case Else
                            strTrace = "Unrecognized output object class: " & olClassIn
                            GoTo ThrowException
                    End Select
            Case Else
                strTrace = "Unrecognized input object class: " & olClassIn
                GoTo ThrowException
        End Select
        
        Set ReferenceObjects = objOut
        Exit Function
    
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        Set ReferenceObjects = Nothing
        
    End Function
    
    ''' <summary>
    ''' Gets the first n character from the Input string
    ''' </summary>
    ''' <param name="strInput">String to summarize</param>
    ''' <returns>String:</returns>
    ''' <remarks></remarks>
    Public Function GetTeaserString(ByVal strInput As String) As String

        Dim strTrace As String
        strTrace = "General Fault."
        
        Dim strReturn As String
        
        If Len(strInput) = 0 Then
            GetTeaserString = ""
            Exit Function
        End If

        Dim iMax As Integer
        iMax = Commands.TeaserLength
        If Len(strInput) > iMax Then
            strReturn = Left(strInput, iMax) & "..."
        Else
            strReturn = strInput
        End If

        GetTeaserString = strReturn
    
    End Function
    
    
    Private Function CloseItem(ByVal oItem As Object, _
                      Optional ByVal SaveMode As Outlook.OlInspectorClose = Outlook.OlInspectorClose.olPromptForSave) As Boolean
                                                
        
        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":CloseItem"
        
        strTrace = "Entered routine."
        
        If IsNothing(oItem) Then
            strTrace = "A null object was encountered - not allowed."
        End If

        On Error Resume Next
        strTrace = "Closing the Outlook Item."
        oItem.Close (SaveMode)
            
        ' Dim errMsg As String = "Cannot close the item(s) - try opening and saving the item."
        ' strTrace = errMsg & vbCrLf & vbCrLf & ex.message
        ' MessageBoxCeptara(strTrace, "Unable to Move Item", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
        ' strTrace = errMsg & " - " & ex.message
        ' Throw New Exception(strDefault)

        strTrace = "Exited routine."
        ' TraceLogger(strTrace, strRoutine)

        CloseItem = True

    End Function
    
    ''' <summary>
    ''' Returns the Outlook full schema name for the specific Outlook User Property
    ''' that ties this item to its Outlook object.
    ''' </summary>
    ''' <returns>String: full schema name</returns>
    ''' <remarks>e.g. FME_ContactItemID, FME_OutlookItemID, POSTaskProperties, ...</remarks>
    Public Function GetOutlookMAPIUserPropertySchemaName(ByVal UserPropertyName As String) As String

        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":GetOutlookMAPIUserPropertySchemaName"
        
        If Len(UserPropertyName) = 0 Then
            strTrace = "No Property name was provided."
            GoTo ThrowException
        End If

        Dim strProperty As String
        strProperty = "http://schemas.microsoft.com/mapi/string/" & _
                        "{00020329-0000-0000-C000-000000000046}/" & UserPropertyName & _
                        "/0x0000001f"

        GetOutlookMAPIUserPropertySchemaName = strProperty
        Exit Function
            
ThrowException:
        LogMessage strTrace, strRoutine

    End Function
    
    ''' Retrieves the Reference ID for the specified Outlook Item
    Public Function GetReferenceID(ByVal oItem As Object) As String
    
        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":GetReferenceID"
        
        If IsNothing(oItem) Then
            strTrace = "A null outlook item encountered."
            GoTo ThrowException
        End If
        
        Dim currID As String
        
        Dim strProperty As String
        strProperty = ReferenceIDUserProperty
        
        Dim oUserProperty As Outlook.UserProperty
        Set oUserProperty = oItem.UserProperties(strProperty)
        
        If Not IsNothing(oUserProperty) Then
            currID = oUserProperty.value
            If Len(currID) = 0 Then
                strTrace = "WARNING: Reference Id not assigned to item: " & oItem.Subject
                GoTo ThrowException
            End If
        Else
            strTrace = "Failed to find a user property named: " & ReferenceIDUserProperty & "."
            GoTo ThrowException
        End If
        
        GetReferenceID = currID
        Exit Function
        
ThrowException:
        LogMessage strTrace, strRoutine
        GetReferenceID = ""
    
    End Function
    
    ''' <summary>
    ''' Sets a unique reference ID for the specified item via the Outlook UserProperties
    ''' - usually is name FMEReferenceID
    ''' </summary>
    ''' <param name="oItem">Outlook Object, e.g. Outlook.TaskItem</param>
    ''' <param name="refId">String (Optional, refId created if empty):</param>
    ''' <returns>String: Returns the assigned ReferenceId</returns>
    ''' <remarks>e.g. FME_ContactItemID, FME_OutlookItemID, POSTaskProperties, ...</remarks>
    Public Function SetReferenceID(ByRef oItem As Object, _
                          Optional ByVal refId As String = "") As String
    
        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":SetReferenceID"
        
        If Len(refId) = 0 Then
            strTrace = "Creating a new reference Id."
            refId = GenerateUniqueID(ReferenceIDLength)
        End If
    
        strTrace = "Setting the reference ID property to " & refId
        Dim strReturn As String
        strReturn = ""
        
        Dim strProperty As String
        strProperty = ReferenceIDUserProperty
        
        Dim oUserProperty As Outlook.UserProperty
        Set oUserProperty = oItem.UserProperties(strProperty)
        
        If Not IsNothing(oUserProperty) Then
            Dim currID As String
            currID = oUserProperty.value
            If Len(currID) = 0 Then
                ' Set the value
                oUserProperty.value = refId
                ' Save to the Outlook datastore
                oItem.Save
            Else
                ' Keep the current reference ID
                strTrace = "Item already assigned a reference id: " & currID & "."
                iReturn = 1
            End If
        Else
            ' Add user property
            Set oUserProperty = oItem.UserProperties.Add(ReferenceIDUserProperty, olText)
            ' Set the value
            oUserProperty.value = refId
            ' Save to the Outlook datastore
            oItem.Save
        End If
        
        SetReferenceID = refId
        Exit Function
        
ThrowException:
        LogMessage strTrace, strRoutine
        SetReferenceID = -1
    
    End Function
    
    ''' Returns the signature for FocusMe tracking
    Public Function GetReferenceMessage(ByVal refId As String) As String
    
        ' Organized and powered by FocusMe for VBA, Ceptara Corp © 2019
        ' Bringing excellence to managing chaos (email) and staying goal focused!
        ' [FSRefID: GV1V72]
        
        ' ReferenceIDString = FSRefID:
    
        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":GetReferenceMessage"
    
        Dim strDate As String
        strDate = Format(Date, "yyyy")
    
        Dim strReturn As String
        
        strReturn = "Organized and powered by FocusMe for VBA, Ceptara Corp (c) " & strDate & vbCrLf & _
                    "Bringing excellence to managing chaos (email) and staying goal focused!" & vbCrLf & _
                    "[" & ReferenceIDString & " " & refId & "]"
                    
        strReturn = "FocusMe for VBA, " & strDate & vbCrLf & _
                    "[" & ReferenceIDString & " " & refId & "]"
        
        GetReferenceMessage = strReturn
        Exit Function
        
ThrowException:
        LogMessage strTrace, strRoutine
        GetReferenceMessage = ""
    
    End Function
    
    ''' Appends the Reference Message to the body of the incoming outlook item
    Public Sub AppendMarketingSignature(ByVal oItem As Object, ByVal refId As String)
    
        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":GetReferenceMessage"
        
        If IsNothing(oItem) Then
            strTrace = "A null Outlook Item encountered."
            GoTo ThrowException
        End If
        If Len(refId) = 0 Then
            strTrace = "An empty ref id found."
            GoTo ThrowException
        End If
        
        On Error GoTo ThrowException
        
        ' Check if oItem already marked
        Dim pos As Long
        pos = InStr(1, UCase(oItem.body), UCase(refId), vbTextCompare)
        If pos > 0 Then
            strTrace = "Outlook item already marked with reference id: " & refId
            GoTo ThrowException
        End If
        
        Dim strNew As String
        Dim strInsert As String
        strInsert = GetReferenceMessage(refId)
        
        Dim oMail As Outlook.MailItem
        
        If TypeOf oItem Is Outlook.MailItem Then
            strTrace = "Casting untyped item to outlook.mailitem."
            Set oMail = oItem
            If oMail.BodyFormat = olFormatHTML Then
                strTrace = "Inserting HTML Text."
                strInsert = Replace(strInsert, vbCrLf, "<br/>")
                Dim strHTML As String
                strHTML = oMail.HTMLBody
                pos = InStr(1, UCase(strHTML), "</BODY>", vbTextCompare)
                If pos <= 0 Then
                    ' No body end tag found - add to the end
                    strNew = strHTML & strInsert
                Else
                    ' Analysis
                    strTrace = "Length: " & Len(strHTML) & " pos: " & pos
                    LogMessage strTrace, strRoutine
                    strTrace = "Left Length: " & pos - 1
                    LogMessage strTrace, strRoutine
                    strTrace = "Right Length: " & Len(strHTML) - pos
                    LogMessage strTrace, strRoutine
                    
                    ' Insert just before the Body end tag
                    Dim strBefore As String
                    strBefore = Left(strHTML, pos - 1)
                    Dim strAfter As String
                    strAfter = Mid(strHTML, pos, Len(strHTML) - pos)
                
                    strNew = strBefore & strInsert & strAfter
                End If
            
                oMail.HTMLBody = strNew
            
            End If
            If oMail.BodyFormat = olFormatPlain Then
                strTrace = "Inserting plain text."
                oMail.body = oMail.body & vbCrLf & strInsert
            End If
            If oMail.BodyFormat = olFormatRichText Then
                strTrace = "WARNING: Linking via FS Reference IDs not supported for RichText formatting."
                GoTo ThrowException
            End If
        Else
           oItem.body = oItem.body & strInsert
        End If
        
        Exit Sub
            
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        
    End Sub

' MailItems ---

    ''' Create a new mailItem
    Public Function CreateMailItem() As Outlook.MailItem
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":CreateMailItem"
        
        On Error GoTo ThrowException
        
        Dim m As Outlook.MailItem
        Set m = app.CreateItem(olMailItem)
        
        Set CreateMailItem = m
        Exit Function
        
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        Set CreateMailItem = Nothing
    
    End Function

    ''' Retrieves a collection of Outlook objects using the specified
    ''' Filter and ItemType.  If Folders is nothing, then the Default
    ''' folder for the given itemType is used.
    Public Function GetOutlookItems(ByVal DASLQuery As String, _
                                    ByVal itemType As OlItemType, _
                           Optional ByVal Folders As ArrayList = Nothing) As ArrayList
                           
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":GetOutlookItems"
        
        On Error GoTo ThrowException
        
        Dim retItems As ArrayList
        Set retItems = New ArrayList
        
        Dim oFolder As Outlook.Folder
        If Folders Is Nothing Then
            Set Folders = New ArrayList
            Select Case itemType
                Case OlItemType.olAppointmentItem
                    Set oFolder = oNS.GetDefaultFolder(olFolderCalendar)
                Case OlItemType.olContactItem
                    Set oFolder = oNS.GetDefaultFolder(olFolderContacts)
                Case OlItemType.olJournalItem
                    Set oFolder = oNS.GetDefaultFolder(olFolderJournal)
                Case OlItemType.olMailItem
                    Set oFolder = oNS.GetDefaultFolder(olFolderInbox)
                Case OlItemType.olTaskItem
                    Set oFolder = oNS.GetDefaultFolder(olFolderTasks)
                Case Else
                    strTrace = "Unsupported itemType: " & itemType
                    GoTo ThrowException
            End Select
            Folders.Add oFolder
        End If
        
        Dim fldr As Outlook.Folder
        Dim oItems As Outlook.Items
        For Each fldr In Folders
            strTrace = "Set the filter: " & DASLQuery
            If Len(DASLQuery) = 0 Then
                Set oItems = fldr.Items
            Else
                Set oItems = fldr.Items.Restrict(DASLQuery)
            End If
            
            strTrace = "Add the Outlook Items collection to the return list."
            For Each o In oItems
                retItems.Add o
            Next
            
        Next
        
        Set GetOutlookItems = retItems
        Exit Function
        
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        Set GetOutlookItems = Nothing
    
    End Function

    ''' <summary>
    ''' Gets a list of Outlook.MailItems for a given folder set and query.
    ''' </summary>
    ''' <param name="DASLQuery">String: Optional DASL query - if empty returns all items.</param>
    ''' <param name="Folders">ArrayList - if nothing returns default task folder.</param>
    ''' <returns>ArrayList Of Outlook.MailItem</returns>
    ''' <remarks>If an error occurs returns Nothing.</remarks>
    Public Function GetOutlookMail(Optional ByVal DASLQuery As String = "", _
                                   Optional ByVal Folders As ArrayList = Nothing) As ArrayList
                                     
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":GetOutlookMail"
        
        On Error GoTo ThrowException
                                                
        Dim retItems As ArrayList
        Set retItems = New ArrayList
                                                
        Dim fldr As Outlook.Folder
        Dim oItems As Outlook.Items
        
        If IsNothing(Folders) Then
            Set Folders = GetMailFolders
        End If
        
        For Each fldr In Folders
            If Len(DASLQuery) = 0 Then
                Set oItems = fldr.Items
            Else
                Set oItems = fldr.Items.Restrict(DASLQuery)
            End If
    
            strTrace = "Add the Outlook Items collection to the return list."
            For Each o In oItems
                If TypeOf o Is Outlook.MailItem Then
                    retItems.Add o
                End If
            Next
        Next
                                                   
        Set GetOutlookMail = retItems
        Exit Function
        
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        Set GetOutlookMail = Nothing
        
    End Function

    ''' <summary>
    ''' Retrieves all messages with the specific topic in the body of the message.  Looks in
    ''' the specified folder.
    ''' </summary>
    ''' <param name="Topic">String: word or phrase to search for, e.g. refId</param>
    ''' <param name="Folder">Outlook.Folder:</param>
    ''' <returns>ArrayList:</returns>
    ''' <remarks></remarks>
    Public Function GetConversationViaBody(ByVal Topic As String, _
                                           ByVal Folder As Outlook.Folder) As ArrayList
                                           
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":GetConversationViaBody"
        
        On Error GoTo ThrowException
        
        If Len(Topic) = 0 Then
            strTrace = "A Topic is required."
            GoTo ThrowException
        End If
        If IsNothing(Folder) Then
            strTrace = "A null folder was encountered."
            GoTo ThrowException
        End If

        Dim retList As New ArrayList

        Dim iCnt As Integer
        iCnt = 0

        Dim oMailFolder As Outlook.Folder
        Set oMailFolder = Folder
        Dim strPath As String
        strPath = oMailFolder.FolderPath

        strTrace = "Get total message count for folder '" & strPath & "'."
        Dim cnt As Integer
        cnt = oMailFolder.Items.count
        
'"urn:schemas:httpmail:textdescription" LIKE '%0B%'

        Dim strTopic As String
        strTopic = Topic
        Dim sFilter As String
        sFilter = "@SQL=" & Chr(34) & "urn:schemas:httpmail:textdescription" & _
                    Chr(34) & " LIKE '%" & strTopic & "%'"

        ' Get related messages from selected item's folder
        Dim oItems As Outlook.Items
        Set oItems = oMailFolder.Items.Restrict(sFilter)

        For Each o In oItems
            Dim m As Outlook.MailItem
            Set m = o
            If Not m Is Nothing Then
                ' found a related item
                retList.Add m
                iCnt = iCnt + 1

                strTrace = "Index = " & m.ConversationIndex
                strTrace = iCnt

            End If
        Next

        strTrace = "Found " & iCnt & " mail items in conversation '" & strTopic & _
                    "' out of " & cnt & " items the folder " & strPath & "."
        LogMessage strTrace, strRoutine
                    
        Set GetConversationViaBody = retList
        Exit Function
        
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        Set GetConversationViaBody = Nothing
                                           
    End Function
    
    ''' <summary>
    ''' Retrieves all Outlook MailItems in a conversation associated with a specified MailItem.
    ''' </summary>
    ''' <param name="originalMail">Outlook.MailItem:</param>
    ''' <returns>ArrayList: collection of Outlook.MailItem</returns>
    ''' <remarks></remarks>
    Public Function GetConversationFromMailItem(ByVal originalMail As Outlook.MailItem) As ArrayList

        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":GetConversationFromMailItem"
                
        If IsNothing(originalMail) Then
            strTrace = "A null message encountered."
            GoTo ThrowException
        End If

        Dim retItems As New ArrayList

        Dim fldr As Outlook.Folder
        Set fldr = originalMail.Parent
        If Not IsNothing(fldr) Then
            ' Check to make sure datastore supports the Conversation object.
            '   Supported stores include POP, IMAP, PST, or Microsoft Exchange Server
            Dim myStore As Outlook.Store
            Set myStore = fldr.Store
            If IsNothing(myStore) Then
                strTrace = "A null store encountered."
                GoTo ThrowException
            End If

            If myStore.IsConversationEnabled Then
                ' Retrieve the Conversation
                Dim myConversation As Outlook.Conversation
                Set myConversation = originalMail.GetConversation

                ' Get the Conversation Items
                Dim myTable As Outlook.Table
                Set myTable = myConversation.GetTable

                'Dim iCnt As Integer = myTable.GetRowCount
                'strTrace = "Found " & iCnt.ToString & " items in the conversation."
                'TraceLogger(strTrace, strRoutine)

                ' Collect the Items and return the Outlook MailItems
                Do Until myTable.EndOfTable
                    Dim nextRow As Outlook.Row
                    Set nextRow = myTable.GetNextRow
                    Dim strID As String
                    strID = nextRow("EntryID")
                    strTrace = "Get the object."
                    Dim o As Object
                    Set o = oNS.GetItemFromId(strID)
                    If Not IsNothing(o) Then
                        '  If TypeOf o Is Outlook.MailItem Then retItems.Add(o)
                        retItems.Add o
                    End If
                Loop

                'Dim oCnt As Integer = retItems.Count
                'strTrace = "Added " & oCnt.ToString & " items to the return conversation."
                'TraceLogger(strTrace, strRoutine)

            Else
                strTrace = "Selected store '" & myStore.DisplayName & "' does not support conversation view."
                GoTo ThrowException
            End If
        Else
            strTrace = "Failed to cast the mailItem parent to a MAPIFolder."
            GoTo ThrowException
        End If

        Set GetConversationFromMailItem = retItems
        Exit Function

ThrowException:
        LogMessageEx strTrace, err, strRoutine
        Set GetConversationFromMailItem = Nothing

    End Function
        
' Contacts --

''' Create a new contact in the default contacts folder
    Public Function CreateContact() As Outlook.ContactItem
    
        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":CreateContact"
        
        On Error GoTo ThrowException
        
        ' Set the destination task folder
        Dim objFolder As Outlook.Folder
        Set objFolder = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts)
        If IsNothing(objFolder) Then
            strTrace = "Unable to retrieve the contact's default folder."
            GoTo ThrowException
        End If
        
        Dim objContact As Outlook.ContactItem
        Set objContact = objFolder.Items.Add

        Set CreateContact = objContact
        Exit Function
                                    
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        Set CreateContact = Nothing
    
    End Function
    
    ''' Returns the SMTP address from an OutlookContact EmailAddress Entry
    Public Function GetSMTPEmailAddressFromContactEmailEntry( _
                            ByVal DisplayName As String, _
                            ByVal emailAddress As String, _
                   Optional ByVal addrType As String = "SMTP")
                                                                                   
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":GetEmailAddressesFromContactEmailEntry"
        
        Dim strSMTPAddress As String
        
        If Len(emailAddress) = 0 Then
            strTrace = "An email address is required."
            GoTo ThrowException
        End If

        If UCase(addrType) = "EX" Then
            strSMTPAddress = GetSMTPAddressFromExchangeUser(DisplayName)
        End If
        If UCase(addrType) = "SMTP" Then
            strSMTPAddress = emailAddress
        End If

        GetSMTPEmailAddressFromContactEmailEntry = strSMTPAddress
        Exit Function

ThrowException:
        LogMessageEx strTrace, err, strRoutine
        GetSMTPEmailAddressFromContactEmailEntry = ""

    End Function
    
    ''' <summary>
    ''' Finds the SMTP address for a given Display Name that matches
    ''' the Exchange Display name entry for a given Exchange user.
    ''' </summary>
    ''' <param name="DisplayName">String: The display in the Exchange GAL.</param>
    ''' <returns>String:</returns>
    ''' <remarks></remarks>
    Public Function GetSMTPAddressFromExchangeUser(ByVal DisplayName As String) As String

        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":GetSMTPAddressFromExchangeUser"
                
        strTrace = "Error checking."
        If Len(DisplayName) = 0 Then
            strTrace = "An empty or null Exchange Display Name was encountered."
            GoTo ThrowException
        End If

        Dim strSMTPAddress As String

        strTrace = "Resolving a Recipient for '" & DisplayName & "'."
        LogMessage strTrace, strRoutine
        
        Dim oExUser As Outlook.ExchangeUser

        Dim oRecipient As Outlook.Recipient
        Set oRecipient = oNS.CreateRecipient(DisplayName)
        oRecipient.Resolve
        If oRecipient.Resolved Then
            Select Case oRecipient.AddressEntry.AddressEntryUserType
                Case Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry
                    Set oExUser = oRecipient.AddressEntry.GetExchangeUser
                    If Not IsNothing(oExUser) Then
                        strSMTPAddress = oExUser.PrimarySmtpAddress
                    Else
                        strTrace = "Failed to retrieve the Exchange User for Display name: '" & DisplayName & "'."
                        GoTo ThrowException
                    End If
                Case Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry
                    Set oExUser = oRecipient.AddressEntry.GetExchangeUser
                    If Not IsNothing(oExUser) Then
                        strSMTPAddress = oExUser.PrimarySmtpAddress
                    Else
                        strTrace = "Failed to retrieve the Remote Exchange User for Display name: '" & DisplayName & "'."
                        GoTo ThrowException
                    End If
                Case Outlook.OlAddressEntryUserType.olOutlookContactAddressEntry
                    Set oExUser = oRecipient.AddressEntry.GetExchangeUser
                    If Not IsNothing(oExUser) Then
                        strSMTPAddress = oExUser.PrimarySmtpAddress
                    Else
                        strTrace = "Failed to retrieve the Exchange User for Display name: '" & DisplayName & "'."
                        GoTo ThrowException
                    End If
                Case Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry
                    Dim oExDistList As Outlook.ExchangeDistributionList
                    Set oExDistList = oRecipient.AddressEntry.GetExchangeDistributionList
                    If Not IsNothing(oExDistList) Then
                        strSMTPAddress = oExDistList.PrimarySmtpAddress
                    End If
                Case Else
                    strTrace = "Failed to retrieve an SMTP address for Display name: '" & DisplayName & "'."
                    GoTo ThrowException
            End Select
        Else
            strTrace = "Failed to resolve the Exchange User for Display name: '" & DisplayName & "'."
            GoTo ThrowException
        End If

        GetSMTPAddressFromExchangeUser = strSMTPAddress
        Exit Function

ThrowException:
        LogMessageEx strTrace, err, strRoutine
        GetSMTPAddressFromExchangeUser = ""

    End Function

' Tasks ---

    ''' Create a new task in the default tasks folder
    Public Function CreateTask() As Outlook.TaskItem
    
        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":CreateTask"
        
        On Error GoTo ThrowException
        
        ' Set the destination task folder
        Dim objFolder As Outlook.Folder
        Set objFolder = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderTasks)
        If IsNothing(objFolder) Then
            strTrace = "Unable to retrieve the task's default folder."
            GoTo ThrowException
        End If

        Dim objTask As Outlook.TaskItem
        Set objTask = objFolder.Items.Add
        
        Set CreateTask = objTask
        Exit Function
                                    
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        Set CreateTask = Nothing
    
    End Function
    
    ''' Creates a related task for the incoming object
    ''' - has the same basic info, e.g. subject, categories and teaser
    Public Function CreateRelatedTask(ByVal objInput As Object, _
                               Optional ByVal AddTeaser As Boolean = True) As Outlook.TaskItem
                                     
        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":CreateTrackingTask"

        Dim objTask As Outlook.TaskItem
        Set objTask = CreateTask
        If IsNothing(objTask) Then
            strTrace = "Failed to create an Outlook Task."
            GoTo ThrowException
        End If
        
        ' Common variables
        Dim strCategories As String
        strCategories = objInput.Categories
        Dim strSubject As String
        strSubject = objInput.Subject
        Dim strBody As String
        strBody = ""
        If AddTeaser Then strBody = GetTeaserString(objInput.body)
        
        objTask.Subject = strSubject
        objTask.Categories = strCategories
        objTask.body = strBody
        
        Set CreateRelatedTask = objTask
        Exit Function
                                    
ThrowException:
        LogMessage strTrace, strRoutine
                                     
    End Function

    ''' <summary>
    ''' Creates an Outlook.TaskItem from the specified object
    ''' </summary>
    ''' <param name="objInput">Outlook Object (e.g. MailItem, AppointmentItem, etc.):</param>
    ''' <param name="linkBy">enuLinkStrategy:</param>
    ''' <param name="AddTeaser">Boolean:</param>
    ''' <param name="closeIncomingItem">Boolean:</param>
    ''' <returns>Outlook.TaskItem:</returns>
    ''' <remarks></remarks>
    Public Function MakeTaskFromItem(ByVal objInput As Object, _
                                     ByVal linkBy As enuLinkStrategy, _
                                     Optional AddTeaser As Boolean, _
                                     Optional closeIncomingItem As Boolean) As Outlook.TaskItem
                                    
        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":MakeTaskFromItem"
        
        On Error GoTo ThrowException
                                    
        If IsNothing(objInput) Then
            strTrace = "A null input object encountered."
            GoTo ThrowException
        End If
        If IsMissing(AddTeaser) Then
            AddTeaser = True
        End If
        If IsMissing(closeIncomingItem) Then
            closeIncomingItem = True
        End If
        
        ' Common variables
        Dim strCategories As String
        strCategories = objInput.Categories
        Dim strSubject As String
        strSubject = objInput.Subject
        Dim strBody As String
        strBody = objInput.body

        Dim dteDueDate As Date
        dteDueDate = DefaultDateOutlook
        
        ' Get the incoming object's Outlook class
        Dim olClass As Outlook.OlObjectClass
        olClass = objInput.Class
        
        Dim retObject As Object
        Set retObject = Nothing
        
        Dim bDoIt As Boolean
        bDoIt = False
        
        Select Case olClass
        
            Case Outlook.OlObjectClass.olMail
            
                bDoIt = True
            
            Case Else
        
        End Select
        
        If bDoIt Then
        
            ' Set the destination task folder
            Dim objFolder As Outlook.Folder
            Set objFolder = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderTasks)
            If IsNothing(objFolder) Then
                strTrace = "Unable to retrieve the task's default folder."
                GoTo ThrowException
            End If

            strTrace = "Adding a " & objInput.Class & " item to a task item: " & strSubject
            Dim objTask As Outlook.TaskItem
            Set objTask = objFolder.Items.Add

            ' Align Categories
            objTask.Categories = strCategories

            ' From the Appointment Item
            objTask.DueDate = dteDueDate

            ' Add Teaser if appropriate
            If linkBy = enuLinkStrategy.embed Then strBody = GetTeaserString(strBody) & vbCrLf

            ' Connect the incoming object with the new task
            Set objTask = LinkObjects(linkBy, objInput, strSubject, strBody, objTask)
           
            ' Pass back the newly created task
            Set retObject = objTask
        
        Else
        
        End If
             
        If closeIncomingItem Then CloseItem objInput
             
        Set MakeTaskFromItem = retObject
        Exit Function
                                    
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        
    End Function
    
    ''' Prepares an Outlook.MailItem, tagged with the task's reference Id to
    ''' send to a recipient
    Public Function GetMessageForTask(ByVal t As Outlook.TaskItem) As Outlook.MailItem
    
        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":GetMessageForTask"
        
        On Error GoTo ThrowException
        
        ' Create new MailItem
        Dim oMail As Outlook.MailItem
        Set oMail = oNS.Application.CreateItem(olMailItem)
        
        ' Add user signature
        Dim insp As Outlook.Inspector
        Set insp = oMail.GetInspector
        
        ' Add Task Reference Id
        ReferenceObjects t, oMail
        
        ' Set Mail Message Topic
        oMail.Subject = t.Subject
        
        ' Return new Email
        Set GetMessageForTask = oMail
        Exit Function
        
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        GetMessageForTask = Nothing
    
    End Function
                                     
                                     
' Appointments ---

    ''' Create a new appointmentItem
    Public Function CreateAppointment() As Outlook.AppointmentItem
    
        Dim strTrace As String
        strTrace = "General Fault."
        Dim strRoutine As String
        strRoutine = rootClass & ":CreateAppointment"
        
        On Error GoTo ThrowException
        
        Dim a As Outlook.AppointmentItem
        Set a = app.CreateItem(olAppointmentItem)
        
        Set CreateAppointment = a
        Exit Function
        
ThrowException:
        LogMessageEx strTrace, err, strRoutine
        Set CreateAppointment = Nothing
    
    End Function
                 
    ''' <summary>
    ''' Creates an Outlook.Appointment from the specified object
    ''' </summary>
    ''' <param name="objInput">Outlook Object (e.g. MailItem, TaskItem, etc.):</param>
    ''' <param name="linkBy">enuLinkStrategy:</param>
    ''' <param name="AddTeaser">Boolean (Optional, defaults to True):</param>
    ''' <param name="closeIncomingItem">Boolean (Optional, defaults to True):</param>
    ''' <returns>Outlook.AppointmentItem:</returns>
    Public Function MakeAppointmentFromItem(ByVal objInput As Object, _
                                            ByVal linkBy As enuLinkStrategy, _
                                            Optional AddTeaser As Boolean, _
                                            Optional closeIncomingItem As Boolean) As Outlook.AppointmentItem
                                            
        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":MakeAppointmentFromItem"
                                    
        If IsNothing(objInput) Then
            strTrace = "A null input object encountered."
            GoTo ThrowException
        End If
        If IsMissing(AddTeaser) Then
            AddTeaser = True
        End If
        If IsMissing(closeIncomingItem) Then
            closeIncomingItem = True
        End If
        
        ' Common variables
        Dim strCategories As String
        strCategories = objInput.Categories
        Dim strSubject As String
        strSubject = objInput.Subject
        Dim strBody As String
        strBody = objInput.body

        Dim dteEnd As Date
        dteEnd = DefaultDateOutlook
        Dim dteStart As Date
        dteStart = DefaultDateOutlook
        
        dteStart = GetDateTimeNearest(Now, Common.DateInterval.Minute, 30)
        dteEnd = DateAdd("n", 30, dteStart)
        
        ' Get the incoming object's Outlook class
        Dim olClass As Outlook.OlObjectClass
        olClass = objInput.Class
        
        Dim retObject As Object
        Set retObject = Nothing
              
        Select Case olClass
            Case Outlook.OlObjectClass.olJournal
                Dim oJournal As Outlook.JournalItem
                Set oJournal = objInput
                If oJournal.Duration > 0 Then
                    dteStart = DateAdd("n", -1 * CDbl(oJournal.Duration), Now)
                    dteEnd = Now
                End If
            
            Case Outlook.OlObjectClass.olTask
                Dim oTask As Outlook.TaskItem
                Set oTask = objInput
                If oTask.ActualWork > 0 Then
                    dteStart = DateAdd("n", -1 * CDbl(oTask.ActualWork), Now)
                    dteEnd = Now
                End If
            
            Case Else
        
        End Select
        
        ' Set the destination calendarfolder
        Dim objFolder As Outlook.Folder
        Set objFolder = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar)
        If IsNothing(objFolder) Then
            strTrace = "Unable to retrieve the calendar's default folder."
            GoTo ThrowException
        End If

        strTrace = "Adding a " & objInput.Class & " item to an appointment: " & strSubject
        Dim oAppt As Outlook.AppointmentItem
        Set oAppt = objFolder.Items.Add
        
        ' Align Categories
        oAppt.Categories = strCategories

        ' Assign the Appointment values
        oAppt.Subject = strSubject
        oAppt.body = strBody

        oAppt.Start = dteStart
        oAppt.End = dteEnd

        ' Connect the incoming object with the new appointment
        Set oAppt = LinkObjects(linkBy, objInput, strSubject, strBody, oAppt)

        ' Pass back the newly created appointment
        Set retObject = oAppt

        If closeIncomingItem Then CloseItem objInput
        
        Set MakeAppointmentFromItem = retObject
        Exit Function
                                    
ThrowException:
        LogMessage strTrace, strRoutine
                                            
    End Function
                 

' JournalItems ---
                 
    ''' <summary>
    ''' Creates an Outlook.JournalItem from the specified object
    ''' </summary>
    ''' <param name="objInput">Outlook Object (e.g. MailItem, TaskItem, etc.):</param>
    ''' <param name="linkBy">enuLinkStrategy:</param>
    ''' <param name="AddTeaser">Boolean (Optional, defaults to True):</param>
    ''' <param name="closeIncomingItem">Boolean (Optional, defaults to True):</param>
    ''' <returns>Outlook.JournalItem:</returns>
    Public Function MakeJournalEntryFromItem(ByVal objInput As Object, _
                                             ByVal linkBy As enuLinkStrategy, _
                                    Optional ByVal AddTeaser As Boolean, _
                                    Optional ByVal closeIncomingItem As Boolean) As Outlook.JournalItem
                                    
        Dim strTrace As String
        Dim strRoutine As String
        strRoutine = rootClass & ":MakeJournalEntryFromItem"
                                    
        If IsNothing(objInput) Then
            strTrace = "A null input object encountered."
            GoTo ThrowException
        End If
        If IsMissing(AddTeaser) Then
            AddTeaser = True
        End If
        If IsMissing(closeIncomingItem) Then
            closeIncomingItem = True
        End If
        
        ' Common variables
        Dim strCategories As String
        strCategories = objInput.Categories
        Dim strSubject As String
        strSubject = objInput.Subject
        Dim strBody As String
        strBody = objInput.body
        Dim strType As String
        strType = "Document"
        
        ' Get the incoming object's Outlook class
        Dim olClass As Outlook.OlObjectClass
        olClass = objInput.Class
        
        Dim retObject As Object
        Set retObject = Nothing
        
        Select Case olClass
            Case Outlook.OlObjectClass.olMail
                strType = "E-mail Message"
                
            Case Outlook.OlObjectClass.olAppointment
                strType = "Meeting"
                
            Case Outlook.OlObjectClass.olJournal
               strType = "Note"
               
            Case Outlook.OlObjectClass.olTask
                strType = "Task"
                            
            Case Else
                ' Document - set before
        
        End Select
        
        ' Set the destination journal folder
        Dim objFolder As Outlook.Folder
        Set objFolder = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJournal)
        If IsNothing(objFolder) Then
            strTrace = "Unable to retrieve the journal's default folder."
            GoTo ThrowException
        End If

        strTrace = "Adding a " & objInput.Class & " item to a journalItem: " & strSubject
        Dim oJournal As Outlook.JournalItem
        Set oJournal = objFolder.Items.Add
        
        ' Align Categories
        oJournal.Categories = strCategories

        ' Assign the JournalItem values
        oJournal.Subject = strSubject
        oJournal.body = strBody
        oJournal.Type = strType

        ' oAppt.Start = dteStart
        ' oAppt.End = dteEnd

        ' Connect the incoming object with the new journalItem
        Set oJournal = LinkObjects(linkBy, objInput, strSubject, strBody, oJournal)

        ' Pass back the newly created journalItem
        Set retObject = oJournal

        If closeIncomingItem Then CloseItem objInput
        
        Set MakeJournalEntryFromItem = retObject
        Exit Function
                                    
ThrowException:
        LogMessage strTrace, strRoutine
                                            
    End Function



    
