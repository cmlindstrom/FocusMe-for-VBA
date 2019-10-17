VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' - Fields

Private Const rootClass As String = "Setup"

' - Events

' - Properties

' - Event Handlers

' - Constructor

Private Sub Class_Initialize()

End Sub

' - Methods

''' Returns True if a DataStore has not been created
Public Function IsFirstTime() As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":IsFirstTime"
    
    On Error GoTo ThrowException
    
    Dim bReturn As Boolean
    bReturn = False
    
    Dim ldb As New dsDataStore
    ldb.Connect
    
    bReturn = Not ldb.DataStoreExists
    
    IsFirstTime = bReturn
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    IsFirstTime = False

End Function

''' Index the user's Outlook folders
Public Function IndexOutlookFolders() As Integer

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":IndexOutlookFolders"
    
    On Error GoTo ThrowException
    
    strTrace = "Get Outlook folders."
    Dim ut As New Utilities
    Dim fldrs As fmeFolders
    Set fldrs = ut.IndexFolders()
    If fldrs Is Nothing Then
        strTrace = "Failed to query the Outlook dataStore folder structure."
        GoTo ThrowException
    End If
    If fldrs.Count = 0 Then
        strTrace = "User has no folders in their folder tree."
        GoTo ThrowException
    End If
    
    Dim ldb As New dsDataStore
    ldb.Connect
    
    strTrace = "Clear the current folder set."
    ldb.ClearEntireCollection "Folder"
    
    strTrace = "Update the local datastore with the indexed folders from Outlook."
    Dim fCnt As Integer
    fCnt = 0
    Dim f As fmeFolder
    For Each f In fldrs.Items
    
        If Not Len(f.EntryId) = 0 Then

            ' Add folder to local datastore
            If ldb.Insert(f, "Folder") Then
                strTrace = "Indexed a new folder: " & f.path
                fCnt = fCnt + 1
            Else
                strTrace = "Failed to insert a new folder (" & f.path & ") into the datastore."
                LogMessage strTrace, strRoutine
            End If
                
        Else
            strTrace = "Ignoring Outlook folder - empty path, id: " & f.id
            LogMessage strTrace, strRoutine
        End If
                
    Next
    
    strTrace = "Indexed " & fCnt & " folders from Outlook."
    LogMessage strTrace, strRoutine
    
    ldb.Disconnect
    
    GoTo Finally
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    fCnt = -1
    
Finally:
    strTrace = "Completed method."
    IndexOutlookFolders = fCnt

End Function

''' Import projects from MCL - Returns # of projects imported
Public Function ImportProjectsFromCategories() As Integer

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ImportProjectsFromCategories"
    
    On Error GoTo ThrowException
    
    Dim pm As New ProjectManager
    Dim myList As Projects
    Set myList = pm.ImportFromMCL
    If myList.Count = 0 Then
        strTrace = "No projects found in the Outlook Master Category List."
        GoTo ThrowException
    End If
    
    Dim ldb As New dsDataStore
    ldb.Connect
    
    Dim arProj As ArrayList
    Set arProj = ldb.GetEntireCollection("Project")
    If arProj Is Nothing Then
        strTrace = "Failed to connect to the DataStore."
        GoTo ThrowException
    End If
    
    Dim pCnt As Integer
    pCnt = 0
    
    Dim p As fmeProject
    Dim ip As fmeProject
    Dim bFnd As Boolean
    
    For Each ip In myList.Items
        bFnd = False
        For Each p In arProj
           If LCase(ip.Subject) = LCase(p.Subject) Then
                bFnd = True
                Exit For
            End If
        Next
        
        If Not bFnd Then
            ldb.Insert ip, "Project"
            pCnt = pCnt + 1
            
            strTrace = "Imported Category '" & ip.Subject & " from Outlook MCL."
            LogMessage strTrace, strRoutine
            
        End If
        
    Next
        
    ldb.Disconnect
    
    ImportProjectsFromCategories = pCnt
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    ImportProjectsFromCategories = -1

End Function
