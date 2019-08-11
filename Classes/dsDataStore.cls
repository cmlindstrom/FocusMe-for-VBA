VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsDataStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' - Fields

Private Const rootClass As String = "dsDataStore"

Private Const collectionNames As String = "Project,Workspace,Focus"

''' Property internal variables
Dim f_source As String
Dim f_name As String

Dim f_lastSaved As Date
Dim f_lastTouched As Date

Dim dsInternal As DataSet


' - Events

Public Event ItemAdded()
Public Event ItemUpdated()
Public Event ItemDeleted()

' - Properties

''' the name part of a file to the xml file path
Public Property Get Name() As String
    Name = f_name
End Property

''' The file where the DataSet is saved
'''     set via the Connect method
Public Property Get Source() As String
    Source = f_source
End Property

''' Returns the date and time the last
''' time the DataSet was saved to its storage file
Public Property Get LastCommitDate() As Date
    LastCommitDate = f_lastSaved
End Property

''' Returns the date and time the last
''' change was made to the DataSet
Public Property Get LastUpdate() As Date
    LastUpdate = f_lastTouched
End Property

' - Constructor

Private Sub Class_Initialize()
    ' Create a DataSet
    Set dsInternal = New DataSet
End Sub

' - Methods

''' Creates a connection for the specified DataSet
Public Sub Connect(Optional ByVal dsName As String = "")

    If Len(dsName) = 0 Then dsName = "Default"

    f_name = CleanFileName(dsName)
    
    Dim Path As String
    Path = GetAppDataPath()
    f_name = dsName
    f_source = Path & "\" & f_name & ".xml"
    
    LoadfromFile f_source
    
End Sub

''' Allows the ability to set a dsDataSource
''' directly to a file path
Public Sub SetConnection(ByVal fullFilePath As String)

    If Len(fullFilePath) = 0 Then
        LogMessageEx "Invalid or empty data source.", Nothing, "dsDataSource:SetConnection"
        Exit Sub
    End If

    ' Reset the current DataSet
    dsInternal.Reset
    
    ' Load the file content
    f_source = fullFilePath
    LoadfromFile f_source
    
End Sub

''' Commits any in-memory changes to the Source
Public Sub AcceptChanges()
    SaveToFile f_source
End Sub

Public Sub Disconnect()
    ' Save to xml file
    SaveToFile f_source
End Sub

' - - Item Management

''' Saves the obj to the specified table
'''  - uses the Created date to differentiate between new and existing items
Public Function Save(ByVal obj As Object, ByVal tblName As String) As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Save"

    Dim bReturn As Boolean
    bReturn = False
       
    On Error GoTo ThrowException

    Dim dte As Date
    dte = GetProperty(obj, "Created")
    If IsDateNone(dte) Then
        ' Insert the new object
        bReturn = Insert(obj, tblName)
    Else
        ' Update the existing object
        bReturn = Update(obj, tblName)
    End If

    Save = bReturn
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Save = False

End Function

''' Inserts a new record in the specified Table
Public Function Insert(ByVal obj As Object, ByVal tblName As String) As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Insert"
    
    If obj Is Nothing Then
        strTrace = "A null object instance encountered."
        GoTo ThrowException
    End If
    If Len(tblName) = 0 Then
        strTrace = "Table name not specified."
        GoTo ThrowException
    End If
    
    On Error GoTo ThrowException

    strTrace = "Find the table in the dataSet."
    Dim dt As DataTable
    Set dt = dsInternal.GetTableByName(tblName)
    If dt Is Nothing Then
        Set dt = New DataTable
        dt.Name = tblName
        
        dsInternal.Tables.Add dt
    End If

    strTrace = "Get the object's property names."
    Dim props() As String
    props = obj.GetListOfProperties
    
    strTrace = "Check columns & update column collection."
    Dim s As Variant
    For Each s In props
        If Not dt.Columns.IsColumnExist(s) Then
            dt.Columns.Add s
        End If
    Next
    
    strTrace = "Check to assure a new object is being inserted."
    Dim dte As Date
    dte = GetProperty(obj, "Created")
    If Not IsDateNone(dte) Then
        strTrace = "Attempting to insert an object that has already been created."
        GoTo ThrowException
    End If
    
    strTrace = "Configure New Object."
    Dim myId As String
    myId = GenerateUniqueID(8)
    SetProperty obj, "id", myId
    SetProperty obj, "Created", Now
    SetProperty obj, "Modified", Now
       
    strTrace = "Insert a new Row into the table."
    Dim dr As New DataRow
    For Each s In props
        dr.Add s, GetProperty(obj, s)
    Next
    dt.Rows.AddRow dr
    
    strTrace = "Capture activity timestamp."
    f_lastTouched = Now
    
    Insert = True
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Insert = False
    
End Function

''' Updates a current item in the dataStore
Public Function Update(ByVal obj As Object, ByVal tblName As String) As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Update"
    
    If obj Is Nothing Then
        strTrace = "A null object instance encountered."
        GoTo ThrowException
    End If
    If Len(tblName) = 0 Then
        strTrace = "Table name not specified."
        GoTo ThrowException
    End If
    
    On Error GoTo ThrowException
    
    strTrace = "Check for existing id."
    Dim myId As Variant
    myId = GetProperty(obj, "id")
    If Len(myId) = 0 Then
        strTrace = "Attempting to update an item that does not yet exist."
        GoTo ThrowException
    End If
    
    strTrace = "Find the table in the dataSet."
    Dim dt As DataTable
    Set dt = dsInternal.GetTableByName(tblName)
    If dt Is Nothing Then
        strTrace = "Failed to find the referenced table: " & tblName
        GoTo ThrowException
    End If
    
    strTrace = "Locate the item of interest."
    Dim drUpdate As DataRow
    Set drUpdate = dt.Rows.GetByColumnValue("id", myId)
    If drUpdate Is Nothing Then
        strTrace = "Failed to locate the item in the collection, id: " & myId
        GoTo ThrowException
    End If
    
    strTrace = "Get the object's property names."
    Dim props() As String
    props = obj.GetListOfProperties
    
    strTrace = "Update the dataRow values."
    Dim s As Variant
    For Each s In props
        If Not s = "id" Then
            strTrace = "Updating property: " & s & "."
            Dim v As Variant
            v = GetProperty(obj, s)
            drUpdate.SetItem s, v
        End If
    Next
    
    ' Update modified timestamp
    drUpdate.SetItem "Modified", Now
    
    ' Do we need to remove and add in the collection? - No

    strTrace = "Capture activity timestamp."
    f_lastTouched = Now
    
    Update = True
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Update = False
    
End Function

''' Marks a dataStore item as deleted
Public Sub Delete(ByVal obj As Object, ByVal tblName As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Delete"
    
    If obj Is Nothing Then
        strTrace = "A null object instance encountered."
        GoTo ThrowException
    End If
    If Len(tblName) = 0 Then
        strTrace = "Table name not specified."
        GoTo ThrowException
    End If
    
    On Error GoTo ThrowException
    
    ' Set the Deleted date to today
    SetProperty obj, "Deleted", Date
    
    ' Update the item
    Update obj, tblName
    
    strTrace = "Capture activity timestamp."
    f_lastTouched = Now
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Removes the item with the specified id from the dataStore
Public Function Remove(ByVal id As String, ByVal tblName As String) As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Remove"
    
    If Len(id) = 0 Then
        strTrace = "Id field missing."
        GoTo ThrowException
    End If
    If Len(tblName) = 0 Then
        strTrace = "Table name not specified."
        GoTo ThrowException
    End If
    
    On Error GoTo ThrowException
    
    strTrace = "Find the table in the dataSet."
    Dim dt As DataTable
    Set dt = dsInternal.GetTableByName(tblName)
    If dt Is Nothing Then
        strTrace = "Failed to find the referenced table: " & tblName
        GoTo ThrowException
    End If
    
    strTrace = "Locate the item of interest."
    Dim drRemove As DataRow
    Set drRemove = dt.Rows.GetByColumnValue("id", id)
    If drRemove Is Nothing Then
        strTrace = "Failed to locate the item in the collection, id: " & id
        GoTo ThrowException
    End If
    
    strTrace = "Remove it from the collection."
    dt.Rows.Remove drRemove

    Remove = True
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Remove = False
    
End Function

''' Gets an item from the dataStore, typed to the originating Table
Public Function GetItemById(ByVal id As String, ByVal tblName As String) As Object

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetItemById"

    If Len(id) = 0 Then
        strTrace = "Id field missing."
        GoTo ThrowException
    End If
    If Len(tblName) = 0 Then
        strTrace = "Table name not specified."
        GoTo ThrowException
    End If
    
    On Error GoTo ThrowException
    
    strTrace = "Find the table in the dataSet."
    Dim dt As DataTable
    Set dt = dsInternal.GetTableByName(tblName)
    If dt Is Nothing Then
        strTrace = "Failed to find the referenced table: " & tblName
        GoTo ThrowException
    End If
    
    strTrace = "Locate the item of interest."
    Dim dr As DataRow
    Set dr = dt.Rows.GetByColumnValue("id", id)
    If dr Is Nothing Then
        strTrace = "Failed to locate the item in the collection, id: " & id
        GoTo ThrowException
    End If
    
    strTrace = "Create an object instance from the dataRow values."
    Dim obj As Object
    Set obj = CreateNewObject(tblName)
    
    strTrace = "Get the object's property names."
    Dim props() As String
    props = obj.GetListOfProperties
    
    strTrace = "Retrieve the dataRow values."
    Dim s As Variant
    For Each s In props
        SetProperty obj, s, dr.GetItem(s)
    Next
    
    Set GetItemById = obj
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetItemById = Nothing
    
End Function

' - - Collections

''' Returns ALL of the objects in the specified collection, including any
''' that are deleted
Public Function GetEntireCollection(ByVal collName As String) As ArrayList

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetEntireCollection"

    If Len(collName) = 0 Then
        strTrace = "Table name not specified."
        GoTo ThrowException
    End If
    
    On Error GoTo ThrowException
    
    Dim retColl As ArrayList
    Set retColl = New ArrayList
   
    strTrace = "Find the table in the dataSet."
    Dim dt As DataTable
    Set dt = dsInternal.GetTableByName(collName)
    If dt Is Nothing Then
        strTrace = "Failed to find the referenced table: " & collName
        GoTo ThrowException
    End If
    
    strTrace = "Retrieve the table rows."
    Dim dr As DataRow
    Dim obj As Object
    For Each dr In dt.Rows.Items
    
        strTrace = "Create an object instance from the dataRow values."
        Set obj = CreateNewObject(collName)
    
        strTrace = "Get the object's property names."
        Dim props() As String
        props = obj.GetListOfProperties
    
        strTrace = "Retrieve the dataRow values."
        Dim s As Variant
        For Each s In props
            SetProperty obj, s, dr.GetItem(s)
        Next
        
        strTrace = "Capture the filled object."
        retColl.Add obj
    
    Next

    Set GetEntireCollection = retColl
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetEntireCollection = New ArrayList

End Function

''' Removes all items from the specified collection
Public Sub ClearEntireCollection(ByVal collName As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ClearEntireCollection"
    
    On Error GoTo ThrowException
    
    strTrace = "Find the table in the dataSet."
    Dim dt As DataTable
    Set dt = dsInternal.GetTableByName(collName)
    If dt Is Nothing Then
        strTrace = "Failed to find the referenced table: " & collName
        GoTo ThrowException
    End If
    
    strTrace = "Clear the Rows from the table."
    Call dt.Rows.Clear
        
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Returns objects within the specified collection that are not marked as deleted.
''' A filter can be specified to further narrow the returned collection
Public Function GetCollection(ByVal collName As String, _
                    Optional ByVal filter As String = "") As ArrayList
                    
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetCollection"

    If Len(collName) = 0 Then
        strTrace = "Table name not specified."
        GoTo ThrowException
    End If
    
    On Error GoTo ThrowException

    Dim retColl As ArrayList
    Set retColl = New ArrayList
    
    Dim tmpColl As ArrayList
    Dim obj As Object
    If Len(filter) = 0 Then
        Set tmpColl = GetEntireCollection(collName)
        If IsNothing(tmpColl) Then
            strTrace = "Collection query failed."
            GoTo ThrowException
        End If
        
        ' Tease out any deleted items
        For Each obj In tmpColl
            Dim dte As Date
            dte = GetProperty(obj, "Deleted")
            If IsDateNone(dte) Then
                retColl.Add obj
            End If
        Next
        
    Else
        ' Not implemented
    End If
    
    Set GetCollection = retColl
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetCollection = New ArrayList

End Function


' - - Utilities

''' Formats a string so it can be used as part of a file path
Public Function CleanFileName(ByVal nme As String) As String

    Dim strReserved As String
    strReserved = "\ / : * ? """" < > "
    
    On Error Resume Next
    
    Dim chrs() As String
    chrs = Split(strReserved, " ")
    
    Dim strReturn As String
    strReturn = nme
    
    Dim c As Variant
    For Each c In chrs
        strReturn = Replace(strReturn, c, "_")
    Next
    strReturn = Replace(strReturn, " ", "_")
    
    CleanFileName = strReturn

End Function

' - Supporting Methods

Private Sub LoadfromFile(ByVal fullPath As String)
    dsInternal.ReadXmlFile fullPath
End Sub

Private Sub SaveToFile(ByVal fullPath As String)
    dsInternal.WriteXmlFile fullPath
    f_lastSaved = Now
End Sub

Private Function CreateNewObject(ByVal tblName As String) As Object

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":CreateNewObject"
    
    If Len(tblName) = 0 Then
        strTrace = "Table name is empty."
        GoTo ThrowException
    End If
    
    On Error GoTo ThrowException
    
    Dim objReturn As Object
    Set objReturn = Nothing
    
    Select Case LCase(tblName)
        Case "project"
            Set objReturn = New fmeProject
        Case "folder"
            Set objReturn = New fmeFolder
        Case "mail"
            Set objReturn = New fmeMailItem
        Case Else
            strTrace = "Table name not recognized: " & tblName
            GoTo ThrowException
    End Select
    
    Set CreateNewObject = objReturn
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set CreateNewObject = Nothing

End Function
