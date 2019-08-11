VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fmeFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' - Fields

Private Const rootClass As String = "fmeFolder"

Private f_Id As String
Private f_entryId As String
Private f_parentEid As String
Private f_name As String
Private f_path As String
Private f_description As String

Private f_itemType As Integer
Private f_messageClass As String

Private f_Created As Date
Private f_Modified As Date
Private f_Deleted As Date

Private Const fields As String = "id,EntryId,ParentEntryId,Name,Path,Description,ItemType,MessageClass," & _
                                    "Created,Modified,Deleted"

' - Properties

''' Unique Identifier for the object
Public Property Let id(ByVal myId As String)
    f_Id = myId
End Property
Public Property Get id() As String
    id = f_Id
End Property

''' Outlook EntryID for this folder
Public Property Let EntryId(ByVal eid As String)
    f_entryId = eid
End Property
Public Property Get EntryId() As String
    EntryId = f_entryId
End Property

''' Outlook EntryID for the parent of this folder
Public Property Let ParentEntryId(ByVal eid As String)
    f_parentEid = eid
End Property
Public Property Get ParentEntryId() As String
    ParentEntryId = f_parentEid
End Property

''' Folder name
Public Property Let Name(ByVal myName As String)
    f_name = myName
End Property
Public Property Get Name() As String
    Name = f_name
End Property

''' Gets / Sets the folder description
Public Property Let Description(ByVal myDesc As String)
    f_description = myDesc
End Property
Public Property Get Description() As String
    Description = f_description
End Property

''' Folder path
Public Property Let Path(ByVal fldrPath As String)
    f_path = fldrPath
End Property
Public Property Get Path() As String
    Path = f_path
End Property

''' Default Item type for the items in the folder - follows Outlook.olItemType
Public Property Let ItemType(ByVal itmType As Integer)
    f_itemType = itmType
End Property
Public Property Get ItemType() As Integer
    ItemType = f_itemType
End Property

''' Default message class for the items in the folder
Public Property Let MessageClass(ByVal mcls As String)
    f_messageClass = mcls
End Property
Public Property Get MessageClass() As String
    MessageClass = f_messageClass
End Property

' - - Basic Properties

''' Date and Time the Item was created the first time
Public Property Let Created(ByVal dte As Date)
    f_Created = dte
End Property
Public Property Get Created() As Date
    Created = f_Created
End Property

''' Date and Time the Item was last modified in the DataStore
Public Property Let Modified(ByVal dte As Date)
    f_Modified = dte
End Property
Public Property Get Modified() As Date
    Modified = f_Modified
End Property

''' Date and Time the Item was marked as deleted in the DataStore
Public Property Let Deleted(ByVal dte As Date)
    f_Deleted = dte
End Property
Public Property Get Deleted() As Date
    Deleted = f_Deleted
End Property

' - Constructor

Private Sub Class_Initialize()

    ' Initialize
    f_Id = Common.GenerateUniqueID(6)
    
    f_name = ""
    f_entryId = ""
    f_parentEid = ""
    f_path = ""
    f_description = ""
    f_messageClass = ""
    
    ' Defaults
    f_Created = #1/1/1970#
    f_Modified = #1/1/1970#
    f_Deleted = #1/1/1970#
    
    SetItemType olMailItem
    
End Sub

' - Methods

''' Sets the ItemType using the Outlook enumeration
Public Sub SetItemType(ByVal itmType As Outlook.OlItemType)
    f_itemType = itmType
End Sub

''' Gets the Outlook enumeration for the default folder item type
Public Function GetItemType() As Outlook.OlItemType
    GetItemType = f_itemType
End Function

''' Creates an instance of this object from the Outlook Folder
Public Sub Fill(ByVal f As Outlook.Folder)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Fill_Outlook"
    
    On Error GoTo ThrowException
    strTrace = "Started method."
    
    If f Is Nothing Then
        strTrace = "A null Outlook folder encountered."
        GoTo ThrowException
    End If
    
    strTrace = "Capture parent folder."
    Dim pf As Outlook.Folder
    Set pf = f.Parent
    If Not pf Is Nothing Then SetProperty Me, "ParentEntryId", pf.EntryId
    
    strTrace = "Capture strings."
    f_name = f.Name
    f_entryId = f.EntryId
    f_description = f.Description
    f_messageClass = f.Class
    
    strTrace = "Capture ItemType."
    f_itemType = f.DefaultItemType
    
    strTrace = "Capture folder path."
    f_path = f.FolderPath
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

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


