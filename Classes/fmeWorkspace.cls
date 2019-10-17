VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fmeWorkspace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' - Fields

Private Const rootClass As String = "fmeWorkspace"

' - Properties

Private f_Id As String
Private f_Name As String
Private f_Code As String
Private f_description As String

Private f_Created As Date
Private f_Modified As Date
Private f_Deleted As Date

Private Const fields As String = "id,Name,Code,Description," & _
                                    "Created,Modified,Deleted"

''' Unique Identifier for the object
Public Property Let id(ByVal myId As String)
    f_Id = myId
End Property
Public Property Get id() As String
    id = f_Id
End Property

''' Name or Title for the Workspace
Public Property Let name(ByVal myName As String)
    f_Name = myName
End Property
Public Property Get name() As String
    name = f_Name
End Property

''' User Code - for organizing / sorting
Public Property Let Code(ByVal myCode As String)
    f_Code = myCode
End Property
Public Property Get Code() As String
    Code = f_Code
End Property

''' Gets / Sets the description
Public Property Let Description(ByVal myDesc As String)
    f_description = myDesc
End Property
Public Property Get Description() As String
    Description = f_description
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

' - Events

' - Event Handlers

' - Constructor

Private Sub Class_Initialize()

    ' Initialize
    f_Id = "W" & Common.GenerateUniqueID(6)

    ' Defaults
    f_Name = ""
    
    f_Created = #1/1/1970#
    f_Modified = #1/1/1970#
    f_Deleted = #1/1/1970#
      
End Sub

' - Methods

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
