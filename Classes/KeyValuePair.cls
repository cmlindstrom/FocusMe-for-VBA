VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyValuePair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Fields

Private Const rootClass As String = "KeyValuePair"

Private dictKey As String
Private dictValue As Variant

' Events

' Properties

''' The name of the Key for the pair
Public Property Let Key(ByVal keyName As String)
    dictKey = keyName
End Property
Public Property Get Key() As String
    Key = dictKey
End Property

''' The value for the pair
Public Property Let Value(ByVal val As Variant)
    dictValue = val
End Property
Public Property Get Value() As Variant
    Value = dictValue
End Property

' Constructor

Private Sub Class_Initialize()
    dictKey = ""
    Set dictValue = Nothing
End Sub

' Methods

' Supporting Methods

