VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' Fields

Private Const rootClass As String = "Setting"

Private strName As String
Private oValue As Variant
Private vType As Integer

''' Properties

' Setting Name
Public Property Let Name(ByVal nme As String)
    strName = nme
End Property
Public Property Get Name() As String
    Name = strName
End Property

' Setting Value
Public Property Let Value(ByVal val As Variant)
    oValue = val
End Property
Public Property Get Value() As Variant
    Value = oValue
End Property

' Value Type
Public Property Let ValueType(ByVal typ As Integer)
    vType = typ
End Property
Public Property Get ValueType() As Integer
    ValueType = vType
End Property


''' Constructor

Private Sub Class_Initialize()
    strName = ""
    oValue = ""
End Sub

