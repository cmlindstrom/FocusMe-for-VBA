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
Public Property Let value(ByVal val As Variant)
    oValue = val
End Property
Public Property Get value() As Variant
    value = oValue
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

''' Methods

''' Compares argument to setting value
''' Returns True if equal otherwise false
Public Function Compare(ByVal v As Variant) As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Compare"
    
    Dim bReturn As Boolean
    bReturn = True
    
    On Error GoTo ThrowException
    
    If VarType(v) = vbString Then
        bReturn = (StrComp(v, oValue) = 0)
    Else
        bReturn = (v = oValue)
    End If
    
    Compare = bReturn
    Exit Function
        
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Compare = False

End Function

