VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fmeEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Fields

Private Const rootClass As String = "fmeEventArgs"

' Properties

Dim f_dataObject As Object
Dim f_dataVariant As Variant
Dim f_Cancel As Boolean

''' Payload for the EventArgs
Public Property Set DataObject(ByVal dobj As Object)
    Set f_dataObject = dobj
End Property
Public Property Get DataObject() As Object
    Set DataObject = f_dataObject
End Property

Public Property Let DataVariant(ByVal var As Variant)
    f_dataVariant = var
End Property
Public Property Get DataVariant() As Variant
    DataVariant = f_dataVariant
End Property


Public Property Let cancel(ByVal b As Boolean)
    f_Cancel = b
End Property
Public Property Get cancel() As Boolean
    cancel = f_Cancel
End Property

'  Constructor
Private Sub Class_Initialize()
    Set f_dataObject = Nothing
End Sub

Private Sub Class_Terminate()
    Set f_dataObject = Nothing
End Sub

' Methods

Public Sub Create(ByVal dobj As Object, ByVal cancel As Boolean)
    Set f_dataObject = dobj
    f_Cancel = cancel
End Sub

Public Sub CreateVariant(ByVal var As Variant, ByVal cancel As Boolean)
    f_dataVariant = var
    f_Cancel = cancel
End Sub
