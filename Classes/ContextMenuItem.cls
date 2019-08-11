VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContextMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' - - Fields

Private Const rootClass As String = "ContextMenuItem"

Public Enum enuMenuItemType
    Button = 1
    Separator = 2
End Enum

Dim strCaption As String
Dim isEnabled As Boolean
Dim iType As enuMenuItemType
Dim strID As String

' - - Properties

''' Menu Title or Caption
Public Property Let Caption(ByVal Name As String)
    strCaption = Name
End Property
Public Property Get Caption() As String
    Caption = strCaption
End Property

''' Menu Item Enabled, True = yes
Public Property Let Enabled(ByVal b As Boolean)
    isEnabled = b
End Property
Public Property Get Enabled() As Boolean
    Enabled = isEnabled
End Property

''' Menu Item type, e.g. Button or Separator
Public Property Let itemtype(ByVal t As enuMenuItemType)
    iType = t
End Property
Public Property Get itemtype() As enuMenuItemType
    itemtype = iType
End Property

''' Unique Identifier - used to process the selection of the menu item
Public Property Let UID(ByVal id As String)
    strID = id
End Property
Public Property Get UID() As String
    UID = strID
End Property

' - - Constructor

Private Sub Class_Initialize()

End Sub

' - - Methods

