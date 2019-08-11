VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContextMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' - - Fields

Private Const rootClass As String = "ContextMenu"

Dim mnuItems As ArrayList

' - - Properties

''' Collection of MenuItems for the Menu
Public Property Get Items() As ArrayList
    Set Items = mnuItems
End Property

' - - Constructor

Private Sub Class_Initialize()
    Set mnuItems = New ArrayList
End Sub

' - - Methods

Public Sub AddItem(ByVal mnu As ContextMenuItem)
    mnuItems.Add mnu
End Sub

Public Sub RemoveItem(ByVal mnu As ContextMenuItem)
    mnuItems.Remove mnu
End Sub

Public Sub RemoveItemAt(ByVal idx As Integer)
    mnuItems.RemoveAt idx
End Sub
