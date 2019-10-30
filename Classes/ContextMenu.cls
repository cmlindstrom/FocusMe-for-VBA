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

Dim WithEvents cmi As ContextMenuItem
Attribute cmi.VB_VarHelpID = -1

' - - Events

Public Event Click(ByVal tag As String)
Public Event DblClick(ByVal tag As String, ByVal Cancel As MSForms.ReturnBoolean)
Public Event MouseMoved(ByVal tag As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

' - - Properties

Dim f_form As UserForm

''' Collection of MenuItems for the Menu
Public Property Get Items() As ArrayList
    Set Items = mnuItems
End Property

''' Track the form that contains the menu
Public Property Set Form(ByVal f As UserForm)
    Set f_form = f
End Property
Public Property Get Form() As UserForm
    Set Form = f_form
End Property

' - - Event Handlers


' - - Constructor

Private Sub Class_Initialize()
    Set mnuItems = New ArrayList
End Sub

Private Sub Class_Terminate()
    Set f_form = Nothing
End Sub

' - - Methods

Public Sub AddItem(ByVal mnu As ContextMenuItem)
    Set cmi = mnu
    Set cmi.Menu = Me
    mnuItems.Add cmi
End Sub

Public Sub RemoveItem(ByVal mnu As ContextMenuItem)
    mnuItems.Remove mnu
End Sub

Public Sub RemoveItemAt(ByVal idx As Integer)
    mnuItems.RemoveAt idx
End Sub

''' Click - the ContextMenuItem calls this method to throw the menu event
''' for the specific item
Public Sub Click(ByVal nme As String)
    RaiseEvent Click(nme)
End Sub

Public Sub DoubleClick(ByVal nme As String, ByVal Cancel As MSForms.ReturnBoolean)
    RaiseEvent DblClick(nme, Cancel)
End Sub

Public Sub MouseMoved(ByVal nme As String, ByVal Button As Integer, _
                       ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RaiseEvent MouseMoved(nme, Button, Shift, X, Y)
End Sub
