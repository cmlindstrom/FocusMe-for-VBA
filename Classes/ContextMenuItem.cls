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
    label = 2
    Separator = 20
End Enum

Dim strCaption As String
Dim isEnabled As Boolean
Dim iType As enuMenuItemType
Dim strID As String
Dim strTag As String
Dim hasControl As Boolean
Dim myControl As control
Dim myMenu As ContextMenu

Dim WithEvents lbl As MSForms.label
Attribute lbl.VB_VarHelpID = -1
Dim WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1

' - - Events

Public Event Click()
Public Event DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Public Event MouseMoved(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

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
Public Property Let itemType(ByVal t As enuMenuItemType)
    iType = t
End Property
Public Property Get itemType() As enuMenuItemType
    itemType = iType
End Property

''' Unique Identifier - used to process the selection of the menu item
Public Property Let UID(ByVal id As String)
    strID = id
End Property
Public Property Get UID() As String
    UID = strID
End Property

''' A tag property
Public Property Let tag(ByVal str As String)
    strTag = str
End Property
Public Property Get tag() As String
    tag = strTag
End Property

''' The ContextMenu this item is contained within
Public Property Set Menu(ByVal cm As ContextMenu)
    Set myMenu = cm
End Property
Public Property Get Menu() As ContextMenu
    Set Menu = myMenu
End Property

' - - Event Handlers

''' Labels
Private Sub lbl_Click()
    myMenu.Click strTag
    RaiseEvent Click
End Sub

Private Sub lbl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    myMenu.DoubleClick strTag, Cancel
    RaiseEvent DblClick(Cancel)
End Sub

Private Sub lbl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                            ByVal X As Single, ByVal Y As Single)
    myMenu.MouseMoved strTag, Button, Shift, X, Y
    RaiseEvent MouseMoved(Button, Shift, X, Y)
End Sub

''' Buttons
Private Sub btn_Click()
    myMenu.Click strTag
    RaiseEvent Click
End Sub

Private Sub btn_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    myMenu.DoubleClick strTag, Cancel
    RaiseEvent DblClick(Cancel)
End Sub

Private Sub btn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    myMenu.MouseMoved strTag, Button, Shift, X, Y
    RaiseEvent MouseMoved(Button, Shift, X, Y)
End Sub

' - - Constructor

Private Sub Class_Initialize()
    Call Initialize
End Sub

Private Sub Class_Terminate()
    Set myControl = Nothing
End Sub

Private Sub Initialize()
    hasControl = False
    isEnabled = True
    Set myControl = Nothing
End Sub

' - - Methods

''' Sets a MSForms Control
Public Sub SetControl(ByVal c As MSForms.control)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SetControl"
    
    On Error GoTo ThrowException
    
    Set myControl = c
    
    If TypeOf c Is MSForms.label Then
        Set lbl = c
        iType = label
        strCaption = lbl.Caption
        isEnabled = lbl.Enabled
        hasControl = True
    End If
    If TypeOf c Is MSForms.CommandButton Then
        Set btn = c
        iType = Button
        strCaption = btn.Caption
        isEnabled = btn.Enabled
        hasControl = True
    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

''' Returns an assigned control
Public Function GetControl() As MSForms.control
    Set GetControl = myControl
End Function


