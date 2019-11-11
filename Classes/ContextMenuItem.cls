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
    Menu = 30
End Enum

Dim f_Name As String

Dim strCaption As String
Dim isEnabled As Boolean
Dim iType As enuMenuItemType
Dim strID As String
Dim strTag As String

Dim hasControl As Boolean
Dim myControl As control

Dim f_ParentMenu As ContextMenu

Dim f_hasSubMenu As Boolean
Dim f_SubMenu As ContextMenu

Dim WithEvents lbl As MSForms.label
Attribute lbl.VB_VarHelpID = -1
Dim WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1

' - - Events

Public Event Click(ByVal sender As ContextMenuItem)
Public Event DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Public Event MouseMoved(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

' - - Properties

''' Name of the CMI (matches attached control name)
Public Property Get Name() As String
    Name = f_Name
End Property

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
Public Property Let Tag(ByVal str As String)
    strTag = str
End Property
Public Property Get Tag() As String
    Tag = strTag
End Property

''' The ContextMenu this item is contained within
Public Property Set ParentMenu(ByVal cm As ContextMenu)
    Set f_ParentMenu = cm
End Property
Public Property Get ParentMenu() As ContextMenu
    Set ParentMenu = f_ParentMenu
End Property

''' The ContextMenu this item launches (if it has a subMenu)
Public Property Set SubMenu(ByVal cm As ContextMenu)
    Set f_SubMenu = cm
    If cm Is Nothing Then
        f_hasSubMenu = False
    Else
        f_hasSubMenu = True
    End If
End Property
Public Property Get SubMenu() As ContextMenu
    Set SubMenu = f_SubMenu
End Property

' - - Event Handlers

''' Labels
Private Sub lbl_Click()
    f_ParentMenu.Click Me
    RaiseEvent Click(Me)
End Sub

Private Sub lbl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    f_ParentMenu.DoubleClick strTag, Cancel
    RaiseEvent DblClick(Cancel)
End Sub

Private Sub lbl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                            ByVal X As Single, ByVal Y As Single)
                            
    If Not iType = Separator Then
        f_ParentMenu.MouseMoved lbl.Name, Button, Shift, X, Y
        RaiseEvent MouseMoved(Button, Shift, X, Y)
    End If
    
End Sub

''' Buttons
Private Sub btn_Click()
    f_ParentMenu.Click Me
    RaiseEvent Click(Me)
End Sub

Private Sub btn_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    f_ParentMenu.DoubleClick strTag, Cancel
    RaiseEvent DblClick(Cancel)
End Sub

Private Sub btn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    f_ParentMenu.MouseMoved btn.Name, Button, Shift, X, Y
    RaiseEvent MouseMoved(Button, Shift, X, Y)
End Sub

' - - Constructor

Private Sub Class_Initialize()
    Call Initialize
End Sub

Private Sub Class_Terminate()
    Set myControl = Nothing
    Set f_Menu = Nothing
    Set f_SubMenu = Nothing
End Sub

Private Sub Initialize()
    hasControl = False
    f_hasSubMenu = False
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
        f_Name = lbl.Name
    End If
    If TypeOf c Is MSForms.CommandButton Then
        Set btn = c
        iType = Button
        strCaption = btn.Caption
        isEnabled = btn.Enabled
        hasControl = True
        f_Name = btn.Name
    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

''' Returns an assigned control
Public Function GetControl() As MSForms.control
    Set GetControl = myControl
End Function

''' Returns a True if CMI has a sub-menu
Public Function HasSubMenu() As Boolean
    HasSubMenu = f_hasSubMenu
End Function

''' Returns the absolute coordinates to launch the SubMenu Form
Public Sub SubMenuLaunchPt(ByRef X As Long, ByRef Y As Long)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SubMenuLaunchPt"
    
    On Error GoTo ThrowException
    
    Dim ufP As UserForm
    Set ufP = f_ParentMenu.Form
    
    Dim fX As Long
    Dim fY As Long
    If Not ufP Is Nothing Then
        GetFormPosition ufP, fX, fY
    Else
        strTrace = "Parent Menu Form not set."
        GoTo ThrowException
    End If
    
    Dim cX As Long
    Dim cY As Long
    Dim cW As Long
    If Not myControl Is Nothing Then
        cX = myControl.Left
        cY = myControl.Top
        cW = myControl.Width
    Else
        strTrace = "No control container found."
        GoTo ThrowException
    End If
    
    X = fX + cX + cW + 26 + 18
    Y = fY + cY + 52
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub
