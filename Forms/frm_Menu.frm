VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Menu 
   Caption         =   "UserForm1"
   ClientHeight    =   3750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2415
   OleObjectBlob   =   "frm_Menu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' - Reference
' https://stackoverflow.com/questions/48382957/how-to-add-events-to-dynamically-created-controls-buttons-listboxes-in-excel

' - Fields

Private Const rootClass As String = "frm_Menu"

Dim leftMouseDown As Boolean
Dim rightMouseDown As Boolean

Private Const BackgroundNormal As Long = &H8000000E
Private Const BackgroundHighlight As Long = &H8000000D
Private Const LineColor As Long = &H8000000A

Public Enum enuControlType
    None = 0
    Labels = 1
    buttons = 2
End Enum

' Form
Dim frmX As Integer
Dim frmY As Integer

Dim frmLocX As Long
Dim frmLocY As Long

' - Events

Public Event Click(ByVal Tag As String)

' - Properties

Dim WithEvents f_cm As ContextMenu
Attribute f_cm.VB_VarHelpID = -1
Dim f_miType As enuControlType
Dim f_cmi As ContextMenuItem

''' ContextMenu to use in the form
Public Property Set Menu(ByVal cm As ContextMenu)
    Set f_cm = cm
    Set f_cm.Form = Me
    Call BuildMenu
End Property
Public Property Get Menu() As ContextMenu
    Set Menu = f_cm
End Property

''' Allows Labels or Buttons to be used on MenuItems
Public Property Let MenuItemType(ByVal typ As enuControlType)
    f_miType = typ
End Property
Public Property Get MenuItemType() As enuControlType
    MenuItemType = f_miType
End Property

''' Selected contextMenuItem
Public Property Get MenuItem() As ContextMenuItem
    Set MenuItem = f_cmi
End Property


' - Event Handlers

''' ContextMenu Events

Private Sub f_cm_Click(ByVal sender As ContextMenu, ByVal Tag As String)

    ' SubMenu launch coordinates
    Dim lX As Long
    Dim lY As Long
    
    Dim cm As ContextMenu
    Set cm = sender

    ' Captured clicked on ContextMenuItem
    Set f_cmi = cm.Selection
    If f_cmi.HasSubMenu Then
    
        ' Get Menu location
        f_cmi.SubMenuLaunchPt lX, lY
    
        ' Menu Node - Display sub-menu
        Tag = ShowMenu(f_cmi.SubMenu, lX, lY)
        
        ' Returns -1 if Menu is cancelled
        If CLng(Tag) > 0 Then
            Set f_cmi = f_cmi.SubMenu.Selection
            RaiseEvent Click(Tag)
            Me.Hide
        End If
        
    Else
        ' Leaf Node - execute the command
        RaiseEvent Click(Tag)
        Me.Hide
             ' MsgBox "Control clicked: " & Tag & " UID: " & f_cmi.UID
    End If

End Sub

Private Sub f_cm_DblClick(ByVal Tag As String, ByVal Cancel As MSForms.ReturnBoolean)
    ' Dim c As control
    ' Set c = FindControl(Tag)
    ' MsgBox "Control double clicked: " & Tag
End Sub

Private Sub f_cm_MouseMoved(ByVal Tag As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim c As control
    Set c = FindControl(Tag)
    Highlight c
End Sub

''' Form Events

Private Sub UserForm_Activate()
    SetFormPosition Me, frmLocX, frmLocY
End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    ' Allow Escape from Menu
    Dim strTrace As String
    strTrace = "Key pressed: " & KeyAscii
    If KeyAscii = 27 Then
        Unload Me
    End If
    
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal X As Single, ByVal Y As Single)
    frmX = X
    frmY = Y
End Sub

' - Constructor

Private Sub UserForm_Initialize()
    HideTitleBar Me
    Initialize
End Sub

Private Sub UserForm_Terminate()
    Set f_cm = Nothing
    Set f_cmi = Nothing
End Sub

Private Sub Initialize()
    f_miType = Labels
    Set f_cmi = Nothing
    
    frmLocX = 10
    frmLocY = 10
    
    ' BuildDefaultMenu
End Sub

' - Methods

''' Places the menu at the specific coordinates
Public Sub SetPosition(ByVal X As Long, ByVal Y As Long)
    frmLocX = X
    frmLocY = Y
    LogMessage "Set position for Menu form (" & X & "," & Y & ")", rootClass & ":SetPosition"
End Sub

''' Mostly for testing purposes
Public Sub ShowDefaultMenu()
    Call BuildDefaultMenu
End Sub

' - Supporting Methods

Private Sub BuildMenu()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":BuildMenu"
    
    On Error GoTo ThrowException
    
    If f_cm Is Nothing Then
        strTrace = "No ContextMenu specified."
        GoTo ThrowException
    End If
    
    Dim cmi As ContextMenuItem
    Dim tmpCtrl As control
    Dim strClassName As String
    Dim strName As String
    
    Dim iRowHgt As Integer
    iRowHgt = 16
    Dim iNextY As Integer
    iNextY = 0
    
    Dim i As Integer
    i = 0
    For Each cmi In f_cm.Items
    
        'Dim lbl As MSForms.label
        'Set lbl = New MSForms.label
        'lbl.BorderStyle = fmBorderStyleSingle
        'lbl.BorderColor = LineColor
        
        Set cmi.ParentMenu = f_cm
        
        Dim f As New MSForms.NewFont
        f.SIZE = 10
        
        If f_miType = Labels Then
            ' Add Label control
            strClassName = "Forms.Label.1"
            strName = "Label" & i
            Set tmpCtrl = Me.Controls.Add(strClassName, strName)
            With tmpCtrl
                .Left = Me.img_Sidebar.Width
                .Width = Me.Width - .Left
                .Top = iNextY ' (X * 20) - 18 'You might have to adjust this spacing.  I just made it up.
                .SpecialEffect = fmSpecialEffectFlat ' fmSpecialEffectEtched
                .Height = iRowHgt
                .Tag = strName
                If cmi.itemType = Separator Then
                    .Height = 1
                    .BorderStyle = fmBorderStyleSingle
                    .BorderColor = LineColor
                End If
                Set .font = f
                
                .Caption = cmi.Caption
                
                If cmi.HasSubMenu Then
                    .Caption = FillSubMenuLabel(.Caption, .font, Me.Width)
                End If
                
            End With
                        
        Else
            ' Add Button control
            strClassName = "Forms.CommandButton.1"
            strName = "btn" & i
            Set tmpCtrl = Me.Controls.Add(strClassName, strName)
            With tmpCtrl
                .Left = Me.img_Sidebar.Width
                .Width = Me.Width - .Left
                .Height = iRowHgt
                .Top = iNextY ' (X * 20) - 18 'You might have to adjust this spacing.  I just made it up.
                .Caption = cmi.Caption
            End With
        End If
                      
        ' Add to ContextMenuItem
        cmi.SetControl tmpCtrl
        ' cmi.UID = tmpCtrl.Name
            
        i = i + 1
        iNextY = iNextY + tmpCtrl.Height + 2
    Next
    
    ' Set the menu height
    Me.Height = iNextY + 4
    Me.img_Sidebar.Height = Me.Height - 4
    
    ' Reset location if beyond bottom of the screen
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Function FindControl(ByVal nme As String) As control

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":FindControl"
    
    On Error GoTo ThrowException
    
    Dim ctrlReturn As control
    Set ctrlReturn = Nothing
    
    Dim ctrl As control
    For Each ctrl In Me.Controls
        If LCase(ctrl.Name) = LCase(nme) Then
            Set ctrlReturn = ctrl
            Exit For
        End If
    Next
    
    Set FindControl = ctrlReturn
    
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set FindControl = Nothing
    
End Function

Private Sub Highlight(ByVal c As control)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Highlight"
    
    On Error GoTo ThrowException
    
    If c Is Nothing Then
        strTrace = "A null control was passed."
        GoTo ThrowException
    End If

    Dim ctrl As control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is label Then
            Dim lbl As label
            Set lbl = ctrl
            If lbl.Tag = c.Tag Then
                lbl.BackColor = BackgroundHighlight
            Else
                lbl.BackColor = BackgroundNormal
            End If
        End If
        If TypeOf ctrl Is CommandButton Then
            Dim btn As CommandButton
            Set btn = ctrl
            If btn.Caption = c.Caption Then
                btn.BackColor = BackgroundHighlight
            Else
                btn.BackColor = BackgroundNormal
            End If
        End If
        
    Next

    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Private Function FillSubMenuLabel(ByVal label As String, ByVal f As MSForms.NewFont, ByVal w As Integer) As String

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":FillSubMenuLabel"
    
    On Error GoTo ThrowException

    Dim strReturn As String
    strReturn = label
    
    Dim iCnt As Integer
    iCnt = 1
    
    Dim strPad As String
    strPad = " " & ChrW(10148)
    strReturn = label & strPad
    
    Dim tot As Integer
    tot = MeasureString(strReturn, f)
    Do While tot < w
        strPad = PadLeft(ChrW(10148), iCnt, " ")
        strReturn = label & strPad
        
        iCnt = iCnt + 1
        tot = MeasureString(strReturn, f)
    Loop
    
    FillSubMenuLabel = strReturn

    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    FillSubMenuLabel = label & ".er"
   
End Function

''' DEPRACATED

Private Sub BuildDefaultMenu()

    Set f_cm = New ContextMenu
    
    Dim cmi As ContextMenuItem
    
    ' File
    Set cmi = New ContextMenuItem
    cmi.Caption = "File"
    cmi.itemType = enuMenuItemType.label
    cmi.Tag = "File"
    f_cm.AddItem cmi
    
    ' Open
    Set cmi = New ContextMenuItem
    cmi.Caption = "Open"
    cmi.itemType = enuMenuItemType.label
    cmi.Tag = "Open"
    f_cm.AddItem cmi
    
    ' Separator
    Set cmi = New ContextMenuItem
    cmi.itemType = enuMenuItemType.Separator
    f_cm.AddItem cmi
    
    ' Properties
    Set cmi = New ContextMenuItem
    cmi.Caption = "Properties..."
    cmi.itemType = enuMenuItemType.label
    cmi.Tag = "Properties"
    f_cm.AddItem cmi

    Call BuildMenu
    
End Sub

Private Sub BuildDefaultMenuOld()

    Set f_cm = New ContextMenu
    
    Dim cmi As ContextMenuItem
    Dim tmpCtrl As control
    Dim strClassName As String
    Dim strName As String
    
    Dim i As Integer
    For i = 1 To 4
    
        Set cmi = New ContextMenuItem
        
        If f_miType = Labels Then
            ' Add Label control
            strClassName = "Forms.Label.1"
            strName = "Label" & i
            Set tmpCtrl = Me.Controls.Add(strClassName, strName)
            With tmpCtrl
                .Left = Me.img_Sidebar.Width
                .Width = Me.Width - .Left
                .Height = 18
                .Top = (i - 1) * 16 ' (X * 20) - 18 'You might have to adjust this spacing.  I just made it up.
                .SpecialEffect = fmSpecialEffectFlat ' fmSpecialEffectEtched
                .Caption = "Label " & i
            End With
        Else
            ' Add Button control
            strClassName = "Forms.CommandButton.1"
            strName = "btn" & i
            Set tmpCtrl = Me.Controls.Add(strClassName, strName)
            With tmpCtrl
                .Left = Me.img_Sidebar.Width
                .Width = Me.Width - 20
                .Height = 18
                .Top = (i - 1) * 18 ' (X * 20) - 18 'You might have to adjust this spacing.  I just made it up.
                .Caption = "Button " & i
            End With
        
        End If
               
        ' Add to ContextMenuItem
        cmi.SetControl tmpCtrl
        cmi.Tag = strName
        cmi.UID = tmpCtrl.Name
        
        ' Add to ContextMenu
        f_cm.AddItem cmi
            
    Next

End Sub
