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

Private Const BackgroundNormal As Long = &H8000000F
Private Const BackgroundHighlight As Long = &H8000000D

Public Enum enuControlType
    None = 0
    Labels = 1
    buttons = 2
End Enum

' Form
Dim frmX As Integer
Dim frmY As Integer

' - Events

Public Event Click(ByVal tag As String)

' - Properties

Dim WithEvents f_cm As ContextMenu
Attribute f_cm.VB_VarHelpID = -1
Dim f_miType As enuControlType

Public Property Set Menu(ByVal cm As ContextMenu)
    Set f_cm = cm
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

' - Event Handlers

''' ContextMenu Events

Private Sub f_cm_Click(ByVal tag As String)
    MsgBox "Control clicked: " & tag
    RaiseEvent Click(tag)
    Me.Hide
End Sub

Private Sub f_cm_DblClick(ByVal tag As String, ByVal Cancel As MSForms.ReturnBoolean)
    Dim c As control
    Set c = FindControl(tag)
    MsgBox "Control double clicked: " & tag
End Sub

Private Sub f_cm_MouseMoved(ByVal tag As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim c As control
    Set c = FindControl(tag)
    Highlight c
End Sub

''' Form Events

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
End Sub

Private Sub Initialize()
    f_miType = Labels
End Sub

' - Methods



' - Supporting Methods

Private Sub BuildMenu()

End Sub

Private Sub BuildDefaultMenu()

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
                .Left = 10
                .Width = Me.Width ' - 20
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
                .Left = 10
                .Width = Me.Width - 20
                .Height = 18
                .Top = (i - 1) * 18 ' (X * 20) - 18 'You might have to adjust this spacing.  I just made it up.
                .Caption = "Button " & i
            End With
        
        End If
               
        ' Add ContextMenuItem
        cmi.SetControl tmpCtrl
        cmi.tag = strName
        
        ' Add to ContextMenu
        f_cm.AddItem cmi
            
    Next

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
            If lbl.Caption = c.Caption Then
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


