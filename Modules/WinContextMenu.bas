Attribute VB_Name = "WinContextMenu"
Option Explicit

Private Const rootClass As String = "WindowLib"

' Type required by TrackPopupMenu although this is ignored !!
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

' Type required by InsertMenuItem
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

' Type required by GetCursorPos
Private Type POINTAPI
        x As Long
        y As Long
End Type

' Constants required by TrackPopupMenu
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_TOPALIGN = &H0
Private Const TPM_RETURNCMD = &H100
Private Const TPM_RIGHTBUTTON = &H2&

' Constants required by MENUITEMINFO type
Private Const MIIM_STATE = &H1
Private Const MIIM_ID = &H2
Private Const MIIM_TYPE = &H10
Private Const MFT_STRING = &H0
Private Const MFT_SEPARATOR = &H800
Private Const MFS_DEFAULT = &H1000
Private Const MFS_ENABLED = &H0
Private Const MFS_GRAYED = &H1

' API declarations
Private Declare PtrSafe Function CreatePopupMenu Lib "user32" () As Long
Private Declare PtrSafe Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare PtrSafe Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long
Private Declare PtrSafe Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Integer, ByVal y As Integer) As Long

' Constants for Keys
Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2

Public Const VK_Tab = &H9

Public Const VK_F15 = 126

Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

''' CONTEXT MENU CALLS

' Contants defined by me for menu item IDs
Private Const ID_Cut = 101
Private Const ID_Copy = 102
Private Const ID_Paste = 103
Private Const ID_Delete = 104
Private Const ID_SelectAll = 105


' Variables declared at module level
Private formCaption As String
Private Cut_Enabled As Long
Private Copy_Enabled As Long
Private Paste_Enabled As Long
Private Delete_Enabled As Long
Private SelectAll_Enabled As Long

Public Function IsKeyPressed(ByVal key As Long) As Boolean

    Dim iReturn As Integer
    iReturn = GetAsyncKeyState(key)
    If iReturn Then
        IsKeyPressed = True
    Else
        IsKeyPressed = False
    End If

End Function

''' Moves the cursor or mouse pointer to the specified absolute coordinates
Public Function SetCursorPosition(ByVal x As Integer, ByVal y As Integer) As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetCursorPosition"
    
    On Error GoTo ThrowException
    
    If x < 0 Or y < 0 Then
        strTrace = "Invalid coordinates specified."
        GoTo ThrowException
    End If
    
    Dim l As Long
    l = SetCursorPos(x, y)
    If l > 0 Then
        strTrace = "Cursor moved to coordinates (" & x & "," & y & ")."
        LogMessage strTrace, strRoutine
    Else
        strTrace = "Failed to set the cursor to a new position using the Windows API."
        GoTo ThrowException
    End If

    SetCursorPosition = True
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    SetCursorPosition = False

End Function

''' Returns the X and Y coordinates of the absolute position of the cursor or
''' mouse position
Public Function GetCursorPosition(ByRef x As Integer, ByRef y As Integer) As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetCursorPosition"
    
    On Error GoTo ThrowException
    
    Dim pt As POINTAPI
    Dim l As Long
    
    l = GetCursorPos(pt)
    If l > 0 Then
        x = pt.x
        y = pt.y
    Else
        strTrace = "An error occurred while using the Windows API."
        GoTo ThrowException
    End If
    
    GetCursorPosition = True
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    GetCursorPosition = False

End Function


Public Function GetUserForm(ByVal ctl As MSForms.control) As UserForm

    Dim tmp As Object
    Set tmp = ctl.Parent
    Do While TypeOf tmp Is MSForms.control
        Set tmp = tmp.Parent
    Loop
' MsgBox tmp.name

    formCaption = tmp.Caption
    Set GetUserForm = tmp

End Function

Public Function ShowPopup(ctl As MSForms.control, cMenu As ContextMenu, x As Single, y As Single) As Long

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ShowPopup"

    ' Dim oControl As Object  'MSForms.TextBox
    ' Static click_flag As Long
    
    ' The following is required because the MouseDown event
    ' fires twice when right-clicked !!
    ' click_flag = click_flag + 1
        
    ' Do nothing on first firing of MouseDown event
    ' If (click_flag Mod 2 <> 0) Then Exit Function
                
    ' Set object reference to the textboxthat was clicked
    ' Set oControl = ctl 'oForm.ActiveControl
        
    ' If click is outside the calling control, do nothing
    If x > ctl.Width Or y > ctl.Height Or x < 0 Or y < 0 Then
        strTrace = "Outside control limits - ignoring menu call."
        GoTo ThrowException
    End If
    
    ' Retrieve caption of UserForm for use in FindWindow API
    '   - needs to set the formCaption variable
    Dim myForm As UserForm
    Set myForm = GetUserForm(ctl)
    ' formCaption = myForm.Caption  'oForm.Caption  'strCaption
    
    ' Call routine that sets menu items as enabled/disabled
    ' Call EnableMenuItems(myForm)
    
    ' Call function that shows the menu and return the ID
    ' of the selected menu item. Subsequent action depends
    ' on the returned ID.
    Dim lSelection As Long
    lSelection = GetSelection(cMenu)
    
'    Select Case GetSelection()
'        Case ID_Cut
'            oControl.Cut
'        Case ID_Copy
'            oControl.Copy
'        Case ID_Paste
'            oControl.Paste
'        Case ID_Delete
'            oControl.SelText = ""
'        Case ID_SelectAll
'            With oControl
'                .SelStart = 0
'                .SelLength = Len(oControl.Text)
'            End With
'    End Select
    
    ShowPopup = lSelection
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Function

Private Sub EnableMenuItems(oForm As UserForm)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":EnableMenuItems"

    Dim oControl As Object ' MSForms.TextBox
    Dim oData As DataObject
    Dim testClipBoard As String
    
    On Error Resume Next
    
    ' Set object variable to clicked textbox
    Set oControl = oForm.ActiveControl
    
    ' Create DataObject to access the clipboard
    Set oData = New DataObject
    
    ' Enable Cut/Copy/Delete menu items if text selected
    ' in textbox
    If oControl.SelLength > 0 Then
        Cut_Enabled = MFS_ENABLED
        Copy_Enabled = MFS_ENABLED
        Delete_Enabled = MFS_ENABLED
    Else
        Cut_Enabled = MFS_GRAYED
        Copy_Enabled = MFS_GRAYED
        Delete_Enabled = MFS_GRAYED
    End If
    
    ' Enable SelectAll menu item if there is any text in textbox
    If Len(oControl.text) > 0 Then
        SelectAll_Enabled = MFS_ENABLED
    Else
        SelectAll_Enabled = MFS_GRAYED
    End If
    
    ' Get data from clipbaord
    oData.GetFromClipboard
    
    ' Following line generates an error if there
    ' is no text in clipboard
    testClipBoard = oData.GetText

    ' If NO error (ie there is text in clipboard) then
    ' enable Paste menu item. Otherwise, diable it.
    If err.Number = 0 Then
        Paste_Enabled = MFS_ENABLED
    Else
        Paste_Enabled = MFS_GRAYED
    End If
    
    ' Clear the error object
    err.Clear
    
    ' Clean up object references
    Set oControl = Nothing
    Set oData = Nothing
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Function GetSelection(ByVal cMenu As ContextMenu) As Long

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetSelection"
    
    On Error GoTo ThrowException

    Dim menu_hwnd As Long
    Dim form_hwnd As Long
    Dim oRect As RECT
    Dim oPointAPI As POINTAPI
    
    ' Find hwnd of UserForm - note different classname
    ' Word 97 vs Word2000
    #If VBA6 Then
        form_hwnd = FindWindow("ThunderDFrame", formCaption)
    #Else
        form_hwnd = FindWindow("ThunderXFrame", formCaption)
    #End If

    ' Get current cursor position
    ' Menu will be drawn at this location
    GetCursorPos oPointAPI
        
    ' Create new popup menu
    menu_hwnd = CreatePopupMenu
    
    ' Process Context Menu
    Dim i As Long
    i = 1
    Dim mni As ContextMenuItem
    For Each mni In cMenu.Items
        
        Dim mInfo As MENUITEMINFO
        mInfo = ConvertMenuItem(mni)
        InsertMenuItem menu_hwnd, i, True, mInfo
        i = i + 1
        
    Next
    
    ' Return the ID of the item selected by the user
    ' and set it the return value of the function
    GetSelection = TrackPopupMenu _
                    (menu_hwnd, _
                     TPM_LEFTALIGN Or TPM_TOPALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, _
                     oPointAPI.x, oPointAPI.y, _
                     0, form_hwnd, oRect)
        
    ' Destroy the menu
    DestroyMenu menu_hwnd
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Function

Private Function ConvertMenuItem(ByVal mnuItem As ContextMenuItem) As MENUITEMINFO

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ConvertMenuItem"
    
    Dim oMenuItemInfo As MENUITEMINFO
    With oMenuItemInfo
        If mnuItem.itemType = Button Then
            .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
            .fType = MFT_STRING
            .dwTypeData = mnuItem.Caption
            .cch = LenB(.dwTypeData)
            .wID = mnuItem.UID
            If mnuItem.Enabled Then
                .fState = MFS_ENABLED
            Else
                .fState = MFS_GRAYED
            End If
        End If
        If mnuItem.itemType = Separator Then
            .fMask = MIIM_TYPE
            .fType = MFT_SEPARATOR
        End If
        
        .cbSize = LenB(oMenuItemInfo)
        
    End With
    
    ConvertMenuItem = oMenuItemInfo

End Function

Private Function GetSelectionOld(ByVal cMenu As ContextMenu) As Long

    Dim menu_hwnd As Long
    Dim form_hwnd As Long
    Dim oMenuItemInfo1 As MENUITEMINFO
    Dim oMenuItemInfo2 As MENUITEMINFO
    Dim oMenuItemInfo3 As MENUITEMINFO
    Dim oMenuItemInfo4 As MENUITEMINFO
    Dim oMenuItemInfo5 As MENUITEMINFO
    Dim oMenuItemInfo6 As MENUITEMINFO
    Dim oRect As RECT
    Dim oPointAPI As POINTAPI
    
    ' Find hwnd of UserForm - note different classname
    ' Word 97 vs Word2000
    #If VBA6 Then
        form_hwnd = FindWindow("ThunderDFrame", formCaption)
    #Else
        form_hwnd = FindWindow("ThunderXFrame", formCaption)
    #End If

    ' Get current cursor position
    ' Menu will be drawn at this location
    GetCursorPos oPointAPI
        
    ' Create new popup menu
    menu_hwnd = CreatePopupMenu
    
    ' Intitialize MenuItemInfo structures for the 6
    ' menu items to be added
    
    ' Cut
    With oMenuItemInfo1
            .cbSize = LenB(oMenuItemInfo1)
            .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
            .fType = MFT_STRING
            .fState = Cut_Enabled
            .wID = ID_Cut
            .dwTypeData = "Cut"
            .cch = LenB(.dwTypeData)
    End With
    
    ' Copy
    With oMenuItemInfo2
            .cbSize = LenB(oMenuItemInfo2)
            .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
            .fType = MFT_STRING
            .fState = Copy_Enabled
            .wID = ID_Copy
            .dwTypeData = "Copy"
            .cch = LenB(.dwTypeData)
    End With
    
    ' Paste
    With oMenuItemInfo3
            .cbSize = LenB(oMenuItemInfo3)
            .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
            .fType = MFT_STRING
            .fState = Paste_Enabled
            .wID = ID_Paste
            .dwTypeData = "Paste"
            .cch = LenB(.dwTypeData)
    End With
    
    ' Separator
    With oMenuItemInfo4
            .cbSize = LenB(oMenuItemInfo4)
            .fMask = MIIM_TYPE
            .fType = MFT_SEPARATOR
    End With
    
    ' Delete
    With oMenuItemInfo5
            .cbSize = LenB(oMenuItemInfo5)
            .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
            .fType = MFT_STRING
            .fState = Delete_Enabled
            .wID = ID_Delete
            .dwTypeData = "Delete"
            .cch = LenB(.dwTypeData)
    End With
    
    ' SelectAll
    With oMenuItemInfo6
            .cbSize = LenB(oMenuItemInfo6)
            .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
            .fType = MFT_STRING
            .fState = SelectAll_Enabled
            .wID = ID_SelectAll
            .dwTypeData = "Select All"
            .cch = LenB(.dwTypeData)
    End With
    
    ' Add the 6 menu items
'    InsertMenuItem menu_hwnd, 1, True, oMenuItemInfo1
'    InsertMenuItem menu_hwnd, 2, True, oMenuItemInfo2
'    InsertMenuItem menu_hwnd, 3, True, oMenuItemInfo3
'    InsertMenuItem menu_hwnd, 4, True, oMenuItemInfo4
'    InsertMenuItem menu_hwnd, 5, True, oMenuItemInfo5
'    InsertMenuItem menu_hwnd, 6, True, oMenuItemInfo6
    
    ' Process Context Menu
    Dim i As Long
    i = 1
    Dim mni As ContextMenuItem
    For Each mni In cMenu.Items
        
        Dim mInfo As MENUITEMINFO
        mInfo = ConvertMenuItem(mni)
        InsertMenuItem menu_hwnd, i, True, mInfo
        i = i + 1
        
    Next
    
    ' Return the ID of the item selected by the user
    ' and set it the return value of the function
    GetSelectionOld = TrackPopupMenu _
                    (menu_hwnd, _
                     TPM_LEFTALIGN Or TPM_TOPALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, _
                     oPointAPI.x, oPointAPI.y, _
                     0, form_hwnd, oRect)
        
    ' Destroy the menu
    DestroyMenu menu_hwnd

End Function




