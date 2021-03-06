Attribute VB_Name = "WinForms"

Option Explicit
Option Compare Text

Private Const rootClass As String = "WinForms"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modFormControl
' By Chip Pearson, chip@cpearson.com, www.cpearson.com
' 21-March-2008
' URL: http://www.cpearson.com/Excel/FormControl.aspx
' Requires: modWindowCaption at http://www.cpearson.com/Excel/FileExtensions.aspx
'
' ----------------------------------
' Functions In This Module:
' ----------------------------------
'   SetFormParent
'       Sets a userform's parent to the Application or the ActiveWindow.
'   IsCloseButtonVisible
'       Returns True or False indicating whether the userform's Close button
'       is visible.
'   ShowCloseButton
'       Displays or hides the userform's Close button.
'   IsCloseButtonEnabled
'       Returns True or False indicating whether the userform's Close button
'       is enabled.
'   EnableCloseButton
'       Enables or disables a userform's Close button.
'   ShowTitleBar
'       Displays or hides a userform's Title Bar. The title bar cannot be
'       hidden if the form is resizable.
'   IsTitleBarVisible
'       Returns True or False indicating if the userform's Title Bar is visible.
'   MakeFormResizable
'       Makes the form resizable or not resizable. If the form is made resizable,
'       the title bar cannot be hidden.
'   IsFormResizable
'       Returns True or False indicating whether the userform is resizable.
'   SetFormOpacity
'       Sets the opacity of a form from fully opaque to fully invisible.
'   HasMaximizeButton
'       Returns True or False indicating whether the userform has a
'       maximize button.
'   HasMinimizeButton
'       Returns True or False indicating whether the userform has a
'       minimize button.
'   ShowMaximizeButton
'       Displays or hides a Maximize Window button on the userform.
'   ShowMinimizeButton
'       Displays or hides a Minimize Window button on the userform.
'   HWndOfUserForm
'       Returns the window handle (HWnd) of a userform.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''' Another reference:
''' https://www.techrepublic.com/blog/10-things/10-plus-of-my-favorite-windows-api-functions-to-use-in-office-applications/

' Windows Class names
Private Const C_USERFORM_CLASSNAME = "ThunderDFrame"
Private Const C_EXCEL_APP_CLASSNAME = "XLMain"
Private Const C_EXCEL_DESK_CLASSNAME = "XLDesk"
Private Const C_EXCEL_WINDOW_CLASSNAME = "Excel7"
Public Const C_OUTLOOK_EXPLORER_CLASSNAME = "rctrl_renwnd32"
Private Const C_ACCESS_APP_CLASSNAME = "OMain"
Private Const C_WORD_APP_CLASSNAME = "OpusApp"

Private Const MF_BYPOSITION = &H400
Private Const MF_REMOVE = &H1000
Private Const MF_ENABLED = &H0&
Private Const MF_DISABLED = &H2&
Private Const MF_GRAYED = &H1&
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)
Private Const GWL_HWNDPARENT = (-8)
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&
Private Const C_ALPHA_FULL_TRANSPARENT As Byte = 0
Private Const C_ALPHA_FULL_OPAQUE As Byte = 255
Private Const WS_DLGFRAME = &H400000
Private Const WS_THICKFRAME = &H40000
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000

''' ShowWindow - use one these constants
Private Const SW_FORCEMINIMIZE = 11     ' Minimizes the window
Private Const SW_HIDE = 0               ' Hides the window & activates another window
Private Const SW_MAXIMIZE = 3           ' Maximizes the window
Private Const SW_MINIMIZE = 6           ' Minimizes the window & activates next top-level window
Private Const SW_RESTORE = 9            ' Activates and displays the window
Private Const SW_SHOW = 5               ' Activates the window
Private Const SW_SHOWMAXIMIZED = 3      ' Activates the window & displays as maximized
Private Const SW_SHOWMINIMIZED = 2      ' Activates the window & displays as minimized
Private Const SW_SHOWMINNOACTIVE = 7    ' Displays the window minimized w/o activating it
Private Const SW_SHOWNA = 8             ' Displays the window in its current size & position w/o activating it
Private Const SW_SHOWNOACTIVATE = 4     ' Displays the window in its most recent size & position
Private Const SW_SHOWNORMAL = 1         ' Activates and displays the window

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5
Private Const WM_SETTEXT = &HC
Private Const WM_CHAR = &H102
Private Const WM_KEYDOWN = &H100

Private Const KEYEVENTF_KEYUP = &H2
Private Const INPUT_MOUSE = 0
Private Const INPUT_KEYBOARD = 1
Private Const INPUT_HARDWARE = 2

Public Enum FORM_PARENT_WINDOW_TYPE
    FORM_PARENT_NONE = 0
    FORM_PARENT_APPLICATION = 1
    FORM_PARENT_WINDOW = 2
End Enum

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type MOUSEINPUT
    dx As Long
    dy As Long
    mouseData As Long
    dwFlags As Long
    time As Long
    dwExtraInfo As Long
End Type

Private Type KEYBDINPUT
    wVK As Integer
    wScan As Integer
    dwFlags As Long
    time As Long
    dwExtraInfo As Long
End Type

Private Type HARDWAREINPUT
    uMsg As Long
    wParamL As Integer
    wParamH As Integer
End Type

Private Type GENERALINPUT
    dwType As Long
    xi(0 To 23) As Byte
End Type

Private Type INPUT_     '   typedef struct tagINPUT ' {
  dwType      As LongPtr
  wVK         As Integer
  wScan       As Integer        '               KEYBDINPUT ki;
  dwFlags     As LongPtr            '               HARDWAREINPUT hi;
  dwTime      As LongPtr            '           '};
  dwExtraInfo As LongPtr            '   '} INPUT, *PINPUT;
  dwPadding   As Currency        '   8 extra bytes, because mouses take more.
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

'''  Get mouse cursor position inside or outside a form anywhere on the screen.
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare PtrSafe Function SetParent Lib "user32" ( _
    ByVal hWndChild As Long, _
    ByVal hWndNewParent As Long) As Long
    
Private Declare PtrSafe Function GetParent Lib "user32" ( _
    ByVal hwnd As Long) As Long

''' Finds a Window on the Desktop
''' Both arguments are not necessary, use vbNullString, e.g. FindWindow(vbNullstring, frm.caption)
''' For class names see constants C_ above
''' The function returns the window's handle
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
    
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, _
    ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long
    
''' After 'finding' the window, you can manipulate its show state (e.g. minimize, maximize, etc)
''' use nCmdShow with declared constants SW_
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdSHow As Long) As Long
    
Private Declare PtrSafe Function MoveWindow Lib "user32.dll" ( _
    ByVal hwnd As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal bRepaint As Long) As Long
    
Private Declare PtrSafe Function GetWindow Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal wCmd As Long) As Long
   
Private Declare PtrSafe Function GetWindowRect Lib "user32.dll" _
                                (ByVal hwnd As Long, _
                                 ByRef lpRect As RECT) As Long
  
''' Retrieves the handle to the desktop window, which covers the entire screen.
''' All other windows are drawn on top of the desktop window.
''' This is one of the easier functions to implement, as there are no parameters.
''' You'll seldom use it alone; rather, you'll combine it with other API functions.
''' For instance, you might combine it with others so you can temporarily drop files
''' onto the desktop or enumerate through all the open windows on the desktop.
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long


Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" _
  (ByVal hwnd As Long, _
   lpdwProcessId As Long) As Long
    
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
    
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal crey As Byte, _
    ByVal bAlpha As Byte, _
    ByVal dwFlags As Long) As Long

''' GetActiveWindow retrieves the window handle for the currently active window
''' -- the new window you last clicked. If there is no active window associated
''' with the thread, the return value is NULL.
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As Long

Private Declare PtrSafe Function DrawMenuBar Lib "user32" ( _
    ByVal hwnd As Long) As Long

Private Declare PtrSafe Function GetMenuItemCount Lib "user32" ( _
    ByVal hMenu As Long) As Long

Private Declare PtrSafe Function GetSystemMenu Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal bRevert As Long) As Long
    
Private Declare PtrSafe Function RemoveMenu Lib "user32" ( _
    ByVal hMenu As Long, _
    ByVal nPosition As Long, _
    ByVal wFlags As Long) As Long
    
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
    ByVal hwnd As Long, _
    ByVal lpString As String, _
    ByVal cch As Long) As Long
    
Private Declare PtrSafe Function SetWindowText Lib "user32" Alias "SetWindowTextA" ( _
    ByVal hwnd As Long, ByVal _
    lpString As String) As Long
    
Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" ( _
    ByVal hwnd As Long) As Long

Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
    ByVal hwnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long

Private Declare PtrSafe Function EnableMenuItem Lib "user32" ( _
    ByVal hMenu As Long, _
    ByVal wIDEnableItem As Long, _
    ByVal wEnable As Long) As Long
    
''' This API function brings the specified window to the top.
''' If the window is a top-level window, the function activates it.
''' If the window is a child window, the function activates the
''' associated top-level parent window.
Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare PtrSafe Function SendInput Lib "user32.dll" ( _
    ByVal nInputs As Long, _
    pInputs As INPUT_, _
    ByVal cbSize As Long) As Long
    
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    ByVal ByteLen As Long)

Function ShowMaximizeButton(UF As MSForms.UserForm, _
    HideButton As Boolean) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShowMaximizeButton
' Displays (if HideButton is False) or hides (if HideButton is True)
' a maximize window button.
' NOTE: If EITHER a Minimize or Maximize button is displayed,
' BOTH buttons are visible but may be disabled.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim WinInfo As Long
Dim r As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    ShowMaximizeButton = False
    Exit Function
End If

WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
If HideButton = False Then
    WinInfo = WinInfo Or WS_MAXIMIZEBOX
Else
    WinInfo = WinInfo And (Not WS_MAXIMIZEBOX)
End If
r = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)

ShowMaximizeButton = (r <> 0)

End Function

Function ShowMinimizeButton(UF As MSForms.UserForm, _
    HideButton As Boolean) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShowMinimizeButton
' Displays (if HideButton is False) or hides (if HideButton is True)
' a minimize window button.
' NOTE: If EITHER a Minimize or Maximize button is displayed,
' BOTH buttons are visible but may be disabled.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim WinInfo As Long
Dim r As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    ShowMinimizeButton = False
    Exit Function
End If

WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
If HideButton = False Then
    WinInfo = WinInfo Or WS_MINIMIZEBOX
Else
    WinInfo = WinInfo And (Not WS_MINIMIZEBOX)
End If
r = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)

ShowMinimizeButton = (r <> 0)

End Function

Function HasMinimizeButton(UF As MSForms.UserForm) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' HasMinimizeButton
' Returns True if the userform has a minimize button, False
' otherwise.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim WinInfo As Long
Dim r As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    HasMinimizeButton = False
    Exit Function
End If

WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)

If WinInfo And WS_MINIMIZEBOX Then
    HasMinimizeButton = True
Else
    HasMinimizeButton = False
End If

End Function

Sub SendKey(bKey As Byte)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SendKey"
    
    On Error GoTo ThrowException

    Dim keyInput()    As INPUT_
    
    ReDim keyInput(0 To 1)

    keyInput(0).dwType = INPUT_KEYBOARD
    keyInput(0).dwFlags = 0        ' Press key
    keyInput(0).wVK = bKey

    keyInput(1).dwType = INPUT_KEYBOARD
    keyInput(1).dwFlags = KEYEVENTF_KEYUP
    keyInput(1).wVK = bKey
    
    Dim lReturn As Long
    lReturn = SendInput(2, keyInput(0), LenB(keyInput(0)))
    If lReturn > 0 Then
        ' Success
        strTrace = "# of keys sent: " & lReturn
        LogMessage strTrace, strRoutine
    Else
        ' Input was blocked
        strTrace = "SendKey failed to send " & CStr(bKey)
        err.Number = -99
        GoTo ThrowException
    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Sub SendKey2(bKey As Byte)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SendKey2"
    
    On Error GoTo ThrowException
    
    Dim GInput(0 To 1) As GENERALINPUT
    Dim KInput As KEYBDINPUT
    
    KInput.wVK = bKey  'the key we're going to press
    KInput.dwFlags = 0 'press the key
    'copy the structure into the input array's buffer.
    GInput(0).dwType = INPUT_KEYBOARD   ' keyboard input
    CopyMemory GInput(0).xi(0), KInput, LenB(KInput)
    
    'do the same as above, but for releasing the key
    KInput.wVK = bKey  ' the key we're going to release
    KInput.dwFlags = KEYEVENTF_KEYUP  ' release the key
    GInput(1).dwType = INPUT_KEYBOARD  ' keyboard input
    CopyMemory GInput(1).xi(0), KInput, LenB(KInput)
    
    'send the input now
    Dim lReturn As Long
    ' lReturn = SendInput(2, GInput(0), LenB(GInput(0)))
    If lReturn > 0 Then
        ' Success
        strTrace = "# of keys sent: " & lReturn
        LogMessage strTrace, strRoutine
    Else
        ' Input was blocked
        strTrace = "SendKey failed to send " & CStr(bKey)
        GoTo ThrowException
    End If
   
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Function SendKeyOld(ByVal UF As MSForms.UserForm, ByVal key As Long) As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SendKeyOld"
    
    On Error GoTo ThrowException

    Dim UFHWnd As Long

    UFHWnd = HWndOfUserForm(UF)
    If UFHWnd = 0 Then
        SendKeyOld = False
        Exit Function
    End If

    ' SendMessage UFHWnd, WM_SETTEXT, 0&, byval msg
    SendMessage UFHWnd, WM_KEYDOWN, key, ByVal 0
    
    strTrace = "Window: " & UFHWnd & " Key: " & key
    LogMessage strTrace, strRoutine
    
    SendKeyOld = True
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    SendKeyOld = False

End Function

Function HasMaximizeButton(UF As MSForms.UserForm) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' HasMaximizeButton
' Returns True if the userform has a maximize button, False
' otherwise.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim WinInfo As Long
Dim r As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    HasMaximizeButton = False
    Exit Function
End If

WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)

If WinInfo And WS_MAXIMIZEBOX Then
    HasMaximizeButton = True
Else
    HasMaximizeButton = False
End If

End Function

''' Get Absolute coordinates of the Mouse - also in WinContextMenu module
Public Sub GetCursorPositionA(ByRef mX As Long, ByRef mY As Long)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetCursorPosition"
    
    On Error GoTo ThrowException
    
    Dim a As POINTAPI
    
    Dim lReturn As Long
    lReturn = GetCursorPos(a)
    
    mX = a.X
    mY = a.Y
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    mX = -1

End Sub

''' Returns the Absolute position of the specified UserForm
Public Sub GetFormPosition(UF As UserForm, _
                            ByRef X As Long, ByRef Y As Long)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetFormPosition"
    
    On Error GoTo ThrowException
    
    If UF Is Nothing Then
        strTrace = "A null form encountered."
        GoTo ThrowException
    End If
    
    ' Inititalize returned values
    X = 0
    Y = 0
    
    Dim UFHWnd As Long

    ' Find the Userform's Windows Handle
    UFHWnd = HWndOfUserForm(UF)
    If UFHWnd = 0 Then
        strTrace = "Failed to get the Userform's Windows Handle"
        GoTo ThrowException
    End If
    
    Dim lr As RECT
    Call GetWindowRect(UFHWnd, lr)
    
    X = lr.Left
    Y = lr.Top
        
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    X = -1

End Sub

''' Sets the Position of any Windows Form based on its handle
Public Function SetWindowPosition(ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long) As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SetWindowPosition"
    
    On Error GoTo ThrowException
    
    If hwnd <= 0 Then
        strTrace = "Invalid windows handle: " & hwnd
        GoTo ThrowException
    End If
    
    ' Get the current position and size of the specified window
    Dim lr As RECT
    Call GetWindowRect(hwnd, lr)
    
    Dim w As Integer
    Dim h As Integer
    w = lr.Right - lr.Left
    h = lr.Bottom - lr.Top
    
    Dim b As Long
    b = 1
    
    ' Move the window to the new origin coordinates
    Dim bWorked As Long
    bWorked = MoveWindow(hwnd, X, Y, w, h, b)
    If Not bWorked Then
        strTrace = "Failed to move the window to (" & X & "," & Y & ") w=" & w & " h=" & h
        GoTo ThrowException
    End If
    
    SetWindowPosition = bWorked
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    SetWindowPosition = False

End Function

''' Sets the Position of a UserForm
Public Function SetFormPosition(ByVal UF As MSForms.UserForm, _
                                ByVal X As Long, ByVal Y As Long) As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SetFormPosition"
    
    On Error GoTo ThrowException
    
    If UF Is Nothing Then
        strTrace = "A null form encountered."
        GoTo ThrowException
    End If
    
    Dim UFHWnd As Long

    ' Find the Userform's Windows Handle
    UFHWnd = HWndOfUserForm(UF)
    If UFHWnd = 0 Then
        strTrace = "Failed to get the Userform's Windows Handle"
        GoTo ThrowException
    End If
    
    Dim bWorked As Boolean
    bWorked = SetWindowPosition(UFHWnd, X, Y)
    
    SetFormPosition = bWorked
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    SetFormPosition = False

End Function

Function GetFormParent(UF As MSForms.UserForm) As Long

    Dim UFHWnd As Long
    Dim WindHWnd As Long

    UFHWnd = HWndOfUserForm(UF)
    If UFHWnd = 0 Then
        GetFormParent = 0
        Exit Function
    End If
    
    WindHWnd = GetParent(UFHWnd)
    
    GetFormParent = WindHWnd

End Function


Function SetFormParent(UF As MSForms.UserForm, _
                       PHWnd As Long) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SetFormParent
' Set the UserForm UF as a child of (1) the Application, (2) the
' Excel ActiveWindow, or (3) no parent. Returns TRUE if successful
' or FALSE if unsuccessful.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim WindHWnd As Long
Dim r As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    SetFormParent = False
    Exit Function
End If
r = SetParent(UFHWnd, PHWnd)

GoTo SkipOut

Select Case Parent
    Case FORM_PARENT_APPLICATION
        Dim l As Long
        l = Application.hwnd
        r = SetParent(UFHWnd, Application.hwnd)
    Case FORM_PARENT_NONE
        r = SetParent(UFHWnd, 0&)
    Case FORM_PARENT_WINDOW
        If Application.ActiveWindow Is Nothing Then
            SetFormParent = False
            Exit Function
        End If
        WindHWnd = WindowHWnd(Application.ActiveWindow)
        If WindHWnd = 0 Then
            SetFormParent = False
            Exit Function
        End If
        r = SetParent(UFHWnd, WindHWnd)
    Case Else
        SetFormParent = False
        Exit Function
End Select

SkipOut:
    SetFormParent = (r <> 0)

End Function


Function IsCloseButtonVisible(UF As MSForms.UserForm) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsCloseButtonVisible
' Returns TRUE if UserForm UF has a close button, FALSE if there
' is no close button.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim WinInfo As Long
Dim r As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    IsCloseButtonVisible = False
    Exit Function
End If

WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
IsCloseButtonVisible = (WinInfo And WS_SYSMENU)

End Function

Function ShowCloseButton(UF As MSForms.UserForm, HideButton As Boolean) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShowCloseButton
' This displays (if HideButton is FALSE) or hides (if HideButton is
' TRUE) the Close button on the userform
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim WinInfo As Long
Dim r As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    Exit Function
End If

WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
If HideButton = False Then
    ' set the SysMenu bit
    WinInfo = WinInfo Or WS_SYSMENU
Else
    ' clear the SysMenu bit
    WinInfo = WinInfo And (Not WS_SYSMENU)
End If

r = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)
ShowCloseButton = (r <> 0)

End Function


Function IsCloseButtonEnabled(UF As MSForms.UserForm) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsCloseButtonEnabled
' This returns TRUE if the close button is enabled or FALSE if
' the close button is disabled.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim hMenu As Long
Dim ItemCount As Long
Dim PrevState As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    IsCloseButtonEnabled = False
    Exit Function
End If
' Get the menu handle
hMenu = GetSystemMenu(UFHWnd, 0&)
If hMenu = 0 Then
    IsCloseButtonEnabled = False
    Exit Function
End If

ItemCount = GetMenuItemCount(hMenu)
' Disable the button. This returns MF_DISABLED or MF_ENABLED indicating
' the previous state of the item.
PrevState = EnableMenuItem(hMenu, ItemCount - 1, MF_DISABLED Or MF_BYPOSITION)

If PrevState = MF_DISABLED Then
    IsCloseButtonEnabled = False
Else
    IsCloseButtonEnabled = True
End If
' restore the previous state
EnableCloseButton UF, (PrevState = MF_DISABLED)

DrawMenuBar UFHWnd

End Function


Function EnableCloseButton(UF As MSForms.UserForm, Disable As Boolean) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EnableCloseButton
' This function enables (if Disable is False) or disables (if
' Disable is True) the "X" button on a UserForm UF.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim hMenu As Long
Dim ItemCount As Long
Dim Res As Long

' Get the HWnd of the UserForm.
UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    EnableCloseButton = False
    Exit Function
End If
' Get the menu handle
hMenu = GetSystemMenu(UFHWnd, 0&)
If hMenu = 0 Then
    EnableCloseButton = False
    Exit Function
End If

ItemCount = GetMenuItemCount(hMenu)
If Disable = True Then
    Res = EnableMenuItem(hMenu, ItemCount - 1, MF_DISABLED Or MF_BYPOSITION)
Else
    Res = EnableMenuItem(hMenu, ItemCount - 1, MF_ENABLED Or MF_BYPOSITION)
End If
If Res = -1 Then
    EnableCloseButton = False
    Exit Function
End If
DrawMenuBar UFHWnd

EnableCloseButton = True


End Function

''' Hides the Forms Title Bar (where the Caption is located)
Sub HideTitleBar(frm As Object)
    Dim lngWindow As Long
    Dim lFrmHdl As Long
    lFrmHdl = FindWindow(vbNullString, frm.Caption)
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
    Call DrawMenuBar(lFrmHdl)
End Sub

''' Shows the Forms Title Bar (where the Caption is located)
Sub ShowTitleBar(frm As Object)
    Dim lngWindow As Long
    Dim lFrmHdl As Long
    lFrmHdl = FindWindow(vbNullString, frm.Caption)
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow Or WS_CAPTION
    Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
    Call DrawMenuBar(lFrmHdl)
End Sub

Function IsTitleBarVisible(UF As MSForms.UserForm) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsTitleBarVisible
' Returns TRUE if the title bar of UF is visible or FALSE if the
' title bar is not visible.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim UFHWnd As Long
Dim WinInfo As Long
Dim r As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    IsTitleBarVisible = False
    Exit Function
End If

WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)

IsTitleBarVisible = (WinInfo And WS_CAPTION)

End Function

Function MakeFormResizable(UF As MSForms.UserForm, Sizable As Boolean) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MakeFormResizable
' This makes the userform UF resizable (if Sizable is TRUE) or not
' resizable (if Sizalbe is FALSE). Returns TRUE if successful or FALSE
' if an error occurred.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim UFHWnd As Long
Dim WinInfo As Long
Dim r As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    MakeFormResizable = False
    Exit Function
End If

WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
If Sizable = True Then
    WinInfo = WinInfo Or WS_SIZEBOX
Else
    WinInfo = WinInfo And (Not WS_SIZEBOX)
End If

r = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)
MakeFormResizable = (r <> 0)


End Function

Function IsFormResizable(UF As MSForms.UserForm) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsFormResizable
' Returns TRUE if UF is resizable, FALSE if UF is not resizable.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim UFHWnd As Long
Dim WinInfo As Long
Dim r As Long

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    IsFormResizable = False
    Exit Function
End If

WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)

IsFormResizable = (WinInfo And WS_SIZEBOX)

End Function


Function SetFormOpacity(UF As MSForms.UserForm, Opacity As Byte) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SetFormOpacity
' This function sets the opacity of the UserForm referenced by the
' UF parameter. Opacity specifies the opacity of the form, from
' 0 = fully transparent (invisible) to 255 = fully opaque. The function
' returns True if successful or False if an error occurred. This
' requires Windows 2000 or later -- it will not work in Windows
' 95, 98, or ME.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim UFHWnd As Long
Dim WinL As Long
Dim Res As Long

SetFormOpacity = False

UFHWnd = HWndOfUserForm(UF)
If UFHWnd = 0 Then
    Exit Function
End If

WinL = GetWindowLong(UFHWnd, GWL_EXSTYLE)
If WinL = 0 Then
    Exit Function
End If

Res = SetWindowLong(UFHWnd, GWL_EXSTYLE, WinL Or WS_EX_LAYERED)
If Res = 0 Then
    Exit Function
End If

Res = SetLayeredWindowAttributes(UFHWnd, 0, Opacity, LWA_ALPHA)
If Res = 0 Then
    Exit Function
End If

SetFormOpacity = True

End Function

''' Returns the Windows handle to a user form based on its caption
Function GetUserFormHandle(ByVal Caption As String) As Long

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetUserFormHandle"
    
    On Error GoTo ThrowException

    Dim UFHWnd As Long
    UFHWnd = FindWindow(C_USERFORM_CLASSNAME, Caption)
    If UFHWnd = 0 Then
        strTrace = "Failed to find a user form using caption: " & Caption
        GoTo ThrowException
    End If
    
    GetUserFormHandle = UFHWnd
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    GetUserFormHandle = -1

End Function

Function FindExplorerWindow(ByVal Caption As String) As Long

    Dim handle As Long
    
    handle = FindWindow(C_OUTLOOK_EXPLORER_CLASSNAME, Caption)
'    If handle <> 0 Then
'        HWndOfUserForm = UFHWnd
'        Exit Function
'    End If
    
    FindExplorerWindow = handle

End Function

Function HWndOfUserForm(UF As MSForms.UserForm) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' HWndOfUserForm
' This returns the window handle (HWnd) of the userform referenced
' by UF. It first looks for a top-level window, then a child
' of the Application window, then a child of the ActiveWindow.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim AppHWnd As Long
Dim DeskHWnd As Long
Dim WinHWnd As Long
Dim UFHWnd As Long
Dim Cap As String
Dim WindowCap As String

Cap = UF.Caption

' First, look in top level windows
UFHWnd = FindWindow(C_USERFORM_CLASSNAME, Cap)
If UFHWnd <> 0 Then
    HWndOfUserForm = UFHWnd
    Exit Function
End If
' Not a top level window. Search for child of application.
'
'AppHWnd = Application.Hwnd
'UFHWnd = FindWindowEx(AppHWnd, 0&, C_USERFORM_CLASSNAME, Cap)
'If UFHWnd <> 0 Then
'    HWndOfUserForm = UFHWnd
'    Exit Function
'End If
' Not a child of the application.
' Search for child of ActiveWindow (Excel's ActiveWindow, not
' Window's ActiveWindow).
If Application.ActiveWindow Is Nothing Then
    HWndOfUserForm = 0
    Exit Function
End If
WinHWnd = WindowHWnd(Application.ActiveWindow)
UFHWnd = FindWindowEx(WinHWnd, 0&, C_USERFORM_CLASSNAME, Cap)
HWndOfUserForm = UFHWnd

End Function


Function ClearBit(value As Long, ByVal BitNumber As Long) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ClearBit
' Clears the specified bit in Value and returns the result. Bits are
' numbered, right (most significant) 31 to left (least significant) 0.
' BitNumber is made positive and then MOD 32 to get a valid bit number.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim SetMask As Long
Dim ClearMask As Long

BitNumber = Abs(BitNumber) Mod 32

SetMask = value
If BitNumber < 30 Then
    ClearMask = Not (2 ^ (BitNumber - 1))
    ClearBit = SetMask And ClearMask
Else
    ClearBit = value And &H7FFFFFFF
End If

End Function

