Attribute VB_Name = "WinCaption"
Option Explicit
Option Compare Text
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modWindowCaption
' By Chip Pearson, 15-March-2008, chip@cpearson.com, www.cpearson.com
' http://www.cpearson.com/Excel/FileExtensions.aspx
'
' This module contains code for working with Excel.Window captions. This code
' is necessary if you are going to use the FindWindowEx API call to get the
' HWnd of an Excel.Window.  Windows has a property called "Hide extensions of
' known file types". If this setting is TRUE, the file extension is not displayed
' (e.g., "Book1.xls" is displayed as just "Book1"). However, the Caption of an
' Excel.Window always includes the ".xls" file extension, regardless of the hide
' extensions setting. FindWindowEx requires that the ".xls" extension be removed
' if the "hide extensions" setting is True.
'
' This module contains a function named DoesWindowsHideFileExtensions, which returns
' TRUE if Windows is hiding file extensions or FALSE if Windows is not hiding file
' extensions. This is determined by a registry key. The module also contains a
' function named WindowCaption that returns the Caption of a specified Excel.Window
' with the ".xls" removed if necessary. The string returned by this function
' is suitable for use in FindWindowEx regardless of the value of the Windows
' "Hide Extensions" setting.
'
' This module also contains a function named WindowHWnd which returns the HWnd of
' a specified Excel.Window object. This function works regardless of the value of the
' Windows "Hide Extensions" setting.
'
' This module also includes the functions WindowText and WindowClassName which are
' just wrappers for the GetWindowText and GetClassName API functions.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, _
    ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long

Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
    ByVal hWnd As Long, _
    ByVal lpString As String, _
    ByVal cch As Long) As Long

Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
    ByVal HKey As Long, _
    ByVal lpSubKey As String, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    phkResult As Long) As Long
    
Private Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
    ByVal HKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    LPType As Long, _
    LPData As Any, _
    lpcbData As Long) As Long

Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" ( _
    ByVal HKey As Long) As Long

Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
    ByVal hWnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long
    
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
 

Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const HKEY_CLASSES_ROOT  As Long = &H80000000
Private Const HKEY_CURRENT_CONFIG  As Long = &H80000005
Private Const HKEY_DYN_DATA  As Long = &H80000006
Private Const HKEY_PERFORMANCE_DATA  As Long = &H80000004
Private Const HKEY_USERS  As Long = &H80000003
Private Const KEY_ALL_ACCESS  As Long = &H3F
Private Const ERROR_SUCCESS  As Long = 0&
Private Const HKCU  As Long = HKEY_CURRENT_USER
Private Const HKLM  As Long = HKEY_LOCAL_MACHINE

Private Enum REG_DATA_TYPE
    REG_DATA_TYPE_DEFAULT = 0   ' Default based on data type of value.
    REG_INVALID = -1            ' Invalid
    REG_SZ = 1                  ' String
    REG_DWORD = 4               ' Long
End Enum

Private Const C_EXCEL_APP_CLASSNAME = "XLMain"
Private Const C_EXCEL_DESK_CLASSNAME = "XLDesk"
Private Const C_EXCEL_WINDOW_CLASSNAME = "EXCEL7"


Function DoesWindowsHideFileExtensions() As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DoesWindowsHideFileExtensions
' This function looks in the registry key
'   HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced
' for the value named "HideFileExt" to determine whether the Windows Explorer
' setting "Hide Extensions Of Known File Types" is enabled. This function returns
' TRUE if this setting is in effect (meaning that Windows displays "Book1" rather
' than "Book1.xls"), or FALSE if this setting is not in effect (meaning that Windows
' displays "Book1.xls").
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Res As Long
Dim RegKey As Long
Dim v As Long

Const KEY_NAME = "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Const VALUE_NAME = "HideFileExt"

''''''''''''''''''''''''''''''''''''''''''''''''''
' Open the registry key to get a handle (RegKey).
''''''''''''''''''''''''''''''''''''''''''''''''''
Res = RegOpenKeyEx(HKey:=HKCU, _
                    lpSubKey:=KEY_NAME, _
                    ulOptions:=0&, _
                    samDesired:=KEY_ALL_ACCESS, _
                    phkResult:=RegKey)

If Res <> ERROR_SUCCESS Then
    Exit Function
End If
''''''''''''''''''''''''''''''''''''''''''''''''''
' Get the value of the "HideFileExt" named value.
''''''''''''''''''''''''''''''''''''''''''''''''''
Res = RegQueryValueEx(HKey:=RegKey, _
                    lpValueName:=VALUE_NAME, _
                    lpReserved:=0&, _
                    LPType:=REG_DWORD, _
                    LPData:=v, _
                    lpcbData:=Len(v))

If Res <> ERROR_SUCCESS Then
    RegCloseKey RegKey
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''''''
' Close the key and return the result.
''''''''''''''''''''''''''''''''''''''''''''''''''
RegCloseKey RegKey
DoesWindowsHideFileExtensions = (v <> 0)


End Function

Function WindowHWnd(ByVal windowCaption As String) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WindowHWnd
' This returns the HWnd of the Window with the specified Caption (Title)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim AppHWnd As Long
Dim DeskHWnd As Long
Dim WHWnd As Long
Dim Cap As String
Cap = windowCaption

AppHWnd = Application.hWnd
DeskHWnd = FindWindowEx(AppHWnd, 0&, C_OUTLOOK_EXPLORER_CLASSNAME, vbNullString)
If DeskHWnd > 0 Then
    WHWnd = FindWindowEx(DeskHWnd, 0&, C_EXCEL_WINDOW_CLASSNAME, Cap)
End If
WindowHWnd = WHWnd

End Function

Function WindowText(hWnd As Long) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WindowText
' This just wraps up GetWindowText.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim s As String
    Dim n As Long
    n = 255
    s = String$(n, vbNullChar)
    n = GetWindowText(hWnd, s, n)
    If n > 0 Then
        WindowText = Left(s, n)
    Else
        WindowText = vbNullString
    End If
End Function

Function WindowClassName(hWnd As Long) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WindowClassName
' This just wraps up GetClassName.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim s As String
    Dim n As Long
    n = 255
    s = String$(n, vbNullChar)
    n = GetClassName(hWnd, s, n)
    If n > 0 Then
        WindowClassName = Left(s, n)
    Else
        WindowClassName = vbNullString
    End If

End Function


