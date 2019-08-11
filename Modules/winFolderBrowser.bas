Attribute VB_Name = "winFolderBrowser"
Option Explicit
Option Private Module
Option Compare Text
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modBrowseFolderEx
' This contains the BrowseFolder function, which displays the standard Windows Browse For Folder
' dialog. It return the complete path of the selected folder or vbNullString if the user cancelled.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_DONTGOBELOWDOMAIN As Long = &H2
Private Const BIF_RETURNFSANCESTORS As Long = &H8
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000
Private Const BIF_BROWSEFORPRINTER As Long = &H2000
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000
Private Const BIF_EDITBOX As Long = &H10
Private Const BIF_VALIDATE As Long = &H20
Private Const BIF_NEWDIALOGSTYLE = &H40

Public Const WM_USER = &H400
Public Const BFFM_SETSTATUSTEXTA = WM_USER + 100
Public Const BFFM_ENABLEOK = WM_USER + 101
Public Const BFFM_SETSELECTIONA = WM_USER + 102
Public Const BFFM_SETSELECTIONW = WM_USER + 103
Public Const BFFM_SETSTATUSTEXTW = WM_USER + 104

Public Const LMEM_FIXED = &H0
Public Const LMEM_ZEROINIT = &H40
Public Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

'Used by callback function to communicate with the browser
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Any) As Long

Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        hpvDest As Any, _
        hpvSource As Any, _
        ByVal cbCopy As Long)

Public Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Public Declare PtrSafe Function LocalAlloc Lib "kernel32" _
       (ByVal uFlags As Long, _
        ByVal uBytes As Long) As Long
    
Public Declare PtrSafe Function LocalFree Lib "kernel32" _
   (ByVal hMem As Long) As Long

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszINSTRUCTIONS As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type


Private Declare PtrSafe Function SHGetPathFromIDListA Lib "shell32.dll" (ByVal pidl As Long, _
    ByVal pszBuffer As String) As Long

Private Declare PtrSafe Function SHBrowseForFolderA Lib "shell32.dll" (lpBrowseInfo As _
    BROWSEINFO) As Long


Private Const MAX_PATH = 260 ' Windows mandated


Function BrowseFolderEx(Optional ByVal DialogTitle As String) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BrowseFolder
' This displays the standard Windows Browse Folder dialog. It returns
' the complete path name of the selected folder or vbNullString if the
' user cancelled.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If DialogTitle = vbNullString Then
    DialogTitle = "Select A Folder"
End If

Dim uBrowseInfo As BROWSEINFO
Dim szBuffer As String
Dim lID As Long
Dim lRet As Long


With uBrowseInfo
    .hOwner = 0
    .pidlRoot = 0
    .pszDisplayName = String$(MAX_PATH, vbNullChar)
    .lpszINSTRUCTIONS = DialogTitle
    .ulFlags = BIF_RETURNONLYFSDIRS ' + BIF_USENEWUI
    .lpfn = 0
End With
szBuffer = String$(MAX_PATH, vbNullChar)
lID = SHBrowseForFolderA(uBrowseInfo)

If lID Then
    ''' Retrieve the path string.
    lRet = SHGetPathFromIDListA(lID, szBuffer)
    If lRet Then
        BrowseFolderEx = Left$(szBuffer, InStr(szBuffer, vbNullChar) - 1)
    End If
End If

End Function


Sub AAA()
Dim FolderName As String
FolderName = BrowseFolderEx("")
If FolderName = vbNullString Then
    MsgBox "No folder selected"
Else
    MsgBox "Folder: " & FolderName
End If

End Sub


