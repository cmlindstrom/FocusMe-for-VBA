Attribute VB_Name = "WinFolder"
Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" _
                Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare PtrSafe Function GetSaveFileName Lib "comdlg32.dll" _
                Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare PtrSafe Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

Private Const rootClass As String = "WinFolder"

#If Win64 Then
    ' 64 bit
#Else
    ' 32 bit
#End If
             
Private Type OPENFILENAME
    lStructSize As LongPtr
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As LongPtr
    nFilterIndex As LongPtr
    lpstrFile As String
    nMaxFile As LongPtr
    lpstrFileTitle As String
    nMaxFileTitle As LongPtr
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As LongPtr
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
End Type

Public Function OpenFileDialog() As String
    OpenFileDialog = SelectFileDialog
End Function

Public Function SaveAsFileDialog() As String
    SaveAsFileDialog = SelectFolderDialog
End Function

Private Function SelectFolderDialog(Optional ByVal WindowTitle As String = "", _
                                   Optional ByVal Filter As String = "")
                                   
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SelectFolderDialog"
    
    On Error GoTo ThrowException
                                   
    If Len(Filter) = 0 Then
        Filter = "Text Files (*.txt)" & Chr$(0) & "*.txt" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*" & Chr$(0)
    End If
    If Len(WindowTitle) = 0 Then
        WindowTitle = "Save As - " & Commands.AppName
    End If
    
    Dim OFName As OPENFILENAME
       
    OFName.lStructSize = LenB(OFName)
    'Set the parent window
    Dim hwnd As Long
    hwnd = FindExplorerWindow(ThisOutlookSession.ActiveExplorer.Caption)
    OFName.hwndOwner = 0& ' hwnd ' Application.hwnd
    'Set the application's instance
    ' OFName.hInstance = hwnd 'Application.hInstance
    'Select a filter
    OFName.lpstrFilter = Filter ' + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    'create a buffer for the file
    OFName.lpstrFile = String(257, 0) ' Space$(254)
    'set the maximum length of a returned file
    OFName.nMaxFile = LenB(OFName.lpstrFile) - 1 '255
    'Create a buffer for the file title
    OFName.lpstrFileTitle = OFName.lpstrFile ' Space$(254)
    'Set the maximum length of a returned file title
    OFName.nMaxFileTitle = OFName.nMaxFile
    'Set the initial directory
    OFName.lpstrInitialDir = "C:\"
    'Set the title
    OFName.lpstrTitle = WindowTitle
    'No flags
    OFName.Flags = 0
    
    OFName.nFileOffset = 0
    OFName.nFileExtension = 0
    OFName.lpstrDefExt = "txt" & Chr$(0)
    OFName.lCustData = 0
    OFName.lpfnHook = 0
    OFName.lpTemplateName = 0
    
    'Show the 'Save As File'-dialog
    Dim APIResults As Long
    APIResults = GetSaveFileName(OFName)
    
    Dim strReturn As String
    strReturn = ""
       
    Dim lErr As Long
    If APIResults <> 0 Then
        ' MsgBox "File to Save: " + Trim$(OFName.lpstrFile)
        strReturn = Trim(OFName.lpstrFile)
    Else
        ' MsgBox "Cancel was pressed OR Error"
        lErr = CommDlgExtendedError()
        If lErr <> 0 Then
            strTrace = GetError(lErr)
            GoTo ThrowException
        End If
        strReturn = vbNullString
    End If
    
    SelectFolderDialog = strReturn
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    SelectFolderDialog = vbNullString

End Function

Private Function SelectFileDialog(Optional ByVal WindowTitle As String = "", _
                               Optional ByVal Filter As String = "") As String
                               
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SelectFileDialog"
    
    On Error GoTo ThrowException
    
    If Len(Filter) = 0 Then
        Filter = "Text Files (*.txt)" & Chr$(0) & "*.txt" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*" & Chr$(0)
    End If
    If Len(WindowTitle) = 0 Then
        WindowTitle = "Open File - " & Commands.AppName
    End If
    
    Dim OFName As OPENFILENAME
       
    OFName.lStructSize = LenB(OFName)
    'Set the parent window
    Dim hwnd As Long
    hwnd = FindExplorerWindow(ThisOutlookSession.ActiveExplorer.Caption)
    OFName.hwndOwner = 0& ' hwnd ' Application.hwnd
    'Set the application's instance
    ' OFName.hInstance = hwnd 'Application.hInstance
    'Select a filter
    OFName.lpstrFilter = Filter ' + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    'create a buffer for the file
    OFName.lpstrFile = String(257, 0) ' Space$(254)
    'set the maximum length of a returned file
    OFName.nMaxFile = LenB(OFName.lpstrFile) - 1 '255
    'Create a buffer for the file title
    OFName.lpstrFileTitle = OFName.lpstrFile ' Space$(254)
    'Set the maximum length of a returned file title
    OFName.nMaxFileTitle = OFName.nMaxFile
    'Set the initial directory
    OFName.lpstrInitialDir = "C:\"
    'Set the title
    OFName.lpstrTitle = WindowTitle
    'No flags
    OFName.Flags = 0
    
    OFName.nFileOffset = 0
    OFName.nFileExtension = 0
    OFName.lpstrDefExt = "txt" & Chr$(0)
    OFName.lCustData = 0
    OFName.lpfnHook = 0
    OFName.lpTemplateName = 0
    
    'Show the 'Open File'-dialog
    Dim APIResults As Long
    APIResults = GetOpenFileName(OFName)
    
    If APIResults <> 0 Then
        ' MsgBox "File to Open: " + Trim$(OFName.lpstrFile)
        SelectFileDialog = Trim$(OFName.lpstrFile)
    Else
        ' MsgBox "Cancel was pressed"
        SelectFileDialog = vbNullString
    End If
    
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    SelectFileDialog = vbNullString
    
End Function

Private Function GetError(ByVal errCode As Long) As String

    ' Error codes per
    ' https://docs.microsoft.com/en-us/windows/desktop/api/commdlg/nf-commdlg-commdlgextendederror

    Dim strErr As String

    Select Case errCode
        Case &H1
            strErr = "The lStructSize member of the initialization structure for the corresponding common dialog box is invalid."
        Case &HFFFF
            strErr = "The dialog box could not be created. The common dialog box function's call to the DialogBox function failed. For example, this error occurs if the common dialog box call specifies an invalid window handle."
        Case &H6
            strErr = "The common dialog box function failed to find a specified resource."
        Case &H2
            strErr = "The common dialog box function failed during initialization. This error often occurs when sufficient memory is not available."
        Case &H7
            strErr = "The common dialog box function failed to load a specified resource. "
        Case &H5
            strErr = "The common dialog box function failed to load a specified string. "
        Case &H8
            strErr = "The common dialog box function failed to lock a specified resource. "
        Case &H9
            strErr = "The common dialog box function was unable to allocate memory for internal structures."
        Case &HA
            strErr = "The common dialog box function was unable to lock the memory associated with a handle."
        Case &H4
            strErr = "The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding instance handle."
        Case &HB
            strErr = "The ENABLEHOOK flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a pointer to a corresponding hook procedure."
        Case &H3
            strErr = "The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding template. "
        Case &HC
            strErr = "The RegisterWindowMessage function returned an error code when it was called by the common dialog box function."
       
        Case Else
            strErr = "Unknown error."
    End Select
    
    GetError = strErr

End Function
