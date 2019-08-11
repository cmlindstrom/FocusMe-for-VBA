VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ProjectProperties 
   Caption         =   "Project Properties"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5310
   OleObjectBlob   =   "frm_ProjectProperties.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_ProjectProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' - Fields

Private Const rootClass As String = "frm_ProjectProperties"

Private f_Project As fmeProject

Dim bDirty As Boolean
Dim bDirtyMCL As Boolean

Private tm As TaskManager

' - Events

Public Event Closing(ByVal sender As Object, ByVal saved As Boolean)

' - Properties

''' The project that is being edited
Public Property Get Project() As fmeProject
    Set Project = f_Project
End Property

' - Event Handlers

Private Sub mp_Tabs_Change()
    Dim myTab As Page
    Set myTab = mp_Tabs.SelectedItem
   
    If InStr(1, LCase(myTab.Caption), "plan") > 0 Then
        tm.LoadByProject f_Project.Name, chkbx_IncludeCompletedTasks.value
    End If
   
End Sub

' - - Title Management

Private Sub chkbx_CombineTitleCode_Change()
    bDirty = True
    Call UpdateMCL
End Sub

Private Sub imgLogo_Click()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":imgLogo_Click"
    
    On Error GoTo ThrowException

    Dim X As Integer
    Dim Y As Integer

    strTrace = "Show Color Picker dialog."
    Dim frm_Color As New frm_ColorPicker
    
    strTrace = "Set up start location."
    Call TryGetRelativePosition(Me.imgLogo, X, Y)
    frm_Color.Left = X
    frm_Color.Top = Y
    
    frm_Color.Show
    
    strTrace = "Check Result."
    If Not IsNothing(frm_Color.SelectedPicture) Then
        
        strTrace = "Change the Logo."
        ' Set the Logo
        f_Project.Color = frm_Color.Selection
        imgLogo.Picture = frm_Color.SelectedPicture
        ' Repaint the graphic elements
        Me.Repaint
        ' Capture the change
        bDirty = True
    End If
    Unload frm_Color

    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Sub txtbx_Code_Change()
    bDirty = True
    Call UpdateMCL
End Sub

Private Sub txtbx_Description_Change()
    bDirty = True
End Sub

Private Sub txtbx_Title_Change()
    bDirty = True
    Call UpdateMCL
End Sub

Private Sub txtbx_Title_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

   If Button = 2 Then
        ' Right Mouse Click
        Dim myExplorer As Outlook.Explorer
        Set myExplorer = ThisOutlookSession.ActiveExplorer
  
        Dim objCommandBars As Office.CommandBars
        Set objCommandBars = myExplorer.CommandBars
  
        Dim myCommandBar As Office.CommandBar
        Set myCommandBar = objCommandBars("TestPopup")
  
        myCommandBar.ShowPopup
  
   End If

End Sub

Private Sub btn_GetNextCode_Click()

    ' Read the projects - get the highest code, add one

End Sub

' - Categories

' - Links

Private Sub btn_SelectWinFolder_Click()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":btn_SelectWinFolder_Click"
    
    On Error GoTo ThrowException
    
'    Dim xlObj As New Excel.Application
'    Dim oDialog As FileDialog
'    Set oDialog = xlObj.FileDialog(msoFileDialogFolderPicker)
'    oDialog.Show
'
'    If oDialog.SelectedItems.Count > 0 Then
'        txtbx_WinFolder.Text = oDialog.SelectedItems(1)
'        bDirty = True
'    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Sub btn_SelectOutlookFolder_Click()

    Dim ut As New Utilities
    Dim destFldr As Outlook.Folder
    Set destFldr = ut.SelectOutlookFolder
    If Not IsNothing(destFldr) Then
        txtbx_OutlookFolder.Text = destFldr.FolderPath
        bDirty = True
    End If

End Sub

Private Sub btn_CreateOutlookFolder_Click()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":btn_CreateOutlookFolder_Click"
    
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

' - Planning Tab

Private Sub chkbx_IncludeCompletedTasks_Change()
    Dim bInclude As Boolean
    bInclude = chkbx_IncludeCompletedTasks.value
    tm.LoadByProject f_Project.Name, bInclude
End Sub

Private Sub btn_AddTask_Click()
    tm.NewTask f_Project
End Sub

Private Sub btn_DeleteTask_Click()
    tm.DeleteTask
End Sub

Private Sub btn_EditTask_Click()
    tm.OpenTask
End Sub


' - Buttons

Private Sub btn_Save_Click()
    ' Me.Hide
    
    ' Save the Project's Properties
    If bDirty Then
        SaveToDataStore
        RaiseEvent Closing(Me, True)
    Else
        RaiseEvent Closing(Me, False)
    End If
    
    Me.Hide
    'Unload Me
    
End Sub

Private Sub btn_Cancel_Click()

    If bDirty Then
        ' Prompt for Save
        
    End If

    Unload Me
End Sub

' - Constructor

Private Sub UserForm_Initialize()

    ' Set up enumerated combo_boxes - olTaskStatus
    cmbobx_Status.AddItem "Not Started" ' 0
    cmbobx_Status.AddItem "In Progress" ' 1
    cmbobx_Status.AddItem "Waiting"     ' 3
    cmbobx_Status.AddItem "Deferred"    ' 4
    cmbobx_Status.AddItem "Complete"    ' 2
    cmbobx_Status.value = "Not Started"
    
    cmbobx_Priority.AddItem "Low"       ' Outlook.OlImportance.olImportanceLow    ' 0
    cmbobx_Priority.AddItem "Normal"    ' Outlook.OlImportance.olImportanceNormal ' 1
    cmbobx_Priority.AddItem "High"      ' Outlook.OlImportance.olImportanceHigh   ' 2
    cmbobx_Priority.value = "Normal"
    
    ' - Setup Task Manager
    Set tm = New TaskManager
    Set tm.ListView = lv_Tasks
    
    ' - - Set up Category Colors ComboBox
    ' Create imglst from Palette Tab
    imglst_Colors.ListImages.Add 1, "None", Image1.Picture
    imglst_Colors.ListImages.Add 2, "Red", Image2.Picture
    imglst_Colors.ListImages.Add 3, "Orange", Image3.Picture
    imglst_Colors.ListImages.Add 4, "Peach", Image4.Picture
    imglst_Colors.ListImages.Add 5, "Yellow", Image5.Picture
    imglst_Colors.ListImages.Add 6, "Green", Image6.Picture
    imglst_Colors.ListImages.Add 7, "Teal", Image7.Picture
    imglst_Colors.ListImages.Add 8, "Olive", Image8.Picture
    imglst_Colors.ListImages.Add 9, "Blue", Image9.Picture
    imglst_Colors.ListImages.Add 10, "Purple", Image10.Picture
    imglst_Colors.ListImages.Add 11, "Maroon", Image11.Picture
    imglst_Colors.ListImages.Add 12, "Steel", Image12.Picture
    imglst_Colors.ListImages.Add 13, "Dark Steel", Image13.Picture
    imglst_Colors.ListImages.Add 14, "Gray", Image14.Picture
    imglst_Colors.ListImages.Add 15, "Dark Gray", Image15.Picture
    imglst_Colors.ListImages.Add 16, "Black", Image16.Picture
    imglst_Colors.ListImages.Add 17, "Dark Red", Image17.Picture
    imglst_Colors.ListImages.Add 18, "Dark Orange", Image18.Picture
    imglst_Colors.ListImages.Add 19, "Dark Peach", Image19.Picture
    imglst_Colors.ListImages.Add 20, "Dark Yellow", Image20.Picture
    imglst_Colors.ListImages.Add 21, "Dark Green", Image21.Picture
    imglst_Colors.ListImages.Add 22, "Dark Teal", Image22.Picture
    imglst_Colors.ListImages.Add 23, "Dark Olive", Image23.Picture
    imglst_Colors.ListImages.Add 24, "Dark Blue", Image24.Picture
    imglst_Colors.ListImages.Add 25, "Dark Purple", Image25.Picture
    imglst_Colors.ListImages.Add 26, "Dark Maroon", Image26.Picture

    ' Add imglst pics to the comboItems
    ' imgcmbobx_Color.ImageList = imglst_Colors
    ' For i = 1 To imglst_Colors.ListImages.Count
    '    imgcmbobx_Color.ComboItems.Add i, , imglst_Colors.ListImages(i).Key, imglst_Colors.ListImages(i).Key
    ' Next
    
    ' Default to 'None' Category
    ' imgcmbobx_Color.selectedItem = imgcmbobx_Color.ComboItems(1)
    
    ' Default Logo
    imgLogo.Picture = imglst_Colors.ListImages(1).Picture
    
    ' Create an empty project
    Set f_Project = New fmeProject
    
    ' Set Dirty Flag
    bDirty = False
    bDirtyMCL = False
    
    ' tb_main.buttons.Add 1, "A", "test"
    
End Sub


Private Sub UserForm_Terminate()

    Set tm = Nothing
    Set f_Project = Nothing

End Sub

Private Sub UserForm_Activate()
    UpdateTitle
End Sub

' - Methods

''' Initialize the UI with the specified Project
Public Sub Load(ByVal p As fmeProject)

    ' Capture internal project
    Set f_Project = p
    
    ' Update UI
    SetUI
    
End Sub


' - Supporting Methods

''' Handles sync'ing the MCLName to the title and code fields
Private Sub UpdateMCL()

    If chkbx_CombineTitleCode.value Then
        ' Combine the code and title
        lbl_MCLName.Caption = txtbx_Code.Text & " - " & txtbx_Title.Text
    Else
        ' Show just the title
        lbl_MCLName.Caption = txtbx_Title.Text
    End If

    bDirty = True
    bDirtyMCL = True
    
End Sub

''' Load internal Project from the control values
Private Sub GetUI()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetUI"
    
    On Error GoTo ThrowException
    
    With f_Project
    
        ' Strings
        .Name = txtbx_Title.Text
        .Code = txtbx_Code.Text
        .Description = txtbx_Description.Text
        .WindowsFolder = txtbx_WinFolder.Text
        .OutlookFolder = txtbx_OutlookFolder.Text
        
        ' Enums
        .SetStatusFromName (cmbobx_Status.value)
        .SetPriorityFromName (cmbobx_Priority.value)
         ' Tracked separately - Set imgLogo.Picture = imglst_Colors.ListImages(.Color + 1).Picture
        
        ' Bools
        .CombineTitleCode = chkbx_CombineTitleCode.value
        .Active = chkbx_Active.value
        
    End With
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

  
End Sub

''' Set control values to the internal Project object instance
Private Sub SetUI()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SetUI"
    
    On Error GoTo ThrowException
    
    With f_Project
    
        ' Strings
        txtbx_Title.Text = .Name
        txtbx_Code.Text = .Code
        txtbx_Description.Text = .Description
        txtbx_WinFolder.Text = .WindowsFolder
        txtbx_OutlookFolder.Text = .OutlookFolder
        
        ' Enums
        cmbobx_Status.value = .GetStatusName
        cmbobx_Priority.value = .GetPriorityName
        Set imgLogo.Picture = imglst_Colors.ListImages(.Color + 1).Picture
        
        ' Bools
        chkbx_CombineTitleCode.value = .CombineTitleCode
        chkbx_Active.value = .Active
        
    End With
    
    UpdateTitle
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Sub Status(Optional ByVal msg As String = "")
    txtbx_Status.Text = msg
End Sub

Private Sub UpdateTitle()

    Dim strCap As String
    If Not f_Project Is Nothing Then
        If Len(f_Project.Name) > 0 Then
            strCap = f_Project.Name & " - " & Commands.AppName
        Else
            strCap = "Project Properties"
        End If
    Else
        strCap = "Project Properties"
    End If
    Me.Caption = strCap

End Sub

''' Saves the internal project to the local Datastore
Private Sub SaveToDataStore()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SaveToDataStore"
    
    On Error GoTo ThrowException
    
    ' Load the internal Project with UI values
    GetUI
    
    Dim ut As New Utilities
    
    With f_Project
        If bDirtyMCL Then
            If Len(.mclId) = 0 Then
                ' Add MCL record
                .mclId = ut.AddtoMCL(.Subject, .Color)
            Else
                ' Update MCL record
                ut.UpdateMCL .mclId, .Subject, .Color
            End If
        End If
    End With
    
    Dim ds As New dsDataStore
    ds.Connect
    
    Dim bDone As Boolean
    bDone = ds.Save(f_Project, "Project")
    If bDone Then ds.AcceptChanges

    GoTo Finally
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
Finally:
    Set ut = Nothing
    Set ds = Nothing

End Sub

Private Function TryGetRelativePosition(ByVal ctrl As control, _
                                         ByRef X As Integer, ByRef Y As Integer, _
                                Optional ByVal sp As Integer = 0) As Boolean
                                         
    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":TryGetRelativePosition"
        
    On Error GoTo ThrowException
    
    Dim tX As Integer
    Dim tY As Integer
    
    Dim titleBarWidth As Integer
    titleBarHeight = 23
       
    ' UserForm screen position
    tX = Me.Left
    tY = Me.Top
    
    ' Return position aligned to the left and under the specified control
    X = tX + ctrl.Left  '(Me.Width / 2)
    Y = tY + ctrl.Top + titleBarHeight + ctrl.Height ' (Me.Height / 2)
    
    '  Assume starts in center of application screen
    TryGetRelativePosition = True
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    TryGetRelativePosition = False
    
End Function


