VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_WorkspaceProperties 
   Caption         =   "Workspace Properties"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4455
   OleObjectBlob   =   "frm_WorkspaceProperties.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_WorkspaceProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' - Fields

Private Const rootClass As String = "frm_WorkspaceProperties"

Private f_Workspace As fmeWorkspace

' - Events

Public Event Closing(ByVal sender As Object, ByVal saved As Boolean)

Dim bDirty As Boolean

' - Properties

''' The workspace that is being edited
Public Property Get Workspace() As fmeWorkspace
    Set Workspace = f_Workspace
End Property

' - Event Handlers

Private Sub txtbx_Description_Change()
    bDirty = True
End Sub

Private Sub txtbx_Title_Change()
    bDirty = True
End Sub

Private Sub btn_Save_Click()

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
    RaiseEvent Closing(Me, False)
    Unload Me
End Sub

Private Sub txtbx_Title_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Caption = txtbx_Title.text & " - " & Commands.AppName
End Sub

' - Constructor

Private Sub UserForm_Initialize()

    ' Set the Window Title
    Me.Caption = "Untitled - " & Commands.AppName

    ' Initialize Variables
    bDirty = False

    ' Create an empty Workspace
    Set f_Workspace = New fmeWorkspace

End Sub

Private Sub UserForm_Activate()
    Dim strTrace As String
    strTrace = ""
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Dim strTrace As String
    
    If bDirty Then
        strTrace = "Changes have been made to the Workspace properties, " & _
                    vbCrLf & vbCrLf & "Would you like to save before closing this dialog?"
        Dim dResult As DialogResult
        dResult = MsgBox(strTrace, vbYesNoCancel, "Closing - Workspace Properties")
        
        If dResult = DialogResult_Cancel Then
            Cancel = True
            Exit Sub
        End If
        
        If dResult = DialogResult_Yes Then
            SaveToDataStore
            RaiseEvent Closing(Me, True)
        Else
            RaiseEvent Closing(Me, False)
        End If

    End If

    Set f_Workspace = Nothing
    
End Sub

Private Sub UserForm_Terminate()
    Set f_Workspace = Nothing
End Sub

' - Methods

''' Initialize the UI with the specified Workspace
Public Sub Load(ByVal w As fmeWorkspace)

    ' Capture internal project
    Set f_Workspace = w
    
    ' Update UI
    SetUI
    
End Sub

''' Initializes the UI from the Workspace Name
Public Sub LoadFromName(ByVal title As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":LoadFromName"
    
    On Error GoTo ThrowException

    Dim ds As New dsDataStore
    ds.Connect
    
    Dim w As fmeWorkspace
    Set w = ds.GetItemByProperty("Workspace", "Name", title)
    If w Is Nothing Then
        Set f_Workspace = New fmeWorkspace
        Me.btn_Save.Caption = "Add"
    Else
        Set f_Workspace = w
        Me.btn_Save.Caption = "Update"
    End If
    SetUI
    
    GoTo Finally
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
Finally:
    Set ds = Nothing
    

End Sub

' - Supporting methods

''' Load internal Project from the control values
Private Sub GetUI()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetUI"
    
    On Error GoTo ThrowException
    
    With f_Workspace
        .Name = txtbx_Title.text
        .Description = txtbx_Description.text
        .Code = txtbx_Code.text
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
    
    With f_Workspace
        txtbx_Title.text = .Name
        txtbx_Description.text = .Description
        txtbx_Code.text = .Code
        
        ' Set the Window Title
        Me.Caption = .Name & " - " & Commands.AppName
        If Len(.Name) = 0 Then Me.Caption = "Untitled - " & Commands.AppName
        
    End With
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub


''' Saves the internal workspace to the local Datastore
Private Sub SaveToDataStore()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SaveToDataStore"
    
    On Error GoTo ThrowException
    
    ' Load the internal object with the UI values
    GetUI
    
    ' Save to the local datastore
    Dim ds As New dsDataStore
    ds.Connect
    
    Dim bDone As Boolean
    bDone = ds.Save(f_Workspace, "Workspace")
    If bDone Then ds.AcceptChanges

    GoTo Finally
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
Finally:
    Set ut = Nothing
    Set ds = Nothing

End Sub

