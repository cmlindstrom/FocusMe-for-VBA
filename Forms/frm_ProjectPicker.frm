VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ProjectPicker 
   Caption         =   "All Projects"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7905
   OleObjectBlob   =   "frm_ProjectPicker.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_ProjectPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' - - Fields

Private Const rootClass As String = "frm_ProjectPicker"
Private Const FormTitle As String = "Project List - " & Commands.AppName

' ListView
Private WithEvents pm As ProjectManager
Attribute pm.VB_VarHelpID = -1

Dim sortingColumn As enuSortOn
Dim sortingDirection As enuSortDirection

' - - Resizing
Private m_clsAnchors As CAnchors

' - - Events

' - - Properties


' - - Event Handlers

' - - Constructor

Private Sub UserForm_Initialize()

    ' Set up for Resizable form
    Set m_clsAnchors = New CAnchors
    
    Set m_clsAnchors.Parent = Me
    
    ' restrict minimum size of userform
    m_clsAnchors.MinimumWidth = 407.25
    m_clsAnchors.MinimumHeight = 375
    
    ' Set Anchors
    With m_clsAnchors
    
        ' Description
        .Anchor("lbl_Description").AnchorStyle = enumAnchorStyleLeft Or _
                                                    enumAnchorStyleTop Or enumAnchorStyleRight
                                                    
        ' Filtering
        .Anchor("txtbx_Filter").AnchorStyle = enumAnchorStyleLeft Or _
                                                    enuanchorstyletop Or enumAnchorStyleRight
        .Anchor("btn_Filter").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleTop
        
        ' ListView
        With .Anchor("lv_Projects")
            .AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleRight _
                            Or enumAnchorStyleTop Or enumAnchorStyleBottom
            .MinimumHeight = 234
        End With
        
        .Anchor("chkbx_ShowInactiveProjects").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleBottom
        
        ' Buttons
        .Anchor("btn_Add").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleTop
        .Anchor("btn_Edit").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleTop
        .Anchor("btn_Delete").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleTop
        
        .Anchor("btn_Open").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleBottom
        
        .Anchor("btn_Save").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleBottom
        .Anchor("btn_Cancel").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleBottom
        
        ' Status Bar
        .Anchor("txtbx_Status").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleRight Or enumAnchorStyleBottom
    
    End With
    
    ' Initialize
    Set pm = New ProjectManager
    Set pm.ListView = lv_Projects
    
End Sub

Private Sub UserForm_Terminate()
     Set m_clsAnchors = Nothing
     Set pm = Nothing
End Sub

Private Sub UserForm_Activate()

    ' Handle Window Title
    Me.Caption = FormTitle
    
    ' Load the Projects
    Dim myImport As Projects
    Set myImport = pm.ImportFromMCL
    Set pm.Items = myImport
    ' pm.ListViewMultiSelect = True
    ' pm.ListViewCheckBox = True
    pm.Refresh
    
    ' Status the User
    Status "Items: " & pm.Items.Count & " Projects..."
    
End Sub

' - - Methods

''' Updates the status bar of the form
Public Sub Status(Optional msg As String = "")
    Me.txtbx_Status.Text = msg
End Sub
