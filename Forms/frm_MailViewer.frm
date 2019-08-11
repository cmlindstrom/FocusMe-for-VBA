VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_MailViewer 
   Caption         =   "Mail Viewer"
   ClientHeight    =   6990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8040
   OleObjectBlob   =   "frm_MailViewer.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_MailViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



' Fields

Private Const rootClass As String = "MailViewer"

' Connect to the Outlook Interface
Private ut As Utilities

' - - Resizing
Private WithEvents m_clsAnchors As CAnchors
Attribute m_clsAnchors.VB_VarHelpID = -1

' Property Declarations
Dim cnv As Conversation
Dim tsk As Outlook.TaskItem
Dim selNode As node

' Events

Public Event collectionUpdated()

' Properties

''' Form Title Bar Text
Public Property Let title(ByVal formCaption As String)
    Me.Caption = formCaption
End Property
Public Property Get title() As String
    Caption = Me.Caption
End Property

''' HeaderText at the top of the form within the form body
Public Property Let Header(ByVal formHeader As String)
    lbl_Header.Caption = formHeader
End Property
Public Property Get Header() As String
    Header = lbl_Header.Caption
End Property

''' Conversation Container
Public Property Get Mail() As Conversation
    Set Mail = cnv
End Property

''' Task for which the Mail is being viewed
Public Property Get Task() As Outlook.TaskItem
    Set Task = tsk
End Property




' Event Handlers

' - - TreeView

Private Sub tv_Mail_Click()

    ' Turn off highlight
    If Not selNode Is Nothing Then
        selNode.ForeColor = GetOleColor(Black)
    End If
    
    If Not tv_Mail.SelectedItem Is Nothing Then
    
        ' Get MailItem
        Set selNode = tv_Mail.SelectedItem
        Dim m As Outlook.MailItem
        Set m = FindMailItem(tv_Mail.SelectedItem)
    
        selNode.ForeColor = GetOleColor(Blue)
    
        ' Populate preview
        FillTextBox m
        Me.txtbx_Properties.SetFocus
        Me.txtbx_Properties.SelStart = 0
        
    End If

End Sub

Private Sub tv_Mail_DblClick()

    ' Get MailItem
    Dim tNode As node
    Set tNode = tv_Mail.SelectedItem
    Dim m As Outlook.MailItem
    Set m = FindMailItem(tv_Mail.SelectedItem)
    
    ' Display MailItem
    If Not IsNothing(m) Then m.Display
    
End Sub

' - - Buttons

Private Sub btn_OK_Click()
    Unload Me
End Sub

Private Sub btn_OpenTask_Click()
    If Not IsNothing(tsk) Then tsk.Display
End Sub

Private Sub btn_Collapse_Click()

    For Each n In Me.tv_Mail.Nodes
        If n.Expanded Then n.Expanded = False
    Next

End Sub

Private Sub btn_Expand_Click()

    For Each n In Me.tv_Mail.Nodes
        If Not n.Expanded Then n.Expanded = True
    Next

End Sub

' Constructor

Private Sub UserForm_Initialize()
    
    Call SetAnchors
    
    Set ut = New Utilities
    
    Set selNode = Nothing
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Outlook reboots with the Title bar Close button selected
    If CloseMode = vbFormControlMenu Then Cancel = 1
End Sub

Private Sub UserForm_Terminate()
    Set m_clsAnchors = Nothing
    Set ut = Nothing
End Sub

' Methods

Public Sub Load(ByVal t As TaskItem)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Load"

    If Not IsNothing(t) Then
        ' Capture the Task
        Set tsk = t
        ' Get the Reference Id
        Dim ut As New Utilities
        Dim rId As String
        rId = ut.GetReferenceID(t)
        ' Load the Task's conversations
        LoadByReference rId
    Else
        strTrace = "WARNING: No TaskItem specified."
        LogMessage strTrace, strRoutine
    End If
    
End Sub

Private Sub LoadByReference(ByVal refId As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":LoadByReference"
    
    On Error GoTo ThrowException
    
    UpdateStatus "Loading related messages: " & refId & "..."
    Me.tv_Mail.Nodes.Add key:="root", text:="Loading..."
    
    If Len(refId) = 0 Then
        strTrace = "Reference Id was missing - showing empty tree."
        LogMessage strTrace, strRoutine
        GoTo NoMessages
    End If
    
    Dim MailItems As New ArrayList
    
    strTrace = "Find referenced MailItems in the Mail Folders."
    Dim fldrList As New ArrayList
    Set fldrList = ut.GetMailFolders
    Dim f As Outlook.Folder
    For Each f In fldrList
        ' Search a tracked folder
        Dim arList As ArrayList
        Set arList = ut.GetConversationViaBody(refId, f)
        If IsNothing(arList) Then GoTo SkipFolder
        ' Capture MailItems
        For Each o In arList
            If Not o Is Nothing Then MailItems.Add o
        Next
SkipFolder:
    Next
    
    Dim s As String
    Dim m As Outlook.MailItem
    Dim tm As Outlook.MailItem
    Dim bFnd As Boolean
    Dim conversations As New ArrayList
    If MailItems.Count > 0 Then
        strTrace = "Narrow to conversations, versus individual mailItems."
        For Each m In MailItems
            bFnd = False
            For Each tm In conversations
                If m.ConversationTopic = tm.ConversationTopic Then
                    bFnd = True
                    Exit For
                End If
            Next
            If Not bFnd Then conversations.Add m
        Next
        
        strTrace = "Found " & MailItems.Count & " related messages in " & conversations.Count & " conversations."
        UpdateStatus strTrace
        
        Dim allItems As ArrayList
        Set allItems = New ArrayList
        Dim tmpItems As ArrayList
        For Each m In conversations
            Set tmpItems = ut.GetConversationFromMailItem(m)
            If Not IsNothing(tmpItems) Then
                If tmpItems.Count > 0 Then
                    For i = 0 To tmpItems.Count - 1
                        allItems.Add tmpItems(i)
                    Next
                End If
            End If
        Next
        
        strTrace = "Found " & allItems.Count & " related messages... in " & conversations.Count & " conversations."
        UpdateStatus strTrace
        
        strTrace = "Creating the Conversation Tree."
        Me.tv_Mail.Nodes.Clear
        Set cnv = New Conversation
        Set cnv.List = allItems
        cnv.CreateTree Me.tv_Mail
    Else
NoMessages:
        strTrace = "Presenting the rendered tree."
        Me.tv_Mail.Nodes.Clear
        Me.tv_Mail.Nodes.Add key:="root", text:="No related mail..."
        strTrace = "No communications associated with this task..."
        UpdateStatus strTrace
    End If

    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

' Supporting Methods

Private Sub UpdateStatus(Optional ByVal msg As String = "")
    
    Me.statusBar.SimpleText = msg

End Sub

Private Function FindMailItem(ByVal tn As node) As Outlook.MailItem

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":FindMailItem"
    
    On Error GoTo ThrowException
    
    If IsNothing(tn) Then
        strTrace = "A null tree node encountered."
        GoTo ThrowException
    End If
    
    Dim retMail As Outlook.MailItem
    
    If Not IsNothing(cnv) Then
        Dim cv As ConversationNode

        For Each cv In cnv.Tree.Nodes
            Dim tmpNode As node
            Set tmpNode = cv.Tag
            If Not IsNothing(tmpNode) Then
                If tmpNode.key = tn.key Then
                    Set retMail = cv.Context
                    Exit For
                End If
            End If
        Next
    Else
        strTrace = "Tree not initialized."
        GoTo ThrowException
    End If
    
    Set FindMailItem = retMail
    
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set FindMailItem = Nothing

End Function

Private Sub FillTextBox(ByVal m As Outlook.MailItem)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":FillTextBox"
    
    On Error GoTo ThrowException
    
    If IsNothing(m) Then
        strTrace = "A null Outlook MailItem encountered."
        GoTo ThrowException
    End If
    
    ' Clear the box
    Me.txtbx_Properties.text = ""
    
    strTrace = "Sent: " & m.SentOn & "  Received: " & m.ReceivedTime & vbCrLf
    strTrace = strTrace & "From: " & m.sender & vbCrLf
    strTrace = strTrace & "Subject: " & m.Subject & vbCrLf
    strTrace = strTrace & String$(20, "-") & vbCrLf
    strTrace = strTrace & m.body ' Mid(m.body, 1, 250)
       
    Me.txtbx_Properties.text = strTrace
    
    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Private Sub SetAnchors()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SetAnchors"

    On Error GoTo ThrowException

    ' Set up for Resizable form
    Set m_clsAnchors = New CAnchors
    
    Set m_clsAnchors.Parent = Me
    
    ' restrict minimum size of userform
    m_clsAnchors.MinimumWidth = 414.75
    m_clsAnchors.MinimumHeight = 374.25
    
        ' Set Anchors
    With m_clsAnchors
    
        ' Header Text
        .Anchor("lbl_Header").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleRight Or _
                                                enumAnchorStyleTop
    
        ' Buttons
        .Anchor("btn_Collapse").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleTop
        .Anchor("btn_Expand").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleTop
        .Anchor("btn_OK").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleTop
        .Anchor("btn_OpenTask").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleTop
                
        ' Panes
        .Anchor("tv_Mail").AnchorStyle = enumAnchorStyleTop Or _
                                            enumAnchorStyleLeft Or enumAnchorStyleRight
        
        ' ListView
        With .Anchor("txtbx_Properties")
            .AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleRight _
                            Or enumAnchorStyleTop Or enumAnchorStyleBottom
            .MinimumHeight = 84
        End With
               
        ' Status Bar
        .Anchor("statusBar").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleRight Or _
                                                enumAnchorStyleBottom
    End With

    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

