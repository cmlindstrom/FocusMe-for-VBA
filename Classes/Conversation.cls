VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Conversation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Fields

Private Const rootClass As String = "Conversation"

Private arList As ArrayList
Dim ctTree As ConversationTree

' Events

' Properties

''' <summary>
''' Flat list of the objects in the Outlook Conversation
''' </summary>
''' <value>ArrayList:</value>
''' <returns>ArrayList:</returns>
''' <remarks></remarks>
Public Property Set List(ByVal lst As ArrayList)
    Set arList = lst
End Property
Public Property Get List() As ArrayList
    Set List = arList
End Property

''' <summary>
''' Hierarchical list of objects in the Outlook Conversation
''' </summary>
''' <value>ConversationTree:</value>
''' <returns>ConversationTree:</returns>
''' <remarks>Can return a TreeNode for a TreeView</remarks>
Public Property Set Tree(ByVal ct As ConversationTree)
    Set ctTree = ct
End Property
Public Property Get Tree() As ConversationTree
    Set Tree = ctTree
End Property

' Constructor

Private Sub Class_Initialize()
    Set arList = New ArrayList
    Set ctTree = New ConversationTree
End Sub

' Methods

''' <summary>
''' Retrieves all Outlook MailItems in a conversation associated with a specified MailItem.
''' </summary>
''' <param name="originalMail">Outlook.MailItem:</param>
''' <returns>ArrayList: collection of Outlook.MailItem</returns>
''' <remarks></remarks>
Public Function GetItems(ByVal originalMail As Outlook.MailItem) As ArrayList

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetItems"
    
    Dim ut As New Utilities
    
    Dim rList As ArrayList
    Set rList = ut.GetConversationFromMailItem(originalMail)
    If Not IsNothing(rList) Then
        ' Capture class collection
        Set arList = rList

        ' Create the Tree
        ' Call CreateTree
    
    Else
        strTrace = "Outlook query failed."
        GoTo ThrowException
    End If
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetItems = Nothing
        
End Function

''' <summary>
''' Creates the Tree from the List property
''' </summary>
''' <remarks>Useful with the List is set and a new Tree is desired.</remarks>
Public Sub CreateTree(Optional ByVal tv As TreeView4 = Nothing)
    Set ctTree = New ConversationTree
    Set ctTree.TreeViewControl = tv
    ctTree.LoadTree arList
End Sub

' Supporting Methods

