VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConversationTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Fields

Private Const rootClass As String = "ConversationTree"
    
Private strTitle As String

Private cvRoot As ConversationNode
Private cvNodes As ArrayList

Private ndRoot As node
Private Tree As TreeView4

' Events

' Properties
         
''' <summary>
''' Title of the Conversation Tree
''' </summary>
''' <value></value>
''' <returns>String:</returns>
''' <remarks></remarks>
Public Property Let title(ByVal t As String)
     strTitle = t
End Property
Public Property Get title() As String
    title = strTitle
End Property

''' The tree's nodes: ArrayList of ConversationNode
Public Property Set Nodes(ByVal nds As ArrayList)
    Set cvNodes = nds
End Property
Public Property Get Nodes() As ArrayList
    Set Nodes = cvNodes
End Property

''' <summary>
''' Root node for the Tree
''' </summary>
''' <value>Node:</value>
''' <returns>Node:</returns>
''' <remarks></remarks>
Public Property Set TreeRoot(ByVal r As node)
    Set ndRoot = r
End Property
Public Property Get TreeRoot() As node
    Set Root = ndRoot
End Property

''' TreeView to populate
Public Property Set TreeViewControl(ByVal tv As TreeView4)
    Set Tree = tv
End Property
Public Property Get TreeViewControl() As TreeView4
    Set TreeViewControl = Tree
End Property
         
' Constructor

Private Sub Class_Initialize()
    Set cvNodes = New ArrayList
End Sub

' Methods

''' <summary>
''' Loads the Conversation Tree from a list of Outlook.MailItems
''' </summary>
''' <param name="List">ArrayList:</param>
''' <remarks></remarks>
Public Sub LoadTree(ByVal List As ArrayList)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":LoadTree"
    
    On Error GoTo ThrowException
                    
    If IsNothing(List) Then
        strTrace = "A null list encountered."
        GoTo ThrowException
    End If
    If List.Count = 0 Then
        strTrace = "An empty list was presented."
        GoTo ThrowException
    End If

    ' Capture the Outlook.MailItems
    Dim nd As ConversationNode
    For Each o In List
                  
        Set nd = New ConversationNode
        nd.Fill o
        cvNodes.Add nd

        If Len(strTitle) = 0 Then
            strTitle = nd.Topic
        Else
            ' If Not _title.ToLower.Contains(nd.Topic.ToLower) Then
            If Not InStr(1, LCase(strTitle), LCase(nd.Topic)) > 0 Then
                strTitle = "Multiple Conversations"
            End If
        End If
                   
    Next

    LogMessage "Before Sort", strRoutine
    For Each nd In cvNodes
        strTrace = Len(nd.Index) & ": " & nd.Index
        LogMessage strTrace, strRoutine
    Next
    
    ' Sort the messages to group their conversations and their levels
    Dim collSort As New SortCollection
    collSort.Sort "Index", List
    
    LogMessage "After Sort", strRoutine
    For Each nd In cvNodes
        strTrace = Len(nd.Index) & ": " & nd.Index
        LogMessage strTrace, strRoutine
    Next

    ' Establish Ranks
    Dim rnk As Integer
    Dim lengths As New ArrayList
    Dim topRank As Integer
    topRank = -1
    
    Dim Level As Integer
    Dim prevNd As ConversationNode
    Set prevNd = cvNodes(0)
    prevNd.Level = 1
    If cvNodes.Count > 1 Then
        For i = 1 To cvNodes.Count - 1
            ' Set the Rank/Level
            Set nd = cvNodes(i)
'            If InStr(nd.Index, prevNd.Index) > 0 Then
'                nd.Level = prevNd.Level + 1
'            Else
'                nd.Level = 1
'            End If
            
            If Contains(prevNd.Index, nd.Index) Then
                nd.Level = prevNd.Level + 1
                Set nd.Parent = FindParentNode(i)
            Else
                Set nd.Parent = FindParentNode(i, Level)
                nd.Level = Level + 1
            End If
            Set prevNd = nd

            
            ' Set the ParentNode
            ' SetParentNode nd
        Next
    End If
       
    If IsNothing(Tree) Then
        Set Tree = New TreeView4
    Else
        Tree.Nodes.Clear
    End If
    Set ndRoot = Tree.Nodes.Add(key:="root", text:=strTitle)

    Dim nRoot As New ConversationNode
    nRoot.name = strTitle
    nRoot.Level = 0
    Set nRoot.Tag = ndRoot
    Set cvRoot = nRoot
    Set ndRoot.Tag = nRoot
                    
    Dim eid As String
    For Each nd In cvNodes
        If nd.Level = 1 Then
            ' Connect first level to the root node
            Dim nde As node
            eid = "K" & Common.GenerateUniqueID(4) ' nd.Context.EntryID
           ' Set nde = Tree.nodes.Add(ndRoot.Key, tvwChild, eid, nd.Name)
            Set nde = Tree.Nodes.Add(relative:=ndRoot.key, relationship:=tvwChild, key:=eid, text:=nd.name)
            ' Set the ConversationNode's TreeNode
            Set nd.Tag = nde
        Else
            Dim desc As String
            ' Connect to the parent node
            Dim parentNode As ConversationNode
            Set parentNode = nd.Parent
            If Not IsNothing(parentNode) Then
                ' Set the TreeView Node
                Dim pNode As node
                Set pNode = parentNode.Tag ' Associated TreeView Node
                ' parentNode.Children.Add nde
                eid = "K" & Common.GenerateUniqueID(4) ' nd.Context.EntryID
                ' desc = Len(nd.Index) & " " & nd.name
                desc = nd.name
                Set nde = Tree.Nodes.Add(pNode.key, tvwChild, eid, desc)
                ' set the ConversationNode's TreeNode
                Set nd.Tag = nde
            End If
        End If
    Next

    strTrace = "Conversation TreeView; node count: " & Tree.Nodes.Count
    LogMessage strTrace, strRoutine

    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub


Public Function FindParentNode(ByVal idxCurrentNode As Integer, _
                      Optional ByRef lvl As Integer = 1) As ConversationNode

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":FindParentNode"
    
    On Error GoTo ThrowException
    
    Dim retNode As ConversationNode
    Set retNode = Nothing
    
    If idxCurrentNode < 1 Then
        strTrace = "Method starts on the second record."
        GoTo ThrowException
    End If
    
    Dim cnd As ConversationNode
    Set cnd = cvNodes(idxCurrentNode)
    If cnd Is Nothing Then
        strTrace = "A null Node encountered."
        GoTo ThrowException
    End If
    
    Dim i As Integer
    Dim nd As ConversationNode
    For i = idxCurrentNode - 1 To 0 Step -1
        Set nd = cvNodes(i)
        If Contains(nd.Index, cnd.Index) Then
            ' Parent node found for current node
            Set retNode = nd
            ' Parent level
            lvl = nd.Level
            Exit For
        End If
    Next
    
    If retNode Is Nothing Then lvl = 0
    
    Set FindParentNode = retNode
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set FindParentNode = Nothing

End Function

''' <summary>
''' Creates a string representation of the ConversationTree
''' </summary>
''' <returns>String:</returns>
''' <remarks></remarks>
Public Function ToString() As String

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ToString"
    
    If IsNothing(ndRoot) Then
        strTrace = "ConversationTree has not been loaded."
        GoTo ThrowException
    End If

    Dim strReturn As String
    PrintTreeRecursive ndRoot, strReturn

    ToString = strReturn
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    ToString = ""

End Function

' Supporting Methods

Private Sub PrintTreeRecursive(ByVal startNode As node, ByRef text As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":PrintTreeRecursive"
    
    On Error GoTo ThrowException
                    
    If IsNothing(startNode) Then
        strTrace = "A null start node encountered."
        GoTo ThrowException
    End If

    Dim nd As ConversationNode
    Set nd = startNode.Tag
    If Not IsNothing(nd) Then
        Dim Fill As String
        Fill = String$(nd.Level, "+")
        text = text & Fill & " " & nd.name & ", " & nd.Index & vbCrLf
        ' sb.Append (fill & " " & nd.Name & ", " & nd.index & vbCrLf)
        If startNode.Nodes.Count > 0 Then
            Dim tn As node
            For Each tn In startNode.Nodes
                PrintTreeRecursive tn, text
            Next
        End If
    End If

ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Function SetRank(ByVal thisNode) As Integer

End Function

''' Returns the rank or level in the tree where the targetLength falls
Private Function GetRankOld(ByVal targetLength As Integer, ByVal lengths As ArrayList) As Integer

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetRank"
    
    On Error GoTo ThrowException
                    
    Dim rank As Integer
    rank = -1
    Dim lg As Integer

    For i = 0 To lengths.Count - 1
        lg = lengths(i)
        If targetLength = lg Then
            rank = i + 1
            Exit For
        End If
    Next

    GetRank = rank
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    GetRank = -1

End Function

''' Uses the Outlook.MailItem's Index property to set the
''' item's preceding MailItem
Private Sub SetParentNode(ByRef childNode As ConversationNode)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SetParentNode"
    
    On Error GoTo ThrowException
    
    If IsNothing(childNode) Then
        strTrace = "A null conversationNode encountered."
        GoTo ThrowException
    End If
    
    If childNode.Level = 1 Then
        ' Set the first level to the root node
        Set childNode.Parent = cvRoot
    Else
        ' Find 2nd and higher levels preceding message
        strTrace = "Evaluating a node at level: " & childNode.Level
        ' Search the list for its parent
        Dim cvNd As ConversationNode
        For Each cvNd In cvNodes
            ' A child message will have it's parent's conversationIndex within it
            If InStr(1, childNode.Index, cvNd.Index) > 0 Then ' string1 = expression being searched, string2 = expression sought
                Set childNode.Parent = cvNd
                Exit For
            End If
        Next
        
    End If
    
    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Private Function FindParentTreeNodeOld(ByVal myNodes As Nodes, ByVal childNode As ConversationNode) As node

    Dim n As node
    Dim nRetVal As node

    For Each n In myNodes
        Dim tNode As ConversationNode
        Set tNode = n.Tag
        If Not IsNothing(tNode) Then
            If InStr(1, childNode.Index, tNode.Index) > 0 Then
                ' Look one level back
                Dim lvls As Integer
                lvls = childNode.Level - tNode.Level
                If lvls = 1 Then
                    ' tNode is the parent to the childNode
                    Set nRetVal = n
                    Exit For
                End If
            End If
        End If
        Set nRetVal = FindParentTreeNodeRecursiveOld(n, childNode)
        If Not IsNothing(nRetVal) Then Exit For
    Next

    Set FindParentTreeNodeOld = nRetVal

End Function


Private Function FindParentTreeNodeRecursiveOld(ByVal treeNode As node, ByVal childNode As ConversationNode) As node

    Dim tn As node
    Dim nR As node

    For Each tn In treeNode.Nodes
        Dim tNode As ConversationNode
        Set tNode = tn.Tag
        If Not IsNothing(tNode) Then
            If InStr(1, childNode.Index, tNode.Index) > 0 Then
                ' Look one level back
                Dim lvls As Integer
                lvls = childNode.Level - tNode.Level
                If lvls = 1 Then
                    ' tNode is the parent to the childNode
                    Set nR = tn
                    Exit For
                End If
            End If
        End If
        Set nR = FindParentTreeNodeRecursiveOld(tn, childNode)
        If Not IsNothing(nR) Then Exit For
    Next

    Set FindParentTreeNodeRecursiveOld = nR

End Function
