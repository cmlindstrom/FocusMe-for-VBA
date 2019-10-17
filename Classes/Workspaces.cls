VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Workspaces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' - - Fields

Private Const rootClass As String = "Workspaces"

Dim f_sortOn As enuSortOn ' Name, Code
Dim sortDir As enuSortDirection
Dim myItems As ArrayList

' - - Properties

''' Latest collection of Items
Public Property Set Items(ByVal ar As ArrayList)
    Set myItems = ar
End Property
Public Property Get Items() As ArrayList
    Set Items = myItems
End Property

''' Sort Property
Property Let SortBy(ByVal srt As enuSortOn)
    f_sortOn = srt
End Property
Property Get SortBy() As enuSortOn
    SortBy = f_sortOn
End Property

' - - Constructor

Private Sub Class_Initialize()

    Set myItems = New ArrayList
    f_sortOn = Subject
    sortDir = Ascending

End Sub

' - - Methods

Public Sub AddItem(ByVal w As fmeWorkspace)
    myItems.Add w
    InternalSort f_sortOn, sortDir
End Sub

Public Sub RemoveItem(ByVal w As fmeWorkspace)
    myItems.Remove w
End Sub

Public Sub RemoveAt(ByVal idx As Integer)
    myItems.RemoveAt idx
End Sub

Public Function Item(ByVal idx As Integer) As fmeWorkspace
    Set Item = myItems(idx)
End Function

Public Function GetItemById(ByVal id As String) As fmeWorkspace

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetItemById"

    Dim retItem As fmeWorkspace
    Set retItem = Nothing

    Dim bFnd As Boolean
    bFnd = False
    Dim w As fmeWorkspace
    For Each w In myItems
        If LCase(w.id) = LCase(id) Then
            Set retItem = w
            bFnd = True
            Exit For
        End If
    Next
    
    Set GetItemById = retItem
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetItemById = Nothing

End Function

Public Sub Sort(ByVal srt As enuSortOn, Optional ByVal dir As enuSortDirection)
    
    If Not IsMissing(dir) Then sortDir = dir
    InternalSort srt, sortDir
   
End Sub

''' Returns the size of the current collection
Public Function Count() As Integer
    Count = myItems.Count
End Function

' - - Supporting Methods

Private Sub InternalSort(ByVal srt As enuSortOn, ByVal dir As enuSortDirection)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":InternalSort"
    
    Dim w As fmeWorkspace
    Dim thisValue As String
    Dim thisItem As fmeWorkspace
    Dim nextValue As String
    Dim nextItem As fmeWorkspace
    Dim bPassDirty As Boolean
    bPassDirty = True
    
    Do While bPassDirty
    
        ' Assume clean pass
        bPassDirty = False
    
        ' Evalute items
        Dim i As Integer
        For i = 0 To myItems.Count - 1
            ' Get current value
            Set w = myItems(i)
            thisValue = GetIndexValue(w)
            Set thisItem = w
            ' Get Next Value
            If Not i = myItems.Count - 1 Then
                Set w = myItems(i + 1)
                nextValue = GetIndexValue(w)
                Set nextItem = w
            Else
                ' Done with the pass
                GoTo PassComplete
            End If
        
            ' Compare
            Dim bSwap As Boolean
            bSwap = False
            If sortDir = Ascending Then bSwap = nextValue < thisValue
            If sortDir = Descending Then bSwap = nextValue > thisValue
        
            ' Swap necessary?
            If bSwap Then
                ' Code Check
                Dim iCurrent As Integer
                iCurrent = myItems.Count
        
                ' Swapped two items, mark the pass as dirty
                ' - Remove nextItem
                myItems.Remove nextItem
                ' - Insert nextItem before thisItem
                myItems.Insert i, nextItem
                
                ' Mark the pass as dirty
                bPassDirty = True
                
                ' Code Check
                If Not iCurrent = myItems.Count Then
                    strTrace = "Elongating the array."
                End If
            End If
        
        Next
PassComplete:
    
    Loop
    
    Exit Sub
    
ThrowException:
    LogMessage strTrace, strRoutine

End Sub
Private Function GetIndexValue(ByVal w As fmeWorkspace) As String

    Dim strReturn As String
    
    Select Case f_sortOn
        Case enuSortOn.name
            strReturn = w.name
        Case enuSortOn.Code
            strReturn = w.Code
    End Select
    
    GetIndexValue = strReturn
    
End Function


