VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SortCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Fields

Private Const rootClass As String = "SortCollection"

Private innerList As ArrayList

' Events

' Properties

''' Returns the internal collection
Public Property Get List() As ArrayList
    Set List = innerList
End Property

' Constructor

Private Sub Class_Initialize()
    Set innerList = New ArrayList
End Sub

' Methods

''' Sorts an ArrayList of Objects on specified property
''' e.g. Sort("Index",myList) or Sort("Index",myList,enuSortDirection.Descending)
Public Sub Sort(ByVal propName As String, ByRef arList As ArrayList, _
       Optional ByVal dir As enuSortDirection = enuSortDirection.Ascending)
                
    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":Sort"
    
    If Len(propName) = 0 Then
        strTrace = "Unidentified search property."
        GoTo ThrowException
    End If
    If IsNothing(arList) Then
        strTrace = "A null incoming collection encountered."
        GoTo ThrowException
    End If
    
    Set innerList = arList
    InternalSort propName, dir
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

' Supporting Methods

Private Sub InternalSort(ByVal propName As String, ByVal sortDir As enuSortDirection)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":InternalSort"
    
    On Error GoTo ThrowException
    
    Dim thisValue As String
    Dim thisItem As Object
    Dim nextValue As String
    Dim nextItem As Object
    Dim bPassDirty As Boolean
    bPassDirty = True
    
    Do While bPassDirty
    
        ' Assume clean pass
        bPassDirty = False
    
        Dim obj As Object
        ' Evalute items
        For i = 0 To innerList.Count - 1
            ' Get current value
            Set thisItem = innerList(i)
            thisValue = GetSortingValue(thisItem, propName)
            ' Get Next Value
            If Not i = innerList.Count - 1 Then
                Set nextItem = innerList(i + 1)
                nextValue = GetSortingValue(nextItem, propName)
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
                iCurrent = innerList.Count
        
                ' Swapped two items
                ' - Remove nextItem
                innerList.Remove nextItem
                ' - Insert nextItem before thisItem
                innerList.Insert i, nextItem
                
                ' Mark the pass as dirty
                bPassDirty = True
                
                ' Code Check
                If Not iCurrent = innerList.Count Then
                    strTrace = "Elongating the array."
                    LogMessage strTrace, strRoutine
                End If
            End If
        
        Next
PassComplete:
    
    Loop
   
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Function GetSortingValue(ByVal itm As Object, ByVal propName As String) As Variant

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":GetSortingValue"
    
    On Error GoTo ThrowException
    
    Dim retValue As Variant
    retValue = GetProperty(itm, propName)
    
    ' Massage if necessary
    If IsDate(retValue) Then
        Dim dt As Date
        dt = retValue
        retValue = Format(dt, "yyyymmdd")
    End If

    GetSortingValue = retValue
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    retValue = ""

End Function

