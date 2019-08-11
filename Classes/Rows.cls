VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Rows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Fields

Private Const rootClass As String = "Rows"

Private innerList As Object ' ArrayList

Private dt As DataTable

' Properties

''' Table the Rows collection reside in
Public Property Let Table(ByVal dTable As DataTable)
    Set dt = dTable
End Property
Public Property Get Table() As DataTable
    Set Table = dt
End Property

''' Returns the collection
Public Property Get Items() As ArrayList
    Set Items = innerList
End Property

''' Number of Rows in the collection
Public Property Get Count() As Integer
    Count = innerList.Count
End Property

' Events

' Constructor

Private Sub Class_Initialize()
    ' Set innerList = New ArrayList
    CreateArrayList
End Sub

' Methods

''' Adds a Row with multiple values
Public Function Add(ByVal values As Variant) As DataRow

    ' Create a row
    Dim retRw As DataRow
    Set retRw = New DataRow
    retRw.Parent = Me
    retRw.ParseValues values
    innerList.Add retRw
    Set Add = retRw

End Function

''' Adds the dataRow to the collection
Public Sub AddRow(ByVal dr As DataRow)
    dr.Parent = Me
    innerList.Add dr
End Sub

''' Returns the DataRow at the specified index
Public Function Row(ByVal idx As Integer) As DataRow
    
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Row"
    
    On Error GoTo ThrowException
    
    Set Item = innerList(idx)
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Function

''' Removes the Row at the specified index
Public Sub RemoveAt(ByVal idx As Integer)
    innerList.RemoveAt idx
End Sub

''' Removes the Row from the collection
Public Sub Remove(ByVal dr As DataRow)
    innerList.Remove dr
End Sub

''' Clears the Rows collection
Public Sub Clear()
    innerList.Clear
End Sub

''' Returns a matching dataRow for the specified column and value
Public Function GetByColumnValue(ByVal colName As String, ByVal colValue As Variant) As DataRow

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetByColumnValue"
    
    If Len(colName) = 0 Then
        strTrace = "Unknown or empty column name encountered."
        GoTo ThrowException
    End If
    
    Dim drReturn As DataRow
    Set drReturn = Nothing
    
    Dim dr As DataRow
    For Each dr In innerList
        Dim v As Variant
        v = dr.GetItem(colName)
        If v = colValue Then
            Set drReturn = dr
            Exit For
        End If
    Next
    
    Set GetByColumnValue = drReturn
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetByColumnValue = Nothing

End Function


' Supporting Methods

Private Sub CreateArrayList()
  Set innerList = CreateObject("System.Collections.ArrayList")
End Sub
