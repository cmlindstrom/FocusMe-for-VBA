VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -  Fields
Private Const rootClass As String = "DataTables"

Private f_dataSet As DataSet

Private innerList As ArrayList

' - Events

' - Properties

''' DataSet the Tables Collection belongs to
Public Property Set DataSet(ByVal myDS As DataSet)
    Set f_dataSet = myDS
End Property
Public Property Get DataSet() As DataSet
    Set DataSet = f_dataSet
End Property

''' Returns the collection
Public Property Get Items() As ArrayList
    Set Items = innerList
End Property

' - Constructor

Private Sub Class_Initialize()
    Set innerList = New ArrayList
End Sub

Private Sub Class_Terminate()
    Set innerList = Nothing
End Sub

' - Methods

''' Adds a new column to the collection
Public Sub Add(ByVal dt As DataTable)
    innerList.Add dt
End Sub

''' Returns the DataTable at the given index
Public Function Table(ByVal idx As Integer) As DataTable
    Set Table = innerList(idx)
End Function

''' Returns a DataTable in the collection with the specified
''' name
Public Function GetTableByName(ByVal nme As String) As DataTable

    Dim dtReturn As DataTable
    Set dtReturn = Nothing
    
    Dim dt As DataTable
    Set dt = Nothing
    For Each dt In innerList
        If LCase(dt.name) = LCase(nme) Then
            Set dtReturn = dt
            Exit For
        End If
    Next

    Set GetTableByName = dtReturn

End Function

''' Removes the specified DataTable from the collection
Public Sub Remove(ByVal dt As DataTable)
    innerList.Remove dt
End Sub

''' Removes the DataTable as the specified index
Public Sub RemoveAt(ByVal idx As Integer)
    innerList.RemoveAt idx
End Sub

' - Supporting Methods


