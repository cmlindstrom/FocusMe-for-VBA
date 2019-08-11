VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Columns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Fields

Private Const rootClass As String = "Columns"

Private dt As DataTable

Private innerList As ArrayList

' Properties

''' Table the Columns collection reside in
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

' Events

' Constructor

Private Sub Class_Initialize()
    Set innerList = New ArrayList
End Sub

' Methods

''' Adds a new column to the collection
Public Function Add(ByVal colName As String) As DataColumn

    Dim col As New DataColumn
    col.Name = colName
    innerList.Add col

    Set Add = col

End Function

Public Function Item(ByVal idx As Integer) As DataColumn
    Set Item = innerList(idx)
End Function

Public Sub Remove(ByVal c As DataColumn)
    innerList.Remove c
End Sub

Public Sub RemoveAt(ByVal idx As Integer)
    innerList.RemoveAt idx
End Sub

''' Checks to see if a column exist with the specified name
Public Function IsColumnExist(ByVal nme As String)

    Dim bReturn As Boolean
    bReturn = False
    
    Dim c As DataColumn
    For Each c In innerList
        If LCase(c.Name) = LCase(nme) Then
            bReturn = True
            Exit For
        End If
    Next
    
    IsColumnExist = bReturn

End Function

' Supporting Methods


