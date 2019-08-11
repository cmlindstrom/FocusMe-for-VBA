VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Fields

Private Const rootClass As String = "DataTable"

Private strName As String
Private colColumns As Columns
Private rowRows As Rows

Dim strData(1, 1) As String

' Properties

''' Name of the table
Public Property Let Name(ByVal nme As String)
    strName = nme
End Property
Public Property Get Name() As String
    Name = strName
End Property

''' Table Columns
Public Property Let Columns(ByVal cols As Columns)
    Set colColumns = cols
End Property
Public Property Get Columns() As Columns
    Set Columns = colColumns
End Property

''' Table Rows
Public Property Let Rows(ByVal rws As Rows)
    Set rowRows = rws
End Property
Public Property Get Rows() As Rows
    Set Rows = rowRows
End Property

' Events

' Constructor

Private Sub Class_Initialize()
    ' set up the tables columns collectin
    Set colColumns = New Columns
    colColumns.Table = Me
    
    ' initialize the table's rows collection
    Set rowRows = New Rows
    rowRows.Table = Me
End Sub

' Methods

' Supporting Methods

