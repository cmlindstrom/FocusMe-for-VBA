VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Fields

Private Const rootClass As String = "DataColumn"

Private f_Name As String
Private f_minWidth As Integer

' Properties

''' Name of the Column
Public Property Let Name(ByVal nme As String)
    f_Name = nme
End Property
Public Property Get Name() As String
    Name = f_Name
End Property

''' Minimum Column Width
''' - the narrowest it can be without truncating the
''' - column header name text
Public Property Get MinimumWidth() As Integer
    MinimumWidth = MeasureString(f_Name)
End Property

' Events

' Constructor

' Methods

' Supporting Methods


