VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmailProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Fields
Private Const rootClass As String = "EmailProcessing"

' Properties

' Operating OutlookItem
Private myItem As Object
Public Property Let CurrentItem(myObj As Object)
    myItem = myObj
End Property
Public Property Get CurrentItem() As Object
    CurrentItem = myItem
End Property




' Constructor
Private Sub Class_Initialize()
    
End Sub

