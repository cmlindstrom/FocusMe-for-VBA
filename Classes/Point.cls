VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' - Fields

' - Properties

Dim f_X As Long
Dim f_Y As Long

'''
Public Property Let x(ByVal fx As Long)
    f_X = fx
End Property
Public Property Get x() As Long
    x = f_X
End Property

'''
Public Property Let y(ByVal fy As Long)
    f_Y = fy
End Property
Public Property Get y() As Long
    y = f_Y
End Property
