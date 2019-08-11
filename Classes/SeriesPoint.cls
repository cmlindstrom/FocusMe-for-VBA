VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeriesPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' #Region "Fields"

Private Const rootClass As String = "SeriesPoint"
        
' #Region "Properties"

Dim f_value As Double
Dim f_label As Variant
Dim f_samples As Integer
Dim f_toolTip As String

Private Const fields As String = "Value,Label,Samples,ToolTip"

Dim f_labelWidth As Integer
Dim f_labelHeight As Integer


''' <summary>
''' Dependent value - often the 'Y' value.
''' </summary>
''' <value>Double</value>
''' <returns>Double</returns>
''' <remarks></remarks>
Public Property Let Value(ByVal dbL As Double)
    f_value = dbL
End Property
Public Property Get Value() As Double
    Value = f_value
End Property

''' <summary>
''' X-Axis label / value
''' </summary>
''' <value>Variant:</value>
''' <returns>Variant:</returns>
''' <remarks>Can be used to represent a category X value, or a #</remarks>
Public Property Let Label(ByVal lbl As Variant)
    f_label = lbl
End Property
Public Property Get Label() As Variant
    Label = f_label
End Property

''' <summary>
''' The number of samples taken for this data point.
''' </summary>
''' <value>Integer</value>
''' <returns>Integer</returns>
''' <remarks></remarks>
Public Property Let Samples(ByVal s As Integer)
    f_samples = s
End Property
Public Property Get Samples() As Integer
    Samples = f_samples
End Property

''' <summary>
''' Tip used for the data point, e.g. a label for a mouse-over.
''' </summary>
''' <value>String</value>
''' <returns>String</returns>
''' <remarks></remarks>
Public Property Let ToolTip(ByVal tip As String)
    f_toolTip = tip
End Property
Public Property Get ToolTip() As String
    ToolTip = f_toolTip
End Property

' Constructor

Private Sub Class_Initialize()

    f_value = 0#
    f_label = ""
    f_samples = 0
    f_toolTip = ""
    f_labelWidth = 0
    f_labelHeight = 10

End Sub

' Methods

''' Initializes the Object instance to the set optional
''' parameters
Public Sub Create(Optional ByVal Label As Variant = "", _
                    Optional ByVal Value As Double = 0#)

    f_label = Label
    f_value = Value
    
    Dim strLabel As String
    strLabel = CStr(Label)
    If Len(strLabel) > 0 Then
        Call SetLabelSize
    End If

End Sub

''' Create a copy of the SeriesPoint
Public Function Clone() As SeriesPoint

    Dim newPt As New SeriesPoint
    Call newPt.Create(Me.Label, Me.Value)
    
    Set Clone = newPt

End Function

''' Returns a string array of the object's property names
Public Function GetListOfProperties() As String()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetListOfProperties"
    
    On Error GoTo ThrowException

    Dim arReturn() As String
    
    arReturn = Split(fields, ",")
    
    GetListOfProperties = arReturn
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    GetListOfProperties = arReturn

End Function

' Supporting Methods

Private Sub SetLabelSize()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SetLabelSize"
    
    On Error GoTo ThrowException

    Dim strLabel As String
    If IsDate(f_label) Then
        strLabel = Format(f_label, "mm/dd/yyyy")
    Else
        strLabel = CStr(f_label)
    End If

    f_labelWidth = Len(strLabel) * 6
    f_labelHeight = 10

    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

