VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csvExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Fields

Private Const rootClass As String = "csvExporter"

Private txtQualifier As String
Private txtDelimiter As String
Private ilExported As Integer

Private lineChar As String

' Events

Public Event ExportComplete()

' Properties

''' The 'Text Qualifier' is the character used to frame each value within the csv row.
Public Property Let TextQualifier(ByVal c As String)
    txtQualifier = c
End Property
Public Property Get TextQualifier() As String
    TextQualifier = txtQualifier
End Property

''' The 'Text Delimiter' is the character used between values within the csv row.
Public Property Let TextDelimiter(ByVal d As String)
    txtDelimiter = d
End Property
Public Property Get TextDelimiter() As String
    TextDelimiter = txtDelimiter
End Property

''' The 'Line Delimiter' is the character used to differentiate the end of a row/line
Public Property Let LineDelimiter(ByVal d As String)
    lineChar = d
End Property
Public Property Get LineDelimiter() As String
    LineDelimiter = lineChar
End Property

'  Constructor

Private Sub Class_Initialize()
    txtDelimiter = ","
    txtQualifier = """"
    lineChar = vbCrLf
End Sub

' Methods

''' <summary>
''' Exports the specified DataTable to a CSV formatted string
''' </summary>
''' <param name="dt">DataTable:</param>
''' <param name="fullFilePath" (Optional, defaults to empty)>String: if specified, saves the
''' exported csv formatted string to the specified file path.</param>
''' <param name="valueDelimiter" (Optional, defaults to empty)>String: is the character used between
''' to differentiate values within the data row</param>
''' <param name="qualifier">String: the character used to frame each value in a data row</param>
''' <param name="lneDelimiter">String: is the character used to differentiate the end of a row/line</param>
''' <returns>String:</returns>
Public Function ExportTable(ByVal dt As DataTable, _
                   Optional ByVal fullFilePath As String = "", _
                   Optional ByVal valueDelimiter As String = ",", _
                   Optional ByVal qualifier As String = "", _
                   Optional ByVal lneDelimiter As String = "") As String
              
              
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ExportTable"
              
    If Len(valueDelimiter) > 0 Then
        txtDelimiter = valueDelimiter
    End If
    If Len(qualifier) > 0 Then
        txtQualifier = qualifier
    End If
    If Len(lneDelimiter) > 0 Then
        lineChar = lneDelimiter
    End If
    If dt Is Nothing Then
        strTrace = "A null DataTable encountered."
        GoTo ThrowException
    End If
    
    On Error GoTo ThrowException
    
    Dim strReturn As String
    strReturn = ""
    
    Dim strTemp As String
    Dim strLine As String
    
    ' Common
    Dim i As Integer
    Dim col As DataColumn
    Dim dr As DataRow
    
    ' Export to CSV
    'strTrace = "Create Header Row."
    'strTemp = ""
    'For i = 0 To dt.Columns.Count - 1
    '    Set col = dt.Columns.Items(i)
    '    strTemp = strTemp & GetCSVValue(col.Name) & txtDelimiter
    'Next
    'strLine = Left(strTemp, Len(strTemp) - 1) & lineChar
    'strReturn = strReturn & strLine
    
    ' Export the table
    For Each dr In dt.Rows.Items
        strTemp = ""
        strLine = ""
        For i = 0 To dt.Columns.Count - 1
            Set col = dt.Columns.Items(i)
            strTemp = strTemp & GetCSVValue(dr.GetItem(col.Name)) & txtDelimiter
        Next
        strLine = Left(strTemp, Len(strTemp) - 1) & lineChar
        strReturn = strReturn & strLine
    Next
    
    If Len(fullFilePath) > 0 Then
        ' Save to Text File
        Call Common.WriteTextFile(fullFilePath, strReturn)
    End If
    
    ExportTable = strReturn
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    ExportTable = ""
                  
End Function

' Supporting Function

Private Function GetCSVValue(ByVal val As Variant) As String

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ExportTable"
    
    On Error GoTo ThrowException
    
    Dim strReturn As String
    strReturn = ""
    
    Select Case VarType(val)
        Case vbArray
        Case vbObject
        Case Else
            If Contains("""", CStr(val)) Then
                strReturn = txtQualifier & CStr(val) & txtQualifier
            Else
                strReturn = val
            End If
    End Select

    GetCSVValue = strReturn
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    GetCSVValue = ""

End Function
