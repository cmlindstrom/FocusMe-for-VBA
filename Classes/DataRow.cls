VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Fields

Private Const rootClass As String = "DataRow"

''' Tracks the values within the row
''' - an array of KeyValuePair, where the key
'''     is the name of the column
Private innerList As ArrayList

Private parentRows As rows

' Properties

''' The Collection this row belongs to
Public Property Let Parent(ByVal rs As rows)
    Set parentRows = rs
End Property
Public Property Get Parent() As rows
    Set Parent = parentRows
End Property

' Events

' Constructor

Private Sub Class_Initialize()
    Set innerList = New ArrayList
End Sub

' Methods

Public Sub Add(ByVal colName As String, ByVal val As Variant)
    
    Dim kvp As New KeyValuePair
    kvp.key = colName
    kvp.value = val
    
    innerList.Add kvp
    
End Sub

''' Loads the DataRow with an array of values
Public Sub ParseValues(ByVal arValues As Variant)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ParseValues"
    
 '   On Error GoTo ThrowException
    
    Dim i As Integer
    i = 0
    For Each o In arValues
    
        Dim kvp As New KeyValuePair
        
        Dim c As Columns
        Set c = parentRows.Table.Columns
        Dim colName As String
        colName = c.Item(i).name
        ' kvp.Key = c.Item(i).Name
        ' kvp.Value = o
    
        If innerList.Count > i + 1 Then
            ' Existing value for this index
            Set innerList(i) = kvp
        Else
            ' Add new value
            Add colName, o
        End If
            
        i = i + 1
    Next
    
    Exit Sub
        
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set Item = Nothing

End Sub

''' Sets the row value at the specific column
''' SetItem(colName,newValue); SetItem(2, newValue); SetItem(dataColumn object, newValue)
Public Sub SetItem(ByVal v As Variant, ByVal val As Variant)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SetItem"
    
    On Error GoTo ThrowException

    Select Case VarType(v)
        Case vbEmpty
            strTrace = "An empty variable encountered."
            GoTo ThrowException
        Case vbNull
            strTrace = "A Null variable encountered."
            GoTo ThrowException
        Case vbInteger
            ' Get the value at an index
            Dim idx As Integer
            idx = v
            innerList(v) = val
        Case vbString
            ' Set the value at a specific column by colName
            SetValueByKey v, val
        Case vbUserDefinedType
            ' Get the value at a specific column
            Dim c As DataColumn
            c = var
            SetValueByKey c.name, val
        Case Else
            strTrace = "Failed to ascertain the variant type."
            GoTo ThrowException
    End Select
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Returns the row value at the specified column
Public Function GetItem(ByVal var As Variant) As Variant

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetItem"
    
    On Error GoTo ThrowException

    Dim retValue As Variant

    Select Case VarType(var)
        Case vbEmpty
            strTrace = "An empty variable encountered."
            GoTo ThrowException
        Case vbNull
            strTrace = "A Null variable encountered."
            GoTo ThrowException
        Case vbInteger
            ' Get the value at an index
            Dim idx As Integer
            idx = var
            Dim kvp As KeyValuePair
            Set kvp = innerList(idx)
            retValue = kvp.value
        Case vbString
            ' Get the value at a specific column by colName
            retValue = GetValueByKey(var)
        Case vbUserDefinedType
            ' Get the value at a specific column
            Dim c As DataColumn
            c = var
            retValue = GetValueByKey(c.name)
        Case Else
            strTrace = "Failed to ascertain the variant type."
            GoTo ThrowException
    End Select
    
    GetItem = retValue
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetItem = Nothing
    
End Function


' Supporting Methods

Private Sub SetValueByKey(ByVal keyName As String, ByVal val As Variant)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SetValueByKey"
    
    On Error GoTo ThrowException
    
    Dim kvp As KeyValuePair
    Set kvp = GetPairByKey(keyName)
    If IsNothing(kvp) Then
        strTrace = "WARNING: Failed to retrieve the column (" & keyName & ") value - adding keyValue pair."
        LogMessage strTrace, strRoutine
        
        Add keyName, val
        ' GoTo ThrowException
    End If
    
    kvp.value = val
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Function GetValueByKey(ByVal keyName As String) As Variant

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetValueByKey"
    
    On Error GoTo ThrowException
    
    Dim retValue As Variant
    Set retValue = Nothing
    
    Dim kvp As KeyValuePair
    Set kvp = GetPairByKey(keyName)
    If IsNothing(kvp) Then
        strTrace = "Failed to retrieve the column (" & keyName & ") value."
        GoTo ThrowException
    End If
    
    retValue = kvp.value
    
    GetValueByKey = retValue
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    GetValueByKey = vbNull

End Function

Private Function GetPairByKey(ByVal keyName As String) As KeyValuePair

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetValueByKey"
    
    On Error GoTo ThrowException
    
    Dim retKVP As KeyValuePair
    
    Dim kvp As KeyValuePair
    For Each kvp In innerList
        If LCase(kvp.key) = LCase(keyName) Then
            Set retKVP = kvp
            Exit For
        End If
    Next
    
    Set GetPairByKey = retKVP
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetPairByKey = Nothing

End Function
