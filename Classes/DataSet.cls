VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' - - Fields
Private Const rootClass As String = "DataSet"

' - - Properties

Dim f_name As String
Dim f_tables As DataTables
Dim f_nameSpace As String

''' Name of the DataSet
Public Property Let Name(ByVal nme As String)
    f_name = nme
End Property
Public Property Get Name() As String
    Name = f_name
End Property

''' Tables contained within the DataSet
Public Property Get Tables() As DataTables
    Set Tables = f_tables
End Property

''' NameSpace for the DataSet
Public Property Let NameSpace(ByVal nme As String)
    f_nameSpace = nme
End Property
Public Property Get NameSpace() As String
    NameSpace = f_nameSpace
End Property

' - - Constructor

Private Sub Class_Initialize()
    Set f_tables = New DataTables
    Set f_tables.DataSet = Me
End Sub

Private Sub Class_Terminate()
    Set f_tables = Nothing
End Sub

' - - Methods

''' Clear - clears all rows in all tables
Public Sub Clear()
    Dim dt As DataTable
    For Each dt In f_tables.Items
        dt.Rows.Clear
    Next
End Sub

' Clone - copies the structure of the DataSet

' Copy - copies the structure and the data

' GetXml - returns the XML representing the DataSet
Public Function GetXml() As String
    Dim dom As DOMDocument60
    Set dom = SaveDataSetViaDOM
    GetXml = dom.XML
End Function

' HasChanges - tracking changes (new, deleted or modified rows)

' Load from the permanent storage location
Public Sub Load()

End Sub

''' Loads the DataSet from and XML string
Public Sub LoadXml(ByVal strXml As String)
    ' Reset the DataSet
    Reset
    ' Load from the xml string
    LoadDataSetViaDOM strXml
End Sub

' Merge (DataRow()) (DataSet) (DataTable)

' ReadXml (Stream) (String-filePath) (TextReader) (XmlReader)
Public Sub ReadXmlFile(ByVal filePath As String)
    ' Reset the DataSet
    Reset
    ' Load from an xml formatted file
    LoadDataSet filePath
End Sub

' RejectChanges

' Reset - Clears all tables and removes all relations, foreign constraints and tables from the DataSet
Public Sub Reset()

    ' Re-initialize the DataSet
    f_name = ""
    f_nameSpace = ""
    
    Dim dt As DataTable
    For Each dt In f_tables.Items
        Set dt = Nothing
    Next
    
    Set f_tables = New DataTables
    Set f_tables.DataSet = Me

End Sub

' WriteXml (Stream) (string-filePath)
Public Sub WriteXmlFile(ByVal filePath As String)
    SaveDataSet filePath
End Sub

' Return a table from the collection
Public Function GetTableByName(ByVal tblName As String) As DataTable

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetTableByName"
    
    On Error GoTo ThrowException
    
    Dim dtReturn As DataTable
    Set dtReturn = Nothing
    
    Dim dt As DataTable
    For Each dt In f_tables.Items
        If LCase(dt.Name) = LCase(tblName) Then
            Set dtReturn = dt
            Exit For
        End If
    Next
    
    Set GetTableByName = dtReturn
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetTableByName = Nothing

End Function

' - - Supporting Methods

Private Function LoadDataSetViaDOM(ByVal strXml As String) As DOMDocument60

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":LoadDataSetViaDOM"
    
    On Error GoTo ThrowException

    ' Create a DOM document
    Dim xmlDoc As MSXML2.DOMDocument60
    Set xmlDoc = New MSXML2.DOMDocument60
    If Not xmlDoc.LoadXml(strXml) Then
        strTrace = "Failed to load the xml content."
        GoTo ThrowException
    End If
    
    ' Load DataSet structure
    
    strTrace = "Retrieve the DataSet node."
    Dim nd As IXMLDOMNode
    Dim oNodes As IXMLDOMNodeList
    Set oNodes = xmlDoc.SelectNodes("//DataSet")
    If oNodes.Length = 0 Then
        strTrace = "Failed to find the DataSet beginning - file is malformed."
        GoTo ThrowException
    End If
    
    strTrace = "Establish the DataSet."
    Set nd = oNodes.NextNode
    f_name = nd.Attributes.getNamedItem("name").Text
    f_nameSpace = nd.Attributes.getNamedItem("nameSpace").Text
    
    strTrace = "Get the table collection."
    Set oNodes = xmlDoc.SelectNodes("//Table")
    If oNodes.Length = 0 Then
        strTrace = "Failed to find any Tables - file is boring."
        GoTo ThrowException
    End If
    
    strTrace = "Loop thru the tables."
    Dim i As Integer
    For i = 0 To oNodes.Length - 1
    
        strtace = "Get the Table Node."
        Set nd = oNodes.NextNode
        If nd Is Nothing Then
            strTrace = "Failed to get the next node: index: " & i
            LogMessage "WARNING: " & strTrace, strRoutine
            GoTo SkipTable
        End If
        
        Dim tblName As String
        tblName = nd.Attributes.getNamedItem("name").Text

        strTrace = "Loading the table: " & tblName
        LoadDataSetTable xmlDoc, tblName
                
SkipTable:

    Next
    
    strTrace = "Return the DOM document."
    Set LoadDataSetViaDOM = xmlDoc
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set LoadDataSetViaDOM = Nothing

End Function
Private Sub LoadDataSetTable(ByVal doc As MSXML2.DOMDocument60, ByVal tableName As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":LoadDataSetTable"
    
    On Error GoTo ThrowException
    
    strTrace = "Create the Table in the DataSet."
    Dim dt As New DataTable
    dt.Name = tableName
        
    strTrace = "Populate the new Table."
    Dim tblNodes As IXMLDOMNodeList
    Set tblNodes = doc.SelectNodes("//" & tableName)
    If tblNodes.Length = 0 Then
        strTrace = "Encountered an empty table: " & tableName & ", not added to DataSet."
        LogMessage "WARNING: " & strTrace, strRoutine
        GoTo ThrowException
    End If
        
    strTrace = "Process the Table Rows."
    Dim j As Integer
    Dim rwNode As IXMLDOMNode
    Dim colValues As New ArrayList
    Dim dr As DataRow
    For j = 0 To tblNodes.Length - 1
        
        strTrace = "Capture the Row Node."
        ' Set rwNode = tblNodes.NextNode
        Set rwNode = tblNodes.Item(j)
            
        strTrace = "Gather the Column Data for the selected Row."
        Dim k As Integer
        Dim vt As Integer
        Dim colNode As IXMLDOMNode
        Dim kvp As KeyValuePair
        colValues.Clear
        For k = 0 To rwNode.ChildNodes.Length - 1
            Set colNode = rwNode.ChildNodes(k)
            Set kvp = New KeyValuePair
            ' Get property name
            kvp.Key = colNode.nodeName
            ' Get Value Type
            vt = CInt(colNode.Attributes.getNamedItem("varType").Text)
            ' Get property value
            kvp.Value = GetTypedValue(colNode.nodeTypedValue, vt)
            ' Capture value pair
            colValues.Add kvp
        Next
            
        strTrace = "Check for first Row."
        If j = 0 Then
            strTrace = "Create the table columns."
            For k = 0 To colValues.Count - 1
                Set kvp = colValues(k)
                dt.Columns.Add kvp.Key
            Next
        End If
            
        strTrace = "Create the Row."
        Set dr = New DataRow
            
        strTrace = "Add the Column values to the new row."
        For k = 0 To colValues.Count - 1
            Set kvp = colValues(k)
            dr.Add kvp.Key, kvp.Value
        Next
            
        strTrace = "Add the row to the table."
        dt.Rows.AddRow dr
        
    Next
                   
    strTrace = "Add the table to the DataSet Collection."
    f_tables.Add dt

    GoTo Finally
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
Finally:
    Set dt = Nothing
    Set tblNodes = Nothing

End Sub
Private Function GetTypedValue(ByVal strVal As String, ByVal vType As Integer) As Variant

    On Error Resume Next

    Dim vReturn As Variant

    If vType = vbBoolean Then vReturn = CBool(strVal)
    If vType = vbDate Then vReturn = CDate(strVal)
    If vType = vbString Then vReturn = strVal
    If vType = vbInteger Then vReturn = CInt(strVal)
    If vType = vbLong Then vReturn = CLng(strVal)
    If vType = vbDecimal Then vReturn = CDbl(strVal)
    If vType = vbSingle Then vType = CSng(strVal)
    If vType = vbDouble Then vType = CDbl(strVal)
    
    GetTypedValue = vReturn
    
End Function

Private Function LoadDataSet(ByVal path As String) As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":LoadDataSet"
    
    On Error GoTo ThrowException
    
    Dim bReturn As Boolean
    bReturn = True
    
    Dim xmlContent As String
    xmlContent = Common.ReadTextFile(path)
    If Len(xmlContent) > 0 Then
        Dim dom As DOMDocument60
        Set dom = LoadDataSetViaDOM(xmlContent)
        If dom Is Nothing Then
            strTrace = "DOM failed to load with content of " & Len(xmlContent) & _
                            " bytes."
            bReturn = False
            LogMessage "WARNING: " & strTrace, strRoutine
        End If
    Else
        strTrace = "Failed to read or encountered empty contents in: " & filePath
        GoTo ThrowException
    End If
    
    LoadDataSet = bReturn
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    LoadDataSet = False

End Function

Private Function SaveDataSetViaDOM() As DOMDocument60

    ' Create a new DOM
    Dim dom As DOMDocument60
    Set dom = New DOMDocument60
    dom.async = True
    dom.validateOnParse = False
    dom.resolveExternals = False
    dom.preserveWhiteSpace = True
    
    Dim node As IXMLDOMNode
    Dim rootNode As IXMLDOMNode
    Dim rwNode As IXMLDOMNode
    Dim rwAttr As IXMLDOMAttribute
    
    ' Set the XML File Header
    Set node = dom.createProcessingInstruction("xml", "version='1.0'")
    dom.appendChild node
    
    ' Set the dataSet root
    Dim dsRoot As IXMLDOMNode
    Set dsRoot = dom.createElement("DataSet")
    
    ' Attach the dataSet name as an Attribute
    Dim dsAttr As IXMLDOMAttribute
    ' DataSet Name
    Set dsAttr = dom.createAttribute("name")
    dsAttr.Value = f_name
    dsRoot.Attributes.setNamedItem dsAttr
    ' DataSet NameSpace
    Set dsAttr = dom.createAttribute("nameSpace")
    dsAttr.Value = f_nameSpace
    dsRoot.Attributes.setNamedItem dsAttr
    
    ' Add the DataSet Node to the DOM
    dom.appendChild dsRoot
    
    ' Go through each table
    Dim dt As DataTable
    For Each dt In f_tables.Items
    
        ' Create the table root node
        Dim eName As String
        eName = ProcessName(dt.Name)
        Dim dtRoot As IXMLDOMNode
        Set dtRoot = dom.createElement("Table")
        
        ' Attach the table name as an attribute
        Dim dtAttr As IXMLDOMAttribute
        Set dtAttr = dom.createAttribute("name")
        dtAttr.Value = eName
        dtRoot.Attributes.setNamedItem dtAttr
        
        ' Create each row node & append to table node
        Dim dr As DataRow
        Dim dc As DataColumn
        For Each dr In dt.Rows.Items
            ' Create the Object's node (row)
            Set node = dom.createElement(eName)
    
            ' Get each Property's value (column)
            For Each dc In dt.Columns.Items
                ' Capture the object's property's value
                Dim v As Variant
                v = dr.GetItem(dc.Name)
                
                ' Create an XML node
                Set rwNode = dom.createElement(dc.Name)
                rwNode.Text = CStr(v)
                
                ' Add the variant type as an attribute to the rwNode
                Set rwAttr = dom.createAttribute("varType")
                rwAttr.Value = VarType(v)
                rwNode.Attributes.setNamedItem rwAttr
                
                ' Append to the row node
                node.appendChild rwNode
            Next
            
            ' Append Row node to Table node
            dtRoot.appendChild node
        Next
        
        ' Append table node to the dataSet root node
        dsRoot.appendChild dtRoot
    Next
    
    Set SaveDataSetViaDOM = dom

End Function
Private Function ProcessName(ByVal nme As String) As String
    Dim strReturn As String
    strReturn = Replace(nme, " ", "_")
    ProcessName = strReturn
End Function

Private Function SaveDataSet(ByVal path As String) As Boolean

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SaveDataSet"
    
    If Len(path) = 0 Then
        strTrace = "A null or empty path encountered."
        GoTo ThrowException
    End If
    
    On Error GoTo ThrowException
    
    Dim bReturn As Boolean
    bReturn = True
    
    ' Create the DOM
    Dim dom As DOMDocument60
    Set dom = SaveDataSetViaDOM
    If dom Is Nothing Then
        strTrace = "Failed to create a DOM Instance."
        GoTo ThrowException
    End If
    
    ' Extract XML from the DOM
    Dim strXml As String
    strXml = dom.XML

    ' Write XML to the path
    Common.WriteTextFile path, strXml
    
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    SaveDataSet = False

End Function

