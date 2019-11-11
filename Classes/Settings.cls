VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' Fields

Private Const rootClass As String = "Settings"

Private innerList As ArrayList

Private settingsFilePath As String

Private Enum enuSettingType
    None = 0
    Binary = 1
    DateTime = 2
    text = 3
    Number = 4
    real = 5
End Enum

''' Predefined Settings
Dim defSetting As Setting

''' Properties

''' Gets/Sets a flag to react to the Application level events
''' - able to turn off the code where FocusMe for Outlook is running
Public Property Let EnableAppEvents(ByVal bl As Boolean)
    SetSetting "EnableAppEvents", bl
End Property
Public Property Get EnableAppEvents() As Boolean
    Set defSetting = GetSetting("EnableAppEvents")
    If IsNothing(defSetting) Then Set defSetting = Create("EnableAppEvents", True)
    EnableAppEvents = defSetting.value
End Property

''' Gets/Sets a flag to automatically moves a 5D processed email
''' to the Project or Destination Folder
Public Property Let AutoMove(ByVal bl As Boolean)
    SetSetting "AutoMove", bl
End Property
Public Property Get AutoMove() As Boolean
    Set defSetting = GetSetting("AutoMove")
    If IsNothing(defSetting) Then Set defSetting = Create("AutoMove", True)
    AutoMove = defSetting.value
End Property

''' Gets/Sets a flag to move on specific Events (True) or All Events (False)
''' provided the AutoMove flag is set to True
Public Property Let MoveOnSpecificEvents(ByVal bl As Boolean)
    SetSetting "MoveOnSpecificEvents", bl
End Property
Public Property Get MoveOnSpecificEvents() As Boolean
    Set defSetting = GetSetting("MoveOnSpecificEvents")
    If IsNothing(defSetting) Then Set defSetting = Create("MoveOnSpecificEvents", False)
    MoveOnSpecificEvents = defSetting.value
End Property

''' Gets/Sets the Move on Defer to Task flag
''' Only checked if MoveOnSpecificEvents is True
Public Property Let MoveOnDeferToTask(ByVal bl As Boolean)
    SetSetting "MoveOnDeferToTask", bl
End Property
Public Property Get MoveOnDeferToTask() As Boolean
    Set defSetting = GetSetting("MoveOnDeferToTask")
    If IsNothing(defSetting) Then Set defSetting = Create("MoveOnDeferToTask", False)
    MoveOnDeferToTask = defSetting.value
End Property

''' Gets/Sets the Move on Defer to Calendar flag
''' Only checked if MoveOnSpecificEvents is True
Public Property Let MoveOnDeferToAppt(ByVal bl As Boolean)
    SetSetting "MoveOnDeferToAppt", bl
End Property
Public Property Get MoveOnDeferToAppt() As Boolean
    Set defSetting = GetSetting("MoveOnDeferToAppt")
    If IsNothing(defSetting) Then Set defSetting = Create("MoveOnDeferToAppt", False)
    MoveOnDeferToAppt = defSetting.value
End Property

''' Gets/Sets the Move on Delegate flag
''' Only checked if MoveOnSpecificEvents is True
Public Property Let MoveOnDelegate(ByVal bl As Boolean)
    SetSetting "MoveOnDelegate", bl
End Property
Public Property Get MoveOnDelegate() As Boolean
    Set defSetting = GetSetting("MoveOnDelegate")
    If IsNothing(defSetting) Then Set defSetting = Create("MoveOnDelegate", False)
    MoveOnDelegate = defSetting.value
End Property

''' Gets/Sets the Move on File in Drawer flag
''' Only checked if MoveOnSpecificEvents is True
Public Property Let MoveOnFileInDrawer(ByVal bl As Boolean)
    SetSetting "MoveOnFileInDrawer", bl
End Property
Public Property Get MoveOnFileInDrawer() As Boolean
    Set defSetting = GetSetting("MoveOnFileInDrawer")
    If IsNothing(defSetting) Then Set defSetting = Create("MoveOnFileInDrawer", False)
    MoveOnFileInDrawer = defSetting.value
End Property

''' Gets/Sets the Move on Reply flag
''' Only checked if MoveOnSpecificEvents is True
Public Property Let MoveOnReply(ByVal bl As Boolean)
   SetSetting "MoveOnReply", bl
End Property
Public Property Get MoveOnReply() As Boolean
    Set defSetting = GetSetting("MoveOnReply")
    If IsNothing(defSetting) Then Set defSetting = Create("MoveOnReply", False)
    MoveOnReply = defSetting.value
End Property

''' Gets/Sets the flag to ignore moving a mailItem if
''' it resides in the 'Sent Items' folder
Public Property Let IgnoreSentMailMove(ByVal bl As Boolean)
   SetSetting "IgnoreSentMailMove", bl
End Property
Public Property Get IgnoreSentMailMove() As Boolean
    Set defSetting = GetSetting("IgnoreSentMailMove")
    If IsNothing(defSetting) Then Set defSetting = Create("IgnoreSentMailMove", True)
    IgnoreSentMailMove = defSetting.value
End Property

''' Gets/Sets a flag to include or exclude a Task as an attachment
''' when forwarding a task in email
Public Property Let AttachTaskToEmail(ByVal bl As Boolean)
   SetSetting "AttachTaskToEmail", bl
End Property
Public Property Get AttachTaskToEmail() As Boolean
    Set defSetting = GetSetting("AttachTaskToEmail")
    If IsNothing(defSetting) Then Set defSetting = Create("AttachTaskToEmail", False)
    AttachTaskToEmail = defSetting.value
End Property

''' Gets/Sets a flag to start the Task Window when Outlook starts up
Public Property Let ShowOnStartup(ByVal bl As Boolean)
    SetSetting "ShowTaskWindowOnStartUp", bl
End Property
Public Property Get ShowOnStartup() As Boolean
    Set defSetting = GetSetting("ShowTaskWindowOnStartUp")
    If IsNothing(defSetting) Then Set defSetting = Create("ShowTaskWindowOnStartUp", True)
    ShowOnStartup = defSetting.value
End Property

''' Gets/Sets the Destination folder setting
Public Property Let DestinationFolder(ByVal fldrPath As String)
    SetSetting "DestinationFolder", fldrPath
End Property
Public Property Get DestinationFolder() As String
    Set defSetting = GetSetting("DestinationFolder")
    If IsNothing(defSetting) Then Set defSetting = Create("DestinationFolder", "")
    DestinationFolder = defSetting.value
End Property

''' Gets/Sets Timecard setting: Include Only Busy Status
Public Property Let TimecardIncludeBusyOnly(ByVal flag As Boolean)
    SetSetting "TimecardIncludeBusyOnly", flag
End Property
Public Property Get TimecardIncludeBusyOnly() As Boolean
    Set defSetting = GetSetting("TimecardIncludeBusyOnly")
    If defSetting Is Nothing Then Set defSetting = Create("TimecardIncludeBusyOnly", True)
    TimecardIncludeBusyOnly = defSetting.value
End Property

''' Gets/Sets Timecard setting: Include 24h Events
Public Property Let TimecardInclude24hEvents(ByVal flag As Boolean)
    SetSetting "TimecardInclude24hEvents", flag
End Property
Public Property Get TimecardInclude24hEvents() As Boolean
    Set defSetting = GetSetting("TimecardInclude24hEvents")
    If defSetting Is Nothing Then Set defSetting = Create("TimecardInclude24hEvents", False)
    TimecardInclude24hEvents = defSetting.value
End Property

''' Gets/Sets Use Windows Context Menu vs internal ContextMenu Window
Public Property Let UseContextMenuWindows(ByVal b As Boolean)
    SetSetting "UseContextMenuWindows", b
End Property
Public Property Get UseContextMenuWindows() As Boolean
    Set defSetting = GetSetting("UseContextMenuWindows")
    If defSetting Is Nothing Then Set defSetting = Create("UseContextMenuWindows", True)
    UseContextMenuWindows = defSetting.value
End Property

''' Gets/Sets starting location of the FME Task Pane
Public Property Set TaskPaneLocation(ByVal pt As Point)
    SetSetting "TaskPaneLocation", pt.X & ";" & pt.y
End Property
Public Property Get TaskPaneLocation() As Point
    Set defSetting = GetSetting("TaskPaneLocation")
    If defSetting Is Nothing Then
        Set defSetting = Create("TaskPaneLocation", "200;200")
    End If
    Dim vars() As String
    vars = Split(defSetting.value, ";")
    
    Dim pt As New Point
    pt.X = CLng(vars(LBound(vars)))
    pt.y = CLng(vars(UBound(vars)))
    Set TaskPaneLocation = pt
    
End Property

''' Retrieves all of the Settings
Public Property Get Items() As ArrayList
    Set Items = innerList
End Property

''' Constructor

Private Sub Class_Initialize()
    Set innerList = New ArrayList
    
    LoadSettings
End Sub

''' Methods

''' Adds a new setting to the collection
Public Function Add(ByVal s As Setting) As Setting
    innerList.Add s
End Function

''' Creates a new setting in the collection
Public Function Create(ByVal nme As String, ByVal val As Variant) As Setting
    ' Create the new Setting
    Dim s As New Setting
    s.Name = nme
    s.value = val
    s.ValueType = GetSettingType(val)
    ' Add it to the collection
    Add s
    ' Return the new Setting
    Set Create = s
End Function

''' Removes a Setting from the collection
Public Sub Remove(ByVal s As Setting)
    innerList.Remove s
End Sub

''' Removes a Setting from the collection at the specified index
Public Sub RemoveAt(ByVal idx As Integer)
    innerList.RemoveAt idx
End Sub

''' Set Setting by name and value
Public Sub SetSetting(ByVal nme As String, ByVal val As Variant)

    Dim s As Setting
    Set s = GetSetting(nme)
    If Not IsNothing(s) Then
        ' Capture new value
        s.value = val
    Else
        ' New Setting
        Create nme, val
    End If
    
    ' Save Settings
    SaveSettings

End Sub

''' Get Setting by name
Public Function GetSetting(ByVal nme As String) As Setting

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetSetting"
    
    On Error GoTo ThrowException
    
    Dim retSetting As Setting
    Set retSetting = Nothing
    
    Dim s As Setting
    For Each s In innerList
        If InStr(1, LCase(s.Name), LCase(nme)) > 0 Then
            Set retSetting = s
            Exit For
        End If
    Next
    
    If IsNothing(retSetting) Then
        strTrace = "WARNING: Failed to find a setting named: '" & nme & "' in the collection."
        LogMessage strTrace, strRoutine
    End If
    
    Set GetSetting = retSetting
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetSetting = Nothing
    
End Function

''' Saves the collection to the XML file
Public Sub Save()
    SaveSettings
End Sub

''' Supporting Methods

Private Function GetSettingType(ByVal v As Variant) As enuSettingType

    Dim vType As enuSettingType
    
    vType = text
    
    If VarType(v) = vbBoolean Then vType = Binary
    If VarType(v) = vbDate Then vType = DateTime
    If VarType(v) = vbString Then vType = text
    If VarType(v) = vbInteger Then vType = Number
    If VarType(v) = vbLong Then vType = Number
    If VarType(v) = vbDecimal Then vType = real
    If VarType(v) = vbSingle Then vType = real
    If VarType(v) = vbDouble Then vType = real
    
    GetSettingType = vType

End Function

Private Function GetStoredValue(ByVal v As Variant) As String
    GetStoredValue = CStr(v)
End Function

Private Function GetSettingValue(ByVal storedValue As String, _
                                ByVal storedType As enuSettingType) As Variant

    Dim v As Variant
    
    Select Case storedType
        Case enuSettingType.Binary
            v = CBool(storedValue)
        Case enuSettingType.DateTime
            v = CDate(storedValue)
        Case enuSettingType.Number
            v = CLng(storedValue)
        Case enuSettingType.real
            v = CDbl(storedValue)
        Case enuSettingType.text
            v = CStr(storedValue)
        Case Else
            v = storedValue
    End Select
    
    GetSettingValue = v

End Function

''' Loads the Settings collection from the XML File
Private Sub LoadSettings()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":LoadSettings"
    
    ' Get the path to the XML file where the Settings are stored
    Dim filePath As String
    filePath = Common.GetAppDataPath & "\Settings.XML"
    
    Dim xmlContent As String
    xmlContent = ReadTextFile(filePath)
    If Len(xmlContent) > 0 Then
        ' A settings file exists and has data
        Dim xmlNodeList As MSXML2.IXMLDOMNodeList
        Dim xmlDoc As MSXML2.DOMDocument60
        Dim xmlNode As MSXML2.IXMLDOMNode
        
        Set xmlDoc = New MSXML2.DOMDocument60
        If Not xmlDoc.LoadXml(xmlContent) Then
            strTrace = "Failed to load the xml content from " & filePath
            GoTo ThrowException
        End If
        
        Dim s As Setting
        Dim nd As IXMLDOMNode
        Dim oNodes As IXMLDOMNodeList
        Set oNodes = xmlDoc.SelectNodes("//Setting")
        If oNodes.Length > 0 Then
        
            innerList.Clear
        
            For i = 0 To oNodes.Length - 1
                Set nd = oNodes.NextNode
                
                Dim settingName As String
                settingName = nd.ChildNodes(0).nodeTypedValue
                
                Dim settingValue As Variant
                settingValue = nd.ChildNodes(1).nodeTypedValue
                
                Dim settingType As enuSettingType
                settingType = nd.ChildNodes(2).nodeTypedValue
                
                Set s = New Setting
                s.Name = settingName
                s.value = GetSettingValue(settingValue, settingType)
                s.ValueType = settingType
                
                innerList.Add s
                
            Next
        End If
        
    Else
        ' Settings file not yet created
        ' Start with an empty collection
    End If
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Private Sub SaveSettings()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SaveSettings"
    
    On Error GoTo ThrowException
    
    ' Get the path to the XML file where the Settings are stored
    Dim filePath As String
    filePath = Common.GetAppDataPath & "\Settings.XML"
    
    ' Create a new DOM
    Dim dom As DOMDocument60
    Set dom = New DOMDocument60
    dom.async = True
    dom.validateOnParse = False
    dom.resolveExternals = False
    dom.preserveWhiteSpace = True
    
    Dim node
    Dim rootNode
    Dim rwNode
    
    Set node = dom.createProcessingInstruction("xml", "version='1.0'")
    dom.appendChild node
    Set rootNode = dom.createElement("Settings")
    dom.appendChild rootNode
    
    If innerList.count > 0 Then
        Dim s As Setting
        For Each s In innerList
            Set node = dom.createElement("Setting")
            
            Set rwNode = dom.createElement("Name")
            rwNode.text = s.Name
            node.appendChild rwNode
            
            Set rwNode = dom.createElement("Value")
            rwNode.text = GetStoredValue(s.value)
            node.appendChild rwNode
            
            Set rwNode = dom.createElement("ValueType")
            rwNode.text = s.ValueType
            node.appendChild rwNode
            
            rootNode.appendChild node

        Next
    End If
    
    Dim strXml As String
    strXml = dom.XML
    
    WriteTextFile filePath, strXml
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

''' Creates/Appends a Text file at the specified filePath with
''' the specified text (fileContent)
Private Sub AppendTextFile(ByVal filePath As String, ByVal fileContent As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":AppendTextFile"
    
    On Error GoTo ThrowException

    Dim TextFile As Integer

    'Determine the next file number available for use by the FileOpen function
    TextFile = FreeFile

    'Open the text file
    Open filePath For Append As TextFile

    'Write the text
    Print #TextFile, fileContent
  
    'Save & Close Text File
    Close TextFile
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
  
End Sub

''' Creates/Overwrites a Text file at the specified filePath with
''' the specified text (fileContent)
Private Sub WriteTextFile(ByVal filePath As String, ByVal fileContent As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":WriteTextFile"
    
    On Error GoTo ThrowException

    Dim TextFile As Integer

    'Determine the next file number available for use by the FileOpen function
    TextFile = FreeFile

    'Open the text file
    Open filePath For Output As TextFile

    'Write the text
    Print #TextFile, fileContent
  
    'Save & Close Text File
    Close TextFile
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
  
End Sub

'''  Reads the text file at the specified filePath
Private Function ReadTextFile(ByVal filePath As String) As String

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ReadTextFile"
    
    On Error GoTo ThrowException

    Dim FileNum As Integer
    Dim DataLine As String
    Dim fileContent As String

    ' Determine next file # available
    FileNum = FreeFile()
    
    ' Open the Text file
    Open filePath For Input As #FileNum

    ' Get the contents of the file
    fileContent = Input(LOF(FileNum), FileNum)

    ' Close the file
    Close FileNum
    
'    ' Or line by line
'    While Not EOF(FileNum)
'        Line Input #FileNum, DataLine ' read in data 1 line at a time
'        ' decide what to do with dataline,
'        ' depending on what processing you need to do for each case
'    Wend
    
    ReadTextFile = fileContent
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    ReadTextFile = ""
    
End Function
