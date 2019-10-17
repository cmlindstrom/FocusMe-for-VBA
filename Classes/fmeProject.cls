VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fmeProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' - Fields

Private Const rootClass As String = "fmeProject"

Private f_Id As String
Private f_Name As String
Private f_Code As String
Private f_description As String

Private f_Active As Boolean
Private f_CombineTitleCode As Boolean

Private f_WinFolder As String
Private f_OutlookFolder As String

Private f_WorkspaceId As String

Private f_Status As Integer
Private f_Priority As Integer
Private f_Color As Integer
Private f_mclId As String

Private f_Created As Date
Private f_Modified As Date
Private f_Deleted As Date

Private Const fields As String = "id,Name,Code,CombineTitleCode,Active,Description,WindowsFolder," & _
                                    "OutlookFolder,WorkspaceId,Status,Priority,Color,mclId,Created," & _
                                    "Modified,Deleted"


' - Properties

''' Unique Identifier for the object
Public Property Let id(ByVal myId As String)
    f_Id = myId
End Property
Public Property Get id() As String
    id = f_Id
End Property

''' Name or Title for the Project
Public Property Let name(ByVal myName As String)
    f_Name = myName
End Property
Public Property Get name() As String
    name = f_Name
End Property

''' Project Code - for organizing / sorting
Public Property Let Code(ByVal myCode As String)
    f_Code = myCode
End Property
Public Property Get Code() As String
    Code = f_Code
End Property

''' Gets / Sets the flag to Combine Code and Name for MCL name
Public Property Let CombineTitleCode(ByVal b As Boolean)
    f_CombineTitleCode = b
End Property
Public Property Get CombineTitleCode() As Boolean
    CombineTitleCode = f_CombineTitleCode
End Property

''' Gets / Sets Active flag, if True project shows in the UI
Public Property Let Active(ByVal b As Boolean)
    f_Active = b
End Property
Public Property Get Active() As Boolean
    Active = f_Active
End Property

''' Gets / Sets the project description
Public Property Let Description(ByVal myDesc As String)
    f_description = myDesc
End Property
Public Property Get Description() As String
    Description = f_description
End Property

''' Gets / Sets the pointer to the Windows Folder
''' Network or local drive location
Public Property Let WindowsFolder(ByVal fldrPath As String)
    f_WinFolder = fldrPath
End Property
Public Property Get WindowsFolder() As String
    WindowsFolder = f_WinFolder
End Property

''' Gets / Sets the pointer to the Outlook Folder
Public Property Let OutlookFolder(ByVal fldrPath As String)
    f_OutlookFolder = fldrPath
End Property
Public Property Get OutlookFolder() As String
    OutlookFolder = f_OutlookFolder
End Property

''' Gets / Sets the pointer to the parent Workspace
Public Property Let WorkspaceId(ByVal id As String)
    f_WorkspaceId = id
End Property
Public Property Get WorkspaceId() As String
    WorkspaceId = f_WorkspaceId
End Property

''' Gets / Sets the Status enumeration index
''' Follows Outlook.OlTaskStatus
Public Property Let Status(ByVal idx As Integer)
    f_Status = idx
End Property
Public Property Get Status() As Integer
    Status = f_Status
End Property

''' Gets / Sets the Priority enumeration index
''' Follows Outlook.olImportance
Public Property Let Priority(ByVal idx As Integer)
    f_Priority = idx
End Property
Public Property Get Priority() As Integer
    Priority = f_Priority
End Property

''' Gets / Sets Color
Public Property Let Color(ByVal idx As Integer)
    f_Color = idx
End Property
Public Property Get Color() As Integer
    Color = f_Color
End Property

''' Gets / Sets Master Category List Id
Public Property Let mclId(ByVal id As String)
    f_mclId = id
End Property
Public Property Get mclId() As String
    mclId = f_mclId
End Property

' - - Basic Properties

''' Date and Time the Item was created the first time
Public Property Let Created(ByVal dte As Date)
    f_Created = dte
End Property
Public Property Get Created() As Date
    Created = f_Created
End Property

''' Date and Time the Item was last modified in the DataStore
Public Property Let Modified(ByVal dte As Date)
    f_Modified = dte
End Property
Public Property Get Modified() As Date
    Modified = f_Modified
End Property

''' Date and Time the Item was marked as deleted in the DataStore
Public Property Let Deleted(ByVal dte As Date)
    f_Deleted = dte
End Property
Public Property Get Deleted() As Date
    Deleted = f_Deleted
End Property

' - Constructor

Private Sub Class_Initialize()

    ' Initialize
    f_Id = "P" & Common.GenerateUniqueID(6)

    ' Defaults
    f_Name = ""
    f_Active = True
    f_CombineTitleCode = True
    f_Color = 0
    
    f_windowsfolder = ""
    f_OutlookFolder = ""
    f_WorkspaceId = ""
    
    f_Created = #1/1/1970#
    f_Modified = #1/1/1970#
    f_Deleted = #1/1/1970#
    
    SetStatus olTaskNotStarted
    SetPriority olImportanceNormal
    
End Sub

' - Methods

''' Returns the preferred representation for the Project's Title
Public Function Subject() As String

    Dim strReturn As String
    strReturn = f_Name
    
    If f_CombineTitleCode Then
        strReturn = f_Code & " - " & f_Name
    End If
    
    Subject = strReturn

End Function

''' Returns True if the Status is equal to olTaskComplete
Public Function Complete() As Boolean
    Complete = (f_Status = Outlook.OlTaskStatus.olTaskComplete)
End Function

''' Sets the Project's parent Workspace
Public Sub SetWorkspace(ByVal w As fmeWorkspace)
    If Not w Is Nothing Then f_WorkspaceId = w.id
End Sub

''' Sets the Project's status using the Outlook enumeration
Public Sub SetStatus(ByVal s As Outlook.OlTaskStatus)
    Dim iVal As Integer
    iVal = s
    f_Status = iVal
End Sub

''' Sets status from a string value
Public Sub SetStatusFromName(ByVal strName As String)

    ' NotStarted = 0
    ' InProgress = 1
    ' Complete = 2
    ' Waiting = 3
    ' Deferred = 4

    Dim iVal As Integer
    iVal = -1
    
    If InStr(LCase(strName), "start") > 0 Then iVal = 0
    If InStr(LCase(strName), "progr") > 0 Then iVal = 1
    If InStr(LCase(strName), "compl") > 0 Then iVal = 2
    If InStr(LCase(strName), "waiti") > 0 Then iVal = 3
    If InStr(LCase(strName), "defer") > 0 Then iVal = 4
    
    f_Status = iVal

End Sub
      
''' Returns the String value for the project's status
Public Function GetStatusName() As String

    Dim strReturn As String
    strReturn = "Not Started"
    
    Select Case f_Status
        Case 0  ' NotStarted
            strReturn = "Not Started"
        Case 1  ' InProgress
            strReturn = "In Progress"
        Case 2  ' Complete
            strReturn = "Complete"
        Case 3  ' Waiting
            strReturn = "Waiting"
        Case 4  ' Deferred
            strReturn = "Deferred"
        Case Else
            strReturn = "Unknown"
    End Select
    
    GetStatusName = strReturn

End Function

''' Sets the Project's Priority from the Outlook enumeration
Public Sub SetPriority(ByVal ip As Outlook.OlImportance)
    Dim iVal As Integer
    iVal = ip
    f_Priority = iVal
End Sub

''' Sets the Project's Priority from a string name
Public Sub SetPriorityFromName(ByVal strName As String)

    Dim iVal As Integer
    iVal = -1
    
    If InStr(LCase(strName), "low") > 0 Then iVal = 0
    If InStr(LCase(strName), "nor") > 0 Then iVal = 1
    If InStr(LCase(strName), "hig") > 0 Then iVal = 2

    f_Priority = iVal

End Sub

''' Returns the name associated with the Priority index value
Public Function GetPriorityName() As String

    Dim strReturn As String
    strReturn = "Normal"
    
    Select Case f_Priority
        Case 0 ' Low
            strReturn = "Low"
        Case 1 ' Normal
            strReturn = "Normal"
        Case 2 ' High
            strReturn = "High"
        Case Else
            strReturn = "Unknown"
    End Select
    
    GetPriorityName = strReturn

End Function

''' Returns the name of the color assigned to the project
Public Function GetColorName() As String
    Dim strReturn As String
    Dim ut As New Utilities
    strReturn = ut.GetCategoryColorName(f_Color)
    Set ut = Nothing
    GetColorName = strReturn
End Function

''' <summary>
''' Returns a random color for the project.
''' </summary>
''' <returns>Integer</returns>
''' <remarks></remarks>
Public Function GetRandomColor() As Integer

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetRandomColor"
    
    On Error GoTo ThrowException

    Dim rnd As New Random
    
    Dim iColor As Integer
    iColor = rnd.Next(1, 24)
    
    GetRandomColor = iColor
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Function


''' <summary>
''' Updates or adds the project's subject to the Master Category List.
''' </summary>
''' <returns>String: Category ID in the MCL</returns>
''' <remarks></remarks>
Public Sub UpdateMCL()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":UpdateMCL_2"
    
    On Error GoTo ThrowException
    
    Dim ut As New Utilities
    
    If Len(f_Name) = 0 Then
        strTrace = "Cannot assign a Master Category without a name."
        GoTo ThrowException
    End If
    
    If Not f_Active Then
        strTrace = "Project '" & f_Name & "' is inactive, skipping request."
        GoTo ThrowException
    End If
    
    Dim bFound As Boolean
    bFound = False
    Dim strMCLID As String

    If Len(f_mclId) = 0 Then
        strTrace = "May be a data issue or adding back an Inactive project."
        strMCLID = ut.GetMCLID(Me.Subject)
        If Len(strMCLID) = 0 Then
            strTrace = "No entry found in the MCL, adding '" & Me.Subject & "'."
            LogMessage strTrace, strRoutine
            
            strMCLID = ut.AddtoMCL(Me.Subject, f_Color)
            If Len(strMCLID) Then
                strTrace = "Error thrown while adding to the MCL for Project: " & Me.Subject
                GoTo ThrowException
            End If
            ' Capture new MCL Id
            f_mclId = strMCLID
            
            GoTo Success
        Else
            f_mclId = strMCLID
            strTrace = "Resync'd MCL with project '" & Me.Subject & "'."
            LogMessage strTrace, strRoutine
        End If
        ' Found an MCLID, so continue on with the routine
    End If
    
    strTrace = "Update MCL with the Project's Subject: " & Me.Subject
    strMCLID = ut.UpdateMCL(f_mclId, Me.Subject, f_Color)
    If Len(strMCLID) Then
        strTrace = "Error thrown while updating the MCL for Project: " & Me.Subject
        GoTo ThrowException
    End If

Success:

    Set ut = Nothing
    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set ut = Nothing

End Sub

' - Common Methods

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

' - Supporting Methods
