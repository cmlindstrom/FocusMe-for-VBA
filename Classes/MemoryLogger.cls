VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MemoryLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' - Fields

    Private Const rootClass As String = "MemoryLogger"
    
    Private Const defaultMax As Integer = 250
    
    Dim iTrimCount As Integer

    Dim f_queue As ArrayList
    Dim f_max As Integer
   
' - Events

    Public Event NewMessage(ByVal id As Long, ByVal msg As String)
    Public Event QueueCleared()

' - Properties

''' Returns the contents of the Message Queue
Public Property Get Messages() As ArrayList
    Set Messages = f_queue
End Property

''' Max Number of records to keep in the memory queue
''' Defaults to 250
Public Property Let Max(ByVal cnt As Integer)
    f_max = cnt
End Property
Public Property Get Max() As Integer
    Max = f_max
End Property

' - Constructor

Private Sub Class_Initialize()
    Set f_queue = New ArrayList
    f_max = defaultMax
    iTrimCount = 0
End Sub

Private Sub Class_Terminate()
    Set f_queue = Nothing
End Sub

' - Methods

''' Clear the Message Queue
Public Sub Clear()

    ' Clear the current ArrayList
    f_queue.Clear
    ' Inform listeners
    RaiseEvent QueueCleared
    
End Sub

''' Adds a message to the queue
Public Sub Post(ByVal channelId As Long, ByVal msg As String)

    ' Add msg to queue
    f_queue.Add channelId & "|" & msg
    ' Inform listeners
    RaiseEvent NewMessage(channelId, msg)
    
    ' Check if queue needs to be Trimmed
    If f_queue.Count > f_max Then Trim f_max
    
End Sub

''' Reduces to the last n messages
Public Sub Trim(Optional n As Integer = 10)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Trim"

    On Error GoTo ThrowException
    
    Dim strtCount As Integer
    strtCount = f_queue.Count
    
    Dim i As Integer
    Dim k As Integer
    If strtCount > n Then
        k = strtCount - 1 - n
        For i = k To 0 Step -1
            f_queue.RemoveAt i
        Next
        
        iTrimCount = iTrimCount + 1
        Dim re As Integer
        re = iTrimCount Mod 5000
        If re = 0 Then
            strTrace = "INFO: Removed " & k + 1 & " messages, (old: " & strtCount & _
                        ") from the queue; (new: " & f_queue.Count & ")."
            LogMessage strTrace, strRoutine
        End If
    Else
        strTrace = "WARNING: Queue trim unnecessary, current length = " & f_queue.Count
        LogMessage strTrace, strRoutine
    End If


    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

