VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Contacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const rootClass As String = "Contacts"

Dim sortOn As enuSortOn
Dim sortDir As enuSortDirection
Dim myItems As ArrayList
Dim sItems As ArrayList

' - - Properties

''' Latest collection of Contacts
Public Property Let Items(ByVal ar As ArrayList)
    Set myItems = ar
End Property
Public Property Get Items() As ArrayList
    Set Items = myItems
End Property

''' Sort Property ' LastName and Name
Property Let SortBy(ByVal srt As enuSortOn)
    sortOn = srt
End Property
Property Get SortBy() As enuSortOn
    SortBy = sortOn
End Property

' - - Constructor

Private Sub Class_Initialize()
    Set myItems = New ArrayList
    Set sItems = New ArrayList
    sortOn = LastName
    sortDir = Ascending
End Sub

' - - Methods

Public Sub AddItem(ByVal c As Outlook.ContactItem)
    myItems.Add c
    InternalSort sortOn, sortDir
End Sub

Public Sub RemoveItem(ByVal c As Outlook.ContactItem)
    myItems.Remove c
End Sub

Public Sub RemoveAt(ByVal idx As Integer)
    myItems.RemoveAt idx
End Sub

Public Function Item(ByVal idx As Integer) As Outlook.ContactItem
    Set Item = myItems(idx)
End Function

Public Function Sort(ByVal srt As enuSortOn, Optional ByVal dir As enuSortDirection) As ArrayList
    If Not IsMissing(dir) Then sortDir = dir
    InternalSort srt, sortDir
End Function

Public Function Count() As Integer
    Count = myItems.Count
End Function

Public Sub Clear()
    myItems.Clear
End Sub

''' Returns True if the specified item is in the collection
Public Function IsItemPresent(ByVal eid As String) As Boolean

    Dim bReturn As Boolean
    bReturn = False

    Dim c As Outlook.ContactItem
    Set c = GetItemById(eid)
    If Not IsNothing(c) Then
        bReturn = True
    Else
        bReturn = False
    End If
    
    IsItemPresent = bReturn

End Function


''' Returns a Contact from the collection matching the specified id
Public Function GetItemById(ByVal id As String) As Outlook.ContactItem

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":GetItemById"
    
    On Error GoTo ThrowException
    
    Dim retItem As Outlook.ContactItem
    Set retItem = Nothing
    
    Dim bFnd As Boolean
    bFnd = False
    Dim c As Outlook.ContactItem
    For Each c In myItems
        If LCase(c.EntryId) = LCase(id) Then
            Set retItem = c
            bFnd = True
            Exit For
        End If
    Next
    
    If Not bFnd Then
        strTrace = "Failed to find a contact in the collection, id: " & id
        LogMessage "WARNING: " & strTrace, strRoutine
    End If
    
    Set GetItemById = retItem
    
    Exit Function
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    Set GetItemById = Nothing

End Function


' - - Supporting Methods

Private Sub InternalSort(ByVal srt As enuSortOn, ByVal dir As enuSortDirection)

    Dim strTrace As String
    strTrace = "General Fault."
    Dim strRoutine As String
    strRoutine = rootClass & ":InternalSort"
    
    ' Set sList = myItems
    
    Dim c As Outlook.ContactItem
    Dim thisValue As String
    Dim thisItem As Outlook.ContactItem
    Dim nextValue As String
    Dim nextItem As Outlook.ContactItem
    Dim bPassDirty As Boolean
    bPassDirty = True
    
    Do While bPassDirty
    
        ' Assume clean pass
        bPassDirty = False
    
        ' Evalute items
        Dim i As Integer
        For i = 0 To myItems.Count - 1
            ' Get current value
            Set c = myItems(i)
            thisValue = GetIndexValue(c)
            Set thisItem = c
            ' Get Next Value
            If Not i = myItems.Count - 1 Then
                Set c = myItems(i + 1)
                nextValue = GetIndexValue(c)
                Set nextItem = c
            Else
                ' Done with the pass
                GoTo PassComplete
            End If
        
            ' Compare
            Dim bSwap As Boolean
            bSwap = False
            If sortDir = Ascending Then bSwap = nextValue < thisValue
            If sortDir = Descending Then bSwap = nextValue > thisValue
        
            ' Swap necessary?
            If bSwap Then
                ' Code Check
                Dim iCurrent As Integer
                iCurrent = myItems.Count
        
                ' Swapped two items, mark the pass as dirty
                ' - Remove nextItem
                myItems.Remove nextItem
                ' - Insert nextItem before thisItem
                myItems.Insert i, nextItem
                
                ' Mark the pass as dirty
                bPassDirty = True
                
                ' Code Check
                If Not iCurrent = myItems.Count Then
                    strTrace = "Elongating the array."
                End If
            End If
        
        Next
PassComplete:
    
    Loop
  
    Exit Sub
    
ThrowException:
    LogMessage strTrace, strRoutine

End Sub
Private Function GetIndexValue(ByVal c As Outlook.ContactItem) As String
    Dim strReturn As String
    
    Select Case sortOn
        Case enuSortOn.CreatedDate
            strReturn = Format(c.CreationTime, "yyyymmdd")
        Case enuSortOn.Subject
            strReturn = c.Subject
        Case enuSortOn.Name
            strReturn = c.FullName
        Case enuSortOn.LastName
            strReturn = c.LastName
    End Select
    
    GetIndexValue = strReturn
End Function

