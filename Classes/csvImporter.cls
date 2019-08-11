VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "csvImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Fields

Private Const rootClass As String = "csvImporter"

Private txtQualifier As String
Private txtDelimiter As String
Private inclHeaders As Boolean
Private ilImported As Integer

Private lineChar As String

Private dtResult As DataTable

' Events

Public Event ImportComplete()

' Properties

Public Enum enuFormat
    None = 0
    Text = 1
    File = 2
End Enum

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

''' Flag to look for Headers in the first row of
''' the imported text
Public Property Let IncludesHeaders(ByVal b As Boolean)
    inclHeaders = b
End Property
Public Property Get IncludesHeaders() As Boolean
    IncludesHeaders = inclHeaders
End Property

''' # of lines imported from the last import call.
Public Property Get LinesImported() As Integer
    LinesImported = ilImported
End Property

''' Last imported table
Public Property Get result() As DataTable
    Set result = dtResult
End Property

' Constructor

Private Sub Class_Initialize()
    txtDelimiter = ","
    txtQualifier = """"
    lineChar = vbCrLf
    inclHeaders = True
End Sub

' Methods

''' <summary>
''' Imports the specified string into a DataTable.
''' </summary>
''' <param name="strInput">String: text to be imported.</param>
''' <param name="fmt">Import.enuFormat:</param>
''' <param name="FirstColumnHeaderLabel" (Optional, defaults to empty)>String: First header label</param>
''' <param name="valueDelimiter" (Optional, defaults to empty)>String: is the character used between
''' to differentiate values within the data row</param>
''' <param name="qualifier">String: the character used to frame each value in a data row</param>
''' <param name="lneDelimiter">String: is the character used to differentiate the end of a row/line</param>
''' <returns>DataTable:</returns>
''' <remarks>If a first column header is designated, the import doesn't
''' start until that value appears in a line read from the text file,
''' otherwise the import assumes the first row is the header row.</remarks>
Public Function Import(ByVal strInput As String, _
                       ByVal fmt As enuFormat, _
              Optional ByVal FirstColumnHeaderLabel As String = "", _
              Optional ByVal valueDelimiter As String = ",", _
              Optional ByVal qualifier As String = "", _
              Optional ByVal lneDelimiter As String = "") As DataTable
              
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":Import"
              
    If Len(valueDelimiter) > 0 Then
        txtDelimiter = valueDelimiter
    End If
    If Len(qualifier) > 0 Then
        txtQualifier = qualifier
    End If
    If Len(lneDelimiter) > 0 Then
        lineChar = lneDelimiter
    End If
    
    Dim dt As DataTable
    
    Select Case fmt
        Case enuFormat.Text
            Set dt = ImportText(strInput, FirstColumnHeaderLabel)
        Case enuFormat.File
            Set dt = ImportFile(strInput, FirstColumnHeaderLabel)
        Case Else
            strTrace = "Unrecognized input format: " & fmt
            GoTo ThrowException
    End Select
    
    Set Import = dt
    Exit Function
    
ThrowException:
    LogMessage strTrace, strRoutine
    Set Import = Nothing

End Function

''' <summary>
''' Imports the specified string into a DataTable.
''' </summary>
''' <param name="strInput">String: text to be imported.</param>
''' <param name="FirstColumnHeaderLabel">String: First header label</param>
''' <returns>DataTable:</returns>
''' <remarks>If a first column header is designated, the import doesn't
''' start until that value appears in a line read from the text file,
''' otherwise the import assumes the first row is the header row.</remarks>
Public Function ImportText(ByVal strInput As String, _
              Optional ByVal FirstColumnHeaderLabel As String = "") As DataTable

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ImportText"
    
    If Len(strInput) = 0 Then
        strTrace = "An empty input encountered."
        GoTo ThrowException
    End If
    
    On Error GoTo ThrowException
    
    Dim bHeaderLabelSpecified As Boolean
    Dim bHeaderFound As Boolean
    bHeaderLabelSpecified = Not Len(FirstColumnHeaderLabel) = 0
  '  If bheaderlablespecified Then bHeaderFound = Contains(strInput, FirstColumnHeaderLabel)
    
    Dim tblName As String
    tblName = "Imported Text"
    
    Set dtResult = New DataTable
    dtResult.Name = tblName

    Dim iLines As Integer
    iLines = 0
    
    strTrace = "Get input lines."
    Dim lines As ArrayList
    Set lines = SplitLines(strInput)
    If lines.Count = 0 Then
        strTrace = "Found zero lines to parse in: " & strInput
        GoTo ThrowException
    End If
    
    Dim strLine As Variant
    For Each strLine In lines
    
        iLines = iLines + 1
    
        If Len(strLine) = 0 Then
            strTrace = "WARNING: Encountered empty row at line: " & iLines
            LogMessage strTrace, strRoutine
        End If
        
        ' Check for Header
        If Not bHeaderFound Then
            If bHeaderLabelSpecified Then
                Dim myHeaders() As String
                myHeaders = SplitLineValues(strLine)
                If UBound(myHeaders) = 0 Then
                    strTrace = "Failed to parse the line - try the next line."
                    LogMessage strTrace, strRoutine
                    GoTo nextLine
                End If

                Dim strFirstHeader As String
                strFirstHeader = myHeaders(0)
                If Contains(LCase(FirstColumnHeaderLabel), LCase(strFirstHeader)) Then
                    ' Found the header line
                    CreateColumns strLine
                    bHeaderFound = True
                    GoTo nextLine
                End If
                
            Else
                If iLines = 1 Then
                    ' Create a default header
                    Dim strHeader As String
                    strHeader = CreateDefaultHeaderLine(strLine)
                    CreateColumns strHeader
                    bHeaderFound = True
                    ' GoTo nextLine
                End If
            End If
        End If
        
        ' Create each row after the Header has been created
        If bHeaderFound Then CreateRow strLine
    
nextLine:
    Next
    
    ' Capture imported line count
    ilImported = iLines
    
    Set ImportText = dtResult
    Exit Function
    
ThrowException:
    LogMessage strTrace, strRoutine

End Function

''' <summary>
''' Imports the specified file into a DataTable.
''' </summary>
''' <param name="fullFilePath">String: full file path to a text file to be imported.</param>
''' <param name="FirstColumnHeaderLabel">String: First header label</param>
''' <returns>System.Data.DataTable:</returns>
''' <remarks>If a first column header is designated, the import doesn't
''' start until that value appears in a line read from the text file,
''' otherwise the import assumes the first row is the header row.</remarks>
Public Function ImportFile(ByVal fullFilePath As String, _
                  Optional ByVal FirstColumnHeaderLabel As String = "") As DataTable
                                
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":ImportFile"
    
    
    
    Exit Function
    
ThrowException:
    LogMessage strTrace, strRoutine
    
End Function

' Supporting Methods

Private Sub CreateColumns(ByVal strLine As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":CreateColumns"
    
    On Error GoTo ThrowException
    
    Dim myValues() As String
    myValues = SplitLineValues(strLine)
    For i = LBound(myValues) To UBound(myValues)
        Dim c As DataColumn
        Set c = dtResult.Columns.Add(myValues(i))
    Next
    
    Exit Sub
    
ThrowException:
    LogMessageEx strTrace, err, strRoutine
    
End Sub

Private Function CreateDefaultHeaderLine(ByVal strLine As String) As String

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SplitLineValues"
    
    On Error GoTo ThrowException
    
    Dim strReturn As String
    strReturn = ""

    Dim myValues() As String
    myValues = SplitLineValues(strLine)
    For i = LBound(myValues) To UBound(myValues)
        If i = 0 Then
            strReturn = strReturn & "Header_" & i
        Else
            strReturn = strReturn & ",Header_" & i
        End If
    Next

    CreateDefaultHeaderLine = strReturn
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    CreateDefaultHeaderLine = ""

End Function

Private Sub CreateRow(ByVal strRow As String)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":CreateRow"

    If Len(strRow) = 0 Then
        strTrace = "An empty or null row string encountered."
        GoTo ThrowException
    End If
    
    On Error GoTo ThrowException
    
    strTrace = "Parsing the row: " & strRow
    Dim myValues() As String
    myValues = SplitLineValues(strRow)
    If UBound(myValues) = 0 Then
        strTrace = "Failed to parse a row string: " & strRow
        GoTo ThrowException
    End If
    
    Dim dr As DataRow
    Set dr = dtResult.Rows.Add(myValues)

    Exit Sub

ThrowException:
    LogMessageEx strTrace, err, strRoutine

End Sub

Private Function SplitLines(ByVal strIn As String) As ArrayList

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SplitLineValues"
    
    On Error GoTo ThrowException
    
    Dim retList As New ArrayList
    
    strTrace = "Parsing input string: " & strIn
    Dim myStrings() As String
    myStrings = Split(strIn, lineChar)
    
    Dim s As Variant
    For Each s In myStrings
        s = Replace(s, Chr(10), "")
        retList.Add s
    Next
    
    Set SplitLines = retList
    Exit Function
    
ThrowException:
    LogMessage strTrace, strRoutine
    LogMessageEx strTrace, err

End Function

Private Function SplitLineValues(ByVal strLine As String) As String()

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":SplitLineValues"
    
    On Error GoTo ThrowException
    
    Dim retList() As String
    
    strTrace = "Parsing input string: " & strLine
    Dim rL As New ArrayList
    Dim myStrings() As String
    myStrings = Split(strLine, txtDelimiter)
    Dim s As Variant
    For Each s In myStrings
        rL.Add s
    Next

    strTrace = "Copying parsed string to output variable, column count: " & rL.Count
    ReDim retList(rL.Count)
    
    For i = 0 To rL.Count - 1
        retList(i) = rL(i)
    Next
    
    SplitLineValues = retList
    Exit Function
    
ThrowException:
    LogMessage strTrace, strRoutine
    LogMessageEx strTrace, err

End Function
