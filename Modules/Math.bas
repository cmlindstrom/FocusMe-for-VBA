Attribute VB_Name = "Math"
' - Fields

Private Const rootClass As String = "Math"

Private Const bayesLearningConstant As Double = 0.1

' Functions

''' <summary>
''' Function used to calculate a Bayesian average statistic.
''' </summary>
''' <param name="dblNewMeasure">New observed value</param>
''' <param name="dblPreviousMeasure">Previous Average</param>
''' <param name="lSampleCount"># of observations including this one</param>
''' <returns>Double</returns>
''' <remarks>This function is useful in tracking a population's mean when all observations are not stored or collected.
''' e.g. Provides a learning algorithm in calculating how long something takes when recording the observed elapsed time for every
''' sample, such as copying a file - can be very useful in predicting how long the entire file might take to copy in it's
''' entirety.</remarks>
Public Function AverageBayes(ByVal dblNewMeasure As Double, _
                         ByVal dblPreviousMeasure As Double, _
                         ByVal lSampleCount As Long) As Double
                         
    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":AverageBayes"
    
    On Error GoTo ThrowException
    
    Dim dblReturn As Double
    dblReturn = 0#

    Dim dbLearningConstant As Double
    dbLearningConstant = bayesLearningConstant ' Property initialized to 0.1

    Dim dbl As Double
    Dim dbLFactor As Double
    Dim dbSFactor As Double

    dbl = 1# / (dbLearningConstant + lSampleCount)
    dbLFactor = dbl / (dbl + 1)
    dbSFactor = 1 / (dbl + 1)

    dblReturn = (dbLFactor * dblNewMeasure) + (dbSFactor * dblPreviousMeasure)
    
    AverageBayes = dblReturn
    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    AverageBayes = -1
                         
End Function

''' Rounds the source real number to the largest number based
''' on the designated increment value (in partial hours)
Public Function RoundHours(ByVal inHrs As Double, ByVal dblInc As Double)

    Dim strTrace As String
    Dim strRoutine As String
    strRoutine = rootClass & ":RoundHours"
    
    On Error GoTo ThrowException

    Dim whole_hr As Double
    Dim decimal_hr As Double
    Dim dblRemainder As Double

    Dim dblRet As Double
    dblRet = 0#

    strTrace = "Incoming number: '" & FormatNumber(inHrs, 3) & "'."
    LogMessage strTrace, strRoutine

    whole_hr = CInt(inHrs) ' Int(inHrs)
    decimal_hr = inHrs - whole_hr

    strTrace = "Calc remainder with: decimal_hr=" & decimal_hr & ", dblInc=" & dblInc
    dblRemainder = XLMod(decimal_hr, dblInc) 'decimal_hr Mod dblInc
    
    strTrace = strTrace & " | Remainder=" & dblRemainder
'    LogMessage strTrace, strRoutine

    ' Always round up...
    If dblRemainder > 0 Then
        decimal_hr = decimal_hr + dblInc - dblRemainder
    End If

    dblRet = whole_hr + decimal_hr

    RoundHours = dblRet

    Exit Function

ThrowException:
    LogMessageEx strTrace, err, strRoutine
    RoundHours = -1

End Function

Function XLMod(a, b) As Double
    ' This replicates the Excel MOD function
    XLMod = a - b * Int(a / b)
End Function
