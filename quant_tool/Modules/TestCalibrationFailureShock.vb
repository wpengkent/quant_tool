Sub testCalibrationFailuresShock()
Dim t As CTimer: Set t = New CTimer: t.StartCounter
'Dim ws0 As Worksheet: Set ws0 = Worksheets("General Settings")
'Dim systemDateRange As Range: Set systemDateRange = ws0.Range("B10:C12") 'all system dates base and current etc

Dim ws1 As Worksheet: Set ws1 = Worksheets("Calibration Testing SHOCK")
Dim dateRange As Range: Set dateRange = ws1.Range("A2")
Dim intDates As Double: intDates = Examine_NumRows(dateRange)
Dim newSystemDate As Date: Dim resultsRangeStart As Range: Set resultsRangeStart = ws1.Range("F1")
Dim failurePointsRange As Range: Set failurePointsRange = ws1.Range("K1")
Dim outputAllVolsBool As Boolean: outputAllVolsBool = ws1.Range("C6").Value
Dim outputCapVolsBool As Boolean: outputCapVolsBool = ws1.Range("C10").Value
Dim intResults As Double: intResults = Examine_NumRows(dateRange.Offset(0, 5))
Dim resultsRange0 As Range

If intResults = 0 Then
Else
Set resultsRange0 = Range(dateRange.Offset(0, 5).Address & ":" & dateRange.Offset(intResults, 12).Address)
resultsRange0.ClearContents
End If

Dim ws0 As Worksheet: Set ws0 = Worksheets("CONT_DL~")

Dim ws2 As Worksheet:  Set ws2 = Worksheets("Setup - CVL")
Dim CVLStartRange As Range: Set CVLStartRange = ws2.Range("A4")
Dim numCVL As Double: numCVL = Examine_NumRows(CVLStartRange)
Dim i As Double: Dim j As Double


'im cvlCodeNames As Dictionary: Set cvlCodeNames = New Dictionary
'im cvlRowNumbers As Dictionary: Set cvlRowNumbers = New Dictionary
'vlCodeNames.CompareMode = CompareMethod.TextCompare
'vlRowNumbers.CompareMode = CompareMethod.TextCompare

Dim cvlCodeNames As Collection: Set cvlCodeNames = New Collection
Dim cvlRowNumbers As Collection: Set cvlRowNumbers = New Collection

Dim str_input As Integer
'make a dictionary to collate all "YES"
Dim tempString As String
For i = 1 To numCVL
If CVLStartRange.Offset(i - 1, 1) = "YES" Then
tempString = CVLStartRange.Offset(i - 1, 0)
cvlCodeNames.Add tempString
cvlRowNumbers.Add (i)
End If

Next i


ws0.Range("O1").Value = ws1.Range("C9").Value   'container

Dim j2 As Double 'for counting calibration failures on a specific strike
Dim j3 As Double 'in case capvols are outputted
'all GenerateSelectedIRCurves

For i = 1 To intDates
newSystemDate = dateRange.Offset(i - 1, 0).Value
'systemDateRange.Value = newSystemDate

ws0.Range("O2").Value = newSystemDate
ws0.Range("O3").Value = newSystemDate

Call DownloadToContSheet
     str_input = (newSystemDate)

   ' If IsNumeric(str_Input) Then
        Dim sct_Scenarios As ScenContainer: Set sct_Scenarios = GetObject_ScenContainer("DL~") 'GetObject_ScenContainer(cbx_Container.Value)
        Call sct_Scenarios.LoadScenario((str_input))
   ' Else
    '    MsgBox "Enter a valid scenario number"
   ' End If

    'tbx_ScenNum.SetFocus
'Call GenerateSelectedIRCurves: Call GenerateSelectedCapVolCurves   'only generates with a "YES"

'record results in testing worksheet
Dim wsTest As Worksheet: Dim calibFailRange As Range
Dim calibfailRangeStart As Range: Dim calibRangeEnd As Double:
Dim capVolsRangeStart As Range: Dim capVolsRangeEnd As Double '17032015
Dim numFailures As Double


    'CHECK FAILURES IN EACH WORKSHEET
    For j = 1 To cvlCodeNames.count
    Set wsTest = Worksheets("CVL_" & (cvlCodeNames.item(j)))
    Set calibfailRangeStart = wsTest.Range("S7")
    calibRangeEnd = Examine_NumRows(calibfailRangeStart)
   Set calibFailRange = wsTest.Range("S7:S" & (6 + calibRangeEnd))

   Set capVolsRangeStart = wsTest.Range("P7")
   capVolsRangeEnd = Examine_NumRows(capVolsRangeStart)

    numFailures = WorksheetFunction.CountIf(calibFailRange, False)
    resultsRangeStart.Offset((i - 1) * cvlCodeNames.count + j, 0).Value = newSystemDate

    resultsRangeStart.Offset((i - 1) * cvlCodeNames.count + j, 1).Value = cvlCodeNames.item(j)
    resultsRangeStart.Offset((i - 1) * cvlCodeNames.count + j, 2).Value = wsTest.Range("M2")
    resultsRangeStart.Offset((i - 1) * cvlCodeNames.count + j, 4).Value = numFailures
    ' resultsRangeStart = ws1.Range("F1")
    Dim tempStringFailures As String: tempStringFailures = ""
    Dim tempStringAllVols As String: tempStringAllVols = ""
    Dim tempStringCapVols As String: tempStringCapVols = ""   '17032015
    Dim calibFailuresAll As Collection: Set calibFailuresAll = New Collection
    Dim capVolsAll As Collection: Set capVolsAll = New Collection '17032015
    If numFailures = 0 Then
    'no need to do anything, no failures
    tempStringFailures = ""
        For j2 = 1 To calibRangeEnd
        'f wsTest.Range("S6").Offset(j2, 0) = False Then
        'alibFailuresAll.Add j2
        If outputAllVolsBool = True Then
        tempStringAllVols = tempStringAllVols & wsTest.Range("R6").Offset(j2, 0).Value & "|"
        End If

        Next j2
    Else
        For j2 = 1 To calibRangeEnd
        If wsTest.Range("S6").Offset(j2, 0) = False Then
        calibFailuresAll.Add j2
        tempStringFailures = tempStringFailures & j2 & "|"
        End If
        If outputAllVolsBool = True Then
        tempStringAllVols = tempStringAllVols & wsTest.Range("R6").Offset(j2, 0).Value & "|"
        'tempStringCapVols = tempStringAllVols & wsTest.Range("R6").Offset(j2, 0).Value & "|"
        End If
        Next j2

    End If

    If outputCapVolsBool = True Then
    For j3 = 1 To capVolsRangeEnd
    tempStringCapVols = tempStringCapVols & wsTest.Range("P6").Offset(j3, 0).Value & "|"
    Next j3
    End If

    resultsRangeStart.Offset((i - 1) * cvlCodeNames.count + j, 5).Value = tempStringFailures
    resultsRangeStart.Offset((i - 1) * cvlCodeNames.count + j, 6).Value = tempStringAllVols
    resultsRangeStart.Offset((i - 1) * cvlCodeNames.count + j, 7).Value = tempStringCapVols
    Next j
Next i
 intResults = Examine_NumRows(dateRange.Offset(0, 5)) - 1
 Dim boolRange As Range
 Set boolRange = Range(dateRange.Offset(0, 8).Address & ":" & dateRange.Offset(intResults, 8).Address)
 boolRange.FormulaR1C1 = "=IF(RC[1]=0,FALSE,TRUE)"

 Call refreshResults(ws1.Name)

ws1.Range("C7") = t.TimeElapsed

End Sub

Public Sub refreshResults(wsname As String)
Dim wsRes As Worksheet: Set wsRes = Worksheets(wsname)
Dim tableRangeStrike As Range: Set tableRangeStrike = wsRes.Range("O1")
Dim tableRangeStrikeEnd As Double
'refresh by Strike
tableRangeStrikeEnd = Examine_NumRows(tableRangeStrike) - 2
Dim i As Double:
For i = 1 To tableRangeStrikeEnd
    tableRangeStrike.Offset(i, 0).Select
    ActiveCell.FormulaR1C1 = "=countfailuresbystrike(RC[-1],R1C7)"
Next i
'refresh dates/Scenarios
End Sub