Option Explicit

' ## MEMBER DATA
Private rng_TopLeft As Range

' ## INITIALIZATION
Public Sub Initialize(wks_Input As Worksheet)
    Set rng_TopLeft = wks_Input.Range("A3")
End Sub

' ## METHODS - LOOKUP
Public Function Lookup_YGenerator(str_YGeneratorName As String) As BondParams
    Dim fld_Output As BondParams, strLst_GeneratorNames As Collection, int_FoundRow As Integer, int_ColCtr As Integer
    Set strLst_GeneratorNames = Convert_RangeToList(Gather_RangeBelow(rng_TopLeft))
    int_FoundRow = Examine_FindIndex(strLst_GeneratorNames, str_YGeneratorName)

    If int_FoundRow = -1 Then
        Debug.Print "## ERROR - Generator not recognized"
    Else
        With fld_Output
            int_ColCtr = 1
            .YieldCalc = rng_TopLeft.Offset(int_FoundRow - 1, int_ColCtr).Value

            int_ColCtr = int_ColCtr + 1
            .RateComputingMode = rng_TopLeft.Offset(int_FoundRow - 1, int_ColCtr).Value

            int_ColCtr = int_ColCtr + 1
            .DaycountConv = rng_TopLeft.Offset(int_FoundRow - 1, int_ColCtr).Value

            int_ColCtr = int_ColCtr + 1
            .Periodicity = rng_TopLeft.Offset(int_FoundRow - 1, int_ColCtr).Value

            int_ColCtr = int_ColCtr + 1
            .FixingCurve = rng_TopLeft.Offset(int_FoundRow - 1, int_ColCtr).Value

            int_ColCtr = int_ColCtr + 1
            .Fix_AI = rng_TopLeft.Offset(int_FoundRow - 1, int_ColCtr).Value

            int_ColCtr = int_ColCtr + 1
            .Fix_IT = rng_TopLeft.Offset(int_FoundRow - 1, int_ColCtr).Value

            int_ColCtr = int_ColCtr + 1
            .YieldSchedule = rng_TopLeft.Offset(int_FoundRow - 1, int_ColCtr).Value

        End With
    End If

    Lookup_YGenerator = fld_Output
End Function