Option Explicit

' ## MEMBER DATA
Private rng_TopLeft As Range


' ## INITIALIZATION
Public Sub Initialize(wks_Input As Worksheet)
    Set rng_TopLeft = wks_Input.Range("A3")
End Sub


' ## METHODS - LOOKUP
Public Function Lookup_Generator(str_GeneratorName As String) As IRLegParams
    Dim fld_Output As IRLegParams, strLst_GeneratorNames As Collection, int_FoundRow As Integer, int_ColCtr As Integer
    Set strLst_GeneratorNames = Convert_RangeToList(Gather_RangeBelow(rng_TopLeft))
    int_FoundRow = Examine_FindIndex(strLst_GeneratorNames, str_GeneratorName)

    If int_FoundRow = -1 Then
        Debug.Print "## ERROR - Generator not recognized"
    Else
        With fld_Output
            .Notional = 100
            .FloatEst = True
            .ForceToMV = True
            .IsFwdGeneration = True
            .IsUniformPeriods = False
            .PExch_Start = False
            .PExch_Intermediate = False
            .PExch_End = False
            .RateOrMargin = 0

            int_ColCtr = 1
            .CCY = UCase(rng_TopLeft.Offset(int_FoundRow - 1, int_ColCtr).Value)

            int_ColCtr = int_ColCtr + 1
            .index = UCase(rng_TopLeft.Offset(int_FoundRow - 1, int_ColCtr).Value)

            int_ColCtr = int_ColCtr + 1
            .PmtFreq = UCase(rng_TopLeft.Offset(int_FoundRow - 1, int_ColCtr).Value)

            int_ColCtr = int_ColCtr + 1
            .Daycount = UCase(rng_TopLeft.Offset(int_FoundRow - 1, int_ColCtr).Value)

            int_ColCtr = int_ColCtr + 1
            .BDC = UCase(rng_TopLeft.Offset(int_FoundRow - 1, int_ColCtr).Value)

            int_ColCtr = int_ColCtr + 1
            .EOM = rng_TopLeft.Offset(int_FoundRow - 1, int_ColCtr).Value

            int_ColCtr = int_ColCtr + 1
            .PmtCal = UCase(rng_TopLeft.Offset(int_FoundRow - 1, int_ColCtr).Value)

            int_ColCtr = int_ColCtr + 1
            .estcal = UCase(rng_TopLeft.Offset(int_FoundRow - 1, int_ColCtr).Value)

            int_ColCtr = int_ColCtr + 1
            .Curve_Disc = rng_TopLeft.Offset(int_FoundRow - 1, int_ColCtr).Value

            int_ColCtr = int_ColCtr + 1
            .Curve_Est = rng_TopLeft.Offset(int_FoundRow - 1, int_ColCtr).Value
        End With
    End If

    Lookup_Generator = fld_Output
End Function