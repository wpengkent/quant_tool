Option Explicit

Public Function Calc_WeekDaysBetween(lng_StartDate As Long, lng_EndDate As Long) As Long
    ' ## Returns the number of weekdays to be added to the start date to obtain the end date

    Dim int_StartDOW As Integer: int_StartDOW = WorksheetFunction.Weekday(lng_StartDate, 2)
    Dim int_EndDOW As Integer: int_EndDOW = WorksheetFunction.Weekday(lng_EndDate, 2)
    Dim lng_NumWholeWeeks As Long: lng_NumWholeWeeks = (Date_PrevDayOfWeek(lng_EndDate, int_StartDOW) - lng_StartDate) / 7
    Dim int_Remainder As Integer

    ' Handle inputs which include weekends
    Dim bln_BothWeekends As Boolean
    If int_StartDOW >= 6 And int_EndDOW >= 6 Then bln_BothWeekends = True
    If int_StartDOW >= 6 Then int_StartDOW = 1  ' Treat starting on Sat or Sun the same as starting on Monday
    If int_EndDOW >= 6 Then int_EndDOW = 5  ' Treat ending on Sat or Sun the same as ending on Friday

    ' Determine the remainder
    If bln_BothWeekends = True Then
        int_Remainder = 0
    Else
        If int_StartDOW <= int_EndDOW Then
            int_Remainder = int_EndDOW - int_StartDOW
        Else
            int_Remainder = 5 - (int_StartDOW - int_EndDOW)
        End If
    End If

    Calc_WeekDaysBetween = 5 * lng_NumWholeWeeks + int_Remainder
End Function

Public Function Calc_SortKey(arr_Input As Variant, Optional bln_InputIsText As Boolean = False) As Integer()
    ' ## Return array containing indexes of each original element if the input array were sorted in ascending order

    Dim int_LowerBound As Integer: int_LowerBound = LBound(arr_Input)
    Dim int_UpperBound As Integer: int_UpperBound = UBound(arr_Input)
    Dim intArr_Output() As Integer: ReDim intArr_Output(int_LowerBound To int_UpperBound) As Integer
    Dim int_KeyCtr As Integer
    Dim var_ActiveVal As Variant
    Dim int_ActiveNumBelow As Integer
    Dim int_Ctr As Integer

    ' Convert to upper case so sorting is not case sensitive
    Dim arr_InputToUse() As Variant: ReDim arr_InputToUse(int_LowerBound To int_UpperBound) As Variant
    For int_Ctr = int_LowerBound To int_UpperBound
        If bln_InputIsText = True Then
            arr_InputToUse(int_Ctr) = UCase(arr_Input(int_Ctr))
        Else
            arr_InputToUse(int_Ctr) = arr_Input(int_Ctr)
        End If
    Next int_Ctr

    For int_Ctr = int_LowerBound To int_UpperBound
        var_ActiveVal = arr_InputToUse(int_Ctr)
        int_ActiveNumBelow = 0

        ' Adjust ordering of other elements in relation to the active value
        For int_KeyCtr = int_LowerBound To int_Ctr - 1
            If var_ActiveVal < arr_InputToUse(int_KeyCtr) Then
                intArr_Output(int_KeyCtr) = intArr_Output(int_KeyCtr) + 1
            Else
                int_ActiveNumBelow = int_ActiveNumBelow + 1
            End If
        Next int_KeyCtr

        ' Set the position of the active value
        intArr_Output(int_Ctr) = int_ActiveNumBelow + int_LowerBound
    Next int_Ctr

    Calc_SortKey = intArr_Output
End Function

Public Function Calc_SortedArray(dblArr_OrigData() As Double, intArr_SortKey() As Integer) As Double()
    ' ## Order output array in a way not related to the data itself, but according to the order specified by the sort key
    Dim int_LowerBound As Integer: int_LowerBound = LBound(dblArr_OrigData)
    Dim int_UpperBound As Integer: int_UpperBound = UBound(dblArr_OrigData)
    Dim dblArr_Output() As Double: ReDim dblArr_Output(int_LowerBound To int_UpperBound) As Double
    Dim int_Ctr As Integer

    For int_Ctr = int_LowerBound To int_UpperBound
        dblArr_Output(intArr_SortKey(int_Ctr)) = dblArr_OrigData(int_Ctr)
    Next int_Ctr

    Calc_SortedArray = dblArr_Output
End Function

Public Function Calc_NumMonths(str_Term As String) As Integer
    ' ## Calculate number of months in the specified term, e.g. "3Y" = 36, "6M" = 6, "2Y,6M" = 30
    Dim arr_Split As Variant: arr_Split = Convert_Split(str_Term, ",")
    Dim var_ActiveTerm As Variant, str_ActiveTerm As String
    Dim int_ActiveQty As Integer, str_ActiveType As String
    Dim int_Output As Integer

    For Each var_ActiveTerm In arr_Split
        str_ActiveTerm = CStr(var_ActiveTerm)
        str_ActiveType = UCase(Examine_TermType(str_ActiveTerm))
        int_ActiveQty = Examine_TermQty(str_ActiveTerm)

        Select Case str_ActiveType
            Case "M": int_Output = int_Output + int_ActiveQty
            Case "Y": int_Output = int_Output + int_ActiveQty * 12
            Case Else
                int_Output = 0
                Exit Function
        End Select
    Next var_ActiveTerm

    Calc_NumMonths = int_Output
End Function

Public Function Calc_NumPeriods(str_Term As String, str_Freq As String) As Integer
    ' ## Calculate number of periods of the specified frequency exist within the specified term.  Assumes term is an exact multiple of frequency
    ' ## E.g. If term = 1Y and frequency = 3M then the number of periods will be 4
    Dim int_TermInMonths As Integer: int_TermInMonths = Calc_NumMonths(str_Term)
    Dim int_FreqInMonths As Integer: int_FreqInMonths = Calc_NumMonths(str_Freq)
    Dim int_Output As Integer

    ' Return 0 if frequency is not recognised
    If int_FreqInMonths <> 0 Then int_Output = int_TermInMonths / int_FreqInMonths

    Calc_NumPeriods = int_Output
End Function

Public Function Calc_YearFrac(var_StartDate As Variant, var_EndDate As Variant, str_Convention As String, _
    Optional str_CouponFreq As String = "6M", Optional bln_IsFwdGeneration As Boolean = False, Optional dbl_Coupon As Double = 0) As Double
    ' ## Calculate term in fractional years between two specified dates according to the specified daycount convention
    ' ## For Act/Act daycount, the coupon frequency is also required as an input.  For other conventions, this input is ignored

    Dim dbl_Output As Double
    Dim lng_CurPeriodStart As Long, lng_CurPeriodEnd As Long
    Dim int_PeriodInMonths As Integer
    Dim int_D1 As Integer, int_D2 As Integer, int_M1 As Integer, int_M2 As Integer, int_Y1 As Integer, int_Y2 As Integer
    Dim dbl_XTEPeriod As Double, lng_LastPeriodEnd As Long, lng_LastPeriodStart As Long, dbl_XTEterm As Double

    Select Case UCase(str_Convention)
        Case "ACT/365": dbl_Output = WorksheetFunction.YearFrac(var_StartDate, var_EndDate, 3)
        Case "ACT/360": dbl_Output = WorksheetFunction.YearFrac(var_StartDate, var_EndDate, 2)
        Case "ACT/ACT"
            Select Case UCase(str_CouponFreq)
                Case ""
                    dbl_Output = WorksheetFunction.YearFrac(var_StartDate, var_EndDate, 1)
                Case Else
                    int_PeriodInMonths = WorksheetFunction.Max(Calc_NumMonths(str_CouponFreq), 1)
                    If bln_IsFwdGeneration = True Then
                        lng_CurPeriodEnd = Date_AddTerm(CLng(var_StartDate), "1M", 1 * int_PeriodInMonths, True)  ' Exactly one coupon period ago
                        'lng_CurPeriodEnd = Date_ApplyBDC(lng_CurPeriodEnd, "MOD FOLL")  ' Not 100% sure this is the rule to apply
                        dbl_Output = (var_EndDate - var_StartDate) / (lng_CurPeriodEnd - var_StartDate) * int_PeriodInMonths / 12
                    Else
                        lng_CurPeriodStart = Date_AddTerm(CLng(var_EndDate), "1M", -1 * int_PeriodInMonths, True)  ' Exactly one coupon period ago
                        'lng_CurPeriodStart = Date_ApplyBDC(lng_CurPeriodStart, "MOD FOLL")  ' Not 100% sure this is the rule to apply
                        dbl_Output = (var_EndDate - var_StartDate) / (var_EndDate - lng_CurPeriodStart) * int_PeriodInMonths / 12
                    End If
            End Select
        Case "ACT/ACT CPN"
            Select Case UCase(str_CouponFreq)
                Case ""
                    dbl_Output = WorksheetFunction.YearFrac(var_StartDate, var_EndDate, 1)
                Case Else
                    int_PeriodInMonths = WorksheetFunction.Max(Calc_NumMonths(str_CouponFreq), 1)
                    If bln_IsFwdGeneration = True Then
                        lng_CurPeriodEnd = Date_AddTerm(CLng(var_StartDate), "1M", 1 * int_PeriodInMonths, False)  ' Exactly one coupon period ago
                        'lng_CurPeriodEnd = Date_ApplyBDC(lng_CurPeriodEnd, "MOD FOLL")  ' Not 100% sure this is the rule to apply
                        dbl_Output = (var_EndDate - var_StartDate) / (lng_CurPeriodEnd - var_StartDate) * int_PeriodInMonths / 12
                    Else
                        lng_CurPeriodStart = Date_AddTerm(CLng(var_EndDate), "1M", -1 * int_PeriodInMonths, False)  ' Exactly one coupon period ago
                        'lng_CurPeriodStart = Date_ApplyBDC(lng_CurPeriodStart, "MOD FOLL")  ' Not 100% sure this is the rule to apply
                        dbl_Output = (var_EndDate - var_StartDate) / (var_EndDate - lng_CurPeriodStart) * int_PeriodInMonths / 12
                    End If
            End Select
        Case "ACT/ACT XTE"
            Select Case UCase(str_CouponFreq)
                Case ""
                    dbl_Output = WorksheetFunction.YearFrac(var_StartDate, var_EndDate, 1)
                Case Else
                    int_PeriodInMonths = WorksheetFunction.Max(Calc_NumMonths(str_CouponFreq), 1)
                    If bln_IsFwdGeneration = True Then
                        lng_CurPeriodEnd = Date_AddTerm(CLng(var_StartDate), "1M", 1 * int_PeriodInMonths, True)  ' Exactly one coupon period ago
                        'lng_CurPeriodEnd = Date_ApplyBDC(lng_CurPeriodEnd, "MOD FOLL")  ' Not 100% sure this is the rule to apply
                        dbl_XTEPeriod = (var_EndDate - var_StartDate) / (lng_CurPeriodEnd - var_StartDate) * int_PeriodInMonths / 12
                        If dbl_XTEPeriod <= 1 Then
                            dbl_Output = dbl_XTEPeriod
                        Else
                            lng_LastPeriodEnd = Date_AddTerm(CLng(lng_CurPeriodEnd), "1M", 1 * int_PeriodInMonths, True)
                            dbl_XTEterm = (var_EndDate - lng_LastPeriodEnd) / (lng_LastPeriodEnd - lng_CurPeriodEnd)
                            dbl_Output = ((1 + dbl_Coupon * dbl_XTEterm / (12 / int_PeriodInMonths)) * (1 + dbl_Coupon / (12 / int_PeriodInMonths)) - 1) / dbl_Coupon
                        End If

                    Else
                        lng_CurPeriodStart = Date_AddTerm(CLng(var_EndDate), "1M", -1 * int_PeriodInMonths, True)  ' Exactly one coupon period ago
                        'lng_CurPeriodStart = Date_ApplyBDC(lng_CurPeriodStart, "MOD FOLL")  ' Not 100% sure this is the rule to apply
                        dbl_XTEPeriod = (var_EndDate - var_StartDate) / (var_EndDate - lng_CurPeriodStart) * int_PeriodInMonths / 12
                        If dbl_XTEPeriod <= 1 Then
                            dbl_Output = dbl_XTEPeriod
                        Else
                            lng_LastPeriodStart = Date_AddTerm(CLng(lng_CurPeriodStart), "1M", -1 * int_PeriodInMonths, True)
                            dbl_XTEterm = (var_StartDate - lng_LastPeriodStart) / (lng_CurPeriodStart - lng_LastPeriodStart)
                            dbl_Output = ((1 + dbl_Coupon * dbl_XTEterm / (12 / int_PeriodInMonths)) * (1 + dbl_Coupon / (12 / int_PeriodInMonths)) - 1) / dbl_Coupon
                        End If

                    End If
            End Select


        Case "ACT/ACT NM": dbl_Output = Round((var_EndDate - var_StartDate) / 30, 0) / 12
        Case "30/360": dbl_Output = WorksheetFunction.YearFrac(var_StartDate, var_EndDate, 0)
        Case "30E/360": dbl_Output = WorksheetFunction.YearFrac(var_StartDate, var_EndDate, 4)
        Case "30/360_ME30"
            int_D1 = Day(var_StartDate)
            int_M1 = Month(var_StartDate)
            int_Y1 = Year(var_StartDate)
            int_D2 = Day(var_EndDate)
            int_M2 = Month(var_EndDate)
            int_Y2 = Year(var_EndDate)

            If int_D1 = 31 Then int_D1 = 30

            If int_D2 = 31 And int_D1 = 30 Then int_D2 = 30


            dbl_Output = ((int_D2 - int_D1) + 30 * (int_M2 - int_M1) + 360 * (int_Y2 - int_Y1)) / 360

    End Select

    If var_StartDate > var_EndDate Then dbl_Output = -dbl_Output

    Calc_YearFrac = dbl_Output
End Function

Public Function Calc_NormalCDF(dbl_Z As Double, Optional enu_Method As NormalCDFMethod = NormalCDFMethod.Excel) As Double
    Dim dbl_Output As Double

    Select Case enu_Method
        Case NormalCDFMethod.Excel
            dbl_Output = WorksheetFunction.NormSDist(dbl_Z)
        Case NormalCDFMethod.Abram
            Const dbl_Pi As Double = 3.14159265358979
            Dim dbl_B0 As Double: dbl_B0 = 0.2316419
            Dim dbl_B1 As Double: dbl_B1 = 0.31938153
            Dim dbl_B2 As Double: dbl_B2 = -0.356563782
            Dim dbl_B3 As Double: dbl_B3 = 1.781477937
            Dim dbl_B4 As Double: dbl_B4 = -1.821255978
            Dim dbl_B5 As Double: dbl_B5 = 1.330274429
            Dim dbl_Transformed As Double, dbl_NormPdf As Double, bln_NegativeX As Double

            If dbl_Z < 0 Then
                dbl_Z = Abs(dbl_Z)
                bln_NegativeX = True
            Else
                bln_NegativeX = False
            End If

            dbl_Transformed = 1 / (1 + dbl_B0 * dbl_Z)
            dbl_NormPdf = Exp(-dbl_Z * dbl_Z / 2) / Sqr(2 * dbl_Pi)

            dbl_Output = 1 - dbl_NormPdf * (dbl_B1 * dbl_Transformed + dbl_B2 * (dbl_Transformed ^ 2) _
                + dbl_B3 * (dbl_Transformed ^ 3) + dbl_B4 * (dbl_Transformed ^ 4) + dbl_B5 * (dbl_Transformed ^ 5))

            If bln_NegativeX = True Then dbl_Output = 1 - dbl_Output
    End Select

    Calc_NormalCDF = dbl_Output
End Function

Public Function Calc_NormalPDF(dbl_Z As Double, Optional dbl_Mean As Double = 0, Optional dbl_Stdev As Double = 1) As Double
    Dim dbl_Pi As Double: dbl_Pi = 3.14159265358979
    Calc_NormalPDF = 1 / (dbl_Stdev * Sqr(2 * dbl_Pi)) * Exp(-(dbl_Z - dbl_Mean) ^ 2 / (2 * dbl_Stdev ^ 2))
End Function

Public Function Calc_BivarNormalCDF(dbl_Z1 As Double, dbl_Z2 As Double, dbl_Correl As Double) As Double
    ' ## Calculates approximate Bivariate normal distribution CDF
    ' ## Adapted from http://finance.bi.no/~bernt/gcc_prog/recipes/recipes/node23.html
    Dim dbl_Output As Double
    If dbl_Correl >= 1 Then dbl_Correl = 0.999999999999999
    If dbl_Correl <= -1 Then dbl_Correl = -0.999999999999999
    Dim dbl_ScalingFac As Double: dbl_ScalingFac = Sqr(2 * (1 - dbl_Correl ^ 2))
    Dim dbl_Z1_Std As Double: dbl_Z1_Std = dbl_Z1 / dbl_ScalingFac
    Dim dbl_Z2_Std As Double: dbl_Z2_Std = dbl_Z2 / dbl_ScalingFac

    If dbl_Z1 <= 0 And dbl_Z2 <= 0 And dbl_Correl <= 0 Then
        Dim dblArr_A(1 To 4) As Double
        dblArr_A(1) = 0.325303
        dblArr_A(2) = 0.4211071
        dblArr_A(3) = 0.1334425
        dblArr_A(4) = 0.006374323

        Dim dblArr_B(1 To 4) As Double
        dblArr_B(1) = 0.1337764
        dblArr_B(2) = 0.6243247
        dblArr_B(3) = 1.3425378
        dblArr_B(4) = 2.2626645

        Dim dbl_RunningTotal As Double: dbl_RunningTotal = 0
        Dim int_OuterCtr As Integer, int_InnerCtr As Integer
        For int_OuterCtr = 1 To 4
            For int_InnerCtr = 1 To 4
                dbl_RunningTotal = dbl_RunningTotal + dblArr_A(int_OuterCtr) * dblArr_A(int_InnerCtr) _
                    * Exp(dbl_Z1_Std * (2 * dblArr_B(int_OuterCtr) - dbl_Z1_Std) + dbl_Z2_Std * _
                    (2 * dblArr_B(int_InnerCtr) - dbl_Z2_Std) + 2 * dbl_Correl * (dblArr_B(int_OuterCtr) _
                    - dbl_Z1_Std) * (dblArr_B(int_InnerCtr) - dbl_Z2_Std))
            Next int_InnerCtr
        Next int_OuterCtr

        dbl_Output = dbl_RunningTotal * dbl_ScalingFac / (Sqr(2) * Pi)
    ElseIf dbl_Z1 <= 0 And dbl_Z2 >= 0 And dbl_Correl >= 0 Then
        dbl_Output = Calc_NormalCDF(dbl_Z1) - Calc_BivarNormalCDF(dbl_Z1, -dbl_Z2, -dbl_Correl)
    ElseIf dbl_Z1 >= 0 And dbl_Z2 <= 0 And dbl_Correl >= 0 Then
        dbl_Output = Calc_NormalCDF(dbl_Z2) - Calc_BivarNormalCDF(-dbl_Z1, dbl_Z2, -dbl_Correl)
    ElseIf dbl_Z1 >= 0 And dbl_Z2 >= 0 And dbl_Correl <= 0 Then
        dbl_Output = Calc_NormalCDF(dbl_Z1) + Calc_NormalCDF(dbl_Z2) - 1 _
            + Calc_BivarNormalCDF(-dbl_Z1, -dbl_Z2, dbl_Correl)
    Else
        Dim dbl_Denominator As Double: dbl_Denominator = Sqr(dbl_Z1 ^ 2 - 2 * dbl_Correl * dbl_Z1 * dbl_Z2 + dbl_Z2 ^ 2)
        Dim int_Sign1 As Integer, int_Sign2 As Integer
        If dbl_Z1 >= 0 Then int_Sign1 = 1 Else int_Sign1 = -1
        If dbl_Z2 >= 0 Then int_Sign2 = 1 Else int_Sign2 = -1

        Dim dbl_Correl_Adj1 As Double: dbl_Correl_Adj1 = (dbl_Correl * dbl_Z1 - dbl_Z2) * int_Sign1 / dbl_Denominator
        Dim dbl_Correl_Adj2 As Double: dbl_Correl_Adj2 = (dbl_Correl * dbl_Z2 - dbl_Z1) * int_Sign2 / dbl_Denominator
        Dim dbl_Delta As Double: dbl_Delta = (1 - int_Sign1 * int_Sign2) / 4

        dbl_Output = Calc_BivarNormalCDF(dbl_Z1, 0, dbl_Correl_Adj1) + Calc_BivarNormalCDF(dbl_Z2, 0, dbl_Correl_Adj2) - dbl_Delta
    End If

    Calc_BivarNormalCDF = dbl_Output
End Function


Public Function Calc_BAW_American(enu_Direction As OptionDirection, dbl_Spot As Double, dbl_Fwd As Double, _
    dbl_Strike As Double, dbl_VolPct As Double, dbl_DomDF As Double, dbl_TimeToMat As Double, _
    dbl_TimeEstPeriod As Double, Optional int_NumIterations As Integer = 1000) As Double
    ' ## Calculate discounted American-Exercise Option price using Barone-Adesi-Whaley Model, valued as at the spot date
    ' ## Adapted from Wallner & Wystup paper online

    Dim start As Single: start = Timer
    Dim dbl_Output As Double: dbl_Output = 0

    ' Static assignments
    Dim dbl_Vol As Double: dbl_Vol = dbl_VolPct / 100
    Dim dbl_FwdSpotRatio As Double: dbl_FwdSpotRatio = dbl_Fwd / dbl_Spot
    Dim dbl_VolSqrT As Double: dbl_VolSqrT = dbl_Vol * Sqr(dbl_TimeToMat)
    Dim dbl_TotalDrift_BN As Double: dbl_TotalDrift_BN = Math.Log(dbl_FwdSpotRatio) + 0.5 * dbl_VolSqrT ^ 2
    Dim dbl_DomZero As Double: dbl_DomZero = -Math.Log(dbl_DomDF) / dbl_TimeEstPeriod
    Dim dbl_FgnDF As Double: dbl_FgnDF = dbl_FwdSpotRatio * dbl_DomDF
    Dim dbl_FgnZero As Double: dbl_FgnZero = -Math.Log(dbl_FgnDF) / dbl_TimeEstPeriod
    Dim dbl_M As Double: dbl_M = 2 * dbl_DomZero / dbl_Vol ^ 2
    Dim dbl_N As Double: dbl_N = dbl_M - 2 * dbl_FgnZero / dbl_Vol ^ 2
    Dim dbl_K As Double: dbl_K = 1 - dbl_DomDF
    Dim dbl_Q As Double: dbl_Q = 0.5 * (-(dbl_N - 1) + enu_Direction * Sqr((dbl_N - 1) ^ 2 + 4 * dbl_M / dbl_K))  ' Roots of the solution to the simplified PDE

    ' Store parameters in dictionary
    Dim dic_StaticParams As Dictionary: Set dic_StaticParams = New Dictionary
    Call dic_StaticParams.Add("dbl_Strike", dbl_Strike)
    Call dic_StaticParams.Add("enu_Direction", enu_Direction)
    Call dic_StaticParams.Add("dbl_FwdSpotRatio", dbl_FwdSpotRatio)
    Call dic_StaticParams.Add("dbl_TimeToMat", dbl_TimeToMat)
    Call dic_StaticParams.Add("dbl_VolPct", dbl_VolPct)
    Call dic_StaticParams.Add("dbl_DomDF", dbl_DomDF)
    Call dic_StaticParams.Add("dbl_TotalDrift_BN", dbl_TotalDrift_BN)
    Call dic_StaticParams.Add("dbl_VolSqrT", dbl_VolSqrT)
    Call dic_StaticParams.Add("dbl_FgnDF", dbl_FgnDF)
    Call dic_StaticParams.Add("dbl_Q", dbl_Q)

    ' Iteration variables
    Dim dbl_ActiveThreshold As Double: dbl_ActiveThreshold = dbl_Spot  ' Optimal exercise spot price, iteratively solved
    Dim dbl_ActiveD1 As Double, dbl_ActiveNd1 As Double
    Dim dic_Outputs As Dictionary: Set dic_Outputs = New Dictionary

    ' Value variables
    Dim dbl_A As Double
    Dim dbl_EuropeanPrice As Double

    ' Iterate to derive optimal exercise spot price
    dbl_ActiveThreshold = Solve_FixedPt(ThisWorkbook, "SolverFuncXX_BAWThreshold", dic_StaticParams, dbl_Spot, 0.000000001, _
        int_NumIterations, -1, dic_Outputs)

    ' Calculate European price
    dbl_EuropeanPrice = Calc_BSPrice_Vanilla(enu_Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct) * dbl_DomDF

    ' Calculate American price, also handling cases where option should be immediately exercised
    If dbl_Spot >= dbl_ActiveThreshold And enu_Direction = OptionDirection.CallOpt Then
        dbl_Output = dbl_Spot - dbl_Strike
    ElseIf dbl_Spot <= dbl_ActiveThreshold And enu_Direction = OptionDirection.PutOpt Then
        dbl_Output = dbl_Strike - dbl_Spot
    Else
        dbl_A = enu_Direction * dbl_ActiveThreshold / dbl_Q * (1 - dbl_FgnDF * dic_Outputs("dbl_ActiveNd1"))
        dbl_Output = dbl_EuropeanPrice + dbl_A * (dbl_Spot / dbl_ActiveThreshold) ^ dbl_Q
    End If

    Calc_BAW_American = dbl_Output
End Function

Public Function Calc_AddLists(ParamArray lstArr_Items() As Variant) As Collection
    ' ## Input is a series of collections
    Dim dblLst_Output As New Collection
    Dim int_NumRows As Integer
    Dim dbl_ActiveSum As Double
    Dim var_Item As Variant

    If IsArray(lstArr_Items(0)) = True Then int_NumRows = UBound(lstArr_Items(0)) Else int_NumRows = lstArr_Items(0).Count

    Dim int_Ctr As Integer
    For int_Ctr = 1 To int_NumRows
        ' Take product of each term in the row
        dbl_ActiveSum = 0
        For Each var_Item In lstArr_Items
            dbl_ActiveSum = dbl_ActiveSum + var_Item(int_Ctr)
        Next var_Item

        ' Add sum to output list
        Call dblLst_Output.Add(dbl_ActiveSum)
    Next int_Ctr

    Set Calc_AddLists = dblLst_Output
End Function

Public Function Calc_AddDicts(dicLst_ToAdd As Collection) As Dictionary
    ' ## Returns a dictionary containing the sum of the specified original dictionaries at each specified key
    ' ## Assumes all dictionaries have the same set of keys
    Debug.Assert dicLst_ToAdd.Count > 0
    Dim dic_Output As New Dictionary: dic_Output.CompareMode = CompareMethod.TextCompare
    Dim lng_NumDicts As Long: lng_NumDicts = dicLst_ToAdd.Count
    Dim lng_RowCtr As Long, lng_DictCtr As Long, dbl_ActiveSum As Double
    Dim dic_Active As Dictionary, var_ActiveKey As Variant
    Dim dic_First As Dictionary: Set dic_First = dicLst_ToAdd(1)
    Dim varArr_Keys() As Variant: varArr_Keys = dic_First.Keys
    Dim lng_NumRows As Long: lng_NumRows = dic_First.Count

    For lng_RowCtr = 1 To lng_NumRows
        dbl_ActiveSum = 0
        var_ActiveKey = varArr_Keys(lng_RowCtr - 1)  ' Array is zero based

        For lng_DictCtr = 1 To lng_NumDicts
            Set dic_Active = dicLst_ToAdd(lng_DictCtr)
            dbl_ActiveSum = dbl_ActiveSum + dic_Active(var_ActiveKey)
        Next lng_DictCtr

        Call dic_Output.Add(var_ActiveKey, dbl_ActiveSum)
    Next lng_RowCtr

    Set Calc_AddDicts = dic_Output
End Function

Public Function Calc_ScaleList(dblLst_Items As Collection, dbl_Factor As Double) As Collection
    ' ## Perform scalar multiplication on a list
    Dim dblLst_Output As New Collection
    Dim int_NumRows As Integer: int_NumRows = dblLst_Items.Count

    Dim int_Ctr As Integer
    For int_Ctr = 1 To int_NumRows
        Call dblLst_Output.Add(dblLst_Items(int_Ctr) * dbl_Factor)
    Next int_Ctr

    Set Calc_ScaleList = dblLst_Output
End Function

Public Function Calc_SumProductOnList(ParamArray lstArr_Items() As Variant) As Double
    ' ## Input is a series of collections
    Dim dbl_Output As Double: dbl_Output = 0
    Dim int_NumRows As Integer
    Dim dbl_ActiveProduct As Double
    Dim var_Item As Variant

    If IsArray(lstArr_Items(0)) = True Then int_NumRows = UBound(lstArr_Items(0)) Else int_NumRows = lstArr_Items(0).Count

    Dim int_Ctr As Integer
    For int_Ctr = 1 To int_NumRows
        ' Take product of each term in the row
        dbl_ActiveProduct = 1
        For Each var_Item In lstArr_Items
            dbl_ActiveProduct = dbl_ActiveProduct * var_Item(int_Ctr)
        Next var_Item

        ' Sum each product
        dbl_Output = dbl_Output + dbl_ActiveProduct
    Next int_Ctr

    Calc_SumProductOnList = dbl_Output
End Function

Public Function Calc_PolyValue(arr_PolyCoefs As Variant, var_LookupX As Variant) As Double
    ' ## Evaluates the polynomial for the specified value of x.  Coefficients are ordered from the highest power to zero
    Dim dbl_Output As Double
    Dim int_LowerBound As Integer: int_LowerBound = Examine_LowerBoundIndex(arr_PolyCoefs)
    Dim int_UpperBound As Integer: int_UpperBound = Examine_UpperBoundIndex(arr_PolyCoefs)
    Dim int_NumPoints As Integer: int_NumPoints = int_UpperBound - int_LowerBound + 1

    ' Evaluate polynomial expression
    Dim int_Ctr As Integer
    For int_Ctr = 1 To int_NumPoints
        dbl_Output = dbl_Output + arr_PolyCoefs(int_LowerBound + int_Ctr - 1) * var_LookupX ^ (int_NumPoints - int_Ctr)
    Next int_Ctr

    Calc_PolyValue = dbl_Output
End Function

Public Function Calc_PolyCoefs(arr_X As Variant, arr_Y As Variant) As Double()
    ' ## Returns coefficients of the polynomial fitting the specified co-ordinates
    ' ## Coefficients are ordered from the highest power to zero
    Dim int_LowerBound As Integer: int_LowerBound = Examine_LowerBoundIndex(arr_X)
    Dim int_UpperBound As Integer: int_UpperBound = Examine_UpperBoundIndex(arr_X)
    Dim int_NumPoints As Integer: int_NumPoints = int_UpperBound - int_LowerBound + 1
    Dim dblArr_Output() As Double: ReDim dblArr_Output(1 To int_NumPoints) As Double
    Dim dblArr_SquareMat() As Double: ReDim dblArr_SquareMat(1 To int_NumPoints, 1 To int_NumPoints)
    Dim dblArr_SquareMatInv As Variant, dblArr_PolyCoefs As Variant
    Dim int_RowCtr As Integer, int_ColCtr As Integer

    ' Convert Y to a two-dimensional array to enable matrix operations
    Dim dblArr_Y2D() As Double: ReDim dblArr_Y2D(1 To int_NumPoints, 1 To 1) As Double
    For int_RowCtr = 1 To int_NumPoints
        dblArr_Y2D(int_RowCtr, 1) = arr_Y(int_LowerBound + int_RowCtr - 1)
    Next int_RowCtr

    ' Determine polynomial coefficients via solution to system of equations
    For int_RowCtr = 1 To int_NumPoints
        For int_ColCtr = 1 To int_NumPoints
            dblArr_SquareMat(int_RowCtr, int_ColCtr) = arr_X(int_LowerBound + int_RowCtr - 1) ^ (int_NumPoints - int_ColCtr)
        Next int_ColCtr
    Next int_RowCtr

    dblArr_SquareMatInv = WorksheetFunction.MInverse(dblArr_SquareMat)
    dblArr_PolyCoefs = WorksheetFunction.MMult(dblArr_SquareMatInv, dblArr_Y2D)

    ' Convert to collection
    For int_RowCtr = 1 To int_NumPoints
        dblArr_Output(int_RowCtr) = dblArr_PolyCoefs(int_RowCtr, 1)
    Next int_RowCtr

    Calc_PolyCoefs = dblArr_Output
End Function

Public Function Calc_PolyDerivCoefs(arr_PolyCoefs As Variant) As Double()
    ' ## Returns coefficients of the derivative of the specified polynomial
    Dim int_LowerBound As Integer: int_LowerBound = Examine_LowerBoundIndex(arr_PolyCoefs)
    Dim int_UpperBound As Integer: int_UpperBound = Examine_UpperBoundIndex(arr_PolyCoefs)
    Dim int_NumPoints As Integer: int_NumPoints = int_UpperBound - int_LowerBound + 1
    Dim dblArr_Output() As Double: ReDim dblArr_Output(int_LowerBound To int_UpperBound - 1) As Double

    If int_NumPoints = 1 Then
        ReDim dblArr_Output(int_LowerBound To int_LowerBound) As Double
        dblArr_Output(int_LowerBound) = 0
    Else
        ReDim dblArr_Output(int_LowerBound To int_UpperBound - 1) As Double

        Dim int_Ctr As Integer, int_ArrIndex As Integer, int_Power As Integer
        For int_Ctr = 1 To int_NumPoints - 1
            int_ArrIndex = int_LowerBound + int_Ctr - 1
            int_Power = int_NumPoints - int_Ctr
            dblArr_Output(int_ArrIndex) = int_Power * arr_PolyCoefs(int_ArrIndex)
        Next int_Ctr
    End If

    Calc_PolyDerivCoefs = dblArr_Output
End Function

Public Function Calc_NumPmtsInWindow(lng_StartDate As Long, lng_EndDate As Long, str_PmtFreq As String, cal_Pmt As Calendar, _
    str_BDC As String, bln_CheckEOM As Boolean, bln_IsFwdGeneration As Boolean, Optional ByRef lngLst_OutputDates As Collection = Nothing) As Integer
    ' ## Returns number of payments in the specified window.  Counts the base date but not the target date
    ' Set up forward or backward generation
    Dim lng_BaseDate As Long, lng_TargetDate As Long, int_Direction As Integer
    If bln_IsFwdGeneration = True Then
        lng_BaseDate = lng_StartDate
        lng_TargetDate = lng_EndDate
        int_Direction = 1
    Else
        lng_BaseDate = lng_EndDate
        lng_TargetDate = lng_StartDate
        int_Direction = -1
    End If

    Dim lng_ActiveDate As Long: lng_ActiveDate = lng_BaseDate
    Dim int_Ctr As Integer: int_Ctr = 0
    While (lng_TargetDate - lng_ActiveDate) * int_Direction > 0
        ' Generate date then see if it has reached or exceeded the target
        int_Ctr = int_Ctr + 1
        lng_ActiveDate = Date_NextCoupon(lng_BaseDate, str_PmtFreq, cal_Pmt, int_Ctr * int_Direction, bln_CheckEOM, str_BDC)
        If Not lngLst_OutputDates Is Nothing Then Call lngLst_OutputDates.Add(lng_ActiveDate)
    Wend

    Calc_NumPmtsInWindow = int_Ctr
End Function

Public Function Calc_Heaviside(var_Input As Variant) As Single
    ' ## Returns value of the Heaviside function for the specified input
    Dim sng_Output As Single
    If var_Input < 0 Then
        sng_Output = 0
    ElseIf var_Input = 0 Then
        sng_Output = 0.5
    Else
        sng_Output = 1
    End If

    Calc_Heaviside = sng_Output
End Function

Public Function Calc_QuantoCorrel(dbl_Vol_XY As Double, dbl_Vol_XQ As Double, dbl_Vol_YQ As Double) As Double
    ' ## Returns correlation between X and Q
    Calc_QuantoCorrel = (dbl_Vol_XQ ^ 2 - dbl_Vol_XY ^ 2 - dbl_Vol_YQ ^ 2) / (2 * dbl_Vol_XY * dbl_Vol_YQ)
End Function

Public Function Calc_DriftAdjFactor(dbl_Vol_XY As Double, dbl_Vol_YQ As Double, dbl_Correl_XQ As Double, dbl_TimeToMat As Double) As Double
    ' ## Return drift adjusted forward FX rate for quantos
    Calc_DriftAdjFactor = Exp(-dbl_Correl_XQ * dbl_Vol_XY * dbl_Vol_YQ * dbl_TimeToMat)
End Function