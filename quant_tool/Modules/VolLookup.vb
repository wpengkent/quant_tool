Option Explicit

Public Function LookupMSurface(dbl_ValDate As Double, dbl_LookupDate As Double, dbl_LookupMoneyness As Double, _
    rng_MatPillars As Range, rng_MonPillars As Range, rng_Surface As Range) As Double

    ' ## Interpolate V2t in time, then linear in moneyness
    ' ## Time is a vertical array, moneyness is a horizontal array

    Dim int_NumMonPillars As Integer: int_NumMonPillars = rng_MonPillars.Columns.count
    Dim int_NumMatPillars As Integer: int_NumMatPillars = rng_MatPillars.Rows.count

    ' Interpolation in time
    Dim dblArr_Smile() As Double: ReDim dblArr_Smile(1 To int_NumMonPillars) As Double
    Dim rng_ActiveCol As Range
    Dim int_ctr As Integer

    For int_ctr = 1 To int_NumMonPillars
        Set rng_ActiveCol = rng_Surface(1, int_ctr).Resize(int_NumMatPillars, 1)
        dblArr_Smile(int_ctr) = Interp_V2t_Range(rng_MatPillars, rng_ActiveCol, dbl_ValDate, dbl_LookupDate)
    Next int_ctr

    ' Interpolation in moneyness
    Dim dblArr_MonPillars() As Double: dblArr_MonPillars = Convert_RangeToDblArr(rng_MonPillars)
    LookupMSurface = Interp_Lin(dblArr_MonPillars, dblArr_Smile, dbl_LookupMoneyness, True)
End Function
Public Function LookupFXSmile(dbl_PutDelta As Double, str_Interp_Delta As String, dic_Params As Dictionary) As Double
    ' ## Reads the volatility surface value corresponding to the lookup date and lookup convention delta

    Dim dbl_Output As Double
    Dim dbl_IterLeftSpread As Double, dbl_IterRightSpread As Double

    Dim dblArr_LookupDeltaPillars() As Double: dblArr_LookupDeltaPillars = dic_Params("dblArr_LookupDeltaPillars")
    Dim dblArr_LeftDeltaPillars() As Double: dblArr_LeftDeltaPillars = dic_Params("dblArr_LeftDeltaPillars")
    Dim dblArr_RightDeltaPillars() As Double: dblArr_RightDeltaPillars = dic_Params("dblArr_RightDeltaPillars")
    Dim dblArr_LookupSmilePillars() As Double: dblArr_LookupSmilePillars = dic_Params("dblArr_LookupSmilePillars")
    Dim dblArr_LeftSmilePillars() As Double: dblArr_LeftSmilePillars = dic_Params("dblArr_LeftSmilePillars")
    Dim dblArr_RightSmilePillars() As Double: dblArr_RightSmilePillars = dic_Params("dblArr_RightSmilePillars")
    Dim dblArr_LookupPolyCoefs() As Double, dblArr_LeftPolyCoefs() As Double, dblArr_RightPolyCoefs() As Double
    Dim bln_IsSpotDeltaInterp As Boolean: bln_IsSpotDeltaInterp = dic_Params("bln_IsSpotDeltaInterp")
    Dim dbl_PutDelta_Lookup As Double, dbl_PutDelta_Left As Double, dbl_PutDelta_Right As Double

    If dic_Params("bln_OnPillar") = True Then
        ' Determine put delta for interpolation
        If bln_IsSpotDeltaInterp = True Then
            dbl_PutDelta_Lookup = dbl_PutDelta * dic_Params("dbl_LookupDF")
        Else
            dbl_PutDelta_Lookup = dbl_PutDelta
        End If

        ' Look up smile directly
        Select Case str_Interp_Delta
            Case "POLYNOMIAL"
                ' Polynomial coefficients must be externally derived and supplied
                dblArr_LookupPolyCoefs = dic_Params("dblArr_LookupPolyCoefs")
                dbl_Output = Calc_PolyValue(dblArr_LookupPolyCoefs, dbl_PutDelta_Lookup)
            Case "SPLINE"
                dbl_Output = Interp_Spline(dblArr_LookupDeltaPillars, dblArr_LookupSmilePillars, dbl_PutDelta_Lookup)
        End Select
    Else
        ' Determine put delta for interpolation
        If bln_IsSpotDeltaInterp = True Then
            dbl_PutDelta_Left = dbl_PutDelta * dic_Params("dbl_LeftPillarDF")
            dbl_PutDelta_Right = dbl_PutDelta * dic_Params("dbl_RightPillarDF")
        Else
            dbl_PutDelta_Left = dbl_PutDelta
            dbl_PutDelta_Right = dbl_PutDelta
        End If

        ' Look up smile for the option delta at maturity pillars and apply the interpolated spread to the ATM vol for the lookup date
        Select Case str_Interp_Delta
            Case "POLYNOMIAL"
                ' Polynomial coefficients must be externally derived and supplied
                dblArr_LeftPolyCoefs = dic_Params("dblArr_LeftPolyCoefs")
                dbl_IterLeftSpread = Calc_PolyValue(dblArr_LeftPolyCoefs, dbl_PutDelta_Left) - dic_Params("dbl_LeftATMVol")

                dblArr_RightPolyCoefs = dic_Params("dblArr_RightPolyCoefs")
                dbl_IterRightSpread = Calc_PolyValue(dblArr_RightPolyCoefs, dbl_PutDelta_Right) - dic_Params("dbl_RightATMVol")
            Case "SPLINE"
                dbl_IterLeftSpread = Interp_Spline(dblArr_LeftDeltaPillars, dblArr_LeftSmilePillars, _
                    dbl_PutDelta_Left) - dic_Params("dbl_LeftATMVol")
                dbl_IterRightSpread = Interp_Spline(dblArr_RightDeltaPillars, dblArr_RightSmilePillars, _
                    dbl_PutDelta_Right) - dic_Params("dbl_RightATMVol")
        End Select

        dbl_Output = dic_Params("dbl_LookupATMVol") + Interp_Lin_Binary(dic_Params("lng_LeftPillarDate"), dic_Params("lng_RightPillarDate"), _
            dbl_IterLeftSpread, dbl_IterRightSpread, dic_Params("lng_LookupDate"))
    End If

    LookupFXSmile = dbl_Output
End Function

Public Function FXSmileSlope(dbl_Delta As Double, str_Interp_Delta As String, dic_Params As Dictionary, Optional dbl_AbsDeviation As Double = 0.0001) As Double
    Dim dbl_Vol_Up As Double, dbl_Vol_Down As Double
    dbl_Vol_Up = LookupFXSmile(dbl_Delta + dbl_AbsDeviation, str_Interp_Delta, dic_Params)
    dbl_Vol_Down = LookupFXSmile(dbl_Delta - dbl_AbsDeviation, str_Interp_Delta, dic_Params)
    FXSmileSlope = (dbl_Vol_Up - dbl_Vol_Down) / (2 * dbl_AbsDeviation)
End Function