Option Explicit


Public Function SolverFuncXY_CapletVolToPriceQJK(dbl_Vol As Double, dic_Params As Dictionary, dic_SecondaryOutputs As Dictionary) As Double
   'new function added QJK 16/12/2014
    ' ## Try the specified caplet vol, and recalculate and read the resulting price of the cap
   ' Dim cvl_Curve As Data_CapVols: Set cvl_Curve = dic_Params("cvl_Curve")
   Dim cvl_Curve As Data_CapVolsQJK: Set cvl_Curve = dic_Params("cvl_Curve")   'Qcode 06/08/2014
    Dim irl_underlying As IRLeg: Set irl_underlying = dic_Params("irl_Underlying")
    Dim int_FinalIndex As Integer: int_FinalIndex = dic_Params("int_FinalIndex")
    'Dim rng_Vol As Range: Set rng_Vol = dic_Params("rng_Vol")
    Dim dblLst_CapletVols As Collection
    Dim cal_Deduction As Calendar

    Set cal_Deduction.HolDates = dic_Params("rng_HolDates")
    cal_Deduction.Weekends = dic_Params("str_Weekends")

    'rng_Vol.Value = dbl_Vol
    Call cvl_Curve.SetFinalVol(int_FinalIndex, dbl_Vol)
   ' Set dblLst_CapletVols = cvl_Curve.Lookup_VolSeries(int_FinalIndex, dic_Params("intLst_InterpPillars"), False)
    Set dblLst_CapletVols = cvl_Curve.Lookup_VolSeriesParVols(int_FinalIndex, dic_Params("intLst_InterpPillars"), False)
    SolverFuncXY_CapletVolToPriceQJK = irl_underlying.Calc_BSOptionValue(dic_Params("enu_Direction"), dic_Params("dbl_ATMStrike"), _
        dic_Params("int_Deduction"), cal_Deduction, True, dblLst_CapletVols)
End Function
Public Function SolverFuncXY_CapletVolToPriceCFSurfaceInterpolateonFWD(dbl_Vol As Double, dic_Params As Dictionary, dic_SecondaryOutputs As Dictionary) As Double
'QJK code 02/02/2015'added to avoid Lookup_VolSeries code interpolating from par vols on value<0.0001
    ' ## Try the specified caplet vol, and recalculate and read the resulting price of the cap
    Dim cvl_Curve As Data_CapVolsQJK: Set cvl_Curve = dic_Params("cvl_Curve")
    Dim irl_underlying As IRLeg: Set irl_underlying = dic_Params("irl_Underlying")
    Dim int_FinalIndex As Integer: int_FinalIndex = dic_Params("int_FinalIndex")
    'Dim rng_Vol As Range: Set rng_Vol = dic_Params("rng_Vol")
    Dim dblLst_CapletVols As Collection
    Dim cal_Deduction As Calendar

    Set cal_Deduction.HolDates = dic_Params("rng_HolDates")
    cal_Deduction.Weekends = dic_Params("str_Weekends")

    'rng_Vol.Value = dbl_Vol
    Call cvl_Curve.SetFinalVol(int_FinalIndex, dbl_Vol)
   ' Set dblLst_CapletVols = cvl_Curve.Lookup_VolSeries(int_FinalIndex, dic_Params("intLst_InterpPillars"), False)
    Set dblLst_CapletVols = cvl_Curve.Lookup_VolSeriesCFSurfaceInterpolateOnFWD(int_FinalIndex, dic_Params("intLst_InterpPillars"), False)   'QJK code 02/02/2015
    'Calc_BSOptionValueForCFSurface
    'SolverFuncXY_CapletVolToPrice = irl_Underlying.Calc_BSOptionValue(dic_Params("enu_Direction"), dic_Params("dbl_ATMStrike"), _
        dic_Params("int_Deduction"), cal_Deduction, True, dblLst_CapletVols)
        SolverFuncXY_CapletVolToPriceCFSurfaceInterpolateonFWD = irl_underlying.Calc_BSOptionValueForCFSurface(dic_Params("enu_Direction"), dic_Params("dbl_ATMStrike"), _
        dic_Params("int_Deduction"), cal_Deduction, True, dblLst_CapletVols)   'QJK code 02022015
End Function

Public Function SolverFuncXY_VolToStrike(dbl_InputVol As Double, dic_Params As Dictionary, dic_SecondaryOutputs As Dictionary) As Double
    ' ## Determine the strike of the option based on the specified vol and delta relationship
    Dim dbl_ConventionDelta As Double, dbl_Nd1 As Double, dbl_IntermediateVol As Double
    Dim dbl_LookupFwd As Double: dbl_LookupFwd = dic_Params("dbl_LookupFwd")
    Dim dbl_TimeToMat As Double: dbl_TimeToMat = dic_Params("dbl_TimeToMat_Lookup")
    Dim dbl_Strike As Double: dbl_Strike = dic_Params("dbl_Strike")

    dbl_ConventionDelta = Calc_BS_FwdDelta(OptionDirection.PutOpt, dbl_LookupFwd, dbl_Strike, dbl_TimeToMat, dbl_InputVol, dic_Params("bln_PID_Interp"))
    dbl_Nd1 = Calc_BS_FwdDelta(OptionDirection.CallOpt, dbl_LookupFwd, dbl_Strike, dbl_TimeToMat, dbl_InputVol, False)
    dbl_IntermediateVol = LookupFXSmile(dbl_ConventionDelta, dic_Params("str_Interp_Delta"), dic_Params)

    ' Calculate output from guess
    SolverFuncXY_VolToStrike = Calc_BS_StrikeFromDelta(dbl_Nd1, dbl_LookupFwd, dbl_TimeToMat, dbl_IntermediateVol)
End Function

Public Function SolverFuncXY_ParToMV(dbl_Par As Double, dic_Params As Dictionary, dic_SecondaryOutputs As Dictionary) As Double
    ' ## For IR leg, determine the MV if the rate or margin was set to the specified value
    Dim irl_Leg As IRLeg: Set irl_Leg = dic_Params("irl_Leg")
    Call irl_Leg.SetRateOrMargin(dbl_Par)

    SolverFuncXY_ParToMV = irl_Leg.marketvalue
End Function

Public Function SolverFuncXY_ZSpreadToMV(dbl_ZSpread As Double, dic_Params As Dictionary, dic_SecondaryOutputs As Dictionary) As Double
    ' ## For IR leg, determine the MV if the ZSpread was set to the specified value
    Dim irl_Leg As IRLeg: Set irl_Leg = dic_Params("irl_Leg")
    Call irl_Leg.SetZSpread(dbl_ZSpread)

    SolverFuncXY_ZSpreadToMV = irl_Leg.marketvalue
End Function

Public Function SolverFuncXY_ZSpreadToMV_Bond(dbl_ZSpread As Double, dic_Params As Dictionary, dic_SecondaryOutputs As Dictionary) As Double
    ' ## For Bond, determine the MV if the ZSpread was set to the specified value
    Dim irl_Leg As IRLeg_Bond: Set irl_Leg = dic_Params("irl_Leg")
    Call irl_Leg.SetZSpread(dbl_ZSpread)

    SolverFuncXY_ZSpreadToMV_Bond = irl_Leg.marketvalue
End Function

Public Function SolverFuncXY_ZeroToDF_Deposit(dbl_Zero As Double, dic_Params As Dictionary, dic_SecondaryOutputs As Dictionary) As Double
    ' ## Try the specified zero rate and read the resulting DF from the curve
    Dim irc_Curve As Data_IRCurve: Set irc_Curve = dic_Params("irc_Curve")

    Call irc_Curve.SetZeroRate(dic_Params("int_Index"), dbl_Zero)
    SolverFuncXY_ZeroToDF_Deposit = irc_Curve.Lookup_Rate(dic_Params("lng_StartDate"), dic_Params("lng_MatDate"), "DF")
End Function

Public Function SolverFuncXY_ZeroToDF_FXFwd(dbl_Zero As Double, dic_Params As Dictionary, dic_SecondaryOutputs As Dictionary) As Double
    ' ## Read the DF from the curve and compare to the target
    Dim irc_Curve As Data_IRCurve: Set irc_Curve = dic_Params("irc_Curve")
    Dim lng_SpotDate As Long: lng_SpotDate = dic_Params("lng_SpotDate")
    Dim lng_MatDate As Long: lng_MatDate = dic_Params("lng_MatDate")
    Dim dbl_FXSpot As Double: dbl_FXSpot = dic_Params("dbl_FXSpot")
    Dim dbl_DF_USDEstFwd As Double: dbl_DF_USDEstFwd = dic_Params("dbl_DF_USDEstFwd")
    Dim dbl_DF_USD As Double: dbl_DF_USD = dic_Params("dbl_DF_USD")
    Dim str_Quotation As String: str_Quotation = dic_Params("str_Quotation")
    Dim dbl_FXFwd_Start As Double: dbl_FXFwd_Start = dic_Params("dbl_FXFwd_Start")
    Dim dbl_FXFwd_End As Double: dbl_FXFwd_End = dic_Params("dbl_FXFwd_End")
    Dim dbl_FXFwd_DF As Double

    Call irc_Curve.SetZeroRate(dic_Params("int_Index"), dbl_Zero)

    If dic_Params("bln_ReqEstimatedFwd") = True Then
        If str_Quotation = "DIRECT" Then
            dbl_FXFwd_End = dbl_FXSpot * dbl_DF_USDEstFwd / irc_Curve.Lookup_Rate(lng_SpotDate, lng_MatDate, "DF")
        Else
            dbl_FXFwd_End = dbl_FXSpot * irc_Curve.Lookup_Rate(lng_SpotDate, lng_MatDate, "DF") / dbl_DF_USDEstFwd
        End If

        dbl_FXFwd_Start = dbl_FXFwd_End - dic_Params("dbl_Par") / 10000
    End If

    If str_Quotation = "DIRECT" Then
        dbl_FXFwd_DF = dbl_DF_USD / (dbl_FXFwd_End / dbl_FXFwd_Start)
    Else
        dbl_FXFwd_DF = dbl_DF_USD * (dbl_FXFwd_End / dbl_FXFwd_Start)
    End If

    SolverFuncXY_ZeroToDF_FXFwd = irc_Curve.Lookup_Rate(dic_Params("lng_StartDate"), lng_MatDate, "DF") - dbl_FXFwd_DF
End Function

Public Function SolverFuncXY_ZeroToMV_Swap(dbl_Zero As Double, dic_Params As Dictionary, dic_SecondaryOutputs As Dictionary) As Double
    ' ## Try the specified zero rate, and recalculate and read the resulting MV of the swap
    Dim irs_Swap As Inst_IRSwap: Set irs_Swap = dic_Params("irs_Swap")
    Dim irc_Curve As Data_IRCurve: Set irc_Curve = dic_Params("irc_Curve")

    Call irc_Curve.SetZeroRate(dic_Params("int_Index"), dbl_Zero)
    Call irs_Swap.HandleUpdate_IRC(dic_Params("str_CurveName"))
    SolverFuncXY_ZeroToMV_Swap = irs_Swap.marketvalue
End Function

Public Function SolverFuncXY_CapletVolToPrice(dbl_Vol As Double, dic_Params As Dictionary, dic_SecondaryOutputs As Dictionary) As Double
    ' ## Try the specified caplet vol, and recalculate and read the resulting price of the cap
    Dim cvl_Curve As Data_CapVols: Set cvl_Curve = dic_Params("cvl_Curve")
    Dim irl_underlying As IRLeg: Set irl_underlying = dic_Params("irl_Underlying")
    Dim int_FinalIndex As Integer: int_FinalIndex = dic_Params("int_FinalIndex")
    'Dim rng_Vol As Range: Set rng_Vol = dic_Params("rng_Vol")
    Dim dblLst_CapletVols As Collection
    Dim cal_Deduction As Calendar

    Set cal_Deduction.HolDates = dic_Params("rng_HolDates")
    cal_Deduction.Weekends = dic_Params("str_Weekends")

    'rng_Vol.Value = dbl_Vol
    Call cvl_Curve.SetFinalVol(int_FinalIndex, dbl_Vol)
    Set dblLst_CapletVols = cvl_Curve.Lookup_VolSeries(int_FinalIndex, dic_Params("intLst_InterpPillars"), False)

    SolverFuncXY_CapletVolToPrice = irl_underlying.Calc_BSOptionValue(dic_Params("enu_Direction"), dic_Params("dbl_ATMStrike"), _
        dic_Params("int_Deduction"), cal_Deduction, True, dblLst_CapletVols)
End Function

Public Function SolverFunc_Yield(dbl_yield As Double, dic_Params As Dictionary, dic_SecondaryOutputs As Dictionary) As Double
    '## by SW: Function to return dirty price based on input yield --> used for yield solving
Dim irl_Leg As IRLeg_Bond: Set irl_Leg = dic_Params("irl_leg")
Call irl_Leg.SetYield(dbl_yield)

SolverFunc_Yield = irl_Leg.DirtyPrice
End Function
Public Function SolverFunc_YieldBA(dbl_yield As Double, dic_Params As Dictionary, dic_SecondaryOutputs As Dictionary) As Double
    '## by SW: Function to return dirty price based on input yield --> used for yield solving
Dim irl_Leg As IRLeg_BA: Set irl_Leg = dic_Params("irl_leg")
Call irl_Leg.SetYield(dbl_yield)

SolverFunc_YieldBA = irl_Leg.DirtyPrice
End Function

Public Function SolverFunc_YieldNID(dbl_yield As Double, dic_Params As Dictionary, dic_SecondaryOutputs As Dictionary) As Double   'QJK 03102014
    '## by SW: Function to return dirty price based on input yield --> used for yield solving
Dim irl_Leg As IRLeg_NID: Set irl_Leg = dic_Params("irl_leg")
Call irl_Leg.SetYield(dbl_yield)

SolverFunc_YieldNID = irl_Leg.DirtyPrice
End Function