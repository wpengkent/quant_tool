Option Explicit


' ## MEMBER DATA

Private Enum Dates
    Calc_Start = 1
    Calc_End
    Fix_Start
    Fix_StartSpot
    Fix_IdStartSpot
    Fix_End
    Fix_EndSpot
    Fix_IdEndSpot
    Fix_NotionalRef
    Pmt_Date
End Enum

Private Enum EqLeg
    EqStart = 1
    FxStart
    EqStartFix
    EqEnd
    FxEnd
    EqEndFix
    NumShare
    Notional
    Rate
    DF

    FlowType
    Flow
    DiscFlow
    DiscFlowPnL
End Enum

Private Enum IdLeg
    Notional = 1
    Rate
    DF

    FlowType
    Flow
    DiscFlow
    DiscFlowPnL
End Enum

Private Enum DivLeg
    EqsDivPmt

    NumOfShare = 1
    Div
    fx
    FixDiv
    DF
    FlowType

    Flow
    DiscFlow

    DiscFlowCash
    DiscFlowMV
    DiscFlowPnL
End Enum



Private Enum Output
    SwapCash = 1
    SwapMV
    SwapPnL
    EQ_NOP
End Enum



' Dependent curves
Private fxs_Spots As Data_FXSpots, eqd_Spots As Data_EQSpots, _
    irc_Disc_A As Data_IRCurve, irc_Est_A As Data_IRCurve, _
    irc_Disc_B As Data_IRCurve, irc_Est_B As Data_IRCurve, _
    irc_Fx_Dom As Data_IRCurve, irc_Fx_For As Data_IRCurve

' Static values
Private dic_GlobalStaticInfo As Dictionary

Private lng_ValDate As Long, lng_SwapStart As Long, lng_DateSchedule_A() As Long, _
    lng_DateSchedule_B() As Long, lng_SpotDate As Long, lng_FutDivPmt As Long, lng_ExDivDate As Long, _
    lng_DateSchedule_B_Fix() As Long ' For case when fixing date = calc start date

Private int_NumFlow_A As Integer, int_NumFlow_B As Integer, _
    int_EqSpotDays As Integer, int_IdSpotDays As Integer, _
    int_sign_A As Integer, int_sign_B As Integer, int_NumDiv As Integer

Private str_GTerm As String, str_freq_A As String, str_freq_B As String, _
    str_BDC_A As String, str_BDC_B As String, _
    str_Ccy_Eq As String, str_Ccy_Fix As String, str_CCY_PnL As String, str_Ccy_Id As String, _
    str_ConstType As String, str_SecCode As String, str_IdDayCnt As String, _
    str_DivType As String, str_OutputType As String

Private dbl_OriNotional As Double, dbl_OriNumShare As Double, _
    dbl_RateOrMagin As Double, dbl_FutDiv As Double


Private cal_Fix_A As Calendar, cal_Pmt_A As Calendar, cal_Fix_B As Calendar, _
    cal_Pmt_B As Calendar, cal_None As Calendar
Private bln_EOM_A As Boolean, bln_EOM_B As Boolean, bln_EqPayer As Boolean, _
    bln_IndexFix As Boolean, bln_DivExist As Boolean, bln_TotalRet As Boolean, bln_CalcEqualFix As Boolean

Private dic_FxFixing As New Dictionary, dic_EqFixing As New Dictionary, _
    dic_IdFixing As New Dictionary, dic_DivInfo As New Dictionary
Private dic_dates_A As New Dictionary, dic_dates_B As New Dictionary
Private dic_flow_A As New Dictionary, dic_flow_B As New Dictionary
Private dic_output As New Dictionary, dic_Div As New Dictionary
Private dic_CustomDate As New Dictionary



' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As InstParams_EQS, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing)

    ' Dependent curves

    If dic_CurveSet Is Nothing Then
        Set irc_Disc_A = GetObject_IRCurve(fld_ParamsInput.EqDiscCurve, True, False)
        Set irc_Est_A = GetObject_IRCurve(fld_ParamsInput.EqEstCurve, True, False)
        Set irc_Disc_B = GetObject_IRCurve(fld_ParamsInput.IdDiscCurve, True, False)

        Set fxs_Spots = GetObject_FXSpots(True)
        Set eqd_Spots = GetObject_EQSpots(True)
    Else
        Set irc_Disc_A = dic_CurveSet(CurveType.IRC)(fld_ParamsInput.EqDiscCurve)
        Set irc_Est_A = dic_CurveSet(CurveType.IRC)(fld_ParamsInput.EqEstCurve)
        Set irc_Disc_B = dic_CurveSet(CurveType.IRC)(fld_ParamsInput.IdDiscCurve)

        Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
        Set eqd_Spots = dic_CurveSet(CurveType.EQSPT)
    End If


    bln_IndexFix = fld_ParamsInput.IdIsFix

    If bln_IndexFix = False Then
        If dic_CurveSet Is Nothing Then
            Set irc_Est_B = GetObject_IRCurve(fld_ParamsInput.IdEstCurve, True, False)
        Else
            Set irc_Est_B = dic_CurveSet(CurveType.IRC)(fld_ParamsInput.IdEstCurve)
        End If
    End If

    ' Store static values

    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    Dim cas_Found As CalendarSet: Set cas_Found = dic_GlobalStaticInfo(StaticInfoType.CalendarSet)


    'Dates Generation

    With fld_ParamsInput
        lng_ValDate = .ValueDate
        lng_SwapStart = .Swapstart
        str_GTerm = .Term
        str_CCY_PnL = .CCY_PnL
        str_OutputType = .PLType
        bln_CalcEqualFix = .IsCalcEqualFix

        Set dic_CustomDate = .CustomDate


        str_freq_A = .EqFreq
        int_EqSpotDays = .EqSpotDays
        str_Ccy_Eq = .CCY_Eq
        str_Ccy_Fix = .CCY_Fix
        str_BDC_A = .EqBDC
        cal_Fix_A = cas_Found.Lookup_Calendar(.EqFixCal)
        cal_Pmt_A = cas_Found.Lookup_Calendar(.EqPmtCal)
        str_ConstType = .ConstantType
        dbl_OriNotional = .Notional
        dbl_OriNumShare = .Quantity
        Set dic_EqFixing = .EqFixing
        Set dic_FxFixing = .FxFixing
        bln_EqPayer = .IsEqPayer

        str_freq_B = .IdFreq
        int_IdSpotDays = .IdSpotDays
        str_Ccy_Id = .CCY_Id
        str_BDC_B = .IdBDC
        cal_Pmt_B = cas_Found.Lookup_Calendar(.IdPmtCal)

        If bln_IndexFix = False Then
            cal_Fix_B = cas_Found.Lookup_Calendar(.IdFixCal)
        End If


        str_SecCode = .Security
        Set dic_IdFixing = .IdFixing
        str_IdDayCnt = .IdDayCnt
        dbl_RateOrMagin = .RateOrMargin






        str_DivType = .EqDivType

        If str_DivType = "Cash" Then
            bln_DivExist = True
            Set dic_DivInfo = .EqDiv
            int_NumDiv = (dic_DivInfo.count - 1) / dic_DivInfo("COUNT")
            lng_FutDivPmt = dic_DivInfo("PMT|" & int_NumDiv)
            dbl_FutDiv = dic_DivInfo("DIV|" & int_NumDiv)
            lng_ExDivDate = dic_DivInfo("EX|" & int_NumDiv)
        Else
            bln_DivExist = False
        End If

        bln_TotalRet = .IsTotalRet
    End With



    Set irc_Fx_For = GetObject_IRCurve(str_Ccy_Eq & "_FX", True, False)
    Set irc_Fx_Dom = GetObject_IRCurve(str_Ccy_Fix & "_FX", True, False)

    cal_None = cas_Found.Lookup_Calendar("NONE")

    If bln_EqPayer = True Then
        int_sign_A = -1
        int_sign_B = 1
    Else
        int_sign_A = 1
        int_sign_B = -1
    End If

    Call FillSpotDate
    Call GenerateDates

End Sub


' ## PROPERTIES - PUBLIC
Public Property Get marketvalue() As Double

    If str_OutputType = "PL" Then
        Call GenerateFlows
        marketvalue = dic_output(Output.SwapMV)
    ElseIf str_OutputType = "Rho" Then
        marketvalue = Rho
    ElseIf str_OutputType = "EQ NOP" Then
        Call GenerateFlows
        marketvalue = dic_output(Output.EQ_NOP)
    End If
End Property

Public Property Get Cash() As Double
    If str_OutputType = "PL" Then
        Cash = dic_output(Output.SwapCash)
    End If
End Property
Public Property Get PnL() As Double
    PnL = marketvalue + Cash
End Property
Public Property Get Rho() As Double

Dim dbl_OriMV As Double
Dim dbl_ShockMV As Double
Dim dbl_Output As Double

Dim dic_curve As New Dictionary

Call GenerateFlows
dbl_OriMV = dic_output(Output.SwapMV)

'## Rho for equity swaps = MV ( rate curve shock up 1 bps point ) - Base MV
'## original code for PL variation for parallel upward curve shift by 1 Bps

'irc_Disc_A.SetCurveState (Zero_Up1BP)
'    Call GenerateFlows
'    dbl_ShockMV = dic_output(Output.SwapMV)
'    dbl_output = dbl_output + (dbl_ShockMV - dbl_OriMV)
'    dic_curve.Add irc_Disc_A.CurveName, irc_Disc_A.CurveName
'irc_Disc_A.SetCurveState (Final)
'
'
'If dic_curve.Exists(irc_Est_A.CurveName) = False Then
'    irc_Est_A.SetCurveState (Zero_Up1BP)
'        Call GenerateFlows
'        dbl_ShockMV = dic_output(Output.SwapMV)
'        dbl_output = dbl_output + (dbl_ShockMV - dbl_OriMV)
'    irc_Est_A.SetCurveState (Final)
'    dic_curve.Add irc_Est_A.CurveName, irc_Est_A.CurveName
'End If
'
'If dic_curve.Exists(irc_Disc_B.CurveName) = False Then
'    irc_Disc_B.SetCurveState (Zero_Up1BP)
'        Call GenerateFlows
'        dbl_ShockMV = dic_output(Output.SwapMV)
'        dbl_output = dbl_output + (dbl_ShockMV - dbl_OriMV)
'    irc_Disc_B.SetCurveState (Final)
'    dic_curve.Add irc_Disc_B.CurveName, irc_Disc_B.CurveName
'End If
'
'If dic_curve.Exists(irc_Est_B.CurveName) = False Then
'    If bln_IndexFix = False Then
'        irc_Est_B.SetCurveState (Zero_Up1BP)
'            Call GenerateFlows
'            dbl_ShockMV = dic_output(Output.SwapMV)
'            dbl_output = dbl_output + (dbl_ShockMV - dbl_OriMV)
'        irc_Est_B.SetCurveState (Final)
'        dic_curve.Add irc_Est_B.CurveName, irc_Est_B.CurveName
'    End If
'End If
'
'
'If str_Ccy_Eq <> str_Ccy_Fix Then
'    If dic_curve.Exists(irc_Fx_For.CurveName) = False Then
'        Call fxs_Spots.SetCurveState(irc_Fx_For.CurveName, Zero_Up1BP)
'            Call GenerateFlows
'            dbl_ShockMV = dic_output(Output.SwapMV)
'            dbl_output = dbl_output + (dbl_ShockMV - dbl_OriMV)
'        Call fxs_Spots.SetCurveState(irc_Fx_For.CurveName, original)
'        dic_curve.Add irc_Fx_For.CurveName, irc_Fx_For.CurveName
'    End If
'
'    If dic_curve.Exists(irc_Fx_Dom.CurveName) = False Then
'        Call fxs_Spots.SetCurveState(irc_Fx_Dom.CurveName, Zero_Up1BP)
'            Call GenerateFlows
'            dbl_ShockMV = dic_output(Output.SwapMV)
'            dbl_output = dbl_output + (dbl_ShockMV - dbl_OriMV)
'        Call fxs_Spots.SetCurveState(irc_Fx_Dom.CurveName, Zero_Up1BP)
'        dic_curve.Add irc_Fx_Dom.CurveName, irc_Fx_Dom.CurveName
'    End If
'End If



'## code for PL variation for 1Bps upward shift of each pillar

Dim int_num As Integer
Dim int_cnt As Integer

int_num = irc_Disc_A.NumPoints
For int_cnt = 1 To int_num
    Call irc_Disc_A.SetCurveState(Zero_Up1BP, int_cnt)
    Call GenerateFlows
    dbl_ShockMV = dic_output(Output.SwapMV)
    dbl_Output = dbl_Output + (dbl_ShockMV - dbl_OriMV)
    irc_Disc_A.SetCurveState (Final)
Next int_cnt


dic_curve.Add irc_Disc_A.CurveName, irc_Disc_A.CurveName


If dic_curve.Exists(irc_Est_A.CurveName) = False Then

    int_num = irc_Est_A.NumPoints
    For int_cnt = 1 To int_num
        Call irc_Est_A.SetCurveState(Zero_Up1BP, int_cnt)
        Call GenerateFlows
        dbl_ShockMV = dic_output(Output.SwapMV)
        dbl_Output = dbl_Output + (dbl_ShockMV - dbl_OriMV)
        irc_Est_A.SetCurveState (Final)
    Next int_cnt

    dic_curve.Add irc_Est_A.CurveName, irc_Est_A.CurveName
End If

If dic_curve.Exists(irc_Disc_B.CurveName) = False Then

    int_num = irc_Disc_B.NumPoints
    For int_cnt = 1 To int_num
        Call irc_Disc_B.SetCurveState(Zero_Up1BP, int_cnt)
        Call GenerateFlows
        dbl_ShockMV = dic_output(Output.SwapMV)
        dbl_Output = dbl_Output + (dbl_ShockMV - dbl_OriMV)
        irc_Disc_B.SetCurveState (Final)
    Next int_cnt


    dic_curve.Add irc_Disc_B.CurveName, irc_Disc_B.CurveName
End If

If bln_IndexFix = False Then
    If dic_curve.Exists(irc_Est_B.CurveName) = False Then

        int_num = irc_Est_B.NumPoints
        For int_cnt = 1 To int_num
            Call irc_Est_B.SetCurveState(Zero_Up1BP, int_cnt)
            Call GenerateFlows
            dbl_ShockMV = dic_output(Output.SwapMV)
            dbl_Output = dbl_Output + (dbl_ShockMV - dbl_OriMV)
            irc_Est_B.SetCurveState (Final)
        Next int_cnt
        dic_curve.Add irc_Est_B.CurveName, irc_Est_B.CurveName
    End If
End If


If str_Ccy_Eq <> str_Ccy_Fix Then
    If dic_curve.Exists(irc_Fx_For.CurveName) = False Then

        int_num = irc_Fx_For.NumPoints
        For int_cnt = 1 To int_num
            Call fxs_Spots.SetCurveState(irc_Fx_For.CurveName, Zero_Up1BP, int_cnt)
            Call GenerateFlows
            dbl_ShockMV = dic_output(Output.SwapMV)
            dbl_Output = dbl_Output + (dbl_ShockMV - dbl_OriMV)
            Call fxs_Spots.SetCurveState(irc_Fx_For.CurveName, Final)
        Next int_cnt


        dic_curve.Add irc_Fx_For.CurveName, irc_Fx_For.CurveName
    End If

    If dic_curve.Exists(irc_Fx_Dom.CurveName) = False Then

        int_num = irc_Fx_Dom.NumPoints
        For int_cnt = 1 To int_num
            Call fxs_Spots.SetCurveState(irc_Fx_Dom.CurveName, Zero_Up1BP, int_cnt)
            Call GenerateFlows
            dbl_ShockMV = dic_output(Output.SwapMV)
            dbl_Output = dbl_Output + (dbl_ShockMV - dbl_OriMV)
            Call fxs_Spots.SetCurveState(irc_Fx_Dom.CurveName, Final)
        Next int_cnt


        dic_curve.Add irc_Fx_Dom.CurveName, irc_Fx_Dom.CurveName
    End If
End If

Rho = dbl_Output

End Property

' ## FUNCTION - PRIVATE
Private Function str_key(int_DateDesc As Integer, int_num As Integer) As String
    str_key = int_DateDesc & "|" & int_num
End Function

Private Function rng_format(wks_output As Worksheet, rng_TopLeft As Range) As Range

    Set rng_format = wks_output.Range( _
        wks_output.Range(rng_TopLeft, rng_TopLeft.End(xlToRight)), _
        wks_output.Range(rng_TopLeft, rng_TopLeft.End(xlToRight).End(xlDown)))
End Function



' ## PROPERTIES - PRIVATE
Private Property Get FillDate(lng_date As Long, int_days As Integer, cal As Calendar) As Long
    ' ## Compute spot date
    If int_days = 0 Then
        FillDate = Date_ApplyBDC(lng_date, "FOLL", cal.HolDates, cal.Weekends)
    Else
        FillDate = date_workday(lng_date, int_days, cal.HolDates, cal.Weekends)
    End If
End Property

Private Property Get FXConvFactor() As Double
    ' ## Discounted spot to translate into the PnL currency
    FXConvFactor = fxs_Spots.Lookup_DiscSpot(str_Ccy_Fix, str_CCY_PnL)
End Property

Private Property Get EqSpot() As Double
    ' Equity spot look up
    EqSpot = eqd_Spots.Lookup_Spot(str_SecCode)
End Property

' ## METHODS - PRIVATE
Private Sub FillSpotDate()
    ' ## Compute and store spot date
    If int_EqSpotDays = 0 Then
        lng_SpotDate = Date_ApplyBDC(lng_ValDate, "FOLL", cal_Pmt_A.HolDates, cal_Pmt_A.Weekends)
    Else
        lng_SpotDate = date_workday(lng_ValDate, int_EqSpotDays, cal_Pmt_A.HolDates, cal_Pmt_A.Weekends)
    End If
End Sub

' ## METHODS - CHANGE PARAMETERS / UPDATE
Private Sub GenerateDates()

    int_NumFlow_A = Calc_NumPeriods(str_GTerm, str_freq_A)
    int_NumFlow_B = Calc_NumPeriods(str_GTerm, str_freq_B)

    lng_DateSchedule_A = Date_CouponSchedule(lng_SwapStart, str_GTerm, cal_Pmt_A, str_freq_A, _
                            str_BDC_A, bln_EOM_A)
    lng_DateSchedule_B = Date_CouponSchedule(lng_SwapStart, str_GTerm, cal_Pmt_B, str_freq_B, _
                            str_BDC_B, bln_EOM_B)

    Dim int_FlowCnt_A As Integer, int_FlowCnt_B As Integer, int_RefNotionalPos As Integer
    Dim lng_CalcStart_A As Long, lng_CalcStart_B As Long
    Dim lng_CalcEnd_A As Long, lng_CalcEnd_B As Long
    Dim lng_EstStart_A As Long, lng_EstStart_B As Long
    Dim lng_EstStartSpot_A As Long, lng_EstStartSpot_B As Long
    Dim lng_EstEnd_A As Long, lng_EstEnd_B As Long
    Dim lng_EstEndSpot_A As Long, lng_EstEndSpot_B As Long
    Dim lng_RefNotional_B As Long
    Dim lng_IdStartSpot_B As Long, lng_IdEndSpot_B As Long

If bln_CalcEqualFix = False Then

    'Handle case when fixing date <> calc start date
    For int_FlowCnt_A = 1 To int_NumFlow_A

        'EQ Calc Start
        If int_FlowCnt_A = 1 Then
            lng_CalcStart_A = lng_SwapStart
        Else
            lng_CalcStart_A = dic_dates_A(str_key(Dates.Calc_End, int_FlowCnt_A - 1))
        End If

        dic_dates_A.Add str_key(Dates.Calc_Start, int_FlowCnt_A), _
            lng_CalcStart_A


        'EQ Calc End &  Eq Pmt Date
        lng_CalcEnd_A = lng_DateSchedule_A(int_FlowCnt_A)
        dic_dates_A.Add str_key(Dates.Calc_End, int_FlowCnt_A), _
            lng_CalcEnd_A
        dic_dates_A.Add str_key(Dates.Pmt_Date, int_FlowCnt_A), _
            lng_CalcEnd_A

        'EQ Fix Start and Fix Start Spot and FX Spot Fix Start

        lng_EstStartSpot_A = lng_CalcStart_A
        dic_dates_A.Add str_key(Dates.Fix_StartSpot, int_FlowCnt_A), _
            lng_EstStartSpot_A

        lng_EstStart_A = FillDate(lng_EstStartSpot_A, -int_EqSpotDays, cal_Pmt_A)
        dic_dates_A.Add str_key(Dates.Fix_Start, int_FlowCnt_A), _
            lng_EstStart_A

        'EQ Fix End and Fix End Spot and FX Spot Fix End

        'lng_EstEndSpot_A = WorksheetFunction.Max(Date_NextCoupon(lng_EstStartSpot_A, _
                        str_freq_A, cal_Fix_A, 1, bln_EOM_A, str_BDC_A), lng_CalcEnd_A)

        lng_EstEndSpot_A = lng_CalcEnd_A

        dic_dates_A.Add str_key(Dates.Fix_EndSpot, int_FlowCnt_A), _
            lng_EstEndSpot_A


        lng_EstEnd_A = FillDate(lng_EstEndSpot_A, -int_EqSpotDays, cal_Pmt_A)
        dic_dates_A.Add str_key(Dates.Fix_End, int_FlowCnt_A), _
            lng_EstEnd_A


    Next int_FlowCnt_A


    For int_FlowCnt_B = 1 To int_NumFlow_B

        'Index Calc Start
        If int_FlowCnt_B = 1 Then
            lng_CalcStart_B = lng_SwapStart
        Else
            lng_CalcStart_B = dic_dates_B(str_key(Dates.Calc_End, int_FlowCnt_B - 1))
        End If

        dic_dates_B.Add str_key(Dates.Calc_Start, int_FlowCnt_B), _
            lng_CalcStart_B


        'Index Calc End &  Index Pmt Date
        lng_CalcEnd_B = lng_DateSchedule_B(int_FlowCnt_B)
        dic_dates_B.Add str_key(Dates.Calc_End, int_FlowCnt_B), _
            lng_CalcEnd_B
        dic_dates_B.Add str_key(Dates.Pmt_Date, int_FlowCnt_B), _
            lng_CalcEnd_B

        'Index Fix Start and Fix Start Spot and FX Spot Fix Start


        lng_EstStartSpot_B = lng_CalcStart_B
        dic_dates_B.Add str_key(Dates.Fix_StartSpot, int_FlowCnt_B), _
            lng_EstStartSpot_B

        lng_EstStart_B = FillDate(lng_EstStartSpot_B, -int_EqSpotDays, cal_Pmt_A)
        dic_dates_B.Add str_key(Dates.Fix_Start, int_FlowCnt_B), _
            lng_EstStart_B

        If bln_IndexFix = False Then
            lng_IdStartSpot_B = FillDate(lng_EstStart_B, int_IdSpotDays, cal_Fix_B)
            dic_dates_B.Add str_key(Dates.Fix_IdStartSpot, int_FlowCnt_B), _
                lng_IdStartSpot_B
        End If

        'Index Fix End and Fix End Spot and FX Spot Fix End

        If bln_IndexFix = False Then

            lng_EstEndSpot_B = WorksheetFunction.Max(Date_NextCoupon(lng_EstStartSpot_B, _
                            str_freq_B, cal_Fix_B, 1, bln_EOM_B, str_BDC_B), lng_CalcEnd_B)
            dic_dates_B.Add str_key(Dates.Fix_EndSpot, int_FlowCnt_B), _
                lng_EstEndSpot_B

            lng_EstEnd_B = FillDate(lng_EstEndSpot_B, -int_EqSpotDays, cal_Pmt_A)
            dic_dates_B.Add str_key(Dates.Fix_End, int_FlowCnt_B), _
                lng_EstEnd_B


            'lng_IdEndSpot_B = FillDate(lng_EstEnd_B, int_IdSpotDays, cal_Pmt_B)


            lng_IdEndSpot_B = Date_NextCoupon(lng_IdStartSpot_B, _
                            str_freq_B, cal_Fix_B, 1, bln_EOM_B, str_BDC_B)



            dic_dates_B.Add str_key(Dates.Fix_IdEndSpot, int_FlowCnt_B), _
                lng_IdEndSpot_B
        End If


        'Index reference notional in Leg A

        If lng_EstStartSpot_B = lng_SwapStart Then
            lng_RefNotional_B = dic_dates_A(str_key(Dates.Fix_Start, 1))
        ElseIf Examine_FindIndex(lng_DateSchedule_A, lng_EstStartSpot_B) <> -1 Then
            int_RefNotionalPos = Examine_FindIndex(lng_DateSchedule_A, lng_EstStartSpot_B) + 1
            lng_RefNotional_B = dic_dates_A(str_key(Dates.Fix_Start, int_RefNotionalPos))
        End If

        dic_dates_B.Add str_key(Dates.Fix_NotionalRef, int_FlowCnt_B), _
            lng_RefNotional_B

    Next int_FlowCnt_B

Else


    'Handle case when fixing date = calc start date
    For int_FlowCnt_A = 1 To int_NumFlow_A

        'EQ Calc Start
        If int_FlowCnt_A = 1 Then
            lng_CalcStart_A = lng_SwapStart
        Else
            lng_CalcStart_A = dic_dates_A(str_key(Dates.Calc_End, int_FlowCnt_A - 1))
        End If

        dic_dates_A.Add str_key(Dates.Calc_Start, int_FlowCnt_A), _
            lng_CalcStart_A


        'EQ Calc End &  Eq Pmt Date
        lng_CalcEnd_A = lng_DateSchedule_A(int_FlowCnt_A)
        dic_dates_A.Add str_key(Dates.Calc_End, int_FlowCnt_A), _
            lng_CalcEnd_A
        dic_dates_A.Add str_key(Dates.Pmt_Date, int_FlowCnt_A), _
            FillDate(lng_CalcEnd_A, int_EqSpotDays, cal_Pmt_A)

        'EQ Fix Start and Fix Start Spot and FX Spot Fix Start

        lng_EstStart_A = lng_CalcStart_A

        dic_dates_A.Add str_key(Dates.Fix_Start, int_FlowCnt_A), _
            lng_EstStart_A


        lng_EstStartSpot_A = FillDate(lng_EstStart_A, int_EqSpotDays, cal_Pmt_A)
        dic_dates_A.Add str_key(Dates.Fix_StartSpot, int_FlowCnt_A), _
            lng_EstStartSpot_A


        'EQ Fix End and Fix End Spot and FX Spot Fix End

        'lng_EstEndSpot_A = WorksheetFunction.Max(Date_NextCoupon(lng_EstStartSpot_A, _
                        str_freq_A, cal_Fix_A, 1, bln_EOM_A, str_BDC_A), lng_CalcEnd_A)

        lng_EstEnd_A = lng_CalcEnd_A
        dic_dates_A.Add str_key(Dates.Fix_End, int_FlowCnt_A), _
            lng_EstEnd_A

        lng_EstEndSpot_A = FillDate(lng_EstEnd_A, int_EqSpotDays, cal_Pmt_A)

        dic_dates_A.Add str_key(Dates.Fix_EndSpot, int_FlowCnt_A), _
            lng_EstEndSpot_A


    Next int_FlowCnt_A


    Dim lng_IndexSpotStartDate As Long


    If bln_IndexFix = False Then
        lng_IndexSpotStartDate = FillDate(lng_SwapStart, int_IdSpotDays, cal_Fix_B)
        lng_DateSchedule_B_Fix = Date_CouponSchedule(lng_IndexSpotStartDate, str_GTerm, cal_Fix_B, str_freq_B, _
                            str_BDC_B, bln_EOM_B)
    End If

    For int_FlowCnt_B = 1 To int_NumFlow_B

        'Index Calc Start
        If int_FlowCnt_B = 1 Then
            lng_CalcStart_B = lng_SwapStart
        Else
            lng_CalcStart_B = dic_dates_B(str_key(Dates.Calc_End, int_FlowCnt_B - 1))
        End If

        dic_dates_B.Add str_key(Dates.Calc_Start, int_FlowCnt_B), _
            lng_CalcStart_B


        'Index Calc End &  Index Pmt Date
        lng_CalcEnd_B = lng_DateSchedule_B(int_FlowCnt_B)
        dic_dates_B.Add str_key(Dates.Calc_End, int_FlowCnt_B), _
            lng_CalcEnd_B
        dic_dates_B.Add str_key(Dates.Pmt_Date, int_FlowCnt_B), _
            FillDate(lng_CalcEnd_B, int_EqSpotDays, cal_Pmt_A)

        'Index Fix Start and Fix Start Spot and FX Spot Fix Start

        lng_EstStart_B = lng_CalcStart_B
        dic_dates_B.Add str_key(Dates.Fix_Start, int_FlowCnt_B), _
            lng_EstStart_B

        lng_EstStartSpot_B = FillDate(lng_EstStart_B, int_EqSpotDays, cal_Pmt_A)
        dic_dates_B.Add str_key(Dates.Fix_StartSpot, int_FlowCnt_B), _
            lng_EstStartSpot_B


        If bln_IndexFix = False Then
            lng_IdStartSpot_B = FillDate(lng_EstStart_B, int_IdSpotDays, cal_Fix_B)
            dic_dates_B.Add str_key(Dates.Fix_IdStartSpot, int_FlowCnt_B), _
                lng_IdStartSpot_B
        End If

        'Index Fix End and Fix End Spot and FX Spot Fix End

        If bln_IndexFix = False Then

            lng_EstEnd_B = Date_ApplyBDC(lng_CalcEnd_B, str_BDC_B, cal_Fix_B.HolDates, cal_Fix_B.Weekends)

            dic_dates_B.Add str_key(Dates.Fix_End, int_FlowCnt_B), _
                lng_EstEnd_B

            lng_EstEndSpot_B = FillDate(lng_EstEnd_B, int_EqSpotDays, cal_Pmt_A)
            dic_dates_B.Add str_key(Dates.Fix_EndSpot, int_FlowCnt_B), _
                lng_EstEndSpot_B

            lng_IdEndSpot_B = lng_DateSchedule_B_Fix(int_FlowCnt_B)

            dic_dates_B.Add str_key(Dates.Fix_IdEndSpot, int_FlowCnt_B), _
                lng_IdEndSpot_B

        End If

        'Index reference notional in Leg A

        If lng_EstStart_B = lng_SwapStart Then
            lng_RefNotional_B = dic_dates_A(str_key(Dates.Fix_Start, 1))
        ElseIf Examine_FindIndex(lng_DateSchedule_A, lng_EstStart_B) <> -1 Then
            int_RefNotionalPos = Examine_FindIndex(lng_DateSchedule_A, lng_EstStart_B) + 1
            lng_RefNotional_B = dic_dates_A(str_key(Dates.Fix_Start, int_RefNotionalPos))
        End If

        dic_dates_B.Add str_key(Dates.Fix_NotionalRef, int_FlowCnt_B), _
            lng_RefNotional_B

    Next int_FlowCnt_B

End If

If dic_CustomDate.count > 0 Then
    Call DateCustomization(dic_CustomDate)
End If

End Sub

Private Sub DateCustomization(dic_date As Dictionary)

Dim int_NumDate As Integer
Dim int_cnt As Integer


int_NumDate = dic_date.count / 4

For int_cnt = 1 To int_NumDate
    Call ModifiedDate(dic_date("ID|" & int_cnt), dic_date("NUM|" & int_cnt), _
        dic_date("TYPE|" & int_cnt), dic_date("DATE|" & int_cnt))
Next int_cnt

End Sub

Private Sub ModifiedDate(int_LegID As Integer, int_FlowNum As Integer, _
    str_type As String, lng_ModifiedDate As Long)
    Select Case int_LegID
        Case 1
            Select Case str_type
                Case "Calc Start"
                    dic_dates_A(str_key(Dates.Calc_Start, int_FlowNum)) = lng_ModifiedDate
                Case "Calc End"
                    dic_dates_A(str_key(Dates.Calc_End, int_FlowNum)) = lng_ModifiedDate
                Case "Fix Start"
                    dic_dates_A(str_key(Dates.Fix_Start, int_FlowNum)) = lng_ModifiedDate
                Case "Fix End"
                    dic_dates_A(str_key(Dates.Fix_End, int_FlowNum)) = lng_ModifiedDate
                Case "Pmt Date"
                    dic_dates_A(str_key(Dates.Pmt_Date, int_FlowNum)) = lng_ModifiedDate
                Case "Fix Start Spot"
                    dic_dates_A(str_key(Dates.Fix_StartSpot, int_FlowNum)) = lng_ModifiedDate
                Case "Fix End Spot"
                    dic_dates_A(str_key(Dates.Fix_EndSpot, int_FlowNum)) = lng_ModifiedDate
            End Select
        Case 2
            Select Case str_type
                Case "Calc Start"
                    dic_dates_B(str_key(Dates.Calc_Start, int_FlowNum)) = lng_ModifiedDate
                Case "Calc End"
                    dic_dates_B(str_key(Dates.Calc_End, int_FlowNum)) = lng_ModifiedDate
                Case "Fix Start"
                    dic_dates_B(str_key(Dates.Fix_Start, int_FlowNum)) = lng_ModifiedDate
                Case "Fix End"
                    dic_dates_B(str_key(Dates.Fix_End, int_FlowNum)) = lng_ModifiedDate
                Case "Pmt Date"
                    dic_dates_B(str_key(Dates.Pmt_Date, int_FlowNum)) = lng_ModifiedDate
'                Case "Fix Start Spot"
'                    dic_dates_B(str_key(Dates.Fix_IdStartSpot, int_FlowNum)) = lng_ModifiedDate
'                Case "Fix End Spot"
'                    dic_dates_B(str_key(Dates.Fix_IdEndSpot, int_FlowNum)) = lng_ModifiedDate
                Case "Fix Start Spot"
                    dic_dates_B(str_key(Dates.Fix_IdStartSpot, int_FlowNum)) = lng_ModifiedDate
                Case "Fix End Spot"
                    dic_dates_B(str_key(Dates.Fix_IdEndSpot, int_FlowNum)) = lng_ModifiedDate


            End Select
    End Select
End Sub


Private Sub GenerateFlows()


dic_flow_A.RemoveAll
dic_flow_B.RemoveAll
dic_output.RemoveAll
dic_Div.RemoveAll

Dim int_FlowCnt_A As Integer, int_FlowCnt_B As Integer
Dim dbl_eqspot As Double: dbl_eqspot = EqSpot

'lng_EqFixStartDate is the spot date for equity

Dim lng_FixStartDate As Long, lng_FixEndDate As Long
Dim lng_EqFixStartDate As Long, lng_EqFixEndDate As Long, lng_PmtDate As Long
Dim lng_IdFixStartSpotDate As Long, lng_IdFixEndSpotDate As Long
Dim lng_IdCalcStart As Long, lng_IdCalcEnd As Long


Dim dbl_EqStart As Double, dbl_EqEnd As Double
Dim dbl_FxStart As Double, dbl_FxEnd As Double
Dim dbl_EqFixStart As Double, dbl_EqFixEnd As Double

Dim dbl_IdNotional As Double

Dim dbl_rate As Double, dbl_flow As Double, dbl_DiscFlow As Double
Dim dbl_DiscFlowPnL As Double, dbl_DF As Double

Dim dbl_NumShare As Double, dbl_notional As Double

Dim str_FlowType As String

Dim dic_EqNotional As New Dictionary, dic_EqNumShare As New Dictionary

dbl_NumShare = dbl_OriNumShare
dbl_notional = dbl_OriNotional

Dim dbl_CASH As Double, dbl_MV As Double, dbl_PnL As Double


dbl_CASH = 0
dbl_MV = 0
dbl_PnL = 0

For int_FlowCnt_A = 1 To int_NumFlow_A

    lng_FixStartDate = dic_dates_A(str_key(Dates.Fix_Start, int_FlowCnt_A))
    lng_FixEndDate = dic_dates_A(str_key(Dates.Fix_End, int_FlowCnt_A))

    lng_EqFixStartDate = dic_dates_A(str_key(Dates.Fix_StartSpot, int_FlowCnt_A))
    lng_EqFixEndDate = dic_dates_A(str_key(Dates.Fix_EndSpot, int_FlowCnt_A))
    lng_PmtDate = dic_dates_A(str_key(Dates.Pmt_Date, int_FlowCnt_A))

    If lng_FixStartDate < lng_ValDate Then
        dbl_EqStart = dic_EqFixing(lng_FixStartDate)
    Else
        If bln_DivExist = True Then
            If lng_FutDivPmt >= lng_SpotDate And lng_FutDivPmt <= lng_EqFixStartDate And lng_ExDivDate > lng_ValDate Then
                dbl_EqStart = (dbl_eqspot / _
                irc_Est_A.Lookup_Rate(lng_SpotDate, lng_FutDivPmt, "DF") - dbl_FutDiv) / _
                irc_Est_A.Lookup_Rate(lng_FutDivPmt, lng_EqFixStartDate, "DF")
            Else
                dbl_EqStart = dbl_eqspot / irc_Est_A.Lookup_Rate(lng_SpotDate, lng_EqFixStartDate, "DF")
            End If
        Else
            dbl_EqStart = dbl_eqspot / irc_Est_A.Lookup_Rate(lng_SpotDate, lng_EqFixStartDate, "DF")
        End If
    End If


    If lng_FixEndDate < lng_ValDate Then
        dbl_EqEnd = dic_EqFixing(lng_FixEndDate)
    Else
        If bln_DivExist = True Then
            If lng_FutDivPmt >= lng_SpotDate And lng_FutDivPmt <= lng_EqFixEndDate And lng_ExDivDate > lng_ValDate Then
                dbl_EqEnd = (dbl_eqspot / _
                irc_Est_A.Lookup_Rate(lng_SpotDate, lng_FutDivPmt, "DF") - dbl_FutDiv) / _
                irc_Est_A.Lookup_Rate(lng_FutDivPmt, lng_EqFixEndDate, "DF")
            Else
                dbl_EqEnd = dbl_eqspot / irc_Est_A.Lookup_Rate(lng_SpotDate, lng_EqFixEndDate, "DF")
            End If
        Else
            dbl_EqEnd = dbl_eqspot / irc_Est_A.Lookup_Rate(lng_SpotDate, lng_EqFixEndDate, "DF")
        End If


    End If

    If str_Ccy_Eq = str_Ccy_Fix Then
        dbl_FxStart = 1
        dbl_FxEnd = 1
    Else
        If lng_FixStartDate < lng_ValDate Then
            dbl_FxStart = dic_FxFixing(lng_FixStartDate)
        Else
            dbl_FxStart = fxs_Spots.Lookup_Fwd(str_Ccy_Eq, str_Ccy_Fix, lng_FixStartDate)
        End If

        If lng_FixEndDate < lng_ValDate Then
            dbl_FxEnd = dic_FxFixing(lng_FixEndDate)
        Else
            dbl_FxEnd = fxs_Spots.Lookup_Fwd(str_Ccy_Eq, str_Ccy_Fix, lng_FixEndDate)
        End If
    End If

    'If lng_FixEndDate < lng_ValDate Then
    If lng_PmtDate < lng_ValDate Then
        str_FlowType = "CASH"
    Else
        str_FlowType = "MV"
    End If

    dbl_EqFixStart = dbl_EqStart * dbl_FxStart
    dbl_EqFixEnd = dbl_EqEnd * dbl_FxEnd

    If dbl_EqFixStart <> 0 Then
        'dbl_rate = ((dbl_EqFixEnd / dbl_EqFixStart) - 1) * 100
        dbl_rate = (dbl_EqFixEnd - dbl_EqFixStart) / Abs(dbl_EqFixStart) * 100

    Else
        dbl_rate = 0
    End If

    If str_ConstType = "SHARE" Then
        dbl_notional = dbl_NumShare * dbl_EqFixStart
    ElseIf str_ConstType = "AMOUNT" Then
        dbl_NumShare = dbl_notional / dbl_EqFixStart
    End If




    dic_EqNotional.Add lng_FixStartDate, dbl_notional
    dic_EqNumShare.Add lng_FixStartDate, dbl_NumShare

    dbl_flow = int_sign_A * dbl_notional * dbl_rate / 100

    If lng_PmtDate < lng_ValDate Then
        dbl_DF = 1
    Else
        dbl_DF = irc_Disc_A.Lookup_Rate(lng_ValDate, lng_PmtDate, "DF")
    End If



    dbl_DiscFlow = dbl_flow * dbl_DF
    dbl_DiscFlowPnL = dbl_DiscFlow * FXConvFactor()




    dic_flow_A.Add str_key(EqLeg.EqStart, int_FlowCnt_A), dbl_EqStart
    dic_flow_A.Add str_key(EqLeg.EqEnd, int_FlowCnt_A), dbl_EqEnd
    dic_flow_A.Add str_key(EqLeg.FxStart, int_FlowCnt_A), dbl_FxStart
    dic_flow_A.Add str_key(EqLeg.FxEnd, int_FlowCnt_A), dbl_FxEnd
    dic_flow_A.Add str_key(EqLeg.EqStartFix, int_FlowCnt_A), dbl_EqFixStart
    dic_flow_A.Add str_key(EqLeg.EqEndFix, int_FlowCnt_A), dbl_EqFixEnd
    dic_flow_A.Add str_key(EqLeg.NumShare, int_FlowCnt_A), dbl_NumShare
    dic_flow_A.Add str_key(EqLeg.Notional, int_FlowCnt_A), dbl_notional
    dic_flow_A.Add str_key(EqLeg.Rate, int_FlowCnt_A), dbl_rate
    dic_flow_A.Add str_key(EqLeg.DF, int_FlowCnt_A), dbl_DF

    dic_flow_A.Add str_key(EqLeg.FlowType, int_FlowCnt_A), str_FlowType
    dic_flow_A.Add str_key(EqLeg.Flow, int_FlowCnt_A), dbl_flow
    dic_flow_A.Add str_key(EqLeg.DiscFlow, int_FlowCnt_A), dbl_DiscFlow
    dic_flow_A.Add str_key(EqLeg.DiscFlowPnL, int_FlowCnt_A), dbl_DiscFlowPnL

    If str_FlowType = "CASH" Then
        dbl_CASH = dbl_CASH + dbl_DiscFlowPnL
    ElseIf str_FlowType = "MV" Then
        dbl_MV = dbl_MV + dbl_DiscFlowPnL
    End If

    If lng_PmtDate >= lng_ValDate And dic_output.Exists(Output.EQ_NOP) = False Then
        dic_output.Add Output.EQ_NOP, int_sign_A * dbl_notional * fxs_Spots.Lookup_DiscSpot(str_Ccy_Fix, str_CCY_PnL)
    End If

Next int_FlowCnt_A

dbl_PnL = dbl_MV + dbl_CASH
dic_flow_A.Add Output.SwapCash, dbl_CASH
dic_flow_A.Add Output.SwapMV, dbl_MV
dic_flow_A.Add Output.SwapPnL, dbl_PnL








dbl_CASH = 0
dbl_MV = 0
dbl_PnL = 0

Dim dbl_NotionalMultiplier As Double
dbl_NotionalMultiplier = int_NumFlow_A / int_NumFlow_B
Dim int_NotCnt As Integer
Dim dbl_AdjNot As Double
Dim int_OriNotPos As Integer

For int_FlowCnt_B = 1 To int_NumFlow_B
    lng_IdCalcStart = dic_dates_B(str_key(Dates.Calc_Start, int_FlowCnt_B))
    lng_IdCalcEnd = dic_dates_B(str_key(Dates.Calc_End, int_FlowCnt_B))


    lng_PmtDate = dic_dates_B(str_key(Dates.Pmt_Date, int_FlowCnt_B))


    lng_FixStartDate = dic_dates_B(str_key(Dates.Fix_Start, int_FlowCnt_B))
    lng_FixEndDate = dic_dates_B(str_key(Dates.Fix_End, int_FlowCnt_B))
    lng_IdFixStartSpotDate = dic_dates_B(str_key(Dates.Fix_IdStartSpot, int_FlowCnt_B))
    lng_IdFixEndSpotDate = dic_dates_B(str_key(Dates.Fix_IdEndSpot, int_FlowCnt_B))



    'dbl_IdNotional = dic_EqNotional(lng_FixStartDate)
    'dbl_IdNotional = dic_EqNotional(dic_dates_B(str_key(Dates.Fix_NotionalRef, int_FlowCnt_B)))
    dbl_IdNotional = dic_EqNotional(dic_EqNotional.Keys(int_FlowCnt_B - 1))


    If bln_IndexFix = True Then
        dbl_rate = dbl_RateOrMagin
    Else
        If lng_FixStartDate < lng_ValDate Then
            dbl_rate = dic_IdFixing(lng_FixStartDate) + dbl_RateOrMagin
        Else
            dbl_rate = irc_Est_B.Lookup_Rate(lng_IdFixStartSpotDate, lng_IdFixEndSpotDate, str_IdDayCnt) + dbl_RateOrMagin
        End If
    End If

    If lng_PmtDate < lng_ValDate Then
        str_FlowType = "CASH"
        dbl_DF = 1
    Else
        str_FlowType = "MV"
        dbl_DF = irc_Disc_B.Lookup_Rate(lng_ValDate, lng_PmtDate, "DF")
    End If

    If dbl_NotionalMultiplier > 1 Then
        dbl_AdjNot = 0

        For int_NotCnt = 1 To CInt(dbl_NotionalMultiplier)
            int_OriNotPos = dbl_NotionalMultiplier * (int_FlowCnt_B - 1) + int_NotCnt
            dbl_AdjNot = dbl_AdjNot + dic_flow_A(str_key(EqLeg.Notional, int_OriNotPos)) * _
                calc_yearfrac(dic_dates_A(str_key(Dates.Calc_Start, int_OriNotPos)), _
                dic_dates_A(str_key(Dates.Calc_End, int_OriNotPos)), str_IdDayCnt)
        Next int_NotCnt
        dbl_AdjNot = dbl_AdjNot / calc_yearfrac(lng_IdCalcStart, lng_IdCalcEnd, str_IdDayCnt)
        dbl_IdNotional = dbl_AdjNot
    End If

    dbl_flow = int_sign_B * dbl_IdNotional * dbl_rate / 100 * _
                calc_yearfrac(lng_IdCalcStart, lng_IdCalcEnd, str_IdDayCnt)

    dbl_DiscFlow = dbl_flow * dbl_DF
    dbl_DiscFlowPnL = dbl_DiscFlow * FXConvFactor()


    dic_flow_B.Add str_key(IdLeg.Notional, int_FlowCnt_B), dbl_IdNotional
    dic_flow_B.Add str_key(IdLeg.Rate, int_FlowCnt_B), dbl_rate
    dic_flow_B.Add str_key(IdLeg.DF, int_FlowCnt_B), dbl_DF

    dic_flow_B.Add str_key(IdLeg.FlowType, int_FlowCnt_B), str_FlowType
    dic_flow_B.Add str_key(IdLeg.Flow, int_FlowCnt_B), dbl_flow
    dic_flow_B.Add str_key(IdLeg.DiscFlow, int_FlowCnt_B), dbl_DiscFlow
    dic_flow_B.Add str_key(IdLeg.DiscFlowPnL, int_FlowCnt_B), dbl_DiscFlowPnL

    If str_FlowType = "CASH" Then
        dbl_CASH = dbl_CASH + dbl_DiscFlowPnL
    ElseIf str_FlowType = "MV" Then
        dbl_MV = dbl_MV + dbl_DiscFlowPnL
    End If

Next int_FlowCnt_B

dbl_PnL = dbl_MV + dbl_CASH

dic_flow_B.Add Output.SwapCash, dbl_CASH
dic_flow_B.Add Output.SwapMV, dbl_MV
dic_flow_B.Add Output.SwapPnL, dbl_PnL

Dim int_DivCnt As Integer
Dim lng_ExDiv As Long, lng_DivPmt As Long, lng_EqsDivPmt As Long
Dim dbl_Div As Double


If bln_TotalRet = True And bln_DivExist = True Then

    For int_DivCnt = 1 To int_NumDiv

        lng_ExDiv = dic_DivInfo("EX|" & int_DivCnt)
        dbl_Div = dic_DivInfo("DIV|" & int_DivCnt)
        lng_DivPmt = dic_DivInfo("PMT|" & int_DivCnt)
        lng_EqsDivPmt = FillDate(lng_DivPmt, dic_DivInfo("SHIFT|" & int_DivCnt), cal_None)

        dic_Div.Add str_key(DivLeg.EqsDivPmt, int_DivCnt), lng_EqsDivPmt

        If lng_ExDiv <= dic_dates_A(str_key(Dates.Fix_End, int_NumFlow_A)) Then
            For int_FlowCnt_A = 1 To int_NumFlow_A
                If dic_dates_A(str_key(Dates.Fix_End, int_FlowCnt_A)) >= lng_ExDiv Then
                    dic_Div.Add str_key(DivLeg.NumOfShare, int_DivCnt), _
                    dic_EqNumShare(dic_dates_A(str_key(Dates.Fix_Start, int_FlowCnt_A)))
                    Exit For
                End If
            Next int_FlowCnt_A
        End If


        dic_Div.Add str_key(DivLeg.Div, int_DivCnt), dbl_Div

        If str_Ccy_Eq = str_Ccy_Fix Then
            dic_Div.Add DivLeg.fx, 1
        Else
            If lng_EqsDivPmt < lng_ValDate Then
                dic_Div.Add str_key(DivLeg.fx, int_DivCnt), dic_FxFixing(lng_EqsDivPmt)
            Else
                dic_Div.Add str_key(DivLeg.fx, int_DivCnt), fxs_Spots.Lookup_Fwd(str_Ccy_Eq, str_Ccy_Fix, lng_EqsDivPmt)
            End If
        End If

        dic_Div.Add str_key(DivLeg.FixDiv, int_DivCnt), dic_Div(str_key(DivLeg.Div, int_DivCnt)) * _
                                    dic_Div(str_key(DivLeg.fx, int_DivCnt))

        If lng_EqsDivPmt >= lng_ValDate Then
            dic_Div.Add str_key(DivLeg.DF, int_DivCnt), irc_Disc_A.Lookup_Rate(lng_ValDate, lng_EqsDivPmt, "DF")
        Else
            dic_Div.Add str_key(DivLeg.DF, int_DivCnt), 1
        End If

        dic_Div.Add str_key(DivLeg.Flow, int_DivCnt), int_sign_A * dic_Div(str_key(DivLeg.NumOfShare, int_DivCnt)) * dic_Div(str_key(DivLeg.FixDiv, int_DivCnt))
        dic_Div.Add str_key(DivLeg.DiscFlow, int_DivCnt), dic_Div(str_key(DivLeg.Flow, int_DivCnt)) * dic_Div(str_key(DivLeg.DF, int_DivCnt))
        dic_Div.Add str_key(DivLeg.DiscFlowPnL, int_DivCnt), dic_Div(str_key(DivLeg.DiscFlow, int_DivCnt)) * FXConvFactor()

        If lng_EqsDivPmt >= lng_ValDate Then
            dic_Div.Add str_key(DivLeg.DiscFlowMV, int_DivCnt), dic_Div(str_key(DivLeg.DiscFlowPnL, int_DivCnt))
            dic_Div.Add str_key(DivLeg.FlowType, int_DivCnt), "MV"
        Else
            dic_Div.Add str_key(DivLeg.DiscFlowCash, int_DivCnt), dic_Div(str_key(DivLeg.DiscFlowPnL, int_DivCnt))
            dic_Div.Add str_key(DivLeg.FlowType, int_DivCnt), "CASH"
        End If

        If dic_Div.Exists(DivLeg.DiscFlowCash) = False Then
            dic_Div.Add DivLeg.DiscFlowCash, dic_Div(str_key(DivLeg.DiscFlowCash, int_DivCnt))
        Else
            dic_Div(DivLeg.DiscFlowCash) = dic_Div(DivLeg.DiscFlowCash) + dic_Div(str_key(DivLeg.DiscFlowCash, int_DivCnt))
        End If

        If dic_Div.Exists(DivLeg.DiscFlowMV) = False Then
            dic_Div.Add DivLeg.DiscFlowMV, dic_Div(str_key(DivLeg.DiscFlowMV, int_DivCnt))
        Else
            dic_Div(DivLeg.DiscFlowMV) = dic_Div(DivLeg.DiscFlowMV) + dic_Div(str_key(DivLeg.DiscFlowMV, int_DivCnt))
        End If

        If dic_Div.Exists(DivLeg.DiscFlowPnL) = False Then
            dic_Div.Add DivLeg.DiscFlowPnL, dic_Div(str_key(DivLeg.DiscFlowPnL, int_DivCnt))
        Else
            dic_Div(DivLeg.DiscFlowPnL) = dic_Div(DivLeg.DiscFlowPnL) + dic_Div(str_key(DivLeg.DiscFlowPnL, int_DivCnt))
        End If

    Next int_DivCnt
End If

dic_output.Add Output.SwapCash, dic_flow_A(Output.SwapCash) + dic_flow_B(Output.SwapCash) + dic_Div(DivLeg.DiscFlowCash)
dic_output.Add Output.SwapMV, dic_flow_A(Output.SwapMV) + dic_flow_B(Output.SwapMV) + dic_Div(DivLeg.DiscFlowMV)
dic_output.Add Output.SwapPnL, dic_flow_A(Output.SwapPnL) + dic_flow_B(Output.SwapPnL) + dic_Div(DivLeg.DiscFlowPnL)

End Sub
' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Sub SetValDate(lng_Input As Long)
    ' ## Set stored value date and refresh values dependent on the value date
    lng_ValDate = lng_Input
    Call FillSpotDate
End Sub

Public Sub OutputReport(wks_output As Worksheet)

    wks_output.Cells.Clear

    Call GenerateFlows



    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    Dim str_CurrencyFormat As String: str_CurrencyFormat = Gather_CurrencyFormat()

    'Summary

    Dim rng_SummaryTopLeft As Range: Set rng_SummaryTopLeft = wks_output.Range("A1")
    Dim int_RowCnt As Integer: int_RowCnt = 0

    rng_SummaryTopLeft.Value = "Overall"
    rng_SummaryTopLeft.Font.Bold = True

    int_RowCnt = int_RowCnt + 1
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
    rng_SummaryTopLeft.Value = "Valuation Date"
    rng_SummaryTopLeft.Offset(0, 1).Value = lng_ValDate
    rng_SummaryTopLeft.Offset(0, 1).NumberFormat = str_DateFormat

    int_RowCnt = int_RowCnt + 1
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
    rng_SummaryTopLeft.Value = "Swap Start"
    rng_SummaryTopLeft.Offset(0, 1).Value = lng_SwapStart
    rng_SummaryTopLeft.Offset(0, 1).NumberFormat = str_DateFormat

    int_RowCnt = int_RowCnt + 1
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
    rng_SummaryTopLeft.Value = "PnL Ccy"
    rng_SummaryTopLeft.Offset(0, 1).Value = str_CCY_PnL
    rng_SummaryTopLeft.Offset(0, 1).NumberFormat = str_DateFormat

    int_RowCnt = int_RowCnt + 1
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
    rng_SummaryTopLeft.Value = "Cash"
    rng_SummaryTopLeft.Offset(0, 1).Value = dic_output(Output.SwapCash)
    rng_SummaryTopLeft.Offset(0, 1).NumberFormat = str_CurrencyFormat

    int_RowCnt = int_RowCnt + 1
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
    rng_SummaryTopLeft.Value = "MV"
    rng_SummaryTopLeft.Offset(0, 1).Value = dic_output(Output.SwapMV)
    rng_SummaryTopLeft.Offset(0, 1).NumberFormat = str_CurrencyFormat

    int_RowCnt = int_RowCnt + 1
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
    rng_SummaryTopLeft.Value = "PnL"
    rng_SummaryTopLeft.Offset(0, 1).Value = dic_output(Output.SwapPnL)
    rng_SummaryTopLeft.Offset(0, 1).NumberFormat = str_CurrencyFormat



    int_RowCnt = int_RowCnt + 2
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(2, 0)
    rng_SummaryTopLeft.Value = "Eq Leg A"
    rng_SummaryTopLeft.Font.Bold = True

    int_RowCnt = int_RowCnt + 1
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
    rng_SummaryTopLeft.Value = "Equity"
    rng_SummaryTopLeft.Offset(0, 1).Value = str_SecCode

    int_RowCnt = int_RowCnt + 1
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
    rng_SummaryTopLeft.Value = "Equity Payer?"
    rng_SummaryTopLeft.Offset(0, 1).Value = bln_EqPayer

    int_RowCnt = int_RowCnt + 1
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
    rng_SummaryTopLeft.Value = "Equity Ccy"
    rng_SummaryTopLeft.Offset(0, 1).Value = str_Ccy_Eq


    int_RowCnt = int_RowCnt + 1
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
    rng_SummaryTopLeft.Value = "Equity Fix Ccy"
    rng_SummaryTopLeft.Offset(0, 1).Value = str_Ccy_Fix


    int_RowCnt = int_RowCnt + 1
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
    rng_SummaryTopLeft.Value = "Equity Freq"
    rng_SummaryTopLeft.Offset(0, 1).Value = str_freq_A


    int_RowCnt = int_RowCnt + 1
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
    rng_SummaryTopLeft.Value = "Equity Cash"
    rng_SummaryTopLeft.Offset(0, 1).Value = dic_flow_A(Output.SwapCash)
    rng_SummaryTopLeft.Offset(0, 1).NumberFormat = str_CurrencyFormat

    int_RowCnt = int_RowCnt + 1
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
    rng_SummaryTopLeft.Value = "Equity MV"
    rng_SummaryTopLeft.Offset(0, 1).Value = dic_flow_A(Output.SwapMV)
    rng_SummaryTopLeft.Offset(0, 1).NumberFormat = str_CurrencyFormat

    int_RowCnt = int_RowCnt + 1
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
    rng_SummaryTopLeft.Value = "Equity PnL"
    rng_SummaryTopLeft.Offset(0, 1).Value = dic_flow_A(Output.SwapPnL)
    rng_SummaryTopLeft.Offset(0, 1).NumberFormat = str_CurrencyFormat

    int_RowCnt = int_RowCnt + 1
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(2, 0)
    rng_SummaryTopLeft.Value = "Index Leg B"
    rng_SummaryTopLeft.Font.Bold = True


    int_RowCnt = int_RowCnt + 1
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
    rng_SummaryTopLeft.Value = "Index Ccy"
    rng_SummaryTopLeft.Offset(0, 1).Value = str_Ccy_Id


    int_RowCnt = int_RowCnt + 1
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
    rng_SummaryTopLeft.Value = "Index Freq"
    rng_SummaryTopLeft.Offset(0, 1).Value = str_freq_B


    int_RowCnt = int_RowCnt + 1
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
    rng_SummaryTopLeft.Value = "Index Cash"
    rng_SummaryTopLeft.Offset(0, 1).Value = dic_flow_B(Output.SwapCash)
    rng_SummaryTopLeft.Offset(0, 1).NumberFormat = str_CurrencyFormat

    int_RowCnt = int_RowCnt + 1
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
    rng_SummaryTopLeft.Value = "Index MV"
    rng_SummaryTopLeft.Offset(0, 1).Value = dic_flow_B(Output.SwapMV)
    rng_SummaryTopLeft.Offset(0, 1).NumberFormat = str_CurrencyFormat

    int_RowCnt = int_RowCnt + 1
    Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
    rng_SummaryTopLeft.Value = "Index PnL"
    rng_SummaryTopLeft.Offset(0, 1).Value = dic_flow_B(Output.SwapPnL)
    rng_SummaryTopLeft.Offset(0, 1).NumberFormat = str_CurrencyFormat

    If bln_TotalRet = True And bln_DivExist = True Then

        int_RowCnt = int_RowCnt + 2
        Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(2, 0)
        rng_SummaryTopLeft.Value = "Div Leg C"
        rng_SummaryTopLeft.Font.Bold = True

        int_RowCnt = int_RowCnt + 1
        Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
        rng_SummaryTopLeft.Value = "Div Cash"
        rng_SummaryTopLeft.Offset(0, 1).Value = dic_Div(DivLeg.DiscFlowCash)
        rng_SummaryTopLeft.Offset(0, 1).NumberFormat = str_CurrencyFormat

        int_RowCnt = int_RowCnt + 1
        Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
        rng_SummaryTopLeft.Value = "Div MV"
        rng_SummaryTopLeft.Offset(0, 1).Value = dic_Div(DivLeg.DiscFlowMV)
        rng_SummaryTopLeft.Offset(0, 1).NumberFormat = str_CurrencyFormat

        int_RowCnt = int_RowCnt + 1
        Set rng_SummaryTopLeft = rng_SummaryTopLeft.Offset(1, 0)
        rng_SummaryTopLeft.Value = "Div PnL"
        rng_SummaryTopLeft.Offset(0, 1).Value = dic_Div(DivLeg.DiscFlowPnL)
        rng_SummaryTopLeft.Offset(0, 1).NumberFormat = str_CurrencyFormat
    End If

    'Leg Information BreakDown


    Dim rng_OutputTopLeft As Range: Set rng_OutputTopLeft = rng_SummaryTopLeft.Offset(5, 0)
    Dim int_FlowCnt_A As Integer, int_FlowCnt_B As Integer, int_FlowCnt_C As Integer

    Dim rng_output_A As Range, rng_output_B As Range, rng_output_C As Range
    Dim int_LegSpace As Integer: int_LegSpace = 5

    Dim int_row_TitleA As Integer: int_row_TitleA = 1
    Dim int_row_outputA As Integer: int_row_outputA = int_row_TitleA + 1
    Dim int_row_TitleB As Integer: int_row_TitleB = int_NumFlow_A + int_LegSpace + 3
    Dim int_row_outputB As Integer: int_row_outputB = int_row_TitleB + 1


    If bln_TotalRet = True And bln_DivExist = True Then
        Dim int_row_TitleC As Integer: int_row_TitleC = int_row_TitleB + 3 + int_LegSpace + 3
        Dim int_row_outputC As Integer: int_row_outputC = int_row_TitleC + 1
    End If




    rng_OutputTopLeft.Value = "Eq Leg A"
    rng_OutputTopLeft.Font.Bold = True

    rng_OutputTopLeft.Offset(int_row_TitleB - 1, 0).Value = "Index Leg B"
    rng_OutputTopLeft.Offset(int_row_TitleB - 1, 0).Font.Bold = True


    If bln_TotalRet = True And bln_DivExist = True Then
        rng_OutputTopLeft.Offset(int_row_TitleC - 1, 0).Value = "Div Leg C"
        rng_OutputTopLeft.Offset(int_row_TitleC - 1, 0).Font.Bold = True
    End If


    Set rng_output_A = rng_OutputTopLeft.Offset(int_row_outputA, 0)
    Set rng_output_B = rng_OutputTopLeft.Offset(int_row_outputB, 0)

    If bln_TotalRet = True And bln_DivExist = True Then
        Set rng_output_C = rng_OutputTopLeft.Offset(int_row_outputC, 0)
    End If

    Dim int_ColCnt As Integer
    Dim int_col_Date_A As Integer, int_col_rate_A As Integer
    Dim int_col_Date_B As Integer, int_col_rate_B As Integer
    Dim int_col_Date_C As Integer, int_col_rate_C As Integer

    For int_FlowCnt_A = 1 To int_NumFlow_A

        int_ColCnt = 0

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "Calc Start"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_dates_A(str_key(Dates.Calc_Start, int_FlowCnt_A))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "Calc End"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_dates_A(str_key(Dates.Calc_End, int_FlowCnt_A))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "Fix Start"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_dates_A(str_key(Dates.Fix_Start, int_FlowCnt_A))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "Fix Start Spot"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_dates_A(str_key(Dates.Fix_StartSpot, int_FlowCnt_A))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "Fix End"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_dates_A(str_key(Dates.Fix_End, int_FlowCnt_A))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "Fix End Spot"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_dates_A(str_key(Dates.Fix_EndSpot, int_FlowCnt_A))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "Pmt Date"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_dates_A(str_key(Dates.Pmt_Date, int_FlowCnt_A))

        int_col_Date_A = int_ColCnt

        int_ColCnt = int_ColCnt + 2
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "Eq Start"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_flow_A(str_key(EqLeg.EqStart, int_FlowCnt_A))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "Eq End"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_flow_A(str_key(EqLeg.EqEnd, int_FlowCnt_A))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "Fx Start"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_flow_A(str_key(EqLeg.FxStart, int_FlowCnt_A))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "Fx End"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_flow_A(str_key(EqLeg.FxEnd, int_FlowCnt_A))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "Eq Fix Start"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_flow_A(str_key(EqLeg.EqStartFix, int_FlowCnt_A))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "Eq Fix End"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_flow_A(str_key(EqLeg.EqEndFix, int_FlowCnt_A))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "No. Share"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_flow_A(str_key(EqLeg.NumShare, int_FlowCnt_A))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "Notional"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_flow_A(str_key(EqLeg.Notional, int_FlowCnt_A))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "Rate"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_flow_A(str_key(EqLeg.Rate, int_FlowCnt_A))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "DF"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_flow_A(str_key(EqLeg.DF, int_FlowCnt_A))


        int_col_rate_A = int_ColCnt

        int_ColCnt = int_ColCnt + 2
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "Flow Type"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_flow_A(str_key(EqLeg.FlowType, int_FlowCnt_A))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "Flow"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_flow_A(str_key(EqLeg.Flow, int_FlowCnt_A))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "Disc Flow"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_flow_A(str_key(EqLeg.DiscFlow, int_FlowCnt_A))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleA, int_ColCnt - 1).Value = "Disc Flow (PnL)"
        rng_output_A(int_FlowCnt_A, int_ColCnt) = dic_flow_A(str_key(EqLeg.DiscFlowPnL, int_FlowCnt_A))


    Next int_FlowCnt_A





    For int_FlowCnt_B = 1 To int_NumFlow_B

        int_ColCnt = 0

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleB, int_ColCnt - 1).Value = "Calc Start"
        rng_output_B(int_FlowCnt_B, int_ColCnt) = dic_dates_B(str_key(Dates.Calc_Start, int_FlowCnt_B))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleB, int_ColCnt - 1).Value = "Calc End"
        rng_output_B(int_FlowCnt_B, int_ColCnt) = dic_dates_B(str_key(Dates.Calc_End, int_FlowCnt_B))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleB, int_ColCnt - 1).Value = "Fix Start"
        rng_output_B(int_FlowCnt_B, int_ColCnt) = dic_dates_B(str_key(Dates.Fix_Start, int_FlowCnt_B))

'        int_ColCnt = int_ColCnt + 1
'        rng_OutputTopLeft.Offset(int_row_TitleB, int_ColCnt - 1).Value = "Fix Start Spot"
'        rng_output_B(int_FlowCnt_B, int_ColCnt) = dic_dates_B(str_key(Dates.Fix_StartSpot, int_FlowCnt_B))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleB, int_ColCnt - 1).Value = "Fix Id Start Spot"
        rng_output_B(int_FlowCnt_B, int_ColCnt) = dic_dates_B(str_key(Dates.Fix_IdStartSpot, int_FlowCnt_B))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleB, int_ColCnt - 1).Value = "Fix End"
        rng_output_B(int_FlowCnt_B, int_ColCnt) = dic_dates_B(str_key(Dates.Fix_End, int_FlowCnt_B))

'        int_ColCnt = int_ColCnt + 1
'        rng_OutputTopLeft.Offset(int_row_TitleB, int_ColCnt - 1).Value = "Fix End Spot"
'        rng_output_B(int_FlowCnt_B, int_ColCnt) = dic_dates_B(str_key(Dates.Fix_EndSpot, int_FlowCnt_B))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleB, int_ColCnt - 1).Value = "Fix Id End Spot"
        rng_output_B(int_FlowCnt_B, int_ColCnt) = dic_dates_B(str_key(Dates.Fix_IdEndSpot, int_FlowCnt_B))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleB, int_ColCnt - 1).Value = "Pmt Date"
        rng_output_B(int_FlowCnt_B, int_ColCnt) = dic_dates_B(str_key(Dates.Pmt_Date, int_FlowCnt_B))

'        int_ColCnt = int_ColCnt + 1
'        rng_OutputTopLeft.Offset(int_row_TitleB, int_ColCnt - 1).Value = "Notional Ref Fix Date"
'        rng_output_B(int_FlowCnt_B, int_ColCnt) = dic_dates_B(str_key(Dates.Fix_NotionalRef, int_FlowCnt_B))

        int_col_Date_B = int_ColCnt

        int_ColCnt = int_ColCnt + 2
        rng_OutputTopLeft.Offset(int_row_TitleB, int_ColCnt - 1).Value = "Notional"
        rng_output_B(int_FlowCnt_B, int_ColCnt) = dic_flow_B(str_key(IdLeg.Notional, int_FlowCnt_B))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleB, int_ColCnt - 1).Value = "Rate"
        rng_output_B(int_FlowCnt_B, int_ColCnt) = dic_flow_B(str_key(IdLeg.Rate, int_FlowCnt_B))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleB, int_ColCnt - 1).Value = "DF"
        rng_output_B(int_FlowCnt_B, int_ColCnt) = dic_flow_B(str_key(IdLeg.DF, int_FlowCnt_B))


        int_col_rate_B = int_ColCnt

        int_ColCnt = int_ColCnt + 2
        rng_OutputTopLeft.Offset(int_row_TitleB, int_ColCnt - 1).Value = "Flow Type"
        rng_output_B(int_FlowCnt_B, int_ColCnt) = dic_flow_B(str_key(IdLeg.FlowType, int_FlowCnt_B))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleB, int_ColCnt - 1).Value = "Flow"
        rng_output_B(int_FlowCnt_B, int_ColCnt) = dic_flow_B(str_key(IdLeg.Flow, int_FlowCnt_B))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleB, int_ColCnt - 1).Value = "Disc Flow"
        rng_output_B(int_FlowCnt_B, int_ColCnt) = dic_flow_B(str_key(IdLeg.DiscFlow, int_FlowCnt_B))

        int_ColCnt = int_ColCnt + 1
        rng_OutputTopLeft.Offset(int_row_TitleB, int_ColCnt - 1).Value = "Disc Flow (PnL)"
        rng_output_B(int_FlowCnt_B, int_ColCnt) = dic_flow_B(str_key(IdLeg.DiscFlowPnL, int_FlowCnt_B))

    Next int_FlowCnt_B

    If bln_TotalRet = True And bln_DivExist = True Then
        For int_FlowCnt_C = 1 To int_NumDiv

            int_ColCnt = 0

            int_ColCnt = int_ColCnt + 1
            rng_OutputTopLeft.Offset(int_row_TitleC, int_ColCnt - 1).Value = "Ex-Div Date"
            rng_output_C(int_FlowCnt_C, int_ColCnt).Value = dic_DivInfo("EX|" & int_FlowCnt_C)

            int_ColCnt = int_ColCnt + 1
            rng_OutputTopLeft.Offset(int_row_TitleC, int_ColCnt - 1).Value = "Div Pmt Date"
            rng_output_C(int_FlowCnt_C, int_ColCnt).Value = dic_DivInfo("PMT|" & int_FlowCnt_C)

            int_ColCnt = int_ColCnt + 1
            rng_OutputTopLeft.Offset(int_row_TitleC, int_ColCnt - 1).Value = "EQS Div Pmt Date"
            rng_output_C(int_FlowCnt_C, int_ColCnt).Value = dic_Div(str_key(DivLeg.EqsDivPmt, int_FlowCnt_C))

            int_ColCnt = int_ColCnt + 2
            rng_OutputTopLeft.Offset(int_row_TitleC, int_ColCnt - 1).Value = "Div Amt"
            rng_output_C(int_FlowCnt_C, int_ColCnt).Value = dic_DivInfo("DIV|" & int_FlowCnt_C)

            int_ColCnt = int_ColCnt + 1
            rng_OutputTopLeft.Offset(int_row_TitleC, int_ColCnt - 1).Value = "FX"
            rng_output_C(int_FlowCnt_C, int_ColCnt).Value = dic_Div(str_key(DivLeg.fx, int_FlowCnt_C))

            int_ColCnt = int_ColCnt + 1
            rng_OutputTopLeft.Offset(int_row_TitleC, int_ColCnt - 1).Value = "Converted Div Amt"
            rng_output_C(int_FlowCnt_C, int_ColCnt).Value = dic_Div(str_key(DivLeg.FixDiv, int_FlowCnt_C))

            int_ColCnt = int_ColCnt + 1
            rng_OutputTopLeft.Offset(int_row_TitleC, int_ColCnt - 1).Value = "Num of Share"
            rng_output_C(int_FlowCnt_C, int_ColCnt).Value = dic_Div(str_key(DivLeg.NumOfShare, int_FlowCnt_C))

            int_ColCnt = int_ColCnt + 1
            rng_OutputTopLeft.Offset(int_row_TitleC, int_ColCnt - 1).Value = "DF"
            rng_output_C(int_FlowCnt_C, int_ColCnt).Value = dic_Div(str_key(DivLeg.DF, int_FlowCnt_C))

            int_ColCnt = int_ColCnt + 1
            rng_OutputTopLeft.Offset(int_row_TitleC, int_ColCnt - 1).Value = "Flow Type"
            rng_output_C(int_FlowCnt_C, int_ColCnt).Value = dic_Div(str_key(DivLeg.FlowType, int_FlowCnt_C))

            int_col_rate_C = int_ColCnt

            int_ColCnt = int_ColCnt + 2
            rng_OutputTopLeft.Offset(int_row_TitleC, int_ColCnt - 1).Value = "Div Flow"
            rng_output_C(int_FlowCnt_C, int_ColCnt).Value = dic_Div(str_key(DivLeg.Flow, int_FlowCnt_C))

            int_ColCnt = int_ColCnt + 1
            rng_OutputTopLeft.Offset(int_row_TitleC, int_ColCnt - 1).Value = "Div Disc Flow"
            rng_output_C(int_FlowCnt_C, int_ColCnt).Value = dic_Div(str_key(DivLeg.DiscFlow, int_FlowCnt_C))

            int_ColCnt = int_ColCnt + 1
            rng_OutputTopLeft.Offset(int_row_TitleC, int_ColCnt - 1).Value = "Div Disc Flow (PnL)"
            rng_output_C(int_FlowCnt_C, int_ColCnt).Value = dic_Div(str_key(DivLeg.DiscFlowPnL, int_FlowCnt_C))

        Next int_FlowCnt_C


        rng_format(wks_output, rng_output_C).NumberFormat = str_DateFormat
        rng_format(wks_output, rng_output_C.Offset(0, int_col_rate_C + 1)).NumberFormat = str_CurrencyFormat

    End If

    rng_format(wks_output, rng_output_A).NumberFormat = str_DateFormat
    rng_format(wks_output, rng_output_B).NumberFormat = str_DateFormat


    rng_format(wks_output, rng_output_A.Offset(0, int_col_Date_A + 1)).NumberFormat = "General"
    rng_format(wks_output, rng_output_A.Offset(0, int_col_rate_A + 1)).NumberFormat = str_CurrencyFormat

    rng_format(wks_output, rng_output_B.Offset(0, int_col_Date_B + 1)).NumberFormat = "General"
    rng_format(wks_output, rng_output_B.Offset(0, int_col_rate_B + 1)).NumberFormat = str_CurrencyFormat

    wks_output.Calculate
    wks_output.Cells.HorizontalAlignment = xlCenter
    wks_output.Columns.AutoFit

End Sub