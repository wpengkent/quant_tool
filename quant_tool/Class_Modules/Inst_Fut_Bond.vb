Option Explicit

' ## MEMBER DATA
' Components
Private irl_Underlying_SpreadOn As IRLeg, irl_Underlying_SpreadOff As IRLeg

' Dependent curves
Private fxs_Spots As Data_FXSpots, irc_Disc As Data_IRCurve

' Static values
Private dic_GlobalStaticInfo As Dictionary, dic_CurveDependencies As Dictionary
Private fld_Params As InstParams_FBN
Private cal_pmt As Calendar, lng_FutMatSpot As Long
Private lng_PrevCpnDate As Long, dbl_UndAccrual As Double
Private int_Sign As Integer
Private str_PmtFreq As String, bln_EOM As Boolean, str_Daycount As String, str_CCY_Notional As String, str_DiscCurve As String


' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As InstParams_FBN, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing)

    ' Store independent static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    fld_Params = fld_ParamsInput
    Const str_Daycount_DV01 As String = "ACT/365"
    If fld_Params.IsBuy = True Then int_Sign = 1 Else int_Sign = -1

    ' Gather generator level information and store related static values
    Dim igs_Generators As IRGeneratorSet: Set igs_Generators = dic_GlobalStaticInfo(StaticInfoType.IRGeneratorSet)
    Dim fld_UndParams As IRLegParams: fld_UndParams = igs_Generators.Lookup_Generator(fld_Params.Generator)
    str_PmtFreq = fld_UndParams.PmtFreq
    bln_EOM = fld_UndParams.EOM
    str_Daycount = fld_UndParams.Daycount
    str_CCY_Notional = fld_UndParams.CCY
    str_DiscCurve = fld_UndParams.Curve_Disc

    ' Gather calendars and dates
    Dim cas_Calendars As CalendarSet: Set cas_Calendars = dic_GlobalStaticInfo(StaticInfoType.CalendarSet)
    Dim cal_pmt As Calendar: cal_pmt = cas_Calendars.Lookup_Calendar(fld_UndParams.PmtCal)
    lng_FutMatSpot = date_workday(fld_Params.FutMat, fld_Params.SettleDays, cal_pmt.HolDates, cal_pmt.Weekends)

    ' Set dependent curves
    If dic_CurveSet Is Nothing Then
        Set irc_Disc = GetObject_IRCurve(fld_UndParams.Curve_Disc, True, False, dic_GlobalStaticInfo)
        Set fxs_Spots = GetObject_FXSpots(True, dic_GlobalStaticInfo)
    Else
        Dim dic_IRCurves As Dictionary: Set dic_IRCurves = dic_CurveSet(CurveType.IRC)
        Set irc_Disc = dic_IRCurves(fld_UndParams.Curve_Disc)
        Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
    End If

    ' Set up components
    Dim int_TermInMonths As Integer: int_TermInMonths = calc_nummonths(str_PmtFreq)
    Dim int_NumPmts As Integer
    int_NumPmts = Calc_NumPmtsInWindow(lng_FutMatSpot, fld_Params.UndMat, str_PmtFreq, cal_pmt, fld_UndParams.BDC, bln_EOM, False)

    ' Set up underlying CTD bond leg
    With fld_UndParams
        .ValueDate = lng_FutMatSpot
        .Swapstart = lng_FutMatSpot
        .GenerationRefPoint = fld_Params.UndMat
        .IsFwdGeneration = False
        .Term = int_TermInMonths * int_NumPmts & "M"
        .PExch_Start = False
        .PExch_Intermediate = False
        .PExch_End = True
        .FloatEst = True
        .ForceToMV = False
        .Notional = 100
        .RateOrMargin = fld_Params.Coupon
        .IsUniformPeriods = fld_Params.IsUniformPeriods
    End With

    If fld_Params.IsSpreadOn_PnL = True Or fld_Params.IsSpreadOn_DV01 = True Then
        Set irl_Underlying_SpreadOn = New IRLeg
        Call irl_Underlying_SpreadOn.Initialize(fld_UndParams, dic_CurveSet, dic_GlobalStaticInfo)
    End If

    If fld_Params.IsSpreadOn_PnL = False Or fld_Params.IsSpreadOn_DV01 = False Then
        Set irl_Underlying_SpreadOff = New IRLeg
        Call irl_Underlying_SpreadOff.Initialize(fld_UndParams, dic_CurveSet, dic_GlobalStaticInfo)
    End If

    ' Derive accrual
    If fld_Params.PriceType = "CLEAN" Then
        lng_PrevCpnDate = Date_NextCoupon(fld_Params.UndMat, str_PmtFreq, cal_pmt, -int_NumPmts, bln_EOM, fld_Params.BDC_Accrual)
        dbl_UndAccrual = fld_Params.Coupon * calc_yearfrac(lng_PrevCpnDate, lng_FutMatSpot, str_Daycount, str_PmtFreq, True)
    Else
        dbl_UndAccrual = 0
    End If

    ' Set ZSpread such that bond clean price will equal the market price under the base scenario
    If fld_Params.IsSpreadOn_PnL = True Or fld_Params.IsSpreadOn_DV01 = True Then
        Call irl_Underlying_SpreadOn.SetCurveState(str_DiscCurve, CurveState_IRC.original)
        Call irl_Underlying_SpreadOn.ForceMVToValue(fld_Params.Price_Mkt * fld_Params.ConvFac / 100 + dbl_UndAccrual)
        Call irl_Underlying_SpreadOn.SetCurveState(str_DiscCurve, CurveState_IRC.Final)
    End If

    ' Determine curve dependencies
    Set dic_CurveDependencies = fxs_Spots.Lookup_CurveDependencies(fld_Params.CCY_PnL, fld_UndParams.CCY)
    If Not irl_Underlying_SpreadOff Is Nothing Then
        Set dic_CurveDependencies = Convert_MergeDicts(dic_CurveDependencies, irl_Underlying_SpreadOff.CurveDependencies)
    End If
    If Not irl_Underlying_SpreadOn Is Nothing Then
        Set dic_CurveDependencies = Convert_MergeDicts(dic_CurveDependencies, irl_Underlying_SpreadOn.CurveDependencies)
    End If
    If dic_CurveDependencies.Exists(irc_Disc.CurveName) = False Then Call dic_CurveDependencies.Add(irc_Disc.CurveName, True)
End Sub


' ## PROPERTIES
Public Property Get PnL() As Double
    ' ## Get the value of the contract in the PnL currency
    PnL = CalcValue(fld_Params.IsSpreadOn_PnL)
End Property


' ## METHODS - GREEKS
Public Function Calc_DV01(str_curve As String, Optional int_PillarIndex As Integer = 0) As Double
    ' ## Return DV01 sensitivity to the specified curve
    Dim dbl_Output As Double
    If dic_CurveDependencies.Exists(str_curve) Then
        ' Remember original setting, then disable DV01 impact on discounted spot
        Dim bln_ZShiftsEnabled_DiscSpot As Boolean: bln_ZShiftsEnabled_DiscSpot = fxs_Spots.ZShiftsEnabled_DiscSpot
        fxs_Spots.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

        Dim irl_ToUse As IRLeg
        If fld_Params.IsSpreadOn_DV01 = True Then
            Set irl_ToUse = irl_Underlying_SpreadOn
        Else
            Set irl_ToUse = irl_Underlying_SpreadOff
        End If

        ' Store shifted values
        Dim dbl_Val_Up As Double, dbl_Val_Down As Double
        Call irl_ToUse.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        dbl_Val_Up = CalcValue(fld_Params.IsSpreadOn_DV01)

        Call irl_ToUse.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        dbl_Val_Down = CalcValue(fld_Params.IsSpreadOn_DV01)

        ' Clear temporary shifts from the underlying leg
        Call irl_ToUse.SetCurveState(str_curve, CurveState_IRC.Final)

        ' Revert to original setting
        fxs_Spots.ZShiftsEnabled_DiscSpot = bln_ZShiftsEnabled_DiscSpot

        ' Calculate by finite differencing and convert to PnL currency
        dbl_Output = (dbl_Val_Up - dbl_Val_Down) / 2
    Else
        dbl_Output = 0
    End If

    Calc_DV01 = dbl_Output
End Function

Public Function Calc_DV02(str_curve As String, Optional int_PillarIndex As Integer = 0) As Double
    ' ## Return second order sensitivity to the specified curve
    Dim dbl_Output As Double
    If dic_CurveDependencies.Exists(str_curve) Then
        ' Remember original setting, then disable DV01 impact on discounted spot
        Dim bln_ZShiftsEnabled_DiscSpot As Boolean: bln_ZShiftsEnabled_DiscSpot = fxs_Spots.ZShiftsEnabled_DiscSpot
        fxs_Spots.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

        Dim irl_ToUse As IRLeg
        If fld_Params.IsSpreadOn_DV01 = True Then
            Set irl_ToUse = irl_Underlying_SpreadOn
        Else
            Set irl_ToUse = irl_Underlying_SpreadOff
        End If

        ' Store shifted values
        Dim dbl_Val_Up As Double, dbl_Val_Down As Double, dbl_Val_Unch As Double
        Call irl_ToUse.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        dbl_Val_Up = CalcValue(fld_Params.IsSpreadOn_DV01)

        Call irl_ToUse.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        dbl_Val_Down = CalcValue(fld_Params.IsSpreadOn_DV01)

        ' Clear temporary shifts from the underlying leg
        Call irl_ToUse.SetCurveState(str_curve, CurveState_IRC.Final)
        dbl_Val_Unch = CalcValue(fld_Params.IsSpreadOn_DV01)

        ' Revert to original setting
        fxs_Spots.ZShiftsEnabled_DiscSpot = bln_ZShiftsEnabled_DiscSpot

        ' Calculate by finite differencing and convert to PnL currency
        dbl_Output = dbl_Val_Up + dbl_Val_Down - 2 * dbl_Val_Unch
    Else
        dbl_Output = 0
    End If

    Calc_DV02 = dbl_Output
End Function


' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Sub SetValDate(lng_Input As Long)
    ' ## No dependency on time
End Sub

Public Sub HandleUpdate_IRC(str_CurveName As String)
    ' ## Update stored values affected by change in the curve values
    If Not irl_Underlying_SpreadOn Is Nothing Then Call irl_Underlying_SpreadOn.HandleUpdate_IRC(str_CurveName)
    If Not irl_Underlying_SpreadOff Is Nothing Then Call irl_Underlying_SpreadOff.HandleUpdate_IRC(str_CurveName)
End Sub


' ## METHODS - INTERMEDIATE CALCULATIONS
Private Function Calc_TheoFutPrice(bln_SpreadOn As Boolean) As Double
    ' ## Get futures price, optionally including a spread to bring the base case in line with the market
    Dim irl_ToUse As IRLeg
    If bln_SpreadOn = True Then Set irl_ToUse = irl_Underlying_SpreadOn Else Set irl_ToUse = irl_Underlying_SpreadOff

    Calc_TheoFutPrice = (irl_ToUse.marketvalue - dbl_UndAccrual) / (fld_Params.ConvFac / 100)
End Function

Private Function CalcValue(bln_SpreadOn As Boolean) As Double
    ' ## Get the value of the contract in the PnL currency
    Dim dbl_DiscFXSpot As Double: dbl_DiscFXSpot = fxs_Spots.Lookup_DiscSpot(str_CCY_Notional, fld_Params.CCY_PnL)
    CalcValue = (Calc_TheoFutPrice(bln_SpreadOn) - fld_Params.Price_Orig) / 100 * fld_Params.Notional * int_Sign * dbl_DiscFXSpot
End Function


' ## METHODS - CALCULATION DETAILS
Public Sub OutputReport(wks_output As Worksheet)
    wks_output.Cells.Clear

    Dim rng_OutputTopLeft As Range: Set rng_OutputTopLeft = wks_output.Range("A1")
    Dim dic_Addresses As New Dictionary, rng_PnL As Range, rng_CurrentPrice As Range
    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    Dim str_CurrencyFormat As String: str_CurrencyFormat = Gather_CurrencyFormat()
    Dim int_ActiveRow As Integer: int_ActiveRow = 0
    Dim irl_UndToUse As IRLeg
    If fld_Params.IsSpreadOn_PnL = True Then
        Set irl_UndToUse = irl_Underlying_SpreadOn
    Else
        Set irl_UndToUse = irl_Underlying_SpreadOff
    End If

    With rng_OutputTopLeft
        .Offset(int_ActiveRow, 0).Value = "Value date:"
        .Offset(int_ActiveRow, 1).Value = fld_Params.ValueDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("ValDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "PnL:"
        .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
        .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
        Set rng_PnL = .Offset(int_ActiveRow, 1)

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Currency:"
        .Offset(int_ActiveRow, 1).Value = fld_Params.CCY_PnL
        Call dic_Addresses.Add("PnLCCY", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 2
        .Offset(int_ActiveRow, 0).Value = "Notional:"
        .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
        .Offset(int_ActiveRow, 1).Value = fld_Params.Notional
        Call dic_Addresses.Add("Notional", .Offset(int_ActiveRow, 1).Address(False, False))
        .Offset(int_ActiveRow, 2).Value = str_CCY_Notional
        Call dic_Addresses.Add("NotionalCCY", .Offset(int_ActiveRow, 2).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Position:"
        If fld_Params.IsBuy = True Then .Offset(int_ActiveRow, 1).Value = "B" Else .Offset(int_ActiveRow, 1).Value = "S"
        Call dic_Addresses.Add("Position", .Offset(int_ActiveRow, 1).Address(False, False))

        If fld_Params.PriceType = "CLEAN" Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Prev Coupon:"
            .Offset(int_ActiveRow, 1).Value = lng_PrevCpnDate
            .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
            Call dic_Addresses.Add("PrevCoupon", .Offset(int_ActiveRow, 1).Address(False, False))
        End If

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Bond Start:"
        .Offset(int_ActiveRow, 1).Value = lng_FutMatSpot
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("BondStart", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Bond Mat:"
        .Offset(int_ActiveRow, 1).Value = fld_Params.UndMat
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("BondMat", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Coupon Rate:"
        .Offset(int_ActiveRow, 1).Value = fld_Params.Coupon
        Call dic_Addresses.Add("CouponRate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "ZSpread:"
        .Offset(int_ActiveRow, 1).Value = irl_UndToUse.ZSpread
        Call dic_Addresses.Add("ZSpread", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Conv Factor:"
        .Offset(int_ActiveRow, 1).Value = fld_Params.ConvFac
        Call dic_Addresses.Add("ConvFac", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Accrued Int:"
        Call dic_Addresses.Add("AccruedInt", .Offset(int_ActiveRow, 1).Address(False, False))
        If fld_Params.PriceType = "CLEAN" Then
            .Offset(int_ActiveRow, 1).Formula = "=" & dic_Addresses("CouponRate") & "*Calc_YearFrac(" & dic_Addresses("PrevCoupon") _
                & "," & dic_Addresses("BondStart") & ",""" & str_Daycount & """,""" & str_PmtFreq & """, True)"
        Else
            .Offset(int_ActiveRow, 1).Value = 0
        End If

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Current Price:"
        Call dic_Addresses.Add("CurrentPrice", .Offset(int_ActiveRow, 1).Address(False, False))
        Set rng_CurrentPrice = .Offset(int_ActiveRow, 1)

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Entry Price:"
        .Offset(int_ActiveRow, 1).Value = fld_Params.Price_Orig
        Call dic_Addresses.Add("EntryPrice", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 2
        Dim dic_BondAddresses As Dictionary
        Set dic_BondAddresses = irl_UndToUse.OutputReport_Bond(.Offset(int_ActiveRow, 0), lng_FutMatSpot, str_CCY_Notional, _
            int_Sign, str_DiscCurve, True, , dic_Addresses("ZSpread"))

        rng_CurrentPrice.Formula = "=(" & dic_BondAddresses("Bond_MV") & "-" & dic_Addresses("AccruedInt") & ")/(" _
            & dic_Addresses("ConvFac") & "/100)"

        rng_PnL.Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" & dic_Addresses("Notional") _
        & "*(" & dic_Addresses("CurrentPrice") & "-" & dic_Addresses("EntryPrice") & ")/100*cyGetFXDiscSpot(" _
        & dic_Addresses("NotionalCCY") & "," & dic_Addresses("PnLCCY") & ")"
    End With

    wks_output.Calculate
    wks_output.Cells.HorizontalAlignment = xlCenter
    wks_output.Columns.AutoFit
End Sub