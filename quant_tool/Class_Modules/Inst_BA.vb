Option Explicit

' ## MEMBER DATA
' Components
Private irl_BALeg As IRLeg_BA, scf_Purchase As SCF

' Dependent curves
Private fxs_Spots As Data_FXSpots, irc_SpotDisc As Data_IRCurve

' Variable dates
Private lng_ValDate As Long, lng_SpotDate As Long

' Static values
Private dic_GlobalStaticInfo As Dictionary, dic_CurveDependencies As Dictionary
Private str_CCY_PnL As String, int_Sign As Integer
Private int_SpotDays As Integer, cal_pmt As Calendar
Private dic_YieldParamsInput As Dictionary

' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As InstParams_BND, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing)

    ' Store static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    lng_ValDate = fld_ParamsInput.ValueDate
    str_CCY_PnL = fld_ParamsInput.CCY_PnL
    int_SpotDays = fld_ParamsInput.SpotDays
    Dim cas_Found As CalendarSet: Set cas_Found = dic_GlobalStaticInfo(StaticInfoType.CalendarSet)
    cal_pmt = cas_Found.Lookup_Calendar(fld_ParamsInput.PmtCal)
    Call FillSpotDate



    Dim fld_BALegParams As BondParams


    ''''' by SW: for yield solving purpose

    Dim ygs_Generators As YieldGeneratorSet: Set ygs_Generators = dic_GlobalStaticInfo(StaticInfoType.YieldGeneratorSet)
    'Dim ygd_Generators As BAParams: ygd_Generators = ygs_Generators.Lookup_YGenerator(fld_ParamsInput.SW_YieldGenerator)
    'Dim fld_BondLegParams As IRLegParams: fld_BondLegParams = ygd_Generators
    fld_BALegParams = ygs_Generators.Lookup_YGenerator(fld_ParamsInput.YieldGenerator)
    '''''

    ' Stored components
    With fld_BALegParams
        .ValueDate = lng_SpotDate
        .StubType = fld_ParamsInput.StubType
        .IsLongCpn = fld_ParamsInput.IsLongCpn
        .FirstAccDate = fld_ParamsInput.FirstAccDate
        .MatDate = fld_ParamsInput.MatDate
        .RollDate = fld_ParamsInput.RollDate
        .PaymentSchType = fld_ParamsInput.PaymentSchType

        .IsFwdGeneration = fld_ParamsInput.IsFwdGeneration

        .PExch_Start = False
        .PExch_Intermediate = True
        .PExch_End = True
        .FloatEst = True
        .ForceToMV = False
        Set .AmortSchedule = fld_ParamsInput.AmortSchedule
        .Notional = fld_ParamsInput.Principal
        .CCY = fld_ParamsInput.CCY_Principal
        .index = fld_ParamsInput.index
        .RateOrMargin = fld_ParamsInput.RateOrMargin
        .IsRoundFlow = fld_ParamsInput.IsRoundFlow
        .PmtFreq = fld_ParamsInput.PmtFreq
        .Daycount = fld_ParamsInput.Daycount
        .BDC = fld_ParamsInput.BDC
        .IsUniformPeriods = fld_ParamsInput.IsUniformPeriods
        .EOM = fld_ParamsInput.EOM
        .IsEOM2830 = fld_ParamsInput.EOM2830

        .PmtCal = fld_ParamsInput.PmtCal
        .estcal = fld_ParamsInput.estcal
        .Curve_Disc = fld_ParamsInput.Curve_Disc
        .Curve_Est = fld_ParamsInput.Curve_Est
        .DiscSpread = fld_ParamsInput.DiscSpread
        .CalType = fld_ParamsInput.CalType
        .BondMarketPrice = fld_ParamsInput.BondMarketPrice

        Set .Fixings = fld_ParamsInput.Fixings
        Set .ModStarts = fld_ParamsInput.ModStarts
    End With

    Set irl_BALeg = New IRLeg_BA
    Call irl_BALeg.Initialize(fld_BALegParams, dic_CurveSet, dic_GlobalStaticInfo)

    Set scf_Purchase = New SCF
    Call scf_Purchase.Initialize(fld_ParamsInput.PurchaseCost, dic_CurveSet, dic_GlobalStaticInfo)
    scf_Purchase.ZShiftsEnabled_DF = GetSetting_IsPremInDV01()
    scf_Purchase.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

    ' Dependent curves
    If dic_CurveSet Is Nothing Then
        Set irc_SpotDisc = GetObject_IRCurve(fld_ParamsInput.Curve_SpotDisc, True, False)
        Set fxs_Spots = GetObject_FXSpots(True)
    Else
        Set irc_SpotDisc = dic_CurveSet(CurveType.IRC)(fld_ParamsInput.Curve_SpotDisc)
        Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
    End If

    ' Calculated values
    If fld_ParamsInput.IsBuy = True Then int_Sign = 1 Else int_Sign = -1

    ' Determine curve dependencies
    Set dic_CurveDependencies = scf_Purchase.CurveDependencies
    Set dic_CurveDependencies = Convert_MergeDicts(dic_CurveDependencies, fxs_Spots.Lookup_CurveDependencies(fld_ParamsInput.Principal, _
        fld_ParamsInput.CCY_PnL))
    If dic_CurveDependencies.Exists(irc_SpotDisc.CurveName) = False Then Call dic_CurveDependencies.Add(irc_SpotDisc.CurveName, True)
    Set dic_CurveDependencies = Convert_MergeDicts(dic_CurveDependencies, irl_BALeg.CurveDependencies)
End Sub


' ## PROPERTIES - PUBLIC
Public Property Get marketvalue() As Double
    ' ## Get discounted value of future flows in the PnL currency
    marketvalue = irl_BALeg.marketvalue * SpotDF * FXConvFactor() * int_Sign
End Property

Public Property Get Cash() As Double
    ' ## Get discounted value of the premium in the PnL currency
    Cash = (irl_BALeg.CustomCashValuation(lng_ValDate, irc_SpotDisc) * FXConvFactor() - scf_Purchase.CalcValue(lng_ValDate, lng_SpotDate, str_CCY_PnL)) * int_Sign
End Property

Public Property Get PnL() As Double
    PnL = Me.marketvalue + Me.Cash
End Property

Public Property Get Yield() As Double
    Yield = irl_BALeg.BondYield
End Property

Public Property Get Duration() As Double
    Duration = irl_BALeg.BondDuration
End Property

Public Property Get ModifiedDuration() As Double
    ModifiedDuration = irl_BALeg.BondModifiedDuration
End Property

Public Function Calc_DV01(str_curve As String, Optional int_PillarIndex As Integer = 0) As Double
    Dim dbl_Output As Double
    If dic_CurveDependencies.Exists(str_curve) Then
        Dim bln_ZShiftsEnabled_DiscSpot As Boolean: bln_ZShiftsEnabled_DiscSpot = fxs_Spots.ZShiftsEnabled_DiscSpot
        fxs_Spots.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

        If int_PillarIndex = 0 Then
            ' DV01 does not include the cost price
            dbl_Output = irl_BALeg.Calc_DV01_Analytical(str_curve) * SpotDF * FXConvFactor() * int_Sign
        Else
            ' Use finite differencing instead of analytical
            ' Store shifted values
            Dim dbl_Val_Up As Double, dbl_Val_Down As Double
            Call irl_BALeg.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
            Call scf_Purchase.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
            dbl_Val_Up = Me.marketvalue

            Call irl_BALeg.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
            Call scf_Purchase.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
            dbl_Val_Down = Me.marketvalue

            ' Clear temporary shifts
            Call irl_BALeg.SetCurveState(str_curve, CurveState_IRC.Final)
            Call scf_Purchase.SetCurveState(str_curve, CurveState_IRC.Final)
            dbl_Output = (dbl_Val_Up - dbl_Val_Down) / 2
        End If

        fxs_Spots.ZShiftsEnabled_DiscSpot = bln_ZShiftsEnabled_DiscSpot
    Else
        dbl_Output = 0
    End If

    Calc_DV01 = dbl_Output
End Function

Public Function Calc_DV02(str_curve As String, Optional int_PillarIndex As Integer = 0) As Double
    ' ## Return second order sensitivity to the specified curve.  There is no impact on FX discounted spot
    Dim dbl_Output As Double
    If dic_CurveDependencies.Exists(str_curve) Then
        Dim bln_ZShiftsEnabled_DiscSpot As Boolean: bln_ZShiftsEnabled_DiscSpot = fxs_Spots.ZShiftsEnabled_DiscSpot
        fxs_Spots.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

        ' Store shifted values
        Dim dbl_Val_Up As Double, dbl_Val_Down As Double, dbl_Val_Unch As Double
        Call irl_BALeg.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        Call scf_Purchase.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        dbl_Val_Up = Me.marketvalue

        Call irl_BALeg.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        Call scf_Purchase.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        dbl_Val_Down = Me.marketvalue

        ' Clear temporary shifts
        Call irl_BALeg.SetCurveState(str_curve, CurveState_IRC.Final)
        Call scf_Purchase.SetCurveState(str_curve, CurveState_IRC.Final)
        dbl_Val_Unch = Me.marketvalue

        fxs_Spots.ZShiftsEnabled_DiscSpot = bln_ZShiftsEnabled_DiscSpot

        ' Calculate by finite differencing and convert to PnL currency
        dbl_Output = dbl_Val_Up + dbl_Val_Down - 2 * dbl_Val_Unch
    Else
        dbl_Output = 0
    End If

    Calc_DV02 = dbl_Output
End Function


' ## PROPERTIES - PRIVATE
Private Property Get SpotDF() As Double
    SpotDF = irc_SpotDisc.Lookup_Rate(lng_ValDate, lng_SpotDate, "DF", , , True)
End Property

Private Property Get FXConvFactor() As Double
    ' ## Discounted spot to translate into the PnL currency
    FXConvFactor = fxs_Spots.Lookup_DiscSpot(irl_BALeg.Params.CCY, str_CCY_PnL)
End Property


' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Sub HandleUpdate_IRC(str_CurveName As String)
    ' ## Update stored values affected by change in the curve values
    Call irl_BALeg.HandleUpdate_IRC(str_CurveName)
End Sub

Public Sub SetValDate(lng_Input As Long)
    ' ## Set stored value date and refresh values dependent on the value date
    lng_ValDate = lng_Input
    Call FillSpotDate
    Call irl_BALeg.SetValDate(lng_SpotDate)
End Sub


' ## METHODS - PRIVATE
Private Sub FillSpotDate()
    ' ## Compute and store spot date
    If int_SpotDays = 0 Then
        lng_SpotDate = Date_ApplyBDC(lng_ValDate, "FOLL", cal_pmt.HolDates, cal_pmt.Weekends)
    Else
        lng_SpotDate = date_workday(lng_ValDate, int_SpotDays, cal_pmt.HolDates, cal_pmt.Weekends)
    End If
End Sub


' ## METHODS - CALCULATION DETAILS
Public Sub OutputReport(wks_output As Worksheet)
    With wks_output
        .Cells.Clear

        Call irl_BALeg.OutputReport_Bond(.Range("A1"), lng_ValDate, str_CCY_PnL, int_Sign, irc_SpotDisc.CurveName, False, scf_Purchase)

        .Columns.AutoFit
        .Cells.HorizontalAlignment = xlCenter
    End With
End Sub