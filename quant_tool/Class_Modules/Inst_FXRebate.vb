Option Explicit

' ## ENUMERATIONS
Private Enum DetailedCat
    None = 0
    Cash_Knocked = 11
    Cash_Matured = 12
    SingleAmerican = 101
    SingleAmericanInstant = 102
    SingleAmericanQ = 103
    SingleAmericanQInstant = 104
    DoubleAmerican = 201
    DoubleAmericanInstant = 202
    DoubleAmericanQ = 203
    DoubleAmericanQInstant = 204
End Enum

Private Enum WindowType
    ' Numbers are related to numbering in DetailedCat
    American = 0
    European = 1
End Enum

Private Enum BarType
    NoBar = 0
    UpperBar = 1
    LowerBar = 2
    DoubleBar = 3
End Enum

Private Enum PayoffType
    Standard = 1
    StandardInstant = 2
    Quanto = 3
    QuantoInstant = 4
End Enum


' ## MEMBER DATA
' Components
Private scf_Premium As SCF

' Curve dependencies
Private fxs_Spots As Data_FXSpots, irc_DiscCurve As Data_IRCurve, irc_SpotDiscCurve As Data_IRCurve
Private fxv_Vols_XY As Data_FXVols, fxv_Vols_XQ As Data_FXVols, fxv_Vols_YQ As Data_FXVols

' Dynamic variables
Private lng_ValDate As Long, lng_SpotDate As Long, dbl_TimeToMat As Double, dbl_TimeEstPeriod As Double
Private dbl_ZShift_Disc As Double, dbl_ZShift_SpotDisc As Double
Private dbl_VolShiftXY_Sens As Double, dbl_VolShiftXQ_Sens As Double, dbl_VolShiftYQ_Sens As Double

' Static values
Private dic_GlobalStaticInfo As Dictionary, dic_CurveDependencies As Dictionary, map_Rules As MappingRules
Private fld_Params As InstParams_FRE
Private int_Sign_BS As Integer
Private lng_MatSpotDate As Long, dbl_SingleBarrier As Double
Private enu_Type_Barrier As BarType, enu_Type_Window As WindowType, enu_Type_Payoff As PayoffType
Private str_Pair_XY As String, str_Pair_XQ As String, str_Pair_YQ As String
Private bln_IsQuanto As Boolean


' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As InstParams_FRE, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing)

    ' Store static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    Set map_Rules = dic_GlobalStaticInfo(StaticInfoType.MappingRules)
    fld_Params = fld_ParamsInput
    If fld_Params.IsBuy = True Then int_Sign_BS = 1 Else int_Sign_BS = -1
    lng_MatSpotDate = cyGetFXCrossSpotDate(fld_Params.CCY_Fgn, fld_Params.CCY_Dom, fld_Params.MatDate, dic_GlobalStaticInfo)

    ' Stored components
    Set scf_Premium = New SCF
    Call scf_Premium.Initialize(fld_Params.Premium, dic_CurveSet, dic_GlobalStaticInfo)
    scf_Premium.ZShiftsEnabled_DF = GetSetting_IsPremInDV01()
    scf_Premium.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

    ' Determine relevant currency pairs
    str_Pair_XY = map_Rules.Lookup_MappedFXVolPair(fld_Params.CCY_Fgn, fld_Params.CCY_Dom)
    bln_IsQuanto = (fld_Params.CCY_Rebate <> fld_Params.CCY_Dom And fld_Params.CCY_Rebate <> fld_Params.CCY_Fgn)
    If bln_IsQuanto = True Then
        str_Pair_XQ = map_Rules.Lookup_MappedFXVolPair(fld_Params.CCY_Fgn, fld_Params.CCY_Rebate)
        str_Pair_YQ = map_Rules.Lookup_MappedFXVolPair(fld_Params.CCY_Dom, fld_Params.CCY_Rebate)
    End If

    ' Gather curves from curve set if available, otherwise use new instances
    If dic_CurveSet Is Nothing Then
        Set fxv_Vols_XY = GetObject_FXVols(str_Pair_XY, True, False)
        If bln_IsQuanto = True Then
            Set fxv_Vols_XQ = GetObject_FXVols(str_Pair_XQ, True, False)
            Set fxv_Vols_YQ = GetObject_FXVols(str_Pair_YQ, True, False)
        End If
        Set fxs_Spots = GetObject_FXSpots(True)
        Set irc_DiscCurve = GetObject_IRCurve(fld_Params.Curve_Disc, True, False)
        Set irc_SpotDiscCurve = GetObject_IRCurve(fld_Params.Curve_SpotDisc, True, False)
    Else
        Set fxv_Vols_XY = dic_CurveSet(CurveType.FXV)(str_Pair_XY)
        If bln_IsQuanto = True Then
            Set fxv_Vols_XQ = dic_CurveSet(CurveType.FXV)(str_Pair_XQ)
            Set fxv_Vols_YQ = dic_CurveSet(CurveType.FXV)(str_Pair_YQ)
        End If
        Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
        Set irc_DiscCurve = dic_CurveSet(CurveType.IRC)(fld_Params.Curve_Disc)
        Set irc_SpotDiscCurve = dic_CurveSet(CurveType.IRC)(fld_Params.Curve_SpotDisc)
    End If

    ' Variable dates and period lengths
    Call Me.SetValDate(fld_Params.ValueDate)

    ' Determine orientation of barriers
    dbl_SingleBarrier = -1
    If fld_Params.LowerBar = -1 And fld_Params.UpperBar <> -1 Then
        dbl_SingleBarrier = fld_Params.UpperBar
        enu_Type_Barrier = BarType.UpperBar
    ElseIf fld_Params.LowerBar <> -1 And fld_Params.UpperBar = -1 Then
        dbl_SingleBarrier = fld_Params.LowerBar
        enu_Type_Barrier = BarType.LowerBar
    ElseIf fld_Params.LowerBar <> -1 And fld_Params.UpperBar <> -1 Then
        enu_Type_Barrier = BarType.DoubleBar
    Else
        enu_Type_Barrier = BarType.NoBar
    End If

    ' Determine window type
    If fld_Params.WindowStart <= lng_ValDate And fld_Params.WindowEnd >= fld_Params.MatDate Then
        enu_Type_Window = WindowType.American
    End If

    ' Determine rebate payoff type
    If bln_IsQuanto = True Then
        enu_Type_Payoff = PayoffType.Quanto
    Else
        enu_Type_Payoff = PayoffType.Standard
    End If
    If fld_Params.IsInstantRebate = True And fld_Params.IsKnockOut = False Then enu_Type_Payoff = enu_Type_Payoff + 1

    ' Determine curve dependencies
    Set dic_CurveDependencies = scf_Premium.CurveDependencies
    Set dic_CurveDependencies = Convert_MergeDicts(dic_CurveDependencies, fxs_Spots.Lookup_CurveDependencies(fld_Params.CCY_Dom, _
        fld_Params.CCY_Fgn, fld_Params.CCY_Rebate, fld_Params.CCY_PnL))
    If dic_CurveDependencies.Exists(irc_DiscCurve.CurveName) = False Then Call dic_CurveDependencies.Add(irc_DiscCurve.CurveName, True)
    If dic_CurveDependencies.Exists(irc_SpotDiscCurve.CurveName) = False Then Call dic_CurveDependencies.Add(irc_SpotDiscCurve.CurveName, True)
End Sub


' ## PROPERTIES
Public Property Get marketvalue() As Double
    ' ## Get discounted value of option and rebate legs in the PnL currency
    ' Prepare values
    Dim enu_Detailed As DetailedCat: enu_Detailed = DetailedCategory()
    Dim dbl_Spot As Double: dbl_Spot = Me.Spot
    Dim dbl_Fwd As Double: dbl_Fwd = Me.Forward
    Dim dbl_VolPct_XY As Double: dbl_VolPct_XY = GetVol(VolPair.XY)
    Dim dbl_DF_Spot As Double
    If lng_SpotDate < lng_MatSpotDate Then
        dbl_DF_Spot = irc_SpotDiscCurve.Lookup_Rate(lng_ValDate, lng_SpotDate, "DF", , , False)
    Else
        dbl_DF_Spot = irc_SpotDiscCurve.Lookup_Rate(lng_ValDate, lng_MatSpotDate, "DF", , , False)
    End If

    Dim dbl_UnitVal As Double: dbl_UnitVal = 0
    Dim dbl_DF_Rebate As Double, bln_PayoutIsDom As Boolean, dbl_RebateAmt As Double
    dbl_DF_Rebate = irc_DiscCurve.Lookup_Rate(lng_SpotDate, lng_MatSpotDate, "DF")
    bln_PayoutIsDom = (fld_Params.CCY_Rebate = fld_Params.CCY_Dom)

    ' Calculate drift-adjusted forward
    Dim dbl_ATMVol_XY As Double, dbl_ATMVol_XQ As Double, dbl_ATMVol_YQ As Double, dbl_Correl_XQ As Double
    If bln_IsQuanto = True Then
        ' Perform drift adjustment on forward
        dbl_ATMVol_XY = GetVol(VolPair.XY) / 100
        dbl_ATMVol_XQ = GetVol(VolPair.XQ) / 100
        dbl_ATMVol_YQ = GetVol(VolPair.YQ) / 100
        dbl_Correl_XQ = Calc_QuantoCorrel(dbl_ATMVol_XY, dbl_ATMVol_XQ, dbl_ATMVol_YQ)
        dbl_Fwd = dbl_Fwd * Calc_DriftAdjFactor(dbl_ATMVol_XY, dbl_ATMVol_YQ, dbl_Correl_XQ, dbl_TimeToMat)

        ' Unit price is per unit of quanto currency
        dbl_RebateAmt = fld_Params.RebateAmt
    Else
        ' Number of units of the domestic currency since unit price is per domestic currency
        dbl_RebateAmt = fld_Params.RebateAmt * fxs_Spots.Lookup_Spot(fld_Params.CCY_Rebate, fld_Params.CCY_Dom)
    End If

    ' Value rebate as at spot date in the domestic currency
    Dim str_CCY_UnitVal As String
    Select Case enu_Detailed
        Case DetailedCat.Cash_Knocked, DetailedCat.Cash_Matured
            str_CCY_UnitVal = fld_Params.CCY_Dom
            If fld_Params.IsInstantRebate = True Then
                dbl_UnitVal = 1
            Else
                dbl_UnitVal = dbl_DF_Rebate
            End If
        Case DetailedCat.SingleAmerican, DetailedCat.SingleAmericanInstant
            str_CCY_UnitVal = fld_Params.CCY_Dom
            dbl_UnitVal = Calc_BSPrice_SingleBarRebate(dbl_Spot, dbl_Fwd, dbl_VolPct_XY, _
                dbl_TimeToMat, dbl_TimeEstPeriod, dbl_DF_Rebate, bln_PayoutIsDom, dbl_SingleBarrier, _
                (enu_Type_Barrier = BarType.UpperBar), fld_Params.IsKnockOut, fld_Params.IsInstantRebate)
        Case DetailedCat.SingleAmericanQ, DetailedCat.SingleAmericanQInstant
            str_CCY_UnitVal = fld_Params.CCY_Rebate
            dbl_UnitVal = Calc_BSPrice_SingleBarRebate(dbl_Spot, dbl_Fwd, dbl_VolPct_XY, _
                dbl_TimeToMat, dbl_TimeEstPeriod, dbl_DF_Rebate, True, dbl_SingleBarrier, _
                (enu_Type_Barrier = BarType.UpperBar), fld_Params.IsKnockOut, fld_Params.IsInstantRebate)
        Case DetailedCat.DoubleAmerican
            str_CCY_UnitVal = fld_Params.CCY_Dom
            dbl_UnitVal = Calc_BSPrice_DoubleBarRebate(dbl_Spot, dbl_Fwd, dbl_VolPct_XY, dbl_TimeToMat, _
                dbl_DF_Rebate, bln_PayoutIsDom, fld_Params.LowerBar, fld_Params.UpperBar, fld_Params.IsKnockOut)
        Case DetailedCat.DoubleAmericanInstant
            str_CCY_UnitVal = fld_Params.CCY_Dom
            dbl_UnitVal = Calc_BSPrice_DoubleBarInstantRebate(dbl_Spot, dbl_Fwd, dbl_VolPct_XY, dbl_TimeToMat, _
                dbl_TimeEstPeriod, dbl_DF_Rebate, bln_PayoutIsDom, fld_Params.LowerBar, fld_Params.UpperBar)
        Case DetailedCat.DoubleAmericanQ
            str_CCY_UnitVal = fld_Params.CCY_Rebate
            dbl_UnitVal = Calc_BSPrice_DoubleBarRebate(dbl_Spot, dbl_Fwd, dbl_VolPct_XY, dbl_TimeToMat, _
                dbl_DF_Rebate, True, fld_Params.LowerBar, fld_Params.UpperBar, fld_Params.IsKnockOut)
        Case DetailedCat.DoubleAmericanQInstant
            str_CCY_UnitVal = fld_Params.CCY_Rebate
            dbl_UnitVal = Calc_BSPrice_DoubleBarInstantRebate(dbl_Spot, dbl_Fwd, dbl_VolPct_XY, dbl_TimeToMat, _
                dbl_TimeEstPeriod, dbl_DF_Rebate, True, fld_Params.LowerBar, fld_Params.UpperBar)
        Case Else
            str_CCY_UnitVal = fld_Params.CCY_Dom
            dbl_UnitVal = 0
    End Select

    ' Convert to PnL currency at valuation date, then output final value
    marketvalue = dbl_RebateAmt * dbl_UnitVal * dbl_DF_Spot * fxs_Spots.Lookup_DiscSpot(str_CCY_UnitVal, fld_Params.CCY_PnL) * int_Sign_BS
End Property

Public Property Get Cash() As Double
    ' ## Get discounted value of the premium in the PnL currency
    Cash = -scf_Premium.CalcValue(lng_ValDate, lng_SpotDate, fld_Params.CCY_PnL) * int_Sign_BS
End Property

Public Property Get PnL() As Double
    PnL = Me.marketvalue + Me.Cash
End Property

Public Property Get Spot() As Double
    Spot = fxs_Spots.Lookup_Spot(fld_Params.CCY_Fgn, fld_Params.CCY_Dom)
End Property

Public Property Get Forward() As Double
    Forward = fxs_Spots.Lookup_Fwd(fld_Params.CCY_Fgn, fld_Params.CCY_Dom, fld_Params.MatDate)
End Property


' ## PROPERTIES - CATEGORIZATION
Private Property Get BarrierAlreadyHit() As Boolean
    Dim bln_Output As Boolean
    Dim dbl_Spot As Double: dbl_Spot = Me.Spot

    Select Case enu_Type_Barrier
        Case BarType.LowerBar: bln_Output = (dbl_Spot <= fld_Params.LowerBar)
        Case BarType.UpperBar: bln_Output = (dbl_Spot >= fld_Params.UpperBar)
        Case BarType.DoubleBar: bln_Output = (dbl_Spot <= fld_Params.LowerBar Or dbl_Spot >= fld_Params.UpperBar)
        Case BarType.NoBar: bln_Output = False
    End Select

    BarrierAlreadyHit = bln_Output
End Property

Private Property Get MaturedAlready() As Boolean
    MaturedAlready = (fld_Params.MatDate <= lng_ValDate)
End Property

Private Property Get ContainsLiveRebate() As Boolean
    ' ## Used for output of calculations, True if there is a rebate which is yet to knock or expire
    Dim bln_Output As Boolean
    Dim enu_DetailedCat As DetailedCat: enu_DetailedCat = DetailedCategory()
    Select Case enu_DetailedCat
        Case DetailedCat.None, DetailedCat.Cash_Knocked, DetailedCat.Cash_Matured: bln_Output = False
        Case Else: bln_Output = True
    End Select

    ContainsLiveRebate = bln_Output
End Property

Private Property Get DetailedCategory() As DetailedCat
    ' ## Determine option category for valuation and output purposes
    Dim enu_Output As DetailedCat: enu_Output = DetailedCat.None
    Dim bln_IsKnockOut As Boolean: bln_IsKnockOut = fld_Params.IsKnockOut
    Dim int_ID_Sides As Integer

    If BarrierAlreadyHit() = True Then
        If bln_IsKnockOut = True Then
            enu_Output = DetailedCat.None
        Else
            enu_Output = DetailedCat.Cash_Knocked
        End If
    ElseIf MaturedAlready() = True Then
        ' Barrier not hit at maturity
        If bln_IsKnockOut = True Then
            enu_Output = DetailedCat.Cash_Matured
        Else
            enu_Output = DetailedCat.None
        End If
    Else
        ' Live barrier
        Select Case enu_Type_Barrier
            Case BarType.LowerBar, BarType.UpperBar: int_ID_Sides = 100
            Case BarType.DoubleBar: int_ID_Sides = 200
        End Select

        enu_Output = int_ID_Sides + enu_Type_Window * 10 + enu_Type_Payoff
    End If

    DetailedCategory = enu_Output
End Property


' ## METHODS - GREEKS
Public Function Calc_DV01(str_curve As String, Optional int_PillarIndex As Integer = 0) As Double
    ' ## Return DV01 sensitivity to the specified curve
    Dim dbl_Output As Double
    If dic_CurveDependencies.Exists(str_curve) Then
        ' Remember original setting, then disable DV01 impact on discounted spot
        Dim bln_ZShiftsEnabled_DiscSpot As Boolean: bln_ZShiftsEnabled_DiscSpot = fxs_Spots.ZShiftsEnabled_DiscSpot
        fxs_Spots.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

        ' Store shifted value
        Dim dbl_Val_Up As Double, dbl_Val_Base As Double
        Call SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        dbl_Val_Up = Me.PnL

        ' Clear temporary shifts
        Call SetCurveState(str_curve, CurveState_IRC.Final)
        dbl_Val_Base = Me.PnL

        ' Restore original settings
        fxs_Spots.ZShiftsEnabled_DiscSpot = bln_ZShiftsEnabled_DiscSpot

        ' Calculate by finite differencing and convert to PnL currency
        dbl_Output = dbl_Val_Up - dbl_Val_Base
    Else
        dbl_Output = 0
    End If

    Calc_DV01 = dbl_Output
End Function

Public Function Calc_DV02(str_curve As String, Optional int_PillarIndex As Integer = 0) As Double
    ' ## Return second order sensitivity to the specified curve
    ' Remember original setting, then disable DV01 impact on discounted spot
    Dim dbl_Output As Double
    If dic_CurveDependencies.Exists(str_curve) Then
        Dim bln_ZShiftsEnabled_DiscSpot As Boolean: bln_ZShiftsEnabled_DiscSpot = fxs_Spots.ZShiftsEnabled_DiscSpot
        fxs_Spots.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

        ' Store shifted values
        Dim dbl_Val_Up As Double, dbl_Val_Down As Double, dbl_Val_Unch As Double
        Call SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        dbl_Val_Up = Me.PnL

        Call SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        dbl_Val_Down = Me.PnL

        ' Clear temporary shifts
        Call SetCurveState(str_curve, CurveState_IRC.Final)
        dbl_Val_Unch = Me.PnL

        ' Restore original settings
        fxs_Spots.ZShiftsEnabled_DiscSpot = bln_ZShiftsEnabled_DiscSpot

        ' Calculate by finite differencing and convert to PnL currency
        dbl_Output = dbl_Val_Up + dbl_Val_Down - 2 * dbl_Val_Unch
    Else
        dbl_Output = 0
    End If

    Calc_DV02 = dbl_Output
End Function

Public Function Calc_Vega(enu_Type As CurveType, str_curve As String) As Double
    ' ## Return sensitivity to a 1 vol increase in the vol
    Dim dbl_Output As Double: dbl_Output = 0
    If enu_Type = CurveType.FXV Then
        Dim dbl_Val_Up As Double, dbl_Val_Unch As Double
        Dim bln_IsShifted As Boolean

        ' Store shifted values and gather values
        bln_IsShifted = ApplyVolShift(str_curve, 0.1)
        If bln_IsShifted = True Then
            dbl_Val_Up = Me.PnL

            ' Clear temporary shifts from the vol curve
            Call ApplyVolShift(str_curve, 0)
            dbl_Val_Unch = Me.PnL

            ' Calculate by finite differencing and convert to PnL currency
            dbl_Output = (dbl_Val_Up - dbl_Val_Unch) * 10
        End If
    End If

    Calc_Vega = dbl_Output
End Function


' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Sub SetValDate(lng_Input As Long)
    ' Set stored value date and also dates & periods dependent on the value date
    lng_ValDate = lng_Input
    If lng_ValDate > lng_MatSpotDate Then lng_ValDate = lng_MatSpotDate  ' Prevent accumulation of expired deal
    lng_SpotDate = cyGetFXCrossSpotDate(fld_Params.CCY_Fgn, fld_Params.CCY_Dom, lng_ValDate, dic_GlobalStaticInfo)
    dbl_TimeToMat = calc_yearfrac(lng_ValDate, fld_Params.MatDate, "ACT/365")
    dbl_TimeEstPeriod = calc_yearfrac(lng_SpotDate, lng_MatSpotDate, "ACT/365")
End Sub

Private Sub SetCurveState(str_curve As String, enu_State As CurveState_IRC, Optional int_PillarIndex As Integer = 0)
    ' ## Set up shift in the market data and underlying components
    If irc_DiscCurve.CurveName = str_curve Then Call irc_DiscCurve.SetCurveState(enu_State, int_PillarIndex)
    If irc_SpotDiscCurve.CurveName = str_curve Then Call irc_SpotDiscCurve.SetCurveState(enu_State, int_PillarIndex)
    Call scf_Premium.SetCurveState(str_curve, enu_State, int_PillarIndex)
    Call fxs_Spots.SetCurveState(str_curve, enu_State, int_PillarIndex)
    Call fxv_Vols_XY.SetCurveState(str_curve, enu_State, int_PillarIndex)
    If bln_IsQuanto = True Then
        Call fxv_Vols_XQ.SetCurveState(str_curve, enu_State, int_PillarIndex)
        Call fxv_Vols_YQ.SetCurveState(str_curve, enu_State, int_PillarIndex)
    End If
End Sub

Private Function ApplyVolShift(str_curve As String, dbl_ShiftSize As Double) As Boolean
    ' ## Set shift for the specified curve, and return whether any curve was shifted
    Dim bln_IsShifted As Boolean: bln_IsShifted = False

    If fxv_Vols_XY.CurveName = str_curve Then
        dbl_VolShiftXY_Sens = dbl_ShiftSize
        'fxv_Vols_XY.VolShift_Sens = dbl_ShiftSize
        bln_IsShifted = True
    End If
    If bln_IsQuanto = True Then
        If fxv_Vols_XQ.CurveName = str_curve Then
            dbl_VolShiftXQ_Sens = dbl_ShiftSize
            'fxv_Vols_XQ.VolShift_Sens = dbl_ShiftSize
            bln_IsShifted = True
        End If
        If fxv_Vols_YQ.CurveName = str_curve Then
            dbl_VolShiftYQ_Sens = dbl_ShiftSize
            'fxv_Vols_YQ.VolShift_Sens = dbl_ShiftSize
            bln_IsShifted = True
        End If
    End If

    ApplyVolShift = bln_IsShifted
End Function


' ## METHODS - INTERMEDIATE CALCULATIONS
Private Function GetVol(enu_Pair As VolPair) As Double
    ' Determine vol curve to look up
    Dim dbl_Output As Double
    Dim fxv_VolCurve As Data_FXVols, dbl_VolShift As Double
    Select Case enu_Pair
        Case VolPair.XY: Set fxv_VolCurve = fxv_Vols_XY
        Case VolPair.XQ: Set fxv_VolCurve = fxv_Vols_XQ
        Case VolPair.YQ: Set fxv_VolCurve = fxv_Vols_YQ
    End Select

    ' Obtain the vol from the curve
    dbl_Output = fxv_VolCurve.Lookup_ATMVol(fld_Params.MatDate)

    Select Case enu_Pair
        Case VolPair.XY: dbl_Output = dbl_Output + dbl_VolShiftXY_Sens
        Case VolPair.XQ: dbl_Output = dbl_Output + dbl_VolShiftXQ_Sens
        Case VolPair.YQ: dbl_Output = dbl_Output + dbl_VolShiftYQ_Sens
    End Select



    GetVol = dbl_Output
End Function


' ## METHODS - CALCULATION DETAILS
Public Sub OutputReport(wks_output As Worksheet)
    wks_output.Cells.Clear

    Dim rng_TopLeft As Range: Set rng_TopLeft = wks_output.Range("A1")
    Dim str_CurrencyFormat As String: str_CurrencyFormat = Gather_CurrencyFormat()
    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    Dim int_ActiveRow As Integer: int_ActiveRow = 0
    Dim enu_Type_Detailed As DetailedCat: enu_Type_Detailed = DetailedCategory()
    Dim bln_Next_Display As Boolean, str_Next_Formula As String
    Dim dic_Addresses As Dictionary: Set dic_Addresses = New Dictionary
    dic_Addresses.CompareMode = CompareMethod.TextCompare

    ' Categorization
    Dim bln_ContainsLiveRebate As Boolean: bln_ContainsLiveRebate = ContainsLiveRebate()
    Dim bln_ContainsKnockedInRebate As Boolean: bln_ContainsKnockedInRebate = (enu_Type_Detailed = DetailedCat.Cash_Knocked)
    Dim bln_ContainsLiveSingle As Boolean: bln_ContainsLiveSingle = (enu_Type_Detailed > 100 And enu_Type_Detailed < 200)
    Dim bln_ContainsLiveDouble As Boolean: bln_ContainsLiveDouble = (enu_Type_Detailed > 200 And enu_Type_Detailed < 300)

    With rng_TopLeft
        ' Display PnL
        .Offset(int_ActiveRow, 0).Value = "OVERALL"
        .Offset(int_ActiveRow, 0).Font.Italic = True

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Value Date:"
        .Offset(int_ActiveRow, 1).Value = lng_ValDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("ValDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Spot Date:"
        .Offset(int_ActiveRow, 1).Value = lng_SpotDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("SpotDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "PnL:"
        .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
        .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
        Call dic_Addresses.Add("Range_PnL", .Offset(int_ActiveRow, 1))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "CCY:"
        .Offset(int_ActiveRow, 1).Value = fld_Params.CCY_PnL
        Call dic_Addresses.Add("PnLCCY", .Offset(int_ActiveRow, 1).Address(False, False))

        ' Display MV for option component
        int_ActiveRow = int_ActiveRow + 2
        .Offset(int_ActiveRow, 0).Value = "OPTION LEG"
        .Offset(int_ActiveRow, 0).Font.Italic = True

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Status:"
        Select Case enu_Type_Detailed
            Case DetailedCat.None: .Offset(int_ActiveRow, 1).Value = "No MV"
            Case DetailedCat.Cash_Matured: .Offset(int_ActiveRow, 1).Value = "Matured"
            Case DetailedCat.Cash_Knocked: .Offset(int_ActiveRow, 1).Value = "Knocked"
            Case Else: .Offset(int_ActiveRow, 1).Value = "Live"
        End Select

        If enu_Type_Detailed <> DetailedCat.None Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Rebate Amt:"
            .Offset(int_ActiveRow, 1).Value = fld_Params.RebateAmt
            .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
            Call dic_Addresses.Add("RebateAmt", .Offset(int_ActiveRow, 1).Address(False, False))
            .Offset(int_ActiveRow, 2).Value = fld_Params.CCY_Rebate
            Call dic_Addresses.Add("RebateCCY", .Offset(int_ActiveRow, 2).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Position:"
            If fld_Params.IsBuy = True Then
                .Offset(int_ActiveRow, 1).Value = "B"
            Else
                .Offset(int_ActiveRow, 1).Value = "S"
            End If
            Call dic_Addresses.Add("Position", .Offset(int_ActiveRow, 1).Address(False, False))
        End If

        If bln_ContainsLiveSingle = True Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Direction:"
            If enu_Type_Barrier = BarType.UpperBar Then
                .Offset(int_ActiveRow, 1).Value = "Up"
            ElseIf enu_Type_Barrier = BarType.LowerBar Then
                .Offset(int_ActiveRow, 1).Value = "Down"
            End If
            Call dic_Addresses.Add("Direction", .Offset(int_ActiveRow, 1).Address(False, False))
        End If

        If enu_Type_Detailed <> DetailedCat.None Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Rebate Knocks:"
            If fld_Params.IsKnockOut = True Then
                .Offset(int_ActiveRow, 1).Value = "Out"
            Else
                .Offset(int_ActiveRow, 1).Value = "In"
            End If
            Call dic_Addresses.Add("Knock_Reb", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Received:"
            If fld_Params.IsInstantRebate Then
                .Offset(int_ActiveRow, 1).Value = "Knock Time"
            Else
                .Offset(int_ActiveRow, 1).Value = "Maturity"
            End If
            Call dic_Addresses.Add("ReceivedAt", .Offset(int_ActiveRow, 1).Address(False, False))
        End If

        If bln_ContainsLiveSingle = True Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Barrier:"
            .Offset(int_ActiveRow, 1).Value = dbl_SingleBarrier
            Call dic_Addresses.Add("Barrier", .Offset(int_ActiveRow, 1).Address(False, False))
        ElseIf bln_ContainsLiveDouble = True Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Lower Bar:"
            .Offset(int_ActiveRow, 1).Value = fld_Params.LowerBar
            Call dic_Addresses.Add("LowerBar", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Upper Bar:"
            .Offset(int_ActiveRow, 1).Value = fld_Params.UpperBar
            Call dic_Addresses.Add("UpperBar", .Offset(int_ActiveRow, 1).Address(False, False))
        End If

        If enu_Type_Detailed <> DetailedCat.None Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Spot:"
            .Offset(int_ActiveRow, 1).Value = Me.Spot
            Call dic_Addresses.Add("Spot", .Offset(int_ActiveRow, 1).Address(False, False))
        End If

        If bln_ContainsLiveRebate = True Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Forward:"
            .Offset(int_ActiveRow, 1).Value = Me.Forward
            Call dic_Addresses.Add("Fwd", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Vol:"
            .Offset(int_ActiveRow, 1).Value = GetVol(VolPair.XY)
            Call dic_Addresses.Add("Vol", .Offset(int_ActiveRow, 1).Address(False, False))
        End If

        If bln_ContainsLiveRebate = True And bln_IsQuanto = True Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "ATM Vol (Fgn/Dom):"
            .Offset(int_ActiveRow, 1).Value = GetVol(VolPair.XY)
            Call dic_Addresses.Add("ATMVol_XY", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "ATM Vol (Fgn/Qto):"
            .Offset(int_ActiveRow, 1).Value = GetVol(VolPair.XQ)
            Call dic_Addresses.Add("ATMVol_XQ", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "ATM Vol (Dom/Qto):"
            .Offset(int_ActiveRow, 1).Value = GetVol(VolPair.YQ)
            Call dic_Addresses.Add("ATMVol_YQ", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Correl (Fgn/Qto):"
            .Offset(int_ActiveRow, 1).Formula = "=((" & dic_Addresses("ATMVol_XQ") & "/100)^2-(" & dic_Addresses("ATMVol_XY") _
                & "/100)^2-(" & dic_Addresses("ATMVol_YQ") & "/100)^2)/(2*" & dic_Addresses("ATMVol_XY") & "/100*" & dic_Addresses("ATMVol_YQ") & "/100)"
            Call dic_Addresses.Add("Correl_XQ", .Offset(int_ActiveRow, 1).Address(False, False))
        End If

        If bln_ContainsLiveRebate = True Or bln_ContainsKnockedInRebate Then
                int_ActiveRow = int_ActiveRow + 1
                .Offset(int_ActiveRow, 0).Value = "Maturity:"
                .Offset(int_ActiveRow, 1).Value = fld_Params.MatDate
                .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
                Call dic_Addresses.Add("MatDate", .Offset(int_ActiveRow, 1).Address(False, False))

                int_ActiveRow = int_ActiveRow + 1
                .Offset(int_ActiveRow, 0).Value = "Spot Date:"
                .Offset(int_ActiveRow, 1).Value = lng_MatSpotDate
                .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
                Call dic_Addresses.Add("MatSpotDate", .Offset(int_ActiveRow, 1).Address(False, False))

                int_ActiveRow = int_ActiveRow + 1
                .Offset(int_ActiveRow, 0).Value = "Delivery:"
                .Offset(int_ActiveRow, 1).Value = fld_Params.DelivDate
                .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
                Call dic_Addresses.Add("DelivDate", .Offset(int_ActiveRow, 1).Address(False, False))
        End If

        If bln_ContainsLiveRebate = True And bln_IsQuanto = True Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Drift Adj Fwd:"
            .Offset(int_ActiveRow, 1).Formula = "=" & dic_Addresses("Fwd") & "*Calc_DriftAdjFactor(" & dic_Addresses("ATMVol_XY") _
                & "/100," & dic_Addresses("ATMVol_YQ") & "/100," & dic_Addresses("Correl_XQ") & ",Calc_YearFrac(" & dic_Addresses("ValDate") _
                & "," & dic_Addresses("MatDate") & ",""ACT/365""))"
            Call dic_Addresses.Add("Fwd_DriftAdj", .Offset(int_ActiveRow, 1).Address(False, False))
        End If

        If enu_Type_Detailed <> DetailedCat.None And enu_Type_Detailed <> Cash_Matured Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Rebate DF:"
            .Offset(int_ActiveRow, 1).Formula = "=cyReadIRCurve(""" & fld_Params.Curve_Disc & """," & dic_Addresses("SpotDate") _
                & "," & dic_Addresses("MatSpotDate") & ",""DF"",,False)"
            Call dic_Addresses.Add("RebateDF", .Offset(int_ActiveRow, 1).Address(False, False))
        End If

        If enu_Type_Detailed <> DetailedCat.None Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Spot DF:"
            .Offset(int_ActiveRow, 1).Formula = "=cyReadIRCurve(""" & fld_Params.Curve_SpotDisc & """," & dic_Addresses("ValDate") _
                & ",MIN(" & dic_Addresses("SpotDate") & "," & dic_Addresses("MatSpotDate") & "),""DF"",,False)"
            Call dic_Addresses.Add("SpotDF", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Dom CCY:"
            .Offset(int_ActiveRow, 1).Value = fld_Params.CCY_Dom
            Call dic_Addresses.Add("DomCCY", .Offset(int_ActiveRow, 1).Address(False, False))
        End If

        ' Output rebate MV
        Dim str_CCY_UnitVal As String
        bln_Next_Display = True
        Select Case enu_Type_Detailed
            Case DetailedCat.SingleAmerican, DetailedCat.SingleAmericanInstant
                str_CCY_UnitVal = fld_Params.CCY_Dom
                str_Next_Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" & dic_Addresses("RebateAmt") _
                    & "*cyGetFXSpot(" & dic_Addresses("RebateCCY") & "," & dic_Addresses("DomCCY") _
                    & ")" & "*Calc_BSPrice_SingleBarRebate(" & dic_Addresses("Spot") & "," & dic_Addresses("Fwd") & "," _
                    & dic_Addresses("Vol") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," & dic_Addresses("MatDate") _
                    & ",""ACT/365""),Calc_YearFrac(" & dic_Addresses("SpotDate") & "," & dic_Addresses("MatSpotDate") & ",""ACT/365"")," _
                    & dic_Addresses("RebateDF") & "," & dic_Addresses("RebateCCY") & "=" & dic_Addresses("DomCCY") & "," _
                    & dic_Addresses("Barrier") & "," & dic_Addresses("Direction") & "=""Up""," & dic_Addresses("Knock_Reb") & "=""Out""," _
                    & dic_Addresses("ReceivedAt") & "=""Knock Time"")*" & dic_Addresses("SpotDF") _
                    & "*cyGetFXDiscSpot(" & dic_Addresses("DomCCY") & "," & dic_Addresses("PnLCCY") & ")"
            Case DetailedCat.SingleAmericanQ, DetailedCat.SingleAmericanQInstant
                str_CCY_UnitVal = fld_Params.CCY_Rebate
                str_Next_Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" & dic_Addresses("RebateAmt") _
                    & "*Calc_BSPrice_SingleBarRebate(" & dic_Addresses("Spot") & "," & dic_Addresses("Fwd_DriftAdj") & "," _
                    & dic_Addresses("Vol") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," & dic_Addresses("MatDate") _
                    & ",""ACT/365""),Calc_YearFrac(" & dic_Addresses("SpotDate") & "," & dic_Addresses("MatSpotDate") & ",""ACT/365"")," _
                    & dic_Addresses("RebateDF") & ",True," & dic_Addresses("Barrier") & "," & dic_Addresses("Direction") & "=""Up""," _
                    & dic_Addresses("Knock_Reb") & "=""Out""," & dic_Addresses("ReceivedAt") & "=""Knock Time"")*" & dic_Addresses("SpotDF") _
                    & "*cyGetFXDiscSpot(" & dic_Addresses("RebateCCY") & "," & dic_Addresses("PnLCCY") & ")"
            Case DetailedCat.DoubleAmerican
                str_CCY_UnitVal = fld_Params.CCY_Dom
                str_Next_Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" & dic_Addresses("RebateAmt") _
                    & "*cyGetFXSpot(" & dic_Addresses("RebateCCY") & "," & dic_Addresses("DomCCY") _
                    & ")" & "*Calc_BSPrice_DoubleBarRebate(" & dic_Addresses("Spot") & "," & dic_Addresses("Fwd") & "," _
                    & dic_Addresses("Vol") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," & dic_Addresses("MatDate") _
                    & ",""ACT/365"")," & dic_Addresses("RebateDF") & "," & dic_Addresses("RebateCCY") & "=" _
                    & dic_Addresses("DomCCY") & "," & dic_Addresses("LowerBar") & "," & dic_Addresses("UpperBar") _
                    & "," & dic_Addresses("Knock_Reb") & "=""Out"")*" & dic_Addresses("SpotDF") _
                    & "*cyGetFXDiscSpot(" & dic_Addresses("DomCCY") & "," & dic_Addresses("PnLCCY") & ")"
            Case DetailedCat.DoubleAmericanInstant
                str_CCY_UnitVal = fld_Params.CCY_Dom
                str_Next_Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" & dic_Addresses("RebateAmt") _
                    & "*cyGetFXSpot(" & dic_Addresses("RebateCCY") & "," & dic_Addresses("DomCCY") _
                    & ")" & "*Calc_BSPrice_DoubleBarInstantRebate(" & dic_Addresses("Spot") & "," & dic_Addresses("Fwd") & "," _
                    & dic_Addresses("Vol") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," & dic_Addresses("MatDate") _
                    & ",""ACT/365""),Calc_YearFrac(" & dic_Addresses("SpotDate") & "," & dic_Addresses("MatSpotDate") & ",""ACT/365"")," _
                    & dic_Addresses("RebateDF") & "," & dic_Addresses("RebateCCY") & "=" & dic_Addresses("DomCCY") & "," _
                    & dic_Addresses("LowerBar") & "," & dic_Addresses("UpperBar") & ")*" & dic_Addresses("SpotDF") & "*cyGetFXDiscSpot(" _
                    & dic_Addresses("DomCCY") & "," & dic_Addresses("PnLCCY") & ")"
            Case DetailedCat.DoubleAmericanQ
                str_CCY_UnitVal = fld_Params.CCY_Fgn
                str_Next_Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" & dic_Addresses("RebateAmt") _
                    & "*Calc_BSPrice_DoubleBarRebate(" & dic_Addresses("Spot") & "," & dic_Addresses("Fwd_DriftAdj") & "," _
                    & dic_Addresses("Vol") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," & dic_Addresses("MatDate") _
                    & ",""ACT/365"")," & dic_Addresses("RebateDF") & ",True," & dic_Addresses("LowerBar") & "," _
                    & dic_Addresses("UpperBar") & "," & dic_Addresses("Knock_Reb") & "=""Out"")*" & dic_Addresses("SpotDF") _
                    & "*cyGetFXDiscSpot(" & dic_Addresses("RebateCCY") & "," & dic_Addresses("PnLCCY") & ")"
            Case DetailedCat.DoubleAmericanQInstant
                str_CCY_UnitVal = fld_Params.CCY_Dom
                str_Next_Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" & dic_Addresses("RebateAmt") _
                    & "*Calc_BSPrice_DoubleBarInstantRebate(" & dic_Addresses("Spot") & "," & dic_Addresses("Fwd_DriftAdj") _
                    & "," & dic_Addresses("Vol") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," & dic_Addresses("MatDate") _
                    & ",""ACT/365""),Calc_YearFrac(" & dic_Addresses("SpotDate") & "," & dic_Addresses("MatSpotDate") & ",""ACT/365"")," _
                    & dic_Addresses("RebateDF") & ",True," & dic_Addresses("LowerBar") & "," & dic_Addresses("UpperBar") & ")*" _
                    & dic_Addresses("SpotDF") & "*cyGetFXDiscSpot(" & dic_Addresses("RebateCCY") & "," & dic_Addresses("PnLCCY") & ")"
            Case DetailedCat.Cash_Knocked
                str_CCY_UnitVal = fld_Params.CCY_Dom
                str_Next_Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" _
                    & dic_Addresses("RebateAmt") & "*IF(" & dic_Addresses("ReceivedAt") & "=""Knock Time"",1," _
                    & dic_Addresses("RebateDF") & ")*cyGetFXSpot(" & dic_Addresses("RebateCCY") & "," & dic_Addresses("DomCCY") _
                    & ")" & "*" & dic_Addresses("SpotDF") & "*cyGetFXDiscSpot(" & dic_Addresses("DomCCY") & "," _
                    & dic_Addresses("PnLCCY") & ")"
            Case DetailedCat.Cash_Matured
                str_CCY_UnitVal = fld_Params.CCY_Dom
                str_Next_Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" _
                    & dic_Addresses("RebateAmt") & "*cyGetFXSpot(" & dic_Addresses("RebateCCY") & "," & dic_Addresses("DomCCY") _
                    & ")" & "*" & dic_Addresses("SpotDF") & "*cyGetFXDiscSpot(" & dic_Addresses("DomCCY") & "," _
                    & dic_Addresses("PnLCCY") & ")"
            Case Else
                str_CCY_UnitVal = fld_Params.CCY_Dom
                bln_Next_Display = False
                str_Next_Formula = ""
        End Select

        If bln_Next_Display = True Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Rebate MV (" & fld_Params.CCY_PnL & "):"
            .Offset(int_ActiveRow, 1).Formula = str_Next_Formula
            .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
            .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
            Call dic_Addresses.Add("MV_Rebate", .Offset(int_ActiveRow, 1).Address(False, False))
        End If

        ' Display cash
        int_ActiveRow = int_ActiveRow + 2
        .Offset(int_ActiveRow, 0).Value = "PREMIUM LEG"
        .Offset(int_ActiveRow, 0).Font.Italic = True

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Amount:"
        .Offset(int_ActiveRow, 1).Value = -fld_Params.Premium.Amount * int_Sign_BS
        .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
        Call dic_Addresses.Add("Premium", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "CCY:"
        .Offset(int_ActiveRow, 1).Value = fld_Params.Premium.CCY
        Call dic_Addresses.Add("PremiumCCY", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Pmt date:"
        .Offset(int_ActiveRow, 1).Value = fld_Params.Premium.PmtDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("PremiumDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "DF:"
        .Offset(int_ActiveRow, 1).Value = "=cyReadIRCurve(""" & fld_Params.Premium.Curve_Disc & """," _
            & dic_Addresses("ValDate") & "," & dic_Addresses("PremiumDate") & ",""DF"",,FALSE)"
        Call dic_Addresses.Add("PremiumDF", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Cash (" & fld_Params.CCY_PnL & "):"
        .Offset(int_ActiveRow, 1).Value = "=" & dic_Addresses("Premium") & "*" & dic_Addresses("PremiumDF") _
            & "*cyGetFXDiscSpot(" & dic_Addresses("PremiumCCY") & "," & dic_Addresses("PnLCCY") & ")"
        .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
        .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
        Call dic_Addresses.Add("Cash", .Offset(int_ActiveRow, 1).Address(False, False))

        ' Display PnL formula
        Select Case enu_Type_Detailed
            Case DetailedCat.None: dic_Addresses("Range_PnL").Formula = "=" & dic_Addresses("Cash")
            Case Else: dic_Addresses("Range_PnL").Formula = "=" & dic_Addresses("MV_Rebate") & "+" & dic_Addresses("Cash")
        End Select
    End With

    wks_output.Calculate
    wks_output.Cells.HorizontalAlignment = xlCenter
    wks_output.Columns.AutoFit
End Sub