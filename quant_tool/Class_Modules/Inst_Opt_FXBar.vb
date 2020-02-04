Option Explicit

' ## ENUMERATIONS
Private Enum DetailedCat
    None = 0
    CashFlow = 1
    Vanilla = 11
    Quanto = 13
    SingleAmerican = 101
    SingleAmericanQ = 103
    DoubleAmerican = 201
    DoubleAmericanQ = 203
    '#Matt edit
    EuropeanBar = 300
    '#Matt Edit end
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
    Digital = 2
    Quanto = 3
    DigitalQuanto = 4
End Enum


' ## MEMBER DATA
' Components
Private scf_Premium As SCF

' Curve dependencies
Private fxs_Spots As Data_FXSpots, irc_DiscCurve As Data_IRCurve, irc_SpotDiscCurve As Data_IRCurve
Private fxv_Vols_XY As Data_FXVols, fxv_Vols_XQ As Data_FXVols, fxv_Vols_YQ As Data_FXVols

' Dynamic variables
Private lng_ValDate As Long, lng_SpotDate As Long, dbl_TimeToMat As Double, dbl_TimeEstPeriod As Double
Private dbl_VolShiftXY_Sens As Double, dbl_VolShiftXQ_Sens As Double, dbl_VolShiftYQ_Sens As Double

' Static values
Private dic_GlobalStaticInfo As Dictionary, dic_CurveDependencies As Dictionary, map_Rules As MappingRules
Private fld_Params As InstParams_FBR
Private int_Sign_BS As Integer
Private lng_MatSpotDate As Long, dbl_SingleBarrier As Double
Private enu_Type_Payoff As PayoffType, enu_Type_Barrier As BarType, enu_Type_Window As WindowType
Private str_Pair_XY As String, str_Pair_XQ As String, str_Pair_YQ As String
Private bln_IsQuanto As Boolean


' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As InstParams_FBR, Optional dic_CurveSet As Dictionary = Nothing, _
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
    bln_IsQuanto = (fld_Params.CCY_Payout <> fld_Params.CCY_Dom And fld_Params.CCY_Payout <> fld_Params.CCY_Fgn)
    If bln_IsQuanto = True Then
        str_Pair_XQ = map_Rules.Lookup_MappedFXVolPair(fld_Params.CCY_Fgn, fld_Params.CCY_Payout)
        str_Pair_YQ = map_Rules.Lookup_MappedFXVolPair(fld_Params.CCY_Dom, fld_Params.CCY_Payout)
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

    ' Initialize variables
    Call Me.SetValDate(fld_Params.ValueDate)
    dbl_VolShiftXY_Sens = 0
    dbl_VolShiftXQ_Sens = 0
    dbl_VolShiftYQ_Sens = 0

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

    ' Determine payoff type
    If bln_IsQuanto = True Then
        enu_Type_Payoff = PayoffType.Quanto
    Else
        enu_Type_Payoff = PayoffType.Standard
    End If

    ' Determine curve dependencies
    Set dic_CurveDependencies = scf_Premium.CurveDependencies
    Set dic_CurveDependencies = Convert_MergeDicts(dic_CurveDependencies, fxs_Spots.Lookup_CurveDependencies(fld_Params.CCY_Dom, _
        fld_Params.CCY_Fgn, fld_Params.CCY_Payout, fld_Params.CCY_PnL))
    If dic_CurveDependencies.Exists(irc_DiscCurve.CurveName) = False Then Call dic_CurveDependencies.Add(irc_DiscCurve.CurveName, True)
    If dic_CurveDependencies.Exists(irc_SpotDiscCurve.CurveName) = False Then Call dic_CurveDependencies.Add(irc_SpotDiscCurve.CurveName, True)
End Sub


' ## PROPERTIES
Public Property Get marketvalue() As Double
    ' ## Get discounted value in the PnL currency

    ' Prepare values for valuing option
    Dim enu_Detailed As DetailedCat: enu_Detailed = DetailedCategory()
    '##Matt Edit
'    If (fld_Params.WindowStart = fld_Params.WindowEnd And fld_Params.WindowEnd = fld_Params.MatDate) Then
'        enu_Detailed = EuropeanBar
'    Else
'        enu_Detailed = DetailedCategory()
'    End If

    '##Matt Edit end
    Dim dbl_Spot As Double: dbl_Spot = Me.Spot
    Dim dbl_Fwd As Double: dbl_Fwd = Me.Forward
    Dim dbl_SmileStrike_Orig As Double, dbl_SmileStrike_Knocked As Double
    If fld_Params.IsSmile_Orig = True Then dbl_SmileStrike_Orig = fld_Params.strike Else dbl_SmileStrike_Orig = -1
    If fld_Params.IsSmile_IfKnocked = True Then dbl_SmileStrike_Knocked = fld_Params.strike Else dbl_SmileStrike_Knocked = -1
    Dim dbl_VolPct_XY As Double: dbl_VolPct_XY = GetVol(VolPair.XY, dbl_SmileStrike_Orig)
    Dim dbl_UnitVal As Double

    ' Calculate discount factors
    Dim dbl_DF_Option As Double: dbl_DF_Option = irc_DiscCurve.Lookup_Rate(lng_SpotDate, lng_MatSpotDate, "DF", , , False)
    Dim dbl_DF_Spot As Double
    If lng_SpotDate < lng_MatSpotDate Then
        dbl_DF_Spot = irc_SpotDiscCurve.Lookup_Rate(lng_ValDate, lng_SpotDate, "DF", , , False)
    Else
        dbl_DF_Spot = irc_SpotDiscCurve.Lookup_Rate(lng_ValDate, lng_MatSpotDate, "DF", , , False)
    End If

    ' Calculate drift-adjusted forward
    Dim dbl_ATMVol_XY As Double, dbl_ATMVol_XQ As Double, dbl_ATMVol_YQ As Double, dbl_Correl_XQ As Double
    Dim dbl_DriftAdjFactor As Double
    If bln_IsQuanto = True Then
        ' Perform drift adjustment on forward
        dbl_ATMVol_XY = GetVol(VolPair.XY, -1) / 100
        dbl_ATMVol_XQ = GetVol(VolPair.XQ, -1) / 100
        dbl_ATMVol_YQ = GetVol(VolPair.YQ, -1) / 100
        dbl_Correl_XQ = Calc_QuantoCorrel(dbl_ATMVol_XY, dbl_ATMVol_XQ, dbl_ATMVol_YQ)
        dbl_DriftAdjFactor = Calc_DriftAdjFactor(dbl_ATMVol_XY, dbl_ATMVol_YQ, dbl_Correl_XQ, dbl_TimeToMat)
        dbl_Fwd = dbl_Fwd * dbl_DriftAdjFactor
    End If

    ' Calculate FX conversion factor
    Dim dbl_DiscFXSpot As Double
    If bln_IsQuanto = True Then
        dbl_DiscFXSpot = fxs_Spots.Lookup_DiscSpot(fld_Params.CCY_Payout, fld_Params.CCY_PnL)
    Else
        dbl_DiscFXSpot = fxs_Spots.Lookup_DiscSpot(fld_Params.CCY_Dom, fld_Params.CCY_PnL)
    End If

    ' Value option as at spot date in the domestic currency
    Select Case enu_Detailed
        Case DetailedCat.CashFlow
            dbl_UnitVal = WorksheetFunction.Max(fld_Params.OptDirection * (dbl_Spot - fld_Params.strike), 0) * dbl_DF_Spot
        Case DetailedCat.SingleAmerican
            dbl_UnitVal = Calc_BSPrice_SingleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, dbl_SingleBarrier, (enu_Type_Barrier = BarType.UpperBar), _
                fld_Params.IsKnockOut) * dbl_DF_Option * dbl_DF_Spot
        Case DetailedCat.SingleAmericanQ
            dbl_UnitVal = Calc_BSPrice_SingleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, dbl_SingleBarrier, (enu_Type_Barrier = BarType.UpperBar), _
                fld_Params.IsKnockOut) * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot
        Case DetailedCat.DoubleAmerican
            dbl_UnitVal = Calc_BSPrice_DoubleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, fld_Params.LowerBar, fld_Params.UpperBar, fld_Params.IsKnockOut) _
                 * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot
        Case DetailedCat.DoubleAmericanQ
            dbl_UnitVal = Calc_BSPrice_DoubleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, fld_Params.LowerBar, fld_Params.UpperBar, fld_Params.IsKnockOut) _
                 * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot
        Case DetailedCat.Vanilla
            dbl_UnitVal = Calc_BSPrice_Vanilla(fld_Params.OptDirection, dbl_Fwd, fld_Params.strike, _
                dbl_TimeToMat, GetVol(VolPair.XY, dbl_SmileStrike_Knocked)) * dbl_DF_Option * dbl_DF_Spot
        Case DetailedCat.Quanto
            dbl_UnitVal = Calc_BSPrice_Vanilla(fld_Params.OptDirection, dbl_Fwd, fld_Params.strike, dbl_TimeToMat, _
            GetVol(VolPair.XY, dbl_SmileStrike_Knocked)) * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot
        '##Matt Edit
        Case DetailedCat.EuropeanBar
            dbl_UnitVal = EuropeanBarType(fld_Params.IsKnockOut, fld_Params.strike, fld_Params.OptDirection, fld_Params.LowerBar, fld_Params.UpperBar)
            If (int_Sign_BS = 1 And int_Sign_BS * dbl_UnitVal < 0) Or (int_Sign_BS = -1 And int_Sign_BS * dbl_UnitVal > 0) Then
                dbl_UnitVal = 0
            End If
            dbl_UnitVal = dbl_UnitVal * dbl_DF_Option * dbl_DF_Spot
        '##Matt Edit end
        Case Else: dbl_UnitVal = 0
    End Select

    ' Convert to PnL currency at valuation date, then output final value
    marketvalue = fld_Params.Notional_Fgn * dbl_UnitVal * int_Sign_BS * dbl_DiscFXSpot
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

Private Property Get ContainsLiveOption() As Boolean
    ' ## Used for output of calculations, True if any sort of option needs to be displayed
    Dim bln_Output As Boolean
    Dim enu_DetailedCat As DetailedCat ': enu_DetailedCat = DetailedCategory()
    If (fld_Params.WindowStart = fld_Params.WindowEnd And fld_Params.WindowEnd = fld_Params.MatDate) Then
        enu_DetailedCat = EuropeanBar
    Else
        enu_DetailedCat = DetailedCategory()
    End If
    Select Case enu_DetailedCat
        Case DetailedCat.None, DetailedCat.CashFlow: bln_Output = False
        Case Else: bln_Output = True
    End Select

    ContainsLiveOption = bln_Output
End Property

Private Property Get DetailedCategory() As DetailedCat
    ' ## Determine option category for valuation and output purposes
    Dim enu_Output As DetailedCat: enu_Output = DetailedCat.None
    Dim bln_IsKnockOut As Boolean: bln_IsKnockOut = fld_Params.IsKnockOut
    Dim int_ID_Sides As Integer


    If (fld_Params.WindowStart = fld_Params.WindowEnd And fld_Params.WindowEnd = fld_Params.MatDate) Then
        'to handle European Barrier
        enu_Output = DetailedCat.EuropeanBar
    Else
        If BarrierAlreadyHit() = True Then
            If bln_IsKnockOut = False Then
                enu_Output = 10 + enu_Type_Payoff
            Else
                enu_Output = DetailedCat.None
            End If
        ElseIf MaturedAlready() = True Then
            ' Barrier not hit at maturity
            If bln_IsKnockOut = True Then
                enu_Output = DetailedCat.CashFlow
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
     End If

    DetailedCategory = enu_Output
End Property


' ## METHODS - GREEKS
Public Function Calc_DV01(str_curve As String, Optional int_PillarIndex As Integer = 0) As Double
    ' ## Return DV01 sensitivity to the specified curve
    Dim dbl_Output As Double
    If dic_CurveDependencies.Exists(str_curve) Then
        Dim bln_ZShiftsEnabled_DiscSpot As Boolean: bln_ZShiftsEnabled_DiscSpot = fxs_Spots.ZShiftsEnabled_DiscSpot
        fxs_Spots.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

        ' Store base value
        Dim dbl_Val_Up As Double, dbl_Val_Unch As Double ', csh_AbsShift As New CurveDaysShift

        ' Store shifted value
        Call SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        dbl_Val_Up = Me.PnL

        ' Clear temporary shifts
        Call SetCurveState(str_curve, CurveState_IRC.Final)
        dbl_Val_Unch = Me.PnL

        ' Revert to original setting
        fxs_Spots.ZShiftsEnabled_DiscSpot = bln_ZShiftsEnabled_DiscSpot

        ' Calculate by finite differencing and convert to PnL currency
        dbl_Output = dbl_Val_Up - dbl_Val_Unch
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

        ' Store base value
        Dim dbl_Val_Up As Double, dbl_Val_Down As Double, dbl_Val_Unch As Double

        ' Store shifted value
        Call SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        dbl_Val_Up = Me.PnL

        Call SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        dbl_Val_Down = Me.PnL

        ' Clear temporary shifts
        Call SetCurveState(str_curve, CurveState_IRC.Final)
        dbl_Val_Unch = Me.PnL

        ' Revert to original setting
        fxs_Spots.ZShiftsEnabled_DiscSpot = bln_ZShiftsEnabled_DiscSpot

        ' Calculate by finite differencing and convert to PnL currency
        dbl_Output = dbl_Val_Up + dbl_Val_Down - 2 * dbl_Val_Unch
    Else
        dbl_Output = 0
    End If

    Calc_DV02 = dbl_Output
End Function

Public Function Calc_Delta() As Double

    ' Added by Dennis Foong on 18th January 2016

    ' ## Get discounted value in the PnL currency

    ' Prepare values for valuing option
    Dim enu_Detailed As DetailedCat: enu_Detailed = DetailedCategory()
    '##Matt Edit
'    If (fld_Params.WindowStart = fld_Params.WindowEnd And fld_Params.WindowEnd = fld_Params.MatDate) Then
'        enu_Detailed = EuropeanBar
'    Else
'        enu_Detailed = DetailedCategory()
'    End If

    '##Matt Edit end
    Dim dbl_Spot As Double: dbl_Spot = Me.Spot
    Dim dbl_Fwd As Double: dbl_Fwd = Me.Forward
    Dim dbl_SmileStrike_Orig As Double, dbl_SmileStrike_Knocked As Double
    If fld_Params.IsSmile_Orig = True Then dbl_SmileStrike_Orig = fld_Params.strike Else dbl_SmileStrike_Orig = -1
    If fld_Params.IsSmile_IfKnocked = True Then dbl_SmileStrike_Knocked = fld_Params.strike Else dbl_SmileStrike_Knocked = -1
    Dim dbl_VolPct_XY As Double: dbl_VolPct_XY = GetVol(VolPair.XY, dbl_SmileStrike_Orig)
    Dim dbl_UnitVal As Double

    ' Calculate discount factors
    Dim dbl_DF_Option As Double: dbl_DF_Option = irc_DiscCurve.Lookup_Rate(lng_SpotDate, lng_MatSpotDate, "DF", , , False)
    Dim dbl_DF_Spot As Double
    If lng_SpotDate < lng_MatSpotDate Then
        dbl_DF_Spot = irc_SpotDiscCurve.Lookup_Rate(lng_ValDate, lng_SpotDate, "DF", , , False)
    Else
        dbl_DF_Spot = irc_SpotDiscCurve.Lookup_Rate(lng_ValDate, lng_MatSpotDate, "DF", , , False)
    End If

    ' Calculate drift-adjusted forward
    Dim dbl_ATMVol_XY As Double, dbl_ATMVol_XQ As Double, dbl_ATMVol_YQ As Double, dbl_Correl_XQ As Double
    Dim dbl_DriftAdjFactor As Double
    If bln_IsQuanto = True Then
        ' Perform drift adjustment on forward
        dbl_ATMVol_XY = GetVol(VolPair.XY, -1) / 100
        dbl_ATMVol_XQ = GetVol(VolPair.XQ, -1) / 100
        dbl_ATMVol_YQ = GetVol(VolPair.YQ, -1) / 100
        dbl_Correl_XQ = Calc_QuantoCorrel(dbl_ATMVol_XY, dbl_ATMVol_XQ, dbl_ATMVol_YQ)
        dbl_DriftAdjFactor = Calc_DriftAdjFactor(dbl_ATMVol_XY, dbl_ATMVol_YQ, dbl_Correl_XQ, dbl_TimeToMat)
        dbl_Fwd = dbl_Fwd * dbl_DriftAdjFactor
    End If

    ' Calculate FX conversion factor
    Dim dbl_DiscFXSpot As Double
    If bln_IsQuanto = True Then
        dbl_DiscFXSpot = fxs_Spots.Lookup_DiscSpot(fld_Params.CCY_Payout, fld_Params.CCY_PnL)
    Else
        dbl_DiscFXSpot = fxs_Spots.Lookup_DiscSpot(fld_Params.CCY_Dom, fld_Params.CCY_PnL)
    End If

    ' Calculate option delta as at spot date in the domestic currency
    Dim dbl_UnitValBase As Double, dbl_UnitValSpotShockUp As Double, dbl_UnitValSpotShockDown As Double
    Dim dbl_ShockSize As Double: dbl_ShockSize = 0.01
    Dim dbl_Output As Double

    Dim str_TargetCCy As String
    If fld_Params.CCY_Fgn = "USD" Then
        str_TargetCCy = fld_Params.CCY_Dom
    ElseIf fld_Params.CCY_Dom = "USD" Then
        str_TargetCCy = fld_Params.CCY_Fgn
    ElseIf (fxs_Spots.Lookup_Quotation(fld_Params.CCY_Dom) = "INDIRECT" And fxs_Spots.Lookup_Quotation(fld_Params.CCY_Fgn) = "INDIRECT") Then
        str_TargetCCy = fld_Params.CCY_Fgn
    Else
        str_TargetCCy = fld_Params.CCY_Dom
    End If

    Select Case enu_Detailed

        Case DetailedCat.CashFlow

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_UnitValSpotShockUp = WorksheetFunction.Max(fld_Params.OptDirection * (dbl_Spot - fld_Params.strike), 0) * dbl_DF_Spot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_UnitValSpotShockDown = WorksheetFunction.Max(fld_Params.OptDirection * (dbl_Spot - fld_Params.strike), 0) * dbl_DF_Spot

        Case DetailedCat.SingleAmerican

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, fld_Params.strike)
            dbl_UnitValSpotShockUp = Calc_BSPrice_SingleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, dbl_SingleBarrier, (enu_Type_Barrier = BarType.UpperBar), _
                fld_Params.IsKnockOut) * dbl_DF_Option * dbl_DF_Spot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, fld_Params.strike)
            dbl_UnitValSpotShockDown = Calc_BSPrice_SingleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, dbl_SingleBarrier, (enu_Type_Barrier = BarType.UpperBar), _
                fld_Params.IsKnockOut) * dbl_DF_Option * dbl_DF_Spot

        Case DetailedCat.SingleAmericanQ

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, fld_Params.strike)
            dbl_UnitValSpotShockUp = Calc_BSPrice_SingleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, dbl_SingleBarrier, (enu_Type_Barrier = BarType.UpperBar), _
                fld_Params.IsKnockOut) * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, fld_Params.strike)
            dbl_UnitValSpotShockDown = Calc_BSPrice_SingleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, dbl_SingleBarrier, (enu_Type_Barrier = BarType.UpperBar), _
                fld_Params.IsKnockOut) * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot

        Case DetailedCat.DoubleAmerican

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, fld_Params.strike)
            dbl_UnitValSpotShockUp = Calc_BSPrice_DoubleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, fld_Params.LowerBar, fld_Params.UpperBar, fld_Params.IsKnockOut) _
                 * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, fld_Params.strike)
            dbl_UnitValSpotShockDown = Calc_BSPrice_DoubleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, fld_Params.LowerBar, fld_Params.UpperBar, fld_Params.IsKnockOut) _
                 * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot

        Case DetailedCat.DoubleAmericanQ

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, fld_Params.strike)
            dbl_UnitValSpotShockUp = Calc_BSPrice_DoubleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, fld_Params.LowerBar, fld_Params.UpperBar, fld_Params.IsKnockOut) _
                 * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, fld_Params.strike)
            dbl_UnitValSpotShockDown = Calc_BSPrice_DoubleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, fld_Params.LowerBar, fld_Params.UpperBar, fld_Params.IsKnockOut) _
                 * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot

        Case DetailedCat.Vanilla

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_UnitValSpotShockUp = Calc_BSPrice_Vanilla(fld_Params.OptDirection, dbl_Fwd, fld_Params.strike, _
                dbl_TimeToMat, GetVol(VolPair.XY, dbl_SmileStrike_Knocked)) * dbl_DF_Option * dbl_DF_Spot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_UnitValSpotShockDown = Calc_BSPrice_Vanilla(fld_Params.OptDirection, dbl_Fwd, fld_Params.strike, _
                dbl_TimeToMat, GetVol(VolPair.XY, dbl_SmileStrike_Knocked)) * dbl_DF_Option * dbl_DF_Spot

        Case DetailedCat.Quanto

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_UnitValSpotShockUp = Calc_BSPrice_Vanilla(fld_Params.OptDirection, dbl_Fwd, fld_Params.strike, dbl_TimeToMat, _
                GetVol(VolPair.XY, dbl_SmileStrike_Knocked)) * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_UnitValSpotShockDown = Calc_BSPrice_Vanilla(fld_Params.OptDirection, dbl_Fwd, fld_Params.strike, dbl_TimeToMat, _
                GetVol(VolPair.XY, dbl_SmileStrike_Knocked)) * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot

        '##Matt Edit
        Case DetailedCat.EuropeanBar

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_UnitVal = EuropeanBarType(fld_Params.IsKnockOut, fld_Params.strike, fld_Params.OptDirection, fld_Params.LowerBar, fld_Params.UpperBar)
            If (int_Sign_BS = 1 And int_Sign_BS * dbl_UnitVal < 0) Or (int_Sign_BS = -1 And int_Sign_BS * dbl_UnitVal > 0) Then
                dbl_UnitVal = 0
            End If
            dbl_UnitValSpotShockUp = dbl_UnitVal * dbl_DF_Option * dbl_DF_Spot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_UnitVal = EuropeanBarType(fld_Params.IsKnockOut, fld_Params.strike, fld_Params.OptDirection, fld_Params.LowerBar, fld_Params.UpperBar)
            If (int_Sign_BS = 1 And int_Sign_BS * dbl_UnitVal < 0) Or (int_Sign_BS = -1 And int_Sign_BS * dbl_UnitVal > 0) Then
                dbl_UnitVal = 0
            End If
            dbl_UnitValSpotShockDown = dbl_UnitVal * dbl_DF_Option * dbl_DF_Spot

        '##Matt Edit end
        Case Else:

            dbl_UnitValSpotShockUp = 0
            dbl_UnitValSpotShockDown = 0

    End Select

    Call fxs_Spots.Scen_ApplyBase
    ' Convert to PnL currency at valuation date, then output final value
    dbl_UnitValSpotShockUp = dbl_UnitValSpotShockUp * int_Sign_BS * dbl_DiscFXSpot
    dbl_UnitValSpotShockDown = dbl_UnitValSpotShockDown * int_Sign_BS * dbl_DiscFXSpot
    dbl_Output = fld_Params.Notional_Fgn * (dbl_UnitValSpotShockUp - dbl_UnitValSpotShockDown) / (2 * dbl_ShockSize / 100 * dbl_Spot)
    Calc_Delta = dbl_Output

End Function
Public Function Calc_Gamma() As Double

    ' Added by Dennis Foong on 18th January 2016

    ' ## Get discounted value in the PnL currency

    ' Prepare values for valuing option
    Dim enu_Detailed As DetailedCat: enu_Detailed = DetailedCategory()
    '##Matt Edit
'    If (fld_Params.WindowStart = fld_Params.WindowEnd And fld_Params.WindowEnd = fld_Params.MatDate) Then
'        enu_Detailed = EuropeanBar
'    Else
'        enu_Detailed = DetailedCategory()
'    End If

    '##Matt Edit end
    Dim dbl_Spot As Double: dbl_Spot = Me.Spot
    Dim dbl_Fwd As Double: dbl_Fwd = Me.Forward
    Dim dbl_SmileStrike_Orig As Double, dbl_SmileStrike_Knocked As Double
    If fld_Params.IsSmile_Orig = True Then dbl_SmileStrike_Orig = fld_Params.strike Else dbl_SmileStrike_Orig = -1
    If fld_Params.IsSmile_IfKnocked = True Then dbl_SmileStrike_Knocked = fld_Params.strike Else dbl_SmileStrike_Knocked = -1
    Dim dbl_VolPct_XY As Double: dbl_VolPct_XY = GetVol(VolPair.XY, dbl_SmileStrike_Orig)
    Dim dbl_UnitVal As Double

    ' Calculate discount factors
    Dim dbl_DF_Option As Double: dbl_DF_Option = irc_DiscCurve.Lookup_Rate(lng_SpotDate, lng_MatSpotDate, "DF", , , False)
    Dim dbl_DF_Spot As Double
    If lng_SpotDate < lng_MatSpotDate Then
        dbl_DF_Spot = irc_SpotDiscCurve.Lookup_Rate(lng_ValDate, lng_SpotDate, "DF", , , False)
    Else
        dbl_DF_Spot = irc_SpotDiscCurve.Lookup_Rate(lng_ValDate, lng_MatSpotDate, "DF", , , False)
    End If

    ' Calculate drift-adjusted forward
    Dim dbl_ATMVol_XY As Double, dbl_ATMVol_XQ As Double, dbl_ATMVol_YQ As Double, dbl_Correl_XQ As Double
    Dim dbl_DriftAdjFactor As Double
    If bln_IsQuanto = True Then
        ' Perform drift adjustment on forward
        dbl_ATMVol_XY = GetVol(VolPair.XY, -1) / 100
        dbl_ATMVol_XQ = GetVol(VolPair.XQ, -1) / 100
        dbl_ATMVol_YQ = GetVol(VolPair.YQ, -1) / 100
        dbl_Correl_XQ = Calc_QuantoCorrel(dbl_ATMVol_XY, dbl_ATMVol_XQ, dbl_ATMVol_YQ)
        dbl_DriftAdjFactor = Calc_DriftAdjFactor(dbl_ATMVol_XY, dbl_ATMVol_YQ, dbl_Correl_XQ, dbl_TimeToMat)
        dbl_Fwd = dbl_Fwd * dbl_DriftAdjFactor
    End If

    ' Calculate FX conversion factor
    Dim dbl_DiscFXSpot As Double
    If bln_IsQuanto = True Then
        dbl_DiscFXSpot = fxs_Spots.Lookup_DiscSpot(fld_Params.CCY_Payout, fld_Params.CCY_PnL)
    Else
        dbl_DiscFXSpot = fxs_Spots.Lookup_DiscSpot(fld_Params.CCY_Dom, fld_Params.CCY_PnL)
    End If

    ' Calculate option delta as at spot date in the domestic currency
    Dim dbl_UnitValBase As Double, dbl_UnitValSpotShockUp As Double, dbl_UnitValSpotShockDown As Double
    Dim dbl_ShockSize As Double: dbl_ShockSize = 0.01
    Dim dbl_Output As Double

    Dim str_TargetCCy As String
    If fld_Params.CCY_Fgn = "USD" Then
        str_TargetCCy = fld_Params.CCY_Dom
    ElseIf fld_Params.CCY_Dom = "USD" Then
        str_TargetCCy = fld_Params.CCY_Fgn
    ElseIf (fxs_Spots.Lookup_Quotation(fld_Params.CCY_Dom) = "INDIRECT" And fxs_Spots.Lookup_Quotation(fld_Params.CCY_Fgn) = "INDIRECT") Then
        str_TargetCCy = fld_Params.CCY_Fgn
    Else
        str_TargetCCy = fld_Params.CCY_Dom
    End If

    Select Case enu_Detailed

        Case DetailedCat.CashFlow

            ' Base
            dbl_Spot = Spot()
            dbl_UnitValBase = WorksheetFunction.Max(fld_Params.OptDirection * (dbl_Spot - fld_Params.strike), 0) * dbl_DF_Spot

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_UnitValSpotShockUp = WorksheetFunction.Max(fld_Params.OptDirection * (dbl_Spot - fld_Params.strike), 0) * dbl_DF_Spot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_UnitValSpotShockDown = WorksheetFunction.Max(fld_Params.OptDirection * (dbl_Spot - fld_Params.strike), 0) * dbl_DF_Spot

        Case DetailedCat.SingleAmerican

            ' Base
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, fld_Params.strike)
            dbl_UnitValBase = Calc_BSPrice_SingleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, dbl_SingleBarrier, (enu_Type_Barrier = BarType.UpperBar), _
                fld_Params.IsKnockOut) * dbl_DF_Option * dbl_DF_Spot

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, fld_Params.strike)
            dbl_UnitValSpotShockUp = Calc_BSPrice_SingleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, dbl_SingleBarrier, (enu_Type_Barrier = BarType.UpperBar), _
                fld_Params.IsKnockOut) * dbl_DF_Option * dbl_DF_Spot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, fld_Params.strike)
            dbl_UnitValSpotShockDown = Calc_BSPrice_SingleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, dbl_SingleBarrier, (enu_Type_Barrier = BarType.UpperBar), _
                fld_Params.IsKnockOut) * dbl_DF_Option * dbl_DF_Spot

        Case DetailedCat.SingleAmericanQ

            ' Base
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, fld_Params.strike)
            dbl_UnitValBase = Calc_BSPrice_SingleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, dbl_SingleBarrier, (enu_Type_Barrier = BarType.UpperBar), _
                fld_Params.IsKnockOut) * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, fld_Params.strike)
            dbl_UnitValSpotShockUp = Calc_BSPrice_SingleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, dbl_SingleBarrier, (enu_Type_Barrier = BarType.UpperBar), _
                fld_Params.IsKnockOut) * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, fld_Params.strike)
            dbl_UnitValSpotShockDown = Calc_BSPrice_SingleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, dbl_SingleBarrier, (enu_Type_Barrier = BarType.UpperBar), _
                fld_Params.IsKnockOut) * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot

        Case DetailedCat.DoubleAmerican

            ' Base
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, fld_Params.strike)
            dbl_UnitValBase = Calc_BSPrice_DoubleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, fld_Params.LowerBar, fld_Params.UpperBar, fld_Params.IsKnockOut) _
                 * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, fld_Params.strike)
            dbl_UnitValSpotShockUp = Calc_BSPrice_DoubleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, fld_Params.LowerBar, fld_Params.UpperBar, fld_Params.IsKnockOut) _
                 * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, fld_Params.strike)
            dbl_UnitValSpotShockDown = Calc_BSPrice_DoubleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, fld_Params.LowerBar, fld_Params.UpperBar, fld_Params.IsKnockOut) _
                 * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot

        Case DetailedCat.DoubleAmericanQ

            ' Base
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, fld_Params.strike)
            dbl_UnitValBase = Calc_BSPrice_DoubleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, fld_Params.LowerBar, fld_Params.UpperBar, fld_Params.IsKnockOut) _
                 * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, fld_Params.strike)
            dbl_UnitValSpotShockUp = Calc_BSPrice_DoubleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, fld_Params.LowerBar, fld_Params.UpperBar, fld_Params.IsKnockOut) _
                 * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, fld_Params.strike)
            dbl_UnitValSpotShockDown = Calc_BSPrice_DoubleAmericanBar(fld_Params.OptDirection, dbl_Spot, dbl_Fwd, fld_Params.strike, _
                dbl_VolPct_XY, dbl_TimeToMat, fld_Params.LowerBar, fld_Params.UpperBar, fld_Params.IsKnockOut) _
                 * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot

        Case DetailedCat.Vanilla

            ' Base
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_UnitValBase = Calc_BSPrice_Vanilla(fld_Params.OptDirection, dbl_Fwd, fld_Params.strike, _
                dbl_TimeToMat, GetVol(VolPair.XY, dbl_SmileStrike_Knocked)) * dbl_DF_Option * dbl_DF_Spot

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_UnitValSpotShockUp = Calc_BSPrice_Vanilla(fld_Params.OptDirection, dbl_Fwd, fld_Params.strike, _
                dbl_TimeToMat, GetVol(VolPair.XY, dbl_SmileStrike_Knocked)) * dbl_DF_Option * dbl_DF_Spot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_UnitValSpotShockDown = Calc_BSPrice_Vanilla(fld_Params.OptDirection, dbl_Fwd, fld_Params.strike, _
                dbl_TimeToMat, GetVol(VolPair.XY, dbl_SmileStrike_Knocked)) * dbl_DF_Option * dbl_DF_Spot

        Case DetailedCat.Quanto

            ' Shock up
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_UnitValBase = Calc_BSPrice_Vanilla(fld_Params.OptDirection, dbl_Fwd, fld_Params.strike, dbl_TimeToMat, _
                GetVol(VolPair.XY, dbl_SmileStrike_Knocked)) * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_UnitValSpotShockUp = Calc_BSPrice_Vanilla(fld_Params.OptDirection, dbl_Fwd, fld_Params.strike, dbl_TimeToMat, _
                GetVol(VolPair.XY, dbl_SmileStrike_Knocked)) * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_UnitValSpotShockDown = Calc_BSPrice_Vanilla(fld_Params.OptDirection, dbl_Fwd, fld_Params.strike, dbl_TimeToMat, _
                GetVol(VolPair.XY, dbl_SmileStrike_Knocked)) * fld_Params.QuantoFactor * dbl_DF_Option * dbl_DF_Spot

        '##Matt Edit
        Case DetailedCat.EuropeanBar

            ' Base
            dbl_UnitVal = EuropeanBarType(fld_Params.IsKnockOut, fld_Params.strike, fld_Params.OptDirection, fld_Params.LowerBar, fld_Params.UpperBar)
            If (int_Sign_BS = 1 And int_Sign_BS * dbl_UnitVal < 0) Or (int_Sign_BS = -1 And int_Sign_BS * dbl_UnitVal > 0) Then
                dbl_UnitVal = 0
            End If
            dbl_UnitValBase = dbl_UnitVal * dbl_DF_Option * dbl_DF_Spot

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_UnitVal = EuropeanBarType(fld_Params.IsKnockOut, fld_Params.strike, fld_Params.OptDirection, fld_Params.LowerBar, fld_Params.UpperBar)
            If (int_Sign_BS = 1 And int_Sign_BS * dbl_UnitVal < 0) Or (int_Sign_BS = -1 And int_Sign_BS * dbl_UnitVal > 0) Then
                dbl_UnitVal = 0
            End If
            dbl_UnitValSpotShockUp = dbl_UnitVal * dbl_DF_Option * dbl_DF_Spot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_UnitVal = EuropeanBarType(fld_Params.IsKnockOut, fld_Params.strike, fld_Params.OptDirection, fld_Params.LowerBar, fld_Params.UpperBar)
            If (int_Sign_BS = 1 And int_Sign_BS * dbl_UnitVal < 0) Or (int_Sign_BS = -1 And int_Sign_BS * dbl_UnitVal > 0) Then
                dbl_UnitVal = 0
            End If
            dbl_UnitValSpotShockDown = dbl_UnitVal * dbl_DF_Option * dbl_DF_Spot

        '##Matt Edit end
        Case Else:

            dbl_UnitVal = 0
            dbl_UnitValSpotShockUp = 0
            dbl_UnitValSpotShockDown = 0

    End Select

    Call fxs_Spots.Scen_ApplyBase
    dbl_Spot = Spot()
    ' Convert to PnL currency at valuation date, then output final value
    dbl_UnitValBase = dbl_UnitValBase * int_Sign_BS * dbl_DiscFXSpot
    dbl_UnitValSpotShockUp = dbl_UnitValSpotShockUp * int_Sign_BS * dbl_DiscFXSpot
    dbl_UnitValSpotShockDown = dbl_UnitValSpotShockDown * int_Sign_BS * dbl_DiscFXSpot
    dbl_Output = fld_Params.Notional_Fgn * (dbl_UnitValSpotShockUp + dbl_UnitValSpotShockDown - 2 * dbl_UnitValBase) / ((dbl_ShockSize / 100 * dbl_Spot) ^ 2) * dbl_Spot / 100
    Calc_Gamma = dbl_Output

End Function

Public Function Calc_Vega(enu_Type As CurveType, str_curve As String) As Double
    ' ## Return sensitivity to a 1 vol increase in the vol
    Dim dbl_Output As Double: dbl_Output = 0

    ' Prepare values for valuing option
    Dim enu_Detailed As DetailedCat: enu_Detailed = DetailedCategory()

    If enu_Type = CurveType.FXV Then

        Dim dbl_Val_Up As Double, dbl_Val_Unch As Double, dbl_Val_Down  As Double
        Dim bln_IsShifted As Boolean

        If enu_Detailed = DetailedCat.EuropeanBar Then
            'To handle European Barrier
            bln_IsShifted = ApplyVolShift(str_curve, 0.01)
            If bln_IsShifted = True Then
                dbl_Val_Up = Me.PnL

                Call ApplyVolShift(str_curve, -0.01)
                dbl_Val_Down = Me.PnL

                ' Clear temporary shifts from the vol curve
                Call ApplyVolShift(str_curve, 0)

                ' Calculate by finite differencing and convert to PnL currency
                dbl_Output = (dbl_Val_Up - dbl_Val_Down) * 50
            End If
        Else
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


    Dim enu_Detailed As DetailedCat: enu_Detailed = DetailedCategory()

    If fxv_Vols_XY.CurveName = str_curve Then

        If enu_Detailed = DetailedCat.EuropeanBar Then
            fxv_Vols_XY.VolShift_Sens = dbl_ShiftSize
        Else
            dbl_VolShiftXY_Sens = dbl_ShiftSize
            'fxv_Vols_XY.VolShift_Sens = dbl_ShiftSize
        End If
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
Private Function GetVol(enu_Pair As VolPair, dbl_LookupStrike As Double) As Double
    ' Determine vol curve to look up
    Dim dbl_Output As Double
    Dim fxv_VolCurve As Data_FXVols
    Select Case enu_Pair
        Case VolPair.XY: Set fxv_VolCurve = fxv_Vols_XY
        Case VolPair.XQ: Set fxv_VolCurve = fxv_Vols_XQ
        Case VolPair.YQ: Set fxv_VolCurve = fxv_Vols_YQ
    End Select

    ' Obtain the vol from the curve
    If dbl_LookupStrike <> -1 Then
        dbl_Output = fxv_VolCurve.Lookup_SmileVol(fld_Params.MatDate, dbl_LookupStrike, , fld_Params.IsRescaling)
    Else
        dbl_Output = fxv_VolCurve.Lookup_ATMVol(fld_Params.MatDate)
    End If

    ' Add on any sensitivity shifts

    Select Case enu_Pair
        Case VolPair.XY: dbl_Output = dbl_Output + dbl_VolShiftXY_Sens
        Case VolPair.XQ: dbl_Output = dbl_Output + dbl_VolShiftXQ_Sens
        Case VolPair.YQ: dbl_Output = dbl_Output + dbl_VolShiftYQ_Sens
    End Select

    GetVol = dbl_Output
End Function

'##Matt Edit
Private Function EuropeanBarType(IsKnockOut As Boolean, strike As Double, OptDirection As OptionDirection, LowerBar As Double, UpperBar As Double) As Double
    Dim Van_Strike As Double, Van_Barrier As Double, Van_ShiftedBarrier As Double, dbl_CSDigi As Double, dbl_Output As Double
    Dim Barrier As Double, ShiftedBarrier As Double
    Dim kvol As Double, bvol As Double, sbvol As Double, dbl_Fwd As Double
    dbl_Fwd = Me.Forward

    If UpperBar = -1 Then
        Barrier = LowerBar
    ElseIf LowerBar = -1 Then
        Barrier = UpperBar
    End If

    ShiftedBarrier = Barrier * (1 + OptDirection * 0.0001)
    kvol = GetVol(VolPair.XY, strike)
    bvol = GetVol(VolPair.XY, Barrier)
    sbvol = GetVol(VolPair.XY, ShiftedBarrier)

    Van_Strike = Calc_BSPrice_Vanilla(OptDirection, dbl_Fwd, strike, dbl_TimeToMat, kvol)
    Van_Barrier = Calc_BSPrice_Vanilla(OptDirection, dbl_Fwd, Barrier, dbl_TimeToMat, bvol)
    Van_ShiftedBarrier = Calc_BSPrice_Vanilla(OptDirection, dbl_Fwd, ShiftedBarrier, dbl_TimeToMat, sbvol)

    dbl_CSDigi = (Van_Barrier - Van_ShiftedBarrier) / (Barrier * 0.0001)

    If (OptDirection = PutOpt And LowerBar = -1 And IsKnockOut = False And OptDirection * strike < OptDirection * Barrier) _
    Or (OptDirection = PutOpt And UpperBar = -1 And IsKnockOut = True And OptDirection * strike < OptDirection * Barrier) _
    Or (OptDirection = CallOpt And UpperBar = -1 And IsKnockOut = False And OptDirection * strike < OptDirection * Barrier) _
    Or (OptDirection = CallOpt And LowerBar = -1 And IsKnockOut = True And OptDirection * strike < OptDirection * Barrier) Then

       dbl_Output = Van_Strike - Van_Barrier - OptDirection * (Barrier - strike) * dbl_CSDigi

    ElseIf (OptDirection = PutOpt And LowerBar = -1 And IsKnockOut = False And OptDirection * strike >= OptDirection * Barrier) _
    Or (OptDirection = PutOpt And UpperBar = -1 And IsKnockOut = True And OptDirection * strike >= OptDirection * Barrier) _
    Or (OptDirection = CallOpt And UpperBar = -1 And IsKnockOut = False And OptDirection * strike >= OptDirection * Barrier) _
    Or (OptDirection = CallOpt And LowerBar = -1 And IsKnockOut = True And OptDirection * strike >= OptDirection * Barrier) Then

        dbl_Output = 0

    ElseIf (OptDirection = PutOpt And LowerBar = -1 And IsKnockOut = True And OptDirection * strike < OptDirection * Barrier) _
    Or (OptDirection = PutOpt And UpperBar = -1 And IsKnockOut = False And OptDirection * strike < OptDirection * Barrier) _
    Or (OptDirection = CallOpt And UpperBar = -1 And IsKnockOut = True And OptDirection * strike < OptDirection * Barrier) _
    Or (OptDirection = CallOpt And LowerBar = -1 And IsKnockOut = False And OptDirection * strike < OptDirection * Barrier) Then

        dbl_Output = Van_Barrier - OptDirection * (strike - Barrier) * dbl_CSDigi

    ElseIf (OptDirection = PutOpt And LowerBar = -1 And IsKnockOut = True And OptDirection * strike >= OptDirection * Barrier) _
    Or (OptDirection = PutOpt And UpperBar = -1 And IsKnockOut = False And OptDirection * strike >= OptDirection * Barrier) _
    Or (OptDirection = CallOpt And UpperBar = -1 And IsKnockOut = True And OptDirection * strike >= OptDirection * Barrier) _
    Or (OptDirection = CallOpt And LowerBar = -1 And IsKnockOut = False And OptDirection * strike >= OptDirection * Barrier) Then

        dbl_Output = Van_Strike
    End If
EuropeanBarType = dbl_Output
End Function
'##Matt Edit end

' ## METHODS - CALCULATION DETAILS
Public Sub OutputReport(wks_output As Worksheet)
    wks_output.Cells.Clear

    Dim rng_TopLeft As Range: Set rng_TopLeft = wks_output.Range("A1")
    Dim str_CurrencyFormat As String: str_CurrencyFormat = Gather_CurrencyFormat()
    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    Dim int_ActiveRow As Integer: int_ActiveRow = 0
    Dim enu_Type_Detailed As DetailedCat ': enu_Type_Detailed = DetailedCategory()
    '##Matt Edit
    If (fld_Params.WindowStart = fld_Params.WindowEnd And fld_Params.WindowEnd = fld_Params.MatDate) Then
        enu_Type_Detailed = EuropeanBar
    Else
        enu_Type_Detailed = DetailedCategory()
    End If
    '##Matt Edit end

    Dim bln_Next_Display As Boolean, str_Next_Formula As String
    Dim dic_Addresses As Dictionary: Set dic_Addresses = New Dictionary
    dic_Addresses.CompareMode = CompareMethod.TextCompare
    Dim dbl_SmileStrike_Orig As Double, dbl_SmileStrike_Knocked As Double
    If fld_Params.IsSmile_Orig = True Then dbl_SmileStrike_Orig = fld_Params.strike Else dbl_SmileStrike_Orig = -1
    If fld_Params.IsSmile_IfKnocked = True Then dbl_SmileStrike_Knocked = fld_Params.strike Else dbl_SmileStrike_Knocked = -1

    ' Categorization
    Dim bln_ContainsLiveOption As Boolean: bln_ContainsLiveOption = ContainsLiveOption()
    Dim bln_ContainsMaturedOption As Boolean: bln_ContainsMaturedOption = (enu_Type_Detailed = DetailedCat.CashFlow)
    Dim bln_ContainsLiveSingle As Boolean: bln_ContainsLiveSingle = (enu_Type_Detailed > 100 And enu_Type_Detailed < 200) Or (enu_Type_Detailed >= 300)
    Dim bln_ContainsLiveDouble As Boolean: bln_ContainsLiveDouble = (enu_Type_Detailed > 200 And enu_Type_Detailed < 300)
    Dim bln_European As Boolean: bln_European = enu_Type_Detailed >= 300


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

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Status:"
        Select Case enu_Type_Detailed
            Case DetailedCat.None: .Offset(int_ActiveRow, 1).Value = "Knocked out"
            Case DetailedCat.CashFlow: .Offset(int_ActiveRow, 1).Value = "Matured"
            Case DetailedCat.Vanilla, DetailedCat.Quanto
                .Offset(int_ActiveRow, 1).Value = "Knocked in"
            Case Else: .Offset(int_ActiveRow, 1).Value = "Live"
        End Select

        If enu_Type_Detailed <> DetailedCat.None Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Fgn Notional:"
            .Offset(int_ActiveRow, 1).Value = fld_Params.Notional_Fgn
            .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
            Call dic_Addresses.Add("Notional", .Offset(int_ActiveRow, 1).Address(False, False))
            .Offset(int_ActiveRow, 2).Value = fld_Params.CCY_Fgn

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Position:"
            If fld_Params.IsBuy = True Then
                .Offset(int_ActiveRow, 1).Value = "B"
            Else
                .Offset(int_ActiveRow, 1).Value = "S"
            End If
            Call dic_Addresses.Add("Position", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Payout:"
            If fld_Params.OptDirection = OptionDirection.CallOpt Then
                .Offset(int_ActiveRow, 1).Value = "Call"
            Else
                .Offset(int_ActiveRow, 1).Value = "Put"
            End If
            Call dic_Addresses.Add("Payout", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Strike:"
            .Offset(int_ActiveRow, 1).Value = fld_Params.strike
            Call dic_Addresses.Add("Strike", .Offset(int_ActiveRow, 1).Address(False, False))
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

        If bln_ContainsLiveSingle = True Or bln_ContainsLiveDouble = True Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Barrier Knocks:"
            If fld_Params.IsKnockOut = True Then
                .Offset(int_ActiveRow, 1).Value = "Out"
            Else
                .Offset(int_ActiveRow, 1).Value = "In"
            End If
            Call dic_Addresses.Add("Knock_Bar", .Offset(int_ActiveRow, 1).Address(False, False))
        End If


        If (bln_ContainsLiveSingle = True And bln_European = True) Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Barrier:"
            .Offset(int_ActiveRow, 1).Value = dbl_SingleBarrier
            Call dic_Addresses.Add("Barrier", .Offset(int_ActiveRow, 1).Address(False, False))


            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Shifted Barrier:"
            .Offset(int_ActiveRow, 1).Value = dbl_SingleBarrier * (1 + fld_Params.OptDirection * 0.0001)
            Call dic_Addresses.Add("Shifted_Barrier", .Offset(int_ActiveRow, 1).Address(False, False))

        ElseIf bln_ContainsLiveSingle = True Then
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

        If bln_ContainsLiveOption = True Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Forward:"
            .Offset(int_ActiveRow, 1).Value = Me.Forward
            Call dic_Addresses.Add("Fwd", .Offset(int_ActiveRow, 1).Address(False, False))

            Select Case enu_Type_Detailed
                Case DetailedCat.Vanilla, DetailedCat.Quanto
                    int_ActiveRow = int_ActiveRow + 1
                    .Offset(int_ActiveRow, 0).Value = "Vol:"
                    .Offset(int_ActiveRow, 1).Value = GetVol(VolPair.XY, dbl_SmileStrike_Knocked)
                    Call dic_Addresses.Add("Vol", .Offset(int_ActiveRow, 1).Address(False, False))

                Case DetailedCat.EuropeanBar
                    int_ActiveRow = int_ActiveRow + 1
                    .Offset(int_ActiveRow, 0).Value = "Vol (Strike):"
                    .Offset(int_ActiveRow, 1).Value = GetVol(VolPair.XY, dbl_SmileStrike_Knocked)
                    Call dic_Addresses.Add("Vol", .Offset(int_ActiveRow, 1).Address(False, False))

                    int_ActiveRow = int_ActiveRow + 1
                    .Offset(int_ActiveRow, 0).Value = "Barrier Vol:"
                    .Offset(int_ActiveRow, 1).Value = GetVol(VolPair.XY, dbl_SingleBarrier)
                    Call dic_Addresses.Add("Vol_Barrier", .Offset(int_ActiveRow, 1).Address(False, False))

                    int_ActiveRow = int_ActiveRow + 1
                    .Offset(int_ActiveRow, 0).Value = "Shifted Barrier Vol:"
                    .Offset(int_ActiveRow, 1).Value = GetVol(VolPair.XY, dbl_SingleBarrier * (1 + fld_Params.OptDirection * 0.0001))
                    Call dic_Addresses.Add("Vol_ShiftedBarrier", .Offset(int_ActiveRow, 1).Address(False, False))

                Case Else
                    int_ActiveRow = int_ActiveRow + 1
                    .Offset(int_ActiveRow, 0).Value = "Vol:"
                    .Offset(int_ActiveRow, 1).Value = GetVol(VolPair.XY, dbl_SmileStrike_Orig)
                    Call dic_Addresses.Add("Vol", .Offset(int_ActiveRow, 1).Address(False, False))
            End Select
        End If

        If bln_IsQuanto = True Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "ATM Vol (Fgn/Dom):"
            .Offset(int_ActiveRow, 1).Value = GetVol(VolPair.XY, -1)
            Call dic_Addresses.Add("ATMVol_XY", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "ATM Vol (Fgn/Qto):"
            .Offset(int_ActiveRow, 1).Value = GetVol(VolPair.XQ, -1)
            Call dic_Addresses.Add("ATMVol_XQ", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "ATM Vol (Dom/Qto):"
            .Offset(int_ActiveRow, 1).Value = GetVol(VolPair.YQ, -1)
            Call dic_Addresses.Add("ATMVol_YQ", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Correl (Fgn/Qto):"
            .Offset(int_ActiveRow, 1).Formula = "=((" & dic_Addresses("ATMVol_XQ") & "/100)^2-(" & dic_Addresses("ATMVol_XY") _
                & "/100)^2-(" & dic_Addresses("ATMVol_YQ") & "/100)^2)/(2*" & dic_Addresses("ATMVol_XY") & "/100*" & dic_Addresses("ATMVol_YQ") & "/100)"
            Call dic_Addresses.Add("Correl_XQ", .Offset(int_ActiveRow, 1).Address(False, False))
        End If

        If bln_ContainsLiveOption = True Then
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

        If bln_IsQuanto = True Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Drift Adj Fwd:"
            .Offset(int_ActiveRow, 1).Formula = "=" & dic_Addresses("Fwd") & "*Exp(-" & dic_Addresses("Correl_XQ") & "*" _
                & dic_Addresses("ATMVol_XY") & "/100*" & dic_Addresses("ATMVol_YQ") & "/100*Calc_YearFrac(" & dic_Addresses("ValDate") _
                & "," & dic_Addresses("MatDate") & ",""ACT/365""))"
            Call dic_Addresses.Add("Fwd_DriftAdj", .Offset(int_ActiveRow, 1).Address(False, False))
        End If

        If bln_ContainsLiveOption = True Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Option DF:"
            .Offset(int_ActiveRow, 1).Formula = "=cyReadIRCurve(""" & fld_Params.Curve_Disc & """," & dic_Addresses("SpotDate") _
                & "," & dic_Addresses("DelivDate") & ",""DF"",,False)"
            Call dic_Addresses.Add("OptionDF", .Offset(int_ActiveRow, 1).Address(False, False))
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

        If bln_IsQuanto = True Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Payout CCY:"
            .Offset(int_ActiveRow, 1).Value = fld_Params.CCY_Payout
            Call dic_Addresses.Add("PayoutCCY", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Quanto Factor:"
            .Offset(int_ActiveRow, 1).Value = fld_Params.QuantoFactor
            Call dic_Addresses.Add("QuantoFactor", .Offset(int_ActiveRow, 1).Address(False, False))
        End If

        ' Output option MV
        Select Case enu_Type_Detailed
            Case DetailedCat.SingleAmerican
                bln_Next_Display = True
                str_Next_Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" _
                    & dic_Addresses("Notional") & "*Calc_BSPrice_SingleAmericanBar(IF(" & dic_Addresses("Payout") & "=""Call"",1,-1)," _
                    & dic_Addresses("Spot") & "," & dic_Addresses("Fwd") & "," & dic_Addresses("Strike") & "," _
                    & dic_Addresses("Vol") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," & dic_Addresses("MatDate") & ",""ACT/365"")," _
                    & dic_Addresses("Barrier") & "," & dic_Addresses("Direction") & "=""Up""," & dic_Addresses("Knock_Bar") _
                    & "=""Out"")*" & dic_Addresses("OptionDF") & "*" & dic_Addresses("SpotDF") & "*cyGetFXDiscSpot(" _
                    & dic_Addresses("DomCCY") & "," & dic_Addresses("PnLCCY") & ")"
            Case DetailedCat.SingleAmericanQ
                bln_Next_Display = True
                str_Next_Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" _
                    & dic_Addresses("Notional") & "*Calc_BSPrice_SingleAmericanBar(IF(" & dic_Addresses("Payout") & "=""Call"",1,-1)," _
                    & dic_Addresses("Spot") & "," & dic_Addresses("Fwd_DriftAdj") & "," & dic_Addresses("Strike") & "," _
                    & dic_Addresses("Vol") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," & dic_Addresses("MatDate") & ",""ACT/365"")," _
                    & dic_Addresses("Barrier") & "," & dic_Addresses("Direction") & "=""Up""," & dic_Addresses("Knock_Bar") _
                    & "=""Out"")*" & dic_Addresses("QuantoFactor") & "*" & dic_Addresses("OptionDF") & "*" & dic_Addresses("SpotDF") _
                    & "*cyGetFXDiscSpot(" & dic_Addresses("PayoutCCY") & "," & dic_Addresses("PnLCCY") & ")"
            Case DetailedCat.DoubleAmerican
                bln_Next_Display = True
                str_Next_Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" _
                    & dic_Addresses("Notional") & "*Calc_BSPrice_DoubleAmericanBar(IF(" & dic_Addresses("Payout") & "=""Call"",1,-1)," _
                    & dic_Addresses("Spot") & "," & dic_Addresses("Fwd") & "," & dic_Addresses("Strike") & "," _
                    & dic_Addresses("Vol") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," & dic_Addresses("MatDate") & ",""ACT/365"")," _
                    & dic_Addresses("LowerBar") & "," & dic_Addresses("UpperBar") & "," & dic_Addresses("Knock_Bar") _
                    & "=""Out"")*" & dic_Addresses("OptionDF") & "*" & dic_Addresses("SpotDF") & "*cyGetFXDiscSpot(" _
                    & dic_Addresses("DomCCY") & "," & dic_Addresses("PnLCCY") & ")"
            Case DetailedCat.DoubleAmericanQ
                bln_Next_Display = True
                str_Next_Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" & dic_Addresses("Notional") & _
                    "*Calc_BSPrice_DoubleAmericanBar(IF(" & dic_Addresses("Payout") & "=""Call"",1,-1)," & dic_Addresses("Spot") & "," _
                    & dic_Addresses("Fwd_DriftAdj") & "," & dic_Addresses("Strike") & "," & dic_Addresses("Vol") & ",Calc_YearFrac(" _
                    & dic_Addresses("ValDate") & "," & dic_Addresses("MatDate") & ",""ACT/365"")," & dic_Addresses("LowerBar") & "," _
                    & dic_Addresses("UpperBar") & "," & dic_Addresses("Knock_Bar") & "=""Out"")*" & dic_Addresses("QuantoFactor") & "*" _
                    & dic_Addresses("OptionDF") & "*" & dic_Addresses("SpotDF") & "*cyGetFXDiscSpot(" & dic_Addresses("PayoutCCY") & "," _
                    & dic_Addresses("PnLCCY") & ")"
            Case DetailedCat.Vanilla
                bln_Next_Display = True
                str_Next_Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" _
                    & dic_Addresses("Notional") & "*Calc_BSPrice_Vanilla(IF(" & dic_Addresses("Payout") & "=""Call"",1,-1)," _
                    & dic_Addresses("Fwd") & "," & dic_Addresses("Strike") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," _
                    & dic_Addresses("MatDate") & ",""ACT/365"")," & dic_Addresses("Vol") & ")*" & dic_Addresses("OptionDF") _
                    & "*" & dic_Addresses("SpotDF") & "*cyGetFXDiscSpot(" & dic_Addresses("DomCCY") & "," & dic_Addresses("PnLCCY") & ")"
            Case DetailedCat.Quanto
                bln_Next_Display = True
                str_Next_Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" & dic_Addresses("Notional") & "*Calc_BSPrice_Vanilla(IF(" _
                    & dic_Addresses("Payout") & "=""Call"",1,-1)," & dic_Addresses("Fwd_DriftAdj") & "," & dic_Addresses("Strike") _
                    & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," & dic_Addresses("MatDate") & ",""ACT/365"")," & dic_Addresses("Vol") _
                    & ")*" & dic_Addresses("QuantoFactor") & "*" & dic_Addresses("OptionDF") & "*" & dic_Addresses("SpotDF") & "*cyGetFXDiscSpot(" _
                    & dic_Addresses("PayoutCCY") & "," & dic_Addresses("PnLCCY") & ")"
            Case DetailedCat.CashFlow
                bln_Next_Display = True
                str_Next_Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" _
                    & dic_Addresses("Notional") & "*MAX(IF(" & dic_Addresses("Payout") & "=""Call"",1,-1)*(" _
                    & dic_Addresses("Spot") & "-" & dic_Addresses("Strike") & "),0)*" & dic_Addresses("SpotDF") _
                    & "*cyGetFXDiscSpot(" & dic_Addresses("DomCCY") & "," & dic_Addresses("PnLCCY") & ")"

            Case DetailedCat.EuropeanBar
                bln_Next_Display = True
                If (fld_Params.OptDirection = PutOpt And enu_Type_Barrier = UpperBar And fld_Params.IsKnockOut = False And fld_Params.OptDirection * fld_Params.strike < fld_Params.OptDirection * dbl_SingleBarrier) _
                Or (fld_Params.OptDirection = PutOpt And enu_Type_Barrier = LowerBar And fld_Params.IsKnockOut = True And fld_Params.OptDirection * fld_Params.strike < fld_Params.OptDirection * dbl_SingleBarrier) _
                Or (fld_Params.OptDirection = CallOpt And enu_Type_Barrier = LowerBar And fld_Params.IsKnockOut = False And fld_Params.OptDirection * fld_Params.strike < fld_Params.OptDirection * dbl_SingleBarrier) _
                Or (fld_Params.OptDirection = CallOpt And enu_Type_Barrier = UpperBar And fld_Params.IsKnockOut = True And fld_Params.OptDirection * fld_Params.strike < fld_Params.OptDirection * dbl_SingleBarrier) Then
                    str_Next_Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" _
                        & dic_Addresses("Notional") & "*(Calc_BSPrice_Vanilla(IF(" & dic_Addresses("Payout") & "=""Call"",1,-1)," _
                        & dic_Addresses("Fwd") & "," & dic_Addresses("Strike") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," _
                        & dic_Addresses("MatDate") & ",""ACT/365"")," & dic_Addresses("Vol") & ")-Calc_BSPrice_Vanilla(IF(" & dic_Addresses("Payout") & "=""Call"",1,-1)," _
                        & dic_Addresses("Fwd") & "," & dic_Addresses("Barrier") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," _
                        & dic_Addresses("MatDate") & ",""ACT/365"")," & dic_Addresses("Vol_Barrier") & ")-IF(" & dic_Addresses("Payout") & "=""Call"",1,-1)*(" & dic_Addresses("Barrier") _
                        & "-" & dic_Addresses("Strike") & ")*(Calc_BSPrice_Vanilla(IF(" & dic_Addresses("Payout") & "=""Call"",1,-1)," _
                        & dic_Addresses("Fwd") & "," & dic_Addresses("Barrier") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," _
                        & dic_Addresses("MatDate") & ",""ACT/365"")," & dic_Addresses("Vol_Barrier") & ")-Calc_BSPrice_Vanilla(IF(" & dic_Addresses("Payout") & "=""Call"",1,-1)," _
                        & dic_Addresses("Fwd") & "," & dic_Addresses("Shifted_Barrier") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," _
                        & dic_Addresses("MatDate") & ",""ACT/365"")," & dic_Addresses("Vol_ShiftedBarrier") & "))/(" & dic_Addresses("Barrier") & "*0.0001))*" & dic_Addresses("OptionDF") _
                        & "*" & dic_Addresses("SpotDF") & "*cyGetFXDiscSpot(" & dic_Addresses("DomCCY") & "," & dic_Addresses("PnLCCY") & ")"

                ElseIf (fld_Params.OptDirection = PutOpt And enu_Type_Barrier = UpperBar And fld_Params.IsKnockOut = False And fld_Params.OptDirection * fld_Params.strike >= fld_Params.OptDirection * dbl_SingleBarrier) _
                Or (fld_Params.OptDirection = PutOpt And enu_Type_Barrier = LowerBar And fld_Params.IsKnockOut = True And fld_Params.OptDirection * fld_Params.strike >= fld_Params.OptDirection * dbl_SingleBarrier) _
                Or (fld_Params.OptDirection = CallOpt And enu_Type_Barrier = LowerBar And fld_Params.IsKnockOut = False And fld_Params.OptDirection * fld_Params.strike >= fld_Params.OptDirection * dbl_SingleBarrier) _
                Or (fld_Params.OptDirection = CallOpt And enu_Type_Barrier = UpperBar And fld_Params.IsKnockOut = True And fld_Params.OptDirection * fld_Params.strike >= fld_Params.OptDirection * dbl_SingleBarrier) Then
                    str_Next_Formula = "0"

                ElseIf (fld_Params.OptDirection = PutOpt And enu_Type_Barrier = UpperBar And fld_Params.IsKnockOut = True And fld_Params.OptDirection * fld_Params.strike < fld_Params.OptDirection * dbl_SingleBarrier) _
                Or (fld_Params.OptDirection = PutOpt And enu_Type_Barrier = LowerBar And fld_Params.IsKnockOut = False And fld_Params.OptDirection * fld_Params.strike < fld_Params.OptDirection * dbl_SingleBarrier) _
                Or (fld_Params.OptDirection = CallOpt And enu_Type_Barrier = LowerBar And fld_Params.IsKnockOut = True And fld_Params.OptDirection * fld_Params.strike < fld_Params.OptDirection * dbl_SingleBarrier) _
                Or (fld_Params.OptDirection = CallOpt And enu_Type_Barrier = UpperBar And fld_Params.IsKnockOut = False And fld_Params.OptDirection * fld_Params.strike < fld_Params.OptDirection * dbl_SingleBarrier) Then
                    str_Next_Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" _
                        & dic_Addresses("Notional") & "*(Calc_BSPrice_Vanilla(IF(" & dic_Addresses("Payout") & "=""Call"",1,-1)," _
                        & dic_Addresses("Fwd") & "," & dic_Addresses("Barrier") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," _
                        & dic_Addresses("MatDate") & ",""ACT/365"")," & dic_Addresses("Vol_Barrier") & ")-IF(" & dic_Addresses("Payout") & "=""Call"",1,-1)*(" & dic_Addresses("Strike") _
                        & "-" & dic_Addresses("Barrier") & ")*(Calc_BSPrice_Vanilla(IF(" & dic_Addresses("Payout") & "=""Call"",1,-1)," _
                        & dic_Addresses("Fwd") & "," & dic_Addresses("Barrier") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," _
                        & dic_Addresses("MatDate") & ",""ACT/365"")," & dic_Addresses("Vol_Barrier") & ")-Calc_BSPrice_Vanilla(IF(" & dic_Addresses("Payout") & "=""Call"",1,-1)," _
                        & dic_Addresses("Fwd") & "," & dic_Addresses("Shifted_Barrier") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," _
                        & dic_Addresses("MatDate") & ",""ACT/365"")," & dic_Addresses("Vol_ShiftedBarrier") & "))/(" & dic_Addresses("Barrier") & "*0.0001))*" & dic_Addresses("OptionDF") _
                        & "*" & dic_Addresses("SpotDF") & "*cyGetFXDiscSpot(" & dic_Addresses("DomCCY") & "," & dic_Addresses("PnLCCY") & ")"

                 ElseIf (fld_Params.OptDirection = PutOpt And enu_Type_Barrier = UpperBar And fld_Params.IsKnockOut = True And fld_Params.OptDirection * fld_Params.strike >= fld_Params.OptDirection * dbl_SingleBarrier) _
                Or (fld_Params.OptDirection = PutOpt And enu_Type_Barrier = LowerBar And fld_Params.IsKnockOut = False And fld_Params.OptDirection * fld_Params.strike >= fld_Params.OptDirection * dbl_SingleBarrier) _
                Or (fld_Params.OptDirection = CallOpt And enu_Type_Barrier = LowerBar And fld_Params.IsKnockOut = True And fld_Params.OptDirection * fld_Params.strike >= fld_Params.OptDirection * dbl_SingleBarrier) _
                Or (fld_Params.OptDirection = CallOpt And enu_Type_Barrier = UpperBar And fld_Params.IsKnockOut = False And fld_Params.OptDirection * fld_Params.strike >= fld_Params.OptDirection * dbl_SingleBarrier) Then
                    str_Next_Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" _
                    & dic_Addresses("Notional") & "*Calc_BSPrice_Vanilla(IF(" & dic_Addresses("Payout") & "=""Call"",1,-1)," _
                    & dic_Addresses("Fwd") & "," & dic_Addresses("Strike") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," _
                    & dic_Addresses("MatDate") & ",""ACT/365"")," & dic_Addresses("Vol") & ")*" & dic_Addresses("OptionDF") _
                    & "*" & dic_Addresses("SpotDF") & "*cyGetFXDiscSpot(" & dic_Addresses("DomCCY") & "," & dic_Addresses("PnLCCY") & ")"
                End If

            Case Else
                bln_Next_Display = False
                str_Next_Formula = ""
        End Select

        If bln_Next_Display = True Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Option MV (" & fld_Params.CCY_PnL & "):"
            .Offset(int_ActiveRow, 1).Formula = str_Next_Formula
            .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
            .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
            Call dic_Addresses.Add("MV_Option", .Offset(int_ActiveRow, 1).Address(False, False))
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
            Case DetailedCat.None
                dic_Addresses("Range_PnL").Formula = "=" & dic_Addresses("Cash")
            Case Else
                dic_Addresses("Range_PnL").Formula = "=" & dic_Addresses("MV_Option") & "+" & dic_Addresses("Cash")
        End Select
    End With

    wks_output.Calculate
    wks_output.Cells.HorizontalAlignment = xlCenter
    wks_output.Columns.AutoFit
End Sub
