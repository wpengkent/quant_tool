Option Explicit

Private Enum DetailedCat
    None = 0
    Eur_Standard = 100
    Eur_Digital = 101
    Eur_DigitalCS = 102
    Eur_Quanto = 110
    Eur_DigitalQuanto = 111
    Amr_Standard = 200
    Amr_Digital = 201
    Amr_Quanto = 210
    Amr_DigitalQuanto = 211
End Enum


' ## MEMBER DATA
' Components
Private scf_Premium As SCF

' Curve dependencies
Private dic_FXCurveNames As Dictionary
Private fxs_Spots As Data_FXSpots, irc_DiscCurve As Data_IRCurve, irc_SpotDiscCurve As Data_IRCurve, irc_Fgn As Data_IRCurve
Private fxv_Vols_XY As Data_FXVols, fxv_Vols_XQ As Data_FXVols, fxv_Vols_YQ As Data_FXVols

' Dynamic variables
Private lng_ValDate As Long, lng_SpotDate As Long
Private dbl_TimeToMat As Double, dbl_TimeEstPeriod As Double

' Static values
Private dic_GlobalStaticInfo As Dictionary, dic_CurveDependencies As Dictionary, map_Rules As MappingRules
Private fld_Params As InstParams_FVN
Private int_Sign As Integer, dbl_Strike As Double
Private lng_MatDate As Long, lng_MatSpotDate As Long, lng_MatSpotDate_Std As Long
Private bln_IsQuanto As Boolean, enu_Payoff As EuropeanPayoff
Private str_Pair_XY As String, str_Pair_XQ As String, str_Pair_YQ As String


' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As InstParams_FVN, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing)

    ' Store static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    Set map_Rules = dic_GlobalStaticInfo(StaticInfoType.MappingRules)
    fld_Params = fld_ParamsInput

    If fld_Params.IsDigital = True Then
        If (fld_Params.CCY_Payout = fld_Params.CCY_Fgn) Then
            enu_Payoff = EuropeanPayoff.Digital_AoN
        Else
            enu_Payoff = EuropeanPayoff.Digital_CoN  ' Includes quantos
        End If
    Else
        enu_Payoff = EuropeanPayoff.Standard
    End If

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
        Set fxs_Spots = GetObject_FXSpots(True)
        Set irc_DiscCurve = GetObject_IRCurve(fld_Params.Curve_Disc, True, False)
        Set irc_SpotDiscCurve = GetObject_IRCurve(fld_Params.Curve_SpotDisc, True, False)

        Set fxv_Vols_XY = GetObject_FXVols(str_Pair_XY, True, False)
        If bln_IsQuanto = True Then
            Set fxv_Vols_XQ = GetObject_FXVols(str_Pair_XQ, True, False)
            Set fxv_Vols_YQ = GetObject_FXVols(str_Pair_YQ, True, False)
        End If
    Else
        Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
        Set irc_DiscCurve = dic_CurveSet(CurveType.IRC)(fld_Params.Curve_Disc)
        Set irc_SpotDiscCurve = dic_CurveSet(CurveType.IRC)(fld_Params.Curve_SpotDisc)

        Set fxv_Vols_XY = dic_CurveSet(CurveType.FXV)(str_Pair_XY)
        If bln_IsQuanto = True Then
            Set fxv_Vols_XQ = dic_CurveSet(CurveType.FXV)(str_Pair_XQ)
            Set fxv_Vols_YQ = dic_CurveSet(CurveType.FXV)(str_Pair_YQ)
        End If
    End If

    ' Static values
    If fld_Params.IsBuy = True Then int_Sign = 1 Else int_Sign = -1
    dbl_Strike = fld_Params.strike
    lng_MatDate = fld_Params.MatDate

    '------------------------------------------------
    ' Delivery date, T+2 for standard type
    '------------------------------------------------
    Select Case fld_Params.LateType
        Case "STANDARD"
            lng_MatSpotDate = cyGetFXCrossSpotDate(fld_Params.CCY_Fgn, fld_Params.CCY_Dom, lng_MatDate, dic_GlobalStaticInfo)
        Case "LATE DELIVERY ATM SPOT"
            lng_MatSpotDate_Std = cyGetFXCrossSpotDate(fld_Params.CCY_Fgn, fld_Params.CCY_Dom, lng_MatDate, dic_GlobalStaticInfo)
            lng_MatSpotDate = fld_Params.DelivDate
        Case Else
            lng_MatSpotDate = fld_Params.DelivDate
    End Select

    ' Initialize dynamic variables
    Call Me.SetValDate(fld_Params.ValueDate)

    ' Determine curve dependencies
    Set dic_CurveDependencies = scf_Premium.CurveDependencies
    Set dic_CurveDependencies = Convert_MergeDicts(dic_CurveDependencies, fxs_Spots.Lookup_CurveDependencies(fld_Params.CCY_Dom, _
        fld_Params.CCY_Fgn, fld_Params.CCY_Payout, fld_Params.CCY_PnL))
    If dic_CurveDependencies.Exists(irc_DiscCurve.CurveName) = False Then Call dic_CurveDependencies.Add(irc_DiscCurve.CurveName, True)
    If dic_CurveDependencies.Exists(irc_SpotDiscCurve.CurveName) = False Then Call dic_CurveDependencies.Add(irc_SpotDiscCurve.CurveName, True)

    Set dic_FXCurveNames = map_Rules.Dict_FXCurveNames
    Dim str_Curve_Fgn As String: str_Curve_Fgn = dic_FXCurveNames(fld_Params.CCY_Fgn)
    Set irc_Fgn = GetObject_IRCurve(str_Curve_Fgn, True, False)

End Sub

'-------------------------------------------------------------------------------------------
' NAME:    marketvalue
'
' PURPOSE: Calculate market value
'
' NOTES: Call in InstrumentCache
'
' INPUT OPTIONS:
'
' MODIFIED:
'    30JAN2020 - KW - Support Late Settlement FX Vanilla Option
'                     Late Cash, Late Delivery, Late Delivery ATM Spot
'
'-------------------------------------------------------------------------------------------
' ## PROPERTIES
Public Property Get marketvalue() As Double
    ' ## Get discounted value of future flows in the PnL currency
    Dim dbl_UnitVal As Double
    Dim dbl_Spot As Double: dbl_Spot = Spot()
    Dim dbl_Fwd As Double: dbl_Fwd = Forward()
    Dim dbl_VolPct_XY As Double, dbl_ATMVol_XY As Double, dbl_ATMVol_XQ As Double, dbl_ATMVol_YQ As Double, dbl_Correl_XQ As Double
    Dim dbl_DF_ValSpot As Double: dbl_DF_ValSpot = irc_SpotDiscCurve.Lookup_Rate(lng_ValDate, lng_SpotDate, "DF", , , True)
    Dim dbl_DF_MatSpot As Double: dbl_DF_MatSpot = irc_DiscCurve.Lookup_Rate(lng_SpotDate, lng_MatSpotDate, "DF", , , True)
    Dim dbl_DF_MatSpot_Std As Double: dbl_DF_MatSpot_Std = irc_DiscCurve.Lookup_Rate(lng_SpotDate, lng_MatSpotDate_Std, "DF", , , True)

    dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)

    If bln_IsQuanto = True Then
        ' Perform drift adjustment on forward
        dbl_ATMVol_XY = GetVol(VolPair.XY, -1) / 100
        dbl_ATMVol_XQ = GetVol(VolPair.XQ, -1) / 100
        dbl_ATMVol_YQ = GetVol(VolPair.YQ, -1) / 100
        dbl_Correl_XQ = Calc_QuantoCorrel(dbl_ATMVol_XY, dbl_ATMVol_XQ, dbl_ATMVol_YQ)
        dbl_Fwd = dbl_Fwd * Calc_DriftAdjFactor(dbl_ATMVol_XY, dbl_ATMVol_YQ, dbl_Correl_XQ, dbl_TimeToMat)
    End If

    ' Determine smile correction for digitals
    Dim dbl_Vega As Double, dbl_Gamma As Double, dbl_Vanna As Double, dbl_SmileSlope As Double, dbl_SmileCorrection_Digital As Double
    Dim enu_DetailedCat As DetailedCat: enu_DetailedCat = DetailedCategory()
    If enu_DetailedCat = Eur_Digital And fld_Params.IsSmile = True Then
    'dbl_Vega = Calc_BS_Vega_Vanilla(dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY) * dbl_DF_MatSpot
    'dbl_Gamma = Calc_BS_Gamma_Vanilla(dbl_Spot, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY) * dbl_DF_MatSpot * dbl_Spot
    'dbl_Vanna = Calc_BS_Vanna_Vanilla(dbl_Spot, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY) * dbl_DF_MatSpot
    'dbl_SmileSlope = fxv_Vols_XY.Lookup_SmileSlope(fld_Params.MatDate, dbl_Strike)
    'dbl_SmileCorrection_Digital = fld_Params.Direction * dbl_Vega * dbl_Gamma * dbl_SmileSlope / (1 - dbl_SmileSlope * dbl_Vanna)
        ' ## NOT READY YET
    End If

    ' Obtain value per unit at the valuation date
    Dim dbl_ShiftedVol As Double
    Select Case DetailedCategory()
        Case DetailedCat.Eur_DigitalQuanto
            ' ## Currently only matches Murex for the smile off case
            dbl_UnitVal = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, _
                enu_Payoff) * dbl_DF_MatSpot * dbl_DF_ValSpot
        Case DetailedCat.Eur_Digital
            dbl_UnitVal = BSDigital(dbl_Strike, dbl_Spot, dbl_DF_ValSpot, dbl_DF_MatSpot)
        Case DetailedCat.Eur_DigitalCS
            If fld_Params.IsSmile = True Then
                dbl_ShiftedVol = GetVol(VolPair.XY, dbl_Strike * (1 + fld_Params.CSFactor))
            Else
                dbl_ShiftedVol = dbl_VolPct_XY
            End If

            dbl_UnitVal = Calc_BSPrice_DigitalCS(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, _
                dbl_ShiftedVol, (fld_Params.CCY_Payout = fld_Params.CCY_Dom), fld_Params.CSFactor) _
                * dbl_DF_MatSpot * dbl_DF_ValSpot

        Case DetailedCat.Eur_Standard
            '------------------------------------------------
            ' Calculate market value depending on Late Type
            '------------------------------------------------
            Select Case fld_Params.LateType
                '------------------------------------------------
                ' Standard type, fix at T+2, deliver at T+2
                ' Late Cash, fix at T+2, deliver at T+x
                ' Late Delivery, fix at T+x, diliver at T+x
                ' Input are depending on late type
                '------------------------------------------------
                Case "STANDARD", "LATE CASH", "LATE DELIVERY"
                    dbl_UnitVal = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, enu_Payoff) _
                                  * dbl_DF_MatSpot * dbl_DF_ValSpot
                '----------------------------------------------------
                ' Late Delivery ATM Spot, fix at T+2, deliver at T+x
                '----------------------------------------------------
                Case "LATE DELIVERY ATM SPOT"
                    dbl_UnitVal = Late_Del_Atm_Spot(dbl_Strike, dbl_Fwd, dbl_Spot, dbl_VolPct_XY, dbl_DF_ValSpot, dbl_DF_MatSpot, dbl_DF_MatSpot_Std, enu_Payoff)
            End Select

        Case DetailedCat.Eur_Quanto
            dbl_UnitVal = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, enu_Payoff) _
                * dbl_DF_MatSpot * dbl_DF_ValSpot
        Case DetailedCat.Amr_Standard
            dbl_UnitVal = Calc_BAW_American(fld_Params.Direction, dbl_Spot, dbl_Fwd, dbl_Strike, dbl_VolPct_XY, dbl_DF_MatSpot, _
                dbl_TimeToMat, dbl_TimeEstPeriod) * dbl_DF_ValSpot
        Case Else: dbl_UnitVal = 0
    End Select

    ' Obtain the value of all units in the PnL currency
    Dim dbl_DiscFXSpot As Double ', bln_Wa
    If bln_IsQuanto = True Then
        dbl_DiscFXSpot = fxs_Spots.Lookup_DiscSpot(fld_Params.CCY_Payout, fld_Params.CCY_PnL)
    Else
        dbl_DiscFXSpot = fxs_Spots.Lookup_DiscSpot(fld_Params.CCY_Dom, fld_Params.CCY_PnL)
    End If

    marketvalue = fld_Params.Notional_Fgn * dbl_UnitVal * int_Sign * dbl_DiscFXSpot * fld_Params.QuantoFactor
End Property

Public Property Get Cash() As Double
    ' ## Get discounted value of the premium in the PnL currency
    Cash = -scf_Premium.CalcValue(lng_ValDate, lng_SpotDate, fld_Params.CCY_PnL) * int_Sign
End Property

Public Property Get PnL() As Double
    PnL = Me.marketvalue + Me.Cash
End Property

Private Property Get Spot() As Double
    Spot = fxs_Spots.Lookup_Spot(fld_Params.CCY_Fgn, fld_Params.CCY_Dom)
End Property

Private Property Get Forward() As Double
    '------------------------------------------------------------
    ' Late Delivery Type: Forward is fixed at late delivery date
    ' Other types: fixed at T+2
    '------------------------------------------------------------
    Select Case fld_Params.LateType
        Case "LATE DELIVERY"
            Forward = fxs_Spots.Lookup_Fwd(fld_Params.CCY_Fgn, fld_Params.CCY_Dom, fld_Params.DelivDate, False)
        Case Else
            Forward = fxs_Spots.Lookup_Fwd(fld_Params.CCY_Fgn, fld_Params.CCY_Dom, lng_MatDate)
    End Select

End Property

Private Property Get DetailedCategory() As DetailedCat
    ' ## Determine option category for valuation and output purposes
    Dim enu_Output As DetailedCat: enu_Output = DetailedCat.None

    ' Determine code of output
    Dim int_Exercise As Integer
    Select Case fld_Params.ExerciseType
        Case "EUROPEAN": int_Exercise = 100
        Case "AMERICAN": int_Exercise = 200
        Case Else: int_Exercise = 0
    End Select
    Dim int_Payout As Integer: If bln_IsQuanto = True Then int_Payout = 10 Else int_Payout = 0

    Dim int_Payoff As Integer
    If fld_Params.IsDigital = True And fld_Params.CSFactor <> 0 Then
        int_Payoff = 2
    ElseIf fld_Params.IsDigital Then
        int_Payoff = 1
    Else
        int_Payoff = 0
    End If

    DetailedCategory = int_Exercise + int_Payout + int_Payoff
End Property


' ## METHODS - GREEKS
Public Function Calc_DV01(str_curve As String, Optional int_PillarIndex As Integer = 0) As Double
    ' ## Return DV01 sensitivity to the specified curve
    Dim dbl_Output As Double
    If dic_CurveDependencies.Exists(str_curve) Then
        ' Remember original setting, then disable DV01 impact on discounted spot
        Dim bln_ZShiftsEnabled_DiscSpot As Boolean: bln_ZShiftsEnabled_DiscSpot = fxs_Spots.ZShiftsEnabled_DiscSpot
        fxs_Spots.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

        ' Store shifted values
        Dim dbl_Val_Up As Double, dbl_Val_Down As Double
        Call SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        dbl_Val_Up = Me.marketvalue

        Call SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        dbl_Val_Down = Me.marketvalue

        ' Clear temporary shifts
        Call SetCurveState(str_curve, CurveState_IRC.Final)

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

        ' Store shifted values
        Dim dbl_Val_Up As Double, dbl_Val_Down As Double, dbl_Val_Unch As Double
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

'-------------------------------------------------------------------------------------------
' NAME:    Calc_Delta
'
' PURPOSE: Calculate delta in booking screen
'
' NOTES: MUREX delta is as of spot date
'
' INPUT OPTIONS:
'
' MODIFIED:
'    03DEC2019 - KW - Remove dbl_DF_ValSpot from dbl_UnitValSpotShockUp and dbl_UnitValSpotShockDown
'                     to reconcile MUREX spot delta
'                   - Project digital option shock up/down price to spot date
'    03FEB2020 - KW - Add Late Delivery Delta
'
'-------------------------------------------------------------------------------------------
Public Function Calc_Delta() As Double

    ' Added by Dennis Foong on 18th January 2016
    Dim dbl_Spot As Double
    Dim dbl_Fwd As Double
    Dim dbl_Strike As Double: dbl_Strike = fld_Params.strike
    Dim dbl_VolPct_XY As Double, dbl_ATMVol_XY As Double, dbl_ATMVol_XQ As Double, dbl_ATMVol_YQ As Double, dbl_Correl_XQ As Double
    Dim dbl_DF_ValSpot As Double: dbl_DF_ValSpot = irc_SpotDiscCurve.Lookup_Rate(lng_ValDate, lng_SpotDate, "DF", , , True)
    Dim dbl_DF_MatSpot As Double: dbl_DF_MatSpot = irc_DiscCurve.Lookup_Rate(lng_SpotDate, lng_MatSpotDate, "DF", , , True)
    Dim dbl_DF_MatSpot_Std As Double: dbl_DF_MatSpot_Std = irc_DiscCurve.Lookup_Rate(lng_SpotDate, lng_MatSpotDate_Std, "DF", , , True)

    Dim dbl_UnitValSpotShockUp As Double, dbl_UnitValSpotShockDown As Double, dbl_Output As Double
    Dim dbl_ShockSize As Double: dbl_ShockSize = 0.01

    dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
    If bln_IsQuanto = True Then
        ' Perform drift adjustment on forward
        dbl_ATMVol_XY = GetVol(VolPair.XY, -1) / 100
        dbl_ATMVol_XQ = GetVol(VolPair.XQ, -1) / 100
        dbl_ATMVol_YQ = GetVol(VolPair.YQ, -1) / 100
        dbl_Correl_XQ = Calc_QuantoCorrel(dbl_ATMVol_XY, dbl_ATMVol_XQ, dbl_ATMVol_YQ)
        dbl_Fwd = dbl_Fwd * Calc_DriftAdjFactor(dbl_ATMVol_XY, dbl_ATMVol_YQ, dbl_Correl_XQ, dbl_TimeToMat)
    End If

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

    Select Case DetailedCategory()

        Case DetailedCat.Eur_DigitalQuanto

            ' ## Currently only matches Murex for the smile off case

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_Fwd = Forward()
            dbl_UnitValSpotShockUp = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, _
                enu_Payoff) * dbl_DF_MatSpot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_Fwd = Forward()
            dbl_UnitValSpotShockDown = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, _
                enu_Payoff) * dbl_DF_MatSpot

        Case DetailedCat.Eur_Digital

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            '------------------------------------------------
            ' Project price from valuation date to spot date
            '------------------------------------------------
            dbl_UnitValSpotShockUp = BSDigital(dbl_Strike, dbl_Spot, dbl_DF_ValSpot, dbl_DF_MatSpot) / dbl_DF_ValSpot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Spot = Spot()
            '------------------------------------------------
            ' Project price from valuation date to spot date
            '------------------------------------------------
            dbl_UnitValSpotShockDown = BSDigital(dbl_Strike, dbl_Spot, dbl_DF_ValSpot, dbl_DF_MatSpot) / dbl_DF_ValSpot

        Case DetailedCat.Eur_DigitalCS

            Dim dbl_ShiftedVol As Double
            If fld_Params.IsSmile = True Then
                dbl_ShiftedVol = GetVol(VolPair.XY, dbl_Strike * (1 + fld_Params.CSFactor))
            Else
                dbl_ShiftedVol = dbl_VolPct_XY
            End If

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_Fwd = Forward()
            dbl_UnitValSpotShockUp = Calc_BSPrice_DigitalCS(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, _
                dbl_ShiftedVol, (fld_Params.CCY_Payout = fld_Params.CCY_Dom), fld_Params.CSFactor) _
                * dbl_DF_MatSpot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_Fwd = Forward()
            dbl_UnitValSpotShockDown = Calc_BSPrice_DigitalCS(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, _
                dbl_ShiftedVol, (fld_Params.CCY_Payout = fld_Params.CCY_Dom), fld_Params.CSFactor) _
                * dbl_DF_MatSpot

        Case DetailedCat.Eur_Standard

            Select Case fld_Params.LateType
                Case "LATE DELIVERY ATM SPOT"
                    ' Shock up
                    Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
                    Call fxs_Spots.Scen_ApplyCurrent
                    dbl_Fwd = Forward()
                    dbl_Spot = Spot()
                    dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
                    '------------------------------------------------
                    ' Project price from valuation date to spot date
                    '------------------------------------------------
                    dbl_UnitValSpotShockUp = Late_Del_Atm_Spot(dbl_Strike, dbl_Fwd, dbl_Spot, dbl_VolPct_XY, dbl_DF_ValSpot, dbl_DF_MatSpot, dbl_DF_MatSpot_Std, enu_Payoff) / dbl_DF_ValSpot

                    ' Shock down
                    Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
                    Call fxs_Spots.Scen_ApplyCurrent
                    dbl_Fwd = Forward()
                    dbl_Spot = Spot()
                    dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
                    '------------------------------------------------
                    ' Project price from valuation date to spot date
                    '------------------------------------------------
                    dbl_UnitValSpotShockDown = Late_Del_Atm_Spot(dbl_Strike, dbl_Fwd, dbl_Spot, dbl_VolPct_XY, dbl_DF_ValSpot, dbl_DF_MatSpot, dbl_DF_MatSpot_Std, enu_Payoff) / dbl_DF_ValSpot

                Case Else
                    ' Shock up
                    Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
                    Call fxs_Spots.Scen_ApplyCurrent
                    dbl_Fwd = Forward()
                    dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
                    dbl_UnitValSpotShockUp = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, enu_Payoff) _
                                             * dbl_DF_MatSpot

                    ' Shock down
                    Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
                    Call fxs_Spots.Scen_ApplyCurrent
                    dbl_Fwd = Forward()
                    dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
                    dbl_UnitValSpotShockDown = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, enu_Payoff) _
                                               * dbl_DF_MatSpot
            End Select

        Case DetailedCat.Eur_Quanto

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_UnitValSpotShockUp = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, enu_Payoff) _
                * dbl_DF_MatSpot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_UnitValSpotShockDown = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, enu_Payoff) _
                * dbl_DF_MatSpot

        Case DetailedCat.Amr_Standard

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_UnitValSpotShockUp = Calc_BAW_American(fld_Params.Direction, dbl_Spot, dbl_Fwd, dbl_Strike, dbl_VolPct_XY, dbl_DF_MatSpot, _
                dbl_TimeToMat, dbl_TimeEstPeriod)

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_Spot = Spot()
            dbl_Fwd = Forward()
            dbl_UnitValSpotShockDown = Calc_BAW_American(fld_Params.Direction, dbl_Spot, dbl_Fwd, dbl_Strike, dbl_VolPct_XY, dbl_DF_MatSpot, _
                dbl_TimeToMat, dbl_TimeEstPeriod)

        Case Else:
            dbl_UnitValSpotShockUp = 0
            dbl_UnitValSpotShockDown = 0

    End Select

    Call fxs_Spots.Scen_ApplyBase
    dbl_Spot = Spot()
    'dbl_Output = int_Sign * fld_Params.Direction * fld_Params.Notional_Fgn * (dbl_UnitValSpotShockUp - dbl_UnitValSpotShockDown) / (2 * dbl_ShockSize / 100 * dbl_Spot)
    dbl_Output = int_Sign * fld_Params.Notional_Fgn * (dbl_UnitValSpotShockUp - dbl_UnitValSpotShockDown) / (2 * dbl_ShockSize / 100 * dbl_Spot)
    If fld_Params.Premium.CCY = fld_Params.CCY_Fgn Then
        If int_Sign = 1 Then
            dbl_Output = dbl_Output - fld_Params.Premium.Amount
        Else
            dbl_Output = dbl_Output + fld_Params.Premium.Amount
        End If
    End If
    Calc_Delta = dbl_Output

End Function

'-------------------------------------------------------------------------------------------
' NAME:    Calc_Gamma
'
' PURPOSE: Calculate gamma in booking screen
'
' NOTES: MUREX gamma is as of today date
'
' INPUT OPTIONS:
'
' MODIFIED:
'   03FEB2020 - KW - Add Late Delivery Gamma
'
'-------------------------------------------------------------------------------------------
Public Function Calc_Gamma() As Double

    ' Added by Dennis Foong on 18th January 2016
    Dim dbl_Spot As Double: dbl_Spot = Spot()
    Dim dbl_Fwd As Double: dbl_Fwd = Forward()
    Dim dbl_Strike As Double: dbl_Strike = fld_Params.strike
    Dim dbl_VolPct_XY As Double, dbl_ATMVol_XY As Double, dbl_ATMVol_XQ As Double, dbl_ATMVol_YQ As Double, dbl_Correl_XQ As Double
    Dim dbl_DF_ValSpot As Double: dbl_DF_ValSpot = irc_SpotDiscCurve.Lookup_Rate(lng_ValDate, lng_SpotDate, "DF", , , True)
    Dim dbl_DF_MatSpot As Double: dbl_DF_MatSpot = irc_DiscCurve.Lookup_Rate(lng_SpotDate, lng_MatSpotDate, "DF", , , True)
    Dim dbl_DF_MatSpot_Std As Double: dbl_DF_MatSpot_Std = irc_DiscCurve.Lookup_Rate(lng_SpotDate, lng_MatSpotDate_Std, "DF", , , True)

    Dim dbl_UnitValBase As Double, dbl_UnitValSpotShockUp As Double, dbl_UnitValSpotShockDown As Double
    Dim dbl_SpotDelta As Double, dbl_Output As Double
    Dim dbl_ShockSize As Double: dbl_ShockSize = 0.01

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

    Select Case DetailedCategory()
        Case DetailedCat.Eur_DigitalQuanto
            ' ## Currently only matches Murex for the smile off case

            ' Base
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_Fwd = Forward()
            dbl_UnitValBase = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, _
                enu_Payoff) * dbl_DF_MatSpot * dbl_DF_ValSpot

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_Fwd = Forward()
            dbl_UnitValSpotShockUp = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, _
                enu_Payoff) * dbl_DF_MatSpot * dbl_DF_ValSpot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_Fwd = Forward()
            dbl_UnitValSpotShockDown = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, _
                enu_Payoff) * dbl_DF_MatSpot * dbl_DF_ValSpot

        Case DetailedCat.Eur_Digital

            ' Base
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_Spot = Spot()

            dbl_UnitValBase = BSDigital(dbl_Strike, dbl_Spot, dbl_DF_ValSpot, dbl_DF_MatSpot)

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_Spot = Spot()

            dbl_UnitValSpotShockUp = BSDigital(dbl_Strike, dbl_Spot, dbl_DF_ValSpot, dbl_DF_MatSpot)

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_Spot = Spot()

            dbl_UnitValSpotShockDown = BSDigital(dbl_Strike, dbl_Spot, dbl_DF_ValSpot, dbl_DF_MatSpot)

        Case DetailedCat.Eur_DigitalCS

            Dim dbl_ShiftedVol As Double
            If fld_Params.IsSmile = True Then
                dbl_ShiftedVol = GetVol(VolPair.XY, dbl_Strike * (1 + fld_Params.CSFactor))
            Else
                dbl_ShiftedVol = dbl_VolPct_XY
            End If

            ' Base
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_Fwd = Forward()
            dbl_UnitValBase = Calc_BSPrice_DigitalCS(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, _
                dbl_ShiftedVol, (fld_Params.CCY_Payout = fld_Params.CCY_Dom), fld_Params.CSFactor) _
                * dbl_DF_MatSpot * dbl_DF_ValSpot

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_Fwd = Forward()
            dbl_UnitValSpotShockUp = Calc_BSPrice_DigitalCS(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, _
                dbl_ShiftedVol, (fld_Params.CCY_Payout = fld_Params.CCY_Dom), fld_Params.CSFactor) _
                * dbl_DF_MatSpot * dbl_DF_ValSpot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_Fwd = Forward()
            dbl_UnitValSpotShockDown = Calc_BSPrice_DigitalCS(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, _
                dbl_ShiftedVol, (fld_Params.CCY_Payout = fld_Params.CCY_Dom), fld_Params.CSFactor) _
                * dbl_DF_MatSpot * dbl_DF_ValSpot

        Case DetailedCat.Eur_Standard
            Select Case fld_Params.LateType
                Case "LATE DELIVERY ATM SPOT"
                    ' Base
                    dbl_Fwd = Forward()
                    dbl_Spot = Spot()
                    dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
                    dbl_UnitValBase = Late_Del_Atm_Spot(dbl_Strike, dbl_Fwd, dbl_Spot, dbl_VolPct_XY, dbl_DF_ValSpot, dbl_DF_MatSpot, dbl_DF_MatSpot_Std, enu_Payoff)

                    ' Shock up
                    Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
                    Call fxs_Spots.Scen_ApplyCurrent
                    dbl_Fwd = Forward()
                    dbl_Spot = Spot()
                    dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
                    dbl_UnitValSpotShockUp = Late_Del_Atm_Spot(dbl_Strike, dbl_Fwd, dbl_Spot, dbl_VolPct_XY, dbl_DF_ValSpot, dbl_DF_MatSpot, dbl_DF_MatSpot_Std, enu_Payoff)

                    ' Shock down
                    Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
                    Call fxs_Spots.Scen_ApplyCurrent
                    dbl_Fwd = Forward()
                    dbl_Spot = Spot()
                    dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
                    dbl_UnitValSpotShockDown = Late_Del_Atm_Spot(dbl_Strike, dbl_Fwd, dbl_Spot, dbl_VolPct_XY, dbl_DF_ValSpot, dbl_DF_MatSpot, dbl_DF_MatSpot_Std, enu_Payoff)

                Case Else
                    ' Base
                    dbl_Fwd = Forward()
                    dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
                    dbl_UnitValBase = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, enu_Payoff) _
                                      * dbl_DF_MatSpot * dbl_DF_ValSpot

                    ' Shock up
                    Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
                    Call fxs_Spots.Scen_ApplyCurrent
                    dbl_Fwd = Forward()
                    dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
                    dbl_UnitValSpotShockUp = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, enu_Payoff) _
                                             * dbl_DF_MatSpot * dbl_DF_ValSpot

                    ' Shock down
                    Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
                    Call fxs_Spots.Scen_ApplyCurrent
                    dbl_Fwd = Forward()
                    dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
                    dbl_UnitValSpotShockDown = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, enu_Payoff) _
                                               * dbl_DF_MatSpot * dbl_DF_ValSpot
            End Select

        Case DetailedCat.Eur_Quanto

            ' Base
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_UnitValBase = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, enu_Payoff) _
                * dbl_DF_MatSpot * dbl_DF_ValSpot

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_UnitValSpotShockUp = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, enu_Payoff) _
                * dbl_DF_MatSpot * dbl_DF_ValSpot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_Fwd = Forward()
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_UnitValSpotShockDown = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, enu_Payoff) _
                * dbl_DF_MatSpot * dbl_DF_ValSpot
        Case DetailedCat.Amr_Standard

            ' Base
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_Spot = Spot()
            dbl_UnitValBase = Calc_BAW_American(fld_Params.Direction, dbl_Spot, dbl_Fwd, dbl_Strike, dbl_VolPct_XY, dbl_DF_MatSpot, _
                dbl_TimeToMat, dbl_TimeEstPeriod) * dbl_DF_ValSpot

            ' Shock up
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_Spot = Spot()
            dbl_UnitValSpotShockUp = Calc_BAW_American(fld_Params.Direction, dbl_Spot, dbl_Fwd, dbl_Strike, dbl_VolPct_XY, dbl_DF_MatSpot, _
                dbl_TimeToMat, dbl_TimeEstPeriod) * dbl_DF_ValSpot

            ' Shock down
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent
            dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
            dbl_Spot = Spot()
            dbl_UnitValSpotShockDown = Calc_BAW_American(fld_Params.Direction, dbl_Spot, dbl_Fwd, dbl_Strike, dbl_VolPct_XY, dbl_DF_MatSpot, _
                dbl_TimeToMat, dbl_TimeEstPeriod) * dbl_DF_ValSpot

        Case Else:

            dbl_UnitValBase = 0
            dbl_UnitValSpotShockUp = 0
            dbl_UnitValSpotShockDown = 0

    End Select

    Call fxs_Spots.Scen_ApplyBase
    dbl_Spot = Spot()
    'dbl_Output = int_Sign * fld_Params.Direction * fld_Params.Notional_Fgn * (dbl_UnitValSpotShockUp + dbl_UnitValSpotShockDown - 2 * dbl_UnitValBase) / ((dbl_ShockSize / 100 * dbl_Spot) ^ 2) * dbl_Spot / 100
    dbl_Output = int_Sign * fld_Params.Notional_Fgn * (dbl_UnitValSpotShockUp + dbl_UnitValSpotShockDown - 2 * dbl_UnitValBase) / ((dbl_ShockSize / 100 * dbl_Spot) ^ 2) * dbl_Spot / 100
    Calc_Gamma = dbl_Output

End Function

Public Function Calc_Vega(enu_Type As CurveType, str_curve As String) As Double
    ' ## Return sensitivity to a 1 vol increase in the vol
    Dim dbl_Output As Double: dbl_Output = 0
    If fxv_Vols_XY.TypeCode = enu_Type Then
        Dim bln_IsShifted As Boolean
        Dim dbl_Val_Up As Double, dbl_Val_Down As Double

        ' Store shifted values and gather values
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
    End If

    Calc_Vega = dbl_Output
End Function


' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Sub SetValDate(lng_Input As Long)
    ' Set stored value date and also dates & periods dependent on the value date
    lng_ValDate = lng_Input
    If lng_ValDate > lng_MatSpotDate Then lng_ValDate = lng_MatSpotDate  ' Prevent accumulation of expired deal
    lng_SpotDate = cyGetFXCrossSpotDate(fld_Params.CCY_Fgn, fld_Params.CCY_Dom, lng_ValDate, dic_GlobalStaticInfo)
    dbl_TimeToMat = calc_yearfrac(lng_ValDate, lng_MatDate, "ACT/365")
    dbl_TimeEstPeriod = calc_yearfrac(lng_SpotDate, lng_MatSpotDate, "ACT/365")
End Sub

Private Sub SetCurveState(str_curve As String, enu_State As CurveState_IRC, Optional int_PillarIndex As Integer = 0)
    ' ## Set up shift in the market data and underlying components
    If irc_DiscCurve.CurveName = str_curve Then Call irc_DiscCurve.SetCurveState(enu_State, int_PillarIndex)
    If irc_SpotDiscCurve.CurveName = str_curve Then Call irc_SpotDiscCurve.SetCurveState(enu_State, int_PillarIndex)
    If irc_Fgn.CurveName = str_curve Then Call irc_Fgn.SetCurveState(enu_State, int_PillarIndex)
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
        fxv_Vols_XY.VolShift_Sens = dbl_ShiftSize
        bln_IsShifted = True
    End If
    If bln_IsQuanto = True Then
        If fxv_Vols_XQ.CurveName = str_curve Then
            fxv_Vols_XQ.VolShift_Sens = dbl_ShiftSize
            bln_IsShifted = True
        End If
        If fxv_Vols_YQ.CurveName = str_curve Then
            fxv_Vols_YQ.VolShift_Sens = dbl_ShiftSize
            bln_IsShifted = True
        End If
    End If

    ApplyVolShift = bln_IsShifted
End Function


' ## METHODS - INTERMEDIATE CALCULATIONS
Private Function GetVol(enu_Pair As VolPair, dbl_LookupStrike As Double) As Double
    ' Determine vol curve to look up
    Dim dbl_Output As Double
    Dim fxv_VolCurve As Data_FXVols, dbl_VolShift As Double
    Select Case enu_Pair
        Case VolPair.XY: Set fxv_VolCurve = fxv_Vols_XY
        Case VolPair.XQ: Set fxv_VolCurve = fxv_Vols_XQ
        Case VolPair.YQ: Set fxv_VolCurve = fxv_Vols_YQ
    End Select

    ' Obtain the vol from the curve
    If fld_Params.IsSmile = True And dbl_LookupStrike <> -1 Then
        If fld_Params.LateType <> "LATE DELIVERY" Then
            dbl_Output = fxv_VolCurve.Lookup_SmileVol(lng_MatDate, dbl_LookupStrike, , fld_Params.IsRescaling)
        Else
            ' Late Delivery uses forward fixed in delivery date later than T+2
            Dim dbl_Fwd As Double
            dbl_Fwd = Forward()
            dbl_Output = fxv_VolCurve.Lookup_SmileVol(lng_MatDate, dbl_LookupStrike, , fld_Params.IsRescaling, , dbl_Fwd)
        End If
    Else
        dbl_Output = fxv_VolCurve.Lookup_ATMVol(lng_MatDate)
    End If

    GetVol = dbl_Output
End Function

Private Function BSDigital(dbl_Strike As Double, dbl_Spot As Double, dbl_DF_ValSpot As Double, dbl_DF_MatSpot As Double) As Double

    Dim dbl_SpotDelta As Double, dbl_UnitValSpotShockUp As Double, dbl_UnitValSpotShockDown As Double, dbl_ShockSize As Double, dbl_Fwd As Double, dbl_VolPct_XY As Double, dbl_Output As Double
    Dim str_TargetCCy As String

    dbl_ShockSize = 0.01

    dbl_Spot = Spot()

    If fld_Params.CCY_Fgn = "USD" Then
        str_TargetCCy = fld_Params.CCY_Dom
    ElseIf fld_Params.CCY_Dom = "USD" Then
        str_TargetCCy = fld_Params.CCY_Fgn
    ElseIf (fxs_Spots.Lookup_Quotation(fld_Params.CCY_Dom) = "INDIRECT" And fxs_Spots.Lookup_Quotation(fld_Params.CCY_Fgn) = "INDIRECT") Then
        str_TargetCCy = fld_Params.CCY_Fgn
    Else
        str_TargetCCy = fld_Params.CCY_Dom
    End If

    Call fxs_Spots.Scen_StoreOrigRate
    Call fxs_Spots.Scen_TempRate

    Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
    Call fxs_Spots.Scen_ApplyCurrent
    dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
    dbl_Fwd = Forward()
    dbl_UnitValSpotShockUp = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, 1) * dbl_DF_MatSpot
    Call fxs_Spots.Scen_ApplyBase

    Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
    Call fxs_Spots.Scen_ApplyCurrent
    dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
    dbl_Fwd = Forward()
    dbl_UnitValSpotShockDown = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, 1) * dbl_DF_MatSpot
    Call fxs_Spots.Scen_ApplyBase

    dbl_SpotDelta = (dbl_UnitValSpotShockUp - dbl_UnitValSpotShockDown) / (2 * dbl_ShockSize / 100 * dbl_Spot)

    dbl_VolPct_XY = GetVol(VolPair.XY, dbl_Strike)
    dbl_Fwd = Forward()

    If enu_Payoff = EuropeanPayoff.Digital_AoN Then
        dbl_Output = fld_Params.Direction * dbl_Spot * dbl_SpotDelta * dbl_DF_ValSpot

    Else
        dbl_Output = fld_Params.Direction * (dbl_Spot * dbl_SpotDelta - Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, 1) _
                    * dbl_DF_MatSpot) * dbl_DF_ValSpot
    End If

    Call fxs_Spots.Scen_RestoreOrigRate
    BSDigital = dbl_Output
End Function

'-------------------------------------------------------------------------------------------
' NAME:    Late_Del_Atm_Spot
'
' PURPOSE: Calculate Late Delivery ATM Spot market value
'
' NOTES: Late Delivery ATM Spot is FX vanilla option, fix at T+2, delivery later than T+2
'        Can be replicated using Late Cash and AoN digital option
'
' INPUT OPTIONS:
'   dbl_DF_MatSpot - DF from late delivery date to spot
'   dbl_DF_MatSpot_Std - DF from T+2 to spot
'
' MODIFIED:
'    31JAN2020 - KW - Creation
'-------------------------------------------------------------------------------------------
Private Function Late_Del_Atm_Spot(dbl_Strike As Double, dbl_Fwd As Double, dbl_Spot As Double, dbl_VolPct_XY As Double, dbl_DF_ValSpot As Double, _
                                   dbl_DF_MatSpot As Double, dbl_DF_MatSpot_Std As Double, enu_Payoff As EuropeanPayoff) As Double

    Dim dbl_UnitAoN As Double, dbl_UnitLateCash As Double, dbl_Output As Double, dbl_df_diff As Double

    '------------------------------------------------
    ' Obtain DF of DOM and FOR between T+2 and delivery
    '------------------------------------------------
    ' Variable to store DF between T+2 and delivery
    Dim dbl_Fgn_LateDel_DF As Double, dbl_Dom_LateDel_DF As Double

    ' irc_DiscCurve is the domestic FX curve, irc_Fgn is the foreign FX curve
    dbl_Fgn_LateDel_DF = irc_Fgn.Lookup_Rate(lng_MatSpotDate_Std, lng_MatSpotDate, "DF")
    dbl_Dom_LateDel_DF = irc_DiscCurve.Lookup_Rate(lng_MatSpotDate_Std, lng_MatSpotDate, "DF")

    '--------------------------------
    ' Calculate difference between DF
    '--------------------------------
    dbl_df_diff = dbl_Fgn_LateDel_DF - dbl_Dom_LateDel_DF

    '----------------------
    ' Calculate Late Cash
    '----------------------
    dbl_UnitLateCash = Calc_BSPrice_Vanilla(fld_Params.Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_XY, enu_Payoff) _
                       * dbl_DF_MatSpot * dbl_DF_ValSpot

    '----------------------
    ' Calculate AoN
    '----------------------
    enu_Payoff = EuropeanPayoff.Digital_AoN
    dbl_UnitAoN = BSDigital(dbl_Strike, dbl_Spot, dbl_DF_ValSpot, dbl_DF_MatSpot_Std)
    enu_Payoff = EuropeanPayoff.Standard

    '---------------------------------
    ' Calculate Latr Delivery ATM Spot
    '---------------------------------
    dbl_Output =  dbl_UnitLateCash + fld_Params.Direction * dbl_df_diff * dbl_UnitAoN

    Late_Del_Atm_Spot = dbl_Output
End Function

' ## METHODS - CALCULATION DETAILS
Public Sub OutputReport(wks_output As Worksheet)
    wks_output.Cells.Clear

    Dim rng_TopLeft As Range: Set rng_TopLeft = wks_output.Range("A1")
    Dim str_CurrencyFormat As String: str_CurrencyFormat = Gather_CurrencyFormat()
    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    Dim dic_Addresses As New Dictionary: dic_Addresses.CompareMode = CompareMethod.TextCompare
    Dim rng_PnL As Range
    Dim int_ActiveRow As Integer: int_ActiveRow = 0
    Dim enu_DetailedCat As DetailedCat: enu_DetailedCat = DetailedCategory()

    With rng_TopLeft
        ' Display PnL
        .Offset(int_ActiveRow, 0).Value = "OVERALL"
        .Offset(int_ActiveRow, 0).Font.Italic = True

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Value date:"
        .Offset(int_ActiveRow, 1).Value = lng_ValDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("ValDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "PnL:"
        .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
        .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
        Set rng_PnL = .Offset(int_ActiveRow, 1)

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "CCY:"
        .Offset(int_ActiveRow, 1).Value = fld_Params.CCY_PnL
        Call dic_Addresses.Add("PnLCCY", .Offset(int_ActiveRow, 1).Address(False, False))

        ' Display MV
        int_ActiveRow = int_ActiveRow + 2
        .Offset(int_ActiveRow, 0).Value = "OPTION LEG"
        .Offset(int_ActiveRow, 0).Font.Italic = True

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Qty:"
        .Offset(int_ActiveRow, 1).Value = fld_Params.Notional_Fgn
        .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
        Call dic_Addresses.Add("Notional", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Fgn CCY:"
        .Offset(int_ActiveRow, 1).Value = fld_Params.CCY_Fgn
        Call dic_Addresses.Add("FgnCCY", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Dom CCY:"
        .Offset(int_ActiveRow, 1).Value = fld_Params.CCY_Dom
        Call dic_Addresses.Add("DomCCY", .Offset(int_ActiveRow, 1).Address(False, False))

        If bln_IsQuanto = True Or fld_Params.IsDigital = True Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Payout CCY:"
            .Offset(int_ActiveRow, 1).Value = fld_Params.CCY_Payout
            Call dic_Addresses.Add("PayoutCCY", .Offset(int_ActiveRow, 1).Address(False, False))
        End If

        If bln_IsQuanto = True Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Quanto Factor:"
            .Offset(int_ActiveRow, 1).Value = fld_Params.QuantoFactor
            Call dic_Addresses.Add("QuantoFactor", .Offset(int_ActiveRow, 1).Address(False, False))
        End If

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Position:"
        If fld_Params.IsBuy = True Then .Offset(int_ActiveRow, 1).Value = "B" Else .Offset(int_ActiveRow, 1).Value = "S"
        Call dic_Addresses.Add("BS", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Payout:"
        If fld_Params.Direction = OptionDirection.CallOpt Then
            .Offset(int_ActiveRow, 1).Value = "Call"
        Else
            .Offset(int_ActiveRow, 1).Value = "Put"
        End If

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Payoff:"
        .Offset(int_ActiveRow, 1).Value = Convert_PayoffCode(enu_Payoff)

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Spot:"
        .Offset(int_ActiveRow, 1).Value = Spot()
        Call dic_Addresses.Add("Spot", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Forward:"
        .Offset(int_ActiveRow, 1).Value = Forward()
        Call dic_Addresses.Add("Fwd", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Strike:"
        .Offset(int_ActiveRow, 1).Value = dbl_Strike
        Call dic_Addresses.Add("Strike", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Option Vol:"
        .Offset(int_ActiveRow, 1).Value = GetVol(VolPair.XY, dbl_Strike)
        Call dic_Addresses.Add("Vol_XY", .Offset(int_ActiveRow, 1).Address(False, False))

        If enu_DetailedCat = DetailedCat.Eur_DigitalCS Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Option Vol (with spread):"
            If fld_Params.IsSmile = True Then
                .Offset(int_ActiveRow, 1).Value = GetVol(VolPair.XY, dbl_Strike * (1 + fld_Params.CSFactor))
            Else
                .Offset(int_ActiveRow, 1).Value = GetVol(VolPair.XY, dbl_Strike)
            End If
            Call dic_Addresses.Add("Vol_XY_WithCS", .Offset(int_ActiveRow, 1).Address(False, False))
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

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Spot Date:"
        .Offset(int_ActiveRow, 1).Value = lng_SpotDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("SpotDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Maturity:"
        .Offset(int_ActiveRow, 1).Value = lng_MatDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("MatDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Mat Spot:"
        .Offset(int_ActiveRow, 1).Value = lng_MatSpotDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("MatSpotDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Delivery:"
        .Offset(int_ActiveRow, 1).Value = fld_Params.DelivDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("DelivDate", .Offset(int_ActiveRow, 1).Address(False, False))

        If bln_IsQuanto = True Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Drift Adj Fwd:"
            .Offset(int_ActiveRow, 1).Formula = "=" & dic_Addresses("Fwd") & "*Exp(-" & dic_Addresses("Correl_XQ") & "*" _
                & dic_Addresses("ATMVol_XY") & "/100*" & dic_Addresses("ATMVol_YQ") & "/100*Calc_YearFrac(" & dic_Addresses("ValDate") _
                & "," & dic_Addresses("MatDate") & ",""ACT/365""))"
            Call dic_Addresses.Add("Fwd_DriftAdj", .Offset(int_ActiveRow, 1).Address(False, False))
        End If

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Spot DF:"
        .Offset(int_ActiveRow, 1).Formula = "=cyReadIRCurve(""" & fld_Params.Curve_SpotDisc & """," & dic_Addresses("ValDate") _
            & "," & dic_Addresses("SpotDate") & ",""DF"")"
        Call dic_Addresses.Add("SpotDF", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Option DF:"
        .Offset(int_ActiveRow, 1).Formula = "=cyReadIRCurve(""" & fld_Params.Curve_Disc & """," & dic_Addresses("SpotDate") _
            & "," & dic_Addresses("DelivDate") & ",""DF"")"
        Call dic_Addresses.Add("OptionDF", .Offset(int_ActiveRow, 1).Address(False, False))

        If enu_DetailedCat = DetailedCat.Eur_Digital Then
            Dim dbl_ShockSize As Double, str_TargetCCy As String
            dbl_ShockSize = 0.01

            If fld_Params.CCY_Fgn = "USD" Then
                str_TargetCCy = fld_Params.CCY_Dom
            ElseIf fld_Params.CCY_Dom = "USD" Then
                    str_TargetCCy = fld_Params.CCY_Fgn
            ElseIf (fxs_Spots.Lookup_Quotation(fld_Params.CCY_Dom) = "INDIRECT" And fxs_Spots.Lookup_Quotation(fld_Params.CCY_Fgn) = "INDIRECT") Then
                str_TargetCCy = fld_Params.CCY_Fgn
            Else
                str_TargetCCy = fld_Params.CCY_Dom
            End If
            Call fxs_Spots.Scen_StoreOrigRate
            Call fxs_Spots.Scen_TempRate
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Shock Spot Up:"
            .Offset(int_ActiveRow, 1).Value = Spot()
            Call dic_Addresses.Add("ShockSpotUp", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Shock Forward Up:"
            .Offset(int_ActiveRow, 1).Value = Forward()
            Call dic_Addresses.Add("ShockFwdUp", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Option Vol (Spot Shock Up):"
            .Offset(int_ActiveRow, 1).Value = GetVol(VolPair.XY, dbl_Strike)
            Call dic_Addresses.Add("Vol_XY_ShockSpotUp", .Offset(int_ActiveRow, 1).Address(False, False))

            Call fxs_Spots.Scen_ApplyBase
            Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
            Call fxs_Spots.Scen_ApplyCurrent

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Shock Spot Down:"
            .Offset(int_ActiveRow, 1).Value = Spot()
            Call dic_Addresses.Add("ShockSpotDown", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Shock Forward Down:"
            .Offset(int_ActiveRow, 1).Value = Forward()
            Call dic_Addresses.Add("ShockFwdDown", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Option Vol (Spot Shock Down):"
            .Offset(int_ActiveRow, 1).Value = GetVol(VolPair.XY, dbl_Strike)
            Call dic_Addresses.Add("Vol_XY_ShockSpotDown", .Offset(int_ActiveRow, 1).Address(False, False))

            Call fxs_Spots.Scen_ApplyBase

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = " Spot Delta(Finite Difference):"
            .Offset(int_ActiveRow, 1).Value = "=(Calc_BSPrice_Vanilla(" & fld_Params.Direction & "," & dic_Addresses("ShockFwdUp") & "," & dic_Addresses("Strike") _
                    & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," & dic_Addresses("MatDate") & ",""ACT/365"")," & dic_Addresses("Vol_XY_ShockSpotUp") _
                    & "," & "1" & ")" & "-" & "Calc_BSPrice_Vanilla(" & fld_Params.Direction & "," & dic_Addresses("ShockFwdDown") & "," & dic_Addresses("Strike") _
                    & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," & dic_Addresses("MatDate") & ",""ACT/365"")," & dic_Addresses("Vol_XY_ShockSpotDown") _
                    & "," & "1" & "))" & "/" & "(2*" & (dbl_ShockSize / 100) & "*" & dic_Addresses("Spot") & ")*" & dic_Addresses("OptionDF")
            Call dic_Addresses.Add("SpotDelta", .Offset(int_ActiveRow, 1).Address(False, False))
            Call fxs_Spots.Scen_RestoreOrigRate
        End If

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "MV (" & fld_Params.CCY_PnL & "):"
        .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
        .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
        Call dic_Addresses.Add("MV", .Offset(int_ActiveRow, 1).Address(False, False))

        Select Case enu_DetailedCat
            Case DetailedCat.Eur_Standard
                .Offset(int_ActiveRow, 1).Formula = "=IF(" & dic_Addresses("BS") & "=""S"",-1,1)*" & dic_Addresses("Notional") _
                    & "*Calc_BSPrice_Vanilla(" & fld_Params.Direction & "," & dic_Addresses("Fwd") & "," & dic_Addresses("Strike") _
                    & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," & dic_Addresses("MatDate") & ",""ACT/365"")," & dic_Addresses("Vol_XY") _
                    & "," & enu_Payoff & ")*" & dic_Addresses("OptionDF") & "*" & dic_Addresses("SpotDF") & "*cyGetFXDiscSpot(" _
                    & dic_Addresses("DomCCY") & "," & dic_Addresses("PnLCCY") & ")"
            Case DetailedCat.Eur_Digital
                If enu_Payoff = EuropeanPayoff.Digital_AoN Then
                    .Offset(int_ActiveRow, 1).Formula = "=IF(" & dic_Addresses("BS") & "=""S"",-1,1)*" & dic_Addresses("Notional") & "*" & fld_Params.Direction _
                        & "*(" & dic_Addresses("Spot") & "*" & dic_Addresses("SpotDelta") & ")" & "*" & dic_Addresses("SpotDF") & _
                        "*cyGetFXDiscSpot(" & dic_Addresses("DomCCY") & "," & dic_Addresses("PnLCCY") & ")"
                Else
                    .Offset(int_ActiveRow, 1).Formula = "=IF(" & dic_Addresses("BS") & "=""S"",-1,1)*" & dic_Addresses("Notional") & "*" & fld_Params.Direction _
                        & "*(" & dic_Addresses("Spot") & "*" & dic_Addresses("SpotDelta") & "-" & "Calc_BSPrice_Vanilla(" & fld_Params.Direction & "," & _
                        dic_Addresses("Fwd") & "," & dic_Addresses("Strike") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," & dic_Addresses("MatDate") _
                        & ",""ACT/365"")," & dic_Addresses("Vol_XY") & "," & 1 & ")*" & dic_Addresses("OptionDF") & ")" & "*" & dic_Addresses("SpotDF") & _
                        "*cyGetFXDiscSpot(" & dic_Addresses("DomCCY") & "," & dic_Addresses("PnLCCY") & ")"
                End If
            Case DetailedCat.Eur_DigitalCS
                .Offset(int_ActiveRow, 1).Formula = "=IF(" & dic_Addresses("BS") & "=""S"",-1,1)*" & dic_Addresses("Notional") _
                    & "*Calc_BSPrice_DigitalCS(" & fld_Params.Direction & "," & dic_Addresses("Fwd") & "," & dic_Addresses("Strike") _
                    & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," & dic_Addresses("MatDate") & ",""ACT/365"")," & dic_Addresses("Vol_XY") & "," _
                    & dic_Addresses("Vol_XY_WithCS") & "," & dic_Addresses("PayoutCCY") & "=" & dic_Addresses("DomCCY") & "," & fld_Params.CSFactor & ")*" _
                    & dic_Addresses("OptionDF") & "*" & dic_Addresses("SpotDF") & "*cyGetFXDiscSpot(" & dic_Addresses("DomCCY") _
                    & "," & dic_Addresses("PnLCCY") & ")"
            Case DetailedCat.Eur_Quanto
                .Offset(int_ActiveRow, 1).Formula = "=IF(" & dic_Addresses("BS") & "=""S"",-1,1)*" & dic_Addresses("Notional") _
                    & "*Calc_BSPrice_Vanilla(" & fld_Params.Direction & "," & dic_Addresses("Fwd_DriftAdj") & "," _
                    & dic_Addresses("Strike") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," & dic_Addresses("MatDate") & ",""ACT/365"")," _
                    & dic_Addresses("Vol_XY") & "," & enu_Payoff & ")*" & dic_Addresses("QuantoFactor") & "*" & dic_Addresses("OptionDF") & "*" _
                    & dic_Addresses("SpotDF") & "*cyGetFXDiscSpot(" & dic_Addresses("PayoutCCY") & "," & dic_Addresses("PnLCCY") & ")"
            Case DetailedCat.Eur_DigitalQuanto
                .Offset(int_ActiveRow, 1).Formula = "=IF(" & dic_Addresses("BS") & "=""S"",-1,1)*" & dic_Addresses("Notional") _
                    & "*Calc_BSPrice_Vanilla(" & fld_Params.Direction & "," & dic_Addresses("Fwd_DriftAdj") & "," _
                    & dic_Addresses("Strike") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," & dic_Addresses("MatDate") & ",""ACT/365"")," _
                    & dic_Addresses("Vol_XY") & "," & enu_Payoff & ")*" & dic_Addresses("QuantoFactor") & "*" & dic_Addresses("OptionDF") _
                    & "*" & dic_Addresses("SpotDF") & "*cyGetFXDiscSpot(" & dic_Addresses("PayoutCCY") & "," _
                    & dic_Addresses("PnLCCY") & ")"
            Case DetailedCat.Amr_Standard
                .Offset(int_ActiveRow, 1).Formula = "=IF(" & dic_Addresses("BS") & "=""S"",-1,1)*" & dic_Addresses("Notional") _
                    & "*Calc_BAW_American(" & fld_Params.Direction & "," & dic_Addresses("Spot") & "," & dic_Addresses("Fwd") & "," _
                    & dic_Addresses("Strike") & "," & dic_Addresses("Vol_XY") & "," & dic_Addresses("OptionDF") & ",Calc_YearFrac(" _
                    & dic_Addresses("ValDate") & "," & dic_Addresses("MatDate") & ",""ACT/365""),Calc_YearFrac(" & dic_Addresses("SpotDate") & "," _
                    & dic_Addresses("MatSpotDate") & ",""ACT/365""))*" & dic_Addresses("SpotDF") & "*cyGetFXDiscSpot(" & dic_Addresses("DomCCY") _
                    & "," & dic_Addresses("PnLCCY") & ")"
        End Select

        ' Output premium flow
        int_ActiveRow = int_ActiveRow + 2
        With rng_TopLeft
            .Offset(int_ActiveRow, 0).Value = "PREMIUM LEG"
            .Offset(int_ActiveRow, 0).Font.Italic = True

            int_ActiveRow = int_ActiveRow + 1
            Call scf_Premium.OutputReport(.Offset(int_ActiveRow, 0), "Cash", _
                fld_Params.CCY_PnL, -int_Sign, True, dic_Addresses, False)
        End With

        ' Display PnL formula
        rng_PnL.Formula = "=" & dic_Addresses("MV") & "+" & dic_Addresses("SCF_PV")
    End With

    wks_output.Calculate
    wks_output.Cells.HorizontalAlignment = xlCenter
    wks_output.Columns.AutoFit
End Sub