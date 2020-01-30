Option Explicit

Private Const dbl_MinFwd As Double = 0.0000000001

Public Function Calc_BSPrice_Vanilla(enu_Direction As OptionDirection, dbl_Fwd As Double, dbl_Strike As Double, _
    dbl_TimeToMat As Double, dbl_VolPct As Double, Optional enu_Payoff As EuropeanPayoff = EuropeanPayoff.Standard, _
    Optional enu_Method As NormalCDFMethod = NormalCDFMethod.Excel) As Double
    ' ## Gives undsicounted vanilla option price per unit of foreign notional, using Black model (on forward), gives payoff if expired
    ' ## Also can give BS price for digital payoff options
    Dim dbl_Output As Double
    Dim dbl_Fwd_ForD1 As Double
    If dbl_Fwd < dbl_MinFwd Then dbl_Fwd_ForD1 = dbl_MinFwd Else dbl_Fwd_ForD1 = dbl_Fwd  ' Floor forward used in D1 calculation, otherwise math error will occur

    If dbl_TimeToMat <= 0 Then
        ' Option expired
        Select Case enu_Payoff
            Case EuropeanPayoff.Standard
                dbl_Output = (dbl_Fwd - dbl_Strike) * enu_Direction
                If dbl_Output < 0 Then dbl_Output = 0
            Case EuropeanPayoff.Digital_CoN
                dbl_Output = Calc_Heaviside((dbl_Fwd - dbl_Strike) * enu_Direction) * dbl_Strike
            Case EuropeanPayoff.Digital_AoN
                dbl_Output = Calc_Heaviside((dbl_Fwd - dbl_Strike) * enu_Direction) * dbl_Fwd
        End Select
    Else
        ' Option is live
        Dim dbl_VolSqrtT As Double: dbl_VolSqrtT = (dbl_VolPct / 100) * Sqr(dbl_TimeToMat)
        Dim dbl_D1 As Double: dbl_D1 = Calc_BS_D1(dbl_Fwd_ForD1, dbl_Strike, dbl_TimeToMat, dbl_VolPct)
        Dim dbl_D2 As Double: dbl_D2 = dbl_D1 - dbl_VolSqrtT
        Dim dbl_AoN_Call As Double: dbl_AoN_Call = dbl_Fwd * Calc_NormalCDF(dbl_D1, enu_Method)
        Dim dbl_CoN_Call As Double: dbl_CoN_Call = Calc_NormalCDF(dbl_D2, enu_Method)

        If enu_Direction = OptionDirection.CallOpt Then
            Select Case enu_Payoff
                Case EuropeanPayoff.Standard: dbl_Output = dbl_AoN_Call - dbl_Strike * dbl_CoN_Call
                Case EuropeanPayoff.Digital_CoN: dbl_Output = dbl_CoN_Call * dbl_Strike
                Case EuropeanPayoff.Digital_AoN: dbl_Output = dbl_AoN_Call
            End Select
        Else
            Dim dbl_AoN_Put As Double: dbl_AoN_Put = dbl_Fwd - dbl_AoN_Call
            Dim dbl_CoN_Put As Double: dbl_CoN_Put = 1 - dbl_CoN_Call

            Select Case enu_Payoff
                Case EuropeanPayoff.Standard: dbl_Output = dbl_Strike * dbl_CoN_Put - dbl_AoN_Put
                Case EuropeanPayoff.Digital_CoN: dbl_Output = dbl_CoN_Put * dbl_Strike
                Case EuropeanPayoff.Digital_AoN: dbl_Output = dbl_AoN_Put
            End Select
        End If
    End If

    Calc_BSPrice_Vanilla = dbl_Output
End Function

Public Function Calc_BSPrice_DigitalCS(enu_Direction As OptionDirection, dbl_Fwd As Double, dbl_Strike As Double, _
    dbl_TimeToMat As Double, dbl_VolPct_Orig As Double, dbl_VolPct_Shifted As Double, _
    bln_IsPayoutDom As Boolean, Optional dbl_RelSpread As Double = 0.0001) As Double
    ' ## Undiscounted price of digital option using call spread replication
    ' ## Result is per unit of FGN currency, even if the payout is DOM
    ' Determine cash-or-nothing digital price
    Dim dbl_ShiftAmt As Double: dbl_ShiftAmt = dbl_Strike * dbl_RelSpread
    Dim dbl_Orig As Double: dbl_Orig = Calc_BSPrice_Vanilla(enu_Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_Orig)
    Dim dbl_Orig_Shifted As Double: dbl_Orig_Shifted = Calc_BSPrice_Vanilla(enu_Direction, dbl_Fwd, dbl_Strike + dbl_ShiftAmt, dbl_TimeToMat, dbl_VolPct_Shifted, EuropeanPayoff.Standard)
    Dim dbl_DigitalMoney As Double: dbl_DigitalMoney = enu_Direction * (dbl_Orig - dbl_Orig_Shifted) / dbl_ShiftAmt * dbl_Strike

    ' Output final price
    If bln_IsPayoutDom Then
        Calc_BSPrice_DigitalCS = dbl_DigitalMoney
    Else
        Calc_BSPrice_DigitalCS = dbl_DigitalMoney + enu_Direction * Calc_BSPrice_Vanilla(enu_Direction, dbl_Fwd, _
            dbl_Strike, dbl_TimeToMat, dbl_VolPct_Orig, EuropeanPayoff.Standard)
    End If
End Function
Public Function Calc_BSPrice_DigitalSmileOn(enu_Direction As OptionDirection, dbl_Fwd As Double, dbl_Strike As Double, _
    dbl_TimeToMat As Double, dbl_VolPct_Orig As Double, dbl_VolPct_Shifted As Double, Optional dbl_RelSpread As Double = 0.0001) As Double
    Dim dbl_Vega As Double: dbl_Vega = Calc_BS_Vega_Vanilla(dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_Orig) / 100

    Dim dbl_DigitalMoney As Double: dbl_DigitalMoney = enu_Direction * dbl_Vega * (dbl_VolPct_Orig - dbl_VolPct_Shifted) / dbl_RelSpread

    Calc_BSPrice_DigitalSmileOn = dbl_DigitalMoney + Calc_BSPrice_Vanilla(enu_Direction, dbl_Fwd, _
            dbl_Strike, dbl_TimeToMat, dbl_VolPct_Orig, EuropeanPayoff.Digital_CoN) / dbl_Strike


End Function
Public Function Calc_BSPrice_SingleAmericanBar(enu_Direction As OptionDirection, dbl_Spot As Double, dbl_Fwd As Double, _
    dbl_Strike As Double, dbl_VolPct As Double, dbl_TimeToMat As Double, dbl_Barrier As Double, _
    bln_IsUpBarrier As Boolean, bln_IsKnockOut As Boolean) As Double
    ' ## Gives undiscounted single barrier option price, using the Black model

    Dim dbl_KnockOutPrice As Double

    ' Determine properties of barrier case for those which are independent of whether barrier is knocked already
    Dim int_Sign_Bar As Integer: If bln_IsUpBarrier = True Then int_Sign_Bar = 1 Else int_Sign_Bar = -1

    Dim bln_AlreadyKnocked As Boolean
    bln_AlreadyKnocked = ((bln_IsUpBarrier = True And dbl_Spot >= dbl_Barrier) _
        Or (bln_IsUpBarrier = False And dbl_Spot <= dbl_Barrier))

    ' ## Handle special cases
    Dim bln_Handled As Boolean, dbl_HandledValue As Double
    If bln_AlreadyKnocked = True Then
        ' Already knocked
        If bln_IsKnockOut = True Then
            dbl_HandledValue = 0
        Else
            ' Convert to vanilla option
            dbl_HandledValue = Calc_BSPrice_Vanilla(enu_Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct)
        End If
        bln_Handled = True
    ElseIf dbl_TimeToMat <= 0 Then
        ' Not knocked but matured
        If bln_IsKnockOut = True Then
            dbl_HandledValue = enu_Direction * (dbl_Spot - dbl_Strike)
            If dbl_HandledValue < 0 Then dbl_HandledValue = 0
        Else
            dbl_HandledValue = 0
        End If
        bln_Handled = True
    Else
        bln_Handled = False
    End If

    If bln_Handled = True Then
        Calc_BSPrice_SingleAmericanBar = dbl_HandledValue
        Exit Function
    End If

    ' ## Handle typical cases
    Dim dbl_VolSqrtT As Double: dbl_VolSqrtT = dbl_VolPct / 100 * Sqr(dbl_TimeToMat)

    ' Calculate values which may be used if knocked in
    Dim dbl_TotalDrift_BN As Double: dbl_TotalDrift_BN = Math.Log(dbl_Fwd / dbl_Spot) - 0.5 * dbl_VolSqrtT ^ 2  ' Drift when using bond as numeraire
    Dim dbl_TotalDrift_SN As Double: dbl_TotalDrift_SN = dbl_TotalDrift_BN + dbl_VolSqrtT ^ 2  ' Drift when using stock as numeraire
    Dim dbl_LogHHSK As Double: dbl_LogHHSK = Math.Log(dbl_Barrier ^ 2 / (dbl_Spot * dbl_Strike))

    ' Calculate normal distribution terms
    ' Kink means strike is between spot and barrier
    Dim dbl_D1_Kink As Double: dbl_D1_Kink = 1 / dbl_VolSqrtT * (Math.Log(dbl_Spot / dbl_Strike) + dbl_TotalDrift_SN)
    Dim dbl_D1_NoKink As Double: dbl_D1_NoKink = 1 / dbl_VolSqrtT * (Math.Log(dbl_Spot / dbl_Barrier) + dbl_TotalDrift_SN)
    Dim dbl_H1_Kink As Double: dbl_H1_Kink = 1 / dbl_VolSqrtT * (dbl_LogHHSK + dbl_TotalDrift_SN)
    Dim dbl_H1_NoKink As Double: dbl_H1_NoKink = 1 / dbl_VolSqrtT * (Math.Log(dbl_Barrier / dbl_Spot) + dbl_TotalDrift_SN)

    ' Calculate drift adjustment terms for distribution of the minimum
    ' ## DEV - Refactor intermediate calcs here to allow for digitals

    ' Calculate price terms
    Dim dbl_PriceD_Kink As Double: dbl_PriceD_Kink = enu_Direction * (dbl_Fwd * Calc_NormalCDF(enu_Direction * dbl_D1_Kink) _
        - dbl_Strike * Calc_NormalCDF(enu_Direction * (dbl_D1_Kink - dbl_VolSqrtT)))  ' Smile-off vanilla price
    Dim dbl_PriceD_NoKink As Double: dbl_PriceD_NoKink = enu_Direction * (dbl_Fwd * Calc_NormalCDF(enu_Direction * dbl_D1_NoKink) _
        - dbl_Strike * Calc_NormalCDF(enu_Direction * (dbl_D1_NoKink - dbl_VolSqrtT)))
    Dim dbl_PriceH_Kink As Double: dbl_PriceH_Kink = enu_Direction * (dbl_Fwd * (dbl_Barrier / dbl_Spot) ^ (dbl_TotalDrift_SN / (0.5 * dbl_VolSqrtT ^ 2)) _
        * Calc_NormalCDF(-int_Sign_Bar * dbl_H1_Kink) - dbl_Strike * (dbl_Barrier / dbl_Spot) ^ (dbl_TotalDrift_BN / (0.5 * dbl_VolSqrtT ^ 2)) _
        * Calc_NormalCDF(-int_Sign_Bar * (dbl_H1_Kink - dbl_VolSqrtT)))
    Dim dbl_PriceH_NoKink As Double: dbl_PriceH_NoKink = enu_Direction * (dbl_Fwd * (dbl_Barrier / dbl_Spot) ^ (dbl_TotalDrift_SN / (0.5 * dbl_VolSqrtT ^ 2)) _
        * Calc_NormalCDF(-int_Sign_Bar * dbl_H1_NoKink) - dbl_Strike * (dbl_Barrier / dbl_Spot) ^ (dbl_TotalDrift_BN / (0.5 * dbl_VolSqrtT ^ 2)) _
        * Calc_NormalCDF(-int_Sign_Bar * (dbl_H1_NoKink - dbl_VolSqrtT)))

    ' Determine type of barrier and handle remaining cases
    Dim bln_BarBetweenKS As Boolean, bln_BackwardKnock As Boolean
    bln_BarBetweenKS = (bln_IsUpBarrier = True And dbl_Strike >= dbl_Barrier) Or (bln_IsUpBarrier = False And dbl_Strike <= dbl_Barrier)
    bln_BackwardKnock = (enu_Direction * int_Sign_Bar = -1) ' True if moving towards the barrier reduces the payoff

    ' Knock out barriers
    If bln_BackwardKnock = True Then
        If bln_BarBetweenKS = True Then
            dbl_KnockOutPrice = dbl_PriceD_NoKink - dbl_PriceH_NoKink
        Else
            dbl_KnockOutPrice = dbl_PriceD_Kink - dbl_PriceH_Kink
        End If
    Else
        If bln_BarBetweenKS = True Then
            dbl_KnockOutPrice = 0
        Else
            dbl_KnockOutPrice = dbl_PriceD_Kink - dbl_PriceD_NoKink + dbl_PriceH_Kink - dbl_PriceH_NoKink
        End If
    End If

    ' Final output
    If bln_IsKnockOut = True Then
        Calc_BSPrice_SingleAmericanBar = dbl_KnockOutPrice
    Else
        ' Use in-out parity
        Calc_BSPrice_SingleAmericanBar = dbl_PriceD_Kink - dbl_KnockOutPrice
    End If
End Function

Public Function Calc_BSPrice_SingleBarRebate(dbl_Spot As Double, dbl_Fwd As Double, dbl_VolPct As Double, _
    dbl_TimeToMat As Double, dbl_TimeEstPeriod As Double, dbl_DF_MatSpot As Double, bln_IsPayoutDom As Boolean, _
    dbl_Barrier As Double, bln_IsUpBarrier As Boolean, bln_IsKnockOut As Boolean, bln_IsInstantPmt As Boolean) As Double
    ' ## Returns discounted value of the rebate at the spot date, expressed per unit of the domestic currency

    Dim dbl_KnockInPrice As Double
    Dim dbl_DF_Rebate As Double
    If bln_IsInstantPmt = True Then dbl_DF_Rebate = 1 Else dbl_DF_Rebate = dbl_DF_MatSpot

    ' Start by valuing knock-in rebate
    If (bln_IsUpBarrier = True And dbl_Spot >= dbl_Barrier) Or (bln_IsUpBarrier = False And dbl_Spot <= dbl_Barrier) = True Then
        ' Barrier already hit
        dbl_KnockInPrice = dbl_DF_Rebate
    ElseIf dbl_TimeToMat <= 0 Then
        ' Contract matured without hitting
        dbl_KnockInPrice = 0
    Else
        Dim dbl_D1 As Double, dbl_D2 As Double, dbl_PowerA As Double, dbl_PowerB As Double

        ' Calculate input values
        Dim int_Sign_Bar As Integer: If bln_IsUpBarrier = True Then int_Sign_Bar = 1 Else int_Sign_Bar = -1
        Dim dbl_Vol As Double: dbl_Vol = dbl_VolPct / 100
        Dim dbl_VolSqrtT As Double: dbl_VolSqrtT = dbl_Vol * Sqr(dbl_TimeToMat)
        Dim dbl_RateDiff_AvgOverMat As Double: dbl_RateDiff_AvgOverMat = Math.Log(dbl_Fwd / dbl_Spot) / dbl_TimeToMat
        Dim dbl_DiscRate_AvgOverMat As Double: dbl_DiscRate_AvgOverMat = -Math.Log(dbl_DF_MatSpot) / dbl_TimeToMat


        ' Derive theta, nu and rebate discount factor, depending on currency of the rebate
        Dim dbl_Theta_Mat As Double
        If bln_IsPayoutDom = True Then
            dbl_Theta_Mat = dbl_RateDiff_AvgOverMat / dbl_Vol - dbl_Vol / 2
        Else
            dbl_Theta_Mat = dbl_RateDiff_AvgOverMat / dbl_Vol + dbl_Vol / 2
        End If

        Dim dbl_Nu As Double
        If bln_IsInstantPmt = True Then
            dbl_Nu = Sqr(dbl_Theta_Mat ^ 2 + 2 * dbl_DiscRate_AvgOverMat * dbl_TimeEstPeriod / dbl_TimeToMat)
            dbl_DF_Rebate = 1
        Else
            dbl_Nu = Abs(dbl_Theta_Mat)
            dbl_DF_Rebate = dbl_DF_MatSpot
        End If

        ' Derive values for within the normal CDF
        dbl_D1 = (Log(dbl_Spot / dbl_Barrier) - dbl_Vol * dbl_Nu * dbl_TimeToMat) / dbl_VolSqrtT
        dbl_D2 = (Log(dbl_Barrier / dbl_Spot) - dbl_Vol * dbl_Nu * dbl_TimeToMat) / dbl_VolSqrtT

        ' Derive powers
        dbl_PowerA = (dbl_Theta_Mat + dbl_Nu) / dbl_Vol
        dbl_PowerB = (dbl_Theta_Mat - dbl_Nu) / dbl_Vol

        ' Derive knock-in rebate value
        dbl_KnockInPrice = dbl_DF_Rebate * (((dbl_Barrier / dbl_Spot) ^ dbl_PowerA) * Calc_NormalCDF(int_Sign_Bar * dbl_D1) + _
            ((dbl_Barrier / dbl_Spot) ^ dbl_PowerB) * Calc_NormalCDF(-int_Sign_Bar * dbl_D2))
    End If

    ' Convert to knock-out rebate value
    If bln_IsKnockOut = True Then
        Calc_BSPrice_SingleBarRebate = dbl_DF_Rebate - dbl_KnockInPrice
    Else
        Calc_BSPrice_SingleBarRebate = dbl_KnockInPrice
    End If
End Function

Public Function Calc_BSPrice_DoubleBarRebate(dbl_Spot As Double, dbl_Fwd As Double, dbl_VolPct As Double, _
    dbl_TimeToMat As Double, dbl_DF_Rebate As Double, bln_IsPayoutDom As Boolean, dbl_LowerBar As Double, _
    dbl_UpperBar As Double, bln_IsKnockOut As Boolean) As Double
    ' ## Returns discounted value of the rebate at the spot date, expressed per unit of the domestic currency
    ' ## Using formula from MathFinance paper
    Const int_Coverage As Integer = 50
    Dim dbl_KnockOutPrice As Double: dbl_KnockOutPrice = 0

    ' Value knock-out rebate
    If dbl_Spot <= dbl_LowerBar Or dbl_Spot >= dbl_UpperBar Then
        ' Barrier already hit
        dbl_KnockOutPrice = 0
    ElseIf dbl_TimeToMat <= 0 Then
        ' Contract matured without hitting, therefore pays out rebate
        dbl_KnockOutPrice = 1
    Else
        ' Calculate static parameters
        Dim dbl_VolSqrtT As Double: dbl_VolSqrtT = dbl_VolPct / 100 * Sqr(dbl_TimeToMat)
        Dim dbl_UpperBar_Z As Double: dbl_UpperBar_Z = Math.Log(dbl_UpperBar / dbl_Spot) / dbl_VolSqrtT
        Dim dbl_LowerBar_Z As Double: dbl_LowerBar_Z = Math.Log(dbl_LowerBar / dbl_Spot) / dbl_VolSqrtT
        If dbl_UpperBar_Z > 10 Then dbl_UpperBar_Z = 10
        If dbl_LowerBar_Z < -10 Then dbl_LowerBar_Z = -10
        Dim dbl_TotalDrift_ToUse As Double
        If bln_IsPayoutDom = True Then
            dbl_TotalDrift_ToUse = Math.Log(dbl_Fwd / dbl_Spot) - 0.5 * dbl_VolSqrtT ^ 2
        Else
            dbl_TotalDrift_ToUse = Math.Log(dbl_Fwd / dbl_Spot) + 0.5 * dbl_VolSqrtT ^ 2
        End If

        Dim dbl_ThetaTilde As Double: dbl_ThetaTilde = dbl_TotalDrift_ToUse / dbl_VolSqrtT  ' Scaled total drift
        Dim dbl_StepSize As Double: dbl_StepSize = 2 * (dbl_UpperBar_Z - dbl_LowerBar_Z)
        Dim dbl_ActiveEpsilon As Double  ' An incrementing offset away from the barrier Z level

        ' Evaluate expectation
        Dim int_Ctr As Integer
        For int_Ctr = -int_Coverage To int_Coverage
            dbl_ActiveEpsilon = -dbl_ThetaTilde + int_Ctr * dbl_StepSize
            dbl_KnockOutPrice = dbl_KnockOutPrice + Math.Exp(-int_Ctr * dbl_StepSize * dbl_ThetaTilde) * ((Calc_NormalCDF(dbl_UpperBar_Z + dbl_ActiveEpsilon) _
                - Calc_NormalCDF(dbl_LowerBar_Z + dbl_ActiveEpsilon)) - Math.Exp(2 * dbl_ThetaTilde * dbl_UpperBar_Z) _
                * (Calc_NormalCDF(-dbl_UpperBar_Z + dbl_ActiveEpsilon) - Calc_NormalCDF(dbl_LowerBar_Z - 2 * dbl_UpperBar_Z + dbl_ActiveEpsilon)))
        Next int_Ctr
    End If

    ' Convert to knock-in rebate value
    If bln_IsKnockOut = False Then
        Calc_BSPrice_DoubleBarRebate = (1 - dbl_KnockOutPrice) * dbl_DF_Rebate
    Else
        Calc_BSPrice_DoubleBarRebate = dbl_KnockOutPrice * dbl_DF_Rebate
    End If
End Function

Public Function Calc_BSPrice_DoubleBarInstantRebate(dbl_Spot As Double, dbl_Fwd As Double, dbl_VolPct As Double, _
    dbl_TimeToMat As Double, dbl_TimeEstPeriod As Double, dbl_DF_Rebate As Double, bln_IsPayoutDom As Boolean, _
    dbl_LowerBar As Double, dbl_UpperBar As Double) As Double
    ' ## Returns discounted value of the rebate at the spot date, expressed per unit of the domestic currency
    ' ## Payment occurs at knock time
    ' ## Adapted from Quant Central code
    Dim dbl_RebatePrice As Double: dbl_RebatePrice = 0

    ' Value knock-out rebate
    If dbl_Spot <= dbl_LowerBar Or dbl_Spot >= dbl_UpperBar Then
        ' Barrier already hit, rebate should be paid already
        dbl_RebatePrice = 1
    ElseIf dbl_TimeToMat <= 0 Then
        ' Contract matured without hitting, therefore pays nothing
        dbl_RebatePrice = 0
    Else
        ' Calculate static parameters
        Const int_Coverage As Integer = 20
        Dim dbl_Vol As Double: dbl_Vol = dbl_VolPct / 100
        Dim dbl_VolSqrtT As Double: dbl_VolSqrtT = dbl_Vol * Sqr(dbl_TimeToMat)
        Dim dbl_V2T As Double: dbl_V2T = dbl_VolSqrtT ^ 2
        Dim dbl_Scaled_UpperBar As Double: dbl_Scaled_UpperBar = Log(dbl_UpperBar / dbl_Spot)
        Dim dbl_Scaled_LowerBar As Double: dbl_Scaled_LowerBar = Log(dbl_LowerBar / dbl_Spot)
        If dbl_Scaled_UpperBar > 10 * dbl_VolSqrtT Then dbl_Scaled_UpperBar = 10 * dbl_VolSqrtT
        If dbl_Scaled_LowerBar < -10 * dbl_VolSqrtT Then dbl_Scaled_LowerBar = -10 * dbl_VolSqrtT
        Dim dbl_Scaled_BarWidth As Double: dbl_Scaled_BarWidth = dbl_Scaled_UpperBar - dbl_Scaled_LowerBar
        Dim dbl_TotalDrift_ToUse As Double
        If bln_IsPayoutDom = True Then
            dbl_TotalDrift_ToUse = Math.Log(dbl_Fwd / dbl_Spot) - 0.5 * dbl_V2T  ' Drift when using bond as numeraire
        Else
            dbl_TotalDrift_ToUse = Math.Log(dbl_Fwd / dbl_Spot) + 0.5 * dbl_V2T  ' Drift when using stock as numeraire
        End If

        Dim dbl_DiscRate_AvgOverMat As Double: dbl_DiscRate_AvgOverMat = -Math.Log(dbl_DF_Rebate) / dbl_TimeToMat
        Dim dbl_VariFactor As Double: dbl_VariFactor = Sqr((dbl_TotalDrift_ToUse / dbl_V2T) ^ 2 + 2 * dbl_DiscRate_AvgOverMat _
            * (dbl_TimeEstPeriod / dbl_TimeToMat) / dbl_Vol ^ 2)
        Dim dbl_CoefU As Double: dbl_CoefU = Exp(dbl_TotalDrift_ToUse / dbl_V2T * dbl_Scaled_UpperBar)
        Dim dbl_CoefL As Double: dbl_CoefL = Exp(dbl_TotalDrift_ToUse / dbl_V2T * dbl_Scaled_LowerBar)

        ' Prepare intermediate values
        Dim dbl_Active_TermU1 As Double, dbl_Active_TermU2 As Double, dbl_Active_TermL1 As Double, dbl_Active_TermL2 As Double
        Dim dbl_Active_FactorU As Double, dbl_Active_FactorL As Double
        Dim dbl_Active_CentreU As Double, dbl_Active_CentreL As Double
        Dim dbl_Active_HSideU As Double, dbl_Active_HSideL As Double

        Dim int_Ctr As Integer
        For int_Ctr = -int_Coverage To int_Coverage
            dbl_Active_CentreU = (dbl_Scaled_UpperBar + 2 * int_Ctr * dbl_Scaled_BarWidth) / dbl_VolSqrtT
            dbl_Active_CentreL = (dbl_Scaled_LowerBar + 2 * int_Ctr * dbl_Scaled_BarWidth) / dbl_VolSqrtT
            dbl_Active_TermU1 = dbl_Active_CentreU - dbl_VariFactor * dbl_VolSqrtT
            dbl_Active_TermU2 = dbl_Active_CentreU + dbl_VariFactor * dbl_VolSqrtT
            dbl_Active_TermL1 = dbl_Active_CentreL - dbl_VariFactor * dbl_VolSqrtT
            dbl_Active_TermL2 = dbl_Active_CentreL + dbl_VariFactor * dbl_VolSqrtT
            dbl_Active_HSideU = Calc_Heaviside(dbl_Active_CentreU)
            dbl_Active_HSideL = Calc_Heaviside(dbl_Active_CentreL)
            dbl_Active_FactorU = Exp(dbl_VariFactor * dbl_Active_CentreU * dbl_VolSqrtT)
            dbl_Active_FactorL = Exp(dbl_VariFactor * dbl_Active_CentreL * dbl_VolSqrtT)

            ' Running sum to determine price
            dbl_RebatePrice = dbl_RebatePrice + dbl_CoefL * ((Calc_NormalCDF(dbl_Active_TermL1) - dbl_Active_HSideL) / dbl_Active_FactorL _
                + (Calc_NormalCDF(dbl_Active_TermL2) - dbl_Active_HSideL) * dbl_Active_FactorL) _
                - dbl_CoefU * ((Calc_NormalCDF(dbl_Active_TermU1) - dbl_Active_HSideU) / dbl_Active_FactorU _
                + (Calc_NormalCDF(dbl_Active_TermU2) - dbl_Active_HSideU) * dbl_Active_FactorU)
        Next int_Ctr
    End If

    Calc_BSPrice_DoubleBarInstantRebate = dbl_RebatePrice
End Function

Public Function Calc_BSPrice_DoubleAmericanBar(enu_Direction As OptionDirection, dbl_Spot As Double, dbl_Fwd As Double, _
    dbl_Strike As Double, dbl_VolPct As Double, dbl_TimeToMat As Double, dbl_LowerBar As Double, _
    dbl_UpperBar As Double, bln_IsKnockOut As Boolean) As Double
    ' ## Returns undiscounted double barrier option price
    ' ## Adapted from Quant Central code, using formulas developed by Douady (1999)

    ' Handle trivial cases
    Dim dbl_KnockOutPrice As Double: dbl_KnockOutPrice = 0
    Dim dbl_VanillaPrice As Double: dbl_VanillaPrice = Calc_BSPrice_Vanilla(enu_Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct)
    If dbl_Spot <= dbl_LowerBar Or dbl_Spot >= dbl_UpperBar Then
        ' Barrier already hit
        dbl_KnockOutPrice = 0
    ElseIf dbl_TimeToMat <= 0 Then
        ' Contract matured without hitting, therefore pays out same as vanilla
        dbl_KnockOutPrice = dbl_VanillaPrice
    Else
        ' Store commonly used values
        Const int_Coverage As Integer = 50
        Dim dbl_VolSqrtT As Double: dbl_VolSqrtT = dbl_VolPct / 100 * Sqr(dbl_TimeToMat)
        Dim dbl_V2T As Double: dbl_V2T = dbl_VolSqrtT ^ 2
        Dim dbl_Scaled_UpperBar As Double: dbl_Scaled_UpperBar = Log(dbl_UpperBar / dbl_Spot)
        Dim dbl_Scaled_LowerBar As Double: dbl_Scaled_LowerBar = Log(dbl_LowerBar / dbl_Spot)
        If dbl_Scaled_UpperBar > 10 * dbl_VolSqrtT Then dbl_Scaled_UpperBar = 10 * dbl_VolSqrtT
        If dbl_Scaled_LowerBar < -10 * dbl_VolSqrtT Then dbl_Scaled_LowerBar = -10 * dbl_VolSqrtT
        Dim dbl_Scaled_Strike As Double: dbl_Scaled_Strike = Log(dbl_Strike / dbl_Spot)
        Dim dbl_Scaled_ITMBound As Double
        Dim dbl_Scaled_BarWidth As Double: dbl_Scaled_BarWidth = dbl_Scaled_UpperBar - dbl_Scaled_LowerBar
        Dim dbl_TotalDrift_BN As Double: dbl_TotalDrift_BN = Math.Log(dbl_Fwd / dbl_Spot) - 0.5 * dbl_V2T  ' Drift when using bond as numeraire
        Dim dbl_TotalDrift_SN As Double: dbl_TotalDrift_SN = dbl_TotalDrift_BN + dbl_VolSqrtT ^ 2  ' Drift when using stock as numeraire

        ' Prepare intermediate values
        Dim dbl_Active_BN_TermA As Double, dbl_Active_BN_TermB As Double, dbl_Active_BN_TermC As Double, dbl_Active_BN_TermD As Double
        Dim dbl_Active_SN_TermA As Double, dbl_Active_SN_TermB As Double, dbl_Active_SN_TermC As Double, dbl_Active_SN_TermD As Double
        Dim dbl_Active_BN_FactorAB As Double, dbl_Active_BN_FactorCD As Double, dbl_Active_SN_FactorAB As Double, dbl_Active_SN_FactorCD As Double
        Dim dbl_Coef_BN As Double: dbl_Coef_BN = 0
        Dim dbl_Coef_SN As Double: dbl_Coef_SN = 0

        Dim int_Ctr As Integer
        Select Case enu_Direction
            Case OptionDirection.CallOpt
                For int_Ctr = -int_Coverage To int_Coverage
                    dbl_Scaled_ITMBound = Examine_MaxOfPair(dbl_Scaled_LowerBar, dbl_Scaled_Strike)

                    ' Calculate values of normal distribution terms
                    dbl_Active_SN_TermA = (dbl_Scaled_UpperBar + 2 * int_Ctr * dbl_Scaled_BarWidth - dbl_TotalDrift_SN) / dbl_VolSqrtT
                    dbl_Active_SN_TermB = (dbl_Scaled_ITMBound + 2 * int_Ctr * dbl_Scaled_BarWidth - dbl_TotalDrift_SN) / dbl_VolSqrtT
                    dbl_Active_SN_TermC = (2 * dbl_Scaled_UpperBar - dbl_Scaled_ITMBound + 2 * int_Ctr * dbl_Scaled_BarWidth _
                        + dbl_TotalDrift_SN) / dbl_VolSqrtT
                    dbl_Active_SN_TermD = (dbl_Scaled_UpperBar + 2 * int_Ctr * dbl_Scaled_BarWidth + dbl_TotalDrift_SN) / dbl_VolSqrtT
                    dbl_Active_BN_TermA = (dbl_Scaled_UpperBar + 2 * int_Ctr * dbl_Scaled_BarWidth - dbl_TotalDrift_BN) / dbl_VolSqrtT
                    dbl_Active_BN_TermB = (dbl_Scaled_ITMBound + 2 * int_Ctr * dbl_Scaled_BarWidth - dbl_TotalDrift_BN) / dbl_VolSqrtT
                    dbl_Active_BN_TermC = (2 * dbl_Scaled_UpperBar - dbl_Scaled_ITMBound + 2 * int_Ctr * dbl_Scaled_BarWidth _
                        + dbl_TotalDrift_BN) / dbl_VolSqrtT
                    dbl_Active_BN_TermD = (dbl_Scaled_UpperBar + 2 * int_Ctr * dbl_Scaled_BarWidth + dbl_TotalDrift_BN) / dbl_VolSqrtT

                    ' Calculate values of multipliers
                    dbl_Active_SN_FactorAB = Exp(-2 * dbl_TotalDrift_SN / dbl_V2T * int_Ctr * dbl_Scaled_BarWidth)
                    dbl_Active_SN_FactorCD = Exp(2 * dbl_TotalDrift_SN / dbl_V2T * (dbl_Scaled_UpperBar + int_Ctr * dbl_Scaled_BarWidth))
                    dbl_Active_BN_FactorAB = Exp(-2 * dbl_TotalDrift_BN / dbl_V2T * int_Ctr * dbl_Scaled_BarWidth)
                    dbl_Active_BN_FactorCD = Exp(2 * dbl_TotalDrift_BN / dbl_V2T * (dbl_Scaled_UpperBar + int_Ctr * dbl_Scaled_BarWidth))

                    ' Running sum to determine coefficients
                    dbl_Coef_SN = dbl_Coef_SN + dbl_Active_SN_FactorAB * (Calc_NormalCDF(dbl_Active_SN_TermA) - Calc_NormalCDF(dbl_Active_SN_TermB)) _
                        - dbl_Active_SN_FactorCD * (Calc_NormalCDF(dbl_Active_SN_TermC) - Calc_NormalCDF(dbl_Active_SN_TermD))
                    dbl_Coef_BN = dbl_Coef_BN + dbl_Active_BN_FactorAB * (Calc_NormalCDF(dbl_Active_BN_TermA) - Calc_NormalCDF(dbl_Active_BN_TermB)) _
                        - dbl_Active_BN_FactorCD * (Calc_NormalCDF(dbl_Active_BN_TermC) - Calc_NormalCDF(dbl_Active_BN_TermD))
                Next int_Ctr

                dbl_KnockOutPrice = dbl_Coef_SN * dbl_Fwd - dbl_Coef_BN * dbl_Strike
            Case OptionDirection.PutOpt
                For int_Ctr = -int_Coverage To int_Coverage
                    dbl_Scaled_ITMBound = Examine_MinOfPair(dbl_Scaled_Strike, dbl_Scaled_UpperBar)

                    ' Calculate values of normal distribution terms
                    dbl_Active_SN_TermA = (dbl_Scaled_ITMBound + 2 * int_Ctr * dbl_Scaled_BarWidth - dbl_TotalDrift_SN) / dbl_VolSqrtT
                    dbl_Active_SN_TermB = (dbl_Scaled_LowerBar + 2 * int_Ctr * dbl_Scaled_BarWidth - dbl_TotalDrift_SN) / dbl_VolSqrtT
                    dbl_Active_SN_TermC = (2 * dbl_Scaled_UpperBar - dbl_Scaled_LowerBar + 2 * int_Ctr * dbl_Scaled_BarWidth _
                        + dbl_TotalDrift_SN) / dbl_VolSqrtT
                    dbl_Active_SN_TermD = (2 * dbl_Scaled_UpperBar - dbl_Scaled_ITMBound + 2 * int_Ctr * dbl_Scaled_BarWidth _
                        + dbl_TotalDrift_SN) / dbl_VolSqrtT
                    dbl_Active_BN_TermA = (dbl_Scaled_ITMBound + 2 * int_Ctr * dbl_Scaled_BarWidth - dbl_TotalDrift_BN) / dbl_VolSqrtT
                    dbl_Active_BN_TermB = (dbl_Scaled_LowerBar + 2 * int_Ctr * dbl_Scaled_BarWidth - dbl_TotalDrift_BN) / dbl_VolSqrtT
                    dbl_Active_BN_TermC = (2 * dbl_Scaled_UpperBar - dbl_Scaled_LowerBar + 2 * int_Ctr * dbl_Scaled_BarWidth _
                        + dbl_TotalDrift_BN) / dbl_VolSqrtT
                    dbl_Active_BN_TermD = (2 * dbl_Scaled_UpperBar - dbl_Scaled_ITMBound + 2 * int_Ctr * dbl_Scaled_BarWidth _
                        + dbl_TotalDrift_BN) / dbl_VolSqrtT

                    ' Calculate values of multipliers
                    dbl_Active_SN_FactorAB = Exp(-2 * dbl_TotalDrift_SN / dbl_V2T * int_Ctr * dbl_Scaled_BarWidth)
                    dbl_Active_SN_FactorCD = Exp(2 * dbl_TotalDrift_SN / dbl_V2T * (dbl_Scaled_UpperBar + int_Ctr * dbl_Scaled_BarWidth))
                    dbl_Active_BN_FactorAB = Exp(-2 * dbl_TotalDrift_BN / dbl_V2T * int_Ctr * dbl_Scaled_BarWidth)
                    dbl_Active_BN_FactorCD = Exp(2 * dbl_TotalDrift_BN / dbl_V2T * (dbl_Scaled_UpperBar + int_Ctr * dbl_Scaled_BarWidth))

                    ' Running sum to determine coefficients
                    dbl_Coef_SN = dbl_Coef_SN + dbl_Active_SN_FactorAB * (Calc_NormalCDF(dbl_Active_SN_TermA) - Calc_NormalCDF(dbl_Active_SN_TermB)) _
                        - dbl_Active_SN_FactorCD * (Calc_NormalCDF(dbl_Active_SN_TermC) - Calc_NormalCDF(dbl_Active_SN_TermD))
                    dbl_Coef_BN = dbl_Coef_BN + dbl_Active_BN_FactorAB * (Calc_NormalCDF(dbl_Active_BN_TermA) - Calc_NormalCDF(dbl_Active_BN_TermB)) _
                        - dbl_Active_BN_FactorCD * (Calc_NormalCDF(dbl_Active_BN_TermC) - Calc_NormalCDF(dbl_Active_BN_TermD))
                Next int_Ctr

                dbl_KnockOutPrice = dbl_Coef_BN * dbl_Strike - dbl_Coef_SN * dbl_Fwd
        End Select

        If dbl_KnockOutPrice < 0 Then
            Debug.Assert dbl_KnockOutPrice > -0.00001
            dbl_KnockOutPrice = 0
        End If
    End If

    ' Handle knock-in barriers using in-out parity
    If bln_IsKnockOut = True Then
        Calc_BSPrice_DoubleAmericanBar = dbl_KnockOutPrice
    Else
        Calc_BSPrice_DoubleAmericanBar = dbl_VanillaPrice - dbl_KnockOutPrice
    End If
End Function