' ## Calculations based on Black-Scholes assumptions
Option Explicit


' ## INTERMEDIATE CALCULATIONS
Public Function Calc_BS_ATMStrike(dbl_Fwd As Double, dbl_ATMVolPct As Double, dbl_TimeToMat As Double, _
    Optional bln_PID As Boolean = False) As Double
    ' ## Calculate at-the-money strike under the model assumptions
    Dim dbl_AdjFactor As Double: dbl_AdjFactor = Math.Exp(0.5 * (dbl_ATMVolPct / 100) ^ 2 * dbl_TimeToMat)

    If bln_PID = True Then
        Calc_BS_ATMStrike = dbl_Fwd / dbl_AdjFactor
    Else
        Calc_BS_ATMStrike = dbl_Fwd * dbl_AdjFactor
    End If
End Function

Public Function Calc_BS_Strike(dbl_TargetDelta As Double, enu_Direction As OptionDirection, dbl_Fwd As Double, _
    dbl_TimeToMat As Double, dbl_VolPct As Double, Optional bln_PID As Boolean = False) As Double
    ' ## Get strike corresponding to a given delta/vol combination

    If bln_PID = True Then
        Dim dbl_DevFactor As Double: dbl_DevFactor = Exp(5 * dbl_VolPct / 100 * Sqr(dbl_TimeToMat))
        Dim dbl_X1 As Double: dbl_X1 = dbl_Fwd * dbl_DevFactor
        Dim dbl_X2 As Double: dbl_X2 = dbl_Fwd / dbl_DevFactor
        Dim dic_StaticParams As Dictionary: Set dic_StaticParams = New Dictionary
        Call dic_StaticParams.Add("dbl_Fwd", dbl_Fwd)
        Call dic_StaticParams.Add("dbl_TimeToMat", dbl_TimeToMat)
        Call dic_StaticParams.Add("dbl_VolPct", dbl_VolPct)
        Call dic_StaticParams.Add("enu_Direction", enu_Direction)
        Call dic_StaticParams.Add("bln_PID", bln_PID)

        Dim dic_SolverOutputs As Dictionary: Set dic_SolverOutputs = New Dictionary
        Calc_BS_Strike = Solve_BrentDekker(ThisWorkbook, "SolverFuncXY_StrikeToDelta", dic_StaticParams, dbl_X1, _
            dbl_X2, dbl_TargetDelta, 0.000000005 * dbl_Fwd, 0, 1000, -1, dic_SolverOutputs)
    Else
        Select Case enu_Direction
            Case OptionDirection.CallOpt
                Calc_BS_Strike = Calc_BS_StrikeFromDelta(dbl_TargetDelta, dbl_Fwd, dbl_TimeToMat, dbl_VolPct)
            Case OptionDirection.PutOpt
                Calc_BS_Strike = Calc_BS_StrikeFromDelta(100 - dbl_TargetDelta, dbl_Fwd, dbl_TimeToMat, dbl_VolPct)
        End Select
    End If
End Function

Public Function Calc_BS_StrikeFromDelta(dbl_TargetNd1 As Double, dbl_Fwd As Double, dbl_TimeToMat As Double, dbl_VolPct As Double) As Double
    ' ## Target driftless call delta ~ N(d1) between 0 and 100
    Dim dbl_Output As Double, dbl_D1 As Double

    Select Case dbl_TargetNd1
        Case Is <= 0: dbl_Output = 1E+20
        Case Is >= 100: dbl_Output = 0
        Case Else
            dbl_D1 = WorksheetFunction.NormSInv(dbl_TargetNd1 / 100)
            dbl_Output = Calc_BS_StrikeFromD1(dbl_D1, dbl_Fwd, dbl_TimeToMat, dbl_VolPct)
    End Select

    Calc_BS_StrikeFromDelta = dbl_Output
End Function

Public Function Calc_BS_StrikeFromD1(dbl_D1 As Double, dbl_Fwd As Double, dbl_TimeToMat As Double, dbl_VolPct As Double) As Double
    Calc_BS_StrikeFromD1 = dbl_Fwd * Exp(0.5 * (dbl_VolPct / 100) ^ 2 * dbl_TimeToMat - dbl_D1 * Sqr(dbl_TimeToMat) * dbl_VolPct / 100)
End Function

Public Function Calc_BS_RemovePID(enu_Direction As OptionDirection, dbl_FwdDelta_PID As Double, dbl_Fwd As Double, _
    dbl_TimeToMat As Double, dbl_VolPct As Double) As Double
    ' ## Convert a premium included delta to a non-premium included delta
    ' ## Input delta must be a forward delta

    Dim dbl_ActiveStrike As Double
    dbl_ActiveStrike = Calc_BS_Strike(dbl_FwdDelta_PID, enu_Direction, dbl_Fwd, dbl_TimeToMat, dbl_VolPct, True)
    Calc_BS_RemovePID = Calc_BS_FwdDelta(enu_Direction, dbl_Fwd, dbl_ActiveStrike, dbl_TimeToMat, dbl_VolPct, False)
End Function

Public Function Calc_BS_CallToPutDelta(dbl_FwdCallDelta As Double, dbl_Fwd As Double, dbl_TimeToMat As Double, _
    dbl_VolPct As Double, dbl_ATMVolPct As Double, Optional bln_PID As Boolean = False) As Double
    ' ## For the PID case, solve for the strike based on the known vol and delta, then convert call to put delta
    ' ## For the non-PID case, conversion is trivial

    If bln_PID = False Then
        Calc_BS_CallToPutDelta = 100 - dbl_FwdCallDelta
    Else
        ' Derive values required by the inputs
        Dim dbl_ATMDelta As Double: dbl_ATMDelta = 50 * Math.Exp(-0.5 * (dbl_ATMVolPct / 100) ^ 2 * dbl_TimeToMat)
        Dim dbl_ATMStrike As Double: dbl_ATMStrike = dbl_ATMDelta * dbl_Fwd / 50
        Dim dbl_X1 As Double: dbl_X1 = dbl_Fwd * Exp(5 * dbl_VolPct / 100 * Sqr(dbl_TimeToMat))
        Dim dbl_X2 As Double: dbl_X2 = dbl_Fwd * Exp(-5 * dbl_VolPct / 100 * Sqr(dbl_TimeToMat))

        ' Prepare static parameters
        Dim dic_StaticParams As Dictionary: Set dic_StaticParams = New Dictionary
        Call dic_StaticParams.Add("dbl_Fwd", dbl_Fwd)
        Call dic_StaticParams.Add("dbl_TimeToMat", dbl_TimeToMat)
        Call dic_StaticParams.Add("dbl_VolPct", dbl_VolPct)
        Call dic_StaticParams.Add("enu_Direction", OptionDirection.CallOpt)
        Call dic_StaticParams.Add("bln_PID", bln_PID)

         ' Find range of strikes known to straddle the solution
        Dim dbl_Guess_LowerDelta As Double, dbl_Guess_HigherDelta As Double
        If dbl_FwdCallDelta <= dbl_ATMDelta Then
            ' Out of the money calls have higher strike than ATM
            dbl_Guess_LowerDelta = dbl_X1
            dbl_Guess_HigherDelta = dbl_ATMStrike
        Else
            ' In the money calls have lower strike than ATM
            dbl_Guess_LowerDelta = dbl_ATMStrike
            dbl_Guess_HigherDelta = dbl_X2
        End If

        ' Solve for strike
        Dim dbl_Strike As Double
        Dim dic_SolverOutputs As Dictionary: Set dic_SolverOutputs = New Dictionary
        dbl_Strike = Solve_BrentDekker(ThisWorkbook, "SolverFuncXY_StrikeToDelta", dic_StaticParams, dbl_Guess_LowerDelta, _
            dbl_Guess_HigherDelta, dbl_FwdCallDelta, 0.000000005 * dbl_Fwd, 0, 1000, -1, dic_SolverOutputs)

        ' Try widening the range (this should rarely if ever be required)
        If dic_SolverOutputs("Solvable") = False Then
            Set dic_SolverOutputs = New Dictionary
            dbl_Strike = Solve_BrentDekker(ThisWorkbook, "SolverFuncXY_StrikeToDelta", dic_StaticParams, dbl_X1, _
                dbl_X2, dbl_FwdCallDelta, 0.000000005 * dbl_Fwd, 0, 1000, -1, dic_SolverOutputs)
        End If

        ' Convert call to put delta
        Calc_BS_CallToPutDelta = 100 * dbl_Strike / dbl_Fwd - dbl_FwdCallDelta
    End If
End Function

Public Function Calc_BS_D1(dbl_Fwd As Double, dbl_Strike As Double, dbl_TimeToMat As Double, dbl_VolPct As Double) As Double
    Dim dbl_VolSqrtT As Double: dbl_VolSqrtT = dbl_VolPct / 100 * Sqr(dbl_TimeToMat)
    Calc_BS_D1 = (Log(dbl_Fwd / dbl_Strike) + 0.5 * dbl_VolSqrtT ^ 2) / dbl_VolSqrtT
End Function


' ## GREEKS
Public Function Calc_BS_FwdDelta(enu_Direction As OptionDirection, dbl_Fwd As Double, dbl_Strike As Double, dbl_TimeToMat As Double, _
    dbl_VolPct As Double, Optional bln_PID As Boolean = False) As Double
    ' ## Get absolute value of delta, expressed as a percentage.  The delta can be call/put, and can optionally be PID
    ' Calculate call delta
    Dim dbl_D1 As Double: dbl_D1 = Calc_BS_D1(dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct)

    Dim dbl_CallDelta As Double
    If bln_PID = True Then
        dbl_CallDelta = Calc_NormalCDF(dbl_D1 - dbl_VolPct / 100 * Sqr(dbl_TimeToMat)) * dbl_Strike / dbl_Fwd * 100

        ' Convert to put delta if necessary
        Select Case enu_Direction
            Case OptionDirection.PutOpt
                    Dim dbl_CallStrike As Double: dbl_CallStrike = Calc_BS_StrikeFromD1(dbl_D1, dbl_Fwd, dbl_TimeToMat, dbl_VolPct)
                    Calc_BS_FwdDelta = 100 * dbl_CallStrike / dbl_Fwd - dbl_CallDelta
            Case OptionDirection.CallOpt
                Calc_BS_FwdDelta = dbl_CallDelta
        End Select
    Else
        dbl_CallDelta = Calc_NormalCDF(dbl_D1) * 100

        ' Convert to put delta if necessary
        Select Case enu_Direction
            Case OptionDirection.PutOpt: Calc_BS_FwdDelta = 100 - dbl_CallDelta
            Case OptionDirection.CallOpt: Calc_BS_FwdDelta = dbl_CallDelta
        End Select
    End If
End Function

Public Function Calc_BS_Vega_Vanilla(dbl_Fwd As Double, dbl_Strike As Double, dbl_TimeToMat As Double, dbl_VolPct As Double) As Double
    ' ## Calculate undiscounted BS vega of a vanilla option
    Dim dbl_Output As Double
    If dbl_TimeToMat > 0 Then
        Dim dbl_D1 As Double: dbl_D1 = Calc_BS_D1(dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct)
        dbl_Output = dbl_Fwd * Sqr(dbl_TimeToMat) * Calc_NormalPDF(dbl_D1)
    Else
        dbl_Output = 0
    End If

    Calc_BS_Vega_Vanilla = dbl_Output
End Function

Public Function Calc_BS_Vanna_Vanilla(dbl_Spot As Double, dbl_Fwd As Double, dbl_Strike As Double, dbl_TimeToMat As Double, dbl_VolPct As Double) As Double
    ' ## Calculate undiscounted BS vanna of a standard vanilla option
    Dim dbl_Output As Double
    If dbl_TimeToMat > 0 Then
        Dim dbl_VolSqrtT As Double: dbl_VolSqrtT = dbl_VolPct / 100 * Sqr(dbl_TimeToMat)
        Dim dbl_D1 As Double: dbl_D1 = Calc_BS_D1(dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct)
        dbl_Output = Calc_BS_Vega_Vanilla(dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct) / dbl_Spot * (1 - dbl_D1 / dbl_VolSqrtT)
    Else
        dbl_Output = 0
    End If

    Calc_BS_Vanna_Vanilla = dbl_Output
End Function

Public Function Calc_BS_Vanna_Digital(enu_Direction As OptionDirection, dbl_Spot As Double, dbl_Fwd As Double, _
    dbl_Strike As Double, dbl_TimeToMat As Double, dbl_VolPct As Double, bln_IsPayoutDom As Boolean) As Double
    ' ## Calculate undiscounted BS vanna of a digital option
    Dim dbl_VolSqrtT As Double: dbl_VolSqrtT = dbl_VolPct / 100 * Sqr(dbl_TimeToMat)
    Dim dbl_D1 As Double: dbl_D1 = Calc_BS_D1(dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct)
    Dim dbl_D2 As Double: dbl_D2 = dbl_D1 - dbl_VolSqrtT
    Dim dbl_DigitalMoney As Double: dbl_DigitalMoney = -enu_Direction * Calc_NormalPDF(dbl_D2) * (1 - dbl_D1 * dbl_D2) / (dbl_Spot * dbl_VolSqrtT * dbl_VolPct / 100) * dbl_Strike

    If bln_IsPayoutDom = True Then
        Calc_BS_Vanna_Digital = dbl_DigitalMoney
    Else
        Calc_BS_Vanna_Digital = dbl_DigitalMoney - enu_Direction * Calc_BS_Vanna_Vanilla(dbl_Spot, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct)
    End If
End Function

Public Function Calc_BS_Vanna_DigitalCS(enu_Direction As OptionDirection, dbl_Spot As Double, dbl_Fwd As Double, _
    dbl_Strike As Double, dbl_TimeToMat As Double, dbl_VolPct_Orig As Double, dbl_VolPct_Shifted As Double, _
    bln_IsPayoutDom As Boolean, Optional dbl_RelSpread As Double = 0.0001) As Double
    ' ## Undiscounted vanna of digital option using call spread replication
    ' ## Result is per unit of FGN currency, even if the payout is DOM
    ' Determine cash-or-nothing digital vanna
    Dim dbl_ShiftAmt As Double: dbl_ShiftAmt = dbl_Strike * dbl_RelSpread
    Dim dbl_Orig As Double: dbl_Orig = Calc_BS_Vanna_Vanilla(dbl_Spot, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_Orig)
    Dim dbl_Orig_Shifted As Double: dbl_Orig_Shifted = Calc_BS_Vanna_Vanilla(dbl_Spot, dbl_Fwd, dbl_Strike + dbl_ShiftAmt, dbl_TimeToMat, dbl_VolPct_Shifted)
    Dim dbl_DigitalMoney As Double: dbl_DigitalMoney = enu_Direction * (dbl_Orig - dbl_Orig_Shifted) / dbl_ShiftAmt * dbl_Strike

    ' Output final price
    If bln_IsPayoutDom = True Then
        Calc_BS_Vanna_DigitalCS = dbl_DigitalMoney
    Else
        Calc_BS_Vanna_DigitalCS = dbl_DigitalMoney + enu_Direction * Calc_BS_Vanna_Vanilla(dbl_Spot, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct_Orig)
    End If
End Function

Public Function Calc_BS_Gamma_Vanilla(dbl_Spot As Double, dbl_Fwd As Double, dbl_Strike As Double, _
    dbl_TimeToMat As Double, dbl_VolPct As Double) As Double
    ' ## Calculate undiscounted BS gamma of a vanilla option, expressed in foreign currency per unit of foreign currency notional
    Dim dbl_Output As Double
    If dbl_TimeToMat > 0 Then
        Dim dbl_D1 As Double: dbl_D1 = Calc_BS_D1(dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct)
        Dim dbl_VolSqrtT As Double: dbl_VolSqrtT = dbl_VolPct / 100 * Sqr(dbl_TimeToMat)
        dbl_Output = dbl_Fwd / dbl_Spot * Calc_NormalPDF(dbl_D1) / (dbl_Spot * dbl_VolSqrtT)
    Else
        dbl_Output = 0
    End If

    Calc_BS_Gamma_Vanilla = dbl_Output
End Function

Public Function Calc_BS_Gamma_Digital(enu_Direction As OptionDirection, dbl_Spot As Double, dbl_Fwd As Double, dbl_Strike As Double, _
    dbl_TimeToMat As Double, dbl_VolPct As Double, bln_IsPayoutDom As Boolean) As Double
    ' ## Calculate undiscounted BS gamma of a digital option, per unit of foreign notional
    Dim dbl_D1 As Double: dbl_D1 = Calc_BS_D1(dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct)
    Dim dbl_VolSqrtT As Double: dbl_VolSqrtT = dbl_VolPct / 100 * Sqr(dbl_TimeToMat)
    Dim dbl_D2 As Double: dbl_D2 = dbl_D1 - dbl_VolSqrtT
    Dim dbl_FullPayout As Double: If bln_IsPayoutDom = True Then dbl_FullPayout = 1 Else dbl_FullPayout = dbl_Fwd

    Calc_BS_Gamma_Digital = enu_Direction * dbl_FullPayout * Calc_NormalPDF(dbl_D1) * dbl_D2 / (dbl_Spot * dbl_VolSqrtT ^ 2)
End Function

Public Function Calc_BS_AdaptedDelta(enu_Direction As OptionDirection, dbl_Spot As Double, dbl_Fwd As Double, dbl_Strike As Double, _
    dbl_TimeToMat As Double, dbl_VolPct As Double, dbl_SmileSlope As Double, dbl_OptionDF As Double, Optional bln_PID As Boolean = False) As Double
    ' ## Calculate spot delta of a standard vanilla option, taking into account the shape of the smile
    Dim dbl_Delta As Double: dbl_Delta = Calc_BS_FwdDelta(enu_Direction, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct, bln_PID) * dbl_OptionDF / 100
    Dim dbl_Vega As Double: dbl_Vega = Calc_BS_Vega_Vanilla(dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct) * dbl_OptionDF
    Dim dbl_Gamma As Double: dbl_Gamma = Calc_BS_Gamma_Vanilla(dbl_Spot, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct) * dbl_OptionDF
    Dim dbl_Vanna As Double:  dbl_Vanna = Calc_BS_Vanna_Vanilla(dbl_Spot, dbl_Fwd, dbl_Strike, dbl_TimeToMat, dbl_VolPct) * dbl_OptionDF

    Calc_BS_AdaptedDelta = dbl_Delta + dbl_Vega * dbl_Gamma * dbl_SmileSlope / (1 - dbl_SmileSlope * dbl_Vanna)
End Function