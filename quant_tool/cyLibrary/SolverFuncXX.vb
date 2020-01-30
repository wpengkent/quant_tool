Option Explicit
' ## Functions are suitable for solver methods: FixedPt

Public Function SolverFuncXX_BAWThreshold(dbl_Threshold As Double, dic_Params As Dictionary, dic_SecondaryOutputs As Dictionary) As Double
    ' ## Find the level of the spot for which exercise is first considered optimal

    Dim dbl_Output As Double
    Dim dbl_EuropeanPrice As Double, dbl_ActiveD1 As Double, dbl_ActiveNd1 As Double
    Dim dbl_Strike As Double: dbl_Strike = dic_Params("dbl_Strike")
    Dim enu_Direction As OptionDirection: enu_Direction = dic_Params("enu_Direction")

    dbl_EuropeanPrice = Calc_BSPrice_Vanilla(enu_Direction, dbl_Threshold * dic_Params("dbl_FwdSpotRatio"), _
        dbl_Strike, dic_Params("dbl_TimeToMat"), dic_Params("dbl_VolPct")) * dic_Params("dbl_DomDF")
    dbl_ActiveD1 = (Math.Log(dbl_Threshold / dbl_Strike) + dic_Params("dbl_TotalDrift_BN")) / dic_Params("dbl_VolSqrT")
    dbl_ActiveNd1 = Calc_NormalCDF(dbl_ActiveD1 * enu_Direction)

    dbl_Output = dbl_Strike + enu_Direction * dbl_EuropeanPrice + (1 - dic_Params("dbl_FgnDF") * dbl_ActiveNd1) * dbl_Threshold / dic_Params("dbl_Q")

    If dbl_Output <= 0 Then dbl_Output = 0.00000000000001  ' Rational bounds on spot

    ' Return outputs
    SolverFuncXX_BAWThreshold = dbl_Output
    If dic_SecondaryOutputs.Exists("dbl_ActiveNd1") = True Then Call dic_SecondaryOutputs.Remove("dbl_ActiveNd1")
    Call dic_SecondaryOutputs.Add("dbl_ActiveNd1", dbl_ActiveNd1)
End Function