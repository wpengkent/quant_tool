Option Explicit
' ## Functions are suitable for solver methods: Secant, Brent-Dekker

Public Function SolverFuncXY_StrikeToDelta(dbl_Strike As Double, dic_Params As Dictionary, dic_SecondaryOutputs As Dictionary) As Double
    ' ## Return the convention delta corresponding to the specified strike
    SolverFuncXY_StrikeToDelta = Calc_BS_FwdDelta(dic_Params("enu_Direction"), dic_Params("dbl_Fwd"), dbl_Strike, dic_Params("dbl_TimeToMat"), _
        dic_Params("dbl_VolPct"), dic_Params("bln_PID"))
End Function