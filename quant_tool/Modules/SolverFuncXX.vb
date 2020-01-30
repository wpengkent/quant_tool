Option Explicit

Public Function SolverFuncXX_FXSmileIteration(dbl_VolPct As Double, dic_Params As Dictionary, dic_Outputs As Dictionary) As Double
    ' ## Find the delta of the option with the specified vol then output the vol corresponding to the delta
    Dim lng_LookupDate As Long: lng_LookupDate = dic_Params("lng_LookupDate")
    Dim dbl_IterPutDelta As Double
    dbl_IterPutDelta = Calc_BS_FwdDelta(OptionDirection.PutOpt, dic_Params("dbl_LookupFwd"), dic_Params("dbl_Strike"), dic_Params("dbl_TimeToMat_Lookup"), _
        dbl_VolPct, dic_Params("bln_PID_Interp"))

    SolverFuncXX_FXSmileIteration = LookupFXSmile(dbl_IterPutDelta, dic_Params("str_Interp_Delta"), dic_Params)
End Function