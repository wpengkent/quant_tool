Public Function HW_B(dbl_MeanRev As Double, dbl_T1 As Double, dbl_T2 As Double) As Double

    HW_B = (1 - Exp(-dbl_MeanRev * (dbl_T2 - dbl_T1))) / dbl_MeanRev

End Function

Public Function HW_B_Disc(dbl_MeanRev As Double, dbl_T1 As Double, dbl_T2 As Double, dbl_DeltaT As Double) As Double

Dim dbl_B_Numerator As Double
    dbl_B_Numerator = HW_B(dbl_MeanRev, dbl_T1, dbl_T2)

Dim dbl_B_Denominator As Double
    dbl_B_Denominator = HW_B(dbl_MeanRev, dbl_T1, dbl_T1 + dbl_DeltaT)

HW_B_Disc = dbl_B_Numerator / dbl_B_Denominator * dbl_DeltaT

End Function

Public Function HW_A(dbl_MeanRev As Double, dbl_T1 As Double, dbl_T2 As Double, dbl_Vol As Double, _
                        dbl_P0T1 As Double, dbl_P0T2 As Double, dbl_F0T1 As Double) As Double


Dim dbl_B As Double
    dbl_B = HW_B(dbl_MeanRev, dbl_T1, dbl_T2)


HW_A = (dbl_P0T2 / dbl_P0T1) * Exp(dbl_B * dbl_F0T1 - (dbl_B ^ 2) * (dbl_Vol ^ 2) * (1 - Exp(-2 * dbl_MeanRev * dbl_T1)) / (4 * dbl_MeanRev))

End Function


Public Function HW_SigmaP(dbl_Vol As Double, dbl_MeanRev As Double, dbl_OptMat As Double, dbl_UndStart As Double, dbl_UndMat As Double) As Double

'HW_SigmaP = dbl_Vol * HW_B(dbl_MeanRev, dbl_UndStart, dbl_UndMat) * Sqr(HW_B(2 * dbl_MeanRev, 0, dbl_OptMat))

HW_SigmaP = dbl_Vol * Sqr(HW_B(2 * dbl_MeanRev, 0, dbl_OptMat)) * _
            (Exp(-dbl_MeanRev * (dbl_UndStart - dbl_OptMat)) - Exp(-dbl_MeanRev * (dbl_UndMat - dbl_OptMat))) / dbl_MeanRev

End Function

Public Function HW_h(dbl_Vol As Double, dbl_MeanRev As Double, dbl_OptMat As Double, dbl_UndStart As Double, dbl_UndMat As Double, _
                dbl_L As Double, dbl_K As Double, dbl_P0T As Double, dbl_P0s As Double) As Double

Dim dbl_SigmaP As Double
dbl_SigmaP = HW_SigmaP(dbl_Vol, dbl_MeanRev, dbl_OptMat, dbl_UndStart, dbl_UndMat)

    HW_h = (1 / dbl_SigmaP) * Log((dbl_L * dbl_P0s) / (dbl_K * dbl_P0T)) + dbl_SigmaP / 2

End Function

Public Function HW_ZcCall(dbl_Vol As Double, dbl_MeanRev As Double, dbl_OptMat As Double, dbl_UndStart As Double, dbl_UndMat As Double, _
                dbl_L As Double, dbl_K As Double, dbl_P0T As Double, dbl_P0s As Double) As Double

Dim dbl_SigmaP As Double
dbl_SigmaP = HW_SigmaP(dbl_Vol, dbl_MeanRev, dbl_OptMat, dbl_UndStart, dbl_UndMat)

Dim dbl_h As Double
'dbl_h = HW_h(dbl_Vol, dbl_MeanRev, dbl_OptMat, dbl_UndStart, dbl_UndMat, dbl_L, dbl_K, dbl_P0T, dbl_P0s)
dbl_h = (1 / dbl_SigmaP) * Log((dbl_L * dbl_P0s) / (dbl_K * dbl_P0T)) + dbl_SigmaP / 2


HW_ZcCall = dbl_L * dbl_P0s * WorksheetFunction.NormSDist(dbl_h) - dbl_K * dbl_P0T * WorksheetFunction.NormSDist(dbl_h - dbl_SigmaP)

End Function

Public Function HW_ZcPut(dbl_Vol As Double, dbl_MeanRev As Double, dbl_OptMat As Double, dbl_UndStart As Double, dbl_UndMat As Double, _
                dbl_L As Double, dbl_K As Double, dbl_P0T As Double, dbl_P0s As Double) As Double

Dim dbl_SigmaP As Double
dbl_SigmaP = HW_SigmaP(dbl_Vol, dbl_MeanRev, dbl_OptMat, dbl_UndStart, dbl_UndMat)


Dim dbl_h As Double
'dbl_h = HW_h(dbl_Vol, dbl_MeanRev, dbl_OptMat, dbl_UndStart, dbl_UndMat, dbl_L, dbl_K, dbl_P0T, dbl_P0s)
dbl_h = (1 / dbl_SigmaP) * Log((dbl_L * dbl_P0s) / (dbl_K * dbl_P0T)) + dbl_SigmaP / 2


HW_ZcPut = dbl_K * dbl_P0T * WorksheetFunction.NormSDist(-dbl_h + dbl_SigmaP) - _
            dbl_L * dbl_P0s * WorksheetFunction.NormSDist(-dbl_h)

End Function

Public Function BootstrapHWVol(dbl_MR As Double, dbl_T1 As Double, dbl_T2 As Double, dbl_vol_1 As Double, dbl_vol_2 As Double, dbl_PrevMaxSigma As Double, Optional bln_HandleClbFailureMXWay As Boolean = False, Optional dbl_FailSigma As Double = 0.0001) As Double

Dim dbl_V1 As Double
Dim dbl_V2 As Double

Dim dbl_EXP_1 As Double
Dim dbl_EXP_2 As Double

dbl_EXP_1 = Exp(2 * dbl_MR * dbl_T1)
dbl_EXP_2 = Exp(2 * dbl_MR * dbl_T2)

dbl_V1 = (dbl_vol_1 ^ 2) * (dbl_EXP_1 - 1) / (2 * dbl_MR)
dbl_V2 = (dbl_vol_2 ^ 2) * (dbl_EXP_2 - 1) / (2 * dbl_MR)

If (dbl_V2 - dbl_V1) < 0 Then
    If bln_HandleClbFailureMXWay = True Then
        BootstrapHWVol = dbl_PrevMaxSigma * 0.1
    Else
        BootstrapHWVol = dbl_FailSigma
    End If
Else
    BootstrapHWVol = Sqr(2 * dbl_MR / (dbl_EXP_2 - dbl_EXP_1) * (dbl_V2 - dbl_V1))
End If

End Function

Public Function BootstrapHWVolFail(dbl_MR As Double, dbl_T1 As Double, dbl_T2 As Double, dbl_vol_1 As Double, dbl_vol_2 As Double) As Boolean

Dim dbl_V1 As Double
Dim dbl_V2 As Double

Dim dbl_EXP_1 As Double
Dim dbl_EXP_2 As Double

dbl_EXP_1 = Exp(2 * dbl_MR * dbl_T1)
dbl_EXP_2 = Exp(2 * dbl_MR * dbl_T2)

dbl_V1 = (dbl_vol_1 ^ 2) * (dbl_EXP_1 - 1) / (2 * dbl_MR)
dbl_V2 = (dbl_vol_2 ^ 2) * (dbl_EXP_2 - 1) / (2 * dbl_MR)

If (dbl_V2 - dbl_V1) < 0 Then
    BootstrapHWVolFail = True
Else
    BootstrapHWVolFail = False
End If

End Function
Public Function BootstrapHWOriSigma(dbl_MR As Double, dbl_T1 As Double, dbl_T2 As Double, dbl_vol_1 As Double, dbl_FinalSigma As Double) As Double

Dim dbl_V1 As Double
Dim dbl_V2 As Double

Dim dbl_EXP_1 As Double
Dim dbl_EXP_2 As Double

dbl_EXP_1 = Exp(2 * dbl_MR * dbl_T1)
dbl_EXP_2 = Exp(2 * dbl_MR * dbl_T2)

dbl_V1 = (dbl_vol_1 ^ 2) * (dbl_EXP_1 - 1) / (2 * dbl_MR)

BootstrapHWOriSigma = Sqr((((dbl_FinalSigma) ^ 2) * (dbl_EXP_2 - dbl_EXP_1) / (2 * dbl_MR) + dbl_V1) * (2 * dbl_MR) / (dbl_EXP_2 - 1))


End Function