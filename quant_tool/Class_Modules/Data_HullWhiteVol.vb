    'Outstanding:
'1. Small difference in USD
'2. Handle of specific scenario: Handle holiday at start and end date !!!! Usage of Gen Start
'3. Initial guess
'4. Deal applicable strike
'5. Mapping of actual trade to the hull white calibration inputs
'6. Memory issue
'TRUE: ATM strike at exercise date for calibration
'FALSE: Deal applicable strike at exercise date for calibration
Const bln_AtmStrikeCal As Boolean = True
'Handle calibration failure due to negative vol or exceed max iterations
Const dbl_CalFailVal As Double = 0.1

' ## MEMBER DATA
Const dbl_Vol_MYR As Double = 0.3 / 100
Const dbl_r_MYR As Double = 3.8 / 100

Const dbl_Vol_USD As Double = 0.6 / 100 '0.6/100
Const dbl_r_USD As Double = 3 / 100 '2/100

Const str_MYR As String = "MYR"
Const str_USD As String = "USD"

Const int_DealAppStrikeDelta As Integer = 2

'applicable for bermudan swaption only currently
'if it is set to true, the assumption is if previous deal applicable strike = strike, deal applicable strike checking is no longer required for the rest of the pillars
Const bln_SpeedUpDAS As Boolean = False

'Handle bootstrap failure using Murex: TRUE = 10% of prev max bootstrapped sigma. FALSE = dbl_SmallValue i.e. very small value
Const bln_HandleClbFailureMXWay = False
Const dbl_SmallValue = 0.0001

' Components
Private Swaption_active As Inst_IRSwaption
Private dic_OriVol As Dictionary
Private dic_FinalVol As Dictionary

' Curve dependencies

' Dynamic variables
Private lng_ValDate As Long

' Static values
Private dic_GlobalStaticInfo As Dictionary, dic_CurveSet As Dictionary

Public Sub Initialize(dbl_MR As Double, str_CCY As String, str_Gen_FixLeg As String, str_Gen_FltLeg As String, col_ExerciseDate As Collection, col_SwapStart As Collection, lng_SwapMat As Long, col_Strike As Collection, str_Svl As String, Optional dbl_notional As Double = 100000000)

Set dic_GlobalStaticInfo = GetStaticInfo()

Set dic_CurveSet = GetAllCurves(True, False, dic_GlobalStaticInfo)
Call FillAllDependencies(dic_CurveSet)

Dim int_num As Integer
int_num = col_ExerciseDate.count
Dim int_i As Integer

Dim fld_Output As InstParams_SWT
Dim dic_OriVolOutput As Dictionary: Set dic_OriVolOutput = New Dictionary
Dim dic_FinalVolOutput As Dictionary: Set dic_FinalVolOutput = New Dictionary
Dim Col_CalibrateOutput As Collection

Dim dbl_r As Double
Dim dbl_Vol As Double

If str_CCY = str_MYR Then
    dbl_r = dbl_r_MYR
    dbl_Vol = dbl_Vol_MYR
ElseIf str_CCY = str_USD Then
    dbl_r = dbl_r_USD
    dbl_Vol = dbl_Vol_USD

End If

Dim dbl_FinalVol As Double
Dim dbl_T1 As Double
Dim dbl_T2 As Double

lng_ValDate = cyGetValDate()

Set Swaption_active = New Inst_IRSwaption

Dim dbl_ATM_strike As Double
Dim dbl_ATM_vol As Double
Dim dbl_Lower_Strike As Double
Dim dbl_Upper_Strike As Double
Dim dbl_Strike As Double
Dim dbl_adj As Double

Dim dbl_PrevStrike As Double
dbl_PrevStrike = 0

Dim dbl_PrevMaxSigma As Double
dbl_PrevMaxSigma = 0

Dim dbl_NewV2OriSigma As Double
Dim lng_DateKey As Long
Dim bln_check As Boolean


For int_i = 1 To int_num

    If dbl_PrevStrike = col_Strike(int_i) And bln_SpeedUpDAS = True Then
        dbl_Strike = col_Strike(int_i)
    Else
        fld_Output = GetInstParams(str_CCY, str_Gen_FixLeg, str_Gen_FltLeg, col_ExerciseDate(int_i), col_SwapStart(int_i), lng_SwapMat, col_Strike(int_i), str_Svl, True, False, True)
        Call Swaption_active.Initialize(fld_Output, dic_CurveSet, dic_GlobalStaticInfo)

        dbl_ATM_strike = Swaption_active.SwapRate

        If bln_AtmStrikeCal = False Then

            fld_Output = GetInstParams(str_CCY, str_Gen_FixLeg, str_Gen_FltLeg, col_ExerciseDate(int_i), col_SwapStart(int_i), lng_SwapMat, dbl_ATM_strike, str_Svl, True, False, True)
            Call Swaption_active.Initialize(fld_Output, dic_CurveSet, dic_GlobalStaticInfo)

            dbl_ATM_vol = Swaption_active.Volatility / 100

            dbl_adj = StrikeAdj(int_DealAppStrikeDelta, dbl_ATM_vol, col_ExerciseDate(int_i), lng_ValDate)
            dbl_Lower_Strike = dbl_ATM_strike / dbl_adj
            dbl_Upper_Strike = dbl_ATM_strike * dbl_adj

            If col_Strike(int_i) <= dbl_Lower_Strike Then
                dbl_Strike = dbl_Lower_Strike
            ElseIf col_Strike(int_i) >= dbl_Upper_Strike Then
                dbl_Strike = dbl_Upper_Strike
            Else
                dbl_Strike = col_Strike(int_i)
            End If
        End If
    End If
    'Debug.Print dbl_Strike
    If bln_AtmStrikeCal = True Then
        dbl_Strike = dbl_ATM_strike
    End If

    dbl_PrevStrike = dbl_Strike

    fld_Output = GetInstParams(str_CCY, str_Gen_FixLeg, str_Gen_FltLeg, col_ExerciseDate(int_i), col_SwapStart(int_i), lng_SwapMat, dbl_Strike, str_Svl)
    Call Swaption_active.Initialize(fld_Output, dic_CurveSet, dic_GlobalStaticInfo)

    Set Col_CalibrateOutput = New Collection
    Set Col_CalibrateOutput = Swaption_active.HW_CalibrateVol(dbl_MR, dbl_r, dbl_Vol)

    If Col_CalibrateOutput(2) = 9999 Then
        Call dic_OriVolOutput.Add(col_ExerciseDate(int_i), dbl_CalFailVal)
    Else
        Call dic_OriVolOutput.Add(col_ExerciseDate(int_i), Col_CalibrateOutput(2))
    End If


    If dic_OriVolOutput(col_ExerciseDate(int_i)) = dbl_CalFailVal Then
        Call dic_FinalVolOutput.Add(col_ExerciseDate(int_i), dbl_CalFailVal)

    ElseIf int_i = 1 Then

    'If int_i = 1 Then
        Call dic_FinalVolOutput.Add(col_ExerciseDate(1), dic_OriVolOutput(dic_OriVolOutput.Keys(0)))
    Else
        dbl_T1 = (col_ExerciseDate(int_i - 1) - lng_ValDate) / 365
        dbl_T2 = (col_ExerciseDate(int_i) - lng_ValDate) / 365
        dbl_FinalVol = BootstrapHWVol(dbl_MR, dbl_T1, dbl_T2, dic_OriVolOutput(dic_OriVolOutput.Keys(int_i - 2)), dic_OriVolOutput(dic_OriVolOutput.Keys(int_i - 1)), dbl_PrevMaxSigma, bln_HandleClbFailureMXWay, dbl_SmallValue)

        bln_check = BootstrapHWVolFail(dbl_MR, dbl_T1, dbl_T2, dic_OriVolOutput(dic_OriVolOutput.Keys(int_i - 2)), dic_OriVolOutput(dic_OriVolOutput.Keys(int_i - 1)))

        If bln_check = True Then
            lng_DateKey = dic_OriVolOutput.Keys(int_i - 1)
            Call dic_OriVolOutput.Remove(lng_DateKey)
            dbl_NewV2OriSigma = BootstrapHWOriSigma(dbl_MR, dbl_T1, dbl_T2, dic_OriVolOutput(dic_OriVolOutput.Keys(int_i - 2)), dbl_FinalVol)
            Call dic_OriVolOutput.Add(lng_DateKey, dbl_NewV2OriSigma)
        End If

        Call dic_FinalVolOutput.Add(col_ExerciseDate(int_i), dbl_FinalVol)

    End If

    If dbl_PrevMaxSigma < dic_FinalVolOutput(col_ExerciseDate(int_i)) Then
        dbl_PrevMaxSigma = dic_FinalVolOutput(col_ExerciseDate(int_i))
    End If


Next int_i

'Hardcode Sigma
'Set dic_FinalVolOutput = New Dictionary
'
'For int_i = 1 To 18
'    Call dic_FinalVolOutput.Add(col_ExerciseDate(int_i), Sheets("RngAcc_Output_MYR KLI Q 6M").Range("C2").Offset(int_i - 1, 0).Value / 100)
'Next int_i


Set dic_OriVol = New Dictionary
Set dic_OriVol = dic_OriVolOutput

Set dic_FinalVol = New Dictionary
Set dic_FinalVol = dic_FinalVolOutput

'Debug.Print dic_OriVolOutput(dic_OriVolOutput.Keys(0))
'Debug.Print dic_OriVolOutput(dic_OriVolOutput.Keys(1))
'Debug.Print dic_OriVolOutput(dic_OriVolOutput.Keys(2))

'Debug.Print dic_FinalVolOutput(dic_FinalVolOutput.Keys(0))
'Debug.Print dic_FinalVolOutput(dic_FinalVolOutput.Keys(1))
'Debug.Print dic_FinalVolOutput(dic_FinalVolOutput.Keys(2))

End Sub
Public Property Get FullOriHwVol() As Dictionary
    Set FullOriHwVol = dic_OriVol
End Property
Public Property Get FullFinalHwVol() As Dictionary
    Set FullFinalHwVol = dic_FinalVol
End Property

Public Property Get FinalHwVol(lng_date As Long) As Double

Dim dbl_Output As Double
Dim int_i As Integer
Dim int_count As Integer

If dic_FinalVol.Exists(lng_date) = True Then
    dbl_Output = dic_FinalVol(lng_date)
Else
    int_count = dic_FinalVol.count
    For int_i = 1 To int_count
        If lng_date >= dic_FinalVol.Keys(int_count - 1) Then
            dbl_Output = dic_FinalVol(dic_FinalVol.Keys(int_count - 1))
            Exit For
        ElseIf lng_date <= dic_FinalVol.Keys(int_i - 1) Then
            dbl_Output = dic_FinalVol(dic_FinalVol.Keys(int_i - 1))
            Exit For
        End If
    Next int_i
End If

FinalHwVol = dbl_Output

End Property

Private Function GetInstParams(str_CCY As String, str_Gen_FixLeg As String, str_Gen_FltLeg As String, lng_OptMat As Long, _
    lng_SwapStart As Long, lng_SwapMat As Long, dbl_Strike As Double, str_Svl As String, Optional bln_USDStubInterp As Boolean = False, Optional bln_FreqChange As Boolean = True, _
    Optional bln_EstFlt As Boolean = False, Optional dbl_notional As Double = 100000000) As InstParams_SWT

Dim fld_Output As InstParams_SWT

Set dic_StaticInfoInput = GetStaticInfo()

    ' Gather generator level information
    Dim igs_Generators As IRGeneratorSet
    Set igs_Generators = New IRGeneratorSet
    Set igs_Generators = dic_StaticInfoInput(StaticInfoType.IRGeneratorSet)

    Dim cfg_Settings As ConfigSheet
    Set cfg_Settings = New ConfigSheet
    Set cfg_Settings = dic_StaticInfoInput(StaticInfoType.ConfigSheet)

    Dim fld_LegA As IRLegParams: fld_LegA = igs_Generators.Lookup_Generator(str_Gen_FixLeg)
    Dim fld_LegB As IRLegParams: fld_LegB = igs_Generators.Lookup_Generator(str_Gen_FltLeg)
    fld_LegA.ForceToMV = True
    fld_LegB.ForceToMV = True

    fld_LegA.IsUniformPeriods = False
    fld_LegB.IsUniformPeriods = False

    fld_Output.ValueDate = cfg_Settings.CurrentValDate

    If FreqValue(fld_LegA.PmtFreq) < FreqValue(fld_LegB.PmtFreq) And bln_FreqChange = True Then
        fld_LegB.PmtFreq = fld_LegA.PmtFreq
        fld_LegB.index = fld_LegA.PmtFreq
    End If

    fld_Output.OptionMat = lng_OptMat
    fld_LegA.Swapstart = lng_SwapStart
    fld_LegB.Swapstart = lng_SwapStart
    fld_LegA.ValueDate = lng_SwapStart
    fld_LegB.ValueDate = lng_SwapStart

    'fld_LegA.Term = "3M" 'not used
    'fld_LegB.Term = "3M" 'not used
    fld_LegA.IsFwdGeneration = False
    fld_LegB.IsFwdGeneration = False


    'Instead of using swap maturity as generation reference point, may need to consider using generation start if there is issue
    fld_LegA.GenerationRefPoint = lng_SwapMat
    fld_LegB.GenerationRefPoint = lng_SwapMat
    fld_LegA.GenerationLimitPoint = lng_SwapStart
    fld_LegB.GenerationLimitPoint = lng_SwapStart

    fld_Output.BuySell = "B"
    fld_Output.Pay_LegA = True
    fld_Output.Exercise = "European"
    fld_LegA.PExch_Start = False
    fld_LegB.PExch_Start = False
    fld_LegA.PExch_Intermediate = False
    fld_LegB.PExch_Intermediate = False
    fld_LegA.PExch_End = False
    fld_LegB.PExch_End = False
    'fld_LegA.FloatEst = False
    'fld_LegB.FloatEst = False

    fld_LegA.FloatEst = bln_EstFlt
    fld_LegB.FloatEst = bln_EstFlt

    If bln_USDStubInterp = True And str_CCY = str_USD Then
        fld_LegB.StubInterpolate = True
    End If

    fld_Output.CCY_PnL = str_CCY
    fld_Output.IsSmile = True
    fld_Output.VolCurve = str_Svl

    ' Leg A
    With fld_LegA
        .Notional = dbl_notional
        .RateOrMargin = dbl_Strike
    End With

    ' Leg B
    With fld_LegB
        .Notional = dbl_notional
        .RateOrMargin = 0
    End With

    fld_Output.LegA = fld_LegA
    fld_Output.LegB = fld_LegB
    GetInstParams = fld_Output

End Function


Private Function FreqValue(str_freq As String) As Integer

Dim int_output As Integer

If str_freq = "3M" Then
    int_output = 3
ElseIf str_freq = "6M" Then
    int_output = 6
ElseIf str_freq = "12M" Or str_freq = "1Y" Then
    int_output = 12
End If

FreqValue = int_output

End Function

Private Function StrikeAdj(int_Delta As Integer, dbl_Vol As Double, lng_ExeDate As Long, lng_ValuationDate As Long) As Double

StrikeAdj = Exp(int_Delta * dbl_Vol * Sqr((lng_ExeDate - lng_ValuationDate) / 365))

End Function