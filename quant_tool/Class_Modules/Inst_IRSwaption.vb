Option Explicit
'----------------------------------------------------------------------------------
'Important Notes (For Bermudan Swaption Enhancement
' Assuming Leg A and Leg B is based on same Disc Curve and Est Curve
'i.e. irc_Disc_LegA & irc_Est_LegA
'Leg A must be fixed/Leg B must be float

'----------------------------------------------------------------------------------
' ## MEMBER DATA

' WL 20181002: Hull White Calibration
Const dbl_rShockSize As Double = 0.01
Const dbl_VolShockSize As Double = 0.01
Const dbl_MaxIter As Long = 100
Const dbl_Tolerance As Double = 0.000001

' Components
Private irl_LegA As IRLeg, irl_legB As IRLeg, scf_Premium As SCF
Private irl_LegA_Orig As IRLeg, irl_LegB_Orig As IRLeg

' Curve dependencies
Private fxs_Spots As Data_FXSpots, svc_VolCurve As Data_SwptVols, irc_Disc_LegA As Data_IRCurve

' Dynamic variables
Private lng_ValDate As Long, dbl_VolShift As Double

' Static values
Private dic_GlobalStaticInfo As Dictionary, dic_CurveDependencies As Dictionary
Private fld_Params As InstParams_SWT
Private enu_Direction As OptionDirection, int_Sign As Integer, dbl_Strike As Double
Private enu_VolType As CurveType

'KL - PDE Implementation
'START
Private HWPDE As PDE_Matrix
Private int_SpotStep As Integer
Private int_TimeStep As Integer
Private dbl_MeanRev As Double
Private irc_Est_LegB As Data_IRCurve
Private arr_PDE_UndVal() As Variant
Private arr_PDE_Option() As Variant
Private HW_vol As Data_HullWhiteVol
'END

' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As InstParams_SWT, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing)

    ' Store static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    fld_Params = fld_ParamsInput
    dbl_Strike = fld_Params.LegA.RateOrMargin

    ' Initialize dynamic variables
    Call Me.SetValDate(fld_Params.ValueDate)
    dbl_VolShift = 0

    ' Store dependent curves
    If dic_CurveSet Is Nothing Then
        Set irc_Disc_LegA = GetObject_IRCurve(fld_Params.LegA.Curve_Disc, True, False)
        Set svc_VolCurve = GetObject_SwptVols(fld_Params.VolCurve, True, False)
        Set fxs_Spots = GetObject_FXSpots(True)
    Else
        Set irc_Disc_LegA = dic_CurveSet(CurveType.IRC)(fld_Params.LegA.Curve_Disc)
        Set svc_VolCurve = dic_CurveSet(CurveType.SVL)(fld_Params.VolCurve)
        Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
    End If

    ' Set up underlying IR legs
    Set irl_LegA = New IRLeg
    Call irl_LegA.Initialize(fld_Params.LegA, dic_CurveSet, dic_GlobalStaticInfo)

    Set irl_legB = New IRLeg
    Call irl_legB.Initialize(fld_Params.LegB, dic_CurveSet, dic_GlobalStaticInfo)

    ' Handle case where margin is non-zero by storing a zero margin instrument with adjusted strike having the same MV
    If irl_legB.Params.RateOrMargin <> 0 Then
        Set irl_LegA_Orig = New IRLeg
        Call irl_LegA_Orig.Initialize(fld_Params.LegA, dic_CurveSet, dic_GlobalStaticInfo)

        Set irl_LegB_Orig = New IRLeg
        Call irl_LegB_Orig.Initialize(fld_Params.LegB, dic_CurveSet, dic_GlobalStaticInfo)

        Call irl_legB.SetRateOrMargin(0)

        ' Set adjusted strike, which is the leg A fixed rate
        Call CalibrateAdjustedStrike
    End If

    ' Set up premium
    Set scf_Premium = New SCF
    Call scf_Premium.Initialize(fld_Params.Premium, dic_CurveSet, dic_GlobalStaticInfo)
    scf_Premium.ZShiftsEnabled_DF = GetSetting_IsPremInDV01()
    scf_Premium.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

    ' Store calculated values
    If fld_Params.Pay_LegA = True Then
        enu_Direction = OptionDirection.CallOpt
    Else
        enu_Direction = OptionDirection.PutOpt
    End If

    Select Case UCase(fld_Params.BuySell)
        Case "B", "BUY": int_Sign = 1
        Case "S", "SELL": int_Sign = -1
        Case Else: int_Sign = 0
    End Select

    ''''''''''KL - PDE Implementation''''''''
    '''''''''''''''''START'''''''''''''''''''

    If UCase(fld_Params.Exercise) = "BERMUDAN" Then

        Set irc_Est_LegB = dic_CurveSet(CurveType.IRC)(fld_Params.LegB.Curve_Est)
        int_SpotStep = fld_Params.SpotStep
        int_TimeStep = fld_Params.TimeStep
        dbl_MeanRev = fld_Params.MeanRev / 100

        'Calibration
        Dim col_ExDate As Collection: Set col_ExDate = New Collection
        Dim col_SwapStart As Collection: Set col_SwapStart = New Collection
        Dim col_Strike As Collection: Set col_Strike = New Collection

        'Add Call Date into Collection
        Dim k As Variant
        For Each k In fld_Params.CallDate.Keys
            col_ExDate.Add (k)
        Next k

        For Each k In fld_Params.Swapstart.Keys
            col_SwapStart.Add (k)
        Next k

        For Each k In fld_Params.Swapstart.Keys
            col_Strike.Add (dbl_Strike)
        Next k

        Set HW_vol = New Data_HullWhiteVol
        Call HW_vol.Initialize(dbl_MeanRev, fld_Params.LegA.CCY, fld_Params.GeneratorA, fld_Params.GeneratorB, col_ExDate, col_SwapStart, irl_LegA.MatDate, col_Strike, fld_Params.VolCurve)
        Dim dic As New Dictionary
        Set dic = HW_vol.FullFinalHwVol

        'Hull White PDE
        Set HWPDE = New PDE_Matrix
        Call HWPDE.Initialize(fld_Params.ValueDate, irl_LegA.MatDate, HW_vol.FullFinalHwVol, irl_LegA, irl_legB, int_SpotStep, int_TimeStep, dbl_MeanRev, BSWAP)

    End If
    '''''''''''''''''END'''''''''''''''''''''

    ' Determine curve dependencies
    Set dic_CurveDependencies = scf_Premium.CurveDependencies
    Set dic_CurveDependencies = Convert_MergeDicts(dic_CurveDependencies, fxs_Spots.Lookup_CurveDependencies(fld_Params.CCY_PnL))
    Set dic_CurveDependencies = Convert_MergeDicts(dic_CurveDependencies, irl_LegA.CurveDependencies)
    Set dic_CurveDependencies = Convert_MergeDicts(dic_CurveDependencies, irl_legB.CurveDependencies)
    If dic_CurveDependencies.Exists(irc_Disc_LegA.CurveName) = False Then Call dic_CurveDependencies.Add(irc_Disc_LegA.CurveName, True)
End Sub

Private Sub HW_Option()

Dim bln_LastUndCF As Boolean: bln_LastUndCF = True
Dim int_LastUndCF As Integer: int_LastUndCF = 0

Dim int_i As Integer
Dim int_j As Integer
Dim int_cnt1 As Integer

Dim dbl_T1 As Double
Dim dbl_T2 As Double
Dim dbl_PT1 As Double
Dim dbl_PT2 As Double
Dim dbl_A As Double

Dim arr_PDE_Matrix() As Variant
Dim arr_PDE_Temp() As Variant
Dim arr_PDE_Temp_Undval() As Variant

Dim arr_output() As Variant
Dim arr_UndVal() As Variant
Dim arr_UndVal_A() As Variant
Dim arr_UndVal_B() As Variant

Dim arr_SpotStep() As Double
Dim arr_TimeStep() As Double
Dim arr_TimeLabel() As Integer
Dim arr_TimeStep_ImpDate() As Double

arr_SpotStep() = HWPDE.SpotStep
arr_TimeStep() = HWPDE.TimeStep
arr_TimeLabel() = HWPDE.TimeLabel
arr_TimeStep_ImpDate() = HWPDE.TimeStep_ImpDate

ReDim arr_PDE_Temp(1 To UBound(arr_SpotStep))
ReDim arr_PDE_Temp_Undval(1 To UBound(arr_SpotStep))
ReDim arr_output(1 To UBound(arr_TimeStep), 1 To UBound(arr_SpotStep))
ReDim arr_UndVal(1 To UBound(arr_TimeStep), 1 To UBound(arr_SpotStep))

Dim x As Double
Dim y As Double
Dim sum_y As Double


arr_UndVal_A() = HW_Underlying(irl_LegA)
arr_UndVal_B() = HW_Underlying(irl_legB)

'Combining cash flow from LegA & LegB
For int_i = UBound(arr_TimeStep) To 1 Step -1
    For int_j = 1 To UBound(arr_SpotStep)
        arr_UndVal(int_i, int_j) = enu_Direction * (-arr_UndVal_A(int_i, int_j) + arr_UndVal_B(int_i, int_j))
    Next int_j
Next int_i

arr_PDE_UndVal = arr_UndVal

'Calculate option value
'First Cash Flow

'Smart Multiphication
If (arr_TimeLabel(UBound(arr_TimeStep)) = 1 Or arr_TimeLabel(UBound(arr_TimeStep)) = 2) Then
    dbl_T1 = arr_TimeStep_ImpDate(UBound(arr_TimeStep_ImpDate) - int_cnt1 - 1)
    dbl_T2 = arr_TimeStep_ImpDate(UBound(arr_TimeStep_ImpDate) - int_cnt1)
    dbl_PT1 = irc_Disc_LegA.Lookup_Rate(lng_ValDate, lng_ValDate + dbl_T1 * 365, "DF", , , True)
    dbl_PT2 = irc_Disc_LegA.Lookup_Rate(lng_ValDate, lng_ValDate + dbl_T2 * 365, "DF", , , True)
    dbl_A = HWPDE.HWBond_A(dbl_PT1, dbl_PT2, dbl_T1, dbl_T2)

    If int_cnt1 < UBound(arr_TimeStep_ImpDate) Then int_cnt1 = int_cnt1 + 1
Else
    dbl_A = 1
End If

If arr_TimeLabel(UBound(arr_TimeStep)) = 1 Then
    For int_j = 1 To UBound(arr_SpotStep)

        'Store Temp-Underlying
        arr_PDE_Temp_Undval(int_j) = arr_UndVal(UBound(arr_TimeStep), int_j)

        'Store Opt-Underlying
        If arr_UndVal(UBound(arr_TimeStep), int_j) < 0 Then
            arr_PDE_Temp(int_j) = 0
        Else
            arr_PDE_Temp(int_j) = arr_UndVal(UBound(arr_TimeStep), int_j) * dbl_A
            arr_PDE_Temp_Undval(int_j) = arr_UndVal(UBound(arr_TimeStep), int_j)
        End If
    Next int_j

    'Smoothing
    If int_LastUndCF > 0 Then bln_LastUndCF = False
    arr_PDE_Temp = HWPDE.PayoffSmooth_Independent(arr_PDE_Temp_Undval, arr_PDE_Temp, arr_SpotStep(2) - arr_SpotStep(1), bln_LastUndCF, True)
    int_LastUndCF = int_LastUndCF + 1
Else
    For int_j = 1 To UBound(arr_SpotStep)
        arr_PDE_Temp(int_j) = 0
        arr_PDE_Temp_Undval(int_j) = 0
    Next int_j
End If

For int_j = 1 To UBound(arr_SpotStep)
    arr_output(UBound(arr_TimeStep), int_j) = arr_PDE_Temp(int_j)
Next int_j

'Subsequent Cash Flow
For int_i = UBound(arr_TimeStep) - 1 To 1 Step -1

    'Diffuse Option Value
    For int_j = 1 To UBound(arr_SpotStep)
        arr_PDE_Temp_Undval(int_j) = arr_UndVal(int_i, int_j)
        arr_PDE_Temp(int_j) = arr_output(int_i + 1, int_j)
    Next int_j

    dbl_T1 = arr_TimeStep(int_i)
    dbl_T2 = arr_TimeStep(int_i + 1)

    arr_PDE_Temp = HWPDE.PDE_Matrix(arr_PDE_Temp, dbl_T1, dbl_T2, arr_TimeLabel(int_i + 1))

    'Smart Multiphication
    If (arr_TimeLabel(int_i) = 1 Or arr_TimeLabel(int_i) = 2) Then
        dbl_T1 = arr_TimeStep_ImpDate(UBound(arr_TimeStep_ImpDate) - int_cnt1 - 1)
        dbl_T2 = arr_TimeStep_ImpDate(UBound(arr_TimeStep_ImpDate) - int_cnt1)
        dbl_PT1 = irc_Disc_LegA.Lookup_Rate(lng_ValDate, lng_ValDate + dbl_T1 * 365, "DF", , , True)
        dbl_PT2 = irc_Disc_LegA.Lookup_Rate(lng_ValDate, lng_ValDate + dbl_T2 * 365, "DF", , , True)
        dbl_A = HWPDE.HWBond_A(dbl_PT1, dbl_PT2, dbl_T1, dbl_T2)

        If int_cnt1 < UBound(arr_TimeStep_ImpDate) Then int_cnt1 = int_cnt1 + 1
    Else
        dbl_A = 1
    End If

    For int_j = 1 To UBound(arr_SpotStep)
       arr_PDE_Temp(int_j) = arr_PDE_Temp(int_j) * dbl_A
    Next int_j

    'Smoothing
    If arr_TimeLabel(int_i) = 1 Then
        If int_LastUndCF > 0 Then bln_LastUndCF = False
        arr_PDE_Temp = HWPDE.PayoffSmooth_Independent(arr_PDE_Temp_Undval, arr_PDE_Temp, arr_SpotStep(2) - arr_SpotStep(1), bln_LastUndCF, True)
        int_LastUndCF = int_LastUndCF + 1
    End If

    'Store in Output
    For int_j = 1 To UBound(arr_SpotStep)
       arr_output(int_i, int_j) = arr_PDE_Temp(int_j)
    Next int_j

Next int_i

arr_PDE_Option = arr_output

End Sub
Private Function HW_Underlying(irl_Leg As IRLeg) As Variant

Dim int_cnt As Integer
Dim int_cnt1 As Integer
Dim int_i As Integer
Dim int_j As Integer

Dim dbl_T1 As Double
Dim dbl_T2 As Double
Dim dbl_PT1 As Double
Dim dbl_PT2 As Double
Dim dbl_A As Double
Dim dbl_B As Double
Dim dbl_X As Double
Dim dbl_HWBond As Double
Dim dbl_CalcPeriod As Double

Dim dbl_T1_F As Double
Dim dbl_T2_F As Double
Dim dbl_PT1_F As Double
Dim dbl_PT2_F As Double
Dim dbl_A_F As Double
Dim dbl_B_F As Double
Dim dbl_X_F As Double
Dim dbl_HWBond_F As Double
Dim dbl_Forward As Double
Dim dbl_CalcPeriod_F As Double

Dim dbl_Payoff As Double

Dim arr_SpotStep() As Double
Dim arr_TimeStep() As Double
Dim arr_TimeLabel() As Integer
Dim arr_TimeStep_ImpDate() As Double
Dim arr_PDE_Matrix() As Variant
Dim arr_PDE_Temp() As Variant
Dim arr_output As Variant

Dim arr_payoff() As Double
Dim arr_Payoff_Time() As Double

Dim col_PeriodStart As New Collection
Dim col_PeriodEnd As New Collection
Dim col_EstStart As New Collection
Dim col_EstEnd As New Collection

arr_SpotStep() = HWPDE.SpotStep
arr_TimeStep() = HWPDE.TimeStep
arr_TimeLabel() = HWPDE.TimeLabel
arr_TimeStep_ImpDate() = HWPDE.TimeStep_ImpDate

Set col_PeriodStart = irl_Leg.PeriodStart
Set col_PeriodEnd = irl_Leg.PeriodEnd

If irl_Leg.IsFixed = False Then
    Set col_EstStart = irl_Leg.EstStart
    Set col_EstEnd = irl_Leg.EstEnd
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Calculate Underlying Value based on Hull White Model'
''''''''''''''''''''''''''''''''''''''''''''''''''''''
For int_cnt = col_PeriodStart.count To 1 Step -1

    If col_PeriodStart(int_cnt) < lng_ValDate Then
        GoTo SkipPayoff
    Else
        int_i = int_i + 1
        ReDim Preserve arr_payoff(1 To UBound(arr_SpotStep), 1 To int_i) As Double
        ReDim Preserve arr_Payoff_Time(1 To int_i) As Double

        For int_j = 1 To UBound(arr_SpotStep)

            dbl_T1 = calc_yearfrac(lng_ValDate, col_PeriodStart(int_cnt), "ACT/365")
            dbl_T2 = calc_yearfrac(lng_ValDate, col_PeriodEnd(int_cnt), "ACT/365")
            dbl_PT1 = irc_Disc_LegA.Lookup_Rate(lng_ValDate, col_PeriodStart(int_cnt), "DF", , , True)
            dbl_PT2 = irc_Disc_LegA.Lookup_Rate(lng_ValDate, col_PeriodEnd(int_cnt), "DF", , , True)
            dbl_A = HWPDE.HWBond_A(dbl_PT1, dbl_PT2, dbl_T1, dbl_T2)
            dbl_B = HW_B(dbl_MeanRev, dbl_T1, dbl_T2)
            dbl_X = arr_SpotStep(int_j)
            dbl_HWBond = dbl_A * Exp(-dbl_B * dbl_X)
            dbl_CalcPeriod = calc_yearfrac(col_PeriodStart(int_cnt), col_PeriodEnd(int_cnt), irl_Leg.Params.Daycount)

            Select Case irl_Leg.IsFixed

            Case True:
                 dbl_Payoff = dbl_Strike / 100 * irl_Leg.Params.Notional * dbl_HWBond * dbl_CalcPeriod
                 arr_payoff(int_j, int_i) = dbl_Payoff
                 arr_Payoff_Time(int_i) = dbl_T1

            Case False:
                 dbl_T1_F = calc_yearfrac(lng_ValDate, col_EstStart(int_cnt), "ACT/365")
                 dbl_T2_F = calc_yearfrac(lng_ValDate, col_EstEnd(int_cnt), "ACT/365")
                 dbl_PT1_F = irc_Disc_LegA.Lookup_Rate(lng_ValDate, col_EstStart(int_cnt), "DF", , , True)
                 dbl_PT2_F = irc_Disc_LegA.Lookup_Rate(lng_ValDate, col_EstEnd(int_cnt), "DF", , , True)

                 dbl_A_F = HWPDE.HWBond_A(dbl_PT1_F, dbl_PT2_F, dbl_T1_F, dbl_T2_F)
                 dbl_B_F = HW_B(dbl_MeanRev, dbl_T1_F, dbl_T2_F)
                 dbl_HWBond_F = dbl_A_F * Exp(-dbl_B_F * dbl_X)
                 dbl_CalcPeriod_F = calc_yearfrac(col_EstStart(int_cnt), col_EstEnd(int_cnt), irl_Leg.Params.Daycount)

                 Dim dbl_spread As Double
                 Dim dbl_spread_est As Double
                 Dim dbl_spread_disc As Double

                 dbl_spread_disc = irc_Disc_LegA.Lookup_Rate(lng_ValDate, col_EstStart(int_cnt), "DF", , , True) / irc_Disc_LegA.Lookup_Rate(lng_ValDate, col_EstEnd(int_cnt), "DF", , , True)
                 dbl_spread_disc = (dbl_spread_disc - 1) / dbl_CalcPeriod_F

                 dbl_spread_est = irc_Est_LegB.Lookup_Rate(lng_ValDate, col_EstStart(int_cnt), "DF", , , True) / irc_Est_LegB.Lookup_Rate(lng_ValDate, col_EstEnd(int_cnt), "DF", , , True)
                 dbl_spread_est = (dbl_spread_est - 1) / dbl_CalcPeriod_F

                 dbl_spread = (dbl_spread_est - dbl_spread_disc)

                 dbl_Forward = (1 / dbl_HWBond_F - 1) / dbl_CalcPeriod_F + dbl_spread
                 dbl_Payoff = (dbl_Forward + irl_Leg.Params.RateOrMargin / 100) * irl_Leg.Params.Notional * dbl_HWBond * dbl_CalcPeriod

                 arr_payoff(int_j, int_i) = dbl_Payoff
                 arr_Payoff_Time(int_i) = dbl_T1_F
            End Select

        Next int_j

    End If

SkipPayoff:

Next int_cnt

'''''''''''''''''''''''''
''''Diffusion Process''''
'''''''''''''''''''''''''
ReDim arr_output(1 To UBound(arr_TimeStep), 1 To UBound(arr_SpotStep)) As Variant
ReDim arr_PDE_Temp(1 To UBound(arr_SpotStep)) As Variant

'First Cashflow
int_cnt = 1
If (arr_TimeStep(UBound(arr_TimeStep)) = arr_Payoff_Time(int_cnt)) Then

    dbl_T1 = arr_TimeStep_ImpDate(UBound(arr_TimeStep_ImpDate) - 1)
    dbl_T2 = arr_TimeStep_ImpDate(UBound(arr_TimeStep_ImpDate))
    dbl_PT1 = irc_Disc_LegA.Lookup_Rate(lng_ValDate, lng_ValDate + dbl_T1 * 365, "DF", , , True)
    dbl_PT2 = irc_Disc_LegA.Lookup_Rate(lng_ValDate, lng_ValDate + dbl_T2 * 365, "DF", , , True)
    dbl_A = HWPDE.HWBond_A(dbl_PT1, dbl_PT2, dbl_T1, dbl_T2)

    For int_j = 1 To UBound(arr_SpotStep)
        arr_PDE_Temp(int_j) = arr_payoff(int_j, int_cnt)
        arr_output(UBound(arr_TimeStep), int_j) = arr_PDE_Temp(int_j) * dbl_A
    Next int_j
    int_cnt = int_cnt + 1
Else
    For int_j = 1 To UBound(arr_SpotStep)
        arr_PDE_Temp(int_j) = 0
        arr_output(UBound(arr_TimeStep), int_j) = 0
    Next int_j
End If

'Subsequent Cashflow
int_cnt1 = 1
For int_i = UBound(arr_TimeStep) - 1 To 1 Step -1

    'Normal Diffusion

    dbl_T1 = arr_TimeStep(int_i)
    dbl_T2 = arr_TimeStep(int_i + 1)
'    dbl_PT1 = irc_Disc_LegA.Lookup_Rate(lng_ValDate, lng_ValDate + dbl_T1 * 365, "DF", , , True)
'    dbl_PT2 = irc_Disc_LegA.Lookup_Rate(lng_ValDate, lng_ValDate + dbl_T2 * 365, "DF", , , True)
'    dbl_A = HWPDE.HWBond_A(dbl_PT1, dbl_PT2, dbl_T1, dbl_T2)

    For int_j = 1 To UBound(arr_SpotStep)
        arr_PDE_Temp(int_j) = arr_output(int_i + 1, int_j)
    Next int_j

     arr_PDE_Temp = HWPDE.PDE_Matrix(arr_PDE_Temp, dbl_T1, dbl_T2, arr_TimeLabel(int_i + 1))

    'Smart Multiplication
    If (arr_TimeLabel(int_i) = 1 Or arr_TimeLabel(int_i) = 2) Then
        dbl_T1 = arr_TimeStep_ImpDate(UBound(arr_TimeStep_ImpDate) - int_cnt1 - 1)
        dbl_T2 = arr_TimeStep_ImpDate(UBound(arr_TimeStep_ImpDate) - int_cnt1)
        dbl_PT1 = irc_Disc_LegA.Lookup_Rate(lng_ValDate, lng_ValDate + dbl_T1 * 365, "DF", , , True)
        dbl_PT2 = irc_Disc_LegA.Lookup_Rate(lng_ValDate, lng_ValDate + dbl_T2 * 365, "DF", , , True)
        dbl_A = HWPDE.HWBond_A(dbl_PT1, dbl_PT2, dbl_T1, dbl_T2)

        If int_cnt1 < UBound(arr_TimeStep_ImpDate) Then int_cnt1 = int_cnt1 + 1
    Else
        dbl_A = 1
    End If

    'Add In Payoff
    If (arr_TimeStep(int_i) = arr_Payoff_Time(int_cnt)) Then
        For int_j = 1 To UBound(arr_SpotStep)
            arr_output(int_i, int_j) = (arr_PDE_Temp(int_j) + arr_payoff(int_j, int_cnt)) * dbl_A
        Next int_j

        If int_cnt < UBound(arr_Payoff_Time) Then int_cnt = int_cnt + 1
    Else
        For int_j = 1 To UBound(arr_SpotStep)
            arr_output(int_i, int_j) = arr_PDE_Temp(int_j) * dbl_A
        Next int_j
    End If

Next int_i

HW_Underlying = arr_output

End Function
' ## PROPERTIES
Public Property Get marketvalue() As Double
    ' ## Get discounted option value
    Dim dbl_Output As Double
    Dim dbl_TimeToMat As Double: dbl_TimeToMat = calc_yearfrac(lng_ValDate, fld_Params.OptionMat, "ACT/365")
    Dim dbl_DFToOptMat_LegA As Double
    dbl_DFToOptMat_LegA = irc_Disc_LegA.Lookup_Rate(lng_ValDate, fld_Params.LegA.ValueDate, "DF", , , True)

    Select Case UCase(fld_Params.Exercise)
        Case "EUROPEAN"
            ' Case for option expired is handled within the BS pricing function
            dbl_Output = Calc_BSPrice_Vanilla(enu_Direction, SwapRate(), dbl_Strike, dbl_TimeToMat, _
                Volatility(), , NormalCDFMethod.Abram) / 100 * fld_Params.LegA.Notional * irl_LegA.SwaptionScalingFactor * dbl_DFToOptMat_LegA
        Case "BERMUDAN"
            Call HW_Option
            dbl_Output = arr_PDE_Option(1, fld_Params.SpotStep + 1)
    End Select

    marketvalue = dbl_Output * GetFXConvFactor() * int_Sign
End Property

Public Property Get Cash() As Double
    ' ## Get value of premium in PnL currency
    Cash = -scf_Premium.CalcValue(lng_ValDate, lng_ValDate, fld_Params.CCY_PnL) * int_Sign
End Property

Public Property Get PnL() As Double
    PnL = Me.marketvalue + Me.Cash
End Property

Public Property Get Volatility() As Double
    If fld_Params.IsSmile = True Then
        Volatility = svc_VolCurve.Lookup_Vol(fld_Params.OptionMat, irl_LegA.MatDate, dbl_Strike)
    Else
        Volatility = svc_VolCurve.Lookup_Vol(fld_Params.OptionMat, irl_LegA.MatDate)
    End If
End Property

Public Property Get SwapRate() As Double
    SwapRate = irl_LegA.SolveParRate(irl_legB)
End Property

Private Property Get GetFXConvFactor() As Double
    ' ## Get factor to convert from the native currency to the PnL reporting currency
    GetFXConvFactor = fxs_Spots.Lookup_DiscSpot(irl_LegA.Params.CCY, fld_Params.CCY_PnL)
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
        dbl_Val_Up = Me.PnL

        Call SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        dbl_Val_Down = Me.PnL

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

Public Function Calc_Vega(enu_Type As CurveType, str_curve As String) As Double
    ' ## Return sensitivity to a 1 vol increase in the vol
    Dim dbl_Output As Double

    If svc_VolCurve.TypeCode = enu_Type And fld_Params.VolCurve = str_curve Then
        ' Store shifted values
        Dim dbl_Val_Up As Double, dbl_Val_Down As Double
        svc_VolCurve.VolShift_Sens = 0.01
        dbl_Val_Up = Me.PnL

        svc_VolCurve.VolShift_Sens = -0.01
        dbl_Val_Down = Me.PnL

        ' Clear temporary shifts from the underlying leg
        svc_VolCurve.VolShift_Sens = 0

        ' Calculate by finite differencing and convert to PnL currency
        dbl_Output = (dbl_Val_Up - dbl_Val_Down) * 50
    Else
        dbl_Output = 0
    End If

    Calc_Vega = dbl_Output
End Function


' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Sub HandleUpdate_IRC(str_CurveName As String)
    ' ## Update stored values affected by change in the curve values
    Call irl_LegA.HandleUpdate_IRC(str_CurveName)
    Call irl_legB.HandleUpdate_IRC(str_CurveName)

    If Not irl_LegA_Orig Is Nothing Then
        Call irl_LegA_Orig.HandleUpdate_IRC(str_CurveName)
        Call irl_LegB_Orig.HandleUpdate_IRC(str_CurveName)
        Call CalibrateAdjustedStrike
    End If
End Sub

Public Sub SetValDate(lng_Input As Long)
    ' ## Set stored value date and refresh values dependent on the value date
    ' ## Don't change valuation date of underlying legs, because these are based on the maturity date
    lng_ValDate = lng_Input
    If lng_ValDate > fld_Params.LegA.ValueDate Then lng_ValDate = fld_Params.LegA.ValueDate  ' Prevent accumulation of expired deal
    If Not irl_LegA_Orig Is Nothing Then Call CalibrateAdjustedStrike
End Sub

Private Sub SetCurveState(str_curve As String, enu_State As CurveState_IRC, Optional int_PillarIndex As Integer = 0)
    ' ## Set up shift in the market data and underlying components
    If irc_Disc_LegA.CurveName = str_curve Then Call irc_Disc_LegA.SetCurveState(enu_State, int_PillarIndex)
    Call irl_LegA.SetCurveState(str_curve, enu_State, int_PillarIndex)
    Call irl_legB.SetCurveState(str_curve, enu_State, int_PillarIndex)
    Call scf_Premium.SetCurveState(str_curve, enu_State, int_PillarIndex)

    If Not irl_LegA_Orig Is Nothing Then
        Call irl_LegA_Orig.SetCurveState(str_curve, enu_State, int_PillarIndex)
        Call irl_LegB_Orig.SetCurveState(str_curve, enu_State, int_PillarIndex)
        Call CalibrateAdjustedStrike
    End If
End Sub

Private Sub CalibrateAdjustedStrike()
    ' Calculate adjusted strike, required initially if margin is non-zero.  Need to recalculate if the rates change
    dbl_Strike = irl_LegA.SolveParRate(irl_legB, False, irl_LegA_Orig.marketvalue, irl_LegB_Orig.marketvalue)
End Sub

' ## METHODS - CALCULATION DETAILS
Public Sub OutputReport(wks_output As Worksheet)
    wks_output.Cells.Clear

    Dim rng_ActiveOutput As Range: Set rng_ActiveOutput = wks_output.Range("A1")
    Dim int_ActiveRow As Integer: int_ActiveRow = 0
    Dim int_ActiveColumn As Integer: int_ActiveColumn = 0
    Dim str_Address_Notional As String
    Dim str_Address_NotionalCCY As String, str_Address_Position As String, str_Address_Payout As String, str_Address_Fwd As String
    Dim str_Address_Strike As String, str_Address_Vol As String, str_Address_MatDate As String
    Dim str_Address_OptionDF As String, str_Address_SwapStart As String, str_Address_SwapMat As String
    Dim str_Address_ScalingFactor As String, str_Address_MV As String, str_Address_Cash As String

    Dim str_CurrencyFormat As String: str_CurrencyFormat = Gather_CurrencyFormat()
    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    Dim rng_PnL As Range, rng_OptionDF As Range
    Dim int_ctr As Integer

    Dim dic_Addresses As Dictionary: Set dic_Addresses = New Dictionary
    dic_Addresses.CompareMode = CompareMethod.TextCompare

    ' Swaption pricing info

    Select Case UCase(fld_Params.Exercise)

    Case "EUROPEAN"
        With rng_ActiveOutput
            .Offset(int_ActiveRow, 0).Value = "OVERALL"
            .Offset(int_ActiveRow, 0).Font.Italic = True

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Value date:"
            .Offset(int_ActiveRow, 1).Value = lng_ValDate
            .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
            Call dic_Addresses.Add("ValDate", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "PnL:"
            .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
            Set rng_PnL = .Offset(int_ActiveRow, 1)

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Currency:"
            .Offset(int_ActiveRow, 1).Value = fld_Params.CCY_PnL
            Call dic_Addresses.Add("PnLCCY", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 2
            .Offset(int_ActiveRow, 0).Value = "OPTION LEG"
            .Offset(int_ActiveRow, 0).Font.Italic = True

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Notional:"
            .Offset(int_ActiveRow, 1).Value = fld_Params.LegA.Notional
            .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
            str_Address_Notional = .Offset(int_ActiveRow, 1).Address(False, False)

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Currency:"
            .Offset(int_ActiveRow, 1).Value = fld_Params.LegA.CCY
            str_Address_NotionalCCY = .Offset(int_ActiveRow, 1).Address(False, False)

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Direction:"
            If enu_Direction = CallOpt Then
                .Offset(int_ActiveRow, 1).Value = "Call"
            Else
                .Offset(int_ActiveRow, 1).Value = "Put"
            End If
            str_Address_Payout = .Offset(int_ActiveRow, 1).Address(False, False)

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Swap rate:"
            .Offset(int_ActiveRow, 1).Value = SwapRate()
            str_Address_Fwd = .Offset(int_ActiveRow, 1).Address(False, False)

            If irl_legB.Params.RateOrMargin <> 0 Then
                int_ActiveRow = int_ActiveRow + 1
                .Offset(int_ActiveRow, 0).Value = "Orig Strike:"
                .Offset(int_ActiveRow, 1).Value = irl_LegA_Orig.RateOrMargin
                str_Address_Strike = .Offset(int_ActiveRow, 1).Address(False, False)
            End If

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Adj Strike:"
            .Offset(int_ActiveRow, 1).Value = irl_LegA.RateOrMargin
            str_Address_Strike = .Offset(int_ActiveRow, 1).Address(False, False)

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Vol:"
            .Offset(int_ActiveRow, 1).Value = Volatility()
            str_Address_Vol = .Offset(int_ActiveRow, 1).Address(False, False)

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Option mat:"
            .Offset(int_ActiveRow, 1).Value = fld_Params.OptionMat
            .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
            str_Address_MatDate = .Offset(int_ActiveRow, 1).Address(False, False)

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Option DF:"
            Set rng_OptionDF = .Offset(int_ActiveRow, 1)
            str_Address_OptionDF = rng_OptionDF.Address(False, False)

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Swap start:"
            .Offset(int_ActiveRow, 1).Value = irl_LegA.Params.Swapstart
            .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
            str_Address_SwapStart = .Offset(int_ActiveRow, 1).Address(False, False)

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Swap mat:"
            .Offset(int_ActiveRow, 1).Value = irl_LegA.MatDate
            .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
            str_Address_SwapMat = .Offset(int_ActiveRow, 1).Address(False, False)

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Scaling factor:"
            .Offset(int_ActiveRow, 1).Value = irl_LegA.SwaptionScalingFactor
            str_Address_ScalingFactor = .Offset(int_ActiveRow, 1).Address(False, False)

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "MV (" & fld_Params.CCY_PnL & "):"
            .Offset(int_ActiveRow, 1).Formula = "=" & int_Sign & "*" & str_Address_Notional & "*Calc_BSPrice_Vanilla(IF(" & str_Address_Payout _
                & "=""Call"",1,-1)," & str_Address_Fwd & "," & str_Address_Strike & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," & str_Address_MatDate _
                & ",""ACT/365"")," & str_Address_Vol & ")/100*" & str_Address_OptionDF & "*" & str_Address_ScalingFactor & "*cyGetFXDiscSpot(" _
                & str_Address_NotionalCCY & "," & dic_Addresses("PnLCCY") & ")"
            .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
            .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
            str_Address_MV = .Offset(int_ActiveRow, 1).Address(False, False)

            ' Swap legs detail
            int_ActiveRow = int_ActiveRow + 3
            Call irl_LegA.OutputReport_Swap(.Offset(int_ActiveRow, 0), "Leg A", fld_Params.Pay_LegA, dic_Addresses("PnLCCY"), _
                dic_Addresses("ValDate"), fld_Params.CCY_PnL, "", "")
            Call irl_legB.OutputReport_Swap(.Offset(int_ActiveRow, 14), "Leg B", Not fld_Params.Pay_LegA, dic_Addresses("PnLCCY"), _
                dic_Addresses("ValDate"), fld_Params.CCY_PnL, "", "")

            ' Cost leg
            int_ActiveRow = 0
            .Offset(int_ActiveRow, 4).Value = "COST LEG"
            .Offset(int_ActiveRow, 4).Font.Italic = True

            int_ActiveRow = int_ActiveRow + 1
            Call scf_Premium.OutputReport(.Offset(int_ActiveRow, 4), "Cash", fld_Params.CCY_PnL, -int_Sign, _
                True, dic_Addresses, False)

            ' Calculated fields
            rng_OptionDF.Formula = "=cyReadIRCurve(""" & fld_Params.LegA.Curve_Disc & """," & dic_Addresses("ValDate") & "," _
                & str_Address_SwapStart & ",""DF""" & ")"
            rng_PnL.Formula = "=" & str_Address_MV & "+" & dic_Addresses("SCF_PV")
        End With

    Case "BERMUDAN"

        'KL - Add for Bermudan Swaption
        Dim int_i As Integer
        Dim int_j As Integer

        Dim arr_SpotStep() As Double
        Dim arr_TimeStep() As Double

        arr_SpotStep() = HWPDE.SpotStep
        arr_TimeStep() = HWPDE.TimeStep

        Call HW_Option
        'Populate Calibrated vol
        rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn).Value = "Calibrated Sigma"

        For int_i = 1 To HW_vol.FullFinalHwVol.count
            int_ActiveRow = int_ActiveRow + 1
            With rng_ActiveOutput
                .Offset(int_ActiveRow, int_ActiveColumn).Value = HW_vol.FullFinalHwVol.Keys(int_i - 1)
                '.Offset(int_ActiveRow, int_ActiveColumn).styl
                .Offset(int_ActiveRow, int_ActiveColumn + 1).Value = HW_vol.FullFinalHwVol.Items(int_i - 1)
            End With

        Next int_i

        int_ActiveRow = int_ActiveRow + 2
        rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn).Value = "Option Value"
        int_ActiveRow = int_ActiveRow + 2
        rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn).Value = "TimeStep/SpotStep"

        ''''For Option Value''''
        'Populate SpotStep
        int_ActiveRow = int_ActiveRow
        int_ActiveColumn = 1
        For int_j = 1 To UBound(arr_SpotStep)
            With rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn)
                .Value = arr_SpotStep(int_j)
                .NumberFormat = "0.00%"
                .Font.Italic = True
                .Font.Bold = True
            End With
            int_ActiveColumn = int_ActiveColumn + 1
        Next int_j

        'Populate TimeStep
        int_ActiveRow = int_ActiveRow + 1
        int_ActiveColumn = 0
        For int_i = UBound(arr_TimeStep) To 1 Step -1
            With rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn)
                .Value = arr_TimeStep(int_i)
                .NumberFormat = "0.0000"
                .Font.Italic = True
                .Font.Bold = True
            End With
            int_ActiveRow = int_ActiveRow + 1
        Next int_i

        'Populate Option Value
        int_ActiveRow = int_ActiveRow - UBound(arr_TimeStep)
        int_ActiveColumn = 1
        For int_i = UBound(arr_TimeStep) To 1 Step -1
            For int_j = 1 To UBound(arr_SpotStep)
                With rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn)
                    .Value = arr_PDE_Option(int_i, int_j)
                    .Style = "Comma"
                End With
                int_ActiveColumn = int_ActiveColumn + 1
            Next int_j
            int_ActiveRow = int_ActiveRow + 1
            int_ActiveColumn = int_ActiveColumn - UBound(arr_SpotStep)
        Next int_i

        ''''For Underlying Value''''
        int_ActiveRow = int_ActiveRow + 2
        int_ActiveColumn = 0
        rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn).Value = "Underlying Value"
        int_ActiveRow = int_ActiveRow + 2
        rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn).Value = "TimeStep/SpotStep"

        'Populate SpotStep
        int_ActiveRow = int_ActiveRow
        int_ActiveColumn = 1
        For int_j = 1 To UBound(arr_SpotStep)
            With rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn)
                .Value = arr_SpotStep(int_j)
                .NumberFormat = "0.00%"
                .Font.Italic = True
                .Font.Bold = True
            End With
            int_ActiveColumn = int_ActiveColumn + 1
        Next int_j

        'Populate TimeStep
        int_ActiveRow = int_ActiveRow + 1
        int_ActiveColumn = 0
        For int_i = UBound(arr_TimeStep) To 1 Step -1
            With rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn)
                .Value = arr_TimeStep(int_i)
                .NumberFormat = "0.0000"
                .Font.Italic = True
                .Font.Bold = True
            End With
            int_ActiveRow = int_ActiveRow + 1
        Next int_i

        'Populate UnderlyingValue Value
        int_ActiveRow = int_ActiveRow - UBound(arr_TimeStep)
        int_ActiveColumn = 1

        For int_i = UBound(arr_TimeStep) To 1 Step -1
            For int_j = 1 To UBound(arr_SpotStep)
                With rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn)
                    .Value = arr_PDE_UndVal(int_i, int_j)
                    .Style = "Comma"
                End With
                int_ActiveColumn = int_ActiveColumn + 1
            Next int_j
            int_ActiveRow = int_ActiveRow + 1
            int_ActiveColumn = int_ActiveColumn - UBound(arr_SpotStep)
        Next int_i

    End Select

    wks_output.Calculate
    wks_output.Columns.AutoFit
    wks_output.Cells.HorizontalAlignment = xlCenter
End Sub

Private Function Col_AdjFixCashFlow(col_fix As Collection, col_est As Collection, col_disc As Collection) As Collection

Dim col_output As Collection
Set col_output = New Collection

Dim dbl_Output As Double
Dim int_count As Integer

int_count = col_fix.count

Dim int_i As Integer
For int_i = 1 To int_count
    col_output.Add (col_fix(int_i) - col_est(int_i) + col_disc(int_i))
    'Debug.Print col_fix(int_i) - col_est(int_i) + col_disc(int_i)
Next int_i

Set Col_AdjFixCashFlow = col_output

End Function

' WL 20181002: ## METHODS - Hull White Calibration
Private Function HW_Cal_OriVol(dbl_MR As Double, dbl_r As Double, dbl_Vol As Double) As Double()

'Input
Dim col_FixCashFlow As Collection
Dim col_FixCashFlowEndDate As Collection
Dim col_FixCashFlowDF As Collection
'Set col_FixCashFlow = irl_LegA.FixLegCashFlow
Set col_FixCashFlowEndDate = irl_LegA.PeriodEnd
Set col_FixCashFlowDF = irl_LegA.FixLegDF

Dim bln_acc As Boolean

If FreqValue(irl_LegA.Params.PmtFreq) > FreqValue(irl_legB.Params.PmtFreq) Then
    bln_acc = True
Else
    bln_acc = False
End If

If irl_legB.Params.Curve_Disc = irl_legB.Params.Curve_Est Then
    Set col_FixCashFlow = irl_LegA.FixLegCashFlow
Else
    Dim col_FixCashFlowOri As Collection
    Dim col_FltCashFlowEst As Collection
    Dim col_FltCashFlowDisc As Collection

    Set col_FixCashFlowOri = irl_LegA.FixLegCashFlow
    Set col_FltCashFlowDisc = irl_legB.FloatLegCashFlow(irl_legB.Params.Curve_Disc, irl_LegA.FixLegPmtDate, bln_acc)
    Set col_FltCashFlowEst = irl_legB.FloatLegCashFlow(irl_legB.Params.Curve_Est, irl_LegA.FixLegPmtDate, bln_acc)
    Set col_FixCashFlow = Col_AdjFixCashFlow(col_FixCashFlowOri, col_FltCashFlowEst, col_FltCashFlowDisc)
End If

Dim dbl_notional As Double
Dim lng_OptDate As Long
Dim lng_SwapStartDate As Long
dbl_notional = irl_LegA.Params.Notional
lng_OptDate = fld_Params.OptionMat
lng_SwapStartDate = irl_LegA.Params.Swapstart


Dim dbl_OptDF As Double
Dim dbl_F As Double
Dim dbl_MV As Double
Dim dbl_SwapStartDF As Double

dbl_OptDF = irc_Disc_LegA.Lookup_Rate(lng_ValDate, lng_OptDate, "DF", , , True)
dbl_SwapStartDF = irc_Disc_LegA.Lookup_Rate(lng_ValDate, lng_SwapStartDate, "DF", , , True)
dbl_F = irc_Disc_LegA.Lookup_Rate(lng_ValDate, lng_OptDate, "ZERO", , , True) / 100
dbl_MV = Me.marketvalue / dbl_notional * 100

Dim int_cnt As Integer
Dim dbl_OptT As Double
Dim dbl_T As Double
int_cnt = irl_LegA.FixLegCashFlow.count
dbl_OptT = (lng_OptDate - lng_ValDate) / 365
dbl_T = (lng_SwapStartDate - lng_ValDate) / 365


Dim dbl_Output() As Double: ReDim dbl_Output(1 To 2) As Double
Dim dbl_output_r() As Double: ReDim dbl_output_r(1 To 2) As Double
Dim dbl_output_vol() As Double: ReDim dbl_output_vol(1 To 2) As Double
Dim dbl_Jac() As Double: ReDim dbl_Jac(1 To 2, 1 To 2) As Double
Dim dbl_JacInv() As Double: ReDim dbl_JacInv(1 To 2, 1 To 2) As Double
Dim dbl_OutputDiff() As Double: ReDim dbl_OutputDiff(1 To 2) As Double


ReDo:
'dbl_Output = HW_Cal_ObjFunc(dbl_MR, dbl_r, dbl_Vol, lng_SwapStartDate, dbl_notional, dbl_MV, dbl_SwapStartDF, dbl_F, int_cnt, dbl_OptT, dbl_T, col_FixCashFlow, col_FixCashFlowEndDate, col_FixCashFlowDF)
'dbl_output_r = HW_Cal_ObjFunc(dbl_MR, dbl_r + dbl_rShockSize, dbl_Vol, lng_SwapStartDate, dbl_notional, dbl_MV, dbl_SwapStartDF, dbl_F, int_cnt, dbl_OptT, dbl_T, col_FixCashFlow, col_FixCashFlowEndDate, col_FixCashFlowDF)
'dbl_output_vol = HW_Cal_ObjFunc(dbl_MR, dbl_r, dbl_Vol + dbl_VolShockSize, lng_SwapStartDate, dbl_notional, dbl_MV, dbl_SwapStartDF, dbl_F, int_cnt, dbl_OptT, dbl_T, col_FixCashFlow, col_FixCashFlowEndDate, col_FixCashFlowDF)

dbl_Output = HW_Cal_ObjFunc(dbl_MR, dbl_r, dbl_Vol, lng_SwapStartDate, dbl_notional, dbl_MV, dbl_OptDF, dbl_SwapStartDF, dbl_F, int_cnt, dbl_OptT, dbl_T, col_FixCashFlow, col_FixCashFlowEndDate, col_FixCashFlowDF)
dbl_output_r = HW_Cal_ObjFunc(dbl_MR, dbl_r + dbl_rShockSize, dbl_Vol, lng_SwapStartDate, dbl_notional, dbl_MV, dbl_OptDF, dbl_SwapStartDF, dbl_F, int_cnt, dbl_OptT, dbl_T, col_FixCashFlow, col_FixCashFlowEndDate, col_FixCashFlowDF)
dbl_output_vol = HW_Cal_ObjFunc(dbl_MR, dbl_r, dbl_Vol + dbl_VolShockSize, lng_SwapStartDate, dbl_notional, dbl_MV, dbl_OptDF, dbl_SwapStartDF, dbl_F, int_cnt, dbl_OptT, dbl_T, col_FixCashFlow, col_FixCashFlowEndDate, col_FixCashFlowDF)

dbl_Jac(1, 1) = (dbl_output_r(1) - dbl_Output(1)) / dbl_rShockSize
dbl_Jac(1, 2) = (dbl_output_vol(1) - dbl_Output(1)) / dbl_VolShockSize
dbl_Jac(2, 1) = (dbl_output_r(2) - dbl_Output(2)) / dbl_rShockSize
dbl_Jac(2, 2) = (dbl_output_vol(2) - dbl_Output(2)) / dbl_VolShockSize

Dim dbl_det As Double
'error handling
If (dbl_Jac(1, 1) * dbl_Jac(2, 2) - dbl_Jac(1, 2) * dbl_Jac(2, 1)) = 0 Then
    dbl_det = 0.01
Else
dbl_det = 1 / (dbl_Jac(1, 1) * dbl_Jac(2, 2) - dbl_Jac(1, 2) * dbl_Jac(2, 1))
End If

dbl_JacInv(1, 1) = dbl_Jac(2, 2) * dbl_det
dbl_JacInv(1, 2) = -dbl_Jac(1, 2) * dbl_det
dbl_JacInv(2, 1) = -dbl_Jac(2, 1) * dbl_det
dbl_JacInv(2, 2) = dbl_Jac(1, 1) * dbl_det

dbl_OutputDiff(1) = dbl_JacInv(1, 1) * dbl_Output(1) + dbl_JacInv(1, 2) * dbl_Output(2)
dbl_OutputDiff(2) = dbl_JacInv(2, 1) * dbl_Output(1) + dbl_JacInv(2, 2) * dbl_Output(2)

Dim int_num As Integer
Dim dbl_CalOutput() As Double
ReDim dbl_CalOutput(1 To 3) As Double

'If (Abs(dbl_OutputDiff(1)) < dbl_rTolerance And Abs(dbl_OutputDiff(2)) < dbl_VolTolerance) Then
'If (Abs(dbl_output(1)) < dbl_rTolerance And Abs(dbl_output(2)) < dbl_VolTolerance) Then
If (Abs(dbl_Output(1)) < 0.01 And Abs(dbl_Output(2)) < dbl_Tolerance) Then

    'Debug.Print dbl_r; dbl_vol; int_num
    dbl_CalOutput(1) = dbl_r
    dbl_CalOutput(2) = dbl_Vol
    dbl_CalOutput(3) = int_num
ElseIf int_num > dbl_MaxIter Or dbl_Vol < 0 Then
    dbl_CalOutput(1) = dbl_r
    dbl_CalOutput(2) = 9999
    dbl_CalOutput(3) = int_num

Else
    If Abs(dbl_r - dbl_OutputDiff(1)) < 1 Then
        dbl_r = dbl_r - dbl_OutputDiff(1)
    End If

    If dbl_Vol > dbl_OutputDiff(2) And Abs(dbl_Vol - dbl_OutputDiff(2)) < 1 Then
        dbl_Vol = dbl_Vol - dbl_OutputDiff(2)
    End If

    int_num = int_num + 1
    GoTo ReDo
End If

HW_Cal_OriVol = dbl_CalOutput

End Function

Private Function HW_Cal_ObjFunc(dbl_MR As Double, dbl_r As Double, dbl_Vol As Double, lng_SwStart As Long, dbl_notional As Double, dbl_MV As Double, dbl_OptDF As Double, dbl_SwDF As Double, dbl_F As Double, int_cnt As Integer, dbl_OptT As Double, dbl_T As Double, _
    col_FixCashFlow As Collection, col_FixCashFlowEndDate As Collection, col_FixCashFlowDF As Collection) As Double()
'Private Function HW_Cal_ObjFunc(dbl_MR As Double, dbl_r As Double, dbl_Vol As Double, lng_OptDate As Long, dbl_notional As Double, dbl_MV As Double, dbl_OptDF As Double, dbl_F As Double, int_cnt As Integer, dbl_OptT As Double, dbl_T As Double, _
    col_FixCashFlow As Collection, col_FixCashFlowEndDate As Collection, col_FixCashFlowDF As Collection) As Double()

'calculate the bond & bond option valuation for single swaption given r* and OriVol
Dim int_i As Integer
Dim dbl_HwDF As Double
Dim dbl_DiscFixFlow As Double
Dim dbl_BondOption As Double
Dim dbl_DiscFixFlowDiff As Double
Dim dbl_BondOptionDiff As Double
Dim dbl_K As Double
Dim dbl_L As Double
Dim dbl_P0s As Double
Dim dbl_YearFrac As Double, dbl_A As Double, dbl_B As Double

Dim dbl_HwDF_SD As Double
Dim dbl_A_SD As Double, dbl_B_SD As Double


If dbl_T = dbl_OptT Then
    dbl_HwDF_SD = 1
Else
    dbl_B_SD = HW_B(dbl_MR, dbl_OptT, dbl_T)
    dbl_A_SD = HW_A(dbl_MR, dbl_OptT, dbl_T, dbl_Vol, dbl_OptDF, dbl_SwDF, dbl_F)
    dbl_HwDF_SD = dbl_A_SD * Exp(-dbl_B_SD * dbl_r)
End If

    For int_i = 1 To int_cnt
        dbl_P0s = dbl_SwDF * col_FixCashFlowDF(int_i)
        dbl_YearFrac = (col_FixCashFlowEndDate(int_i) - lng_SwStart) / 365
        'dbl_P0s = dbl_OptDF * col_FixCashFlowDF(int_i)
        'dbl_YearFrac = (col_FixCashFlowEndDate(int_i) - lng_OptDate) / 365

        dbl_A = HW_A(dbl_MR, dbl_T, dbl_T + dbl_YearFrac, dbl_Vol, dbl_OptDF, dbl_P0s, dbl_F)
        dbl_B = HW_B(dbl_MR, dbl_T, dbl_T + dbl_YearFrac)
        dbl_HwDF = dbl_A * Exp(-dbl_B * dbl_r)

        dbl_L = col_FixCashFlow(int_i)
        dbl_K = dbl_HwDF * dbl_L / dbl_HwDF_SD
        'dbl_K = dbl_HwDF * dbl_L


        If int_i = int_cnt Then
            dbl_K = dbl_K + dbl_notional * dbl_HwDF / dbl_HwDF_SD
            'dbl_K = dbl_K + dbl_notional * dbl_HwDF
            dbl_L = dbl_L + dbl_notional
        End If

        dbl_DiscFixFlow = dbl_DiscFixFlow + dbl_K
        dbl_BondOption = dbl_BondOption + HW_ZcPut(dbl_Vol, dbl_MR, dbl_OptT, dbl_T, dbl_T + dbl_YearFrac, dbl_L, dbl_K, dbl_SwDF, dbl_P0s)
        'dbl_BondOption = dbl_BondOption + HW_ZcPut(dbl_Vol, dbl_MR, dbl_OptT, dbl_T, dbl_T + dbl_YearFrac, dbl_L, dbl_K, dbl_OptDF, dbl_P0s)
    Next int_i

    dbl_BondOption = dbl_BondOption / dbl_notional * 100
    dbl_DiscFixFlowDiff = dbl_DiscFixFlow - dbl_notional
    dbl_BondOptionDiff = dbl_BondOption - dbl_MV

'Output
Dim dbl_Output() As Double
ReDim dbl_Output(1 To 2) As Double
dbl_Output(1) = dbl_DiscFixFlowDiff
dbl_Output(2) = dbl_BondOptionDiff

HW_Cal_ObjFunc = dbl_Output
'Debug.Print dbl_output(1); dbl_output(2)

End Function

Public Property Get HW_CalibrateVol(dbl_MR As Double, dbl_r As Double, dbl_Vol As Double) As Collection

    Dim dbl_Output() As Double
    ReDim dbl_Output(1 To 3) As Double

    dbl_Output = HW_Cal_OriVol(dbl_MR, dbl_r, dbl_Vol)

    Dim col_output As Collection
    Set col_output = New Collection

    col_output.Add dbl_Output(1)
    col_output.Add dbl_Output(2)
    col_output.Add dbl_Output(3)

    Set HW_CalibrateVol = col_output

End Property

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