VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inst_RngAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ## MEMBER DATA
' Components
Private irl_LegA As IRLeg, irl_legB As IRLeg
Private irl_LegADigi As IRLegDigi, irl_LegBDigi As IRLegDigi

'Constants
Private Const maxday As Integer = "365"
Private Const dbl_StrikeGap As Double = 0.0001

' Dependent curves
Private fxs_Spots As Data_FXSpots
    
Private cvl_VolCurve_LegA As Data_CapVolsQJK
Private cvl_VolCurve_LegADigi_upper As Data_CapVolsQJK, cvl_VolCurve_LegADigi_lower As Data_CapVolsQJK

Private cvl_VolCurve_LegB As Data_CapVolsQJK
Private cvl_VolCurve_LegBDigi_upper As Data_CapVolsQJK, cvl_VolCurve_LegBDigi_lower As Data_CapVolsQJK

' Collection
Private dblLst_VolCurve_LegA As Collection
Private dblLst_VolCurve_LegADigi_upper As Collection, dblLst_VolCurve_LegADigi_lower As Collection
Private dblLst_VolCurve_LegADigi_upperUp As Collection, dblLst_VolCurve_LegADigi_lowerUp As Collection

Private dblLst_ATMVolCurve_LegA As Collection

Private dblLst_VolCurve_LegB As Collection
Private dblLst_VolCurve_LegBDigi_upper As Collection, dblLst_VolCurve_LegBDigi_lower As Collection
Private dblLst_VolCurve_LegBDigi_upperUp As Collection, dblLst_VolCurve_LegBDigi_lowerUp As Collection

Private dblLst_ATMVolCurve_LegB As Collection

' Static values
Private dic_GlobalStaticInfo As Dictionary, dic_CurveDependencies As Dictionary

Private str_CCY_PnL As String, int_Sign As Integer
Private cal_pmt As Calendar

Private fld_Params As InstParams_RngAcc

Private scf_Premium As SCF
Private lng_ValDate As Long

Private str_VolCurve_LegA As String
Private str_VolCurve_LegB As String
Private str_VolCurve_LegA_ATM As String
Private str_VolCurve_LegB_ATM As String

'KL 201901 HW1F Enhancement
Private irc_Disc_LegA As Data_IRCurve
Private irc_Est_LegB As Data_IRCurve

Private HWPDE As PDE_Matrix
Private int_Direction As Integer
Private int_SpotStep As Integer
Private int_TimeStep As Integer
Private dbl_MeanRev As Double
Private arr_PDE_CalcPeriod() As Variant
Private arr_PDE_UndVal() As Variant
Private arr_PDE_Option() As Variant
Private HW_vol As Data_HullWhiteVol
Private HW_Payoff_Time() As Variant
Private dbl_PDE_MV As Double
Private Const Erase_Switch As Boolean = True
Private Const bln_Smooth As Boolean = True 'Payoff Independent Smoothing Toggle

' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As InstParams_RngAcc, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing)
    
    ' Store static values
    lng_ValDate = fld_ParamsInput.LegA.ValueDate
    
    ' Store static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    str_CCY_PnL = fld_ParamsInput.CCY_PnL
    If fld_ParamsInput.Pay_LegA = True Then int_Sign = -1 Else int_Sign = 1
    If fld_ParamsInput.Callable_LegA = True Then int_Direction = 1 Else int_Direction = -1
    
    
    fld_Params = fld_ParamsInput

'+++++++++++++++++++++++++++++++++++++++++++++++++LEG A+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Store components for LEG A
Set irl_LegA = New IRLeg
Call irl_LegA.Initialize(fld_ParamsInput.LegA, dic_CurveSet, dic_StaticInfoInput)
    
'****************************ExoticLeg_A**************************************************************************
Select Case UCase(fld_ParamsInput.LegA.ExoticType)
Case "RANGE"
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        'Recycle PeriodStart & PeriodEnd into fld_ParamsInput
        Dim int_ctrA As Integer
        Dim dic_PeriodStartA As New Scripting.Dictionary
        Dim dic_PeriodEndA As New Scripting.Dictionary

        For int_ctrA = 1 To irl_LegA.periodstart.count
            dic_PeriodStartA.Add (irl_LegA.periodstart(int_ctrA)), 0
            dic_PeriodEndA.Add (irl_LegA.PeriodEnd(int_ctrA)), 0
        Next int_ctrA
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        
    'Store Vol Curve
    Set irl_LegADigi = New IRLegDigi
    Call irl_LegADigi.Initialize(fld_ParamsInput.LegA2Digi, dic_CurveSet, dic_StaticInfoInput, dic_PeriodStartA, dic_PeriodEndA)
    

    If fld_ParamsInput.LegA.AboveUpper <> "-" Then
        str_VolCurve_LegA = fld_ParamsInput.LegA.CCY & "_" & Right(fld_ParamsInput.LegA.RangeIndex, 2)
        
        Set cvl_VolCurve_LegADigi_upper = GetObject_CapVolSurf(str_VolCurve_LegA, fld_ParamsInput.LegA.upper, True, False)
        Set dblLst_VolCurve_LegADigi_upper = StoreVolsDigi(irl_LegADigi, cvl_VolCurve_LegADigi_upper)
        
        Set cvl_VolCurve_LegADigi_upper = GetObject_CapVolSurf(str_VolCurve_LegA, fld_ParamsInput.LegA.upper + dbl_StrikeGap, True, False)
        Set dblLst_VolCurve_LegADigi_upperUp = StoreVolsDigi(irl_LegADigi, cvl_VolCurve_LegADigi_upper)
    End If
    
    If fld_ParamsInput.LegA.AboveLower <> "-" Then
        str_VolCurve_LegA = fld_ParamsInput.LegA.CCY & "_" & Right(fld_ParamsInput.LegA.RangeIndex, 2)
        
        Set cvl_VolCurve_LegADigi_lower = GetObject_CapVolSurf(str_VolCurve_LegA, fld_ParamsInput.LegA.lower, True, False)
        Set dblLst_VolCurve_LegADigi_lower = StoreVolsDigi(irl_LegADigi, cvl_VolCurve_LegADigi_lower)
        
        Set cvl_VolCurve_LegADigi_lower = GetObject_CapVolSurf(str_VolCurve_LegA, fld_ParamsInput.LegA.lower + dbl_StrikeGap, True, False)
        Set dblLst_VolCurve_LegADigi_lowerUp = StoreVolsDigi(irl_LegADigi, cvl_VolCurve_LegADigi_lower)
    End If
    
    'Store ATM Vol (For Timing Adjustment)
    If UCase(fld_ParamsInput.LegA.FixedFloat) = "FLOAT" Then

        str_VolCurve_LegA_ATM = fld_ParamsInput.LegA.CCY & "_" & fld_ParamsInput.LegA.index
        Set dblLst_ATMVolCurve_LegA = StoreATMvols(irl_LegA, str_VolCurve_LegA_ATM, True, False)

    End If

End Select
'+++++++++++++++++++++++++++++++++++++++++++++++++LEG B+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' Store components for LEG B
Set irl_legB = New IRLeg
Call irl_legB.Initialize(fld_ParamsInput.LegB, dic_CurveSet, dic_StaticInfoInput)

'****************************ExoticLeg_B**************************************************************************
Select Case UCase(fld_ParamsInput.LegB.ExoticType)
Case "RANGE"
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        'Recycle PeriodStart & PeriodEnd into fld_ParamsInput
        Dim int_ctrB As Integer
        Dim dic_PeriodStartB As New Scripting.Dictionary
        Dim dic_PeriodEndB As New Scripting.Dictionary

        For int_ctrB = 1 To irl_legB.periodstart.count
            dic_PeriodStartB.Add (irl_legB.periodstart(int_ctrB)), 0
            dic_PeriodEndB.Add (irl_legB.PeriodEnd(int_ctrB)), 0
        Next int_ctrB
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        
    'Populate digitalcaplets for legB
    Set irl_LegBDigi = New IRLegDigi
    Call irl_LegBDigi.Initialize(fld_ParamsInput.LegB2Digi, dic_CurveSet, dic_StaticInfoInput, dic_PeriodStartB, dic_PeriodEndB)
        
    'Store Vol Curve
    If fld_ParamsInput.LegB.AboveUpper <> "-" Then
    
        str_VolCurve_LegB = fld_ParamsInput.LegB.CCY & "_" & Right(fld_ParamsInput.LegB.RangeIndex, 2)
        
        Set cvl_VolCurve_LegBDigi_upper = GetObject_CapVolSurf(str_VolCurve_LegB, fld_ParamsInput.LegB.upper, True, False)
        Set dblLst_VolCurve_LegBDigi_upper = StoreVolsDigi(irl_LegBDigi, cvl_VolCurve_LegBDigi_upper)
        
        Set cvl_VolCurve_LegBDigi_upper = GetObject_CapVolSurf(str_VolCurve_LegB, fld_ParamsInput.LegB.upper + dbl_StrikeGap, True, False)
        Set dblLst_VolCurve_LegBDigi_upperUp = StoreVolsDigi(irl_LegBDigi, cvl_VolCurve_LegBDigi_upper)
    End If
    
    If fld_ParamsInput.LegB.AboveLower <> "-" Then
    
        str_VolCurve_LegB = fld_ParamsInput.LegB.CCY & "_" & Right(fld_ParamsInput.LegB.RangeIndex, 2)
         
        Set cvl_VolCurve_LegBDigi_lower = GetObject_CapVolSurf(str_VolCurve_LegB, fld_ParamsInput.LegB.lower, True, False)
        Set dblLst_VolCurve_LegBDigi_lower = StoreVolsDigi(irl_LegBDigi, cvl_VolCurve_LegBDigi_lower)

        Set cvl_VolCurve_LegBDigi_lower = GetObject_CapVolSurf(str_VolCurve_LegB, fld_ParamsInput.LegB.lower + dbl_StrikeGap, True, False)
        Set dblLst_VolCurve_LegBDigi_lowerUp = StoreVolsDigi(irl_LegBDigi, cvl_VolCurve_LegBDigi_lower)
        
    End If
    
    'Store ATM Vol (For Timing Adjustment)
    If UCase(fld_ParamsInput.LegB.FixedFloat) = "FLOAT" Then

        str_VolCurve_LegB_ATM = fld_ParamsInput.LegB.CCY & "_" & fld_ParamsInput.LegB.index
        Set dblLst_ATMVolCurve_LegB = StoreATMvols(irl_legB, str_VolCurve_LegB_ATM, True, False)
    

    End If
End Select

'******************************ExoticLeg_B**************************************************************************

''''''''''KL - PDE Implementation''''''''
'''''''''''''''''START'''''''''''''''''''
If fld_Params.IsCallable = True Then
    
    Set irc_Disc_LegA = dic_CurveSet(CurveType.IRC)(fld_Params.LegA.Curve_Disc)
    Set irc_Est_LegB = dic_CurveSet(CurveType.IRC)(fld_Params.LegB.Curve_Est)
    
    Dim cas_Found As CalendarSet: Set cas_Found = dic_GlobalStaticInfo(StaticInfoType.CalendarSet)
    
    int_SpotStep = fld_Params.spotstep
    int_TimeStep = fld_Params.timestep
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
        col_Strike.Add (fld_Params.LegA.RateOrMargin)
    Next k
    
    Set HW_vol = New Data_HullWhiteVol
    Call HW_vol.Initialize(dbl_MeanRev, fld_Params.LegA.CCY, fld_Params.GeneratorA, fld_Params.GeneratorB, col_ExDate, col_SwapStart, irl_LegA.MatDate, col_Strike, fld_Params.VolCurve)
    
    'Hull White PDE
    Set HWPDE = New PDE_Matrix
    Call HWPDE.Initialize(fld_Params.LegA.ValueDate, irl_LegA.MatDate, HW_vol.FullFinalHwVol, irl_LegA, irl_legB, int_SpotStep, int_TimeStep, dbl_MeanRev, CRA)
    
    'Call HW_CalcPeriod(irl_legA, irl_LegADigi)
    Dim arr_UndVal_A() As Variant
    Dim arr_UndVal_B() As Variant
    Dim arr_SpotStep() As Double
    Dim arr_TimeStep() As Double
    
    arr_TimeStep = HWPDE.timestep
    arr_SpotStep = HWPDE.spotstep
'   arr_UndVal_A = HW_Underlying(irl_legA)
'   arr_UndVal_B = HW_Underlying(irl_legB)
    
    'ReDim arr_PDE_UndVal(1 To UBound(HWPDE.TimeStep), 1 To UBound(HWPDE.SpotStep))
    Dim int_i As Integer, int_j As Integer
'   For int_i = UBound(HWPDE.TimeStep) To 1 Step -1
'       For int_j = 1 To UBound(HWPDE.SpotStep)
'           arr_PDE_UndVal(int_i, int_j) = int_Sign * (arr_UndVal_A(int_i, int_j) - arr_UndVal_B(int_i, int_j))
'       Next int_j
'   Next int_i
    Call HW_Option
End If
'''''''''''''''''END'''''''''''''''''''''

Set fxs_Spots = GetObject_FXSpots(True)

' Determine curve dependencies
Set dic_CurveDependencies = fxs_Spots.Lookup_CurveDependencies(fld_Params.CCY_PnL)
Set dic_CurveDependencies = Convert_MergeDicts(dic_CurveDependencies, irl_LegA.CurveDependencies)
Set dic_CurveDependencies = Convert_MergeDicts(dic_CurveDependencies, irl_legB.CurveDependencies)

End Sub
Private Function HW_CalcPeriod(irl_Leg As IRLeg, ird_Leg As IRLegDigi) As Variant()

Dim int_i As Integer
Dim int_j As Integer
Dim int_cnt As Integer
Dim int_cnt1 As Integer

Dim lng_T1_Date As Long
Dim lng_T2_Date As Long
Dim lng_PeriodEnd As Long

Dim dbl_T1 As Double
Dim dbl_T2 As Double
Dim dbl_PT1 As Double
Dim dbl_PT2 As Double
Dim dbl_A As Double
Dim dbl_B As Double

Dim dbl_T1_F As Double
Dim dbl_T2_F As Double
Dim dbl_PT1_F As Double
Dim dbl_PT2_F As Double
Dim dbl_A_F As Double
Dim dbl_B_F As Double
Dim dbl_CalcPeriod_F As Double

Dim dbl_X As Double
Dim dbl_EffVol As Double
Dim dbl_CalcPeriod As Double
Dim dbl_HWBond As Double
Dim dbl_HWBond_F As Double
Dim dbl_Forward As Double

Dim str_RangeType As String
Dim str_AboveUpper As String
Dim str_abovelower As String
Dim dbl_Upper As Double
Dim dbl_Lower As Double

Dim cal_Leg As Calendar: cal_Leg = irl_Leg.Calendar
Dim col_PeriodStart As New Collection
Dim col_PeriodEnd As New Collection

Dim arr_SpotStep() As Double
Dim arr_TimeStep() As Double
Dim arr_TimeLabel() As Integer
Dim arr_TimeStep_ImpDate() As Double
Dim arr_PDE_Temp() As Double
Dim arr_output() As Variant
Dim arr_output1() As Variant

arr_SpotStep() = HWPDE.spotstep
arr_TimeStep() = HWPDE.timestep
arr_TimeLabel() = HWPDE.timelabel
arr_TimeStep_ImpDate() = HWPDE.TimeStep_ImpDate

Set col_PeriodStart = irl_Leg.periodstart
Set col_PeriodEnd = irl_Leg.PeriodEnd

'Curve Setup
Dim irc_Est As Data_IRCurve
Dim irc_Disc As Data_IRCurve

Set irc_Est = GetObject_IRCurve(ird_Leg.Params.Curve_Est, True, False)
Set irc_Disc = GetObject_IRCurve(ird_Leg.Params.Curve_Disc, True, False)

'For python
Dim arr_A() As Variant
Dim arr_B() As Variant
Dim arr_A_F() As Variant
Dim arr_B_F() As Variant
Dim arr_vol As Variant
Dim arr_cal_F As Variant
Dim arr_spread As Variant
Dim arr_upper As Variant
Dim arr_lower As Variant
Dim arr_rangetype As Variant
Dim arr_CalcPeriod As Variant

ReDim arr_A(1 To UBound(arr_TimeStep))
ReDim arr_B(1 To UBound(arr_TimeStep))
ReDim arr_A_F(1 To UBound(arr_TimeStep))
ReDim arr_B_F(1 To UBound(arr_TimeStep))
ReDim arr_vol(1 To UBound(arr_TimeStep))
ReDim arr_cal_F(1 To UBound(arr_TimeStep))
ReDim arr_spread(1 To UBound(arr_TimeStep))
ReDim arr_upper(1 To UBound(arr_TimeStep))
ReDim arr_lower(1 To UBound(arr_TimeStep))
ReDim arr_rangetype(1 To UBound(arr_TimeStep))
ReDim arr_CalcPeriod(1 To UBound(arr_TimeStep))

'Process Start:
int_cnt = 1
int_cnt1 = 0
ReDim arr_PDE_Temp(1 To UBound(arr_SpotStep)) As Double
ReDim arr_output(1 To UBound(arr_TimeStep), 1 To UBound(arr_SpotStep))
ReDim arr_output1(1 To UBound(arr_TimeStep), 1 To UBound(arr_SpotStep))
lng_PeriodEnd = col_PeriodEnd(col_PeriodEnd.count - int_cnt + 1)

Dim x As Double
Dim sum_x As Double
Dim y As Double
Dim sum_y As Double

'Smart Multiplication
Dim arr_SmartMult() As Double
ReDim arr_SmartMult(1 To UBound(arr_TimeStep)) As Double

arr_SmartMult(UBound(arr_TimeStep)) = 1

For int_i = UBound(arr_TimeStep) - 1 To 1 Step -1
    dbl_T1 = arr_TimeStep(int_i)
    dbl_T2 = arr_TimeStep(int_i + 1)
    dbl_PT1 = irc_Disc_LegA.Lookup_Rate(lng_ValDate, lng_ValDate + dbl_T1 * 365, "DF", , , True)
    dbl_PT2 = irc_Disc_LegA.Lookup_Rate(lng_ValDate, lng_ValDate + dbl_T2 * 365, "DF", , , True)
    dbl_A = HWPDE.HWBond_A(dbl_PT1, dbl_PT2, dbl_T1, dbl_T2)
    
    arr_SmartMult(int_i) = dbl_A
Next int_i

'Convert PeriodStart Collection into Array
Dim arr_PeriodStart() As Double
ReDim arr_PeriodStart(1 To col_PeriodStart.count)

For int_i = 1 To col_PeriodStart.count
    arr_PeriodStart(int_i) = (col_PeriodStart(col_PeriodStart.count - int_i + 1) - lng_ValDate) / 365
Next int_i

'Collect the CalcPeriod

arr_A(UBound(arr_TimeStep)) = 1
arr_B(UBound(arr_TimeStep)) = 0
arr_A_F(UBound(arr_TimeStep)) = 1
arr_B_F(UBound(arr_TimeStep)) = 0
arr_vol(UBound(arr_TimeStep)) = 0.0001
arr_CalcPeriod(UBound(arr_TimeStep)) = 0.0001
arr_cal_F(UBound(arr_TimeStep)) = 0.0001
arr_spread(UBound(arr_TimeStep)) = 0
arr_upper(UBound(arr_TimeStep)) = 0
arr_lower(UBound(arr_TimeStep)) = 0
arr_rangetype(UBound(arr_TimeStep)) = ""

Dim arr_forward As Variant
ReDim arr_forward(1 To UBound(arr_TimeStep), 1 To UBound(arr_SpotStep))
    
For int_i = UBound(arr_TimeStep) - 1 To 1 Step -1

    If arr_TimeStep(int_i + 1) = (col_PeriodStart(col_PeriodStart.count - int_cnt + 1) - lng_ValDate) / 365 Then
        If int_cnt < col_PeriodStart.count Then int_cnt = int_cnt + 1
        lng_PeriodEnd = col_PeriodEnd(col_PeriodEnd.count - int_cnt + 1)
    End If
    
    'To find the Effective Volatility
    Dim int_StartDelay1 As Integer
    Dim int_StartDelay2 As Integer
    
    If UCase(irl_Leg.Params.CCY) = "USD" Then
        int_StartDelay1 = 0
        int_StartDelay2 = 2
    Else
        int_StartDelay1 = -1
        int_StartDelay2 = 1
    End If
    
    dbl_T1_F = arr_TimeStep(int_i)
    lng_T1_Date = date_workday(lng_ValDate + 365 * dbl_T1_F + int_StartDelay1, int_StartDelay2, cal_Leg.HolDates, cal_Leg.Weekends)
    dbl_T1_F = (lng_T1_Date - lng_ValDate) / 365
    lng_T2_Date = date_workday(Date_NextCoupon(lng_ValDate + 365 * dbl_T1_F, ird_Leg.Params.index, cal_Leg, 1, irl_Leg.EOM, "MOD FOLL") - 1, 1, cal_Leg.HolDates, cal_Leg.Weekends)
    dbl_T2_F = (lng_T2_Date - lng_ValDate) / 365
    dbl_PT1_F = irc_Disc.Lookup_Rate(lng_ValDate, lng_ValDate + dbl_T1_F * 365, "DF", , , True)
    dbl_PT2_F = irc_Disc.Lookup_Rate(lng_ValDate, lng_ValDate + dbl_T2_F * 365, "DF", , , True)
    dbl_A_F = HWPDE.HWBond_A(dbl_PT1_F, dbl_PT2_F, dbl_T1_F, dbl_T2_F)
    dbl_B_F = HW_B(dbl_MeanRev, dbl_T1_F, dbl_T2_F)
    
    dbl_T1 = arr_TimeStep(int_i)
    dbl_T2 = (lng_PeriodEnd - lng_ValDate) / 365
    dbl_PT1 = irc_Disc.Lookup_Rate(lng_ValDate, lng_ValDate + dbl_T1 * 365, "DF", , , True)
    dbl_PT2 = irc_Disc.Lookup_Rate(lng_ValDate, lng_ValDate + dbl_T2 * 365, "DF", , , True)
    dbl_A = HWPDE.HWBond_A(dbl_PT1, dbl_PT2, dbl_T1, dbl_T2)
    dbl_B = HW_B(dbl_MeanRev, dbl_T1, dbl_T2)
    dbl_CalcPeriod = calc_yearfrac(lng_ValDate + dbl_T1 * 365, lng_ValDate + arr_TimeStep(int_i + 1) * 365, irl_Leg.Params.Daycount)
    
    
    If dbl_PT1 = 1 Then
        dbl_EffVol = 0.00001
    Else
        dbl_EffVol = (-(HWPDE.LookUpSigma(dbl_T1)) / Log(dbl_PT1) * arr_TimeStep(int_i)) * Sqr(dbl_CalcPeriod)
    End If

    'To find the Forward Rate
    Dim dbl_spread As Double
    Dim dbl_spread_est As Double
    Dim dbl_spread_disc As Double

    dbl_CalcPeriod_F = calc_yearfrac(lng_T1_Date, lng_T2_Date, irl_Leg.Params.Daycount)

    dbl_spread_disc = irc_Disc.Lookup_Rate(lng_ValDate, lng_T1_Date, "DF", , , True) / irc_Disc.Lookup_Rate(lng_ValDate, lng_T2_Date, "DF", , , True)
    dbl_spread_disc = (dbl_spread_disc - 1) / dbl_CalcPeriod_F

    dbl_spread_est = irc_Est.Lookup_Rate(lng_ValDate, lng_T1_Date, "DF", , , True) / irc_Est.Lookup_Rate(lng_ValDate, lng_T2_Date, "DF", , , True)
    dbl_spread_est = (dbl_spread_est - 1) / dbl_CalcPeriod_F

    dbl_spread = (dbl_spread_est - dbl_spread_disc)
    
    'Variable Range
    If Not irl_Leg.Params.VariableRange1 Is Nothing Then
        str_AboveUpper = irl_Leg.Params.VariableRange1(col_PeriodStart(col_PeriodEnd.count - int_cnt + 1))
        dbl_Upper = irl_Leg.Params.VariableRange2(col_PeriodStart(col_PeriodEnd.count - int_cnt + 1))
        str_abovelower = irl_Leg.Params.VariableRange3(col_PeriodStart(col_PeriodEnd.count - int_cnt + 1))
        dbl_Lower = irl_Leg.Params.VariableRange4(col_PeriodStart(col_PeriodEnd.count - int_cnt + 1))
        str_RangeType = ird_Leg.GetRangeType(str_AboveUpper, str_abovelower)
    Else
        str_RangeType = ird_Leg.rangetype
        dbl_Upper = ird_Leg.Params.upper
        dbl_Lower = ird_Leg.Params.lower
    End If
    
    'Add in New CalcPeriod
    arr_A(int_i) = dbl_A
    arr_B(int_i) = dbl_B
    arr_A_F(int_i) = dbl_A_F
    arr_B_F(int_i) = dbl_B_F
    arr_vol(int_i) = dbl_EffVol
    arr_cal_F(int_i) = dbl_CalcPeriod_F
    arr_spread(int_i) = dbl_spread
    arr_upper(int_i) = dbl_Upper
    arr_lower(int_i) = dbl_Lower
    arr_rangetype(int_i) = str_RangeType
    arr_CalcPeriod(int_i) = dbl_CalcPeriod
    
Next int_i

'Sigma
Dim arr_sigma() As Double
arr_sigma = HWPDE.HW_Sigma

arr_output1 = calcperiod(arr_SpotStep, arr_A, arr_B, arr_A_F, arr_B_F, arr_CalcPeriod, arr_cal_F, arr_spread, arr_upper, arr_lower, arr_vol, arr_rangetype)
arr_output1 = PDE_CalcPeriod(arr_SpotStep, arr_TimeStep, arr_TimeLabel, dbl_MeanRev, arr_sigma, arr_PeriodStart, arr_output1, arr_SmartMult)

'Put into Final Output
For int_i = 1 To UBound(arr_TimeStep)
    For int_j = 1 To UBound(arr_SpotStep)
        arr_output(int_i, int_j) = arr_output1(int_i - 1, int_j - 1)
    Next int_j
Next int_i

arr_PDE_CalcPeriod = arr_output
HW_CalcPeriod = arr_output

If Erase_Switch = True Then
    Erase arr_output
    Erase arr_A
    Erase arr_B
    Erase arr_A_F
    Erase arr_B_F
    Erase arr_vol
    Erase arr_cal_F
    Erase arr_spread
    Erase arr_upper
    Erase arr_lower
    Erase arr_rangetype
    Erase arr_CalcPeriod
End If

End Function


Private Function Payoff_Dependent_Smoothing(str_RangeType As String, dbl_Upper As Double, dbl_Lower As Double, dbl_Forward As Double, dbl_EffVol As Double, bln_Switch As Boolean) As Double

Dim dbl_Norm1 As Double
Dim dbl_Norm2 As Double
Dim dbl_Output As Double
    
On Error Resume Next
If dbl_Forward < 0.000001 Then dbl_Forward = 0.000001

If bln_Switch = True Then

    Select Case str_RangeType
        Case "AboveLower"
            dbl_Norm1 = Calc_NormalCDF((Log(dbl_Forward / dbl_Lower * 100) - dbl_EffVol ^ 2 * 0.5) / dbl_EffVol)
            dbl_Output = dbl_Norm1
        Case "Outside"
            dbl_Norm1 = 1 - Calc_NormalCDF((Log(dbl_Forward / dbl_Lower * 100) - dbl_EffVol ^ 2 * 0.5) / dbl_EffVol)
            dbl_Norm2 = Calc_NormalCDF((Log(dbl_Forward / dbl_Upper * 100) - dbl_EffVol ^ 2 * 0.5) / dbl_EffVol)
            dbl_Output = dbl_Norm1 + dbl_Norm2
        Case "Between"
            dbl_Norm1 = Calc_NormalCDF((Log(dbl_Forward / dbl_Lower * 100) - dbl_EffVol ^ 2 * 0.5) / dbl_EffVol)
            dbl_Norm2 = Calc_NormalCDF((Log(dbl_Forward / dbl_Upper * 100) - dbl_EffVol ^ 2 * 0.5) / dbl_EffVol)
            dbl_Output = dbl_Norm1 - dbl_Norm2
        Case "BelowUpper"
            dbl_Norm2 = 1 - Calc_NormalCDF((Log(dbl_Forward / dbl_Upper * 100) - dbl_EffVol ^ 2 * 0.5) / dbl_EffVol)
            dbl_Output = dbl_Norm2
    End Select

Else

    Select Case str_RangeType
        Case "AboveLower"
            If dbl_Forward > dbl_Lower / 100 Then dbl_Output = 1

        Case "Outside"
            If dbl_Lower / 100 > dbl_Forward Or dbl_Upper / 100 < dbl_Forward Then dbl_Output = 1

        Case "Between"
            If dbl_Lower / 100 < dbl_Forward And dbl_Upper / 100 > dbl_Forward Then dbl_Output = 1
        
        Case "BelowUpper"
            If dbl_Upper / 100 > dbl_Forward Then dbl_Output = 1
            
    End Select

End If

Payoff_Dependent_Smoothing = dbl_Output

End Function


Private Function HW_Underlying(irl_Leg As IRLeg, ird_Leg As IRLegDigi) As Double()

Dim dbl_Strike As Double

Dim int_cnt As Integer
Dim int_cnt1 As Integer
Dim int_i As Integer
Dim int_j As Integer
Dim int_k As Integer

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
Dim arr_output As Double
Dim arr_CalcPeriod() As Variant

Dim arr_payoff() As Double
Dim arr_Payoff_Time() As Double

Dim col_PeriodStart As New Collection
Dim col_PeriodEnd As New Collection
Dim col_EstStart As New Collection
Dim col_EstEnd As New Collection

arr_SpotStep() = HWPDE.spotstep
arr_TimeStep() = HWPDE.timestep

Set col_PeriodStart = irl_Leg.periodstart
Set col_PeriodEnd = irl_Leg.PeriodEnd

If irl_Leg.IsFixed = False Then
    Set col_EstStart = irl_Leg.EstStart
    Set col_EstEnd = irl_Leg.EstEnd
End If

'Range Accrual Treatment
Dim x As Double
x = Timer()
ReDim arr_CalcPeriod(1 To UBound(arr_TimeStep), 1 To UBound(arr_SpotStep))
If UCase(irl_Leg.Params.ExoticType) = "RANGE" Then
    arr_CalcPeriod = HW_CalcPeriod(irl_Leg, ird_Leg)
End If
Debug.Print ("CalcPeriod:" & (Timer() - x))

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
                    
        dbl_T1 = calc_yearfrac(lng_ValDate, col_PeriodStart(int_cnt), "ACT/365")
        dbl_T2 = calc_yearfrac(lng_ValDate, col_PeriodEnd(int_cnt), "ACT/365")
        dbl_PT1 = irc_Disc_LegA.Lookup_Rate(lng_ValDate, col_PeriodStart(int_cnt), "DF", , , True)
        dbl_PT2 = irc_Disc_LegA.Lookup_Rate(lng_ValDate, col_PeriodEnd(int_cnt), "DF", , , True)
        dbl_A = HWPDE.HWBond_A(dbl_PT1, dbl_PT2, dbl_T1, dbl_T2)
        dbl_B = HW_B(dbl_MeanRev, dbl_T1, dbl_T2)
        
        dbl_CalcPeriod = calc_yearfrac(col_PeriodStart(int_cnt), col_PeriodEnd(int_cnt), irl_Leg.Params.Daycount)
        
        For int_k = 1 To UBound(arr_TimeStep)
            If arr_TimeStep(int_k) = dbl_T1 Then
                Exit For
            End If
        Next int_k
        
        Select Case irl_Leg.IsFixed
         
        Case True:
        
            'Variable Rate
            If Not irl_Leg.Params.VariableRate Is Nothing Then
                dbl_Strike = irl_Leg.Params.VariableRate(col_PeriodStart(int_cnt))
            Else
                dbl_Strike = irl_Leg.RateOrMargin
            End If
            
            arr_Payoff_Time(int_i) = dbl_T1
            
            For int_j = 1 To UBound(arr_SpotStep)
    
                dbl_X = arr_SpotStep(int_j)
                dbl_HWBond = dbl_A * Exp(-dbl_B * dbl_X)
                
                'Range Accrual Treatment
                If UCase(irl_Leg.Params.ExoticType) = "RANGE" Then
                    arr_payoff(int_j, int_i) = dbl_Strike / 100 * irl_Leg.Params.Notional * arr_CalcPeriod(int_k, int_j)
                Else
                    arr_payoff(int_j, int_i) = dbl_Strike / 100 * irl_Leg.Params.Notional * dbl_HWBond * dbl_CalcPeriod
                End If
            Next int_j
            
        Case False:
            
             'Variable Rate
             Dim dbl_Margin As Double
             If Not irl_Leg.Params.VariableRate Is Nothing Then
                dbl_Margin = irl_Leg.Params.VariableRate(col_PeriodStart(int_cnt))
             Else
                dbl_Margin = irl_Leg.RateOrMargin
             End If
            
             dbl_T1_F = calc_yearfrac(lng_ValDate, col_EstStart(int_cnt), "ACT/365")
             dbl_T2_F = calc_yearfrac(lng_ValDate, col_EstEnd(int_cnt), "ACT/365")
             dbl_PT1_F = irc_Disc_LegA.Lookup_Rate(lng_ValDate, col_EstStart(int_cnt), "DF", , , True)
             dbl_PT2_F = irc_Disc_LegA.Lookup_Rate(lng_ValDate, col_EstEnd(int_cnt), "DF", , , True)
             
             dbl_A_F = HWPDE.HWBond_A(dbl_PT1_F, dbl_PT2_F, dbl_T1_F, dbl_T2_F)
             dbl_B_F = HW_B(dbl_MeanRev, dbl_T1_F, dbl_T2_F)
             
             dbl_CalcPeriod_F = calc_yearfrac(col_EstStart(int_cnt), col_EstEnd(int_cnt), irl_Leg.Params.Daycount)
             
             arr_Payoff_Time(int_i) = dbl_T1_F
             
             Dim dbl_spread As Double
             Dim dbl_spread_est As Double
             Dim dbl_spread_disc As Double
             
             dbl_spread_disc = irc_Disc_LegA.Lookup_Rate(lng_ValDate, col_EstStart(int_cnt), "DF", , , True) / irc_Disc_LegA.Lookup_Rate(lng_ValDate, col_EstEnd(int_cnt), "DF", , , True)
             dbl_spread_disc = (dbl_spread_disc - 1) / dbl_CalcPeriod_F
             dbl_spread_est = irc_Est_LegB.Lookup_Rate(lng_ValDate, col_EstStart(int_cnt), "DF", , , True) / irc_Est_LegB.Lookup_Rate(lng_ValDate, col_EstEnd(int_cnt), "DF", , , True)
             dbl_spread_est = (dbl_spread_est - 1) / dbl_CalcPeriod_F

             dbl_spread = (dbl_spread_est - dbl_spread_disc)
             
             For int_j = 1 To UBound(arr_SpotStep)
        
                dbl_X = arr_SpotStep(int_j)
                dbl_HWBond = dbl_A * Exp(-dbl_B * dbl_X)
                dbl_HWBond_F = dbl_A_F * Exp(-dbl_B_F * dbl_X)
                dbl_Forward = (1 / dbl_HWBond_F - 1) / dbl_CalcPeriod_F + dbl_spread
             
                'Range Accrual Treatment
                If UCase(irl_Leg.Params.ExoticType) = "RANGE" Then
                    arr_payoff(int_j, int_i) = (dbl_Forward + dbl_Margin / 100) * irl_Leg.Params.Notional * arr_CalcPeriod(int_k, int_j)
                Else
                    arr_payoff(int_j, int_i) = (dbl_Forward + dbl_Margin / 100) * irl_Leg.Params.Notional * dbl_HWBond * dbl_CalcPeriod
                End If
             Next int_j
        End Select
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Calculating MV when the Range Accrual is partially fixed and partially not fixed'''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

SkipPayoff:
If lng_ValDate <= col_PeriodEnd(int_cnt) And lng_ValDate > col_PeriodStart(int_cnt) Then
    
    int_i = int_i + 1
    ReDim Preserve arr_payoff(1 To UBound(arr_SpotStep), 1 To int_i) As Double
    ReDim Preserve arr_Payoff_Time(1 To int_i) As Double
    
    arr_Payoff_Time(int_i) = 0

    'HW Discounting Factor
    dbl_T1 = 0
    dbl_T2 = calc_yearfrac(lng_ValDate, col_PeriodEnd(int_cnt), "ACT/365")
    dbl_PT1 = 1
    dbl_PT2 = irc_Disc_LegA.Lookup_Rate(lng_ValDate, col_PeriodEnd(int_cnt), "DF", , , True)
    dbl_A = HWPDE.HWBond_A(dbl_PT1, dbl_PT2, dbl_T1, dbl_T2)
    dbl_B = HW_B(dbl_MeanRev, dbl_T1, dbl_T2)
    dbl_CalcPeriod = calc_yearfrac(col_PeriodStart(int_cnt), col_PeriodEnd(int_cnt), irl_Leg.Params.Daycount)
    
    If UCase(irl_Leg.Params.ExoticType) = "RANGE" Then
    
        'Setting the Fixing FixedRate or FloatRate
        dbl_Strike = irl_Leg.Params.Fixings(col_PeriodStart(int_cnt))
        
        'Setting the Range
        Dim str_AboveUpper As String, str_abovelower As String, str_RangeType As String
        Dim dbl_Upper As Double, dbl_Lower As Double
    
        If Not irl_Leg.Params.VariableRange1 Is Nothing Then
            str_AboveUpper = irl_Leg.Params.VariableRange1(col_PeriodStart(int_cnt))
            dbl_Upper = irl_Leg.Params.VariableRange2(col_PeriodStart(int_cnt))
            str_abovelower = irl_Leg.Params.VariableRange3(col_PeriodStart(int_cnt))
            dbl_Lower = irl_Leg.Params.VariableRange4(col_PeriodStart(int_cnt))
            str_RangeType = ird_Leg.GetRangeType(str_AboveUpper, str_abovelower)
        Else
            str_RangeType = ird_Leg.rangetype
            dbl_Upper = ird_Leg.Params.upper
            dbl_Lower = ird_Leg.Params.lower
        End If
        
        'Find the CalcPeriod(Prob) of the past dated
        Dim dbl_StoreProb As Double:  dbl_StoreProb = 0
        Dim item As Variant
        
        For Each item In ird_Leg.Params.FixingsDigi.Keys
        
            If item >= col_PeriodStart(int_cnt) And item < lng_ValDate Then

                dbl_Forward = ird_Leg.Params.FixingsDigi.item(item) / 100
                dbl_StoreProb = dbl_StoreProb + Payoff_Dependent_Smoothing(str_RangeType, dbl_Upper, dbl_Lower, dbl_Forward, 0, False)
                
            End If
        Next item

        dbl_StoreProb = dbl_StoreProb / (col_PeriodEnd(int_cnt) - col_PeriodStart(int_cnt))
        
        'Rate Fixing
        dbl_Strike = irl_Leg.Params.Fixings(col_PeriodStart(int_cnt))
        
        For int_j = 1 To UBound(arr_SpotStep)
            dbl_X = arr_SpotStep(int_j)
            dbl_HWBond = dbl_A * Exp(-dbl_B * dbl_X)
            
            arr_payoff(int_j, int_i) = (dbl_Strike / 100) * irl_Leg.Params.Notional * (dbl_StoreProb * dbl_CalcPeriod * dbl_HWBond + arr_CalcPeriod(1, int_j))

        Next int_j

    Else
        'if not range, take the fixings from the custom sheet
        For int_j = 1 To UBound(arr_SpotStep)
        
            dbl_X = arr_SpotStep(int_j)
            dbl_HWBond = dbl_A * Exp(-dbl_B * dbl_X)
            
            arr_payoff(int_j, int_i) = irl_Leg.Params.Fixings(col_PeriodStart(int_cnt)) / 100 * dbl_CalcPeriod * irl_Leg.Params.Notional * dbl_HWBond
        Next int_j

    End If
Exit For
End If

Next int_cnt

HW_Payoff_Time = arr_Payoff_Time
HW_Underlying = arr_payoff

If Erase_Switch = True Then
    Erase arr_SpotStep
    Erase arr_TimeStep
    Erase arr_Payoff_Time
    Erase arr_payoff
End If

End Function


Private Sub HW_Option()

Dim int_i As Integer
Dim int_j As Integer
Dim int_cnt1 As Integer

Dim dbl_T1 As Double
Dim dbl_T2 As Double
Dim dbl_PT1 As Double
Dim dbl_PT2 As Double
Dim dbl_A As Double

Dim arr_output() As Variant
Dim arr_UndVal() As Double
Dim arr_UndVal_A() As Double
Dim arr_UndVal_B() As Double

Dim arr_SpotStep() As Double
Dim arr_TimeStep() As Double
Dim arr_TimeLabel() As Integer
Dim arr_TimeStep_ImpDate() As Double

arr_SpotStep() = HWPDE.spotstep
arr_TimeStep() = HWPDE.timestep
arr_TimeLabel() = HWPDE.timelabel
arr_TimeStep_ImpDate() = HWPDE.TimeStep_ImpDate

ReDim arr_output(1 To UBound(arr_TimeStep), 1 To UBound(arr_SpotStep))
ReDim arr_UndVal(1 To UBound(arr_TimeStep), 1 To UBound(arr_SpotStep))
ReDim arr_PDE_UndVal(1 To UBound(arr_TimeStep), 1 To UBound(arr_SpotStep))
ReDim arr_PDE_Option(1 To UBound(arr_TimeStep), 1 To UBound(arr_SpotStep))

Dim arr_payoff_t_a() As Variant
Dim arr_payoff_t_b() As Variant
Dim x As Double

arr_UndVal_A() = HW_Underlying(irl_LegA, irl_LegADigi)
arr_payoff_t_a() = HW_Payoff_Time
arr_UndVal_B() = HW_Underlying(irl_legB, irl_LegBDigi)
arr_payoff_t_b() = HW_Payoff_Time

''Get Sigma
Dim arr_sigma() As Double
arr_sigma = HWPDE.HW_Sigma

''Smart Multiplication
Dim arr_SmartMult() As Double
ReDim arr_SmartMult(1 To UBound(arr_TimeStep)) As Double

arr_SmartMult(1) = 1

For int_i = UBound(arr_TimeStep) To 2 Step -1
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
    arr_SmartMult(int_i) = dbl_A
Next int_i


'Pass to Python in one step
arr_output = PDE(int_Direction, int_Sign, arr_SpotStep, arr_TimeStep, arr_TimeLabel, dbl_MeanRev, arr_sigma, arr_payoff_t_a, arr_UndVal_A, arr_payoff_t_b, arr_UndVal_B, arr_SmartMult, bln_Smooth)

For int_i = 0 To UBound(arr_TimeStep) - 1
    For int_j = 0 To UBound(arr_SpotStep) - 1
        arr_PDE_UndVal(int_i + 1, int_j + 1) = arr_output(0, int_i, int_j)
        arr_PDE_Option(int_i + 1, int_j + 1) = arr_output(1, int_i, int_j)
    Next int_j
Next int_i

dbl_PDE_MV = arr_PDE_UndVal(1, int_SpotStep + 1) + int_Direction * arr_PDE_Option(1, int_SpotStep + 1)

If Erase_Switch = True Then
    Erase arr_output
    Erase HW_Payoff_Time
    Erase arr_UndVal_A
    Erase arr_payoff_t_a
    Erase arr_UndVal_B
    Erase arr_payoff_t_b
End If

End Sub



Private Function StoreATMvols(irl_underlying As IRLeg, str_VolCurve As String, bln_DataExists As Boolean, bln_AddIfMissing As Boolean) As Collection

    ' ## Store vol values, store cap vol as each caplet vol if curve setting is to specify cap vols
    
    Dim lngLst_PeriodStart As Collection: Set lngLst_PeriodStart = irl_underlying.periodstart
    Dim dblLst_CapletVols As New Collection
    Dim int_NumPeriods As Integer: int_NumPeriods = lngLst_PeriodStart.count
    
    Dim cvl_volcurve As Data_CapVolsQJK
    
    Dim dbl_CapVol As Double
    Dim dbl_ATMRates As Double
    
    Dim int_ctr As Integer
    Dim bln_Bootstrappable As Boolean
    
    For int_ctr = 1 To int_NumPeriods
    
        dbl_ATMRates = irl_underlying.GetRates(int_ctr)

        Set cvl_volcurve = GetObject_CapVolSurf(str_VolCurve, dbl_ATMRates, bln_DataExists, bln_AddIfMissing)
        bln_Bootstrappable = cvl_volcurve.IsBootstrappable
    
        If bln_Bootstrappable = True Then
            If lngLst_PeriodStart(int_ctr) > lng_ValDate Then
                Call dblLst_CapletVols.Add(cvl_volcurve.Lookup_Vol(lngLst_PeriodStart(int_ctr), , True))
            Else
                Call dblLst_CapletVols.Add(0)
            End If
        Else
            
            Call dblLst_CapletVols.Add(dbl_CapVol)
        End If
    Next int_ctr

Set StoreATMvols = dblLst_CapletVols

End Function

Private Function StoreVols(irl_underlying As IRLeg, cvl_volcurve As Data_CapVolsQJK) As Collection
    ' ## Store vol values, store cap vol as each caplet vol if curve setting is to specify cap vols
    Dim lngLst_PeriodStart As Collection: Set lngLst_PeriodStart = irl_underlying.periodstart
    Dim dblLst_CapletVols As Collection
    Set dblLst_CapletVols = New Collection
    Dim int_NumPeriods As Integer: int_NumPeriods = lngLst_PeriodStart.count
    Dim dbl_CapVol As Double
    Dim bln_Bootstrappable As Boolean: bln_Bootstrappable = cvl_volcurve.IsBootstrappable
    
    If bln_Bootstrappable = False Then dbl_CapVol = cvl_volcurve.Lookup_Vol(irl_underlying.PeriodEnd(int_NumPeriods), , False)

    Dim int_ctr As Integer
    For int_ctr = 1 To int_NumPeriods
        If bln_Bootstrappable = True Then
            If lngLst_PeriodStart(int_ctr) > lng_ValDate Then
                 Call dblLst_CapletVols.Add(cvl_volcurve.Lookup_Vol(lngLst_PeriodStart(int_ctr), , False))
            Else
                Call dblLst_CapletVols.Add(0)
            End If
        Else
            Call dblLst_CapletVols.Add(dbl_CapVol)
        End If
    Next int_ctr
    
StoreVols = dblLst_CapletVols
End Function

Private Function StoreVolsDigi(irl_underlying As IRLegDigi, cvl_volcurve As Data_CapVolsQJK) As Collection
    ' ## Store vol values, store cap vol as each caplet vol if curve setting is to specify cap vols
    Dim lngLst_PeriodStart As Collection: Set lngLst_PeriodStart = irl_underlying.periodstart
    Dim dblLst_CapletVols As Collection
    Set dblLst_CapletVols = New Collection
    Dim int_NumPeriods As Integer: int_NumPeriods = lngLst_PeriodStart.count
    Dim dbl_CapVol As Double
    Dim bln_Bootstrappable As Boolean: bln_Bootstrappable = cvl_volcurve.IsBootstrappable
    If bln_Bootstrappable = False Then dbl_CapVol = cvl_volcurve.Lookup_Vol(irl_underlying.PeriodEnd(int_NumPeriods), , False)
    
    Dim int_ctr As Integer
    For int_ctr = 1 To int_NumPeriods
        If bln_Bootstrappable = True Then
            If lngLst_PeriodStart(int_ctr) > lng_ValDate Then
                Call dblLst_CapletVols.Add(cvl_volcurve.Lookup_Vol(lngLst_PeriodStart(int_ctr), , False))
            Else
                Call dblLst_CapletVols.Add(0)
            End If
        Else
            Call dblLst_CapletVols.Add(dbl_CapVol)
        End If
    Next int_ctr
    
Set StoreVolsDigi = dblLst_CapletVols
End Function

' ## PROPERTIES
Public Property Get marketvalue() As Double
    marketvalue = CalcValue("MV")
End Property

Public Property Get Cash() As Double
    Cash = CalcValue("CASH")
End Property

Public Property Get PnL() As Double
    PnL = Me.marketvalue + Me.Cash
End Property

' ## METHODS - PRIVATE
Private Function GetFXConvFactor() As Double
    ' ## Get factor to convert from the native currency to the PnL reporting currency
    GetFXConvFactor = fxs_Spots.Lookup_DiscSpot(irl_LegA.Params.CCY, fld_Params.CCY_PnL)

End Function
Private Function CalcValue(str_type As String) As Double

Dim dbl_Output As Double
Dim dbl_legA As Double
Dim dbl_legB As Double

Select Case UCase(fld_Params.IsCallable)
    Case "TRUE"
    
        Select Case UCase(str_type)
            Dim x As Double
           
            Case "MV"
        
                dbl_Output = dbl_PDE_MV
                
                Erase arr_PDE_UndVal
                Erase arr_PDE_Option
                Erase arr_PDE_CalcPeriod
                
            Case "CASH"
                dbl_Output = 0
            
        End Select
        
        dbl_Output = dbl_Output
    
    Case "FALSE"
        Select Case UCase(fld_Params.LegA.ExoticType)
            Case "RANGE"
                dbl_legA = CalcValue_digi(str_type, irl_LegA, irl_LegADigi, dblLst_VolCurve_LegADigi_upper, dblLst_VolCurve_LegADigi_upperUp, _
                            dblLst_VolCurve_LegADigi_lower, dblLst_VolCurve_LegADigi_lowerUp, dblLst_ATMVolCurve_LegA)
            Case "-"
                dbl_legA = irl_LegA.CalcValue(str_type)
        End Select
        
        Select Case UCase(fld_Params.LegB.ExoticType)
            Case "RANGE"
                dbl_legB = CalcValue_digi(str_type, irl_legB, irl_LegBDigi, dblLst_VolCurve_LegBDigi_upper, dblLst_VolCurve_LegBDigi_upperUp, _
                            dblLst_VolCurve_LegBDigi_lower, dblLst_VolCurve_LegBDigi_lowerUp, dblLst_ATMVolCurve_LegB)
            Case "-"
                dbl_legB = irl_legB.CalcValue(str_type)
        End Select
        
    dbl_Output = int_Sign * (dbl_legA - dbl_legB)
    
End Select

CalcValue = dbl_Output * GetFXConvFactor()
    
End Function
Private Function CalcValue_digi(str_type As String, irl_Leg As IRLeg, irl_LegDigi As IRLegDigi, _
                            Optional dblLst_UpperVol As Collection, Optional dblLst_ShiftedUpperVol As Collection, _
                            Optional dblLst_LowerVol As Collection, Optional dblLst_ShiftedLowerVol As Collection, _
                            Optional dblLst_ATMVol As Collection) As Double

Dim dbl_Output As Double
Dim dbl_Upper As Double
Dim dbl_Lower As Double
 
Dim dbl_strike_upper As Double: dbl_strike_upper = irl_LegDigi.dbl_Upper
Dim dbl_strike_lower As Double: dbl_strike_lower = irl_LegDigi.dbl_Lower

Dim dblLst_StoreProb As Collection
Dim dblLst_StoreProb_Corr As Collection

Select Case irl_LegDigi.rangetype

    Case "AboveLower"
        Set dblLst_StoreProb = irl_LegDigi.StoreProb(irl_Leg, dbl_strike_lower, dblLst_LowerVol, dblLst_ShiftedLowerVol)
        Set dblLst_StoreProb_Corr = irl_LegDigi.StoreProb(irl_Leg, dbl_strike_lower, dblLst_LowerVol, dblLst_ShiftedLowerVol, dblLst_ATMVol)
        dbl_Lower = irl_Leg.CalcValue_RA(str_type, dblLst_StoreProb, dblLst_StoreProb_Corr)
        
        dbl_Output = dbl_Lower
        
    Case "Outside"
        Set dblLst_StoreProb = irl_LegDigi.StoreProb(irl_Leg, dbl_strike_upper, dblLst_UpperVol, dblLst_ShiftedUpperVol)
        Set dblLst_StoreProb_Corr = irl_LegDigi.StoreProb(irl_Leg, dbl_strike_upper, dblLst_UpperVol, dblLst_ShiftedUpperVol, dblLst_ATMVol)
        dbl_Upper = irl_Leg.CalcValue_RA(str_type, dblLst_StoreProb, dblLst_StoreProb_Corr)
        
        Set dblLst_StoreProb = irl_LegDigi.StoreProb(irl_Leg, dbl_strike_lower, dblLst_LowerVol, dblLst_ShiftedLowerVol)
        Set dblLst_StoreProb_Corr = irl_LegDigi.StoreProb(irl_Leg, dbl_strike_lower, dblLst_LowerVol, dblLst_ShiftedLowerVol, dblLst_ATMVol, PutOpt)
        dbl_Lower = irl_Leg.CalcValue_RA(str_type, dblLst_StoreProb, dblLst_StoreProb_Corr)
        
        dbl_Output = dbl_Lower + dbl_Upper
        
    Case "Between"
        Set dblLst_StoreProb = irl_LegDigi.StoreProb(irl_Leg, dbl_strike_upper, dblLst_UpperVol, dblLst_ShiftedUpperVol)
        Set dblLst_StoreProb_Corr = irl_LegDigi.StoreProb(irl_Leg, dbl_strike_upper, dblLst_UpperVol, dblLst_ShiftedUpperVol, dblLst_ATMVol)
        dbl_Upper = irl_Leg.CalcValue_RA(str_type, dblLst_StoreProb, dblLst_StoreProb_Corr)
        
        Set dblLst_StoreProb = irl_LegDigi.StoreProb(irl_Leg, dbl_strike_lower, dblLst_LowerVol, dblLst_ShiftedLowerVol)
        Set dblLst_StoreProb_Corr = irl_LegDigi.StoreProb(irl_Leg, dbl_strike_lower, dblLst_LowerVol, dblLst_ShiftedLowerVol, dblLst_ATMVol)
        dbl_Lower = irl_Leg.CalcValue_RA(str_type, dblLst_StoreProb, dblLst_StoreProb_Corr)
        
        dbl_Output = dbl_Lower - dbl_Upper
        
    Case "BelowUpper"
        Set dblLst_StoreProb = irl_LegDigi.StoreProb(irl_Leg, dbl_strike_upper, dblLst_UpperVol, dblLst_ShiftedUpperVol)
        Set dblLst_StoreProb_Corr = irl_LegDigi.StoreProb(irl_Leg, dbl_strike_upper, dblLst_UpperVol, dblLst_ShiftedUpperVol, dblLst_ATMVol, PutOpt)
        dbl_Upper = irl_Leg.CalcValue_RA(str_type, dblLst_StoreProb, dblLst_StoreProb_Corr)
        
        dbl_Output = dbl_Upper
        
    End Select

CalcValue_digi = dbl_Output

End Function
' ## METHODS - CALCULATION DETAILS
Public Sub OutputReport(wks_output As Worksheet)
    wks_output.Cells.Clear
    
    Dim rng_ActiveOutput As Range: Set rng_ActiveOutput = wks_output.Range("A1")
    Dim int_ActiveRow As Integer: int_ActiveRow = 0
    Dim int_ActiveColumn As Integer: int_ActiveColumn = 0
    
    Dim rng_PnL As Range, rng_OptionDF As Range
    Dim int_ctr As Integer
    
    Dim dic_Addresses As Dictionary: Set dic_Addresses = New Dictionary
    dic_Addresses.CompareMode = CompareMethod.TextCompare
     
        'KL - For CRA
        Dim int_i As Integer
        Dim int_j As Integer
        
        Dim arr_SpotStep() As Double
        Dim arr_TimeStep() As Double
        
        arr_SpotStep() = HWPDE.spotstep
        arr_TimeStep() = HWPDE.timestep
        
        'To show how many SpotStep around central points
        Dim int_LBound As Integer
        Dim int_UBound As Integer
        
        int_LBound = int_SpotStep - 2
        int_UBound = int_SpotStep + 4
        
        'Call HW_Option
        'Populate Calibrated vol
        rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn).Value = "Calibrated Sigma"
        
        For int_i = 1 To HW_vol.FullFinalHwVol.count
            int_ActiveRow = int_ActiveRow + 1
            With rng_ActiveOutput
                .Offset(int_ActiveRow, int_ActiveColumn).Value = HW_vol.FullFinalHwVol.Keys(int_i - 1)
                .Offset(int_ActiveRow, int_ActiveColumn + 1).Value = HW_vol.FullFinalHwVol.Items(int_i - 1)
            End With

        Next int_i
        
        'For Calc Period
        int_ActiveRow = int_ActiveRow + 2
        rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn).Value = "CalcPeriod"
        int_ActiveRow = int_ActiveRow + 2
        rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn).Value = "TimeStep/SpotStep"

        ''''For Calc Period''''
        'Populate SpotStep
        int_ActiveRow = int_ActiveRow
        int_ActiveColumn = 1
        For int_j = int_LBound To int_UBound
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

        'Populate Calc Period
        int_ActiveRow = int_ActiveRow - UBound(arr_TimeStep)
        int_ActiveColumn = 1
        For int_i = UBound(arr_TimeStep) To 1 Step -1
            For int_j = int_LBound To int_UBound
                With rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn)
                    .Value = arr_PDE_CalcPeriod(int_i, int_j)
                    .NumberFormat = "0.00%"
                End With
                int_ActiveColumn = int_ActiveColumn + 1
            Next int_j
            int_ActiveRow = int_ActiveRow + 1
            int_ActiveColumn = 1
        Next int_i

        int_ActiveColumn = 0
        int_ActiveRow = int_ActiveRow + 2
        rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn).Value = "Option Value"
        int_ActiveRow = int_ActiveRow + 2
        rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn).Value = "TimeStep/SpotStep"

        ''''For Option Value''''
        'Populate SpotStep
        int_ActiveRow = int_ActiveRow
        int_ActiveColumn = 1
        For int_j = int_LBound To int_UBound
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
            For int_j = int_LBound To int_UBound
                With rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn)
                    .Value = arr_PDE_Option(int_i, int_j)
                    .Style = "Comma"
                End With
                int_ActiveColumn = int_ActiveColumn + 1
            Next int_j
            int_ActiveRow = int_ActiveRow + 1
            int_ActiveColumn = 1
        Next int_i
        
'        ''''For Underlying Value''''
        int_ActiveRow = int_ActiveRow + 2
        int_ActiveColumn = 0
        rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn).Value = "Underlying Value"
        int_ActiveRow = int_ActiveRow + 2
        rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn).Value = "TimeStep/SpotStep"

        'Populate SpotStep
        int_ActiveRow = int_ActiveRow
        int_ActiveColumn = 1
        For int_j = int_LBound To int_UBound
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

        'Populate Underlying Value
        int_ActiveRow = int_ActiveRow - UBound(arr_TimeStep)
        int_ActiveColumn = 1

        For int_i = UBound(arr_TimeStep) To 1 Step -1
            For int_j = int_LBound To int_UBound
                With rng_ActiveOutput.Offset(int_ActiveRow, int_ActiveColumn)
                    .Value = arr_PDE_UndVal(int_i, int_j)
                    .Style = "Comma"
                End With
                int_ActiveColumn = int_ActiveColumn + 1
            Next int_j
            int_ActiveRow = int_ActiveRow + 1
            int_ActiveColumn = 1
        Next int_i

    wks_output.Calculate
    wks_output.Columns.AutoFit
    wks_output.Cells.HorizontalAlignment = xlCenter
End Sub
