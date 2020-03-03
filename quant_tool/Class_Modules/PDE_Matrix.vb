VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PDE_Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum PDE_TimeStep
    major = 1
    important
    CN 'CrankNicolson
    Implicit
End Enum

Public Enum PDE_Product
    CRA = 1
    BSWAP = 2
End Enum

' ## MEMBER DATA
''''''''''''''''START - KL - For Testing'''''''''''''''''
Private lng_MatDate As Long 'Valuation Date
Private lng_ValDate As Long 'Valuation Date
Private lng_LastCallDate As Long 'Last Call Date
Private arr_sigma() As Double 'Vol Series
Private arr_ImpDate() As Double
Private arr_CallDate() As Double
Private irl_LegA As IRLeg
Private bln_BusinessDay As Boolean

' Hull White Setting
Private int_SpotStep As Integer
Private int_TimeStep As Integer
Private dbl_MeanRev As Double

'Output
Private arr_SpotStep() As Double
Private arr_TimeStep() As Double
Private arr_TimeLabel() As Integer
Private arr_TimeStep_ImpDate() As Double
''''''''''''''''END - KL - For Testing'''''''''''''''''

Public Property Get spotstep() As Double()
    spotstep = arr_SpotStep
End Property
Public Property Get timestep() As Double()
    timestep = arr_TimeStep
End Property
Public Property Get timelabel() As Integer()
    timelabel = arr_TimeLabel
End Property
Public Property Get TimeStep_ImpDate() As Double()
    TimeStep_ImpDate = arr_TimeStep_ImpDate
End Property
Public Property Get HW_Sigma() As Double()
    HW_Sigma = arr_sigma()
End Property
Public Sub Initialize(ValDate As Long, MatDate As Long, CalibratedVol As Dictionary, LegA As IRLeg, LegB As IRLeg, _
                      spotstep As Integer, timestep As Integer, MeanRev As Double, ProductType As PDE_Product)

Dim int_i As Integer
Dim dbl_T1 As Double
Dim dic_ImpDate As New Dictionary

int_SpotStep = spotstep
int_TimeStep = timestep
dbl_MeanRev = MeanRev

lng_ValDate = ValDate 'Valuation Date
lng_LastCallDate = CalibratedVol.Keys(CalibratedVol.count - 1) 'Last Call Date
lng_MatDate = MatDate
 
'Vol Series
ReDim arr_sigma(1 To CalibratedVol.count + 2, 1 To 2)
ReDim arr_CallDate(1 To CalibratedVol.count)

arr_sigma(1, 1) = CalibratedVol.Items(0)
arr_sigma(1, 2) = calc_yearfrac(lng_ValDate, lng_ValDate, "ACT/365")

For int_i = 0 To CalibratedVol.count - 1

    'Sigma Series
    arr_sigma(int_i + 2, 1) = CalibratedVol.Items(int_i)
    arr_sigma(int_i + 2, 2) = calc_yearfrac(lng_ValDate, CalibratedVol.Keys(int_i), "ACT/365")
    
    'Call Date
    arr_CallDate(int_i + 1) = calc_yearfrac(lng_ValDate, CalibratedVol.Keys(int_i), "ACT/365")
    
Next int_i

'To include Start Date and End Date(Beyond MatDate) for sigma
arr_sigma(CalibratedVol.count + 2, 1) = CalibratedVol.Items(CalibratedVol.count - 1)
arr_sigma(CalibratedVol.count + 2, 2) = calc_yearfrac(lng_ValDate, lng_MatDate + 365, "ACT/365")

'Important Date
For int_i = 1 To LegA.periodstart.count
    dbl_T1 = calc_yearfrac(lng_ValDate, LegA.periodstart(int_i), "ACT/365")
    If dbl_T1 > 0 Then
        dic_ImpDate.Add dbl_T1, vbNull
    End If
Next int_i

For int_i = 1 To LegA.PeriodEnd.count
    dbl_T1 = calc_yearfrac(lng_ValDate, LegA.PeriodEnd(int_i), "ACT/365")
    If dic_ImpDate.Exists(dbl_T1) = False And dbl_T1 > 0 Then
        dic_ImpDate.Add dbl_T1, vbNull
    End If
Next int_i

For int_i = 1 To LegB.periodstart.count
    dbl_T1 = calc_yearfrac(lng_ValDate, LegB.periodstart(int_i), "ACT/365")
    If dic_ImpDate.Exists(dbl_T1) = False And dbl_T1 > 0 Then
        dic_ImpDate.Add dbl_T1, vbNull
    End If
Next int_i

For int_i = 1 To LegB.PeriodEnd.count
    dbl_T1 = calc_yearfrac(lng_ValDate, LegB.PeriodEnd(int_i), "ACT/365")
    If dic_ImpDate.Exists(dbl_T1) = False And dbl_T1 > 0 Then
        dic_ImpDate.Add dbl_T1, vbNull
    End If
Next int_i

'Remove time = 0
If dic_ImpDate.Exists(0) Then: dic_ImpDate.Remove (0)

ReDim arr_ImpDate(1 To dic_ImpDate.count)

For int_i = 1 To dic_ImpDate.count
    arr_ImpDate(int_i) = dic_ImpDate.Keys(int_i - 1)
Next int_i

Call BubbleSort_ImpDate(arr_ImpDate)
Call CreateSpotStep

'Revised for Business TimeStep
Set irl_LegA = LegA

If UCase(LegA.Params.NbofDays) = "BUSINESS DAYS" Or UCase(LegB.Params.NbofDays) = "BUSINESS DAYS" Then
    bln_BusinessDay = True
Else: bln_BusinessDay = False
End If

Select Case ProductType
    Case PDE_Product.CRA: CreateTimeStep_CRA
    Case PDE_Product.BSWAP: CreateTimeStep_BSWAP
End Select

Call CreateTimeStep_ImpDate

'Overwritting TimeStep
'Call OverwriteTimeStep_4
'Call OverwriteTimeStep_30

End Sub
Private Sub OverwriteTimeStep_4()
 
ReDim arr_TimeStep(1 To 44) As Double
ReDim arr_TimeLabel(1 To 44) As Integer

arr_TimeStep(44) = 1
arr_TimeStep(43) = 0.8
arr_TimeStep(42) = (43628 - 43355) / 365
arr_TimeStep(41) = 0.697717
arr_TimeStep(40) = 0.647489
arr_TimeStep(39) = 0.6
arr_TimeStep(38) = (43536 - 43355) / 365
arr_TimeStep(37) = 0.445662
arr_TimeStep(36) = 0.4
arr_TimeStep(35) = 0.349772
arr_TimeStep(34) = (43446 - 43355) / 365
arr_TimeStep(33) = 0.2
arr_TimeStep(32) = 0.0821918
arr_TimeStep(31) = 0.0794521
arr_TimeStep(30) = 0.0767123
arr_TimeStep(29) = 0.0739726
arr_TimeStep(28) = 0.0712329
arr_TimeStep(27) = 0.0684932
arr_TimeStep(26) = 0.0657534
arr_TimeStep(25) = 0.0630137
arr_TimeStep(24) = 0.060274
arr_TimeStep(23) = 0.0575342
arr_TimeStep(22) = 0.0547945
arr_TimeStep(21) = 0.0520548
arr_TimeStep(20) = 0.0493151
arr_TimeStep(19) = 0.0465753
arr_TimeStep(18) = 0.0438356
arr_TimeStep(17) = 0.0410959
arr_TimeStep(16) = 0.0383562
arr_TimeStep(15) = 0.0356164
arr_TimeStep(14) = 0.0328767
arr_TimeStep(13) = 0.030137
arr_TimeStep(12) = 0.0273973
arr_TimeStep(11) = 0.0246575
arr_TimeStep(10) = 0.0219178
arr_TimeStep(9) = 0.0191781
arr_TimeStep(8) = 0.0164384
arr_TimeStep(7) = 0.0136986
arr_TimeStep(6) = 0.0109589
arr_TimeStep(5) = 0.00821918
arr_TimeStep(4) = 0.00547945
arr_TimeStep(3) = 0.00273973
arr_TimeStep(2) = 0.00001
arr_TimeStep(1) = 0

arr_TimeLabel(44) = 2
arr_TimeLabel(43) = 3
arr_TimeLabel(42) = 1
arr_TimeLabel(41) = 4
arr_TimeLabel(40) = 4
arr_TimeLabel(39) = 3
arr_TimeLabel(38) = 1
arr_TimeLabel(37) = 4
arr_TimeLabel(36) = 4
arr_TimeLabel(35) = 3
arr_TimeLabel(34) = 2
arr_TimeLabel(33) = 3
arr_TimeLabel(32) = 3
arr_TimeLabel(31) = 3
arr_TimeLabel(30) = 3
arr_TimeLabel(29) = 3
arr_TimeLabel(28) = 3
arr_TimeLabel(27) = 3
arr_TimeLabel(26) = 3
arr_TimeLabel(25) = 3
arr_TimeLabel(24) = 3
arr_TimeLabel(23) = 3
arr_TimeLabel(22) = 3
arr_TimeLabel(21) = 3
arr_TimeLabel(20) = 3
arr_TimeLabel(19) = 3
arr_TimeLabel(18) = 3
arr_TimeLabel(17) = 3
arr_TimeLabel(16) = 3
arr_TimeLabel(15) = 3
arr_TimeLabel(14) = 3
arr_TimeLabel(13) = 3
arr_TimeLabel(12) = 3
arr_TimeLabel(11) = 3
arr_TimeLabel(10) = 3
arr_TimeLabel(9) = 3
arr_TimeLabel(8) = 3
arr_TimeLabel(7) = 3
arr_TimeLabel(6) = 3
arr_TimeLabel(5) = 3
arr_TimeLabel(4) = 3
arr_TimeLabel(3) = 3
arr_TimeLabel(2) = 3
arr_TimeLabel(1) = 3

End Sub

Private Sub OverwriteTimeStep_30()
 
ReDim arr_TimeStep(1 To 68) As Double
ReDim arr_TimeLabel(1 To 68) As Integer

arr_TimeStep(68) = 1
arr_TimeStep(67) = 0.986301
arr_TimeStep(66) = 0.953425
arr_TimeStep(65) = 0.920548
arr_TimeStep(64) = 0.887671
arr_TimeStep(63) = 0.854795
arr_TimeStep(62) = 0.821918
arr_TimeStep(61) = 0.789041
arr_TimeStep(60) = 0.756164
arr_TimeStep(59) = (43628 - 43355) / 365
arr_TimeStep(58) = 0.736986
arr_TimeStep(57) = 0.726027
arr_TimeStep(56) = 0.723288
arr_TimeStep(55) = 0.690411
arr_TimeStep(54) = 0.657534
arr_TimeStep(53) = 0.624658
arr_TimeStep(52) = 0.591781
arr_TimeStep(51) = 0.558904
arr_TimeStep(50) = 0.526027
arr_TimeStep(49) = (43536 - 43355) / 365
arr_TimeStep(48) = 0.493151
arr_TimeStep(47) = 0.482192
arr_TimeStep(46) = 0.471233
arr_TimeStep(45) = 0.460274
arr_TimeStep(44) = 0.427397
arr_TimeStep(43) = 0.394521
arr_TimeStep(42) = 0.361644
arr_TimeStep(41) = 0.328767
arr_TimeStep(40) = 0.29589
arr_TimeStep(39) = 0.263014
arr_TimeStep(38) = (43446 - 43355) / 365
arr_TimeStep(37) = 0.230137
arr_TimeStep(36) = 0.19726
arr_TimeStep(35) = 0.164384
arr_TimeStep(34) = 0.131507
arr_TimeStep(33) = 0.0986301
arr_TimeStep(32) = 0.0821918
arr_TimeStep(31) = 0.0794521
arr_TimeStep(30) = 0.0767123
arr_TimeStep(29) = 0.0739726
arr_TimeStep(28) = 0.0712329
arr_TimeStep(27) = 0.0684932
arr_TimeStep(26) = 0.0657534
arr_TimeStep(25) = 0.0630137
arr_TimeStep(24) = 0.060274
arr_TimeStep(23) = 0.0575342
arr_TimeStep(22) = 0.0547945
arr_TimeStep(21) = 0.0520548
arr_TimeStep(20) = 0.0493151
arr_TimeStep(19) = 0.0465753
arr_TimeStep(18) = 0.0438356
arr_TimeStep(17) = 0.0410959
arr_TimeStep(16) = 0.0383562
arr_TimeStep(15) = 0.0356164
arr_TimeStep(14) = 0.0328767
arr_TimeStep(13) = 0.030137
arr_TimeStep(12) = 0.0273973
arr_TimeStep(11) = 0.0246575
arr_TimeStep(10) = 0.0219178
arr_TimeStep(9) = 0.0191781
arr_TimeStep(8) = 0.0164384
arr_TimeStep(7) = 0.0136986
arr_TimeStep(6) = 0.0109589
arr_TimeStep(5) = 0.00821918
arr_TimeStep(4) = 0.00547945
arr_TimeStep(3) = 0.00273973
arr_TimeStep(2) = 0.00001
arr_TimeStep(1) = 0

arr_TimeLabel(68) = 2
arr_TimeLabel(67) = 3
arr_TimeLabel(66) = 3
arr_TimeLabel(65) = 3
arr_TimeLabel(64) = 3
arr_TimeLabel(63) = 3
arr_TimeLabel(62) = 3
arr_TimeLabel(61) = 3
arr_TimeLabel(60) = 3
arr_TimeLabel(59) = 1
arr_TimeLabel(58) = 4
arr_TimeLabel(57) = 4
arr_TimeLabel(56) = 3
arr_TimeLabel(55) = 3
arr_TimeLabel(54) = 3
arr_TimeLabel(53) = 3
arr_TimeLabel(52) = 3
arr_TimeLabel(51) = 3
arr_TimeLabel(50) = 3
arr_TimeLabel(49) = 1
arr_TimeLabel(48) = 4
arr_TimeLabel(47) = 4
arr_TimeLabel(46) = 3
arr_TimeLabel(45) = 3
arr_TimeLabel(44) = 3
arr_TimeLabel(43) = 3
arr_TimeLabel(42) = 3
arr_TimeLabel(41) = 3
arr_TimeLabel(40) = 3
arr_TimeLabel(39) = 3
arr_TimeLabel(38) = 2
arr_TimeLabel(37) = 3
arr_TimeLabel(36) = 3
arr_TimeLabel(35) = 3
arr_TimeLabel(34) = 3
arr_TimeLabel(33) = 3
arr_TimeLabel(32) = 3
arr_TimeLabel(31) = 3
arr_TimeLabel(30) = 3
arr_TimeLabel(29) = 3
arr_TimeLabel(28) = 3
arr_TimeLabel(27) = 3
arr_TimeLabel(26) = 3
arr_TimeLabel(25) = 3
arr_TimeLabel(24) = 3
arr_TimeLabel(23) = 3
arr_TimeLabel(22) = 3
arr_TimeLabel(21) = 3
arr_TimeLabel(20) = 3
arr_TimeLabel(19) = 3
arr_TimeLabel(18) = 3
arr_TimeLabel(17) = 3
arr_TimeLabel(16) = 3
arr_TimeLabel(15) = 3
arr_TimeLabel(14) = 3
arr_TimeLabel(13) = 3
arr_TimeLabel(12) = 3
arr_TimeLabel(11) = 3
arr_TimeLabel(10) = 3
arr_TimeLabel(9) = 3
arr_TimeLabel(8) = 3
arr_TimeLabel(7) = 3
arr_TimeLabel(6) = 3
arr_TimeLabel(5) = 3
arr_TimeLabel(4) = 3
arr_TimeLabel(3) = 3
arr_TimeLabel(2) = 3
arr_TimeLabel(1) = 3

End Sub


Public Function HWBond_A(dbl_PT1 As Double, dbl_PT2 As Double, dbl_T1 As Double, dbl_T2 As Double) As Double

Dim int_cnt As Integer
Dim int_max As Integer

Dim dbl_Output As Double
Dim dbl_output_V_T1 As Double
Dim dbl_output_V_T2 As Double

Dim dbl_A As Double
Dim dbl_B As Double
Dim dbl_C As Double

Dim dbl_T1_loop As Double
Dim dbl_T2_Loop As Double
Dim dbl_Sigma As Double

For int_cnt = 1 To UBound(arr_sigma)
    If arr_sigma(int_cnt, 2) > dbl_T1 Then
        int_max = int_cnt - 1
        Exit For
    End If
Next int_cnt

For int_cnt = 1 To int_max - 1

    dbl_T1_loop = arr_sigma(int_cnt, 2)
    dbl_T2_Loop = arr_sigma(int_cnt + 1, 2)
    dbl_Sigma = arr_sigma(int_cnt + 1, 1)
    
    dbl_A = dbl_T2_Loop - dbl_T1_loop
    dbl_B = 1 / (2 * dbl_MeanRev) * (Exp(2 * dbl_MeanRev * dbl_T2_Loop) - Exp(2 * dbl_MeanRev * dbl_T1_loop))
    dbl_C = 2 / (dbl_MeanRev) * (Exp(dbl_MeanRev * dbl_T2_Loop) - Exp(dbl_MeanRev * dbl_T1_loop))
    
    dbl_output_V_T1 = dbl_output_V_T1 + dbl_Sigma * dbl_Sigma / dbl_MeanRev / dbl_MeanRev * (dbl_A + Exp(-2 * dbl_MeanRev * dbl_T1) * dbl_B - Exp(-dbl_MeanRev * dbl_T1) * dbl_C)
    dbl_output_V_T2 = dbl_output_V_T2 + dbl_Sigma * dbl_Sigma / dbl_MeanRev / dbl_MeanRev * (dbl_A + Exp(-2 * dbl_MeanRev * dbl_T2) * dbl_B - Exp(-dbl_MeanRev * dbl_T2) * dbl_C)
    
Next int_cnt

dbl_T1_loop = arr_sigma(int_max, 2)
dbl_T2_Loop = dbl_T1
dbl_Sigma = arr_sigma(int_max + 1, 1)

dbl_A = dbl_T2_Loop - dbl_T1_loop
dbl_B = 1 / (2 * dbl_MeanRev) * (Exp(2 * dbl_MeanRev * dbl_T2_Loop) - Exp(2 * dbl_MeanRev * dbl_T1_loop))
dbl_C = 2 / (dbl_MeanRev) * (Exp(dbl_MeanRev * dbl_T2_Loop) - Exp(dbl_MeanRev * dbl_T1_loop))

dbl_output_V_T1 = dbl_output_V_T1 + dbl_Sigma * dbl_Sigma / dbl_MeanRev / dbl_MeanRev * (dbl_A + Exp(-2 * dbl_MeanRev * dbl_T1) * dbl_B - Exp(-dbl_MeanRev * dbl_T1) * dbl_C)
dbl_output_V_T2 = dbl_output_V_T2 + dbl_Sigma * dbl_Sigma / dbl_MeanRev / dbl_MeanRev * (dbl_A + Exp(-2 * dbl_MeanRev * dbl_T2) * dbl_B - Exp(-dbl_MeanRev * dbl_T2) * dbl_C)

dbl_Output = dbl_PT2 / dbl_PT1 * Exp((dbl_output_V_T1 - dbl_output_V_T2) / 2)
HWBond_A = dbl_Output

End Function

Private Sub CreateSpotStep()

Dim int_cnt As Integer

Dim dbl_T1 As Double
Dim dbl_T2 As Double
Dim dbl_MaxVar As Double
Dim dbl_V1 As Double
Dim dbl_SM As Double
Dim dbl_SpotScale As Double

Dim arr_output() As Double
Dim arr_MaxVol() As Double

'Construct the array of variance
ReDim arr_MaxVol(1 To UBound(arr_sigma) - 1)

For int_cnt = 1 To UBound(arr_sigma) - 1
    
    dbl_T1 = arr_sigma(int_cnt, 2)
    dbl_T2 = arr_sigma(int_cnt + 1, 2)
    
    If int_cnt = UBound(arr_sigma) - 1 Then
        dbl_T2 = (lng_MatDate - lng_ValDate) / 365
    End If
    
    If int_cnt = UBound(arr_sigma) - 1 And UBound(arr_sigma) < 1 Then
        dbl_T2 = 1
    End If
    
    dbl_SM = arr_sigma(int_cnt + 1, 1) ^ 2 / 2 / dbl_MeanRev
    dbl_V1 = (dbl_V1 - dbl_SM) * Exp(-2 * (dbl_T2 - dbl_T1) * dbl_MeanRev) + dbl_SM
    
    arr_MaxVol(int_cnt) = dbl_V1

Next int_cnt

'Define the max variance and spotscale
dbl_MaxVar = Application.WorksheetFunction.Max(arr_MaxVol)
dbl_SpotScale = Sqr(dbl_MaxVar) * 5 / int_SpotStep

'''''Create Spot Step Series'''''
ReDim arr_output(1 To int_SpotStep * 2 + 1)

For int_cnt = 1 To int_SpotStep * 2 + 1
    arr_output(int_cnt) = -dbl_SpotScale * (int_SpotStep + 1 - int_cnt)
Next int_cnt

arr_SpotStep = arr_output

End Sub

Private Sub CreateTimeStep_BSWAP()

Dim int_i As Integer
Dim int_j As Integer
Dim int_k As Integer

Dim int_N_iteration As Integer
Dim dbl_IteStep As Double
Dim dbl_IteTemp As Double
Dim dbl_T1 As Double
Dim dbl_T2 As Double
Dim dbl_Mat As Double
Dim dbl_EffStep As Double
Dim dbl_Rann As Double

Dim arr_OutputDate() As Double
Dim arr_OutputLabel() As Integer

''''''''''''''''''''''''''''''''''''''
''''Step 1 - Defined Important Date''''
''''''''''''''''''''''''''''''''''''''
arr_OutputDate = arr_ImpDate
ReDim arr_OutputLabel(1 To UBound(arr_OutputDate))

For int_i = 1 To UBound(arr_ImpDate)
    For int_j = 1 To UBound(arr_CallDate)
        If arr_CallDate(int_j) = arr_ImpDate(int_i) Then
            arr_OutputLabel(int_i) = PDE_TimeStep.major
            Exit For
        Else
            arr_OutputLabel(int_i) = PDE_TimeStep.important
        End If
    Next int_j
Next int_i

''''''''''''''''''''''''''''''''''''''
''''Step 2 - Add in Iteration Step''''
''''''''''''''''''''''''''''''''''''''
dbl_Mat = calc_yearfrac(lng_ValDate, lng_MatDate, "ACT/365")
'dbl_EffStep = arr_CallDate(UBound(arr_CallDate)) * int_TimeStep
dbl_EffStep = dbl_Mat * int_TimeStep
int_N_iteration = (Int(dbl_EffStep / UBound(arr_ImpDate)) - 1)

ReDim Preserve arr_OutputDate(1 To UBound(arr_ImpDate) * (int_N_iteration + 1))
ReDim Preserve arr_OutputLabel(1 To UBound(arr_ImpDate) * (int_N_iteration + 1))

int_j = 0
For int_i = 1 To UBound(arr_ImpDate)
    If int_i = 1 Then
        dbl_T1 = 0
        dbl_T2 = arr_ImpDate(int_i)
    Else
        dbl_T1 = arr_ImpDate(int_i - 1)
        dbl_T2 = arr_ImpDate(int_i)
    End If
    
        dbl_IteStep = (dbl_T2 - dbl_T1) / (int_N_iteration + 1)
        If dbl_IteStep < 0.00001 Then GoTo SkipIteration
        int_k = 0

        Do
            If int_N_iteration = 0 Then: Exit Do
            
            int_j = int_j + 1
            int_k = int_k + 1
            dbl_IteTemp = dbl_T1 + dbl_IteStep * int_k
            arr_OutputDate(UBound(arr_ImpDate) + int_j) = dbl_IteTemp
            arr_OutputLabel(UBound(arr_ImpDate) + int_j) = PDE_TimeStep.CN
            
            If (dbl_T2 - (dbl_IteTemp + dbl_IteStep)) < 0.00001 Then: Exit Do
        Loop
SkipIteration:
Next int_i

'Sorting the TimeStep - BubbleSort
Call BubbleSort(arr_OutputDate, arr_OutputLabel)

''''''''''''''''''''''''''''''''''''''
''''''Step 3 - Add in Call Date'''''''
''''''''''''''''''''''''''''''''''''''
For int_i = 1 To UBound(arr_CallDate)

    For int_j = 1 To UBound(arr_OutputDate)
        If arr_CallDate(int_i) = arr_OutputDate(int_j) Then
            arr_OutputLabel(int_j) = PDE_TimeStep.major
            GoTo Skip
        End If
    Next int_j

ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 1)
ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 1)
arr_OutputDate(UBound(arr_OutputDate)) = arr_CallDate(int_i)
arr_OutputLabel(UBound(arr_OutputDate)) = PDE_TimeStep.major

Skip:

Next int_i

Call BubbleSort(arr_OutputDate, arr_OutputLabel)
''''''''''''''''''''''''''''''''''''''
''''Step 4 - Add in Start Date    ''''
''''''''''''''''''''''''''''''''''''''
ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 1)
ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 1)

arr_OutputDate(UBound(arr_OutputDate)) = 0
arr_OutputLabel(UBound(arr_OutputLabel)) = PDE_TimeStep.CN

'Sorting the TimeStep - BubbleSort
Call BubbleSort(arr_OutputDate, arr_OutputLabel)

''''''''''''''''''''''''''''''''''''''
''''Step 5 - Add in Rannacher Step''''
''''''''''''''''''''''''''''''''''''''

'dbl_Rann = 1 / int_TimeStep

For int_i = 2 To UBound(arr_OutputDate)
    If arr_OutputLabel(int_i) = PDE_TimeStep.major Then
        
        dbl_IteTemp = arr_OutputDate(int_i) - arr_OutputDate(int_i - 1)
        dbl_IteTemp = dbl_IteTemp / 3
        
        ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 2)
        ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 2)
        
        arr_OutputDate(UBound(arr_OutputDate)) = arr_OutputDate(int_i) - dbl_IteTemp
        arr_OutputDate(UBound(arr_OutputDate) - 1) = arr_OutputDate(int_i) - dbl_IteTemp * 2
        
        arr_OutputLabel(UBound(arr_OutputDate)) = PDE_TimeStep.Implicit
        arr_OutputLabel(UBound(arr_OutputDate) - 1) = PDE_TimeStep.Implicit
    
    End If

Next int_i

'Sorting the TimeStep - BubbleSort
Call BubbleSort(arr_OutputDate, arr_OutputLabel)

arr_TimeStep = arr_OutputDate
arr_TimeLabel = arr_OutputLabel

End Sub
Private Sub CreateTimeStep_CRA_2()

Dim int_i As Integer
Dim int_j As Integer
Dim int_k As Integer

Dim int_N_iteration As Integer
Dim dbl_IteStep As Double
Dim dbl_IteTemp As Double
Dim dbl_T1 As Double
Dim dbl_T2 As Double
Dim dbl_Mat As Double
Dim dbl_EffStep As Double
Dim dbl_Rann As Double

Dim arr_OutputDate() As Double
Dim arr_OutputLabel() As Integer

''''''''''''''''''''''''''''''''''''''
''''Step 1 - Defined Important Date''''
''''''''''''''''''''''''''''''''''''''
arr_OutputDate = arr_ImpDate
ReDim arr_OutputLabel(1 To UBound(arr_OutputDate))

For int_i = 1 To UBound(arr_ImpDate)
    For int_j = 1 To UBound(arr_CallDate)
        If arr_CallDate(int_j) = arr_ImpDate(int_i) Then
            arr_OutputLabel(int_i) = PDE_TimeStep.major
            Exit For
        Else
            arr_OutputLabel(int_i) = PDE_TimeStep.important
        End If
    Next int_j
Next int_i

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''Step 2 - Add in Iteration Step for Non-First Period''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'dbl_Mat = calc_yearfrac(lng_ValDate, lng_MatDate, "ACT/365")
'dbl_EffStep = arr_CallDate(UBound(arr_CallDate)) * int_TimeStep
'dbl_EffStep = dbl_Mat * int_TimeStep
dbl_IteStep = Int(365 / int_TimeStep) / 365


For int_i = 1 To UBound(arr_ImpDate)

    int_j = 0
    
    If int_i = 1 Then
        dbl_T1 = 0
        dbl_T2 = arr_ImpDate(int_i)
    Else
        dbl_T1 = arr_ImpDate(int_i - 1)
        dbl_T2 = arr_ImpDate(int_i)
    End If

    Do
        ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 1)
        ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 1)
        
        int_j = int_j + 1
        dbl_IteTemp = dbl_T1 + dbl_IteStep * int_j
        arr_OutputDate(UBound(arr_OutputDate)) = dbl_IteTemp
        arr_OutputLabel(UBound(arr_OutputLabel)) = PDE_TimeStep.CN
        
        If (dbl_IteTemp + dbl_IteStep) > dbl_T2 Then: Exit Do

    Loop
SkipIteration:
Next int_i

'Sorting the TimeStep - BubbleSort
Call BubbleSort(arr_OutputDate, arr_OutputLabel)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''Step 3 - Add in Iteration Step (for last Period)''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

dbl_IteStep = 1 / 365
Dim dbl_Ubound As Integer
dbl_Ubound = UBound(arr_OutputDate)

For int_i = 2 To dbl_Ubound

    int_j = 0
    dbl_T1 = arr_OutputDate(int_i - 1)
    dbl_T2 = arr_OutputDate(int_i)
    
    If arr_OutputDate(int_i) <= 30 / 365 And (dbl_T2 - dbl_T1) >= dbl_IteStep Then
    
        Do
        ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 1)
        ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 1)
        
        int_j = int_j + 1
        dbl_IteTemp = dbl_T1 + dbl_IteStep * int_j
        arr_OutputDate(UBound(arr_OutputDate)) = dbl_IteTemp
        arr_OutputLabel(UBound(arr_OutputLabel)) = PDE_TimeStep.CN
        
        If (dbl_IteTemp + dbl_IteStep) > dbl_T2 Then: Exit Do
        Loop
    ElseIf arr_OutputDate(int_i) = 30 / 365 Then
    
        Exit For
    ElseIf arr_OutputDate(int_i) > 30 / 365 And arr_OutputDate(int_i + 1) < 30 / 365 Then
                
        ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 1)
        ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 1)
            
            arr_OutputDate(UBound(arr_OutputDate)) = 30 / 365
            arr_OutputLabel(UBound(arr_OutputLabel)) = PDE_TimeStep.CN
            dbl_T2 = 30 / 365
            
        Do
        ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 1)
        ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 1)
        
            int_j = int_j + 1
            dbl_IteTemp = dbl_T1 + dbl_IteStep * int_j
            arr_OutputDate(UBound(arr_OutputDate)) = dbl_IteTemp
            arr_OutputLabel(UBound(arr_OutputLabel)) = PDE_TimeStep.CN
            
        If (dbl_IteTemp + dbl_IteStep) > dbl_T2 Then: Exit Do
        Loop
        
        Exit For
    Else
    
    End If
    
Next int_i
'Sorting the TimeStep - BubbleSort
Call BubbleSort(arr_OutputDate, arr_OutputLabel)

''''''''''''''''''''''''''''''''''''''
''''''Step 3 - Add in Call Date'''''''
''''''''''''''''''''''''''''''''''''''
For int_i = 1 To UBound(arr_CallDate)

    For int_j = 1 To UBound(arr_OutputDate)
        If arr_CallDate(int_i) = arr_OutputDate(int_j) Then
            arr_OutputLabel(int_j) = PDE_TimeStep.major
            GoTo Skip
        End If
    Next int_j

ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 1)
ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 1)
arr_OutputDate(UBound(arr_OutputDate)) = arr_CallDate(int_i)
arr_OutputLabel(UBound(arr_OutputDate)) = PDE_TimeStep.major

Skip:

Next int_i

Call BubbleSort(arr_OutputDate, arr_OutputLabel)

''''''''''''''''''''''''''''''''''''''
''''Step 4 - Add in Start Date    ''''
''''''''''''''''''''''''''''''''''''''
ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 2)
ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 2)

arr_OutputDate(UBound(arr_OutputDate)) = 0
arr_OutputDate(UBound(arr_OutputDate)) = 0.00001

arr_OutputLabel(UBound(arr_OutputLabel)) = PDE_TimeStep.CN
arr_OutputLabel(UBound(arr_OutputLabel)) = PDE_TimeStep.CN

'Sorting the TimeStep - BubbleSort
Call BubbleSort(arr_OutputDate, arr_OutputLabel)

''''''''''''''''''''''''''''''''''''''
''''Step 5 - Add in Rannacher Step''''
''''''''''''''''''''''''''''''''''''''

For int_i = 1 To UBound(arr_OutputDate)
    If arr_OutputLabel(int_i) = PDE_TimeStep.major And arr_OutputDate(int_i) <> arr_ImpDate(1) Then
        
        dbl_IteTemp = arr_OutputDate(int_i) - arr_OutputDate(int_i - 1)
        dbl_IteTemp = dbl_IteTemp / 3
        
        ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 2)
        ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 2)
        
        arr_OutputDate(UBound(arr_OutputDate)) = arr_OutputDate(int_i) - dbl_IteTemp
        arr_OutputDate(UBound(arr_OutputDate) - 1) = arr_OutputDate(int_i) - dbl_IteTemp * 2
        
        arr_OutputLabel(UBound(arr_OutputDate)) = PDE_TimeStep.Implicit
        arr_OutputLabel(UBound(arr_OutputDate) - 1) = PDE_TimeStep.Implicit
    
    End If

Next int_i

'Sorting the TimeStep - BubbleSort
Call BubbleSort(arr_OutputDate, arr_OutputLabel)

arr_TimeStep = arr_OutputDate
arr_TimeLabel = arr_OutputLabel

End Sub
Private Sub CreateTimeStep_CRA_1()

Dim int_i As Integer
Dim int_j As Integer
Dim int_k As Integer

Dim int_N_iteration As Integer
Dim dbl_IteStep As Double
Dim dbl_IteTemp As Double
Dim dbl_T1 As Double
Dim dbl_T2 As Double
Dim dbl_Mat As Double
Dim dbl_EffStep As Double
Dim dbl_Rann As Double

Dim arr_OutputDate() As Double
Dim arr_OutputLabel() As Integer

''''''''''''''''''''''''''''''''''''''
''''Step 1 - Defined Important Date''''
''''''''''''''''''''''''''''''''''''''
arr_OutputDate = arr_ImpDate
ReDim arr_OutputLabel(1 To UBound(arr_OutputDate))

For int_i = 1 To UBound(arr_ImpDate)
    For int_j = 1 To UBound(arr_CallDate)
        If arr_CallDate(int_j) = arr_ImpDate(int_i) Then
            arr_OutputLabel(int_i) = PDE_TimeStep.major
            Exit For
        Else
            arr_OutputLabel(int_i) = PDE_TimeStep.important
        End If
    Next int_j
Next int_i

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''Step 2 - Add in Iteration Step for Non-First 30 Period''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

dbl_IteStep = Int(365 / int_TimeStep) / 365

int_j = 0
dbl_T1 = 0
dbl_T2 = arr_ImpDate(UBound(arr_ImpDate))

Do
    ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 1)
    ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 1)
    
    For int_i = 1 To UBound(arr_ImpDate)
        If Abs((dbl_IteTemp + dbl_IteStep) - arr_ImpDate(int_i)) < 0.00001 Then
            int_j = int_j + 1
        End If
    Next int_i
    
    int_j = int_j + 1
    dbl_IteTemp = dbl_T1 + dbl_IteStep * int_j
    arr_OutputDate(UBound(arr_OutputDate)) = dbl_IteTemp
    arr_OutputLabel(UBound(arr_OutputLabel)) = PDE_TimeStep.CN
    
    If (dbl_IteTemp + dbl_IteStep) >= dbl_T2 Or Abs((dbl_IteTemp + dbl_IteStep - dbl_T2)) < 0.00001 Then: Exit Do

Loop

'Sorting the TimeStep - BubbleSort
Call BubbleSort(arr_OutputDate, arr_OutputLabel)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''Step 3 - Add in Iteration Step (for last 30-days)''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Add in 30/365
For int_i = 1 To UBound(arr_OutputDate)
    If arr_OutputDate(int_i) = 30 / 365 Then
        Exit For
    End If
    
    If int_i = UBound(arr_OutputDate) Then
        ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 1)
        ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 1)
            
        arr_OutputDate(UBound(arr_OutputDate)) = 30 / 365
        arr_OutputLabel(UBound(arr_OutputLabel)) = PDE_TimeStep.CN
    End If

Next int_i

'Sorting the TimeStep - BubbleSort
Call BubbleSort(arr_OutputDate, arr_OutputLabel)

'Loop for daily accrual for the last 30/365
dbl_IteStep = 1 / 365
Dim dbl_Ubound As Integer
dbl_Ubound = UBound(arr_OutputDate) - 1

For int_i = 1 To dbl_Ubound
    
    int_j = 0
    If int_i = 1 Then
        dbl_T1 = 0
        dbl_T2 = arr_OutputDate(int_i)
    Else
        dbl_T1 = arr_OutputDate(int_i - 1)
        dbl_T2 = arr_OutputDate(int_i)
    End If
    
    
    If arr_OutputDate(int_i) <= 30 / 365 And Abs(dbl_T2 - dbl_T1 - dbl_IteStep) > 0.00001 Then
    
        Do
        ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 1)
        ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 1)
        
        int_j = int_j + 1
        dbl_IteTemp = dbl_T1 + dbl_IteStep * int_j
        arr_OutputDate(UBound(arr_OutputDate)) = dbl_IteTemp
        arr_OutputLabel(UBound(arr_OutputLabel)) = PDE_TimeStep.CN
        
        If (dbl_IteTemp + dbl_IteStep) >= dbl_T2 Or Abs((dbl_IteTemp + dbl_IteStep - dbl_T2)) < 0.00001 Then: Exit Do
        Loop
    
    End If
    
Next int_i
'Sorting the TimeStep - BubbleSort
Call BubbleSort(arr_OutputDate, arr_OutputLabel)

''''''''''''''''''''''''''''''''''''''
''''''Step 3 - Add in Call Date'''''''
''''''''''''''''''''''''''''''''''''''
For int_i = 1 To UBound(arr_CallDate)

    For int_j = 1 To UBound(arr_OutputDate)
        If Abs(arr_CallDate(int_i) - arr_OutputDate(int_j)) < 0.00001 Then
            arr_OutputLabel(int_j) = PDE_TimeStep.major
            GoTo Skip
        End If
    Next int_j

ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 1)
ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 1)
arr_OutputDate(UBound(arr_OutputDate)) = arr_CallDate(int_i)
arr_OutputLabel(UBound(arr_OutputDate)) = PDE_TimeStep.major

Skip:

Next int_i

Call BubbleSort(arr_OutputDate, arr_OutputLabel)

''''''''''''''''''''''''''''''''''''''
''''Step 4 - Add in Start Date    ''''
''''''''''''''''''''''''''''''''''''''
ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 2)
ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 2)

arr_OutputDate(UBound(arr_OutputDate) - 1) = 0
arr_OutputDate(UBound(arr_OutputDate)) = 0.00001

arr_OutputLabel(UBound(arr_OutputLabel) - 1) = PDE_TimeStep.CN
arr_OutputLabel(UBound(arr_OutputLabel)) = PDE_TimeStep.CN

'Sorting the TimeStep - BubbleSort
Call BubbleSort(arr_OutputDate, arr_OutputLabel)

''''''''''''''''''''''''''''''''''''''
''''Step 5 - Add in Rannacher Step''''
''''''''''''''''''''''''''''''''''''''

For int_i = 1 To UBound(arr_OutputDate)

    If arr_OutputLabel(int_i) = PDE_TimeStep.major Then

        
        dbl_IteTemp = arr_OutputDate(int_i) - arr_OutputDate(int_i - 1)
        
        If Round(dbl_IteTemp, 6) / 3 <= Round(1 / 365, 6) Then
        
            dbl_IteTemp = (arr_OutputDate(int_i - 1) - arr_OutputDate(int_i - 2)) / 3
        
            ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 2)
            ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 2)
            
            arr_OutputDate(UBound(arr_OutputDate)) = arr_OutputDate(int_i - 1) - dbl_IteTemp
            arr_OutputDate(UBound(arr_OutputDate) - 1) = arr_OutputDate(int_i - 1) - dbl_IteTemp * 2
            
            arr_OutputLabel(int_i - 1) = PDE_TimeStep.Implicit
            arr_OutputLabel(UBound(arr_OutputDate)) = PDE_TimeStep.Implicit
            arr_OutputLabel(UBound(arr_OutputDate) - 1) = PDE_TimeStep.CN
        
        Else
            dbl_IteTemp = dbl_IteTemp / 3
        
            ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 2)
            ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 2)
            
            arr_OutputDate(UBound(arr_OutputDate)) = arr_OutputDate(int_i) - dbl_IteTemp
            arr_OutputDate(UBound(arr_OutputDate) - 1) = arr_OutputDate(int_i) - dbl_IteTemp * 2
            
            arr_OutputLabel(UBound(arr_OutputDate)) = PDE_TimeStep.Implicit
            arr_OutputLabel(UBound(arr_OutputDate) - 1) = PDE_TimeStep.Implicit
        End If
        

    
    End If

Next int_i

'Sorting the TimeStep - BubbleSort
Call BubbleSort(arr_OutputDate, arr_OutputLabel)

arr_TimeStep = arr_OutputDate
arr_TimeLabel = arr_OutputLabel

End Sub
Private Sub CreateTimeStep_CRA()

Dim int_i As Integer
Dim int_j As Integer
Dim int_k As Integer

Dim int_N_iteration As Integer
Dim dbl_IteStep As Double
Dim dbl_IteTemp As Double
Dim dbl_T1 As Double
Dim dbl_T2 As Double
Dim dbl_Mat As Double
Dim dbl_EffStep As Double
Dim dbl_Rann As Double

Dim arr_OutputDate() As Double
Dim arr_OutputLabel() As Integer

''''''''''''''''''''''''''''''''''''''
''''Step 1 - Defined Important Date''''
''''''''''''''''''''''''''''''''''''''
arr_OutputDate = arr_ImpDate
ReDim arr_OutputLabel(1 To UBound(arr_OutputDate))

For int_i = 1 To UBound(arr_ImpDate)
    For int_j = 1 To UBound(arr_CallDate)
        If arr_CallDate(int_j) = arr_ImpDate(int_i) Then
            arr_OutputLabel(int_i) = PDE_TimeStep.major
            Exit For
        Else
            arr_OutputLabel(int_i) = PDE_TimeStep.important
        End If
    Next int_j
Next int_i

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''Step 2 - Add in Iteration Step for Non-First 30 Period''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

dbl_IteStep = Int(365 / int_TimeStep) / 365
'dbl_IteStep = Round(365 / int_TimeStep, 0) / 365

int_j = 0
dbl_T1 = 0
dbl_T2 = arr_ImpDate(UBound(arr_ImpDate))

Do
    ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 1)
    ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 1)
    
    For int_i = 1 To UBound(arr_ImpDate)
        If Abs((dbl_IteTemp + dbl_IteStep) - arr_ImpDate(int_i)) < 0.00001 Then
            int_j = int_j + 1
        End If
    Next int_i
    
    int_j = int_j + 1
    dbl_IteTemp = dbl_T1 + dbl_IteStep * int_j
    arr_OutputDate(UBound(arr_OutputDate)) = dbl_IteTemp
    arr_OutputLabel(UBound(arr_OutputLabel)) = PDE_TimeStep.CN
    
    If (dbl_IteTemp + dbl_IteStep) >= dbl_T2 Or Abs((dbl_IteTemp + dbl_IteStep - dbl_T2)) < 0.00001 Then: Exit Do

Loop

'Sorting the TimeStep - BubbleSort
Call BubbleSort(arr_OutputDate, arr_OutputLabel)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''Step 3 - Add in Iteration Step (for last 30-days)''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Add in 30/365
For int_i = 1 To UBound(arr_OutputDate)
    If arr_OutputDate(int_i) = 30 / 365 Then
        Exit For
    End If
    
    If int_i = UBound(arr_OutputDate) Then
        ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 1)
        ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 1)
            
        arr_OutputDate(UBound(arr_OutputDate)) = 30 / 365
        arr_OutputLabel(UBound(arr_OutputLabel)) = PDE_TimeStep.CN
    End If

Next int_i

'Sorting the TimeStep - BubbleSort
Call BubbleSort(arr_OutputDate, arr_OutputLabel)

'Loop for daily accrual for the last 30/365
dbl_IteStep = 1 / 365
Dim dbl_Ubound As Integer
dbl_Ubound = UBound(arr_OutputDate) - 1

For int_i = 1 To dbl_Ubound
    
    int_j = 0
    If int_i = 1 Then
        dbl_T1 = 0
        dbl_T2 = arr_OutputDate(int_i)
    Else
        dbl_T1 = arr_OutputDate(int_i - 1)
        dbl_T2 = arr_OutputDate(int_i)
    End If
    
    
    If arr_OutputDate(int_i) <= 30 / 365 And Abs(dbl_T2 - dbl_T1 - dbl_IteStep) > 0.00001 Then
    
        Do
        ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 1)
        ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 1)
        
        int_j = int_j + 1
        dbl_IteTemp = dbl_T1 + dbl_IteStep * int_j
        arr_OutputDate(UBound(arr_OutputDate)) = dbl_IteTemp
        arr_OutputLabel(UBound(arr_OutputLabel)) = PDE_TimeStep.CN
        
        If (dbl_IteTemp + dbl_IteStep) >= dbl_T2 Or Abs((dbl_IteTemp + dbl_IteStep - dbl_T2)) < 0.00001 Then: Exit Do
        Loop
    
    End If
    
Next int_i
'Sorting the TimeStep - BubbleSort
Call BubbleSort(arr_OutputDate, arr_OutputLabel)

''''''''''''''''''''''''''''''''''''''
''''''Step 3 - Add in Call Date'''''''
''''''''''''''''''''''''''''''''''''''
For int_i = 1 To UBound(arr_CallDate)

    For int_j = 1 To UBound(arr_OutputDate)
        If Abs(arr_CallDate(int_i) - arr_OutputDate(int_j)) < 0.00001 Then
            arr_OutputLabel(int_j) = PDE_TimeStep.major
            GoTo Skip
        End If
    Next int_j

ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 1)
ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 1)
arr_OutputDate(UBound(arr_OutputDate)) = arr_CallDate(int_i)
arr_OutputLabel(UBound(arr_OutputDate)) = PDE_TimeStep.major

Skip:

Next int_i

Call BubbleSort(arr_OutputDate, arr_OutputLabel)

''''''''''''''''''''''''''''''''''''''
''''Step 4 - Add in Start Date    ''''
''''''''''''''''''''''''''''''''''''''
ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 2)
ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 2)

arr_OutputDate(UBound(arr_OutputDate) - 1) = 0
arr_OutputDate(UBound(arr_OutputDate)) = 0.00001

arr_OutputLabel(UBound(arr_OutputLabel) - 1) = PDE_TimeStep.CN
arr_OutputLabel(UBound(arr_OutputLabel)) = PDE_TimeStep.CN

'Sorting the TimeStep - BubbleSort
Call BubbleSort(arr_OutputDate, arr_OutputLabel)

''''''''''''''''''''''''''''''''''''''
''''Step 5 - Add in Rannacher Step''''
''''''''''''''''''''''''''''''''''''''

For int_i = 1 To UBound(arr_OutputDate)

    If arr_OutputLabel(int_i) = PDE_TimeStep.major Then

        
        dbl_IteTemp = arr_OutputDate(int_i) - arr_OutputDate(int_i - 1)
        
        If Round(dbl_IteTemp, 6) / 3 <= Round(1 / 365, 6) Then
        
            dbl_IteTemp = (arr_OutputDate(int_i - 1) - arr_OutputDate(int_i - 2)) / 3
        
            ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 2)
            ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 2)
            
            arr_OutputDate(UBound(arr_OutputDate)) = arr_OutputDate(int_i - 1) - dbl_IteTemp
            arr_OutputDate(UBound(arr_OutputDate) - 1) = arr_OutputDate(int_i - 1) - dbl_IteTemp * 2
            
            arr_OutputLabel(int_i - 1) = PDE_TimeStep.Implicit
            arr_OutputLabel(UBound(arr_OutputDate)) = PDE_TimeStep.Implicit
            arr_OutputLabel(UBound(arr_OutputDate) - 1) = PDE_TimeStep.CN
        
        Else
            dbl_IteTemp = dbl_IteTemp / 3
        
            ReDim Preserve arr_OutputDate(1 To UBound(arr_OutputDate) + 2)
            ReDim Preserve arr_OutputLabel(1 To UBound(arr_OutputLabel) + 2)
            
            arr_OutputDate(UBound(arr_OutputDate)) = arr_OutputDate(int_i) - dbl_IteTemp
            arr_OutputDate(UBound(arr_OutputDate) - 1) = arr_OutputDate(int_i) - dbl_IteTemp * 2
            
            arr_OutputLabel(UBound(arr_OutputDate)) = PDE_TimeStep.Implicit
            arr_OutputLabel(UBound(arr_OutputDate) - 1) = PDE_TimeStep.Implicit
        End If
        

    
    End If

Next int_i

'Sorting the TimeStep - BubbleSort
Call BubbleSort(arr_OutputDate, arr_OutputLabel)

'''''''''''''''''''''''''''''''''''''''''''''''''
''''Step 6 - Cater for Business Days and Days''''
'''''''''''''''''''''''''''''''''''''''''''''''''
If bln_BusinessDay = True Then

    Dim cal_Leg As Calendar
    cal_Leg = irl_LegA.Calendar
    
    ReDim arr_TimeStep(1 To 2) As Double
    ReDim arr_TimeLabel(1 To 2) As Integer
    
    For int_i = 1 To 2
        arr_TimeStep(int_i) = arr_OutputDate(int_i)
        arr_TimeLabel(int_i) = arr_OutputLabel(int_i)
    Next int_i
    
    For int_i = 3 To UBound(arr_OutputDate)
        If arr_OutputDate(int_i) * 365 + lng_ValDate = date_workday(lng_ValDate + 365 * arr_OutputDate(int_i) - 1, 1, cal_Leg.HolDates, cal_Leg.Weekends) _
           Or arr_OutputLabel(int_i) = PDE_TimeStep.major Or arr_OutputLabel(int_i) = PDE_TimeStep.important Then
           'Or arr_OutputDate(int_i) <= 30 / 365
        
                ReDim Preserve arr_TimeStep(1 To UBound(arr_TimeStep) + 1)
                ReDim Preserve arr_TimeLabel(1 To UBound(arr_TimeLabel) + 1)
                
                arr_TimeStep(UBound(arr_TimeStep)) = arr_OutputDate(int_i)
                arr_TimeLabel(UBound(arr_TimeStep)) = arr_OutputLabel(int_i)
        End If
    Next int_i

Else
    arr_TimeStep() = arr_OutputDate()
    arr_TimeLabel() = arr_OutputLabel()
End If

End Sub

Private Sub CreateTimeStep_ImpDate()
    
Dim int_i As Integer
Dim int_cnt As Integer: int_cnt = 1
Dim arr_output() As Double

For int_i = 1 To UBound(arr_TimeStep)
    
    If arr_TimeLabel(int_i) = PDE_TimeStep.important Or arr_TimeLabel(int_i) = PDE_TimeStep.major Or arr_TimeStep(int_i) = 0 Then
        ReDim Preserve arr_output(1 To int_cnt)
        arr_output(int_cnt) = arr_TimeStep(int_i)
        int_cnt = int_cnt + 1
    End If
Next int_i

arr_TimeStep_ImpDate = arr_output

End Sub
Public Function PDE_Matrix(arr_PDE As Variant, dbl_T1 As Double, dbl_T2 As Double, pde_TS As Integer) As Variant

Dim int_i As Double
Dim int_j As Double

Dim dbl_Pu As Double
Dim dbl_Pm As Double
Dim dbl_Pd As Double

Dim dbl_Pu_LBound As Double
Dim dbl_Pm_LBound As Double
Dim dbl_Pd_LBound As Double

Dim dbl_Pu_UBound As Double
Dim dbl_Pm_UBound As Double
Dim dbl_Pd_UBound As Double

Dim dbl_Theta As Double
Dim dbl_DeltaT As Double: dbl_DeltaT = dbl_T2 - dbl_T1
Dim dbl_Sigma As Double: dbl_Sigma = LookUpSigma(dbl_T2)

Dim dbl_spotstep As Double

Dim arr_PDE_A() As Double
Dim arr_PDE_B() As Double

ReDim arr_PDE_A(1 To UBound(arr_SpotStep), 1 To UBound(arr_SpotStep)) As Double
ReDim arr_PDE_B(1 To UBound(arr_SpotStep), 1 To UBound(arr_SpotStep)) As Double

If pde_TS = PDE_TimeStep.Implicit Or pde_TS = PDE_TimeStep.major Then
    dbl_Theta = 2
Else
    dbl_Theta = 1
End If

For int_i = 1 To UBound(arr_SpotStep)

    'Defining Pu/Pm/Pd
    dbl_spotstep = arr_SpotStep(2) - arr_SpotStep(1)
    dbl_Pu = 1 / 4 * dbl_DeltaT * (-dbl_Sigma * dbl_Sigma / dbl_spotstep / dbl_spotstep + dbl_MeanRev * arr_SpotStep(int_i) / dbl_spotstep)
    dbl_Pm = dbl_DeltaT / 2 * (dbl_Sigma * dbl_Sigma / dbl_spotstep / dbl_spotstep + arr_SpotStep(int_i))
    dbl_Pd = 1 / 4 * dbl_DeltaT * (-dbl_Sigma * dbl_Sigma / dbl_spotstep / dbl_spotstep - dbl_MeanRev * arr_SpotStep(int_i) / dbl_spotstep)
    
    'Forming PDE_A & PDE_B
    For int_j = 1 To UBound(arr_SpotStep)
        
        If int_j = int_i Then
            arr_PDE_A(int_i, int_j) = 1 + dbl_Theta * dbl_Pm
        ElseIf int_j - int_i = 1 Then arr_PDE_A(int_i, int_j) = dbl_Theta * dbl_Pu
        ElseIf int_j - int_i = -1 Then: arr_PDE_A(int_i, int_j) = dbl_Theta * dbl_Pd
        End If
    
        If int_j = int_i Then
            arr_PDE_B(int_i, int_j) = 1 - dbl_Theta * dbl_Pm
        ElseIf int_j - int_i = 1 Then: arr_PDE_B(int_i, int_j) = -dbl_Theta * dbl_Pu
        ElseIf int_j - int_i = -1 Then: arr_PDE_B(int_i, int_j) = -dbl_Theta * dbl_Pd
        End If

    Next int_j
    
        'Adjusting for Boundary Condition (LBOUND)
        If int_i = 1 Then
        
            dbl_Pu_LBound = 1 / 2 * dbl_DeltaT * (dbl_MeanRev * arr_SpotStep(int_i) / dbl_spotstep)
            dbl_Pm_LBound = dbl_DeltaT / 2 * (arr_SpotStep(int_i) - dbl_MeanRev * arr_SpotStep(int_i) / dbl_spotstep)
            dbl_Pd_LBound = 0
            
            arr_PDE_A(int_i, int_i) = 1 + dbl_Theta * dbl_Pm_LBound
            arr_PDE_A(int_i, int_i + 1) = dbl_Theta * dbl_Pu_LBound
            
            arr_PDE_B(int_i, int_i) = 1 - dbl_Theta * dbl_Pm_LBound
            arr_PDE_B(int_i, int_i + 1) = -dbl_Theta * dbl_Pu_LBound
        
        End If
        
        'Adjusting for Boundary Condition (UBOUND)
        If int_i = UBound(arr_SpotStep) Then
        
            dbl_Pu_UBound = 0
            dbl_Pm_UBound = dbl_DeltaT / 2 * (arr_SpotStep(int_i) + dbl_MeanRev * arr_SpotStep(int_i) / dbl_spotstep)
            dbl_Pd_UBound = 1 / 2 * dbl_DeltaT * (-dbl_MeanRev * arr_SpotStep(int_i) / dbl_spotstep)
            
            arr_PDE_A(int_i, int_i) = 1 + dbl_Theta * dbl_Pm_UBound
            arr_PDE_A(int_i, int_i - 1) = dbl_Theta * dbl_Pd_UBound
            
            arr_PDE_B(int_i, int_i) = 1 - dbl_Theta * dbl_Pm_UBound
            arr_PDE_B(int_i, int_i - 1) = -dbl_Theta * dbl_Pd_UBound
        End If
    
Next int_i

If pde_TS = PDE_TimeStep.CN Or pde_TS = PDE_TimeStep.important Then
    arr_PDE = PDE_MMULT(arr_PDE_B, arr_PDE)
End If

'Solving Using LU Decomposition
'Decompose Matrix A
Dim arr_PDE_A_LU() As Double

Dim arr_PDE_X() As Variant
Dim arr_PDE_Y() As Variant

ReDim arr_PDE_X(1 To UBound(arr_SpotStep)) As Variant
ReDim arr_PDE_Y(1 To UBound(arr_SpotStep)) As Variant

arr_PDE_A_LU = LU_Decompose_Simplified(arr_PDE_A)

'''Solving Y'''
For int_i = 1 To UBound(arr_SpotStep)
    If int_i = 1 Then
        arr_PDE_Y(int_i) = arr_PDE(int_i) / arr_PDE_A_LU(1, int_i, int_i)
    Else
        arr_PDE_Y(int_i) = (arr_PDE(int_i) - arr_PDE_Y(int_i - 1) * arr_PDE_A_LU(1, int_i, int_i - 1)) / arr_PDE_A_LU(1, int_i, int_i)
    End If

Next int_i

'''Solving X'''
For int_i = UBound(arr_SpotStep) To 1 Step -1
    If int_i = UBound(arr_SpotStep) Then
        arr_PDE_X(int_i) = arr_PDE_Y(int_i)
    Else
        arr_PDE_X(int_i) = arr_PDE_Y(int_i) - arr_PDE_X(int_i + 1) * arr_PDE_A_LU(2, int_i, int_i + 1)
    End If
Next int_i

PDE_Matrix = arr_PDE_X

End Function
Private Function PDE_MMULT(arr_PDE_B() As Double, arr_x As Variant) As Variant()

Dim int_i As Double
Dim arr_output() As Variant

ReDim arr_output(1 To UBound(arr_x))

For int_i = 1 To UBound(arr_x)
    If int_i = 1 Then
        arr_output(int_i) = arr_PDE_B(int_i, int_i) * arr_x(int_i) + arr_PDE_B(int_i, int_i + 1) * arr_x(int_i + 1)
    ElseIf int_i = UBound(arr_x) Then
        arr_output(int_i) = arr_PDE_B(int_i, int_i) * arr_x(int_i) + arr_PDE_B(int_i, int_i - 1) * arr_x(int_i - 1)
    Else
        arr_output(int_i) = arr_PDE_B(int_i, int_i + 1) * arr_x(int_i + 1) + arr_PDE_B(int_i, int_i) * arr_x(int_i) + arr_PDE_B(int_i, int_i - 1) * arr_x(int_i - 1)
    End If

Next int_i

PDE_MMULT = arr_output

End Function
Public Function LookUpSigma(dbl_T1 As Double) As Double

Dim int_cnt As Integer
Dim int_max As Integer

For int_cnt = 1 To UBound(arr_sigma)
    If arr_sigma(int_cnt, 2) >= dbl_T1 Then
        int_max = int_cnt
        Exit For
    End If
Next int_cnt

LookUpSigma = arr_sigma(int_max, 1)

End Function

Private Sub BubbleSort(arr_OutputDate() As Double, arr_OutputLabel() As Integer)
' To Sort the TimeStep
Dim int_i As Integer
Dim int_j As Integer
Dim dbl_SortTempDate As Double
Dim dbl_SortTempLabel As Double

For int_i = LBound(arr_OutputDate, 1) To UBound(arr_OutputDate, 1)
     dbl_SortTempDate = arr_OutputDate(int_i)
     dbl_SortTempLabel = arr_OutputLabel(int_i)
     
     For int_j = LBound(arr_OutputDate, 1) To UBound(arr_OutputDate, 1)
         If arr_OutputDate(int_j) > dbl_SortTempDate Then
         
             arr_OutputDate(int_i) = arr_OutputDate(int_j)
             arr_OutputLabel(int_i) = arr_OutputLabel(int_j)
             
             arr_OutputDate(int_j) = dbl_SortTempDate
             arr_OutputLabel(int_j) = dbl_SortTempLabel
             
             dbl_SortTempDate = arr_OutputDate(int_i)
             dbl_SortTempLabel = arr_OutputLabel(int_i)
         End If
     Next int_j
 Next int_i

End Sub

Private Sub BubbleSort_ImpDate(arr_OutputDate() As Double)
' To Sort the ImpDate
Dim int_i As Integer
Dim int_j As Integer
Dim dbl_SortTempDate As Double

For int_i = LBound(arr_OutputDate, 1) To UBound(arr_OutputDate, 1)
     dbl_SortTempDate = arr_OutputDate(int_i)
     
     For int_j = LBound(arr_OutputDate, 1) To UBound(arr_OutputDate, 1)
         If arr_OutputDate(int_j) > dbl_SortTempDate Then
         
             arr_OutputDate(int_i) = arr_OutputDate(int_j)
             arr_OutputDate(int_j) = dbl_SortTempDate
            
             dbl_SortTempDate = arr_OutputDate(int_i)
             
         End If
     Next int_j
 Next int_i

End Sub
'ALVIN  20181210 - Payoff Independent Smoothing
Public Function PayoffSmooth_Independent(arr_UndVal() As Variant, arr_optval() As Variant, dbl_spotstep As Double, _
Optional bln_LastUndCF As Boolean, Optional bln_Smooth As Boolean = True) As Variant

'Constant
Const int_pts4smoothing = 11
Dim int_count As Integer: int_count = UBound(arr_optval) - LBound(arr_optval) + 1

ReDim arr_payoff(1 To int_count) As Variant
ReDim arr_gamma(1 To int_count) As Variant
ReDim arr_abs_gamma(1 To int_count) As Variant
Dim col_pt2smooth As New Collection
Dim dbl_max_gamma As Double, dbl_avg_gamma As Double, dbl_X As Double, dbl_Y As Double, _
dbl_Z As Double
Dim j As Integer, n As Integer, k As Variant, l As Integer


'########################Payoff = MAX(Und, Opt, 0)#####################
For j = 1 To int_count
    arr_payoff(j) = dbl_3_max(arr_UndVal(j), arr_optval(j), 0)
    'arr_payoff(0, j) = Application.Max(arr_undval(0, j), arr_optval(0, j), 0)
Next j
'########################Payoff = MAX(Und, Opt, 0)#####################


'######################## Gamma Calc #########################################
For j = 1 To int_count
'~~Zerorizing i=0 and i =N-1 because not checked~~
 If (j = 1) Or (j = int_count) Then
    arr_gamma(j) = 0
Else
    dbl_X = arr_payoff(j - 1)
    dbl_Y = arr_payoff(j)
    dbl_Z = arr_payoff(j + 1)
     
    arr_gamma(j) = (dbl_X - 2 * dbl_Y + dbl_Z) / (dbl_spotstep * dbl_spotstep)
    arr_abs_gamma(j) = Abs(arr_gamma(j))
End If
Next j
'######################## Gamma Calc ##########################################


'######################## Max & Avg Gamma ####################################
dbl_max_gamma = max_array(arr_gamma)
'dbl_max_gamma = Application.Max(arr_gamma)

'~~ int_count -2 to exclude i=0 & i=N-1 ~~
dbl_avg_gamma = sumarray(arr_gamma) / (int_count - 2)
'dbl_avg_gamma = Application.Sum(arr_gamma) / (int_count - 2)
'######################## Max & Avg Gamma ####################################


'####################### Smoothing Screener ##################################
'~~Which spot step to smooth~~
n = 1
If bln_Smooth = True Then
    For j = 2 To int_count - 1
        If arr_abs_gamma(j) > (0.9 * dbl_max_gamma + 0.1 * dbl_avg_gamma) Then
             col_pt2smooth.Add j
        End If
    Next j
End If

'####################### Smoothing Screener ###################################


'<<<<<<<<<<<<<<<<<<<<<<<<< Payoff Smoothing >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
For Each k In col_pt2smooth
        arr_payoff = LinIntp_Payoff_Trapeze(k, arr_UndVal, int_pts4smoothing, _
        dbl_spotstep, arr_optval, arr_payoff, int_count, bln_LastUndCF)
Next k
'<<<<<<<<<<<<<<<<<<<<<<<<< Payoff Smoothing >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

PayoffSmooth_Independent = arr_payoff

End Function


Private Function LinIntp_Payoff_Trapeze(k As Variant, arr_UndVal() As Variant, int_pts4smoothing As Integer, _
dbl_spotstep As Double, arr_optval() As Variant, arr_payoff() As Variant, _
int_count As Integer, bln_LastUndCF As Boolean) As Variant

Dim dbl_left_smooth As Double, dbl_mid_smooth As Double, dbl_right_smooth As Double

ReDim arr_LinIntpUnd(0, 1 To int_pts4smoothing) As Variant
ReDim arr_LinIntpOpt(0, 1 To int_pts4smoothing) As Variant

        '__________________mid smoothing___________________________________
        '~~LinIntp ~~
        arr_LinIntpUnd = Linear_pts(k, arr_UndVal(), int_pts4smoothing)
        
        If bln_LastUndCF = "FALSE" Then _
        arr_LinIntpOpt = Linear_pts(k, arr_optval(), int_pts4smoothing)

        '~~Payoff, Trapeze & Smoothing ~~
        dbl_mid_smooth = dbl_smoothval(dbl_spotstep, int_pts4smoothing, arr_LinIntpUnd(), arr_LinIntpOpt(), bln_LastUndCF)
        '__________________________________________________________________

        '________________adjacent left smoothing__________________________
        '~~LinIntp ~~
   
        If (k - 1) > LBound(arr_UndVal) Then
            arr_LinIntpUnd = Linear_pts(k - 1, arr_UndVal(), int_pts4smoothing)
        End If

        If bln_LastUndCF = "FALSE" Then
            If (k - 1) > LBound(arr_optval) Then
                arr_LinIntpOpt = Linear_pts(k - 1, arr_optval(), int_pts4smoothing)
            End If
        End If

                
        '~~Payoff, Trapeze & Smoothing ~~
        If (k - 1) > LBound(arr_optval) Then
            dbl_left_smooth = dbl_smoothval(dbl_spotstep, int_pts4smoothing, arr_LinIntpUnd(), arr_LinIntpOpt(), bln_LastUndCF)
        End If

        '__________________________________________________________________

       '________________adjacent right smoothing____________________________
        '~~LinIntp ~~
        If (k + 1) < UBound(arr_UndVal) Then
        arr_LinIntpUnd = Linear_pts(k + 1, arr_UndVal(), int_pts4smoothing)
        End If
        
        If bln_LastUndCF = "FALSE" Then
            If (k + 1) < UBound(arr_optval) Then
                arr_LinIntpOpt = Linear_pts(k + 1, arr_optval(), int_pts4smoothing)
            End If
        End If
        
        '~~Payoff, Trapeze & Smoothing ~~
        If (k + 1) < UBound(arr_optval) Then
            dbl_right_smooth = dbl_smoothval(dbl_spotstep, int_pts4smoothing, arr_LinIntpUnd(), arr_LinIntpOpt(), bln_LastUndCF)
        End If

        '___________________________________________________________________


        '________________replace into Payoff = MAX(Und, Opt, 0)______________
         If dbl_3_max(0, dbl_mid_smooth) <> 0 Then arr_payoff(k) = dbl_mid_smooth
        
        If (k - 1) > LBound(arr_optval) Then
            If dbl_3_max(0, dbl_left_smooth) <> 0 Then arr_payoff(k - 1) = dbl_3_max(0, dbl_left_smooth)
        End If
        
        If (k + 1) < UBound(arr_optval) Then
            If dbl_3_max(0, dbl_right_smooth) <> 0 Then arr_payoff(k + 1) = dbl_3_max(0, dbl_right_smooth)
        End If

        '___________________________________________________________________

LinIntp_Payoff_Trapeze = arr_payoff

End Function

Private Function dbl_smoothval(dbl_spotstep As Double, int_pts4smoothing As Integer, _
arr_LinIntpUnd() As Variant, arr_LinIntpOpt() As Variant, Optional LastUndCF As Boolean) As Double
 
ReDim arr_payoffaftersmooth(1 To int_pts4smoothing) As Variant
ReDim arr_trapeze(1 To int_pts4smoothing) As Variant

Dim dbl_sumval As Double
Dim dbl_trapeze_const As Double: dbl_trapeze_const = dbl_spotstep * 2 / 10 / dbl_spotstep / 2
Dim i As Integer

For i = 1 To int_pts4smoothing
    
    If LastUndCF = "TRUE" Then
        arr_payoffaftersmooth(i) = dbl_3_max(arr_LinIntpUnd(i), 0)
    Else
      arr_payoffaftersmooth(i) = dbl_3_max(arr_LinIntpUnd(i), arr_LinIntpOpt(i), 0)
    End If
    
    
    If (i = 1) Or (i = int_pts4smoothing) Then
        arr_trapeze(i) = 0
    Else
        arr_trapeze(i) = arr_payoffaftersmooth(i)
    End If
                
   dbl_sumval = dbl_sumval + arr_payoffaftersmooth(i) + arr_trapeze(i)
        
  
Next i
    
    dbl_smoothval = dbl_sumval / 2 * dbl_trapeze_const
    

End Function

Private Function Linear_pts(k As Variant, arr_val() As Variant, int_pts4smoothing As Integer) As Variant

Dim i As Integer
Dim upper_p As Double, lower_p As Double, middle_p As Double, int_midpt As Integer
ReDim arr_LinIntp(1 To int_pts4smoothing) As Variant

upper_p = arr_val(k + 1)
middle_p = arr_val(k)
lower_p = arr_val(k - 1)

int_midpt = (int_pts4smoothing + 1) / 2

arr_LinIntp(1) = (lower_p + middle_p) / 2
arr_LinIntp(int_pts4smoothing) = (upper_p + middle_p) / 2
arr_LinIntp(int_midpt) = middle_p

For i = 2 To int_midpt - 1
    arr_LinIntp(i) = arr_LinIntp(i - 1) _
    + (arr_LinIntp(int_midpt) - arr_LinIntp(1)) / (int_midpt - 1)
Next i
    
For i = (int_midpt + 1) To (int_pts4smoothing - 1)
    arr_LinIntp(i) = arr_LinIntp(i - 1) + _
    (arr_LinIntp(int_pts4smoothing) - arr_LinIntp(int_midpt)) / (int_pts4smoothing - int_midpt)
Next i

Linear_pts = arr_LinIntp
 
        
End Function

Private Function dbl_3_max(x As Variant, y As Variant, Optional z As Variant) As Double
Const aa = 3
Dim i As Integer
ReDim arr_xyz(1 To aa) As Double

arr_xyz(1) = x
arr_xyz(2) = y

        If IsMissing(z) Then
            arr_xyz(3) = 0
        Else
            arr_xyz(3) = z
        End If

dbl_3_max = x

For i = 2 To aa
    If arr_xyz(i) > dbl_3_max Then _
    dbl_3_max = arr_xyz(i)
Next i

End Function

Private Function sumarray(arr_sum() As Variant) As Double

Dim int_k As Integer: int_k = UBound(arr_sum) - LBound(arr_sum) + 1
Dim dbl_sumbot As Double: dbl_sumbot = 0
Dim p As Integer

sumarray = 0

For p = 1 To int_k

sumarray = sumarray + arr_sum(p)

Next p
 
End Function
Private Function max_array(arr() As Variant) As Double

Dim int_l As Integer: int_l = UBound(arr) - LBound(arr) + 1
Dim qq As Integer

max_array = arr(1)

For qq = 2 To int_l

If max_array < arr(qq) Then _
max_array = arr(qq)

Next qq

End Function
Private Function LU_Decompose(arr_x() As Double, bln_Lower As Boolean) As Double()

'Provide L or U matrix
Dim i, j, k As Integer

Dim int_length As Integer
int_length = UBound(arr_x)

Dim arr_L() As Double
ReDim arr_L(1 To int_length, 1 To int_length) As Double

Dim arr_U() As Double
ReDim arr_U(1 To int_length, 1 To int_length) As Double

For i = 1 To int_length
    arr_L(i, 1) = arr_x(i, 1)
    
    arr_U(i, i) = 1
    
    If i > 1 Then
        arr_U(1, i) = arr_x(1, i) / arr_L(1, 1)
    End If
Next i

Dim dbl_sum As Double

For j = 2 To int_length

    For i = 2 To j - 1
        dbl_sum = 0
        
        For k = 1 To i - 1
            dbl_sum = dbl_sum + arr_L(i, k) * arr_U(k, j)
        Next k
        
        arr_U(i, j) = (arr_x(i, j) - dbl_sum) / arr_L(i, i)
    Next i

    For i = j To int_length
        dbl_sum = 0
        For k = 1 To j - 1
            dbl_sum = dbl_sum + arr_L(i, k) * arr_U(k, j)
        Next k
        arr_L(i, j) = (arr_x(i, j) - dbl_sum)
    Next i

Next j

If bln_Lower = True Then
    LU_Decompose = arr_L
Else
    LU_Decompose = arr_U
End If

End Function

Private Function LU_Decompose_Simplified(arr_x() As Double) As Double()

'Provide L or U matrix
Dim int_i As Integer

Dim arr_output() As Double
ReDim arr_output(1 To 2, 1 To UBound(arr_x), 1 To UBound(arr_x)) As Double

For int_i = 1 To UBound(arr_x)

arr_output(2, int_i, int_i) = 1

    If int_i = 1 Then
        arr_output(1, int_i, int_i) = arr_x(int_i, int_i)
        arr_output(2, int_i, int_i + 1) = arr_x(int_i, int_i + 1) / arr_output(1, int_i, int_i)
    ElseIf int_i = UBound(arr_x) Then
        arr_output(1, int_i, int_i - 1) = arr_x(int_i, int_i - 1)
        arr_output(1, int_i, int_i) = arr_x(int_i, int_i) - arr_output(1, int_i, int_i - 1) * arr_output(2, int_i - 1, int_i)
    Else
        arr_output(1, int_i, int_i - 1) = arr_x(int_i, int_i - 1)
        arr_output(1, int_i, int_i) = arr_x(int_i, int_i) - arr_output(1, int_i, int_i - 1) * arr_output(2, int_i - 1, int_i)
        arr_output(2, int_i, int_i + 1) = arr_x(int_i, int_i + 1) / arr_output(1, int_i, int_i)
    End If

Next int_i

LU_Decompose_Simplified = arr_output

End Function


'''''''''TO BE PUT UNDER MODULE - HW_FUNCTION'''''''''
Private Function HW_Variance(dbl_Sigma As Double, dbl_MeanRev As Double, dbl_T1 As Double, dbl_T2 As Double) As Double

Dim dbl_Output As Double
    
dbl_Output = dbl_Sigma * dbl_Sigma / (2 * dbl_MeanRev) * (Exp(2 * dbl_MeanRev * dbl_T2) - Exp(2 * dbl_MeanRev * dbl_T1))
HW_Variance = dbl_Output

End Function
'''''''''TO BE PUT UNDER MODULE - HW_FUNCTION'''''''''


