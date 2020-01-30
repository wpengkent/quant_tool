Option Explicit

' ## MEMBER DATA
' Curve dependencies
Private irc_Disc As Data_IRCurve, irc_Est As Data_IRCurve, cvl_volcurve As Data_CapVols

' Variable dates
Private lng_ValDate As Long

' Dynamic stored values
Private dblLst_Rates As Collection, dblLst_IntFlows As Collection, dblLst_DFs As Collection, dbl_StartDF As Double
Private intLst_IsMV As Collection, intLst_IsCash As Collection, str_PrnAtStartPnlType As String
Private dblLst_FlowDurations As Collection  ' Time between valuation and payment date
Private dbl_ZSpread As Double

' Static values - general
Private fld_Params As IRLegParams, dic_CurveDependencies As Dictionary
Private Const str_NotUsed As String = "-", str_Daycount_Duration As String = "ACT/365", bln_ActActFwdGeneration As Boolean = False
Private bln_IsFixed As Boolean
Private bln_IsFixInArrears As Boolean, bln_IsDisableConvAdj As Boolean, bln_StubInterpolate As Boolean
Private bln_StubUpfront As Boolean, bln_StubArrears As Boolean

' Static values - counts and measures
Private int_CalcsPerPmt As Integer
Private str_PeriodLength As String

' Static values - dates, calendars and durations
Private lngLst_PeriodStart As Collection, lngLst_PeriodEnd As Collection
Private lngLst_EstStart As Collection, lngLst_EstEnd As Collection, lngLst_PmtDates As Collection
Private lng_startdate As Long, lng_EndDate As Long
Private cal_pmt As Calendar, cal_est As Calendar
Private dblLst_CalcPeriodDurations As Collection  ' Time between period start and end
Private dblLst_EstPeriodDurations As Collection  ' Time between estimation period start and end

' Static values - rates, values and factors
Private dic_GlobalStaticInfo As Dictionary, dic_CurveSet As Dictionary
Private dblLst_Margins As Collection
Private dblLst_AmortFactors As Collection
Private dbl_PrnAtStart As Double, dblLst_PrnFlows As Collection

' Range Accrual
Public str_Daycount As String
Private dblLst_IntFlows_Rate As Collection
Private dblLst_IntFlows_Margin As Collection

' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As IRLegParams, Optional dic_CurveSetInput As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing)

    ' Set static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    fld_Params = fld_ParamsInput

    ' Initialize dynamic values
    lng_ValDate = fld_ParamsInput.ValueDate

    ' Determine leg type
    bln_IsFixed = (fld_Params.index = str_NotUsed)

    ' Force floating rate estimation if margin is specified
    If bln_IsFixed = False And fld_Params.RateOrMargin <> 0 Then fld_Params.FloatEst = True

    ' Range Accrual
    str_Daycount = fld_ParamsInput.Daycount

    ' Set calendars
    Dim cas_Found As CalendarSet: Set cas_Found = dic_GlobalStaticInfo(StaticInfoType.CalendarSet)
    cal_pmt = cas_Found.Lookup_Calendar(fld_Params.PmtCal)
    If bln_IsFixed = False Then cal_est = cas_Found.Lookup_Calendar(fld_Params.estcal)

    ' Derive number of payment dates
    Dim int_NumPmts As Integer, lng_WindowStart As Long, lng_WindowEnd As Long
    If fld_Params.GenerationLimitPoint = 0 Then
        int_NumPmts = Calc_NumPeriods(fld_Params.Term, fld_Params.PmtFreq)

    Else
        ' Determine number of payments based on the specified start and end dates
        If fld_Params.IsFwdGeneration = True Then
            lng_WindowStart = fld_Params.GenerationRefPoint
            lng_WindowEnd = fld_Params.GenerationLimitPoint
        Else
            lng_WindowStart = fld_Params.GenerationLimitPoint
            lng_WindowEnd = fld_Params.GenerationRefPoint
        End If

        int_NumPmts = Calc_NumPmtsInWindow(lng_WindowStart, lng_WindowEnd, fld_Params.PmtFreq, cal_pmt, fld_Params.BDC, _
            fld_Params.EOM, fld_Params.IsFwdGeneration)
    End If

    ' Derive number of calculation dates
    Dim int_NumCalcs As Integer
    If bln_IsFixed = True Then
        int_NumCalcs = int_NumPmts
        str_PeriodLength = fld_Params.PmtFreq
    Else
        If fld_Params.index = fld_Params.PmtFreq Then
            int_NumCalcs = int_NumPmts
        ElseIf fld_Params.GenerationLimitPoint = 0 Then
            int_NumCalcs = Calc_NumPeriods(fld_Params.Term, fld_Params.index)
        Else
            int_NumCalcs = Calc_NumPmtsInWindow(lng_WindowStart, lng_WindowEnd, fld_Params.index, cal_pmt, fld_Params.BDC, _
                fld_Params.EOM, fld_Params.IsFwdGeneration)
        End If
        str_PeriodLength = fld_Params.index
    End If

    int_CalcsPerPmt = int_NumCalcs / int_NumPmts

    ' Derive dates
    Call FillPeriodDates(int_NumCalcs)

    If bln_IsFixed = False Then Call FillEstDates
    Call ModifyStartDates
    Call FillPmtDates  ' For bonds with past fixings, floating estimation must be turned on
    Call CategorizeFlows

    ' Store durations
    Call FillFlowDurations
    Call FillPeriodDurations

    ' Read amortization schedule if it exists
    Call FillAmortFactors

    'Inputs for In Arrears Swap
    bln_IsFixInArrears = fld_Params.FixInArrears
    bln_IsDisableConvAdj = fld_Params.DisableConvAdj
    bln_StubInterpolate = fld_Params.StubInterpolate

    ' Set dependent curves
    If dic_CurveSetInput Is Nothing Then
        Set irc_Disc = GetObject_IRCurve(fld_Params.Curve_Disc, True, False)
        If bln_IsFixed = False Then Set irc_Est = GetObject_IRCurve(fld_Params.Curve_Est, True, False)

        'Inputs for In Arrears Swap
        If bln_IsFixed = False And bln_IsFixInArrears = True And bln_IsDisableConvAdj = False Then
            Set cvl_volcurve = GetObject_CapVolSurf(fld_Params.CCY & "_" & fld_Params.index, fld_Params.RateOrMargin, True, False)
        End If

    Else
        Set dic_CurveSet = dic_CurveSetInput
        Dim dic_IRCurves As Dictionary: Set dic_IRCurves = dic_CurveSet(CurveType.IRC)
        Set irc_Disc = dic_IRCurves(fld_Params.Curve_Disc)
        If bln_IsFixed = False Then Set irc_Est = dic_IRCurves(fld_Params.Curve_Est)

        'Inputs for In Arrears Swap
        If bln_IsFixed = False And bln_IsFixInArrears = True And bln_IsDisableConvAdj = False Then
            Set cvl_volcurve = GetObject_CapVolSurf(fld_Params.CCY & "_" & fld_Params.index, fld_Params.RateOrMargin, True, False)
        End If
    End If

    ' Fill rates and flows
    Call FillPrincipalFlows  ' Fill if intermediate payments are enabled, otherwise fill with zero
    Call FillRates
    If bln_IsFixed = False Then Call FillFixings
    Call FillIntFlows  ' Calculate undiscounted flows

    Call FillIntFlows_Rate  ' Calculate undiscounted Rates Flows (for Range Accrual)
    Call FillIntFlows_Margin  ' Calculate undiscounted Margin Flows (for Range Accrual)

    Call FillDFs  ' Read and store discount factors

    If Me.IsMissingFixings = True Then Debug.Print "## Error - Missing fixing(s) for trade: " & fld_Params.TradeID

    ' Determine curve dependencies
    Set dic_CurveDependencies = New Dictionary
    dic_CurveDependencies.CompareMode = CompareMethod.TextCompare
    Call dic_CurveDependencies.Add(irc_Disc.CurveName, True)
    If Not irc_Est Is Nothing Then
        If dic_CurveDependencies.Exists(irc_Est.CurveName) = False Then Call dic_CurveDependencies.Add(irc_Est.CurveName, True)
    End If

    Dim fxs_Spots As Data_FXSpots
    If dic_CurveSet Is Nothing Then
        Set fxs_Spots = GetObject_FXSpots(True, dic_GlobalStaticInfo)
    Else
        Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
    End If

    Set dic_CurveDependencies = Convert_MergeDicts(dic_CurveDependencies, fxs_Spots.Lookup_CurveDependencies(fld_Params.CCY))
End Sub

Public Property Get marketvalue() As Double
    marketvalue = CalcValue("MV")
End Property

Public Property Get Cash() As Double
    Cash = CalcValue("CASH")
End Property

Public Property Get PnL() As Double
    PnL = marketvalue + Cash
End Property

Public Property Get Params() As IRLegParams
    Params = fld_Params
End Property

Public Property Get PeriodStart() As Collection
    Set PeriodStart = lngLst_PeriodStart
End Property

Public Property Get PeriodEnd() As Collection
    Set PeriodEnd = lngLst_PeriodEnd
End Property

Public Property Get EstStart() As Collection
    Set EstStart = lngLst_EstStart
End Property

Public Property Get EstEnd() As Collection
    Set EstEnd = lngLst_EstEnd
End Property

Public Property Get MatDate() As Long
    MatDate = lng_EndDate
End Property

Public Property Get FixedFloat() As String
    FixedFloat = fld_Params.FixedFloat
End Property
Public Property Get ForwRates() As Collection
    Set ForwRates = dblLst_Rates
End Property
Public Property Get Lockoutmode() As String
    Lockoutmode = fld_Params.Lockoutmode
End Property
Public Property Get IsFixed() As Boolean
    IsFixed = bln_IsFixed
End Property
'KL - ADD for CRA
'START
Public Property Get Calendar() As Calendar
    Calendar = cal_pmt
End Property
Public Property Get Term() As String
    Term = str_PeriodLength
End Property
Public Property Get EOM() As String
    EOM = fld_Params.EOM
End Property
'END
Public Property Get IsMissingFixings() As Boolean
    ' ## Returns true if any required fixings are not supplied
    Dim int_ctr As Integer
    Dim bln_Output As Boolean: bln_Output = False

    If bln_IsFixed = False Then
        For int_ctr = 1 To lngLst_PeriodEnd.count
            If lngLst_PeriodStart(int_ctr) < lng_ValDate Then
                ' A fixing is required because cannot estimate the rate
                If fld_Params.Fixings Is Nothing Then
                    ' No fixings exist at all
                    bln_Output = True
                    Exit For
                Else
                    If fld_Params.Fixings.Exists(lngLst_PeriodStart(int_ctr)) = False Then
                        ' No fixings exist for the particular date
                        bln_Output = True
                        Exit For
                    End If
                End If
            Else
                ' Already past value date, all subsequent periods will also be in the future
                Exit For
            End If
        Next int_ctr
    End If

    IsMissingFixings = bln_Output
End Property

Public Property Get SwaptionScalingFactor() As Double
    ' ## Return sumproduct of days and DF
    Dim dbl_Output As Double: dbl_Output = 0
    Dim int_CalcsToPmt As Integer: int_CalcsToPmt = int_CalcsPerPmt

    Dim int_ctr As Integer, int_PmtNum As Integer
    For int_ctr = 1 To lngLst_PeriodEnd.count
        int_PmtNum = WorksheetFunction.RoundUp(int_ctr / int_CalcsPerPmt, 1)
        dbl_Output = dbl_Output + calc_yearfrac(lngLst_PeriodStart(int_ctr), lngLst_PeriodEnd(int_ctr), _
            fld_Params.Daycount, fld_Params.PmtFreq, bln_ActActFwdGeneration) * dblLst_DFs(int_PmtNum) _
            * dblLst_AmortFactors(int_ctr)
    Next int_ctr

    SwaptionScalingFactor = dbl_Output
End Property

Public Property Get ZSpread() As Double
    ZSpread = dbl_ZSpread
End Property

Public Property Get CurveDependencies() As Dictionary
    Set CurveDependencies = dic_CurveDependencies
End Property

Public Property Get RateOrMargin() As Double
    RateOrMargin = fld_Params.RateOrMargin
End Property
'Alv 20181017 for RngAcc
Public Property Get NbofDays() As String
    NbofDays = fld_Params.NbofDays
End Property
'Alv 20181017 for RngAcc
Public Property Get calendar_est() As Calendar
    calendar_est = fld_Params.cal_est
End Property

'WL 20181002 for hull white calibration code
Public Property Get FixLegCashFlow() As Collection
    Set FixLegCashFlow = dblLst_IntFlows
End Property
'WL 20181002 for hull white calibration code
Public Property Get FixLegDF() As Collection
    Set FixLegDF = dblLst_DFs
End Property

' ## METHODS - GREEKS
Public Function Calc_DV01_Analytical(str_curve As String) As Double
    Dim dbl_Output As Double
    If dic_CurveDependencies.Exists(str_curve) Then
        Dim dbl_DV01_DiscCurve As Double, dbl_DV01_EstCurve As Double

        ' Discount curve DV01
        If irc_Disc.CurveName = str_curve Then
            dbl_DV01_DiscCurve = (Calc_SumProductOnList(dblLst_IntFlows, dblLst_DFs, dblLst_FlowDurations) _
                + Calc_SumProductOnList(dblLst_PrnFlows, dblLst_DFs, dblLst_FlowDurations)) * -0.0001
        Else
            dbl_DV01_DiscCurve = 0
        End If

        ' Estimation curve DV01
        Dim int_NumPeriods As Integer, dblArr_NotionalFactors() As Double, dblLst_DFs_PeriodEnd As Collection, int_ctr As Integer
        dbl_DV01_EstCurve = 0
        If Not irc_Est Is Nothing Then
            If irc_Est.CurveName = str_curve Then
            ' Gather notional required for a unit interest flow
            int_NumPeriods = lngLst_PeriodEnd.count
            ReDim dblArr_NotionalFactors(1 To int_NumPeriods) As Double
            Set dblLst_DFs_PeriodEnd = New Collection

            For int_ctr = 1 To int_NumPeriods
                dblArr_NotionalFactors(int_ctr) = irc_Est.Lookup_Rate(lngLst_EstStart(int_ctr), lngLst_EstEnd(int_ctr), _
                    fld_Params.Daycount, , fld_Params.PmtFreq, , True)
                Call dblLst_DFs_PeriodEnd.Add(irc_Disc.Lookup_Rate(lng_ValDate, lngLst_PeriodEnd(int_ctr), "DF", , , False))
            Next int_ctr

            ' Create array with full principal as each element
            dbl_DV01_EstCurve = Calc_SumProductOnList(dblArr_NotionalFactors, dblLst_CalcPeriodDurations, _
                dblLst_EstPeriodDurations, dblLst_DFs_PeriodEnd, dblLst_AmortFactors) * fld_Params.Notional / 100 * 0.0001
            End If
        End If

        dbl_Output = dbl_DV01_DiscCurve + dbl_DV01_EstCurve
    Else
        dbl_Output = 0
    End If

    Calc_DV01_Analytical = dbl_Output
End Function


' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Sub SetRateOrMargin(dbl_NewValue As Double)
    ' ## Alter stored parameter, used for solving par rate
    fld_Params.RateOrMargin = dbl_NewValue
    Call RecalcIntFlows
End Sub

Public Sub SetZSpread(dbl_spread As Double)
    ' ## Used when forcing to a known market price.  Affects only discounting
    dbl_ZSpread = dbl_spread
    Call FillDFs
End Sub

Public Sub SetCurveState(str_curve As String, enu_State As CurveState_IRC, Optional int_PillarIndex As Integer = 0)
    ' ## For temporary zero shifts only, such as during a finite differencing calculation
    If irc_Disc.CurveName = str_curve Then
        Call irc_Disc.SetCurveState(enu_State, int_PillarIndex)
        Call FillDFs
    End If

    If Not irc_Est Is Nothing Then
        If irc_Est.CurveName = str_curve Then
            Call irc_Est.SetCurveState(enu_State, int_PillarIndex)
            Call RecalcIntFlows
        End If
    End If
End Sub

Public Sub SetValDate(lng_Input As Long)
    ' ## Set stored value date and refresh values dependent on the value date
    lng_ValDate = lng_Input
    Call FillDFs
    Call CategorizeFlows
    Call FillFlowDurations
End Sub

Public Sub ReplaceCurveObject(str_CurveName As String, irc_Curve As Data_IRCurve)
    ' ## If any curve names match the name specified, replace with the specified curve object
    ' ## Used for bootstrapping procedure to ensure the curve underlying the swap is the same as the curve being updated by the process

    If irc_Disc.CurveName = str_CurveName Then Set irc_Disc = irc_Curve
    If Not irc_Est Is Nothing Then
        If irc_Est.CurveName = str_CurveName Then Set irc_Est = irc_Curve
    End If
End Sub

Public Sub HandleUpdate_IRC(str_CurveName As String)
    ' ## Update stored information given that the specified curve has been updated
    If fld_Params.Curve_Disc = str_CurveName Or str_CurveName = "ALL" Then Call FillDFs

    If Not irc_Est Is Nothing Then
        If fld_Params.Curve_Est = str_CurveName Or str_CurveName = "ALL" Then Call RecalcIntFlows
    End If
End Sub


' ## METHODS - PUBLIC
Public Function DependsOnFuture(str_BootstrappedCurve As String, lng_MatDate) As Boolean
    ' ## If estimation depends on the bootstrapped curve and the last estimation date is beyond the maturity date, return True
    Dim bln_Output As Boolean: bln_Output = False

    If fld_Params.FloatEst = True Then
        If fld_Params.Curve_Est = str_BootstrappedCurve Then
            If lngLst_EstEnd.item(lngLst_EstEnd.count) > lng_MatDate Then
                bln_Output = True
            End If
        End If
    End If

    DependsOnFuture = bln_Output
End Function

Public Function SolveParRate(irl_LegToMatch As IRLeg, Optional bln_ResetToOrigPar As Boolean = True, _
    Optional dbl_ExistingMV_SolveLeg As Double = 0, Optional dbl_ExistingMV_MatchLeg As Double = 0) As Double
    ' ## Find rate or margin at which the NPV is zero, leaving the other leg unchanged
    ' ## Existing MVs are expressed in their native currencies
    Dim dbl_Output As Double
    Dim dbl_OrigRate As Double: dbl_OrigRate = fld_Params.RateOrMargin

    ' Convert MV of leg B to a MV in the currency of leg A at the value date
    Dim dbl_FXConv As Double
    If dic_CurveSet Is Nothing Then
        dbl_FXConv = cyGetFXFwd(irl_LegToMatch.Params.CCY, fld_Params.CCY, irl_LegToMatch.Params.ValueDate, False)
    Else
        Dim fxs_Spots As Data_FXSpots
        Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
        dbl_FXConv = fxs_Spots.Lookup_Fwd(irl_LegToMatch.Params.CCY, fld_Params.CCY, irl_LegToMatch.Params.ValueDate, False)
    End If

    ' Store static parameters for secant solver
    Dim dic_SecantParams As Dictionary: Set dic_SecantParams = New Dictionary
    Call dic_SecantParams.Add("irl_Leg", Me)

    ' Solve
    Dim dic_SecantOutputs As Dictionary: Set dic_SecantOutputs = New Dictionary
    SolveParRate = Solve_Secant_IRleg(ThisWorkbook, "SolverFuncXY_ParToMV", dic_SecantParams, fld_Params.RateOrMargin, _
        fld_Params.RateOrMargin + 1, (irl_LegToMatch.marketvalue - dbl_ExistingMV_MatchLeg) * dbl_FXConv + dbl_ExistingMV_SolveLeg, _
        fld_Params.Notional * 0.000000000000001, 50, -100, dic_SecantOutputs)

    ' Reset back to original rate if required
    If bln_ResetToOrigPar = True Then Call SetRateOrMargin(dbl_OrigRate)
End Function

Private Function Solve_Secant_IRleg(wbk_Caller As Workbook, str_XYFunction As String, dic_StaticParams As Dictionary, _
    dbl_InitialX1 As Double, dbl_InitialX2 As Double, dbl_TargetY As Double, dbl_Tolerance As Double, _
    int_MaxIterations As Integer, dbl_FallBackValue As Double, ByRef dic_SecondaryOutputs As Dictionary) As Double
    ' ## Perform the secant method to solve for the input which sets the function to the target value

    Dim dbl_Output As Double
    Dim dbl_SecantX1 As Double, dbl_SecantX2 As Double, dbl_SecantX3 As Double
    Dim dbl_SecantY1 As Double, dbl_SecantY2 As Double, dbl_SecantY3 As Double
    Call dic_SecondaryOutputs.RemoveAll

    ' Set first initial guess
    dbl_SecantX1 = dbl_InitialX1
    dbl_SecantY1 = Application.Run("'" & wbk_Caller.Name & "'!" & str_XYFunction, dbl_SecantX1, dic_StaticParams, dic_SecondaryOutputs) - dbl_TargetY

    ' Set second intitial guess
    dbl_SecantX2 = dbl_InitialX2
    dbl_SecantY2 = Application.Run("'" & wbk_Caller.Name & "'!" & str_XYFunction, dbl_SecantX2, dic_StaticParams, dic_SecondaryOutputs) - dbl_TargetY

    ' Prepare for iteration
    Dim int_IterCtr As Integer: int_IterCtr = 0
    Dim bln_Solvable As Boolean: bln_Solvable = True

    Do
        If dbl_SecantY2 - dbl_SecantY1 = 0 Then
            ' Allow greater tolerance if having difficulty solving
            If Abs(dbl_SecantY3) > (dbl_Tolerance * 100) Or int_IterCtr = 0 Then
                ' No solution even with looser tolerance
                dbl_SecantY3 = 0
                bln_Solvable = False
            Else
                ' Solved to looser tolerance
                Exit Do
            End If
        End If

        If bln_Solvable = True Then
            int_IterCtr = int_IterCtr + 1

            ' Set new guess
            dbl_SecantX3 = dbl_SecantX2 - dbl_SecantY2 * (dbl_SecantX2 - dbl_SecantX1) / (dbl_SecantY2 - dbl_SecantY1)
            dbl_SecantY3 = Application.Run("'" & wbk_Caller.Name & "'!" & str_XYFunction, dbl_SecantX3, dic_StaticParams, dic_SecondaryOutputs) - dbl_TargetY
        End If

        dbl_SecantX1 = dbl_SecantX2
        dbl_SecantY1 = dbl_SecantY2
        dbl_SecantX2 = dbl_SecantX3
        dbl_SecantY2 = dbl_SecantY3
    Loop Until Abs(dbl_SecantY3) < dbl_Tolerance Or int_IterCtr >= int_MaxIterations

    ' Output final solution if possible, otherwise output the fallback
    If bln_Solvable = True Then dbl_Output = dbl_SecantX3 Else dbl_Output = dbl_FallBackValue

    Solve_Secant_IRleg = dbl_Output
    Call dic_SecondaryOutputs.Add("NumIterations", int_IterCtr)
    Call dic_SecondaryOutputs.Add("Solvable", bln_Solvable)
End Function
Public Sub ForceMVToValue(dbl_TargetMV As Double)
    ' ## Set ZSpread such that the MV equals the market price
    ' Store static parameters for secant solver
    Dim dic_SecantParams As Dictionary: Set dic_SecantParams = New Dictionary
    Call dic_SecantParams.Add("irl_Leg", Me)

    ' Solve
    Dim dic_SecantOutputs As Dictionary: Set dic_SecantOutputs = New Dictionary
    Call Solve_Secant(ThisWorkbook, "SolverFuncXY_ZSpreadToMV", dic_SecantParams, dbl_ZSpread, _
        dbl_ZSpread + 1, dbl_TargetMV, fld_Params.Notional * 0.000000000000001, 50, -100, dic_SecantOutputs)

    ' Handle unsolvable case
    If dic_SecantOutputs("Solvable") = False Then
        Call Me.SetZSpread(0)
        Debug.Print "## ERROR - Could not force MV to value.  Trade: " & fld_Params.TradeID
        Debug.Assert False
    End If
End Sub

Public Function Calc_BSOptionValue(enu_Direction As OptionDirection, dbl_Strike As Double, int_Deduction As Integer, _
    cal_Deduction As Calendar, bln_IsDiscounted As Boolean, Optional dblLst_CapletVols As Collection = Nothing, _
    Optional dbl_CapVol As Double = -1, Optional str_ValueType As String = "PNL", Optional int_CapletIndex As Integer = -1) As Double
    ' ## Values cap (1) or floor (-1), can use specific vols for each caplet or an overall cap vol
    Dim dbl_Output As Double
    Dim int_ctr As Integer
    Dim lng_ActivePeriodStart As Long, lng_ActiveOptionMat As Long, lng_ActivePeriodEnd As Long
    Dim bln_ActiveInclude As Boolean

    If bln_IsFixed = False Then
        ' Value each caplet using either the cap vol or specified caplet vol, then sum
        Dim dbl_ActiveTenor As Double, dbl_ActiveFwd As Double, dbl_ActiveDF As Double
        Dim dbl_ActiveCapletVol As Double, int_PmtNum As Integer
        Dim dbl_ActiveTimeToMat As Double
        For int_ctr = 1 To lngLst_PeriodEnd.count
            lng_ActivePeriodStart = lngLst_PeriodStart(int_ctr)
            lng_ActiveOptionMat = date_workday(lng_ActivePeriodStart, int_Deduction, cal_Deduction.HolDates, cal_Deduction.Weekends)
            lng_ActivePeriodEnd = lngLst_PeriodEnd(int_ctr)
            int_PmtNum = WorksheetFunction.RoundUp(int_ctr / int_CalcsPerPmt, 1)

            ' Filter out caplets not relevant to the specified calculation
            Select Case str_ValueType
                Case "MV"
                    If lng_ActivePeriodEnd > lng_ValDate Then bln_ActiveInclude = True Else bln_ActiveInclude = False
                Case "CASH"
                    If lng_ActivePeriodEnd <= lng_ValDate Then bln_ActiveInclude = True Else bln_ActiveInclude = False
                Case "PNL"
                    bln_ActiveInclude = True
            End Select

            ' Only include caplet if it is the specified index, or if all caplets are included (-1)
            If int_CapletIndex <> -1 And int_ctr <> int_CapletIndex Then bln_ActiveInclude = False

            If bln_ActiveInclude = True Then
                dbl_ActiveTimeToMat = calc_yearfrac(lng_ValDate, lng_ActiveOptionMat, "ACT/365")
                dbl_ActiveTenor = calc_yearfrac(lng_ActivePeriodStart, lng_ActivePeriodEnd, fld_Params.Daycount, fld_Params.PmtFreq, bln_ActActFwdGeneration)
                dbl_ActiveFwd = dblLst_Rates(int_ctr) + dblLst_Margins(int_ctr)
                If bln_IsDiscounted = True Then dbl_ActiveDF = dblLst_DFs(int_PmtNum) Else dbl_ActiveDF = 1

                If dbl_CapVol = -1 Then
                    dbl_ActiveCapletVol = dblLst_CapletVols(int_ctr)
                Else
                    dbl_ActiveCapletVol = dbl_CapVol
                End If

                dbl_Output = dbl_Output + Calc_BSPrice_Vanilla(enu_Direction, dbl_ActiveFwd, dbl_Strike, dbl_ActiveTimeToMat, _
                    dbl_ActiveCapletVol) / 100 * dbl_ActiveDF * dbl_ActiveTenor * fld_Params.Notional * dblLst_AmortFactors(int_ctr)
            End If
        Next int_ctr
    End If

    Calc_BSOptionValue = dbl_Output
End Function

Public Function Calc_BSOptionValueDigitalSmileOn(enu_Direction As OptionDirection, dbl_Strike As Double, int_Deduction As Integer, _
    cal_Deduction As Calendar, bln_IsDiscounted As Boolean, Optional dblLst_CapletVols As Collection = Nothing, _
    Optional dblLst_ShiftedCapletVols As Collection = Nothing, _
    Optional dbl_CapVol As Double = -1, Optional str_ValueType As String = "PNL", Optional int_CapletIndex As Integer = -1) As Double
    ' ## Values cap (1) or floor (-1), can use specific vols for each caplet or an overall cap vol
    Dim dbl_Output As Double
    Dim int_ctr As Integer
    Dim lng_ActivePeriodStart As Long, lng_ActiveOptionMat As Long, lng_ActivePeriodEnd As Long
    Dim bln_ActiveInclude As Boolean

    If bln_IsFixed = False Then
        ' Value each caplet using either the cap vol or specified caplet vol, then sum
        Dim dbl_ActiveTenor As Double, dbl_ActiveFwd As Double, dbl_ActiveDF As Double
        Dim dbl_ActiveCapletVol As Double, dbl_ShiftedCapletVol As Double, int_PmtNum As Integer
        Dim dbl_ActiveTimeToMat As Double
        For int_ctr = 1 To lngLst_PeriodEnd.count
            lng_ActivePeriodStart = lngLst_PeriodStart(int_ctr)
            lng_ActiveOptionMat = date_workday(lng_ActivePeriodStart, int_Deduction, cal_Deduction.HolDates, cal_Deduction.Weekends)
            lng_ActivePeriodEnd = lngLst_PeriodEnd(int_ctr)
            int_PmtNum = WorksheetFunction.RoundUp(int_ctr / int_CalcsPerPmt, 1)

            ' Filter out caplets not relevant to the specified calculation
            Select Case str_ValueType
                Case "MV"
                    If lng_ActivePeriodEnd > lng_ValDate Then bln_ActiveInclude = True Else bln_ActiveInclude = False
                Case "CASH"
                    If lng_ActivePeriodEnd <= lng_ValDate Then bln_ActiveInclude = True Else bln_ActiveInclude = False
                Case "PNL"
                    bln_ActiveInclude = True
            End Select

            ' Only include caplet if it is the specified index, or if all caplets are included (-1)
            If int_CapletIndex <> -1 And int_ctr <> int_CapletIndex Then bln_ActiveInclude = False

            If bln_ActiveInclude = True Then
                dbl_ActiveTimeToMat = calc_yearfrac(lng_ValDate, lng_ActiveOptionMat, "ACT/365")
                dbl_ActiveTenor = calc_yearfrac(lng_ActivePeriodStart, lng_ActivePeriodEnd, fld_Params.Daycount, fld_Params.PmtFreq, bln_ActActFwdGeneration)
                dbl_ActiveFwd = dblLst_Rates(int_ctr) + dblLst_Margins(int_ctr)
                If bln_IsDiscounted = True Then dbl_ActiveDF = dblLst_DFs(int_PmtNum) Else dbl_ActiveDF = 1

                If dbl_CapVol = -1 Then
                    dbl_ActiveCapletVol = dblLst_CapletVols(int_ctr)
                    dbl_ShiftedCapletVol = dblLst_ShiftedCapletVols(int_ctr)
                Else
                    dbl_ActiveCapletVol = dbl_CapVol
                    dbl_ShiftedCapletVol = dbl_CapVol
                End If

                dbl_Output = dbl_Output + Calc_BSPrice_DigitalSmileOn(enu_Direction, dbl_ActiveFwd, dbl_Strike, dbl_ActiveTimeToMat, _
                    dbl_ActiveCapletVol, dbl_ShiftedCapletVol) / 100 * dbl_ActiveDF * dbl_ActiveTenor * fld_Params.Notional * dblLst_AmortFactors(int_ctr)
            End If
        Next int_ctr
    End If

    Calc_BSOptionValueDigitalSmileOn = dbl_Output
End Function


Public Function CustomCashValuation(lng_CustomValDate As Long, irc_CustomCurve As Data_IRCurve) As Double
    ' ## Discount cash using specified curve and valuation date
    Dim int_ctr As Integer
    Dim int_NumPmts As Integer: int_NumPmts = lngLst_PmtDates.count
    Dim dblLst_CustomDFs As Collection: Set dblLst_CustomDFs = New Collection

    For int_ctr = 1 To int_NumPmts
        If intLst_IsCash(int_ctr) = 1 Then
            Call dblLst_CustomDFs.Add(irc_CustomCurve.Lookup_Rate(lng_CustomValDate, lngLst_PmtDates(int_ctr), "DF", , , False))
        Else
            Call dblLst_CustomDFs.Add(dblLst_DFs(int_ctr))
        End If
    Next int_ctr

    CustomCashValuation = CalcValue("CASH", dblLst_CustomDFs)
End Function

' ## METHODS - PRIVATE
Private Sub FillPeriodDates(int_NumPeriods As Integer)
    ' ## Generate start and end dates for cash flow calculation
    Dim lngLst_Follower As New Collection, lngLst_Driver As New Collection
    Dim int_Sign As Integer: If fld_Params.IsFwdGeneration = True Then int_Sign = 1 Else int_Sign = -1

    ' Perform generation
    Dim int_ctr As Integer
    For int_ctr = 1 To int_NumPeriods
        If int_ctr = 1 Then
            Call lngLst_Follower.Add(Date_ApplyBDC(fld_Params.GenerationRefPoint, "FOLL", cal_pmt.HolDates, cal_pmt.Weekends))
        Else
            Call lngLst_Follower.Add(lngLst_Driver(int_ctr - 1))
        End If

        Call lngLst_Driver.Add(Date_NextCoupon(fld_Params.GenerationRefPoint, str_PeriodLength, cal_pmt, _
            int_ctr * int_Sign, fld_Params.EOM, fld_Params.BDC))
    Next int_ctr

    ' Generation limit point if defined, will be the last generated date
    If fld_Params.GenerationLimitPoint <> 0 Then
        Call lngLst_Driver.Remove(int_NumPeriods)
        Call lngLst_Driver.Add(fld_Params.GenerationLimitPoint)
    End If

    ' Store period start and end dates depending on the generation method
    If fld_Params.IsFwdGeneration = True Then
        Set lngLst_PeriodStart = lngLst_Follower
        Set lngLst_PeriodEnd = lngLst_Driver
    Else
        Set lngLst_PeriodStart = Convert_Reverse_List(lngLst_Driver)
        Set lngLst_PeriodEnd = Convert_Reverse_List(lngLst_Follower)
    End If

    ' Store general start and end dates
    lng_startdate = lngLst_PeriodStart(1)
    lng_EndDate = lngLst_PeriodEnd(int_NumPeriods)
End Sub


Private Sub FillEstDates()
    ' ## Generate start and end dates for rate estimation
    Set lngLst_EstStart = New Collection
    Set lngLst_EstEnd = New Collection

    Dim int_ctr As Integer

    If fld_Params.FixInArrears = False Then
        For int_ctr = 1 To lngLst_PeriodStart.count
            Call lngLst_EstStart.Add(date_workday(lngLst_PeriodStart(int_ctr) - 1, 1, cal_est.HolDates, cal_est.Weekends))

            If fld_Params.FloatEst = True Then
                Call lngLst_EstEnd.Add(Date_NextCoupon(lngLst_EstStart(int_ctr), str_PeriodLength, cal_est, 1, _
                fld_Params.EOM, "MOD FOLL"))
            Else
                Call lngLst_EstEnd.Add(lngLst_PeriodEnd(int_ctr))
            End If
        Next int_ctr
    Else
        For int_ctr = 1 To lngLst_PeriodEnd.count
            Call lngLst_EstStart.Add(date_workday(lngLst_PeriodEnd(int_ctr) - 1, 1, cal_pmt.HolDates, cal_pmt.Weekends))

            If fld_Params.FloatEst = True Then
                Call lngLst_EstEnd.Add(date_workday(Date_NextCoupon(lngLst_EstStart(int_ctr), str_PeriodLength, cal_pmt, 1, _
                fld_Params.EOM, "MOD FOLL") - 1, 1, cal_pmt.HolDates, cal_pmt.Weekends))

            Else
                If int_ctr < lngLst_PeriodEnd.count Then
                    Call lngLst_EstEnd.Add(lngLst_PeriodStart(int_ctr + 1))
                Else
                    Call lngLst_EstEnd.Add(date_workday(Date_NextCoupon(lngLst_EstStart(int_ctr), str_PeriodLength, cal_pmt, 1, _
                    fld_Params.EOM, "MOD FOLL") - 1, 1, cal_pmt.HolDates, cal_pmt.Weekends))
                End If
            End If
        Next int_ctr
    End If
End Sub

Private Sub FillPmtDates()
    ' ## Return the list of cash flow payment dates
    ' ## Also classifies each payment date as belonging to MV or not

    Dim int_NumPeriods As Integer: int_NumPeriods = lngLst_PeriodEnd.count
    Set lngLst_PmtDates = New Collection

    Dim int_ctr As Integer
    For int_ctr = 1 To int_NumPeriods
        If int_ctr Mod int_CalcsPerPmt = 0 Then
            Call lngLst_PmtDates.Add(date_workday(lngLst_PeriodEnd(int_ctr) - 1, 1, cal_pmt.HolDates, cal_pmt.Weekends))
        End If
    Next int_ctr
End Sub

Private Sub CategorizeFlows()
   ' ## Classify each payment as cash or MV
    Set intLst_IsMV = New Collection
    Set intLst_IsCash = New Collection

    Dim int_ctr As Integer
    For int_ctr = 1 To lngLst_PmtDates.count
        If lngLst_PmtDates(int_ctr) > fld_Params.ValueDate Then
            Call intLst_IsMV.Add(1)
            Call intLst_IsCash.Add(0)
        ElseIf fld_Params.ForceToMV = True Then
            Call intLst_IsMV.Add(1)
            Call intLst_IsCash.Add(0)
        Else
            Call intLst_IsMV.Add(0)
            Call intLst_IsCash.Add(1)
        End If
    Next int_ctr

    ' Classify initial principal exchange (only relevant if setting is turned on)
    If lng_ValDate < fld_Params.Swapstart Or fld_Params.ForceToMV = True Then
        str_PrnAtStartPnlType = "MV"
    Else
        str_PrnAtStartPnlType = "CASH"
    End If
End Sub

Private Sub FillFlowDurations()
    ' ## Store length of period between value date and payment date of each flow
    Dim dbl_ActiveDuration As Double
    Set dblLst_FlowDurations = New Collection

    Dim int_ctr As Integer
    For int_ctr = 1 To lngLst_PmtDates.count
        ' Store duration of each flow
        dbl_ActiveDuration = calc_yearfrac(fld_Params.ValueDate, lngLst_PmtDates(int_ctr), str_Daycount_Duration, fld_Params.PmtFreq, bln_ActActFwdGeneration)
        If dbl_ActiveDuration < 0 Then Call dblLst_FlowDurations.Add(0) Else Call dblLst_FlowDurations.Add(dbl_ActiveDuration)
    Next int_ctr
End Sub

Private Sub ModifyStartDates()
    ' ## Either modify or remove the start date based on the list of changes.  If start date is removed, then corresponding end date is also removed

    If Not fld_Params.ModStarts Is Nothing Then
        Dim var_ActiveOrig As Variant, str_ActiveMod As String
        Dim lng_ActiveOrig As Long, lng_ActiveMod As Long
        Dim int_ctr As Integer, int_ctr2 As Integer

        'Determine stub upfront or stub arrears
        bln_StubUpfront = False
        bln_StubArrears = False

        ' Go through list of dates to modify
        For Each var_ActiveOrig In fld_Params.ModStarts.Keys
            lng_ActiveOrig = DateValue(var_ActiveOrig)
            str_ActiveMod = fld_Params.ModStarts(var_ActiveOrig)

            ' Search for date to modify
            For int_ctr = 1 To lngLst_PeriodStart.count
                If lngLst_PeriodStart(int_ctr) = lng_ActiveOrig Then
                    If str_ActiveMod = str_NotUsed Then
                        ' Remove the point containing the specified start date
                        Call lngLst_PeriodStart.Remove(int_ctr)
                        Call lngLst_PeriodEnd.Remove(int_ctr)

                        If Not lngLst_EstStart Is Nothing Then
                            Call lngLst_EstStart.Remove(int_ctr)
                            Call lngLst_EstEnd.Remove(int_ctr)
                        End If
                    Else
                        ' Modify the date and handle fixing in arrears case
                        lng_ActiveMod = DateValue(str_ActiveMod)
                        Call lngLst_PeriodStart.Remove(int_ctr)
                        Call lngLst_PeriodStart.Add(lng_ActiveMod, , int_ctr)

                        If Not lngLst_EstStart Is Nothing Then
                            If fld_Params.FixInArrears = False Then
                                Call lngLst_EstStart.Remove(int_ctr)
                                Call lngLst_EstStart.Add(lng_ActiveMod, , int_ctr)
                            ElseIf fld_Params.FixInArrears = True Then
                                If int_ctr > 1 Then
                                    Call lngLst_EstStart.Remove(int_ctr - 1)
                                    Call lngLst_EstStart.Add(lngLst_PeriodStart(int_ctr - 1), , int_ctr - 1)
                                End If
                            End If

                            If int_ctr = 1 Then
                                bln_StubUpfront = True
                            End If
                        End If
                    End If

                    Exit For
                End If
            Next int_ctr

            If lngLst_PeriodEnd(lngLst_PeriodStart.count) = lng_ActiveOrig Then
                lng_ActiveMod = DateValue(str_ActiveMod)
                Call lngLst_PeriodEnd.Remove(lngLst_PeriodStart.count)
                Call lngLst_PeriodEnd.Add(lng_ActiveMod)

                If Not lngLst_EstStart Is Nothing And fld_Params.FixInArrears = True Then
                    Call lngLst_EstStart.Remove(lngLst_PeriodStart.count)
                    Call lngLst_EstStart.Add(lng_ActiveMod)

                    If int_ctr = lngLst_PeriodStart.count Or lngLst_PeriodStart.count = 1 Then
                        bln_StubArrears = True
                    End If
                End If
            End If
        Next var_ActiveOrig
    End If
End Sub

Private Sub FillPeriodDurations()
    ' ## Store length of each period
    Dim int_NumPeriods As Integer: int_NumPeriods = lngLst_PeriodEnd.count
    Set dblLst_CalcPeriodDurations = New Collection
    Set dblLst_EstPeriodDurations = New Collection

    Dim dbl_UniformPeriod As Double
    If fld_Params.IsUniformPeriods = True Then dbl_UniformPeriod = calc_nummonths(fld_Params.PmtFreq) / 12

    Dim int_ctr As Integer
    For int_ctr = 1 To int_NumPeriods
        If fld_Params.IsUniformPeriods = True Then
            Call dblLst_CalcPeriodDurations.Add(dbl_UniformPeriod)
        Else
            Call dblLst_CalcPeriodDurations.Add(calc_yearfrac(lngLst_PeriodStart(int_ctr), lngLst_PeriodEnd(int_ctr), _
                fld_Params.Daycount, fld_Params.PmtFreq, bln_ActActFwdGeneration))
        End If

        If fld_Params.index <> str_NotUsed Then
            Call dblLst_EstPeriodDurations.Add(calc_yearfrac(lngLst_EstStart(int_ctr), lngLst_EstEnd(int_ctr), "ACT/365"))
        End If
    Next int_ctr
End Sub

Private Sub FillAmortFactors()
    ' ## Store amortization factors if defined on sheet, otherwise store values of 1
    ' ## Factors for all payment dates must be stored on the sheet

    Dim int_NumPeriods As Integer: int_NumPeriods = lngLst_PeriodStart.count
    Dim lng_ActivePeriodStart As Long
    Set dblLst_AmortFactors = New Collection
    Dim int_ctr As Integer

    If fld_Params.AmortSchedule Is Nothing Then
        For int_ctr = 1 To int_NumPeriods
            Call dblLst_AmortFactors.Add(1)
        Next int_ctr
    Else
        For int_ctr = 1 To int_NumPeriods
            lng_ActivePeriodStart = lngLst_PeriodStart(int_ctr)
            If fld_Params.AmortSchedule.Exists(lng_ActivePeriodStart) = True Then
                Call dblLst_AmortFactors.Add(fld_Params.AmortSchedule(lng_ActivePeriodStart))
            Else
                Call dblLst_AmortFactors.Add(0)
                Debug.Print "## ERROR - Amort factor lookup failed.  Trade: " & fld_Params.TradeID & "  Date: " _
                    & Format(lng_ActivePeriodStart, Gather_DateFormat())
            End If
        Next int_ctr
    End If
End Sub

Private Sub FillPrincipalFlows()
    ' ## Derive principal flows from amortization schedule if intermediate principal flows are enabled
    ' ## Include final exchange of notional if end principal flows are enabled
    ' ## If multiple periods per payment, assumes principal flow accrues over periods and is returned on the payment date

    Dim int_NumPmts As Integer: int_NumPmts = lngLst_PmtDates.count
    Dim int_ctr As Integer, int_PeriodNumAtPmtStart As Integer

    ' Include principal exchange at start if setting is on
    If fld_Params.PExch_Start = True Then dbl_PrnAtStart = -fld_Params.Notional Else dbl_PrnAtStart = 0

    ' Assumes that principal flows occur on payment dates
    Set dblLst_PrnFlows = New Collection

    For int_ctr = 1 To int_NumPmts - 1
        If fld_Params.PExch_Intermediate = True Then
            int_PeriodNumAtPmtStart = (int_ctr - 1) * int_CalcsPerPmt + 1
            Call dblLst_PrnFlows.Add(fld_Params.Notional * (dblLst_AmortFactors(int_PeriodNumAtPmtStart) - _
                dblLst_AmortFactors(int_PeriodNumAtPmtStart + int_CalcsPerPmt)))
        Else
            Call dblLst_PrnFlows.Add(0)
        End If
    Next int_ctr

    If fld_Params.PExch_End = True Then
    int_PeriodNumAtPmtStart = (int_ctr - 1) * int_CalcsPerPmt + 1
        Call dblLst_PrnFlows.Add(fld_Params.Notional * dblLst_AmortFactors(int_PeriodNumAtPmtStart))
    Else
        Call dblLst_PrnFlows.Add(0)
    End If
End Sub

Private Sub FillRates()
    ' ## Store rates for each estimation period
    ' ## Used for floating legs only
    Dim lng_BuildDate As Long, lng_ActiveEstStart As Long
    Dim dbl_ActiveRate As Double
    Dim int_ctr As Integer

    ' At each estimation period, store the rate
    Set dblLst_Rates = New Collection
    Set dblLst_Margins = New Collection

    If bln_IsFixed = True Then
        For int_ctr = 1 To lngLst_PeriodEnd.count
            Call dblLst_Rates.Add(fld_Params.RateOrMargin)
            Call dblLst_Margins.Add(0)
        Next int_ctr
    Else
        lng_BuildDate = irc_Est.BuildDate

        For int_ctr = 1 To lngLst_PeriodEnd.count
            lng_ActiveEstStart = lngLst_EstStart(int_ctr)

            ' Estimate floating rates for future fixings, past fixin0gs will be filled in later
            If lng_ActiveEstStart >= lng_BuildDate And lng_ActiveEstStart >= fld_Params.ValueDate Then
                'dbl_ActiveRate = irc_Est.Lookup_Rate(lng_ActiveEstStart, lngLst_EstEnd(int_Ctr), fld_Params.Daycount, _
                    , fld_Params.index, False, , , bln_ActActFwdGeneration)

                dbl_ActiveRate = CalcFwdRate(lng_ActiveEstStart, lngLst_EstEnd(int_ctr))

                If ((int_ctr = 1 And bln_StubUpfront = True) Or (int_ctr = lngLst_PeriodEnd.count And bln_StubArrears = True) And _
                    bln_StubInterpolate = True) Then

                    dbl_ActiveRate = CalcStubInterpFwdRate(lngLst_PeriodEnd(int_ctr) - lngLst_PeriodStart(int_ctr), lng_ActiveEstStart)

                ElseIf int_ctr = 1 And bln_StubInterpolate = True Then
                    dbl_ActiveRate = CalcStubInterpFwdRate(lngLst_PeriodEnd(int_ctr) - lngLst_PeriodStart(int_ctr), lng_ActiveEstStart, False)

                End If
            Else
                dbl_ActiveRate = 0
            End If

            Call dblLst_Rates.Add(dbl_ActiveRate)
            Call dblLst_Margins.Add(fld_Params.RateOrMargin)
        Next int_ctr
    End If
End Sub

Private Sub FillIntFlows()
    ' ## Calculate flow amounts from the already set rates based on the specification of the swap

    Dim int_NumPeriods As Integer: int_NumPeriods = lngLst_PeriodStart.count
    Dim dbl_ActiveCalcFlow As Double, dbl_ActiveAccum As Double
    Dim int_CalcsToPmt As Integer: int_CalcsToPmt = int_CalcsPerPmt
    Dim int_ctr As Integer

    Set dblLst_IntFlows = New Collection
    For int_ctr = 1 To int_NumPeriods

        dbl_ActiveCalcFlow = fld_Params.Notional * (dblLst_Rates(int_ctr) + dblLst_Margins(int_ctr)) / 100 _
            * dblLst_CalcPeriodDurations(int_ctr) * dblLst_AmortFactors(int_ctr)
        dbl_ActiveAccum = dbl_ActiveAccum * (1 + dblLst_Rates(int_ctr) / 100 * dblLst_CalcPeriodDurations(int_ctr)) + dbl_ActiveCalcFlow

        ' Accumulate unpaid calculation flows until end of payment period using simple interest per calculation period
        If int_CalcsToPmt = 1 Then
            ' Pay accumulated amount at end of payment period
            Call dblLst_IntFlows.Add(dbl_ActiveAccum)

            ' Reset for next payment period
            dbl_ActiveAccum = 0
            int_CalcsToPmt = int_CalcsPerPmt
        Else
            int_CalcsToPmt = int_CalcsToPmt - 1
        End If
    Next int_ctr
End Sub

Private Sub FillIntFlows_Rate()
    ' ## Calculate flow amounts from the already set rates based on the specification of the swap

    Dim int_NumPeriods As Integer: int_NumPeriods = lngLst_PeriodStart.count
    Dim dbl_ActiveCalcFlow As Double, dbl_ActiveAccum As Double
    Dim int_CalcsToPmt As Integer: int_CalcsToPmt = int_CalcsPerPmt
    Dim int_ctr As Integer

    Set dblLst_IntFlows_Rate = New Collection
    For int_ctr = 1 To int_NumPeriods

        dbl_ActiveCalcFlow = fld_Params.Notional * dblLst_Rates(int_ctr) / 100 _
            * dblLst_CalcPeriodDurations(int_ctr) * dblLst_AmortFactors(int_ctr)
        dbl_ActiveAccum = dbl_ActiveAccum * (1 + dblLst_Rates(int_ctr) / 100 * dblLst_CalcPeriodDurations(int_ctr)) + dbl_ActiveCalcFlow

        ' Accumulate unpaid calculation flows until end of payment period using simple interest per calculation period
        If int_CalcsToPmt = 1 Then
            ' Pay accumulated amount at end of payment period
            Call dblLst_IntFlows_Rate.Add(dbl_ActiveAccum)

            ' Reset for next payment period
            dbl_ActiveAccum = 0
            int_CalcsToPmt = int_CalcsPerPmt
        Else
            int_CalcsToPmt = int_CalcsToPmt - 1
        End If
    Next int_ctr
End Sub
Private Sub FillIntFlows_Margin()
    ' ## Calculate flow amounts from the already set rates based on the specification of the swap

    Dim int_NumPeriods As Integer: int_NumPeriods = lngLst_PeriodStart.count
    Dim dbl_ActiveCalcFlow As Double, dbl_ActiveAccum As Double
    Dim int_CalcsToPmt As Integer: int_CalcsToPmt = int_CalcsPerPmt
    Dim int_ctr As Integer

    Set dblLst_IntFlows_Margin = New Collection
    For int_ctr = 1 To int_NumPeriods

        dbl_ActiveCalcFlow = fld_Params.Notional * dblLst_Margins(int_ctr) / 100 _
            * dblLst_CalcPeriodDurations(int_ctr) * dblLst_AmortFactors(int_ctr)
        dbl_ActiveAccum = dbl_ActiveAccum * (1 + dblLst_Rates(int_ctr) / 100 * dblLst_CalcPeriodDurations(int_ctr)) + dbl_ActiveCalcFlow

        ' Accumulate unpaid calculation flows until end of payment period using simple interest per calculation period
        If int_CalcsToPmt = 1 Then
            ' Pay accumulated amount at end of payment period
            Call dblLst_IntFlows_Margin.Add(dbl_ActiveAccum)

            ' Reset for next payment period
            dbl_ActiveAccum = 0
            int_CalcsToPmt = int_CalcsPerPmt
        Else
            int_CalcsToPmt = int_CalcsToPmt - 1
        End If
    Next int_ctr
End Sub

Private Sub FillFixings()
    ' ## Read from custom sheet.  Relevant only for floating legs

    If Not fld_Params.Fixings Is Nothing Then
        Dim var_ActiveDate As Variant, dbl_ActiveFixing As Double
        Dim int_NumPeriods As Integer: int_NumPeriods = lngLst_PeriodStart.count
        Dim int_ctr As Integer

        ' Go through list of dates to modify
        For Each var_ActiveDate In fld_Params.Fixings.Keys
            dbl_ActiveFixing = fld_Params.Fixings(var_ActiveDate)

            ' Search for affected date to modify
            For int_ctr = 1 To int_NumPeriods
                If lngLst_PeriodStart(int_ctr) = CLng(var_ActiveDate) Then
                    ' Replace stored rate with fixing
                    Call dblLst_Rates.Remove(int_ctr)

                    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                     'Alvin Edit 28/09/2018
                     ' This code allows mod on Last CF
                    If dblLst_Rates.count < int_ctr Then
                        Call dblLst_Rates.Add(dbl_ActiveFixing)
                    Else
                        Call dblLst_Rates.Add(dbl_ActiveFixing, , int_ctr)
                    End If
                    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

                    'Call dblLst_Rates.Add(dbl_ActiveFixing, , int_ctr)  #ORI Alvin Edit 28/09/2018


                    ' Period is no longer relevant for estimation curve DV01
                    Call dblLst_EstPeriodDurations.Remove(int_ctr)

                    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                     'Alvin Edit 28/09/2018
                     ' This code allows mod on Last CF
                    If dblLst_EstPeriodDurations.count < int_ctr Then
                        Call dblLst_EstPeriodDurations.Add(0)
                    Else
                        Call dblLst_EstPeriodDurations.Add(0, , int_ctr)
                    End If
                    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

                    'Call dblLst_EstPeriodDurations.Add(0, , int_ctr) #ORI Alvin Edit 28/09/2018

                    Exit For
                ElseIf lngLst_PeriodStart(int_ctr) > CLng(var_ActiveDate) Then
                    ' Already past fixing date, all subsequent periods will also be beyond the fixing date
                    Exit For
                End If
            Next int_ctr
        Next var_ActiveDate
    End If
End Sub

Private Sub RecalcIntFlows()
    ' ## Call this after changing the rate
    Call FillRates
    If bln_IsFixed = False Then Call FillFixings
    Call FillIntFlows
End Sub

Private Sub FillDFs()
    ' ## Store discount factors by reading from curve
    Set dblLst_DFs = irc_Disc.Lookup_DFs(lng_ValDate, lngLst_PmtDates, False, dbl_ZSpread)

    If lng_ValDate >= fld_Params.Swapstart Then
        dbl_StartDF = 1
    Else
        dbl_StartDF = irc_Disc.Lookup_Rate(lng_ValDate, fld_Params.Swapstart, "DF", , , False, , dbl_ZSpread)
    End If
End Sub

Public Function CalcValue(str_type As String, Optional dblLst_CustomDFs As Variant, Optional dbl_CustomStartDF As Double = -1) As Double
    ' ## Calculate either the MV or cash for the swap
    Dim dbl_Output As Double
    Dim intLst_Inclusions As Collection
    Select Case str_type
        Case "MV"
            Set intLst_Inclusions = intLst_IsMV
        Case "CASH"
            Set intLst_Inclusions = intLst_IsCash
        Case Else: Debug.Assert False
    End Select

    ' Use custom DFs if specified
    Dim dblLst_DFs_ToUse As Collection, dbl_StartDF_ToUse As Double
    If IsMissing(dblLst_CustomDFs) Then Set dblLst_DFs_ToUse = dblLst_DFs Else Set dblLst_DFs_ToUse = dblLst_CustomDFs
    If dbl_CustomStartDF = -1 Then dbl_StartDF_ToUse = dbl_StartDF Else dbl_StartDF_ToUse = dbl_CustomStartDF

    ' Calculate value based on specified PnL type
    dbl_Output = Calc_SumProductOnList(dblLst_IntFlows, dblLst_DFs_ToUse, intLst_Inclusions) _
        + Calc_SumProductOnList(dblLst_PrnFlows, dblLst_DFs_ToUse, intLst_Inclusions)

    ' Include principal exchange at start if exists
    If str_type = str_PrnAtStartPnlType Then dbl_Output = dbl_Output + dbl_PrnAtStart * dbl_StartDF_ToUse

    CalcValue = dbl_Output
End Function

  ' ************** Kuan Lim Added*****************
Public Function CalcValue_RA(str_type As String, dblLst_digi As Collection, dblLst_digi_corr As Collection, _
                        Optional dblLst_CustomDFs As Variant, Optional dbl_CustomStartDF As Double = -1) As Double

    ' ## Calculate either the MV or cash for the swap
    Dim dbl_Output As Double
    Dim intLst_Inclusions As Collection
    Select Case str_type
        Case "MV"
            Set intLst_Inclusions = intLst_IsMV
        Case "CASH"
            Set intLst_Inclusions = intLst_IsCash
        Case Else: Debug.Assert False
    End Select

    ' Use custom DFs if specified
    Dim dblLst_DFs_ToUse As Collection, dbl_StartDF_ToUse As Double
    If IsMissing(dblLst_CustomDFs) Then Set dblLst_DFs_ToUse = dblLst_DFs Else Set dblLst_DFs_ToUse = dblLst_CustomDFs
    If dbl_CustomStartDF = -1 Then dbl_StartDF_ToUse = dbl_StartDF Else dbl_StartDF_ToUse = dbl_CustomStartDF

    ' Calculate value based on specified PnL type
    Dim aa As Variant, bb As Variant, cc As Variant

    aa = Calc_SumProductOnList(dblLst_IntFlows_Rate, dblLst_DFs_ToUse, intLst_Inclusions, dblLst_digi_corr)
    bb = Calc_SumProductOnList(dblLst_IntFlows_Margin, dblLst_DFs_ToUse, intLst_Inclusions, dblLst_digi)
    cc = Calc_SumProductOnList(dblLst_PrnFlows, dblLst_DFs_ToUse, intLst_Inclusions)

    dbl_Output = aa + bb + cc

    ' Include principal exchange at start if exists
    If str_type = str_PrnAtStartPnlType Then dbl_Output = dbl_Output + dbl_PrnAtStart * dbl_StartDF_ToUse

    CalcValue_RA = dbl_Output
End Function
Private Function CalcConvAdj(dbl_Tal As Double, dbl_CapFac As Double, dbl_FwdRate As Double, dbl_Vol As Double, dbl_T As Double) As Double

'Calculate Convexity Adjustment

Dim dbl_Pdy As Double
Dim dbl_Pdy2 As Double

dbl_Pdy = -dbl_Tal / 100 / (dbl_CapFac ^ 2)
dbl_Pdy2 = 2 * ((dbl_Tal / 100) ^ 2) / (dbl_CapFac ^ 3)
CalcConvAdj = -0.5 * (dbl_FwdRate ^ 2) * (dbl_Vol ^ 2) * dbl_T * dbl_Pdy2 / dbl_Pdy

End Function
Public Function GetRates(int_ctr As Integer) As Double
    GetRates = dblLst_Rates(int_ctr)
End Function
Private Function CalcFwdRate(lng_EstStart As Long, lngLst_EstEnd As Long) As Double

'Calculate Forward Rate

Dim dbl_rate As Double
Dim lng_FixDate As Long
Dim dbl_ActiveVol As Double
Dim dbl_ActiveTimeToEstStart As Double
Dim dbl_ActiveTal As Double
Dim dbl_ActiveCapFac As Double
Dim dbl_ConvAdj As Double


dbl_rate = irc_Est.Lookup_Rate(lng_EstStart, lngLst_EstEnd, fld_Params.Daycount, _
    , fld_Params.index, False, , , bln_ActActFwdGeneration)

'Convexity adjustment
If bln_IsFixInArrears = True And bln_IsDisableConvAdj = False Then
    lng_FixDate = date_workday(lng_EstStart, cvl_volcurve.Deduction, cal_est.HolDates, cal_est.Weekends)
    dbl_ActiveVol = cyGetCapVolSurf(fld_Params.CCY & "_" & fld_Params.index, lng_EstStart, dbl_rate) / 100
    dbl_ActiveTimeToEstStart = (lng_FixDate - lng_ValDate) / 365
    dbl_ActiveTal = (lngLst_EstEnd - lng_EstStart) / 365
    dbl_ActiveCapFac = 1 + dbl_rate / 100 * calc_yearfrac(lng_EstStart, lngLst_EstEnd, _
        fld_Params.Daycount, fld_Params.PmtFreq, bln_ActActFwdGeneration)
    dbl_ConvAdj = CalcConvAdj(dbl_ActiveTal, dbl_ActiveCapFac, dbl_rate, dbl_ActiveVol, dbl_ActiveTimeToEstStart)
    dbl_rate = dbl_rate + dbl_ConvAdj
End If

CalcFwdRate = dbl_rate

End Function

Private Function CalcStubInterpFwdRate(int_numdays As Integer, lng_EstStartDate As Long, Optional bln_rounding = True) As Double

'Calculate interpolated forward rate for stub period
Dim lng_FirstEndDate As Long, lng_SecondEndDate As Long
Dim int_FirstDays As Integer, int_SecondDays As Integer, int_CtrLoop As Integer
Dim dbl_FirstRate As Double, dbl_SecondRate As Double, dbl_ActiveRate As Double


    lng_FirstEndDate = date_workday(lng_EstStartDate + 7 - 1, 1, cal_pmt.HolDates, cal_pmt.Weekends)

        If int_numdays <= 7 Then
            dbl_ActiveRate = CalcFwdRate(lng_EstStartDate, lng_FirstEndDate)
        Else
            For int_CtrLoop = 1 To 6
                lng_SecondEndDate = Date_NextCoupon(lng_EstStartDate, int_CtrLoop & "M", cal_pmt, _
                                    1, fld_Params.EOM, "MOD FOLL")

                lng_SecondEndDate = date_workday(lng_SecondEndDate - 1, 1, cal_pmt.HolDates, cal_pmt.Weekends)

                int_SecondDays = lng_SecondEndDate - lng_EstStartDate

                If int_SecondDays >= int_numdays Then
                    int_FirstDays = lng_FirstEndDate - lng_EstStartDate
                    dbl_FirstRate = CalcFwdRate(lng_EstStartDate, lng_FirstEndDate)
                    dbl_SecondRate = CalcFwdRate(lng_EstStartDate, lng_SecondEndDate)
                    dbl_ActiveRate = dbl_FirstRate + (dbl_SecondRate - dbl_FirstRate) / (int_SecondDays - int_FirstDays) * (int_numdays - int_FirstDays)

                    If bln_rounding = True Then
                    dbl_ActiveRate = WorksheetFunction.Round(dbl_ActiveRate, 5)
                    End If

                    Exit For
                Else
                    lng_FirstEndDate = lng_SecondEndDate
                End If
            Next int_CtrLoop
        End If

CalcStubInterpFwdRate = dbl_ActiveRate

End Function

' ## METHODS - CALCULATION DETAILS
Public Sub OutputReport_Swap(rng_OutputTopLeft As Range, str_LegName As String, bln_PayLeg As Boolean, str_Address_PnLCCY As String, _
    str_Address_ValDate As String, str_CCY_PnL As String, ByRef str_Address_MV As String, ByRef str_Address_Cash As String)
    ' ## Display rates, flows and discount factors
    Dim str_CurrencyFormat As String: str_CurrencyFormat = Gather_CurrencyFormat()
    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    Dim rng_MV As Range, rng_Cash As Range, rng_NativeCCY As Range
    Dim int_ActiveRow As Integer: int_ActiveRow = 0
    Dim int_ActiveCol As Integer: int_ActiveCol = 0
    Dim str_Address_PmtDate As String
    Dim int_ctr As Integer

    ' General info
    With rng_OutputTopLeft
        .Offset(int_ActiveRow, int_ActiveCol).Value = str_LegName

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "MV (" & str_CCY_PnL & "):"
        Set rng_MV = .Offset(int_ActiveRow, int_ActiveCol + 1)
        str_Address_MV = rng_MV.Address(False, False)

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Cash (" & str_CCY_PnL & "):"
        Set rng_Cash = .Offset(int_ActiveRow, int_ActiveCol + 1)
        str_Address_Cash = rng_Cash.Address(False, False)

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Currency:"
        .Offset(int_ActiveRow, int_ActiveCol + 1).Value = fld_Params.CCY
        Set rng_NativeCCY = .Offset(int_ActiveRow, int_ActiveCol + 1)

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Direction:"
        If bln_PayLeg = True Then .Offset(int_ActiveRow, int_ActiveCol + 1).Value = "Pay" Else .Offset(int_ActiveRow, int_ActiveCol + 1).Value = "Receive"

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Leg type:"
        If bln_IsFixed = True Then
            .Offset(int_ActiveRow, int_ActiveCol + 1).Value = "Fixed"
        Else
            .Offset(int_ActiveRow, int_ActiveCol + 1).Value = "Floating"
        End If

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Daycount:"
        .Offset(int_ActiveRow, int_ActiveCol + 1).Value = fld_Params.Daycount
    End With

    ' Rate headings
    int_ActiveRow = int_ActiveRow + 2
    With rng_OutputTopLeft
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Period Start"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Period End"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Est Start"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Est End"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Notional"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Rate"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Margin"

        int_ActiveCol = int_ActiveCol - 6
        If fld_Params.PExch_Start = True Then int_ActiveRow = int_ActiveRow + 1

        ' Rates
        For int_ctr = 1 To lngLst_PeriodEnd.count
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = lngLst_PeriodStart(int_ctr)

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = lngLst_PeriodEnd(int_ctr)

            int_ActiveCol = int_ActiveCol + 3
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = fld_Params.Notional * dblLst_AmortFactors(int_ctr)
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).NumberFormat = str_CurrencyFormat

            If bln_IsFixed = True Then
                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = fld_Params.RateOrMargin

                int_ActiveCol = int_ActiveCol - 5
            Else
                int_ActiveCol = int_ActiveCol - 2
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = lngLst_EstStart(int_ctr)

                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = lngLst_EstEnd(int_ctr)

                int_ActiveCol = int_ActiveCol + 2
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = dblLst_Rates(int_ctr)

                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = dblLst_Margins(int_ctr)

                int_ActiveCol = int_ActiveCol - 6
            End If

            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Resize(1, 4).NumberFormat = str_DateFormat
        Next int_ctr
    End With

    ' Payment headings
    If fld_Params.PExch_Start = True Then int_ActiveRow = int_ActiveRow - 1
    int_ActiveCol = int_ActiveCol + 8
    With rng_OutputTopLeft
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Pmt Date"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Type"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Int Flow"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Prn Flow"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "DF"

        int_ActiveCol = int_ActiveCol - 4
    End With

    ' Optional principal exchange at start
    If fld_Params.PExch_Start = True Then
        int_ActiveRow = int_ActiveRow + 1
        With rng_OutputTopLeft
            .Offset(int_ActiveRow, int_ActiveCol).Value = fld_Params.Swapstart
            .Offset(int_ActiveRow, int_ActiveCol).NumberFormat = str_DateFormat

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ActiveRow, int_ActiveCol).Value = str_PrnAtStartPnlType

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ActiveRow, int_ActiveCol).Value = 0
            .Offset(int_ActiveRow, int_ActiveCol).NumberFormat = str_CurrencyFormat

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ActiveRow, int_ActiveCol).Value = dbl_PrnAtStart
            .Offset(int_ActiveRow, int_ActiveCol).NumberFormat = str_CurrencyFormat

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ActiveRow, int_ActiveCol).Value = dbl_StartDF

            int_ActiveCol = int_ActiveCol - 4
        End With
    End If

    ' Payments
    Dim int_PmtSectionHeight As Integer: int_PmtSectionHeight = lngLst_PmtDates.count * int_CalcsPerPmt
    Dim int_ActiveIndex As Integer: int_ActiveIndex = 0
    With rng_OutputTopLeft
        For int_ctr = 1 To int_PmtSectionHeight
            If int_ctr Mod int_CalcsPerPmt = 0 Then
                int_ActiveIndex = int_ActiveIndex + 1
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = lngLst_PmtDates(int_ActiveIndex)
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).NumberFormat = str_DateFormat
                str_Address_PmtDate = .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Address(False, False)

                int_ActiveCol = int_ActiveCol + 1
                If intLst_IsMV(int_ActiveIndex) = 1 Then
                    .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = "MV"
                Else
                    .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = "CASH"
                End If

                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = dblLst_IntFlows(int_ActiveIndex)
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).NumberFormat = str_CurrencyFormat

                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = dblLst_PrnFlows(int_ActiveIndex)
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).NumberFormat = str_CurrencyFormat

                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Formula = "=cyReadIRCurve(""" & fld_Params.Curve_Disc & """," & str_Address_ValDate _
                    & "," & str_Address_PmtDate & ",""DF"",,False)"

                int_ActiveCol = int_ActiveCol - 4
            End If
        Next int_ctr
    End With

    ' Show formula for NPV
    Dim int_SumproductRowOffset As Integer: If fld_Params.PExch_Start = True Then int_SumproductRowOffset = 0 Else int_SumproductRowOffset = 1
    Dim rng_IntFlows As Range
    If fld_Params.PExch_Start = True Then
        Set rng_IntFlows = rng_OutputTopLeft.Offset(int_ActiveRow, int_ActiveCol + 2).Resize(int_PmtSectionHeight + 1, 1)
    Else
        Set rng_IntFlows = rng_OutputTopLeft.Offset(int_ActiveRow + 1, int_ActiveCol + 2).Resize(int_PmtSectionHeight, 1)
    End If

    Dim str_Sign As String: If bln_PayLeg = True Then str_Sign = "=-" Else str_Sign = "=+"
    rng_MV.Value = str_Sign & "SUMPRODUCT((" & rng_IntFlows.Address(False, False) & "+" & rng_IntFlows.Offset(0, 1).Address(False, False) & ")*(" _
        & rng_IntFlows.Offset(0, 2).Address(False, False) & ")*(" & rng_IntFlows.Offset(0, -1).Address(False, False) & "=""MV""))*cyGetFXDiscSpot(" _
        & rng_NativeCCY.Address(False, False) & "," & str_Address_PnLCCY & ")"
    rng_MV.NumberFormat = str_CurrencyFormat

    rng_Cash.Value = str_Sign & "SUMPRODUCT((" & rng_IntFlows.Address(False, False) & "+" & rng_IntFlows.Offset(0, 1).Address(False, False) & ")*(" _
    & rng_IntFlows.Offset(0, 2).Address(False, False) & ")*(" & rng_IntFlows.Offset(0, -1).Address(False, False) & "=""CASH""))*cyGetFXDiscSpot(" _
    & rng_NativeCCY.Address(False, False) & "," & str_Address_PnLCCY & ")"
    rng_Cash.NumberFormat = str_CurrencyFormat
End Sub


Public Sub OutputReport_Option(rng_OutputTopLeft As Range, str_LegName As String, dbl_Strike As Double, enu_Direction As OptionDirection, _
    dblLst_Vols As Collection, int_Deduction As Integer, cal_Deduction As Calendar, str_Position As String, _
    str_CCY_PnL As String, scf_Premium As SCF, int_Sign As Integer)
    ' ## Display rates, flows and discount factors
    Dim str_CurrencyFormat As String: str_CurrencyFormat = Gather_CurrencyFormat()
    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    Dim str_Address_Cash As String
    Dim rng_MV As Range
    Dim int_ctr As Integer
    Dim int_ActiveRow As Integer: int_ActiveRow = 0
    Dim int_ActiveCol As Integer: int_ActiveCol = 0

    Dim dic_Addresses As Dictionary: Set dic_Addresses = New Dictionary
    dic_Addresses.CompareMode = CompareMethod.TextCompare

    Dim dic_TempAddresses As Dictionary: Set dic_TempAddresses = New Dictionary
    dic_TempAddresses.CompareMode = CompareMethod.TextCompare
    ' General info
    With rng_OutputTopLeft
        .Offset(int_ActiveRow, int_ActiveCol).Value = "OVERALL"
        .Offset(int_ActiveRow, int_ActiveCol).Font.Italic = True

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Value Date:"
        .Offset(int_ActiveRow, int_ActiveCol + 1).Value = fld_Params.ValueDate
        .Offset(int_ActiveRow, int_ActiveCol + 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("ValDate", .Offset(int_ActiveRow, int_ActiveCol + 1).Address(True, True))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "PnL:"
        .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
        Call dic_Addresses.Add("Range_PnL", .Offset(int_ActiveRow, int_ActiveCol + 1))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "CCY:"
        .Offset(int_ActiveRow, int_ActiveCol + 1).Value = str_CCY_PnL
        Call dic_Addresses.Add("PnLCCY", .Offset(int_ActiveRow, int_ActiveCol + 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 2
        .Offset(int_ActiveRow, int_ActiveCol).Value = "MV (" & str_CCY_PnL & "):"
        .Offset(int_ActiveRow, int_ActiveCol + 1).NumberFormat = str_CurrencyFormat
        .Offset(int_ActiveRow, int_ActiveCol + 1).Interior.ColorIndex = 20
        Call dic_Addresses.Add("Range_TotalMV", .Offset(int_ActiveRow, int_ActiveCol + 1))
        Call dic_Addresses.Add("TotalMV", .Offset(int_ActiveRow, int_ActiveCol + 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Cash (" & str_CCY_PnL & "):"
        .Offset(int_ActiveRow, int_ActiveCol + 1).NumberFormat = str_CurrencyFormat
        .Offset(int_ActiveRow, int_ActiveCol + 1).Interior.ColorIndex = 20
        Call dic_Addresses.Add("Range_TotalCash", .Offset(int_ActiveRow, int_ActiveCol + 1))
        Call dic_Addresses.Add("TotalCash", .Offset(int_ActiveRow, int_ActiveCol + 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 3
        .Offset(int_ActiveRow, int_ActiveCol).Value = "OPTION LEG"
        .Offset(int_ActiveRow, int_ActiveCol).Font.Italic = True

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "CCY:"
        .Offset(int_ActiveRow, int_ActiveCol + 1).Value = fld_Params.CCY
        Call dic_Addresses.Add("OptionCCY", .Offset(int_ActiveRow, int_ActiveCol + 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Payout:"
        If enu_Direction = OptionDirection.CallOpt Then
            .Offset(int_ActiveRow, int_ActiveCol + 1).Value = "Call"
        Else
            .Offset(int_ActiveRow, int_ActiveCol + 1).Value = "Put"
        End If
        Call dic_Addresses.Add("Payout", .Offset(int_ActiveRow, int_ActiveCol + 1).Address(True, True))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Position:"
        .Offset(int_ActiveRow, int_ActiveCol + 1).Value = str_Position
        Call dic_Addresses.Add("Position", .Offset(int_ActiveRow, int_ActiveCol + 1).Address(False, False))
    End With

    int_ActiveRow = int_ActiveRow + 2
    With rng_OutputTopLeft
        ' Rate headings
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Period Start"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Period End"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Est Start"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Est End"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Notional"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Rate"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Strike"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Vol"

        int_ActiveCol = int_ActiveCol - 7

        ' Rates
        For int_ctr = 1 To lngLst_PeriodEnd.count
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = lngLst_PeriodStart(int_ctr)

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = lngLst_PeriodEnd(int_ctr)

            If bln_IsFixed = True Then
                int_ActiveCol = int_ActiveCol + 4
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = fld_Params.RateOrMargin
            Else
                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = lngLst_EstStart(int_ctr)

                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = lngLst_EstEnd(int_ctr)

                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = fld_Params.Notional * dblLst_AmortFactors(int_ctr)
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).NumberFormat = str_CurrencyFormat

                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = dblLst_Rates(int_ctr) + dblLst_Margins(int_ctr)
            End If

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = dbl_Strike

            int_ActiveCol = int_ActiveCol + 1
            If dblLst_Vols(int_ctr) <> 0 Then
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = dblLst_Vols(int_ctr)
            End If

            int_ActiveCol = int_ActiveCol - 7
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Resize(1, 4).NumberFormat = str_DateFormat
        Next int_ctr
    End With

    int_ActiveCol = int_ActiveCol + 9
    With rng_OutputTopLeft
        ' Payment headings
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Opt Mat"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Pmt Date"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Type"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Flow"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "DF"

        int_ActiveCol = int_ActiveCol - 4

        ' Payments
        Dim int_NumPmts As Integer: int_NumPmts = lngLst_PmtDates.count
        For int_ctr = 1 To int_NumPmts
            Call dic_TempAddresses.RemoveAll

            int_ActiveCol = int_ActiveCol - 9
            Call dic_TempAddresses.Add("CalcStart", .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Address(False, False))

            int_ActiveCol = int_ActiveCol + 1
            Call dic_TempAddresses.Add("CalcEnd", .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Address(False, False))

            int_ActiveCol = int_ActiveCol + 3
            Call dic_TempAddresses.Add("Notional", .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Address(False, False))

            int_ActiveCol = int_ActiveCol + 1
            Call dic_TempAddresses.Add("Rate", .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Address(False, False))

            int_ActiveCol = int_ActiveCol + 1
            Call dic_TempAddresses.Add("Strike", .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Address(False, False))

            int_ActiveCol = int_ActiveCol + 1
            Call dic_TempAddresses.Add("Vol", .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Address(False, False))

            int_ActiveCol = int_ActiveCol + 2
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = date_workday(lngLst_PeriodStart(int_ctr), int_Deduction, cal_Deduction.HolDates, cal_Deduction.Weekends)
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).NumberFormat = str_DateFormat
            Call dic_TempAddresses.Add("OptMat", .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Address(False, False))

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = lngLst_PeriodEnd(int_ctr)
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).NumberFormat = str_DateFormat
            Call dic_TempAddresses.Add("PmtDate", .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Address(False, False))

            int_ActiveCol = int_ActiveCol + 1
            If intLst_IsMV(int_ctr) = 1 Then
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = "MV"
            Else
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = "CASH"
            End If

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Formula = "=" & dic_TempAddresses("Notional") _
                & "*Calc_BSPrice_Vanilla(IF(" & dic_Addresses("Payout") & "=""Call"",1,-1)," & dic_TempAddresses("Rate") & "," _
                & dic_TempAddresses("Strike") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," & dic_TempAddresses("OptMat") _
                & ",""ACT/365"")," & dic_TempAddresses("Vol") & ")/100*Calc_YearFrac(" & dic_TempAddresses("CalcStart") & "," _
                & dic_TempAddresses("CalcEnd") & ",""" & fld_Params.Daycount & """,""" & fld_Params.PmtFreq & """," _
                & bln_ActActFwdGeneration & ")"
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).NumberFormat = str_CurrencyFormat

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Formula = "=cyReadIRCurve(""" & fld_Params.Curve_Disc & """," & dic_Addresses("ValDate") _
                & "," & dic_TempAddresses("PmtDate") & ",""DF"",,False)"

            int_ActiveCol = int_ActiveCol - 4
        Next int_ctr

        ' Store flows range
        Dim rng_CapletPrices As Range: Set rng_CapletPrices = .Offset(int_ActiveRow + 1, int_ActiveCol + 3).Resize(int_NumPmts, 1)
    End With

    ' Output premium flow
    int_ActiveCol = int_ActiveCol - 9
    int_ActiveRow = int_ActiveRow + int_NumPmts + 3
    With rng_OutputTopLeft
        .Offset(int_ActiveRow, int_ActiveCol).Value = "COST LEG"
        .Offset(int_ActiveRow, int_ActiveCol).Font.Italic = True

        int_ActiveRow = int_ActiveRow + 1
        Call scf_Premium.OutputReport(.Offset(int_ActiveRow, int_ActiveCol), "Cash", str_CCY_PnL, -int_Sign, True, _
            dic_Addresses, False)
    End With

    ' Fill in overall formulas
    dic_Addresses("Range_TotalMV").Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*SUMPRODUCT((" _
        & rng_CapletPrices.Address(False, False) & ")*(" & rng_CapletPrices.Offset(0, 1).Address(False, False) & ")*(" _
        & rng_CapletPrices.Offset(0, -1).Address(False, False) & "=""MV""))*cyGetFXDiscSpot(" _
        & dic_Addresses("OptionCCY") & "," & dic_Addresses("PnLCCY") & ")"

    dic_Addresses("Range_TotalCash").Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*SUMPRODUCT((" _
        & rng_CapletPrices.Address(False, False) & ")*(" & rng_CapletPrices.Offset(0, 1).Address(False, False) & ")*(" _
        & rng_CapletPrices.Offset(0, -1).Address(False, False) & "=""CASH""))*cyGetFXDiscSpot(" _
        & dic_Addresses("OptionCCY") & "," & dic_Addresses("PnLCCY") & ")+" & dic_Addresses("SCF_PV")

    dic_Addresses("Range_PnL").Formula = "=" & dic_Addresses("TotalMV") & "+" & dic_Addresses("TotalCash")
End Sub

Public Sub OutputReport_IRDig(rng_OutputTopLeft As Range, str_LegName As String, dbl_Strike As Double, enu_Direction As OptionDirection, _
    dblLst_Vols As Collection, dblLst_ShiftedVols As Collection, int_Deduction As Integer, cal_Deduction As Calendar, str_Position As String, _
    bln_IsDigital As Boolean, str_CCY_PnL As String, scf_Premium As SCF, int_Sign As Integer)
    ' ## Display rates, flows and discount factors
    Dim str_CurrencyFormat As String: str_CurrencyFormat = Gather_CurrencyFormat()
    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    Dim str_Address_Cash As String
    Dim rng_MV As Range
    Dim int_ctr As Integer
    Dim int_ActiveRow As Integer: int_ActiveRow = 0
    Dim int_ActiveCol As Integer: int_ActiveCol = 0

    Dim dic_Addresses As Dictionary: Set dic_Addresses = New Dictionary
    dic_Addresses.CompareMode = CompareMethod.TextCompare

    Dim dic_TempAddresses As Dictionary: Set dic_TempAddresses = New Dictionary
    dic_TempAddresses.CompareMode = CompareMethod.TextCompare
    ' General info
    With rng_OutputTopLeft
        .Offset(int_ActiveRow, int_ActiveCol).Value = "OVERALL"
        .Offset(int_ActiveRow, int_ActiveCol).Font.Italic = True

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Value Date:"
        .Offset(int_ActiveRow, int_ActiveCol + 1).Value = fld_Params.ValueDate
        .Offset(int_ActiveRow, int_ActiveCol + 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("ValDate", .Offset(int_ActiveRow, int_ActiveCol + 1).Address(True, True))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "PnL:"
        .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
        Call dic_Addresses.Add("Range_PnL", .Offset(int_ActiveRow, int_ActiveCol + 1))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "CCY:"
        .Offset(int_ActiveRow, int_ActiveCol + 1).Value = str_CCY_PnL
        Call dic_Addresses.Add("PnLCCY", .Offset(int_ActiveRow, int_ActiveCol + 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 2
        .Offset(int_ActiveRow, int_ActiveCol).Value = "MV (" & str_CCY_PnL & "):"
        .Offset(int_ActiveRow, int_ActiveCol + 1).NumberFormat = str_CurrencyFormat
        .Offset(int_ActiveRow, int_ActiveCol + 1).Interior.ColorIndex = 20
        Call dic_Addresses.Add("Range_TotalMV", .Offset(int_ActiveRow, int_ActiveCol + 1))
        Call dic_Addresses.Add("TotalMV", .Offset(int_ActiveRow, int_ActiveCol + 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Cash (" & str_CCY_PnL & "):"
        .Offset(int_ActiveRow, int_ActiveCol + 1).NumberFormat = str_CurrencyFormat
        .Offset(int_ActiveRow, int_ActiveCol + 1).Interior.ColorIndex = 20
        Call dic_Addresses.Add("Range_TotalCash", .Offset(int_ActiveRow, int_ActiveCol + 1))
        Call dic_Addresses.Add("TotalCash", .Offset(int_ActiveRow, int_ActiveCol + 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 3
        .Offset(int_ActiveRow, int_ActiveCol).Value = "OPTION LEG"
        .Offset(int_ActiveRow, int_ActiveCol).Font.Italic = True

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "CCY:"
        .Offset(int_ActiveRow, int_ActiveCol + 1).Value = fld_Params.CCY
        Call dic_Addresses.Add("OptionCCY", .Offset(int_ActiveRow, int_ActiveCol + 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Payout:"
        If enu_Direction = OptionDirection.CallOpt Then
            .Offset(int_ActiveRow, int_ActiveCol + 1).Value = "Call"
        Else
            .Offset(int_ActiveRow, int_ActiveCol + 1).Value = "Put"
        End If
        Call dic_Addresses.Add("Payout", .Offset(int_ActiveRow, int_ActiveCol + 1).Address(True, True))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Position:"
        .Offset(int_ActiveRow, int_ActiveCol + 1).Value = str_Position

        Call dic_Addresses.Add("Position", .Offset(int_ActiveRow, int_ActiveCol + 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "IsDigital:"
        .Offset(int_ActiveRow, int_ActiveCol + 1).Value = bln_IsDigital


     End With

    int_ActiveRow = int_ActiveRow + 2
    With rng_OutputTopLeft
        ' Rate headings
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Period Start"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Period End"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Est Start"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Est End"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Notional"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Rate"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Strike"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "OriginalVol"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "ShiftedVol"

        int_ActiveCol = int_ActiveCol - 8

        ' Rates
        For int_ctr = 1 To lngLst_PeriodEnd.count
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = lngLst_PeriodStart(int_ctr)

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = lngLst_PeriodEnd(int_ctr)

            If bln_IsFixed = True Then
                int_ActiveCol = int_ActiveCol + 4
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = fld_Params.RateOrMargin
            Else
                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = lngLst_EstStart(int_ctr)

                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = lngLst_EstEnd(int_ctr)

                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = fld_Params.Notional * dblLst_AmortFactors(int_ctr)
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).NumberFormat = str_CurrencyFormat

                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = dblLst_Rates(int_ctr) + dblLst_Margins(int_ctr)
            End If

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = dbl_Strike

            int_ActiveCol = int_ActiveCol + 1
            If dblLst_Vols(int_ctr) <> 0 Then
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = dblLst_Vols(int_ctr)
            End If

            int_ActiveCol = int_ActiveCol + 1
            If dblLst_ShiftedVols(int_ctr) <> 0 Then
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = dblLst_ShiftedVols(int_ctr)
            End If

            int_ActiveCol = int_ActiveCol - 8
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Resize(1, 4).NumberFormat = str_DateFormat
        Next int_ctr
    End With

    int_ActiveCol = int_ActiveCol + 9
    With rng_OutputTopLeft
        ' Payment headings
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Opt Mat"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Pmt Date"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Type"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Flow"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "DF"

        int_ActiveCol = int_ActiveCol - 4

        ' Payments
        Dim int_NumPmts As Integer: int_NumPmts = lngLst_PmtDates.count
        For int_ctr = 1 To int_NumPmts
            Call dic_TempAddresses.RemoveAll

            int_ActiveCol = int_ActiveCol - 9
            Call dic_TempAddresses.Add("CalcStart", .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Address(False, False))

            int_ActiveCol = int_ActiveCol + 1
            Call dic_TempAddresses.Add("CalcEnd", .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Address(False, False))

            int_ActiveCol = int_ActiveCol + 3
            Call dic_TempAddresses.Add("Notional", .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Address(False, False))

            int_ActiveCol = int_ActiveCol + 1
            Call dic_TempAddresses.Add("Rate", .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Address(False, False))

            int_ActiveCol = int_ActiveCol + 1
            Call dic_TempAddresses.Add("Strike", .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Address(False, False))

            int_ActiveCol = int_ActiveCol + 1
            Call dic_TempAddresses.Add("OriginalVol", .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Address(False, False))

            int_ActiveCol = int_ActiveCol + 1
            Call dic_TempAddresses.Add("ShiftedVol", .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Address(False, False))

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = date_workday(lngLst_PeriodStart(int_ctr), int_Deduction, cal_Deduction.HolDates, cal_Deduction.Weekends)
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).NumberFormat = str_DateFormat
            Call dic_TempAddresses.Add("OptMat", .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Address(False, False))

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = lngLst_PeriodEnd(int_ctr)
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).NumberFormat = str_DateFormat
            Call dic_TempAddresses.Add("PmtDate", .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Address(False, False))

            int_ActiveCol = int_ActiveCol + 1
            If intLst_IsMV(int_ctr) = 1 Then
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = "MV"
            Else
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = "CASH"
            End If

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Formula = "=" & dic_TempAddresses("Notional") _
                & "*Calc_BSPrice_DigitalSmileOn(IF(" & dic_Addresses("Payout") & "=""Call"",1,-1)," & dic_TempAddresses("Rate") & "," _
                & dic_TempAddresses("Strike") & ",Calc_YearFrac(" & dic_Addresses("ValDate") & "," & dic_TempAddresses("OptMat") _
                & ",""ACT/365"")," & dic_TempAddresses("OriginalVol") & "," & dic_TempAddresses("ShiftedVol") & ")/100*Calc_YearFrac(" & dic_TempAddresses("CalcStart") & "," _
                & dic_TempAddresses("CalcEnd") & ",""" & fld_Params.Daycount & """,""" & fld_Params.PmtFreq & """," _
                & bln_ActActFwdGeneration & ")"
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).NumberFormat = str_CurrencyFormat

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Formula = "=cyReadIRCurve(""" & fld_Params.Curve_Disc & """," & dic_Addresses("ValDate") _
                & "," & dic_TempAddresses("PmtDate") & ",""DF"",,False)"

            int_ActiveCol = int_ActiveCol - 4
        Next int_ctr

        ' Store flows range
        Dim rng_CapletPrices As Range: Set rng_CapletPrices = .Offset(int_ActiveRow + 1, int_ActiveCol + 3).Resize(int_NumPmts, 1)
    End With

    ' Output premium flow
    int_ActiveCol = int_ActiveCol - 9
    int_ActiveRow = int_ActiveRow + int_NumPmts + 3
    With rng_OutputTopLeft
        .Offset(int_ActiveRow, int_ActiveCol).Value = "COST LEG"
        .Offset(int_ActiveRow, int_ActiveCol).Font.Italic = True

        int_ActiveRow = int_ActiveRow + 1
        Call scf_Premium.OutputReport(.Offset(int_ActiveRow, int_ActiveCol), "Cash", str_CCY_PnL, -int_Sign, True, _
            dic_Addresses, False)
    End With

    ' Fill in overall formulas
    dic_Addresses("Range_TotalMV").Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*SUMPRODUCT((" _
        & rng_CapletPrices.Address(False, False) & ")*(" & rng_CapletPrices.Offset(0, 1).Address(False, False) & ")*(" _
        & rng_CapletPrices.Offset(0, -1).Address(False, False) & "=""MV""))*cyGetFXDiscSpot(" _
        & dic_Addresses("OptionCCY") & "," & dic_Addresses("PnLCCY") & ")"

    dic_Addresses("Range_TotalCash").Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*SUMPRODUCT((" _
        & rng_CapletPrices.Address(False, False) & ")*(" & rng_CapletPrices.Offset(0, 1).Address(False, False) & ")*(" _
        & rng_CapletPrices.Offset(0, -1).Address(False, False) & "=""CASH""))*cyGetFXDiscSpot(" _
        & dic_Addresses("OptionCCY") & "," & dic_Addresses("PnLCCY") & ")+" & dic_Addresses("SCF_PV")

    dic_Addresses("Range_PnL").Formula = "=" & dic_Addresses("TotalMV") & "+" & dic_Addresses("TotalCash")
End Sub

Public Function OutputReport_Bond(rng_OutputTopLeft As Range, lng_ValDate As Long, str_CCY_PnL As String, _
    int_Sign As Integer, str_Curve_SpotDisc As String, bln_IsFuturesUnd As Boolean, Optional scf_Purchase As SCF = Nothing, _
    Optional str_Address_ZSpread As String = "-") As Dictionary

    ' ## Display rates, flows and discount factors
    Dim dic_output As New Dictionary: dic_output.CompareMode = CompareMethod.TextCompare
    Dim int_ActiveRow As Integer: int_ActiveRow = 0
    Dim int_ActiveCol As Integer: int_ActiveCol = 0
    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    Dim str_CurrencyFormat As String: str_CurrencyFormat = Gather_CurrencyFormat()
    Dim dic_Addresses As New Dictionary: dic_Addresses.CompareMode = CompareMethod.TextCompare
    Dim dic_TempAddresses As New Dictionary: dic_TempAddresses.CompareMode = CompareMethod.TextCompare
    Dim rng_PnL As Range, rng_MV As Range, rng_Cash As Range
    Dim int_ctr As Integer


    ' General info
    With rng_OutputTopLeft
        If bln_IsFuturesUnd = True Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, int_ActiveCol).Value = "UND FLOWS"
        Else
            .Offset(int_ActiveRow, int_ActiveCol).Value = "OVERALL"
        End If
        .Offset(int_ActiveRow, 0).Font.Italic = True

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Value date:"
        .Offset(int_ActiveRow, int_ActiveCol + 1).Value = lng_ValDate
        .Offset(int_ActiveRow, int_ActiveCol + 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("ValDate", .Offset(int_ActiveRow, int_ActiveCol + 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Spot date:"
        .Offset(int_ActiveRow, int_ActiveCol + 1).Value = fld_Params.ValueDate
        .Offset(int_ActiveRow, int_ActiveCol + 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("SpotDate", .Offset(int_ActiveRow, int_ActiveCol + 1).Address(False, False))

        If bln_IsFuturesUnd = True Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, int_ActiveCol).Value = "MV (" & str_CCY_PnL & "):"
            .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
            Set rng_MV = .Offset(int_ActiveRow, 1)
        Else
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, int_ActiveCol).Value = "PnL:"
            .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
            .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
            Set rng_PnL = .Offset(int_ActiveRow, 1)

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, int_ActiveCol).Value = "Currency:"
            .Offset(int_ActiveRow, int_ActiveCol + 1).Value = str_CCY_PnL
            Call dic_Addresses.Add("PnLCCY", .Offset(int_ActiveRow, int_ActiveCol + 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 2
            .Offset(int_ActiveRow, int_ActiveCol).Value = "MV (" & str_CCY_PnL & "):"
            .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
            .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
            Set rng_MV = .Offset(int_ActiveRow, 1)

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, int_ActiveCol).Value = "Cash (" & str_CCY_PnL & "):"
            .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
            .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
            Set rng_Cash = .Offset(int_ActiveRow, 1)

            int_ActiveRow = int_ActiveRow + 3
            .Offset(int_ActiveRow, int_ActiveCol).Value = "BOND LEG"
            .Offset(int_ActiveRow, 0).Font.Italic = True

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, int_ActiveCol).Value = "Position:"
            If int_Sign = 1 Then
                .Offset(int_ActiveRow, int_ActiveCol + 1).Value = "B"
            Else
                .Offset(int_ActiveRow, int_ActiveCol + 1).Value = "S"
            End If

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, int_ActiveCol).Value = "Rate type:"
            If bln_IsFixed = True Then
                .Offset(int_ActiveRow, int_ActiveCol + 1).Value = "Fixed"
            Else
                .Offset(int_ActiveRow, int_ActiveCol + 1).Value = "Floating"
            End If

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, int_ActiveCol).Value = "Currency:"
            .Offset(int_ActiveRow, int_ActiveCol + 1).Value = fld_Params.CCY
            Call dic_Addresses.Add("BondCCY", .Offset(int_ActiveRow, int_ActiveCol + 1).Address(False, False))
        End If
    End With

    Call dic_output.Add("Bond_MV", rng_MV.Address(False, False))

    ' Rate headings
    int_ActiveRow = int_ActiveRow + 2
    With rng_OutputTopLeft
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Period Start"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Period End"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Est Start"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Est End"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Rate"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Margin"

        ' Rates
        If fld_Params.PExch_Start = True Then int_ActiveRow = int_ActiveRow + 1
        int_ActiveCol = int_ActiveCol - 5
        For int_ctr = 1 To lngLst_PeriodEnd.count
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = lngLst_PeriodStart(int_ctr)

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = lngLst_PeriodEnd(int_ctr)

            If bln_IsFixed = True Then
                int_ActiveCol = int_ActiveCol + 3
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = fld_Params.RateOrMargin
                int_ActiveCol = int_ActiveCol - 4
            Else
                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = lngLst_EstStart(int_ctr)

                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = lngLst_EstEnd(int_ctr)

                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = dblLst_Rates(int_ctr)

                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = dblLst_Margins(int_ctr)

                int_ActiveCol = int_ActiveCol - 5
            End If

            .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Resize(1, 4).NumberFormat = str_DateFormat
        Next int_ctr

        If fld_Params.PExch_Start = True Then int_ActiveRow = int_ActiveRow - 1
    End With

    ' Payment headings
    int_ActiveCol = int_ActiveCol + 7
    With rng_OutputTopLeft
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Pmt Date"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Type"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Int Flow"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "Prn Flow"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "DF to spot"

        int_ActiveCol = int_ActiveCol + 1
        .Offset(int_ActiveRow, int_ActiveCol).Value = "DF to val"

        int_ActiveCol = int_ActiveCol - 5
    End With

    ' Optional principal exchange at start
    If fld_Params.PExch_Start = True Then
        int_ActiveRow = int_ActiveRow + 1
        With rng_OutputTopLeft
            .Offset(int_ActiveRow, int_ActiveCol).Value = fld_Params.Swapstart
            .Offset(int_ActiveRow, int_ActiveCol).NumberFormat = str_DateFormat
            Call dic_TempAddresses.Add("StartDate", .Offset(int_ActiveRow, int_ActiveCol).Address(False, False))

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ActiveRow, int_ActiveCol).Value = str_PrnAtStartPnlType

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ActiveRow, int_ActiveCol).Value = 0
            .Offset(int_ActiveRow, int_ActiveCol).NumberFormat = str_CurrencyFormat

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ActiveRow, int_ActiveCol).Value = dbl_PrnAtStart
            .Offset(int_ActiveRow, int_ActiveCol).NumberFormat = str_CurrencyFormat

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ActiveRow, int_ActiveCol).Value = dbl_StartDF

            int_ActiveCol = int_ActiveCol + 1
            .Offset(int_ActiveRow, int_ActiveCol).Formula = "=cyReadIRCurve(""" & str_Curve_SpotDisc & """," _
            & dic_Addresses("ValDate") & "," & dic_TempAddresses("StartDate") & ",""DF"")"

            int_ActiveCol = int_ActiveCol - 5
            Call dic_TempAddresses.RemoveAll
        End With
    End If

    ' Payments
    Dim int_ActiveIndex As Integer: int_ActiveIndex = 0
    Dim int_RangeHeight As Integer: int_RangeHeight = lngLst_PmtDates.count * int_CalcsPerPmt
    Dim rng_FlowTypes As Range
    With rng_OutputTopLeft
        ' Set ranges for calculation of MV and Cash
        If fld_Params.PExch_Start = True Then
            Set rng_FlowTypes = .Offset(int_ActiveRow, int_ActiveCol + 1).Resize(int_RangeHeight + 1, 1)
        Else
            Set rng_FlowTypes = .Offset(int_ActiveRow + 1, int_ActiveCol + 1).Resize(int_RangeHeight, 1)
        End If

        For int_ctr = 1 To int_RangeHeight
            int_ActiveRow = int_ActiveRow + 1

            If int_ctr Mod int_CalcsPerPmt = 0 Then
                int_ActiveIndex = int_ActiveIndex + 1
                .Offset(int_ActiveRow, int_ActiveCol).Value = lngLst_PmtDates(int_ActiveIndex)
                .Offset(int_ActiveRow, int_ActiveCol).NumberFormat = str_DateFormat
                Call dic_TempAddresses.Add("PmtDate", .Offset(int_ActiveRow, int_ActiveCol).Address(False, False))

                int_ActiveCol = int_ActiveCol + 1
                If intLst_IsMV(int_ActiveIndex) = 1 Then
                    .Offset(int_ActiveRow, int_ActiveCol).Value = "MV"
                Else
                    .Offset(int_ActiveRow, int_ActiveCol).Value = "CASH"
                End If

                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ActiveRow, int_ActiveCol).Value = dblLst_IntFlows(int_ActiveIndex)
                .Offset(int_ActiveRow, int_ActiveCol).NumberFormat = str_CurrencyFormat

                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ActiveRow, int_ActiveCol).Value = dblLst_PrnFlows(int_ActiveIndex)
                .Offset(int_ActiveRow, int_ActiveCol).NumberFormat = str_CurrencyFormat

                int_ActiveCol = int_ActiveCol + 1
                If bln_IsFuturesUnd = True Then
                    .Offset(int_ActiveRow, int_ActiveCol).Formula = "=cyReadIRCurve(""" & fld_Params.Curve_Disc & """," _
                        & dic_Addresses("SpotDate") & "," & dic_TempAddresses("PmtDate") & ",""DF"",,False," & str_Address_ZSpread & ")"
                Else
                    .Offset(int_ActiveRow, int_ActiveCol).Formula = "=cyReadIRCurve(""" & fld_Params.Curve_Disc & """," _
                        & dic_Addresses("SpotDate") & "," & dic_TempAddresses("PmtDate") & ",""DF"",,False)"
                End If

                int_ActiveCol = int_ActiveCol + 1
                .Offset(int_ActiveRow, int_ActiveCol).Formula = "=cyReadIRCurve(""" & str_Curve_SpotDisc & """," _
                    & dic_Addresses("ValDate") & ",MIN(" & dic_Addresses("SpotDate") & "," & dic_TempAddresses("PmtDate") & "),""DF"",,False)"

                int_ActiveCol = int_ActiveCol - 5
                Call dic_TempAddresses.RemoveAll
            End If
        Next int_ctr
    End With
    int_ActiveCol = int_ActiveCol - 7

    ' Purchase cost
    If Not scf_Purchase Is Nothing Then
        With rng_OutputTopLeft
            int_ActiveRow = int_ActiveRow + 3
            .Offset(int_ActiveRow, int_ActiveCol).Value = "COST LEG"
            .Offset(int_ActiveRow, int_ActiveCol).Font.Italic = True

            int_ActiveRow = int_ActiveRow + 1
            Call scf_Purchase.OutputReport(.Offset(int_ActiveRow, int_ActiveCol), "PV", _
                str_CCY_PnL, -int_Sign, False, dic_Addresses, False)
        End With
    End If

    ' Fill valuation formulas
    Dim str_Sign As String: If int_Sign = -1 And bln_IsFuturesUnd = False Then str_Sign = "-" Else str_Sign = ""
    If bln_IsFuturesUnd = True Then
        rng_MV.Formula = "=" & str_Sign & "SUMPRODUCT((" & rng_FlowTypes.Address(False, False) & "=""MV"")*(" _
            & rng_FlowTypes.Offset(0, 1).Address(False, False) & "+" & rng_FlowTypes.Offset(0, 2).Address(False, False) _
            & ")*(" & rng_FlowTypes.Offset(0, 3).Address(False, False) & ")*" & rng_FlowTypes.Offset(0, 4).Address(False, False) _
            & ")"
    Else
        rng_MV.Formula = "=" & str_Sign & "SUMPRODUCT((" & rng_FlowTypes.Address(False, False) & "=""MV"")*(" _
            & rng_FlowTypes.Offset(0, 1).Address(False, False) & "+" & rng_FlowTypes.Offset(0, 2).Address(False, False) _
            & ")*(" & rng_FlowTypes.Offset(0, 3).Address(False, False) & ")*" & rng_FlowTypes.Offset(0, 4).Address(False, False) _
            & ")*cyGetFXDiscSpot(""" & fld_Params.CCY & """," & dic_Addresses("PnLCCY") & ")"
    End If

    Dim str_OptionalSCF As String
    If scf_Purchase Is Nothing Then str_OptionalSCF = "" Else str_OptionalSCF = "+" & dic_Addresses("SCF_PV")

    If bln_IsFuturesUnd = False Then
        rng_Cash.Formula = "=" & str_Sign & "SUMPRODUCT((" & rng_FlowTypes.Address(False, False) & "=""CASH"")*(" _
            & rng_FlowTypes.Offset(0, 1).Address(False, False) & "+" & rng_FlowTypes.Offset(0, 2).Address(False, False) _
            & ")*(" & rng_FlowTypes.Offset(0, 3).Address(False, False) & ")*(" & rng_FlowTypes.Offset(0, 4).Address(False, False) _
            & "))*cyGetFXDiscSpot(" & dic_Addresses("BondCCY") & "," & dic_Addresses("PnLCCY") & ")" & str_OptionalSCF
        rng_PnL.Formula = "=" & rng_MV.Address(False, False) & "+" & rng_Cash.Address(False, False)
    End If

    Set OutputReport_Bond = dic_output
End Function


Public Property Get FixLegPmtDate() As Collection
    Set FixLegPmtDate = lngLst_PmtDates
End Property


Public Property Get FloatLegCashFlow(str_curve As String, col_FixLegPmtDate As Collection, bln_acc As Boolean) As Collection

    Set irc_Est = GetObject_IRCurve(str_curve, True, False)
    RecalcIntFlows

If bln_acc = True Then

    Dim int_NumFixPmt As Integer: int_NumFixPmt = col_FixLegPmtDate.count
    Dim int_NumFltPmt As Integer: int_NumFltPmt = lngLst_PmtDates.count

    Dim int_i As Integer
    Dim int_j As Integer
    Dim int_k As Integer

    Dim dbl_Output As Double

    Dim col_output As Collection
    Set col_output = New Collection

    Dim int_FltIdentifier As Integer
    int_FltIdentifier = 1

    For int_i = 1 To int_NumFixPmt
        dbl_Output = 0
        For int_j = int_FltIdentifier To int_NumFltPmt

            If col_FixLegPmtDate(int_i) = lngLst_PmtDates(int_j) Then
                For int_k = int_FltIdentifier To int_j
                    dbl_Output = dbl_Output + dblLst_IntFlows(int_k) * dblLst_DFs(int_k) / dblLst_DFs(int_j)
                Next int_k

                int_FltIdentifier = int_j + 1
                Exit For
            End If
        Next int_j
        col_output.Add dbl_Output
    Next int_i

    Set FloatLegCashFlow = col_output

Else
    Set FloatLegCashFlow = dblLst_IntFlows
End If

End Property