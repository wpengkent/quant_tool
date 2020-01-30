Option Explicit

' ## MEMBER DATA
' Curve dependencies
Private irc_Disc As Data_IRCurve, irc_Est As Data_IRCurve

' Variable dates
Private lng_ValDate As Long

' Dynamic stored values
Private dblLst_Rates As Collection, dblLst_IntFlows As Collection, dblLst_DFs As Collection, dbl_StartDF As Double
Private intLst_IsMV As Collection, intLst_IsCash As Collection, str_PrnAtStartPnlType As String
Private dblLst_FlowDurations As Collection  ' Time between valuation and payment date
Private dbl_ZSpread As Double

' Static values - general
Private fld_Params As BondParams, dic_CurveDependencies As Dictionary
Private Const str_NotUsed As String = "-", str_Daycount_Duration As String = "ACT/365", bln_ActActFwdGeneration As Boolean = False
Private bln_IsFixed As Boolean

' Static values - counts and measures
Private int_CalcsPerPmt As Integer
Private str_PeriodLength As String, str_EstPeriodLength As String
Private dbl_periodicity As Double
Private int_CtrStart As Integer

' Static values - dates, calendars and durations
Private lngLst_PeriodStart As Collection, lngLst_PeriodEnd As Collection
Private lngLst_EstStart As Collection, lngLst_EstEnd As Collection, lngLst_PmtDates As Collection
Private lng_startdate As Long, lng_EndDate As Long, lng_TradeDate As Long
Private cal_pmt As Calendar, cal_est As Calendar
Private dblLst_CalcPeriodDurations As Collection  ' Time between period start and end
Private dblLst_EstPeriodDurations As Collection  ' Time between estimation period start and end

' Static values - rates, values and factors
Private dic_GlobalStaticInfo As Dictionary, dic_CurveSet As Dictionary
Private dblLst_Margins As Collection
Private dblLst_AmortFactors As Collection
Private dbl_PrnAtStart As Double, dblLst_PrnFlows As Collection


' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As BondParams, Optional dic_CurveSetInput As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing)

    ' Set static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    fld_Params = fld_ParamsInput

    ' Initialize dynamic values
    lng_ValDate = fld_ParamsInput.ValueDate

    ' Trade Date for Bond
    lng_TradeDate = fld_ParamsInput.TradeDate

    ' Determine leg type
    bln_IsFixed = (fld_Params.index = str_NotUsed)


    ' Force floating rate estimation if margin is specified
    If bln_IsFixed = False And fld_Params.RateOrMargin <> 0 Then fld_Params.FloatEst = True

    ' Set calendars
    Dim cas_Found As CalendarSet: Set cas_Found = dic_GlobalStaticInfo(StaticInfoType.CalendarSet)
    cal_pmt = cas_Found.Lookup_Calendar(fld_Params.PmtCal)
    If bln_IsFixed = False Then cal_est = cas_Found.Lookup_Calendar(fld_Params.estcal)

    ' Derive number of payment dates
    Dim int_NumPmts As Integer, lng_WindowStart As Long, lng_WindowEnd As Long

    ' Determine number of payments based on the specified start and end dates
    If Left(UCase(fld_Params.StubType), 8) <> "BOTH END" Then
        lng_WindowStart = fld_Params.FirstAccDate
        lng_WindowEnd = fld_Params.MatDate
    ElseIf Right(UCase(fld_Params.StubType), 9) = "(FORWARD)" Then
        lng_WindowStart = fld_Params.RollDate
        lng_WindowEnd = fld_Params.MatDate
    ElseIf Right(UCase(fld_Params.StubType), 10) = "(BACKWARD)" Then
        lng_WindowStart = fld_Params.FirstAccDate
        lng_WindowEnd = fld_Params.RollDate
    End If

    Dim bln_ForwardDirection As Boolean: If UCase(fld_Params.StubType) = "UP FRONT" Or Right(UCase(fld_Params.StubType), 10) = "(BACKWARD)" Then _
        bln_ForwardDirection = False Else bln_ForwardDirection = True
    int_NumPmts = Calc_NumPmtsInWindow(lng_WindowStart, lng_WindowEnd, fld_Params.PmtFreq, cal_pmt, fld_Params.BDC, _
        fld_Params.EOM, bln_ForwardDirection)
    If Left(UCase(fld_Params.StubType), 8) = "BOTH END" Then int_NumPmts = int_NumPmts + 1

    ' Derive number of calculation dates
    Dim int_NumCalcs As Integer

    int_NumCalcs = int_NumPmts
    str_PeriodLength = fld_Params.PmtFreq

    ' Input length of floating rates
    If bln_IsFixed = False Then str_EstPeriodLength = fld_Params.index

    int_CalcsPerPmt = int_NumCalcs / int_NumPmts


    ' Derive dates
    Call FillPeriodDates(fld_Params.StubType, int_NumCalcs)
    ' Special treatment to date convention of EOM can only be either 28(for Feb) or 30 (for other months)
    If fld_Params.IsEOM2830 = True Then Call AjustPeriodDates_2830

    If bln_IsFixed = False Then Call FillEstDates
    Call ModifyStartDates
    Call FillPmtDates  ' For bonds with past fixings, floating estimation must be turned on
    Call CategorizeFlows

    ' Store durations
    Call FillFlowDurations
    Call FillPeriodDurations

    ' Read amortization schedule if it exists
    Call FillAmortFactors

    ' Set dependent curves
    If dic_CurveSetInput Is Nothing Then
        Set irc_Disc = GetObject_IRCurve(fld_Params.Curve_Disc, True, False)
        If bln_IsFixed = False Then Set irc_Est = GetObject_IRCurve(fld_Params.Curve_Est, True, False)
    Else
        Set dic_CurveSet = dic_CurveSetInput
        Dim dic_IRCurves As Dictionary: Set dic_IRCurves = dic_CurveSet(CurveType.IRC)
        Set irc_Disc = dic_IRCurves(fld_Params.Curve_Disc)
        If bln_IsFixed = False Then Set irc_Est = dic_IRCurves(fld_Params.Curve_Est)
    End If

    ' Fill rates and flows
    Call FillPrincipalFlows  ' Fill if intermediate payments are enabled, otherwise fill with zero
    Call FillRates
    Call FillFixings 'Fixing is called for floating or for stepping rates bond

    Call FillIntFlows  ' Calculate undiscounted flows

    ' Get Spread on top of discount curve.
    ' If calculation type = Zero Curve, constant spread is added on top of the discount curve
    ' If calculation type - MTM, Market Price of the bond is an input and spread is solved using secant method
    Call getZeroSpread


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


    'By SW for yield solving --> reseting the parameters.
    fld_Params.Yield = 0
    fld_Params.Duration = 0
    fld_Params.ModifiedDuration = 0
    dbl_periodicity = 0
    ''''''

End Sub


' ## PROPERTIES
Public Property Get marketvalue() As Double
    marketvalue = CalcValue("MV")
End Property

Public Property Get Cash() As Double
    Cash = CalcValue("CASH")
End Property

Public Property Get PnL() As Double
    PnL = marketvalue + Cash
End Property

Public Property Get BondYield() As Double
    BondYield = CalcYield
End Property

Public Property Get BondDuration() As Double
'Murex computes duration = modified duration * one period capitalization factor --> fld_Params.Duration
'Generally accepted duration = weighted tenor of bond with its cash flow present value --> fld_Params.MacaulayDuration

    'BondDuration = fld_Params.Duration
    BondDuration = fld_Params.MacaulayDuration
End Property

Public Property Get BondModifiedDuration() As Double
    BondModifiedDuration = fld_Params.ModifiedDuration
End Property

Public Property Get Params() As BondParams
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
        Dim intArr_StartDateBeforeFixing() As Integer

        dbl_DV01_EstCurve = 0
        If Not irc_Est Is Nothing Then
            If irc_Est.CurveName = str_curve Then
            ' Gather notional required for a unit interest flow
            int_NumPeriods = lngLst_PeriodEnd.count
            ReDim dblArr_NotionalFactors(1 To int_NumPeriods) As Double
            ReDim intArr_StartDateBeforeFixing(1 To int_NumPeriods) As Integer
            Set dblLst_DFs_PeriodEnd = New Collection

            For int_ctr = 1 To int_NumPeriods
                dblArr_NotionalFactors(int_ctr) = irc_Est.Lookup_Rate(lngLst_EstStart(int_ctr), lngLst_EstEnd(int_ctr), _
                    fld_Params.Daycount, , fld_Params.PmtFreq, , True)
                If (lngLst_EstStart(int_ctr) > lng_ValDate) Then intArr_StartDateBeforeFixing(int_ctr) = 1 Else intArr_StartDateBeforeFixing(int_ctr) = 0
                Call dblLst_DFs_PeriodEnd.Add(irc_Disc.Lookup_Rate(lng_ValDate, lngLst_PeriodEnd(int_ctr), "DF", , , False))
            Next int_ctr

            ' Create array with full principal as each element
            dbl_DV01_EstCurve = Calc_SumProductOnList(dblArr_NotionalFactors, dblLst_CalcPeriodDurations, _
                dblLst_EstPeriodDurations, intArr_StartDateBeforeFixing, dblLst_DFs_PeriodEnd, dblLst_AmortFactors) * fld_Params.Notional / 100 * 0.0001
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

Private Sub getZeroSpread()
' ## zero spread is determined based on calculation type
' ## If "Zero Curve" then spread is an input
' ## If "MTM" then input is Bond Price and spread is solved

    If fld_Params.CalType = "Zero Curve" Then

        dbl_ZSpread = fld_Params.DiscSpread

    ElseIf fld_Params.CalType = "MTM" Then
        'Solve for spread
        irc_Disc.SetCurveState (original)
        'the following will output ZSpread
        Call ForceMVToValue(fld_Params.BondMarketPrice / 100 * fld_Params.Notional)
        irc_Disc.SetCurveState (Final)

    Else
        dbl_ZSpread = 0
    End If

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
    SolveParRate = Solve_Secant(ThisWorkbook, "SolverFuncXY_ParToMV", dic_SecantParams, fld_Params.RateOrMargin, _
        fld_Params.RateOrMargin + 1, (irl_LegToMatch.marketvalue - dbl_ExistingMV_MatchLeg) * dbl_FXConv + dbl_ExistingMV_SolveLeg, _
        fld_Params.Notional * 0.000000000000001, 50, -100, dic_SecantOutputs)

    ' Reset back to original rate if required
    If bln_ResetToOrigPar = True Then Call SetRateOrMargin(dbl_OrigRate)
End Function

Public Sub ForceMVToValue(dbl_TargetMV As Double)
    ' ## Set ZSpread such that the MV equals the market price
    ' Store static parameters for secant solver
    Dim dic_SecantParams As Dictionary: Set dic_SecantParams = New Dictionary
    Call dic_SecantParams.Add("irl_Leg", Me)

    ' Solve
    Dim dic_SecantOutputs As Dictionary: Set dic_SecantOutputs = New Dictionary
    Call Solve_Secant(ThisWorkbook, "SolverFuncXY_ZSpreadToMV_Bond", dic_SecantParams, dbl_ZSpread, _
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
Private Sub FillPeriodDates(str_StubType As String, int_NumPeriods As Integer)
    ' ## Generate start and end dates for cash flow calculation
    Dim lngLst_Follower As New Collection, lngLst_Driver As New Collection
    Dim int_Sign As Integer: If fld_Params.IsFwdGeneration = True Then int_Sign = 1 Else int_Sign = -1
    Dim int_Direction As Integer: If UCase(str_StubType) = "UP FRONT" Or Right(UCase(str_StubType), 10) = "(BACKWARD)" Then int_Direction = -1 Else int_Direction = 1
    Dim lng_tempDate As Long

    ' Perform generation
    Dim int_ctr As Integer
    For int_ctr = 1 To int_NumPeriods

        If int_ctr = 1 Then
            ' Set the first date (Maturity date or First Accrual date) depends on the direction of generation
            If int_Direction = -1 Then _
            Call lngLst_Follower.Add(Date_ApplyBDC(fld_Params.MatDate, "FOLL", cal_pmt.HolDates, cal_pmt.Weekends))
            If int_Direction = 1 Then _
            Call lngLst_Follower.Add(Date_ApplyBDC(fld_Params.FirstAccDate, "FOLL", cal_pmt.HolDates, cal_pmt.Weekends))

        Else
            ' Set the period start date based on the end date generated for previous period
            Call lngLst_Follower.Add(lngLst_Driver(int_ctr - 1))
        End If

        ' Generate end date of the period
        If UCase(str_StubType) = "UP FRONT" And int_ctr <> int_NumPeriods Then
            If int_Sign <> int_Direction Then
                lng_tempDate = Date_NextCoupon(fld_Params.MatDate, str_PeriodLength, cal_pmt, int_ctr * int_Direction - int_Sign, fld_Params.EOM, "UNADJ")
                lng_tempDate = Date_NextCoupon(lng_tempDate, str_PeriodLength, cal_pmt, int_Sign, fld_Params.EOM, fld_Params.BDC)
                Call lngLst_Driver.Add(lng_tempDate)
            Else
                Call lngLst_Driver.Add(Date_NextCoupon(fld_Params.MatDate, str_PeriodLength, cal_pmt, _
                int_ctr * int_Sign, fld_Params.EOM, fld_Params.BDC))
            End If
        End If

        If UCase(str_StubType) = "ARREARS" And int_ctr <> int_NumPeriods Then
            If int_Sign <> int_Direction Then
                lng_tempDate = Date_NextCoupon(fld_Params.FirstAccDate, str_PeriodLength, cal_pmt, int_ctr * int_Direction - int_Sign, fld_Params.EOM, "UNADJ")
                lng_tempDate = Date_NextCoupon(lng_tempDate, str_PeriodLength, cal_pmt, int_Sign, fld_Params.EOM, fld_Params.BDC)
                Call lngLst_Driver.Add(lng_tempDate)
            Else
                Call lngLst_Driver.Add(Date_NextCoupon(fld_Params.FirstAccDate, str_PeriodLength, cal_pmt, _
                int_ctr * int_Sign, fld_Params.EOM, fld_Params.BDC))
            End If
        End If

        If Left(UCase(str_StubType), 8) = "BOTH END" And int_ctr = 1 Then _
        Call lngLst_Driver.Add(Date_ApplyBDC(fld_Params.RollDate, fld_Params.BDC, cal_pmt.HolDates, cal_pmt.Weekends))

        If Left(UCase(str_StubType), 8) = "BOTH END" And int_ctr <> 1 And int_ctr <> int_NumPeriods Then
            If int_Sign <> int_Direction Then
                lng_tempDate = Date_NextCoupon(fld_Params.RollDate, str_PeriodLength, cal_pmt, (int_ctr - 1) * int_Direction - int_Sign, fld_Params.EOM, "UNADJ")
                lng_tempDate = Date_NextCoupon(lng_tempDate, str_PeriodLength, cal_pmt, int_Sign, fld_Params.EOM, fld_Params.BDC)
                Call lngLst_Driver.Add(lng_tempDate)
            Else
                Call lngLst_Driver.Add(Date_NextCoupon(fld_Params.RollDate, str_PeriodLength, cal_pmt, _
                (int_ctr - 1) * int_Sign, fld_Params.EOM, fld_Params.BDC))
            End If
        End If

        ' Set the end date of the last period
        If Right(UCase(str_StubType), 9) = "(FORWARD)" And int_ctr = int_NumPeriods Then _
        Call lngLst_Driver.Add(Date_ApplyBDC(fld_Params.MatDate, "FOLL", cal_pmt.HolDates, cal_pmt.Weekends))

        If Right(UCase(str_StubType), 10) = "(BACKWARD)" And int_ctr = int_NumPeriods Then _
        Call lngLst_Driver.Add(Date_ApplyBDC(fld_Params.FirstAccDate, "FOLL", cal_pmt.HolDates, cal_pmt.Weekends))

        If UCase(str_StubType) = "ARREARS" And int_ctr = int_NumPeriods Then _
        Call lngLst_Driver.Add(Date_ApplyBDC(fld_Params.MatDate, "FOLL", cal_pmt.HolDates, cal_pmt.Weekends))

        If UCase(str_StubType) = "UP FRONT" And int_ctr = int_NumPeriods Then _
        Call lngLst_Driver.Add(Date_ApplyBDC(fld_Params.FirstAccDate, "FOLL", cal_pmt.HolDates, cal_pmt.Weekends))

    Next int_ctr

    ' Generation limit point if defined, will be the last generated date
    If fld_Params.GenerationLimitPoint <> 0 Then
        Call lngLst_Driver.Remove(int_NumPeriods)
        Call lngLst_Driver.Add(fld_Params.GenerationLimitPoint)
    End If

    ' Store period start and end dates depending on the generation method
    If int_Direction = 1 Then
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

'Mandy: add this for the 28/30 convention
Private Sub AjustPeriodDates_2830()
    '## Adjust date according to special convention where EOM dates are either 28 (for Feb) or 30
    '## The following code assume Roll convention is Indifferent
    Dim lng_newDate As Long

    Dim int_ctr As Integer
    For int_ctr = 2 To lngLst_PeriodStart.count

        If Month(lngLst_PeriodStart(int_ctr)) <> 2 And Day(lngLst_PeriodStart(int_ctr)) = 31 Then
            lng_newDate = DateSerial(year(lngLst_PeriodStart(int_ctr)), Month(lngLst_PeriodStart(int_ctr)), 30)

            Call lngLst_PeriodStart.Remove(int_ctr)
            If int_ctr > lngLst_PeriodStart.count Then _
                Call lngLst_PeriodStart.Add(lng_newDate) Else _
                Call lngLst_PeriodStart.Add(lng_newDate, , int_ctr)
            Call lngLst_PeriodEnd.Remove(int_ctr - 1)
            If int_ctr - 1 > lngLst_PeriodEnd.count Then _
                Call lngLst_PeriodEnd.Add(lng_newDate) Else _
                Call lngLst_PeriodEnd.Add(lng_newDate, , int_ctr - 1)

        ElseIf Month(lngLst_PeriodStart(int_ctr)) = 2 And Day(lngLst_PeriodStart(int_ctr)) = 29 Then
            lng_newDate = DateSerial(year(lngLst_PeriodStart(int_ctr)), Month(lngLst_PeriodStart(int_ctr)), 28)

            Call lngLst_PeriodStart.Remove(int_ctr)
            If int_ctr > lngLst_PeriodStart.count Then _
                Call lngLst_PeriodStart.Add(lng_newDate) Else _
                Call lngLst_PeriodStart.Add(lng_newDate, , int_ctr)
            Call lngLst_PeriodEnd.Remove(int_ctr - 1)
            If int_ctr - 1 > lngLst_PeriodEnd.count Then _
                Call lngLst_PeriodEnd.Add(lng_newDate) Else _
                Call lngLst_PeriodEnd.Add(lng_newDate, , int_ctr - 1)
        End If

    Next int_ctr



End Sub

Private Sub FillEstDates()
    ' ## Generate start and end dates for rate estimation
    Set lngLst_EstStart = New Collection
    Set lngLst_EstEnd = New Collection

    Dim int_ctr As Integer
    For int_ctr = 1 To lngLst_PeriodStart.count
        Call lngLst_EstStart.Add(date_workday(lngLst_PeriodStart(int_ctr) - 1, 1, cal_est.HolDates, cal_est.Weekends))

        If fld_Params.FloatEst = True Then
            Call lngLst_EstEnd.Add(Date_NextCoupon(lngLst_EstStart(int_ctr), str_EstPeriodLength, cal_est, 1, _
            fld_Params.EOM, fld_Params.BDC))
        Else
            Call lngLst_EstEnd.Add(lngLst_PeriodEnd(int_ctr))
        End If
    Next int_ctr
End Sub

Private Sub FillPmtDates()
    ' ## Return the list of cash flow payment dates
    ' ## Also classifies each payment date as belonging to MV or not
    ' Determine generation direction by Stub Type
    Dim bln_IsForwardDirection As Boolean
    If UCase(fld_Params.StubType) = "UP FRONT" Or Right(UCase(fld_Params.StubType), 10) = "(BACKWARD)" Then bln_IsForwardDirection = False Else bln_IsForwardDirection = True

    Dim int_NumPeriods As Integer: int_NumPeriods = lngLst_PeriodEnd.count

    Set lngLst_PmtDates = New Collection

    Dim int_ctr As Integer
    For int_ctr = 1 To int_NumPeriods
'        If UCase(fld_Params.PaymentSchType) = "ZERO COUPON" Then
'        ' Set all payment dates is the same as maturity date
'            Call lngLst_PmtDates.Add(lngLst_PeriodEnd(int_NumPeriods))
'        Else
            ' Regular cases where all payment dates equal to the end date and apply BDC when necessary
            If fld_Params.IsLongCpn = False And int_ctr Mod int_CalcsPerPmt = 0 Then _
            Call lngLst_PmtDates.Add(Date_ApplyBDC(lngLst_PeriodEnd(int_ctr), fld_Params.PaymentSchType, cal_pmt.HolDates, cal_pmt.Weekends))

            If fld_Params.IsLongCpn = True Then
            ' Changing payment date of long coupon here
                If bln_IsForwardDirection = True Then
                ' Set Payment date of the second last period to be the same as that of the last period
                    If int_ctr <> (lngLst_PeriodEnd.count - 1) Then _
                    Call lngLst_PmtDates.Add(Date_ApplyBDC(lngLst_PeriodEnd(int_ctr), fld_Params.PaymentSchType, cal_pmt.HolDates, cal_pmt.Weekends))
                    If int_ctr = (lngLst_PeriodEnd.count - 1) Then _
                    Call lngLst_PmtDates.Add(Date_ApplyBDC(lngLst_PeriodEnd(int_ctr + 1), fld_Params.PaymentSchType, cal_pmt.HolDates, cal_pmt.Weekends))
                Else
                ' Set first payment date to be the same as that of the second period
                    If int_ctr <> 1 Then _
                    Call lngLst_PmtDates.Add(Date_ApplyBDC(lngLst_PeriodEnd(int_ctr), fld_Params.PaymentSchType, cal_pmt.HolDates, cal_pmt.Weekends))
                    If int_ctr = 1 Then _
                    Call lngLst_PmtDates.Add(Date_ApplyBDC(lngLst_PeriodEnd(int_ctr + 1), fld_Params.PaymentSchType, cal_pmt.HolDates, cal_pmt.Weekends))

                End If
            End If
'        End If

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
        ElseIf lngLst_PmtDates(int_ctr) > fld_Params.TradeDate Then
            Call intLst_IsMV.Add(0)
            Call intLst_IsCash.Add(1)
        Else
            Call intLst_IsMV.Add(0)
            Call intLst_IsCash.Add(0)
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
        Dim int_ctr As Integer

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
                        ' Modify the date
                        lng_ActiveMod = DateValue(str_ActiveMod)
                        Call lngLst_PeriodStart.Remove(int_ctr)
                        Call lngLst_PeriodStart.Add(lng_ActiveMod, , int_ctr)

                        If Not lngLst_EstStart Is Nothing Then
                            Call lngLst_EstStart.Remove(int_ctr)
                            Call lngLst_EstStart.Add(lng_ActiveMod, , int_ctr)
                            'Call dblLst_EstPeriodDurations.Remove(int_Ctr)
                            'Call dblLst_EstPeriodDurations.Add(Calc_YearFrac(lngLst_EstStart(int_Ctr), lngLst_EstEnd(int_Ctr), "ACT/365"), , int_Ctr)
                        End If
                    End If

                    Exit For
                End If
            Next int_ctr
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
                fld_Params.Daycount, fld_Params.PmtFreq, bln_ActActFwdGeneration, fld_Params.RateOrMargin))
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

            ' Estimate floating rates for future fixings, past fixings will be filled in later
            If lng_ActiveEstStart >= lng_BuildDate And lng_ActiveEstStart >= fld_Params.ValueDate Then
                dbl_ActiveRate = irc_Est.Lookup_Rate(lng_ActiveEstStart, lngLst_EstEnd(int_ctr), "ACT/365", _
                    , fld_Params.index, False, , , bln_ActActFwdGeneration)

                'dbl_ActiveRate = irc_Est.Lookup_Rate(lng_ActiveEstStart, lngLst_EstEnd(int_Ctr), fld_Params.Daycount, _
                    , fld_Params.Index, False, , , bln_ActActFwdGeneration)
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
            If fld_Params.IsRoundFlow = False Then Call dblLst_IntFlows.Add(dbl_ActiveAccum) _
            Else Call dblLst_IntFlows.Add(Round(dbl_ActiveAccum, 2))

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
                    If int_ctr > dblLst_Rates.count Then _
                    Call dblLst_Rates.Add(dbl_ActiveFixing) Else _
                    Call dblLst_Rates.Add(dbl_ActiveFixing, , int_ctr)

                    ' Period is no longer relevant for estimation curve DV01
                    If bln_IsFixed = False Then
                        Call dblLst_EstPeriodDurations.Remove(int_ctr)
                        If int_ctr > dblLst_EstPeriodDurations.count Then _
                        Call dblLst_EstPeriodDurations.Add(dbl_ActiveFixing) Else _
                        Call dblLst_EstPeriodDurations.Add(dbl_ActiveFixing, , int_ctr)
                    End If

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

Private Function CalcValue(str_type As String, Optional dblLst_CustomDFs As Variant, Optional dbl_CustomStartDF As Double = -1) As Double
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


Private Function CalcYield() As Double

'Solve Yield

Dim dbl_Target As Double
Dim dic_SecantParams As Dictionary: Set dic_SecantParams = New Dictionary
Dim dbl_AI As Double

Call dic_SecantParams.Add("irl_leg", Me)
dbl_Target = CalcValue("MV")

Dim dic_SecantOutputs As Dictionary: Set dic_SecantOutputs = New Dictionary
CalcYield = Solve_Secant(ThisWorkbook, "SolverFunc_Yield", dic_SecantParams, fld_Params.Yield, 0.01, dbl_Target, 0.00000000001 * fld_Params.Notional, 100, -100, dic_SecantOutputs)

fld_Params.ModifiedDuration = fld_Params.ModifiedDuration / dbl_Target

If fld_Params.FixingCurve <> "-" Then
    dbl_periodicity = dbl_periodicity * 360 / 365.25
    fld_Params.MacaulayDuration = fld_Params.ModifiedDuration * (1 + (fld_Params.Fix_AI / 100 + CalcYield) / dbl_periodicity)
Else
    fld_Params.Fix_AI = 0
    fld_Params.MacaulayDuration = fld_Params.MacaulayDuration / dbl_Target
End If

fld_Params.Duration = fld_Params.ModifiedDuration * (1 + (fld_Params.Fix_AI / 100 + CalcYield) / dbl_periodicity)

End Function

Public Function DirtyPrice_AIBDandAIBDEffective() As Double

If fld_Params.StubType = "Up Front" And fld_Params.YieldCalc <> "AIBD AM/Step" Then
    DirtyPrice_AIBDandAIBDEffective = DirtyPrice_AIBDandAIBDEffectiveUpFront
Else
    DirtyPrice_AIBDandAIBDEffective = DirtyPrice_AIBDandAIBDEffectiveOthers
End If

End Function

Public Function DirtyPrice_AIBDandAIBDEffectiveOthers() As Double

Dim dbl_CF As Double, dbl_Price As Double, dbl_yield As Double, dbl_ExpTerm As Double, dbl_ModifiedDurations As Double
Dim int_ctr As Integer
Dim str_DayCountConv As String
Dim int_ModifiedCtr As Integer
Dim dbl_MacaulayDurations As Double

If fld_Params.DaycountConv = "-" Then
    str_DayCountConv = fld_Params.Daycount
Else
    str_DayCountConv = fld_Params.DaycountConv
End If

dbl_periodicity = 12 / calc_nummonths(fld_Params.PmtFreq)

dbl_Price = 0: dbl_ModifiedDurations = 0
dbl_yield = fld_Params.Yield

For int_ctr = int_CtrStart To lngLst_PeriodEnd.count
    int_ModifiedCtr = PmtDateChk(int_ctr)
    dbl_ExpTerm = ExpTerm(int_ModifiedCtr, str_DayCountConv, dbl_periodicity)
    dbl_CF = dblLst_IntFlows(int_ctr) + dblLst_PrnFlows(int_ctr)
    dbl_Price = dbl_Price + dbl_CF / (1 + dbl_yield / dbl_periodicity) ^ dbl_ExpTerm
    dbl_MacaulayDurations = dbl_MacaulayDurations + (dbl_CF * dbl_ExpTerm) / ((1 + dbl_yield / dbl_periodicity) ^ dbl_ExpTerm)
    dbl_ModifiedDurations = dbl_ModifiedDurations + (dbl_CF * dbl_ExpTerm) / ((1 + dbl_yield / dbl_periodicity) ^ dbl_ExpTerm) / (1 + dbl_yield / dbl_periodicity)
Next int_ctr

fld_Params.MacaulayDuration = dbl_MacaulayDurations / dbl_periodicity
fld_Params.ModifiedDuration = dbl_ModifiedDurations / dbl_periodicity
DirtyPrice_AIBDandAIBDEffectiveOthers = dbl_Price

End Function

Public Function DirtyPrice_AIBDandAIBDEffectiveUpFront() As Double

'Compute Dirty Price for Bond for AIBD or AIBD Effective or ACT/ACT specific convention
' if it is a stub both end and cash flow is no difference from the stub upfront, adjustment on the pricing worksheet is required as the code will not be able to detect it.

Dim dbl_CF As Double, dbl_Price As Double, dbl_yield As Double, dbl_BrokenPeriod As Double, dbl_ModifiedDurations As Double
Dim int_ctr As Integer, int_ModifiedCtr As Integer
Dim str_DayCountConv As String

Dim dbl_MacaulayDurations As Double

If fld_Params.DaycountConv = "-" Then
    str_DayCountConv = fld_Params.Daycount
Else
    str_DayCountConv = fld_Params.DaycountConv
End If

dbl_periodicity = 12 / calc_nummonths(fld_Params.PmtFreq)

dbl_Price = 0: dbl_ModifiedDurations = 0
dbl_yield = fld_Params.Yield
dbl_BrokenPeriod = calc_yearfrac(fld_Params.ValueDate, lngLst_PmtDates(int_CtrStart), str_DayCountConv) * dbl_periodicity

For int_ctr = int_CtrStart To lngLst_PeriodEnd.count
    int_ModifiedCtr = PmtDateChk(int_ctr)
    If fld_Params.YieldCalc = "AIBD" And int_ctr <> 1 Then
        dbl_CF = fld_Params.RateOrMargin / 100 / dbl_periodicity * fld_Params.Notional + dblLst_PrnFlows(int_ctr)
    Else
        dbl_CF = dblLst_IntFlows(int_ctr) + dblLst_PrnFlows(int_ctr)
    End If
    dbl_Price = dbl_Price + dbl_CF / (1 + dbl_yield / dbl_periodicity) ^ (dbl_BrokenPeriod + (int_ModifiedCtr - int_CtrStart))
    dbl_MacaulayDurations = dbl_MacaulayDurations + (dbl_CF * (dbl_BrokenPeriod + (int_ModifiedCtr - int_CtrStart))) / ((1 + dbl_yield / dbl_periodicity) ^ (dbl_BrokenPeriod + (int_ModifiedCtr - int_CtrStart)))
    dbl_ModifiedDurations = dbl_ModifiedDurations + (dbl_CF * (dbl_BrokenPeriod + (int_ModifiedCtr - int_CtrStart))) / ((1 + dbl_yield / dbl_periodicity) ^ (dbl_BrokenPeriod + (int_ModifiedCtr - int_CtrStart))) / (1 + dbl_yield / dbl_periodicity)
Next int_ctr

fld_Params.MacaulayDuration = dbl_MacaulayDurations / dbl_periodicity
fld_Params.ModifiedDuration = dbl_ModifiedDurations / dbl_periodicity
DirtyPrice_AIBDandAIBDEffectiveUpFront = dbl_Price

End Function

Public Function DirtyPrice_SpecificConvention() As Double

Dim dbl_Price As Double, dbl_principal As Double, dbl_yield As Double, dbl_AnniversaryDate As Double, dbl_CF As Double, dbl_dt As Double, dbl_ModifiedDurations As Double
Dim int_ctr As Integer, int_ModifiedCtr As Integer
Dim str_DayCountConv As String
Dim str_RateComputingMode As String

Dim dbl_MacaulayDurations As Double

' The default Day Count Convention is ACT/365
If fld_Params.DaycountConv = "-" Then
    str_DayCountConv = fld_Params.Daycount
Else
    str_DayCountConv = fld_Params.DaycountConv
End If

' The default Rate Computing Mode is Linear
If fld_Params.RateComputingMode = "-" Then
    str_RateComputingMode = "Linear"
Else
    str_RateComputingMode = fld_Params.RateComputingMode
End If

' assign the periodicity setting
If fld_Params.Periodicity = "Coupon periodicity" Then
    dbl_periodicity = 12 / calc_nummonths(fld_Params.PmtFreq)
Else
    dbl_periodicity = 1
End If

dbl_principal = 0: dbl_Price = 0
dbl_yield = fld_Params.Yield
dbl_ModifiedDurations = 0: dbl_CF = 0: dbl_dt = 0

For int_ctr = int_CtrStart To lngLst_PeriodEnd.count
    int_ModifiedCtr = PmtDateChk(int_ctr)

    ' The Yield Schedule is assumed to be Anniversary Schedule
'    If int_ModifiedCtr = lngLst_PeriodEnd.count Then
'         dbl_AnniversaryDate = lngLst_PmtDates(int_ModifiedCtr)
'    Else
'        dbl_AnniversaryDate = WorksheetFunction.EDate(lngLst_PmtDates(1), Calc_NumMonths(fld_Params.PmtFreq) * (int_ModifiedCtr - 1))
'    End If
'
'    dbl_dt = Calc_yearfrac(fld_Params.ValueDate, dbl_AnniversaryDate, str_DayCountConv)

    If fld_Params.YieldSchedule = "Anniversary Schedule" Then
        If int_ModifiedCtr = lngLst_PeriodEnd.count Then
            dbl_AnniversaryDate = lngLst_PmtDates(int_ModifiedCtr)
        Else
            dbl_AnniversaryDate = WorksheetFunction.EDate(lngLst_PmtDates(1), calc_nummonths(fld_Params.PmtFreq) * (int_ModifiedCtr - 1))
        End If
        dbl_dt = calc_yearfrac(fld_Params.ValueDate, dbl_AnniversaryDate, str_DayCountConv)
    Else
        dbl_dt = calc_yearfrac(fld_Params.ValueDate, lngLst_PmtDates(int_ModifiedCtr), str_DayCountConv)
    End If


    Select Case str_RateComputingMode
        Case "Linear": dbl_CF = (dblLst_IntFlows(int_ctr) + dblLst_PrnFlows(int_ctr)) / (1 + dbl_yield * dbl_dt)
                    dbl_Price = dbl_Price + dbl_CF
                    dbl_MacaulayDurations = dbl_MacaulayDurations + dbl_CF * dbl_dt
                    dbl_ModifiedDurations = dbl_ModifiedDurations + dbl_CF * dbl_dt / (1 + dbl_yield * dbl_dt)
        Case "Yield": dbl_CF = (dblLst_IntFlows(int_ctr) + dblLst_PrnFlows(int_ctr)) / (1 + dbl_yield / dbl_periodicity) ^ (dbl_periodicity * dbl_dt)
                    dbl_Price = dbl_Price + dbl_CF
                    dbl_MacaulayDurations = dbl_MacaulayDurations + dbl_CF * dbl_dt
                    dbl_ModifiedDurations = dbl_ModifiedDurations + dbl_CF * dbl_dt / (1 + dbl_yield / dbl_periodicity)
        Case "Discount": dbl_CF = (1 - dbl_yield * dbl_dt) * (dblLst_IntFlows(int_ctr) + dblLst_PrnFlows(int_ctr))
                    dbl_Price = dbl_Price + dbl_CF
                    dbl_MacaulayDurations = dbl_MacaulayDurations + dbl_CF * dbl_dt
                    dbl_ModifiedDurations = dbl_ModifiedDurations + dbl_dt * (dblLst_IntFlows(int_ctr) + dblLst_PrnFlows(int_ctr))
    End Select
Next int_ctr
fld_Params.MacaulayDuration = dbl_MacaulayDurations
fld_Params.ModifiedDuration = dbl_ModifiedDurations
DirtyPrice_SpecificConvention = dbl_Price

End Function

Public Function ExpTerm(int_ctr As Integer, str_DayCountConv As String, dbl_periodicity As Double) As Double

Dim AdjustedDate, AdjustedDate1 As Double
Dim index As Integer
AdjustedDate = 0: AdjustedDate1 = 0

Select Case str_DayCountConv
    Case "ACT/365":
        If fld_Params.StubType = "Arrears" And int_ctr <> lngLst_PeriodEnd.count Then
            ExpTerm = calc_yearfrac(fld_Params.ValueDate, lngLst_PmtDates(int_CtrStart), str_DayCountConv) * dbl_periodicity + int_ctr - int_CtrStart
        Else
            index = DateAdjIndex(int_ctr)
            AdjustedDate = WorksheetFunction.EDate(lngLst_PmtDates(int_ctr), -calc_nummonths(fld_Params.PmtFreq) * index)
            ExpTerm = (AdjustedDate - fld_Params.ValueDate) / 365 * dbl_periodicity + index
        End If
    Case "30/360":
        index = DateAdjIndex(int_ctr)
        AdjustedDate = WorksheetFunction.EDate(lngLst_PmtDates(int_ctr), -calc_nummonths(fld_Params.PmtFreq) * index)
        ExpTerm = ((Month(AdjustedDate) - Month(fld_Params.ValueDate)) * 30 + Day(AdjustedDate) - Day(fld_Params.ValueDate) + (year(AdjustedDate) - year(fld_Params.ValueDate)) * 360) / 360 * dbl_periodicity + index
    Case "ACT/ACT":
        index = DateAdjIndex(int_ctr)
        AdjustedDate = WorksheetFunction.EDate(lngLst_PmtDates(int_ctr), -calc_nummonths(fld_Params.PmtFreq) * index)
        AdjustedDate1 = WorksheetFunction.EDate(AdjustedDate, -calc_nummonths(fld_Params.PmtFreq))
        ExpTerm = (AdjustedDate - fld_Params.ValueDate) / (AdjustedDate - AdjustedDate1) + index
    Case "ACT/ACT XTE":
        index = DateAdjIndex(int_ctr)
        AdjustedDate = WorksheetFunction.EDate(lngLst_PmtDates(int_ctr), -calc_nummonths(fld_Params.PmtFreq) * index)
        ExpTerm = calc_yearfrac(fld_Params.ValueDate, AdjustedDate, "ACT/ACT XTE") * dbl_periodicity + index
End Select

End Function

Public Function PmtDateChk(int_ctr As Integer) As Integer

''If int_Ctr = int_CtrStart Or int_Ctr = lngLst_PeriodEnd.count Then
'If int_Ctr = lngLst_PeriodEnd.count Then
'    PmtDateChk = int_Ctr
'ElseIf lngLst_PmtDates(int_Ctr) = lngLst_PmtDates(int_Ctr + 1) Then
'    PmtDateChk = int_Ctr + 1
'ElseIf int_Ctr = int_CtrStart Then
'    PmtDateChk = int_Ctr
'ElseIf lngLst_PmtDates(int_Ctr) = lngLst_PmtDates(int_Ctr - 1) Then
'    PmtDateChk = int_Ctr - 1
'Else
'    PmtDateChk = int_Ctr
'End If

'If dblLst_IntFlows(int_Ctr) = 0 And int_Ctr <> int_CtrStart Then
'     PmtDateChk = int_Ctr - 1
If int_ctr = lngLst_PeriodEnd.count Then
    PmtDateChk = int_ctr
ElseIf lngLst_PmtDates(int_ctr) = lngLst_PmtDates(int_ctr + 1) Then
    PmtDateChk = int_ctr + 1
ElseIf int_ctr = int_CtrStart Then
    PmtDateChk = int_ctr
ElseIf lngLst_PmtDates(int_ctr) = lngLst_PmtDates(int_ctr - 1) Then
    PmtDateChk = int_ctr - 1
Else
    PmtDateChk = int_ctr
End If



End Function

Public Function DirtyPrice_DiscountMargin() As Double

'Compute Dirty Price for Bond for AIBD or AIBD Effective or ACT/ACT specific convention

Dim dbl_CF As Double, dbl_Price As Double, dbl_yield As Double, dbl_BrokenPeriod As Double, dbl_ModifiedDurations As Double
Dim int_ctr As Integer
Dim str_DayCountConv As String
Dim dbl_AI_DM As Double, dbl_IT_DM As Double, dbl_AI As Double, dbl_IT As Double, lng_IndexMat As Long
Dim dbl_DF_broken As Double, dbl_DF As Double, dbl_adjPeriod, dbl_disCF As Double

If fld_Params.DaycountConv = "-" Then
'    str_DayCountConv = "ACT/365"
    str_DayCountConv = fld_Params.Daycount
Else
    str_DayCountConv = fld_Params.DaycountConv
End If

dbl_periodicity = 12 / calc_nummonths(fld_Params.PmtFreq)
dbl_adjPeriod = dbl_periodicity * 360 / 365.25

dbl_Price = 0: dbl_ModifiedDurations = 0
dbl_yield = fld_Params.Yield
dbl_BrokenPeriod = calc_yearfrac(fld_Params.ValueDate, lngLst_PmtDates(int_CtrStart), str_DayCountConv, fld_Params.PmtFreq)

'check this
If IsNumeric(fld_Params.Fix_AI) And IsNumeric(fld_Params.Fix_IT) Then

    dbl_AI = fld_Params.Fix_AI / 100
    dbl_IT = fld_Params.Fix_IT / 100
Else
    lng_IndexMat = Date_NextCoupon(fld_Params.ValueDate, str_EstPeriodLength, cal_est, 1, _
            fld_Params.EOM, fld_Params.BDC)
    dbl_AI = cyReadIRCurve(fld_Params.FixingCurve, fld_Params.ValueDate, lng_IndexMat, "ZERO") / 100
    dbl_IT = cyReadIRCurve(fld_Params.FixingCurve, fld_Params.ValueDate, lngLst_PmtDates(int_CtrStart), "ZERO") / 100
    fld_Params.Fix_AI = dbl_AI * 100
    fld_Params.Fix_IT = dbl_IT * 100
End If

dbl_AI_DM = dbl_AI + dbl_yield
dbl_IT_DM = dbl_IT + dbl_yield


dbl_DF_broken = 1 / (1 + dbl_IT_DM * dbl_BrokenPeriod)
dbl_DF = 1 / (1 + dbl_AI_DM / dbl_adjPeriod)


For int_ctr = int_CtrStart To lngLst_PeriodEnd.count
    If int_ctr <> int_CtrStart Then
        dbl_CF = (fld_Params.RateOrMargin + dbl_AI * 100) / 100 / dbl_adjPeriod * fld_Params.Notional
    Else
        dbl_CF = dblLst_IntFlows(int_ctr)
    End If
    dbl_disCF = dbl_CF * dbl_DF_broken * dbl_DF ^ (int_ctr - int_CtrStart)
    dbl_Price = dbl_Price + dbl_disCF
    dbl_ModifiedDurations = dbl_ModifiedDurations + dbl_disCF * (dbl_DF_broken * dbl_BrokenPeriod + dbl_DF * (int_ctr - int_CtrStart) / dbl_adjPeriod)
Next int_ctr

dbl_disCF = fld_Params.Notional * dbl_DF_broken * dbl_DF ^ (lngLst_PeriodEnd.count - int_CtrStart)
dbl_Price = dbl_Price + dbl_disCF
dbl_ModifiedDurations = dbl_ModifiedDurations + dbl_disCF * (dbl_DF_broken * dbl_BrokenPeriod + dbl_DF * (lngLst_PeriodEnd.count - int_CtrStart) / dbl_adjPeriod)

fld_Params.ModifiedDuration = dbl_ModifiedDurations
DirtyPrice_DiscountMargin = dbl_Price

End Function

Public Function DirtyPrice() As Double

int_CtrStart = CtrStart

If fld_Params.FixingCurve <> "-" Then
    DirtyPrice = DirtyPrice_DiscountMargin
ElseIf fld_Params.YieldCalc = "AIBD" Or fld_Params.YieldCalc = "AIBD (Effective flows)" Or fld_Params.YieldCalc = "AIBD AM/Step" Then 'Or fld_Params.DaycountConv = "ACT/ACT" Then
' if the Yield Convention is AIBD, AIBD Effective Flow, YLD ACT/ACT, LIN ACT/ACT and DIS ACT/ACT, it will flow to seperate function called "SWDirtyPrice_AIBDandAIBDEffective" to calculate the Dirty Price
    DirtyPrice = DirtyPrice_AIBDandAIBDEffective
Else
    DirtyPrice = DirtyPrice_SpecificConvention
End If

End Function

Public Function CtrStart() As Double

Dim int_ctr  As Long

int_CtrStart = 1
For int_ctr = 1 To lngLst_PeriodEnd.count
    If lngLst_PmtDates(int_ctr) < fld_Params.ValueDate Then
        int_CtrStart = int_CtrStart + 1
    Else
        Exit For
    End If
Next int_ctr

CtrStart = int_CtrStart

End Function

Public Function DateAdjIndex(int_ctr) As Double

Dim count As Integer, StartDate As Long, EndDate As Long
count = 0

With WorksheetFunction
Do
     StartDate = .EDate(lngLst_PmtDates(int_ctr), -calc_nummonths(fld_Params.PmtFreq) * count)
     EndDate = .EDate(lngLst_PmtDates(int_ctr), -calc_nummonths(fld_Params.PmtFreq) * (count + 1))
     count = count + 1
     If count = 50 Then
        Exit Do
    End If
Loop Until StartDate > fld_Params.ValueDate And fld_Params.ValueDate > EndDate
DateAdjIndex = count - 1
End With
End Function

Public Function SetYield(dbl_yield As Double) As Double
    fld_Params.Yield = dbl_yield
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
                .Offset(int_ctr + int_ActiveRow, int_ActiveCol).Value = dblLst_Rates(int_ctr)
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
                ElseIf intLst_IsCash(int_ActiveIndex) = 1 Then
                    .Offset(int_ActiveRow, int_ActiveCol).Value = "CASH"
                Else
                    .Offset(int_ActiveRow, int_ActiveCol).Value = "EXCLUDED"
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
                        & dic_Addresses("SpotDate") & "," & dic_TempAddresses("PmtDate") & ",""DF"",,False," & dbl_ZSpread & ")"
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