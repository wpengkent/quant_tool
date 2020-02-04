Option Explicit

Public Enum CurveState_IRC
    original = 1
    Final
    Zero_Up1BP
    Zero_Down1BP
End Enum

' ## MEMBER DATA
Private Const bln_ShowMessages As Boolean = False
Private Const str_Daycount_Zero As String = "ACT/365"
Private wks_Location As Worksheet

' Stored ranges
Private rng_InstTypes As Range, rng_MXTerms As Range, rng_MXSortTerms As Range, rng_FutMats As Range
Private rng_OrigPars As Range, rng_DayCounts As Range, rng_BDCs As Range, rng_SpotDays As Range, rng_FlowsFreq As Range
Private rng_PmtCalendars As Range, rng_EstCalendars As Range, rng_StartDates As Range, rng_MatDates As Range
Private rng_OrigZeros As Range, rng_Days_TopLeft As Range, rng_RelPar_TopLeft As Range, rng_AbsPar_TopLeft As Range
Private rng_RelZero_TopLeft As Range, rng_AbsZero_TopLeft As Range, rng_FinalPars As Range, rng_FinalZeros As Range

' Dynamic variables
Private lngArr_PillarDates() As Long, dblArr_LiveZeros() As Double
Private csh_ZeroShifts_Rel As CurveDaysShift, csh_ZeroShifts_Abs As CurveDaysShift
Private csh_ParShifts_Rel As CurveDaysShift, csh_ParShifts_Abs As CurveDaysShift
Private csh_ZeroSens_Abs As CurveDaysShift
Private dic_CurveStates As Dictionary, enu_LiveState As CurveState_IRC, int_LiveState_Pillar As Integer

' Static values
Private dic_GlobalStaticInfo As Dictionary, iqs_Queries As IRQuerySet, map_Rules As MappingRules
Private cas_Calendars As CalendarSet, cfg_Settings As ConfigSheet
Private lng_BuildDate As Long, str_InterpConv As String, str_BaseCurve As String, str_CurveName As String
Private bln_IsAdjMatForEst As Boolean, str_Priorities As String, bln_InBlocks As Boolean, bln_AcceptNil As Boolean
Private bln_EOMRule As Boolean, bln_EstFloat As Boolean, int_NumPoints As Integer


' ## INITIALIZATION
Public Sub Initialize(wks_Input As Worksheet, bln_DataExists As Boolean, Optional dic_StaticInfoInput As Dictionary = Nothing)
    Set wks_Location = wks_Input

    ' Store static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    Set iqs_Queries = dic_GlobalStaticInfo(StaticInfoType.IRQuerySet)
    Set map_Rules = dic_GlobalStaticInfo(StaticInfoType.MappingRules)
    Set cas_Calendars = dic_GlobalStaticInfo(StaticInfoType.CalendarSet)
    Set cfg_Settings = dic_GlobalStaticInfo(StaticInfoType.ConfigSheet)

    ' Prepare shift objects
    Set csh_ZeroShifts_Rel = New CurveDaysShift
    Set csh_ZeroShifts_Abs = New CurveDaysShift
    Set csh_ParShifts_Rel = New CurveDaysShift
    Set csh_ParShifts_Abs = New CurveDaysShift
    Call csh_ZeroShifts_Rel.Initialize(ShockType.Relative)
    Call csh_ZeroShifts_Abs.Initialize(ShockType.Absolute)
    Call csh_ParShifts_Rel.Initialize(ShockType.Relative)
    Call csh_ParShifts_Abs.Initialize(ShockType.Absolute)

    ' DataExists = True when reading a rate
    If bln_DataExists = True Then
        Call StoreStaticValues
        Call Me.AssignRanges
    End If
End Sub


' ## PROPERTIES
Public Property Get BuildDate() As Long
    BuildDate = lng_BuildDate
End Property

Public Property Get CurveName() As String
    CurveName = str_CurveName
End Property

Public Property Get NumPoints() As Integer
    NumPoints = int_NumPoints
End Property

Public Property Get NumShifts() As Integer
    NumShifts = csh_ParShifts_Abs.NumShifts + csh_ParShifts_Rel.NumShifts + csh_ZeroShifts_Abs.NumShifts + csh_ZeroShifts_Rel.NumShifts
End Property

Public Property Get CurveState() As CurveState_IRC
    CurveState = enu_LiveState
End Property

Public Property Get SensPillar() As Integer
    SensPillar = int_LiveState_Pillar
End Property

Public Property Get ConfigSheet() As ConfigSheet
    Set ConfigSheet = cfg_Settings
End Property


' ## METHODS - LOOKUP
Public Function Lookup_Rate(lng_startdate As Long, lng_EndDate As Long, str_RateType As String, _
    Optional str_InterpPillars As String = "", _
    Optional str_CouponFreq As String = "6M", Optional bln_AllowBackwards = True, _
    Optional bln_GetNotionalFactor As Boolean = False, Optional dbl_ZSpread As Double = 0, _
    Optional bln_IsFwdGeneration As Boolean = False) As Double

    ' ## Reads and interpolates zero rates from sheet
    Dim dbl_Output As Double

    str_RateType = UCase(str_RateType)
    Dim dbl_DF As Double, dbl_FwdZero As Double

    ' Return a rate of 0 if the start and end dates are identical, or backwards is disabled
    If lng_startdate = lng_EndDate Then
        dbl_DF = 1
        dbl_FwdZero = 0
    ElseIf lng_startdate > lng_EndDate And bln_AllowBackwards = False Then
        dbl_DF = 1
        dbl_FwdZero = 0
    Else
        Dim dbl_EndZero As Double, dbl_StartZero As Double

        If str_InterpPillars = "" Then
            ' Use curve pillar points.  Take into account spreads and shifts
            dbl_EndZero = Interp_Lin(lngArr_PillarDates, dblArr_LiveZeros, lng_EndDate, True) + dbl_ZSpread
            dbl_StartZero = Interp_Lin(lngArr_PillarDates, dblArr_LiveZeros, lng_startdate, True) + dbl_ZSpread
        Else
            ' Use specified pillar points
            Dim arr_InterpPillars As Variant: arr_InterpPillars = Split(str_InterpPillars, "|")
            Dim int_NumInterpPillars As Integer: int_NumInterpPillars = UBound(arr_InterpPillars) - LBound(arr_InterpPillars) + 1
            Dim int_ctr As Integer
            Dim lngArr_InterpMatDates() As Long: ReDim lngArr_InterpMatDates(1 To int_NumInterpPillars) As Long
            Dim dblArr_InterpRates() As Double: ReDim dblArr_InterpRates(1 To int_NumInterpPillars) As Double

            For int_ctr = 1 To int_NumInterpPillars
                lngArr_InterpMatDates(int_ctr) = lng_BuildDate + arr_InterpPillars(LBound(arr_InterpPillars) + int_ctr - 1)
                dblArr_InterpRates(int_ctr) = Me.Lookup_Rate(lng_BuildDate, lngArr_InterpMatDates(int_ctr), "ZERO", "", _
                str_CouponFreq, bln_AllowBackwards, bln_GetNotionalFactor, dbl_ZSpread)
            Next int_ctr

            ' Calculated rates after taking into account spreads and shifts
            dbl_EndZero = Interp_Lin(lngArr_InterpMatDates, dblArr_InterpRates, lng_EndDate, True) + dbl_ZSpread
            dbl_StartZero = Interp_Lin(lngArr_InterpMatDates, dblArr_InterpRates, lng_startdate, True) + dbl_ZSpread
        End If

        Select Case UCase(str_InterpConv)
            Case "ZERO"
                dbl_FwdZero = (dbl_EndZero * (lng_EndDate - lng_BuildDate) - dbl_StartZero * (lng_startdate - lng_BuildDate)) / (lng_EndDate - lng_startdate)
                dbl_DF = Exp(-dbl_FwdZero / 100 * calc_yearfrac(lng_startdate, lng_EndDate, str_Daycount_Zero))
            Case "ACT/360"
                dbl_DF = (1 + dbl_StartZero / 100 * calc_yearfrac(lng_BuildDate, lng_startdate, "ACT/360")) / (1 + dbl_EndZero / 100 * calc_yearfrac(lng_BuildDate, lng_EndDate, "ACT/360"))
                dbl_FwdZero = Convert_DFToZero(dbl_DF, lng_startdate, lng_EndDate, "ZERO", "")
            Case "ACT/365"
                dbl_DF = (1 + dbl_StartZero / 100 * calc_yearfrac(lng_BuildDate, lng_startdate, "ACT/365")) / (1 + dbl_EndZero / 100 * calc_yearfrac(lng_BuildDate, lng_EndDate, "ACT/365"))
                dbl_FwdZero = Convert_DFToZero(dbl_DF, lng_startdate, lng_EndDate, "ZERO", "")
        End Select
    End If

    If bln_GetNotionalFactor = True Then
        dbl_Output = (1 / dbl_DF) / calc_yearfrac(lng_startdate, lng_EndDate, str_RateType) * 100
    Else
        Select Case str_RateType
            Case "ZERO": dbl_Output = dbl_FwdZero
            Case "DF"
                dbl_Output = dbl_DF
            Case "ACT/365": dbl_Output = Convert_DFToZero(dbl_DF, lng_startdate, lng_EndDate, str_RateType)
            Case "ACT/360": dbl_Output = Convert_DFToZero(dbl_DF, lng_startdate, lng_EndDate, str_RateType)
            'Case "ACT/ACT": dbl_output = Convert_DFToZero(dbl_DF, lng_StartDate, lng_EndDate, str_RateType, str_CouponFreq, bln_IsFwdGeneration)
            'Case "ACT/ACT NM": dbl_output = (1 / dbl_DF - 1) / Calc_YearFrac(lng_StartDate, lng_EndDate, str_RateType, str_CouponFreq, bln_IsFwdGeneration) * 100
            'Case Else: dbl_output = (1 / dbl_DF - 1) / Calc_YearFrac(lng_StartDate, lng_EndDate, str_RateType) * 100
            Case "ACT/ACT", "ACT/ACT NM", "ACT/ACT XTE", "ACT/ACT CPN": dbl_Output = Convert_DFToZero(dbl_DF, lng_startdate, lng_EndDate, str_RateType, str_CouponFreq, bln_IsFwdGeneration)
            Case Else: dbl_Output = Convert_DFToZero(dbl_DF, lng_startdate, lng_EndDate, str_RateType)
        End Select
    End If

    Lookup_Rate = dbl_Output
End Function

Public Function Lookup_DFs(lng_ValDate As Long, lngLst_PmtDates As Collection, Optional bln_AllowBackwards As Boolean = True, _
    Optional dbl_ZSpread As Double = 0) As Collection
    ' ## Return discount factors from the curve, from the valuation date specified to the payment dates specified
    Dim dblLst_output As Collection: Set dblLst_output = New Collection
    Dim int_ctr As Integer

    For int_ctr = 1 To lngLst_PmtDates.count
        If bln_AllowBackwards = False And lngLst_PmtDates(int_ctr) <= lng_ValDate Then
            Call dblLst_output.Add(1)
        Else
            Call dblLst_output.Add(Me.Lookup_Rate(lng_ValDate, lngLst_PmtDates(int_ctr), "DF", , , , , dbl_ZSpread))
        End If
    Next int_ctr

    Set Lookup_DFs = dblLst_output
End Function

Public Function Lookup_MaturityFromIndex(int_Index As Integer) As Long
    ' ## Return the maturity date corresponding to the specified pillar index
    Lookup_MaturityFromIndex = lngArr_PillarDates(int_Index)
End Function

Private Function Lookup_MaturityFromLabel(str_Label As String) As Long
    ' ## Return the maturity date corresponding to the specified pillar label
    Dim int_FoundIndex As Integer: int_FoundIndex = Examine_FindIndex(Convert_RangeToList(rng_MXTerms), str_Label)
    Debug.Assert int_FoundIndex <> -1
    Lookup_MaturityFromLabel = lngArr_PillarDates(int_FoundIndex)
End Function

Private Function Lookup_NumDays(str_Label As String) As Integer
    ' ## Return the maturity date corresponding to the specified pillar label
    Lookup_NumDays = Lookup_MaturityFromLabel(str_Label) - lng_BuildDate
End Function


' ## METHODS - SCENARIOS
Public Sub Scen_ApplyBase()
    ' ## Can only run this after initial bootstrap

    'Handle error caused by blank worksheet

    'If Me.CurveName = "NID_ISL" Then
    '    Debug.Print "HAHA"
    'End If
    If Me.CurveName = "" Then Exit Sub
    If rng_InstTypes(1, 1).Offset(0, -2).Value = "" Then Exit Sub

    Call csh_ParShifts_Rel.Initialize(ShockType.Relative)
    Call csh_ParShifts_Abs.Initialize(ShockType.Absolute)
    Call csh_ZeroShifts_Rel.Initialize(ShockType.Relative)
    Call csh_ZeroShifts_Abs.Initialize(ShockType.Absolute)

    Call Action_ClearBelow(rng_Days_TopLeft, 5)

    rng_FinalPars.Value = rng_OrigPars.Value
    rng_FinalZeros.Value = rng_OrigZeros.Value

    Call Me.SetCurveState(CurveState_IRC.original)
    Call StoreAsFinalState
    Call PurgeTempStates
End Sub

Public Sub Scen_AddByDays(int_numdays As Integer, str_ShockType As String, dbl_Amount As Double)
    ' ## Shift appleis at the specified number of days and is interpolated/extrapolated at each actual pillar
    Select Case UCase(str_ShockType)
        Case "REL PAR": Call csh_ParShifts_Rel.AddShift(int_numdays, dbl_Amount)
        Case "REL ZERO": Call csh_ZeroShifts_Rel.AddShift(int_numdays, dbl_Amount)
        Case "ABS PAR": Call csh_ParShifts_Abs.AddShift(int_numdays, dbl_Amount)
        Case "ABS ZERO": Call csh_ZeroShifts_Abs.AddShift(int_numdays, dbl_Amount)
        Case Else: Debug.Assert False
    End Select
End Sub

Public Sub Scen_AddByLabel(str_Label As String, str_ShockType As String, dbl_Amount As Double)
    ' ## Shift applies only to the specified pillar
    Dim int_numdays As Integer: int_numdays = Lookup_MaturityFromLabel(str_Label) - lng_BuildDate

    Select Case UCase(str_ShockType)
        Case "REL PAR": Call csh_ParShifts_Rel.AddIsolatedShift(int_numdays, dbl_Amount)
        Case "REL ZERO": Call csh_ZeroShifts_Rel.AddIsolatedShift(int_numdays, dbl_Amount)
        Case "ABS PAR": Call csh_ParShifts_Abs.AddIsolatedShift(int_numdays, dbl_Amount)
        Case "ABS ZERO": Call csh_ZeroShifts_Abs.AddIsolatedShift(int_numdays, dbl_Amount)
        Case Else: Debug.Assert False
    End Select
End Sub

Public Sub Scen_AddUniform(str_ShockType As String, dbl_Amount As Double)
    ' ## Shift applies to all maturity pillars
    Select Case UCase(str_ShockType)
        Case "REL PAR": Call csh_ParShifts_Rel.AddUniformShift(dbl_Amount)
        Case "REL ZERO": Call csh_ZeroShifts_Rel.AddUniformShift(dbl_Amount)
        Case "ABS PAR": Call csh_ParShifts_Abs.AddUniformShift(dbl_Amount)
        Case "ABS ZERO": Call csh_ZeroShifts_Abs.AddUniformShift(dbl_Amount)
        Case Else: Debug.Assert False
    End Select
End Sub

Public Sub Scen_ApplyCurrent(bln_Propagation As Boolean)
    ' ## Update final rates to reflect the specified shifts

    'Handle error caused by blank worksheet
    If Me.CurveName = "" Then Exit Sub
    If rng_InstTypes(1, 1).Offset(0, -2).Value = "" Then Exit Sub

    Call Me.SetCurveState(CurveState_IRC.Final)

    ' Other shock types
    Dim int_NumShifts_RelPar As Integer: int_NumShifts_RelPar = csh_ParShifts_Rel.NumShifts
    Dim int_NumShifts_AbsPar As Integer: int_NumShifts_AbsPar = csh_ParShifts_Abs.NumShifts
    Dim int_NumShifts_RelZero As Integer: int_NumShifts_RelZero = csh_ZeroShifts_Rel.NumShifts
    Dim int_NumShifts_AbsZero As Integer: int_NumShifts_AbsZero = csh_ZeroShifts_Abs.NumShifts

    If int_NumShifts_AbsPar + int_NumShifts_AbsZero + int_NumShifts_RelPar + int_NumShifts_RelZero > 0 Then

        Dim int_ActiveDTM As Integer, int_RowCtr As Integer

        ' If both pars and zeros are shocked, ignores the zero shocks
        If int_NumShifts_AbsPar + int_NumShifts_RelPar > 0 Then
            ' Prepare arrays for interpolation if required
            Dim dblArr_OrigPillarPars() As Variant: dblArr_OrigPillarPars = Convert_RangeToArr2D(rng_OrigPars)
            Dim dblArr_FinalPillarPars() As Variant: dblArr_FinalPillarPars = Convert_RangeToArr2D(rng_FinalPars)

            ' Convert shocks for number of days to shocks for each pillar
            For int_RowCtr = 1 To int_NumPoints
                int_ActiveDTM = lngArr_PillarDates(int_RowCtr) - lng_BuildDate
                dblArr_FinalPillarPars(int_RowCtr, 1) = dblArr_OrigPillarPars(int_RowCtr, 1) * (1 + csh_ParShifts_Rel.ReadShift(int_ActiveDTM) / 100) _
                    + csh_ParShifts_Abs.ReadShift(int_ActiveDTM)
            Next int_RowCtr

            ' Output shifts to sheet
            If int_NumShifts_RelPar > 0 Then
                rng_Days_TopLeft.Resize(int_NumShifts_RelPar, 1).Value = csh_ParShifts_Rel.Days_Arr
                rng_RelPar_TopLeft.Resize(int_NumShifts_RelPar, 1).Value = csh_ParShifts_Rel.Shifts_Arr
            End If

            If int_NumShifts_AbsPar > 0 Then
                rng_Days_TopLeft.Offset(int_NumShifts_RelPar, 0).Resize(int_NumShifts_AbsPar, 1).Value = csh_ParShifts_Abs.Days_Arr
                rng_AbsPar_TopLeft.Offset(int_NumShifts_RelPar, 0).Resize(int_NumShifts_AbsPar, 1).Value = csh_ParShifts_Abs.Shifts_Arr
            End If

            ' Write shifted values back to sheet
            rng_FinalPars.Value = dblArr_FinalPillarPars

            ' Generate final zero rates
            Me.Action_Bootstrap False
        ElseIf int_NumShifts_AbsZero + int_NumShifts_RelZero > 0 Then
            ' Update final zero rates, final par rates are not needed
            rng_FinalPars.ClearContents
            Dim dblArr_OrigPillarZeros() As Double: dblArr_OrigPillarZeros = dic_CurveStates(CurveState_IRC.original)

            ' Convert shocks for number of days to shocks for each pillar
            For int_RowCtr = 1 To int_NumPoints
                int_ActiveDTM = lngArr_PillarDates(int_RowCtr) - lng_BuildDate
                dblArr_LiveZeros(int_RowCtr) = dblArr_OrigPillarZeros(int_RowCtr) * (1 + csh_ZeroShifts_Rel.ReadShift(int_ActiveDTM) / 100) _
                    + csh_ZeroShifts_Abs.ReadShift(int_ActiveDTM)
            Next int_RowCtr

            ' Output shifts to sheet
            If int_NumShifts_RelZero > 0 Then
                rng_Days_TopLeft.Resize(int_NumShifts_RelZero, 1).Value = csh_ZeroShifts_Rel.Days_Arr
                rng_RelZero_TopLeft.Resize(int_NumShifts_RelZero, 1).Value = csh_ZeroShifts_Rel.Shifts_Arr
            End If

            If int_NumShifts_AbsZero > 0 Then
                rng_Days_TopLeft.Resize(int_NumShifts_AbsZero, 1).Value = csh_ZeroShifts_Abs.Days_Arr
                rng_AbsZero_TopLeft.Resize(int_NumShifts_AbsZero, 1).Value = csh_ZeroShifts_Abs.Shifts_Arr
            End If

            ' Write shifted values back to sheet
            rng_FinalZeros.Value = Convert_Array1Dto2D(dblArr_LiveZeros)
        End If
    ElseIf str_BaseCurve <> "<NONE>" And bln_Propagation = True Then
        ' Need to re-bootstrap if base curve is shocked and propagation is turned on
        Dim irc_BaseCurve As Data_IRCurve
        If str_BaseCurve <> "<NONE>" Then Set irc_BaseCurve = GetObject_IRCurve(str_BaseCurve, True, False)
        If irc_BaseCurve.NumShifts > 0 Then
            Call Me.Action_Bootstrap(False)
        Else
            rng_FinalPars.Value = rng_OrigPars.Value
            rng_FinalZeros.Value = rng_OrigZeros.Value
            dblArr_LiveZeros = Convert_RangeToDblArr(rng_FinalZeros)

        End If
    Else
        rng_FinalPars.Value = rng_OrigPars.Value
        rng_FinalZeros.Value = rng_OrigZeros.Value
        dblArr_LiveZeros = Convert_RangeToDblArr(rng_FinalZeros)
    End If

    Call StoreAsFinalState

End Sub

Public Sub Scen_ReceivePropagation()
    Dim str_PropSource As String: str_PropSource = wks_Location.Range("C2").Value
    Dim irc_Source As Data_IRCurve
    Dim lng_ActiveMat As Long, dbl_ActiveShift As Double
    Dim int_RowCtr As Integer

    If str_PropSource <> "<NONE>" Then
        Set irc_Source = GetObject_IRCurve(str_PropSource, True, False)

        If irc_Source.NumShifts > 0 Then
            ' Ignore shocks to the curve and only use propagated shocks from the source
            Me.Scen_ApplyBase

            ' For each pillar, find difference between shocked and original zero of the source, output it as the required zero shock
            For int_RowCtr = 1 To int_NumPoints
                lng_ActiveMat = lngArr_PillarDates(int_RowCtr)
                dbl_ActiveShift = irc_Source.Lookup_Rate(lng_BuildDate, lng_ActiveMat, "ZERO", False) - irc_Source.Lookup_Rate(lng_BuildDate, lng_ActiveMat, "ZERO", True)
                Call csh_ZeroShifts_Abs.AddShift(lng_ActiveMat - lng_BuildDate, dbl_ActiveShift)
            Next int_RowCtr

            ' Build shocked curve
            Call Me.Scen_ApplyCurrent(True)
        End If
    End If
End Sub


' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Sub SetZeroRate(int_Index As Integer, dbl_NewRate As Double)
    ' ## Used by bootstrapping solver function
    dblArr_LiveZeros(int_Index) = dbl_NewRate
End Sub

Public Sub SetCurveState(enu_NewState As CurveState_IRC, Optional int_PillarIndex As Integer = 0)
    ' ## Change the set of zero rates being used for rate lookups
    Dim csh_ToApply As CurveDaysShift

    If enu_NewState <> enu_LiveState Then
        enu_LiveState = enu_NewState

        If int_PillarIndex = 0 Then
            ' Uniform shifts or standard cases
            If dic_CurveStates.Exists(enu_LiveState) Then
                dblArr_LiveZeros = dic_CurveStates(enu_LiveState)
            Else
                ' Gather the shift for the new curve state
                Set csh_ToApply = BuildDaysShifts_IRCurve(enu_NewState)

                ' Derive the zero pillars under the state
                Dim dblArr_Pillars() As Double: ReDim dblArr_Pillars(1 To int_NumPoints) As Double
                Dim dblArr_FinalBase() As Double: dblArr_FinalBase = dic_CurveStates(CurveState_IRC.Final)

                Dim int_ctr As Integer, int_ActiveNumDays As Integer
                For int_ctr = 1 To int_NumPoints
                    int_ActiveNumDays = lngArr_PillarDates(int_ctr) - lng_BuildDate

                    Select Case csh_ToApply.ShockType
                        Case ShockType.Absolute
                            dblArr_Pillars(int_ctr) = dblArr_FinalBase(int_ctr) + csh_ToApply.ReadShift(int_ActiveNumDays)
                        Case ShockType.Relative
                            dblArr_Pillars(int_ctr) = dblArr_FinalBase(int_ctr) * (1 + csh_ToApply.ReadShift(int_ActiveNumDays) / 100)
                    End Select
                Next int_ctr

                ' Store and use the pillar set
                Call dic_CurveStates.Add(enu_NewState, dblArr_Pillars)
                dblArr_LiveZeros = dblArr_Pillars
            End If
        Else
            ' Apply the shift only to the specified pillar of the curve
            Dim str_key As String: str_key = enu_NewState & "|" & int_PillarIndex
            If dic_CurveStates.Exists(str_key) Then
                dblArr_LiveZeros = dic_CurveStates(str_key)
            Else
                ' Gather the shift for the new curve state
                Set csh_ToApply = BuildDaysShifts_IRCurve(enu_NewState)

                dblArr_LiveZeros = dic_CurveStates(CurveState_IRC.Final)
                dblArr_LiveZeros(int_PillarIndex) = dblArr_LiveZeros(int_PillarIndex) + csh_ToApply.ReadShift(0)
                Call dic_CurveStates.Add(str_key, dblArr_LiveZeros)
            End If
        End If
    End If
End Sub

Private Sub PurgeTempStates()
    ' ## Remove states used for finite differencing
    ' Remember permament states
    Dim dblArr_Orig() As Double, dblArr_Final() As Double
    dblArr_Orig = dic_CurveStates(CurveState_IRC.original)
    dblArr_Final = dic_CurveStates(CurveState_IRC.Final)

    ' Repopulate dictionary
    Call dic_CurveStates.RemoveAll
    Call dic_CurveStates.Add(CurveState_IRC.original, dblArr_Orig)
    Call dic_CurveStates.Add(CurveState_IRC.Final, dblArr_Final)
End Sub

Private Sub StoreAsFinalState()
    ' ## Store the current live state as the 'final' state
    Call dic_CurveStates.Remove(CurveState_IRC.Final)
    Call dic_CurveStates.Add(CurveState_IRC.Final, dblArr_LiveZeros)
    enu_LiveState = CurveState_IRC.Final
End Sub


' ## METHODS - SETUP
Public Sub SetParams(rng_Params As Range, str_CurveName As String)
    Dim rng_HolsOther As Range

    With wks_Location
        .Range("A2").Value = cfg_Settings.CurrentBuildDate
        .Range("B2").Value = str_CurveName
        .Range("C2:S2").Value = rng_Params.Value
    End With

    Call StoreStaticValues
End Sub

Private Sub StoreStaticValues()
    Dim int_ColCtr As Integer
    Dim rng_TopLeft As Range: Set rng_TopLeft = wks_Location.Range("A2")
    Dim str_ActiveValue As String

    With rng_TopLeft
        int_ColCtr = 0
        lng_BuildDate = .Offset(0, int_ColCtr).Value

        int_ColCtr = int_ColCtr + 1
        str_CurveName = .Offset(0, int_ColCtr).Value

        int_ColCtr = int_ColCtr + 2
        str_BaseCurve = .Offset(0, int_ColCtr).Value

        int_ColCtr = int_ColCtr + 1
        str_InterpConv = .Offset(0, int_ColCtr).Value

        int_ColCtr = int_ColCtr + 2
        str_ActiveValue = UCase(.Offset(0, int_ColCtr).Value)
        bln_IsAdjMatForEst = (str_ActiveValue = "YES")
        bln_EstFloat = (str_ActiveValue = "YES" Or str_ActiveValue = "CURVE_ASMT")

        int_ColCtr = int_ColCtr + 1
        bln_EOMRule = (UCase(.Offset(0, int_ColCtr).Value) = "YES")

        int_ColCtr = int_ColCtr + 1
        bln_InBlocks = (UCase(.Offset(0, int_ColCtr).Value) = "YES")

        int_ColCtr = int_ColCtr + 1
        bln_AcceptNil = (UCase(.Offset(0, int_ColCtr).Value) = "YES")

        int_ColCtr = int_ColCtr + 9
        str_Priorities = UCase(.Offset(0, int_ColCtr).Value)
    End With
End Sub

Public Sub LoadRates()
    With wks_Location
        Dim lng_DataDate As Long: lng_DataDate = cfg_Settings.CurrentDataDate
        Dim cal_Active As Calendar
        Dim int_ActiveSettleDays As Integer
        Dim lng_ActiveStartDate As Long
        Dim str_SQLCode As String: str_SQLCode = iqs_Queries.Lookup_SQL(str_CurveName, lng_DataDate)
        Dim rng_TopLeft As Range: Set rng_TopLeft = .Range("A7")

        ' Clear out existing data and shifts
        Call Action_ClearBelow(.Range("A7"), 16)
        Call Action_ClearBelow(.Range("R7"), 5)
        Call Action_ClearBelow(.Range("X7"), 3)

        Call Action_Query_Access(cfg_Settings.RatesDBPath, str_SQLCode, rng_TopLeft)

        ' Data cleansing
        Dim int_ctr As Integer
        int_NumPoints = Examine_NumRows(.Range("A7"))
        Dim str_Currency As String: str_Currency = UCase(.Range("B7").Value)
        Dim str_ActiveTerm As String, str_ActiveType As String, lng_ExDate As Long, str_ActiveBDC As String, str_ActiveFlowsFreq As String
        Dim str_FutRule As String: str_FutRule = .Range("F2").Value
        Dim cal_ActiveUSD As Calendar, cal_ActiveFGN As Calendar, str_ActiveCalendars As String
        Dim str_ActiveFgnBDC As String, str_ActiveUSDBDC As String

        Dim lng_Temp As Long
        For int_ctr = 1 To int_NumPoints
            str_ActiveTerm = UCase(.Range("D6").Offset(int_ctr, 0).Value)
            str_ActiveType = UCase(.Range("C6").Offset(int_ctr, 0).Value)
            str_ActiveCalendars = UCase(.Range("L6").Offset(int_ctr, 0).Value)
            str_ActiveBDC = UCase(.Range("I6").Offset(int_ctr, 0).Value)
            int_ActiveSettleDays = .Range("J6").Offset(int_ctr, 0).Value

            ' Generate start and end dates for each instrument
            If str_ActiveType = "IRBSWP" Then
                ' Special case for xccy basis swaps
                str_ActiveTerm = .Range("D6").Offset(int_ctr, 0).Value
                cal_ActiveUSD = cas_Calendars.Lookup_Calendar(ReadSetting_LegA(str_ActiveCalendars))
                cal_ActiveFGN = cas_Calendars.Lookup_Calendar(ReadSetting_LegB(str_ActiveCalendars))
                str_ActiveUSDBDC = ReadSetting_LegA(str_ActiveBDC)
                str_ActiveFgnBDC = ReadSetting_LegB(str_ActiveBDC)
                str_ActiveFlowsFreq = ""

                lng_ActiveStartDate = GetIRStartDate(str_CurveName, str_ActiveType, str_ActiveTerm, lng_ExDate, lng_BuildDate, int_ActiveSettleDays, _
                    cal_ActiveFGN, str_ActiveFgnBDC, str_Currency, dic_GlobalStaticInfo)
                .Range("N6").Offset(int_ctr, 0).Value = lng_ActiveStartDate
                .Range("O6").Offset(int_ctr, 0).Value = GetIRMatDate(str_CurveName, str_ActiveType, str_ActiveTerm, lng_ExDate, lng_BuildDate, _
                    lng_ActiveStartDate, cal_ActiveUSD, str_ActiveUSDBDC, str_ActiveFlowsFreq, bln_IsAdjMatForEst, bln_EOMRule, str_Currency, _
                    str_FutRule, dic_GlobalStaticInfo)
            Else
                str_ActiveFlowsFreq = UCase(.Range("K6").Offset(int_ctr, 0).Value)
                'cal_Active = cas_Calendars.Lookup_Calendar(.Range("L6").Offset(int_Ctr, 0).Value)

                If bln_IsAdjMatForEst = True Then
                    If .Range("M6").Offset(int_ctr, 0).Value = "" Then
                        cal_Active = cas_Calendars.Lookup_Calendar(.Range("L6").Offset(int_ctr, 0).Value)
                    Else
                        cal_Active = cas_Calendars.Lookup_Calendar(.Range("M6").Offset(int_ctr, 0).Value)
                    End If
                Else
                    cal_Active = cas_Calendars.Lookup_Calendar(.Range("L6").Offset(int_ctr, 0).Value)
                End If

                lng_ExDate = .Range("F6").Offset(int_ctr, 0).Value
                lng_ActiveStartDate = GetIRStartDate(str_CurveName, str_ActiveType, str_ActiveTerm, lng_ExDate, lng_BuildDate, int_ActiveSettleDays, _
                    cal_Active, str_ActiveBDC, str_Currency, dic_GlobalStaticInfo)
                .Range("N6").Offset(int_ctr, 0).Value = lng_ActiveStartDate
                .Range("O6").Offset(int_ctr, 0).Value = GetIRMatDate(str_CurveName, str_ActiveType, str_ActiveTerm, lng_ExDate, lng_BuildDate, lng_ActiveStartDate, _
                    cal_Active, str_ActiveBDC, str_ActiveFlowsFreq, bln_IsAdjMatForEst, bln_EOMRule, str_Currency, str_FutRule, dic_GlobalStaticInfo)
            End If

            ' Convert futures prices to par rates
            If str_ActiveType = "IRFUTB" Then
                .Range("G6").Offset(int_ctr, 0).Value = 100 - .Range("G6").Offset(int_ctr, 0).Value
            End If
        Next int_ctr

        ' Unless AcceptNil = True, delete points where market rate is zero
        Dim rng_ActiveRow As Range: Set rng_ActiveRow = .Range("A7").Resize(1, 15)
        If bln_AcceptNil = False Then
            While rng_ActiveRow(1, 1).Value <> ""
                If rng_ActiveRow(1, 7).Value = 0 Then
                    Set rng_ActiveRow = rng_ActiveRow.Offset(-1, 0)
                    rng_ActiveRow.Offset(1, 0).Delete Shift:=xlUp
                End If
                Set rng_ActiveRow = rng_ActiveRow.Offset(1, 0)
            Wend
        End If

        ' Update ranges as number of points may have changed
        int_NumPoints = Examine_NumRows(.Range("A7"))

        ' Sort by maturity date
        Call .Range("A6").Resize(int_NumPoints + 1, 15).Sort(Key1:=.Range("O6"), Order1:=xlAscending, Header:=xlYes, Orientation:=xlSortColumns)

        ' Update ranges as number of points may have changed
        Me.AssignRanges

        ' Take note of priority order if specified
        Dim strLst_Priorities As Collection, dic_Priorities As Dictionary
        Dim arr_Split As Variant, str_ActiveElement As Variant
        If str_Priorities <> "-" Then
            int_ctr = 1
            Set strLst_Priorities = New Collection
            Set dic_Priorities = New Dictionary
            arr_Split = Convert_Split(str_Priorities, "|")
            For Each str_ActiveElement In arr_Split
                Call strLst_Priorities.Add(UCase(CStr(str_ActiveElement)))
                Call dic_Priorities.Add(UCase(CStr(str_ActiveElement)), int_ctr)
                int_ctr = int_ctr + 1
            Next str_ActiveElement
        End If

        ' Check block consistency
        Dim int_BlockStart As Integer, int_BlockEnd As Integer, int_HeaderRowsAbove As Integer
        If bln_InBlocks = True Then
            Dim rng_Active As Range
            ' Act in the order of priorities
            For Each str_ActiveElement In strLst_Priorities
                int_BlockStart = WorksheetFunction.Match(str_ActiveElement, rng_InstTypes, 0)
                int_HeaderRowsAbove = rng_InstTypes(1, 1).Row - 1
                int_BlockEnd = Evaluate("SUMPRODUCT(MAX((" & rng_InstTypes.Address(, , , True) & "=""" & str_ActiveElement & """)*ROW(" & rng_InstTypes.Address(, , , True) _
                    & ")))") - int_HeaderRowsAbove

                ' Delete any rows in between which do not match the instrument type
                Set rng_Active = rng_InstTypes(int_BlockStart, 1)
                While rng_Active.Row - int_HeaderRowsAbove <= int_BlockEnd
                    If rng_Active.Value <> str_ActiveElement Then
                        Set rng_Active = rng_Active.Offset(-1, 0)
                        Call rng_Active.Offset(1, -2).Resize(1, 15).Delete(Shift:=xlShiftUp)
                        int_BlockEnd = int_BlockEnd - 1
                    End If
                    Set rng_Active = rng_Active.Offset(1, 0)
                Wend
            Next str_ActiveElement
        End If

        ' Update ranges as number of points may have changed
        Me.AssignRanges

        ' Handle cases where dates are duplicated
        Const int_ColsToDelete As Integer = 16
        int_ctr = 1
        While int_ctr <= int_NumPoints - 1
            str_ActiveType = UCase(rng_InstTypes(int_ctr, 1).Value)
            If str_ActiveType = "FXFWDPT" Then
                ' Assumes that FXFWDPT will never clash with another type
                If rng_MatDates(int_ctr, 1).Value = rng_MatDates(int_ctr + 1, 1).Value Then
                    ' Earlier has priority, delete later point
                    Call rng_InstTypes(int_ctr, 1).Offset(1, -2).Resize(1, int_ColsToDelete).Delete(Shift:=xlShiftUp)
                    int_NumPoints = int_NumPoints - 1
                ElseIf rng_MatDates(int_ctr, 1).Value = rng_StartDates(int_ctr, 1).Value Then
                    Call rng_InstTypes(int_ctr, 1).Offset(0, -2).Resize(1, int_ColsToDelete).Delete(Shift:=xlShiftUp)
                    int_NumPoints = int_NumPoints - 1
                End If
            End If

            If rng_MatDates(int_ctr, 1).Value = rng_MatDates(int_ctr + 1, 1).Value Then
                If dic_Priorities Is Nothing Then
                    ' Later has priority, delete earlier point
                    Call rng_InstTypes(int_ctr, 1).Offset(0, -2).Resize(1, int_ColsToDelete).Delete(Shift:=xlShiftUp)
                    int_NumPoints = int_NumPoints - 1
                    int_ctr = int_ctr - 1
                Else
                    If dic_Priorities(str_ActiveType) > dic_Priorities(UCase(rng_InstTypes(int_ctr + 1, 1).Value)) Then
                        ' Earlier has priority, delete later point
                        Call rng_InstTypes(int_ctr, 1).Offset(1, -2).Resize(1, int_ColsToDelete).Delete(Shift:=xlShiftUp)
                        int_NumPoints = int_NumPoints - 1
                    Else
                        ' Later has priority, delete earlier point
                        Call rng_InstTypes(int_ctr, 1).Offset(0, -2).Resize(1, int_ColsToDelete).Delete(Shift:=xlShiftUp)
                        int_NumPoints = int_NumPoints - 1
                        int_ctr = int_ctr - 1
                    End If
                End If
            End If

            int_ctr = int_ctr + 1
        Wend

        ' Update ranges as number of points may have changed
        Me.AssignRanges

        ' Correct formats
        Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
        rng_OrigPars.NumberFormat = "General"
        rng_FinalPars.NumberFormat = "General"
        rng_OrigZeros.NumberFormat = "General"
        rng_FinalZeros.NumberFormat = "General"
        rng_SpotDays.NumberFormat = "General"
        rng_StartDates.NumberFormat = str_DateFormat
        rng_MatDates.NumberFormat = str_DateFormat
        rng_FutMats.NumberFormat = str_DateFormat
        rng_MXSortTerms.NumberFormat = "General"

        Dim int_FutCtr As Integer: int_FutCtr = 0
        For int_ctr = 1 To int_NumPoints
            ' Label futures contracts by order to mature
            str_ActiveType = UCase(.Range("C6").Offset(int_ctr, 0).Value)
            If str_ActiveType = "IRFUTB" Then
                int_FutCtr = int_FutCtr + 1
                .Range("D6").Offset(int_ctr, 0).Value = "FUT" & int_FutCtr
            End If
        Next int_ctr

        ' Load base scenario and display terms for ease of reference
        rng_FinalPars.Offset(0, -1).Value = rng_MXTerms.Value
        rng_FinalPars.Value = rng_OrigPars.Value2

        Me.Action_Bootstrap True
    End With
End Sub

Public Sub AssignRanges()
    int_NumPoints = Examine_NumRows(wks_Location.Range("A7"))

    If int_NumPoints > 0 Then
        Set rng_InstTypes = wks_Location.Range("C7").Resize(int_NumPoints, 1)
        Set rng_MXTerms = rng_InstTypes.Offset(0, 1)
        Set rng_MXSortTerms = rng_MXTerms.Offset(0, 1)
        Set rng_FutMats = rng_MXSortTerms.Offset(0, 1)
        Set rng_OrigPars = rng_FutMats.Offset(0, 1)
        Set rng_DayCounts = rng_OrigPars.Offset(0, 1)
        Set rng_BDCs = rng_DayCounts.Offset(0, 1)
        Set rng_SpotDays = rng_BDCs.Offset(0, 1)
        Set rng_FlowsFreq = rng_SpotDays.Offset(0, 1)
        Set rng_PmtCalendars = rng_FlowsFreq.Offset(0, 1)
        Set rng_EstCalendars = rng_PmtCalendars.Offset(0, 1)
        Set rng_StartDates = rng_EstCalendars.Offset(0, 1)
        Set rng_MatDates = rng_StartDates.Offset(0, 1)
        Set rng_OrigZeros = rng_MatDates.Offset(0, 1)
        Set rng_Days_TopLeft = rng_OrigZeros(1, 1).Offset(0, 2)
        Set rng_RelPar_TopLeft = rng_Days_TopLeft.Offset(0, 1)
        Set rng_AbsPar_TopLeft = rng_RelPar_TopLeft.Offset(0, 1)
        Set rng_RelZero_TopLeft = rng_AbsPar_TopLeft.Offset(0, 1)
        Set rng_AbsZero_TopLeft = rng_RelZero_TopLeft.Offset(0, 1)
        Set rng_FinalPars = rng_AbsZero_TopLeft.Offset(0, 3).Resize(int_NumPoints, 1)
        Set rng_FinalZeros = rng_FinalPars.Offset(0, 1)

        ' Fill interpolation cache
        lngArr_PillarDates = Convert_RangeToLngArr(rng_MatDates)

        ' Set up curve state dictionary
        Set dic_CurveStates = New Dictionary
        dic_CurveStates.CompareMode = CompareMethod.TextCompare
        Call dic_CurveStates.Add(CurveState_IRC.original, Convert_RangeToDblArr(rng_OrigZeros))
        Dim dblArr_FinalZeros() As Double: dblArr_FinalZeros = Convert_RangeToDblArr(rng_FinalZeros)
        Call dic_CurveStates.Add(CurveState_IRC.Final, dblArr_FinalZeros)

        ' Force mode to final state.  Don't use SetCurveState because it won't update the live zeros array if already in final state
        enu_LiveState = CurveState_IRC.Final
        dblArr_LiveZeros = dblArr_FinalZeros
    End If
End Sub


' ## METHODS - ACTIONS / OPERATIONS
Public Sub Action_Bootstrap(bln_IsInitialLoad As Boolean)
    Dim fld_AppState_Orig As ApplicationState: fld_AppState_Orig = Gather_ApplicationState(ApplicationStateType.Current)
    Dim fld_AppState_Opt As ApplicationState: fld_AppState_Opt = Gather_ApplicationState(ApplicationStateType.Optimized)
    Call Action_SetAppState(fld_AppState_Opt)

    Application.StatusBar = "Data date: " & Format(cfg_Settings.CurrentDataDate, "dd/mm/yyyy") & "     IRC: " & str_CurveName & " (bootstrapping)"
    Const str_Self As String = "<SELF>"

    ' Ensure final zeros are being modified and updated
    Call SetCurveState(CurveState_IRC.Final)

    ' Read parameters from sheet
    Dim dbl_StartDF As Double, dbl_EndDF As Double, dbl_FwdDF As Double
    Dim lng_startdate As Long, lng_MatDate As Long
    Dim str_CCY As String: str_CCY = wks_Location.Range("B7").Value
    Dim str_FutInterp As String: str_FutInterp = UCase(wks_Location.Range("F2").Value)
    Dim lng_UnAdjEndDate As Long, lng_UnAdjDays As Long
    Dim int_ActiveRow As Integer: int_ActiveRow = 1
    Dim dbl_ActivePar As Double, dbl_NextActivePar As Double
    Dim str_InstType As String
    Dim str_ActiveTerm As String, int_ActiveSortTerm As Integer
    Dim dbl_ActiveZero As Double
    Dim str_ActiveDCC As String, str_ActiveBDC As String, int_ActiveDayCount As Integer
    Dim dbl_templogic As Double

    Dim irc_BaseCurve As Data_IRCurve
    If str_BaseCurve <> "<NONE>" Then Set irc_BaseCurve = GetObject_IRCurve(str_BaseCurve, True, False)
    Dim dbl_BaseCurveDF As Double

    ' Read curve assignments
    Dim str_DiscRateCurve_Basis_LegA As String: str_DiscRateCurve_Basis_LegA = wks_Location.Range("K2").Value
    If str_DiscRateCurve_Basis_LegA = str_Self Then str_DiscRateCurve_Basis_LegA = str_CurveName
    Dim str_EstRateCurve_Basis_LegA As String: str_EstRateCurve_Basis_LegA = wks_Location.Range("L2").Value
    If str_EstRateCurve_Basis_LegA = str_Self Then str_EstRateCurve_Basis_LegA = str_CurveName
    Dim str_DiscRateCurve_Basis_LegB As String: str_DiscRateCurve_Basis_LegB = wks_Location.Range("M2").Value
    If str_DiscRateCurve_Basis_LegB = str_Self Then str_DiscRateCurve_Basis_LegB = str_CurveName
    Dim str_EstRateCurve_Basis_LegB As String: str_EstRateCurve_Basis_LegB = wks_Location.Range("N2").Value
    If str_EstRateCurve_Basis_LegB = str_Self Then str_EstRateCurve_Basis_LegB = str_CurveName

    Dim str_DiscRateCurve_Swap_LegA As String: str_DiscRateCurve_Swap_LegA = wks_Location.Range("O2").Value
    If str_DiscRateCurve_Swap_LegA = str_Self Then str_DiscRateCurve_Swap_LegA = str_CurveName
    Dim str_EstRateCurve_Swap_LegA As String: str_EstRateCurve_Swap_LegA = wks_Location.Range("P2").Value
    If str_EstRateCurve_Swap_LegA = str_Self Then str_EstRateCurve_Swap_LegA = str_CurveName
    Dim str_DiscRateCurve_Swap_LegB As String: str_DiscRateCurve_Swap_LegB = wks_Location.Range("Q2").Value
    If str_DiscRateCurve_Swap_LegB = str_Self Then str_DiscRateCurve_Swap_LegB = str_CurveName
    Dim str_EstRateCurve_Swap_LegB As String: str_EstRateCurve_Swap_LegB = wks_Location.Range("R2").Value
    If str_EstRateCurve_Swap_LegB = str_Self Then str_EstRateCurve_Swap_LegB = str_CurveName

    ' For FX curves
    Dim fxs_Spots As Data_FXSpots
    Dim dbl_Spot As Double, str_DomCurrency As String, str_FgnCurrency As String, cal_ActiveCCY As Calendar
    Dim str_Quotation As String
    Dim dbl_ActiveFXFwdStart As Double, dbl_ActiveFXFwdEnd As Double
    Dim int_ActiveSpotDays As Integer
    Dim bln_1DNo2D As Boolean, dbl_MissingFwd As Double, lng_FXSpotDate As Long
    Dim dbl_DF_USD As Double, dbl_DF_USDEstFwd As Double

    If str_CCY <> "USD" And str_BaseCurve <> "<NONE>" Then
        Set fxs_Spots = GetObject_FXSpots(True)
        str_Quotation = fxs_Spots.Lookup_Quotation(str_CCY)
        If str_Quotation = "DIRECT" Then dbl_Spot = fxs_Spots.Lookup_Spot("USD", str_CCY) Else dbl_Spot = fxs_Spots.Lookup_Spot(str_CCY, "USD")
    End If

    ' Determine what to update
    Dim rng_ParsToBootstrap As Range
    If bln_IsInitialLoad = True Then
        Set rng_ParsToBootstrap = rng_OrigPars
        rng_OrigZeros.ClearContents
    Else
        Set rng_ParsToBootstrap = rng_FinalPars
    End If
    rng_FinalZeros.ClearContents

    ' Re-iteration settings
    Dim bln_ReIterateAll As Boolean
    Const int_MaxReIterations = 3
    Dim dic_RowsToReIterate As New Dictionary
    Dim int_ActiveNumReIterations As Integer

    ' Swap objects
    Dim fldArr_Swaps() As InstParams_IRS: ReDim fldArr_Swaps(1 To int_NumPoints) As InstParams_IRS
    Dim dic_Swaps As New Dictionary
    Dim fld_ActiveParams_Swap As InstParams_IRS
    Dim fld_ActiveParams_LegA As IRLegParams, fld_ActiveParams_LegB As IRLegParams
    Dim irs_Active As Inst_IRSwap
    Dim dbl_FallBackZero As Double
    Dim bln_ActiveIsSwap As Boolean

    ' Secant method
    Dim dic_SecantParams As Dictionary, dic_SecantOutputs As Dictionary
    Dim bln_ReqEstimatedFwd As Boolean

    While int_ActiveRow <= int_NumPoints
        ' Fall back to earlier pillar zero rate if possible
        If int_ActiveRow = 1 Then dbl_FallBackZero = 0 Else dbl_FallBackZero = rng_FinalZeros(int_ActiveRow - 1, 1).Value

        ' Bootstrap point if not been solved, or if re-iteration requested and not yet completed, or if using fixed point iteration
        str_InstType = UCase(rng_InstTypes(int_ActiveRow, 1).Value)
        str_ActiveTerm = rng_MXTerms(int_ActiveRow, 1).Value
        int_ActiveSortTerm = rng_MXSortTerms(int_ActiveRow, 1).Value
        str_ActiveDCC = rng_DayCounts(int_ActiveRow, 1).Value
        int_ActiveDayCount = Examine_DaysPerYear(str_ActiveDCC)
        int_ActiveSpotDays = rng_SpotDays(int_ActiveRow, 1).Value

        lng_startdate = rng_StartDates(int_ActiveRow, 1).Value
        lng_MatDate = rng_MatDates(int_ActiveRow, 1).Value
        dbl_ActivePar = rng_ParsToBootstrap(int_ActiveRow, 1).Value

        dbl_NextActivePar = rng_ParsToBootstrap(int_ActiveRow + 1, 1).Value
        dbl_StartDF = Me.Lookup_Rate(lng_BuildDate, lng_startdate, "DF")

        ' Special case for FX curves with no 2D point
        If rng_MXSortTerms(int_ActiveRow, 1).Value = 1 And rng_MXSortTerms(int_ActiveRow + 1, 1).Value <> 2 Then
            bln_1DNo2D = True
            cal_ActiveCCY = cas_Calendars.Lookup_Calendar(map_Rules.Lookup_CCYCalendar(str_CCY))
            'lng_FXSpotDate = Date_WorkDay(lng_MatDate, int_ActiveSpotDays - 1, cal_ActiveCCY.HolDates, cal_ActiveCCY.weekends)
            lng_FXSpotDate = cyGetFXSpotDate(str_CCY, lng_BuildDate, dic_GlobalStaticInfo)
        Else
            bln_1DNo2D = False
        End If

        ' Store instrument object if not yet stored
        If dic_Swaps.Exists(int_ActiveRow) = False Then
            Select Case str_InstType
                Case "IRSWAP", "YTM_LONG"
                    With fld_ActiveParams_LegA
                        .ValueDate = lng_startdate
                        .Swapstart = lng_startdate
                        .GenerationRefPoint = lng_startdate
                        .IsFwdGeneration = True
                        .Term = str_ActiveTerm
                        .PExch_Start = False
                        .PExch_Intermediate = False
                        .PExch_End = False
                        .FloatEst = bln_EstFloat
                        .Notional = 1000000
                        .CCY = str_CCY
                        .index = "-"
                        .RateOrMargin = dbl_ActivePar
                        .PmtFreq = ReadSetting_LegA(rng_FlowsFreq(int_ActiveRow, 1).Value)
                        .Daycount = ReadSetting_LegA(rng_DayCounts(int_ActiveRow, 1).Value)
                        .BDC = rng_BDCs(int_ActiveRow, 1).Value
                        .EOM = bln_EOMRule
                        .PmtCal = rng_PmtCalendars(int_ActiveRow, 1).Value
                        .estcal = "-"
                        .Curve_Disc = str_DiscRateCurve_Swap_LegA
                        .Curve_Est = "-"
                    End With

                    With fld_ActiveParams_LegB
                        .ValueDate = lng_startdate
                        .Swapstart = lng_startdate
                        .GenerationRefPoint = lng_startdate
                        .IsFwdGeneration = True
                        .Term = str_ActiveTerm
                        .PExch_Start = False
                        .PExch_Intermediate = False
                        .PExch_End = False
                        .FloatEst = bln_EstFloat
                        .Notional = fld_ActiveParams_LegA.Notional
                        .CCY = fld_ActiveParams_LegA.CCY
                        .index = ReadSetting_LegB(rng_FlowsFreq(int_ActiveRow, 1).Value)
                        .RateOrMargin = 0
                        .PmtFreq = fld_ActiveParams_LegB.index
                        .Daycount = ReadSetting_LegB(rng_DayCounts(int_ActiveRow, 1).Value)
                        .BDC = fld_ActiveParams_LegA.BDC
                        .EOM = bln_EOMRule
                        .PmtCal = fld_ActiveParams_LegA.PmtCal
                        .estcal = rng_EstCalendars(int_ActiveRow, 1).Value
                        .Curve_Disc = str_DiscRateCurve_Swap_LegB
                        .Curve_Est = str_EstRateCurve_Swap_LegB
                    End With

                    With fld_ActiveParams_Swap
                        .Pay_LegA = True
                        .CCY_PnL = str_CCY
                        .LegA = fld_ActiveParams_LegA
                        .LegB = fld_ActiveParams_LegB
                    End With

                    bln_ActiveIsSwap = True
                Case "BASIS_SCCY"
                    With fld_ActiveParams_LegA
                        .ValueDate = lng_startdate
                        .Swapstart = lng_startdate
                        .GenerationRefPoint = lng_startdate
                        .IsFwdGeneration = True
                        .Term = str_ActiveTerm
                        .PExch_Start = False
                        .PExch_Intermediate = False
                        .PExch_End = False
                        .FloatEst = bln_EstFloat
                        .Notional = 1000000
                        .CCY = str_CCY
                        .index = ReadSetting_LegA_Index(rng_FlowsFreq(int_ActiveRow, 1).Value)
                        .RateOrMargin = 0
                        .PmtFreq = ReadSetting_LegA_PmtFreq(rng_FlowsFreq(int_ActiveRow, 1).Value)
                        .Daycount = ReadSetting_LegA(rng_DayCounts(int_ActiveRow, 1).Value)
                        .BDC = rng_BDCs(int_ActiveRow, 1).Value
                        .EOM = bln_EOMRule
                        .PmtCal = rng_PmtCalendars(int_ActiveRow, 1).Value
                        .estcal = rng_EstCalendars(int_ActiveRow, 1).Value
                        .Curve_Disc = str_DiscRateCurve_Basis_LegA
                        .Curve_Est = str_EstRateCurve_Basis_LegA
                    End With

                    With fld_ActiveParams_LegB
                        .ValueDate = lng_startdate
                        .Swapstart = lng_startdate
                        .GenerationRefPoint = lng_startdate
                        .IsFwdGeneration = True
                        .Term = str_ActiveTerm
                        .PExch_Start = False
                        .PExch_Intermediate = False
                        .PExch_End = False
                        .FloatEst = bln_EstFloat
                        .Notional = fld_ActiveParams_LegA.Notional
                        .CCY = fld_ActiveParams_LegA.CCY
                        .index = ReadSetting_LegB_Index(rng_FlowsFreq(int_ActiveRow, 1).Value)
                        .RateOrMargin = dbl_ActivePar / 100
                        .PmtFreq = ReadSetting_LegB_PmtFreq(rng_FlowsFreq(int_ActiveRow, 1).Value)
                        .Daycount = ReadSetting_LegB(rng_DayCounts(int_ActiveRow, 1).Value)
                        .BDC = fld_ActiveParams_LegA.BDC
                        .EOM = bln_EOMRule
                        .PmtCal = fld_ActiveParams_LegA.PmtCal
                        .estcal = fld_ActiveParams_LegA.estcal
                        .Curve_Disc = str_DiscRateCurve_Basis_LegB
                        .Curve_Est = str_EstRateCurve_Basis_LegB
                    End With

                    With fld_ActiveParams_Swap
                        .Pay_LegA = True
                        .CCY_PnL = str_CCY
                        .LegA = fld_ActiveParams_LegA
                        .LegB = fld_ActiveParams_LegB
                    End With

                    bln_ActiveIsSwap = True
                Case "IRBSWP"
                    ' For bootstrapping, able to use a single currency since FX is not a varying factor
                    With fld_ActiveParams_LegA
                        .ValueDate = lng_startdate
                        .Swapstart = lng_startdate
                        .GenerationRefPoint = lng_startdate
                        .IsFwdGeneration = True
                        .Term = str_ActiveTerm
                        .PExch_Start = False
                        .PExch_Intermediate = False
                        .PExch_End = True
                        .FloatEst = bln_EstFloat
                        .Notional = 1000000
                        .CCY = str_CCY
                        .index = ReadSetting_LegA_Index(rng_FlowsFreq(int_ActiveRow, 1).Value)
                        .RateOrMargin = 0
                        .PmtFreq = ReadSetting_LegA_PmtFreq(rng_FlowsFreq(int_ActiveRow, 1).Value)
                        .Daycount = ReadSetting_LegA(rng_DayCounts(int_ActiveRow, 1).Value)
                        .BDC = ReadSetting_LegA(rng_BDCs(int_ActiveRow, 1).Value)
                        .EOM = bln_EOMRule
                        .PmtCal = ReadSetting_LegA(rng_PmtCalendars(int_ActiveRow, 1).Value)
                        .estcal = ReadSetting_LegA(rng_EstCalendars(int_ActiveRow, 1).Value)
                        .Curve_Disc = str_DiscRateCurve_Basis_LegA
                        .Curve_Est = str_EstRateCurve_Basis_LegA
                    End With

                    With fld_ActiveParams_LegB
                        .ValueDate = lng_startdate
                        .Swapstart = lng_startdate
                        .GenerationRefPoint = lng_startdate
                        .IsFwdGeneration = True
                        .Term = str_ActiveTerm
                        .PExch_Start = False
                        .PExch_Intermediate = False
                        .PExch_End = True
                        .FloatEst = bln_EstFloat
                        .Notional = fld_ActiveParams_LegA.Notional
                        .CCY = str_CCY
                        .index = ReadSetting_LegB_Index(rng_FlowsFreq(int_ActiveRow, 1).Value)
                        .RateOrMargin = dbl_ActivePar / 100
                        .PmtFreq = ReadSetting_LegB_PmtFreq(rng_FlowsFreq(int_ActiveRow, 1).Value)
                        .Daycount = ReadSetting_LegB(rng_DayCounts(int_ActiveRow, 1).Value)
                        .BDC = ReadSetting_LegB(rng_BDCs(int_ActiveRow, 1).Value)
                        .EOM = bln_EOMRule
                        .PmtCal = ReadSetting_LegB(rng_PmtCalendars(int_ActiveRow, 1).Value)
                        .estcal = ReadSetting_LegB(rng_EstCalendars(int_ActiveRow, 1).Value)
                        .Curve_Disc = str_DiscRateCurve_Basis_LegB
                        .Curve_Est = str_EstRateCurve_Basis_LegB
                    End With

                    With fld_ActiveParams_Swap
                        .Pay_LegA = True
                        .CCY_PnL = str_CCY
                        .LegA = fld_ActiveParams_LegA
                        .LegB = fld_ActiveParams_LegB
                    End With

                    bln_ActiveIsSwap = True
                Case Else
                    bln_ActiveIsSwap = False
            End Select

            ' Add swap to dictionary and check if re-iteration is required
            If bln_ActiveIsSwap = True Then
                Set irs_Active = GetInst_IRS(fld_ActiveParams_Swap, , dic_GlobalStaticInfo)
                Call irs_Active.ReplaceCurveObject(str_CurveName, Me)
                Call dic_Swaps.Add(int_ActiveRow, irs_Active)

                If irs_Active.DependsOnFuture(str_CurveName, lng_MatDate) = True Then
                    Call dic_RowsToReIterate.Add(dic_RowsToReIterate.count + 1, int_ActiveRow)
                End If
            End If
        End If

        If dic_Swaps.Exists(int_ActiveRow) = True Then
            Set irs_Active = dic_Swaps(int_ActiveRow)
        Else
            Set irs_Active = Nothing
        End If

        ' Calculate static parameters
        Select Case str_InstType
            Case "IRBILL", "YTM_SHORT"
                dbl_FwdDF = Convert_SimpToDF(dbl_ActivePar, lng_startdate, lng_MatDate, str_ActiveDCC, "1Y")
            Case "YTM_DISC"
                'dbl_FwdDF = (1 - dbl_ActivePar / 100 * Calc_YearFrac(lng_StartDate, lng_MatDate, str_ActiveDCC, "1Y"))
                dbl_FwdDF = Convert_DiscToDF(dbl_ActivePar, lng_startdate, lng_MatDate, str_ActiveDCC, "1Y")
            Case "IRFUTB"
                Select Case str_FutInterp
                    Case "NII"  ' No intermediate interpolation
                        lng_UnAdjEndDate = date_addterm(lng_startdate, "3M", 1)
                        lng_UnAdjDays = lng_UnAdjEndDate - lng_startdate
                        dbl_FwdDF = 1 / (1 + dbl_ActivePar / 100 * lng_UnAdjDays / int_ActiveDayCount) ^ ((lng_MatDate - lng_startdate) / lng_UnAdjDays)
                    Case Else
                        dbl_FwdDF = Convert_SimpToDF(dbl_ActivePar, lng_startdate, lng_MatDate, str_ActiveDCC, "1Y")
                End Select
            Case "FXFWDPT"
                bln_ReqEstimatedFwd = False
                dbl_DF_USD = irc_BaseCurve.Lookup_Rate(lng_startdate, lng_MatDate, "DF")

                Select Case str_ActiveTerm
                    Case "1D"
                        Select Case int_ActiveSpotDays
                            Case 1
                                dbl_ActiveFXFwdStart = dbl_Spot - dbl_ActivePar / 10000
                                dbl_ActiveFXFwdEnd = dbl_Spot
                            Case Else
                                If bln_1DNo2D = True Then
                                    ' Forward between 1D and 2D is implied from the curve, iteration is then required
                                    bln_ReqEstimatedFwd = True
                                    dbl_DF_USDEstFwd = irc_BaseCurve.Lookup_Rate(lng_FXSpotDate, lng_MatDate, "DF")
                                Else
                                    dbl_ActiveFXFwdStart = dbl_Spot - (dbl_ActivePar + dbl_NextActivePar) / 10000
                                    dbl_ActiveFXFwdEnd = dbl_Spot - dbl_NextActivePar / 10000
                                End If
                        End Select
                    Case "2D"
                        Select Case int_ActiveSpotDays
                            Case 1
                                dbl_ActiveFXFwdStart = dbl_Spot
                                dbl_ActiveFXFwdEnd = dbl_Spot + dbl_ActivePar / 10000
                            Case Else
                                dbl_ActiveFXFwdStart = dbl_Spot - dbl_ActivePar / 10000
                                dbl_ActiveFXFwdEnd = dbl_Spot
                        End Select
                    Case Else
                        dbl_ActiveFXFwdStart = dbl_Spot
                        dbl_ActiveFXFwdEnd = dbl_Spot + dbl_ActivePar / 10000
                End Select
        End Select

        ' Solve zero rate
        Select Case str_InstType
            Case "IRBILL", "YTM_SHORT", "IRFUTB", "YTM_DISC"
                ' Store static parameters for secant solver
                Set dic_SecantParams = New Dictionary
                Call dic_SecantParams.Add("irc_Curve", Me)
                Call dic_SecantParams.Add("int_Index", int_ActiveRow)
                Call dic_SecantParams.Add("lng_StartDate", lng_startdate)
                Call dic_SecantParams.Add("lng_MatDate", lng_MatDate)
                Call dic_SecantParams.Add("dbl_FwdDF", dbl_FwdDF)

                ' Solve using secant method
                Set dic_SecantOutputs = New Dictionary
                Call Solve_Secant(ThisWorkbook, "SolverFuncXY_ZeroToDF_Deposit", dic_SecantParams, dbl_FallBackZero, _
                    dbl_FallBackZero + 1, dbl_FwdDF, 0.000000000000001, 50, dbl_FallBackZero, dic_SecantOutputs)
            Case "BASIS_SCCY", "IRBSWP", "IRSWAP", "YTM_LONG"
                ' Store static parameters for secant solver
                Set dic_SecantParams = New Dictionary
                Call dic_SecantParams.Add("irs_Swap", irs_Active)
                Call dic_SecantParams.Add("irc_Curve", Me)
                Call dic_SecantParams.Add("int_Index", int_ActiveRow)
                Call dic_SecantParams.Add("str_CurveName", str_CurveName)

                ' Solve using secant method
                Set dic_SecantOutputs = New Dictionary
                'Call Solve_Secant(ThisWorkbook, "SolverFuncXY_ZeroToMV_Swap", dic_secantParams, dbl_FallBackZero, _
                '   dbl_FallBackZero + 1, 0, 0.000000000000001, 50, dbl_FallBackZero, dic_SecantOutputs)
                ' SW : Modified to have the second intial guess to be the last solved rate.
                If int_ActiveRow = 1 Then
                    dbl_templogic = 1
                Else
                    dbl_templogic = dblArr_LiveZeros(int_ActiveRow - 1)
                End If
                Call Solve_Secant(ThisWorkbook, "SolverFuncXY_ZeroToMV_Swap", dic_SecantParams, dbl_FallBackZero, _
                    dbl_FallBackZero + dbl_templogic, 0, 0.000000000000001, 50, dbl_FallBackZero, dic_SecantOutputs)
            Case "FXFWDPT"
                ' Store static parameters for secant solver
                Set dic_SecantParams = New Dictionary
                Call dic_SecantParams.Add("irc_Curve", Me)
                Call dic_SecantParams.Add("int_Index", int_ActiveRow)
                Call dic_SecantParams.Add("lng_SpotDate", lng_FXSpotDate)
                Call dic_SecantParams.Add("lng_StartDate", lng_startdate)
                Call dic_SecantParams.Add("lng_MatDate", lng_MatDate)
                Call dic_SecantParams.Add("dbl_FXSpot", dbl_Spot)
                Call dic_SecantParams.Add("dbl_DF_USDEstFwd", dbl_DF_USDEstFwd)
                Call dic_SecantParams.Add("dbl_DF_USD", dbl_DF_USD)
                Call dic_SecantParams.Add("str_Quotation", str_Quotation)
                Call dic_SecantParams.Add("dbl_Par", dbl_ActivePar)
                Call dic_SecantParams.Add("bln_ReqEstimatedFwd", bln_ReqEstimatedFwd)
                Call dic_SecantParams.Add("dbl_FXFwd_Start", dbl_ActiveFXFwdStart)
                Call dic_SecantParams.Add("dbl_FXFwd_End", dbl_ActiveFXFwdEnd)

                ' Solve using secant method
                Set dic_SecantOutputs = New Dictionary
                Call Solve_Secant(ThisWorkbook, "SolverFuncXY_ZeroToDF_FXFwd", dic_SecantParams, dbl_FallBackZero, _
                    dbl_FallBackZero + 1, 0, 0.000000000000001, 50, dbl_FallBackZero, dic_SecantOutputs)
            Case "ZERO"
                dbl_EndDF = Math.Exp(-dbl_ActivePar / 100 * (lng_MatDate - lng_startdate) / int_ActiveDayCount) * dbl_StartDF
                Call Me.SetZeroRate(int_ActiveRow, Convert_DFToZero(dbl_EndDF, lng_BuildDate, lng_MatDate, str_InterpConv, ""))
        End Select

        ' Re-iterate if required, after any dependent future pillars have been evaluated
        If dic_RowsToReIterate.count > 0 Then
            If dic_RowsToReIterate(dic_RowsToReIterate.count) = int_ActiveRow - 1 Or int_ActiveRow = int_NumPoints Then
                int_ActiveNumReIterations = int_ActiveNumReIterations + 1

                ' Re-iterate until the maximum number of re-iterations, starting from first pillar requiring re-iteration
                If int_ActiveNumReIterations <= int_MaxReIterations Then
                    int_ActiveRow = dic_RowsToReIterate(1)
                Else
                    Call dic_RowsToReIterate.RemoveAll
                    int_ActiveRow = int_ActiveRow + 1
                    int_ActiveNumReIterations = 0
                End If
            Else
                int_ActiveRow = int_ActiveRow + 1
            End If
        Else
            int_ActiveRow = int_ActiveRow + 1
        End If
    Wend

    ' Sync rates
    rng_FinalZeros.Value = Convert_Array1Dto2D(dblArr_LiveZeros)
    Call StoreAsFinalState

    ' Copy for record-keeping purposes
    If bln_IsInitialLoad = True Then
        rng_OrigZeros.Value = rng_FinalZeros.Value
        Call dic_CurveStates.Remove(CurveState_IRC.original)
        Call dic_CurveStates.Add(CurveState_IRC.original, dblArr_LiveZeros)
    End If

    Call Action_SetAppState(fld_AppState_Orig)
End Sub


' ## METHODS - OUTPUT
Public Sub Output_ZeroRates(rng_OutputStart As Range)
    Dim int_RowCtr As Integer
    For int_RowCtr = 1 To int_NumPoints
        rng_OutputStart(int_RowCtr, 1).Value = str_CurveName
        rng_OutputStart(int_RowCtr, 2).Value = lng_BuildDate
        rng_OutputStart(int_RowCtr, 3).Value = CLng(rng_MatDates(int_RowCtr, 1).Value) - lng_BuildDate
        rng_OutputStart(int_RowCtr, 4).Value = dblArr_LiveZeros(int_RowCtr)
        rng_OutputStart(int_RowCtr, 5).Value = "Zero"
        rng_OutputStart(int_RowCtr, 6).Value = "Native"
    Next int_RowCtr
End Sub

Public Sub Output_ZeroRatesAtSelected(rng_OutputStart As Range, str_Pillars As String)
    ' ## Return zero rates interpolated at the specified pillars, defined in number of days
    Dim strArr_Pillars() As String: strArr_Pillars = Split(str_Pillars, "|")

    Dim int_RowCtr As Integer
    Dim lng_ActivePillar As Long
    Dim int_LBound As Integer: int_LBound = LBound(strArr_Pillars)
    Dim int_OutputRowNum As Integer
    For int_RowCtr = int_LBound To UBound(strArr_Pillars)
        int_OutputRowNum = int_RowCtr - int_LBound + 1
        lng_ActivePillar = CLng(strArr_Pillars(int_RowCtr))
        rng_OutputStart(int_OutputRowNum, 1).Value = str_CurveName
        rng_OutputStart(int_OutputRowNum, 2).Value = lng_BuildDate
        rng_OutputStart(int_OutputRowNum, 3).Value = lng_ActivePillar
        rng_OutputStart(int_OutputRowNum, 4).Value = Me.Lookup_Rate(lng_BuildDate, lng_BuildDate + lng_ActivePillar, "ZERO")
        rng_OutputStart(int_OutputRowNum, 5).Value = "Zero"
        rng_OutputStart(int_OutputRowNum, 6).Value = "Defined"
    Next int_RowCtr
End Sub

Public Sub Output_ParRates(rng_OutputStart As Range)
    Dim int_RowCtr As Integer
    For int_RowCtr = 1 To int_NumPoints
        rng_OutputStart(int_RowCtr, 1).Value = str_CurveName
        rng_OutputStart(int_RowCtr, 2).Value = lng_BuildDate
        rng_OutputStart(int_RowCtr, 3).Value = CLng(rng_MatDates(int_RowCtr, 1).Value) - lng_BuildDate
        rng_OutputStart(int_RowCtr, 4).Value = rng_OrigPars(int_RowCtr, 1).Value
        rng_OutputStart(int_RowCtr, 5).Value = "Par"
        rng_OutputStart(int_RowCtr, 6).Value = "Native"
    Next int_RowCtr
End Sub