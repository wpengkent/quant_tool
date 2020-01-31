Option Explicit


' ## MEMBER DATA
Private wks_Location As Worksheet
Private rng_TopLeft_Query As Range, rng_TopLeft_MXTerms As Range
Private rng_PID_Quote As Range, rng_PID_Interp As Range
Private rng_SpotDeltaPvt As Range, rng_IsSpotDeltaInterp As Range, rng_LessThan1MRule As Range

Private rng_MXTerms As Range
Private rng_OrigDeltas As Range, rng_OrigDates As Range, rng_OrigVols As Range
Private rng_Days_TopLeft As Range, rng_DeltasToShift_TopLeft As Range, rng_RelShifts_TopLeft As Range, rng_AbsShifts_TopLeft As Range
Private rng_FinalBuildDate As Range, rng_FinalDates As Range, rng_FinalDeltas As Range, rng_FinalVols As Range

Private dblLst_ATMVols_Final As Collection, dblLst_ATMVols_Orig As Collection
Private dic_Cache_DeltaPillars As Dictionary, dic_Cache_SmileVols As Dictionary
Private dic_Cache_PolyCoefs As Dictionary, dic_Cache_SmileSlopes As Dictionary
Private dic_FXCurveNames As Dictionary

' Dependent curves
Private fxs_Spots As Data_FXSpots, irc_SmileCCY As Data_IRCurve

' Dynamic variables
Private lng_ValDate As Long, dbl_VolShift_Sens As Double, dic_Vols_Orig As Dictionary, dic_Vols_Final As Dictionary

' Static values
Private dic_GlobalStaticInfo As Dictionary, igs_Generators As IRGeneratorSet, cas_Calendars As CalendarSet
Private map_Rules As MappingRules, cfg_Settings As ConfigSheet
Private lng_BuildDate_Orig As Long, str_CCY_Fgn As String, str_CCY_Dom As String
Private str_Calendar_DTV As String, str_Calendar_STD As String, str_FwdShifter As String, str_BackShifter As String
Private str_Interp_Time As String, str_Interp_Delta As String, bln_PID_Quote As Boolean, bln_PID_Interp As Boolean
Private str_CCY_Smile As String, str_RateCurve_Smile As String
Private dic_RelShifts As Dictionary, dic_AbsShifts As Dictionary
Private str_SolverMethod As String


' ## INITIALIZATION
Public Sub Initialize(wks_Input As Worksheet, bln_DataExists As Boolean, Optional dic_StaticInfoInput As Dictionary = Nothing)
    Set wks_Location = wks_Input

    ' Static info
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    Set igs_Generators = dic_GlobalStaticInfo(StaticInfoType.IRGeneratorSet)
    Set cas_Calendars = dic_GlobalStaticInfo(StaticInfoType.CalendarSet)
    Set cfg_Settings = dic_GlobalStaticInfo(StaticInfoType.ConfigSheet)
    Set map_Rules = dic_GlobalStaticInfo(StaticInfoType.MappingRules)
    Set dic_FXCurveNames = map_Rules.Dict_FXCurveNames
    Set dic_RelShifts = New Dictionary
    Set dic_AbsShifts = New Dictionary
    lng_ValDate = cfg_Settings.CurrentValDate

    ' Ranges required before data is populated
    Set rng_TopLeft_Query = wks_Location.Range("A7")
    Set rng_TopLeft_MXTerms = rng_TopLeft_Query.Offset(0, 9)

    If bln_DataExists = True Then
        Call StoreStaticValues
        Me.AssignRanges
    End If
End Sub


' ## PROPERTIES
Public Property Get CurveName() As String
    CurveName = str_CCY_Fgn & str_CCY_Dom
End Property

Public Property Get SmileCCY() As String
    SmileCCY = str_CCY_Smile
End Property

Private Property Get NumQueryRows() As Integer
    NumQueryRows = Examine_NumRows(rng_TopLeft_Query)
End Property

Public Property Get TypeCode() As CurveType
    TypeCode = CurveType.FXV
End Property

Public Property Get VolShift_Sens() As Double
    VolShift_Sens = dbl_VolShift_Sens
End Property

Public Property Let VolShift_Sens(dbl_Shift As Double)
    dbl_VolShift_Sens = dbl_Shift
End Property

Public Property Get ConfigSheet() As ConfigSheet
    Set ConfigSheet = cfg_Settings
End Property


' ## METHODS - SETUP
Private Sub StoreStaticValues()
    Dim rng_FirstParam As Range: Set rng_FirstParam = wks_Location.Range("A2")
    Dim int_ActiveCol As Integer: int_ActiveCol = 0
    lng_BuildDate_Orig = rng_FirstParam.Offset(0, int_ActiveCol).Value

    int_ActiveCol = int_ActiveCol + 1
    str_CCY_Fgn = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)

    int_ActiveCol = int_ActiveCol + 1
    str_CCY_Dom = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)

    int_ActiveCol = int_ActiveCol + 2
    str_Calendar_DTV = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)

    int_ActiveCol = int_ActiveCol + 1
    str_Calendar_STD = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)

    int_ActiveCol = int_ActiveCol + 1
    str_FwdShifter = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)

    int_ActiveCol = int_ActiveCol + 1
    str_BackShifter = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)

    int_ActiveCol = int_ActiveCol + 1
    str_Interp_Time = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)

    int_ActiveCol = int_ActiveCol + 1
    str_Interp_Delta = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)

    int_ActiveCol = int_ActiveCol + 1
    bln_PID_Quote = (UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value) = "YES")

    int_ActiveCol = int_ActiveCol + 1
    bln_PID_Interp = (UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value) = "YES")

    int_ActiveCol = int_ActiveCol + 1
    str_CCY_Smile = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)

    int_ActiveCol = int_ActiveCol + 4
    str_SolverMethod = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)

    ' Store dependent curves
    Set fxs_Spots = GetObject_FXSpots(True, dic_GlobalStaticInfo)  ' The instance actually used can be set via the property

    If dic_FXCurveNames.Exists(str_CCY_Smile) = True Then
        str_RateCurve_Smile = dic_FXCurveNames(str_CCY_Smile)
        Set irc_SmileCCY = GetObject_IRCurve(str_RateCurve_Smile, True, False)
    Else
        Debug.Assert False
    End If
End Sub

Public Sub LoadRates()
    With wks_Location
        'Dim lng_DataDate As Long: lng_DataDate = cyGetDataDate()
        Dim str_SQLCode As String

        Dim str_OptionalExclusions As String
        If .Range("D2").Value = "-" Then
            str_OptionalExclusions = ""
        Else
            str_OptionalExclusions = "AND SortTerm NOT IN (" & Replace(.Range("D2").Value, "|", ", ") & ") "
        End If
    End With

    ' Determine table name
    Debug.Assert map_Rules.Dict_SourceTables.Exists("FXVOL")
    Dim str_TableName As String: str_TableName = map_Rules.Dict_SourceTables("FXVOL")

    ' Query
    str_SQLCode = "SELECT [Data Date], Term, CallDelta, Rate " _
            & "FROM " & str_TableName _
            & " WHERE [Data Date] = #" & Convert_SQLDate(cfg_Settings.CurrentDataDate) & "# AND Fgn = '" & str_CCY_Fgn & "' " _
                & "AND Dom = '" & str_CCY_Dom & "' " & str_OptionalExclusions _
            & "ORDER BY CallDelta, SortTerm;"

    Me.ClearCurve
    Call Action_Query_Access(cfg_Settings.RatesDBPath, str_SQLCode, rng_TopLeft_Query)

    If NumQueryRows() > 0 Then
        Me.AssignRanges  ' Update ranges as number of points may have changed
        Call Me.GeneratePillarDates(False)

        rng_FinalDeltas.Value = rng_OrigDeltas.Value
        Me.Scen_ApplyBase
    End If
End Sub

Public Sub SetParams(rng_QueryParams As Range)
    With wks_Location
        .Range("A2").Value = cfg_Settings.CurrentBuildDate
        .Range("B2:Q2").Value = rng_QueryParams.Value
    End With

    Call StoreStaticValues
End Sub

Public Sub ClearCurve()
    With wks_Location
        Call Action_ClearBelow(.Range("A7"), 5)
        Call Action_ClearBelow(.Range("G7"), 4)
        Call Action_ClearBelow(.Range("L7"), 3)
    End With
End Sub

Public Sub AssignRanges()
    Dim int_NumRows As Integer: int_NumRows = NumQueryRows()
    If int_NumRows > 0 Then
        With wks_Location
            ' Parameters
            Set rng_PID_Quote = .Range("K2")
            Set rng_PID_Interp = rng_PID_Quote.Offset(0, 1)
            Set rng_SpotDeltaPvt = rng_PID_Interp.Offset(0, 2)
            Set rng_IsSpotDeltaInterp = rng_SpotDeltaPvt.Offset(0, 1)
            Set rng_LessThan1MRule = rng_IsSpotDeltaInterp.Offset(0, 1)
            Set rng_FinalBuildDate = rng_LessThan1MRule.Offset(0, 3)

            ' Data
            Set rng_MXTerms = .Range("B7").Resize(int_NumRows, 1)
            Set rng_OrigDeltas = rng_MXTerms.Offset(0, 1)
            Set rng_OrigVols = rng_OrigDeltas.Offset(0, 1)
            Set rng_OrigDates = rng_OrigVols.Offset(0, 1)
            Set rng_Days_TopLeft = .Range("G7")
            Set rng_DeltasToShift_TopLeft = rng_Days_TopLeft.Offset(0, 1)
            Set rng_RelShifts_TopLeft = rng_DeltasToShift_TopLeft.Offset(0, 1)
            Set rng_AbsShifts_TopLeft = rng_RelShifts_TopLeft.Offset(0, 1)
            Set rng_FinalDates = rng_AbsShifts_TopLeft.Offset(0, 2).Resize(int_NumRows)
            Set rng_FinalDeltas = rng_FinalDates.Offset(0, 1)
            Set rng_FinalVols = rng_FinalDeltas.Offset(0, 1)
        End With

        Set dic_Vols_Orig = Gather_InterpDict(rng_OrigDeltas, rng_OrigDates, rng_OrigVols)
        Call FillFinalVolsDict
    End If

    Call ResetCache_Lookups
End Sub


' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Sub SetValDate(lng_Input As Long)
    ' ## Set stored value date and initialize cache
    lng_ValDate = lng_Input
    Call ResetCache_Lookups
End Sub

Public Sub SetCurveState(str_curve As String, enu_State As CurveState_IRC, Optional int_PillarIndex As Integer = 0)
    ' ## For temporary zero shifts only, such as during a finite differencing calculation
    If irc_SmileCCY.CurveName = str_curve Then
        Call irc_SmileCCY.SetCurveState(enu_State, int_PillarIndex)
        Call ResetCache_Lookups
    End If
End Sub

Public Sub FillDependency_FXS(fxs_Input As Data_FXSpots)
    Set fxs_Spots = fxs_Input
End Sub

Public Sub FillDependency_IRC(dic_IRCurves As Dictionary)
    Set irc_SmileCCY = dic_IRCurves(str_RateCurve_Smile)
End Sub


' ## METHODS - LOOKUP
Public Function Lookup_ATMVol(lng_date As Long, Optional bln_GetOrig As Boolean = False) As Double
    Dim dbl_Output As Double
    ' Determine source
    Dim dic_ToUse As Dictionary
    Dim dblLst_VolsToUse As Collection, lngLst_DatesToUse As Collection
    If bln_GetOrig = True Then Set dic_ToUse = dic_Vols_Orig Else Set dic_ToUse = dic_Vols_Final

    Set lngLst_DatesToUse = dic_ToUse(50)(InterpAxis.Keys)
    Set dblLst_VolsToUse = dic_ToUse(50)(InterpAxis.Values)

    ' Apply spread, if any
    If dbl_VolShift_Sens <> 0 Then
        Dim dblLst_BeforeShift As Collection: Set dblLst_BeforeShift = dblLst_VolsToUse
        Set dblLst_VolsToUse = New Collection

        Dim int_ctr As Integer
        For int_ctr = 1 To dblLst_BeforeShift.count
            Call dblLst_VolsToUse.Add(dblLst_BeforeShift(int_ctr) + dbl_VolShift_Sens)
        Next int_ctr
    End If

    ' Look up vol
    Select Case str_Interp_Time
        Case "LIN", "LINEAR": dbl_Output = Interp_Lin(lngLst_DatesToUse, dblLst_VolsToUse, lng_date, True)
        Case "V2T": dbl_Output = Interp_V2t(lngLst_DatesToUse, dblLst_VolsToUse, rng_FinalBuildDate.Value, lng_date)
        Case "V2T_SDM"
            Dim lng_1YDate As Long: lng_1YDate = Lookup_MXTermDate("1Y")
            If lng_date <= lng_1YDate Then
                ' Short date V2t model
                dbl_Output = Interp_V2t_WeekDays(lngLst_DatesToUse, dblLst_VolsToUse, rng_FinalBuildDate.Value, lng_date)
            Else
                ' Default V2t model
                dbl_Output = Interp_V2t(lngLst_DatesToUse, dblLst_VolsToUse, rng_FinalBuildDate.Value, lng_date)
            End If
        Case Else: Debug.Assert False
    End Select

    If dbl_Output <= 0 Then dbl_Output = 0.000001
    Lookup_ATMVol = dbl_Output
End Function

'-------------------------------------------------------------------------------------------
' NAME:    Lookup_SmileVol
'
' PURPOSE: Bootstrap FX Smile
'
' NOTES:
'
' INPUT OPTIONS:
'
' MODIFIED:
'    31JAN2020 - KW - Add parameter dbl_fwd_input to allow user overwrite forward rate
'
'-------------------------------------------------------------------------------------------
Public Function Lookup_SmileVol(lng_LookupDateRaw As Long, dbl_Strike As Double, Optional bln_GetOrig As Boolean = False, _
    Optional bln_AxisRescale As Boolean = True, Optional bln_CalcAxisOnly As Boolean = False, Optional dbl_fwd_input As Double = 0) As Double
    ' Determine dates
    Dim lng_LookupDate As Long: lng_LookupDate = CleanLookupDate(lng_LookupDateRaw)
    Dim lng_LeftPillarDate As Long: lng_LeftPillarDate = Lookup_PrevOptionMat(lng_LookupDate, bln_GetOrig)
    If lng_LeftPillarDate = 0 Then lng_LeftPillarDate = lng_LookupDate
    Dim lng_RightPillarDate As Long: lng_RightPillarDate = Lookup_NextOptionMat(lng_LookupDate, bln_GetOrig)

    ' Handle case where valuation date is greater than the lookup date (could occur in theta scenario)
    If lng_ValDate >= lng_LookupDate Then
        Lookup_SmileVol = 0
        Exit Function
    End If

    ' Store period lengths used for each purpose
    Dim dbl_TimeToMat_Lookup As Double, dbl_TimeToMat_Left As Double, dbl_TimeToMat_Right As Double
    If bln_GetOrig = True Then
        dbl_TimeToMat_Lookup = calc_yearfrac(lng_BuildDate_Orig, lng_LookupDate, "ACT/365")
        dbl_TimeToMat_Left = calc_yearfrac(lng_BuildDate_Orig, lng_LeftPillarDate, "ACT/365")
        dbl_TimeToMat_Right = calc_yearfrac(lng_BuildDate_Orig, lng_RightPillarDate, "ACT/365")
    Else
        dbl_TimeToMat_Lookup = calc_yearfrac(lng_ValDate, lng_LookupDate, "ACT/365")
        dbl_TimeToMat_Left = calc_yearfrac(lng_ValDate, lng_LeftPillarDate, "ACT/365")
        dbl_TimeToMat_Right = calc_yearfrac(lng_ValDate, lng_RightPillarDate, "ACT/365")
    End If

    Dim bln_OnPillar As Boolean: bln_OnPillar = (lng_LookupDate = lng_LeftPillarDate Or lng_LookupDate = lng_RightPillarDate)
    Dim int_Index_Lookup As Integer, int_Index_Left As Integer, int_Index_Right As Integer
    Dim str_ActiveCacheKey_Lookup As String, str_ActiveCacheKey_Left As String, str_ActiveCacheKey_Right As String

    ' Determine forwards and vols
    Dim dblArr_LeftSmilePillars() As Double, dblArr_RightSmilePillars() As Double, dblArr_LookupSmilePillars() As Double
    Dim dbl_LeftATMVol As Double, dbl_RightATMVol As Double, dbl_LookupATMVol As Double
    Dim dbl_LeftPillarFwd As Double, dbl_RightPillarFwd As Double, dbl_LookupFwd As Double
    Dim enu_CurveState_Prev As CurveState_IRC, int_SensPillar_Prev As Integer

    Dim lngLst_PillarDates As Collection
    If bln_GetOrig = True Then
        Set lngLst_PillarDates = dic_Vols_Orig(50)(InterpAxis.Keys)
    Else
        Set lngLst_PillarDates = dic_Vols_Final(50)(InterpAxis.Keys)
    End If

    If bln_OnPillar = True Then
        int_Index_Lookup = Examine_FindIndex(lngLst_PillarDates, lng_LookupDate)
        Debug.Assert int_Index_Lookup <> -1
        dblArr_LookupSmilePillars = Lookup_SmilePillars(int_Index_Lookup, dbl_VolShift_Sens, bln_GetOrig)

        ' Check if delta scale already stored for this maturity
        str_ActiveCacheKey_Lookup = BuildCacheKey_DeltaScale(int_Index_Lookup, 3, dbl_VolShift_Sens, bln_GetOrig)
        If dic_Cache_DeltaPillars.Exists(str_ActiveCacheKey_Lookup) Then
            ' If purpose of calling function is to store the delta, then no need to continue
            If bln_CalcAxisOnly = True Then Exit Function
        Else
            ' If axis rescaling is off, need to have the original scale stored before looking it up
            If bln_AxisRescale = False Then
                enu_CurveState_Prev = irc_SmileCCY.CurveState  ' Note original values
                int_SensPillar_Prev = irc_SmileCCY.SensPillar
                Call irc_SmileCCY.SetCurveState(CurveState_IRC.original)  ' Disable shifts
                Call Me.Lookup_SmileVol(lng_LookupDateRaw, dbl_Strike, True, True, True)  ' Store original put delta scale
                Call irc_SmileCCY.SetCurveState(enu_CurveState_Prev, int_SensPillar_Prev)  ' Revert to previous settings
            End If
        End If
    Else
        int_Index_Left = Examine_FindIndex(lngLst_PillarDates, lng_LeftPillarDate)
        int_Index_Right = Examine_FindIndex(lngLst_PillarDates, lng_RightPillarDate)
        Debug.Assert int_Index_Left <> -1
        Debug.Assert int_Index_Right <> -1

        ' Check if delta scale already stored for this maturity.  Only left check required because right always follows it when being stored
        str_ActiveCacheKey_Left = BuildCacheKey_DeltaScale(int_Index_Left, 3, dbl_VolShift_Sens, bln_GetOrig)
        str_ActiveCacheKey_Right = BuildCacheKey_DeltaScale(int_Index_Right, 3, dbl_VolShift_Sens, bln_GetOrig)
        If dic_Cache_DeltaPillars.Exists(str_ActiveCacheKey_Left) And dic_Cache_DeltaPillars.Exists(str_ActiveCacheKey_Right) Then
            ' If purpose of calling function is to store the delta, then no need to continue
            If bln_CalcAxisOnly = True Then Exit Function
        Else
            ' If axis rescaling is off, need to have the original scale stored before looking it up
            If bln_AxisRescale = False Then
                enu_CurveState_Prev = irc_SmileCCY.CurveState  ' Note original values
                int_SensPillar_Prev = irc_SmileCCY.SensPillar
                Call irc_SmileCCY.SetCurveState(CurveState_IRC.original)  ' Disable shifts
                Call Me.Lookup_SmileVol(lng_LookupDateRaw, dbl_Strike, True, True, True)  ' Store original put delta scale
                Call irc_SmileCCY.SetCurveState(enu_CurveState_Prev, int_SensPillar_Prev)  ' Revert to previous settings
            End If
        End If

        ' Gather values required for smile lookup
        dblArr_LeftSmilePillars = Lookup_SmilePillars(int_Index_Left, dbl_VolShift_Sens, bln_GetOrig)
        dblArr_RightSmilePillars = Lookup_SmilePillars(int_Index_Right, dbl_VolShift_Sens, bln_GetOrig)
        dbl_LeftATMVol = dblArr_LeftSmilePillars(3)
        dbl_RightATMVol = dblArr_RightSmilePillars(3)
        dbl_LeftPillarFwd = fxs_Spots.Lookup_Fwd(str_CCY_Fgn, str_CCY_Dom, lng_LeftPillarDate, , bln_GetOrig)
        dbl_RightPillarFwd = fxs_Spots.Lookup_Fwd(str_CCY_Fgn, str_CCY_Dom, lng_RightPillarDate, , bln_GetOrig)
    End If

    dbl_LookupATMVol = Me.Lookup_ATMVol(lng_LookupDate, bln_GetOrig)

    If dbl_fwd_input = 0 Then
        dbl_LookupFwd = fxs_Spots.Lookup_Fwd(str_CCY_Fgn, str_CCY_Dom, lng_LookupDate, , bln_GetOrig)
    Else
        dbl_LookupFwd = dbl_fwd_input
    End IF

    ' Static parameters
    Dim lng_SpotDate As Long: lng_SpotDate = cyGetFXSpotDate(str_CCY_Smile, lng_ValDate, dic_GlobalStaticInfo)
    Dim lng_MatSpotDate As Long
    Dim lng_SpotDeltaLastDate As Long
    Dim dbl_LookupDF As Double: dbl_LookupDF = 1
    Dim dbl_LeftPillarDF As Double: dbl_LeftPillarDF = 1
    Dim dbl_RightPillarDF As Double: dbl_RightPillarDF = 1
    Dim lng_LeftPillarSpotDate As Long: lng_LeftPillarSpotDate = cyGetFXSpotDate(str_CCY_Smile, lng_LeftPillarDate, dic_GlobalStaticInfo)
    Dim lng_RightPillarSpotDate As Long: lng_RightPillarSpotDate = cyGetFXSpotDate(str_CCY_Smile, lng_RightPillarDate, dic_GlobalStaticInfo)

    ' Determine whether spot delta is used
    lng_MatSpotDate = cyGetFXSpotDate(str_CCY_Smile, lng_LookupDate, dic_GlobalStaticInfo)
    If rng_SpotDeltaPvt <> "-" Then
        lng_SpotDeltaLastDate = Lookup_MXTermDate(rng_SpotDeltaPvt.Value)

        If lng_LookupDate <= lng_SpotDeltaLastDate Then dbl_LookupDF = irc_SmileCCY.Lookup_Rate(lng_SpotDate, lng_MatSpotDate, "DF")
        If lng_LeftPillarDate <= lng_SpotDeltaLastDate Then dbl_LeftPillarDF = irc_SmileCCY.Lookup_Rate(lng_SpotDate, lng_LeftPillarSpotDate, "DF")
        If lng_RightPillarDate <= lng_SpotDeltaLastDate Then dbl_RightPillarDF = irc_SmileCCY.Lookup_Rate(lng_SpotDate, lng_RightPillarSpotDate, "DF")
    End If

    ' Determine put delta scale
    Dim dblArr_LeftDeltaPillars() As Double, dblArr_RightDeltaPillars() As Double, dblArr_LookupDeltaPillars() As Double
    Dim dbl_IterLeftSpread As Double, dbl_IterRightSpread As Double, dbl_ActiveStrike As Double
    Dim int_DeltaPillarCtr As Integer
    Dim bln_OriginalDeltaScale As Boolean: bln_OriginalDeltaScale = (bln_GetOrig = True Or bln_AxisRescale = False)
    Dim bln_IsSpotDeltaInterp As Boolean: bln_IsSpotDeltaInterp = (rng_IsSpotDeltaInterp.Value = "YES")

    If bln_OnPillar = True Then
        dblArr_LookupDeltaPillars = DeriveDeltaAxis(dblArr_LookupSmilePillars, int_Index_Lookup, bln_PID_Quote, bln_PID_Interp, _
            dbl_TimeToMat_Lookup, dbl_LookupFwd, dbl_LookupDF, bln_IsSpotDeltaInterp, bln_OriginalDeltaScale)
    Else
        dblArr_LeftDeltaPillars = DeriveDeltaAxis(dblArr_LeftSmilePillars, int_Index_Left, bln_PID_Quote, bln_PID_Interp, _
            dbl_TimeToMat_Left, dbl_LeftPillarFwd, dbl_LeftPillarDF, bln_IsSpotDeltaInterp, bln_OriginalDeltaScale)

        dblArr_RightDeltaPillars = DeriveDeltaAxis(dblArr_RightSmilePillars, int_Index_Right, bln_PID_Quote, bln_PID_Interp, _
            dbl_TimeToMat_Right, dbl_RightPillarFwd, dbl_RightPillarDF, bln_IsSpotDeltaInterp, bln_OriginalDeltaScale)
    End If

    If bln_CalcAxisOnly = True Then Exit Function

    ' Store static parameters for secant solver
    Dim dic_StaticParams As Dictionary: Set dic_StaticParams = New Dictionary
    Call dic_StaticParams.Add("dbl_LookupFwd", dbl_LookupFwd)
    Call dic_StaticParams.Add("dbl_TimeToMat_Lookup", dbl_TimeToMat_Lookup)
    Call dic_StaticParams.Add("lng_LookupDate", lng_LookupDate)
    Call dic_StaticParams.Add("dbl_Strike", dbl_Strike)
    Call dic_StaticParams.Add("bln_PID_Interp", bln_PID_Interp)
    Call dic_StaticParams.Add("str_Interp_Delta", str_Interp_Delta)
    Call dic_StaticParams.Add("bln_OnPillar", bln_OnPillar)
    Call dic_StaticParams.Add("dbl_TimeToMat_Left", dbl_TimeToMat_Left)
    Call dic_StaticParams.Add("dbl_TimeToMat_Right", dbl_TimeToMat_Right)
    Call dic_StaticParams.Add("lng_LeftPillarDate", lng_LeftPillarDate)
    Call dic_StaticParams.Add("lng_RightPillarDate", lng_RightPillarDate)
    Call dic_StaticParams.Add("dblArr_LookupDeltaPillars", dblArr_LookupDeltaPillars)
    Call dic_StaticParams.Add("dblArr_LeftDeltaPillars", dblArr_LeftDeltaPillars)
    Call dic_StaticParams.Add("dblArr_RightDeltaPillars", dblArr_RightDeltaPillars)
    Call dic_StaticParams.Add("dblArr_LookupSmilePillars", dblArr_LookupSmilePillars)
    Call dic_StaticParams.Add("dblArr_LeftSmilePillars", dblArr_LeftSmilePillars)
    Call dic_StaticParams.Add("dblArr_RightSmilePillars", dblArr_RightSmilePillars)
    Call dic_StaticParams.Add("dbl_LookupDF", dbl_LookupDF)
    Call dic_StaticParams.Add("dbl_LeftPillarDF", dbl_LeftPillarDF)
    Call dic_StaticParams.Add("dbl_RightPillarDF", dbl_RightPillarDF)
    Call dic_StaticParams.Add("dbl_LookupATMVol", dbl_LookupATMVol)
    Call dic_StaticParams.Add("dbl_LeftATMVol", dbl_LeftATMVol)
    Call dic_StaticParams.Add("dbl_RightATMVol", dbl_RightATMVol)
    Call dic_StaticParams.Add("bln_IsSpotDeltaInterp", bln_IsSpotDeltaInterp)

    If str_Interp_Delta = "POLYNOMIAL" Then
        Dim dblArr_ActivePolyCoefs() As Double
        Dim str_CacheKey As String
        If bln_OnPillar = True Then
            ' Retrieve or derive coefficients
            str_CacheKey = int_Index_Lookup & "|" & dbl_VolShift_Sens & "|" & bln_GetOrig
            If dic_Cache_PolyCoefs.Exists(str_CacheKey) = True Then
                dblArr_ActivePolyCoefs = dic_Cache_PolyCoefs(str_CacheKey)
            Else
                dblArr_ActivePolyCoefs = Calc_PolyCoefs(dblArr_LookupDeltaPillars, dblArr_LookupSmilePillars)
                Call dic_Cache_PolyCoefs.Add(str_CacheKey, dblArr_ActivePolyCoefs)
            End If

            ' Store found coefficients
            Call dic_StaticParams.Add("dblArr_LookupPolyCoefs", dblArr_ActivePolyCoefs)
        Else
            ' Retrieve or derive coefficients
            str_CacheKey = int_Index_Left & "|" & dbl_VolShift_Sens & "|" & bln_GetOrig
            If dic_Cache_PolyCoefs.Exists(str_CacheKey) = True Then
                dblArr_ActivePolyCoefs = dic_Cache_PolyCoefs(str_CacheKey)
            Else
                dblArr_ActivePolyCoefs = Calc_PolyCoefs(dblArr_LeftDeltaPillars, dblArr_LeftSmilePillars)
                Call dic_Cache_PolyCoefs.Add(str_CacheKey, dblArr_ActivePolyCoefs)
            End If

            ' Store found coefficients
            Call dic_StaticParams.Add("dblArr_LeftPolyCoefs", dblArr_ActivePolyCoefs)

            ' Retrieve or derive coefficients
            str_CacheKey = int_Index_Right & "|" & dbl_VolShift_Sens & "|" & bln_GetOrig
            If dic_Cache_PolyCoefs.Exists(str_CacheKey) = True Then
                dblArr_ActivePolyCoefs = dic_Cache_PolyCoefs(str_CacheKey)
            Else
                dblArr_ActivePolyCoefs = Calc_PolyCoefs(dblArr_RightDeltaPillars, dblArr_RightSmilePillars)
                Call dic_Cache_PolyCoefs.Add(str_CacheKey, dblArr_ActivePolyCoefs)
            End If

            ' Store found coefficients
            Call dic_StaticParams.Add("dblArr_RightPolyCoefs", dblArr_ActivePolyCoefs)
        End If
    End If

    ' Solve using secant method
    Dim dic_Outputs As Dictionary: Set dic_Outputs = New Dictionary
    Dim dbl_FinalVol As Double

    'str_SolverMethod = "FIXEDPT"
    Dim dbl_ShockForSeed As Double

SOLVER:

    Select Case str_SolverMethod
        Case "SECANT"
            'dbl_FinalVol = Solve_Secant(ThisWorkbook, "SolverFuncXY_VolToStrike", dic_StaticParams, dbl_LookupATMVol, _
                dbl_LookupATMVol + 1, dbl_Strike, 0.0000000001, 50, -1, dic_Outputs)

            If dbl_LookupATMVol >= 1 Then
                dbl_ShockForSeed = 1
            Else
                dbl_ShockForSeed = 0.01
            End If

            dbl_FinalVol = Solve_Secant(ThisWorkbook, "SolverFuncXY_VolToStrike", dic_StaticParams, dbl_LookupATMVol, _
                dbl_LookupATMVol + dbl_ShockForSeed, dbl_Strike, 0.0000000001, 50, -1, dic_Outputs)

            ' Check secant method actually produced the correct solution
            If dic_Outputs("Solvable") = True Then
                dbl_FinalVol = Solve_FixedPt(ThisWorkbook, "SolverFuncXX_FXSmileIteration", dic_StaticParams, dbl_FinalVol, _
                    0.0001, 1, -1, dic_Outputs)
            End If

            ' Output final solution
            If dic_Outputs("Solvable") = False Then
                ' Try fixed point iteration, if fails output -1 as the vol
                Debug.Print "Secant method did not find a solution.  Falling back to fixed point iteration"
                dbl_FinalVol = Solve_FixedPt(ThisWorkbook, "SolverFuncXX_FXSmileIteration", dic_StaticParams, dbl_LookupATMVol, _
                    0.0000000001, 50, -1, dic_Outputs)
                Debug.Assert dbl_FinalVol > 0
            End If
        Case "FIXEDPT"
            dbl_FinalVol = Solve_FixedPt(ThisWorkbook, "SolverFuncXX_FXSmileIteration", dic_StaticParams, dbl_LookupATMVol, _
                0.0000000001, 50, -1, dic_Outputs)

            ' Output final solution
            If dic_Outputs("Solvable") = False Then
                ' Try fixed point iteration, if fails output -1 as the vol
                Debug.Print "Fixed point iteration did not find a solution.  Falling back to secant method"
                dbl_FinalVol = Solve_Secant(ThisWorkbook, "SolverFuncXY_VolToStrike", dic_StaticParams, dbl_LookupATMVol, _
                    dbl_LookupATMVol + 1, dbl_Strike, 0.0000000001, 50, -1, dic_Outputs)


                If dbl_FinalVol <= 0 Then
                    str_SolverMethod = "SECANT"
                    GoTo SOLVER
                End If

                Debug.Assert dbl_FinalVol > 0
            End If
        Case Else: Debug.Assert False
    End Select

    ' Calculate and store smile slope if not already stored
    str_CacheKey = lng_LookupDate & "|" & dbl_Strike & "|" & dbl_VolShift_Sens
    If dic_Cache_SmileSlopes.Exists(str_CacheKey) = False Then
        ' Convert to forward call delta scale (no PID)
        Dim dblArr_CallDeltas(1 To 5) As Double, dblArr_ActiveSmilePillars() As Double
        dblArr_CallDeltas(1) = 10
        dblArr_CallDeltas(2) = 25
        dblArr_CallDeltas(3) = 50
        dblArr_CallDeltas(4) = 75
        dblArr_CallDeltas(5) = 90

        If bln_OnPillar = True Then
            Call dic_StaticParams.Remove("dblArr_LookupDeltaPillars")
            Call dic_StaticParams.Add("dblArr_LookupDeltaPillars", dblArr_CallDeltas)
            dblArr_ActiveSmilePillars = Convert_Reverse(dblArr_LookupSmilePillars)
            Call dic_StaticParams.Remove("dblArr_LookupSmilePillars")
            Call dic_StaticParams.Add("dblArr_LookupSmilePillars", dblArr_ActiveSmilePillars)

            dblArr_ActivePolyCoefs = Calc_PolyCoefs(dblArr_CallDeltas, dblArr_ActiveSmilePillars)
            Call dic_StaticParams.Remove("dblArr_LookupPolyCoefs")
            Call dic_StaticParams.Add("dblArr_LookupPolyCoefs", dblArr_ActivePolyCoefs)
        Else
            Call dic_StaticParams.Remove("dblArr_LeftDeltaPillars")
            Call dic_StaticParams.Add("dblArr_LeftDeltaPillars", dblArr_CallDeltas)
            dblArr_ActiveSmilePillars = Convert_Reverse(dblArr_LeftSmilePillars)
            Call dic_StaticParams.Remove("dblArr_LeftSmilePillars")
            Call dic_StaticParams.Add("dblArr_LeftSmilePillars", dblArr_ActiveSmilePillars)

            dblArr_ActivePolyCoefs = Calc_PolyCoefs(dblArr_CallDeltas, dblArr_ActiveSmilePillars)
            Call dic_StaticParams.Remove("dblArr_LeftPolyCoefs")
            Call dic_StaticParams.Add("dblArr_LeftPolyCoefs", dblArr_ActivePolyCoefs)

            Call dic_StaticParams.Remove("dblArr_RightDeltaPillars")
            Call dic_StaticParams.Add("dblArr_RightDeltaPillars", dblArr_CallDeltas)
            dblArr_ActiveSmilePillars = Convert_Reverse(dblArr_RightSmilePillars)
            Call dic_StaticParams.Remove("dblArr_RightSmilePillars")
            Call dic_StaticParams.Add("dblArr_RightSmilePillars", dblArr_ActiveSmilePillars)

            dblArr_ActivePolyCoefs = Calc_PolyCoefs(dblArr_CallDeltas, dblArr_ActiveSmilePillars)
            Call dic_StaticParams.Remove("dblArr_RightPolyCoefs")
            Call dic_StaticParams.Add("dblArr_RightPolyCoefs", dblArr_ActivePolyCoefs)
        End If

        Dim dbl_FwdCallDelta As Double, dbl_SmileSlope As Double
        dbl_FwdCallDelta = Calc_BS_FwdDelta(OptionDirection.CallOpt, dbl_LookupFwd, dbl_Strike, dbl_TimeToMat_Lookup, dbl_FinalVol, False)
        dbl_SmileSlope = FXSmileSlope(dbl_FwdCallDelta, str_Interp_Delta, dic_StaticParams)
        Call dic_Cache_SmileSlopes.Add(str_CacheKey, dbl_SmileSlope)
    End If

    If dbl_FinalVol <= 0 Then dbl_FinalVol = 0.000001
    Lookup_SmileVol = dbl_FinalVol
End Function

Public Function Lookup_SmileSlope(lng_LookupDateRaw As Long, dbl_Strike As Double) As Double
    ' ## Read stored smile slope, which should have been stored by vol calculation.  If absent, execute the vol calculation to store the slope as a byproduct
    Dim dbl_Output As Double
    Dim lng_LookupDate As Long: lng_LookupDate = CleanLookupDate(lng_LookupDateRaw)
    Dim str_CacheKey As String: str_CacheKey = lng_LookupDate & "|" & dbl_Strike
    If dic_Cache_SmileSlopes.Exists(str_CacheKey) Then
        dbl_Output = dic_Cache_SmileSlopes(str_CacheKey)
    Else
        Call Me.Lookup_SmileVol(lng_LookupDateRaw, dbl_Strike)
        dbl_Output = dic_Cache_SmileSlopes(str_CacheKey)
    End If

    Lookup_SmileSlope = dbl_Output
End Function

Private Function Lookup_SmilePillars(int_Index As Integer, Optional dbl_spread As Double = 0, Optional bln_GetOrig As Boolean = False) As Double()
    ' ## Returns smile volatilities in format { 10P, 25P, ATM, 25C, 10C }
    Dim dblArr_Output() As Double
    Dim str_CacheKey As String: str_CacheKey = int_Index & dbl_spread & bln_GetOrig
    Dim int_ctr As Integer

    If dic_Cache_SmileVols.Exists(str_CacheKey) = True Then
        ' Read from cache
        dblArr_Output = dic_Cache_SmileVols(str_CacheKey)
    Else
        ' Gather and store vols in cache.  Reverse order because interpolating in put delta mode
        Dim dic_ToUse As Dictionary
        If bln_GetOrig = True Then Set dic_ToUse = dic_Vols_Orig Else Set dic_ToUse = dic_Vols_Final

        ReDim dblArr_Output(1 To 5) As Double
        dblArr_Output(1) = dic_ToUse(90)(InterpAxis.Values)(int_Index)
        dblArr_Output(2) = dic_ToUse(75)(InterpAxis.Values)(int_Index)
        dblArr_Output(3) = dic_ToUse(50)(InterpAxis.Values)(int_Index)
        dblArr_Output(4) = dic_ToUse(25)(InterpAxis.Values)(int_Index)
        dblArr_Output(5) = dic_ToUse(10)(InterpAxis.Values)(int_Index)

        ' Prevent negative vols
        For int_ctr = 1 To 5
            If dblArr_Output(int_ctr) <= 0 Then dblArr_Output(int_ctr) = 0.000001
        Next int_ctr

        Call dic_Cache_SmileVols.Add(str_CacheKey, dblArr_Output)
    End If

    ' Apply spread, if any.  Don't apply spread when getting base case vol for disabled rescaling case
    If dbl_spread <> 0 And bln_GetOrig = False Then
        For int_ctr = LBound(dblArr_Output) To UBound(dblArr_Output)
            dblArr_Output(int_ctr) = dblArr_Output(int_ctr) + dbl_spread
        Next int_ctr
    End If

    Lookup_SmilePillars = dblArr_Output
End Function

Private Function Lookup_PutDeltaPillars(dblArr_SmilePillars() As Double, int_Index As Integer, bln_PID As Boolean, _
    dbl_Fwd As Double, dbl_SmileDF As Double, dbl_VolSpread As Double, bln_FlatSmile As Boolean, bln_IsSpotDeltaInterp As Boolean, _
    Optional bln_GetOrig As Boolean = False) As Double()
    ' ## Returns put delta %, in some cases call deltas need conversion.  If spot delta is in use, the output is a spot delta
    ' ## Call to put delta conversion requires forward deltas
    Dim dblArr_Output() As Double: ReDim dblArr_Output(1 To 5) As Double
    dblArr_Output(1) = 10
    dblArr_Output(2) = 25

    If bln_FlatSmile = True Then
        ' If smile is flat, delta pillars are inconsequential
        dblArr_Output(3) = 50
        dblArr_Output(4) = 75
        dblArr_Output(5) = 90
    ElseIf bln_PID = True Then
        Dim lng_PillarDate As Long
        Dim dic_ToUse As Dictionary
        If bln_GetOrig = True Then Set dic_ToUse = dic_Vols_Orig Else Set dic_ToUse = dic_Vols_Final
        lng_PillarDate = dic_ToUse(50)(InterpAxis.Keys)(int_Index)

        Dim dbl_TimeToMat As Double: dbl_TimeToMat = calc_yearfrac(lng_ValDate, lng_PillarDate, "ACT/365")
        Dim dbl_ATMvol As Double: dbl_ATMvol = dblArr_SmilePillars(3)

        dblArr_Output(3) = 50 * Math.Exp(-0.5 * (dbl_ATMvol / 100) ^ 2 * dbl_TimeToMat) * dbl_SmileDF
        dblArr_Output(4) = Calc_BS_CallToPutDelta(25 / dbl_SmileDF, dbl_Fwd, dbl_TimeToMat, dblArr_SmilePillars(4), dbl_ATMvol, True) * dbl_SmileDF
        dblArr_Output(5) = Calc_BS_CallToPutDelta(10 / dbl_SmileDF, dbl_Fwd, dbl_TimeToMat, dblArr_SmilePillars(5), dbl_ATMvol, True) * dbl_SmileDF

        ' Cause function to return an error if deltas are invalid
        Debug.Assert (dblArr_Output(4) > 0 And dblArr_Output(5) > 0)
    Else
        dblArr_Output(3) = 50 * dbl_SmileDF
        dblArr_Output(4) = 100 * dbl_SmileDF - 25
        dblArr_Output(5) = 100 * dbl_SmileDF - 10
    End If

    ' Convert pillars to fwd delta if these are quoted in spot delta but interpolated in fwd delta
    If bln_IsSpotDeltaInterp = False Then
        Dim lng_Ctr As Long
        For lng_Ctr = 1 To 5
            dblArr_Output(lng_Ctr) = dblArr_Output(lng_Ctr) / dbl_SmileDF
        Next lng_Ctr
    End If

    Lookup_PutDeltaPillars = dblArr_Output
End Function

Private Function Lookup_NextOptionMat(lng_RefDate As Long, Optional bln_GetOrig As Boolean = False) As Long
    ' ## Get the option maturity date on or immediately after the reference date
    Lookup_NextOptionMat = Lookup_OptionMat(lng_RefDate, 0, bln_GetOrig)
End Function

Private Function Lookup_PrevOptionMat(lng_RefDate As Long, Optional bln_GetOrig As Boolean = False) As Long
    ' ## Get the option maturity date immediately before the reference date
    Lookup_PrevOptionMat = Lookup_OptionMat(lng_RefDate, -1, bln_GetOrig)
End Function

Private Function Lookup_OptionMat(lng_RefDate As Long, int_PillarOffset As Integer, Optional bln_GetOrig As Boolean = False) As Long
    ' ## Use int_PillarOffset = -1 to get previous option maturity
    ' ## Returns zero if requested contract offset is not available

    Dim lng_FoundDate As Long
    Dim lngLst_PillarDates_ToUse As Collection
    If bln_GetOrig = True Then
        Set lngLst_PillarDates_ToUse = dic_Vols_Orig(50)(InterpAxis.Keys)
    Else
        Set lngLst_PillarDates_ToUse = dic_Vols_Final(50)(InterpAxis.Keys)
    End If

    Dim int_ctr As Integer
    For int_ctr = 1 To lngLst_PillarDates_ToUse.count
        lng_FoundDate = lngLst_PillarDates_ToUse(int_ctr)

        If lng_RefDate <= lng_FoundDate Then
            If int_ctr + int_PillarOffset >= 1 Then
                Lookup_OptionMat = lngLst_PillarDates_ToUse(int_ctr + int_PillarOffset)
            Else
                Lookup_OptionMat = 0
            End If

            Exit Function
        End If
    Next int_ctr
End Function

Private Function Lookup_MXTermDate(str_Term As String) As Long
    ' ## Returns final maturity date corresponding to the specified term
    Dim int_FoundIndex As Integer: int_FoundIndex = Examine_FindIndex(Convert_RangeToList(rng_MXTerms), str_Term)
    Debug.Assert int_FoundIndex <> -1
    Lookup_MXTermDate = rng_FinalDates(int_FoundIndex, 1).Value
End Function

Public Function Lookup_BucketWeights(lng_RefDate As Long, Optional bln_GetOrig As Boolean = False) As Variant()
    ' ## Returns the proportion of the total vega to be assigned to the adjacent pillars straddling the specified option maturity date
    ' ## Return format is a array containing: Left date, left weight, right date, right weight
    Dim varArr_Output(1 To 1, 1 To 4) As Variant

    ' Obtain pillar dates
    Dim lng_Pillar_Left As Long: lng_Pillar_Left = Lookup_PrevOptionMat(lng_RefDate, bln_GetOrig)
    Dim lng_Pillar_Right As Long: lng_Pillar_Right = Lookup_NextOptionMat(lng_RefDate, bln_GetOrig)
    If lng_Pillar_Left = 0 Then varArr_Output(1, 1) = "-" Else varArr_Output(1, 1) = lng_Pillar_Left
    varArr_Output(1, 3) = lng_Pillar_Right

    If lng_Pillar_Left = 0 Or lng_RefDate = lng_Pillar_Right Then
        varArr_Output(1, 2) = 0
        varArr_Output(1, 4) = 1
    Else
        Dim lng_BuildDate As Long
        If bln_GetOrig = True Then lng_BuildDate = lng_BuildDate_Orig Else lng_BuildDate = rng_FinalBuildDate.Value

        ' Obtain #days
        Dim lng_Days_Left As Long: lng_Days_Left = lng_Pillar_Left - lng_BuildDate
        Dim lng_Days_Right As Long: lng_Days_Right = lng_Pillar_Right - lng_BuildDate
        Dim lng_Days_Lookup As Long: lng_Days_Lookup = lng_RefDate - lng_BuildDate

        ' Obtain ATM vols
        Dim dbl_Vol_Left As Double: dbl_Vol_Left = Me.Lookup_ATMVol(lng_Pillar_Left, bln_GetOrig)
        Dim dbl_Vol_Right As Double: dbl_Vol_Right = Me.Lookup_ATMVol(lng_Pillar_Right, bln_GetOrig)
        Dim dbl_Vol_Lookup As Double: dbl_Vol_Lookup = Me.Lookup_ATMVol(lng_RefDate, bln_GetOrig)

        ' Evaluate formula for V2t weightings
        varArr_Output(1, 4) = (lng_Days_Lookup - lng_Days_Left) / (lng_Days_Right - lng_Days_Left) _
            * (dbl_Vol_Right * lng_Days_Right) / (dbl_Vol_Lookup * lng_Days_Lookup)
        varArr_Output(1, 2) = 1 - varArr_Output(1, 4)  ' ## Disputing whether this is correct or not
    End If

    Lookup_BucketWeights = varArr_Output
End Function


' ## METHODS - SUPPORT
Private Function Gather_InterpDict(rng_Deltas As Range, rng_MatDates As Range, rng_Vols As Range) As Dictionary
    ' ## Reads values from sheet into dictionary containing lists for interpolation
    ' ## Assumes data is sorted in ascending order by maturity
    Dim dic_output As New Dictionary
    Dim int_NumRows As Integer: int_NumRows = NumQueryRows()
    Dim intArr_Deltas() As Variant: intArr_Deltas = rng_Deltas.Value
    Dim lngArr_MatDates() As Variant: lngArr_MatDates = rng_MatDates.Value
    Dim dblArr_Vols() As Variant: dblArr_Vols = rng_Vols.Value2

    Dim int_ctr As Integer, int_ActiveDelta As Integer
    Dim varArr_ActiveSet() As Variant, lngLst_ActiveMatDates As Collection, dblLst_ActiveVols As Collection
    For int_ctr = 1 To int_NumRows
        int_ActiveDelta = intArr_Deltas(int_ctr, 1)

        ' Find collections containing dates and vols
        If dic_output.Exists(int_ActiveDelta) Then
            varArr_ActiveSet = dic_output(int_ActiveDelta)
            Set lngLst_ActiveMatDates = varArr_ActiveSet(InterpAxis.Keys)
            Set dblLst_ActiveVols = varArr_ActiveSet(InterpAxis.Values)
        Else
            ReDim varArr_ActiveSet(1 To 2) As Variant
            Set lngLst_ActiveMatDates = New Collection
            Set dblLst_ActiveVols = New Collection
            Set varArr_ActiveSet(InterpAxis.Keys) = lngLst_ActiveMatDates
            Set varArr_ActiveSet(2) = dblLst_ActiveVols
            Call dic_output.Add(int_ActiveDelta, varArr_ActiveSet)
        End If

        Call lngLst_ActiveMatDates.Add(lngArr_MatDates(int_ctr, 1))
        Call dblLst_ActiveVols.Add(dblArr_Vols(int_ctr, 1))
    Next int_ctr

    Set Gather_InterpDict = dic_output
End Function

Private Function DeriveDeltaAxis(dblArr_SmilePillars() As Double, int_Index_MatPillar As Integer, bln_PID_Quote As Boolean, _
    bln_PID_Interp As Boolean, dbl_TimeToMat As Double, dbl_Fwd As Double, dbl_DF As Double, bln_IsSpotDeltaInterp As Boolean, _
    bln_OriginalDeltaScale As Boolean) As Double()
    ' ## Returns set of deltas for a given maturity pillar.  Uses cached value if available
    ' Convert to non-PID put delta pillars by first computing the common strike
    Dim int_LBound As Integer: int_LBound = LBound(dblArr_SmilePillars)
    Dim int_UBound As Integer: int_UBound = UBound(dblArr_SmilePillars)
    Dim dblArr_DeltaPillars() As Double: ReDim dblArr_DeltaPillars(int_LBound To int_UBound) As Double
    Dim bln_FlatSmile As Boolean: bln_FlatSmile = Examine_IsUniform(dblArr_SmilePillars)  ' If smile turns out to be flat, skip some inconsequential steps
    Dim int_ctr As Integer, str_ActiveCacheKey As String

    For int_ctr = int_LBound To int_UBound
        str_ActiveCacheKey = BuildCacheKey_DeltaScale(int_Index_MatPillar, int_ctr, dbl_VolShift_Sens, bln_OriginalDeltaScale)
        If dic_Cache_DeltaPillars.Exists(str_ActiveCacheKey) Then
            ' Use stored delta scale
            dblArr_DeltaPillars(int_ctr) = dic_Cache_DeltaPillars(str_ActiveCacheKey)
        Else
            ' Only need to look up the raw put delta pillars the first time as it is constant
            If int_ctr = int_LBound Then
                dblArr_DeltaPillars = Lookup_PutDeltaPillars(dblArr_SmilePillars, int_Index_MatPillar, bln_PID_Quote, dbl_Fwd, _
                    dbl_DF, dbl_VolShift_Sens, bln_FlatSmile, bln_IsSpotDeltaInterp, bln_OriginalDeltaScale)
            End If

            ' Build and store delta scale
            If bln_PID_Quote = True And bln_PID_Interp = False And bln_FlatSmile = False Then
                If bln_IsSpotDeltaInterp = True Then
                    ' Pass the forward deltas to the conversion function but interpolate based on spot delta
                    dblArr_DeltaPillars(int_ctr) = Calc_BS_RemovePID(OptionDirection.PutOpt, dblArr_DeltaPillars(int_ctr) / dbl_DF, _
                        dbl_Fwd, dbl_TimeToMat, dblArr_SmilePillars(int_ctr)) * dbl_DF
                Else
                    dblArr_DeltaPillars(int_ctr) = Calc_BS_RemovePID(OptionDirection.PutOpt, dblArr_DeltaPillars(int_ctr), _
                        dbl_Fwd, dbl_TimeToMat, dblArr_SmilePillars(int_ctr))
                End If
            End If
            Call dic_Cache_DeltaPillars.Add(str_ActiveCacheKey, dblArr_DeltaPillars(int_ctr))
        End If
    Next int_ctr

    DeriveDeltaAxis = dblArr_DeltaPillars
End Function



Private Function CleanLookupDate(lng_LookupDateRaw As Long) As Long
    ' ## Prevent lookup date from being before the first maturity pillar
    Dim lng_Output As Long
    Dim lng_FirstPillarDate As Long: lng_FirstPillarDate = rng_FinalDates(1, 1).Value
    If lng_LookupDateRaw < lng_FirstPillarDate Then lng_Output = lng_FirstPillarDate Else lng_Output = lng_LookupDateRaw
    CleanLookupDate = lng_Output
End Function

Private Function BuildCacheKey_DeltaScale(int_Index_Time As Integer, int_Index_Delta As Integer, dbl_VolSpread As Double, _
    bln_GetOrig As Boolean) As String
    ' ## Build key to look up stored delta scale values
    BuildCacheKey_DeltaScale = int_Index_Time & "|" & int_Index_Delta & "|" & dbl_VolSpread & "|" & bln_GetOrig
End Function

Public Sub GeneratePillarDates(var_Final As Variant)
    ' Generate option maturity dates from terms
    Dim bln_Final As Boolean: bln_Final = CBool(var_Final)
    Dim str_Code As String: str_Code = str_CCY_Fgn & str_CCY_Dom
    Dim cal_DTV As Calendar: cal_DTV = cas_Calendars.Lookup_Calendar(str_Calendar_DTV)
    Dim cal_STD As Calendar: cal_STD = cas_Calendars.Lookup_Calendar(str_Calendar_STD)

    ' If generating shifted pillar set under theta, set the build date to the shifted valuation date
    Dim lng_BuildDate As Long
    If bln_Final = True Then
        lng_BuildDate = cfg_Settings.CurrentValDate
        rng_FinalBuildDate.Value = lng_BuildDate
    Else
        lng_BuildDate = lng_BuildDate_Orig
    End If

    Dim str_ActiveTerm As String
    Dim int_NumRows As Integer: int_NumRows = NumQueryRows()
    Dim lngArr_Dates() As Long: ReDim lngArr_Dates(1 To int_NumRows, 1 To 1) As Long
    Dim dic_Cache_Dates As New Dictionary: dic_Cache_Dates.CompareMode = CompareMethod.TextCompare

    ' Obtain shifters
    Dim dss_Shifters As DateShifterSet: Set dss_Shifters = dic_GlobalStaticInfo(StaticInfoType.DateShifterSet)
    Dim shi_Forward As DateShifter: Set shi_Forward = dss_Shifters.Lookup_Shifter(str_FwdShifter)
    Dim shi_Backward As DateShifter: Set shi_Backward = dss_Shifters.Lookup_Shifter(str_BackShifter)

    ' Set up external calendars in shifters
    Call shi_Forward.IncludeExternalCalendar(str_Calendar_DTV)
    If shi_Forward.IsRelShifter Then Call shi_Forward.BaseShifter.IncludeExternalCalendar(str_Calendar_DTV)
    Call shi_Backward.IncludeExternalCalendar(str_Calendar_DTV)
    If shi_Backward.IsRelShifter Then Call shi_Forward.BaseShifter.IncludeExternalCalendar(str_Calendar_DTV)

    ' Determine dates
    Dim lng_SpotDate As Long: lng_SpotDate = shi_Forward.Lookup_ShiftedDate(lng_BuildDate)
    lng_SpotDate = Date_ApplyBDC(lng_SpotDate, "FOLL", cal_STD.HolDates, cal_STD.Weekends)  ' In case the spot date is a holiday in the other calendar

    Dim int_ctr As Integer, lng_ActiveMat As Long, lng_ActiveTermDate As Long, lng_ActiveDelivDate As Long
    For int_ctr = 1 To int_NumRows
        str_ActiveTerm = UCase(rng_MXTerms(int_ctr, 1).Value)

        If dic_Cache_Dates.Exists(str_ActiveTerm) Then
            lng_ActiveMat = dic_Cache_Dates(str_ActiveTerm)
        Else
            Select Case str_ActiveTerm
                Case "1D", "1W"
                    ' Murex appears to use a common set of dates across all pairs
                    lng_ActiveTermDate = date_addterm(lng_BuildDate, str_ActiveTerm, 1)
                    lng_ActiveMat = Date_ApplyBDC(lng_ActiveTermDate, "FOLL")
                Case Else
                    ' Add term using modified following business day convention and apply shifters
                    lng_ActiveTermDate = date_addterm(lng_SpotDate, str_ActiveTerm, 1)
                    lng_ActiveDelivDate = Date_ApplyBDC(lng_ActiveTermDate, "FOLL", cal_STD.HolDates, cal_STD.Weekends)
                    lng_ActiveMat = shi_Backward.Lookup_ShiftedDate(lng_ActiveDelivDate)
            End Select

            Call dic_Cache_Dates.Add(str_ActiveTerm, lng_ActiveMat)
        End If

        lngArr_Dates(int_ctr, 1) = lng_ActiveMat
    Next int_ctr

    ' Remove external calendars from shifters
    shi_Forward.RemoveExternalCalendar
    If shi_Forward.IsRelShifter Then shi_Forward.BaseShifter.RemoveExternalCalendar
    shi_Backward.RemoveExternalCalendar
    If shi_Backward.IsRelShifter Then shi_Backward.BaseShifter.RemoveExternalCalendar

    ' Output to sheet
    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    If bln_Final = True Then rng_FinalDates.Value = lngArr_Dates Else rng_OrigDates.Value = lngArr_Dates
    rng_OrigDates.NumberFormat = str_DateFormat
    rng_FinalDates.NumberFormat = str_DateFormat

    ' Update interpolation scales
    Call FillFinalVolsDict
End Sub

Public Sub ResetCache_Lookups()
    ' ## Reset cached lookup values
    Set dic_Cache_SmileVols = New Dictionary
    Set dic_Cache_DeltaPillars = New Dictionary
    Set dic_Cache_PolyCoefs = New Dictionary
    Set dic_Cache_SmileSlopes = New Dictionary
End Sub

Private Function Gather_ShiftObj(var_Delta As Variant, enu_ShockType As ShockType) As CurveDaysShift
    ' ## Return object containing the shifts for the specified delta.  For a uniform shift across deltas, specify 0 as the delta input
    ' Determine shift type
    Dim dic_ToUse As Dictionary
    Select Case enu_ShockType
        Case ShockType.Absolute: Set dic_ToUse = dic_AbsShifts
        Case ShockType.Relative: Set dic_ToUse = dic_RelShifts
    End Select

    ' Gather shift object with the specified delta, otherwise create it
    Dim csh_Found As CurveDaysShift
    If dic_ToUse.Exists(var_Delta) Then
        Set csh_Found = dic_ToUse(var_Delta)
    Else
        Set csh_Found = New CurveDaysShift
        Call csh_Found.Initialize(enu_ShockType)
        Call dic_ToUse.Add(var_Delta, csh_Found)
    End If

    Set Gather_ShiftObj = csh_Found
End Function

Private Sub FillFinalVolsDict()
    Set dic_Vols_Final = Gather_InterpDict(rng_FinalDeltas, rng_FinalDates, rng_FinalVols)
End Sub

Private Sub OutputShifts(ByRef rng_ActiveOutput_TopLeft As Range, dic_Outer As Dictionary, int_ShiftsOffset As Integer)
    Dim var_ActiveDelta As Variant, csh_ActiveOutput As CurveDaysShift, int_ActiveNumRows As Integer
    For Each var_ActiveDelta In dic_Outer.Keys
        Set csh_ActiveOutput = dic_Outer(var_ActiveDelta)
        int_ActiveNumRows = csh_ActiveOutput.NumShifts
        rng_ActiveOutput_TopLeft.Resize(int_ActiveNumRows, 1).Value = csh_ActiveOutput.Days_Arr
        rng_ActiveOutput_TopLeft.Offset(0, 1).Resize(int_ActiveNumRows, 1).Value = var_ActiveDelta
        rng_ActiveOutput_TopLeft.Offset(0, int_ShiftsOffset).Resize(int_ActiveNumRows, 1).Value = csh_ActiveOutput.Shifts_Arr

        Set rng_ActiveOutput_TopLeft = rng_ActiveOutput_TopLeft.Offset(int_ActiveNumRows, 0)
    Next var_ActiveDelta
End Sub


' ## METHODS - SECNARIOS
Public Sub Scen_ApplyBase()
    ' Clear out shifts
    Call dic_RelShifts.RemoveAll
    Call dic_AbsShifts.RemoveAll
    Call Action_ClearBelow(rng_Days_TopLeft, 4)

    ' Reset data
    rng_FinalDates.Value = rng_OrigDates.Value
    rng_FinalVols.Value = rng_OrigVols.Value
    rng_FinalBuildDate.Value = lng_BuildDate_Orig

    ' Sync memory
    Call FillFinalVolsDict
    Call ResetCache_Lookups
End Sub

Public Sub Scen_AddShock_DaysDelta(int_numdays As Integer, var_Delta As Variant, enu_ShockType As ShockType, dbl_Amount As Double)
    Dim csh_Found As CurveDaysShift: Set csh_Found = Gather_ShiftObj(var_Delta, enu_ShockType)
    Call csh_Found.AddShift(int_numdays, dbl_Amount)
End Sub

Public Sub Scen_AddShock_Days(int_numdays As Integer, enu_ShockType As ShockType, dbl_Amount As Double)
    Call Me.Scen_AddShock_DaysDelta(int_numdays, "-", enu_ShockType, dbl_Amount)
End Sub

Public Sub Scen_AddShock_TermDelta(str_Term As String, var_Delta As Variant, enu_ShockType As ShockType, dbl_Amount As Double)
    Dim csh_Found As CurveDaysShift: Set csh_Found = Gather_ShiftObj(var_Delta, enu_ShockType)
    Dim lng_MatDate As Long: lng_MatDate = Lookup_MXTermDate(str_Term)
    Call csh_Found.AddIsolatedShift(lng_MatDate - rng_FinalBuildDate.Value, dbl_Amount)
End Sub

Public Sub Scen_AddShock_Term(str_Term As String, enu_ShockType As ShockType, dbl_Amount As Double)
    Call Me.Scen_AddShock_TermDelta(str_Term, "-", enu_ShockType, dbl_Amount)
End Sub

Public Sub Scen_AddShock_Delta(var_Delta As Variant, enu_ShockType As ShockType, dbl_Amount As Double)
    Dim csh_Found As CurveDaysShift: Set csh_Found = Gather_ShiftObj(var_Delta, enu_ShockType)
    Call csh_Found.AddUniformShift(dbl_Amount)
End Sub

Public Sub Scen_AddShock_Uniform(enu_ShockType As ShockType, dbl_Amount As Double)
    Call Me.Scen_AddShock_Delta("-", enu_ShockType, dbl_Amount)
End Sub

Public Sub Scen_ApplyCurrent()
    Dim bln_RelShifts As Boolean: bln_RelShifts = (dic_RelShifts.count > 0)
    Dim bln_AbsShifts As Boolean: bln_AbsShifts = (dic_AbsShifts.count > 0)

    If bln_RelShifts = True Or bln_AbsShifts = True Then
        Dim int_NumRows As Integer: int_NumRows = NumQueryRows()
        Dim lngArr_MatDates() As Variant: lngArr_MatDates = rng_FinalDates.Value
        Dim lng_BuildDate As Long: lng_BuildDate = rng_FinalBuildDate.Value
        Dim intArr_Deltas() As Variant: intArr_Deltas = rng_FinalDeltas.Value
        Dim dblArr_OrigVols() As Variant: dblArr_OrigVols = rng_OrigVols.Value
        Dim dblArr_FinalVols() As Double: ReDim dblArr_FinalVols(1 To int_NumRows, 1 To 1) As Double
        Dim csh_ActiveDelta_Rel As CurveDaysShift, csh_ActiveDelta_Abs As CurveDaysShift

        ' Gather uniform shift objects if these exist
        Dim csh_Uniform_Rel As CurveDaysShift, csh_Uniform_Abs As CurveDaysShift
        If dic_RelShifts.Exists("-") Then Set csh_Uniform_Rel = dic_RelShifts("-")
        If dic_AbsShifts.Exists("-") Then Set csh_Uniform_Abs = dic_AbsShifts("-")

        Dim int_ctr As Integer, int_ActiveDelta As Integer, int_ActiveDays As Integer
        Dim dbl_ActiveShift_Rel As Double, dbl_ActiveShift_Abs As Double
        For int_ctr = 1 To int_NumRows
            int_ActiveDelta = intArr_Deltas(int_ctr, 1)
            int_ActiveDays = lngArr_MatDates(int_ctr, 1) - lng_BuildDate

            ' Gather the shift objects for the specified delta
            If dic_RelShifts.Exists(int_ActiveDelta) Then
                Set csh_ActiveDelta_Rel = dic_RelShifts(int_ActiveDelta)
            Else
                Set csh_ActiveDelta_Rel = Nothing
            End If

            If dic_AbsShifts.Exists(int_ActiveDelta) Then
                Set csh_ActiveDelta_Abs = dic_AbsShifts(int_ActiveDelta)
            Else
                Set csh_ActiveDelta_Abs = Nothing
            End If

            ' Determine the shift to apply to the current pillar
            dbl_ActiveShift_Rel = 0
            dbl_ActiveShift_Abs = 0
            If Not csh_ActiveDelta_Rel Is Nothing Then dbl_ActiveShift_Rel = csh_ActiveDelta_Rel.ReadShift(int_ActiveDays)
            If Not csh_ActiveDelta_Abs Is Nothing Then dbl_ActiveShift_Abs = csh_ActiveDelta_Abs.ReadShift(int_ActiveDays)
            If Not csh_Uniform_Rel Is Nothing Then dbl_ActiveShift_Rel = dbl_ActiveShift_Rel + csh_Uniform_Rel.ReadShift(int_ActiveDays)
            If Not csh_Uniform_Abs Is Nothing Then dbl_ActiveShift_Abs = dbl_ActiveShift_Abs + csh_Uniform_Abs.ReadShift(int_ActiveDays)

            dblArr_FinalVols(int_ctr, 1) = dblArr_OrigVols(int_ctr, 1) * (1 + dbl_ActiveShift_Rel / 100) + dbl_ActiveShift_Abs
            If dblArr_FinalVols(int_ctr, 1) <= 0 Then dblArr_FinalVols(int_ctr, 1) = 0.000001
        Next int_ctr

        ' Output shifts to sheet
        Dim rng_ActiveOutput_TopLeft As Range: Set rng_ActiveOutput_TopLeft = rng_Days_TopLeft
        Dim var_ActiveDelta As Variant, csh_ActiveOutput As CurveDaysShift, int_ActiveNumRows As Integer
        If bln_RelShifts = True Then Call OutputShifts(rng_ActiveOutput_TopLeft, dic_RelShifts, 2)
        If bln_AbsShifts = True Then Call OutputShifts(rng_ActiveOutput_TopLeft, dic_AbsShifts, 3)

        ' Write shifted values back to sheet
        rng_FinalVols.Value = dblArr_FinalVols
    Else
        rng_FinalVols.Value = rng_OrigVols.Value
    End If

    Call FillFinalVolsDict
    Call ResetCache_Lookups
End Sub


' ## METHODS - OUTPUT
Public Sub OutputFinalVols(rng_OutputStart As Range)
    Dim int_NumRows As Integer: int_NumRows = NumQueryRows()
    If int_NumRows > 0 Then
        Dim str_Fgn As String: str_Fgn = Left(Me.CurveName, 3)
        Dim str_Dom As String: str_Dom = Right(Me.CurveName, 3)
        Dim varArr_Output() As Variant: ReDim varArr_Output(1 To int_NumRows, 1 To 6) As Variant
        Dim intArr_Deltas() As Variant: intArr_Deltas = rng_FinalDeltas.Value
        Dim lngArr_FinalDates() As Variant: lngArr_FinalDates = rng_FinalDates.Value
        Dim dblArr_FinalVols() As Variant: dblArr_FinalVols = rng_FinalVols.Value

        Dim int_RowCtr As Integer
        For int_RowCtr = 1 To int_NumRows
            varArr_Output(int_RowCtr, 1) = str_Fgn
            varArr_Output(int_RowCtr, 2) = str_Dom
            varArr_Output(int_RowCtr, 3) = CInt(lngArr_FinalDates(int_RowCtr, 1) - lng_BuildDate_Orig)
            varArr_Output(int_RowCtr, 4) = intArr_Deltas(int_RowCtr, 1)
            varArr_Output(int_RowCtr, 5) = lng_BuildDate_Orig
            varArr_Output(int_RowCtr, 6) = dblArr_FinalVols(int_RowCtr, 1)
        Next int_RowCtr

        rng_OutputStart.Resize(int_NumRows, 6).Value = varArr_Output
    End If
End Sub