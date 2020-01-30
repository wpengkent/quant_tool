Option Explicit

' ## READ INFO
Public Function cyGetEqSpot(str_Code As String) As Double
    Application.Volatile
    Dim eqs_All As Data_EQSpots: Set eqs_All = GetObject_EQSpots(True)
    cyGetEqSpot = eqs_All.Lookup_Spot(str_Code)
End Function

Public Function cyGetEqVol(str_Code As String) As Double
    Application.Volatile
    Dim eqv_All As Data_EQVols: Set eqv_All = GetObject_EQVols(True)
    cyGetEqVol = eqv_All.Lookup_Vol(str_Code)
End Function


Public Function cyGetFXFwd(str_Fgn As String, str_Dom As String, Optional lng_Maturity As Long = 0, _
    Optional bln_UseSpotDelay As Boolean = True) As Double

    Application.Volatile
    Dim fxs_All As Data_FXSpots: Set fxs_All = GetObject_FXSpots(True)
    cyGetFXFwd = fxs_All.Lookup_Fwd(str_Fgn, str_Dom, lng_Maturity, bln_UseSpotDelay)
End Function

Public Function cyGetFXSpot(str_Fgn As String, str_Dom As String) As Double
    cyGetFXSpot = cyGetFXFwd(str_Fgn, str_Dom)
End Function

Public Function cyGetFXDiscSpot(str_Fgn As String, str_Dom As String, _
    Optional str_SpecialScen As String = "<NONE>") As Double

    ' ## Returns spot used for immediate translation between currencies
    Application.Volatile

    If str_Fgn = str_Dom Then
        cyGetFXDiscSpot = 1
    Else
        Dim fxs_All As Data_FXSpots: Set fxs_All = GetObject_FXSpots(True)
        cyGetFXDiscSpot = fxs_All.Lookup_DiscSpot(str_Fgn, str_Dom)
    End If
End Function

Public Function cyGetFXVol(str_Fgn As String, str_Dom As String, lng_date As Long, Optional dbl_Strike As Double = -1, _
    Optional bln_GetOrig As Boolean = False, Optional bln_AxisRescale As Boolean = True) As Double
    ' Translate inputs to a code to look up sheet
    Application.Volatile

    ' Allow both quotations of a pair to be recognized, then gather the curve
    Dim map_Rules As MappingRules: Set map_Rules = GetObject_MappingRules()
    Dim str_MappedCode As String: str_MappedCode = map_Rules.Lookup_MappedFXVolPair(str_Fgn, str_Dom)
    Dim fxv_Found As Data_FXVols: Set fxv_Found = GetObject_FXVols(str_MappedCode, True, False)

    ' Look up smile if strike is given, otherwise look up ATM volatility
    If dbl_Strike = -1 Then
        cyGetFXVol = fxv_Found.Lookup_ATMVol(lng_date)
    Else
        If fxv_Found.SmileCCY = str_Fgn Then
            cyGetFXVol = fxv_Found.Lookup_SmileVol(lng_date, dbl_Strike, bln_GetOrig, bln_AxisRescale)
        Else
            ' Invert quotation
            cyGetFXVol = fxv_Found.Lookup_SmileVol(lng_date, 1 / dbl_Strike, bln_GetOrig, bln_AxisRescale)
        End If
    End If
End Function

Public Function cyGetFXVegaBucketWeight(str_Fgn As String, str_Dom As String, lng_date As Long, _
    Optional bln_GetOrig As Boolean = False) As Variant()
    ' ## Obtain the weight applied to the total vega to obtain the bucketed vega for an option with the specified maturity
    Application.Volatile

    ' Allow both quotations of a pair to be recognized, then gather the curve
    Dim map_Rules As MappingRules: Set map_Rules = GetObject_MappingRules()
    Dim str_MappedCode As String: str_MappedCode = map_Rules.Lookup_MappedFXVolPair(str_Fgn, str_Dom)
    Dim fxv_Found As Data_FXVols: Set fxv_Found = GetObject_FXVols(str_MappedCode, True, False)

    cyGetFXVegaBucketWeight = fxv_Found.Lookup_BucketWeights(lng_date, bln_GetOrig)
End Function

Public Function cyReadIRCurve(str_curve As String, lng_startdate As Long, lng_EndDate As Long, str_RateType As String, _
    Optional str_System As String = "", Optional bln_AllowBackwards = True, Optional dbl_ZSpread As Double = 0) As Double

    ' ## Returns either a zero rate or a discount factor
    Application.Volatile
    Dim str_InterpPillars As String
    If str_System <> "" Then str_InterpPillars = GetObject_MappingRules().Dict_PillarSets(str_System)
    Dim irc_Curve As Data_IRCurve: Set irc_Curve = GetObject_IRCurve(str_curve, True, False)
    cyReadIRCurve = irc_Curve.Lookup_Rate(lng_startdate, lng_EndDate, str_RateType, str_InterpPillars, , bln_AllowBackwards, , dbl_ZSpread)
End Function

Public Function cyGetCapVol(str_Code As String, lng_date As Long) As Double
    Application.Volatile
   ' Dim cvl_Found As Data_CapVols: Set cvl_Found = GetObject_CapVols(str_Code, True, False)
   Dim cvl_Found As Data_CapVolsQJK: Set cvl_Found = GetObject_CapVols(str_Code, True, False)  'QJK code 16/12/2014
    cyGetCapVol = cvl_Found.Lookup_Vol(lng_date)
End Function

Public Function cyGetSwptVol(str_Code As String, lng_OptionMat As Long, lng_SwapMat As Long, Optional dbl_Strike As Double = -1) As Double
    Application.Volatile
    Dim svc_Found As Data_SwptVols: Set svc_Found = GetObject_SwptVols(str_Code, True, False)
    cyGetSwptVol = svc_Found.Lookup_Vol(lng_OptionMat, lng_SwapMat, dbl_Strike)
End Function

Public Function cyGetCodeList(enu_Type As CurveType) As Variant
    Application.Volatile
    Dim wks_Location As Worksheet

    ' Find relevant sheet
    Set wks_Location = ThisWorkbook.Worksheets("Setup - " & GetCurveTypeName(enu_Type))

    ' Determine number of rows
    Dim rng_Active As Range: Set rng_Active = wks_Location.Range("A4")
    Dim int_NumRows As Integer
    If rng_Active.Offset(1, 0).Value = "" Then
        int_NumRows = 1
    Else
        int_NumRows = rng_Active.End(xlDown).Row - 3
    End If
    Dim strArr_Output() As String: ReDim strArr_Output(1 To int_NumRows)

    Dim int_ctr As Integer

    ' Copy data from column into array
    For int_ctr = 1 To int_NumRows
        strArr_Output(int_ctr) = rng_Active.Value
        Set rng_Active = rng_Active.Offset(1, 0)
    Next int_ctr

    cyGetCodeList = strArr_Output
End Function

Public Function cyGetValDate() As Long
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = GetObject_ConfigSheet()
    cyGetValDate = cfg_Settings.CurrentValDate
End Function

Public Function cyGetFXSpotDate(str_Currency As String, lng_ValDate As Long, dic_StaticInfo As Dictionary) As Long
    ' ## Currency is the non-USD currency of any non-cross pair
    Application.Volatile

    Dim lng_Output As Long
    If dic_StaticInfo Is Nothing Then Set dic_StaticInfo = GetStaticInfo()
    Dim map_Rules As MappingRules: Set map_Rules = dic_StaticInfo(StaticInfoType.MappingRules)
    Dim cas_Calendars As CalendarSet: Set cas_Calendars = dic_StaticInfo(StaticInfoType.CalendarSet)
    Dim cal_Calendar As Calendar: cal_Calendar = cas_Calendars.Lookup_Calendar(map_Rules.Lookup_CCYCalendar(str_Currency))
    Dim cal_Calendar_USD As Calendar: cal_Calendar_USD = cas_Calendars.Lookup_Calendar(map_Rules.Lookup_CCYCalendar("USD"))
    Dim int_SpotDays As Integer: int_SpotDays = map_Rules.Lookup_CCYSpotDays(str_Currency)

    lng_Output = date_workday(lng_ValDate, int_SpotDays, cal_Calendar.HolDates, cal_Calendar.Weekends)

    ' Don't allow spot date to be a USD holiday or a foreign holiday
    lng_Output = Date_ApplyBDC(lng_Output, "FOLL", cal_Calendar_USD.HolDates, cal_Calendar_USD.Weekends)
    cyGetFXSpotDate = Date_ApplyBDC(lng_Output, "FOLL", cal_Calendar.HolDates, cal_Calendar.Weekends)
End Function

Public Function cyGetFXTomDate(str_Currency As String, lng_ValDate As Long, dic_StaticInfo As Dictionary) As Long
    ' ## Currency is the non-USD currency of any non-cross pair
    ' ## Used for 1D from today for 2D spot currencies
    Application.Volatile

    Dim lng_Output As Long
    If dic_StaticInfo Is Nothing Then Set dic_StaticInfo = GetStaticInfo()
    Dim map_Rules As MappingRules: Set map_Rules = dic_StaticInfo(StaticInfoType.MappingRules)
    Dim cas_Calendars As CalendarSet: Set cas_Calendars = dic_StaticInfo(StaticInfoType.CalendarSet)
    Dim cal_Calendar As Calendar: cal_Calendar = cas_Calendars.Lookup_Calendar(map_Rules.Lookup_CCYCalendar(str_Currency))
    Dim cal_Calendar_USD As Calendar: cal_Calendar_USD = cas_Calendars.Lookup_Calendar(map_Rules.Lookup_CCYCalendar("USD"))

    lng_Output = date_workday(lng_ValDate, 1, cal_Calendar.HolDates, cal_Calendar.Weekends)

    ' Don't allow spot date to be a USD holiday or a foreign holiday
    lng_Output = Date_ApplyBDC(lng_Output, "FOLL", cal_Calendar_USD.HolDates, cal_Calendar_USD.Weekends)
    cyGetFXTomDate = Date_ApplyBDC(lng_Output, "FOLL", cal_Calendar.HolDates, cal_Calendar.Weekends)
End Function


Public Function cyGetFXCrossSpotDate(str_Fgn As String, str_Dom As String, lng_ValDate As Long, Optional dic_StaticInfo As Dictionary) As Long
   Dim lng_FgnSpotDate As Long: lng_FgnSpotDate = cyGetFXSpotDate(str_Fgn, lng_ValDate, dic_StaticInfo)
    Dim lng_DomSpotDate As Long: lng_DomSpotDate = cyGetFXSpotDate(str_Dom, lng_ValDate, dic_StaticInfo)

    cyGetFXCrossSpotDate = Examine_MaxOfPair(lng_FgnSpotDate, lng_DomSpotDate)
End Function


' ## OPERATIONS
Public Sub cyRefreshAll()
    ' ## Retrieve all market data from the rates DB, irrespective of whether the curve is selected in the setup sheet
    Dim fld_AppState_Orig As ApplicationState: fld_AppState_Orig = Gather_ApplicationState(ApplicationStateType.Current)
    Dim fld_AppState_Opt As ApplicationState: fld_AppState_Opt = Gather_ApplicationState(ApplicationStateType.Optimized)
    Call Action_SetAppState(fld_AppState_Opt)

    Dim dic_CurveSet As Dictionary: Set dic_CurveSet = GetAllCurves(False, True)

    Dim eqs_Spots As Data_EQSpots: Set eqs_Spots = dic_CurveSet(CurveType.EQSPT)
    Call eqs_Spots.LoadRates

    Dim eqs_Vols As Data_EQVols: Set eqs_Vols = dic_CurveSet(CurveType.EQVOL)
    Call eqs_Vols.LoadRates

    Dim fxs_Spots As Data_FXSpots: Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
    Call fxs_Spots.LoadRates

    Call GenerateAllFXVolCurves(dic_CurveSet(CurveType.FXV))
    Call GenerateAllIRCurves(dic_CurveSet(CurveType.IRC))
    Call GenerateAllCapVolCurves(dic_CurveSet(CurveType.cvl))
    Call GenerateAllSwptVolCurves(dic_CurveSet(CurveType.SVL))

    Call Action_SetAppState(fld_AppState_Orig)
End Sub

Public Sub cyRefreshAllSelected()
    ' ## Retrieve all market data from the rates DB for the curves selected in the setup sheet
    Dim fld_AppState_Orig As ApplicationState: fld_AppState_Orig = Gather_ApplicationState(ApplicationStateType.Current)
    Dim fld_AppState_Opt As ApplicationState: fld_AppState_Opt = Gather_ApplicationState(ApplicationStateType.Optimized)
    Call Action_SetAppState(fld_AppState_Opt)

    Dim wks_Current As Worksheet: Set wks_Current = ThisWorkbook.ActiveSheet
    Dim dic_CurveSet As Dictionary: Set dic_CurveSet = GetAllCurves(False, True)

    Dim eqs_Spots As Data_EQSpots: Set eqs_Spots = dic_CurveSet(CurveType.EQSPT)
    Call eqs_Spots.LoadRates

    Dim eqs_Vols As Data_EQVols: Set eqs_Vols = dic_CurveSet(CurveType.EQVOL)
    Call eqs_Vols.LoadRates

    Dim fxs_Spots As Data_FXSpots: Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
    Call fxs_Spots.LoadRates

    Call GenerateSelectedFXVolCurves(dic_CurveSet(CurveType.FXV))
    Call GenerateSelectedIRCurves(dic_CurveSet(CurveType.IRC))
    Call GenerateSelectedCapVolCurves(dic_CurveSet(CurveType.cvl))
    Call GenerateSelectedSwptVolCurves(dic_CurveSet(CurveType.SVL))

    Call GotoSheet(wks_Current.Name)
    Call Action_SetAppState(fld_AppState_Orig)
End Sub

Public Sub cyRefreshFXVolCurve(str_Code As String, dic_CurveSet_FXV As Dictionary)
    Dim fxv_Found As Data_FXVols: Set fxv_Found = dic_CurveSet_FXV(str_Code)
    Application.StatusBar = "Data date: " & Format(fxv_Found.ConfigSheet.CurrentDataDate, "dd/mm/yyyy") & "     FXV: " & str_Code
    Call fxv_Found.SetParams(GetRange_CurveParams(CurveType.FXV, str_Code))
    fxv_Found.LoadRates
End Sub

Public Sub cyRebootstrapIRCurve(str_curve As String, dic_CurveSet_IRC As Dictionary)
    Dim irc_Found As Data_IRCurve: Set irc_Found = dic_CurveSet_IRC(str_curve)
    Application.StatusBar = "Data date: " & Format(irc_Found.ConfigSheet.CurrentDataDate, "dd/mm/yyyy") & "     IRC: " & str_curve
    Call irc_Found.Action_Bootstrap(True)
End Sub

Public Sub cyRefreshIRCurve(str_curve As String, dic_CurveSet_IRC As Dictionary)
    Dim irc_Found As Data_IRCurve: Set irc_Found = dic_CurveSet_IRC(str_curve)
    Application.StatusBar = "Data date: " & Format(irc_Found.ConfigSheet.CurrentDataDate, "dd/mm/yyyy") & "     IRC: " & str_curve
    Call irc_Found.SetParams(GetRange_CurveParams(CurveType.IRC, str_curve), str_curve)
    irc_Found.LoadRates
End Sub

Public Sub cyRefreshCapVolCurve(str_Code As String, dic_CurveSet_CVL As Dictionary)
    'Dim cvl_Found As Data_CapVols: Set cvl_Found = dic_CurveSet_CVL(str_Code)
    Dim cvl_Found As Data_CapVolsQJK: Set cvl_Found = dic_CurveSet_CVL(str_Code)   'QJK code 16/12/2014
    Application.StatusBar = "Data date: " & Format(cvl_Found.ConfigSheet.CurrentDataDate, "dd/mm/yyyy") & "     CVL: " & str_Code
    Call cvl_Found.SetParams(GetRange_CurveParams(CurveType.cvl, str_Code))
    cvl_Found.LoadRates
End Sub
Public Function cyGetCapVolSurf(str_Code As String, lng_date As Long, strike As Double) As Double
    Application.Volatile
   'replace cyGetCapVol because cap vol surface is being introduced
   Dim cvl_Found As Data_CapVolsQJK: Set cvl_Found = GetObject_CapVolSurf(str_Code, strike, True, False)
    cyGetCapVolSurf = cvl_Found.Lookup_Vol(lng_date)
End Function

Public Sub cyRebootstrapCapVolCurve(str_Code As String, dic_CurveSet_CVL As Dictionary)
    Dim cvl_Found As Data_CapVols: Set cvl_Found = dic_CurveSet_CVL(str_Code)
    Application.StatusBar = "Data date: " & Format(cvl_Found.ConfigSheet.CurrentDataDate, "dd/mm/yyyy") & "     CVL: " & str_Code
    Call cvl_Found.Bootstrap(False)
End Sub

Public Sub cyRefreshSwptVolCurve(str_Code As String, dic_CurveSet_SVL As Dictionary)
    Dim svl_Found As Data_SwptVols: Set svl_Found = dic_CurveSet_SVL(str_Code)
    Application.StatusBar = "Data date: " & Format(svl_Found.ConfigSheet.CurrentDataDate, "dd/mm/yyyy") & "     SVL: " & str_Code
    Call svl_Found.SetParams(GetRange_CurveParams(CurveType.SVL, str_Code))
    svl_Found.LoadRates
End Sub

'##Equity smile - matt edit
Public Sub cyRefreshEQSmileCurve(str_Code As String, dic_CurveSet_EVL As Dictionary)
    Dim evl_Found As Data_EQSmile: Set evl_Found = dic_CurveSet_EVL(str_Code)
    Application.StatusBar = "Data date: " & Format(evl_Found.ConfigSheet.CurrentDataDate, "dd/mm/yyyy") & "     EVL: " & str_Code
    Call evl_Found.SetParams(GetRange_CurveParams(CurveType.EVL, str_Code))
    evl_Found.LoadRates
End Sub


' ## SCENARIOS & SHOCKS
Public Sub ApplyListedIRCurveShocks(dic_Curves As Dictionary, bln_Propagation As Boolean)
    Dim str_ActiveCode As Variant, irc_Active As Data_IRCurve

    For Each str_ActiveCode In dic_Curves.Keys
        Set irc_Active = dic_Curves(str_ActiveCode)
        Call irc_Active.Scen_ApplyCurrent(bln_Propagation)
    Next str_ActiveCode
End Sub

Public Sub cyApplyIRCurvePropagation(str_Code As String)
    If str_Code <> "" Then
        Dim irc_Found As Data_IRCurve: Set irc_Found = GetObject_IRCurve(str_Code, True, False)  ' ## DEV - Remove GetObject_IRCurve
        irc_Found.Scen_ReceivePropagation
    End If
End Sub


' ## CALCULATION TOOLS
Public Function cyApplyBucketing(rng_SourceMatDates As Range, rng_BucketedMatDates As Range, rng_SourceValues As Range, _
    lng_TargetBucketedMatDate As Long) As Double
    ' ## Dates should be pre-sorted in ascending order

    Dim int_TargetBucketPos As Integer: int_TargetBucketPos = WorksheetFunction.Match(lng_TargetBucketedMatDate, rng_BucketedMatDates, 0)
    Dim int_NumBuckets As Integer: int_NumBuckets = rng_BucketedMatDates.Rows.count
    Dim int_NumSourcePillars As Integer: int_NumSourcePillars = rng_SourceMatDates.count
    Dim int_LowerBucketPos As Integer, int_UpperBucketPos As Integer
    Dim lng_LowerBucketMatDate As Long, lng_UpperBucketMatDate As Long

    ' Set bounds for relevant pillar dates
    If int_TargetBucketPos = 1 Then
        int_LowerBucketPos = 1
        lng_LowerBucketMatDate = rng_SourceMatDates(1, 1).Value
    Else
        int_LowerBucketPos = int_TargetBucketPos - 1
        lng_LowerBucketMatDate = rng_BucketedMatDates(int_LowerBucketPos, 1).Value
    End If

    If int_TargetBucketPos = int_NumBuckets Then
        int_UpperBucketPos = int_NumBuckets
        lng_UpperBucketMatDate = rng_SourceMatDates(int_NumSourcePillars).Value
    Else
        int_UpperBucketPos = int_TargetBucketPos + 1
        lng_UpperBucketMatDate = rng_BucketedMatDates(int_UpperBucketPos, 1).Value
    End If

    ' Determine index of relevant source pillars
    Dim int_StartRelevant As Integer
    Dim int_EndRelevant As Integer: int_EndRelevant = int_NumSourcePillars
    Dim lng_ActiveDate As Long
    Dim int_RowCtr As Integer

    For int_RowCtr = 1 To rng_SourceMatDates.Rows.count
        lng_ActiveDate = rng_SourceMatDates(int_RowCtr, 1).Value

        ' Conditions are guaranteed to hit at some stage
        If int_StartRelevant = 0 And lng_ActiveDate >= lng_LowerBucketMatDate And lng_ActiveDate <= lng_UpperBucketMatDate Then
            int_StartRelevant = int_RowCtr
        End If

        If int_EndRelevant = int_NumSourcePillars And lng_ActiveDate > lng_UpperBucketMatDate Then
            int_EndRelevant = int_RowCtr
        End If
    Next int_RowCtr

    ' Set dates and weights for interpolation.  Weights are zero at adjacent buckets, and one at the target
    Dim lngArr_RelevantBucketMatDates() As Long: ReDim lngArr_RelevantBucketMatDates(int_LowerBucketPos To int_UpperBucketPos) As Long
    Dim intArr_RelevantBucketWeights() As Integer: ReDim intArr_RelevantBucketWeights(int_LowerBucketPos To int_UpperBucketPos) As Integer
    lngArr_RelevantBucketMatDates(int_LowerBucketPos) = lng_LowerBucketMatDate
    lngArr_RelevantBucketMatDates(int_UpperBucketPos) = lng_UpperBucketMatDate
    lngArr_RelevantBucketMatDates(int_TargetBucketPos) = lng_TargetBucketedMatDate
    intArr_RelevantBucketWeights(int_TargetBucketPos) = 1

    ' Find weight by interpolating the bucket maturity dates and weights.  Apply this weight to the source value and sum across all relevant pillars
    Dim dbl_ActiveWeight As Double
    Dim dbl_Output As Double
    For int_RowCtr = int_StartRelevant To int_EndRelevant
        dbl_ActiveWeight = Interp_Lin(lngArr_RelevantBucketMatDates, intArr_RelevantBucketWeights, rng_SourceMatDates(int_RowCtr, 1).Value, True)
        dbl_Output = dbl_Output + rng_SourceValues(int_RowCtr, 1).Value * dbl_ActiveWeight
    Next int_RowCtr

    ' Output result
    cyApplyBucketing = dbl_Output
End Function

Public Function cyGetCalendarAttribute(str_CalendarName As String, str_AttributeType As String) As Variant
    ' ## Returns code for determining weekends under the specified calendar
    Dim cal_Found As Calendar: cal_Found = GetObject_Calendar(str_CalendarName)
    Select Case UCase(str_AttributeType)
        Case "HOLS": Set cyGetCalendarAttribute = cal_Found.HolDates
        Case "WEEKENDS": cyGetCalendarAttribute = cal_Found.Weekends
    End Select
End Function

Public Function cyApplyDateShifter(str_ShifterName As String, lng_OrigDate As Long, Optional str_ExternalCal As String = "-", _
    Optional str_BaseExternalCal As String = "-") As Long
    Dim shi_Found As DateShifter: Set shi_Found = GetObject_DateShifter(str_ShifterName)
    Dim lng_Output As Long
    Dim bln_RemoveExt As Boolean: bln_RemoveExt = False
    Dim bln_RemoveBaseExt As Boolean: bln_RemoveBaseExt = False

    ' Include external calendars
    If str_ExternalCal <> "-" Then
        Call shi_Found.IncludeExternalCalendar(str_ExternalCal)
        bln_RemoveExt = True
    End If

    If str_BaseExternalCal <> "-" And Not shi_Found.BaseShifter Is Nothing Then
        Call shi_Found.BaseShifter.IncludeExternalCalendar(str_BaseExternalCal)
        bln_RemoveBaseExt = True
    End If

    lng_Output = shi_Found.Lookup_ShiftedDate(lng_OrigDate)

    ' Remove external calendars
    If bln_RemoveExt = True Then shi_Found.RemoveExternalCalendar
    If bln_RemoveBaseExt = True Then shi_Found.BaseShifter.RemoveExternalCalendar

    cyApplyDateShifter = lng_Output
End Function

Public Function cyGetIRCurveQuery(str_CurveName As String, lng_DataDate As Long) As String
    ' ## Return SQL used to read curve from rates database
    Dim iqs_Output As IRQuerySet: Set iqs_Output = GetObject_IRQuerySet()
    cyGetIRCurveQuery = iqs_Output.Lookup_SQL(str_CurveName, lng_DataDate)
End Function