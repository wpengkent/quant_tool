Option Explicit

' ## MEMBER DATA
Private wks_Location As Worksheet
Private rng_DataTopLeft As Range
Private rng_ScenNums As Range, rng_IsSmile As Range, rng_IsPropagation As Range
Private rng_DataTypes As Range, rng_Codes As Range, rng_Values As Range, rng_ShiftTypes As Range
Private rng_PillarTypes As Range, rng_Pillars As Range, rng_Deltas As Range, rng_IsActive As Range
Private dic_CurveSet As Dictionary, dic_GlobalStaticInfo As Dictionary
Private map_Rules As MappingRules, cas_Calendars As CalendarSet, cfg_Settings As ConfigSheet
Private lng_ValDate_Orig As Long, lng_ValDate_Current As Long
Private str_DBPath As String
Private Const int_NumCols As Integer = 11


' ## INITIALIZATION
Public Sub Initialize(wks_LocationInput As Worksheet, Optional dic_CurveSetInput As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing)

    Application.StatusBar = "Gathering curves..."

    Set wks_Location = wks_LocationInput
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    If dic_CurveSetInput Is Nothing Then Set dic_CurveSet = GetAllCurves(True, False, dic_GlobalStaticInfo) Else Set dic_CurveSet = dic_CurveSetInput
    Set map_Rules = dic_GlobalStaticInfo(StaticInfoType.MappingRules)
    Set cfg_Settings = dic_GlobalStaticInfo(StaticInfoType.ConfigSheet)
    Set cas_Calendars = dic_GlobalStaticInfo(StaticInfoType.CalendarSet)
    str_DBPath = cfg_Settings.ScenDBPath

    lng_ValDate_Orig = cfg_Settings.OrigValDate
    Call AssignRanges

    Application.StatusBar = False
End Sub


' ## PROPERTIES
Public Property Get NumShifts() As Long
    NumShifts = Examine_NumRows(rng_DataTopLeft)
End Property

Public Property Get ValDate_Current() As Long
    ValDate_Current = lng_ValDate_Current
End Property


' ## METHODS - PUBLIC
Public Sub LoadScenario(int_TargetScen As Integer)
    ' ## Load a scenario which has been already been entered into the container sheet
    ' ## Assumes scenarios are not overlapping

    Dim fld_AppState_Orig As ApplicationState: fld_AppState_Orig = Gather_ApplicationState(ApplicationStateType.Current)
    Dim fld_AppState_Opt As ApplicationState: fld_AppState_Opt = Gather_ApplicationState(ApplicationStateType.Optimized)
    Call Action_SetAppState(fld_AppState_Opt)

    ' Store name of container selected
    GetRange_ActiveContainer().Value = Right(wks_Location.Name, Len(wks_Location.Name) - Len("CONT_"))

    Dim str_ActiveDataType As String, str_ActiveCode As String, str_ActiveShockType As String, str_ActivePillarType As String
    Dim var_ActiveDelta As Variant, enu_ActiveShockType As ShockType
    Dim str_ActiveMappedType_Shift As String, str_ActiveMappedType_Storage As String, str_ActiveCurveName As String
    Dim var_ActivePillar As Variant, dbl_ActiveShockAmt As Double, str_UnderlyingSwap As String, str_ActiveMappedCode As String
    Dim cal_Active As Calendar
    Dim bln_Propagate As Boolean: bln_Propagate = False
    Dim bln_BuildDateShift As Boolean: bln_BuildDateShift = False
    Dim bln_DataDateShift As Boolean: bln_DataDateShift = False
    Dim int_ScenLoaded As Integer: int_ScenLoaded = 0
    Dim dic_DataTypeMapping As Dictionary

    ' Market data objects
    Dim eqs_AllSpots As Data_EQSpots
    Dim eqv_AllVols As Data_EQVols
    Dim fxs_AllSpots As Data_FXSpots
    Dim fxv_ActiveCurve As Data_FXVols
    Dim irc_ActiveCurve As Data_IRCurve
    Dim cvl_ActiveCurve As Data_CapVolsQJK
    Dim svc_ActiveCurve As Data_SwptVols
    Dim evl_ActiveCurve As Data_EQSmile 'Matt Edit evl
    Dim strArr_AllIRC() As String: strArr_AllIRC = cyGetCodeList(CurveType.IRC)
    Dim var_ActiveCurveName As Variant
    Dim var_ActiveCurve As Variant

    ' Clear out existing shocks
    Call ResetShiftedCurves

    ' May need to reset queried data back to the original data date
    Dim bln_NeedInitRefresh As Boolean, bln_ValDateReset As Boolean
    Dim lng_OrigBuildDate As Long: lng_OrigBuildDate = cfg_Settings.OrigBuildDate
    Dim lng_OrigDataDate As Long: lng_OrigDataDate = cfg_Settings.OrigDataDate
    bln_NeedInitRefresh = cfg_Settings.CurrentBuildDate <> lng_OrigBuildDate Or cfg_Settings.CurrentDataDate <> lng_OrigDataDate
    lng_ValDate_Current = lng_ValDate_Orig
    bln_ValDateReset = cfg_Settings.SetCurrentValDate(lng_ValDate_Current)
    cfg_Settings.CurrentBuildDate = lng_OrigBuildDate
    cfg_Settings.CurrentDataDate = lng_OrigDataDate
    If bln_NeedInitRefresh Then
        Call cyRefreshAll
    ElseIf bln_ValDateReset = True Then
        Call HandleValDateChange
    End If

    If int_TargetScen <> 0 And Me.NumShifts > 0 Then
        Dim lng_FirstMatchingRow As Long: lng_FirstMatchingRow = Examine_FindIndex(rng_ScenNums, int_TargetScen) ' WorksheetFunction.Match(int_TargetScen, rng_ScenNums, 0)
        Set dic_DataTypeMapping = map_Rules.Dict_DataTypes()

        If lng_FirstMatchingRow <> -1 Then
            Dim lng_NumRows As Long: lng_NumRows = WorksheetFunction.CountIf(rng_ScenNums, int_TargetScen)
            Dim lng_RowCtr As Long: lng_RowCtr = lng_FirstMatchingRow
            Dim lng_ActiveCount As Long
            Dim bln_WasValDateShift As Boolean
            Application.StatusBar = "Preparing scenario " & int_TargetScen & " - (0 of " & lng_NumRows & " shifts registered)"
            bln_ValDateReset = False

            ' Setup of shocks for target scenario
            While rng_ScenNums(lng_RowCtr, 1).Value = int_TargetScen
                If UCase(rng_IsActive(lng_RowCtr, 1).Value) = "YES" Then
                    lng_ActiveCount = lng_RowCtr - lng_FirstMatchingRow + 1
                    If lng_ActiveCount Mod 1000 = 0 Then
                        Application.StatusBar = "Preparing scenario " & int_TargetScen & " - (" & lng_ActiveCount & " of " _
                            & lng_NumRows & " shifts registered)"
                    End If

                    int_ScenLoaded = int_TargetScen

                    ' Read shocks from sheet
                    str_ActiveDataType = rng_DataTypes(lng_RowCtr, 1).Value
                    str_ActiveCode = rng_Codes(lng_RowCtr, 1).Value
                    dbl_ActiveShockAmt = rng_Values(lng_RowCtr, 1).Value
                    str_ActiveShockType = UCase(rng_ShiftTypes(lng_RowCtr, 1).Value)
                    Select Case UCase(rng_ShiftTypes(lng_RowCtr, 1).Value)
                        Case "ABS", "ABSOLUTE": enu_ActiveShockType = ShockType.Absolute
                        Case "REL", "RELATIVE": enu_ActiveShockType = ShockType.Relative
                    End Select

                    str_ActivePillarType = UCase(rng_PillarTypes(lng_RowCtr, 1).Value)
                    var_ActivePillar = rng_Pillars(lng_RowCtr, 1).Value
                    var_ActiveDelta = rng_Deltas(lng_RowCtr, 1).Value

                    If rng_IsPropagation(lng_RowCtr, 1).Value = "YES" Then bln_Propagate = True

                    ' Determine datatype codes
                    str_ActiveMappedType_Shift = dic_DataTypeMapping(str_ActiveDataType)
                    If str_ActiveMappedType_Shift = "CLT" Then
                        str_ActiveMappedType_Storage = "CVL"
                    Else
                        str_ActiveMappedType_Storage = str_ActiveMappedType_Shift
                    End If

                    ' Gather the curve object being shifted, and store in the set of shifted curves
                    Select Case str_ActiveMappedType_Storage
                        Case "FXS"
                            ' Gather FXS object directly from outer dictionary
                            Set fxs_AllSpots = dic_CurveSet(CurveType.FXSPT)
                        Case "EQS"
                            ' Gather EQS object directly from outer dictionary
                            Set eqs_AllSpots = dic_CurveSet(CurveType.EQSPT)
                        Case "EQVOL"
                            ' Gather EQVOL object directly from outer dictionary
                            Set eqv_AllVols = dic_CurveSet(CurveType.EQVOL)
                        Case "FXV", "IRC", "CVL", "SVL", "EVL"
                            ' Determine storage code name
                            If str_ActiveMappedType_Storage = "SVL" And str_ActivePillarType <> "UNIFORM_GROUP" Then
                                ' Code consists of name and underlying swap maturity label
                                str_UnderlyingSwap = Convert_Split(str_ActiveCode, "_", 3)
                                str_ActiveMappedCode = Convert_Split(str_ActiveCode, "_", 1)
                            Else
                                ' Code consists of name only
                                str_ActiveMappedCode = str_ActiveCode
                            End If

                            ' Gather curve object of the specified code from the dictionary of the specified type
                            Set var_ActiveCurve = dic_CurveSet(GetCurveType(str_ActiveMappedType_Storage))(str_ActiveMappedCode)
                    End Select

                    ' Load shocks into market data
                    Select Case str_ActiveMappedType_Shift
                        Case "FXS"
                            Select Case str_ActivePillarType
                                Case "NATIVE": Call fxs_AllSpots.Scen_AddNativeShock(str_ActiveCode, str_ActiveShockType, dbl_ActiveShockAmt)
                                Case Else: Call fxs_AllSpots.Scen_AddShock(str_ActiveCode, str_ActiveShockType, dbl_ActiveShockAmt)
                            End Select
                        Case "FXV"
                            Set fxv_ActiveCurve = var_ActiveCurve

                            Select Case str_ActivePillarType
                                Case "DAYS": Call fxv_ActiveCurve.Scen_AddShock_Days(CInt(var_ActivePillar), enu_ActiveShockType, dbl_ActiveShockAmt)
                                Case "DAYS_DELTA": Call fxv_ActiveCurve.Scen_AddShock_DaysDelta(CInt(var_ActivePillar), CInt(var_ActiveDelta), enu_ActiveShockType, dbl_ActiveShockAmt)
                                Case "LABEL": Call fxv_ActiveCurve.Scen_AddShock_Term(CStr(var_ActivePillar), enu_ActiveShockType, dbl_ActiveShockAmt)
                                Case "LABEL_DELTA": Call fxv_ActiveCurve.Scen_AddShock_TermDelta(CStr(var_ActivePillar), CInt(var_ActiveDelta), enu_ActiveShockType, dbl_ActiveShockAmt)
                                Case "DELTA": Call fxv_ActiveCurve.Scen_AddShock_Delta(CInt(var_ActiveDelta), enu_ActiveShockType, dbl_ActiveShockAmt)
                                Case "UNIFORM": Call fxv_ActiveCurve.Scen_AddShock_Uniform(enu_ActiveShockType, dbl_ActiveShockAmt)
                            End Select
                        Case "IRC"
                            Set irc_ActiveCurve = var_ActiveCurve

                            Select Case str_ActivePillarType
                                Case "DAYS": Call irc_ActiveCurve.Scen_AddByDays(CInt(var_ActivePillar), str_ActiveShockType, dbl_ActiveShockAmt)
                                Case "LABEL": Call irc_ActiveCurve.Scen_AddByLabel(CStr(var_ActivePillar), str_ActiveShockType, dbl_ActiveShockAmt)
                                Case "UNIFORM": Call irc_ActiveCurve.Scen_AddUniform(str_ActiveShockType, dbl_ActiveShockAmt)
                            End Select
                        Case "EQS"
                            Select Case str_ActivePillarType
                                Case "MARKET"
                                    Call eqs_AllSpots.Scen_AddShockMarket(str_ActiveCode, enu_ActiveShockType, dbl_ActiveShockAmt)
                                Case Else
                                    Call eqs_AllSpots.Scen_AddShock(str_ActiveCode, enu_ActiveShockType, dbl_ActiveShockAmt)
                            End Select
                        Case "EQVOL"

                            Select Case str_ActivePillarType
                                Case "MARKET"
                                    Call eqv_AllVols.Scen_AddShockMarket(str_ActiveCode, enu_ActiveShockType, dbl_ActiveShockAmt)
                                Case Else
                                    Call eqv_AllVols.Scen_AddShock(str_ActiveCode, enu_ActiveShockType, dbl_ActiveShockAmt)
                            End Select

                        Case "EVL"  'Matt edit start
                            Set evl_ActiveCurve = var_ActiveCurve
                            Select Case str_ActivePillarType
                                Case "DAYS_SMILE": Call evl_ActiveCurve.Scen_AddShock_Days(var_ActiveDelta, CInt(var_ActivePillar), enu_ActiveShockType, dbl_ActiveShockAmt)
                                Case "UNIFORM_STRIKE": Call evl_ActiveCurve.Scen_AddShock_UniformStrike(var_ActiveDelta, enu_ActiveShockType, dbl_ActiveShockAmt)
                                Case "UNIFORM_ALL": Call evl_ActiveCurve.Scen_AddShock_UniformAll(enu_ActiveShockType, dbl_ActiveShockAmt)
                            End Select

                        Case "CVL"
                            ' Cap vol shift
                            Set cvl_ActiveCurve = var_ActiveCurve
                            Call cvl_ActiveCurve.SetShockInst("CAP")

                            Select Case str_ActivePillarType
                                Case "DAYS": Call cvl_ActiveCurve.Scen_AddByDays(CInt(var_ActivePillar), enu_ActiveShockType, dbl_ActiveShockAmt)
                                Case "LABEL": Call cvl_ActiveCurve.Scen_AddByTerm(CStr(var_ActivePillar), enu_ActiveShockType, dbl_ActiveShockAmt)
                                Case "UNIFORM": Call cvl_ActiveCurve.Scen_AddUniform(enu_ActiveShockType, dbl_ActiveShockAmt)
                            End Select
                        Case "CLT"
                            ' Caplet vol shift
                            Set cvl_ActiveCurve = var_ActiveCurve

                            ' Only allow caplet vol shifts if curve bootstraps input cap vols
                            If cvl_ActiveCurve.IsBootstrappable = True Then
                                Call cvl_ActiveCurve.SetShockInst("CAPLET")

                                Select Case str_ActivePillarType
                                    Case "DAYS": Call cvl_ActiveCurve.Scen_AddByDays(CInt(var_ActivePillar), enu_ActiveShockType, dbl_ActiveShockAmt)
                                    Case "LABEL": Call cvl_ActiveCurve.Scen_AddByTerm(CStr(var_ActivePillar), enu_ActiveShockType, dbl_ActiveShockAmt)
                                    Case "UNIFORM": Call cvl_ActiveCurve.Scen_AddUniform(enu_ActiveShockType, dbl_ActiveShockAmt)
                                End Select
                            End If
                        Case "SVL"
                            Set svc_ActiveCurve = var_ActiveCurve

                            Select Case str_ActivePillarType
                                Case "UNIFORM_GROUP": Call svc_ActiveCurve.Scen_AddShock_UniformAll(enu_ActiveShockType, dbl_ActiveShockAmt)
                                Case "DAYS": Call svc_ActiveCurve.Scen_AddShock_Days(str_UnderlyingSwap, CInt(var_ActivePillar), enu_ActiveShockType, dbl_ActiveShockAmt)
                                Case "DAYS_SMILE": Call svc_ActiveCurve.Scen_AddShock_Days(str_UnderlyingSwap, _
                                    CInt(var_ActivePillar), enu_ActiveShockType, dbl_ActiveShockAmt, CDbl(var_ActiveDelta))
                                Case "LABEL": Call svc_ActiveCurve.Scen_AddShock_Pillar(str_UnderlyingSwap, CStr(var_ActivePillar), enu_ActiveShockType, dbl_ActiveShockAmt)
                                Case "LABEL_SMILE": Call svc_ActiveCurve.Scen_AddShock_Pillar(str_UnderlyingSwap, _
                                    CStr(var_ActivePillar), enu_ActiveShockType, dbl_ActiveShockAmt, CDbl(var_ActiveDelta))
                                Case "UNIFORM_CURVE": Call svc_ActiveCurve.Scen_AddShock_Uniform(str_UnderlyingSwap, enu_ActiveShockType, dbl_ActiveShockAmt)
                                Case "CURVE_SMILE": Call svc_ActiveCurve.Scen_AddShock_Uniform(str_UnderlyingSwap, enu_ActiveShockType, dbl_ActiveShockAmt, CDbl(var_ActiveDelta))
                            End Select
                        Case "DTV"
                            Select Case UCase(str_ActiveShockType)
                                Case "CALENDAR": lng_ValDate_Current = lng_ValDate_Orig + CLng(dbl_ActiveShockAmt)
                                Case "BUSINESS"
                                    cal_Active = cas_Calendars.Lookup_Calendar(str_ActiveCode)
                                    lng_ValDate_Current = date_workday(lng_ValDate_Orig, CInt(dbl_ActiveShockAmt), _
                                        cal_Active.HolDates, cal_Active.Weekends)
                            End Select

                            bln_ValDateReset = bln_ValDateReset Or cfg_Settings.SetCurrentValDate(lng_ValDate_Current)  ' Any true means the flag stays true

                        Case "DTB"
                            Select Case UCase(str_ActiveShockType)
                                Case "CALENDAR": cfg_Settings.CurrentBuildDate = lng_OrigBuildDate + CLng(dbl_ActiveShockAmt)
                                Case "BUSINESS": cfg_Settings.CurrentBuildDate = date_workday(lng_OrigBuildDate, CInt(dbl_ActiveShockAmt))
                            End Select

                            bln_BuildDateShift = True
                        Case "DTD"
                            Select Case UCase(str_ActiveShockType)
                                Case "CALENDAR": cfg_Settings.CurrentDataDate = lng_OrigDataDate + CLng(dbl_ActiveShockAmt)
                                Case "BUSINESS": cfg_Settings.CurrentDataDate = date_workday(lng_OrigDataDate, CInt(dbl_ActiveShockAmt))
                            End Select

                            bln_DataDateShift = True
                        Case "ALR"
                            ' All interest rate curves
                            For Each var_ActiveCurveName In strArr_AllIRC
                                str_ActiveCurveName = CStr(var_ActiveCurveName)
                                Set irc_ActiveCurve = dic_CurveSet(CurveType.IRC)(str_ActiveCurveName)

                                Select Case str_ActivePillarType
                                    Case "DAYS": Call irc_ActiveCurve.Scen_AddByDays(CInt(var_ActivePillar), str_ActiveShockType, dbl_ActiveShockAmt)
                                    Case "LABEL": Call irc_ActiveCurve.Scen_AddByLabel(CStr(var_ActivePillar), str_ActiveShockType, dbl_ActiveShockAmt)
                                    Case "UNIFORM": Call irc_ActiveCurve.Scen_AddUniform(str_ActiveShockType, dbl_ActiveShockAmt)
                                End Select
                            Next var_ActiveCurveName
                    End Select
                End If

                lng_RowCtr = lng_RowCtr + 1
            Wend

            Application.StatusBar = "Applying scenario " & int_TargetScen
            If bln_DataDateShift = True Or bln_BuildDateShift = True Then Call cyRefreshAll

            Call dic_CurveSet(CurveType.EQSPT).Scen_ApplyCurrent
            Call dic_CurveSet(CurveType.EQVOL).Scen_ApplyCurrent
            Call dic_CurveSet(CurveType.FXSPT).Scen_ApplyCurrent
            Call ApplyListedShocks(dic_CurveSet(CurveType.FXV))
            If bln_ValDateReset = True Then Call HandleValDateChange
            Call ApplyListedIRCurveShocks(dic_CurveSet(CurveType.IRC), bln_Propagate)
            If bln_Propagate = True Then Call OperateOnCurves_Method(dic_CurveSet(CurveType.IRC), "Scen_ReceivePropagation")
            Call ApplyListedShocks(dic_CurveSet(CurveType.cvl))
            Call ApplyListedShocks(dic_CurveSet(CurveType.SVL))
            Call ApplyListedShocks(dic_CurveSet(CurveType.EVL))
        End If

        cfg_Settings.CurrentScen = int_TargetScen
    Else
        cfg_Settings.CurrentScen = 0
    End If

    Call Action_SetAppState(fld_AppState_Orig)
End Sub

Private Sub HandleValDateChange()
    ' Update internal dates within objects (required if using cached objects)
    Dim fxs_Spots As Data_FXSpots
    If Not dic_CurveSet Is Nothing Then
        Call SetCurveValDate(dic_CurveSet(CurveType.FXV), lng_ValDate_Current)

        ' Clear cached lookup values
        Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
        Call fxs_Spots.ResetCache_Lookups
    End If

    ' Review pillar dates
    Dim dic_Modes As Dictionary: Set dic_Modes = map_Rules.Dict_ThetaModes
    If UCase(dic_Modes("FXV")) = "SHIFTED PILLARS" Then Call OperateOnCurves_Method(dic_CurveSet(CurveType.FXV), "GeneratePillarDates", True)
    If UCase(dic_Modes("CVL")) = "SHIFTED PILLARS" Then Call UpdateCapVolPillarDates(dic_CurveSet(CurveType.cvl))
    If UCase(dic_Modes("SVL")) = "SHIFTED PILLARS" Then Call OperateOnCurves_Method(dic_CurveSet(CurveType.SVL), "GeneratePillarDates")
     If UCase(dic_Modes("EVL")) = "SHIFTED PILLARS" Then Call OperateOnCurves_Method(dic_CurveSet(CurveType.EVL), "GenerateMaturityDates", True) 'Alvin added Matt's 20180809
End Sub

Private Sub ResetShiftedCurves()
    ' ## Apply base scenario for all curves
    Call dic_CurveSet(CurveType.EQSPT).Scen_ApplyBase
    Call dic_CurveSet(CurveType.EQVOL).Scen_ApplyBase
    Call dic_CurveSet(CurveType.FXSPT).Scen_ApplyBase
    Call ClearShocks(dic_CurveSet(CurveType.FXV))
    Call ClearShocks(dic_CurveSet(CurveType.IRC))
    Call ClearShocks(dic_CurveSet(CurveType.cvl))
    Call ClearShocks(dic_CurveSet(CurveType.SVL))
    Call ClearShocks(dic_CurveSet(CurveType.EVL))

    ' Reset other variables
    lng_ValDate_Current = lng_ValDate_Orig
End Sub

Public Sub LoadBaseScenario()
    Application.StatusBar = "Base scenario"
    Call Me.LoadScenario(0)
End Sub

Public Sub ClearScenarios()
    ' ## Remove all scenarios from the sheet.  Must clear and not delete, because range stored in sct_Orig is affected by other containers
    Dim lng_NumRows As Long: lng_NumRows = Me.NumShifts
    If lng_NumRows > 0 Then Call Action_ClearBelow(rng_DataTopLeft, int_NumCols)
    Call AssignRanges
End Sub

Public Sub DownloadFromDB(str_Cont As String, int_ScenMin As Integer, int_ScenMax As Integer)
    ' ## Download specified scenario from scenario DB
    Dim str_SQL As String
    str_SQL = "SELECT ScenNum, ScenName, IsPropagation, DataType, Code, ShiftValue, ShiftType, PillarType, Pillar, SmileAxis, IsActive " _
        & "FROM Scenarios WHERE Container = '" & str_Cont & "' AND ScenNum >= " & int_ScenMin & " AND ScenNum <= " & int_ScenMax
    Call Action_Query_Access(str_DBPath, str_SQL, rng_DataTopLeft)
    Call AssignRanges
End Sub


' ## METHODS - PRIVATE
Private Sub AssignRanges()
    Set rng_DataTopLeft = wks_Location.Range("A4")
    Dim lng_NumShifts As Long: lng_NumShifts = Me.NumShifts

    If lng_NumShifts > 0 Then
        Set rng_ScenNums = rng_DataTopLeft.Resize(lng_NumShifts, 1)
        Set rng_IsPropagation = rng_ScenNums.Offset(0, 2)
        Set rng_DataTypes = rng_IsPropagation.Offset(0, 1)
        Set rng_Codes = rng_DataTypes.Offset(0, 1)
        Set rng_Values = rng_Codes.Offset(0, 1)
        Set rng_ShiftTypes = rng_Values.Offset(0, 1)
        Set rng_PillarTypes = rng_ShiftTypes.Offset(0, 1)
        Set rng_Pillars = rng_PillarTypes.Offset(0, 1)
        Set rng_Deltas = rng_Pillars.Offset(0, 1)
        Set rng_IsActive = rng_Deltas.Offset(0, 1)
    End If
End Sub