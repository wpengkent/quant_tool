Option Explicit

' ## MEMBER DATA
Private wks_Location As Worksheet
Private rng_QueryTopLeft As Range
Private rng_CapTerms As Range, rng_OrigDates_Cap As Range, rng_OrigDates_Caplet As Range
Private rng_OrigCapVols As Range, rng_OrigCapletVols As Range, rng_ShockedCapVols As Range
Private rng_FinalDates_Cap As Range, rng_FinalDates_Caplet As Range, rng_FinalVols As Range
Private rng_DaysTopLeft As Range, rng_RelShifts_TopLeft As Range, rng_AbsShifts_TopLeft As Range
Private rng_FinalBuildDate As Range, rng_SpotDate As Range, rng_ShockInst As Range
Private rng_FirstFixing As Range

' Dependent curves
Private dic_CurveSet As Dictionary

' Dynamic variables
Private lngArr_FinalDates_Cap() As Long, lngArr_FinalDates_Caplet() As Long, dblArr_FinalVols() As Double
Private csh_Shifts_Rel As CurveDaysShift, csh_Shifts_Abs As CurveDaysShift
Private dbl_VolShift_Sens As Double

' Static values
Private dic_GlobalStaticInfo As Dictionary, igs_Generators As IRGeneratorSet, cas_Calendars As CalendarSet
Private map_Rules As MappingRules, cfg_Settings As ConfigSheet
Private lng_BuildDate As Long, str_NameInDB As String, bln_Bootstrap As Boolean, int_SpotDays As Integer, int_Deduction As Integer
Private cal_Deduction As Calendar, str_Interp_Time As String, str_ShockInst As String, str_OnFail As String, fld_LegParams As IRLegParams
Private str_CurveName As String
Private Const dbl_MinVol As Double = 0.000001, dbl_MinStrike As Double = 0.0000000001


' ## INITIALIZATION
Public Sub Initialize(wks_Input As Worksheet, bln_DataExists As Boolean, Optional dic_StaticInfoInput As Dictionary = Nothing)
    Set wks_Location = wks_Input

    ' Static info
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    Set igs_Generators = dic_GlobalStaticInfo(StaticInfoType.IRGeneratorSet)
    Set cas_Calendars = dic_GlobalStaticInfo(StaticInfoType.CalendarSet)
    Set map_Rules = dic_GlobalStaticInfo(StaticInfoType.MappingRules)
    Set cfg_Settings = dic_GlobalStaticInfo(StaticInfoType.ConfigSheet)
    Set csh_Shifts_Rel = New CurveDaysShift
    Set csh_Shifts_Abs = New CurveDaysShift

    If bln_DataExists = True Then
        Call StoreStaticValues
        Call AssignRanges ' Used when reading a rate
    End If
End Sub


' ## PROPERTIES
Public Property Get IsBootstrappable() As Boolean
    IsBootstrappable = bln_Bootstrap
End Property

Public Property Get NumCaps() As Integer
    NumCaps = Examine_NumRows(rng_QueryTopLeft)
End Property

Public Property Get NumCaplets() As Integer
    NumCaplets = Examine_NumRows(rng_OrigDates_Caplet(1, 1))
End Property

Public Property Get TypeCode() As CurveType
    TypeCode = CurveType.cvl
End Property

Private Property Get NumPoints() As Integer
    Dim int_output As Integer
    Select Case UCase(rng_ShockInst.Value)
        Case "CAP": int_output = Me.NumCaps
        Case "CAPLET": int_output = Me.NumCaplets
        Case Else: int_output = 0
    End Select

    NumPoints = int_output
End Property

Public Property Get Deduction() As Integer
    Deduction = int_Deduction
End Property

Public Property Get DeductionCalendar() As Calendar
    DeductionCalendar = cal_Deduction
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


' ## METHODS - LOOKUPS
Public Function Lookup_Vol(lng_MatDate As Long, Optional intLst_IndexFilters As Collection = Nothing, Optional bln_ApplyDeduction As Boolean = True) As Double
    Dim dbl_Output As Double
    Dim lng_DeductedMat As Long

    ' Apply deduction formula to lookup date
    If int_Deduction <> 0 And bln_ApplyDeduction = True Then
        lng_DeductedMat = date_workday(lng_MatDate, int_Deduction, cal_Deduction.HolDates, cal_Deduction.Weekends)
    Else
        lng_DeductedMat = lng_MatDate
    End If

    ' Determine which array to use
    Dim lngArr_FinalDates_ToUse() As Long
    If bln_Bootstrap = True Then
        lngArr_FinalDates_ToUse = lngArr_FinalDates_Caplet
    Else
        lngArr_FinalDates_ToUse = lngArr_FinalDates_Cap
    End If

    ' Apply index filters
    Dim lngArr_FilteredDates As Variant, dblArr_FilteredVols As Variant
    If intLst_IndexFilters Is Nothing Then
        lngArr_FilteredDates = lngArr_FinalDates_ToUse
        dblArr_FilteredVols = dblArr_FinalVols
    Else
        lngArr_FilteredDates = Convert_FilterArr(lngArr_FinalDates_ToUse, intLst_IndexFilters)
        dblArr_FilteredVols = Convert_FilterArr(dblArr_FinalVols, intLst_IndexFilters)
    End If

    Select Case str_Interp_Time
        Case "LIN", "LINEAR":
            dbl_Output = Interp_Lin(lngArr_FilteredDates, dblArr_FilteredVols, lng_DeductedMat, True) + dbl_VolShift_Sens
        Case "V2T"
            dbl_Output = Interp_V2t(lngArr_FilteredDates, dblArr_FilteredVols, rng_FinalBuildDate.Value, lng_DeductedMat) + dbl_VolShift_Sens
    End Select

    If dbl_Output <= 0 Then dbl_Output = dbl_MinVol
    Lookup_Vol = dbl_Output
End Function

Public Function Lookup_VolSeries(int_FinalIndex As Integer, Optional intLst_InterpPillars As Collection = Nothing, _
    Optional bln_ApplyDeduction As Boolean = True) As Collection
    ' ## Obtain collection of caplet vols
    Dim dblLst_output As New Collection
    Dim int_CapletCtr As Integer

    For int_CapletCtr = 1 To int_FinalIndex
        Call dblLst_output.Add(Me.Lookup_Vol(rng_FinalDates_Caplet(int_CapletCtr, 1).Value, intLst_InterpPillars, bln_ApplyDeduction))
    Next int_CapletCtr

    If int_FinalIndex > 0 Then
        ' Use second vol as first vol since extrapolation is flat
        Call dblLst_output.Add(dblLst_output(1), , 1)
    End If

    Set Lookup_VolSeries = dblLst_output
End Function


' ## METHODS - SETUP
Public Sub LoadRates()
    ' Read parameters
    With wks_Location
        Dim str_OptionalExclusions As String

        If .Range("C2").Value = "-" Then
            str_OptionalExclusions = ""
        Else
            str_OptionalExclusions = "AND SortTerm NOT IN (" & Replace(.Range("C2").Value, "|", ", ") & ") "
        End If

        Dim str_SQLCode As String
        Set rng_QueryTopLeft = .Range("A7")
    End With

    ' Determine table name
    Debug.Assert map_Rules.Dict_SourceTables.Exists("CAPVOL")
    Dim str_TableName As String: str_TableName = map_Rules.Dict_SourceTables("CAPVOL")

    ' Query
    str_SQLCode = "SELECT [Data Date], Currency, Term, SortTerm, Rate " _
            & "FROM " & str_TableName _
            & " WHERE [Data Date] = #" & Convert_SQLDate(cfg_Settings.CurrentDataDate) & "# AND Currency = '" & str_NameInDB & "' " & str_OptionalExclusions _
            & "ORDER BY [SortTerm];"

    Me.ClearCurve
    Call Action_Query_Access(cfg_Settings.RatesDBPath, str_SQLCode, rng_QueryTopLeft)
    Call AssignRanges

    ' Set instrument being shocked (cap or caplet) and set up associated ranges
    rng_ShockedCapVols.Value = rng_OrigCapVols.Value

    ' Derive cap vol pillar dates
    Call GeneratePillarDates(True, True, True)
    Dim cal_pmt As Calendar: cal_pmt = cas_Calendars.Lookup_Calendar(fld_LegParams.PmtCal)

    If bln_Bootstrap = True Then
        Call Me.Bootstrap(True)
    Else
        rng_FinalVols.Value = rng_OrigCapVols.Value
    End If
End Sub

Public Sub SetParams(rng_QueryParams As Range)
    With wks_Location
        .Range("A2").Value = cfg_Settings.CurrentBuildDate
        .Range("B2:K2").Value = rng_QueryParams.Value
    End With

    Call StoreStaticValues
End Sub

Private Sub StoreStaticValues()
    ' ## Read static values from the sheet and store in memory
    Dim rng_FirstParam As Range: Set rng_FirstParam = wks_Location.Range("A2")
    Dim int_ActiveCol As Integer: int_ActiveCol = 0
    lng_BuildDate = rng_FirstParam.Offset(0, int_ActiveCol).Value

    int_ActiveCol = int_ActiveCol + 1
    str_NameInDB = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)
    str_CurveName = Right(wks_Location.Name, Len(wks_Location.Name) - 4)

    int_ActiveCol = int_ActiveCol + 2
    bln_Bootstrap = (UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value) = "YES")

    int_ActiveCol = int_ActiveCol + 1
    fld_LegParams = igs_Generators.Lookup_Generator(UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value))

    int_ActiveCol = int_ActiveCol + 1
    int_SpotDays = rng_FirstParam.Offset(0, int_ActiveCol).Value

    int_ActiveCol = int_ActiveCol + 1
    int_Deduction = rng_FirstParam.Offset(0, int_ActiveCol).Value  ' Coupon period start minus deduction gives the option maturity

    int_ActiveCol = int_ActiveCol + 1
    cal_Deduction = cas_Calendars.Lookup_Calendar(UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value))

    int_ActiveCol = int_ActiveCol + 1
    str_Interp_Time = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)

    int_ActiveCol = int_ActiveCol + 1
    str_OnFail = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)
End Sub

Public Sub ClearCurve()
    With wks_Location
        Call Action_ClearBelow(.Range("A7"), 6)
        Call Action_ClearBelow(.Range("H7"), 2)
        Call Action_ClearBelow(.Range("K7"), 3)
        Call Action_ClearBelow(.Range("O7"), 1)
        Call Action_ClearBelow(.Range("P7"), 2)
        .Range("O2:Q2").ClearContents
    End With
End Sub

Private Sub AssignRanges()
    Dim int_NumCaps As Integer: int_NumCaps = Examine_NumRows(wks_Location.Range("A7"))
    If int_NumCaps > 0 Then
        Set rng_FirstFixing = wks_Location.Range("K2")
        Set rng_FinalBuildDate = wks_Location.Range("O2")
        Set rng_SpotDate = rng_FinalBuildDate.Offset(0, 1)
        Set rng_ShockInst = rng_SpotDate.Offset(0, 1)
        Set rng_QueryTopLeft = wks_Location.Range("A7")
        Set rng_CapTerms = rng_QueryTopLeft.Offset(0, 2).Resize(int_NumCaps, 1)
        Set rng_OrigCapVols = rng_CapTerms.Offset(0, 2)
        Set rng_OrigDates_Cap = rng_OrigCapVols.Offset(0, 1)
        Set rng_DaysTopLeft = wks_Location.Range("K7")
        Set rng_RelShifts_TopLeft = rng_DaysTopLeft.Offset(0, 1)
        Set rng_AbsShifts_TopLeft = rng_RelShifts_TopLeft.Offset(0, 1)

        Dim int_NumPoints As Integer
        If Me.IsBootstrappable = True Then
            Dim str_LastCapMat As String: str_LastCapMat = rng_CapTerms(int_NumCaps, 1).Value
            Dim int_NumCaplets As Integer: int_NumCaplets = Calc_NumPeriods(str_LastCapMat, fld_LegParams.PmtFreq) - 1
            int_NumPoints = int_NumCaplets

            Set rng_OrigDates_Caplet = rng_OrigDates_Cap(1, 1).Offset(0, 2).Resize(int_NumCaplets, 1)
            Set rng_OrigCapletVols = rng_OrigDates_Caplet.Offset(0, 1)
        Else
            int_NumPoints = int_NumCaps
        End If

        Set rng_FinalDates_Cap = wks_Location.Range("O7").Resize(int_NumCaps, 1)
        Set rng_ShockedCapVols = rng_FinalDates_Cap.Offset(0, 1)
        Set rng_FinalDates_Caplet = rng_ShockedCapVols(1, 1).Offset(0, 1).Resize(int_NumPoints, 1)
        Set rng_FinalVols = rng_FinalDates_Caplet.Offset(0, 1)

        ' Fill interpolation cache
        lngArr_FinalDates_Cap = Convert_RangeToLngArr(rng_FinalDates_Cap)
        lngArr_FinalDates_Caplet = Convert_RangeToLngArr(rng_FinalDates_Caplet)
        dblArr_FinalVols = Convert_RangeToDblArr(rng_FinalVols)
    End If
End Sub

Public Sub GeneratePillarDates(bln_GenCapDates As Boolean, bln_GenCapletDates As Boolean, bln_StoreAsOrig As Boolean)
    ' ## Derive cap vol pillar dates, building from the current valuation date
    Dim cal_pmt As Calendar: cal_pmt = cas_Calendars.Lookup_Calendar(fld_LegParams.PmtCal)
    Dim lng_ValDate As Long: lng_ValDate = cfg_Settings.CurrentValDate
    Dim int_NumCapVols As Integer: int_NumCapVols = Me.NumCaps
    Dim int_NumCapletVols As Integer: int_NumCapletVols = rng_OrigDates_Caplet.Rows.count
    Dim lng_ActivePillarDate As Long
    Dim int_ctr As Integer

    ' Determine spot date
    Dim lng_SpotDate As Long
    If int_SpotDays = 0 Then
        lng_SpotDate = Date_ApplyBDC(lng_ValDate, "FOLL", cal_pmt.HolDates, cal_pmt.Weekends)
    Else
        lng_SpotDate = date_workday(lng_ValDate, int_SpotDays, cal_pmt.HolDates, cal_pmt.Weekends)
    End If

    rng_FinalBuildDate.Value = lng_ValDate
    rng_SpotDate.Value = lng_SpotDate

    ReDim lngArr_FinalDates_Cap(1 To int_NumCapVols) As Long
    ReDim lngArr_FinalDates_Caplet(1 To int_NumCapletVols) As Long

    ' Generate cap vol dates if depending on these
    If bln_GenCapDates = True Then
        For int_ctr = 1 To int_NumCapVols
            lngArr_FinalDates_Cap(int_ctr) = date_addterm(lng_SpotDate, rng_CapTerms(int_ctr, 1).Value, 1, True)
        Next int_ctr

        rng_FinalDates_Cap.Resize(int_NumCapVols, 1).Value = Convert_Array1Dto2D(lngArr_FinalDates_Cap)
        If bln_StoreAsOrig = True Then rng_OrigDates_Cap.Value = rng_FinalDates_Cap.Value
    End If

    If Me.IsBootstrappable = True Then
        If bln_GenCapletDates = True Then
            ' Derive caplet vol pillar dates
            For int_ctr = 1 To int_NumCapletVols
                lng_ActivePillarDate = date_addterm(lng_SpotDate, fld_LegParams.PmtFreq, int_ctr, True)
                lng_ActivePillarDate = Date_ApplyBDC(lng_ActivePillarDate, fld_LegParams.BDC, cal_pmt.HolDates, cal_pmt.Weekends)
                lngArr_FinalDates_Caplet(int_ctr) = date_workday(lng_ActivePillarDate, int_Deduction, cal_Deduction.HolDates, cal_Deduction.Weekends)
            Next int_ctr

            rng_FinalDates_Caplet.Resize(int_NumCapletVols, 1).Value = Convert_Array1Dto2D(lngArr_FinalDates_Caplet)
            If bln_StoreAsOrig = True Then rng_OrigDates_Caplet.Value = rng_FinalDates_Caplet.Value
        End If
    End If
End Sub

Public Sub Bootstrap(bln_CopyToOrig As Boolean)
    Dim bln_ScreenUpdating As Boolean: bln_ScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim lng_ValueDate As Long: lng_ValueDate = rng_FinalBuildDate.Value
    Dim lng_SpotDate As Long: lng_SpotDate = rng_SpotDate.Value
    Dim fld_Params As InstParams_IRS
    Dim fld_LegA As IRLegParams: fld_LegA = fld_LegParams
    Dim fld_LegB As IRLegParams: fld_LegB = fld_LegParams
    Dim irs_Active As Inst_IRSwap
    Dim irl_Floating As IRLeg
    Application.StatusBar = "Data date: " & Format(cfg_Settings.CurrentDataDate, "dd/mm/yyyy") & "     CVL: " & str_CurveName & " (bootstrapping)"
    Dim enu_Direction As OptionDirection: enu_Direction = OptionDirection.CallOpt

    ' Ensure curve dependencies are available
    If dic_CurveSet Is Nothing Then Set dic_CurveSet = GetAllCurves(True, False, dic_GlobalStaticInfo)

    ' Fill parameters for underlying
    ' Leg A - fixed
    With fld_LegA
        .ValueDate = lng_ValueDate
        .Swapstart = lng_SpotDate
        .GenerationRefPoint = lng_SpotDate
        .IsFwdGeneration = True
        .PExch_Start = False
        .PExch_Intermediate = False
        .PExch_End = False
        .FloatEst = True
        .Notional = 1000000
        .index = "-"
        .estcal = "-"
        .Curve_Est = "-"
    End With

    ' Leg B - floating
    With fld_LegB
        .ValueDate = lng_ValueDate
        .Swapstart = lng_SpotDate
        .GenerationRefPoint = lng_SpotDate
        .IsFwdGeneration = True
        .PExch_Start = False
        .PExch_Intermediate = False
        .PExch_End = False
        .FloatEst = True
        .Notional = 1000000

        ' Apply customized first fixing if it exists
        If rng_FirstFixing.Value <> "-" And lng_ValueDate = lng_BuildDate Then
            Dim dic_Fixings As New Dictionary
            Call dic_Fixings.Add(lng_SpotDate, CDbl(rng_FirstFixing.Value))
            Set .Fixings = dic_Fixings
        End If
    End With

    ' Shared
    With fld_Params
        .Pay_LegA = True
        .CCY_PnL = fld_LegParams.CCY
        .LegA = fld_LegA
        .LegB = fld_LegB
    End With

    ' Loop through cap pillars
    Dim int_NumCaps As Integer: int_NumCaps = Me.NumCaps
    Dim int_NumCaplets As Integer: int_NumCaplets = Me.NumCaplets
    Dim int_CapCtr As Integer, int_CapletCtr As Integer
    Dim dbl_ActiveATMStrike As Double, dbl_ActiveCapVol As Double, lng_ActiveCapMatDate As Long
    Dim dbl_ActiveCapPrice As Double, dbl_ActiveCapletStripPrice As Double
    Dim int_ActiveFinalIndex As Integer
    Dim intLst_InterpPillars As New Collection
    Dim dblLst_CapletVols As New Collection
    Dim dbl_ActiveCapletVol As Double
    Dim dbl_PrevCapletVol As Double: dbl_PrevCapletVol = 0
    ReDim dblArr_FinalVols(1 To int_NumCaplets) As Double  ' Reset final vols
    Dim intLst_FailedPoints As New Collection

    ' Secant method variables
    Dim dic_SecantParams As Dictionary, dic_SecantOutputs As Dictionary
    Dim bln_SolutionPossible As Boolean
    Dim int_PrevFinalIndex As Integer: int_PrevFinalIndex = 0

    For int_CapCtr = 1 To int_NumCaps
        fld_Params.LegA.Term = rng_CapTerms(int_CapCtr, 1).Value
        fld_Params.LegB.Term = fld_Params.LegA.Term
        Set irs_Active = GetInst_IRS(fld_Params, dic_CurveSet, dic_GlobalStaticInfo)
        Set irl_Floating = irs_Active.LegB

        dbl_ActiveATMStrike = irs_Active.ParRate_LegA
        If dbl_ActiveATMStrike < dbl_MinStrike Then dbl_ActiveATMStrike = dbl_MinStrike
        dbl_ActiveCapVol = rng_ShockedCapVols(int_CapCtr, 1).Value

        dbl_ActiveCapPrice = irl_Floating.Calc_BSOptionValue(enu_Direction, dbl_ActiveATMStrike, int_Deduction, cal_Deduction, True, , _
            dbl_ActiveCapVol)

        ' Find index of latest caplet falling within the cap period
        int_ActiveFinalIndex = Calc_NumPeriods(irl_Floating.Params.Term, fld_LegParams.PmtFreq) - 1
        Call intLst_InterpPillars.Add(int_ActiveFinalIndex)

        ' Store static parameters for secant solver
        Set dic_SecantParams = New Dictionary
        Call dic_SecantParams.Add("cvl_Curve", Me)
        Call dic_SecantParams.Add("irl_Underlying", irl_Floating)
        Call dic_SecantParams.Add("int_FinalIndex", int_ActiveFinalIndex)
        Call dic_SecantParams.Add("intLst_InterpPillars", intLst_InterpPillars)
        Call dic_SecantParams.Add("enu_Direction", enu_Direction)
        Call dic_SecantParams.Add("dbl_ATMStrike", dbl_ActiveATMStrike)
        Call dic_SecantParams.Add("int_Deduction", int_Deduction)
        Call dic_SecantParams.Add("rng_HolDates", cal_Deduction.HolDates)
        Call dic_SecantParams.Add("str_Weekends", cal_Deduction.Weekends)

        ' Solve using secant method
        Set dic_SecantOutputs = New Dictionary
        dbl_ActiveCapletVol = Solve_Secant(ThisWorkbook, "SolverFuncXY_CapletVolToPrice", dic_SecantParams, _
            dbl_ActiveCapVol, dbl_ActiveCapVol + 1, dbl_ActiveCapPrice, 0.0000000001, 50, -1, dic_SecantOutputs)

        ' Final solution will be shown in the cell, if no solution found, show error value
        If dic_SecantOutputs("Solvable") = True And dbl_ActiveCapletVol > 0 Then
            dbl_PrevCapletVol = dbl_ActiveCapletVol
        Else
            Call intLst_FailedPoints.Add(int_ActiveFinalIndex)
            Debug.Print "## ERROR - Caplet volatility could not be solved for " & str_CurveName & " " _
                & irl_Floating.Params.Term

            ' Fall back to the previous pillar vol
            If int_PrevFinalIndex <> 0 Then
                Select Case str_OnFail
                    Case "FLAT": Call Me.SetFinalVol(int_ActiveFinalIndex, dbl_PrevCapletVol)
                    Case "ZERO", "PAR": Call Me.SetFinalVol(int_ActiveFinalIndex, dbl_MinVol)
                End Select
            End If
        End If

        int_PrevFinalIndex = int_ActiveFinalIndex
    Next int_CapCtr

    ' Fill in interpolated caplet pillars
    For int_CapletCtr = 1 To int_NumCaplets
        dblArr_FinalVols(int_CapletCtr) = Me.Lookup_Vol(lngArr_FinalDates_Caplet(int_CapletCtr), intLst_InterpPillars, False)
    Next int_CapletCtr


    If str_OnFail = "PAR" Then
        ' After interpolation, replace failed pillars with interpolated par vols
        Dim dblArr_CapVols() As Double: dblArr_CapVols = Convert_RangeToDblArr(rng_ShockedCapVols)
        Dim var_ActiveIndex As Variant, lng_ActiveCapletMat As Long

        For Each var_ActiveIndex In intLst_FailedPoints
            lng_ActiveCapletMat = lngArr_FinalDates_Caplet(var_ActiveIndex)
            dblArr_FinalVols(var_ActiveIndex) = Interp_Lin(lngArr_FinalDates_Cap, dblArr_CapVols, lng_ActiveCapletMat, True)
        Next var_ActiveIndex

        ' Check zero vols should really be zero by re-interpolating
        For int_CapletCtr = 1 To int_NumCaplets
            If Round(dblArr_FinalVols(int_CapletCtr), 8) = dbl_MinVol Then
                dblArr_FinalVols(int_CapletCtr) = Me.Lookup_Vol(lngArr_FinalDates_Caplet(int_CapletCtr), intLst_InterpPillars, False)
            End If
        Next int_CapletCtr
    End If

    ' Output to sheet
    rng_FinalVols.Value = Convert_Array1Dto2D(dblArr_FinalVols)
    If bln_CopyToOrig = True Then rng_OrigCapletVols.Value = rng_FinalVols.Value

    Application.StatusBar = False
    Application.ScreenUpdating = bln_ScreenUpdating
End Sub


' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Sub SetFinalVol(int_Index As Integer, dbl_NewVol As Double)
    ' ## Used by bootstrapping solver function
    dblArr_FinalVols(int_Index) = dbl_NewVol
End Sub

Public Sub SetShockInst(str_Instrument As String)
    rng_ShockInst.Value = str_Instrument
End Sub

Public Sub FillDependency_AllCurves(dic_Curves As Dictionary)
    ' ## Store curve set, so it does not have to be created each time a value is looked up
    Set dic_CurveSet = dic_Curves
End Sub


' ## METHODS - SCENARIOS
Public Sub Scen_AddByTerm(str_Term As String, enu_ShockType As ShockType, dbl_Amount As Double)
    Dim int_Index As Integer
    Dim int_NumMonths As Integer, int_FreqInMonths As Integer, int_numdays As Integer

    Select Case rng_ShockInst.Value
        Case "CAP"
            int_Index = Examine_FindIndex(Convert_RangeToList(rng_CapTerms), str_Term)
            int_numdays = rng_OrigDates_Cap(int_Index, 1).Value - lng_BuildDate
        Case "CAPLET"
            int_NumMonths = calc_nummonths(str_Term)
            int_FreqInMonths = calc_nummonths(fld_LegParams.PmtFreq)

            Debug.Assert (int_NumMonths Mod int_FreqInMonths = 0)
            int_Index = int_NumMonths / int_FreqInMonths
            int_numdays = rng_OrigDates_Caplet(int_Index, 1).Value - lng_BuildDate
        Case Else
            Debug.Assert False
    End Select

    If int_Index > 0 Then
        Select Case enu_ShockType
            Case ShockType.Absolute: Call csh_Shifts_Abs.AddIsolatedShift(int_numdays, dbl_Amount)
            Case ShockType.Relative: Call csh_Shifts_Rel.AddIsolatedShift(int_numdays, dbl_Amount)
        End Select
    End If
End Sub

Public Sub Scen_AddByDays(int_numdays As Integer, enu_ShockType As ShockType, dbl_Amount As Double)
    Select Case enu_ShockType
        Case ShockType.Absolute: Call csh_Shifts_Abs.AddShift(int_numdays, dbl_Amount)
        Case ShockType.Relative: Call csh_Shifts_Rel.AddShift(int_numdays, dbl_Amount)
    End Select
End Sub

Public Sub Scen_AddUniform(enu_ShockType As ShockType, dbl_Amount As Double)
    ' ## Shock entire curve by the same amount
    Select Case enu_ShockType
        Case ShockType.Absolute: Call csh_Shifts_Abs.AddUniformShift(dbl_Amount)
        Case ShockType.Relative: Call csh_Shifts_Rel.AddUniformShift(dbl_Amount)
    End Select
End Sub

Public Sub Scen_ApplyBase()
    ' Clear shifts
    Call csh_Shifts_Rel.Initialize(ShockType.Relative)
    Call csh_Shifts_Abs.Initialize(ShockType.Absolute)

    Call Action_ClearBelow(rng_DaysTopLeft, 3)

    ' Reset vols
    If Me.IsBootstrappable = True Then
        rng_FinalVols.Value = rng_OrigCapletVols.Value
    Else
        rng_FinalVols.Value = rng_OrigCapVols.Value
    End If

    rng_ShockedCapVols.Value = rng_OrigCapVols.Value
    rng_ShockInst.ClearContents

    ' Read original vols back into memory
    dblArr_FinalVols = Convert_RangeToDblArr(rng_FinalVols)
End Sub

Public Sub Scen_ApplyCurrent()
    Dim int_NumPoints As Integer: int_NumPoints = NumPoints()
    Dim int_NumShifts_Rel As Integer: int_NumShifts_Rel = csh_Shifts_Rel.NumShifts
    Dim int_NumShifts_Abs As Integer: int_NumShifts_Abs = csh_Shifts_Abs.NumShifts
    Dim bln_Bootstrap As Boolean: bln_Bootstrap = Me.IsBootstrappable

    If int_NumShifts_Rel + int_NumShifts_Abs > 0 Then
        ' Update final vols
        Dim str_ShockInst As String: str_ShockInst = rng_ShockInst.Value
        Dim int_ActiveDTM As Integer
        Dim dbl_ActiveDaysAbs As Double, dbl_ActiveDaysRel As Double
        Dim int_RowCtr As Integer
        Dim rng_Orig As Range, rng_Final As Range
        Dim dblArr_OrigVols() As Variant
        Dim dblArr_ShockedCapVols() As Double: ReDim dblArr_ShockedCapVols(1 To int_NumPoints, 1 To 1) As Double

        ' Convert shocks for number of days to shocks for each pillar
        Select Case str_ShockInst
            Case "CAP"
                ' Determine shifts
                dblArr_OrigVols = rng_OrigCapVols.Value2
                For int_RowCtr = 1 To int_NumPoints
                    int_ActiveDTM = lngArr_FinalDates_Cap(int_RowCtr) - lng_BuildDate
                    dblArr_ShockedCapVols(int_RowCtr, 1) = dblArr_OrigVols(int_RowCtr, 1) * (1 + csh_Shifts_Rel.ReadShift(int_ActiveDTM) / 100) _
                        + csh_Shifts_Abs.ReadShift(int_ActiveDTM)
                    If dblArr_ShockedCapVols(int_RowCtr, 1) <= 0 Then dblArr_ShockedCapVols(int_RowCtr, 1) = dbl_MinVol
                Next int_RowCtr

                ' Output to sheet and bootstrap if required
                rng_ShockedCapVols.Value = dblArr_ShockedCapVols
                If bln_Bootstrap = True Then
                    Call Me.Bootstrap(False)
                Else
                    rng_FinalVols.Value = rng_ShockedCapVols.Value
                End If
            Case "CAPLET"
                ' Determine shifts
                dblArr_OrigVols = rng_OrigCapletVols.Value2
                For int_RowCtr = 1 To int_NumPoints
                    int_ActiveDTM = lngArr_FinalDates_Caplet(int_RowCtr) - lng_BuildDate
                    dblArr_FinalVols(int_RowCtr) = dblArr_OrigVols(int_RowCtr, 1) * (1 + csh_Shifts_Rel.ReadShift(int_ActiveDTM) / 100) _
                        + csh_Shifts_Abs.ReadShift(int_ActiveDTM)
                    If dblArr_FinalVols(int_RowCtr) <= 0 Then dblArr_FinalVols(int_RowCtr) = dbl_MinVol
                Next int_RowCtr

                ' Output to sheet
                rng_FinalVols.Value = Convert_Array1Dto2D(dblArr_FinalVols)
            Case Else: Debug.Assert False
        End Select

        ' Output shifts to sheet
        If int_NumShifts_Rel > 0 Then
            rng_DaysTopLeft.Resize(int_NumShifts_Rel, 1).Value = csh_Shifts_Rel.Days_Arr
            rng_RelShifts_TopLeft.Resize(int_NumShifts_Rel, 1).Value = csh_Shifts_Rel.Shifts_Arr
        End If

        If int_NumShifts_Abs > 0 Then
            rng_DaysTopLeft.Resize(int_NumShifts_Abs, 1).Value = csh_Shifts_Abs.Days_Arr
            rng_AbsShifts_TopLeft.Resize(int_NumShifts_Abs, 1).Value = csh_Shifts_Abs.Shifts_Arr
        End If
    End If
End Sub


' ## METHODS - OUTPUT
Public Sub OutputFinalVols(rng_OutputStart As Range)
    With wks_Location
        Dim str_type As String: If bln_Bootstrap = True Then str_type = "Caplet" Else str_type = "Cap"
    End With

    Dim int_RowCtr As Integer
    For int_RowCtr = 1 To Me.NumCaplets
        rng_OutputStart(int_RowCtr, 1).Value = str_CurveName
        rng_OutputStart(int_RowCtr, 2).Value = lng_BuildDate
        If bln_Bootstrap = True Then
            rng_OutputStart(int_RowCtr, 3).Value = lngArr_FinalDates_Caplet(int_RowCtr) - lng_BuildDate
        Else
            rng_OutputStart(int_RowCtr, 3).Value = lngArr_FinalDates_Cap(int_RowCtr) - lng_BuildDate
        End If
        rng_OutputStart(int_RowCtr, 4).Value = dblArr_FinalVols(int_RowCtr)
        rng_OutputStart(int_RowCtr, 5).Value = str_type
        rng_OutputStart(int_RowCtr, 6).Value = "Native"
    Next int_RowCtr
End Sub

Public Sub OutputOrigCapVols(rng_OutputStart As Range)
    ' ## Return the vols queried from the database
    Dim int_RowCtr As Integer
    For int_RowCtr = 1 To Me.NumCaps
        rng_OutputStart(int_RowCtr, 1).Value = str_CurveName
        rng_OutputStart(int_RowCtr, 2).Value = lng_BuildDate
        rng_OutputStart(int_RowCtr, 3).Value = CLng(rng_OrigDates_Cap(int_RowCtr, 1).Value) - lng_BuildDate
        rng_OutputStart(int_RowCtr, 4).Value = rng_OrigCapVols(int_RowCtr, 1).Value
        rng_OutputStart(int_RowCtr, 5).Value = "Cap"
        rng_OutputStart(int_RowCtr, 6).Value = "Native"
    Next int_RowCtr
End Sub