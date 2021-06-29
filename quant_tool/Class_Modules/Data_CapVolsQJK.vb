Option Explicit

' ## MEMBER DATA
Private wks_Location As Worksheet
Private wks_LocationATMVolsName As String: Private wks_LocationATMVols As Worksheet 'QJK code 14/1/2015
Private rng_QueryTopLeft As Range
Private rng_CapTerms As Range, rng_OrigDates_Cap As Range, rng_OrigDates_Caplet As Range
Private rng_OrigCapVols As Range, rng_OrigCapletVols As Range, rng_ShockedCapVols As Range
Private rng_FinalDates_Cap As Range, rng_FinalDates_Caplet As Range, rng_FinalVols As Range
Private rng_DaysTopLeft As Range, rng_RelShifts_TopLeft As Range, rng_AbsShifts_TopLeft As Range
Private rng_FinalBuildDate As Range, rng_SpotDate As Range, rng_ShockInst As Range
Private rng_FirstFixing As Range
Private strikeQJK As Double, rngStrikeQJK As Range:
Private interpolateOn As String, rngInterpolateOn As Range:   'QJK code 16/12/2014
Private rng_QueryTopLeftATM As Range: Private rng_ATMDates As Range: Private rng_ATMVols As Range 'QJK code 14/1/2015
Private rng_ATMVolShocked As Range 'Mandy 15/1/2014
Private rng_CalibrationSolved As Range
' Dependent curves
Private dic_CurveSet As Dictionary

' Dynamic variables
Private lngArr_FinalDates_Cap() As Long, lngArr_FinalDates_Caplet() As Long, dblArr_FinalVols() As Double
Private lngArr_FinalDates_ATMDates() As Long: Private dblArr_FinalATMVols() As Double 'QJK code 14/1/2015
Private csh_Shifts_Rel As CurveDaysShift, csh_Shifts_Abs As CurveDaysShift
Private dbl_VolShift_Sens As Double
Private ParVolDates() As Double:
Private bln_CalibrationSolved() As Boolean   'QJK code 14/1/2015
'variable for interpolation between strike pillars
Private dblArr_PillarVols_1() As Double, dblArr_PillarVols_2() As Double
Private dbl_StrikePillar_1 As Double, dbl_StrikePillar_2 As Double
Private bln_ATMvols As Boolean: 'QJK code 28/1/2015 get the range right for FinalVols for K=0 ATM
' Static values
Private dic_GlobalStaticInfo As Dictionary, igs_Generators As IRGeneratorSet, cas_Calendars As CalendarSet
Private map_Rules As MappingRules, cfg_Settings As ConfigSheet:
Private lng_BuildDate As Long, str_NameInDB As String, bln_Bootstrap As Boolean, int_SpotDays As Integer, int_Deduction As Integer
Private cal_Deduction As Calendar, str_Interp_Time As String, str_ShockInst As String, str_OnFail As String, fld_LegParams As IRLegParams
Private str_CurveName As String
Private Const dbl_MinVol As Double = 0.000001, dbl_MinStrike As Double = 0.0000000001

Private fwdrate As Collection
Private int_PreviousCapletsNum As Integer
Private bln_calibrationfailed As Boolean



' ## INITIALIZATION
Public Sub Initialize(wks_Input As Worksheet, bln_DataExists As Boolean, Optional dic_StaticInfoInput As Dictionary = Nothing, _
Optional bln_betweenStrikePillar As Boolean = False, Optional bln_IsFirstStrikePillar As Boolean = True, Optional dbl_Strike As Double)
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
        If bln_betweenStrikePillar = False Then
            Call AssignRanges ' Used when reading a rate
        Else
            Call AssignRanges(bln_betweenStrikePillar, bln_IsFirstStrikePillar)
        End If

    End If

    'Interpolate the vols according to the strike
    If bln_betweenStrikePillar = True And dbl_StrikePillar_1 < dbl_StrikePillar_2 Then Call interpolateVols(dbl_Strike)

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

'QJK added function 21/10/2016
Public Function getStrCurveName() As String
'Dim tempString As String
'tempString = str_CurveName
getStrCurveName = str_CurveName
End Function
'qjk ADDED 26102016

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
       ' If interpolateOn = "FORWARD" Then 'QJK code 28/01/2015 SHOULD BE AUTOMATICALLY ANY THING THAT USES THIS SHOULD BE IN FORWARD MODE AND  using function BootstraP_QJK
        'lngArr_FilteredDates = lngArr_FinalDates_Cap 'QJK code 28/01/2015
        'dblArr_FilteredVols = Convert_FilterArr(dblArr_FinalVols, intLst_IndexFilters)  'QJK code 28/01/2015
        'Else    'no condition, in case not only for surface  'QJK code 28/01/2015
        lngArr_FilteredDates = Convert_FilterArr(lngArr_FinalDates_ToUse, intLst_IndexFilters)
        dblArr_FilteredVols = Convert_FilterArr(dblArr_FinalVols, intLst_IndexFilters)
        'End If 'QJK code 28/01/2015
    End If
 'If interpolateOn = "FORWARD" Then   'QJK CODE 17/12/2014
    Select Case str_Interp_Time
        Case "LIN", "LINEAR":
            dbl_Output = Interp_Lin(lngArr_FilteredDates, dblArr_FilteredVols, lng_DeductedMat, True) + dbl_VolShift_Sens
        Case "V2T"
            dbl_Output = Interp_V2t(lngArr_FilteredDates, dblArr_FilteredVols, rng_FinalBuildDate.Value, lng_DeductedMat) + dbl_VolShift_Sens
    End Select
  'Else 'QJK CODE 17/12/2014
  'dbl_Output =dblArr_FinalVols
  'End If                             'QJK CODE 17/12/2014
    'If dbl_Output <= 0 Then dbl_Output = dbl_MinVol
    If dbl_Output <= 0.000101 Then 'Mandy 15/1/2015  'CHANGED TO MINVOL QJK  'QJK ADDED 16012017
        Select Case str_Interp_Time
        Case "LIN", "LINEAR":
            'QJK added 08/11/2016 in case vega shock down is negative vol
            'qjk ADDED 16012017 added or condition for vega 1bp+minvol=minvol so doesnt calculate a value for this
            If (Interp_Lin(rng_ATMDates, rng_ATMVolShocked, lng_DeductedMat, True) = dbl_MinVol) Or (Interp_Lin(rng_ATMDates, rng_ATMVolShocked, lng_DeductedMat, True) <= 0.0001 + dbl_MinVol) Then
             dbl_Output = dbl_MinVol
            Else
            'end of QJK added 08/11/2016
            dbl_Output = Interp_Lin(rng_ATMDates, rng_ATMVolShocked, lng_DeductedMat, True) + dbl_VolShift_Sens
            End If
        Case "V2T"
            dbl_Output = Interp_V2t(rng_ATMDates, rng_ATMVolShocked, rng_FinalBuildDate.Value, lng_DeductedMat) + dbl_VolShift_Sens
        End Select
    End If
    Lookup_Vol = dbl_Output
   ' Debug.Print "dbl_output is:" & dbl_Output & "and vol date is:" & CDate(lng_DeductedMat)
End Function
Public Function Lookup_VolCFSurfaceInterpolateOnFWD(lng_MatDate As Long, Optional intLst_IndexFilters As Collection = Nothing, Optional bln_ApplyDeduction As Boolean = True) As Double
    'QJK code 02/02/2014 added to stop this code looking up par vols when vols<0.0001

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
       ' If interpolateOn = "FORWARD" Then 'QJK code 28/01/2015 SHOULD BE AUTOMATICALLY ANY THING THAT USES THIS SHOULD BE IN FORWARD MODE AND  using function BootstraP_QJK
        'lngArr_FilteredDates = lngArr_FinalDates_Cap 'QJK code 28/01/2015
        'dblArr_FilteredVols = Convert_FilterArr(dblArr_FinalVols, intLst_IndexFilters)  'QJK code 28/01/2015
        'Else    'no condition, in case not only for surface  'QJK code 28/01/2015
        lngArr_FilteredDates = Convert_FilterArr(lngArr_FinalDates_ToUse, intLst_IndexFilters)
        dblArr_FilteredVols = Convert_FilterArr(dblArr_FinalVols, intLst_IndexFilters)
        'End If 'QJK code 28/01/2015
    End If

    Select Case str_Interp_Time
        Case "LIN", "LINEAR":
            dbl_Output = Interp_Lin(lngArr_FilteredDates, dblArr_FilteredVols, lng_DeductedMat, True) + dbl_VolShift_Sens
        Case "V2T"
            dbl_Output = Interp_V2t(lngArr_FilteredDates, dblArr_FilteredVols, rng_FinalBuildDate.Value, lng_DeductedMat) + dbl_VolShift_Sens
    End Select

   ' If dbl_Output <= 0.0001 Then 'Mandy 15/1/2015
       ' Select Case str_Interp_Time
       ' Case "LIN", "LINEAR":
       '     dbl_Output = Interp_Lin(rng_ATMDates, rng_ATMVolShocked, lng_DeductedMat, True) + dbl_VolShift_Sens
       ' Case "V2T"
      '      dbl_Output = Interp_V2t(rng_ATMDates, rng_ATMVolShocked, rng_FinalBuildDate.Value, lng_DeductedMat) + dbl_VolShift_Sens
       ' End Select
    'End If
    Lookup_VolCFSurfaceInterpolateOnFWD = dbl_Output
   ' Debug.Print "dbl_output is:" & dbl_Output & "and vol date is:" & CDate(lng_DeductedMat)
End Function
Public Function Lookup_VolSeries(int_FinalIndex As Integer, Optional intLst_InterpPillars As Collection = Nothing, _
    Optional bln_ApplyDeduction As Boolean = True) As Collection
    ' ## Obtain collection of caplet vols
    Dim dblLst_output As New Collection
    Dim int_CapletCtr As Integer

    For int_CapletCtr = 1 To int_FinalIndex
        Call dblLst_output.Add(Me.Lookup_Vol(rng_FinalDates_Caplet(int_CapletCtr, 1).Value, intLst_InterpPillars, bln_ApplyDeduction))
    Next int_CapletCtr

   ' If int_FinalIndex > 0 Then   '***no need for this anymore QJK comment out 31/01/2015. check with Mandy
        ' Use second vol as first vol since extrapolation is flat
    '    Call dblLst_Output.Add(dblLst_Output(1), , 1)
    'End If

    Set Lookup_VolSeries = dblLst_output
End Function

Public Function Lookup_VolSeriesCFSurfaceInterpolateOnFWD(int_FinalIndex As Integer, Optional intLst_InterpPillars As Collection = Nothing, _
    Optional bln_ApplyDeduction As Boolean = True) As Collection   'QJK code 02/02/2015
    ' ## Obtain collection of caplet vols
    Dim dblLst_output As New Collection
    Dim int_CapletCtr As Integer

    For int_CapletCtr = 1 To int_FinalIndex
       ' Call dblLst_Output.Add(Me.Lookup_Vol(rng_FinalDates_Caplet(int_CapletCtr, 1).Value, intLst_InterpPillars, bln_ApplyDeduction))
       Call dblLst_output.Add(Me.Lookup_VolCFSurfaceInterpolateOnFWD(rng_FinalDates_Caplet(int_CapletCtr, 1).Value, intLst_InterpPillars, bln_ApplyDeduction))   'QJK code 02/02/2015
    Next int_CapletCtr


    Set Lookup_VolSeriesCFSurfaceInterpolateOnFWD = dblLst_output
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
    ' QMOD changed SQL Query
    str_SQLCode = "SELECT [Data Date], Currency, Term, SortTerm, Rate " _
            & "FROM " & str_TableName _
            & " WHERE [Data Date] = #" & Convert_SQLDate(cfg_Settings.CurrentDataDate) & "#  AND Currency = '" & str_NameInDB & "' " & str_OptionalExclusions & " AND [Strike] ='" _
            & getStrikeQJK(str_CurveName) & "' AND [Name] = '" & getCurveNameSQL(str_CurveName) _
            & "' ORDER BY [SortTerm];"

    Me.ClearCurve

    Call Action_Query_Access(cfg_Settings.RatesDBPath, str_SQLCode, rng_QueryTopLeft)

    bln_ATMvols = False

    If getStrikeQJK(str_CurveName) = 0 Then
        bln_ATMvols = True
    End If

    Call AssignRanges

    ' Set instrument being shocked (cap or caplet) and set up associated ranges
    rng_ShockedCapVols.Value = rng_OrigCapVols.Value
    interpolateOn = rngInterpolateOn.Value

    ' Derive cap vol pillar dates
    Call GeneratePillarDates(True, True, True)
    Dim cal_pmt As Calendar: cal_pmt = cas_Calendars.Lookup_Calendar(fld_LegParams.PmtCal)

    If bln_ATMvols = False Then
        If interpolateOn = "FORWARD" Then   ' interpolate on forward vols
            Call Me.Bootstrap_QJK(True)
        Else
            Call Me.Bootstrap_ParVols(True) ' interpolate on par vols
        End If
    End If

End Sub

Public Sub SetParams(rng_QueryParams As Range)
    With wks_Location   'sets in CVL_MYR worksheet 'QJK comment 16/12/2014
        .Range("A2").Value = cfg_Settings.CurrentBuildDate
        .Range("B2:M2").Value = rng_QueryParams.Value
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


    'START FROM K onwards QJK comment 16/12/2014
    'int_ActiveCol = int_ActiveCol + 2  'SEEMS TO BE ON NEW WORKSHEET CVL_QJK_MYR
    'StrikeQJK = rng_FirstParam.Offset(0, 11).Value
End Sub

Public Sub ClearCurve()
    With wks_Location
        Call Action_ClearBelow(.Range("A7"), 6)
        Call Action_ClearBelow(.Range("H7"), 2)
        Call Action_ClearBelow(.Range("K7"), 3)
        Call Action_ClearBelow(.Range("O7"), 1)
        Call Action_ClearBelow(.Range("P7"), 2)
        Call Action_ClearBelow(.Range("S7"), 1)   'QJK code 14/1/2015  going to put Boolean=true/false here
        Call Action_ClearBelow(.Range("W7"), 6)   'QJK code 14/1/2015
        .Range("O2:Q2").ClearContents
    End With
End Sub

Private Sub AssignRanges(Optional bln_betweenStrikePillars As Boolean = False, Optional bln_IsFirstStrikePillar As Boolean = True)
    Dim int_NumCaps As Integer: int_NumCaps = Examine_NumRows(wks_Location.Range("A7"))
     wks_LocationATMVolsName = Left(wks_Location.Name, WorksheetFunction.Find("=", wks_Location.Name, 1)) & "0"  'QJK 14/01/2015
     Set wks_LocationATMVols = Worksheets(wks_LocationATMVolsName)  'QJK 14/01/2015


    If int_NumCaps > 0 Then
        Set rng_FirstFixing = wks_Location.Range("K2")

        Set rngStrikeQJK = wks_Location.Range("L2")   'Qcode 06/08/2014
       ' strikeQJK = wks_Location.Range("L2").Value                'Qcode 06/08/2014
        Set rngInterpolateOn = wks_Location.Range("M2")  'QJK CODE 16/12/2014

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
        Set rng_QueryTopLeftATM = wks_LocationATMVols.Range("A7") 'QJK code 14/1/2015
        Set rng_ATMDates = rng_QueryTopLeftATM.Offset(0, 5).Resize(int_NumCaps, 1) 'QJK code 14/1/2015
        Set rng_ATMVols = rng_QueryTopLeftATM.Offset(0, 4)  'QJK code 14/1/2015
        Set rng_ATMVolShocked = rng_QueryTopLeftATM.Offset(0, 15).Resize(int_NumCaps, 1) 'Mandy 15/1/2015

        Dim int_NumPoints As Integer
        If Me.IsBootstrappable = True Then
            Dim str_LastCapMat As String: str_LastCapMat = rng_CapTerms(int_NumCaps, 1).Value
            Dim int_NumCaplets As Integer:
            If rngInterpolateOn.Value = "FORWARD" Then 'QJK code 17/12/2014
            '' int_NumCaplets = Calc_NumPeriods(str_LastCapMat, fld_LegParams.PmtFreq) - 1
             int_NumCaplets = Calc_NumPeriods(str_LastCapMat, fld_LegParams.PmtFreq) 'QJK code 28/1/2015
           Else 'QJK code 17/12/2014  no minus 1 on numcaplets.
            int_NumCaplets = Calc_NumPeriods(str_LastCapMat, fld_LegParams.PmtFreq)  'QJK code 17/12/2014
            End If 'QJK code 17/12/2014
            int_NumPoints = int_NumCaplets

            Set rng_OrigDates_Caplet = rng_OrigDates_Cap(1, 1).Offset(0, 2).Resize(int_NumCaplets, 1)
            Set rng_OrigCapletVols = rng_OrigDates_Caplet.Offset(0, 1)
        Else
            int_NumPoints = int_NumCaps
        End If

        Set rng_FinalDates_Cap = wks_Location.Range("O7").Resize(int_NumCaps, 1)
        Set rng_ShockedCapVols = rng_FinalDates_Cap.Offset(0, 1)
        Set rng_FinalDates_Caplet = rng_ShockedCapVols(1, 1).Offset(0, 1).Resize(int_NumPoints, 1)

        If bln_ATMvols = True Then     'QJK code 28/01/2015, to get finalVolsRange right
        Set rng_FinalVols = rng_FinalDates_Caplet.Offset(0, 1).Resize(int_NumCaps, 1) 'QJK code 28/01/2015
        Else    'QJK code 28/01/2015
        Set rng_FinalVols = rng_FinalDates_Caplet.Offset(0, 1)
        End If   'QJK code 28/01/2015

        Set rng_CalibrationSolved = rng_FinalDates_Caplet.Offset(0, 2) 'QJK code 14/1/2015
        ' Fill interpolation cache
        lngArr_FinalDates_Cap = Convert_RangeToLngArr(rng_FinalDates_Cap)
        lngArr_FinalDates_Caplet = Convert_RangeToLngArr(rng_FinalDates_Caplet)
        dblArr_FinalVols = Convert_RangeToDblArr(rng_FinalVols)
       lngArr_FinalDates_ATMDates = Convert_RangeToLngArr(rng_ATMDates)  'QJK code 14/1/2015 take into account ATM vols for failure
        dblArr_FinalATMVols = Convert_RangeToDblArr(rng_ATMVols)    'QJK code 14/1/2015

        ' To get Vol of upper bound and lower bound strikes
        If bln_betweenStrikePillars = False Then
            dblArr_FinalVols = Convert_RangeToDblArr(rng_FinalVols)
        ElseIf bln_IsFirstStrikePillar = True Then
            dbl_StrikePillar_1 = rngStrikeQJK.Value
            dblArr_PillarVols_1 = Convert_RangeToDblArr(rng_FinalVols)
        ElseIf bln_IsFirstStrikePillar = False Then
            dbl_StrikePillar_2 = rngStrikeQJK.Value
            dblArr_PillarVols_2 = Convert_RangeToDblArr(rng_FinalVols)
        End If

   End If
End Sub

Private Sub interpolateVols(dbl_Strike As Double)
    'Interpolate Vol according to strike and then store dblArr_FinalVols

    Dim int_count As Integer
    ReDim dblArr_FinalVols(UBound(dblArr_PillarVols_1))

    For int_count = 1 To UBound(dblArr_PillarVols_1)
        dblArr_FinalVols(int_count) = dblArr_PillarVols_1(int_count) + (dblArr_PillarVols_2(int_count) - dblArr_PillarVols_1(int_count)) / (dbl_StrikePillar_2 - dbl_StrikePillar_1) _
            * (dbl_Strike - dbl_StrikePillar_1)
    Next int_count

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
    Dim lng_SpotDate As Long: lng_SpotDate = date_workday(lng_ValDate - 1, int_SpotDays + 1, cal_pmt.HolDates, cal_pmt.Weekends)
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

    If Me.IsBootstrappable = True And bln_ATMvols = False Then 'QJK code 28/1/2014 And bln_ATMvols = False Then
        If bln_GenCapletDates = True Then
            ' Derive caplet vol pillar dates
            For int_ctr = 1 To int_NumCapletVols   'QJK 29122014 mod to get same dates as Murex

            If int_ctr = 1 Then     'QJK 29122014 mod to get same dates as Murex
            lng_ActivePillarDate = lng_SpotDate    'QJK 29122014 mod to get same dates as Murex
                If int_Deduction = 0 Then
                    lngArr_FinalDates_Caplet(int_ctr) = date_workday(lng_ActivePillarDate - 1, 1, cal_Deduction.HolDates, cal_Deduction.Weekends) 'QJK 29122014 mod to get same dates as Murex
                Else
                    lngArr_FinalDates_Caplet(int_ctr) = date_workday(lng_ActivePillarDate, int_Deduction, cal_Deduction.HolDates, cal_Deduction.Weekends)
                End If
            Else    'QJK 29122014 mod to get same dates as Murex
            lng_ActivePillarDate = date_addterm(lng_SpotDate, fld_LegParams.PmtFreq, int_ctr - 1, True)
              '  lng_ActivePillarDate = Date_AddTerm(lng_SpotDate, fld_LegParams.PmtFreq, int_Ctr, True)
                lng_ActivePillarDate = Date_ApplyBDC(lng_ActivePillarDate, fld_LegParams.BDC, cal_pmt.HolDates, cal_pmt.Weekends)
                If int_Deduction = 0 Then
                    lngArr_FinalDates_Caplet(int_ctr) = date_workday(lng_ActivePillarDate - 1, 1, cal_Deduction.HolDates, cal_Deduction.Weekends)
                Else
                    lngArr_FinalDates_Caplet(int_ctr) = date_workday(lng_ActivePillarDate, int_Deduction, cal_Deduction.HolDates, cal_Deduction.Weekends)
                End If
            End If    'QJK 29122014 mod to get same dates as Murex
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
        'QJKmod
        strikeQJK = rngStrikeQJK.Value
        dbl_ActiveATMStrike = strikeQJK   'QCode 06/08/2014
        'dbl_ActiveATMStrike = irs_Active.ParRate_LegA
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
        interpolateOn = rngInterpolateOn.Value   'QJK code 13012015
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

                If getStrikeQJK(str_CurveName) = 0 Then bln_Bootstrap = False 'QJK code 27/1 2014 added to stop bootstrapping at K=0


                If bln_Bootstrap = True Then
                   If interpolateOn = "PAR" Then 'QJK code 13012015
                    Call Me.Bootstrap_ParVols(False)  'QJK code 13012015
                   Else  'QJK code 13012015
                   Call Me.Bootstrap_QJK(False)  'for forward vols 'QJK code 13012015
                    'Call Me.Bootstrap(False)
                  End If  'QJK code 13012015
                ElseIf getStrikeQJK(str_CurveName) = 0 Then  'QJK code 27/1 2014 added to stop bootstrapping at K=0
                Else    ' must be the case when no need to bootstrap, and not ATM vols
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


'QMOD added code for curvename SQL filter
'gets curvename for RER and loks it up in Microsoft Access
Public Function getCurveNameSQL(curvenameRER As String) As String
    Dim temp As String

    Select Case Left(curvenameRER, 6)
        Case "MYR_6M"
            temp = "MYR 6M"
        Case "USD_6M"
            temp = "USD 6M"
        Case "MYR_3M"
            temp = "MYR 3M"
        Case "USD_3M"
            temp = "USD 3M"
    End Select

    getCurveNameSQL = temp
End Function

Public Function getStrikeQJK(curvenameRER As String) As Double
    Dim temp As Double
    Dim tempLength As Double, tempStart As Double
    tempLength = Len(curvenameRER)
    tempStart = WorksheetFunction.Find("=", curvenameRER) + 1

    temp = Mid(curvenameRER, tempStart, tempLength - tempStart + 1)

    getStrikeQJK = temp
End Function

Public Sub Bootstrap_QJK(bln_CopyToOrig As Boolean)
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
    ReDim bln_CalibrationSolved(1 To int_NumCaplets) As Boolean   ' number of caplets used for true or false

    ' Secant method variables
    Dim dic_SecantParams As Dictionary, dic_SecantOutputs As Dictionary
    Dim bln_SolutionPossible As Boolean
    Dim int_PrevFinalIndex As Integer: int_PrevFinalIndex = 0
    Dim sumCapPremiums() As Double: ReDim sumCapPremiums(1 To int_NumCaps) As Double
    Dim end_FirstCap As Double: Dim j As Double
    Dim temp() As Double: ReDim temp(1 To int_NumCaplets) As Double:

    '-------------------------------------------------------------------
    ' Calculate cap premium for each pillar cap using par vol
    '-------------------------------------------------------------------
    For int_CapCtr = 1 To int_NumCaps
        fld_Params.LegA.Term = rng_CapTerms(int_CapCtr, 1).Value
        fld_Params.LegB.Term = fld_Params.LegA.Term
        Set irs_Active = GetInst_IRS(fld_Params, dic_CurveSet, dic_GlobalStaticInfo)
        Set irl_Floating = irs_Active.LegB

        strikeQJK = rngStrikeQJK.Value
        dbl_ActiveATMStrike = strikeQJK
        If dbl_ActiveATMStrike < dbl_MinStrike Then dbl_ActiveATMStrike = dbl_MinStrike
        dbl_ActiveCapVol = rng_ShockedCapVols(int_CapCtr, 1).Value
        If dbl_ActiveCapVol = 0 Then dbl_ActiveCapVol = 0.000001

        ' Doesnt take first caplet into account.
        dbl_ActiveCapPrice = irl_Floating.Calc_BSOptionValue(enu_Direction, dbl_ActiveATMStrike, int_Deduction, cal_Deduction, True, ,dbl_ActiveCapVol)

        Debug.Print " cap price  " & fld_Params.LegA.Term & dbl_ActiveCapPrice
        sumCapPremiums(int_CapCtr) = dbl_ActiveCapPrice   'i.e. (1)=1,(2)=1+2,(3)=1+2+3, etc
    Next int_CapCtr

    For int_CapCtr = 1 To int_NumCaps
        ' Find index of latest caplet falling within the cap period
        fld_Params.LegA.Term = rng_CapTerms(int_CapCtr, 1).Value
        fld_Params.LegB.Term = fld_Params.LegA.Term

        Set irs_Active = GetInst_IRS(fld_Params, dic_CurveSet, dic_GlobalStaticInfo)
        Set irl_Floating = irs_Active.LegB

        int_ActiveFinalIndex = Calc_NumPeriods(irl_Floating.Params.Term, fld_LegParams.PmtFreq)
        Call intLst_InterpPillars.Add(int_ActiveFinalIndex)
        dbl_ActiveCapVol = rng_ShockedCapVols(int_CapCtr, 1).Value

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

        ' Consistency Check
        Dim col_fwdrate As Collection
        Set col_fwdrate = New Collection
        Dim k As Variant

        For Each k In irl_Floating.ForwRates
            col_fwdrate.Add (k)
        Next k

        '-------------------------------------------------------------------
        ' Each pillar cap goes through consistency check
        ' Discard cap premium for those pillars that fail consistency check
        '-------------------------------------------------------------------
        Dim int_1stccfailpillar As Integer
        Dim int_CCFail As Integer
        Dim CCResult As String
        Dim PrevCapPremium As Double
        If int_CapCtr <> 1 Then
            PrevCapPremium = sumCapPremiums(int_CapCtr - int_CCFail - 1)
            CCResult = Me.ConsistencyCheck(enu_Direction, irl_Floating, col_fwdrate, int_CapCtr, int_CCFail, dbl_ActiveATMStrike, PrevCapPremium, sumCapPremiums(int_CapCtr), intLst_InterpPillars)
        End If

        If int_CapCtr = 1 Then
            ' Just find last caplet that falls within the first cap period
            end_FirstCap = int_ActiveFinalIndex
            For j = 1 To end_FirstCap
                dblArr_FinalVols(j) = dbl_ActiveCapVol
                bln_CalibrationSolved(j) = True
            Next j

            GoTo Label2:
        End If

        '-------------------------------------------------------------------
        ' Pass consistency check, use secant method to solve pillar cap vol
        '-------------------------------------------------------------------
        If CCResult = True Then
            ' Solve using secant method
            Set dic_SecantOutputs = New Dictionary

            dbl_ActiveCapletVol = Solve_SecantQJK(ThisWorkbook, "SolverFuncXY_CapletVolToPriceCFSurfaceInterpolateonFWD", dic_SecantParams, _
                dbl_ActiveCapVol, dbl_ActiveCapVol + 1, sumCapPremiums(int_CapCtr), 0.0000000001, 60, -1, dic_SecantOutputs)

            ' Final solution will be shown in the cell, if no solution found, show error value
            If dic_SecantOutputs("Solvable") = True And dbl_ActiveCapletVol >= 0 Then
                dbl_PrevCapletVol = dbl_ActiveCapletVol
                Debug.Print dbl_ActiveCapletVol & " was successful  " & int_ActiveFinalIndex

                For j = int_ActiveFinalIndex To (int_PrevFinalIndex + 1) Step -1
                    bln_CalibrationSolved(j) = True
                Next j

                '-----------------------------------------------------------------------
                ' Linear interpolate caplet vol between calibrated successful pillar cap
                '-----------------------------------------------------------------------
                If int_CCFail = 0 Then
                    For int_CapletCtr = int_PrevFinalIndex To int_ActiveFinalIndex
                        If bln_CalibrationSolved(int_CapletCtr) = True Then
                            dblArr_FinalVols(int_CapletCtr) = Me.Lookup_VolCFSurfaceInterpolateOnFWD(lngArr_FinalDates_Caplet(int_CapletCtr), intLst_InterpPillars, False)
                        End If
                    Next int_CapletCtr
                Else
                    For int_CapletCtr = int_PreviousCapletsNum To int_ActiveFinalIndex
                        If bln_CalibrationSolved(int_CapletCtr) = True Then
                            dblArr_FinalVols(int_CapletCtr) = Me.Lookup_VolCFSurfaceInterpolateOnFWD(lngArr_FinalDates_Caplet(int_CapletCtr), intLst_InterpPillars, False)
                        End If
                    Next int_CapletCtr
                End If

            '---------------------
            ' Calibration failure
            '---------------------
            Else
                Call intLst_FailedPoints.Add(int_ActiveFinalIndex)
                For j = int_ActiveFinalIndex To (int_PrevFinalIndex + 1) Step -1
                    bln_CalibrationSolved(j) = False
                Next j

                Debug.Print "## ERROR - Caplet volatility could not be solved for " & str_CurveName & " " & irl_Floating.Params.Term

                bln_calibrationfailed = True

                '  dblArr_FinalVols(int_Index)--final vols filled
                Select Case str_OnFail
                    Case "FLAT": Call Me.SetFinalVol(int_ActiveFinalIndex, dbl_PrevCapletVol)
                    Case "ZERO", "PAR":
                        Call Me.SetFinalVol(int_ActiveFinalIndex, dbl_MinVol)

                        If bln_CalibrationSolved(int_PrevFinalIndex) = True Then
                            'last caplet was solved so linearly interpolate between last caplet and min vol
                            For j = int_ActiveFinalIndex To (int_PrevFinalIndex + 1) Step -1
                                dblArr_FinalVols(j) = Me.Lookup_VolCFSurfaceInterpolateOnFWD(lngArr_FinalDates_Caplet(j), intLst_InterpPillars, False)
                            Next j
                        ElseIf bln_CalibrationSolved(int_PrevFinalIndex) = False Then
                            'last caplet failed, keep points in between at minvol
                            For j = int_ActiveFinalIndex To (int_PrevFinalIndex + 1) Step -1
                                dblArr_FinalVols(j) = dbl_MinVol
                            Next j
                        End If
                End Select
            End If
Label2:
            int_PrevFinalIndex = int_ActiveFinalIndex
            int_CCFail = 0

        '------------------------
        ' Fail consistency check
        '------------------------
        Else
            '----------------------------------------------------------
            ' Skip to the next pillar if this is not the final pillar
            '----------------------------------------------------------
            If int_CapCtr <> int_NumCaps Then
                int_CCFail = int_CCFail + 1

                If int_CCFail = 1 Then
                    int_1stccfailpillar = intLst_InterpPillars(intLst_InterpPillars.count)
                End If

                Dim int_skipcap As Integer
                int_skipcap = intLst_InterpPillars.count
                Call intLst_InterpPillars.Remove(int_skipcap)

                Debug.Print "consistency check failed - " & " " & str_CurveName & " " & irl_Floating.Params.Term

            '-----------------------------------
            ' Last pillar fail consitency check
            '-----------------------------------
            Else
                Select Case str_OnFail
                    Case "FLAT": Call Me.SetFinalVol(int_ActiveFinalIndex, dbl_PrevCapletVol)
                    Case "ZERO", "PAR":
                        If bln_CalibrationSolved(int_PrevFinalIndex) = True Then
                            'last caplet was solved so linearly interpolate between last caplet and min vol for first failed pillar
                            'the remaining caplets are using minvol
                            If int_CCFail <> 0 And int_CapCtr = int_NumCaps Then
                                Call Me.SetFinalVol(int_1stccfailpillar, dbl_MinVol)
                                Dim int_lastcap As Integer
                                int_lastcap = intLst_InterpPillars.count
                                Call intLst_InterpPillars.Remove(int_lastcap)
                                Call intLst_InterpPillars.Add(int_1stccfailpillar)

                                For j = int_ActiveFinalIndex To (int_PrevFinalIndex + 1) Step -1
                                    dblArr_FinalVols(j) = dblArr_FinalVols(int_PrevFinalIndex)
                                Next j
                                For j = int_ActiveFinalIndex To (int_1stccfailpillar + 1) Step -1
                                    dblArr_FinalVols(j) = dblArr_FinalVols(int_PrevFinalIndex)
                                Next j
                            Else
                                Call Me.SetFinalVol(int_ActiveFinalIndex, dbl_PrevCapletVol)
                                For j = int_ActiveFinalIndex To (int_PrevFinalIndex + 1) Step -1
                                    dblArr_FinalVols(j) = dblArr_FinalVols(int_PrevFinalIndex)
                                Next j
                            End If
                        ElseIf bln_CalibrationSolved(int_PrevFinalIndex) = False Then
                            'last caplet failed, keep points in between at minvol
                            For j = int_ActiveFinalIndex To (int_PrevFinalIndex + 1) Step -1
                                dblArr_FinalVols(j) = dblArr_FinalVols(int_PrevFinalIndex)
                            Next j
                        End If
                    Debug.Print "consistency check failed - " & " " & str_CurveName & " " & irl_Floating.Params.Term
                End Select
            End If
        End If
    Next int_CapCtr

    ' Fill in interpolated caplet pillars
    For int_CapletCtr = end_FirstCap To int_NumCaplets
        If bln_CalibrationSolved(int_CapletCtr) = True Then
           dblArr_FinalVols(int_CapletCtr) = Me.Lookup_VolCFSurfaceInterpolateOnFWD(lngArr_FinalDates_Caplet(int_CapletCtr), intLst_InterpPillars, False)
        End If
    Next int_CapletCtr

    ' Output to sheet
    rng_FinalVols.Value = Convert_Array1Dto2D(dblArr_FinalVols)
    If bln_CopyToOrig = True Then rng_OrigCapletVols.Value = rng_FinalVols.Value
    rng_CalibrationSolved = Convert_Array1Dto2D(bln_CalibrationSolved)
    Application.StatusBar = False
    Application.ScreenUpdating = bln_ScreenUpdating

End Sub

Public Sub Bootstrap_ParVols(bln_CopyToOrig As Boolean)
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
        Dim sumCapPremiums() As Double: ReDim sumCapPremiums(1 To int_NumCaps) As Double   'QCode 11/08/2014
        Dim end_FirstCap As Double: Dim j As Double   'QCode 11/08/2014
        Dim temp() As Double: ReDim temp(1 To int_NumCaplets) As Double:   'QCode 11/08/2014
        strikeQJK = rngStrikeQJK.Value
        dbl_ActiveATMStrike = strikeQJK   'QCode 06/08/2014
  'For int_CapletCtr = 1 To int_NumCaplets
   '     fld_Params.LegA.Term = rng_CapTerms(int_CapCtr, 1).Value
   '     fld_Params.LegB.Term = fld_Params.LegA.Term
    '    Set irs_Active = GetInst_IRS(fld_Params, dic_CurveSet, dic_GlobalStaticInfo)
    '    Set irl_Floating = irs_Active.LegB
    '    'QJKmod

        'dbl_ActiveATMStrike = irs_Active.ParRate_LegA
      '  If dbl_ActiveATMStrike < dbl_MinStrike Then dbl_ActiveATMStrike = dbl_MinStrike
       ' dbl_ActiveCapVol = rng_ShockedCapVols(int_CapCtr, 1).Value
       ' dbl_ActiveCapPrice = irl_Floating.Calc_BSOptionValue(enu_Direction, dbl_ActiveATMStrike, int_Deduction, cal_Deduction, True, , _
            dbl_ActiveCapVol)

        'sumCapPremiums(int_CapCtr) = dbl_ActiveCapPrice   'QCode 11/08/2014--i.e. (1)=1,(2)=1+2,(3)=1+2+3, etc
        'calculates cap premium
        'Public Function Calc_BSOptionValue(enu_Direction As OptionDirection, dbl_Strike As Double, int_Deduction As Integer, _
    cal_Deduction As Calendar, bln_IsDiscounted As Boolean, Optional dblLst_CapletVols As Collection = Nothing, _
    Optional dbl_CapVol As Double = -1, Optional str_ValueType As String = "PNL", Optional int_CapletIndex As Integer = -1) As Double

   'Next int_CapletCtr                       'Qcode 11/08/2014


     ParVolDates = getParVolDates(fld_LegA.PmtFreq, lng_SpotDate, int_NumCaplets)
    ' ParVolDates = (lngArr_FinalDates_Caplet)
    ReDim bln_CalibrationSolved(1 To int_NumCaplets) As Boolean   'QJK code 14/1/2015
    For int_CapletCtr = 1 To (int_NumCaplets)   'Qcode 11/08/2014
            ' Find index of latest caplet falling within the cap period
        'fld_Params.LegA.Term = rng_CapTerms(int_NumCaps, 1).Value
        fld_Params.LegA.Term = getCapletTerms(int_CapletCtr, fld_Params.LegA.PmtFreq)
        fld_Params.LegB.Term = fld_Params.LegA.Term
        Set irs_Active = GetInst_IRS(fld_Params, dic_CurveSet, dic_GlobalStaticInfo)
        Set irl_Floating = irs_Active.LegB
       ' int_ActiveFinalIndex = Calc_NumPeriods(irl_Floating.Params.Term, fld_LegParams.PmtFreq) - 1
        int_ActiveFinalIndex = int_CapletCtr
        Call intLst_InterpPillars.Add(int_ActiveFinalIndex)
        'dbl_ActiveCapVol = rng_ShockedCapVols(int_CapCtr, 1).Value    ''Qcode 11/08/2014
        'dbl_ActiveCapVol = Interp_Spline(rng_FinalDates_Cap, rng_ShockedCapVols, rng_FinalDates_Caplet(int_CapletCtr, 1))
         'dbl_ActiveCapVol = TMR_KSpline(rng_FinalDates_Cap, rng_ShockedCapVols, rng_FinalDates_Caplet(int_CapletCtr, 1), 0, 0, False, False)
dbl_ActiveCapVol = TMR_KSpline(rng_FinalDates_Cap, rng_ShockedCapVols, ParVolDates(int_CapletCtr), 0, 0, False, False)
'dbl_ActiveCapVol = TMR_KSpline(rng_FinalDates_Cap, rng_ShockedCapVols, lngArr_FinalDates_Caplet(int_CapletCtr), 0, 0, False, False)
Dim dbl_fallBackVol As Double: 'QJK code 14/01/2014
    If int_CapletCtr >= 19 Then
    Dim g As Double: g = 1
    End If
        'Interp_Spline(arr_X As Variant, arr_Y As Variant, var_LookupX As Variant) As Double
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
        ' dbl_ActiveCapVol = rng_ShockedCapVols(int_CapCtr, 1).Value
         If int_CapletCtr <= 2 Then
        ' dbl_ActiveCapVol = TMR_KSpline(rng_FinalDates_Cap, rng_ShockedCapVols, rng_FinalDates_Caplet(1, 1), 0, 0, False, False)
         dbl_ActiveCapVol = TMR_KSpline(rng_FinalDates_Cap, rng_ShockedCapVols, ParVolDates(int_CapletCtr), 0, 0, False, False)
          'QJK 05/10/2016
         If dbl_ActiveCapVol <= 0 Then
          dblArr_FinalVols(int_CapletCtr) = 0.000001
            bln_CalibrationSolved(int_CapletCtr) = False
          Else
         dblArr_FinalVols(int_CapletCtr) = dbl_ActiveCapVol
            bln_CalibrationSolved(int_CapletCtr) = True
         End If





         GoTo Label2:
         End If
        dbl_ActiveCapPrice = irl_Floating.Calc_BSOptionValue(enu_Direction, dbl_ActiveATMStrike, int_Deduction, cal_Deduction, True, , _
            dbl_ActiveCapVol)


       ' dbl_ActiveCapPrice = irl_Floating.Calc_BSOptionValue(enu_Direction, dbl_ActiveATMStrike, int_Deduction, cal_Deduction, True, dbl_ActiveCapVol)
            '(enu_Direction As OptionDirection, dbl_Strike As Double, int_Deduction As Integer, _
    cal_Deduction As Calendar, bln_IsDiscounted As Boolean, Optional dblLst_CapletVols As Collection = Nothing, _
    Optional dbl_CapVol As Double = -1, Optional str_ValueType As String = "PNL", Optional int_CapletIndex As Integer = -1)
        'QCode 11/08/2014, just find last caplet that falls within the first cap period"

      '  For j = 2 To int_NumCaplets       'Qcode 11/08/2014

        ' Solve using secant method
        Set dic_SecantOutputs = New Dictionary
            'If int_CapCtr <= end_FirstCap Then     'Qcode 08/08/2014
            'temp(j) = dbl_ActiveCapVol    'Qcode 08/08/2014: caplets in the first cap pillar are FLAT.
           ' Else 'Qcode 08/08/2014
           ' temp(j)=

          'dbl_ActiveCapletVol = Solve_Secant(ThisWorkbook, "SolverFuncXY_CapletVolToPriceQJK", dic_SecantParams, _
                dbl_ActiveCapVol, dbl_ActiveCapVol + 1, sumCapPremiums(int_CapCtr), 0.0000000001, 50, -1, dic_SecantOutputs)
                  ' Select Case str_OnFail  'QJK code 14/01/2014
                   ' Case "FLAT": dbl_fallBackVol = dbl_PrevCapletVol 'QJK code 14/01/2014
                   ' Case "ZERO", "PAR": dbl_fallBackVol = dbl_MinVol   'QJK code 14/01/2014
                  'End Select  'QJK code 14/01/2014
           ' dbl_ActiveCapletVol = Solve_SecantQJK(ThisWorkbook, "SolverFuncXY_CapletVolToPriceQJK", dic_SecantParams, _
                dbl_ActiveCapVol, dbl_ActiveCapVol + 1, dbl_ActiveCapPrice, 0.0000000001, 50, -1, dic_SecantOutputs)

                '0.000001
                dbl_ActiveCapletVol = Solve_SecantQJK(ThisWorkbook, "SolverFuncXY_CapletVolToPriceQJK", dic_SecantParams, _
                dbl_ActiveCapVol, dbl_ActiveCapVol + 1, dbl_ActiveCapPrice, 0.0000000001, 60, dbl_MinVol, dic_SecantOutputs)

            'End If                     'Qcode 08/08/2014


        ' Final solution will be shown in the cell, if no solution found, show error value
        If dic_SecantOutputs("Solvable") = True And dbl_ActiveCapletVol > 0 Then
            dbl_PrevCapletVol = dbl_ActiveCapletVol
             bln_CalibrationSolved(int_CapletCtr) = True  'QJK code 14/1/2015
       ' Debug.Print "YAY!!!!Caplet volatility SOLVED " & str_CurveName & " " _
                & irl_Floating.Params.Term & "  " & dbl_ActiveCapletVol & "CAPLET NUMBER:" & int_CapletCtr
        Else
            bln_CalibrationSolved(int_CapletCtr) = False

            Call intLst_FailedPoints.Add(int_ActiveFinalIndex)
                    Select Case str_OnFail
                    Case "FLAT": dblArr_FinalVols(int_CapletCtr) = dbl_PrevCapletVol ' Call Me.SetFinalVol(int_ActiveFinalIndex, dbl_PrevCapletVol)
                    Case "ZERO", "PAR": dblArr_FinalVols(int_CapletCtr) = dbl_MinVol
                    'Case "PAR":  dblArr_FinalVols(int_CapletCtr) = Interp_Lin(lngArr_FinalDates_ATMDates, dblArr_FinalATMVols, lngArr_FinalDates_Caplet(int_CapletCtr), True) 'QJK code 14/1/2015
                  'Private lngArr_FinalDates_ATMDates() As Long: Private dblArr_FinalATMVols() As Double 'QJK code 14/1/2015
                  ' dblArr_FinalVols(var_ActiveIndex) = Interp_Lin(lngArr_FinalDates_Cap, dblArr_CapVols, lng_ActiveCapletMat, True)
                   End Select
            Debug.Print "## ERROR - Caplet volatility could not be solved for " & str_CurveName & " " _
                & irl_Floating.Params.Term & dbl_ActiveCapVol & "CAPLET NUMBER:" & int_CapletCtr
            ' Fall back to the previous pillar vol
           ' If int_PrevFinalIndex <> 0 Then

           ' End If
        End If

       ' int_PrevFinalIndex = int_ActiveFinalIndex

Label2:
  '  Next j   'Qcode 11/08/2014
    Next int_CapletCtr

    ' Fill in interpolated caplet pillars
   ' For int_CapletCtr = 1 To int_NumCaplets
   '     dblArr_FinalVols(int_CapletCtr) = Me.Lookup_Vol(lngArr_FinalDates_Caplet(int_CapletCtr), intLst_InterpPillars, False)
   ' Next int_CapletCtr


    If str_OnFail = "PAR" Then
        ' After interpolation, replace failed pillars with interpolated par vols
        'Dim dblArr_CapVols() As Double: dblArr_CapVols = Convert_RangeToDblArr(rng_ShockedCapVols)
        'Dim var_ActiveIndex As Variant, lng_ActiveCapletMat As Long

        For int_CapletCtr = 1 To (int_NumCaplets) 'QJK code 14/1/2015
            If bln_CalibrationSolved(int_CapletCtr) = False Then 'QJK code 14/1/2015
            'dblArr_FinalVols(int_CapletCtr) = Interp_Lin(lngArr_FinalDates_ATMDates, dblArr_FinalATMVols, lngArr_FinalDates_Caplet(int_CapletCtr), True) 'QJK code 14/1/2015
            dblArr_FinalVols(int_CapletCtr) = Interp_Lin(rng_ATMDates, rng_ATMVols, lngArr_FinalDates_Caplet(int_CapletCtr), True) 'QJK code 14/1/2015
            End If 'QJK code 14/1/2015
        Next int_CapletCtr 'QJK code 14/1/2015
        '      Set rng_ATMDates = rng_QueryTopLeftATM.Offset(0, 5).Resize(int_NumCaps, 1) 'QJK code 14/1/2015
        'Set rng_ATMVols = rng_QueryTopLeftATM.Offset(0, 4)  'QJK code 14/1/2015
        ' Check zero vols should really be zero by re-interpolating
        'For int_CapletCtr = 1 To int_NumCaplets
        '    If Round(dblArr_FinalVols(int_CapletCtr), 8) = dbl_MinVol Then
         '       dblArr_FinalVols(int_CapletCtr) = Me.Lookup_Vol(lngArr_FinalDates_Caplet(int_CapletCtr), intLst_InterpPillars, False)
         '   End If
        'Next int_CapletCtr
    End If

    ' Output to sheet
    rng_FinalVols.Value = Convert_Array1Dto2D(dblArr_FinalVols)
    If bln_CopyToOrig = True Then rng_OrigCapletVols.Value = rng_FinalVols.Value
    rng_CalibrationSolved = Convert_Array1Dto2D(bln_CalibrationSolved)   'QJK code 14/1/2015
    Application.StatusBar = False
    Application.ScreenUpdating = bln_ScreenUpdating
End Sub

Public Function getCapletTerms(intCaplet As Integer, freq As String) As String
    Dim temp As String
    Dim yearsInt As Integer, monthsInt As Integer
    Dim freqInt As Integer: freqInt = calc_nummonths(freq)

    monthsInt = freqInt * (intCaplet)

    If monthsInt >= 12 Then
    yearsInt = Int(monthsInt / 12)
    monthsInt = monthsInt - (yearsInt * 12)
        If monthsInt = 0 Then
        temp = yearsInt & "Y"
        Else
        temp = yearsInt & "Y," & monthsInt & "M"
        End If
    Else
    temp = monthsInt & "M"
    End If

    getCapletTerms = temp

End Function


Public Function Lookup_VolSeriesParVols(int_FinalIndex As Integer, Optional intLst_InterpPillars As Collection = Nothing, _
    Optional bln_ApplyDeduction As Boolean = True) As Collection

'Public Function Lookup_VolSeries(int_FinalIndex As Integer, Optional intLst_InterpPillars As Collection = Nothing, _
    Optional bln_ApplyDeduction As Boolean = True) As Collection
    ' ## Obtain collection of caplet vols
   Dim dblLst_output As New Collection
    Dim int_CapletCtr As Integer

  '
   For int_CapletCtr = 1 To int_FinalIndex
        Call dblLst_output.Add(dblArr_FinalVols(int_CapletCtr))
    Next int_CapletCtr
  '
  '  If int_FinalIndex > 0 Then
  '      ' Use second vol as first vol since extrapolation is flat
  '      Call dblLst_Output.Add(dblLst_Output(1), , 1)
  '  End If

    Set Lookup_VolSeriesParVols = dblLst_output
'End Function
End Function

Public Function TMR_KSpline(X_Array As Variant, Y_Array As Variant, x As Double, _
                    Optional yt_1 As Double = 0, Optional q_n As Double = 0, _
                    Optional FlatEnds As Boolean = False, _
                    Optional LinearEnds As Boolean = True) As Double
    'normally YT_1=-0.5 AND q_n=0.5
    'Edited SRS Cubic spline - adapted from Numerical Recipes in C
    Dim iCnt As Integer
    '' Kspline.TMR_KSplineInternal(capMatDates, capVols, parVolDates(i), 0, 0, False, False)
    iCnt = WorksheetFunction.CountA(X_Array)

    '''''''''''''''''''''''''''''''''''''''
    ' values are populated
    '''''''''''''''''''''''''''''''''''''''
    Dim n As Integer 'n=iCnt
    Dim i As Integer, j As Integer, k As Integer 'these are loop counting integers
    Dim p, qn, sig, un As Double
    ReDim U(iCnt - 1) As Double
    ReDim yt(iCnt) As Double 'these are the 2nd deriv values

    n = iCnt
    yt(1) = yt_1
    U(1) = 0

    For i = 2 To n - 1
        sig = (X_Array(i) - X_Array(i - 1)) / (X_Array(i + 1) - X_Array(i - 1))
        p = sig * yt(i - 1) + 2
        yt(i) = (sig - 1) / p
        U(i) = (Y_Array(i + 1) - Y_Array(i)) / (X_Array(i + 1) - X_Array(i)) - (Y_Array(i) - Y_Array(i - 1)) / (X_Array(i) - X_Array(i - 1))
        U(i) = (6 * U(i) / (X_Array(i + 1) - X_Array(i - 1)) - sig * U(i - 1)) / p
    Next i

    qn = q_n
    un = 0

    yt(n) = (un - qn * U(n - 1)) / (qn * yt(n - 1) + 1)

    For k = n - 1 To 1 Step -1
        yt(k) = yt(k) * yt(k + 1) + U(k)
    Next k

    ''''''''''''''''''''
    'now eval spline at one point
    '''''''''''''''''''''
    Dim klo As Integer, khi As Integer, h As Double, B As Double, A As Double, outCnt As Integer
    outCnt = WorksheetFunction.CountA(x)
    ' first find correct interval
    ReDim y(1 To outCnt, 1 To 1)
    For i = 1 To outCnt
        klo = 1: khi = 2
        If FlatEnds And x <= X_Array(1) Then
            y(i, 1) = Y_Array(1)
        ElseIf FlatEnds And x >= X_Array(n) Then
            y(i, 1) = Y_Array(n)
        ElseIf LinearEnds And x <= X_Array(1) Then
            y(i, 1) = Y_Array(1) + (Y_Array(2) - Y_Array(1)) / (X_Array(2) - X_Array(1)) * (x - X_Array(1))
        ElseIf LinearEnds And x >= X_Array(n) Then
            y(i, 1) = Y_Array(n) + (Y_Array(n) - Y_Array(n - 1)) / (X_Array(n) - X_Array(n - 1)) * (x - X_Array(n))
        Else
            For j = 1 To n - 2
                If x < X_Array(khi) Then Exit For
                klo = klo + 1
                khi = khi + 1
            Next j

            h = X_Array(khi) - X_Array(klo)
            A = (X_Array(khi) - x) / h
            B = (x - X_Array(klo)) / h
            y(i, 1) = A * Y_Array(klo) + B * Y_Array(khi) + ((A ^ 3 - A) * yt(klo) + (B ^ 3 - B) * yt(khi)) * (h ^ 2) / 6
        End If
    Next i

    TMR_KSpline = y(1, 1)

End Function

Public Function TMR_KSplineInternal(X_Array() As Double, Y_Array() As Double, x As Double, _
                    Optional yt_1 As Double = 0, Optional q_n As Double = 0, _
                    Optional FlatEnds As Boolean = False, _
                    Optional LinearEnds As Boolean = True) As Double
    'normally YT_1=-0.5 AND q_n=0.5
    'Edited SRS Cubic spline - adapted from Numerical Recipes in C
    Dim iCnt As Integer

    iCnt = WorksheetFunction.CountA(X_Array)

    '''''''''''''''''''''''''''''''''''''''
    ' values are populated
    '''''''''''''''''''''''''''''''''''''''
    Dim n As Integer 'n=iCnt
    Dim i As Integer, k As Integer  'these are loop counting integers
    Dim p, qn, sig, un As Double
    ReDim U(iCnt - 1) As Double
    ReDim yt(iCnt) As Double 'these are the 2nd deriv values

    n = iCnt
    yt(1) = yt_1
    U(1) = 0

    For i = 2 To n - 1
        sig = (X_Array(i) - X_Array(i - 1)) / (X_Array(i + 1) - X_Array(i - 1))
        p = sig * yt(i - 1) + 2
        yt(i) = (sig - 1) / p
        U(i) = (Y_Array(i + 1) - Y_Array(i)) / (X_Array(i + 1) - X_Array(i)) - (Y_Array(i) - Y_Array(i - 1)) / (X_Array(i) - X_Array(i - 1))
        U(i) = (6 * U(i) / (X_Array(i + 1) - X_Array(i - 1)) - sig * U(i - 1)) / p
    Next i

    qn = q_n
    un = 0

    yt(n) = (un - qn * U(n - 1)) / (qn * yt(n - 1) + 1)

    For k = n - 1 To 1 Step -1
        yt(k) = yt(k) * yt(k + 1) + U(k)
    Next k

    ''''''''''''''''''''
    'now eval spline at one point
    '''''''''''''''''''''
    Dim klo As Integer, khi As Integer, h As Double, B As Double, A As Double, outCnt As Integer
    outCnt = WorksheetFunction.CountA(x)
    ' first find correct interval
    ReDim y(1 To outCnt, 1 To 1)
    For i = 1 To outCnt
        klo = 1: khi = 2
        If FlatEnds And x <= X_Array(1) Then
            y(i, 1) = Y_Array(1)
        ElseIf FlatEnds And x >= X_Array(n) Then
            y(i, 1) = Y_Array(n)
        ElseIf LinearEnds And x <= X_Array(1) Then
            y(i, 1) = Y_Array(1) + (Y_Array(2) - Y_Array(1)) / (X_Array(2) - X_Array(1)) * (x - X_Array(1))
        ElseIf LinearEnds And x >= X_Array(n) Then
            y(i, 1) = Y_Array(n) + (Y_Array(n) - Y_Array(n - 1)) / (X_Array(n) - X_Array(n - 1)) * (x - X_Array(n))
        Else
            For j = 1 To n - 2
                If x < X_Array(khi) Then Exit For
                klo = klo + 1
                khi = khi + 1
            Next j

            h = X_Array(khi) - X_Array(klo)
            A = (X_Array(khi) - x) / h
            B = (x - X_Array(klo)) / h
            y(i, 1) = A * Y_Array(klo) + B * Y_Array(khi) + ((A ^ 3 - A) * yt(klo) + (B ^ 3 - B) * yt(khi)) * (h ^ 2) / 6
        End If
    Next i

    TMR_KSplineInternal = y(1, 1)

End Function

Public Function getParVolDates(freq As String, couponstartdate As Long, nRows As Integer) As Double()
    Dim temp() As Double, freqInt As Integer: 'nrows=number of caplets
    ReDim temp(1 To nRows) As Double
    Dim i As Double, tempInt As Integer: tempInt = calc_nummonths(freq)

    For i = 1 To nRows
    'temp(i) = WorksheetFunction.EDate(couponstartdate, (i) * tempInt)
    temp(i) = date_addterm(couponstartdate, (i) * tempInt & "M", 1, True)
    Next i

    getParVolDates = temp
End Function

Public Function Solve_SecantQJK(wbk_Caller As Workbook, str_XYFunction As String, dic_StaticParams As Dictionary, _
    dbl_InitialX1 As Double, dbl_InitialX2 As Double, dbl_TargetY As Double, dbl_Tolerance As Double, _
    int_MaxIterations As Integer, dbl_FallBackValue As Double, ByRef dic_SecondaryOutputs As Dictionary) As Double
    ' ## Perform the secant method to solve for the input which sets the function to the target value
    'QJK code 14/01/2015 FUNCTION ADDED
    Dim dbl_Output As Double
    Dim dbl_SecantX1 As Double, dbl_SecantX2 As Double, dbl_SecantX3 As Double
    Dim dbl_SecantY1 As Double, dbl_SecantY2 As Double, dbl_SecantY3 As Double
    Call dic_SecondaryOutputs.RemoveAll

    ' Set first initial guess
    dbl_SecantX1 = dbl_InitialX1
    'dbl_SecantY1 = Application.Run("'" & wbk_Caller.Name & "'!" & str_XYFunction, dbl_SecantX1, dic_StaticParams, dic_SecondaryOutputs) - dbl_TargetY
    dbl_SecantY1 = Application.Run("'" & wbk_Caller.Name & "'!" & str_XYFunction, dbl_SecantX1, dic_StaticParams, dic_SecondaryOutputs)
    ' Set second intitial guess
    dbl_SecantX2 = dbl_InitialX2
    'dbl_SecantY2 = Application.Run("'" & wbk_Caller.Name & "'!" & str_XYFunction, dbl_SecantX2, dic_StaticParams, dic_SecondaryOutputs) - dbl_TargetY
    dbl_SecantY2 = Application.Run("'" & wbk_Caller.Name & "'!" & str_XYFunction, dbl_SecantX2, dic_StaticParams, dic_SecondaryOutputs)
    ' Prepare for iteration
    Dim int_IterCtr As Integer: int_IterCtr = 0
    Dim bln_Solvable As Boolean: bln_Solvable = True

    Do
       ' If dbl_SecantY2 - dbl_SecantY1 = 0 Or dbl_SecantX2 < 0 Then
        If dbl_SecantY2 - dbl_SecantY1 = 0 Then 'dbl_SecantY2 = 0.00000001: dbl_SecantY1 = 0 '02022015 QJK code commented out line for experiment
            ' Allow greater tolerance if having difficulty solving
          '  If (Abs(dbl_SecantY3) > (dbl_Tolerance * 100)) Or int_IterCtr = 0 Then
           '     ' No solution even with looser tolerance
           '     dbl_SecantY3 = 0
                'bln_Solvable = False
         Exit Do
       ' ElseIf dbl_SecantX2 < 0 Then  'QJK code 14/01/2015

                ' Solved to looser tolerance
           ' Exit Do
           ' End If
       ' ElseIf dbl_SecantX3 < 0 Then 'QJK code 14012014
       ' dbl_SecantY3 = 0   'QJK code 14012014
       ' bln_Solvable = False   'QJK code 14012014
        End If

        If bln_Solvable = True Then
            int_IterCtr = int_IterCtr + 1

            ' Set new guess
            dbl_SecantX3 = dbl_SecantX2 - (dbl_SecantY2 - dbl_TargetY) * (dbl_SecantX2 - dbl_SecantX1) / (dbl_SecantY2 - dbl_SecantY1)
            'dbl_SecantY3 = Application.Run("'" & wbk_Caller.Name & "'!" & str_XYFunction, dbl_SecantX3, dic_StaticParams, dic_SecondaryOutputs) - dbl_TargetY
           dbl_SecantY3 = Application.Run("'" & wbk_Caller.Name & "'!" & str_XYFunction, dbl_SecantX3, dic_StaticParams, dic_SecondaryOutputs)
         '  Debug.Print "secant_x3 is" & dbl_SecantX3 & " loopcount is:" & int_IterCtr & "  caplet premiums are:" & dbl_SecantY3 & " and cap prems are: " & dbl_TargetY
        End If

        'If dbl_SecantX3 <= 0 Then Exit Do 'Or dbl_SecantX3 >= 250

        dbl_SecantX1 = dbl_SecantX2
        dbl_SecantY1 = dbl_SecantY2
        dbl_SecantX2 = dbl_SecantX3
        dbl_SecantY2 = dbl_SecantY3
    Loop Until Abs(dbl_SecantY3 - dbl_TargetY) < dbl_Tolerance Or int_IterCtr >= int_MaxIterations

   ' If dbl_SecantX2 < 0 Then
     'bln_Solvable = False
    If dbl_SecantX3 <= 0 Or dbl_SecantX3 > 1000 Then
    bln_Solvable = False
    ElseIf Abs(dbl_SecantY3 - dbl_TargetY) <= dbl_Tolerance And int_IterCtr <= int_MaxIterations Then
    bln_Solvable = True
    ElseIf (Abs(dbl_SecantY3 - dbl_TargetY) > dbl_Tolerance) Or int_IterCtr >= int_MaxIterations Then
    bln_Solvable = False
    'ElseIf dbl_SecantX3 <= 0 Then 'Or dbl_SecantX3 >= 250  '(dbl_SecantY2 - dbl_SecantY1 = 0) Or dbl_SecantX2 < 0 Then  'QJK code 14/01/2015
        'If Abs(dbl_SecantY3 - dbl_TargetY) > dbl_Tolerance And int_IterCtr >= int_MaxIterations Then 'QJK code 14/01/2015
        'bln_Solvable = False   'QJK code 14/01/2015
       'ElseIf (Abs(dbl_SecantY3 - dbl_TargetY) > dbl_Tolerance) Or int_IterCtr >= int_MaxIterations Then  'QJK code 14/01/2015
       ' bln_Solvable = False   'QJK code 14/01/2015
       ' End If  'QJK code 14/01/2015
    End If   'QJK code 14/01/2015



    ' Output final solution if possible, otherwise output the fallback
    If bln_Solvable = True Then '02022015 QJK code commented out experiment with no dbl_FallBackVALUE=-1
    dbl_Output = dbl_SecantX3
    Else  '02022015 QJK code commented out experiment with no dbl_FallBackVALUE=-1
    dbl_Output = dbl_FallBackValue  '02022015 QJK code commented out experiment with no dbl_FallBackVALUE=-1
    End If  '02022015 QJK code commented out experiment with no dbl_FallBackVALUE=-1
    Solve_SecantQJK = dbl_Output
    Call dic_SecondaryOutputs.Add("NumIterations", int_IterCtr)
    Call dic_SecondaryOutputs.Add("Solvable", bln_Solvable)
End Function



Public Function ConsistencyCheck(enu_Direction As OptionDirection, irl_Floating As IRLeg, col_fwdrate As Collection, int_CapCtr As Integer, int_CCFail As Integer, _
                                    strike As Double, PreviousCapPremiums As Double, CapPremiums As Double, Optional intLst_InterpPillars As Collection = Nothing) As String
    Dim str_PreviousCapMat As String
    Dim str_LastCapMat As String
    Dim int_NumCaplets As Integer: int_NumCaplets = Me.NumCaplets
    Dim int_LastCapletsNum As Integer
    Dim dbl_LminusK As Double
    Dim dbl_max As Double
    Dim dbl_ActiveCapletPrice As Double
    Dim str_CCpass1 As String
    Dim str_CCpass2 As String
    Dim dbl_PreviousCapPrice As Double
    Dim dbl_CurrentCapPrice As Double
    Dim int_IntCaplet As Integer
    Dim dbl_IntCapletPrice As Double
    Dim int_IntCapletCtr As Integer
    Dim dblArr_IntCapletVol() As Double
    ReDim dblArr_IntCapletVol(1 To int_NumCaplets) As Double
    Dim fld_Params As InstParams_IRS

    str_PreviousCapMat = rng_CapTerms(int_CapCtr - int_CCFail - 1, 1).Value
    str_LastCapMat = rng_CapTerms(int_CapCtr, 1).Value
    int_PreviousCapletsNum = Calc_NumPeriods(str_PreviousCapMat, fld_LegParams.PmtFreq)
    int_LastCapletsNum = Calc_NumPeriods(str_LastCapMat, fld_LegParams.PmtFreq)
    'max(L(j+1)-K,0)
    dbl_LminusK = col_fwdrate(int_LastCapletsNum) - strike
    dbl_max = WorksheetFunction.Max(dbl_LminusK, 0)
    'caplet(j+1,j+2)
    dbl_ActiveCapletPrice = irl_Floating.Calc_BSCapletValue(enu_Direction, strike, int_Deduction, cal_Deduction, True, 0.000001, int_LastCapletsNum)

    If dbl_ActiveCapletPrice >= dbl_max Then
        str_CCpass1 = True
    Else
        str_CCpass1 = False
    End If

    If bln_calibrationfailed = True Then
        Dim i As Integer
        dbl_PreviousCapPrice = 0
        For i = 1 To int_PreviousCapletsNum
            dbl_PreviousCapPrice = dbl_PreviousCapPrice + irl_Floating.Calc_BSCapletValue(enu_Direction, strike, int_Deduction, cal_Deduction, True, dblArr_FinalVols(i), i, True)
        Next i
        bln_calibrationfailed = False
    ElseIf int_CCFail = 0 Then
        dbl_PreviousCapPrice = PreviousCapPremiums
    ElseIf int_CCFail <> 0 Then
        dbl_PreviousCapPrice = 0
        For i = 1 To int_PreviousCapletsNum
            dbl_PreviousCapPrice = dbl_PreviousCapPrice + irl_Floating.Calc_BSCapletValue(enu_Direction, strike, int_Deduction, cal_Deduction, True, dblArr_FinalVols(i), i, True)
        Next i
    End If

    dbl_CurrentCapPrice = CapPremiums
    int_IntCaplet = int_LastCapletsNum - int_PreviousCapletsNum

    dbl_IntCapletPrice = 0

    For int_IntCapletCtr = int_PreviousCapletsNum + 1 To int_LastCapletsNum
        If int_IntCapletCtr <> int_LastCapletsNum Then
            dblArr_IntCapletVol(int_IntCapletCtr - int_PreviousCapletsNum) = Me.Lookup_VolCFSurfaceInterpolateOnFWD(lngArr_FinalDates_Caplet(int_IntCapletCtr), intLst_InterpPillars, False)
        Else
            dblArr_IntCapletVol(int_IntCapletCtr - int_PreviousCapletsNum) = 0.000001
        End If

        dbl_IntCapletPrice = dbl_IntCapletPrice + irl_Floating.Calc_BSCapletValue(enu_Direction, strike, int_Deduction, cal_Deduction, True, dblArr_IntCapletVol(int_IntCapletCtr - int_PreviousCapletsNum), int_IntCapletCtr, True)
    Next int_IntCapletCtr

    If dbl_CurrentCapPrice >= dbl_IntCapletPrice + dbl_PreviousCapPrice Then
        str_CCpass2 = True
    Else
        str_CCpass2 = False
    End If

    If str_CCpass1 = True And str_CCpass2 = True Then
        ConsistencyCheck = True
    Else
        ConsistencyCheck = False
    End If

End Function