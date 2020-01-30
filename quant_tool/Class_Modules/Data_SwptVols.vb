Option Explicit

Private Enum SwapPillarLookup
    StartDates = 1
    MatDates = 2
    Days = 3
    ATMVols = 4
    Terms = 5
End Enum

Private Enum InfoCol
    OptionMats = 1
    SwapStarts = 2
    SwapMats = 3
    SwapTerms = 4
    StrikeSprds = 5
    Finalvols = 6
End Enum

Private Enum SmilePillarLookup
    Strikes = 1
    vols = 2
End Enum


' ## MEMBER DATA
Private wks_Location As Worksheet
Private rng_QueryTopLeft As Range
Private rng_OptionTerms As Range, rng_SwapTerms As Range, rng_StrikeSprds As Range, rng_OrigVols As Range, rng_FinalVols As Range
Private rng_OptionMats As Range, rng_SwapStarts As Range, rng_SwapMats As Range
Private rng_Days_TopLeft As Range, rng_FirstOptionPillarSet As Range

' Dynamic variables
Private dbl_VolShift_Sens As Double
Private dic_RelShifts As Dictionary, dic_AbsShifts As Dictionary

' Static values
Private dic_GlobalStaticInfo As Dictionary, igs_Generators As IRGeneratorSet, map_Rules As MappingRules
Private cfg_Settings As ConfigSheet
Private lng_BuildDate As Long, str_CurveName As String, cal_pmt As Calendar, int_SpotDays As Integer, str_BDC_Opt As String
Private str_BDC_Swap As String, str_Interp_Time As String, str_Interp_Underlying As String, str_Gen_LegA As String, str_Gen_LegB As String
Private Const str_Interp_Strike As String = "SPLINE"


' ## INITIALIZATION
Public Sub Initialize(wks_Input As Worksheet, bln_DataExists As Boolean, Optional dic_StaticInfoInput As Dictionary = Nothing)
    Set wks_Location = wks_Input

    ' Static info
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    Set igs_Generators = dic_GlobalStaticInfo(StaticInfoType.IRGeneratorSet)
    Set map_Rules = dic_GlobalStaticInfo(StaticInfoType.MappingRules)
    Set cfg_Settings = dic_GlobalStaticInfo(StaticInfoType.ConfigSheet)
    Set dic_RelShifts = New Dictionary
    dic_RelShifts.CompareMode = CompareMethod.TextCompare
    Set dic_AbsShifts = New Dictionary
    dic_AbsShifts.CompareMode = CompareMethod.TextCompare

    If bln_DataExists = True Then
        Call ReadSheetStatic
        Call AssignRanges ' Used when reading a rate
    End If
End Sub


' ## PROPERTIES
Public Property Get NumPoints() As Integer
    NumPoints = Examine_NumRows(rng_QueryTopLeft)
End Property

Public Property Get TypeCode() As CurveType
    TypeCode = CurveType.SVL
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
Private Sub ReadSheetStatic()
    ' ## Read static values from the sheet and store in memory
    Dim rng_FirstParam As Range: Set rng_FirstParam = wks_Location.Range("A2")
    Dim int_ActiveCol As Integer: int_ActiveCol = 0
    lng_BuildDate = rng_FirstParam.Offset(0, int_ActiveCol).Value

    int_ActiveCol = int_ActiveCol + 1
    str_CurveName = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)

    int_ActiveCol = int_ActiveCol + 2
    cal_pmt = GetObject_Calendar(UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value))

    int_ActiveCol = int_ActiveCol + 1
    int_SpotDays = rng_FirstParam.Offset(0, int_ActiveCol).Value

    int_ActiveCol = int_ActiveCol + 1
    str_BDC_Opt = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)

    int_ActiveCol = int_ActiveCol + 1
    str_BDC_Swap = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)

    int_ActiveCol = int_ActiveCol + 1
    str_Interp_Time = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)

    int_ActiveCol = int_ActiveCol + 1
    str_Interp_Underlying = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)

    int_ActiveCol = int_ActiveCol + 1
    str_Gen_LegA = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)

    int_ActiveCol = int_ActiveCol + 1
    str_Gen_LegB = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)
End Sub

Public Sub LoadRates()
    ' Read parameters
    With wks_Location
        Dim str_name As String: str_name = .Range("B2").Value

        Dim str_OptionalExclusions As String
        If .Range("C2").Value = "-" Then
            str_OptionalExclusions = ""
        Else
            str_OptionalExclusions = "AND [Option SortTerm] NOT IN (" & Replace(.Range("C2").Value, "|", ", ") & ") "
        End If

        Dim str_SQLCode As String
        Set rng_QueryTopLeft = .Range("A7")
    End With

    ' Determine table name
    Debug.Assert map_Rules.Dict_SourceTables.Exists("SWPTVOL")
    Dim str_TableName As String: str_TableName = map_Rules.Dict_SourceTables("SWPTVOL")

    ' Query
    str_SQLCode = "SELECT [Data Date], Currency, [Option Term], [Option SortTerm], [Swap Term], [Swap SortTerm], StrikeSprd, Rate " _
            & "FROM " & str_TableName _
            & " WHERE [Data Date] = #" & Convert_SQLDate(cfg_Settings.CurrentDataDate) & "# AND Currency = '" & str_name & "' " & str_OptionalExclusions _
            & "ORDER BY [Swap SortTerm], [Option SortTerm], StrikeSprd;"

    Me.ClearCurve
    Call Action_Query_Access(cfg_Settings.RatesDBPath, str_SQLCode, rng_QueryTopLeft)
    Call AssignRanges

    ' Derive dates
    Call Me.GeneratePillarDates

    ' Format output
    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    Dim str_FloatNumFormat As String: str_FloatNumFormat = Gather_FloatNumFormat()
    rng_OptionMats.NumberFormat = str_DateFormat
    rng_SwapStarts.NumberFormat = str_DateFormat
    rng_SwapMats.NumberFormat = str_DateFormat
    rng_FinalVols.NumberFormat = str_FloatNumFormat

    ' Apply base scenario
    rng_FinalVols.Value = rng_OrigVols.Value
End Sub

Public Sub SetParams(rng_QueryParams As Range)
    With wks_Location
        .Range("A2").Value = cfg_Settings.CurrentBuildDate
        .Range("B2:K2").Value = rng_QueryParams.Value
    End With

    Call ReadSheetStatic
End Sub

Public Sub ClearParams()
    With wks_Location
        .Range("A2:K2").ClearContents
    End With
End Sub

Public Sub ClearCurve()
    With wks_Location
        Call Action_ClearBelow(.Range("A7"), 11)
        Call Action_ClearBelow(.Range("M7"), 5)
        Call Action_ClearBelow(.Range("S7"), 1)
    End With
End Sub

Private Sub AssignRanges()
    Set rng_QueryTopLeft = wks_Location.Range("A7")
    Dim int_NumPoints As Integer: int_NumPoints = Examine_NumRows(rng_QueryTopLeft)
    If int_NumPoints > 0 Then
        Set rng_OptionTerms = rng_QueryTopLeft.Offset(0, 2).Resize(int_NumPoints, 1)
        Set rng_SwapTerms = rng_OptionTerms.Offset(0, 2)
        Set rng_StrikeSprds = rng_SwapTerms.Offset(0, 2)
        Set rng_OrigVols = rng_StrikeSprds.Offset(0, 1)
        Set rng_OptionMats = rng_OrigVols.Offset(0, 1)
        Set rng_SwapStarts = rng_OptionMats.Offset(0, 1)
        Set rng_SwapMats = rng_SwapStarts.Offset(0, 1)
        Set rng_Days_TopLeft = rng_SwapMats(1, 1).Offset(0, 2)
        Set rng_FinalVols = rng_Days_TopLeft.Offset(0, 6).Resize(int_NumPoints, 1)

        ' Store option maturity pillar set for the first swap maturity, used for interpolation in option maturity
        Dim rng_ActiveRow As Range: Set rng_ActiveRow = rng_SwapTerms(1, 1)
        Dim str_FirstSwapTerm As String: str_FirstSwapTerm = rng_ActiveRow.Value
        Dim int_NumOptionTerms As Integer: int_NumOptionTerms = WorksheetFunction.CountIf(rng_SwapTerms, str_FirstSwapTerm)
        Set rng_FirstOptionPillarSet = rng_OptionMats(1, 1).Resize(int_NumOptionTerms, 1)
    End If
End Sub


' ## METHODS - LOOKUP
Public Function Lookup_Vol(lng_OptionMat As Long, lng_SwapMat As Long, Optional dbl_Strike As Double = -1) As Double
    ' ## Perform vol lookup based on option and swap maturity dates.  Look up the smile if a strike is provided
    Dim dbl_Output As Double

    ' Determine number of days of the actual swap
    Dim lng_SwapStart As Long: lng_SwapStart = date_workday(lng_OptionMat, int_SpotDays, cal_pmt.HolDates, cal_pmt.Weekends)
    Dim lng_SwapDays As Long: lng_SwapDays = lng_SwapMat - lng_SwapStart

    ' Find adjacent option maturity pillars
    Dim lngLst_OPillars As Collection: Set lngLst_OPillars = Convert_RangeToList(rng_FirstOptionPillarSet)
    Dim dic_AdjOPillars As Dictionary: Set dic_AdjOPillars = Gather_AdjacentPillars(lng_OptionMat, lngLst_OPillars)
    Dim lng_OPillar_Left As Long: lng_OPillar_Left = dic_AdjOPillars("Pillar_Below")
    Dim lng_OPillar_Right As Long: lng_OPillar_Right = dic_AdjOPillars("Pillar_Above")

    ' Find index of each adjacent swap maturity pillar
    Dim dic_InfoCols As Dictionary: Set dic_InfoCols = Gather_InfoColumns()
    Dim lstArr_SMatPillarInfo_Left() As Collection: lstArr_SMatPillarInfo_Left = Lookup_SwapMatPillars(lng_OPillar_Left, dic_InfoCols)
    Dim lngArr_SwapDays_Left() As Long, dblArr_Vols_Left() As Double, dbl_Vol_Left As Double
    Dim lngLst_SwapDays_Left As Collection: Set lngLst_SwapDays_Left = lstArr_SMatPillarInfo_Left(SwapPillarLookup.Days)
    Dim dic_AdjSPillars As Dictionary: Set dic_AdjSPillars = Gather_AdjacentPillars(lng_SwapDays, lngLst_SwapDays_Left)
    Dim int_SPillarIndex_Below As Integer: int_SPillarIndex_Below = dic_AdjSPillars("Index_Below")
    Dim int_SPillarIndex_Above As Integer: int_SPillarIndex_Above = dic_AdjSPillars("Index_Above")

    ' Determine days at each adjacent swap maturity pillar
    Dim lng_ActiveSwapDays_Below, lng_ActiveSwapDays_Above
    lngArr_SwapDays_Left = Convert_ListToLngArr(lngLst_SwapDays_Left)
    lng_ActiveSwapDays_Below = lngArr_SwapDays_Left(int_SPillarIndex_Below)
    lng_ActiveSwapDays_Above = lngArr_SwapDays_Left(int_SPillarIndex_Above)

    ' Determine vols at each adjacent swap maturity pillar
    Dim dbl_ActivePillarVol_Below As Double, dbl_ActivePillarVol_Above As Double
    Dim irl_ActiveLegA As IRLeg, irl_ActiveLegB As IRLeg
    Dim fld_ActiveParams_LegA As IRLegParams, fld_ActiveParams_LegB As IRLegParams
    Dim dic_SmilePillars_Below As Dictionary, dic_SmilePillars_Above As Dictionary
    Dim dblLst_ActiveSmile_Strikes As Collection, dblLst_ActiveSmile_Vols As Collection
    If dbl_Strike = -1 Then
        ' Smile off
        dblArr_Vols_Left = Convert_ListToDblArr(lstArr_SMatPillarInfo_Left(SwapPillarLookup.ATMVols))
        dbl_ActivePillarVol_Below = dblArr_Vols_Left(int_SPillarIndex_Below)
        dbl_ActivePillarVol_Above = dblArr_Vols_Left(int_SPillarIndex_Above)
    Else
        ' Smile on
        ' Gather swap for derivation of the implied swap rate
        fld_ActiveParams_LegA = igs_Generators.Lookup_Generator(str_Gen_LegA)
        fld_ActiveParams_LegB = igs_Generators.Lookup_Generator(str_Gen_LegB)
        fld_ActiveParams_LegA.ValueDate = lng_BuildDate
        fld_ActiveParams_LegB.ValueDate = lng_BuildDate

        ' Gather pillar vols
        dbl_ActivePillarVol_Below = Lookup_PillarSmileVol(fld_ActiveParams_LegA, fld_ActiveParams_LegB, lstArr_SMatPillarInfo_Left, _
            lng_OPillar_Left, int_SPillarIndex_Below, dbl_Strike)

        dbl_ActivePillarVol_Above = Lookup_PillarSmileVol(fld_ActiveParams_LegA, fld_ActiveParams_LegB, lstArr_SMatPillarInfo_Left, _
            lng_OPillar_Left, int_SPillarIndex_Above, dbl_Strike)
    End If

    ' At each adjacent pillar, interpolate the vol based on the number of days between swap start and end dates
    Select Case str_Interp_Underlying
        Case "LINEAR"
            dbl_Vol_Left = Interp_Lin_Binary(lng_ActiveSwapDays_Below, lng_ActiveSwapDays_Above, dbl_ActivePillarVol_Below, _
                dbl_ActivePillarVol_Above, lng_SwapDays)
    End Select

    If lng_OPillar_Right = lng_OPillar_Left Then
        dbl_Output = dbl_Vol_Left
    Else
        ' Determine days at each adjacent swap maturity pillar for the optional maturity pillar on the right of the lookup date
        Dim lstArr_SMatPillarInfo_Right() As Collection: lstArr_SMatPillarInfo_Right = Lookup_SwapMatPillars(lng_OPillar_Right, dic_InfoCols)
        Dim lngArr_SwapDays_Right() As Long: lngArr_SwapDays_Right = Convert_ListToLngArr(lstArr_SMatPillarInfo_Right(SwapPillarLookup.Days))

'        ' Disable this code if emulating Murex
'        Set dic_AdjSPillars = Gather_AdjacentPillars(lng_SwapDays, lstArr_SMatPillarInfo_Right(SwapPillarLookup.Days))
'        int_SPillarIndex_Below = dic_AdjSPillars("Index_Below")
'        int_SPillarIndex_Above = dic_AdjSPillars("Index_Above")

        ' to emulate Murex behavior when the deal swap days is the same as the lower option pillar swap days, remain the if statement below.
        ' to disable this logic, comment the if statement below

        'If lng_SwapDays <> lng_ActiveSwapDays_Below Then
            lng_ActiveSwapDays_Below = lngArr_SwapDays_Right(int_SPillarIndex_Below)
            lng_ActiveSwapDays_Above = lngArr_SwapDays_Right(int_SPillarIndex_Above)
        'End If

        ' Determine vols at each adjacent swap maturity pillar
        Dim dblArr_Vols_Right() As Double, dbl_Vol_Right As Double
        If dbl_Strike = -1 Then
            ' Smile off
            dblArr_Vols_Right = Convert_ListToDblArr(lstArr_SMatPillarInfo_Right(SwapPillarLookup.ATMVols))
            dbl_ActivePillarVol_Below = dblArr_Vols_Right(int_SPillarIndex_Below)
            dbl_ActivePillarVol_Above = dblArr_Vols_Right(int_SPillarIndex_Above)
        Else
            ' Gather pillar vols
            dbl_ActivePillarVol_Below = Lookup_PillarSmileVol(fld_ActiveParams_LegA, fld_ActiveParams_LegB, _
                lstArr_SMatPillarInfo_Right, lng_OPillar_Right, int_SPillarIndex_Below, dbl_Strike)

            dbl_ActivePillarVol_Above = Lookup_PillarSmileVol(fld_ActiveParams_LegA, fld_ActiveParams_LegB, _
                lstArr_SMatPillarInfo_Right, lng_OPillar_Right, int_SPillarIndex_Above, dbl_Strike)
        End If

        Select Case str_Interp_Underlying
            Case "LINEAR"
                ' Regardless of the pillars that the swap maturity actually falls between, interpolate using the same indicies as for the left option maturity
                ' Note that int_SPillarIndex_Below and int_SPillarIndex_Above are based on the left option pillar
                dbl_Vol_Right = Interp_Lin_Binary(lng_ActiveSwapDays_Below, lng_ActiveSwapDays_Above, dbl_ActivePillarVol_Below, _
                    dbl_ActivePillarVol_Above, lng_SwapDays)
        End Select

        ' Interpolate the adjacent pillar vols using the actual option maturity
        Select Case str_Interp_Time
            Case "LINEAR"
                dbl_Output = Interp_Lin_Binary(lng_OPillar_Left, lng_OPillar_Right, dbl_Vol_Left, dbl_Vol_Right, lng_OptionMat)
        End Select
    End If

    If dbl_Output <= 0 Then dbl_Output = 0.000001
    Lookup_Vol = dbl_Output
End Function

Private Function Lookup_PillarSmileVol(fld_Params_LegA As IRLegParams, fld_Params_LegB As IRLegParams, lstArr_SwapMatPillarInfo, _
    lng_OptionMat As Long, int_SPillarIndex As Integer, dbl_Strike As Double) As Double
    ' ## Look up the smile at the specified swap and option maturity combination
    Dim dbl_Output As Double

    ' Determine information about the swap from the sheet
    Dim strLst_SwapTerms As Collection, str_SwapTerm As String
    Dim lngLst_SwapStarts As Collection, lng_SwapStart As Long
    Dim lngLst_SwapMats As Collection, lng_SwapMat As Long
    If dbl_Strike <> -1 Then
        Set strLst_SwapTerms = lstArr_SwapMatPillarInfo(SwapPillarLookup.Terms)
        str_SwapTerm = strLst_SwapTerms(int_SPillarIndex)
        Set lngLst_SwapStarts = lstArr_SwapMatPillarInfo(SwapPillarLookup.StartDates)
        lng_SwapStart = lngLst_SwapStarts(int_SPillarIndex)
        Set lngLst_SwapMats = lstArr_SwapMatPillarInfo(SwapPillarLookup.MatDates)
        lng_SwapMat = lngLst_SwapMats(int_SPillarIndex)
    End If

    ' Set up the swap according to the maturity of the swap and option
    With fld_Params_LegA
        .Term = str_SwapTerm
        .GenerationRefPoint = lng_SwapStart
    End With

    With fld_Params_LegB
        .Term = str_SwapTerm
        .GenerationRefPoint = lng_SwapStart
    End With

    ' Set up swap legs and solve for ATM swap rate
    Dim irl_LegA As New IRLeg, irl_legB As New IRLeg, dbl_ATMStrike As Double
    Call irl_LegA.Initialize(fld_Params_LegA, , dic_GlobalStaticInfo)
    Call irl_legB.Initialize(fld_Params_LegB, , dic_GlobalStaticInfo)
    dbl_ATMStrike = irl_LegA.SolveParRate(irl_legB)
    'Debug.Print "Option mat: " & Format(lng_OptionMat, "dd/mm/yyyy") & "  Swap mat: " & Format(lng_SwapMat, "dd/mm/yyyy") & "  Strike: " & dbl_ATMStrike

    ' Interpolate smile vol at actual strike
    Dim dic_SmilePillars As Dictionary, dblLst_Strikes As Collection, dblLst_Vols As Collection
    Set dic_SmilePillars = Lookup_SmilePillars(lng_OptionMat, lng_SwapMat, dbl_ATMStrike)
    Set dblLst_Strikes = dic_SmilePillars(SmilePillarLookup.Strikes)
    Set dblLst_Vols = dic_SmilePillars(SmilePillarLookup.vols)

    Select Case str_Interp_Strike
        Case "LINEAR": dbl_Output = Interp_Lin(dblLst_Strikes, dblLst_Vols, dbl_Strike, True)
        Case "SPLINE": dbl_Output = Interp_Spline(dblLst_Strikes, dblLst_Vols, dbl_Strike)
    End Select

    Lookup_PillarSmileVol = dbl_Output
    'Debug.Print "Option mat: " & Format(lng_OptionMat, "dd/mm/yyyy") & "  Swap mat: " & Format(lng_SwapMat, "dd/mm/yyyy") & "  Vol: " & dbl_Output
End Function

Private Function Lookup_SwapMatPillars(lng_OptionPillar As Long, dic_InfoCols As Dictionary) As Collection()
    ' ## Return underlying swap maturity dates or ATM vols based on the specified option maturity
    Dim lstArr_Output() As Collection: ReDim lstArr_Output(1 To 5) As Collection
    Dim lngArr_OptionMats() As Variant, lngArr_SwapStarts() As Variant, lngArr_SwapMats() As Variant
    Dim strArr_SwapTerms() As Variant, dblArr_StrikeSprds() As Variant, dblArr_FinalVols() As Variant

    Dim lngLst_Output_SStarts As New Collection, lngLst_Output_SMats As New Collection, intLst_Output_Days As New Collection
    Dim dblLst_Output_ATMVols As New Collection, strLst_Output_STerms As New Collection

    lngArr_OptionMats = dic_InfoCols(InfoCol.OptionMats)
    lngArr_SwapStarts = dic_InfoCols(InfoCol.SwapStarts)
    lngArr_SwapMats = dic_InfoCols(InfoCol.SwapMats)
    strArr_SwapTerms = dic_InfoCols(InfoCol.SwapTerms)
    dblArr_StrikeSprds = dic_InfoCols(InfoCol.StrikeSprds)
    dblArr_FinalVols = dic_InfoCols(InfoCol.Finalvols)

    Dim int_ctr As Integer
    For int_ctr = 1 To Me.NumPoints
        If lngArr_OptionMats(int_ctr, 1) = lng_OptionPillar Then
            If dblArr_StrikeSprds(int_ctr, 1) = 0 Then
                Call lngLst_Output_SStarts.Add(lngArr_SwapStarts(int_ctr, 1))
                Call lngLst_Output_SMats.Add(lngArr_SwapMats(int_ctr, 1))
                Call intLst_Output_Days.Add(lngArr_SwapMats(int_ctr, 1) - lngArr_SwapStarts(int_ctr, 1))
                Call dblLst_Output_ATMVols.Add(dblArr_FinalVols(int_ctr, 1) + dbl_VolShift_Sens)
                Call strLst_Output_STerms.Add(strArr_SwapTerms(int_ctr, 1))
            End If
        End If
    Next int_ctr

    Set lstArr_Output(SwapPillarLookup.StartDates) = lngLst_Output_SStarts
    Set lstArr_Output(SwapPillarLookup.MatDates) = lngLst_Output_SMats
    Set lstArr_Output(SwapPillarLookup.Days) = intLst_Output_Days
    Set lstArr_Output(SwapPillarLookup.ATMVols) = dblLst_Output_ATMVols
    Set lstArr_Output(SwapPillarLookup.Terms) = strLst_Output_STerms

    Lookup_SwapMatPillars = lstArr_Output
End Function

Private Function Lookup_SmilePillars(lng_OptionPillar As Long, lng_SwapPillar As Long, dbl_ATMStrike As Double) As Dictionary
    ' ## Return smile and ATM volatilities at the specified pillar location.  Assumes query orders the strikes in ascending order
    Dim dic_output As New Dictionary: dic_output.CompareMode = CompareMethod.TextCompare
    Dim dblLst_Strikes As New Collection, dblLst_Vols As New Collection
    Dim lngArr_OptionMats() As Variant: lngArr_OptionMats = rng_OptionMats.Value
    Dim lngArr_SwapMats() As Variant: lngArr_SwapMats = rng_SwapMats.Value
    Dim int_ctr As Integer

    ' Find relevant section
    For int_ctr = 1 To Me.NumPoints
        If lngArr_OptionMats(int_ctr, 1) = lng_OptionPillar Then
            If lngArr_SwapMats(int_ctr, 1) = lng_SwapPillar Then
                Call dblLst_Strikes.Add(dbl_ATMStrike + rng_StrikeSprds(int_ctr, 1).Value)
                Call dblLst_Vols.Add(rng_FinalVols(int_ctr, 1).Value + dbl_VolShift_Sens)
            End If
        End If
    Next int_ctr

    ' Create output dictionary
    Call dic_output.Add(SmilePillarLookup.Strikes, dblLst_Strikes)
    Call dic_output.Add(SmilePillarLookup.vols, dblLst_Vols)
    Set Lookup_SmilePillars = dic_output
End Function

Private Function Lookup_OptionMat(str_OptionTerm As String) As Long
    ' ## Return option maturity date shown on the sheet for the specified term
    Dim int_FoundIndex As Integer: int_FoundIndex = Examine_FindIndex(Convert_RangeToList(rng_OptionTerms), str_OptionTerm)
    Debug.Assert int_FoundIndex <> -1
    Lookup_OptionMat = rng_OptionMats(int_FoundIndex, 1).Value
End Function


' ## METHODS - SECNARIOS
Public Sub Scen_ApplyBase()
    ' ## Clear shifts and reset final vols to the original vols
    Call dic_RelShifts.RemoveAll
    Call dic_AbsShifts.RemoveAll
    Call Action_ClearBelow(rng_Days_TopLeft, 5)
    rng_FinalVols.Value = rng_OrigVols.Value
End Sub

Public Sub Scen_AddShock_UniformAll(enu_ShockType As ShockType, dbl_Amount As Double)
    ' ## Shock all swap and option maturity pillars by the same amount
    Call Me.Scen_AddShock_Uniform("-", enu_ShockType, dbl_Amount)
End Sub

Public Sub Scen_AddShock_Uniform(str_SwapMat As String, enu_ShockType As ShockType, dbl_Amount As Double, Optional var_StrikeSprd As Variant = "-")
    ' ## Shock all option maturity pillars by the same amount for the specified swap term
    Dim csh_Found As CurveDaysShift: Set csh_Found = Gather_ShiftObj(str_SwapMat, var_StrikeSprd, enu_ShockType)
    Call csh_Found.AddUniformShift(dbl_Amount)
End Sub

Public Sub Scen_AddShock_Pillar(str_SwapMat As String, str_OptMat As String, enu_ShockType As ShockType, dbl_Amount As Double, _
    Optional var_StrikeSprd As Variant = "-")
    ' ## Shock the specified option maturity pillar on the specified underlying swap maturity
    Dim csh_Found As CurveDaysShift: Set csh_Found = Gather_ShiftObj(str_SwapMat, var_StrikeSprd, enu_ShockType)
    Dim lng_MatDate As Long: lng_MatDate = Lookup_OptionMat(str_OptMat)
    Call csh_Found.AddIsolatedShift(lng_MatDate - lng_BuildDate, dbl_Amount)
End Sub

Public Sub Scen_AddShock_Days(str_SwapMat As String, int_numdays As Integer, enu_ShockType As ShockType, dbl_Amount As Double, _
    Optional var_StrikeSprd As Variant = "-")
    ' ## Shock the specified number of days on the specified underlying swap maturity
    Dim csh_Found As CurveDaysShift: Set csh_Found = Gather_ShiftObj(str_SwapMat, var_StrikeSprd, enu_ShockType)
    Call csh_Found.AddShift(int_numdays, dbl_Amount)
End Sub

Public Sub Scen_ApplyCurrent()
    ' ## Alter final vols to reflect current shifts
    Dim bln_RelShifts As Boolean: bln_RelShifts = (dic_RelShifts.count > 0)
    Dim bln_AbsShifts As Boolean: bln_AbsShifts = (dic_AbsShifts.count > 0)

    If bln_RelShifts = True Or bln_AbsShifts = True Then
        Dim int_NumRows As Integer: int_NumRows = Me.NumPoints
        Dim lngArr_OptionMats() As Variant: lngArr_OptionMats = rng_OptionMats.Value
        Dim strArr_SwapTerms() As Variant: strArr_SwapTerms = rng_SwapTerms.Value
        Dim dblArr_StrikeSprds() As Variant: dblArr_StrikeSprds = rng_StrikeSprds.Value
        Dim dblArr_OrigVols() As Variant: dblArr_OrigVols = rng_OrigVols.Value
        Dim dblArr_FinalVols() As Double: ReDim dblArr_FinalVols(1 To int_NumRows, 1 To 1) As Double
        Dim csh_ActiveDelta_Rel As CurveDaysShift, csh_ActiveDelta_Abs As CurveDaysShift

        ' Gather uniform shift objects if these exist (swap maturity)
        Dim csh_UniformOnSwap_Rel As CurveDaysShift, csh_UniformOnSwap_Abs As CurveDaysShift
        If dic_RelShifts.Exists("-") Then Set csh_UniformOnSwap_Rel = dic_RelShifts("-")("-")
        If dic_AbsShifts.Exists("-") Then Set csh_UniformOnSwap_Abs = dic_AbsShifts("-")("-")

        ' Other shift objects that may potentially exist
        Dim dic_ActiveSwap_Rel As Dictionary, dic_ActiveSwap_Abs As Dictionary
        Dim csh_ActiveUniformOnSprd_Rel As CurveDaysShift, csh_ActiveUniformOnSprd_Abs As CurveDaysShift
        Dim csh_ActiveSprd_Rel As CurveDaysShift, csh_ActiveSprd_Abs As CurveDaysShift

        Dim int_ctr As Integer, str_ActiveSwapMat As String, dbl_ActiveStrikeSprd As Double, int_ActiveDays As Integer
        Dim dbl_ActiveShift_Rel As Double, dbl_ActiveShift_Abs As Double
        For int_ctr = 1 To int_NumRows
            str_ActiveSwapMat = strArr_SwapTerms(int_ctr, 1)
            dbl_ActiveStrikeSprd = dblArr_StrikeSprds(int_ctr, 1)
            int_ActiveDays = lngArr_OptionMats(int_ctr, 1) - lng_BuildDate

            ' Gather the shift objects which exist, under the specified swap term
            Set csh_ActiveUniformOnSprd_Rel = Nothing
            Set csh_ActiveSprd_Rel = Nothing
            If dic_RelShifts.Exists(str_ActiveSwapMat) Then
                Set dic_ActiveSwap_Rel = dic_RelShifts(str_ActiveSwapMat)

                If dic_ActiveSwap_Rel.Exists("-") Then Set csh_ActiveUniformOnSprd_Rel = dic_ActiveSwap_Rel("-")
                If dic_ActiveSwap_Rel.Exists(dbl_ActiveStrikeSprd) Then Set csh_ActiveSprd_Rel = dic_ActiveSwap_Rel(dbl_ActiveStrikeSprd)
            End If

            Set csh_ActiveUniformOnSprd_Abs = Nothing
            Set csh_ActiveSprd_Abs = Nothing
            If dic_AbsShifts.Exists(str_ActiveSwapMat) Then
                Set dic_ActiveSwap_Abs = dic_AbsShifts(str_ActiveSwapMat)

                If dic_ActiveSwap_Abs.Exists("-") Then Set csh_ActiveUniformOnSprd_Abs = dic_ActiveSwap_Abs("-")
                If dic_ActiveSwap_Abs.Exists(dbl_ActiveStrikeSprd) Then Set csh_ActiveSprd_Abs = dic_ActiveSwap_Abs(dbl_ActiveStrikeSprd)
            End If

            ' Determine the shift to apply to the current pillar
            dbl_ActiveShift_Rel = 0
            dbl_ActiveShift_Abs = 0
            If Not csh_UniformOnSwap_Rel Is Nothing Then dbl_ActiveShift_Rel = csh_UniformOnSwap_Rel.ReadShift(int_ActiveDays)
            If Not csh_UniformOnSwap_Abs Is Nothing Then dbl_ActiveShift_Abs = csh_UniformOnSwap_Abs.ReadShift(int_ActiveDays)
            If Not csh_ActiveUniformOnSprd_Rel Is Nothing Then dbl_ActiveShift_Rel = dbl_ActiveShift_Rel + csh_ActiveUniformOnSprd_Rel.ReadShift(int_ActiveDays)
            If Not csh_ActiveUniformOnSprd_Abs Is Nothing Then dbl_ActiveShift_Abs = dbl_ActiveShift_Abs + csh_ActiveUniformOnSprd_Abs.ReadShift(int_ActiveDays)
            If Not csh_ActiveSprd_Rel Is Nothing Then dbl_ActiveShift_Rel = dbl_ActiveShift_Rel + csh_ActiveSprd_Rel.ReadShift(int_ActiveDays)
            If Not csh_ActiveSprd_Abs Is Nothing Then dbl_ActiveShift_Abs = dbl_ActiveShift_Abs + csh_ActiveSprd_Abs.ReadShift(int_ActiveDays)

            dblArr_FinalVols(int_ctr, 1) = dblArr_OrigVols(int_ctr, 1) * (1 + dbl_ActiveShift_Rel / 100) + dbl_ActiveShift_Abs
            If dblArr_FinalVols(int_ctr, 1) <= 0 Then dblArr_FinalVols(int_ctr, 1) = 0.000001
        Next int_ctr

        ' Output shifts to sheet
        Dim rng_ActiveOutput_TopLeft As Range: Set rng_ActiveOutput_TopLeft = rng_Days_TopLeft
        If bln_RelShifts = True Then Call OutputShifts(rng_ActiveOutput_TopLeft, dic_RelShifts, 3)
        If bln_AbsShifts = True Then Call OutputShifts(rng_ActiveOutput_TopLeft, dic_AbsShifts, 4)

        ' Write shifted values back to sheet
        rng_FinalVols.Value = dblArr_FinalVols
    Else
        rng_FinalVols.Value = rng_OrigVols.Value
    End If
End Sub


' ## METHODS - SUPPORT
Public Sub GeneratePillarDates()
    ' Derive dates
    Dim dic_OptionMats As New Dictionary: dic_OptionMats.CompareMode = CompareMethod.TextCompare
    Dim lngArr_ActiveMatInfo() As Long: ReDim lngArr_ActiveMatInfo(1 To 2) As Long
    Dim dic_SwapMats As New Dictionary: dic_SwapMats.CompareMode = CompareMethod.TextCompare
    Dim str_ActiveSwapMatKey As String
    Dim str_ActiveTerm_Option As String, str_ActiveTerm_Swap As String, lng_ActiveMat_Option As Long
    Dim lng_ValDate As Long: lng_ValDate = cfg_Settings.CurrentValDate
    Dim lng_ActiveStart_Swap As Long, lng_ActiveMat_Swap As Long

    ' Read terms from sheet and prepare output array
    Dim int_NumPoints As Integer: int_NumPoints = Me.NumPoints
    Dim strArr_OptionTerms() As Variant
    Dim strArr_SwapTerms() As Variant
    strArr_OptionTerms = rng_OptionTerms(1, 1).Resize(int_NumPoints, 1).Value
    strArr_SwapTerms = rng_SwapTerms(1, 1).Resize(int_NumPoints, 1).Value
    Dim lngArr_OutputDates() As Long: ReDim lngArr_OutputDates(1 To int_NumPoints, 1 To 3) As Long

    Dim int_ctr As Integer
    For int_ctr = 1 To int_NumPoints
        ' Derive option maturity pillar dates
        str_ActiveTerm_Option = strArr_OptionTerms(int_ctr, 1)
        str_ActiveTerm_Swap = strArr_SwapTerms(int_ctr, 1)

        If dic_OptionMats.Exists(str_ActiveTerm_Option) = True Then
            lngArr_ActiveMatInfo = dic_OptionMats(str_ActiveTerm_Option)
            lng_ActiveMat_Option = lngArr_ActiveMatInfo(1)
            lng_ActiveStart_Swap = lngArr_ActiveMatInfo(2)
        Else
            lng_ActiveMat_Option = date_addterm(lng_ValDate, str_ActiveTerm_Option, 1, True)
            lng_ActiveMat_Option = Date_ApplyBDC(lng_ActiveMat_Option, str_BDC_Opt, cal_pmt.HolDates, cal_pmt.Weekends)
            lng_ActiveStart_Swap = date_workday(lng_ActiveMat_Option, int_SpotDays, cal_pmt.HolDates, cal_pmt.Weekends)
            lngArr_ActiveMatInfo(1) = lng_ActiveMat_Option
            lngArr_ActiveMatInfo(2) = lng_ActiveStart_Swap
            Call dic_OptionMats.Add(str_ActiveTerm_Option, lngArr_ActiveMatInfo)
        End If

        ' Derive underlying swap maturity pillar dates
        str_ActiveSwapMatKey = str_ActiveTerm_Option & str_ActiveTerm_Swap
        If dic_SwapMats.Exists(str_ActiveSwapMatKey) Then
            lng_ActiveMat_Swap = dic_SwapMats(str_ActiveSwapMatKey)
        Else
            lng_ActiveMat_Swap = date_addterm(lng_ActiveStart_Swap, str_ActiveTerm_Swap, 1, True)
            lng_ActiveMat_Swap = Date_ApplyBDC(lng_ActiveMat_Swap, str_BDC_Swap, cal_pmt.HolDates, cal_pmt.Weekends)
            Call dic_SwapMats.Add(str_ActiveSwapMatKey, lng_ActiveMat_Swap)
        End If

        ' Store dates in output array
        lngArr_OutputDates(int_ctr, 1) = lng_ActiveMat_Option
        lngArr_OutputDates(int_ctr, 2) = lng_ActiveStart_Swap
        lngArr_OutputDates(int_ctr, 3) = lng_ActiveMat_Swap
    Next int_ctr

    ' Output to sheet
    rng_OptionMats(1, 1).Resize(int_NumPoints, 3).Value = lngArr_OutputDates
End Sub

Private Sub OutputShifts(ByRef rng_ActiveOutput_TopLeft As Range, dic_Outer As Dictionary, int_ShiftsOffset As Integer)
    ' ## Output all shifts in the dictionary, then update the output top left to the next row below
    Dim dic_ActiveSwap As Dictionary, csh_ActiveOutput As CurveDaysShift, int_ActiveNumRows As Integer
    Dim var_ActiveSwapMat As Variant, var_ActiveSprd As Variant

    For Each var_ActiveSwapMat In dic_Outer.Keys
        Set dic_ActiveSwap = dic_Outer(var_ActiveSwapMat)

        For Each var_ActiveSprd In dic_ActiveSwap.Keys
            Set csh_ActiveOutput = dic_ActiveSwap(var_ActiveSprd)
            int_ActiveNumRows = csh_ActiveOutput.NumShifts
            rng_ActiveOutput_TopLeft.Resize(int_ActiveNumRows, 1).Value = var_ActiveSwapMat
            rng_ActiveOutput_TopLeft.Offset(0, 1).Resize(int_ActiveNumRows, 1).Value = csh_ActiveOutput.Days_Arr
            rng_ActiveOutput_TopLeft.Offset(0, 2).Resize(int_ActiveNumRows, 1).Value = var_ActiveSprd
            rng_ActiveOutput_TopLeft.Offset(0, int_ShiftsOffset).Resize(int_ActiveNumRows, 1).Value = csh_ActiveOutput.Shifts_Arr

            Set rng_ActiveOutput_TopLeft = rng_ActiveOutput_TopLeft.Offset(int_ActiveNumRows, 0)
        Next var_ActiveSprd
    Next var_ActiveSwapMat
End Sub

Private Function Gather_InfoColumns() As Dictionary
    ' ## Return dictionary containing various columns in array form.  To improve performance
    Dim dic_output As New Dictionary: dic_output.CompareMode = CompareMethod.TextCompare
    Dim int_NumPoints As Integer: int_NumPoints = Me.NumPoints
    Dim varArr_ActiveCol() As Variant

    varArr_ActiveCol = rng_OptionMats(1, 1).Resize(int_NumPoints, 1).Value
    Call dic_output.Add(InfoCol.OptionMats, varArr_ActiveCol)

    varArr_ActiveCol = rng_SwapStarts(1, 1).Resize(int_NumPoints, 1).Value
    Call dic_output.Add(InfoCol.SwapStarts, varArr_ActiveCol)

    varArr_ActiveCol = rng_SwapMats(1, 1).Resize(int_NumPoints, 1).Value
    Call dic_output.Add(InfoCol.SwapMats, varArr_ActiveCol)

    varArr_ActiveCol = rng_SwapTerms(1, 1).Resize(int_NumPoints, 1).Value
    Call dic_output.Add(InfoCol.SwapTerms, varArr_ActiveCol)

    varArr_ActiveCol = rng_StrikeSprds(1, 1).Resize(int_NumPoints, 1).Value
    Call dic_output.Add(InfoCol.StrikeSprds, varArr_ActiveCol)

    varArr_ActiveCol = rng_FinalVols(1, 1).Resize(int_NumPoints, 1).Value
    Call dic_output.Add(InfoCol.Finalvols, varArr_ActiveCol)

    Set Gather_InfoColumns = dic_output
End Function

Private Function Gather_ShiftObj(str_SwapTerm As String, var_StrikeSprd As Variant, enu_ShockType As ShockType) As CurveDaysShift
    ' ## Return object containing the shifts for the specified delta.  For a uniform shift across deltas, specify 0 as the delta input
    ' Determine shift type
    Dim dic_ToUse As Dictionary
    Select Case enu_ShockType
        Case ShockType.Absolute: Set dic_ToUse = dic_AbsShifts
        Case ShockType.Relative: Set dic_ToUse = dic_RelShifts
    End Select

    ' Filter to the specified swap term, otherwise create it
    Dim dic_FoundSwapTerm As Dictionary
    If dic_ToUse.Exists(str_SwapTerm) Then
        Set dic_FoundSwapTerm = dic_ToUse(str_SwapTerm)
    Else
        Set dic_FoundSwapTerm = New Dictionary
        Call dic_ToUse.Add(str_SwapTerm, dic_FoundSwapTerm)
    End If

    ' Gather shift object with the specified strike spread, otherwise create it
    Dim csh_Found As CurveDaysShift
    If dic_FoundSwapTerm.Exists(var_StrikeSprd) Then
        Set csh_Found = dic_FoundSwapTerm(var_StrikeSprd)
    Else
        Set csh_Found = New CurveDaysShift
        Call csh_Found.Initialize(enu_ShockType)
        Call dic_FoundSwapTerm.Add(var_StrikeSprd, csh_Found)
    End If

    Set Gather_ShiftObj = csh_Found
End Function


' ## METHODS - OUTPUT
Public Sub OutputFinalVols(rng_OutputStart As Range)
    Dim int_NumRows As Integer: int_NumRows = NumPoints()
    If int_NumRows > 0 Then
        Dim varArr_Output() As Variant: ReDim varArr_Output(1 To int_NumRows, 1 To 6) As Variant
        Dim strArr_SwapTerms() As Variant: strArr_SwapTerms = rng_SwapTerms.Value
        Dim lngArr_OptionMats() As Variant: lngArr_OptionMats = rng_OptionMats.Value
        Dim dblArr_FinalVols() As Variant: dblArr_FinalVols = rng_FinalVols.Value
        Dim dblArr_StrikeSprds() As Variant: dblArr_StrikeSprds = rng_StrikeSprds.Value

        Dim int_RowCtr As Integer
        For int_RowCtr = 1 To int_NumRows
            varArr_Output(int_RowCtr, 1) = str_CurveName
            varArr_Output(int_RowCtr, 2) = lng_BuildDate
            varArr_Output(int_RowCtr, 3) = CInt(lngArr_OptionMats(int_RowCtr, 1) - lng_BuildDate)
            varArr_Output(int_RowCtr, 4) = dblArr_FinalVols(int_RowCtr, 1)
            varArr_Output(int_RowCtr, 5) = strArr_SwapTerms(int_RowCtr, 1)
            varArr_Output(int_RowCtr, 6) = dblArr_StrikeSprds(int_RowCtr, 1)
        Next int_RowCtr

        rng_OutputStart.Resize(int_NumRows, 6).Value = varArr_Output
    End If
End Subs