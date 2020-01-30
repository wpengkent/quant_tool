Option Explicit
' ## MEMBER DATA
Private wks_Location As Worksheet
Private rng_QueryTopLeft As Range
Private rng_Terms As Range, rng_StrikePerc As Range, rng_OrigVols As Range, rng_Maturity As Range, rng_FinalVols As Range
Private rng_ShiftStrikes As Range, rng_FinalMaturity As Range, rng_FinalBuildDate As Range

'Constants
Private Const str_flat_extp As String = "Smile_FlatExtp"

' Dynamic variables
Private dic_RelShifts As Dictionary, dic_AbsShifts As Dictionary

' Static values
Private dic_GlobalStaticInfo As Dictionary, map_Rules As MappingRules, cfg_Settings As ConfigSheet
Private str_CurveName As String, str_Market As String, str_EQCode As String
Private eq_spot As Data_EQSpots, dbl_eqspot As Double

' ## INITIALIZATION
Public Sub Initialize(wks_Input As Worksheet, bln_DataExists As Boolean, Optional dic_StaticInfoInput As Dictionary = Nothing)
    Set wks_Location = wks_Input

    ' Static info
    If dic_StaticInfoInput Is Nothing Then
        Set dic_GlobalStaticInfo = GetStaticInfo()
    Else
        Set dic_GlobalStaticInfo = dic_StaticInfoInput
    End If
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
    TypeCode = CurveType.EVL
End Property

Public Property Get ConfigSheet() As ConfigSheet
    Set ConfigSheet = cfg_Settings
End Property

Public Property Get EQCode() As String
    EQCode = str_EQCode
End Property

Public Function Lookup_Vol(dbl_Strike As Double, lng_MatDate As Long, Optional var_smile_type As Variant = "") As Double
    Dim dic_Interp As Dictionary: Set dic_Interp = Gather_InterpDictEQSmile()
    Dim n As Integer: n = dic_Interp.count '# number of strikes
    Dim nt As Integer '# number of time pillars
    Dim lng_BuildDate As Long: lng_BuildDate = cfg_Settings.CurrentValDate
    Dim t As Double
    Dim int_lower As Integer, int_upper As Integer ' doesnt work for beyond time pillar

    Dim i As Integer, j As Integer
    Dim vol1 As Double, vol2 As Double, pillar1 As Double, pillar2 As Double

    '## Array construction for interpolation
    Dim arr_vol1() As Variant, arr_vol2() As Variant
    Dim arr_temp1() As Variant, arr_temp2() As Variant
    Dim arr_InterpVol() As Variant
    Dim arr_strikescale() As Variant

    '## determine the time - Flat extrapolation at time
    t = (lng_MatDate - lng_BuildDate) / 365
    nt = UBound(dic_Interp.Items()(0), 2)

    '## Lower and upper index of the pillar array
    If dic_Interp.Items()(0)(2, 1) > t Then

            '''-----Alvin's workaround 5/6/2018----''
            If WorksheetFunction.Round(dic_Interp.Items()(0)(2, 1), 0) = 0 Then
                        int_lower = WorksheetFunction.RoundUp(dic_Interp.Items()(0)(2, 1), 0)
                        int_upper = WorksheetFunction.RoundUp(dic_Interp.Items()(0)(2, 2), 0)
            Else
                        int_lower = dic_Interp.Items()(0)(2, 1)
                        int_upper = dic_Interp.Items()(0)(2, 2)
            End If
           '''-----Alvin's workaround 5/6/2018----''

    ElseIf dic_Interp.Items()(0)(2, nt) < t Then
'
'        int_lower = dic_Interp.Items()(0)(2, nt - 1)
'        int_upper = dic_Interp.Items()(0)(2, nt)

'Alv Edit for Extrp
        int_lower = nt
        int_upper = nt
'Alv Edit for Extrp

    Else
        arr_temp1 = dic_Interp.Items()(0)
        arr_temp2 = GetArrDim(2, arr_temp1)
        int_lower = Application.Match(t, arr_temp2)

                    '''-----Alvin's workaround 5/6/2018----''
                     If int_lower = nt Then
                         int_upper = int_lower
                     Else
                         int_upper = int_lower + 1
                     End If

                    '''-----Alvin's workaround 5/6/2018----''
        End If


    '# Interpolate across strike - extrapolate at strike end points
    ReDim arr_strikescale(1 To n)
    ReDim arr_vol1(1 To n), arr_vol2(1 To n)

    For i = 1 To n
        arr_vol1(i) = dic_Interp.Items()(i - 1)(1, int_lower)
        arr_vol2(i) = dic_Interp.Items()(i - 1)(1, int_upper)
        arr_strikescale(i) = dic_Interp.Keys(i - 1)
    Next i


    '''-----Alvin: Flat Extp or Linear 28/6/2018----''
If var_smile_type = str_flat_extp Then    'Flat extrapolation
    vol1 = Interp_Lin(arr_strikescale, arr_vol1, dbl_Strike, True)
    If vol1 < 0 Then
        vol1 = 0.0000001
    End If

    vol2 = Interp_Lin(arr_strikescale, arr_vol2, dbl_Strike, True)
    If vol2 < 0 Then
        vol2 = 0.0000001
     End If
Else
        vol1 = Interp_Lin(arr_strikescale, arr_vol1, dbl_Strike, False)
    If vol1 < 0 Then
        vol1 = 0.0000001
    End If

    vol2 = Interp_Lin(arr_strikescale, arr_vol2, dbl_Strike, False)
    If vol2 < 0 Then
        vol2 = 0.0000001
    End If
End If
    '''-----Alvin: Flat Extp or Linear 28/6/2018-----''



    '#v2t Interpolate - flat at end points
    pillar1 = dic_Interp.Items()(0)(2, int_lower)
    pillar2 = dic_Interp.Items()(0)(2, int_upper)

    If t < dic_Interp.Items()(0)(2, 1) Then
        Lookup_Vol = vol1
    ElseIf t > dic_Interp.Items()(0)(2, nt) Then
        Lookup_Vol = vol2
    Else
        Lookup_Vol = Interp_V2t_Binary(pillar1, pillar2, vol1, vol2, 0, t)


    End If

End Function

' ## METHODS - SETUP
Private Sub ReadSheetStatic()
    ' ## Read static values from the sheet and store in memory
    Dim rng_FirstParam As Range: Set rng_FirstParam = wks_Location.Range("A2")
    Dim int_ActiveCol As Integer: int_ActiveCol = 0
    'lng_BuildDate = rng_FirstParam.Offset(0, int_ActiveCol).Value

    int_ActiveCol = int_ActiveCol + 1
    str_CurveName = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)

    int_ActiveCol = int_ActiveCol + 1
    str_Market = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)

    int_ActiveCol = int_ActiveCol + 1
    str_EQCode = UCase(rng_FirstParam.Offset(0, int_ActiveCol).Value)

    'Read Equity Spot data
    Set eq_spot = GetObject_EQSpots(True, dic_GlobalStaticInfo)
    dbl_eqspot = eq_spot.Lookup_Spot(str_EQCode)


End Sub
Public Sub LoadRates()
    ' Read parameters
    With wks_Location
        Dim str_name As String: str_name = .Range("B2").Value
        Dim str_Code As String: str_Code = .Range("D2").Value
        Dim str_SQLCode As String
        Set rng_QueryTopLeft = .Range("A7")
    End With

    ' Determine table name
    Debug.Assert map_Rules.Dict_SourceTables.Exists("EQSMILE")
    Dim str_TableName As String: str_TableName = map_Rules.Dict_SourceTables("EQSMILE")

    ' Query
    str_SQLCode = "SELECT [Data Date], Currency, [Term], [SortTerm], Strike, Rate " _
            & "FROM " & str_TableName _
            & " WHERE [Data Date] = #" & Convert_SQLDate(cfg_Settings.CurrentDataDate) & "# AND Currency = '" & str_name & "' " _
            & " AND [Equity] ='" & str_Code & "' " _
            & "ORDER BY [SortTerm], Strike;"

    Me.ClearCurve
    Call Action_Query_Access(cfg_Settings.RatesDBPath, str_SQLCode, rng_QueryTopLeft)
    Call AssignRanges

    ' Derive dates
    Call Me.GenerateMaturityDates(False)

    ' Format output
    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    Dim str_FloatNumFormat As String: str_FloatNumFormat = Gather_FloatNumFormat()
    rng_Maturity.NumberFormat = str_DateFormat
    rng_FinalMaturity.NumberFormat = str_DateFormat
    rng_FinalVols.NumberFormat = str_FloatNumFormat

    ' Apply base scenario
    rng_FinalVols.Value = rng_OrigVols.Value
End Sub

Public Sub SetParams(rng_QueryParams As Range)
    With wks_Location
        .Range("A2").Value = cfg_Settings.CurrentBuildDate
        .Range("B2:C2").Value = rng_QueryParams.Value
        .Range("D2").Value = rng_QueryParams(, -1)
    End With

    Call ReadSheetStatic
End Sub

Public Sub ClearParams()
    With wks_Location
        .Range("A2:D2").ClearContents
    End With
End Sub

Public Sub ClearCurve()
    With wks_Location
        Call Action_ClearBelow(.Range("A7"), 7)
        Call Action_ClearBelow(.Range("I7"), 4)
        Call Action_ClearBelow(.Range("N7"), 1)
    End With
End Sub

Private Sub AssignRanges()
    Set rng_QueryTopLeft = wks_Location.Range("A7")
    Dim int_NumPoints As Integer: int_NumPoints = Examine_NumRows(rng_QueryTopLeft)
    If int_NumPoints > 0 Then
        Set rng_Terms = rng_QueryTopLeft.Offset(0, 2).Resize(int_NumPoints, 1)
        Set rng_StrikePerc = rng_Terms.Offset(0, 2)
        Set rng_OrigVols = rng_StrikePerc.Offset(0, 1)
        Set rng_Maturity = rng_OrigVols.Offset(0, 1)
        Set rng_ShiftStrikes = rng_Maturity.Offset(0, 2)
        Set rng_FinalVols = rng_ShiftStrikes.Offset(0, 5).Resize(int_NumPoints, 1)
        Set rng_FinalMaturity = rng_FinalVols.Offset(0, 1).Resize(int_NumPoints, 1)
        Set rng_FinalBuildDate = wks_Location.Range("N2")
    End If
End Sub

' ## METHODS - LOOKUP
' ## METHODS - SECNARIOS
Public Sub Scen_ApplyBase()
    ' ## Clear shifts and reset final vols to the original vols
    Call dic_RelShifts.RemoveAll
    Call dic_AbsShifts.RemoveAll
    Call Action_ClearBelow(rng_ShiftStrikes, 4)
    rng_FinalVols.Value = rng_OrigVols.Value
End Sub

Public Sub Scen_AddShock_Days(var_StrikeSprd As Variant, int_numdays As Integer, enu_ShockType As ShockType, dbl_Amount As Double)
    '## Shock the specific number of days on the specified strike spread
       Dim csh_Found As CurveDaysShift: Set csh_Found = Gather_ShiftObjEvl(var_StrikeSprd, enu_ShockType)
       Call csh_Found.AddShift(int_numdays, dbl_Amount)
End Sub

Public Sub Scen_AddShock_UniformStrike(var_StrikeSprd As Variant, enu_ShockType As ShockType, dbl_Amount As Double)
    '## Shock all maturity pillars by the same amount for a specific strike spread
       Dim csh_Found As CurveDaysShift: Set csh_Found = Gather_ShiftObjEvl(var_StrikeSprd, enu_ShockType)
       Call csh_Found.AddUniformShift(dbl_Amount)
End Sub

Public Sub Scen_AddShock_UniformAll(enu_ShockType As ShockType, dbl_Amount As Double)
    '## Shock all strikes and all maturity pillars
    Call Me.Scen_AddShock_UniformStrike("-", enu_ShockType, dbl_Amount)
End Sub

Public Sub Scen_ApplyCurrent()
    ' ## Alter final vols to reflect current shifts
    Dim bln_RelShifts As Boolean: bln_RelShifts = (dic_RelShifts.count > 0)
    Dim bln_AbsShifts As Boolean: bln_AbsShifts = (dic_AbsShifts.count > 0)

    If bln_RelShifts = True Or bln_AbsShifts = True Then
        Dim int_NumRows As Integer: int_NumRows = Me.NumPoints
        Dim Arr_MatTerm() As Variant: Arr_MatTerm = rng_Terms.Value
        Dim arr_strike() As Variant: arr_strike = rng_StrikePerc.Value
        Dim lng_BuildDate As Long: lng_BuildDate = cfg_Settings.CurrentValDate
        Dim Arr_OrigVols() As Variant: Arr_OrigVols = rng_OrigVols.Value
        Dim Arr_StrikeSprds() As Variant: Arr_StrikeSprds = rng_StrikePerc.Value
        Dim Arr_Maturity() As Variant: Arr_Maturity = rng_Maturity.Value
        Dim Arr_FinalVols() As Variant: ReDim Arr_FinalVols(1 To int_NumRows, 1 To 1) As Variant

        Dim int_ctr As Integer, str_ActiveMatTerm As String, dbl_ActiveStrikeSprd As Variant, int_ActiveDays As Integer
        Dim dbl_ActiveShift_Rel As Variant, dbl_ActiveShift_Abs As Variant



        For int_ctr = 1 To int_NumRows
            str_ActiveMatTerm = Arr_MatTerm(int_ctr, 1)
            dbl_ActiveStrikeSprd = Arr_StrikeSprds(int_ctr, 1)
            int_ActiveDays = Arr_Maturity(int_ctr, 1) - lng_BuildDate

            dbl_ActiveShift_Rel = 0
            dbl_ActiveShift_Abs = 0

            Dim csh_ActiveShift_Rel As CurveDaysShift: Set csh_ActiveShift_Rel = Nothing
            Dim csh_ActiveShift_Abs As CurveDaysShift: Set csh_ActiveShift_Abs = Nothing
            Dim csh_UniformShift_Rel As CurveDaysShift: Set csh_UniformShift_Rel = Nothing
            Dim csh_UniformShift_Abs As CurveDaysShift: Set csh_UniformShift_Abs = Nothing

            If dic_RelShifts.Exists(dbl_ActiveStrikeSprd) Then
                Set csh_ActiveShift_Rel = dic_RelShifts(dbl_ActiveStrikeSprd)
                dbl_ActiveShift_Rel = dbl_ActiveShift_Rel + csh_ActiveShift_Rel.ReadShift(int_ActiveDays)
            End If

            If dic_AbsShifts.Exists(dbl_ActiveStrikeSprd) Then
                Set csh_ActiveShift_Abs = dic_AbsShifts(dbl_ActiveStrikeSprd)
                dbl_ActiveShift_Abs = dbl_ActiveShift_Abs + csh_ActiveShift_Abs.ReadShift(int_ActiveDays)
            End If

            If dic_RelShifts.Exists("-") Then
                Set csh_UniformShift_Rel = dic_RelShifts("-")
                dbl_ActiveShift_Rel = dbl_ActiveShift_Rel + csh_UniformShift_Rel.ReadShift(0)
            End If

            If dic_AbsShifts.Exists("-") Then
                Set csh_UniformShift_Abs = dic_AbsShifts("-")
                dbl_ActiveShift_Abs = dbl_ActiveShift_Abs + csh_UniformShift_Abs.ReadShift(0)
            End If

            Arr_FinalVols(int_ctr, 1) = Arr_OrigVols(int_ctr, 1) * (1 + dbl_ActiveShift_Rel / 100) + dbl_ActiveShift_Abs
            If Arr_FinalVols(int_ctr, 1) <= 0 Then
                Arr_FinalVols(int_ctr, 1) = 0.000001
            End If
        Next int_ctr
        'Output Shifts to sheet
        Dim rng_ActiveOutput_TopLeft As Range: Set rng_ActiveOutput_TopLeft = rng_ShiftStrikes
        If bln_RelShifts = True Then Call OutputShifts(rng_ActiveOutput_TopLeft, dic_RelShifts, 2)
        If bln_AbsShifts = True Then Call OutputShifts(rng_ActiveOutput_TopLeft, dic_AbsShifts, 3)

        rng_FinalVols.Value = Arr_FinalVols
    Else
        rng_FinalVols.Value = rng_OrigVols.Value
    End If
End Sub

' ## METHODS - SUPPORT
Public Sub GenerateMaturityDates(var_Final As Variant)
    'Generate volatility maturity dates based on term
    Dim bln_Final As Boolean: bln_Final = CBool(var_Final)
    Dim lng_BuildDate As Long: lng_BuildDate = cfg_Settings.CurrentValDate
    Dim int_NumPoints As Integer: int_NumPoints = Me.NumPoints
    Dim int_ctr As Integer
    Dim lng_ValDate As Long
    '## ----Normal Case Date generation from Load Rates----
    If bln_Final = False Then
        lng_ValDate = cfg_Settings.CurrentBuildDate
        rng_FinalBuildDate = lng_ValDate
        For int_ctr = 1 To int_NumPoints
            rng_Maturity(int_ctr, 1) = date_addterm(lng_ValDate, rng_Terms(int_ctr, 1), 1, True)
            rng_FinalMaturity(int_ctr, 1) = rng_Maturity(int_ctr, 1).Value
        Next int_ctr
    Else
    '## ----Var Scenario Date ----
        lng_ValDate = cfg_Settings.CurrentValDate
        rng_FinalBuildDate = lng_ValDate
        For int_ctr = 1 To int_NumPoints
            rng_FinalMaturity(int_ctr, 1) = date_addterm(lng_ValDate, rng_Terms(int_ctr, 1), 1, True)
        Next int_ctr
    End If

End Sub

Private Function Gather_InterpDictEQSmile() As Dictionary
    Dim dic_output As New Dictionary
    Dim int_NumPoints As Integer: int_NumPoints = Me.NumPoints
    Dim i As Integer, n As Integer
    Dim arr_temp() As Variant, arr_temp2() As Variant
    Dim lng_BuildDate As Long: lng_BuildDate = cfg_Settings.CurrentValDate

    ReDim arr_temp2(1 To 2, 1 To 1)
    Call eq_spot.ResetCache_Lookups
    dbl_eqspot = eq_spot.Lookup_Spot(str_EQCode)
    For i = 1 To int_NumPoints
        If dic_output.Exists(rng_StrikePerc(i).Value * dbl_eqspot) = False Then
            ReDim arr_temp(1 To 2, 1 To 1)
            arr_temp(1, 1) = rng_FinalVols(i).Value
            arr_temp(2, 1) = (rng_FinalMaturity(i).Value - lng_BuildDate) / 365
            Call dic_output.Add(rng_StrikePerc(i).Value * dbl_eqspot, arr_temp)

        ElseIf dic_output.Exists(rng_StrikePerc(i).Value * dbl_eqspot) = True Then
            n = UBound(dic_output(rng_StrikePerc(i).Value * dbl_eqspot), 2) + 1
            arr_temp2 = dic_output(rng_StrikePerc(i).Value * dbl_eqspot)
            ReDim Preserve arr_temp2(1 To 2, 1 To n)
            arr_temp2(1, n) = rng_FinalVols(i).Value
            arr_temp2(2, n) = (rng_FinalMaturity(i).Value - lng_BuildDate) / 365
            dic_output(rng_StrikePerc(i).Value * dbl_eqspot) = arr_temp2
        End If
    Next i
    Set Gather_InterpDictEQSmile = dic_output
End Function

Private Function Gather_ShiftObjEvl(var_StrikeSprd As Variant, enu_ShockType As ShockType)
    '## Return object containing the shifts for the specified strikespread
    'Determine shift type
    Dim dic_ToUse As Dictionary
    Select Case enu_ShockType
        Case ShockType.Absolute: Set dic_ToUse = dic_AbsShifts
        Case ShockType.Relative: Set dic_ToUse = dic_RelShifts
    End Select

    Dim csh_Found As CurveDaysShift
    If dic_ToUse.Exists(var_StrikeSprd) Then
        Set csh_Found = dic_ToUse(var_StrikeSprd)
    Else
        Set csh_Found = New CurveDaysShift
        Call csh_Found.Initialize(enu_ShockType)
        Call dic_ToUse.Add(var_StrikeSprd, csh_Found)
    End If

    Set Gather_ShiftObjEvl = csh_Found
End Function
Private Sub OutputShifts(rng_ActiveOutput_TopLeft As Range, dic_Outer As Dictionary, int_ShiftsOffset As Integer)
    ' ## Output all shifts in the dictionary, then update the output top left to the next row below
    Dim csh_ActiveOutput As CurveDaysShift, int_ActiveNumRows As Integer
    Dim var_ActiveSprd As Variant

    For Each var_ActiveSprd In dic_Outer.Keys
        Set csh_ActiveOutput = dic_Outer(var_ActiveSprd)
        int_ActiveNumRows = csh_ActiveOutput.NumShifts
        rng_ActiveOutput_TopLeft.Resize(int_ActiveNumRows, 1).Value = var_ActiveSprd
        rng_ActiveOutput_TopLeft.Offset(0, 1).Resize(int_ActiveNumRows, 1).Value = csh_ActiveOutput.Days_Arr
        rng_ActiveOutput_TopLeft.Offset(0, int_ShiftsOffset).Resize(int_ActiveNumRows, 1).Value = csh_ActiveOutput.Shifts_Arr

        Set rng_ActiveOutput_TopLeft = rng_ActiveOutput_TopLeft.Offset(int_ActiveNumRows, 0)
    Next var_ActiveSprd
End Sub

'## Supporting functions

Public Function GetArrDim(x As Integer, arr() As Variant) As Variant()
Dim n As Integer, i As Integer
n = UBound(arr, 2)
Dim arr_t() As Variant
ReDim arr_t(1 To n)
For i = 1 To n
    arr_t(i) = arr(x, i)
Next i
GetArrDim = arr_t
End Function