Option Explicit


' ## MEMBER DATA
Private wks_Location As Worksheet
Private rng_QueryTopLeft As Range, rng_Codes As Range, rng_Orig As Range
Private rng_Market As Range, rng_Ccy As Range
Private rng_RelShocks As Range, rng_AbsShocks As Range, rng_Final As Range
Private dic_Cache_Vols As Dictionary

' Static values
Private dic_GlobalStaticInfo As Dictionary, map_Rules As MappingRules, cfg_Settings As ConfigSheet


' ## INITIALIZATION
Public Sub Initialize(wks_Input As Worksheet, bln_DataExists As Boolean, Optional dic_StaticInfoInput As Dictionary = Nothing)
    Set wks_Location = wks_Input
    Set dic_Cache_Vols = New Dictionary
    Set rng_QueryTopLeft = wks_Location.Range("A7")
    If bln_DataExists = True Then Me.AssignRanges

    ' Store static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    Set map_Rules = dic_GlobalStaticInfo(StaticInfoType.MappingRules)
    Set cfg_Settings = dic_GlobalStaticInfo(StaticInfoType.ConfigSheet)
End Sub


' ## PROPERTIES
Public Property Get NumPoints() As Long
    NumPoints = Examine_NumRows(rng_QueryTopLeft)
End Property

Public Property Get ConfigSheet() As ConfigSheet
    Set ConfigSheet = cfg_Settings
End Property


' ## METHODS - SETUP
Public Sub LoadRates()
    Dim str_SQLCode As String
    Debug.Assert map_Rules.Dict_SourceTables.Exists("EQVOL")
    Dim str_TableName As String: str_TableName = map_Rules.Dict_SourceTables("EQVOL")

    ' Query
    str_SQLCode = "SELECT [Data Date], [Code], Rate, [Market], [Currency] " _
                & "FROM " & str_TableName _
                & " WHERE [Data Date] = #" & Convert_SQLDate(cfg_Settings.CurrentDataDate) & "# " _
                & "ORDER BY Code;"

    Me.ClearCurve
    Call Action_Query_Access(cfg_Settings.RatesEqDBPath, str_SQLCode, rng_QueryTopLeft)

    Me.AssignRanges  ' Update ranges as number of points may have changed
    rng_Orig.NumberFormat = "General"
    rng_Final.NumberFormat = "General"
    Me.Scen_ApplyBase
End Sub

Public Sub AssignRanges()
    Dim int_NumPoints As Integer: int_NumPoints = Me.NumPoints  ' Requires QueryTopLeft to be known
    If int_NumPoints > 0 Then
        Set rng_Codes = rng_QueryTopLeft.Offset(0, 1).Resize(int_NumPoints, 1)
        Set rng_Orig = rng_Codes.Offset(0, 1)
        Set rng_Market = rng_Orig.Offset(0, 1)
        Set rng_Ccy = rng_Market.Offset(0, 1)
        Set rng_RelShocks = rng_Ccy.Offset(0, 2)
        Set rng_AbsShocks = rng_RelShocks.Offset(0, 1)
        Set rng_Final = rng_AbsShocks.Offset(0, 2)
    Else
        Set rng_Codes = rng_QueryTopLeft
        Set rng_Orig = rng_QueryTopLeft
        Set rng_Market = rng_QueryTopLeft
        Set rng_Ccy = rng_QueryTopLeft
        Set rng_RelShocks = rng_QueryTopLeft
        Set rng_AbsShocks = rng_QueryTopLeft
        Set rng_Final = rng_QueryTopLeft
    End If
End Sub

Public Sub ClearCurve()
    Call Action_ClearBelow(rng_QueryTopLeft, 8)
End Sub

Public Sub ResetCache_Lookups()
    Call dic_Cache_Vols.RemoveAll
End Sub


' ## METHODS - LOOKUP
Public Function Lookup_Vol(str_Code As String, Optional bln_GetOrig As Boolean = False) As Double
    ' ## Returns spot price of the equity, and stores in dictionary if key didn't already exist
    Dim dbl_Spot As Double
    Dim str_CacheKey As String: str_CacheKey = str_Code & "|" & bln_GetOrig
    If dic_Cache_Vols.Exists(str_CacheKey) = True Then
        dbl_Spot = dic_Cache_Vols(str_CacheKey)
    Else
        Dim int_FoundIndex As Integer: int_FoundIndex = Examine_FindIndex(Convert_RangeToList(rng_Codes), str_Code)
        Debug.Assert int_FoundIndex <> -1

        If bln_GetOrig = True Then
            dbl_Spot = rng_Orig(int_FoundIndex, 1).Value
        Else
            dbl_Spot = rng_Final(int_FoundIndex, 1).Value
        End If

        Call dic_Cache_Vols.Add(str_CacheKey, dbl_Spot)
    End If

    Lookup_Vol = dbl_Spot
End Function


' ## METHODS - SCENARIOS
Public Sub Scen_ApplyBase()
    ' ## Remove shifts and reset final spots to the original values
    rng_RelShocks.ClearContents
    rng_AbsShocks.ClearContents
    rng_Final.Value = rng_Orig.Value
    Call Me.ResetCache_Lookups
End Sub

Public Sub Scen_AddShock(str_Code As String, enu_ShockType As ShockType, dbl_Amount As Double)
    ' ## Apply shift to the original price
    Dim int_FoundIndex As Integer: int_FoundIndex = Examine_FindIndex(Convert_RangeToList(rng_Codes), str_Code)
    'Debug.Assert int_FoundIndex <> -1

    Select Case enu_ShockType
        Case ShockType.Absolute: rng_AbsShocks(int_FoundIndex, 1).Value = dbl_Amount
        Case ShockType.Relative: rng_RelShocks(int_FoundIndex, 1).Value = dbl_Amount
    End Select
End Sub

Public Sub Scen_AddShockMarket(str_Code As String, enu_ShockType As ShockType, dbl_Amount As Double)
    ' ## Apply shift to the original price based on Market

    Dim lng_NumRows As Long: lng_NumRows = Me.NumPoints
    Dim int_RowCtr As Integer

    For int_RowCtr = 1 To lng_NumRows

        If str_Code = rng_Market(int_RowCtr) Then

            Select Case enu_ShockType
                Case ShockType.Absolute: rng_AbsShocks(int_RowCtr, 1).Value = dbl_Amount
                Case ShockType.Relative: rng_RelShocks(int_RowCtr, 1).Value = dbl_Amount
            End Select
        End If
    Next int_RowCtr
End Sub



Public Sub Scen_ApplyCurrent()
    ' ## Calculates final spots taking into account the shocks applied
    Dim lng_NumRows As Long: lng_NumRows = Me.NumPoints
    Dim int_RowCtr As Integer

    ' Read input data
    Dim dblArr_Orig() As Double, dblArr_AbsShifts() As Double, dblArr_RelShifts() As Double
    If lng_NumRows = 1 Then
        ReDim dblArr_Orig(1 To 1, 1 To 1) As Double
        ReDim dblArr_AbsShifts(1 To 1, 1 To 1) As Double
        ReDim dblArr_RelShifts(1 To 1, 1 To 1) As Double
        dblArr_Orig(1, 1) = rng_Orig.Value
        dblArr_AbsShifts(1, 1) = rng_AbsShocks.Value
        dblArr_RelShifts(1, 1) = rng_RelShocks.Value
    Else
        ReDim dblArr_Orig(1 To lng_NumRows, 1 To 1) As Double
        ReDim dblArr_AbsShifts(1 To lng_NumRows, 1 To 1) As Double
        ReDim dblArr_RelShifts(1 To lng_NumRows, 1 To 1) As Double

        For int_RowCtr = 1 To lng_NumRows
            dblArr_Orig(int_RowCtr, 1) = rng_Orig(int_RowCtr).Value
            dblArr_AbsShifts(int_RowCtr, 1) = rng_AbsShocks(int_RowCtr).Value
            dblArr_RelShifts(int_RowCtr, 1) = rng_RelShocks(int_RowCtr).Value
        Next int_RowCtr
    End If

    ' Calculate final shift
    Dim dblArr_Final() As Double: ReDim dblArr_Final(1 To lng_NumRows, 1 To 1) As Double
    For int_RowCtr = 1 To lng_NumRows
        dblArr_Final(int_RowCtr, 1) = dblArr_Orig(int_RowCtr, 1) * (1 + dblArr_RelShifts(int_RowCtr, 1) / 100) _
            + dblArr_AbsShifts(int_RowCtr, 1)
    Next int_RowCtr

    ' Output final shift
    If lng_NumRows = 1 Then rng_Final.Value = dblArr_Final(1, 1) Else rng_Final.Value = dblArr_Final

    Call Me.ResetCache_Lookups
End Sub