Option Explicit


' ## MEMBER DATA
Private wks_Location As Worksheet
Private rng_QueryTopLeft As Range, rng_CCYs As Range, rng_Basis As Range, rng_OrigSpots As Range, rng_tempRel As Range, rng_tempAbs As Range
Private rng_RelShocks As Range, rng_AbsShocks As Range, rng_FinalSpots As Range
Private dic_Cache_DiscSpots As Dictionary, dic_Cache_USDPerFgn As Dictionary
Private dic_FXCurveNames As Dictionary

' Dependent curves
Private dic_IRCurves As Dictionary

' Dynamic variables
Private bln_Caching As Boolean, bln_ZShiftsEnabled_DiscSpot As Boolean

' Static values
Private dic_GlobalStaticInfo As Dictionary, map_Rules As MappingRules, cfg_Settings As ConfigSheet

Private rng_OrigRate As Range, rng_OrigFinalRate As Range


' ## INITIALIZATION
Public Sub Initialize(wks_Input As Worksheet, bln_DataExists As Boolean, Optional dic_StaticInfoInput As Dictionary = Nothing)
    Set wks_Location = wks_Input
    Set dic_Cache_DiscSpots = New Dictionary
    Set dic_Cache_USDPerFgn = New Dictionary
    Set rng_QueryTopLeft = wks_Location.Range("A7")
    If bln_DataExists = True Then Me.AssignRanges

    ' Initialize dynamic variables
    bln_Caching = True
    bln_ZShiftsEnabled_DiscSpot = True

    ' Store static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    Set map_Rules = dic_GlobalStaticInfo(StaticInfoType.MappingRules)
    Set cfg_Settings = dic_GlobalStaticInfo(StaticInfoType.ConfigSheet)
    Set dic_FXCurveNames = map_Rules.Dict_FXCurveNames
End Sub


' ## PROPERTIES
Public Property Get NumPoints() As Integer
    NumPoints = Examine_NumRows(rng_QueryTopLeft)
End Property

Public Property Let Caching(bln_Setting As Boolean)
    bln_Caching = bln_Setting
End Property

Public Property Get ZShiftsEnabled_DiscSpot() As Boolean
    ZShiftsEnabled_DiscSpot = bln_ZShiftsEnabled_DiscSpot
End Property

Public Property Let ZShiftsEnabled_DiscSpot(bln_Setting As Boolean)
    bln_ZShiftsEnabled_DiscSpot = bln_Setting
End Property

Public Property Get ConfigSheet() As ConfigSheet
    Set ConfigSheet = cfg_Settings
End Property


' ## METHODS - SETUP
Public Sub LoadRates()
    Dim str_SQLCode As String
    'Dim lng_BuildDate As Long: lng_BuildDate = cfg_Settings.CurrentBuildDate
    'Dim lng_DataDate As Long: lng_DataDate = cfg_Settings.CurrentDataDate
    Debug.Assert map_Rules.Dict_SourceTables.Exists("FXSPT")
    Dim str_TableName As String: str_TableName = map_Rules.Dict_SourceTables("FXSPT")

    ' Query
    str_SQLCode = "SELECT [Data Date], Currency, Quotation, Rate " _
                & "FROM " & str_TableName _
                & " WHERE [Data Date] = #" & Convert_SQLDate(cfg_Settings.CurrentDataDate) & "# " _
                & "ORDER BY Currency;"

    Me.ClearCurve
    Call Action_Query_Access(cfg_Settings.RatesDBPath, str_SQLCode, rng_QueryTopLeft)

    Me.AssignRanges  ' Update ranges as number of points may have changed
    rng_OrigSpots.NumberFormat = "General"
    rng_FinalSpots.NumberFormat = "General"
    Me.Scen_ApplyBase
End Sub

Public Sub AssignRanges()
    Dim int_NumPoints As Integer: int_NumPoints = Me.NumPoints  ' Requires QueryTopLeft to be known
    If int_NumPoints > 0 Then
        Set rng_CCYs = rng_QueryTopLeft.Offset(0, 1).Resize(int_NumPoints, 1)
        Set rng_Basis = rng_CCYs.Offset(0, 1)
        Set rng_OrigSpots = rng_Basis.Offset(0, 1)
        Set rng_RelShocks = rng_OrigSpots.Offset(0, 2)
        Set rng_AbsShocks = rng_RelShocks.Offset(0, 1)
        Set rng_FinalSpots = rng_AbsShocks.Offset(0, 2)
    End If
End Sub

Public Sub ClearCurve()
    Call Action_ClearBelow(rng_QueryTopLeft, 9)
End Sub

Public Sub ResetCache_Lookups()
    Call dic_Cache_DiscSpots.RemoveAll
    Call dic_Cache_USDPerFgn.RemoveAll
End Sub


' ## METHODS - LOOKUP
Public Function Lookup_Fwd(str_Fgn As String, str_Dom As String, Optional lng_Maturity As Long = 0, _
    Optional bln_UseSpotDelay As Boolean = True, Optional bln_GetOrig As Boolean = False) As Double
    ' ## Return forward to specified maturity in the quotation FGN/DOM
    ' ## Discount factors are calculated between spot dates

    ' Check trivial case
    If UCase(str_Fgn) = UCase(str_Dom) Then
        Lookup_Fwd = 1
        Exit Function
    End If

    Dim lng_ValDate As Long: lng_ValDate = cfg_Settings.CurrentValDate

    ' If requesting for quoted spot against USD, no need to perform any calculations
    If lng_Maturity = 0 Then
        If str_Fgn = "USD" Then
            Lookup_Fwd = 1 / Lookup_USDPerFgn(str_Dom)
            Exit Function
        ElseIf str_Dom = "USD" Then
            Lookup_Fwd = Lookup_USDPerFgn(str_Fgn)
            Exit Function
        End If
    End If

    ' For maturity date, use the spot date of the pair if maturity is not specified
    Dim lng_MappedMaturity As Long: lng_MappedMaturity = lng_Maturity
    If lng_Maturity = 0 Then lng_MappedMaturity = lng_ValDate

    ' Calculate maturity spot date of the pair
    Dim lng_FgnMatSpotDate As Long, lng_DomMatSpotDate As Long, lng_MatSpotDate As Long

    If bln_UseSpotDelay = True Then
        lng_MatSpotDate = cyGetFXCrossSpotDate(str_Fgn, str_Dom, lng_MappedMaturity, dic_GlobalStaticInfo)
    Else
        lng_MatSpotDate = lng_MappedMaturity
    End If

    ' Gather curves
    Dim irc_Fgn As Data_IRCurve, irc_Dom As Data_IRCurve
    Dim str_Curve_Fgn As String: str_Curve_Fgn = dic_FXCurveNames(str_Fgn)
    Dim str_Curve_Dom As String: str_Curve_Dom = dic_FXCurveNames(str_Dom)

    If dic_IRCurves Is Nothing Then
        Set irc_Fgn = GetObject_IRCurve(str_Curve_Fgn, True, False)
        Set irc_Dom = GetObject_IRCurve(str_Curve_Dom, True, False)
    Else
        Set irc_Fgn = dic_IRCurves(str_Curve_Fgn)
        Set irc_Dom = dic_IRCurves(str_Curve_Dom)
    End If

    ' Calculate forward to maturity
    Dim dbl_FgnMatSpotDF As Double, dbl_DomMatSpotDF As Double
    dbl_FgnMatSpotDF = irc_Fgn.Lookup_Rate(lng_ValDate, lng_MatSpotDate, "DF")
    dbl_DomMatSpotDF = irc_Dom.Lookup_Rate(lng_ValDate, lng_MatSpotDate, "DF")

    Lookup_Fwd = Me.Lookup_DiscSpot(str_Fgn, str_Dom, bln_GetOrig) * (dbl_FgnMatSpotDF / dbl_DomMatSpotDF)
End Function

Public Function Lookup_DiscSpot(str_Fgn As String, str_Dom As String, Optional bln_GetOrig As Boolean = False) As Double
    ' ## Returns discounted spot in the quotation FGN/DOM
    Dim str_Cache_Key As String
    If bln_Caching = True Then
        str_Cache_Key = str_Fgn & "|" & str_Dom & "|" & "|" & bln_GetOrig
    End If

    Dim bln_NeedEval As Boolean: bln_NeedEval = True
    Dim dbl_Output As Double

    If bln_Caching = True And dic_Cache_DiscSpots.Exists(str_Cache_Key) Then
        ' Use already known value if available
        dbl_Output = dic_Cache_DiscSpots(str_Cache_Key)
    ElseIf str_Fgn = str_Dom Then
        ' Trivial case
        dbl_Output = 1
        If bln_Caching = True Then Call dic_Cache_DiscSpots.Add(str_Cache_Key, dbl_Output)
    Else
        Dim lng_ValDate As Long: lng_ValDate = cfg_Settings.CurrentValDate
        Dim lng_FgnSpotDate As Long: lng_FgnSpotDate = cyGetFXSpotDate(str_Fgn, lng_ValDate, dic_GlobalStaticInfo)
        Dim lng_DomSpotDate As Long: lng_DomSpotDate = cyGetFXSpotDate(str_Dom, lng_ValDate, dic_GlobalStaticInfo)

        ' Gather curves
        Dim irc_Fgn As Data_IRCurve, irc_Dom As Data_IRCurve, irc_USD As Data_IRCurve
        Dim str_Curve_Fgn As String: str_Curve_Fgn = dic_FXCurveNames(str_Fgn)
        Dim str_Curve_Dom As String: str_Curve_Dom = dic_FXCurveNames(str_Dom)
        Dim str_Curve_USD As String: str_Curve_USD = dic_FXCurveNames("USD")

        If dic_IRCurves Is Nothing Then
            Set irc_Fgn = GetObject_IRCurve(str_Curve_Fgn, True, False)
            Set irc_Dom = GetObject_IRCurve(str_Curve_Dom, True, False)
            Set irc_USD = GetObject_IRCurve(str_Curve_USD, True, False)
        Else
            Set irc_Fgn = dic_IRCurves(str_Curve_Fgn)
            Set irc_Dom = dic_IRCurves(str_Curve_Dom)
            Set irc_USD = dic_IRCurves(str_Curve_USD)
        End If

        ' Remember original shifts, then set to zero if shifts are disabled for discounted spots
        Dim csh_OrigZShift_USD As CurveDaysShift, csh_OrigZShift_Fgn As CurveDaysShift, csh_OrigZShift_Dom As CurveDaysShift
        Dim enu_OrigState_USD As CurveState_IRC, enu_OrigState_Fgn As CurveState_IRC, enu_OrigState_Dom As CurveState_IRC, enu_ForcedState As CurveState_IRC
        Dim int_OrigSensPillar_USD As Integer, int_OrigSensPillar_Fgn As Integer, int_OrigSensPillar_Dom As Integer
        If bln_ZShiftsEnabled_DiscSpot = False Then
            enu_OrigState_USD = irc_USD.CurveState
            int_OrigSensPillar_USD = irc_USD.SensPillar
            enu_OrigState_Fgn = irc_Fgn.CurveState
            int_OrigSensPillar_Fgn = irc_Fgn.SensPillar
            enu_OrigState_Dom = irc_Dom.CurveState
            int_OrigSensPillar_Dom = irc_Dom.SensPillar

            If bln_GetOrig = True Then enu_ForcedState = CurveState_IRC.original Else enu_ForcedState = CurveState_IRC.Final
            Call irc_USD.SetCurveState(enu_ForcedState)
            Call irc_USD.SetCurveState(enu_ForcedState)
            Call irc_USD.SetCurveState(enu_ForcedState)
        End If

        ' Obtain discount factors
        Dim dbl_USDFgnSpotDF As Double, dbl_USDDomSpotDF As Double, dbl_FgnSpotDF As Double, dbl_DomSpotDF As Double
        dbl_USDFgnSpotDF = irc_USD.Lookup_Rate(lng_ValDate, lng_FgnSpotDate, "DF")
        dbl_USDDomSpotDF = irc_USD.Lookup_Rate(lng_ValDate, lng_DomSpotDate, "DF")
        dbl_FgnSpotDF = irc_Fgn.Lookup_Rate(lng_ValDate, lng_FgnSpotDate, "DF")
        dbl_DomSpotDF = irc_Dom.Lookup_Rate(lng_ValDate, lng_DomSpotDate, "DF")

        ' Reset to original shifts
        If bln_ZShiftsEnabled_DiscSpot = False Then
            Call irc_USD.SetCurveState(enu_OrigState_USD, int_OrigSensPillar_USD)
            Call irc_Fgn.SetCurveState(enu_OrigState_Fgn, int_OrigSensPillar_Fgn)
            Call irc_Dom.SetCurveState(enu_OrigState_Dom, int_OrigSensPillar_Dom)
        End If

        ' Obtain discounted spot for foreign and domestic currency vs USD
        Dim dbl_DiscFgnSpot As Double, dbl_DiscDomSpot As Double
        If str_Fgn = "USD" Then dbl_DiscFgnSpot = 1 Else dbl_DiscFgnSpot = Lookup_USDPerFgn(str_Fgn, bln_GetOrig) / (dbl_FgnSpotDF / dbl_USDFgnSpotDF)
        If str_Dom = "USD" Then dbl_DiscDomSpot = 1 Else dbl_DiscDomSpot = Lookup_USDPerFgn(str_Dom, bln_GetOrig) / (dbl_DomSpotDF / dbl_USDDomSpotDF)

        ' Obtain the cross spot
        dbl_Output = dbl_DiscFgnSpot / dbl_DiscDomSpot
        If bln_Caching = True Then Call dic_Cache_DiscSpots.Add(str_Cache_Key, dbl_Output)
    End If

    Lookup_DiscSpot = dbl_Output
End Function

Public Function Lookup_Spot(str_Fgn As String, str_Dom As String) As Double
    ' ## Return spot as at the spot date.  Performs cross spot calculation
    Lookup_Spot = Me.Lookup_Fwd(str_Fgn, str_Dom)
End Function

Public Function Lookup_Quotation(str_Currency As String) As String
    ' ## Returns DIRECT or INDIRECT, based on the quotation convention of the currency vs USD
    Dim int_FoundIndex As Integer: int_FoundIndex = WorksheetFunction.Match(str_Currency, rng_CCYs, 0)
    Lookup_Quotation = rng_Basis(int_FoundIndex, 1).Value
End Function

Public Function Lookup_NativeSpot(str_Currency As String, Optional bln_GetOrig As Boolean = False) As Double
    ' ## Returns spot based on the quotation convention of the currency vs USD
    Dim int_FoundIndex As Integer: int_FoundIndex = WorksheetFunction.Match(str_Currency, rng_CCYs, 0)
    Dim dbl_Spot As Double
    If bln_GetOrig = True Then
        dbl_Spot = rng_OrigSpots(int_FoundIndex, 1).Value
    Else
        dbl_Spot = rng_FinalSpots(int_FoundIndex, 1).Value
    End If

    Lookup_NativeSpot = dbl_Spot
End Function

Private Function Lookup_USDPerFgn(str_Fgn As String, Optional bln_GetOrig As Boolean = False) As Double
    ' ## Return spot in the quotation FGN/USD
    Dim dbl_Output As Double
    Dim str_Cache_Key As String: str_Cache_Key = str_Fgn & bln_GetOrig

    If bln_Caching = True And dic_Cache_USDPerFgn.Exists(str_Cache_Key) Then
        ' Use already known value if stored
        dbl_Output = dic_Cache_USDPerFgn(str_Cache_Key)
    Else
        Dim int_FoundIndex As Integer: int_FoundIndex = WorksheetFunction.Match(str_Fgn, rng_CCYs, 0)
        Dim str_Basis As String: str_Basis = rng_Basis(int_FoundIndex, 1).Value
        Dim dbl_Spot As Double
        If bln_GetOrig = True Then
            dbl_Spot = rng_OrigSpots(int_FoundIndex, 1).Value
        Else
            dbl_Spot = rng_FinalSpots(int_FoundIndex, 1).Value
        End If

        Select Case UCase(str_Basis)
            Case "DIRECT": dbl_Output = 1 / dbl_Spot
            Case "INDIRECT": dbl_Output = dbl_Spot
        End Select

        If bln_Caching = True Then Call dic_Cache_USDPerFgn.Add(str_Cache_Key, dbl_Output)
    End If

    Lookup_USDPerFgn = dbl_Output
End Function

Public Function Lookup_CurveDependencies(ParamArray strArr_Currencies() As Variant) As Dictionary
    Dim dic_output As New Dictionary: dic_output.CompareMode = CompareMethod.TextCompare

    ' Always dependent on USD which is the root currency for crosses
    Dim str_Curve_USD As String: str_Curve_USD = dic_FXCurveNames("USD")
    Call dic_output.Add(str_Curve_USD, True)

    ' Add FX curves for other currencies
    Dim str_Curve_ActiveCcy As String, int_ctr As Integer
    For int_ctr = LBound(strArr_Currencies) To UBound(strArr_Currencies)
        str_Curve_ActiveCcy = dic_FXCurveNames(strArr_Currencies(int_ctr))
        If dic_output.Exists(str_Curve_ActiveCcy) = False Then Call dic_output.Add(str_Curve_ActiveCcy, True)
    Next int_ctr

    Set Lookup_CurveDependencies = dic_output
End Function


' ## METHODS - SCENARIOS
Public Sub Scen_ApplyBase()
    ' ## Remove shifts and reset final spots to the original values
    rng_RelShocks.ClearContents
    rng_AbsShocks.ClearContents
    rng_FinalSpots.Value = rng_OrigSpots.Value
    Call Me.ResetCache_Lookups
End Sub

Public Sub Scen_AddShock(str_CCY As String, str_ShockType As String, dbl_AmountIndQuote As Double)
    ' ## Converts the indirect quotation shock specified to the basis as defined by the currency, then outputs to sheet
    Dim str_Basis As String, dbl_OrigSpot As Double, dbl_MappedAmount As Double

    If str_CCY = "USD" Then
        ' ## Special case: USD is present in every currency pair
        Dim int_RowCtr As Integer

        For int_RowCtr = 1 To rng_CCYs.count
            str_Basis = UCase(rng_Basis(int_RowCtr, 1).Value)
            Select Case UCase(str_ShockType)
                Case "ABS", "ABSOLUTE"
                    If str_Basis = "DIRECT" Then
                        dbl_MappedAmount = dbl_AmountIndQuote
                    Else
                        dbl_OrigSpot = rng_OrigSpots(int_RowCtr, 1).Value
                        dbl_MappedAmount = 1 / (1 / dbl_OrigSpot + dbl_AmountIndQuote) - dbl_OrigSpot
                    End If

                    rng_AbsShocks(int_RowCtr, 1).Value = dbl_MappedAmount
                Case "REL", "RELATIVE"
                    If str_Basis = "DIRECT" Then
                        dbl_MappedAmount = dbl_AmountIndQuote
                    Else
                        If dbl_AmountIndQuote = -1 Then
                            ' Murex seems to set spot to 1 if shock results in a zero spot rate
                            dbl_MappedAmount = (1 / rng_OrigSpots(int_RowCtr, 1).Value - 1) * 100
                        Else
                            dbl_MappedAmount = (1 / (1 + dbl_AmountIndQuote / 100) - 1) * 100
                        End If
                    End If

                    rng_RelShocks(int_RowCtr, 1).Value = dbl_MappedAmount
            End Select
        Next int_RowCtr
    Else
        Dim int_FoundIndex As Integer: int_FoundIndex = WorksheetFunction.Match(str_CCY, rng_CCYs, 0)

        str_Basis = UCase(rng_Basis(int_FoundIndex, 1).Value)
        Select Case str_ShockType
            Case "ABS", "ABSOLUTE"
                If str_Basis = "DIRECT" Then
                    dbl_OrigSpot = rng_OrigSpots(int_FoundIndex, 1).Value
                    dbl_MappedAmount = 1 / (1 / dbl_OrigSpot + dbl_AmountIndQuote) - dbl_OrigSpot
                Else
                    dbl_MappedAmount = dbl_AmountIndQuote
                End If

                rng_AbsShocks(int_FoundIndex, 1).Value = dbl_MappedAmount
            Case "REL", "RELATIVE"
                If str_Basis = "DIRECT" Then
                    If dbl_AmountIndQuote = -100 Then
                        ' Murex seems to set spot to 1 if shock results in a zero spot rate
                        dbl_MappedAmount = (1 / rng_OrigSpots(int_FoundIndex, 1).Value - 1) * 100
                    Else
                        dbl_MappedAmount = (1 / (1 + dbl_AmountIndQuote / 100) - 1) * 100
                    End If
                Else
                    dbl_MappedAmount = dbl_AmountIndQuote
                End If

                rng_RelShocks(int_FoundIndex, 1).Value = dbl_MappedAmount
        End Select
    End If
End Sub

Public Sub Scen_AddNativeShock(str_CCY As String, str_ShockType As String, dbl_Amount As Double)
    ' ## Adds shift based on the quotation convention
    If str_CCY <> "USD" Then
        Dim int_FoundIndex As Integer: int_FoundIndex = WorksheetFunction.Match(str_CCY, rng_CCYs, 0)
        Select Case str_ShockType
            Case "ABS", "ABSOLUTE": rng_AbsShocks(int_FoundIndex, 1).Value = dbl_Amount
            Case "REL", "RELATIVE": rng_RelShocks(int_FoundIndex, 1).Value = dbl_Amount
        End Select
    End If
End Sub

Public Sub Scen_ApplyCurrent()
    ' ## Calculates final spots taking into account the shocks applied
    Dim int_RowCtr As Integer
    For int_RowCtr = 1 To Me.NumPoints
        rng_FinalSpots(int_RowCtr, 1).Value = rng_OrigSpots(int_RowCtr, 1).Value * (1 + rng_RelShocks(int_RowCtr, 1).Value / 100) _
            + rng_AbsShocks(int_RowCtr, 1).Value
    Next int_RowCtr

    Call Me.ResetCache_Lookups
End Sub


' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Sub SetCurveState(str_curve As String, enu_State As CurveState_IRC, Optional int_PillarIndex As Integer = 0)
    ' ## For temporary zero shifts only, such as during a finite differencing calculation
    Dim irc_CurveToShift As Data_IRCurve: Set irc_CurveToShift = dic_IRCurves(str_curve)
    Call irc_CurveToShift.SetCurveState(enu_State, int_PillarIndex)

    If Not enu_State = CurveState_IRC.Final Then bln_Caching = False
End Sub

Public Sub FillDependency_IRC(dic_Curves As Dictionary)
    ' ## Store curve set, so it does not have to be created each time a value is looked up
    Set dic_IRCurves = dic_Curves
End Sub

Public Sub Scen_StoreOrigRate()
        Set rng_OrigRate = rng_OrigSpots.Offset(0, 1)
        Set rng_OrigFinalRate = rng_FinalSpots.Offset(0, 1)
        Set rng_tempRel = rng_RelShocks.Offset(0, 5)
        Set rng_tempAbs = rng_AbsShocks.Offset(0, 5)

        rng_tempRel = rng_RelShocks.Value
        rng_tempAbs = rng_AbsShocks.Value
        rng_RelShocks.ClearContents
        rng_AbsShocks.ClearContents
        rng_OrigRate = rng_OrigSpots.Value
        rng_OrigFinalRate = rng_FinalSpots.Value
End Sub

Public Sub Scen_TempRate()
    Dim int_RowCtr As Integer
    For int_RowCtr = 1 To Me.NumPoints
        rng_OrigSpots(int_RowCtr, 1) = rng_FinalSpots(int_RowCtr, 1).Value
    Next int_RowCtr

End Sub

Public Sub Scen_RestoreOrigRate()
        rng_OrigSpots.Value = rng_OrigRate.Value
        rng_FinalSpots.Value = rng_OrigFinalRate.Value
        rng_RelShocks = rng_tempRel.Value
        rng_AbsShocks = rng_tempAbs.Value

        rng_tempRel.ClearContents
        rng_tempAbs.ClearContents
        rng_OrigRate.ClearContents
        rng_OrigFinalRate.ClearContents
End Sub