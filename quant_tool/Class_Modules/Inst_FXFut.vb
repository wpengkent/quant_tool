Option Explicit
' ## MEMBER DATA
' Dependent curves
Private fxs_Spots As Data_FXSpots, irc_CcyA As Data_IRCurve, irc_CcyB As Data_IRCurve

' Variable dates
Private lng_ValDate As Long, lng_SpotDate As Long

' Static values
Private dic_GlobalStaticInfo As Dictionary, dic_CurveDependencies As Dictionary, map_Rules As MappingRules
Private fld_Params As InstParams_FXFut
Private str_CCY_FlowA As String, str_CCY_FlowB As String
Private str_CCY_PnL As String, int_Sign As Integer
Private str_MatDate As Long, str_FutCode As String, str_Und As String, str_LotSizeCCY As String
Private dbl_Amount As Double, dbl_FutPrice As Double, dbl_ContractPrice As Double, dbl_FxFwdBase As Double, dbl_FXFutSpread As Double, dbl_FxSpotBase As Double
Private dbl_LotSize As Double, dbl_Quantity As Double, dbl_TotalAmount As Double

Private Const bln_IsSpotDFInDV01 As Boolean = True


' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As InstParams_FXFut, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing)

    ' Store static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    Set map_Rules = dic_GlobalStaticInfo(StaticInfoType.MappingRules)
    str_CCY_PnL = fld_ParamsInput.CCY_PnL
    If fld_ParamsInput.IsBuy = True Then int_Sign = 1 Else int_Sign = -1

    fld_Params = fld_ParamsInput
    str_CCY_FlowA = fld_ParamsInput.FlowA.CCY
    str_CCY_FlowB = fld_ParamsInput.FlowB.CCY
    str_CCY_PnL = fld_ParamsInput.CCY_PnL

    Call SetValDate(fld_ParamsInput.ValueDate)

     If dic_CurveSet Is Nothing Then
        Set fxs_Spots = GetObject_FXSpots(True)
        Set irc_CcyA = GetObject_IRCurve(fld_ParamsInput.FlowA.Curve_Disc, True, False)
        Set irc_CcyB = GetObject_IRCurve(fld_ParamsInput.FlowB.Curve_Disc, True, False)
    Else
        Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
        Set irc_CcyA = dic_CurveSet(CurveType.IRC)(fld_ParamsInput.FlowA.Curve_Disc)
        Set irc_CcyB = dic_CurveSet(CurveType.IRC)(fld_ParamsInput.FlowB.Curve_Disc)
    End If

    'Get Spread
    '##Temporary set Market Data to Original values
    fxs_Spots.Scen_StoreOrigRate
    fxs_Spots.Scen_ApplyBase
    irc_CcyA.SetCurveState (original)
    irc_CcyB.SetCurveState (original)
    dbl_FxFwdBase = Forward()
    dbl_FxSpotBase = Spot()
    '##Set Market Data back to Final Values
    fxs_Spots.Scen_RestoreOrigRate
    fxs_Spots.Scen_ApplyCurrent
    irc_CcyA.SetCurveState (Final)
    irc_CcyB.SetCurveState (Final)

     ' Determine curve dependencies
    Set dic_CurveDependencies = fxs_Spots.Lookup_CurveDependencies(str_CCY_FlowA, str_CCY_FlowB, str_CCY_PnL)

    With fld_ParamsInput
        str_MatDate = .MatDate
        str_FutCode = .Futures
        str_Und = .Underlying
        dbl_LotSize = .LotSize
        str_LotSizeCCY = .LotSizeCCY
        dbl_Quantity = .Quantity
        dbl_FutPrice = .Fut_MktPrice
        dbl_ContractPrice = .Fut_ContractPrice
    End With


    dbl_FXFutSpread = Forward() - dbl_FxFwdBase
    dbl_TotalAmount = dbl_Quantity * dbl_LotSize
    Debug.Print dbl_FXFutSpread
End Sub

' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Sub SetValDate(lng_Input As Long)
    ' Set stored value date
    lng_ValDate = lng_Input
    lng_SpotDate = cyGetFXCrossSpotDate(str_CCY_FlowA, str_CCY_FlowB, lng_ValDate, dic_GlobalStaticInfo)
End Sub

Private Sub SetCurveState(str_curve As String, enu_State As CurveState_IRC, Optional int_PillarIndex As Integer = 0)
    ' ## Set up shift in the market data and underlying components
    If irc_CcyA.CurveName = str_curve Then Call irc_CcyA.SetCurveState(enu_State, int_PillarIndex)
    If irc_CcyB.CurveName = str_curve Then Call irc_CcyB.SetCurveState(enu_State, int_PillarIndex)
    Call fxs_Spots.SetCurveState(str_curve, enu_State, int_PillarIndex)
End Sub

' ## PROPERTIES - PUBLIC
Public Property Get marketvalue(Optional dbl_spread As Double = 0) As Double
    Dim dbl_TheoFutPrice As Double
    dbl_TheoFutPrice = dbl_FutPrice + dbl_FXFutSpread
    marketvalue = int_Sign * dbl_TotalAmount * (dbl_TheoFutPrice - dbl_ContractPrice) * cyGetFXDiscSpot(str_CCY_FlowB, str_CCY_PnL)
End Property

Public Property Get Cash() As Double
    Cash = 0
End Property

Public Property Get PnL() As Double
    PnL = Me.marketvalue
End Property

' ## PROPERTIES - PRIVATE
Private Property Get Spot() As Double
    Spot = fxs_Spots.Lookup_Spot(fld_Params.FlowA.CCY, fld_Params.FlowB.CCY)
End Property

Private Property Get Forward() As Double
    Forward = fxs_Spots.Lookup_Fwd(fld_Params.FlowA.CCY, fld_Params.FlowB.CCY, fld_Params.MatDate, False)
End Property


' ## METHODS - GREEKS
Public Function Calc_Delta() As Double
    Dim dbl_Output As Double
    Dim dbl_SpotUp As Double, dbl_SpotDown As Double, dbl_FxFwdBase As Double, dbl_FxFwdShockUp As Double, dbl_FxFwdShockDown As Double
    Dim dbl_SpreadShockUp As Double, dbl_SpreadShockDown As Double
    Dim dbl_ShockSize As Double: dbl_ShockSize = 1
    Dim dbl_TheoFutPriceShockUp As Double, dbl_TheoFutPriceShockDown As Double, dbl_PnLShockUp As Double, dbl_PnLShockDown As Double
    Dim dbl_DF_MatSpotCcyA As Double: dbl_DF_MatSpotCcyA = irc_CcyA.Lookup_Rate(lng_ValDate, lng_SpotDate, "DF", , , True)
    Dim dbl_DF_MatSpotCcyB As Double: dbl_DF_MatSpotCcyB = irc_CcyB.Lookup_Rate(lng_ValDate, lng_SpotDate, "DF", , , True)

    Dim str_TargetCCy As String
    'str_TargetCCy = fld_Params.FlowA.CCY
      If fld_Params.FlowA.CCY = "USD" Then
        str_TargetCCy = fld_Params.FlowB.CCY
    ElseIf fld_Params.FlowB.CCY = "USD" Then
        str_TargetCCy = fld_Params.FlowA.CCY
    ElseIf (fxs_Spots.Lookup_Quotation(fld_Params.FlowB.CCY) = "INDIRECT" And fxs_Spots.Lookup_Quotation(fld_Params.FlowA.CCY) = "INDIRECT") Then
        str_TargetCCy = fld_Params.FlowA.CCY
    Else
        str_TargetCCy = fld_Params.FlowB.CCY
    End If

    ' Shock up
    Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", dbl_ShockSize)
    Call fxs_Spots.Scen_ApplyCurrent
    dbl_SpotUp = Spot()
    dbl_FxFwdShockUp = Forward()
    dbl_SpreadShockUp = dbl_FxFwdShockUp - dbl_FxFwdBase
    dbl_PnLShockUp = int_Sign * dbl_TotalAmount * ((dbl_FutPrice + dbl_SpreadShockUp) - dbl_ContractPrice) / dbl_DF_MatSpotCcyB
    Call fxs_Spots.Scen_ApplyBase

    ' Shock down
    Call fxs_Spots.Scen_AddNativeShock(str_TargetCCy, "REL", -dbl_ShockSize)
    Call fxs_Spots.Scen_ApplyCurrent
    dbl_SpotDown = Spot()
    dbl_FxFwdShockDown = Forward()
    dbl_SpreadShockDown = dbl_FxFwdShockDown - dbl_FxFwdBase
    dbl_PnLShockDown = int_Sign * dbl_TotalAmount * ((dbl_FutPrice + dbl_SpreadShockDown) - dbl_ContractPrice) / dbl_DF_MatSpotCcyB
    Call fxs_Spots.Scen_ApplyBase

    dbl_Output = (dbl_PnLShockUp - dbl_PnLShockDown) / (dbl_SpotUp - dbl_SpotDown)
    Calc_Delta = dbl_Output
End Function

Public Function Calc_DV01(str_curve As String, Optional int_PillarIndex As Integer = 0) As Double
    Dim dbl_FxFwdShockUp As Double, dbl_FxFwdShockDown As Double, dbl_rd As Double, dbl_rf As Double
    Dim dbl_FutTheoPriceShockUp As Double, dbl_FutTheoPriceShockDown As Double
    Dim dbl_PnLShockUp As Double, dbl_PnLShockDown As Double
    Dim dbl_FxFutSpreadShockUp As Double, dbl_FxFutSpreadShockDown As Double
    Dim dbl_Output As Double
    If dic_CurveDependencies.Exists(str_curve) Then
        ' Remember original setting, then disable DV01 impact on discounted spot
        Dim bln_ZShiftsEnabled_DiscSpot As Boolean: bln_ZShiftsEnabled_DiscSpot = fxs_Spots.ZShiftsEnabled_DiscSpot
        fxs_Spots.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

        Call SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        'dbl_FxFwdShockUp = Forward()

        dbl_rd = irc_CcyB.Lookup_Rate(lng_SpotDate, fld_Params.MatDate, "DF", , , True)
        dbl_rf = irc_CcyA.Lookup_Rate(lng_SpotDate, fld_Params.MatDate, "DF", , , True)
        dbl_FxFwdShockUp = dbl_FxSpotBase * dbl_rf / dbl_rd
        dbl_FxFutSpreadShockUp = dbl_FxFwdShockUp / dbl_FxFwdBase
        dbl_FutTheoPriceShockUp = dbl_FutPrice * dbl_FxFutSpreadShockUp
        dbl_PnLShockUp = int_Sign * dbl_TotalAmount * (dbl_FutTheoPriceShockUp - dbl_ContractPrice)

        Call SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        'dbl_FxFwdShockDown = Forward()
        dbl_rd = irc_CcyB.Lookup_Rate(lng_SpotDate, fld_Params.MatDate, "DF", , , True)
        dbl_rf = irc_CcyA.Lookup_Rate(lng_SpotDate, fld_Params.MatDate, "DF", , , True)
        dbl_FxFwdShockDown = dbl_FxSpotBase * dbl_rf / dbl_rd
        dbl_FxFutSpreadShockDown = dbl_FxFwdShockDown / dbl_FxFwdBase
        dbl_FutTheoPriceShockDown = dbl_FutPrice * dbl_FxFutSpreadShockDown
        dbl_PnLShockDown = int_Sign * dbl_TotalAmount * (dbl_FutTheoPriceShockDown - dbl_ContractPrice)


        ' Revert to original setting
        fxs_Spots.ZShiftsEnabled_DiscSpot = bln_ZShiftsEnabled_DiscSpot
        Call SetCurveState(str_curve, CurveState_IRC.Final, int_PillarIndex)
        dbl_Output = (dbl_PnLShockUp - dbl_PnLShockDown) * cyGetFXDiscSpot(str_CCY_FlowB, str_CCY_PnL) / 2

    Else
        dbl_Output = 0
    End If

Calc_DV01 = dbl_Output
End Function

' ## METHODS - CALCULATION DETAILS
Public Sub OutputReport(wks_output As Worksheet)
    wks_output.Cells.Clear

    Dim rng_OutputTopLeft As Range: Set rng_OutputTopLeft = wks_output.Range("A1")
    Dim rng_PnL As Range, rng_MV As Range, rng_Cash As Range
    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    Dim str_CurrencyFormat As String: str_CurrencyFormat = Gather_CurrencyFormat()
    Dim int_ActiveRow As Integer: int_ActiveRow = 0

    Dim dic_Addresses As Dictionary: Set dic_Addresses = New Dictionary
    dic_Addresses.CompareMode = CompareMethod.TextCompare
    ' Output overall info

     With rng_OutputTopLeft
        .Offset(int_ActiveRow, 0).Value = "Value date:"
        .Offset(int_ActiveRow, 1).Value = lng_ValDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("ValDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 2
        .Offset(int_ActiveRow, 0).Value = "Futures Contract:"
        .Offset(int_ActiveRow, 1).Value = str_FutCode
        Call dic_Addresses.Add("FutCode", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Underlying:"
        .Offset(int_ActiveRow, 1).Value = str_Und
        Call dic_Addresses.Add("Underlying", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Maturity date:"
        .Offset(int_ActiveRow, 1).Value = str_MatDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("MatDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Spot date:"
        .Offset(int_ActiveRow, 1).Value = lng_SpotDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("SpotDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 2
        .Offset(int_ActiveRow, 0).Value = "B/S Sign:"
        .Offset(int_ActiveRow, 1).Value = int_Sign
        Call dic_Addresses.Add("B/S Sign", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Quantity:"
        .Offset(int_ActiveRow, 1).Value = dbl_Quantity
        Call dic_Addresses.Add("Quantity", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Lot Size:"
        .Offset(int_ActiveRow, 1).Value = dbl_LotSize
        Call dic_Addresses.Add("LotSize", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Lot Size Currency:"
        .Offset(int_ActiveRow, 1).Value = str_LotSizeCCY
        Call dic_Addresses.Add("LotSizeCCY", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Contract Price:"
        .Offset(int_ActiveRow, 1).Value = dbl_ContractPrice
        Call dic_Addresses.Add("ContractPrice", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Market Price:"
        .Offset(int_ActiveRow, 1).Value = dbl_FutPrice + dbl_FXFutSpread
        Call dic_Addresses.Add("MarketPrice", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "PnL Currency:"
        .Offset(int_ActiveRow, 1).Value = str_CCY_PnL
        Call dic_Addresses.Add("PnLCcy", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Fgn Currency:"
        .Offset(int_ActiveRow, 1).Value = str_CCY_FlowA
        Call dic_Addresses.Add("FgnCcy", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Dom Currency:"
        .Offset(int_ActiveRow, 1).Value = str_CCY_FlowB
        Call dic_Addresses.Add("DomCcy", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Market Value:"
        .Offset(int_ActiveRow, 1).Value = "=" & dic_Addresses("B/S Sign") & "*" & dic_Addresses("Quantity") & "*" & dic_Addresses("LotSize") & _
            "*(" & dic_Addresses("MarketPrice") & "-" & dic_Addresses("ContractPrice") & ")"
        Call dic_Addresses.Add("MarketValue", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Cash:"
        .Offset(int_ActiveRow, 1).Value = 0
        Call dic_Addresses.Add("Cash", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "PnL:"
        .Offset(int_ActiveRow, 1).Value = "=(" & dic_Addresses("MarketValue") & "+" & dic_Addresses("Cash") & ")" & "*cyGetFXDiscSpot(" & dic_Addresses("DomCcy") & "," & dic_Addresses("PnLCcy") & ")"
        .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
        Call dic_Addresses.Add("PnL", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 2
        .Offset(int_ActiveRow, 0).Value = "Fx Future Spread:"
        .Offset(int_ActiveRow, 1).Value = dbl_FXFutSpread
        .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
        Call dic_Addresses.Add("FxFutSpread", .Offset(int_ActiveRow, 1).Address(False, False))


     End With

    wks_output.Calculate
    wks_output.Cells.HorizontalAlignment = xlCenter
    wks_output.Columns.AutoFit
End Sub