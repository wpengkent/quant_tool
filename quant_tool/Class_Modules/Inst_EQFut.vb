Option Explicit

' ## MEMBER DATA

' Dependent curves
Private fxs_Spots As Data_FXSpots, irc_Disc As Data_IRCurve, eqd_Spots As Data_EQSpots

' Static values
Private dic_GlobalStaticInfo As Dictionary, dic_CurveDependencies As Dictionary
Private fld_Params As InstParams_EQF
Private int_Sign As Integer, int_SpotDays As Integer
Private cal_Spot As Calendar
Private str_DiscCurve As String, str_SecCode As String, str_SecCcy As String, str_PnLCcy As String
Private str_DivCurve As String, str_OutputType As String, str_FutCode As String
Private lng_Units As Long, lng_Quantity As Long, lng_LotSize As Long
Private dbl_settlement As Double, dbl_FutSpread As Double, dbl_Div_Yield As Double
Private dbl_FutPrice As Double, lng_ContractPrice As Double, lng_MktPrice As Double

'Greeks
Private bln_CalRho As Boolean

'in % term
Private Const dbl_RhoShift As Double = 0.01

' Variable dates
Private lng_ValDate As Long, lng_OriValDate As Long, lng_ValDate_fixed As Long, lng_SpotDate As Long, lng_MatDate As Long, lng_MatSpotDate As Long

' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As InstParams_EQF, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing)

    ' Store static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    lng_ValDate = fld_ParamsInput.ValueDate
    lng_OriValDate = fld_ParamsInput.OriValueDate
    int_SpotDays = fld_ParamsInput.SpotDays
    lng_MatDate = fld_ParamsInput.MatDate

    Dim cas_Found As CalendarSet: Set cas_Found = dic_GlobalStaticInfo(StaticInfoType.CalendarSet)
    cal_Spot = cas_Found.Lookup_Calendar(fld_ParamsInput.SpotCal)

    Call FillSpotDate
    Call FillMatSpotDate

    ' Dependent curves

    str_DiscCurve = fld_ParamsInput.Curve_Disc
    str_DivCurve = fld_ParamsInput.Curve_Div


    If dic_CurveSet Is Nothing Then
        Set irc_Disc = GetObject_IRCurve(str_DiscCurve, True, False)
        Set fxs_Spots = GetObject_FXSpots(True)
        Set eqd_Spots = GetObject_EQSpots(True)
    Else
        Set irc_Disc = dic_CurveSet(CurveType.IRC)(str_DiscCurve)
        Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
        Set eqd_Spots = dic_CurveSet(CurveType.EQSPT)
    End If

    ' Calculated values
    If fld_ParamsInput.IsBuy = True Then int_Sign = 1 Else int_Sign = -1

    lng_Units = fld_ParamsInput.Quantity * fld_ParamsInput.LotSize
    dbl_settlement = lng_Units * fld_ParamsInput.Fut_ContractPrice * (-int_Sign)

    ' Other inputs
    With fld_ParamsInput
        str_FutCode = .Futures 'new
        str_SecCode = .Security
        str_SecCcy = .CCY_Sec
        str_PnLCcy = .CCY_PnL
        str_OutputType = .PLType
        lng_Quantity = .Quantity 'new
        lng_LotSize = .LotSize 'new
        lng_ContractPrice = .Fut_ContractPrice 'new
        lng_MktPrice = .Fut_MktPrice 'new
    End With

    If str_DivCurve = "" Then
        dbl_Div_Yield = fld_ParamsInput.Div_Yield
    Else
        dbl_Div_Yield = cyReadIRCurve(fld_ParamsInput.Curve_Div, lng_SpotDate, lng_MatSpotDate, "ZERO")
    End If

    dbl_FutPrice = fld_ParamsInput.Fut_MktPrice

    Call GetFutureSpread

End Sub

' ## PROPERTIES - PUBLIC
Public Property Get marketvalue(Optional dbl_spread As Double = 0) As Double

    Dim dbl_OriMV As Double
    Dim dbl_ShockMV As Double

    If bln_CalRho = False Then Call GetFutureSpread


    Select Case str_OutputType
        Case "PL"
            ' ## Get futures market value in the PnL currency
            marketvalue = (lng_Units * int_Sign * (EqSpot * CapFactor(dbl_spread) + dbl_FutSpread) + dbl_settlement) * FXConvFactor()
        Case "Delta Inv Amount"
            ' ## Delta Inv Amt for equity futures = PL Variation in MYR when Eq Spot move up 1 unit _
                @ valuation date / spot DF * Share Price
            ' ## Make sure setting Put equity future sensitivities on index is set to YES. It can be accessed through _
                Securities simulation settings >> Sensitivities


            dbl_OriMV = (lng_Units * int_Sign * (EqSpot * CapFactor(dbl_spread) + dbl_FutSpread) + dbl_settlement) * FXConvFactor()
            dbl_ShockMV = (lng_Units * int_Sign * ((EqSpot + 1) * CapFactor(dbl_spread) + dbl_FutSpread) + dbl_settlement) * FXConvFactor()

            marketvalue = (dbl_ShockMV - dbl_OriMV) / SpotDF * EqSpot

        Case "Rho"
            ' ## Rho for equity futures = MV ( rate curve shock up 1 bps point ) - Base MV
            marketvalue = Rho

        Case "EQ NOP"

            marketvalue = lng_Units * int_Sign * dbl_FutPrice * FXConvFactor()

    End Select

End Property

Public Property Get Cash() As Double
    Cash = 0
End Property

Public Property Get PnL() As Double
    PnL = Me.marketvalue
End Property

Public Property Get Rho() As Double

'## Rho for equity futures = MV ( rate curve shock up 1 bps point ) - Base MV
'## original code for PL variation for parallel upward curve shift by 1 Bps
'        str_OutputType = "PL"
'            Rho = MarketValue(dbl_RhoShift) - MarketValue
'        str_OutputType = "Rho"



'## code for PL variation for 1Bps upward shift of each pillar

Dim int_num As Integer: int_num = irc_Disc.NumPoints
Dim int_cnt As Integer

Dim dbl_OriMV As Double

str_OutputType = "PL"
    dbl_OriMV = marketvalue
str_OutputType = "Rho"

Dim dbl_Output As Double

bln_CalRho = True

For int_cnt = 1 To int_num
    Call irc_Disc.SetCurveState(Zero_Up1BP, int_cnt)
        str_OutputType = "PL"
            dbl_Output = dbl_Output + (marketvalue - dbl_OriMV)
        str_OutputType = "Rho"
    Call irc_Disc.SetCurveState(Final)
Next int_cnt

Rho = dbl_Output

End Property

' ## PROPERTIES - PRIVATE
Private Property Get CapFactor(Optional dbl_spread As Double = 0) As Double
    ' Capitalization Factor for forward projection
    CapFactor = 1 / irc_Disc.Lookup_Rate(lng_SpotDate, lng_MatSpotDate, "DF", , , True, , dbl_spread - dbl_Div_Yield)
End Property

Private Property Get SpotDF() As Double
    ' Spot discount factor look up
    SpotDF = irc_Disc.Lookup_Rate(lng_ValDate, lng_SpotDate, "DF", , , True)
End Property

Private Property Get FXConvFactor() As Double
    ' ## Discounted spot to translate into the PnL currency
    FXConvFactor = fxs_Spots.Lookup_DiscSpot(str_SecCcy, str_PnLCcy)
End Property

Private Property Get EqSpot() As Double
    ' Equity spot look up
    EqSpot = eqd_Spots.Lookup_Spot(str_SecCode)
End Property

' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Sub SetValDate(lng_Input As Long)
    ' ## Set stored value date and refresh values dependent on the value date
    lng_ValDate = lng_Input
    Call FillSpotDate
End Sub
' ## METHODS - PRIVATE
Private Sub FillSpotDate()
    ' ## Compute and store spot date
    If int_SpotDays = 0 Then
        lng_SpotDate = Date_ApplyBDC(lng_ValDate, "FOLL", cal_Spot.HolDates, cal_Spot.Weekends)
    Else
        lng_SpotDate = date_workday(lng_ValDate, int_SpotDays, cal_Spot.HolDates, cal_Spot.Weekends)
    End If
End Sub

Private Sub FillMatSpotDate()
    ' ## Compute and store maturity spot date
    If int_SpotDays = 0 Then
        lng_MatSpotDate = Date_ApplyBDC(lng_MatDate, "FOLL", cal_Spot.HolDates, cal_Spot.Weekends)
    Else
        lng_MatSpotDate = date_workday(lng_MatDate, int_SpotDays, cal_Spot.HolDates, cal_Spot.Weekends)
    End If
End Sub

Private Sub GetFutureSpread()
   'lng_ValDate_fixed = lng_ValDate
    irc_Disc.SetCurveState (original)
    'SetValDate (lng_OriValDate)
    dbl_FutSpread = dbl_FutPrice - eqd_Spots.Lookup_Spot(str_SecCode, True) _
                    / irc_Disc.Lookup_Rate(lng_SpotDate, lng_MatSpotDate, "DF", , , True, , -dbl_Div_Yield)
       irc_Disc.SetCurveState (Final)
    'SetValDate (lng_ValDate_fixed)

End Sub


' ## METHODS - CALCULATION DETAILS

Public Sub OutputReport(wks_output As Worksheet)
    wks_output.Cells.Clear

    Dim rng_OutputTopLeft As Range: Set rng_OutputTopLeft = wks_output.Range("A1")
    Dim str_Address_InitialPV As String, str_Address_FinalPV As String, str_Address_InitialType As String, str_Address_FinalType As String
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
        Call dic_Addresses.Add("SecCode", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Underlying:"
        .Offset(int_ActiveRow, 1).Value = str_SecCode
        Call dic_Addresses.Add("Underlying", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Maturity date:"
        .Offset(int_ActiveRow, 1).Value = lng_MatDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("MatDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Maturity Spot date:"
        .Offset(int_ActiveRow, 1).Value = lng_MatSpotDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("MatSpotDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Spot date:"
        .Offset(int_ActiveRow, 1).Value = lng_SpotDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("SpotDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Native Currency:"
        .Offset(int_ActiveRow, 1).Value = str_SecCcy
        Call dic_Addresses.Add("SecCcy", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "PnL Currency:"
        .Offset(int_ActiveRow, 1).Value = str_PnLCcy
        Call dic_Addresses.Add("PnLCcy", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Rate Curve:"
        .Offset(int_ActiveRow, 1).Value = str_DiscCurve
        Call dic_Addresses.Add("RateCurve", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 2
        .Offset(int_ActiveRow, 0).Value = "B/S Sign:"
        .Offset(int_ActiveRow, 1).Value = int_Sign
        Call dic_Addresses.Add("B/S Sign", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Quantity:"
        .Offset(int_ActiveRow, 1).Value = lng_Quantity
        Call dic_Addresses.Add("Quantity", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Lot Size:"
        .Offset(int_ActiveRow, 1).Value = lng_LotSize
        Call dic_Addresses.Add("LotSize", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Contract Price:"
        .Offset(int_ActiveRow, 1).Value = lng_ContractPrice
        Call dic_Addresses.Add("ContractPrice", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Market Price:"
        .Offset(int_ActiveRow, 1).Value = lng_MktPrice

       int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Underlying Price:"
        .Offset(int_ActiveRow, 1).Value = EqSpot
        Call dic_Addresses.Add("UndPrice", .Offset(int_ActiveRow, 1).Address(False, False))

        'PnL calculation
        int_ActiveRow = int_ActiveRow + 3

        .Offset(int_ActiveRow, 0).Value = "PnL Calculation:"
        .Offset(int_ActiveRow, 0).Font.Italic = True

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Cap Factor:"
        .Offset(int_ActiveRow, 1).Value = CapFactor()
        Call dic_Addresses.Add("PnL_Cap", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Model Price:"
        .Offset(int_ActiveRow, 1).Value = "=" & dic_Addresses("UndPrice") & "*" & dic_Addresses("PnL_Cap")
        Call dic_Addresses.Add("ModelPrice", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Future Spread:"
        .Offset(int_ActiveRow, 1).Value = dbl_FutSpread
        Call dic_Addresses.Add("FutSpread", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Market Price:"
        .Offset(int_ActiveRow, 1).Value = "=" & dic_Addresses("ModelPrice") & "+" & dic_Addresses("FutSpread")
        Call dic_Addresses.Add("MktPrice", .Offset(int_ActiveRow, 1).Address(False, False))

         int_ActiveRow = int_ActiveRow + 1
         .Offset(int_ActiveRow, 0).Value = "PnL (" & str_SecCcy & "):"
         .Offset(int_ActiveRow, 1).Value = "=" & dic_Addresses("B/S Sign") & _
            " *(" & dic_Addresses("MktPrice") & "-" & dic_Addresses("ContractPrice") & ")*" & _
            dic_Addresses("Quantity") & "*" & dic_Addresses("LotSize")
        Call dic_Addresses.Add("PnL_native", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Discounted Spot:"
        .Offset(int_ActiveRow, 1).Value = "=cyGetFXDiscSpot(" & dic_Addresses("SecCcy") & _
            " ," & dic_Addresses("PnLCcy") & ")"
        Call dic_Addresses.Add("DiscFXSpot", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "PnL (" & str_PnLCcy & "):"
        .Offset(int_ActiveRow, 1).Value = "=" & dic_Addresses("PnL_native") & _
            " * " & dic_Addresses("DiscFXSpot")
        .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20

        'Rho Calculation
        int_ActiveRow = int_ActiveRow + 3
        .Offset(int_ActiveRow, 0).Value = "Rho Calculation:"
        .Offset(int_ActiveRow, 0).Font.Italic = True

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Dividend Yield:"
        .Offset(int_ActiveRow, 1).Value = dbl_Div_Yield

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Cap Factor (Before Shock):"
        .Offset(int_ActiveRow, 1).Value = CapFactor()
        Call dic_Addresses.Add("Cap_before_shock", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Cap Factor (After Shock):"
        .Offset(int_ActiveRow, 1).Value = CapFactor(dbl_RhoShift)
        Call dic_Addresses.Add("Cap_after_shock", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Rho (" & str_SecCcy & "):"
        .Offset(int_ActiveRow, 1).Value = "=(" & dic_Addresses("Cap_after_shock") & "-" & _
            dic_Addresses("Cap_before_shock") & ")*" & dic_Addresses("UndPrice") & "*" & _
            dic_Addresses("Quantity") & "*" & dic_Addresses("LotSize") & "*" & dic_Addresses("B/S Sign")
        Call dic_Addresses.Add("rho_native", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Rho (" & str_PnLCcy & "):"
        .Offset(int_ActiveRow, 1).Value = "=" & dic_Addresses("rho_native") & _
            " * " & dic_Addresses("DiscFXSpot")
        .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20

    End With

    wks_output.Calculate
    wks_output.Cells.HorizontalAlignment = xlCenter
    wks_output.Columns.AutoFit
End Sub