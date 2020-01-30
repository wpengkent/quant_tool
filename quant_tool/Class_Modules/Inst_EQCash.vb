Option Explicit

' ## MEMBER DATA

' Components
Private scf_Purchase As SCF

' Dependent curves
Private fxs_Spots As Data_FXSpots, irc_SpotDisc As Data_IRCurve, eqd_Spots As Data_EQSpots

' Static values
Private dic_GlobalStaticInfo As Dictionary, dic_CurveDependencies As Dictionary
Private fld_Params As InstParams_ECS
Private str_PnLCcy As String, int_Sign As Integer
Private int_SpotDays As Integer, cal_Spot As Calendar
Private str_SecCode As String
Private lng_Quantity As Long
Private str_SecCcy As String
Private str_SpotDiscCurve As String
Private str_OutputType As String

' Variable dates
Private lng_ValDate As Long, lng_SpotDate As Long


' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As InstParams_ECS, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing)

    ' Store static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    lng_ValDate = fld_ParamsInput.ValueDate
    str_PnLCcy = fld_ParamsInput.CCY_PnL
    int_SpotDays = fld_ParamsInput.SpotDays
    Dim cas_Found As CalendarSet: Set cas_Found = dic_GlobalStaticInfo(StaticInfoType.CalendarSet)
    cal_Spot = cas_Found.Lookup_Calendar(fld_ParamsInput.SpotCal)
    Call FillSpotDate

    ' Dependent curves

    str_SpotDiscCurve = fld_ParamsInput.Curve_SpotDisc

    If dic_CurveSet Is Nothing Then
        Set irc_SpotDisc = GetObject_IRCurve(str_SpotDiscCurve, True, False)
        Set fxs_Spots = GetObject_FXSpots(True)
        Set eqd_Spots = GetObject_EQSpots(True)
    Else
        Set irc_SpotDisc = dic_CurveSet(CurveType.IRC)(str_SpotDiscCurve)
        Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
        Set eqd_Spots = dic_CurveSet(CurveType.EQSPT)
    End If

    ' Calculated values
    If fld_ParamsInput.IsBuy = True Then int_Sign = 1 Else int_Sign = -1

    ' Other inputs
    With fld_ParamsInput
        str_SecCode = .Security
        lng_Quantity = .Quantity
        str_SecCcy = .CCY_Sec
        str_OutputType = .OutputType
    End With

    Set scf_Purchase = New SCF
    Call scf_Purchase.Initialize(fld_ParamsInput.PurchaseCost, dic_CurveSet, dic_GlobalStaticInfo)
    scf_Purchase.ZShiftsEnabled_DF = GetSetting_IsPremInDV01()
    scf_Purchase.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

End Sub


' ## PROPERTIES - PUBLIC
Public Property Get marketvalue() As Double

    Select Case str_OutputType
        Case "PL"
            ' ## Get discounted value of future flows in the PnL currency
            marketvalue = EqSpot * lng_Quantity * SpotDF * FXConvFactor() * int_Sign
        Case "EQ NOP"
            ' ## EQ NOP = Market Value @ Spot Date
            marketvalue = EqSpot * lng_Quantity * FXConvFactor() * int_Sign
        Case "Rho"
            ' ## EQ NOP = Market Value @ Spot Date
            marketvalue = Rho
    End Select

End Property

Public Property Get Cash() As Double

    Select Case str_OutputType
        Case "PL"
            ' ## Get discounted value of the cost price in the PnL currency
            Cash = -scf_Purchase.CalcValue(lng_ValDate, lng_SpotDate, str_PnLCcy) * int_Sign
    End Select
End Property

Public Property Get PnL() As Double
    PnL = Me.marketvalue + Me.Cash
End Property
Private Property Get Rho() As Double

Dim dbl_OriMV As Double
Dim dbl_RhoMV As Double

'get pre-shock value
str_OutputType = "PL"
dbl_OriMV = marketvalue

Call irc_SpotDisc.SetCurveState(Zero_Up1BP)
dbl_RhoMV = marketvalue

Rho = dbl_RhoMV - dbl_OriMV
str_OutputType = "Rho"


Call irc_SpotDisc.SetCurveState(Final)

End Property

' ## PROPERTIES - PRIVATE
Private Property Get SpotDF() As Double
    ' Spot discount factor look up
    SpotDF = irc_SpotDisc.Lookup_Rate(lng_ValDate, lng_SpotDate, "DF", , , True)
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
        .Offset(int_ActiveRow, 0).Value = "Security:"
        .Offset(int_ActiveRow, 1).Value = str_SecCode
        Call dic_Addresses.Add("SecCode", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Spot Date:"
        .Offset(int_ActiveRow, 1).Value = lng_SpotDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("SpotDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Currency:"
        .Offset(int_ActiveRow, 1).Value = str_SecCcy
        Call dic_Addresses.Add("SecCcy", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 2
        .Offset(int_ActiveRow, 0).Value = "B/S Sign:"
        .Offset(int_ActiveRow, 1).Value = int_Sign
        Call dic_Addresses.Add("B/S Sign", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Quantity:"
        .Offset(int_ActiveRow, 1).Value = lng_Quantity
        Call dic_Addresses.Add("Quantity", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Market Price:"
        .Offset(int_ActiveRow, 1).Value = EqSpot
        Call dic_Addresses.Add("MktPrice", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "PnL Ccy:"
        .Offset(int_ActiveRow, 1).Value = str_PnLCcy
        Call dic_Addresses.Add("PnLCCY", .Offset(int_ActiveRow, 1).Address(False, False))


        'Output premium flow
        int_ActiveRow = int_ActiveRow + 3

        With rng_OutputTopLeft
            .Offset(int_ActiveRow, 0).Value = "PURCHASE COST"
            .Offset(int_ActiveRow, 0).Font.Italic = True

            int_ActiveRow = int_ActiveRow + 1
            Call scf_Purchase.OutputReport(.Offset(int_ActiveRow, 0), "Cash", _
                str_PnLCcy, -int_Sign, True, dic_Addresses, False)
        End With

        int_ActiveRow = wks_output.Range(dic_Addresses("SCF_PV")).Row

        int_ActiveRow = int_ActiveRow + 2
        .Offset(int_ActiveRow, 0).Value = "MV:"
        .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
        .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
        Set rng_MV = .Offset(int_ActiveRow, 1)

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "PnL:"
        .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
        .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
        Set rng_PnL = .Offset(int_ActiveRow, 1)

    End With

    ' Calculate values
    rng_MV.Formula = "=" & dic_Addresses("B/S Sign") & "*" & dic_Addresses("Quantity") & _
                    "*" & dic_Addresses("MktPrice") & "*cyreadIrCurve(" & _
                    """" & str_SpotDiscCurve & """" & "," & _
                    dic_Addresses("ValDate") & "," & dic_Addresses("SpotDate") & _
                    ",""DF"")" & "*cyGetFXDiscSpot(" & dic_Addresses("SecCcy") & "," & _
                    dic_Addresses("PnLCCY") & ")"

    'rng_Cash.Formula =
    rng_PnL.Formula = "=" & rng_MV.Address & "+" & dic_Addresses("SCF_PV")

    wks_output.Calculate
    wks_output.Cells.HorizontalAlignment = xlCenter
    wks_output.Columns.AutoFit
End Sub
