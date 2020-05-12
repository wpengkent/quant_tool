Option Explicit

' ## MEMBER DATA

' Components
Private scf_Purchase As SCF

' Dependent curves
Private fxs_Spots As Data_FXSpots, irc_SpotDisc As Data_IRCurve, irc_MarketDisc As Data_IRCurve, _
eqd_Spots As Data_EQSpots, eqd_Vols As Data_EQVols, eqd_Smile As Data_EQSmile

' Static values
Private dic_GlobalStaticInfo As Dictionary, dic_CurveDependencies As Dictionary
Private fld_Params As InstParams_ECS
Private str_PnLCcy As String, int_Sign As Integer
Private int_SpotDays As Integer, cal_Spot As Calendar
Private str_SecCode As String
Private lng_Quantity As Long
Private str_SecCcy As String
Private str_SpotDiscCurve As String
Private str_MarketDiscCurve As String
Private str_ExerciseType As String
Private str_Settlement As String
Private dbl_DivAmount As Double, str_DivType As String
Private dbl_VolSpread As Double
Private dbl_Strike As Double
Private bln_IsCall As Boolean
Private dbl_Parity As Double
Private str_VolCode As String
Private str_PLType As String
Private int_RhoDays As Integer
Private bln_IsFutures As Boolean 'new
Private str_FuturesContract As String 'new
Private lng_LotSize As Long 'new
Private bln_IsSmile As Boolean
'Private lng_PaymentSpot As Long
Private dbl_underlying As Double
Private int_OptionSpot As Integer
Private int_DivSpot As Integer
Private var_smile_type As Variant


' To be included in Output Info/ BS Price
Private dbl_Forward As Double
Private dbl_Spot As Double
Private dbl_FXConvFactor As Double
Private dbl_Vol As Double
Private dbl_Vol_factor As Double
Private dbl_ForwardPremium As Double
Private d1 As Double, d2 As Double
Private dbl_FutPrice As Double 'new
Private dbl_r As Double
Private dbl_r_Und As Double




' Variable dates
Private lng_ValDate As Long, lng_SpotDate As Long
Private lng_MatDate As Long, lng_DelivDate As Long
Private lng_SettlementDate As Long ' Dividend Date & Cash Settlement Date
Private lng_DivPaymentDate As Long, lng_DivExDate As Long
Private lng_RhoDate As Long
Private lng_FutMatDate As Long 'new
Private lng_FutMatSpotDate As Long 'new
Private lng_OriValDate As Long 'new
Private lng_ValDate_fixed As Long 'new
Private lng_DivSpotDate As Long
Private lng_DivExpiryDate As Long
Private lng_PaymentSpotDate As Long

'DF of Dates
Private dbl_DFValSpotToVal As Double
Private dbl_DFValSpotToVal_market As Double
Private dbl_DFMatSpotToValSpot As Double
Private dbl_DFSettleToValSpot As Double
Private dbl_DFDivPaymentToValSpot As Double
Private dbl_DFMatSpotToDivPayment As Double
Private dbl_DFRhoToVal As Double
Private arr_Binomial As Variant
Private dbl_DivDF As Double
Private dbl_DFMatToVal As Double
Private dbl_DFMatSpotToVal As Double
Private dbl_DFSettleToVal As Double
Private dbl_DFExDivToDivPayment As Double
Private dbl_DFValToPaymentSpot As Double 'new

'Other variable used for computation
Private dbl_Time_MatToVal As Double 'Te
Private dbl_Time_MatSpotToValSpot As Double 'Td
Private dbl_Time_ExDivToValSpot As Double
Private int_optdirection As Integer

'other info used for output result
Private str_description As String

'Variables used for binomial tree and output result
Private int_n As Integer, int_DivPeriod As Integer
Private dbl_bin_u As Double, dbl_bin_d As Double, dbl_bin_p As Double, dbl_bin_q As Double
Private dbl_dt As Double, dbl_DF_dt As Double, dbl_PVDiv As Double

' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As InstParams_EQO, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing)

    ' Store static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    lng_ValDate = fld_ParamsInput.ValueDate
    lng_MatDate = fld_ParamsInput.MatDate
    lng_DelivDate = fld_ParamsInput.DelivDate
    lng_DivPaymentDate = fld_ParamsInput.DivPayment_Date
    lng_DivExDate = fld_ParamsInput.DivEx_Date
    lng_SettlementDate = fld_ParamsInput.SettlementDate
    str_PnLCcy = fld_ParamsInput.CCY_PnL
    int_SpotDays = fld_ParamsInput.SpotDays
    int_OptionSpot = fld_ParamsInput.OptionSpot
    int_DivSpot = fld_ParamsInput.DividendSpot
    'lng_PaymentSpot = fld_ParamsInput.OptionSpot
    str_PLType = fld_ParamsInput.PLType
    var_smile_type = fld_ParamsInput.VolSpread  '****Alvin 28/06/2018****
    Dim cas_Found As CalendarSet: Set cas_Found = dic_GlobalStaticInfo(StaticInfoType.CalendarSet)
    cal_Spot = cas_Found.Lookup_Calendar(fld_ParamsInput.SpotCal)
    Call FillSpotDate
    Call FillSettlementDate

    '## Check for smile
    If IsNumeric(fld_ParamsInput.VolSpread) = False Then
        bln_IsSmile = True
    End If
    ' Dependent curves

    str_SpotDiscCurve = fld_ParamsInput.Curve_SpotDisc
    str_MarketDiscCurve = fld_ParamsInput.Curve_MarketDisc

    If dic_CurveSet Is Nothing Then
        Set irc_SpotDisc = GetObject_IRCurve(str_SpotDiscCurve, True, False)
        Set fxs_Spots = GetObject_FXSpots(True)
        Set eqd_Spots = GetObject_EQSpots(True)
        Set eqd_Vols = GetObject_EQVols(True)
        If str_MarketDiscCurve <> "" Then _
        Set irc_MarketDisc = GetObject_IRCurve(str_MarketDiscCurve, True, False)

        '# Matt Edit
        If bln_IsSmile = True Then
            Set eqd_Smile = GetObject_EQSmile(fld_ParamsInput.VolCode, True, False)
        End If


    Else
        Set irc_SpotDisc = dic_CurveSet(CurveType.IRC)(str_SpotDiscCurve)
        Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
        Set eqd_Spots = dic_CurveSet(CurveType.EQSPT)
        Set eqd_Vols = dic_CurveSet(CurveType.EQVOL)
        If str_MarketDiscCurve <> "" Then _
        Set irc_MarketDisc = dic_CurveSet(CurveType.IRC)(str_MarketDiscCurve)

        '# Matt Edit
        If bln_IsSmile = True Then
            Set eqd_Smile = dic_CurveSet(CurveType.EVL)(fld_ParamsInput.VolCode)
        End If

    End If

    ' Calculated values
    If fld_ParamsInput.IsBuy = True Then int_Sign = 1 Else int_Sign = -1

    ' Other inputs
    With fld_ParamsInput
        str_SecCode = .Security
        lng_Quantity = .Quantity
        str_SecCcy = .CCY_Sec
        str_ExerciseType = .ExerciseType
        str_Settlement = .Settlement
        dbl_DivAmount = .Div_Amount
        dbl_Strike = .strike
        bln_IsCall = .IsCall
        str_DivType = .Div_Type
        dbl_Parity = .Parity
        str_VolCode = .VolCode
        str_description = .Description
        bln_IsFutures = .IsFutures
        str_FuturesContract = .FuturesContract
        lng_LotSize = .LotSize
        lng_OriValDate = .OriValueDate

        If bln_IsSmile = True Then
            dbl_VolSpread = 0
        Else
            dbl_VolSpread = .VolSpread
        End If

    End With

    Set scf_Purchase = New SCF
    Call scf_Purchase.Initialize(fld_ParamsInput.PurchaseCost, dic_CurveSet, dic_GlobalStaticInfo)
    scf_Purchase.ZShiftsEnabled_DF = GetSetting_IsPremInDV01()
    scf_Purchase.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

    'For futures option
    If bln_IsFutures = True Then
        'only cover European case for futures option
        lng_FutMatDate = fld_ParamsInput.FutMat_Date 'new
        str_ExerciseType = "EUROPEAN" 'new
        Call FillFutMatSpotDate
    End If

End Sub

Private Sub FillMktData()
'Alv: Updated 4/7/2018
    'Other variables
    'If lng_MatDate > lng_ValDate Then
        dbl_Time_MatToVal = (lng_MatDate - lng_ValDate) / 365   'Te
        dbl_Time_MatSpotToValSpot = (lng_DelivDate - lng_SpotDate) / 365    'Td
        dbl_Time_ExDivToValSpot = (lng_DivExDate - lng_SpotDate) / 365
    'End If

    If bln_IsCall = True Then int_optdirection = 1 Else int_optdirection = -1


    ' Storing Additional Static Values : Spot Price , Curve DF , Vega
    dbl_Spot = eqd_Spots.Lookup_Spot(str_SecCode)
    'dbl_Vol = eqd_Vols.Lookup_Vol(str_VolCode) + dbl_VolSpread

    If bln_IsSmile = True Then
'**********Alvin add str_smile_boundary on  28/06/2018 for Flat Smile Extrp********
        dbl_Vol = eqd_Smile.Lookup_Vol(dbl_Strike, lng_MatDate, var_smile_type)
'**********Alvin add str_smile_boundary on  28/06/2018 for Flat Smile Extrp********

    Else
        dbl_Vol = eqd_Vols.Lookup_Vol(str_VolCode) + dbl_VolSpread
    End If

    If bln_IsFutures = True Then dbl_FutPrice = eqd_Spots.Lookup_Spot(str_FuturesContract)

    ' DF
    If lng_MatDate > lng_ValDate Then
        dbl_DFValSpotToVal = irc_SpotDisc.Lookup_Rate(lng_ValDate, lng_SpotDate, "DF", , , True)
        dbl_DFMatSpotToValSpot = irc_SpotDisc.Lookup_Rate(lng_SpotDate, lng_DelivDate, "DF", , , True)
        dbl_DFMatToVal = irc_SpotDisc.Lookup_Rate(lng_ValDate, lng_MatDate, "DF", , , True)
        dbl_DFSettleToValSpot = irc_SpotDisc.Lookup_Rate(lng_SpotDate, lng_SettlementDate, "DF", , , True)
        dbl_DFDivPaymentToValSpot = irc_SpotDisc.Lookup_Rate(lng_SpotDate, lng_DivPaymentDate, "DF", , , True)
        dbl_DFMatSpotToDivPayment = irc_SpotDisc.Lookup_Rate(lng_DivPaymentDate, lng_DelivDate, "DF", , , True)
        dbl_DFRhoToVal = irc_SpotDisc.Lookup_Rate(lng_ValDate, lng_RhoDate, "DF", , , True)
        dbl_DivDF = irc_SpotDisc.Lookup_Rate(lng_SpotDate, lng_DivPaymentDate, "DF", , , True) 'lng_SpotDate
        If str_MarketDiscCurve <> "" Then _
        dbl_DFValSpotToVal_market = irc_MarketDisc.Lookup_Rate(lng_ValDate, lng_SpotDate, "DF", , , True)
        dbl_DFExDivToDivPayment = irc_SpotDisc.Lookup_Rate(lng_DivExDate, lng_DivPaymentDate, "DF", , , True)
        'DF if MarketDiscCurve is different from SpotDiscCurve
        If str_MarketDiscCurve <> "" Then dbl_DFValSpotToVal = dbl_DFValSpotToVal_market
         If lng_PaymentSpotDate >= lng_ValDate Then dbl_DFValToPaymentSpot = irc_SpotDisc.Lookup_Rate(lng_ValDate, lng_PaymentSpotDate, "DF", , , True) 'new
    Else
        'For valuation date after or at expiry option
        dbl_DFMatSpotToVal = irc_SpotDisc.Lookup_Rate(lng_ValDate, lng_DelivDate, "DF", , , True)
        dbl_DFSettleToVal = irc_SpotDisc.Lookup_Rate(lng_ValDate, lng_SettlementDate, "DF", , , True)
        dbl_DFValSpotToVal = irc_SpotDisc.Lookup_Rate(lng_ValDate, lng_SpotDate, "DF", , , True)
        If lng_PaymentSpotDate > lng_ValDate Then dbl_DFValToPaymentSpot = irc_SpotDisc.Lookup_Rate(lng_ValDate, lng_PaymentSpotDate, "DF", , , True) 'new
    End If



End Sub


' ## PROPERTIES - PUBLIC
Public Property Get marketvalue() As Double
    ' ## Get discounted value of future flows in the PnL currency
Call FillMktData

Select Case str_PLType
    Case "PL"
        If bln_IsFutures = True Then dbl_underlying = FutForward()

        marketvalue = Option_Premium * lng_Quantity * lng_LotSize * FXConvFactor() * int_Sign * dbl_Parity


    Case "Delta"
        If bln_IsFutures = True Then dbl_underlying = dbl_FutPrice
        marketvalue = Delta * lng_Quantity * lng_LotSize * int_Sign * dbl_Parity

    Case "Gamma"
        If bln_IsFutures = True Then dbl_underlying = dbl_FutPrice
        marketvalue = Gamma * lng_Quantity * lng_LotSize * int_Sign * dbl_Parity

    Case "Delta Inv Amount"
        If bln_IsFutures = True Then dbl_underlying = dbl_FutPrice

        If bln_IsFutures = True Then
            marketvalue = Delta * lng_Quantity * lng_LotSize * FXConvFactor() * int_Sign * dbl_Parity * dbl_underlying
        Else
            marketvalue = Delta * lng_Quantity * lng_LotSize * FXConvFactor() * int_Sign * dbl_Parity * dbl_Spot / dbl_DFValSpotToVal
        End If

    Case "Vega"
        If bln_IsFutures = True Then dbl_underlying = dbl_FutPrice
        marketvalue = Vega * lng_Quantity * lng_LotSize * FXConvFactor() * int_Sign * dbl_Parity
    Case "Rho"
        If bln_IsFutures = True Then dbl_underlying = dbl_FutPrice
        marketvalue = Rho * lng_Quantity * lng_LotSize * FXConvFactor() * int_Sign * dbl_Parity
    Case "EQ NOP"
        If bln_IsFutures = True Then dbl_underlying = dbl_FutPrice

        If bln_IsFutures = True Then
            marketvalue = Delta * lng_Quantity * lng_LotSize * FXConvFactor() * int_Sign * dbl_Parity * dbl_underlying
        Else
            marketvalue = Delta * lng_Quantity * lng_LotSize * FXConvFactor() * int_Sign * dbl_Parity * dbl_Spot / dbl_DFValSpotToVal
        End If

End Select

End Property

Private Property Get Option_Premium() As Double

'use this to handle theta
Dim lng_DelivDate_temp As Long

If str_Settlement = "CASH" Then
    lng_DelivDate_temp = lng_SettlementDate
Else
    lng_DelivDate_temp = lng_DelivDate
End If


If lng_MatDate > lng_ValDate Then
    Select Case str_ExerciseType
        Case "EUROPEAN"
            Option_Premium = BSPrice_European

        Case "AMERICAN"
            'control variate
            'Option_Premium = BSPrice_European - Cox_American(False) + Cox_American(True)
            'Option_Premium = BSPrice_European - Cox_Fwd(False) + Cox_Fwd(True)
            Option_Premium = Cox_Fwd
    End Select

ElseIf lng_DelivDate_temp < lng_ValDate Then
    If str_Settlement <> "CASH" Then dbl_Spot = dbl_Spot * dbl_DFValSpotToVal
    Option_Premium = WorksheetFunction.Max(0, (dbl_Spot - dbl_Strike) * int_optdirection)
Else
    If str_Settlement = "CASH" Then dbl_DFMatSpotToVal = dbl_DFSettleToVal
    If str_Settlement <> "CASH" Then dbl_Spot = dbl_Spot * dbl_DFValSpotToVal / dbl_DFMatSpotToVal
    Option_Premium = WorksheetFunction.Max(0, (dbl_Spot - dbl_Strike) * int_optdirection) * _
    dbl_DFMatSpotToVal

End If

End Property

Public Property Get Cash() As Double
    ' ## Get discounted value of the cost price in the PnL currency

Call FillMktData

If str_PLType = "PL" Then
    Cash = -scf_Purchase.CalcValue(lng_ValDate, lng_SpotDate, str_PnLCcy) * int_Sign
Else
    Cash = 0
End If

End Property

Public Property Get PnL() As Double
    PnL = Me.marketvalue + Me.Cash
End Property

' Private property : Calculating European BS Price
' Taking into account the Delivery/Cash Settlement Date & Dividend

Private Property Get BSPrice_European() As Double

Dim dbl_Premium As Double

'Converting equity spot price to equity forward price (handling the dividend here)


If str_DivType <> "" And lng_ValDate < lng_DivExDate Then
    dbl_Forward = (dbl_Spot / dbl_DFDivPaymentToValSpot - dbl_DivAmount) / dbl_DFMatSpotToDivPayment
ElseIf str_DivType <> "" And lng_ValDate >= lng_DivExDate Then
    dbl_Forward = (dbl_Spot / dbl_DFDivPaymentToValSpot - dbl_DivAmount) / dbl_DFMatSpotToDivPayment
Else
    dbl_Forward = dbl_Spot * 1 / dbl_DFMatSpotToValSpot
End If

'Future option
If bln_IsFutures = True Then dbl_Forward = dbl_underlying

'Black-Scholes Formula
dbl_Vol_factor = dbl_Vol / 100 * (dbl_Time_MatToVal) ^ 0.5
d1 = (Log(dbl_Forward / dbl_Strike) + dbl_Vol_factor * dbl_Vol_factor / 2) / dbl_Vol_factor
d2 = d1 - dbl_Vol_factor

'option price before discount (included payoff and probability)
dbl_ForwardPremium = dbl_Forward * WorksheetFunction.NormSDist(d1 * int_optdirection) * int_optdirection - _
dbl_Strike * WorksheetFunction.NormSDist(d2 * int_optdirection) * int_optdirection


'discounting to valuation date (different discounting for cash settlement/delivery)
If str_Settlement = "CASH" Then
    dbl_Premium = dbl_ForwardPremium * dbl_DFSettleToValSpot * dbl_DFValSpotToVal
Else
    dbl_Premium = dbl_ForwardPremium * dbl_DFMatSpotToValSpot * dbl_DFValSpotToVal
End If


'output
BSPrice_European = dbl_Premium


End Property

Private Property Get Cox_American(Bln_IsAmerican As Boolean) As Double

Dim i As Integer, j As Integer
Dim dbl_TDiv As Double 'dividend date to spot date

Dim lng_DivNode As Long
Dim dbl_CFDivToDivNode As Long
Dim temp As Double
Dim dbl_DivDF_dt As Double

'Premilinaries before Contructing Binomial Tree
int_n = 500

'computation of binomial trees input
dbl_dt = dbl_Time_MatSpotToValSpot / int_n
dbl_bin_u = Exp(dbl_Vol / 100 * (dbl_Time_MatToVal / int_n) ^ 0.5)
dbl_bin_d = 1 / dbl_bin_u
dbl_DF_dt = dbl_DFMatSpotToValSpot ^ (dbl_dt / dbl_Time_MatSpotToValSpot)
dbl_bin_p = ((1 / dbl_DF_dt) - dbl_bin_d) / (dbl_bin_u - dbl_bin_d)
dbl_bin_q = 1 - dbl_bin_p
dbl_Forward = dbl_Spot * 1 / dbl_DFMatSpotToValSpot
dbl_TDiv = (lng_DivExDate - lng_ValDate) / 365 '(lng_DivDate - lng_SpotDate) / 365

''define array for the tree. arr_Binomial(time interval, branches at each time interval, 1=stock price simulation; 2= payoff
ReDim arr_Binomial(0 To int_n, 1 To int_n + 1, 1 To 2) As Variant


'' Finding div node. If dividend falls between node t and t+1, return t
'' Setting the stock price at t=0 for div-paying and non-div-paying stock
If str_DivType <> "" And lng_ValDate < lng_DivExDate Then

    int_DivPeriod = WorksheetFunction.RoundDown(dbl_TDiv / dbl_dt, 0)
    dbl_PVDiv = dbl_DivAmount * dbl_DivDF
    arr_Binomial(0, 1, 1) = dbl_Spot - dbl_PVDiv

Else
    arr_Binomial(0, 1, 1) = dbl_Spot
End If

'' Binomial tree starts here
For i = 1 To int_n
    For j = 1 To i
        arr_Binomial(i, j, 1) = arr_Binomial(i - 1, j, 1) * dbl_bin_u
    Next j
        arr_Binomial(i, i + 1, 1) = arr_Binomial(i - 1, i, 1) * dbl_bin_d
Next i

''Simulated stock price adjustment for div-paying stock
''PV of Div will be added back to the tree here
''This step is skip if calculating for European cases
If Bln_IsAmerican = True Then
    For i = 0 To int_n
        For j = 1 To i + 1

                If i <= int_DivPeriod And lng_ValDate < lng_DivExDate Then
                    arr_Binomial(i, j, 1) = arr_Binomial(i, j, 1) + dbl_PVDiv / (dbl_DF_dt ^ i)
                End If

        Next j
    Next i
End If

''Initialization of payoff at the last node
For j = 1 To int_n + 1
    arr_Binomial(int_n, j, 2) = WorksheetFunction.Max(0, int_optdirection * (arr_Binomial(int_n, j, 1) - dbl_Strike))
Next j

''Discounting the tree backward
For i = int_n - 1 To 0 Step -1
    For j = 1 To i + 1
            arr_Binomial(i, j, 2) = (dbl_bin_p * arr_Binomial(i + 1, j, 2) + dbl_bin_q * arr_Binomial(i + 1, j + 1, 2)) * dbl_DF_dt
        'only for American option
        If Bln_IsAmerican = True Then
            arr_Binomial(i, j, 2) = WorksheetFunction.Max(arr_Binomial(i, j, 2), int_optdirection * (arr_Binomial(i, j, 1) - dbl_Strike))
        End If
    Next j
Next i

''Final step

Cox_American = arr_Binomial(0, 1, 2) * dbl_DFValSpotToVal


End Property

Private Function Black76(bln_Call As Boolean, k As Double, F As Variant, Vol As Double, t As Double, te As Double, r As Double) As Double

'Black Forward Model used specifically for American Option Initialization

Dim x As Double
Dim d1 As Double
Dim d2 As Double

x = Vol / 100 * (te ^ 0.5)

d1 = (Log(F / k) + (x ^ 2) / 2) / x
d2 = d1 - x

Dim dbl_Output As Double

'If bln_Call = True Then
'    dbl_output = (F * WorksheetFunction.NormSDist(d1) - K * WorksheetFunction.NormSDist(d2)) * Exp(-r * t)
'Else: dbl_output = (K * WorksheetFunction.NormSDist(-d2) - F * WorksheetFunction.NormSDist(-d1)) * Exp(-r * t)
'End If

If bln_Call = True Then
    dbl_Output = (F * WorksheetFunction.NormSDist(d1) - k * WorksheetFunction.NormSDist(d2)) * Exp(-r * t)
Else: dbl_Output = (k * WorksheetFunction.NormSDist(-d2) - F * WorksheetFunction.NormSDist(-d1)) * Exp(-r * t)
End If





Black76 = WorksheetFunction.Max(0, dbl_Output)

End Function

Private Property Get Cox_Fwd(Optional bln_American As Boolean = True) As Double
    'Kindly note the diff between the code and the quant test report formula
    'Quant Test Report: Payment Spot Date = Maturity Date + Option Spot for cash and Maturity Date + Spot delay for delivery
    'Code: Payment Spot Date = Valuation Date + Option Spot

    'Quant Test Report: Settlement Date = Valuation Date + Option Spot
    'Code: Settlement Date = Maturity Date + Option Spot for cash and Maturity Date + Spot delay for delivery (Part of Deal Entry Info in worksheet)


    'Cox Forward Model

    Dim i As Integer, j As Integer
    Dim dbl_TDiv As Double ' ex dividend date to spot date
    Dim int_DivPeriod As Integer
    Dim lng_DivNode As Long
    Dim dbl_CFDivToDivNode As Long
    Dim temp As Double
    Dim dbl_DivDF_dt As Double

    'Premilinaries before Contructing Binomial Tree
    'int_n = 30
    int_n = 500

    If str_DivType = "Cash" Then Call FillDivSpotDate

    Call FillSettlementDate ' fill up payment spot date!! not settlement date


    dbl_r = irc_SpotDisc.Lookup_Rate(lng_PaymentSpotDate, lng_SettlementDate, "ZERO", , , True) / 100


    'computation of binomial trees input
    Dim dbl_dt_Te As Double
    dbl_dt_Te = dbl_Time_MatToVal / int_n


    dbl_dt = (lng_DelivDate - lng_SpotDate) / 365 / int_n

    If str_Settlement = "CASH" Then dbl_dt = (lng_SettlementDate - lng_PaymentSpotDate) / 365 / int_n

    Dim dbl_dt_Opt As Double
    Dim dbl_dt_Und As Double

    dbl_dt_Opt = (lng_SettlementDate - lng_PaymentSpotDate) / 365 / int_n
    dbl_dt_Und = (lng_DelivDate - lng_SpotDate) / 365 / int_n
    dbl_r_Und = irc_SpotDisc.Lookup_Rate(lng_SpotDate, lng_DelivDate, "ZERO", , , True) / 100


    dbl_bin_u = Exp(dbl_Vol / 100 * (dbl_Time_MatToVal / int_n) ^ 0.5)
    dbl_bin_d = 1 / dbl_bin_u


    dbl_DF_dt = Exp(-dbl_r * dbl_dt_Opt)

    dbl_bin_p = (1 - dbl_bin_d) / (dbl_bin_u - dbl_bin_d)
    dbl_bin_q = 1 - dbl_bin_p

    'Converting equity spot price to equity forward price (handling the dividend here)
    If str_DivType = "" Then
        dbl_Forward = dbl_Spot * 1 / dbl_DFMatSpotToValSpot
    ElseIf str_DivType = "Cash" Then
        dbl_Forward = (dbl_Spot / dbl_DFDivPaymentToValSpot - dbl_DivAmount) / dbl_DFMatSpotToDivPayment
    End If

    Dim dbl_r1 As Double
    Dim dbl_r2 As Double


    If str_DivType = "Cash" Then
        dbl_TDiv = (lng_DivExDate - lng_DivSpotDate) / (lng_DivExpiryDate - lng_DivSpotDate) * dbl_dt_Und * int_n
        'Set dbl_r1 to 0 when valuation date is on or after ex-div date
        If dbl_TDiv <= 0 Then
            dbl_r1 = 0
        Else
            dbl_r1 = Log(1 / irc_SpotDisc.Lookup_Rate(lng_SpotDate, lng_DivExDate, "DF", , , True)) / dbl_TDiv
        dbl_r2 = Log(1 / irc_SpotDisc.Lookup_Rate(lng_DivExDate, lng_DelivDate, "DF", , , True)) / (dbl_dt_Und * int_n - dbl_TDiv)
    End If

    ''define array for the tree. arr_Binomial(time interval, branches at each time interval, 1=forward price simulation; 2= payoff, 3 = spot price
    ReDim arr_Binomial(0 To int_n, 1 To int_n + 1, 1 To 5) As Variant
    'arr_Binomial(x,y,1) = Forward Value
    'arr_Binomial(x,y,2) = Max (Option Value, Exercise Value, 0)
    'arr_Binomial(x,y,3) = Spot Value


    '' Finding div node. If dividend falls between node t and t+1, return t
    '' Setting the stock price at t=0 for div-paying and non-div-paying stock
    Dim dbl_NodeDiv As Double

    If str_DivType <> "" And lng_ValDate < lng_DivExDate Then
        dbl_NodeDiv = dbl_TDiv / dbl_dt_Und
        int_DivPeriod = WorksheetFunction.RoundDown(dbl_NodeDiv, 0)
    End If


    Dim dbl_DfDivExDiv As Double
    dbl_DfDivExDiv = irc_SpotDisc.Lookup_Rate(lng_DivExDate, lng_DivPaymentDate, "DF", , , True)


    Dim dbl_ExDiv As Double
    dbl_ExDiv = dbl_DivAmount * dbl_DfDivExDiv

    arr_Binomial(0, 1, 1) = dbl_Forward

    '' Binomial tree starts here
    For i = 1 To int_n
        For j = 1 To i
            arr_Binomial(i, j, 1) = arr_Binomial(i - 1, j, 1) * dbl_bin_u
        Next j
            arr_Binomial(i, i + 1, 1) = arr_Binomial(i - 1, i, 1) * dbl_bin_d
    Next i

        For i = 0 To int_n - 1
            For j = 1 To i + 1
                    If str_DivType = "" Then
                        arr_Binomial(i, j, 3) = arr_Binomial(i, j, 1) * Exp(-dbl_r_Und * (int_n - i) * dbl_dt_Und)
                    Else
                        If i <= int_DivPeriod And lng_ValDate < lng_DivExDate Then
                            arr_Binomial(i, j, 3) = arr_Binomial(i, j, 1) * Exp(-(dbl_r2 * (int_n - dbl_NodeDiv) * dbl_dt_Und + dbl_r1 * (dbl_NodeDiv - i) * dbl_dt_Und)) + _
                                dbl_ExDiv / Exp(dbl_r1 * (dbl_NodeDiv - i) * dbl_dt_Und)
                        Else
                            arr_Binomial(i, j, 3) = arr_Binomial(i, j, 1) * Exp(-dbl_r2 * (int_n - i) * dbl_dt_Und)
                        End If
                    End If
            Next j
        Next i

    ''Initialization of payoff at the last node

    For j = 1 To int_n
        arr_Binomial(int_n - 1, j, 2) = Black76(bln_IsCall, dbl_Strike, arr_Binomial(int_n - 1, j, 1), dbl_Vol, dbl_dt_Opt, dbl_dt_Te, dbl_r)
    Next j

    ''Discounting the tree backward
    For i = int_n - 1 To 0 Step -1

        For j = 1 To i + 1

            If i <> int_n - 1 Then
                arr_Binomial(i, j, 2) = (dbl_bin_p * arr_Binomial(i + 1, j, 2) + dbl_bin_q * arr_Binomial(i + 1, j + 1, 2)) * dbl_DF_dt
            End If

            If bln_American = True Then
                arr_Binomial(i, j, 2) = WorksheetFunction.Max(arr_Binomial(i, j, 2), int_optdirection * (arr_Binomial(i, j, 3) - dbl_Strike))
            End If

            Dim dbl_S_BeforeDiv As Double
            Dim dbl_S_AfterDiv As Double

            'Early exercise is optimal right before the ex-div date when there is dividend
            'Only execute this when valuation date is before ex-div date
            If str_DivType <> "" Then
                If i = int_DivPeriod and int_DivPeriod >0 Then

                    dbl_S_AfterDiv = (arr_Binomial(i, j, 1)) _
                                        * Exp(-dbl_r2 * (int_n - dbl_NodeDiv) * dbl_dt_Und)

                    dbl_S_BeforeDiv = dbl_S_AfterDiv + dbl_ExDiv

                    arr_Binomial(i, j, 2) = WorksheetFunction.Max(arr_Binomial(i, j, 2), _
                        int_optdirection * (dbl_S_BeforeDiv - dbl_Strike) * Exp(-dbl_r * (dbl_NodeDiv - int_DivPeriod) * dbl_dt_Opt), _
                        int_optdirection * (dbl_S_AfterDiv - dbl_Strike) * Exp(-dbl_r * (dbl_NodeDiv - int_DivPeriod) * dbl_dt_Opt))


                End If
            End If

        Next j
    Next i

    ''Final step

    If dbl_r <= 0 And bln_IsCall = True Then Cox_Fwd = (dbl_Spot - dbl_Strike)

    Cox_Fwd = arr_Binomial(0, 1, 2) * irc_SpotDisc.Lookup_Rate(lng_ValDate, lng_PaymentSpotDate, "DF", , , True)


End Property


Private Property Get FutForward() As Double
    Dim dbl_FutSpread As Double

    'lng_ValDate_fixed = lng_ValDate
    irc_SpotDisc.SetCurveState (original)
    'SetValDate (lng_OriValDate)
    dbl_FutSpread = dbl_FutPrice - eqd_Spots.Lookup_Spot(str_SecCode, True) _
                    / irc_SpotDisc.Lookup_Rate(lng_SpotDate, lng_FutMatSpotDate, "DF", , , True, , 0)
       irc_SpotDisc.SetCurveState (Final)
    'SetValDate (lng_ValDate_fixed)

    FutForward = dbl_Spot * FutCapFactor() + dbl_FutSpread
End Property

Private Property Get FutCapFactor(Optional dbl_spread As Double = 0) As Double
    ' Capitalization Factor for forward projection
    FutCapFactor = 1 / irc_SpotDisc.Lookup_Rate(lng_SpotDate, lng_FutMatSpotDate, "DF", , , True, , dbl_spread)
End Property

Private Property Get Delta() As Double

Dim dbl_OriPremium As Double
Dim dbl_DeltaPremium As Double
Dim dbl_DeltaMagnitude As Double

'Select Case str_ExerciseType

'Case "EUROPEAN"
    dbl_DeltaMagnitude = 0.0001
    dbl_OriPremium = Option_Premium
    If bln_IsFutures = True Then dbl_underlying = dbl_underlying * (1 + dbl_DeltaMagnitude)
    dbl_Spot = dbl_Spot * (1 + dbl_DeltaMagnitude)
    dbl_DeltaPremium = Option_Premium
    If bln_IsFutures = True Then dbl_underlying = dbl_underlying / (1 + dbl_DeltaMagnitude)
    dbl_Spot = dbl_Spot / (1 + dbl_DeltaMagnitude)

    Delta = (dbl_DeltaPremium - dbl_OriPremium) / dbl_Spot / dbl_DeltaMagnitude
    If bln_IsFutures = True Then Delta = (dbl_DeltaPremium - dbl_OriPremium) / dbl_underlying / dbl_DeltaMagnitude
'Case "AMERICAN"

'    Delta = Cox_American(True)
'    Delta = (arr_Binomial(1, 1, 2) - arr_Binomial(1, 2, 2)) / (arr_Binomial(1, 1, 1) - arr_Binomial(1, 2, 1)) * dbl_DFValSpotToVal


'End Select

End Property

Private Property Get Gamma() As Double

Dim dbl_OriPremium As Double
Dim dbl_DeltaPremium As Double
Dim dbl_2DeltaPremium As Double
Dim dbl_DeltaMagnitude As Double

    dbl_DeltaMagnitude = 0.0001
    dbl_OriPremium = Option_Premium
    If bln_IsFutures = True Then
        dbl_underlying = dbl_underlying * (1 + dbl_DeltaMagnitude)
        dbl_DeltaPremium = Option_Premium
        dbl_underlying = dbl_underlying / (1 + dbl_DeltaMagnitude) * (1 + 2 * dbl_DeltaMagnitude)
        dbl_2DeltaPremium = Option_Premium
        dbl_underlying = dbl_underlying / (1 + 2 * dbl_DeltaMagnitude)
        Gamma = (dbl_2DeltaPremium - 2 * dbl_DeltaPremium + dbl_OriPremium) / (dbl_underlying * dbl_DeltaMagnitude) ^ 2
    Else
        dbl_Spot = dbl_Spot * (1 + dbl_DeltaMagnitude)
        dbl_DeltaPremium = Option_Premium
        dbl_Spot = dbl_Spot / (1 + dbl_DeltaMagnitude) * (1 + 2 * dbl_DeltaMagnitude)
        dbl_2DeltaPremium = Option_Premium
        dbl_Spot = dbl_Spot / (1 + 2 * dbl_DeltaMagnitude)
        Gamma = (dbl_2DeltaPremium - 2 * dbl_DeltaPremium + dbl_OriPremium) / (dbl_Spot * dbl_DeltaMagnitude) ^ 2
    End If

    Gamma = Gamma / dbl_DFValToPaymentSpot
End Property
Private Property Get Vega() As Double
'Alv Updated 4/7/2018
Dim dbl_OriPremium As Double
Dim dbl_VegaPremium As Double
Dim dbl_VegaMagnitude As Double



If lng_PaymentSpotDate <= lng_ValDate Then dbl_DFValToPaymentSpot = 1

'Alv muted  "'*100" 7/6/2018
dbl_VegaMagnitude = 0.0001 '* 100 'in % form
'Alv muted  "'*100" 7/6/2018

dbl_OriPremium = Option_Premium / dbl_DFValToPaymentSpot
dbl_Vol = dbl_Vol + dbl_VegaMagnitude
dbl_VegaPremium = Option_Premium / dbl_DFValToPaymentSpot
dbl_Vol = dbl_Vol - dbl_VegaMagnitude
Vega = (dbl_VegaPremium - dbl_OriPremium) / dbl_VegaMagnitude

End Property


Private Property Get Rho() As Double

Dim dbl_OriPremium As Double
Dim dbl_RhoPremium As Double
Dim int_num As Integer
Dim int_cnt As Integer

If lng_PaymentSpotDate <= lng_ValDate Then dbl_DFValToPaymentSpot = 1

'get pre-shock value
dbl_OriPremium = Option_Premium / dbl_DFValToPaymentSpot

'get number of pillars of the curve
int_num = irc_SpotDisc.NumPoints

'Start the loop to shock pillar one by one
For int_cnt = 1 To int_num
    Call irc_SpotDisc.SetCurveState(Zero_Up1BP, int_cnt)
    Call FillMktData
    dbl_RhoPremium = dbl_RhoPremium + (Option_Premium / dbl_DFValToPaymentSpot - dbl_OriPremium)
    irc_SpotDisc.SetCurveState (Final)
    'Debug.Print int_cnt & "," & (Option_Premium - dbl_OriPremium)
Next int_cnt

'val date = expiry date
If lng_MatDate = lng_ValDate Then dbl_RhoPremium = 0

'output
Rho = dbl_RhoPremium

End Property


Private Property Get FXConvFactor() As Double

    ' ## Discounted spot to translate into the PnL currency
    FXConvFactor = fxs_Spots.Lookup_DiscSpot(str_SecCcy, str_PnLCcy)
    If str_MarketDiscCurve <> "" Then _
    FXConvFactor = fxs_Spots.Lookup_Spot(str_SecCcy, str_PnLCcy)

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
' ## METHODS - PRIVATE
Private Sub FillDivSpotDate()
    ' ## Compute and store dividend start date and expiry date
    If int_DivSpot = 0 Then
        lng_DivSpotDate = Date_ApplyBDC(lng_ValDate, "FOLL", cal_Spot.HolDates, cal_Spot.Weekends)
        lng_DivExpiryDate = Date_ApplyBDC(lng_MatDate, "FOLL", cal_Spot.HolDates, cal_Spot.Weekends)
    Else
        lng_DivSpotDate = date_workday(lng_ValDate, int_DivSpot, cal_Spot.HolDates, cal_Spot.Weekends)
        lng_DivExpiryDate = date_workday(lng_MatDate, int_DivSpot, cal_Spot.HolDates, cal_Spot.Weekends)
    End If
End Sub
' ## METHODS - PRIVATE
Private Sub FillSettlementDate()

    If int_OptionSpot = 0 Then
        lng_PaymentSpotDate = Date_ApplyBDC(lng_ValDate, "FOLL", cal_Spot.HolDates, cal_Spot.Weekends)
    Else
        lng_PaymentSpotDate = date_workday(lng_ValDate, int_OptionSpot, cal_Spot.HolDates, cal_Spot.Weekends)

    End If
End Sub

Private Property Get Option_AmericanDiv_Theta() As Double
Dim lng_SpotDate_temp As Long
Dim dbl_Option_premium_temp As Double

lng_SpotDate_temp = lng_SpotDate
lng_SpotDate = lng_DivExDate

dbl_Option_premium_temp = Cox_Fwd

lng_SpotDate = lng_SpotDate_temp

Option_AmericanDiv_Theta = dbl_Option_premium_temp * dbl_DFValSpotToVal

End Property
Private Sub FillFutMatSpotDate()
    ' ## Compute and store maturity spot date
    If int_SpotDays = 0 Then
        lng_FutMatSpotDate = Date_ApplyBDC(lng_FutMatDate, "FOLL", cal_Spot.HolDates, cal_Spot.Weekends)
    Else
        lng_FutMatSpotDate = date_workday(lng_FutMatDate, int_SpotDays, cal_Spot.HolDates, cal_Spot.Weekends)
    End If
End Sub

' ## METHODS - CALCULATION DETAILS
Public Sub OutputReport(wks_output As Worksheet)
'Alv: Updated on 4/7/2018
wks_output.Cells.Clear

    Dim rng_OutputTopLeft As Range: Set rng_OutputTopLeft = wks_output.Range("A1")
    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    Dim str_CurrencyFormat As String: str_CurrencyFormat = Gather_CurrencyFormat()
    Dim int_ActiveRow As Integer: int_ActiveRow = 0
    Dim initiate_all_value As Double

    Dim dic_Addresses As Dictionary: Set dic_Addresses = New Dictionary
    dic_Addresses.CompareMode = CompareMethod.TextCompare

    initiate_all_value = PnL

    ' Output overall info
    With rng_OutputTopLeft
        .Offset(int_ActiveRow, 0).Value = "Value date:"
        .Offset(int_ActiveRow, 1).Value = lng_ValDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("ValDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 2
        .Offset(int_ActiveRow, 0).Value = "Contract:"
        .Offset(int_ActiveRow, 1).Value = str_description

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Underlying:"
        .Offset(int_ActiveRow, 1).Value = str_SecCode

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Expiry date:"
        .Offset(int_ActiveRow, 1).Value = lng_MatDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("MatDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Delivery/Settlement date:"
        .Offset(int_ActiveRow, 1).Value = lng_DelivDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        If str_Settlement = "CASH" Then .Offset(int_ActiveRow, 2).Value = "Cash Settlement" Else .Offset(int_ActiveRow, 2).Value = "Delivery"
        Call dic_Addresses.Add("MatSpotDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Valuation Spot date:"
        .Offset(int_ActiveRow, 1).Value = lng_SpotDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("ValSpotDate", .Offset(int_ActiveRow, 1).Address(False, False))

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
        .Offset(int_ActiveRow, 1).Value = str_SpotDiscCurve
        Call dic_Addresses.Add("RateCurve", .Offset(int_ActiveRow, 1).Address(False, False))

        'Option details
        int_ActiveRow = int_ActiveRow + 2

        .Offset(int_ActiveRow, 0).Value = "Option details:"
        .Offset(int_ActiveRow, 0).Font.Italic = True

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "B/S Sign (1=Buy, -1=Sell):"
        .Offset(int_ActiveRow, 1).Value = int_Sign
        Call dic_Addresses.Add("B/S Sign", .Offset(int_ActiveRow, 1).Address(False, False))


        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Put/Call:"
        If bln_IsCall = True Then
            .Offset(int_ActiveRow, 1).Value = "Call"
        Else
            .Offset(int_ActiveRow, 1).Value = "Put"
        End If
        .Offset(int_ActiveRow, 2).Value = int_optdirection
        Call dic_Addresses.Add("optdirection", .Offset(int_ActiveRow, 2).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "European/American:"
        If str_ExerciseType = "EUROPEAN" Then
            .Offset(int_ActiveRow, 1).Value = "European"
        Else
            .Offset(int_ActiveRow, 1).Value = "American"
        End If


        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Strike Price:"
        .Offset(int_ActiveRow, 1).Value = dbl_Strike
        Call dic_Addresses.Add("strike", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Underlying Spot:"
        .Offset(int_ActiveRow, 1).Value = dbl_Spot
        Call dic_Addresses.Add("Und_Spot", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Underlying Vol:"
        .Offset(int_ActiveRow, 1).Value = dbl_Vol
        Call dic_Addresses.Add("Vol", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Fwd To Spot Disc Rate"
        .Offset(int_ActiveRow, 1).Value = dbl_r

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Option Disc Rate"
        .Offset(int_ActiveRow, 1).Value = dbl_r_Und

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Valuation + Option Spot"
        .Offset(int_ActiveRow, 1).Value = lng_PaymentSpotDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat


        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Dividend:"
        If str_DivType <> "" Then
            .Offset(int_ActiveRow, 1).Value = "Yes"

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Cash dividend:"
            .Offset(int_ActiveRow, 1).Value = dbl_DivAmount
            Call dic_Addresses.Add("DivAmount", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Ex Dividend date:"
            .Offset(int_ActiveRow, 1).Value = lng_DivExDate
            .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
            Call dic_Addresses.Add("DivDate", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Dividend Payment date:"
            .Offset(int_ActiveRow, 1).Value = lng_DivPaymentDate
            .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
            Call dic_Addresses.Add("DivPaymentDate", .Offset(int_ActiveRow, 1).Address(False, False))
        Else
            .Offset(int_ActiveRow, 1).Value = "None"
        End If

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Parity:"
        .Offset(int_ActiveRow, 1).Value = dbl_Parity

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Quantity:"
        .Offset(int_ActiveRow, 1).Value = lng_Quantity

        'PnL calculation
        int_ActiveRow = int_ActiveRow + 3

        .Offset(int_ActiveRow, 0).Value = "PnL Calculation:"
        .Offset(int_ActiveRow, 0).Font.Italic = True

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Option Premium:"
        .Offset(int_ActiveRow, 1).Value = Option_Premium

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "FX Spot:"
        .Offset(int_ActiveRow, 1).Value = FXConvFactor()

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Market Value (" & str_PnLCcy & "):"
        str_PLType = "PL"
        .Offset(int_ActiveRow, 1).Value = marketvalue
        Call dic_Addresses.Add("MV", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Cost (" & str_PnLCcy & "):"
        .Offset(int_ActiveRow, 1).Value = Cash
        Call dic_Addresses.Add("Cost", .Offset(int_ActiveRow, 1).Address(False, False))

         int_ActiveRow = int_ActiveRow + 1
         .Offset(int_ActiveRow, 0).Value = "PnL (" & str_PnLCcy & "):"
         .Offset(int_ActiveRow, 1).Value = "=" & dic_Addresses("MV") & "+" & dic_Addresses("Cost")
         .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20

        'Option Premium Calculation
        'Output is different for American and European option
        int_ActiveRow = int_ActiveRow + 3
        .Offset(int_ActiveRow, 0).Value = "Option Premium Calculation:"
        .Offset(int_ActiveRow, 0).Font.Italic = True

        'For European option
        If str_ExerciseType = "EUROPEAN" Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Method:"
            .Offset(int_ActiveRow, 1).Value = "Black-Scholes"

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "DFValuation Spot to Maturity Spot:"
            .Offset(int_ActiveRow, 0).Characters(Start:=3, Length:=31).Font.Subscript = True

            'handle cash settlement
            If str_Settlement = "CASH" Then .Offset(int_ActiveRow, 1).Value = dbl_DFSettleToValSpot _
            Else .Offset(int_ActiveRow, 1).Value = dbl_DFMatSpotToValSpot
            Call dic_Addresses.Add("DF_MatSpotToValSpot", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "DFValuation to Valuation Spot:"
            .Offset(int_ActiveRow, 0).Characters(Start:=3, Length:=27).Font.Subscript = True
            .Offset(int_ActiveRow, 1).Value = dbl_DFValSpotToVal
            Call dic_Addresses.Add("DF_ValSpotToVal", .Offset(int_ActiveRow, 1).Address(False, False))

            If str_DivType <> "" Then
                int_ActiveRow = int_ActiveRow + 1
                .Offset(int_ActiveRow, 0).Value = "DFValuation Spot to Dividend:"
                .Offset(int_ActiveRow, 0).Characters(Start:=3, Length:=26).Font.Subscript = True
                .Offset(int_ActiveRow, 1).Value = dbl_DFDivPaymentToValSpot
                Call dic_Addresses.Add("DF_DivToValSpot", .Offset(int_ActiveRow, 1).Address(False, False))

                int_ActiveRow = int_ActiveRow + 1
                .Offset(int_ActiveRow, 0).Value = "DFDividend to Maturity Spot:"
                .Offset(int_ActiveRow, 0).Characters(Start:=3, Length:=25).Font.Subscript = True
                .Offset(int_ActiveRow, 1).Value = dbl_DFMatSpotToDivPayment
                Call dic_Addresses.Add("DF_MatSpotToDiv", .Offset(int_ActiveRow, 1).Address(False, False))

                int_ActiveRow = int_ActiveRow + 1
                .Offset(int_ActiveRow, 0).Value = "Forward:"


                'WL modification 6-Jun-2018
                .Offset(int_ActiveRow, 1).Value = "=" & dic_Addresses("Und_Spot") & "/" & dic_Addresses("DF_DivToValSpot") _
                    & "-" & dic_Addresses("DivAmount") & "/" & dic_Addresses("DF_MatSpotToDiv")
                Call dic_Addresses.Add("forward", .Offset(int_ActiveRow, 1).Address(False, False))
                'WL modification 6-Jun-2018

                .Offset(int_ActiveRow, 2).Value = dbl_Forward

            Else
                int_ActiveRow = int_ActiveRow + 1
                .Offset(int_ActiveRow, 0).Value = "Forward:"
                .Offset(int_ActiveRow, 1).Value = "=" & dic_Addresses("Und_Spot") & "/" & dic_Addresses("DF_MatSpotToValSpot") 'dbl_DFMatSpotToValSpot
                Call dic_Addresses.Add("forward", .Offset(int_ActiveRow, 1).Address(False, False))
                .Offset(int_ActiveRow, 2).Value = dbl_Forward
            End If

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Volatility factor:"
            .Offset(int_ActiveRow, 1).Value = "=" & dic_Addresses("Vol") & " / 100 * ((" & dic_Addresses("MatDate") & "-" & dic_Addresses("ValDate") & ")/365) ^ 0.5"
            Call dic_Addresses.Add("vol_factor", .Offset(int_ActiveRow, 1).Address(False, False))
            .Offset(int_ActiveRow, 2).Value = dbl_Vol_factor

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "d1:"
            .Offset(int_ActiveRow, 0).Characters(Start:=2, Length:=1).Font.Subscript = True
            .Offset(int_ActiveRow, 1).Value = "=(ln(" & dic_Addresses("forward") & "/" & dic_Addresses("strike") _
                    & ")+(" & dic_Addresses("vol_factor") & "^2)/2)/" & dic_Addresses("vol_factor")
            .Offset(int_ActiveRow, 2).Value = d1
            Call dic_Addresses.Add("d1", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "d2:"
            .Offset(int_ActiveRow, 0).Characters(Start:=2, Length:=1).Font.Subscript = True
            .Offset(int_ActiveRow, 1).Value = "=" & dic_Addresses("d1") & "-" & dic_Addresses("vol_factor")
            .Offset(int_ActiveRow, 2).Value = d2
            Call dic_Addresses.Add("d2", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Forward Premium"
            .Offset(int_ActiveRow, 1).Value = "=" & dic_Addresses("forward") & "*NORMSDIST(" & dic_Addresses("d1") & " * " & dic_Addresses("optdirection") _
             & ")*" & dic_Addresses("optdirection") & "-" & dic_Addresses("strike") & "*NORMSDIST(" & dic_Addresses("d2") & " * " & dic_Addresses("optdirection") _
             & ")*" & dic_Addresses("optdirection")
            .Offset(int_ActiveRow, 2).Value = dbl_ForwardPremium
            Call dic_Addresses.Add("forwardPremium", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Premium:"
            .Offset(int_ActiveRow, 1).Value = "=" & dic_Addresses("forwardPremium") & "*" & dic_Addresses("DF_MatSpotToValSpot") & "*" & dic_Addresses("DF_ValSpotToVal")
            .Offset(int_ActiveRow, 2).Value = Option_Premium
            .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20

       Else
       'For American option

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Method:"
            .Offset(int_ActiveRow, 1).Value = "Binomial Tree"

'            int_ActiveRow = int_ActiveRow + 1
'            .Offset(int_ActiveRow, 0).Value = "Control Variate Technique:"
'            .Offset(int_ActiveRow, 0).Font.Italic = True

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "European Black-Scholes:"
            .Offset(int_ActiveRow, 1).Value = BSPrice_European
            Call dic_Addresses.Add("BS_European", .Offset(int_ActiveRow, 1).Address(False, False))

'            int_ActiveRow = int_ActiveRow + 1
'            .Offset(int_ActiveRow, 0).Value = "European Tree:"
'            .Offset(int_ActiveRow, 1).Value = Cox_American(False)
'            Call dic_Addresses.Add("Tree_European", .Offset(int_ActiveRow, 1).Address(False, False))

             int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "American Tree:"
            '.Offset(int_ActiveRow, 1).Value = Cox_American(True)
            .Offset(int_ActiveRow, 1).Value = Cox_Fwd(True)
            Call dic_Addresses.Add("Tree_American", .Offset(int_ActiveRow, 1).Address(False, False))

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Option Premium:"
            '.Offset(int_ActiveRow, 1).Value = "=" & dic_Addresses("BS_European") & "+(" & dic_Addresses("Tree_American") & "-" & dic_Addresses("Tree_European") & ")"
            .Offset(int_ActiveRow, 1).Value = "=" & dic_Addresses("Tree_American")


            'Binomial tree details
            int_ActiveRow = int_ActiveRow + 2
            .Offset(int_ActiveRow, 0).Value = "Binomial Tree Details:"
            .Offset(int_ActiveRow, 0).Font.Italic = True

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Valuation Spot to Maturity Spot (year):"
            .Offset(int_ActiveRow, 1).Value = "=(" & dic_Addresses("MatSpotDate") & "-" & dic_Addresses("ValSpotDate") & ")/365"

            If str_DivType <> "" Then
                int_ActiveRow = int_ActiveRow + 1
                .Offset(int_ActiveRow, 0).Value = "Valuation Spot to Dividend (year):"
                .Offset(int_ActiveRow, 1).Value = "=(" & dic_Addresses("DivDate") & "-" & dic_Addresses("ValSpotDate") & ")/365"
            End If

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Number of Steps:"
            .Offset(int_ActiveRow, 1).Value = int_n

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "dt:"
            .Offset(int_ActiveRow, 1).Value = dbl_dt

             int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "DFdt:"
            .Offset(int_ActiveRow, 0).Characters(Start:=3, Length:=2).Font.Subscript = True
            .Offset(int_ActiveRow, 1).Value = dbl_DF_dt

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "u:"
            .Offset(int_ActiveRow, 1).Value = dbl_bin_u

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "d:"
            .Offset(int_ActiveRow, 1).Value = dbl_bin_d

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "p:"
            .Offset(int_ActiveRow, 1).Value = dbl_bin_p

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "q:"
            .Offset(int_ActiveRow, 1).Value = dbl_bin_q

            'Binomial tree construction
            'loops to output array starts here
            Dim i As Integer, j As Integer

            int_ActiveRow = int_ActiveRow + 2
            .Offset(int_ActiveRow, 0).Value = "Binomial Tree Nodes:"
            .Offset(int_ActiveRow, 0).Font.Italic = True

            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Steps:"
            For i = 0 To int_n
                .Offset(int_ActiveRow, i + 1).Value = i
            Next i

            If str_DivType <> "" Then
                int_ActiveRow = int_ActiveRow + 1
                .Offset(int_ActiveRow, 0).Value = "PV of Dividends:"
                For i = 0 To int_DivPeriod
                .Offset(int_ActiveRow, i + 1).Value = dbl_PVDiv / (dbl_DF_dt ^ i)
                Next i
            End If
            int_ActiveRow = int_ActiveRow + 1

'            .Offset(int_ActiveRow, 0).Value = "Stock Price Simulation:"
'            For i = 0 To int_n
'                For j = 1 To int_n + 1
'                .Offset(int_ActiveRow + j - 1, i + 1).Value = arr_Binomial(i, j, 1)
'                Next j
'            Next i
'
'            int_ActiveRow = int_ActiveRow + int_n + 1
'            .Offset(int_ActiveRow, 0).Value = "Option Value:"
'            For i = 0 To int_n
'                For j = 1 To int_n + 1
'                .Offset(int_ActiveRow + j - 1, i + 1).Value = arr_Binomial(i, j, 2)
'                Next j
'            Next i

            .Offset(int_ActiveRow, 0).Value = "Forward:"
            For i = 0 To int_n
                For j = 1 To int_n + 1
                .Offset(int_ActiveRow + j - 1, i + 1).Value = arr_Binomial(i, j, 1)
                Next j
            Next i

            int_ActiveRow = int_ActiveRow + int_n + 1
            .Offset(int_ActiveRow, 0).Value = "Spot:"
            For i = 0 To int_n
                For j = 1 To int_n + 1
                .Offset(int_ActiveRow + j - 1, i + 1).Value = arr_Binomial(i, j, 3)
                Next j
            Next i

            int_ActiveRow = int_ActiveRow + int_n + 1
            .Offset(int_ActiveRow, 0).Value = "Option Value:"
            For i = 0 To int_n
                For j = 1 To int_n + 1
                .Offset(int_ActiveRow + j - 1, i + 1).Value = arr_Binomial(i, j, 2)
                Next j
            Next i

    End If
    End With

    wks_output.Calculate
    wks_output.Cells.HorizontalAlignment = xlCenter
    wks_output.Columns.AutoFit
End Sub