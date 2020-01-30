Option Explicit

' ## MEMBER DATA
' Dependent curves
Private fxs_Spots As Data_FXSpots, irc_Est As Data_IRCurve

' Static values
Private dic_GlobalStaticInfo As Dictionary, dic_CurveDependencies As Dictionary
Private fld_Params As InstParams_FTB
Private lng_FutMatSpot As Long, lng_UndMat_Generated As Long, lng_UndMat_Adjusted As Long
Private cal_pmt As Calendar
Private dbl_Time_Gen As Double, dbl_Time_Adj As Double, dbl_Time_DV01 As Double, dbl_SpreadToFut As Double
Private dbl_YearFrac_TickMode As Double, dbl_YearFrac_YieldMode As Double
Private int_Sign As Integer
Private dbl_Yield_Orig As Double, dbl_Yield_Mkt As Double


' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As InstParams_FTB, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing)

    ' Store static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    fld_Params = fld_ParamsInput
    Const str_Daycount_DV01 As String = "ACT/365"

    ' Set dependent curves
    If dic_CurveSet Is Nothing Then
        Set irc_Est = GetObject_IRCurve(fld_Params.Curve_Est, True, False, dic_GlobalStaticInfo)
        Set fxs_Spots = GetObject_FXSpots(True, dic_GlobalStaticInfo)
    Else
        Dim dic_IRCurves As Dictionary: Set dic_IRCurves = dic_CurveSet(CurveType.IRC)
        Set irc_Est = dic_IRCurves(fld_Params.Curve_Est)
        Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
    End If

    ' Store further static values
    If fld_Params.IsBuy = True Then int_Sign = 1 Else int_Sign = -1
    dbl_Yield_Orig = (100 - fld_Params.Price_Orig) / 100
    dbl_Yield_Mkt = (100 - fld_Params.Price_Mkt) / 100
    Dim cas_Found As CalendarSet: Set cas_Found = dic_GlobalStaticInfo(StaticInfoType.CalendarSet)
    cal_pmt = cas_Found.Lookup_Calendar(fld_Params.PmtCal)
    lng_FutMatSpot = date_workday(fld_Params.FutMat, fld_Params.SettleDays, cal_pmt.HolDates, cal_pmt.Weekends)
    lng_UndMat_Generated = date_addterm(lng_FutMatSpot, fld_Params.UndTerm, 1, True)

    If fld_Params.AdjUndMatDate = fld_Params.FutMat Then
        lng_UndMat_Adjusted = lng_UndMat_Generated  ' No override specified
    Else
        lng_UndMat_Adjusted = fld_Params.AdjUndMatDate
    End If

    Dim int_NumMonths As Integer: int_NumMonths = calc_nummonths(fld_Params.UndTerm)
    dbl_YearFrac_TickMode = int_NumMonths / 12
    dbl_YearFrac_YieldMode = calc_yearfrac(lng_FutMatSpot, lng_FutMatSpot + int_NumMonths * 30, fld_Params.Daycount, fld_Params.UndTerm, True)
    dbl_Time_Gen = calc_yearfrac(lng_FutMatSpot, lng_UndMat_Generated, fld_Params.Daycount, fld_Params.UndTerm, True)
    dbl_Time_Adj = calc_yearfrac(lng_FutMatSpot, lng_UndMat_Adjusted, fld_Params.Daycount, fld_Params.UndTerm, True)
    dbl_Time_DV01 = calc_yearfrac(lng_FutMatSpot, lng_UndMat_Generated, str_Daycount_DV01, fld_Params.UndTerm, True)

    If fld_Params.IsSpreadOn_PnL = True Or fld_Params.IsSpreadOn_DV01 = True Then
        Dim enu_CurveState_Prev As CurveState_IRC: enu_CurveState_Prev = irc_Est.CurveState
        Dim int_SensPillar_Prev As Integer: int_SensPillar_Prev = irc_Est.SensPillar
        Call irc_Est.SetCurveState(CurveState_IRC.original)
        dbl_SpreadToFut = (1 + dbl_Yield_Mkt * dbl_Time_Gen) * irc_Est.Lookup_Rate(lng_FutMatSpot, lng_UndMat_Adjusted, "DF", , , False) ^ (dbl_Time_Gen / dbl_Time_Adj)
        Call irc_Est.SetCurveState(enu_CurveState_Prev, int_SensPillar_Prev)
    End If

    ' Determine curve dependencies
    Set dic_CurveDependencies = fxs_Spots.Lookup_CurveDependencies(fld_Params.CCY_Notional, fld_Params.CCY_PnL)
    If dic_CurveDependencies.Exists(irc_Est.CurveName) = False Then Call dic_CurveDependencies.Add(irc_Est.CurveName, True)
End Sub


' ## PROPERTIES
Public Property Get PnL() As Double
    ' ## Get the value of the contract in the PnL currency
    PnL = CalcValue(fld_Params.IsSpreadOn_PnL)
End Property


' ## METHODS - GREEKS
Public Function Calc_DV01(str_curve As String, Optional int_PillarIndex As Integer = 0) As Double
    ' ## Return DV01 sensitivity to the specified curve
    Dim dbl_Output As Double: dbl_Output = 0
    If dic_CurveDependencies.Exists(str_curve) Then
        ' Remember original setting, then disable DV01 impact on discounted spot
        Dim bln_ZShiftsEnabled_DiscSpot As Boolean: bln_ZShiftsEnabled_DiscSpot = fxs_Spots.ZShiftsEnabled_DiscSpot
        fxs_Spots.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

        Dim dbl_Val_Up As Double, dbl_Val_Down As Double
        If irc_Est.CurveName = str_curve Then
            Call irc_Est.SetCurveState(CurveState_IRC.Zero_Up1BP, int_PillarIndex)
            dbl_Val_Up = CalcValue(fld_Params.IsSpreadOn_DV01)

            Call irc_Est.SetCurveState(CurveState_IRC.Zero_Down1BP, int_PillarIndex)
            dbl_Val_Down = CalcValue(fld_Params.IsSpreadOn_DV01)

            Call irc_Est.SetCurveState(CurveState_IRC.Final)
            dbl_Output = (dbl_Val_Up - dbl_Val_Down) / 2
        End If

        ' Revert to original setting
        fxs_Spots.ZShiftsEnabled_DiscSpot = bln_ZShiftsEnabled_DiscSpot
    End If

    Calc_DV01 = dbl_Output
End Function

Public Function Calc_DV02(str_curve As String, Optional int_PillarIndex As Integer = 0) As Double
    ' ## Return second order sensitivity to the specified curve
    ' Remember original setting, then disable DV01 impact on discounted spot
    Dim dbl_Output As Double: dbl_Output = 0
    If dic_CurveDependencies.Exists(str_curve) Then
        Dim bln_ZShiftsEnabled_DiscSpot As Boolean: bln_ZShiftsEnabled_DiscSpot = fxs_Spots.ZShiftsEnabled_DiscSpot
        fxs_Spots.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

        Dim dbl_Val_Up As Double, dbl_Val_Down As Double, dbl_Val_Unch As Double
        If irc_Est.CurveName = str_curve Then
            Call irc_Est.SetCurveState(CurveState_IRC.Zero_Up1BP, int_PillarIndex)
            dbl_Val_Up = CalcValue(fld_Params.IsSpreadOn_DV01)

            Call irc_Est.SetCurveState(CurveState_IRC.Zero_Down1BP, int_PillarIndex)
            dbl_Val_Down = CalcValue(fld_Params.IsSpreadOn_DV01)

            Call irc_Est.SetCurveState(CurveState_IRC.Final)
            dbl_Val_Unch = CalcValue(fld_Params.IsSpreadOn_DV01)

            dbl_Output = dbl_Val_Up + dbl_Val_Down - 2 * dbl_Val_Unch
        End If

        fxs_Spots.ZShiftsEnabled_DiscSpot = bln_ZShiftsEnabled_DiscSpot
    End If

    Calc_DV02 = dbl_Output
End Function


' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Sub SetValDate(lng_Input As Long)
    ' ## No dependency on time
End Sub


' ## METHODS - INTERMEDIATE CALCULATIONS
Private Function GetTheoFutYield(bln_WithSpread As Boolean) As Double
    ' ## Get annualized simple yield, optionally including a spread to bring the base case in line with the market
    Dim dbl_SpreadToUse As Double
    If bln_WithSpread = True Then dbl_SpreadToUse = dbl_SpreadToFut Else dbl_SpreadToUse = 1

    GetTheoFutYield = (irc_Est.Lookup_Rate(lng_FutMatSpot, lng_UndMat_Adjusted, "DF", , , False) ^ (-dbl_Time_Gen / dbl_Time_Adj) * dbl_SpreadToUse - 1) / dbl_Time_Gen
End Function

Private Function GetUnitVal(dbl_Yield_Current As Double, dbl_Yield_Base As Double) As Double
    Dim dbl_Output As Double

    ' Calculate value per unit
    Select Case fld_Params.MeasureMode
        Case "TICK VALUE"
            dbl_Output = (dbl_Yield_Base - dbl_Yield_Current) * dbl_YearFrac_TickMode
        Case "YIELD BASIS"
            dbl_Output = 1 / (1 + dbl_Yield_Current * dbl_YearFrac_YieldMode) - 1 / (1 + dbl_Yield_Base * dbl_YearFrac_YieldMode)
    End Select

    GetUnitVal = dbl_Output
End Function

Private Function CalcValue(bln_SpreadOn As Boolean) As Double
    ' ## Get the value of the contract in the PnL currency
    Dim dbl_TheoFutYield As Double: dbl_TheoFutYield = GetTheoFutYield(bln_SpreadOn)
    Dim dbl_ValPerUnit As Double: dbl_ValPerUnit = GetUnitVal(dbl_TheoFutYield, dbl_Yield_Orig)
    Dim dbl_DiscFXSpot As Double: dbl_DiscFXSpot = fxs_Spots.Lookup_DiscSpot(fld_Params.CCY_Notional, fld_Params.CCY_PnL)
    CalcValue = fld_Params.Notional * dbl_ValPerUnit * int_Sign * dbl_DiscFXSpot
End Function


' ## METHODS - CALCULATION DETAILS
Public Sub OutputReport(wks_output As Worksheet)
    wks_output.Cells.Clear

    Dim rng_OutputTopLeft As Range: Set rng_OutputTopLeft = wks_output.Range("A1")
    Dim dic_Addresses As New Dictionary, rng_PnL As Range
    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    Dim str_CurrencyFormat As String: str_CurrencyFormat = Gather_CurrencyFormat()
    Dim int_ActiveRow As Integer: int_ActiveRow = 0

    With rng_OutputTopLeft
        .Offset(int_ActiveRow, 0).Value = "Value date:"
        .Offset(int_ActiveRow, 1).Value = fld_Params.ValueDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("ValDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "PnL:"
        .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
        .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
        Set rng_PnL = .Offset(int_ActiveRow, 1)

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Currency:"
        .Offset(int_ActiveRow, 1).Value = fld_Params.CCY_PnL
        Call dic_Addresses.Add("PnLCCY", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 2
        .Offset(int_ActiveRow, 0).Value = "Notional:"
        .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
        .Offset(int_ActiveRow, 1).Value = fld_Params.Notional
        Call dic_Addresses.Add("Notional", .Offset(int_ActiveRow, 1).Address(False, False))
        .Offset(int_ActiveRow, 2).Value = fld_Params.CCY_Notional
        Call dic_Addresses.Add("NotionalCCY", .Offset(int_ActiveRow, 2).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Position:"
        If fld_Params.IsBuy = True Then .Offset(int_ActiveRow, 1).Value = "B" Else .Offset(int_ActiveRow, 1).Value = "S"
        Call dic_Addresses.Add("Position", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Bill Start:"
        .Offset(int_ActiveRow, 1).Value = lng_FutMatSpot
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("BillStart", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Bill End:"
        .Offset(int_ActiveRow, 1).Value = lng_UndMat_Generated
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("BillEnd_Gen", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Bill Est End:"
        .Offset(int_ActiveRow, 1).Value = lng_UndMat_Adjusted
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("BillEnd_Adj", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Term:"
        .Offset(int_ActiveRow, 1).Formula = "=Calc_YearFrac(" & dic_Addresses("BillStart") & "," & dic_Addresses("BillEnd_Gen") _
            & ",""" & fld_Params.Daycount & """,""" & fld_Params.UndTerm & """,True)"
        Call dic_Addresses.Add("Term_Gen", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Est Term:"
        .Offset(int_ActiveRow, 1).Formula = "=Calc_YearFrac(" & dic_Addresses("BillStart") & "," & dic_Addresses("BillEnd_Adj") _
            & ",""" & fld_Params.Daycount & """,""" & fld_Params.UndTerm & """,True)"
        Call dic_Addresses.Add("Term_Adj", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Spread Factor:"
        If fld_Params.IsSpreadOn_PnL = True Then
            .Offset(int_ActiveRow, 1).Value = dbl_SpreadToFut
        Else
            .Offset(int_ActiveRow, 1).Value = 1
        End If
        Call dic_Addresses.Add("Spread", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Current Price:"
        .Offset(int_ActiveRow, 1).Formula = "=100*(1-(cyReadIRCurve(""" & fld_Params.Curve_Est & """," & dic_Addresses("BillStart") _
            & "," & dic_Addresses("BillEnd_Adj") & ",""DF"")^(-" & dic_Addresses("Term_Gen") & "/" & dic_Addresses("Term_Adj") _
            & ")*" & dic_Addresses("Spread") & "-1)/Calc_YearFrac(" & dic_Addresses("BillStart") & "," & dic_Addresses("BillEnd_Gen") _
            & ",""" & fld_Params.Daycount & """,""" & fld_Params.UndTerm & """,True))"
        Call dic_Addresses.Add("CurrentPrice", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Entry Price:"
        .Offset(int_ActiveRow, 1).Value = 100 * (1 - dbl_Yield_Orig)
        Call dic_Addresses.Add("EntryPrice", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Calc Term:"
        Call dic_Addresses.Add("Term_MV", .Offset(int_ActiveRow, 1).Address(False, False))
        Select Case fld_Params.MeasureMode
            Case "TICK VALUE"
                .Offset(int_ActiveRow, 1).Formula = "=" & calc_nummonths(fld_Params.UndTerm) & "/12"
            Case "YIELD BASIS"
                .Offset(int_ActiveRow, 1).Formula = "=Calc_YearFrac(" & dic_Addresses("BillStart") & "," & dic_Addresses("BillStart") _
                & "+" & calc_nummonths(fld_Params.UndTerm) & "*30,""" & fld_Params.Daycount & """,""" & fld_Params.UndTerm & """,True)"
        End Select

        ' Fill in formula for PnL
        Select Case fld_Params.MeasureMode
            Case "TICK VALUE"
                rng_PnL.Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" & dic_Addresses("Notional") _
                & "*(" & dic_Addresses("CurrentPrice") & "-" & dic_Addresses("EntryPrice") & ")/100*" & dic_Addresses("Term_MV") _
                & "*cyGetFXDiscSpot(" & dic_Addresses("NotionalCCY") & "," & dic_Addresses("PnLCCY") & ")"
            Case "YIELD BASIS"
                rng_PnL.Formula = "=IF(" & dic_Addresses("Position") & "=""S"",-1,1)*" & dic_Addresses("Notional") _
                & "*(1/(1+(1-" & dic_Addresses("CurrentPrice") & "/100)*" & dic_Addresses("Term_MV") & ")-1/(1+(1-" & dic_Addresses("EntryPrice") _
                & "/100)*" & dic_Addresses("Term_MV") & "))*cyGetFXDiscSpot(" & dic_Addresses("NotionalCCY") & "," & dic_Addresses("PnLCCY") & ")"
        End Select
    End With

    wks_output.Calculate
    wks_output.Cells.HorizontalAlignment = xlCenter
    wks_output.Columns.AutoFit
End Sub