Option Explicit

' ## MEMBER DATA
' Components
Private irl_LegA As IRLeg, irl_legB As IRLeg

' Dependent curves
Private fxs_Spots As Data_FXSpots

' Static values
Private dic_GlobalStaticInfo As Dictionary, dic_CurveDependencies As Dictionary
Private str_CCY_PnL As String, int_Sign As Integer


' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As InstParams_IRS, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing)

    ' Store static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    str_CCY_PnL = fld_ParamsInput.CCY_PnL
    If fld_ParamsInput.Pay_LegA = True Then int_Sign = -1 Else int_Sign = 1

    ' Store components
    Set irl_LegA = New IRLeg
    Call irl_LegA.Initialize(fld_ParamsInput.LegA, dic_CurveSet, dic_StaticInfoInput)

    Set irl_legB = New IRLeg
    Call irl_legB.Initialize(fld_ParamsInput.LegB, dic_CurveSet, dic_StaticInfoInput)

    ' Store dependent curves
    If dic_CurveSet Is Nothing Then
        Set fxs_Spots = GetObject_FXSpots(True)
    Else
        Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
    End If

    ' Determine curve dependencies
    Set dic_CurveDependencies = irl_LegA.CurveDependencies
    Set dic_CurveDependencies = Convert_MergeDicts(dic_CurveDependencies, irl_legB.CurveDependencies)
    Set dic_CurveDependencies = Convert_MergeDicts(dic_CurveDependencies, fxs_Spots.Lookup_CurveDependencies(fld_ParamsInput.CCY_PnL))
End Sub


' ## PROPERTIES
Public Property Get marketvalue() As Double
    marketvalue = CalcValue(ValType.MV)
End Property

Public Property Get Cash() As Double
    Cash = CalcValue(ValType.Cash)
End Property

Public Property Get PnL() As Double
    PnL = marketvalue + Cash
End Property

Public Property Get LegA() As IRLeg
    Set LegA = irl_LegA
End Property

Public Property Get LegB() As IRLeg
    Set LegB = irl_legB
End Property

Public Property Get ParRate_LegA() As Double
    ' ## Find rate for leg A which makes NPV equal to zero
    ParRate_LegA = irl_LegA.SolveParRate(irl_legB)
End Property

Public Property Get ParRate_LegB() As Double
    ' ## Find rate for leg B which makes NPV equal to zero
    ParRate_LegB = irl_legB.SolveParRate(irl_LegA)
End Property

Public Property Get PnLCurrency() As String
    PnLCurrency = str_CCY_PnL
End Property

Public Property Get IsPayLegA() As Boolean
    IsPayLegA = (int_Sign = -1)
End Property

Public Function DependsOnFuture(str_BootstrappedCurve As String, lng_MatDate) As Boolean
    ' ## If estimation depends on the bootstrapped curve and the last estimation date is beyond the maturity date, return True
    DependsOnFuture = irl_LegA.DependsOnFuture(str_BootstrappedCurve, lng_MatDate) _
        Or irl_legB.DependsOnFuture(str_BootstrappedCurve, lng_MatDate)
End Function


' ## METHODS - GREEKS
Public Function Calc_DV01(str_curve As String, Optional int_PillarIndex As Integer = 0) As Double
    ' ## Return DV01 sensitivity to the specified curve.  There is no impact on FX discounted spot
    Dim dbl_Output As Double
    If dic_CurveDependencies.Exists(str_curve) Then
        Dim bln_ZShiftsEnabled_DiscSpot As Boolean: bln_ZShiftsEnabled_DiscSpot = fxs_Spots.ZShiftsEnabled_DiscSpot
        fxs_Spots.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

        ' Store shifted values
        Dim dbl_Val_Up As Double, dbl_Val_Down As Double
        Call irl_LegA.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        Call irl_legB.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        dbl_Val_Up = Me.marketvalue

        Call irl_LegA.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        Call irl_legB.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        dbl_Val_Down = Me.marketvalue

        ' Clear temporary shifts
        Call irl_LegA.SetCurveState(str_curve, CurveState_IRC.Final)
        Call irl_legB.SetCurveState(str_curve, CurveState_IRC.Final)

        ' Revert to original settings
        fxs_Spots.ZShiftsEnabled_DiscSpot = bln_ZShiftsEnabled_DiscSpot

        ' Calculate by finite differencing and convert to PnL currency
        dbl_Output = (dbl_Val_Up - dbl_Val_Down) / 2
    Else
        dbl_Output = 0
    End If

    Calc_DV01 = dbl_Output
End Function

Public Function Calc_DV02(str_curve As String, Optional int_PillarIndex As Integer = 0) As Double
    ' ## Return second order sensitivity to the specified curve.  There is no impact on FX discounted spot
    Dim dbl_Output As Double
    If dic_CurveDependencies.Exists(str_curve) Then
        Dim bln_ZShiftsEnabled_DiscSpot As Boolean: bln_ZShiftsEnabled_DiscSpot = fxs_Spots.ZShiftsEnabled_DiscSpot
        fxs_Spots.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

        ' Store shifted values
        Dim dbl_Val_Up As Double, dbl_Val_Down As Double, dbl_Val_Unch As Double
        Call irl_LegA.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        Call irl_legB.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        dbl_Val_Up = Me.marketvalue

        Call irl_LegA.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        Call irl_legB.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        dbl_Val_Down = Me.marketvalue

        ' Clear temporary shifts
        Call irl_LegA.SetCurveState(str_curve, CurveState_IRC.Final)
        Call irl_legB.SetCurveState(str_curve, CurveState_IRC.Final)
        dbl_Val_Unch = Me.marketvalue

        ' Revert to original settings
        fxs_Spots.ZShiftsEnabled_DiscSpot = bln_ZShiftsEnabled_DiscSpot

        ' Calculate by finite differencing and convert to PnL currency
        dbl_Output = dbl_Val_Up + dbl_Val_Down - 2 * dbl_Val_Unch
    Else
        dbl_Output = 0
    End If

    Calc_DV02 = dbl_Output
End Function


' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Sub SetValDate(lng_Input As Long)
    ' ## Set stored value date and refresh values dependent on the value date
    Call irl_LegA.SetValDate(lng_Input)
    Call irl_legB.SetValDate(lng_Input)
End Sub

Public Sub ReplaceCurveObject(str_CurveName As String, irc_Curve As Data_IRCurve)
    ' ## If any curve names match the name specified, replace with the specified curve object
    ' ## Used for bootstrapping procedure to ensure the curve underlying the swap is the same as the curve being updated by the process
    Call irl_LegA.ReplaceCurveObject(str_CurveName, irc_Curve)
    Call irl_legB.ReplaceCurveObject(str_CurveName, irc_Curve)
End Sub

Public Sub HandleUpdate_IRC(str_CurveName As String)
    ' ## Update stored values affected by change in the curve values
    Call irl_LegA.HandleUpdate_IRC(str_CurveName)
    Call irl_legB.HandleUpdate_IRC(str_CurveName)
End Sub


' ## METHODS - PRIVATE
Private Function CalcValue(enu_ValType As ValType) As Double
    ' ## Calculate either the MV or cash for the swap

    Dim dbl_FXConv_LegA As Double, dbl_FXConv_LegB As Double
    Dim dbl_Val_LegA As Double, dbl_Val_LegB As Double

    ' Find FX conversion factors to translate to PnL currency
    dbl_FXConv_LegA = fxs_Spots.Lookup_DiscSpot(irl_LegA.Params.CCY, str_CCY_PnL)
    dbl_FXConv_LegB = fxs_Spots.Lookup_DiscSpot(irl_legB.Params.CCY, str_CCY_PnL)

    Select Case enu_ValType
        Case ValType.MV
            dbl_Val_LegA = irl_LegA.marketvalue
            dbl_Val_LegB = irl_legB.marketvalue
        Case ValType.Cash
            dbl_Val_LegA = irl_LegA.Cash
            dbl_Val_LegB = irl_legB.Cash
    End Select

    ' Calculate NPV
    CalcValue = int_Sign * (dbl_Val_LegA * dbl_FXConv_LegA - dbl_Val_LegB * dbl_FXConv_LegB)
End Function


' ## METHODS - CALCULATION DETAILS
Public Sub OutputReport(wks_output As Worksheet)
    wks_output.Cells.Clear

    Dim rng_OutputTopLeft As Range: Set rng_OutputTopLeft = wks_output.Range("A1")
    Dim str_Address_PnLCCY As String, str_Address_ValDate As String
    Dim rng_PnL As Range, rng_MV As Range, rng_Cash As Range
    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    Dim str_CurrencyFormat As String: str_CurrencyFormat = Gather_CurrencyFormat()
    Dim int_ActiveRow As Integer: int_ActiveRow = 0

    ' Output overall info
    With rng_OutputTopLeft
        .Offset(int_ActiveRow, 0).Value = "Value date:"
        .Offset(int_ActiveRow, 1).Value = irl_LegA.Params.ValueDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        str_Address_ValDate = .Offset(int_ActiveRow, 1).Address(False, False)

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Currency:"
        .Offset(int_ActiveRow, 1).Value = str_CCY_PnL
        str_Address_PnLCCY = .Offset(int_ActiveRow, 1).Address(False, False)

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "MV:"
        .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
        .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
        Set rng_MV = .Offset(int_ActiveRow, 1)

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Cash:"
        .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
        .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
        Set rng_Cash = .Offset(int_ActiveRow, 1)

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "PnL:"
        .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
        .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
        Set rng_PnL = .Offset(int_ActiveRow, 1)
    End With

    ' Output individual leg info
    int_ActiveRow = int_ActiveRow + 3
    Dim str_Address_MVA As String, str_Address_MVB As String
    Dim str_Address_CashA As String, str_Address_CashB As String
    Call irl_LegA.OutputReport_Swap(rng_OutputTopLeft.Offset(int_ActiveRow, 0), "Leg A", (int_Sign = -1), str_Address_PnLCCY, _
        str_Address_ValDate, str_CCY_PnL, str_Address_MVA, str_Address_CashA)
    Call irl_legB.OutputReport_Swap(rng_OutputTopLeft.Offset(int_ActiveRow, 14), "Leg B", (int_Sign = 1), str_Address_PnLCCY, _
        str_Address_ValDate, str_CCY_PnL, str_Address_MVB, str_Address_CashB)

    ' Calculate values
    rng_MV.Formula = "=" & str_Address_MVA & "+" & str_Address_MVB
    rng_Cash.Formula = "=" & str_Address_CashA & "+" & str_Address_CashB
    rng_PnL.Formula = "=" & rng_MV.Address(False, False) & "+" & rng_Cash.Address(False, False)

    wks_output.Calculate
    wks_output.Cells.HorizontalAlignment = xlCenter
    wks_output.Columns.AutoFit
End Sub