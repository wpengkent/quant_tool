Option Explicit

' ## MEMBER DATA
' Components
Private irs_SinglePeriod As Inst_IRSwap

' Static values
Private dic_GlobalStaticInfo As Dictionary


' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As InstParams_FRA, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing)
    Dim fld_SwapParams As InstParams_IRS, lng_fixingdate As Long, dic_Fixings As Dictionary

    ' Store static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    Dim cas_Calendars As CalendarSet: Set cas_Calendars = dic_GlobalStaticInfo(StaticInfoType.CalendarSet)

    ' Gather generator level information
    Dim igs_Generators As IRGeneratorSet: Set igs_Generators = dic_GlobalStaticInfo(StaticInfoType.IRGeneratorSet)
    Dim fld_FloatLegParams As IRLegParams: fld_FloatLegParams = igs_Generators.Lookup_Generator(fld_ParamsInput.Generator)
    Dim fld_FixedLegParams As IRLegParams: fld_FixedLegParams = fld_FloatLegParams

    ' Initialize fixed leg of the underlying single period swap
    With fld_FixedLegParams
        .ValueDate = fld_ParamsInput.ValueDate
        .Swapstart = fld_ParamsInput.StartDate
        .GenerationRefPoint = fld_ParamsInput.StartDate
        .IsFwdGeneration = True
        .Term = .PmtFreq
        .PExch_Start = False
        .PExch_Intermediate = False
        .PExch_End = False
        .FloatEst = True
        .ForceToMV = False
        .Notional = fld_ParamsInput.Notional
        .index = "-"
        .RateOrMargin = fld_ParamsInput.Rate
        .IsUniformPeriods = False
        .estcal = "-"
        .Curve_Est = "-"
    End With

    ' Initialize floating leg of the underlying single period swap
    With fld_FloatLegParams
        .ValueDate = fld_ParamsInput.ValueDate
        .Swapstart = fld_ParamsInput.StartDate
        .GenerationRefPoint = fld_ParamsInput.StartDate
        .IsFwdGeneration = True
        .Term = .PmtFreq
        .PExch_Start = False
        .PExch_Intermediate = False
        .PExch_End = False
        .FloatEst = True
        .ForceToMV = False
        .Notional = fld_ParamsInput.Notional
        .RateOrMargin = 0

        ' Store fixing if supplied
        If fld_ParamsInput.Fixing <> "-" Then
            lng_fixingdate = Date_NextCoupon(.Swapstart, .Term, cas_Calendars.Lookup_Calendar(.PmtCal), 1, .EOM, .BDC)
            Set dic_Fixings = New Dictionary
            Call dic_Fixings.Add(lng_fixingdate, CDbl(fld_ParamsInput.Fixing))
            Set .Fixings = dic_Fixings
        End If
    End With

    ' Initialize remaining parameters and create the swap
    With fld_SwapParams
        .CCY_PnL = fld_ParamsInput.CCY_PnL
        .Pay_LegA = fld_ParamsInput.IsBuy
        .LegA = fld_FixedLegParams
        .LegB = fld_FloatLegParams
    End With

    Set irs_SinglePeriod = New Inst_IRSwap
    Call irs_SinglePeriod.Initialize(fld_SwapParams, dic_CurveSet, dic_GlobalStaticInfo)
End Sub


' ## PROPERTIES
Public Property Get marketvalue() As Double
    ' ## Get discounted value of future flows in the PnL currency
    marketvalue = irs_SinglePeriod.marketvalue
End Property

Public Property Get Cash() As Double
    ' ## Get discounted value of past and present flows in the PnL currency
    Cash = irs_SinglePeriod.Cash
End Property

Public Property Get PnL() As Double
    PnL = Me.marketvalue + Me.Cash
End Property


' ## METHODS - GREEKS
Public Function Calc_DV01(str_curve As String, Optional int_PillarIndex As Integer = 0) As Double
    Calc_DV01 = irs_SinglePeriod.Calc_DV01(str_curve, int_PillarIndex)
End Function

Public Function Calc_DV02(str_curve As String, Optional int_PillarIndex As Integer = 0) As Double
    Calc_DV02 = irs_SinglePeriod.Calc_DV02(str_curve, int_PillarIndex)
End Function


' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Sub HandleUpdate_IRC(str_CurveName As String)
    ' ## Update stored values affected by change in the curve values
    Call irs_SinglePeriod.HandleUpdate_IRC(str_CurveName)
End Sub

Public Sub SetValDate(lng_Input As Long)
    ' ## Set stored value date and refresh values dependent on the value date
    Call irs_SinglePeriod.SetValDate(lng_Input)
End Sub


' ## METHODS - CALCULATION DETAILS
Public Sub OutputReport(wks_output As Worksheet)
    wks_output.Cells.Clear

    Dim rng_OutputTopLeft As Range: Set rng_OutputTopLeft = wks_output.Range("A1")
    Dim str_Address_PnLCCY As String, str_Address_ValDate As String
    Dim rng_PnL As Range, rng_MV As Range, rng_Cash As Range
    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    Dim str_CurrencyFormat As String: str_CurrencyFormat = Gather_CurrencyFormat()
    Dim str_CCY_PnL As String: str_CCY_PnL = irs_SinglePeriod.PnLCurrency
    Dim int_ActiveRow As Integer: int_ActiveRow = 0

    ' Output overall info
    With rng_OutputTopLeft
        .Offset(int_ActiveRow, 0).Value = "Value date:"
        .Offset(int_ActiveRow, 1).Value = irs_SinglePeriod.LegA.Params.ValueDate
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
    Call irs_SinglePeriod.LegA.OutputReport_Swap(rng_OutputTopLeft.Offset(int_ActiveRow, 0), "Leg A", irs_SinglePeriod.IsPayLegA, str_Address_PnLCCY, _
        str_Address_ValDate, str_CCY_PnL, str_Address_MVA, str_Address_CashA)

    int_ActiveRow = int_ActiveRow + 12
    Call irs_SinglePeriod.LegB.OutputReport_Swap(rng_OutputTopLeft.Offset(int_ActiveRow, 0), "Leg B", Not irs_SinglePeriod.IsPayLegA, str_Address_PnLCCY, _
        str_Address_ValDate, str_CCY_PnL, str_Address_MVB, str_Address_CashB)

    ' Calculate values
    rng_MV.Formula = "=" & str_Address_MVA & "+" & str_Address_MVB
    rng_Cash.Formula = "=" & str_Address_CashA & "+" & str_Address_CashB
    rng_PnL.Formula = "=" & rng_MV.Address(False, False) & "+" & rng_Cash.Address(False, False)

    wks_output.Calculate
    wks_output.Cells.HorizontalAlignment = xlCenter
    wks_output.Columns.AutoFit
End Sub
