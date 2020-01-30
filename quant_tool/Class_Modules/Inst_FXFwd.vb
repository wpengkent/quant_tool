Option Explicit

' ## MEMBER DATA
' Components
Private scf_FlowA As SCF, scf_FlowB As SCF

' Variable dates
Private lng_ValDate As Long, lng_SpotDate As Long

' Static values
Private dic_GlobalStaticInfo As Dictionary, dic_CurveDependencies As Dictionary, map_Rules As MappingRules
Private str_CCY_FlowA As String, str_CCY_FlowB As String
Private str_CCY_PnL As String, int_Sign As Integer
Private Const bln_IsSpotDFInDV01 As Boolean = True


' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As InstParams_FXF, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing)

    ' Store static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    Set map_Rules = dic_GlobalStaticInfo(StaticInfoType.MappingRules)
    str_CCY_PnL = fld_ParamsInput.CCY_PnL
    If fld_ParamsInput.Pay_FlowA = True Then int_Sign = -1 Else int_Sign = 1
    str_CCY_FlowA = fld_ParamsInput.FlowA.CCY
    str_CCY_FlowB = fld_ParamsInput.FlowB.CCY

    Call SetValDate(fld_ParamsInput.ValueDate)

    ' Set up components
    Set scf_FlowA = New SCF
    Call scf_FlowA.Initialize(fld_ParamsInput.FlowA, dic_CurveSet, dic_GlobalStaticInfo)
    scf_FlowA.ZShiftsEnabled_DF = True
    scf_FlowA.ZShiftsEnabled_SpotDF = bln_IsSpotDFInDV01
    scf_FlowA.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

    Set scf_FlowB = New SCF
    Call scf_FlowB.Initialize(fld_ParamsInput.FlowB, dic_CurveSet, dic_GlobalStaticInfo)
    scf_FlowB.ZShiftsEnabled_DF = True
    scf_FlowB.ZShiftsEnabled_SpotDF = bln_IsSpotDFInDV01
    scf_FlowB.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

    ' Determine discount curve dependencies
    Set dic_CurveDependencies = scf_FlowA.CurveDependencies
    Set dic_CurveDependencies = Convert_MergeDicts(dic_CurveDependencies, scf_FlowB.CurveDependencies)

    ' Determine additional FX curve dependencies
    Dim dic_FXCurves As Dictionary: Set dic_FXCurves = map_Rules.Dict_FXCurveNames
    Dim str_Curve_PnL As String: str_Curve_PnL = dic_FXCurves(str_CCY_PnL)
    If dic_CurveDependencies.Exists(str_Curve_PnL) = False Then Call dic_CurveDependencies.Add(str_Curve_PnL, True)
End Sub


' ## PROPERTIES
Public Property Get marketvalue() As Double
    ' ## Get discounted value of future flows in the PnL currency
    marketvalue = CalcValue(ValType.MV)
End Property

Public Property Get Cash() As Double
    ' ## Get discounted value of past and present flows in the PnL currency
    Cash = CalcValue(ValType.Cash)
End Property

Public Property Get PnL() As Double
    PnL = CalcValue(ValType.PnL)
End Property


' ## METHODS - GREEKS
Public Function Calc_DV01(str_curve As String, Optional int_PillarIndex As Integer = 0) As Double
    ' ## Return DV01 sensitivity to the specified curve.  There is no impact on FX discounted spot
    Dim dbl_Output As Double
    If dic_CurveDependencies.Exists(str_curve) Then
        ' Store shifted values
        Dim dbl_Val_Up As Double, dbl_Val_Down As Double
        Call scf_FlowA.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        Call scf_FlowB.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        dbl_Val_Up = Me.marketvalue

        Call scf_FlowA.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        Call scf_FlowB.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        dbl_Val_Down = Me.marketvalue

        ' Clear temporary shifts
        Call scf_FlowA.SetCurveState(str_curve, CurveState_IRC.Final)
        Call scf_FlowB.SetCurveState(str_curve, CurveState_IRC.Final)

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
        ' Store shifted values
        Dim dbl_Val_Up As Double, dbl_Val_Down As Double, dbl_Val_Unch As Double
        Call scf_FlowA.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        Call scf_FlowB.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        dbl_Val_Up = Me.marketvalue

        Call scf_FlowA.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        Call scf_FlowB.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        dbl_Val_Down = Me.marketvalue

        ' Clear temporary shifts
        Call scf_FlowA.SetCurveState(str_curve, CurveState_IRC.Final)
        Call scf_FlowB.SetCurveState(str_curve, CurveState_IRC.Final)
        dbl_Val_Unch = Me.marketvalue

        ' Calculate by finite differencing and convert to PnL currency
        dbl_Output = dbl_Val_Up + dbl_Val_Down - 2 * dbl_Val_Unch
    Else
        dbl_Output = 0
    End If

    Calc_DV02 = dbl_Output
End Function


' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Sub SetValDate(lng_Input As Long)
    ' Set stored value date
    lng_ValDate = lng_Input
    lng_SpotDate = cyGetFXCrossSpotDate(str_CCY_FlowA, str_CCY_FlowB, lng_ValDate, dic_GlobalStaticInfo)
End Sub


' ## METHODS - SUPPORT
Private Function CalcValue(enu_ValType As ValType) As Double
    CalcValue = (scf_FlowA.CalcValue(lng_ValDate, lng_SpotDate, str_CCY_PnL, enu_ValType) _
        - scf_FlowB.CalcValue(lng_ValDate, lng_SpotDate, str_CCY_PnL, enu_ValType)) * int_Sign
End Function


' ## METHODS - CALCULATION DETAILS
Public Sub OutputReport(wks_output As Worksheet)
    wks_output.Cells.Clear

    Dim rng_OutputTopLeft As Range: Set rng_OutputTopLeft = wks_output.Range("A1")
    Dim rng_PnL As Range, rng_MV As Range, rng_Cash As Range
    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    Dim str_CurrencyFormat As String: str_CurrencyFormat = Gather_CurrencyFormat()
    Dim int_ActiveRow As Integer: int_ActiveRow = 0

    Dim dic_Addresses_A As Dictionary: Set dic_Addresses_A = New Dictionary
    dic_Addresses_A.CompareMode = CompareMethod.TextCompare
    Dim dic_Addresses_B As Dictionary: Set dic_Addresses_B = New Dictionary
    dic_Addresses_B.CompareMode = CompareMethod.TextCompare

    ' Output overall info
    With rng_OutputTopLeft
        .Offset(int_ActiveRow, 0).Value = "Value date:"
        .Offset(int_ActiveRow, 1).Value = lng_ValDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses_A.Add("ValDate", .Offset(int_ActiveRow, 1).Address(False, False))
        Call dic_Addresses_B.Add("ValDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Currency:"
        .Offset(int_ActiveRow, 1).Value = str_CCY_PnL
        Call dic_Addresses_A.Add("PnLCCY", .Offset(int_ActiveRow, 1).Address(False, False))
        Call dic_Addresses_B.Add("PnLCCY", .Offset(int_ActiveRow, 1).Address(False, False))

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

        ' Output individual leg info - leg A
        int_ActiveRow = int_ActiveRow + 3
        .Offset(int_ActiveRow, 0).Value = "LEG A"
        .Offset(int_ActiveRow, 0).Font.Italic = True

        int_ActiveRow = int_ActiveRow + 1
        Call scf_FlowA.OutputReport(rng_OutputTopLeft.Offset(int_ActiveRow, 0), "PV", _
            str_CCY_PnL, int_Sign, False, dic_Addresses_A, True)

        ' Output individual leg info - leg B
        int_ActiveRow = int_ActiveRow + 8
        .Offset(int_ActiveRow, 0).Value = "LEG B"
        .Offset(int_ActiveRow, 0).Font.Italic = True

        int_ActiveRow = int_ActiveRow + 1
        Call scf_FlowB.OutputReport(rng_OutputTopLeft.Offset(int_ActiveRow, 0), "PV", str_CCY_PnL, -int_Sign, False, _
            dic_Addresses_B, True)
    End With

    ' Calculate values
    rng_MV.Formula = "=IF(" & dic_Addresses_A("SCF_Type") & "=""MV""," & dic_Addresses_A("SCF_PV") & ",0)" & _
        "+IF(" & dic_Addresses_B("SCF_Type") & "=""MV""," & dic_Addresses_B("SCF_PV") & ",0)"
    rng_Cash.Formula = "=IF(" & dic_Addresses_A("SCF_Type") & "=""CASH""," & dic_Addresses_A("SCF_PV") & ",0)" & _
        "+IF(" & dic_Addresses_B("SCF_Type") & "=""CASH""," & dic_Addresses_B("SCF_PV") & ",0)"
    rng_PnL.Formula = "=" & rng_MV.Address(False, False) & "+" & rng_Cash.Address(False, False)

    wks_output.Calculate
    wks_output.Cells.HorizontalAlignment = xlCenter
    wks_output.Columns.AutoFit
End Sub