Option Explicit

' ## MEMBER DATA
' Components
Private scf_InitialFlow As SCF, scf_FinalFlow As SCF

' Variable dates
Private lng_ValDate As Long

' Static values
Private dic_GlobalStaticInfo As Dictionary, dic_CurveDependencies As Dictionary, map_Rules As MappingRules
Private str_CCY_PnL As String, int_Sign_Initial As Integer, int_Sign_Final As Integer


' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As InstParams_DEP, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing)

    ' Store static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    str_CCY_PnL = fld_ParamsInput.CCY_PnL

    lng_ValDate = fld_ParamsInput.ValueDate

    ' Determine whether loan or deposit
    Dim int_Sign As Integer
    If fld_ParamsInput.IsLoan = True Then
        int_Sign_Initial = 1
        int_Sign_Final = -1
    Else
        int_Sign_Initial = -1
        int_Sign_Final = 1
    End If

    ' Set up initial flow
    Dim fld_InitialFlowParams As SCFParams
    With fld_InitialFlowParams
        If fld_ParamsInput.PExch = True Then .Amount = fld_ParamsInput.Principal Else .Amount = 0

        .CCY = fld_ParamsInput.CCY_Principal
        .PmtDate = fld_ParamsInput.StartDate
        .Curve_Disc = fld_ParamsInput.Curve_Disc
        .Curve_SpotDisc = .Curve_Disc
    End With

    Set scf_InitialFlow = New SCF
    Call scf_InitialFlow.Initialize(fld_InitialFlowParams, dic_CurveSet, dic_GlobalStaticInfo)
    scf_InitialFlow.ZShiftsEnabled_DF = False
    scf_InitialFlow.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

    ' Set up final flow
    Dim fld_FinalFlowParams As SCFParams
    With fld_FinalFlowParams
        .Amount = fld_ParamsInput.Principal * fld_ParamsInput.Rate / 100 * calc_yearfrac(fld_ParamsInput.StartDate, _
            fld_ParamsInput.MatDate, fld_ParamsInput.Daycount)
        If fld_ParamsInput.PExch = True Then .Amount = .Amount + fld_ParamsInput.Principal

        .CCY = fld_ParamsInput.CCY_Principal
        .PmtDate = fld_ParamsInput.MatDate
        .Curve_Disc = fld_ParamsInput.Curve_Disc
        .Curve_SpotDisc = .Curve_Disc
    End With

    Set scf_FinalFlow = New SCF
    Call scf_FinalFlow.Initialize(fld_FinalFlowParams, dic_CurveSet, dic_GlobalStaticInfo)
    scf_FinalFlow.ZShiftsEnabled_DF = True
    scf_FinalFlow.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

    ' Determine discount curve dependencies
    Set dic_CurveDependencies = scf_InitialFlow.CurveDependencies
    Set dic_CurveDependencies = Convert_MergeDicts(dic_CurveDependencies, scf_FinalFlow.CurveDependencies)

    ' Determine additional FX curve dependencies
    Set map_Rules = dic_GlobalStaticInfo(StaticInfoType.MappingRules)
    Dim dic_FXCurves As Dictionary: Set dic_FXCurves = map_Rules.Dict_FXCurveNames
    Dim str_Curve_PnL As String: str_Curve_PnL = dic_FXCurves(str_CCY_PnL)
    If dic_CurveDependencies.Exists(str_Curve_PnL) = False Then Call dic_CurveDependencies.Add(str_Curve_PnL, True)
End Sub


' ## PROPERTIES
Public Property Get marketvalue() As Double
    ' ## Get discounted value of future flows in the PnL currency
    marketvalue = GetValue(ValType.MV)
End Property

Public Property Get Cash() As Double
    ' ## Get discounted value of past and present flows in the PnL currency
    Cash = GetValue(ValType.Cash)
End Property

Public Property Get PnL() As Double
    PnL = GetValue(ValType.PnL)
End Property


' ## METHODS - GREEKS
Public Function Calc_DV01(str_curve As String, Optional int_PillarIndex As Integer = 0) As Double
    ' ## Return DV01 sensitivity to the specified curve.  There is no impact on FX discounted spot
    Dim dbl_Output As Double
    If dic_CurveDependencies.Exists(str_curve) Then
        ' Store shifted values
        Dim dbl_Val_Up As Double, dbl_Val_Down As Double

        Call scf_InitialFlow.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        Call scf_FinalFlow.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        dbl_Val_Up = Me.marketvalue

        Call scf_InitialFlow.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        Call scf_FinalFlow.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        dbl_Val_Down = Me.marketvalue

        ' Clear temporary shifts
        Call scf_InitialFlow.SetCurveState(str_curve, CurveState_IRC.Final)
        Call scf_FinalFlow.SetCurveState(str_curve, CurveState_IRC.Final)

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
        Call scf_InitialFlow.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        Call scf_FinalFlow.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        dbl_Val_Up = Me.marketvalue

        Call scf_InitialFlow.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        Call scf_FinalFlow.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        dbl_Val_Down = Me.marketvalue

        ' Clear temporary shifts
        Call scf_InitialFlow.SetCurveState(str_curve, CurveState_IRC.Final)
        Call scf_FinalFlow.SetCurveState(str_curve, CurveState_IRC.Final)
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
End Sub


' ## METHODS - CALCULATION
Private Function GetValue(enu_ValType As ValType) As Double
    GetValue = scf_InitialFlow.CalcValue(lng_ValDate, lng_ValDate, str_CCY_PnL, enu_ValType) * int_Sign_Initial _
        + scf_FinalFlow.CalcValue(lng_ValDate, lng_ValDate, str_CCY_PnL, enu_ValType) * int_Sign_Final
End Function


' ## METHODS - CALCULATION DETAILS
Public Sub OutputReport(wks_output As Worksheet)
    wks_output.Cells.Clear

    Dim rng_OutputTopLeft As Range: Set rng_OutputTopLeft = wks_output.Range("A1")
    Dim str_Address_InitialPV As String, str_Address_FinalPV As String, str_Address_InitialType As String, str_Address_FinalType As String
    Dim rng_PnL As Range, rng_MV As Range, rng_Cash As Range
    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    Dim str_CurrencyFormat As String: str_CurrencyFormat = Gather_CurrencyFormat()
    Dim int_ActiveRow As Integer: int_ActiveRow = 0

    Dim dic_Addresses_Initial As Dictionary: Set dic_Addresses_Initial = New Dictionary
    dic_Addresses_Initial.CompareMode = CompareMethod.TextCompare
    Dim dic_Addresses_Final As Dictionary: Set dic_Addresses_Final = New Dictionary
    dic_Addresses_Final.CompareMode = CompareMethod.TextCompare

    ' Output overall info
    With rng_OutputTopLeft
        .Offset(int_ActiveRow, 0).Value = "Value date:"
        .Offset(int_ActiveRow, 1).Value = lng_ValDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses_Initial.Add("ValDate", .Offset(int_ActiveRow, 1).Address(False, False))
        Call dic_Addresses_Final.Add("ValDate", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Currency:"
        .Offset(int_ActiveRow, 1).Value = str_CCY_PnL
        Call dic_Addresses_Initial.Add("PnLCCY", .Offset(int_ActiveRow, 1).Address(False, False))
        Call dic_Addresses_Final.Add("PnLCCY", .Offset(int_ActiveRow, 1).Address(False, False))

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

        ' Output individual leg info
        int_ActiveRow = int_ActiveRow + 3
        .Offset(int_ActiveRow, 0).Value = "INITIAL LEG"
        .Offset(int_ActiveRow, 0).Font.Italic = True

        int_ActiveRow = int_ActiveRow + 1
        Call scf_InitialFlow.OutputReport(rng_OutputTopLeft.Offset(int_ActiveRow, 0), "PV", str_CCY_PnL, int_Sign_Initial, _
            False, dic_Addresses_Initial, True)

        int_ActiveRow = int_ActiveRow + 8
        .Offset(int_ActiveRow, 0).Value = "FINAL LEG"
        .Offset(int_ActiveRow, 0).Font.Italic = True

        int_ActiveRow = int_ActiveRow + 1
        Call scf_FinalFlow.OutputReport(rng_OutputTopLeft.Offset(int_ActiveRow, 0), "PV", str_CCY_PnL, int_Sign_Final, _
            False, dic_Addresses_Final, True)
    End With

    ' Calculate values
    rng_MV.Formula = "=IF(" & dic_Addresses_Initial("SCF_Type") & "=""MV""," & dic_Addresses_Initial("SCF_PV") & ",0)" & _
        "+IF(" & dic_Addresses_Final("SCF_Type") & "=""MV""," & dic_Addresses_Final("SCF_PV") & ",0)"
    rng_Cash.Formula = "=IF(" & dic_Addresses_Initial("SCF_Type") & "=""CASH""," & dic_Addresses_Initial("SCF_PV") & ",0)" & _
        "+IF(" & dic_Addresses_Final("SCF_Type") & "=""CASH""," & dic_Addresses_Final("SCF_PV") & ",0)"
    rng_PnL.Formula = "=" & rng_MV.Address(False, False) & "+" & rng_Cash.Address(False, False)

    wks_output.Calculate
    wks_output.Cells.HorizontalAlignment = xlCenter
    wks_output.Columns.AutoFit
End Sub