Option Explicit

' ## MEMBER DATA
' Curve dependencies
Private fxs_Spots As Data_FXSpots, irc_Disc As Data_IRCurve, irc_SpotDisc As Data_IRCurve

' Dynamic variables
Private bln_ZShiftsEnabled_DF As Boolean, bln_ZShiftsEnabled_DiscSpot As Boolean, bln_ZShiftsEnabled_SpotDF As Boolean

' Static values
Private dic_GlobalStaticInfo As Dictionary, dic_CurveDependencies As Dictionary
Private Const str_Daycount_Duration As String = "ACT/365"
Private fld_Params As SCFParams


' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As SCFParams, dic_CurveSet As Dictionary, Optional dic_StaticInfoInput As Dictionary = Nothing)
    ' Store static info
    fld_Params = fld_ParamsInput
    bln_ZShiftsEnabled_DF = True
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput

    If fld_Params.Amount <> 0 Then
        If dic_CurveSet Is Nothing Then
            Set irc_Disc = GetObject_IRCurve(fld_Params.Curve_Disc, True, False)
            Set irc_Disc = GetObject_IRCurve(fld_Params.Curve_SpotDisc, True, False)
            Set fxs_Spots = GetObject_FXSpots(True)
        Else
            Set irc_Disc = dic_CurveSet(CurveType.IRC)(fld_Params.Curve_Disc)
            Set irc_SpotDisc = dic_CurveSet(CurveType.IRC)(fld_Params.Curve_SpotDisc)
            Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
        End If
    End If

    ' Determine curve dependencies
    Set dic_CurveDependencies = New Dictionary
    dic_CurveDependencies.CompareMode = CompareMethod.TextCompare
    If fld_Params.Amount <> 0 Then
        Call dic_CurveDependencies.Add(irc_Disc.CurveName, True)
        If dic_CurveDependencies.Exists(irc_SpotDisc.CurveName) = False Then Call dic_CurveDependencies.Add(irc_SpotDisc, True)
        Set dic_CurveDependencies = Convert_MergeDicts(dic_CurveDependencies, fxs_Spots.Lookup_CurveDependencies(fld_Params.CCY))
    End If
End Sub


' ## PROPERTIES
Public Property Get PmtDate() As Double
    PmtDate = fld_Params.PmtDate
End Property

Public Property Get ZShiftsEnabled_DF() As Boolean
    ZShiftsEnabled_DF = bln_ZShiftsEnabled_DF
End Property

Public Property Let ZShiftsEnabled_DF(bln_Setting As Boolean)
    bln_ZShiftsEnabled_DF = bln_Setting
End Property

Public Property Get ZShiftsEnabled_DiscSpot() As Boolean
    ZShiftsEnabled_DiscSpot = bln_ZShiftsEnabled_DiscSpot
End Property

Public Property Let ZShiftsEnabled_DiscSpot(bln_Setting As Boolean)
    bln_ZShiftsEnabled_DiscSpot = bln_Setting
End Property

Public Property Get ZShiftsEnabled_SpotDF() As Boolean
    ZShiftsEnabled_SpotDF = bln_ZShiftsEnabled_SpotDF
End Property

Public Property Let ZShiftsEnabled_SpotDF(bln_Setting As Boolean)
    bln_ZShiftsEnabled_SpotDF = bln_Setting
End Property

Public Property Get CurveDependencies() As Dictionary
    Set CurveDependencies = dic_CurveDependencies
End Property


' ## METHODS - LOOKUP
Public Function CalcValue(lng_ValueDate As Long, lng_SpotDate As Long, str_CCY_Report As String, _
    Optional enu_ValueType As ValType = ValType.PnL) As Double
    ' ## Return discounted value in the specified currency for the specified value type
    Dim bln_IsZero As Boolean
    Dim dbl_Output As Double
    Dim lng_PmtDate As Long: lng_PmtDate = fld_Params.PmtDate

    ' Disable zero shifts if directed to
    Dim csh_OrigZShift As CurveDaysShift, csh_OrigZShift_DiscSpot As CurveDaysShift, bln_OrigEnabledZShift_DiscSpot As Double
    Dim enu_OrigState As CurveState_IRC, int_SensPillar_Orig As Integer
    If bln_ZShiftsEnabled_DF = False And Not irc_Disc Is Nothing Then
        enu_OrigState = irc_Disc.CurveState
        int_SensPillar_Orig = irc_Disc.SensPillar
        Call irc_Disc.SetCurveState(CurveState_IRC.Final)
        Call irc_SpotDisc.SetCurveState(CurveState_IRC.Final)
    End If

    If bln_ZShiftsEnabled_DiscSpot = False And Not fxs_Spots Is Nothing Then
        bln_OrigEnabledZShift_DiscSpot = fxs_Spots.ZShiftsEnabled_DiscSpot
        fxs_Spots.ZShiftsEnabled_DiscSpot = bln_ZShiftsEnabled_DiscSpot
    End If

    ' Determine if value is definitely zero
    If fld_Params.Amount = 0 Then
        bln_IsZero = True
    Else
        Select Case enu_ValueType
            Case ValType.Cash: If lng_PmtDate > lng_ValueDate Then bln_IsZero = True Else bln_IsZero = False
            Case ValType.MV: If lng_PmtDate > lng_ValueDate Then bln_IsZero = False Else bln_IsZero = True
            Case ValType.PnL: bln_IsZero = False
        End Select
    End If

    If bln_IsZero = True Then
        dbl_Output = 0
    Else
        Dim dbl_FXConv As Double: dbl_FXConv = fxs_Spots.Lookup_DiscSpot(fld_Params.CCY, str_CCY_Report)
        Dim dbl_DFToPmt As Double, dbl_DFToSpot As Double
        If lng_SpotDate >= fld_Params.PmtDate Then
            dbl_DFToPmt = 1
            dbl_DFToSpot = irc_SpotDisc.Lookup_Rate(lng_ValueDate, fld_Params.PmtDate, "DF", , , False)
        Else
            dbl_DFToPmt = irc_Disc.Lookup_Rate(lng_SpotDate, fld_Params.PmtDate, "DF", , , False)
            dbl_DFToSpot = irc_SpotDisc.Lookup_Rate(lng_ValueDate, lng_SpotDate, "DF", , , False)
        End If
        dbl_Output = fld_Params.Amount * dbl_DFToPmt * dbl_DFToSpot * dbl_FXConv
    End If

    ' Return to original shifts
    If bln_ZShiftsEnabled_DF = False And Not irc_Disc Is Nothing Then
        Call irc_Disc.SetCurveState(enu_OrigState, int_SensPillar_Orig)
        Call irc_SpotDisc.SetCurveState(enu_OrigState, int_SensPillar_Orig)
    End If
    If bln_ZShiftsEnabled_DiscSpot = False And Not fxs_Spots Is Nothing Then fxs_Spots.ZShiftsEnabled_DiscSpot = bln_OrigEnabledZShift_DiscSpot

    CalcValue = dbl_Output
End Function


' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Sub SetCurveState(str_curve As String, enu_State As CurveState_IRC, Optional int_PillarIndex As Integer = 0)
    ' ## For temporary zero shifts only, such as during a finite differencing calculation
    If fld_Params.Amount <> 0 Then
        If irc_Disc.CurveName = str_curve And bln_ZShiftsEnabled_DF = True Then Call irc_Disc.SetCurveState(enu_State, int_PillarIndex)
        If irc_SpotDisc.CurveName = str_curve And bln_ZShiftsEnabled_SpotDF = True Then Call irc_SpotDisc.SetCurveState(enu_State, int_PillarIndex)

        Call fxs_Spots.SetCurveState(str_curve, enu_State, int_PillarIndex)
    End If
End Sub


' ## METHODS - CALCULATION DETAILS
Public Sub OutputReport(rng_OutputTopLeft As Range, str_PVTitle As String, str_PnLCcy As String, int_Sign As Integer, _
    bln_HighlightPV As Boolean, ByRef dic_Addresses As Dictionary, Optional bln_ShowType As Boolean = False)
    ' ## Output details of valuation of the flow
    ' ## Requires address dictionary to contain keys: ValDate, PnLCCY

    Dim str_DateFormat As String: str_DateFormat = Gather_DateFormat()
    Dim str_CurrencyFormat As String: str_CurrencyFormat = Gather_CurrencyFormat()
    Dim int_ActiveRow As Integer: int_ActiveRow = 0

    With rng_OutputTopLeft
        .Offset(int_ActiveRow, 0).Value = "Flow:"
        .Offset(int_ActiveRow, 1).Value = fld_Params.Amount * int_Sign
        .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
        Call dic_Addresses.Add("SCF_Amt", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "CCY:"
        .Offset(int_ActiveRow, 1).Value = fld_Params.CCY
        Call dic_Addresses.Add("SCF_CCY", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "Pmt date:"
        .Offset(int_ActiveRow, 1).Value = fld_Params.PmtDate
        .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
        Call dic_Addresses.Add("SCF_Date", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = "DF:"
        .Offset(int_ActiveRow, 1).Value = "=cyReadIRCurve(""" & fld_Params.Curve_Disc & """," & dic_Addresses("ValDate") _
            & "," & dic_Addresses("SCF_Date") & ",""DF"",,FALSE)"
        Call dic_Addresses.Add("SCF_DF", .Offset(int_ActiveRow, 1).Address(False, False))

        int_ActiveRow = int_ActiveRow + 1
        .Offset(int_ActiveRow, 0).Value = str_PVTitle & " (" & str_PnLCcy & "):"
        .Offset(int_ActiveRow, 1).Value = "=" & dic_Addresses("SCF_Amt") & "*" & dic_Addresses("SCF_DF") _
            & "*cyGetFXDiscSpot(" & dic_Addresses("SCF_CCY") & "," & dic_Addresses("PnLCCY") & ")"
        .Offset(int_ActiveRow, 1).NumberFormat = str_CurrencyFormat
        If bln_HighlightPV = True Then .Offset(int_ActiveRow, 1).Interior.ColorIndex = 20
        Call dic_Addresses.Add("SCF_PV", .Offset(int_ActiveRow, 1).Address(False, False))

        If bln_ShowType = True Then
            int_ActiveRow = int_ActiveRow + 1
            .Offset(int_ActiveRow, 0).Value = "Type:"
            If fld_Params.PmtDate > rng_OutputTopLeft.Worksheet.Range(dic_Addresses("ValDate")).Value Then
                .Offset(int_ActiveRow, 1).Value = "MV"
            Else
                .Offset(int_ActiveRow, 1).Value = "CASH"
            End If

            .Offset(int_ActiveRow, 1).NumberFormat = str_DateFormat
            Call dic_Addresses.Add("SCF_Type", .Offset(int_ActiveRow, 1).Address(False, False))
        End If
    End With
End Sub