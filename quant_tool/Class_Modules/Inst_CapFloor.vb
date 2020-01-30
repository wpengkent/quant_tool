Option Explicit

' ## MEMBER DATA
' Components
Private irl_underlying As IRLeg, scf_Premium As SCF

' Dependent curves
Private fxs_Spots As Data_FXSpots, cvl_volcurve As Data_CapVolsQJK, cvl_ShiftedVolCurve As Data_CapVolsQJK

' Variable dates
Private lng_ValDate As Long

' Dynamic values
Private dblLst_CapletVols As Collection
Private dblLst_ShiftedCapletVols As Collection

' Static values
Private dic_GlobalStaticInfo As Dictionary, dic_CurveDependencies As Dictionary
Private fld_Params As InstParams_CFL
Private int_Sign As Integer

Private Const dbl_StrikeGap As Double = 0.0001

'QJK 04/11/2016 added for flat vega
Dim inBetweenStrikes As Boolean
Dim upperStrike As Double
Dim lowerStrike As Double

Dim cvl_VolCurveUpperStrike As Data_CapVolsQJK
Dim cvl_VolCurveLowerStrike As Data_CapVolsQJK

'end of
'QJK 04/11/2016 added for flat vega

' ## INITIALIZATION
Public Sub Initialize(fld_ParamsInput As InstParams_CFL, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional dic_StaticInfoInput As Dictionary = Nothing)

    ' Store static values
    If dic_StaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_StaticInfoInput
    fld_Params = fld_ParamsInput

    lng_ValDate = fld_Params.Underlying.ValueDate

    ' Set up underlying IR leg
    Set irl_underlying = New IRLeg
    Call irl_underlying.Initialize(fld_Params.Underlying, dic_CurveSet, dic_GlobalStaticInfo)

    ' Store dependent curves
    'If dic_CurveSet Is Nothing Then
    Set cvl_volcurve = GetObject_CapVolSurf(fld_Params.VolCurve, fld_Params.strike, True, False)        'Mandy-interpolate between strike pillars
    Set cvl_ShiftedVolCurve = GetObject_CapVolSurf(fld_Params.VolCurve, fld_Params.strike + dbl_StrikeGap, True, False)

    Set fxs_Spots = GetObject_FXSpots(True)

    Call StoreCapletVols
    Call StoreShiftedCapletVols

    'Else
    '    Set cvl_VolCurve = dic_CurveSet(CurveType.CVL)(fld_Params.VolCurve)
    '    Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
    'End If

    ' Set up premium
    Set scf_Premium = New SCF
    Call scf_Premium.Initialize(fld_Params.Premium, dic_CurveSet, dic_GlobalStaticInfo)
    scf_Premium.ZShiftsEnabled_DF = GetSetting_IsPremInDV01()
    scf_Premium.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

    ' Store calculated values
    Select Case UCase(fld_Params.BuySell)
        Case "B", "BUY": int_Sign = 1
        Case "S", "SELL": int_Sign = -1
    End Select

    ' Determine curve dependencies
    Set dic_CurveDependencies = scf_Premium.CurveDependencies
    Set dic_CurveDependencies = Convert_MergeDicts(dic_CurveDependencies, fxs_Spots.Lookup_CurveDependencies(fld_Params.CCY_PnL))
    Set dic_CurveDependencies = Convert_MergeDicts(dic_CurveDependencies, irl_underlying.CurveDependencies)
End Sub


' ## PROPERTIES

Public Property Get marketvalue() As Double
    marketvalue = CalcValue("MV", fld_Params.IsDigital)
End Property

Public Property Get Cash() As Double
    Cash = CalcValue("CASH", fld_Params.IsDigital) - scf_Premium.CalcValue(lng_ValDate, lng_ValDate, fld_Params.CCY_PnL) * int_Sign
End Property

Public Property Get PnL() As Double
    PnL = Me.marketvalue + Me.Cash
End Property


' ## METHODS - GREEKS
Public Function Calc_DV01(str_curve As String, Optional int_PillarIndex As Integer = 0) As Double
    ' ## Return DV01 sensitivity to the specified curve
    Dim dbl_Output As Double
    If dic_CurveDependencies.Exists(str_curve) Then
        ' Remember original setting, then disable DV01 impact on discounted spot
        Dim bln_ZShiftsEnabled_DiscSpot As Boolean: bln_ZShiftsEnabled_DiscSpot = fxs_Spots.ZShiftsEnabled_DiscSpot
        fxs_Spots.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

        ' Store shifted values
        Dim dbl_Val_Up As Double, dbl_Val_Down As Double
        Call irl_underlying.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        Call scf_Premium.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        dbl_Val_Up = Me.PnL

        Call irl_underlying.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        Call scf_Premium.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        dbl_Val_Down = Me.PnL

        ' Clear temporary shifts from the underlying leg
        Call irl_underlying.SetCurveState(str_curve, CurveState_IRC.Final)
        Call scf_Premium.SetCurveState(str_curve, CurveState_IRC.Final)

        ' Restore original settings
        fxs_Spots.ZShiftsEnabled_DiscSpot = bln_ZShiftsEnabled_DiscSpot

        ' Calculate by finite differencing and convert to PnL currency
        dbl_Output = (dbl_Val_Up - dbl_Val_Down) / 2
    Else
        dbl_Output = 0
    End If

    Calc_DV01 = dbl_Output
End Function

Public Function Calc_DV02(str_curve As String, Optional int_PillarIndex As Integer = 0) As Double
    ' ## Return second order sensitivity to the specified curve
    Dim dbl_Output As Double
    If dic_CurveDependencies.Exists(str_curve) Then
        ' Remember original setting, then disable DV01 impact on discounted spot
        Dim bln_ZShiftsEnabled_DiscSpot As Boolean: bln_ZShiftsEnabled_DiscSpot = fxs_Spots.ZShiftsEnabled_DiscSpot
        fxs_Spots.ZShiftsEnabled_DiscSpot = GetSetting_IsDiscSpotInDV01()

        ' Store shifted values
        Dim dbl_Val_Up As Double, dbl_Val_Down As Double, dbl_Val_Unch As Double
        Call irl_underlying.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        Call scf_Premium.SetCurveState(str_curve, CurveState_IRC.Zero_Up1BP, int_PillarIndex)
        dbl_Val_Up = Me.PnL

        Call irl_underlying.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        Call scf_Premium.SetCurveState(str_curve, CurveState_IRC.Zero_Down1BP, int_PillarIndex)
        dbl_Val_Down = Me.PnL

        ' Clear temporary shifts from the underlying leg
        Call irl_underlying.SetCurveState(str_curve, CurveState_IRC.Final)
        Call scf_Premium.SetCurveState(str_curve, CurveState_IRC.Final)
        dbl_Val_Unch = Me.PnL

        ' Restore original settings
        fxs_Spots.ZShiftsEnabled_DiscSpot = bln_ZShiftsEnabled_DiscSpot

        ' Calculate by finite differencing and convert to PnL currency
        dbl_Output = dbl_Val_Up + dbl_Val_Down - 2 * dbl_Val_Unch
    Else
        dbl_Output = 0
    End If

    Calc_DV02 = dbl_Output
End Function

Public Function Calc_Vega(enu_Type As CurveType, str_curve As String) As Double
    ' ## Return sensitivity to a 1 vol increase in the vol
    Dim dbl_Output As Double

    If cvl_volcurve.TypeCode = enu_Type And fld_Params.VolCurve = str_curve Then
        ' Store shifted values
        Dim dbl_Val_Up As Double, dbl_Val_Down As Double
        cvl_volcurve.VolShift_Sens = 0.01
        cvl_ShiftedVolCurve.VolShift_Sens = 0.01

        Call StoreCapletVols
        Call StoreShiftedCapletVols
        dbl_Val_Up = Me.PnL

        cvl_volcurve.VolShift_Sens = -0.01
        cvl_ShiftedVolCurve.VolShift_Sens = -0.01
        Call StoreCapletVols
        Call StoreShiftedCapletVols
        dbl_Val_Down = Me.PnL

        ' Clear temporary shifts from the underlying leg
        cvl_volcurve.VolShift_Sens = 0
        cvl_ShiftedVolCurve.VolShift_Sens = 0
        Call StoreCapletVols
        Call StoreShiftedCapletVols

        ' Calculate by finite differencing and convert to PnL currency
        dbl_Output = (dbl_Val_Up - dbl_Val_Down) * 50
    Else
        dbl_Output = 0
    End If

    Calc_Vega = dbl_Output
End Function
Public Function Calc_FlatVega_QJK(enuType As CurveType, str_curve As String) As Double
        Set cvl_VolCurveUpperStrike = New Data_CapVolsQJK
        Set cvl_VolCurveLowerStrike = New Data_CapVolsQJK
'only intialise before flat vega  calculations
Call initialiseCapVolSurfFlatVega(fld_Params.VolCurve, fld_Params.strike, True, False)

    ' ## Return sensitivity to a 1 vol increase in the vol
    Dim dbl_Output As Double
    Dim dbl_Val_Up As Double, dbl_Val_Down As Double

    If inBetweenStrikes = False Then  'on pillar, not off pillar.

'1. shock caps
Call cvl_volcurve.Scen_ApplyBase

    If cvl_volcurve.TypeCode = enuType And fld_Params.VolCurve = str_curve Then
    'shock up
    cvl_volcurve.SetShockInst ("CAP")  'shock the cap vols
    Call cvl_volcurve.Scen_AddUniform(ShockType.Absolute, 0.01)
    'bootstraps caplet vols with the current scenario
    Call cvl_volcurve.Scen_ApplyCurrent   'set str_ShockInst="CAP" to shock cap,int_ActiveDTM=1/100?  csh_Shifts_Abs.ReadShift(int_ActiveDTM)
    Call StoreCapletVols   'stores new caplets from CVL_VOLCURVE
    dbl_Val_Up = Me.PnL
    'rebase
    Call cvl_volcurve.Scen_ApplyBase    '        Dim dblArr_CapVols() As Double: dblArr_CapVols = Convert_RangeToDblArr(rng_ShockedCapVols)

    'shock down
    cvl_volcurve.SetShockInst ("CAP")  'shock the cap vols
    Call cvl_volcurve.Scen_AddUniform(ShockType.Absolute, -0.01)
    'bootstraps caplet vols with the current scenario
    Call cvl_volcurve.Scen_ApplyCurrent   'set str_ShockInst="CAP" to shock cap,int_ActiveDTM=1/100?  csh_Shifts_Abs.ReadShift(int_ActiveDTM)
    Call StoreCapletVols   'stores new caplets from CVL_VOLCURVE
    dbl_Val_Down = Me.PnL
    dbl_Output = (dbl_Val_Up - dbl_Val_Down) * 50

    Else
    dbl_Output = 0
    End If

    'then just set scenario to 0  i.e. RESET
    Call cvl_volcurve.Scen_AddUniform(ShockType.Absolute, 0)
    'bootstraps caplet vols with the current scenario
    Call cvl_volcurve.Scen_ApplyCurrent
    Call StoreCapletVols

    Calc_FlatVega_QJK = dbl_Output



Else   'IF IN BWTEEN STRIKES=TRUE?
         'what happens if in between strikes?

'1. shock caps
Call cvl_VolCurveUpperStrike.Scen_ApplyBase
Call cvl_VolCurveLowerStrike.Scen_ApplyBase
Call StoreInBetweenStrikeCapletVols
    If cvl_volcurve.TypeCode = enuType And fld_Params.VolCurve = str_curve Then
    'shock up both upper strike and lower strike

    cvl_VolCurveUpperStrike.SetShockInst ("CAP")  'shock the cap vols
    Call cvl_VolCurveUpperStrike.Scen_AddUniform(ShockType.Absolute, 0.01)
    Call cvl_VolCurveUpperStrike.Scen_ApplyCurrent      'bootstraps caplet vols with the current scenario
    'Call StoreCapletVols   'stores new caplets from CVL_VOLCURVE
    cvl_VolCurveLowerStrike.SetShockInst ("CAP")  'shock the cap vols
    Call cvl_VolCurveLowerStrike.Scen_AddUniform(ShockType.Absolute, 0.01)
    Call cvl_VolCurveLowerStrike.Scen_ApplyCurrent       'bootstraps caplet vols with the current scenario

    Call StoreInBetweenStrikeCapletVols 'stores new caplets ksplined from upperStrike and lowerStrike
    dbl_Val_Up = Me.PnL


    'rebase
    Call cvl_VolCurveUpperStrike.Scen_ApplyBase
    Call cvl_VolCurveLowerStrike.Scen_ApplyBase

    'shock down both upper strike and lower strike
    cvl_VolCurveUpperStrike.SetShockInst ("CAP")  'shock the cap vols
    Call cvl_VolCurveUpperStrike.Scen_AddUniform(ShockType.Absolute, -0.01)
    Call cvl_VolCurveUpperStrike.Scen_ApplyCurrent      'bootstraps caplet vols with the current scenario

    cvl_VolCurveLowerStrike.SetShockInst ("CAP")  'shock the cap vols
    Call cvl_VolCurveLowerStrike.Scen_AddUniform(ShockType.Absolute, -0.01)
    Call cvl_VolCurveLowerStrike.Scen_ApplyCurrent       'bootstraps caplet vols with the current scenario

   Call StoreInBetweenStrikeCapletVols 'stores new caplets ksplined from upperStrike and lowerStrike

   ' Call StoreCapletVols   'stores new caplets from CVL_VOLCURVE
    dbl_Val_Down = Me.PnL
    dbl_Output = (dbl_Val_Up - dbl_Val_Down) * 50

    Else
    dbl_Output = 0
    End If

    'then just set scenario to 0  i.e. RESET BOTH UPPER STRIKE AND LOWER STRIKE
    Call cvl_VolCurveUpperStrike.Scen_AddUniform(ShockType.Absolute, 0): Call cvl_VolCurveLowerStrike.Scen_AddUniform(ShockType.Absolute, 0)
    'bootstraps caplet vols with the current scenario
    Call cvl_VolCurveUpperStrike.Scen_ApplyCurrent:    Call cvl_VolCurveLowerStrike.Scen_ApplyCurrent
     Call StoreInBetweenStrikeCapletVols

    Calc_FlatVega_QJK = dbl_Output


End If
End Function


'QJK added 04/11/2016 FOR FLAT VEGA IN BETWEEN STRIKES
Public Function initialiseCapVolSurfFlatVega(str_Code As String, dbl_Strike As Double, bln_DataExists As Boolean, bln_AddIfMissing As Boolean, _
    Optional dic_GlobalStaticInfo As Dictionary = Nothing)


    'Check all worksheets name to get the strike pillars
    Dim int_count As Integer
    Dim arr_allSheetName As Variant
    ReDim arr_allSheetName(Sheets.count)

    For int_count = 1 To Sheets.count
        arr_allSheetName(int_count) = Sheets(int_count).Name
    Next int_count

    'capture the relevant strike pillars
    Dim arr_strike() As Double
    Dim int_count2 As Integer: int_count2 = 0
    Dim str_tempCaption As String

    For int_count = 1 To UBound(arr_allSheetName)
        str_tempCaption = UCase(arr_allSheetName(int_count))
        If InStr(str_tempCaption, "CVL_" & str_Code) > 0 And InStr(str_tempCaption, "=") > 0 And _
        Mid(str_tempCaption, InStr(str_tempCaption, "=") + 1, Len(str_tempCaption) - InStr(str_tempCaption, "=")) <> 0 Then
            int_count2 = int_count2 + 1
            ReDim Preserve arr_strike(int_count2)
            arr_strike(int_count2) = Mid(str_tempCaption, InStr(str_tempCaption, "=") + 1, Len(str_tempCaption) - InStr(str_tempCaption, "="))
        End If
    Next int_count

    'sort strikes in ascending order
    Dim int_sort As Integer, int_sort2 As Integer
    Dim dbl_temp As Double

    For int_sort = 1 To UBound(arr_strike)
        For int_sort2 = int_sort + 1 To UBound(arr_strike)
            If arr_strike(int_sort) > arr_strike(int_sort2) Then
                dbl_temp = arr_strike(int_sort2)
                arr_strike(int_sort2) = arr_strike(int_sort)
                arr_strike(int_sort) = dbl_temp
            End If
        Next int_sort2
    Next int_sort

    'find upper pillar and lower pillar of strike
    Dim bln_isOnPillar As Boolean: bln_isOnPillar = False
    Dim int_StrikePillarCount As Integer

    For int_count = 1 To UBound(arr_strike)
        If dbl_Strike - arr_strike(int_count) = 0 Then
            bln_isOnPillar = True
            int_StrikePillarCount = int_count
            Exit For
        End If
        If dbl_Strike < arr_strike(1) Then
            bln_isOnPillar = True
            int_StrikePillarCount = 1
            Exit For
        ElseIf dbl_Strike > arr_strike(UBound(arr_strike)) Then
            bln_isOnPillar = True
            int_StrikePillarCount = UBound(arr_strike)
            Exit For
        ElseIf dbl_Strike - arr_strike(int_count) < 0 Then
            bln_isOnPillar = False
            int_StrikePillarCount = int_count
            Exit For
        End If
    Next int_count

    'Dim cvl_OutputTemp As Data_CapVolsQJK: Set cvl_OutputTemp = New Data_CapVolsQJK
    Dim wks_Location As Worksheet
    On Error GoTo errHandler
    If bln_isOnPillar = True Then   'don't need to do anything, this has already been handled in Intialise
      '  Set wks_Location = ThisWorkbook.Worksheets("CVL_" & str_Code & "K=" & arr_strike(int_StrikePillarCount))
      '  Call cvl_OutputTemp.Initialize(wks_Location, bln_DataExists, dic_GlobalStaticInfo)
      '  Set initialiseCapVolSurfFlatVega = cvl_Output
      inBetweenStrikes = False
    Else
        'Strike falls between pillars
        inBetweenStrikes = True
        'Lower Strike Pillar
        upperStrike = arr_strike(int_StrikePillarCount)
        lowerStrike = arr_strike(int_StrikePillarCount - 1)
       ' cvl_VolCurveUpperStrike = New Data_CapVolsQJK
        'cvl_VolCurveLowerStrike = New Data_CapVolsQJK
        Set wks_Location = ThisWorkbook.Worksheets("CVL_" & str_Code & "K=" & arr_strike(int_StrikePillarCount - 1))
        Call cvl_VolCurveLowerStrike.Initialize(wks_Location, bln_DataExists, dic_GlobalStaticInfo, True, True, dbl_Strike)


        'Upper Strike Pillar
        Set wks_Location = ThisWorkbook.Worksheets("CVL_" & str_Code & "K=" & arr_strike(int_StrikePillarCount))
        Call cvl_VolCurveUpperStrike.Initialize(wks_Location, bln_DataExists, dic_GlobalStaticInfo, True, False, dbl_Strike)
       ' Set initialiseCapVolSurfFlatVega = cvl_Output
    End If

errHandler:
    Select Case Err
        Case 0
        Case 9  ' Worksheet does not exist
            If bln_AddIfMissing = True Then
                ThisWorkbook.Worksheets("CVL").Copy After:=ThisWorkbook.Worksheets("CVL")
                Set wks_Location = ThisWorkbook.Worksheets("CVL (2)")
                wks_Location.Name = "CVL_" & str_Code
                wks_Location.Visible = xlSheetVisible
                Resume Next
            End If
        Case Else
            MsgBox "Error code " & Err.Number & ": " & Err.Description
    End Select
End Function
'END OF QJK ADDED 04/11/2016


' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Sub HandleUpdate_IRC(str_CurveName As String)
    ' ## Update stored values affected by change in the curve values
    Call irl_underlying.HandleUpdate_IRC(str_CurveName)
    Call StoreCapletVols
End Sub

Public Sub SetValDate(lng_Input As Long)
    ' ## Set stored value date and refresh values dependent on the value date
    lng_ValDate = lng_Input

    Call irl_underlying.SetValDate(lng_Input)
    Call StoreCapletVols
End Sub


' ## METHODS - PRIVATE
Private Function GetFXConvFactor() As Double
    ' ## Get factor to convert from the native currency to the PnL reporting currency
    GetFXConvFactor = fxs_Spots.Lookup_DiscSpot(irl_underlying.Params.CCY, fld_Params.CCY_PnL)
End Function

Private Function CalcValue(str_type As String, Optional IsDigital As Boolean = False) As Double
    ' Get discounted option value
    Dim dbl_Output As Double

    Select Case IsDigital
    Case False
    dbl_Output = irl_underlying.Calc_BSOptionValue(fld_Params.Direction, fld_Params.strike, cvl_volcurve.Deduction, _
        cvl_volcurve.DeductionCalendar, True, dblLst_CapletVols, , str_type)
    Case True
    dbl_Output = irl_underlying.Calc_BSOptionValueDigitalSmileOn(fld_Params.Direction, fld_Params.strike, cvl_volcurve.Deduction, _
        cvl_volcurve.DeductionCalendar, True, dblLst_CapletVols, dblLst_ShiftedCapletVols, , str_type)
    End Select

    CalcValue = dbl_Output * GetFXConvFactor() * int_Sign

End Function

Private Sub StoreCapletVols()
    ' ## Store vol values, store cap vol as each caplet vol if curve setting is to specify cap vols
    Dim lngLst_PeriodStart As Collection: Set lngLst_PeriodStart = irl_underlying.PeriodStart
    Set dblLst_CapletVols = New Collection
    Dim int_NumPeriods As Integer: int_NumPeriods = lngLst_PeriodStart.count
    Dim dbl_CapVol As Double
    Dim bln_Bootstrappable As Boolean: bln_Bootstrappable = cvl_volcurve.IsBootstrappable
    If bln_Bootstrappable = False Then dbl_CapVol = cvl_volcurve.Lookup_Vol(irl_underlying.PeriodEnd(int_NumPeriods))

    Dim int_ctr As Integer
    For int_ctr = 1 To int_NumPeriods
        If bln_Bootstrappable = True Then
            If lngLst_PeriodStart(int_ctr) > lng_ValDate Then
                Call dblLst_CapletVols.Add(cvl_volcurve.Lookup_Vol(lngLst_PeriodStart(int_ctr), , True))
            Else
                Call dblLst_CapletVols.Add(0)
            End If
        Else
            Call dblLst_CapletVols.Add(dbl_CapVol)
        End If
    Next int_ctr
End Sub
Private Sub StoreShiftedCapletVols()
    ' ## Store vol values, store cap vol as each caplet vol if curve setting is to specify cap vols
    Dim lngLst_PeriodStart As Collection: Set lngLst_PeriodStart = irl_underlying.PeriodStart
    Set dblLst_ShiftedCapletVols = New Collection
    Dim int_NumPeriods As Integer: int_NumPeriods = lngLst_PeriodStart.count
    Dim dbl_CapVol As Double
    Dim bln_Bootstrappable As Boolean: bln_Bootstrappable = cvl_ShiftedVolCurve.IsBootstrappable
    If bln_Bootstrappable = False Then dbl_CapVol = cvl_ShiftedVolCurve.Lookup_Vol(irl_underlying.PeriodEnd(int_NumPeriods))

    Dim int_ctr As Integer
    For int_ctr = 1 To int_NumPeriods
        If bln_Bootstrappable = True Then
            If lngLst_PeriodStart(int_ctr) > lng_ValDate Then
                Call dblLst_ShiftedCapletVols.Add(cvl_ShiftedVolCurve.Lookup_Vol(lngLst_PeriodStart(int_ctr), , True))
                Else
                Call dblLst_ShiftedCapletVols.Add(0)
            End If
        Else
            Call dblLst_ShiftedCapletVols.Add(dbl_CapVol)
        End If
    Next int_ctr
End Sub

'QJK added 04/11/2016
Private Sub StoreInBetweenStrikeCapletVols()
    ' ## Store vol values, store cap vol as each caplet vol if curve setting is to specify cap vols
    Dim lngLst_PeriodStart As Collection: Set lngLst_PeriodStart = irl_underlying.PeriodStart
    Set dblLst_CapletVols = New Collection
    Dim int_NumPeriods As Integer: int_NumPeriods = lngLst_PeriodStart.count
    Dim dbl_CapVol As Double
    Dim bln_Bootstrappable As Boolean: bln_Bootstrappable = cvl_VolCurveUpperStrike.IsBootstrappable
    If bln_Bootstrappable = False Then dbl_CapVol = cvl_VolCurveUpperStrike.Lookup_Vol(irl_underlying.PeriodEnd(int_NumPeriods))

    Dim int_ctr As Integer
    For int_ctr = 1 To int_NumPeriods
        If bln_Bootstrappable = True Then
            If lngLst_PeriodStart(int_ctr) > lng_ValDate Then
            'interpolate between vols

            'Public Function TMR_KSplineInternal(X_Array() As Double, Y_Array() As Double, x As Double, _
                    Optional yt_1 As Double = 0, Optional q_n As Double = 0, _
                    Optional FlatEnds As Boolean = False, _
                    Optional LinearEnds As Boolean = True) As Double

                    Dim strikeArray(1 To 2) As Double
                    Dim volArray(1 To 2) As Double
                    strikeArray(1) = (lowerStrike)
                    strikeArray(2) = upperStrike
                    volArray(1) = cvl_VolCurveLowerStrike.Lookup_Vol(lngLst_PeriodStart(int_ctr), , True)
                    volArray(2) = cvl_VolCurveUpperStrike.Lookup_Vol(lngLst_PeriodStart(int_ctr), , True)


                Call dblLst_CapletVols.Add(TMR_KSplineInternal(strikeArray, volArray, fld_Params.strike))
            Else
                Call dblLst_CapletVols.Add(0)
            End If
        Else
            Call dblLst_CapletVols.Add(dbl_CapVol)   'this should actually be min vol
        End If
    Next int_ctr
End Sub

Public Function TMR_KSplineInternal(X_Array() As Double, Y_Array() As Double, x As Double, _
                    Optional yt_1 As Double = 0, Optional q_n As Double = 0, _
                    Optional FlatEnds As Boolean = False, _
                    Optional LinearEnds As Boolean = True) As Double
'normally YT_1=-0.5 AND q_n=0.5
'Edited SRS Cubic spline - adapted from Numerical Recipes in C
Dim iCnt As Integer

iCnt = WorksheetFunction.CountA(X_Array)

'''''''''''''''''''''''''''''''''''''''
' values are populated
'''''''''''''''''''''''''''''''''''''''
Dim n As Integer 'n=iCnt
Dim i As Integer, k As Integer, j As Integer  'these are loop counting integers
Dim p, qn, sig, un As Double
ReDim U(iCnt - 1) As Double
ReDim yt(iCnt) As Double 'these are the 2nd deriv values

n = iCnt
yt(1) = yt_1
U(1) = 0

For i = 2 To n - 1
    sig = (X_Array(i) - X_Array(i - 1)) / (X_Array(i + 1) - X_Array(i - 1))
    p = sig * yt(i - 1) + 2
    yt(i) = (sig - 1) / p
    U(i) = (Y_Array(i + 1) - Y_Array(i)) / (X_Array(i + 1) - X_Array(i)) - (Y_Array(i) - Y_Array(i - 1)) / (X_Array(i) - X_Array(i - 1))
    U(i) = (6 * U(i) / (X_Array(i + 1) - X_Array(i - 1)) - sig * U(i - 1)) / p
Next i

qn = q_n
un = 0

yt(n) = (un - qn * U(n - 1)) / (qn * yt(n - 1) + 1)

For k = n - 1 To 1 Step -1
    yt(k) = yt(k) * yt(k + 1) + U(k)
Next k

''''''''''''''''''''
'now eval spline at one point
'''''''''''''''''''''
Dim klo As Integer, khi As Integer, h As Double, b As Double, a As Double, outCnt As Integer
outCnt = WorksheetFunction.CountA(x)
' first find correct interval
ReDim y(1 To outCnt, 1 To 1)
For i = 1 To outCnt
    klo = 1: khi = 2
    If FlatEnds And x <= X_Array(1) Then
        y(i, 1) = Y_Array(1)
    ElseIf FlatEnds And x >= X_Array(n) Then
        y(i, 1) = Y_Array(n)
    ElseIf LinearEnds And x <= X_Array(1) Then
        y(i, 1) = Y_Array(1) + (Y_Array(2) - Y_Array(1)) / (X_Array(2) - X_Array(1)) * (x - X_Array(1))
    ElseIf LinearEnds And x >= X_Array(n) Then
        y(i, 1) = Y_Array(n) + (Y_Array(n) - Y_Array(n - 1)) / (X_Array(n) - X_Array(n - 1)) * (x - X_Array(n))
    Else
        For j = 1 To n - 2
            If x < X_Array(khi) Then Exit For
            klo = klo + 1
            khi = khi + 1
        Next j

        h = X_Array(khi) - X_Array(klo)
        a = (X_Array(khi) - x) / h
        b = (x - X_Array(klo)) / h
        y(i, 1) = a * Y_Array(klo) + b * Y_Array(khi) + ((a ^ 3 - a) * yt(klo) + (b ^ 3 - b) * yt(khi)) * (h ^ 2) / 6
    End If
Next i

TMR_KSplineInternal = y(1, 1)

End Function
'end of QJK added 04/11/2016


' ## METHODS - CALCULATION DETAILS
Public Sub OutputReport(wks_output As Worksheet)
    With wks_output
        .Cells.Clear

        If fld_Params.IsDigital Then

            Call irl_underlying.OutputReport_IRDig(.Range("A1"), "Underlying", fld_Params.strike, fld_Params.Direction, _
                dblLst_CapletVols, dblLst_ShiftedCapletVols, cvl_volcurve.Deduction, cvl_volcurve.DeductionCalendar, fld_Params.BuySell, _
                fld_Params.IsDigital, fld_Params.CCY_PnL, scf_Premium, int_Sign)

            .Columns.AutoFit
            .Cells.HorizontalAlignment = xlCenter

        Else

            Call irl_underlying.OutputReport_Option(.Range("A1"), "Underlying", fld_Params.strike, fld_Params.Direction, _
                dblLst_CapletVols, cvl_volcurve.Deduction, cvl_volcurve.DeductionCalendar, fld_Params.BuySell, _
                fld_Params.CCY_PnL, scf_Premium, int_Sign)

        End If

        .Columns.AutoFit
        .Cells.HorizontalAlignment = xlCenter
    End With
End Sub