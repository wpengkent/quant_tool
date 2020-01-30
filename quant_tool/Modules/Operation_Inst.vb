Option Explicit

' ## GENERAL
Public Sub PerformInstrumentAction(enu_Inst As InstType, enu_Action As InstAction)
    ' ## Output values of selected trades on the booking sheet.  Type is a 3 letter code
    Dim fld_AppState_Orig As ApplicationState: fld_AppState_Orig = Gather_ApplicationState(ApplicationStateType.Current)
    Dim fld_AppState_Opt As ApplicationState: fld_AppState_Opt = Gather_ApplicationState(ApplicationStateType.Optimized)
    Call Action_SetAppState(fld_AppState_Opt)

    Application.StatusBar = "Initializing instruments of type: " & GetInstType_String(enu_Inst)

    ' Determine if instruments need to be set up and stored
    Dim bln_StoreInst As Boolean
    Select Case enu_Action
        Case InstAction.Select_All, InstAction.Select_None, InstAction.Rebase, InstAction.DefineTarget: bln_StoreInst = False
        Case Else: bln_StoreInst = True
    End Select

    Dim ica_Inst As InstrumentCache: Set ica_Inst = GetObject_InstCache(enu_Inst, bln_StoreInst)
    Dim str_curve As String, enu_VegaType As CurveType
    Select Case enu_Action
        Case InstAction.Calc_PnL: Call ica_Inst.OutputValues(RevalType.PnL)
        Case InstAction.Calc_Yield: Call ica_Inst.OutputValues(RevalType.Yield)
        Case InstAction.Calc_Delta: Call ica_Inst.OutputValues(RevalType.Delta)
        Case InstAction.Calc_Gamma: Call ica_Inst.OutputValues(RevalType.Gamma)
        Case InstAction.Calc_DV01, InstAction.Calc_DV02
            str_curve = InputBox("Underlying IR curve to shift:")
            If StrPtr(str_curve) <> 0 Then
                ica_Inst.Curve_IRSens = str_curve

                Select Case enu_Action
                    Case InstAction.Calc_DV01: Call ica_Inst.OutputValues(RevalType.DV01)
                    Case InstAction.Calc_DV02: Call ica_Inst.OutputValues(RevalType.DV02)
                End Select
            End If

        Case InstAction.Calc_Vega
            str_curve = InputBox("Underlying volatility curve to shift:")
            enu_VegaType = GetCurveType(InputBox("Type of curve:"))

            If StrPtr(str_curve) <> 0 Then
                Call ica_Inst.SetVegaCurve(enu_VegaType, str_curve)
                Call ica_Inst.OutputValues(RevalType.Vega)
            End If
          'QJK added 26102016 for flat vega
         Case InstAction.Calc_FlatVega
            str_curve = InputBox("Underlying volatility curve to shift:")
            enu_VegaType = GetCurveType(InputBox("Type of curve:"))

            If StrPtr(str_curve) <> 0 Then
                Call ica_Inst.SetVegaCurve(enu_VegaType, str_curve)
                Call ica_Inst.OutputValues(RevalType.Flat_Vega)
            End If
              'END OF QJK added  26102016 for flat vega


        Case InstAction.Display_Flows: Call ica_Inst.OutputFlows
        Case InstAction.Select_All: Call ica_Inst.SelectAllTrades(True)
        Case InstAction.Select_None: Call ica_Inst.SelectAllTrades(False)
        Case InstAction.Rebase: Call ica_Inst.StoreAsBase
        Case InstAction.DefineTarget: Call ica_Inst.SetTarget
    End Select

    Call Action_SetAppState(fld_AppState_Orig)
End Sub


' ## IR SWAPS
Public Sub SolveLegsIRS()
    ' ## Display in debug window the value of the rate/margin which makes the swap NPV equal to zero
    Dim dic_FoundBooking As Dictionary: Set dic_FoundBooking = GetBookings_IRS()
    Dim rng_ActiveID As Range: Set rng_ActiveID = dic_FoundBooking(BookingAttribute.IDSelection)
    Dim rng_ActiveParams As Range: Set rng_ActiveParams = dic_FoundBooking(BookingAttribute.Params)
    Dim fld_Params As InstParams_IRS
    Dim irs_Swap As Inst_IRSwap
    Dim strLst_Output As New Collection
    Dim str_ActiveTrade As String, str_ActiveLegA As String, str_ActiveLegB As String
    Dim dic_GlobalStaticInfo As Dictionary: Set dic_GlobalStaticInfo = GetStaticInfo()

    ' Value all trades marked as "YES"
    While rng_ActiveID(1, 1).Value <> ""
        If rng_ActiveID(1, 2).Value = "YES" Then
            fld_Params = GetInstParams_IRS(rng_ActiveParams, dic_GlobalStaticInfo)
            Set irs_Swap = GetInst_IRS(fld_Params)

            str_ActiveTrade = Left("Trade: " & rng_ActiveID(1, 1).Value & Space(25), 25)
            str_ActiveLegA = Left("Leg A: " & Format(irs_Swap.ParRate_LegA, "0.00000000") & Space(25), 25)
            str_ActiveLegB = "Leg B: " & Format(irs_Swap.ParRate_LegB, "0.00000000")
            Call strLst_Output.Add(str_ActiveTrade & str_ActiveLegA & str_ActiveLegB)
        End If

        Set rng_ActiveID = rng_ActiveID.Offset(1, 0)
        Set rng_ActiveParams = rng_ActiveParams.Offset(1, 0)
    Wend

    Call frm_DisplayResult.DisplayResult(strLst_Output)
End Sub