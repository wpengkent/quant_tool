Option Explicit


' ## MEMBER DATA
Private Const int_SelectionCol As Integer = 4
Private enu_InstType As InstType
Private dic_AllBookings As Dictionary, dic_InstDescriptions As Dictionary, dic_AllStoredInst As Dictionary
Private dic_CurveSet As Dictionary, dic_GlobalStaticInfo As Dictionary, cfg_Settings As ConfigSheet
Private lng_ValDate As Long
Private dic_Cache_AllIDInfo As Dictionary
Private eng_Risk As RiskEngine

' Variable parameters
Private str_Curve_IRSens As String, int_Pillar_IRSens As Integer
Private str_Curve_Vega As String, enu_DataType_Vega As CurveType


' ## INITIALIZATION
Public Sub Initialize(enu_InstTypeInput As InstType, bln_StoreInst As Boolean, Optional dic_CurveSetInput As Dictionary = Nothing, _
    Optional dic_GlobalStaticInfoInput As Dictionary = Nothing)
    ' Gather curves and static data sets
    If dic_GlobalStaticInfoInput Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo() Else Set dic_GlobalStaticInfo = dic_GlobalStaticInfoInput
    If dic_CurveSetInput Is Nothing Then
        Set dic_CurveSet = GetAllCurves(True, False, dic_GlobalStaticInfo)
        Call FillAllDependencies(dic_CurveSet)
    Else
        Set dic_CurveSet = dic_CurveSetInput
    End If
    Set cfg_Settings = dic_GlobalStaticInfo(StaticInfoType.ConfigSheet)

    enu_InstType = enu_InstTypeInput
    lng_ValDate = cfg_Settings.CurrentValDate
    Set dic_Cache_AllIDInfo = New Dictionary
    Set eng_Risk = GetObject_RiskEngine()

    ' Read and store selected trades for each instrument type
    Set dic_AllStoredInst = New Dictionary
    Set dic_AllBookings = New Dictionary
    Set dic_InstDescriptions = New Dictionary

    If enu_InstType = InstType.RngAcc Or enu_InstType = InstType.All Then
        Call dic_AllBookings.Add(InstType.RngAcc, GetBookings_RngAcc())
        If bln_StoreInst = True Then Call StoreInst(InstType.RngAcc)
        Call dic_InstDescriptions.Add(InstType.RngAcc, "RngAcc")
    End If

    If enu_InstType = InstType.IRS Or enu_InstType = InstType.All Then
        Call dic_AllBookings.Add(InstType.IRS, GetBookings_IRS())
        If bln_StoreInst = True Then Call StoreInst(InstType.IRS)
        Call dic_InstDescriptions.Add(InstType.IRS, "IR swaps")
    End If

    If enu_InstType = InstType.CFL Or enu_InstType = InstType.All Then
        Call dic_AllBookings.Add(InstType.CFL, GetBookings_CFL())
        If bln_StoreInst = True Then Call StoreInst(InstType.CFL)
        Call dic_InstDescriptions.Add(InstType.CFL, "caps / floors")
    End If

    If enu_InstType = InstType.SWT Or enu_InstType = InstType.All Then
        Call dic_AllBookings.Add(InstType.SWT, GetBookings_SWT())
        If bln_StoreInst = True Then Call StoreInst(InstType.SWT)
        Call dic_InstDescriptions.Add(InstType.SWT, "swaptions")
    End If

    If enu_InstType = InstType.FXF Or enu_InstType = InstType.All Then
        Call dic_AllBookings.Add(InstType.FXF, GetBookings_FXF())
        If bln_StoreInst = True Then Call StoreInst(InstType.FXF)
        Call dic_InstDescriptions.Add(InstType.FXF, "FX forwards")
    End If

    If enu_InstType = InstType.DEP Or enu_InstType = InstType.All Then
        Call dic_AllBookings.Add(InstType.DEP, GetBookings_DEP())
        If bln_StoreInst = True Then Call StoreInst(InstType.DEP)
        Call dic_InstDescriptions.Add(InstType.DEP, "simple deposits")
    End If

    If enu_InstType = InstType.FRA Or enu_InstType = InstType.All Then
        Call dic_AllBookings.Add(InstType.FRA, GetBookings_FRA())
        If bln_StoreInst = True Then Call StoreInst(InstType.FRA)
        Call dic_InstDescriptions.Add(InstType.FRA, "FRAs")
    End If

    If enu_InstType = InstType.FVN Or enu_InstType = InstType.All Then
        Call dic_AllBookings.Add(InstType.FVN, GetBookings_FVN())
        If bln_StoreInst = True Then Call StoreInst(InstType.FVN)
        Call dic_InstDescriptions.Add(InstType.FVN, "FX vanilla options")
    End If

    If enu_InstType = InstType.BND Or enu_InstType = InstType.All Then
        Call dic_AllBookings.Add(InstType.BND, GetBookings_BND())
        If bln_StoreInst = True Then Call StoreInst(InstType.BND)
        Call dic_InstDescriptions.Add(InstType.BND, "bonds")
    End If

    If enu_InstType = InstType.BA Or enu_InstType = InstType.All Then
        Call dic_AllBookings.Add(InstType.BA, GetBookings_BA())
        If bln_StoreInst = True Then Call StoreInst(InstType.BA)
        Call dic_InstDescriptions.Add(InstType.BA, "BAs")
    End If

    If enu_InstType = InstType.NID Or enu_InstType = InstType.All Then   'QJK 03102014
        Call dic_AllBookings.Add(InstType.NID, GetBookings_NID())
        If bln_StoreInst = True Then Call StoreInst(InstType.NID)
        Call dic_InstDescriptions.Add(InstType.NID, "NIDs")
    End If

    If enu_InstType = InstType.FBR Or enu_InstType = InstType.All Then
        Call dic_AllBookings.Add(InstType.FBR, GetBookings_FBR())
        If bln_StoreInst = True Then Call StoreInst(InstType.FBR)
        Call dic_InstDescriptions.Add(InstType.FBR, "FX barrier options")
    End If

    If enu_InstType = InstType.FTB Or enu_InstType = InstType.All Then
        Call dic_AllBookings.Add(InstType.FTB, GetBookings_FTB())
        If bln_StoreInst = True Then Call StoreInst(InstType.FTB)
        Call dic_InstDescriptions.Add(InstType.FTB, "bill futures")
    End If

    If enu_InstType = InstType.FBN Or enu_InstType = InstType.All Then
        Call dic_AllBookings.Add(InstType.FBN, GetBookings_FBN())
        If bln_StoreInst = True Then Call StoreInst(InstType.FBN)
        Call dic_InstDescriptions.Add(InstType.FBN, "bond futures")
    End If

    If enu_InstType = InstType.FRE Or enu_InstType = InstType.All Then
        Call dic_AllBookings.Add(InstType.FRE, GetBookings_FRE())
        If bln_StoreInst = True Then Call StoreInst(InstType.FRE)
        Call dic_InstDescriptions.Add(InstType.FRE, "FX rebates")
    End If

    If enu_InstType = InstType.ECS Or enu_InstType = InstType.All Then
        Call dic_AllBookings.Add(InstType.ECS, GetBookings_ECS())
        If bln_StoreInst = True Then Call StoreInst(InstType.ECS)
        Call dic_InstDescriptions.Add(InstType.ECS, "EQ cash")
    End If

    If enu_InstType = InstType.EQO Or enu_InstType = InstType.All Then
        Call dic_AllBookings.Add(InstType.EQO, GetBookings_EQO())
        If bln_StoreInst = True Then Call StoreInst(InstType.EQO)
        Call dic_InstDescriptions.Add(InstType.EQO, "EQ options")
    End If

    If enu_InstType = InstType.EQF Or enu_InstType = InstType.All Then
        Call dic_AllBookings.Add(InstType.EQF, GetBookings_EQF())
        If bln_StoreInst = True Then Call StoreInst(InstType.EQF)
        Call dic_InstDescriptions.Add(InstType.EQF, "EQ futures")
    End If

    If enu_InstType = InstType.EQS Or enu_InstType = InstType.All Then
        Call dic_AllBookings.Add(InstType.EQS, GetBookings_EQS())
        If bln_StoreInst = True Then Call StoreInst(InstType.EQS)
        Call dic_InstDescriptions.Add(InstType.EQS, "EQ swap")
    End If

    If enu_InstType = InstType.FXFut Or enu_InstType = InstType.All Then
        Call dic_AllBookings.Add(InstType.FXFut, GetBookings_FXFut())
        If bln_StoreInst = True Then Call StoreInst(InstType.FXFut)
        Call dic_InstDescriptions.Add(InstType.FXFut, "FX Futures")
    End If

End Sub


' ## PROPERTIES
Public Property Let Curve_IRSens(str_CurveName As String)
    ' ## Set curve being shifted for DV01
    str_Curve_IRSens = str_CurveName
End Property

Public Property Let Pillar_IRSens(int_Pillar As Integer)
    ' ## Set curve being shifted for DV01.  For uniform shifts, use 0 as the pillar
    int_Pillar_IRSens = int_Pillar
End Property


' ## METHODS - SET PARAMETERS
Public Sub SetVegaCurve(enu_Type As CurveType, str_CurveName As String)
    ' ## Set curve being shifted for vega
    enu_DataType_Vega = enu_Type
    str_Curve_Vega = str_CurveName
End Sub


' ## METHODS - INSTRUMENT ACTIONS
Public Sub OutputValues(enu_RevalType As RevalType)
    ' Output values for selected trades for each instrument type
    Dim dic_ActiveBooking As Dictionary
    If enu_InstType = InstType.All Then
        ' Perform action on all instrument types
        Dim var_ActiveInstType As Variant
        For Each var_ActiveInstType In dic_AllBookings.Keys
            Call OutputInstValues(CStr(var_ActiveInstType), enu_RevalType)
            Set dic_ActiveBooking = dic_AllBookings(var_ActiveInstType)
            Call dic_ActiveBooking(BookingAttribute.Sheet).Calculate
        Next var_ActiveInstType
    Else
        ' Perform action on specified instrument type
        Call OutputInstValues(enu_InstType, enu_RevalType)
        Set dic_ActiveBooking = dic_AllBookings(enu_InstType)
        Call dic_ActiveBooking(BookingAttribute.Sheet).Calculate
    End If
End Sub

Public Sub OutputFlows()
    ' Output values for selected trades for each instrument type
    If enu_InstType = InstType.All Then
        ' Perform action on all instrument types
        Dim var_ActiveInstType As Variant
        For Each var_ActiveInstType In dic_AllBookings.Keys
            Call OutputInstFlows(CStr(var_ActiveInstType))
        Next var_ActiveInstType
    Else
        ' Perform action on specified instrument type
        Call OutputInstFlows(enu_InstType)
    End If
End Sub

Public Sub PerformRecalc(lng_ValDateInput As Long)
    ' Determine if valuation date needs to be shifted, and store the new date
    Dim bln_Shifted_ValDate As Boolean
    If lng_ValDateInput <> lng_ValDate Then
        bln_Shifted_ValDate = True
        lng_ValDate = lng_ValDateInput
    Else
        bln_Shifted_ValDate = False
    End If

    ' Recalculate stored values for each instrument type
    If enu_InstType = InstType.All Then
        ' Perform action on all instrument types
        Dim var_ActiveInstType As Variant, enu_ActiveInstType As InstType
        For Each var_ActiveInstType In dic_AllBookings.Keys
            enu_ActiveInstType = var_ActiveInstType
            Call PerformInstRecalc(enu_ActiveInstType, bln_Shifted_ValDate)
        Next var_ActiveInstType
    Else
        ' Perform action on specified instrument type
        Call PerformInstRecalc(enu_InstType, bln_Shifted_ValDate)
    End If
End Sub

Public Sub SelectAllTrades(bln_Select As Boolean)
    ' Output values for selected trades for each instrument type
    If enu_InstType = InstType.All Then
        ' Perform action on all instrument types
        Dim var_ActiveInstType As Variant, enu_ActiveInstType As InstType
        For Each var_ActiveInstType In dic_AllBookings.Keys
            enu_ActiveInstType = var_ActiveInstType
            Call SelectAllInstTrades(enu_ActiveInstType, bln_Select)
        Next var_ActiveInstType
    Else
        ' Perform action on specified instrument type
        Call SelectAllInstTrades(enu_InstType, bln_Select)
    End If
End Sub

Public Sub StoreAsBase()
    ' Output values for selected trades for each instrument type
    Dim int_Result As Integer: int_Result = MsgBox("Set current values as base?", vbOKCancel)
    If int_Result = vbOK Then
        If enu_InstType = InstType.All Then
            ' Perform action on all instrument types
            Dim var_ActiveInstType As Variant, enu_ActiveInstType As InstType
            For Each var_ActiveInstType In dic_AllBookings.Keys
                enu_ActiveInstType = var_ActiveInstType
                Call StoreAsBase_Inst(enu_ActiveInstType)
            Next var_ActiveInstType
        Else
            ' Perform action on specified instrument type
            Call StoreAsBase_Inst(enu_InstType)
        End If
    End If
End Sub

Public Sub SetTarget()
    ' ## Replace trade section of RE Output sheet with info for the selected trades and a cell reference to each value being collected
    eng_Risk.ClearTrades

    ' Set the target
    Dim rng_ActiveOutput As Range: Set rng_ActiveOutput = eng_Risk.TopLeft_Trades  ' This will get shifted within the instrument level function
    If enu_InstType = InstType.All Then
        ' Perform action on all instrument types
        Dim var_ActiveInstType As Variant, enu_ActiveInstType As InstType
        For Each var_ActiveInstType In dic_AllBookings.Keys
            enu_ActiveInstType = var_ActiveInstType
            Call SetTarget_Inst(enu_ActiveInstType, rng_ActiveOutput)
        Next var_ActiveInstType
    Else
        ' Perform action on specified instrument type
        Call SetTarget_Inst(enu_InstType, rng_ActiveOutput)
    End If

    Call GotoSheet(eng_Risk.SetupSheet.Name)
End Sub


' ## METHODS - PRIVATE
Private Sub StoreInst(enu_ActiveInstType As InstType)
    ' ## Create and store instrument objects for the selected trades of the specified instrument type
    Dim rng_IDTopLeft As Range, rng_ActiveParams As Range
    Dim dic_FoundBooking As Dictionary: Set dic_FoundBooking = dic_AllBookings(enu_ActiveInstType)
    Set rng_IDTopLeft = dic_FoundBooking(BookingAttribute.IDSelection)
    Set rng_ActiveParams = dic_FoundBooking(BookingAttribute.Params)

    ' Set up ranges and objects
    Dim dic_ActiveStoredInst As New Dictionary, dic_Cache_ActiveIDInfo As New Dictionary
    Call dic_AllStoredInst.Add(enu_ActiveInstType, dic_ActiveStoredInst)
    Call dic_Cache_AllIDInfo.Add(enu_ActiveInstType, dic_Cache_ActiveIDInfo)
    Dim strArr_ActiveIDInfo(1 To 4) As String

    ' Store selected instruments
    Dim int_NumRows As Integer: int_NumRows = Examine_NumRows(rng_IDTopLeft)
    Dim str_ActiveID As String

    Dim int_ctr As Integer
    For int_ctr = 1 To int_NumRows
        str_ActiveID = rng_IDTopLeft(int_ctr, 1).Value
        strArr_ActiveIDInfo(1) = str_ActiveID
        strArr_ActiveIDInfo(2) = rng_IDTopLeft(int_ctr, 2).Value
        strArr_ActiveIDInfo(3) = rng_IDTopLeft(int_ctr, 3).Value
        strArr_ActiveIDInfo(4) = rng_IDTopLeft(int_ctr, 4).Value
        Call dic_Cache_ActiveIDInfo.Add(int_ctr, strArr_ActiveIDInfo)

        If strArr_ActiveIDInfo(int_SelectionCol) = "YES" Then
            Select Case enu_ActiveInstType
                Case InstType.RngAcc  '#Alvin
                    Call dic_ActiveStoredInst.Add(str_ActiveID, GetInst_RngAcc(GetInstParams_RngAcc(rng_ActiveParams.Offset(int_ctr - 1, 0), _
                        dic_GlobalStaticInfo, str_ActiveID), dic_CurveSet, dic_GlobalStaticInfo))
                Case InstType.IRS
                    Call dic_ActiveStoredInst.Add(str_ActiveID, GetInst_IRS(GetInstParams_IRS(rng_ActiveParams.Offset(int_ctr - 1, 0), _
                        dic_GlobalStaticInfo, str_ActiveID), dic_CurveSet, dic_GlobalStaticInfo))
                Case InstType.CFL
                    Call dic_ActiveStoredInst.Add(str_ActiveID, GetInst_CFL(GetInstParams_CFL(rng_ActiveParams.Offset(int_ctr - 1, 0), _
                        dic_GlobalStaticInfo, str_ActiveID), dic_CurveSet, dic_GlobalStaticInfo))
                Case InstType.SWT
                    Call dic_ActiveStoredInst.Add(str_ActiveID, GetInst_SWT(GetInstParams_SWT(rng_ActiveParams.Offset(int_ctr - 1, 0), _
                        dic_GlobalStaticInfo, str_ActiveID), dic_CurveSet, dic_GlobalStaticInfo))
                Case InstType.FXF
                    Call dic_ActiveStoredInst.Add(str_ActiveID, GetInst_FXF(GetInstParams_FXF(rng_ActiveParams.Offset(int_ctr - 1, 0), _
                        dic_GlobalStaticInfo, str_ActiveID), dic_CurveSet, dic_GlobalStaticInfo))
                Case InstType.DEP
                    Call dic_ActiveStoredInst.Add(str_ActiveID, GetInst_DEP(GetInstParams_DEP(rng_ActiveParams.Offset(int_ctr - 1, 0), _
                        dic_GlobalStaticInfo, str_ActiveID), dic_CurveSet, dic_GlobalStaticInfo))
                Case InstType.FRA
                    Call dic_ActiveStoredInst.Add(str_ActiveID, GetInst_FRA(GetInstParams_FRA(rng_ActiveParams.Offset(int_ctr - 1, 0), _
                        dic_GlobalStaticInfo, str_ActiveID), dic_CurveSet, dic_GlobalStaticInfo))
                Case InstType.FVN
                    Call dic_ActiveStoredInst.Add(str_ActiveID, GetInst_FVN(GetInstParams_FVN(rng_ActiveParams.Offset(int_ctr - 1, 0), _
                        dic_GlobalStaticInfo, str_ActiveID), dic_CurveSet, dic_GlobalStaticInfo))
                Case InstType.BND
                    Call dic_ActiveStoredInst.Add(str_ActiveID, GetInst_BND(GetInstParams_BND(rng_ActiveParams.Offset(int_ctr - 1, 0), _
                        dic_GlobalStaticInfo, str_ActiveID), dic_CurveSet, dic_GlobalStaticInfo))
                Case InstType.BA
                    Call dic_ActiveStoredInst.Add(str_ActiveID, GetInst_BA(GetInstParams_BA(rng_ActiveParams.Offset(int_ctr - 1, 0), _
                        dic_GlobalStaticInfo, str_ActiveID), dic_CurveSet, dic_GlobalStaticInfo))
                                        'QJK 03102014
                Case InstType.NID
                    Call dic_ActiveStoredInst.Add(str_ActiveID, GetInst_NID(GetInstParams_NID(rng_ActiveParams.Offset(int_ctr - 1, 0), _
                        dic_GlobalStaticInfo, str_ActiveID), dic_CurveSet, dic_GlobalStaticInfo))
                Case InstType.FBR
                    Call dic_ActiveStoredInst.Add(str_ActiveID, GetInst_FBR(GetInstParams_FBR(rng_ActiveParams.Offset(int_ctr - 1, 0), _
                        dic_GlobalStaticInfo, str_ActiveID), dic_CurveSet, dic_GlobalStaticInfo))
                Case InstType.FTB
                    Call dic_ActiveStoredInst.Add(str_ActiveID, GetInst_FTB(GetInstParams_FTB(rng_ActiveParams.Offset(int_ctr - 1, 0), _
                        dic_GlobalStaticInfo, str_ActiveID), dic_CurveSet, dic_GlobalStaticInfo))
                Case InstType.FBN
                    Call dic_ActiveStoredInst.Add(str_ActiveID, GetInst_FBN(GetInstParams_FBN(rng_ActiveParams.Offset(int_ctr - 1, 0), _
                        dic_GlobalStaticInfo, str_ActiveID), dic_CurveSet, dic_GlobalStaticInfo))
                Case InstType.FRE
                    Call dic_ActiveStoredInst.Add(str_ActiveID, GetInst_FRE(GetInstParams_FRE(rng_ActiveParams.Offset(int_ctr - 1, 0), _
                        dic_GlobalStaticInfo, str_ActiveID), dic_CurveSet, dic_GlobalStaticInfo))
                Case InstType.ECS
                    Call dic_ActiveStoredInst.Add(str_ActiveID, GetInst_ECS(GetInstParams_ECS(rng_ActiveParams.Offset(int_ctr - 1, 0), _
                        dic_GlobalStaticInfo, str_ActiveID), dic_CurveSet, dic_GlobalStaticInfo))
                Case InstType.EQO
                    Call dic_ActiveStoredInst.Add(str_ActiveID, GetInst_EQO(GetInstParams_EQO(rng_ActiveParams.Offset(int_ctr - 1, 0), _
                        dic_GlobalStaticInfo, str_ActiveID), dic_CurveSet, dic_GlobalStaticInfo))
                Case InstType.EQF
                    Call dic_ActiveStoredInst.Add(str_ActiveID, GetInst_EQF(GetInstParams_EQF(rng_ActiveParams.Offset(int_ctr - 1, 0), _
                        dic_GlobalStaticInfo, str_ActiveID), dic_CurveSet, dic_GlobalStaticInfo))
                Case InstType.EQS
                    Call dic_ActiveStoredInst.Add(str_ActiveID, GetInst_EQS(GetInstParams_EQS(rng_ActiveParams.Offset(int_ctr - 1, 0), _
                        dic_GlobalStaticInfo, str_ActiveID), dic_CurveSet, dic_GlobalStaticInfo))
                Case InstType.FXFut
                    Call dic_ActiveStoredInst.Add(str_ActiveID, GetInst_FXFut(GetInstParams_FXFut(rng_ActiveParams.Offset(int_ctr - 1, 0), _
                        dic_GlobalStaticInfo, str_ActiveID), dic_CurveSet, dic_GlobalStaticInfo))
            End Select
        End If
    Next int_ctr
End Sub

Private Function Gather_OutputColIndexes(enu_ActiveInstType As InstType) As Dictionary
    ' ## Return instructions for which values can be calculated and in which column to find them
    Dim dic_output As New Dictionary: dic_output.CompareMode = CompareMethod.TextCompare

    Select Case enu_ActiveInstType
        Case InstType.FTB, InstType.FBN
            Call dic_output.Add("PNL", 1)
            Call dic_output.Add("DV01", 2)
            Call dic_output.Add("DV02", 3)
        Case InstType.CFL, InstType.SWT, InstType.FRE
            Call dic_output.Add("MV", 1)
            Call dic_output.Add("CASH", 2)
            Call dic_output.Add("PNL", 3)
            Call dic_output.Add("DV01", 4)
            Call dic_output.Add("DV02", 5)
            Call dic_output.Add("VEGA", 6)
        Case InstType.FVN, InstType.FBR
            Call dic_output.Add("MV", 1)
            Call dic_output.Add("CASH", 2)
            Call dic_output.Add("PNL", 3)
            Call dic_output.Add("DV01", 4)
            Call dic_output.Add("DV02", 5)
            Call dic_output.Add("VEGA", 6)
            Call dic_output.Add("DELTA", 7)
            Call dic_output.Add("GAMMA", 8)
        Case InstType.ECS, InstType.EQO, InstType.EQF, InstType.EQS
            Call dic_output.Add("MV", 1)
            Call dic_output.Add("CASH", 2)
            Call dic_output.Add("PNL", 3)
        Case InstType.BND, InstType.BA, InstType.NID 'QJK code 03102014
            Call dic_output.Add("MV", 1)
            Call dic_output.Add("CASH", 2)
            Call dic_output.Add("PNL", 3)
            Call dic_output.Add("DV01", 4)
            Call dic_output.Add("DV02", 5)
            Call dic_output.Add("YIELD", 6)
            Call dic_output.Add("DURATION", 7)
            Call dic_output.Add("MODIFIED DURATION", 8)
        Case InstType.FXFut
            Call dic_output.Add("MV", 1)
            Call dic_output.Add("CASH", 2)
            Call dic_output.Add("PNL", 3)
            Call dic_output.Add("DV01", 4)
            Call dic_output.Add("DELTA", 5)
        Case InstType.RngAcc
            Call dic_output.Add("MV", 1)
            Call dic_output.Add("CASH", 2)
            Call dic_output.Add("PNL", 3)
            Call dic_output.Add("DV01", 4)
            Call dic_output.Add("DV02", 5)
            Call dic_output.Add("VEGA", 6)


        Case Else
            Call dic_output.Add("MV", 1)
            Call dic_output.Add("CASH", 2)
            Call dic_output.Add("PNL", 3)
            Call dic_output.Add("DV01", 4)
            Call dic_output.Add("DV02", 5)

    End Select

    Set Gather_OutputColIndexes = dic_output
End Function

Private Sub OutputInstValues(enu_ActiveInstType As InstType, enu_RevalType As RevalType)
    ' ## Output values for the selected trades of the specified instrument type
    Dim rng_ActiveIDTopLeft As Range, rng_OutputTopLeft As Range
    Dim dic_FoundBooking As Dictionary: Set dic_FoundBooking = dic_AllBookings(enu_ActiveInstType)
    Set rng_ActiveIDTopLeft = dic_FoundBooking(BookingAttribute.IDSelection)
    Set rng_OutputTopLeft = dic_FoundBooking(BookingAttribute.Outputs)
    Dim str_MessageStart As String: str_MessageStart = "Scenario: " & cfg_Settings.CurrentScen & "    Valuing " & dic_InstDescriptions(enu_ActiveInstType) & " - trades completed: "
    Dim int_NumOutputCols As Integer: int_NumOutputCols = rng_OutputTopLeft.Columns.count

    Dim rng_AllTradeIDs As Range: Set rng_AllTradeIDs = Gather_RangeBelow(rng_ActiveIDTopLeft(1, 1))
    If Not rng_AllTradeIDs Is Nothing Then
        Dim int_NumTotalRows As Integer: int_NumTotalRows = rng_AllTradeIDs.Rows.count
        Dim int_NumSelectedRows As Integer: int_NumSelectedRows = WorksheetFunction.CountIf(rng_AllTradeIDs.Offset(0, int_SelectionCol - 1), "YES")
        Dim int_TotalRowCtr As Integer
        Dim int_SelectedTradeCtr As Integer: int_SelectedTradeCtr = 0
        Dim str_ActiveID As String
        Dim dic_ActiveStoredInst As Dictionary: Set dic_ActiveStoredInst = dic_AllStoredInst(enu_ActiveInstType)

        Application.StatusBar = str_MessageStart & int_SelectedTradeCtr & " of " & int_NumSelectedRows

        ' Determine relevant output columns
        Dim dic_OutputCols As New Dictionary: Set dic_OutputCols = Gather_OutputColIndexes(enu_ActiveInstType)

        ' Read existing values
        Dim rng_OutputValues As Range: Set rng_OutputValues = rng_OutputTopLeft.Resize(int_NumTotalRows, int_NumOutputCols)
        Dim dblArr_OutputValues As Variant: dblArr_OutputValues = rng_OutputValues.Value2

        ' Output selected instrument values
        Dim dic_Cache_ActiveIDInfo As Dictionary: Set dic_Cache_ActiveIDInfo = dic_Cache_AllIDInfo(enu_ActiveInstType)
        Dim dbl_ActiveMV As Double, dbl_ActiveCash As Double, dbl_ActiveYield As Double, dbl_ActiveDuration  As Double, dbl_ActiveModifiedDuration As Double

        Dim strArr_ActiveIDInfo() As String
        For int_TotalRowCtr = 1 To int_NumTotalRows
            strArr_ActiveIDInfo = dic_Cache_ActiveIDInfo(int_TotalRowCtr)

            If strArr_ActiveIDInfo(int_SelectionCol) = "YES" Then
                ' Update status periodically
                int_SelectedTradeCtr = int_SelectedTradeCtr + 1
                If int_SelectedTradeCtr Mod 20 = 0 Then Application.StatusBar = str_MessageStart & int_SelectedTradeCtr & " of " & int_NumSelectedRows
                str_ActiveID = strArr_ActiveIDInfo(1)

                ' Output PnL if selected and available for the instrument type
                With dic_ActiveStoredInst(str_ActiveID)
                    If enu_RevalType = RevalType.All Or enu_RevalType = RevalType.PnL Then
                        If dic_OutputCols.Exists("MV") Then
                            dbl_ActiveMV = .marketvalue
                            dblArr_OutputValues(int_TotalRowCtr, dic_OutputCols("MV")) = dbl_ActiveMV
                        End If
                        If dic_OutputCols.Exists("CASH") Then
                            dbl_ActiveCash = .Cash
                            dblArr_OutputValues(int_TotalRowCtr, dic_OutputCols("CASH")) = dbl_ActiveCash
                        End If

                        If dic_OutputCols.Exists("PNL") Then
                            If dic_OutputCols.Exists("MV") And dic_OutputCols.Exists("CASH") Then
                                dblArr_OutputValues(int_TotalRowCtr, dic_OutputCols("PNL")) = dbl_ActiveMV + dbl_ActiveCash
                            Else
                                dblArr_OutputValues(int_TotalRowCtr, dic_OutputCols("PNL")) = .PnL
                            End If
                        End If
                    End If

                    ' Output DV01 if selected and available for the instrument type
                    If enu_RevalType = RevalType.All Or enu_RevalType = RevalType.DV01 Then
                        If dic_OutputCols.Exists("DV01") Then
                            dblArr_OutputValues(int_TotalRowCtr, dic_OutputCols("DV01")) = .Calc_DV01(str_Curve_IRSens, int_Pillar_IRSens)
                        End If
                    End If

                    ' Output DV02 if selected and available for the instrument type
                    If enu_RevalType = RevalType.All Or enu_RevalType = RevalType.DV02 Then
                        If dic_OutputCols.Exists("DV02") Then
                            dblArr_OutputValues(int_TotalRowCtr, dic_OutputCols("DV02")) = .Calc_DV02(str_Curve_IRSens, int_Pillar_IRSens)
                        End If
                    End If

                    ' Output Vega if selected and available for the instrument type
                    If enu_RevalType = RevalType.All Or enu_RevalType = RevalType.Vega Then
                        If dic_OutputCols.Exists("VEGA") Then
                            dblArr_OutputValues(int_TotalRowCtr, dic_OutputCols("VEGA")) = .Calc_Vega(enu_DataType_Vega, str_Curve_Vega)
                        End If
                    End If
                    'QJK added 26102016 FOR FLAT VEGA
                 ' Output Flat Vega if selected and available for the instrument type
                    If enu_RevalType = RevalType.Flat_Vega Then   'FLAT VEGA APPEARS IN VEGA COLUMN QJK
                        If dic_OutputCols.Exists("VEGA") Then
                            dblArr_OutputValues(int_TotalRowCtr, dic_OutputCols("VEGA")) = .Calc_FlatVega_QJK(enu_DataType_Vega, str_Curve_Vega)
                        End If
                    End If
                    'QJK added 26102016 FOR FLAT VEGA

                    ' Output Delta if selected and available for the instrument type
                    If enu_RevalType = RevalType.All Or enu_RevalType = RevalType.Delta Then
                        If dic_OutputCols.Exists("DELTA") Then
                            dblArr_OutputValues(int_TotalRowCtr, dic_OutputCols("DELTA")) = .Calc_Delta()
                        End If
                    End If

                    ' Output Gamma if selected and available for the instrument type
                    If enu_RevalType = RevalType.All Or enu_RevalType = RevalType.Gamma Then
                        If dic_OutputCols.Exists("GAMMA") Then
                            dblArr_OutputValues(int_TotalRowCtr, dic_OutputCols("GAMMA")) = .Calc_Gamma()
                        End If
                    End If

                    ' Output Yield if selected and available for the instrument type
                    If enu_RevalType = RevalType.All Or enu_RevalType = RevalType.Yield Then
                        If dic_OutputCols.Exists("Yield") Then
                            dbl_ActiveYield = .Yield
                            dblArr_OutputValues(int_TotalRowCtr, dic_OutputCols("YIELD")) = dbl_ActiveYield * 100
                        End If

                        If dic_OutputCols.Exists("Duration") Then
                            dbl_ActiveDuration = .Duration
                            dblArr_OutputValues(int_TotalRowCtr, dic_OutputCols("DURATION")) = dbl_ActiveDuration
                        End If

                        If dic_OutputCols.Exists("Modified Duration") Then
                            dbl_ActiveModifiedDuration = .ModifiedDuration
                            dblArr_OutputValues(int_TotalRowCtr, dic_OutputCols("MODIFIED DURATION")) = dbl_ActiveModifiedDuration
                        End If
                    End If
                End With
            End If
        Next int_TotalRowCtr

        ' Output values to sheet
        rng_OutputTopLeft.Resize(int_NumTotalRows, int_NumOutputCols).Value = dblArr_OutputValues

        Application.StatusBar = False
    End If
End Sub

Private Sub OutputInstFlows(enu_ActiveInstType As InstType)
    ' ## Output flow details for the selected trades of the specified instrument type
    Dim str_ActiveSheetName As String
    Dim wks_ActiveOutput As Worksheet, wks_PrevOutput As Worksheet
    Dim dic_FoundBooking As Dictionary: Set dic_FoundBooking = dic_AllBookings(enu_ActiveInstType)
    Dim rng_ActiveIDTopLeft As Range: Set rng_ActiveIDTopLeft = dic_FoundBooking(BookingAttribute.IDSelection)
    Dim wks_BookingSheet As Worksheet: Set wks_BookingSheet = dic_FoundBooking(BookingAttribute.Sheet)

    Dim str_MessageStart As String: str_MessageStart = "Gathering flows (" & dic_InstDescriptions(enu_ActiveInstType) & ") - trades completed: "

    Dim rng_AllTradeIDs As Range: Set rng_AllTradeIDs = Gather_RangeBelow(rng_ActiveIDTopLeft(1, 1))
    If Not rng_AllTradeIDs Is Nothing Then
        Dim int_NumTotalRows As Integer: int_NumTotalRows = rng_AllTradeIDs.Rows.count
        Dim int_NumSelectedRows As Integer: int_NumSelectedRows = WorksheetFunction.CountIf(rng_AllTradeIDs.Offset(0, int_SelectionCol - 1), "YES")
        Dim int_TotalRowCtr As Integer
        Dim int_SelectedTradeCtr As Integer: int_SelectedTradeCtr = 0
        Dim dic_ActiveStoredInst As Dictionary: Set dic_ActiveStoredInst = dic_AllStoredInst(enu_ActiveInstType)

        Set wks_PrevOutput = wks_BookingSheet
        Application.DisplayAlerts = False
        Application.StatusBar = str_MessageStart & int_SelectedTradeCtr & " of " & int_NumSelectedRows

        ' Output selected instrument values
        Dim dic_Cache_ActiveIDInfo As Dictionary: Set dic_Cache_ActiveIDInfo = dic_Cache_AllIDInfo(enu_ActiveInstType)
        Dim strArr_ActiveIDInfo() As String
        Dim str_ActiveInstType As String: str_ActiveInstType = GetInstType_String(enu_ActiveInstType)
        For int_TotalRowCtr = 1 To int_NumTotalRows
            strArr_ActiveIDInfo = dic_Cache_ActiveIDInfo(int_TotalRowCtr)

            If strArr_ActiveIDInfo(int_SelectionCol) = "YES" Then
                ' Update status periodically
                int_SelectedTradeCtr = int_SelectedTradeCtr + 1
                If int_SelectedTradeCtr Mod 5 = 0 Then Application.StatusBar = str_MessageStart & int_SelectedTradeCtr & " of " & int_NumSelectedRows


                ' Prepare output sheet
                str_ActiveSheetName = str_ActiveInstType & "_Output_" & rng_ActiveIDTopLeft(int_TotalRowCtr, 1).Value
                If Examine_WorksheetExists(ThisWorkbook, str_ActiveSheetName) = True Then ThisWorkbook.Worksheets(str_ActiveSheetName).Delete
                Set wks_ActiveOutput = ThisWorkbook.Worksheets.Add(, wks_PrevOutput)
                wks_ActiveOutput.Name = str_ActiveSheetName
                Set wks_PrevOutput = wks_ActiveOutput

                ' Output
                Call dic_ActiveStoredInst(strArr_ActiveIDInfo(1)).OutputReport(wks_ActiveOutput)
                End If

        Next int_TotalRowCtr

        Application.StatusBar = False
        Application.DisplayAlerts = True
    End If
End Sub

Private Sub PerformInstRecalc(enu_ActiveInstType As InstType, bln_Shifted_ValDate As Boolean)
    ' Update stored values within the instruments so that they re-value correctly under the currently applied scenario
    Dim var_ActiveObj As Variant
    Dim dic_ActiveStoredInst As Dictionary: Set dic_ActiveStoredInst = dic_AllStoredInst(enu_ActiveInstType)

    ' Handle interest rate shifts
    Select Case enu_ActiveInstType
        Case InstType.IRS, InstType.CFL, InstType.SWT, InstType.BND, InstType.FRA, InstType.FBN
            For Each var_ActiveObj In dic_ActiveStoredInst.Items
                Call var_ActiveObj.HandleUpdate_IRC("ALL")
            Next var_ActiveObj
    End Select

    ' Handle horizon shifts
    If bln_Shifted_ValDate = True Then
        For Each var_ActiveObj In dic_ActiveStoredInst.Items
            Call var_ActiveObj.SetValDate(lng_ValDate)
        Next var_ActiveObj
    End If
End Sub

Private Sub SelectAllInstTrades(enu_ActiveInstType As InstType, bln_Select As Boolean)
    ' ## Set the 'Selected?' field for all trades to YES
    Dim dic_ActiveBooking As Dictionary: Set dic_ActiveBooking = dic_AllBookings(enu_ActiveInstType)
    Dim rng_IDTop As Range: Set rng_IDTop = dic_ActiveBooking(BookingAttribute.IDSelection)
    Dim rng_Selections As Range: Set rng_Selections = Gather_RangeBelow(rng_IDTop(1, 1)).Offset(0, int_SelectionCol - 1)
    If bln_Select = True Then rng_Selections.Value = "YES" Else rng_Selections.ClearContents
End Sub

Private Sub StoreAsBase_Inst(enu_ActiveInstType As InstType)
    ' ## Set the current PnL as base PnL for all selected trades
    Dim dic_ActiveBooking As Dictionary: Set dic_ActiveBooking = dic_AllBookings(enu_ActiveInstType)
    Dim rng_IDTop As Range: Set rng_IDTop = dic_ActiveBooking(BookingAttribute.IDSelection)
    Dim rng_CurrentValsTop As Range: Set rng_CurrentValsTop = dic_ActiveBooking(BookingAttribute.Outputs)
    Dim rng_BaseValTop As Range: Set rng_BaseValTop = dic_ActiveBooking(BookingAttribute.BaseChg)
    Dim rng_AllTradeIDs As Range: Set rng_AllTradeIDs = Gather_RangeBelow(rng_IDTop(1, 1))

    If Not rng_AllTradeIDs Is Nothing Then
        ' Determine column of outputs which contains the PnL
        Dim int_PnLCol As Integer
        Select Case enu_ActiveInstType
            Case InstType.FTB, InstType.FBN: int_PnLCol = 1
            Case Else: int_PnLCol = 3
        End Select

        ' Set the values
        Dim lng_NumTotalRows As Long: lng_NumTotalRows = rng_AllTradeIDs.Rows.count
        Dim lng_TotalRowCtr As Long
        For lng_TotalRowCtr = 1 To lng_NumTotalRows
            If rng_AllTradeIDs(lng_TotalRowCtr, int_SelectionCol).Value = "YES" Then
                rng_BaseValTop(1, 1).Offset(lng_TotalRowCtr - 1, 0).Value = rng_CurrentValsTop.Offset(lng_TotalRowCtr - 1, int_PnLCol - 1).Value2
            End If
        Next lng_TotalRowCtr

        dic_ActiveBooking(BookingAttribute.Sheet).Calculate
    End If
End Sub

Private Sub SetTarget_Inst(enu_ActiveInstType As InstType, ByRef rng_ActiveOutput As Range)
    ' ## Replace trade section of RE Output sheet with info for the selected trades
    ' Preparation
    Dim dic_ActiveBooking As Dictionary: Set dic_ActiveBooking = dic_AllBookings(enu_ActiveInstType)
    Dim rng_IDTop As Range: Set rng_IDTop = dic_ActiveBooking(BookingAttribute.IDSelection)
    Dim rng_AllTradeIDs As Range: Set rng_AllTradeIDs = Gather_RangeBelow(rng_IDTop(1, 1))
    Dim int_NumTotalRows As Integer: int_NumTotalRows = rng_AllTradeIDs.Rows.count
    Dim wks_Source As Worksheet: Set wks_Source = dic_ActiveBooking(BookingAttribute.Sheet)
    Dim str_SheetName As String: str_SheetName = wks_Source.Name
    Dim rng_Top_PnLChg As Range: Set rng_Top_PnLChg = dic_ActiveBooking(BookingAttribute.BaseChg)(1, 2)
    Dim rng_TopLeft_Values As Range: Set rng_TopLeft_Values = dic_ActiveBooking(BookingAttribute.Outputs)(1, 1)
    Dim dic_ValueCols As Dictionary: Set dic_ValueCols = Gather_OutputColIndexes(enu_ActiveInstType)

    ' Loop through selected trades and output
    If Not rng_AllTradeIDs Is Nothing Then
        Dim int_TotalRowCtr As Integer
        For int_TotalRowCtr = 1 To int_NumTotalRows
            If rng_AllTradeIDs(int_TotalRowCtr, int_SelectionCol).Value = "YES" Then
                rng_ActiveOutput(1, 1).Resize(1, 4).Value = rng_AllTradeIDs(int_TotalRowCtr, 1).Resize(1, 4).Value
                rng_ActiveOutput(1, 5).Value = str_SheetName
                rng_ActiveOutput(1, 6).Value = rng_TopLeft_Values.Offset(int_TotalRowCtr - 1, dic_ValueCols("PNL") - 1).Address(False, False)
                rng_ActiveOutput(1, 7).Value = rng_Top_PnLChg.Offset(int_TotalRowCtr - 1, 0).Address(False, False)

                If dic_ValueCols.Exists("DV01") Then
                    rng_ActiveOutput(1, 8).Value = rng_TopLeft_Values.Offset(int_TotalRowCtr - 1, dic_ValueCols("DV01") - 1).Address(False, False)
                Else
                    rng_ActiveOutput(1, 8).Value = "-"
                End If

                If dic_ValueCols.Exists("DV02") Then
                    rng_ActiveOutput(1, 9).Value = rng_TopLeft_Values.Offset(int_TotalRowCtr - 1, dic_ValueCols("DV02") - 1).Address(False, False)
                Else
                    rng_ActiveOutput(1, 9).Value = "-"
                End If

                If dic_ValueCols.Exists("VEGA") Then
                    rng_ActiveOutput(1, 10).Value = rng_TopLeft_Values.Offset(int_TotalRowCtr - 1, dic_ValueCols("VEGA") - 1).Address(False, False)
                Else
                    rng_ActiveOutput(1, 10).Value = "-"
                End If

                Set rng_ActiveOutput = rng_ActiveOutput.Offset(1, 0)
            End If
        Next int_TotalRowCtr
    End If
End Sub