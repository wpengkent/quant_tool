Attribute VB_Name = "Operation_Curves"
Option Explicit

' ## EQUITY SPOTS
Public Sub RefreshEQSpots()
    Dim eqs_Spots As Data_EQSpots: Set eqs_Spots = GetObject_EQSpots(True)
    Call eqs_Spots.LoadRates
End Sub

' ## EQUITY VOLS
Public Sub RefreshEQVols()
    Dim eqs_Vols As Data_EQVols: Set eqs_Vols = GetObject_EQVols(True)
    Call eqs_Vols.LoadRates
End Sub

' ## FX SPOTS
Public Sub RefreshFXSpots()
    Dim fxs_Spots As Data_FXSpots: Set fxs_Spots = GetObject_FXSpots(True)
    Call fxs_Spots.LoadRates
End Sub


' ## FX VOL CURVES
Public Sub GenerateSelectedFXVolCurves(Optional dic_Curves_FXV As Dictionary = Nothing, Optional bln_ReturnToOrigSheet As Boolean = True)
    If dic_Curves_FXV Is Nothing Then Set dic_Curves_FXV = GetAllCurves_FXV(False, True)
    Call OperateOnCurves(dic_Curves_FXV, CurveType.FXV, "cyRefreshFXVolCurve", bln_ReturnToOrigSheet, True, 2)
End Sub

Public Sub GenerateAllFXVolCurves(Optional dic_Curves_FXV As Dictionary = Nothing, Optional bln_ReturnToOrigSheet As Boolean = False)
    If dic_Curves_FXV Is Nothing Then Set dic_Curves_FXV = GetAllCurves_FXV(False, True)
    Call OperateOnCurves(dic_Curves_FXV, CurveType.FXV, "cyRefreshFXVolCurve", bln_ReturnToOrigSheet, False)
End Sub


' ## IR CURVES
Public Sub GenerateSelectedIRCurves(Optional dic_Curves_IRC As Dictionary = Nothing, Optional bln_ReturnToOrigSheet As Boolean = True)
    If dic_Curves_IRC Is Nothing Then Set dic_Curves_IRC = GetAllCurves_IRC(False, True)
    Call OperateOnCurves(dic_Curves_IRC, CurveType.IRC, "cyRefreshIRCurve", bln_ReturnToOrigSheet, True, 2)
End Sub

Public Sub RebootstrapSelectedIRCurves(Optional dic_Curves_IRC As Dictionary = Nothing, Optional bln_ReturnToOrigSheet As Boolean = True)
    If dic_Curves_IRC Is Nothing Then Set dic_Curves_IRC = GetAllCurves_IRC(True, False)
    Call OperateOnCurves(dic_Curves_IRC, CurveType.IRC, "cyRebootstrapIRCurve", bln_ReturnToOrigSheet, True, 2)
End Sub

Public Sub GenerateAllIRCurves(Optional dic_Curves_IRC As Dictionary = Nothing, Optional bln_ReturnToOrigSheet As Boolean = False)
    If dic_Curves_IRC Is Nothing Then Set dic_Curves_IRC = GetAllCurves_IRC(False, True)
    Call OperateOnCurves(dic_Curves_IRC, CurveType.IRC, "cyRefreshIRCurve", bln_ReturnToOrigSheet, False)
End Sub


' ## CAP VOL CURVES
Public Sub GenerateSelectedCapVolCurves(Optional dic_Curves_CVL As Dictionary = Nothing, Optional bln_ReturnToOrigSheet As Boolean = True)
    Dim dic_StaticInfo As Dictionary: Set dic_StaticInfo = GetStaticInfo()
    If dic_Curves_CVL Is Nothing Then Set dic_Curves_CVL = GetAllCurves_CVL(False, True, dic_StaticInfo)
    Dim dic_CurveSet As Dictionary: Set dic_CurveSet = GetAllCurves_ExceptCVL(True, False, dic_StaticInfo)
    Call FillDependency_AllCurves(dic_Curves_CVL, dic_CurveSet)
    Call OperateOnCurves(dic_Curves_CVL, CurveType.cvl, "cyRefreshCapVolCurve", bln_ReturnToOrigSheet, True, 2)
End Sub

Public Sub GenerateAllCapVolCurves(Optional dic_Curves_CVL As Dictionary = Nothing, Optional bln_ReturnToOrigSheet As Boolean = False)
    Dim dic_StaticInfo As Dictionary: Set dic_StaticInfo = GetStaticInfo()
    If dic_Curves_CVL Is Nothing Then Set dic_Curves_CVL = GetAllCurves_CVL(False, True, dic_StaticInfo)
    Dim dic_CurveSet As Dictionary: Set dic_CurveSet = GetAllCurves(True, False, dic_StaticInfo)
    Call FillDependency_AllCurves(dic_Curves_CVL, dic_CurveSet)
    Call OperateOnCurves(dic_Curves_CVL, CurveType.cvl, "cyRefreshCapVolCurve", bln_ReturnToOrigSheet, False)
End Sub

Public Sub UpdateCapVolPillarDates(dic_Curves_CVL As Dictionary)
    ' ## To update dates after changing the valuation date
    Dim var_Active As Variant, cvl_Active As Data_CapVolsQJK
    Dim bln_ActiveGenCapDates As Boolean
    For Each var_Active In dic_Curves_CVL.Items
        Set cvl_Active = var_Active
        Call cvl_Active.GeneratePillarDates(True, cvl_Active.IsBootstrappable, False)
        'QJK added 21102016
        'original
        'Call cvl_Active.Bootstrap(False)
       If (cvl_Active.getStrikeQJK(cvl_Active.getStrCurveName()) = 0) Then
       'do nothing for ATM  ''(curvenameRER As String)
       Else
          Call cvl_Active.Bootstrap_ParVols(False)
       End If
     
    Next var_Active
End Sub


' ## SWAPTION VOL CURVES
Public Sub GenerateSelectedSwptVolCurves(Optional dic_Curves_SVL As Dictionary = Nothing, Optional bln_ReturnToOrigSheet As Boolean = True)
    If dic_Curves_SVL Is Nothing Then Set dic_Curves_SVL = GetAllCurves_SVL(False, True)
    Call OperateOnCurves(dic_Curves_SVL, CurveType.svl, "cyRefreshSwptVolCurve", bln_ReturnToOrigSheet, True, 2)
End Sub

Public Sub GenerateAllSwptVolCurves(Optional dic_Curves_SVL As Dictionary = Nothing, Optional bln_ReturnToOrigSheet As Boolean = False)
    If dic_Curves_SVL Is Nothing Then Set dic_Curves_SVL = GetAllCurves_SVL(False, True)
    Call OperateOnCurves(dic_Curves_SVL, CurveType.svl, "cyRefreshSwptVolCurve", bln_ReturnToOrigSheet, False)
End Sub


'## EQUITY SMILE CURVE - matt edit
Public Sub GenerateSelectedEQSmileCurves(Optional dic_Curves_EVL As Dictionary = Nothing, Optional bln_ReturnToOrigSheet As Boolean = True)
    If dic_Curves_EVL Is Nothing Then Set dic_Curves_EVL = GetAllCurves_EVL(False, True)
    Call OperateOnCurves(dic_Curves_EVL, CurveType.EVL, "cyRefreshEQSmileCurve", bln_ReturnToOrigSheet, True, 2)
End Sub

Public Sub GenerateAllEQSmileCurves(Optional dic_Curves_EVL As Dictionary = Nothing, Optional bln_ReturnToOrigSheet As Boolean = False)
    If dic_Curves_EVL Is Nothing Then Set dic_Curves_EVL = GetAllCurves_EVL(False, True)
    Call OperateOnCurves(dic_Curves_EVL, CurveType.EVL, "cyRefreshEQSmileCurve", bln_ReturnToOrigSheet, False)
End Sub



' ## GENERAL OPERATIONS
Public Sub ClearShocks(dic_CurvesOfType As Dictionary)
    Call OperateOnCurves_Method(dic_CurvesOfType, "Scen_ApplyBase")
End Sub

Public Sub ApplyListedShocks(dic_CurvesOfType As Dictionary)
    Call OperateOnCurves_Method(dic_CurvesOfType, "Scen_ApplyCurrent")
End Sub

Public Sub SetCurveValDate(dic_CurvesOfType As Dictionary, lng_ValDate As Long)
    Call OperateOnCurves_Method(dic_CurvesOfType, "SetValDate", lng_ValDate)
End Sub

Public Sub ResetCache_Lookups(dic_CurvesOfType As Dictionary)
    Call OperateOnCurves_Method(dic_CurvesOfType, "ResetCache_Lookups")
End Sub

Public Sub FillDependency_FXS(dic_CurvesOfType As Dictionary, fxs_Input As Data_FXSpots)
    Call OperateOnCurves_Method(dic_CurvesOfType, "FillDependency_FXS", fxs_Input)
End Sub

Public Sub FillDependency_IRC(dic_CurvesOfType As Dictionary, dic_IRCurves As Dictionary)
    Call OperateOnCurves_Method(dic_CurvesOfType, "FillDependency_IRC", dic_IRCurves)
End Sub

Public Sub FillDependency_AllCurves(dic_CurvesOfType As Dictionary, dic_CurveSet As Dictionary)
    Call OperateOnCurves_Method(dic_CurvesOfType, "FillDependency_AllCurves", dic_CurveSet)
End Sub

Public Sub FillAllDependencies(dic_CurveSet As Dictionary)
    ' ## Set instance relationships between curves
    Dim fxs_Spots As Data_FXSpots: Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
    Call fxs_Spots.FillDependency_IRC(dic_CurveSet(CurveType.IRC))
    Call FillDependency_FXS(dic_CurveSet(CurveType.FXV), fxs_Spots)
    Call FillDependency_IRC(dic_CurveSet(CurveType.FXV), dic_CurveSet(CurveType.IRC))
    Call FillDependency_AllCurves(dic_CurveSet(CurveType.cvl), dic_CurveSet)
End Sub
Public Sub FillAllDependencies_HW(dic_CurveSet As Dictionary)
    ' ## Set instance relationships between curves
    Dim fxs_Spots As Data_FXSpots: Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
    Call fxs_Spots.FillDependency_IRC(dic_CurveSet(CurveType.IRC))
End Sub


' ## WORKER OPERATIONS
Private Sub OperateOnCurves(dic_CurveSet As Dictionary, enu_Type As CurveType, str_SubName As String, bln_ReturnToOrigSheet As Boolean, _
    bln_SelectedOnly As Boolean, Optional int_SelectionCol As Integer = -1)
    ' ## Download and rebuild selected curves as specified by 'YES' values in the selection column
    Dim fld_AppState_Orig As ApplicationState: fld_AppState_Orig = Gather_ApplicationState(ApplicationStateType.Current)
    Dim fld_AppState_Opt As ApplicationState: fld_AppState_Opt = Gather_ApplicationState(ApplicationStateType.Optimized)
    Call Action_SetAppState(fld_AppState_Opt)
    
    Dim wks_Orig As Worksheet: If bln_ReturnToOrigSheet = True Then Set wks_Orig = ActiveSheet
    Dim rng_Active As Range: Set rng_Active = GetRange_CurveSetup(enu_Type)
    Dim str_ActiveCode As String
    Dim bln_InScope As Boolean
    While rng_Active(1, 1).Value <> ""
        ' Determine whether to generate
        bln_InScope = True
        If bln_SelectedOnly = True Then
            If rng_Active(1, int_SelectionCol).Value <> "YES" Then bln_InScope = False
        End If
        
        ' Generate
        If bln_InScope = True Then
            str_ActiveCode = rng_Active(1, 1).Value
            Call Application.Run(str_SubName, str_ActiveCode, dic_CurveSet)
        End If
        
        Set rng_Active = rng_Active.Offset(1, 0)
    Wend
    
    If bln_ReturnToOrigSheet = True Then Call GotoSheet(wks_Orig.Name)
    Call Action_SetAppState(fld_AppState_Orig)
End Sub

Public Sub OperateOnCurves_Method(dic_Curves As Dictionary, str_Method As String, ParamArray arr_Args() As Variant)
    ' ## Run the specified method on all the curves specified
    Dim fld_AppState_Orig As ApplicationState: fld_AppState_Orig = Gather_ApplicationState(ApplicationStateType.Current)
    Dim fld_AppState_Opt As ApplicationState: fld_AppState_Opt = Gather_ApplicationState(ApplicationStateType.Optimized)
    Call Action_SetAppState(fld_AppState_Opt)

    Dim int_NumArgs As Integer: int_NumArgs = UBound(arr_Args) + 1
    Dim var_Active As Variant
    For Each var_Active In dic_Curves.Items
        Select Case int_NumArgs
            Case 0: Call CallByName(var_Active, str_Method, VbCallType.VbMethod)
            Case 1: Call CallByName(var_Active, str_Method, VbCallType.VbMethod, arr_Args(0))
            Case 2: Debug.Assert False
        End Select
    Next var_Active
    
    Call Action_SetAppState(fld_AppState_Orig)
End Sub
