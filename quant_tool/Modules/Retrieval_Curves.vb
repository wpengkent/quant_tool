Option Explicit

' ## BULK GATHERING FUNCTIONS
Public Function GetAllCurves(bln_DataExists As Boolean, bln_AddIfMissing As Boolean, Optional dic_GlobalStaticInfo As Dictionary = Nothing) As Dictionary
    Dim dic_output As Dictionary: Set dic_output = GetAllCurves_ExceptCVL(bln_DataExists, bln_AddIfMissing, dic_GlobalStaticInfo)
    Call dic_output.Add(CurveType.cvl, GetAllCurves_CVL(bln_DataExists, bln_AddIfMissing, dic_GlobalStaticInfo))
    Set GetAllCurves = dic_output
End Function

Public Function GetAllCurves_ExceptCVL(bln_DataExists As Boolean, bln_AddIfMissing As Boolean, Optional dic_GlobalStaticInfo As Dictionary = Nothing) As Dictionary
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    dic_output.CompareMode = CompareMethod.TextCompare
    If dic_GlobalStaticInfo Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo()
    Call dic_output.Add(CurveType.EQSPT, GetObject_EQSpots(True, dic_GlobalStaticInfo))
    Call dic_output.Add(CurveType.EQVOL, GetObject_EQVols(True, dic_GlobalStaticInfo))
    Call dic_output.Add(CurveType.FXSPT, GetObject_FXSpots(True, dic_GlobalStaticInfo))
    Call dic_output.Add(CurveType.IRC, GetAllCurves_IRC(bln_DataExists, bln_AddIfMissing, dic_GlobalStaticInfo))
    Call dic_output.Add(CurveType.FXV, GetAllCurves_FXV(bln_DataExists, bln_AddIfMissing, dic_GlobalStaticInfo))
    Call dic_output.Add(CurveType.SVL, GetAllCurves_SVL(bln_DataExists, bln_AddIfMissing, dic_GlobalStaticInfo))
    Call dic_output.Add(CurveType.EVL, GetAllCurves_EVL(bln_DataExists, bln_AddIfMissing, dic_GlobalStaticInfo))
    Set GetAllCurves_ExceptCVL = dic_output
End Function

Public Function GetAllCurves_IRC(bln_DataExists As Boolean, bln_AddIfMissing As Boolean, Optional dic_GlobalStaticInfo As Dictionary = Nothing) As Dictionary
    Set GetAllCurves_IRC = GetAllCurvesByType(CurveType.IRC, "GetObject_IRCurve", False, bln_DataExists, bln_AddIfMissing, , dic_GlobalStaticInfo)
End Function

Public Function GetAllCurves_FXV(bln_DataExists As Boolean, bln_AddIfMissing As Boolean, Optional dic_GlobalStaticInfo As Dictionary = Nothing) As Dictionary
    Set GetAllCurves_FXV = GetAllCurvesByType(CurveType.FXV, "GetObject_FXVols", False, bln_DataExists, bln_AddIfMissing, , dic_GlobalStaticInfo)
End Function

Public Function GetAllCurves_CVL(bln_DataExists As Boolean, bln_AddIfMissing As Boolean, Optional dic_GlobalStaticInfo As Dictionary = Nothing) As Dictionary
    Set GetAllCurves_CVL = GetAllCurvesByType(CurveType.cvl, "GetObject_CapVols", False, bln_DataExists, bln_AddIfMissing, , dic_GlobalStaticInfo)
End Function

Public Function GetAllCurves_SVL(bln_DataExists As Boolean, bln_AddIfMissing As Boolean, Optional dic_GlobalStaticInfo As Dictionary = Nothing) As Dictionary
    Set GetAllCurves_SVL = GetAllCurvesByType(CurveType.SVL, "GetObject_SwptVols", False, bln_DataExists, bln_AddIfMissing, , dic_GlobalStaticInfo)
End Function

'## Equity smile get all curve - matt edit
Public Function GetAllCurves_EVL(bln_DataExists As Boolean, bln_AddIfMissing As Boolean, Optional dic_GlobalStaticInfo As Dictionary = Nothing) As Dictionary
    Set GetAllCurves_EVL = GetAllCurvesByType(CurveType.EVL, "GetObject_EQSmile", False, bln_DataExists, bln_AddIfMissing, , dic_GlobalStaticInfo)
End Function


' ## WORKER FUNCTIONS
Private Function GetAllCurvesByType(enu_Type As CurveType, str_GatheringFunc As String, bln_SelectedOnly As Boolean, _
    bln_DataExists As Boolean, bln_AddIfMissing As Boolean, Optional int_SelectionCol As Integer = -1, _
    Optional dic_GlobalStaticInfo As Dictionary = Nothing) As Dictionary
     ' ## Gather all curves in the sheet, can also filter for selected curves only
    Dim fld_AppState_Orig As ApplicationState: fld_AppState_Orig = Gather_ApplicationState(ApplicationStateType.Current)
    Dim fld_AppState_Opt As ApplicationState: fld_AppState_Opt = Gather_ApplicationState(ApplicationStateType.Optimized)
    Call Action_SetAppState(fld_AppState_Opt)

    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    dic_output.CompareMode = CompareMethod.TextCompare
    If dic_GlobalStaticInfo Is Nothing Then Set dic_GlobalStaticInfo = GetStaticInfo()
    Dim rng_Active As Range: Set rng_Active = GetRange_CurveSetup(enu_Type)
    Dim str_ActiveCode As String, bln_InScope As Boolean

    ' Remember original sheet if allowing creation of new sheets for missing curves.  Since if a sheet is created, it becomes the active sheet
    Dim wks_Original As Worksheet
    If bln_AddIfMissing = True Then Set wks_Original = ActiveSheet

    While rng_Active(1, 1).Value <> ""
        ' Determine whether to generate
        bln_InScope = True
        If bln_SelectedOnly = True Then
            If rng_Active(1, int_SelectionCol).Value <> "YES" Then bln_InScope = False
        End If

        ' Generate
        If bln_InScope = True Then
            str_ActiveCode = rng_Active(1, 1).Value
            Call dic_output.Add(str_ActiveCode, Application.Run(str_GatheringFunc, str_ActiveCode, bln_DataExists, _
                bln_AddIfMissing, dic_GlobalStaticInfo))
        End If

        Set rng_Active = rng_Active.Offset(1, 0)
    Wend

    ' Ensure remain on original sheet
    If bln_AddIfMissing = True Then wks_Original.Activate

    Call Action_SetAppState(fld_AppState_Orig)

    Set GetAllCurvesByType = dic_output
End Function