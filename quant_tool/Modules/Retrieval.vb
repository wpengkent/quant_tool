Option Explicit

Public Function GetStaticInfo() As Dictionary
    ' ## Retrieve complete set of static information
    Dim dic_output As New Dictionary
    dic_output.CompareMode = CompareMethod.TextCompare
    Call dic_output.Add(StaticInfoType.ConfigSheet, GetObject_ConfigSheet())
    Call dic_output.Add(StaticInfoType.CalendarSet, GetObject_CalendarSet())
    Call dic_output.Add(StaticInfoType.IRGeneratorSet, GetObject_IRGeneratorSet())
    Call dic_output.Add(StaticInfoType.YieldGeneratorSet, GetObject_YieldGeneratorSet())
    Call dic_output.Add(StaticInfoType.DateShifterSet, GetObject_DateShifterSet())
    Call dic_output.Add(StaticInfoType.IRQuerySet, GetObject_IRQuerySet())
    Call dic_output.Add(StaticInfoType.MappingRules, GetObject_MappingRules())

    Set GetStaticInfo = dic_output
End Function

Public Function GetCurveType(str_type As String) As CurveType
    Dim enu_Output As CurveType
    Select Case UCase(str_type)
        Case "EQSPT": enu_Output = CurveType.EQSPT
        Case "FXSPT": enu_Output = CurveType.FXSPT
        Case "FXV": enu_Output = CurveType.FXV
        Case "IRC": enu_Output = CurveType.IRC
        Case "CVL": enu_Output = CurveType.cvl
        Case "SVL": enu_Output = CurveType.SVL
        Case "EQVOL": enu_Output = CurveType.EQVOL
        Case "EVL": enu_Output = CurveType.EVL
        Case Else: Debug.Assert False
    End Select

    GetCurveType = enu_Output
End Function

Public Function GetCurveTypeName(enu_Type As CurveType) As String
    Dim str_Output As String
    Select Case enu_Type
        Case CurveType.EQSPT: str_Output = "EQSPT"
        Case CurveType.FXSPT: str_Output = "FXSPT"
        Case CurveType.FXV: str_Output = "FXV"
        Case CurveType.IRC: str_Output = "IRC"
        Case CurveType.cvl: str_Output = "CVL"
        Case CurveType.SVL: str_Output = "SVL"
        Case CurveType.EQVOL: str_Output = "EQVOL"
        Case CurveType.EVL: str_Output = "EVL"
        Case Else: Debug.Assert False
    End Select

    GetCurveTypeName = str_Output
End Function


Public Function BuildDaysShifts_IRCurve(enu_CurveState As CurveState_IRC) As CurveDaysShift
    Dim csh_Output As New CurveDaysShift

    Select Case enu_CurveState
        Case CurveState_IRC.Zero_Up1BP
            Call csh_Output.Initialize(ShockType.Absolute)
            Call csh_Output.AddUniformShift(0.01)
        Case CurveState_IRC.Zero_Down1BP
            Call csh_Output.Initialize(ShockType.Absolute)
            Call csh_Output.AddUniformShift(-0.01)
        Case Else: Debug.Assert False
    End Select

    Set BuildDaysShifts_IRCurve = csh_Output
End Function