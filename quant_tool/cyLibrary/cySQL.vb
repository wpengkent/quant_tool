Option Explicit
' ## For Murex queries, LTRIM(RTRIM()) is required for text fields to remove padded spaces
Const str_DBOwnerPrefix As String = "MUREXDB."


Public Function SQL_MX_ScenContainer(str_Container As String, int_ScenMin As Integer, int_ScenMax As Integer, _
    int_DataTypeFilter As Integer) As String
    ' ## Query to extract scenario and shift information for a container from the VaR database

    Dim str_SQL As String
    Dim str_OptionalTypeFilter As String
    If int_DataTypeFilter = -1 Then
        str_OptionalTypeFilter = ""
    Else
        str_OptionalTypeFilter = " AND T1.M_SCNTYPE = " & int_DataTypeFilter
    End If

    str_SQL = "SELECT T1.M_SCNNUM, T1.M_SCNTYPE, LTRIM(RTRIM(T1.M_MLABEL0)) AS M_MLABEL0, T1.M_FAMILY0, " _
        & "T1.M_FAMILY1, LTRIM(RTRIM(T1.M_SLABEL0)) AS M_SLABEL0, LTRIM(RTRIM(T1.M_MAT0)) AS M_MAT0, " _
        & "LTRIM(RTRIM(T1.M_OPTMAT)) AS M_OPTMAT, T1.M_DAYSNB, T1.M_STRIKE, T1.M_VALUE, T1.M_VARI, T1.M_VARITYP, " _
        & "T2.M_HDATE0, T2.M_HDATE1, LTRIM(RTRIM(T2.M_DESC0)), LTRIM(RTRIM(T2.M_DESC1)) " _
        & "FROM " & str_DBOwnerPrefix & "SE_VARCS_DBF AS T1 INNER JOIN " & str_DBOwnerPrefix & "SE_VARCM_DBF AS T2 ON T1.M_SCNLABEL = T2.M_SCNLABEL " _
            & "AND T1.M_SCNNUM = T2.M_SCNNUM " _
            & "WHERE T1.M_SCNLABEL = '" & str_Container & "' AND T1.M_SCNNUM >= " & int_ScenMin _
            & " AND T1.M_SCNNUM <= " & int_ScenMax & str_OptionalTypeFilter & " " _
        & "ORDER BY T1.M_SCNNUM, T1.M_SCNTYPE, T1.M_MLABEL0, T1.M_SLABEL0, T1.M_MAT0, T1.M_OPTMAT, T1.M_DAYSNB"

    SQL_MX_ScenContainer = str_SQL
End Function

Public Function SQL_MX_ScenContainerFilter(str_Container As String, str_Filter() As String, int_ScenMin As Integer, int_ScenMax As Integer, _
    int_DataTypeFilter As Integer) As String
    ' ## Query to extract scenario and shift information for a container from the VaR database

    Dim str_SQL As String
    Dim str_OptionalTypeFilter As String
    Dim int_FilterCount As Integer
    Dim str_sqlFilter As String
    Dim int_i As Integer

    int_FilterCount = UBound(str_Filter)

    For int_i = 1 To int_FilterCount
        If int_i = 1 Then
            str_sqlFilter = str_sqlFilter & "AND (LTRIM(RTRIM(T1.M_MLABEL0)) = '" & str_Filter(int_i) & "'"
        Else
            str_sqlFilter = str_sqlFilter & "OR LTRIM(RTRIM(T1.M_MLABEL0)) = '" & str_Filter(int_i) & "'"
        End If
        If int_i = int_FilterCount Then str_sqlFilter = str_sqlFilter & ")"
    Next int_i


    If int_DataTypeFilter = -1 Then
        str_OptionalTypeFilter = ""
    Else
        str_OptionalTypeFilter = " AND T1.M_SCNTYPE = " & int_DataTypeFilter
    End If

    str_SQL = "SELECT T1.M_SCNNUM, T1.M_SCNTYPE, LTRIM(RTRIM(T1.M_MLABEL0)) AS M_MLABEL0, T1.M_FAMILY0, " _
        & "T1.M_FAMILY1, LTRIM(RTRIM(T1.M_SLABEL0)) AS M_SLABEL0, LTRIM(RTRIM(T1.M_MAT0)) AS M_MAT0, " _
        & "LTRIM(RTRIM(T1.M_OPTMAT)) AS M_OPTMAT, T1.M_DAYSNB, T1.M_STRIKE, T1.M_VALUE, T1.M_VARI, T1.M_VARITYP, " _
        & "T2.M_HDATE0, T2.M_HDATE1, LTRIM(RTRIM(T2.M_DESC0)), LTRIM(RTRIM(T2.M_DESC1)) " _
        & "FROM " & str_DBOwnerPrefix & "SE_VARCS_DBF AS T1 INNER JOIN " & str_DBOwnerPrefix & "SE_VARCM_DBF AS T2 ON T1.M_SCNLABEL = T2.M_SCNLABEL " _
            & "AND T1.M_SCNNUM = T2.M_SCNNUM " _
            & "WHERE T1.M_SCNLABEL = '" & str_Container & "' " & str_sqlFilter & " AND T1.M_SCNNUM >= " & int_ScenMin _
            & " AND T1.M_SCNNUM <= " & int_ScenMax & str_OptionalTypeFilter & " " _
        & "ORDER BY T1.M_SCNNUM, T1.M_SCNTYPE, T1.M_MLABEL0, T1.M_SLABEL0, T1.M_MAT0, T1.M_OPTMAT, T1.M_DAYSNB"

    SQL_MX_ScenContainerFilter = str_SQL
End Function


Public Function SQL_MX_CurveMapping(str_Version_MX As String) As String
    ' ## Query to extract mapping from code to curve name from the FO database
    Dim str_SQL As String
    Select Case str_Version_MX
        Case "3.1.21"
            str_SQL = "SELECT LTRIM(RTRIM(T1.M_K_102)), LTRIM(RTRIM(T1.M_LABEL)), LTRIM(RTRIM(T1.M_K_101)) " _
                & "FROM " & str_DBOwnerPrefix & "RT_CURVE_PK_DBF AS T1 ORDER BY T1.M_LABEL"
        Case "3.1.29"
            str_SQL = "SELECT LTRIM(RTRIM(T1.M_LABEL)), LTRIM(RTRIM(T1.M_DLABEL)), LTRIM(RTRIM(T1.M_CURRENCY)) " _
                & "FROM " & str_DBOwnerPrefix & "RT_CT_DBF AS T1 ORDER BY T1.M_LABEL"
    End Select
    SQL_MX_CurveMapping = str_SQL
End Function

Public Function SQL_MX_FutMapping() As String
    ' ## Query to extract list of futures and their codes from the FO database
    Dim str_SQL As String
    str_SQL = "SELECT LTRIM(RTRIM(T1.M_SE_LABEL)), LTRIM(RTRIM(T1.M_SE_D_LABEL)) " _
        & "FROM " & str_DBOwnerPrefix & "SE_HEAD_DBF AS T1 " _
        & "WHERE T1.M_SE_GROUP = 'Future' " _
        & "ORDER BY T1.M_SE_LABEL"

    SQL_MX_FutMapping = str_SQL
End Function

Public Function SQL_MX_SecurityMapping() As String
    ' ## Query to extract list of equities and their codes from the FO database
    Dim str_SQL As String
    str_SQL = "SELECT LTRIM(RTRIM(T1.M_SE_LABEL)), LTRIM(RTRIM(T1.M_SE_D_LABEL)) " _
        & "FROM " & str_DBOwnerPrefix & "SE_HEAD_DBF AS T1 " _
        & "WHERE T1.M_SE_GROUP = 'Equity' " _
        & "ORDER BY T1.M_SE_LABEL"

    SQL_MX_SecurityMapping = str_SQL
End Function

Public Function SQL_MX_ScenContainerInfo() As String
    ' ## Query to extract list of containers and number of scenarios from the VaR database
    Dim str_SQL As String
    str_SQL = "SELECT LTRIM(RTRIM(T1.M_SCNLABEL)), COUNT(T1.M_SCNNUM) " _
        & "FROM " & str_DBOwnerPrefix & "SE_VARCM_DBF AS T1 " _
        & "GROUP BY T1.M_SCNLABEL " _
        & "ORDER BY T1.M_SCNLABEL"

    SQL_MX_ScenContainerInfo = str_SQL
End Function

Public Function SQL_MX_ResultTableInfo() As String
    ' ## Query to extract list of result tables for PSRs
    Dim str_SQL As String
    str_SQL = "SELECT LTRIM(RTRIM(T1.M_LABEL)), LTRIM(RTRIM(T1.M_DESC)) " _
        & "FROM " & str_DBOwnerPrefix & "VAR_CSCF_DBF AS T1 " _
        & "ORDER BY M_LABEL"

    SQL_MX_ResultTableInfo = str_SQL
End Function

Public Function SQL_MX_RateCurve(mqp_Params As MxQP_RateCurve) As String
    ' ## NOT USED AT THE MOMENT
    ' ## Query to extract rate curves (market quotes) from the FO database
    ' ## Values come from MPX_IN_DBF
    Dim str_SQL As String
    Dim str_DataDate As String: str_DataDate = Format(mqp_Params.SystemDate, "yyyymmdd")

    str_SQL = "SELECT DISTINCT LTRIM(RTRIM(T1.M__ALIAS_)), T1.M__DATE_, LTRIM(RTRIM(T4.M_LABEL)), LTRIM(RTRIM(T2.M_TYPE)), " _
        & "LTRIM(RTRIM(T3.M_LABEL)), T3.M_DATE AS [MAT DATE], (T3.M_BID0 + T3.M_ASK0) / 2 AS MID, LTRIM(RTRIM(T2.M_GENERAT)), T1.M_UND_CURVE " _
        & "FROM " & str_DBOwnerPrefix & "MPX_RTC_DBF AS T1 INNER JOIN " & str_DBOwnerPrefix & "RT_CURVE_PK_DBF AS T4 ON T1.M_RT_KEY = T4.M_REFERENCE, " _
            & str_DBOwnerPrefix & "MPY_RTC_DBF AS T2, " & str_DBOwnerPrefix & "MPX_IN_DBF AS T3 " _
        & "WHERE LTRIM(RTRIM(T4.M_LABEL))='" & mqp_Params.Curve & "' and T1.M__DATE_='" & str_DataDate _
            & "' AND T1.M__ALIAS_= '" & mqp_Params.DataSet & "' and T1.M__INDEX_=T2.M__INDEX_ and T1.M_CURRENCY=T3.M_CURRENCY " _
            & "AND T1.M__DATE_=T3.M__DATE_ and T2.M_TYPE=T3.M_TYPE and T2.M_GENINTNB=T3.M_GENINTNB and " _
            & "T2.M_GENERAT=T3.M_GENERAT and T2.M_LABEL_D=T3.M_LABEL " _
        & "ORDER BY [MAT DATE]"

    SQL_MX_RateCurve = str_SQL
End Function

Public Function SQL_MX_ZeroCurve(lng_StartDate As Long, lng_EndDate As Long) As String
    ' ## Query to extract zero rate curves from the VaR database

    Dim str_SQL As String
    Dim str_StartDate As String: str_StartDate = Convert_SQLDate(lng_StartDate)
    Dim str_EndDate As String: str_EndDate = Convert_SQLDate(lng_EndDate)

    str_SQL = "SELECT LTRIM(RTRIM(T1.M_MLABEL0)) AS CurveCode, T2.M_HDATE1 AS Date, T1.M_DAYSNB AS NumDays, T1.M_VALUE + T1.M_VARI AS Rate " _
                & "FROM " & str_DBOwnerPrefix & "SE_VARCS_DBF AS T1 " _
                & "INNER JOIN " & str_DBOwnerPrefix & "SE_VARCM_DBF AS T2 ON T1.M_SCNLABEL = T2.M_SCNLABEL AND T1.M_SCNNUM = T2.M_SCNNUM " _
                & "WHERE T1.M_SCNTYPE = 5 AND T1.M_SCNLABEL = 'HSVAR_SCN_1D' AND T2.M_HDATE1 >= '" & str_StartDate & "' AND T2.M_HDATE1 <= '" & str_EndDate _
                & "' ORDER BY T2.M_HDATE1, T1.M_MLABEL0, T1.M_DAYSNB"

    SQL_MX_ZeroCurve = str_SQL
End Function

Public Function SQL_MX_FXSpots(lng_StartDate As Long, lng_EndDate As Long, str_DataSet As String) As String
    ' ## Query to extract FX spots from the FO database, exclude rates that don't get refreshed for > 100 days and cross pairs
    Dim str_SQL As String

    str_SQL = "SELECT LTRIM(RTRIM(T1.M_REF_QUOT)), T1.M__DATE_, (T1.M_SPOT_RF_B + T1.M_SPOT_RF_A)/2 AS FX_SPOT " _
                & "FROM " & str_DBOwnerPrefix & "MPX_SPOT_DBF AS T1 WHERE T1.M__ALIAS_ = '" & str_DataSet & "' AND T1.M__DATE_ >= '" _
                & Convert_SQLDate(lng_StartDate) & "' AND T1.M__DATE_ <= '" & Convert_SQLDate(lng_EndDate) & "' " _
                & "AND T1.M_REF_QUOT LIKE '%USD%' AND T1.M_SPOT_RF_D = T1.M__DATE_ " _
                & "ORDER BY T1.M__DATE_, T1.M_REF_QUOT"

    SQL_MX_FXSpots = str_SQL
End Function

Public Function SQL_MX_EqSpots(lng_StartDate As Long, lng_EndDate As Long, str_DataSet As String) As String
    ' ## Query to extract equity spots from the FO database
    Dim str_SQL As String

    str_SQL = "SELECT LTRIM(RTRIM(T3.M_SE_D_LABEL)), LTRIM(RTRIM(T1.M_MARKET)), LTRIM(RTRIM(T5.M_SE_CUR)), T1.M__DATE_, (T2.M_BID+T2.M_ASK)/2 AS Spot " _
                & "FROM " & str_DBOwnerPrefix & "MPX_PRIC_DBF AS T1 " _
                & "INNER JOIN " & str_DBOwnerPrefix & "MPY_PRIC_DBF AS T2 ON T1.M__INDEX_ = T2.M__INDEX_ " _
                & "INNER JOIN " & str_DBOwnerPrefix & "SE_HEAD_DBF AS T3 ON T1.M_INSTRUM = T3.M_SE_LABEL " _
                & "INNER JOIN " & str_DBOwnerPrefix & "SE_ROOT_DBF AS T4 ON T3.M_SE_LABEL = T4.M_SE_LABEL and T1.M_MARKET = T4.M_SE_MARKET " _
                & "INNER JOIN " & str_DBOwnerPrefix & "SE_TRDC_DBF AS T5 ON T4.M_SE_TRDCL = T5.M_SE_TRDCL " _
                & "WHERE T1.M__DATE_ >= '" & Convert_SQLDate(lng_StartDate) & "' AND T1.M__DATE_ <= '" & Convert_SQLDate(lng_EndDate) & "' " _
                & "AND T3.M_SE_GROUP = 'Equity' AND T2.M_BID > 0 AND T1.M__ALIAS_ = '" & str_DataSet & "' " _
                & "ORDER BY T1.M__DATE_, T3.M_SE_D_LABEL"

    'Debug.Print str_SQL

    SQL_MX_EqSpots = str_SQL
End Function

Public Function SQL_MX_EqVol(lng_StartDate As Long, lng_EndDate As Long, str_DataSet As String) As String
    ' ## Query to extract equity spots from the FO database
    Dim str_SQL As String

    str_SQL = "SELECT LTRIM(RTRIM(T3.M_SE_D_LABEL)), LTRIM(RTRIM(T1.M_MARKET)), LTRIM(RTRIM(T5.M_SE_CUR)), T1.M__DATE_, (T2.M_CALLBID+T2.M_CALLASK)/2 AS Spot " _
                & "FROM " & str_DBOwnerPrefix & "MPX_VOL_DBF AS T1 " _
                & "INNER JOIN " & str_DBOwnerPrefix & "MPY_VOL_DBF AS T2 ON T1.M__INDEX_ = T2.M__INDEX_ " _
                & "INNER JOIN " & str_DBOwnerPrefix & "SE_HEAD_DBF AS T3 ON T1.M_INSTRUM = T3.M_SE_LABEL " _
                & "INNER JOIN " & str_DBOwnerPrefix & "SE_ROOT_DBF AS T4 ON T3.M_SE_LABEL = T4.M_SE_LABEL and T1.M_MARKET = T4.M_SE_MARKET " _
                & "INNER JOIN " & str_DBOwnerPrefix & "SE_TRDC_DBF AS T5 ON T4.M_SE_TRDCL = T5.M_SE_TRDCL " _
                & "WHERE T1.M__DATE_ >= '" & Convert_SQLDate(lng_StartDate) & "' AND T1.M__DATE_ <= '" & Convert_SQLDate(lng_EndDate) & "' " _
                & "AND T3.M_SE_GROUP = 'Equity' AND T2.M_RELDATE = '12/31/1980' AND T2.M_CALLBID > 0 AND T1.M__ALIAS_ = '" & str_DataSet & "' " _
                & "ORDER BY T1.M__DATE_, T3.M_SE_D_LABEL"

    SQL_MX_EqVol = str_SQL
End Function

Public Function SQL_MX_FXSmile(lng_StartDate As Long, lng_EndDate As Long, str_DataSet As String, _
    strLst_Inclusions As Collection) As String
    ' ## Query to extract FX smile vols from the database.  If inclusions list is empty, return all pairs
    Dim str_SQL As String
    Dim str_OptionalInclusionsFilter As String
    If strLst_Inclusions.Count > 0 Then
        str_OptionalInclusionsFilter = "AND LTRIM(RTRIM(T1.M_INSTRUM)) IN " & Convert_ListToParams(strLst_Inclusions, "'") & " "
    Else
        str_OptionalInclusionsFilter = ""
    End If

    str_SQL = "SELECT LTRIM(RTRIM(T1.M_INSTRUM)) AS CcyPair, T3.M_LABEL AS Pillar, T1.M_REFVAL + T2.M_STRIKE AS Delta, " _
                & "T1.M__DATE_ as DataDate, ROUND((T2.M_CALLBID + T2.M_CALLASK + T5.M_CALLBID + T5.M_CALLASK) / 2, 3) " _
                & "FROM " & str_DBOwnerPrefix & "MPX_SMC_DBF AS T1 " _
                & "INNER JOIN " & str_DBOwnerPrefix & "MPY_SMC_DBF AS T2 ON T1.M__INDEX_ = T2.M__INDEX_ " _
                & "INNER JOIN " & str_DBOwnerPrefix & "OM_MAT_DBF AS T3 ON T2.M_MATCOD = T3.M_CODE " _
                & "INNER JOIN " & str_DBOwnerPrefix & "MPX_VOL_DBF AS T4 ON T1.M_INSTRUM = T4.M_INSTRUM " _
                    & "AND T1.M__DATE_ = T4.M__DATE_ AND T1.M__ALIAS_ = T4.M__ALIAS_ " _
                & "INNER JOIN " & str_DBOwnerPrefix & "MPY_VOL_DBF AS T5 ON T4.M__INDEX_ = T5.M__INDEX_ " _
                    & "AND T2.M_MATCOD = T5.M_MATCOD " _
                & "WHERE T1.M__ALIAS_ = '" & str_DataSet & "' AND T1.M__DATE_ >='" & Convert_SQLDate(lng_StartDate) _
                    & "' AND T1.M__DATE_ <= '" & Convert_SQLDate(lng_EndDate) & "' AND T1.M_INSTYP = 'CU' " & str_OptionalInclusionsFilter _
                & "ORDER BY DataDate, CcyPair, Delta, T2.M_MATCOD"

    SQL_MX_FXSmile = str_SQL
End Function

Public Function SQL_MX_AnnualHols(mqp_Params As MxQP_Hols) As String
    ' ## Query to extract annual holidays from the FO database
    Dim str_SQL As String

    str_SQL = "SELECT T1.M_DATE " _
        & "FROM " & str_DBOwnerPrefix & "CAL_HOL_DBF AS T1 " _
        & "WHERE T1.M_CAL_LABEL = '" & mqp_Params.Calendar & "' AND T1.M_GENERAL = 1" _
        & "ORDER BY T1.M_DATE"

    SQL_MX_AnnualHols = str_SQL
End Function

Public Function SQL_MX_SpecialHols(mqp_Params As MxQP_Hols) As String
    ' ## Query to extract ad-hoc special holidays from the FO database
    Dim str_SQL As String

    str_SQL = "SELECT T1.M_DATE " _
        & "FROM " & str_DBOwnerPrefix & "CAL_HOL_DBF AS T1 " _
        & "WHERE T1.M_CAL_LABEL = '" & mqp_Params.Calendar & "' AND M_GENERAL = 0 AND T1.M_DATE >= '" & Convert_SQLDate(mqp_Params.MinDate) _
            & "' AND T1.M_DATE <= '" & Convert_SQLDate(mqp_Params.MaxDate) _
        & "' ORDER BY T1.M_DATE"

    SQL_MX_SpecialHols = str_SQL
End Function

Public Function SQL_MX_Results(mqp_Params As MxQP_Result) As String
    ' ## Query to extract PnL or market values for an executed report from the VaR database, for scenarios between a specified range of numbers based on a specific run date
    Dim str_DateStamp As String
    If mqp_Params.IsSplitByDate = True Then str_DateStamp = Format(mqp_Params.SystemDate, "ddmmyy")

    ' Optionally filter on trades - if no filter, then array will have index zero.  Otherwise index will start from 1
    Dim str_OptionalFilterTrades As String
    Dim int_TradeCtr As Integer
    If mqp_Params.TradeSet.Count > 0 Then
        Select Case mqp_Params.IncExcl
            Case "INCLUSIVE"
                str_OptionalFilterTrades = "AND T1.M_DEALNUM IN " & Convert_ListToParams(mqp_Params.TradeSet, "") & " "
            Case "EXCLUSIVE"
                str_OptionalFilterTrades = "AND T1.M_DEALNUM NOT IN " & Convert_ListToParams(mqp_Params.TradeSet, "") & " "
        End Select
    Else
        str_OptionalFilterTrades = ""
    End If

    Dim str_Table1 As String: str_Table1 = mqp_Params.ResultTable & "#VR" & str_DateStamp & "_VR1"
    Dim str_Table2 As String: str_Table2 = mqp_Params.ResultTable & "#VR" & str_DateStamp & "_VR2"
    Dim int_ResultType As Integer

    Select Case UCase(mqp_Params.ResultType)
        Case "PNL": int_ResultType = 1
        Case "MARKET VALUE": int_ResultType = 2
    End Select

    ' Build query depending on form of output that the user requests
    Dim str_SQL As String
    Dim str_Fields As String, str_GroupBy As String, str_OrderBy As String
    Select Case UCase(mqp_Params.OutputForm)
        Case "BREAKDOWN"
            str_Fields = "SELECT LTRIM(RTRIM(T1.M_PTFOLIO)), LTRIM(RTRIM(T1.M_TRN_FMLY)), LTRIM(RTRIM(T1.M_TRN_GRP)), " _
                & "LTRIM(RTRIM(T1.M_TRN_TYPE)), T1.M_DEALNUM, T2.M_SCENARIO, T2.M_RESULT, T2.M_RESULTV "
            str_GroupBy = ""
            str_OrderBy = "ORDER BY T1.M_DEALNUM, T2.M_SCENARIO"
        Case "AGGREGATED"
            str_Fields = "SELECT '-', '-', '-', '-', '-', T2.M_SCENARIO, SUM(T2.M_RESULT), SUM(T2.M_RESULTV) "
            str_GroupBy = "GROUP BY T2.M_SCENARIO "
            str_OrderBy = "ORDER BY T2.M_SCENARIO"
        Case "BY PORTFOLIO"
            str_Fields = "SELECT LTRIM(RTRIM(T1.M_PTFOLIO)), '-', '-', '-', '-', T2.M_SCENARIO, SUM(T2.M_RESULT), SUM(T2.M_RESULTV) "
            str_GroupBy = "GROUP BY T2.M_SCENARIO, T1.M_PTFOLIO "
            str_OrderBy = "ORDER BY T1.M_PTFOLIO, T2.M_SCENARIO"
    End Select

    ' Build any trade set formulas
    Dim str_TSFormula As String
    Select Case mqp_Params.TradeSetFormula
        Case "NON-MYR BONDS"
            Select Case mqp_Params.IncExcl
                Case "INCLUSIVE"
                    str_TSFormula = "AND T1.M_CUR <> 'MYR' AND T1.M_TRN_FMLY = 'IRD' AND T1.M_TRN_GRP = 'BOND' "
                Case "EXCLUSIVE"
                    str_TSFormula = "AND NOT (T1.M_CUR <> 'MYR' AND T1.M_TRN_FMLY = 'IRD' AND T1.M_TRN_GRP = 'BOND') "
            End Select
            str_OptionalFilterTrades = ""  ' Disable manual trade set, otherwise the logic in the query would not make sense
    End Select

    str_SQL = str_Fields & "FROM " & str_DBOwnerPrefix & str_Table1 & " AS T1 INNER JOIN " & str_DBOwnerPrefix & str_Table2 _
            & " AS T2 ON T1.M_KEY_ID = T2.M_KEY_ID " _
        & "WHERE M_OTYPE = " & int_ResultType & " AND M_SCENARIO >= " & mqp_Params.ScenMin & " AND M_SCENARIO <= " _
            & mqp_Params.ScenMax & " AND T1.M_DATE = '" & Convert_SQLDate(mqp_Params.SystemDate) & "' " _
            & str_OptionalFilterTrades & str_TSFormula & str_GroupBy & str_OrderBy

    SQL_MX_Results = str_SQL
End Function

Public Function SQL_MX_Nostro() As String
    ' ## Query to extract historical purchase cost for non-MYR bonds, split by portfolio and currency
    Dim str_SQL As String

    str_SQL = "SELECT LTRIM(RTRIM(Portfolio)), Q1.Currency, SUM(Nominal * CleanPrice * BSFactor) / 100 AS Nostro " _
        & "FROM (" _
            & "SELECT T1.M_NB AS TradeNum, T1.M_BRW_NOM1 AS Nominal, " _
                & "CASE T1.M_COMMENT_BS WHEN 'B' THEN T1.M_BPFOLIO ELSE T1.M_SPFOLIO END AS Portfolio, " _
                & "CASE T1.M_COMMENT_BS WHEN 'B' THEN 1 ELSE -1 END AS BSFactor, " _
                & "T1.M_TRN_STATUS as Status, T1.M_BRW_RTE1 as CleanPrice, T1.M_PL_INSCUR as Currency, T1.M_TRN_FMLY, " _
                & "T1.M_TRN_GRP, T1.M_TRN_TYPE " _
            & "FROM " & str_DBOwnerPrefix & "TRN_HDR_DBF AS T1" _
        & ") AS Q1 " _
        & "WHERE Q1.Currency <> 'MYR' and Q1.M_TRN_FMLY = 'IRD' and Q1.M_TRN_GRP = 'BOND' " _
        & "GROUP BY Q1.Portfolio, Q1.Currency"

    SQL_MX_Nostro = str_SQL
End Function

Public Function SQL_MX_NostroBreakdown() As String
    ' ## Query to extract historical purchase cost for non-MYR bonds, split by trade ID

    Dim str_SQL As String
    Dim str_CostField As String
    Dim str_PortfolioField As String

    ' Contingent fields
    str_PortfolioField = "CASE T1.M_COMMENT_BS WHEN 'B' THEN LTRIM(RTRIM(T1.M_BPFOLIO)) WHEN 'S' THEN LTRIM(RTRIM(T1.M_SPFOLIO)) ELSE '-' END AS Portfolio"
    str_CostField = "CASE T1.M_COMMENT_BS WHEN 'B' THEN T1.M_BRW_NOM1 * T1.M_BRW_RTE1 * 0.01 ELSE T1.M_BRW_NOM1 * T1.M_BRW_RTE1 * -0.01 END AS HistCost"

    ' Select necessary fields
    str_SQL = "SELECT T1.M_NB, T1.M_GID, " & str_PortfolioField & ", T1.M_COMMENT_BS, " & str_CostField & ", T1.M_PL_INSCUR " _
        & "FROM " & str_DBOwnerPrefix & "TRN_HDR_DBF AS T1 " _
        & "WHERE T1.M_PL_INSCUR <> 'MYR' and T1.M_TRN_FMLY = 'IRD' and T1.M_TRN_GRP = 'BOND' " _
        & "ORDER BY T1.Portfolio, T1.M_NB"

    ' Filter by portfolio
    SQL_MX_NostroBreakdown = str_SQL
End Function

Public Function SQL_MX_FXSpotEqualsOne(str_VaRDBName As String) As String
    ' ## Find cases where mid rate = 1, output relative shift for user to update scenario with
    ' ## Find scenarios where start date matches exception date, then find scenarios where end date matches exception date

    Dim str_SQL As String
    str_SQL = "SELECT T2.M_SCNLABEL AS CONTAINER, T2.M_SCNNUM AS SCENARIO, T1.M_REF_QUOT AS [CCY PAIR], T2.M_HDATE0 AS [START DATE], T2.M_HDATE1 AS [END DATE], " _
            & "(T1.M_SPOT_RF_B + T1.M_SPOT_RF_A) / 2 AS [START RATE], (T3.M_SPOT_RF_B + T3.M_SPOT_RF_A) / 2 AS [END DATE], " _
            & "100 * ((T3.M_SPOT_RF_B + T3.M_SPOT_RF_A) / (T1.M_SPOT_RF_B + T1.M_SPOT_RF_A) - 1) AS [CORRECT SHIFT %] " _
        & "FROM ((" & str_DBOwnerPrefix & "MPX_SPOT_DBF AS T1 INNER JOIN " & str_VaRDBName & "." & str_DBOwnerPrefix & "SE_VARCM_DBF AS T2 ON T1.M__DATE_ = T2.M_HDATE0) " _
            & "INNER JOIN " & str_DBOwnerPrefix & "MPX_SPOT_DBF AS T3 ON T2.M_HDATE1 = T3.M__DATE_ AND T1.M_REF_QUOT = T3.M_REF_QUOT) " _
        & "WHERE ((T1.M_SPOT_RF_B + T1.M_SPOT_RF_A) / 2 = 1 AND (T3.M_SPOT_RF_B + T3.M_SPOT_RF_A) / (T1.M_SPOT_RF_B + T1.M_SPOT_RF_A) <> 1) AND T1.M_REF_QUOT LIKE '%USD%' AND T1.M_NUM != 'PAB' " _
        & "UNION " _
        & "SELECT T2.M_SCNLABEL AS REPORT, T2.M_SCNNUM AS SCENARIO, T1.M_REF_QUOT AS [CCY PAIR], T2.M_HDATE0 AS [START DATE], T2.M_HDATE1 AS [END DATE], " _
            & "(T3.M_SPOT_RF_B + T3.M_SPOT_RF_A) / 2 AS [START RATE], (T1.M_SPOT_RF_B + T1.M_SPOT_RF_A) / 2 AS [END DATE], " _
            & "100 * ((T1.M_SPOT_RF_B + T1.M_SPOT_RF_A) / (T3.M_SPOT_RF_B + T3.M_SPOT_RF_A) - 1) AS [CORRECT SHIFT %] " _
        & "FROM (" & str_DBOwnerPrefix & "MPX_SPOT_DBF AS T1 INNER JOIN " & str_VaRDBName & "." & str_DBOwnerPrefix & "SE_VARCM_DBF AS T2 ON T1.M__DATE_ = T2.M_HDATE1) " _
            & "INNER JOIN " & str_DBOwnerPrefix & "MPX_SPOT_DBF AS T3 ON T2.M_HDATE0 = T3.M__DATE_ AND T1.M_REF_QUOT = T3.M_REF_QUOT " _
        & "WHERE ((T1.M_SPOT_RF_B + T1.M_SPOT_RF_A) / 2 = 1 AND (T1.M_SPOT_RF_B + T1.M_SPOT_RF_A) / (T3.M_SPOT_RF_B + T3.M_SPOT_RF_A) <> 1) AND T1.M_REF_QUOT LIKE '%USD%' AND T1.M_NUM != 'PAB' " _
        & "ORDER BY M_SCNLABEL, M_SCNNUM"

    SQL_MX_FXSpotEqualsOne = str_SQL
End Function

Public Function SQL_MX_TradeQuery(rng_PortfoliosTop As Range, str_IntExt As String, lng_TradeNum As Long) As String
    ' ## Output details of trades belonging to a specified portfolio
    ' ## Includes trades cancelled and replaced by other trades, which VaR target will remove

    Dim str_SQL As String
    Dim rng_ActivePort As Range: Set rng_ActivePort = rng_PortfoliosTop
    Dim str_PortfoliosFilter As String
    If rng_PortfoliosTop.Value = "" Then
        str_PortfoliosFilter = "Null"
    Else
        While rng_ActivePort.Value <> ""
            str_PortfoliosFilter = str_PortfoliosFilter & "'" & rng_ActivePort.Value & "', "
            Set rng_ActivePort = rng_ActivePort.Offset(1, 0)
        Wend

        str_PortfoliosFilter = Left(str_PortfoliosFilter, Len(str_PortfoliosFilter) - 2)
    End If

    Dim str_TargetFilter_A As String, str_TargetFilter_B As String, str_TargetFilter_C As String
    Select Case lng_TradeNum
        Case -1
            ' Searching by portfolio rather than trade
            str_TargetFilter_A = "M_BPFOLIO IN (" & str_PortfoliosFilter & ")"
            str_TargetFilter_B = "M_SPFOLIO IN (" & str_PortfoliosFilter & ")"
            str_TargetFilter_C = "Portfolio IN (" & str_PortfoliosFilter & ")"
        Case Else
            ' Searching by trade
            str_TargetFilter_A = "M_NB = " & lng_TradeNum
            str_TargetFilter_B = "M_NB = " & lng_TradeNum
            str_TargetFilter_C = "M_NB = " & lng_TradeNum
    End Select

    Select Case UCase(str_IntExt)
        Case "INTERNAL"
            ' Trades between two Murex portfolios rather than with a counterparty
            Dim str_Query_V2 As String, str_Query_V3 As String

            ' Where buy is in scope
            str_Query_V2 = "SELECT U2.M_NB AS TradeNum, LTRIM(RTRIM(U2.M_GID)), LTRIM(RTRIM(U2.M_BPFOLIO)), 'B', U2.M_BRW_NOM1, LTRIM(RTRIM(U2.M_BRW_ODNC0)), " _
                & "LTRIM(RTRIM(U2.M_TRN_FMLY)), LTRIM(RTRIM(U2.M_TRN_GRP)), LTRIM(RTRIM(U2.M_TRN_TYPE)), U2.M_TRN_DATE, LTRIM(RTRIM(U2.M_TRN_STATUS)), U2.M_CONTRACT, LTRIM(RTRIM(U2.M_SPFOLIO)) " _
                & "FROM " & str_DBOwnerPrefix & "TRN_HDR_DBF AS U2 " _
                & "WHERE U2." & str_TargetFilter_A & " AND M_COUNTRPART = 0"

            ' Where sell is in scope
            str_Query_V3 = "SELECT U3.M_NB AS TradeNum, LTRIM(RTRIM(U3.M_GID)), LTRIM(RTRIM(U3.M_SPFOLIO)), 'S', U3.M_BRW_NOM1 * -1, LTRIM(RTRIM(U3.M_BRW_ODNC0)), " _
                & "LTRIM(RTRIM(U3.M_TRN_FMLY)), LTRIM(RTRIM(U3.M_TRN_GRP)), LTRIM(RTRIM(U3.M_TRN_TYPE)), U3.M_TRN_DATE, LTRIM(RTRIM(U3.M_TRN_STATUS)), U3.M_CONTRACT, LTRIM(RTRIM(U3.M_BPFOLIO)) " _
                & "FROM " & str_DBOwnerPrefix & "TRN_HDR_DBF AS U3 " _
                & "WHERE U3." & str_TargetFilter_B & " AND M_COUNTRPART = 0"

            str_SQL = str_Query_V2 & " UNION " & str_Query_V3 & " ORDER BY TradeNum"
        Case "EXTERNAL"
            Dim str_Query_U1 As String, str_Query_V1 As String

            ' Contingent fields
            Dim str_NetPayField As String: str_NetPayField = "CASE T1.M_PAY_NET WHEN 1 THEN -1 ELSE 1 END AS PayRecFactor"
            Dim str_BuySellField As String: str_BuySellField = "CASE T1.M_COMMENT_BS WHEN 'B' THEN 1 ELSE -1 END AS BuySellFactor"
            Dim str_PortfolioField As String: str_PortfolioField = "CASE T1.M_COMMENT_BS WHEN 'B' THEN T1.M_BPFOLIO WHEN 'S' THEN T1.M_SPFOLIO ELSE '-' END AS Portfolio"

            ' Select necessary fields - external trades
            str_Query_U1 = "SELECT T1.M_NB, T1.M_GID, " & str_PortfolioField & ", T1.M_COMMENT_BS, T1.M_BRW_NOM1, T1.M_BRW_ODNC0, " _
                & "T1.M_TRN_FMLY, T1.M_TRN_GRP, T1.M_TRN_TYPE, T1.M_TRN_DATE, T1.M_TRN_STATUS, T1.M_CONTRACT, T1.M_RSKSECTION, T1.M_COUNTRPART, " _
                & str_NetPayField & ", " & str_BuySellField & " " _
                & "FROM " & str_DBOwnerPrefix & "TRN_HDR_DBF AS T1 " _
                & "WHERE M_COUNTRPART <> 0"

            ' Filter by portfolio
            str_Query_V1 = "SELECT * FROM (" & str_Query_U1 & ") AS U1 WHERE U1." & str_TargetFilter_C

            ' Logic for nominal
            Dim str_NominalField As String
            str_NominalField = "CASE V1.M_TRN_GRP WHEN 'SCF' THEN V2.M_FLOW_AMT * V1.PayRecFactor " _
                & "WHEN 'LFUT' THEN V1.M_BRW_NOM1 * V3.M_SE_SEC_LS0 * V1.BuySellFactor ELSE V1.M_BRW_NOM1 * V1.BuySellFactor END AS Nominal"

            ' Final query
            str_SQL = "SELECT V1.M_NB, LTRIM(RTRIM(V1.M_GID)), LTRIM(RTRIM(V1.Portfolio)), LTRIM(RTRIM(V1.M_COMMENT_BS)), " & str_NominalField & ", LTRIM(RTRIM(V1.M_BRW_ODNC0)), " _
                & "LTRIM(RTRIM(V1.M_TRN_FMLY)), LTRIM(RTRIM(V1.M_TRN_GRP)), LTRIM(RTRIM(V1.M_TRN_TYPE)), V1.M_TRN_DATE, LTRIM(RTRIM(V1.M_TRN_STATUS)), V1.M_CONTRACT, V1.M_COUNTRPART " _
                & "FROM (" & str_Query_V1 & ") AS V1 LEFT JOIN " & str_DBOwnerPrefix & "TRN_SCFB_DBF AS V2 ON V1.M_NB = V2.M_NB " _
                & "LEFT JOIN " & str_DBOwnerPrefix & "SE_ROOT_DBF AS V3 ON V1.M_RSKSECTION = V3.M_SE_LABEL " _
                & "ORDER BY V1.M_NB"
    End Select

    SQL_MX_TradeQuery = str_SQL
End Function

Public Function SQL_MX_Portfolios() As String
    ' ## Output the list of simple portfolios within Murex
    Dim str_SQL As String
    str_SQL = "SELECT DISTINCT T1.M_PTF_LABEL " _
        & "FROM " & str_DBOwnerPrefix & "GRP_SPTF_DBF AS T1 " _
        & "ORDER BY T1.M_PTF_LABEL" _

    SQL_MX_Portfolios = str_SQL
End Function

Public Function SQL_MX_PortTree() As String
    ' ## Output MX FO portfolio hierarchy
    Dim str_SQL As String
    str_SQL = "SELECT T1.M_HEIGHT, LTRIM(RTRIM(T1.M_LABEL)), LTRIM(RTRIM(T1.M_FATHER_L)) " _
        & "FROM " & str_DBOwnerPrefix & "MUB#MUB_TREE_DBF AS T1 " _
        & "ORDER BY T1.M_HEIGHT, T1.M_FATHER_L"

    SQL_MX_PortTree = str_SQL
End Function

Public Function SQL_MX_CombinedPorts() As String
    ' ## Output the list of MX combined portfolios
    Dim str_SQL As String
    str_SQL = "SELECT DISTINCT LTRIM(RTRIM(T1.M_LABEL)) " _
        & "FROM " & str_DBOwnerPrefix & "MUB#GRP_COMB_DBF AS T1 " _
        & "ORDER BY T1.M_LABEL"

    SQL_MX_CombinedPorts = str_SQL
End Function

Public Function SQL_MX_CombinedPortDefs() As String
    ' ## Output MX combined portfolio definitions
    Dim str_SQL As String
    str_SQL = "SELECT LTRIM(RTRIM(T1.M_LABEL)), LTRIM(RTRIM(T1.M_UNIT)) " _
        & "FROM " & str_DBOwnerPrefix & "MUB#GRP_COMB_DBF AS T1 " _
        & "ORDER BY T1.M_LABEL, T1.M_UNIT"

    SQL_MX_CombinedPortDefs = str_SQL
End Function

Public Function SQL_MX_BondRCA(str_Label As String, lng_Date As Long, str_MXVersion As String) As String
    ' ## Output rate curve assignments by bond for the specified date and curve assignment group
    Dim str_SQL As String

    Dim str_RateCurveNameCol As String
    Dim str_RateCurveComponent As String

    Select Case str_MXVersion
        Case "3.1.21"
            str_RateCurveNameCol = "T3.M_LABEL"
            str_RateCurveComponent = "RT_CURVE_PK_DBF AS T3 ON T2.M_CURVE = T3.M_K_102 "
        Case "3.1.29"
            str_RateCurveNameCol = "T3.M_DLABEL"
            str_RateCurveComponent = "RT_CT_DBF AS T3 ON T2.M_CURVE = T3.M_LABEL "
    End Select

    str_SQL = "SELECT T1.M_NB, T5.M_SE_CODE, LTRIM(RTRIM(T2.M_CALC)), LTRIM(RTRIM(" & str_RateCurveNameCol & ")) " _
        & "FROM " & str_DBOwnerPrefix & "TRN_HDR_DBF AS T1 " _
        & "INNER JOIN " & str_DBOwnerPrefix & "MPY_ASGRC_DBF AS T2 ON T1.M_RSKSECTION = T2.M_ISSUE " _
        & "INNER JOIN " & str_DBOwnerPrefix & str_RateCurveComponent _
        & "INNER JOIN " & str_DBOwnerPrefix & "MPX_ASGRC_DBF AS T4 ON T2.M__INDEX_ = T4.M__INDEX_ " _
        & "INNER JOIN " & str_DBOwnerPrefix & "SE_HEAD_DBF AS T5 ON T1.M_RSKSECTION = T5.M_SE_LABEL " _
        & "WHERE T1.M_TRN_GRP = 'BOND' AND T4.M_LABEL = '" & str_Label & "' AND T4.M__DATE_ = '" & Convert_SQLDate(lng_Date) & "' " _
        & "ORDER BY T1.M_NB"

    SQL_MX_BondRCA = str_SQL
End Function

Public Function SQL_MX_SystemDates() As String
    ' ## Output MX system date by data set
    Dim str_SQL As String
    str_SQL = "SELECT LTRIM(RTRIM(T1.M_ALIAS)), MAX(T1.M_DATE) " _
        & "FROM " & str_DBOwnerPrefix & "MPX_HIS_DBF AS T1 " _
        & "GROUP BY T1.M_ALIAS " _
        & "ORDER BY T1.M_ALIAS"

    SQL_MX_SystemDates = str_SQL
End Function

Public Function SQL_MX_IRIndexMapping() As String
    ' ## Output MX interest rate indicies and their codes in the Murex DB
    Dim str_SQL As String
    str_SQL = "SELECT LTRIM(RTRIM(T1.M_INDEX)), LTRIM(RTRIM(T1.M_IND_LAB)) " _
        & "FROM " & str_DBOwnerPrefix & "RT_INDEX_DBF AS T1 " _
        & "ORDER BY T1.M_INDEX"

    SQL_MX_IRIndexMapping = str_SQL
End Function

Public Function SQL_MX_FindTablesContainingCol(str_ColName As String) As String
    ' ## Output list of all tables containing the specified column name
    Dim str_SQL As String
    str_SQL = "SELECT T1.name " _
        & "FROM dbo.sysobjects AS T1 INNER JOIN dbo.syscolumns AS T2 ON T1.id = T2.id " _
        & "WHERE T2.name = '" & str_ColName & "'"

    SQL_MX_FindTablesContainingCol = str_SQL
End Function