Option Explicit

Private Enum QueryComponent
    ColSet = 1
    DataDate
    TableName
    FilterSet
    GroupBy
End Enum


' ## MEMBER DATA
Private wks_Location As Worksheet, fxs_Spots As Data_FXSpots, cas_Calendars As CalendarSet, map_Rules As MappingRules
Private rng_Names_TopLeft As Range
Private dic_InstParams As Dictionary, dic_Cache_SQL As Dictionary, dic_DBTables As Dictionary
Private Const int_NumCols As Integer = 15


' ## INITIALIZATION
Public Sub Initialize(wks_Input As Worksheet, cas_Input As CalendarSet, map_Input As MappingRules)
    Set wks_Location = wks_Input
    Set cas_Calendars = cas_Input
    Set map_Rules = map_Input
    Set rng_Names_TopLeft = wks_Input.Range("A3")
    Call FillInstParamsDict
    Set dic_DBTables = map_Rules.Dict_SourceTables
End Sub


' ## METHODS - LOOKUP
Public Function Lookup_SQL(str_CurveName As String, lng_DataDate As Long) As String
    Dim str_Output As String
    Dim str_CacheKey As String: str_CacheKey = str_CurveName & "|" & lng_DataDate

    If dic_Cache_SQL.Exists(str_CacheKey) Then
        ' Use cached SQL
        str_Output = dic_Cache_SQL(str_CacheKey)
    Else
        ' Query-specific values
        Dim arrLst_FoundInsts As Collection, var_ActiveRow As Variant, str_ActiveTable As String
        Dim bln_ActiveIsAverage As Boolean
        Dim dic_QueriesReq As New Dictionary: dic_QueriesReq.CompareMode = CompareMethod.TextCompare
        Dim dic_ActiveQuery As Dictionary, strLst_ActiveFilters As Collection
        Dim str_ActiveQueryKey As String

        ' Filter-specific values
        Dim bln_ActiveIsFutures As Boolean, bln_ActiveIsFwdPt As Boolean, bln_ConvertOutrightFwd As Boolean
        Dim str_ActiveInst As String, int_ActiveMax As Integer, strLst_ActiveExcl As Collection
        Dim str_ActiveCcy As String
        Dim lng_ActiveFormFactor As Long, str_OptionalFut As String, dbl_DeliverableSpot As Double
        Dim str_SQL_ActiveInst As String, str_SQL_ActiveCCY As String, str_SQL_OptionalCurveName As String
        Dim str_SQL_ActiveMin As String, str_SQL_ActiveMax As String, str_SQL_ActiveOptionalExcl As String
        Dim str_SQL_OptionalFutMonth As String
        Dim strLst_ActiveDetail As Collection, str_ActiveRateFormula As String
        Dim lng_ActiveLastFut As Long, str_ActiveFutRule As String, cal_Active As Calendar

        ' Gather instrument set for the specified curve
        Debug.Assert dic_InstParams.Exists(str_CurveName)
        Set arrLst_FoundInsts = dic_InstParams(str_CurveName)

        ' Parse each instrument
        For Each var_ActiveRow In arrLst_FoundInsts
            ' Gather information used frequently
            Set strLst_ActiveDetail = Convert_SplitToList(CStr(var_ActiveRow(4)), "|")
            bln_ActiveIsAverage = (strLst_ActiveDetail.count > 1)
            str_ActiveCcy = var_ActiveRow(2)
            bln_ActiveIsFutures = var_ActiveRow(8)
            bln_ActiveIsFwdPt = var_ActiveRow(12)
            If bln_ActiveIsFwdPt = True Then bln_ConvertOutrightFwd = var_ActiveRow(14)
            str_SQL_OptionalFutMonth = ""

            ' Determine instrument type and table containing the data
            str_ActiveInst = var_ActiveRow(3)
            Debug.Assert dic_DBTables.Exists(str_ActiveInst)
            str_ActiveTable = dic_DBTables(str_ActiveInst)
            str_ActiveQueryKey = str_ActiveTable & "|" & bln_ActiveIsAverage & "|" & bln_ActiveIsFwdPt  ' Determines what requires a separate query

            ' Obtain info for the current query
            If dic_QueriesReq.Exists(str_ActiveQueryKey) Then
                Set dic_ActiveQuery = dic_QueriesReq(str_ActiveQueryKey)
            Else
                Set dic_ActiveQuery = New Dictionary
                dic_ActiveQuery.CompareMode = CompareMethod.TextCompare

                ' Determine value to extract
                If bln_ActiveIsFwdPt = True Then
                    If bln_ConvertOutrightFwd = True Then
                        If fxs_Spots Is Nothing Then Set fxs_Spots = GetObject_FXSpots(True)
                        dbl_DeliverableSpot = fxs_Spots.Lookup_NativeSpot(CStr(var_ActiveRow(15)), True)
                        str_ActiveRateFormula = "(Rate - " & dbl_DeliverableSpot & ") * 10000"
                    Else
                        lng_ActiveFormFactor = var_ActiveRow(13)
                        str_ActiveRateFormula = "Rate * " & (10000 / lng_ActiveFormFactor)
                    End If
                Else
                    str_ActiveRateFormula = "Rate"
                End If

                If bln_ActiveIsAverage = True Then str_ActiveRateFormula = "AVG(" & str_ActiveRateFormula & ")"

                ' Store columns to select
                If bln_ActiveIsFutures = True Then
                    Call dic_ActiveQuery.Add(QueryComponent.ColSet, "SELECT [Data Date], Currency, 'IRFUTB', Null AS Term, " _
                        & "Null AS SortTerm, FutMat, " & str_ActiveRateFormula & ", Daycount, BDC, [Spot Days], " _
                        & "'NONE' AS [Flows Freq], [Pmt Calendar], Null AS [Est Calendar]")
                Else
                    Call dic_ActiveQuery.Add(QueryComponent.ColSet, "SELECT [Data Date], Currency, [Mapped Type], Term, " _
                        & "SortTerm, IIf(True,Null,#01/01/1900#) AS FutMat, " & str_ActiveRateFormula & ", Daycount, BDC, [Spot Days], " _
                        & "[Flows Freq], [Pmt Calendar], [Est Calendar]")
                End If

                ' Store source table
                Call dic_ActiveQuery.Add(QueryComponent.TableName, " FROM " & str_ActiveTable)

                ' Store data date
                Call dic_ActiveQuery.Add(QueryComponent.DataDate, " WHERE [Data Date] = #" & Convert_SQLDate(lng_DataDate) & "#")

                ' Initialize filter list
                Call dic_ActiveQuery.Add(QueryComponent.FilterSet, New Collection)

                ' Store current query in the set of queries
                Call dic_QueriesReq.Add(str_ActiveQueryKey, dic_ActiveQuery)
            End If

            ' Gather filter specific values
            str_SQL_ActiveOptionalExcl = ""
            If var_ActiveRow(7) <> "-" And bln_ActiveIsFutures = False Then
                str_SQL_ActiveOptionalExcl = " AND SortTerm NOT IN " & Convert_ListToParams(Convert_SplitToList(CStr(var_ActiveRow(7)), "|"))
            End If

            If bln_ActiveIsFutures = True Then
                str_SQL_ActiveInst = ""
                str_SQL_ActiveCCY = "Currency = '" & str_ActiveCcy & "'"
            Else
                str_SQL_ActiveInst = "[Mapped Type] = '" & str_ActiveInst & "'"
                str_SQL_ActiveCCY = " AND Currency = '" & str_ActiveCcy & "'"
            End If

            If var_ActiveRow(4) = "" Then
                str_SQL_OptionalCurveName = ""
            ElseIf bln_ActiveIsAverage = True Then
                str_SQL_OptionalCurveName = " AND CurveName IN " & Convert_ListToParams(strLst_ActiveDetail, "'")
            Else
                str_SQL_OptionalCurveName = " AND CurveName = '" & strLst_ActiveDetail(1) & "'"
            End If

            ' Determine bounds for current instrument
            If var_ActiveRow(6) = "-" Then int_ActiveMax = 30000 Else int_ActiveMax = var_ActiveRow(6)

            ' Filter based on bounds.  For futures, only upper bound is used
            If bln_ActiveIsFutures = True Then
                cal_Active = cas_Calendars.Lookup_Calendar(CStr(var_ActiveRow(10)))
                str_ActiveFutRule = var_ActiveRow(9)

                If lng_DataDate = Date_NextFutMat(lng_DataDate, str_ActiveFutRule, 0, cal_Active) Then
                    ' Query will capture the specified number of futures maturities, including the one falling on the data date
                    lng_ActiveLastFut = Date_NextFutMat(lng_DataDate, str_ActiveFutRule, int_ActiveMax - 1, cal_Active)
                Else
                    ' Query will capture the specified number futures maturities, all in the future
                    lng_ActiveLastFut = Date_NextFutMat(lng_DataDate, str_ActiveFutRule, int_ActiveMax, cal_Active)
                End If

                str_SQL_ActiveMin = " AND FutMat >= #" & Convert_SQLDate(lng_DataDate) & "#"
                str_SQL_ActiveMax = " AND FutMat <= #" & Convert_SQLDate(lng_ActiveLastFut) & "#"

                ' Only allow contracts maturing in the specified months
                If var_ActiveRow(11) <> "-" Then
                    str_SQL_OptionalFutMonth = "AND Month(FutMat) IN " & Convert_ListToParams(Convert_SplitToList(CStr(var_ActiveRow(11)), "|"))
                End If
            Else
                If var_ActiveRow(5) = "-" Then str_SQL_ActiveMin = "" Else str_SQL_ActiveMin = " AND SortTerm >= " & var_ActiveRow(5)
                If var_ActiveRow(6) = "-" Then str_SQL_ActiveMax = "" Else str_SQL_ActiveMax = " AND SortTerm <= " & int_ActiveMax
            End If

            ' Add group by clause for averaged queries
            If bln_ActiveIsAverage = True Then
                If bln_ActiveIsFutures = True Then
                    Call dic_ActiveQuery.Add(QueryComponent.GroupBy, " GROUP BY [Data Date], Currency, FutMat, Daycount, BDC, " _
                        & "[Spot Days], [Pmt Calendar]")
                Else
                    Call dic_ActiveQuery.Add(QueryComponent.GroupBy, " GROUP BY [Data Date], Currency, [Mapped Type], Term, " _
                        & "SortTerm, Daycount, BDC, [Spot Days], [Flows Freq], [Pmt Calendar], [Est Calendar]")
                End If
            Else
                If dic_ActiveQuery.Exists(QueryComponent.GroupBy) = False Then Call dic_ActiveQuery.Add(QueryComponent.GroupBy, "")
            End If

            ' Store finished filter
            Set strLst_ActiveFilters = dic_ActiveQuery(QueryComponent.FilterSet)
            Call strLst_ActiveFilters.Add("(" & str_SQL_ActiveInst & str_SQL_ActiveCCY & str_SQL_OptionalCurveName _
                & str_SQL_ActiveMin & str_SQL_ActiveMax & str_SQL_OptionalFutMonth & str_SQL_ActiveOptionalExcl & ")")
        Next var_ActiveRow

        ' Build final query
        Dim str_SQL_ActiveFilters As String
        Dim bln_IsFirstQuery As Boolean: bln_IsFirstQuery = True
        Dim bln_IsFirstFilter As Boolean
        Dim var_ActiveQuery As Variant, var_ActiveFilterSet As Variant

        For Each var_ActiveQuery In dic_QueriesReq.Items
            ' Determine filters applicable to the current query
            bln_IsFirstFilter = True
            For Each var_ActiveFilterSet In var_ActiveQuery(QueryComponent.FilterSet)
                If bln_IsFirstFilter = True Then
                    str_SQL_ActiveFilters = " AND (" & var_ActiveFilterSet
                    bln_IsFirstFilter = False
                Else
                    str_SQL_ActiveFilters = str_SQL_ActiveFilters & " OR " & var_ActiveFilterSet
                End If
            Next var_ActiveFilterSet
            str_SQL_ActiveFilters = str_SQL_ActiveFilters & ")"

            ' Incorporate the current query in the final query
            If bln_IsFirstQuery = True Then
                str_Output = ""
                bln_IsFirstQuery = False
            Else
                str_Output = str_Output & " UNION "
            End If

            str_Output = str_Output & var_ActiveQuery(QueryComponent.ColSet) & var_ActiveQuery(QueryComponent.TableName) _
                & var_ActiveQuery(QueryComponent.DataDate) & str_SQL_ActiveFilters & var_ActiveQuery(QueryComponent.GroupBy)
        Next var_ActiveQuery

        ' Ensure query is sorted
        str_Output = str_Output & " ORDER BY SortTerm;"

        ' Store query in cache
        Call dic_Cache_SQL.Add(str_CacheKey, str_Output)
    End If

    Lookup_SQL = str_Output
End Function


' ## METHODS - SUPPORT
Private Sub FillInstParamsDict()
    ' ## Read information from sheet and store in a dictionary catalogued by curve name
    Set dic_InstParams = New Dictionary
    dic_InstParams.CompareMode = CompareMethod.TextCompare
    Dim int_NumRows As Integer: int_NumRows = Examine_NumRows(rng_Names_TopLeft)
    Dim varArr_Data() As Variant: varArr_Data = rng_Names_TopLeft.Resize(int_NumRows, int_NumCols).Value

    ' Add each instrument to the relevant collection stored in the dictionary
    Dim int_RowCtr As Integer, int_ColCtr As Integer
    Dim str_ActiveCurve As String, arrLst_FoundCurve As Collection, varArr_ActiveRow(1 To int_NumCols) As Variant
    For int_RowCtr = 1 To int_NumRows
        ' Gather current row to add
        For int_ColCtr = 1 To int_NumCols
            varArr_ActiveRow(int_ColCtr) = varArr_Data(int_RowCtr, int_ColCtr)
        Next int_ColCtr

        ' Gather set of instruments for the specified curve
        str_ActiveCurve = varArr_Data(int_RowCtr, 1)
        If dic_InstParams.Exists(str_ActiveCurve) Then
            Set arrLst_FoundCurve = dic_InstParams(str_ActiveCurve)
        Else
            Set arrLst_FoundCurve = New Collection
            Call dic_InstParams.Add(str_ActiveCurve, arrLst_FoundCurve)
        End If

        ' Add current row into the set of instruments
        Call arrLst_FoundCurve.Add(varArr_ActiveRow)
    Next int_RowCtr

    Call ResetCache
End Sub

Private Sub ResetCache()
    ' ## Clear stored SQL queries
    Set dic_Cache_SQL = New Dictionary
    dic_Cache_SQL.CompareMode = CompareMethod.TextCompare
End Sub