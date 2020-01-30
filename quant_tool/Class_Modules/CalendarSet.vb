Option Explicit

' ## MEMBER DATA
Const int_RowOffset_Weekends As Integer = 1, int_RowOffset_HolDates As Integer = 2
Private rng_TopLeft As Range
Private dic_Cache_Calendars As Dictionary


' ## INITIALIZATION
Public Sub Initialize(wks_Input As Worksheet)
    Set rng_TopLeft = wks_Input.Range("A1")
    Set dic_Cache_Calendars = New Dictionary
End Sub


' ## METHODS - LOOKUP
Public Function Lookup_Calendar(str_CalendarName As String) As Calendar
    ' ## Return the calendar object based on the specified calendar name
    Dim arr_CachedLine() As Variant
    Dim cal_Output As Calendar

    If dic_Cache_Calendars.Exists(str_CalendarName) Then
        ' Read output from cache
        arr_CachedLine = dic_Cache_Calendars(str_CalendarName)
        Set cal_Output.HolDates = arr_CachedLine(1)
        cal_Output.Weekends = arr_CachedLine(2)
    Else
        Dim rng_ActiveCal As Range: Set rng_ActiveCal = rng_TopLeft
        Dim int_NumCals As Integer: int_NumCals = Examine_NumCols(rng_TopLeft)

        Dim int_ctr As Integer
        For int_ctr = 1 To int_NumCals
            If rng_TopLeft.Offset(0, int_ctr - 1).Value = str_CalendarName Then
                ' Read output from sheet
                Set cal_Output.HolDates = Gather_RangeBelow(rng_TopLeft.Offset(int_RowOffset_HolDates, int_ctr - 1))
                cal_Output.Weekends = rng_ActiveCal.Offset(int_RowOffset_Weekends, int_ctr - 1).Value

                ' Store in cache
                ReDim arr_CachedLine(1 To 2) As Variant
                Set arr_CachedLine(1) = cal_Output.HolDates
                arr_CachedLine(2) = cal_Output.Weekends
                Call dic_Cache_Calendars.Add(str_CalendarName, arr_CachedLine)
                Exit For

            End If
        Next int_ctr
    End If

    Debug.Assert Not cal_Output.HolDates Is Nothing
    Lookup_Calendar = cal_Output
End Function


' ## METHODS - OPERATIONS
Public Sub Fill_UnionCalendars()
    ' ## Fill in dates for calendars which are the union of other calendars
    Dim fld_AppState_Orig As ApplicationState: fld_AppState_Orig = Gather_ApplicationState(ApplicationStateType.Current)
    Dim fld_AppState_Opt As ApplicationState: fld_AppState_Opt = Gather_ApplicationState(ApplicationStateType.Optimized)
    Call Action_SetAppState(fld_AppState_Opt)

    ' Preparation
    Dim int_NumCols As Integer: int_NumCols = Examine_NumCols(rng_TopLeft)
    Dim strLst_Names As Collection: Set strLst_Names = Convert_RangeToList(rng_TopLeft.Resize(1, int_NumCols))
    Dim str_ActiveName As String
    Call ClearCache

    ' Find index of NZB and construct the holidays from its components
    Dim int_FoundIndex_NZB As Integer: int_FoundIndex_NZB = Examine_FindIndex(strLst_Names, "NZB")
    If int_FoundIndex_NZB <> -1 Then Call OutputUnionHolidays(rng_TopLeft, "NZB", int_FoundIndex_NZB)

    ' Create union holiday list for other calendars
    Dim int_ColCtr As Integer
    For int_ColCtr = 1 To int_NumCols
        str_ActiveName = strLst_Names(int_ColCtr)
        If Len(str_ActiveName) > 3 Then
            Call OutputUnionHolidays(rng_TopLeft, str_ActiveName, int_ColCtr)
        End If
    Next int_ColCtr

    Call Action_SetAppState(fld_AppState_Orig)
End Sub

Private Sub OutputUnionHolidays(rng_TopLeft As Range, str_CalendarName As String, int_Col As Integer)
    ' ## Create and output union holiday list for the specified column
    ' Clear holidays
    Dim rng_BodyTop As Range: Set rng_BodyTop = rng_TopLeft.Offset(int_RowOffset_HolDates, int_Col - 1)
    Call Action_ClearBelow(rng_BodyTop, 1)

    ' Handle special case for NZB
    Dim strArr_Calendars() As String
    If str_CalendarName = "NZB" Then
        strArr_Calendars = Split("WEB_AUB", "_")
    Else
        strArr_Calendars = Split(str_CalendarName, "_")
    End If

    ' Copy holidays from component calendars
    Dim str_ActiveComponent As Variant, cal_ActiveComponent As Calendar
    For Each str_ActiveComponent In strArr_Calendars
        Debug.Assert Len(str_ActiveComponent) = 3
        cal_ActiveComponent = Me.Lookup_Calendar(CStr(str_ActiveComponent))
        Gather_RowForAppend(rng_TopLeft.Offset(0, int_Col - 1)).Resize(cal_ActiveComponent.HolDates.Rows.count, 1).Value = cal_ActiveComponent.HolDates.Value
    Next str_ActiveComponent

    ' Remove duplicates
    Dim int_NumRows As Integer: int_NumRows = Examine_NumRows(rng_BodyTop)
    Call rng_BodyTop.Resize(int_NumRows, 1).Sort(Key1:=rng_BodyTop(1, 1), Order1:=xlAscending, Header:=xlNo, Orientation:=xlSortColumns)
    Call rng_BodyTop.Resize(int_NumRows, 1).RemoveDuplicates(1, Header:=xlNo)

    ' Format range
    Dim rng_Final As Range: Set rng_Final = Gather_RangeBelow(rng_BodyTop)
    rng_Final.NumberFormat = Gather_DateFormat()
    rng_Final.HorizontalAlignment = xlCenter
End Sub


Private Sub ClearCache()
    Call dic_Cache_Calendars.RemoveAll
End Sub