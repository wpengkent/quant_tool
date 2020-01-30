Option Explicit

Public Function Date_WorkDay(lng_StartDate As Long, int_Days As Integer, Optional rng_Holidays As Range = Nothing, _
    Optional str_WeekendDays As String = "SAT_SUN") As Long
    ' ## Replication of Excel 2007 function - WORKDAY
    ' ## The WORKDAY function exists in Excel 2003 as a cell formula if Analysis Toolpak VBA is added in

    If str_WeekendDays <> "SAT_SUN" Then
        ' Custom weekends, no Excel function available, use manually defined function
        If rng_Holidays Is Nothing Then
            Date_WorkDay = Date_Workday_Manual(lng_StartDate, int_Days, str_WeekendDays)
        Else
            Dim lngArr_Hols() As Long: lngArr_Hols = Convert_RangeToLngArr(rng_Holidays)
            Date_WorkDay = Date_Workday_Manual(lng_StartDate, int_Days, str_WeekendDays, , lngArr_Hols)
        End If
    Else
        Select Case CInt(Application.Version)
            Case 11
                ' Excel 2003
                If rng_Holidays Is Nothing Then
                    Date_WorkDay = Application.Run("WorkDay", lng_StartDate, int_Days)
                Else
                    Date_WorkDay = Application.Run("WorkDay", lng_StartDate, int_Days, rng_Holidays)
                End If
            Case Is >= 12
                ' Excel 2007
                If rng_Holidays Is Nothing Then
                    Date_WorkDay = WorksheetFunction.WorkDay(lng_StartDate, int_Days)
                Else
                    Date_WorkDay = WorksheetFunction.WorkDay(lng_StartDate, int_Days, rng_Holidays)
                End If
        End Select
    End If
End Function

Public Function Date_Workday_Manual(lng_RefDate As Long, int_DaysToShift As Integer, str_WeekendsType As String, _
    Optional intArr_AnnualHols As Variant = Nothing, Optional dteArr_SpecialHols As Variant = Nothing) As Long
    ' ## Slower but more flexible WorkDay function, allowing for user defined weekends and annual holidays

    Dim int_Ctr As Integer
    Dim lng_Output As Long: lng_Output = lng_RefDate

    For int_Ctr = 1 To Abs(int_DaysToShift)
        Select Case int_DaysToShift
            Case Is > 0: lng_Output = Date_WorkDay_Single(lng_Output, 1, str_WeekendsType, intArr_AnnualHols, dteArr_SpecialHols)
            Case Is < 0: lng_Output = Date_WorkDay_Single(lng_Output, -1, str_WeekendsType, intArr_AnnualHols, dteArr_SpecialHols)
        End Select
    Next int_Ctr

    Date_Workday_Manual = lng_Output
End Function

Private Function Date_WorkDay_Single(lng_RefDate As Long, int_Direction As Integer, str_WeekendsType As String, _
    Optional ByRef intArr_AnnualHols As Variant = Nothing, Optional ByRef dteArr_SpecialHols As Variant = Nothing) As Long
    ' ## Used by Date_Workday_Manual

    Dim lng_Output As Long, int_Ctr As Integer
    Dim enu_WeekStart As VbDayOfWeek, int_WeekendLength As Integer
    Select Case str_WeekendsType
        ' Week start is the first day after the weekend block.  Weekends are assumed to be in a block
        Case "FRI"
            enu_WeekStart = vbSaturday
            int_WeekendLength = 1
        Case "FRI_SAT"
            enu_WeekStart = vbSunday
            int_WeekendLength = 2
        Case "SAT_SUN"
            enu_WeekStart = vbMonday
            int_WeekendLength = 2
        Case "FRI_SAT_SUN"
            enu_WeekStart = vbMonday
            int_WeekendLength = 3
        Case Else
            Debug.Print "## ERROR - Unsupported weekend type"
            Debug.Assert False
    End Select

    Dim int_DayOfWeek As Integer: int_DayOfWeek = Weekday(lng_RefDate, enu_WeekStart)
    Select Case int_Direction
        Case 1
            ' Shift forward
            Select Case int_DayOfWeek
                Case 7: lng_Output = lng_RefDate + 1
                Case Is < (7 - int_WeekendLength): lng_Output = lng_RefDate + 1
                Case Else: lng_Output = lng_RefDate + 8 - int_DayOfWeek  ' In a standard week, this will catch Fri and Sat
            End Select
        Case -1
            ' Shift backward
            Select Case int_DayOfWeek
                Case 1: lng_Output = lng_RefDate - (int_WeekendLength + 1)
                Case Is > (8 - int_WeekendLength): lng_Output = lng_RefDate - 1  ' In a standard week, this will catch Sun and Mon
                Case Else: lng_Output = lng_RefDate - 1
            End Select
    End Select

    If IsArray(intArr_AnnualHols) = True Then
        For int_Ctr = LBound(intArr_AnnualHols) To UBound(intArr_AnnualHols)
            If intArr_AnnualHols(int_Ctr, 1) = Month(lng_Output) And intArr_AnnualHols(int_Ctr, 2) = Day(lng_Output) Then
                ' Go forward another day if found Long is an annual holiday (format: [month, day])
                lng_Output = Date_WorkDay_Single(lng_Output, int_Direction, str_WeekendsType, intArr_AnnualHols, dteArr_SpecialHols)
                Exit For
            End If
        Next int_Ctr
    End If

    If IsArray(dteArr_SpecialHols) = True Then
        For int_Ctr = LBound(dteArr_SpecialHols) To UBound(dteArr_SpecialHols)
            If dteArr_SpecialHols(int_Ctr) = lng_Output Then
                ' Go forward another day if found Long is a special holiday
                lng_Output = Date_WorkDay_Single(lng_Output, int_Direction, str_WeekendsType, intArr_AnnualHols, dteArr_SpecialHols)
            End If
        Next int_Ctr
    End If

    Date_WorkDay_Single = lng_Output
End Function

Public Function Date_EOMonth(lng_Date As Long, int_MonthShift As Integer) As Long
    ' ## Replication of Excel 2007 function - EOMONTH
    ' ## Month 13 is treated by Excel as first month of the next year
    Date_EOMonth = DateSerial(Year(lng_Date), Month(lng_Date) + int_MonthShift + 1, 1) - 1
End Function

Public Function Date_PrevDayOfWeek(lng_OrigDate As Long, int_Day As Integer) As Long
    ' ## Returns previous Monday, Tuesday, etc
    ' ## Enter 1 - 7 (Mon - Sun) as int_Day to specify desired day of the week

    Dim int_OrigDayOfWeek As Integer: int_OrigDayOfWeek = Weekday(lng_OrigDate, vbMonday)
    Dim int_Shift As Integer: int_Shift = (int_OrigDayOfWeek + (7 - int_Day)) Mod 7
    Date_PrevDayOfWeek = lng_OrigDate - int_Shift
End Function

Public Function Date_FirstBDOfMonth(lng_OrigDate As Long, Optional rng_Hols As Range = Nothing, _
    Optional str_WeekendDays As String = "SAT_SUN") As Long

    ' ## Return first business day of the month that the specified date is within
    Dim lng_StartMonth As Long: lng_StartMonth = lng_OrigDate - Day(lng_OrigDate) + 1
    Date_FirstBDOfMonth = Date_WorkDay(lng_StartMonth - 1, 1, rng_Hols, str_WeekendDays)
End Function

Public Function Date_LastBDOfMonth(lng_OrigDate As Long, Optional rng_Hols As Range = Nothing, _
    Optional str_WeekendDays As String = "SAT_SUN") As Long
    ' ## Return last business day of the month that the specified date is within
    Dim lng_EndMonth As Long: lng_EndMonth = Date_EOMonth(lng_OrigDate, 0)

    ' If not a business day, go to previous business day
    Date_LastBDOfMonth = Date_WorkDay(lng_EndMonth + 1, -1, rng_Hols, str_WeekendDays)
End Function

Public Function Date_AddTerm(lng_OrigDate As Long, str_Term As String, int_Multiples As Integer, _
    Optional bln_CheckEOM As Boolean = False) As Long
    ' ## Returns a date which has added a descriptive term such as "1D", "1W", "1M", "1Y" to the original date
    ' ## The multiples parameter specifies how many of these terms to add
    ' ## E.g. if term = "3M" and multiples = -2 then the function will return a date 6M prior to the original date
    ' ## If EOM check is turned on, then adding months to an end of month date will return another end of month date
    ' ## E.g. if EOM check is turned on, 3M from 28 Feb will be 31 May.  If it is turned off, it will be 28 May

    Dim str_TermType As String
    Dim int_TermQty As Long
    Select Case UCase(str_Term)
        Case "O/N"
            str_TermType = "D"
            int_TermQty = int_Multiples
        Case Else
            str_TermType = Examine_TermType(str_Term)
            int_TermQty = CLng(Examine_TermQty(str_Term)) * CLng(int_Multiples)

    End Select

    Dim str_MSTermType As String
    Dim int_NextDay As Integer: int_NextDay = Day(DateAdd("d", 1, lng_OrigDate))
    Dim lng_Output As Long

    Select Case UCase(str_TermType)
        Case "D": str_MSTermType = "d"
        Case "W": str_MSTermType = "ww"
        Case "M": str_MSTermType = "m"
        Case "Y": str_MSTermType = "yyyy"
    End Select

    lng_Output = DateAdd(str_MSTermType, int_TermQty, lng_OrigDate)

    ' Return EOM if the original date was EOM and this feature is requested
    If bln_CheckEOM = True And int_NextDay = 1 And (str_TermType = "M" Or str_TermType = "Y") Then
        lng_Output = Date_EOMonth(lng_Output, 0)
    End If

    Date_AddTerm = lng_Output
End Function

Public Function Date_UseDay(lng_RawDate As Long, int_DayOfMonth As Integer) As Long
    ' ## Corrects raw date, changing it to the specified day of the month if it exists in that month, otherwise the end of month is used
    Dim lng_Output As Long
    Dim lng_EOM As Long: lng_EOM = WorksheetFunction.EoMonth(lng_RawDate, 0)

    If int_DayOfMonth <= Day(lng_EOM) Then
        lng_Output = lng_RawDate + int_DayOfMonth - Day(lng_RawDate)
    Else
        lng_Output = lng_EOM
    End If

    Date_UseDay = lng_Output
End Function

Public Function Date_ApplyBDC(lng_RawDate As Long, str_BDC As String, Optional rng_Hols As Range = Nothing, _
    Optional str_Weekends As String = "SAT_SUN") As Long
    ' ## Apply specified business day convention to the raw date, optionally taking into account a holiday calendar

    Dim lng_Output As Long
    Select Case UCase(str_BDC)
        Case "FOLL"
            lng_Output = Date_WorkDay(lng_RawDate - 1, 1, rng_Hols, str_Weekends)
        Case "MOD FOLL"
            lng_Output = Date_WorkDay(lng_RawDate - 1, 1, rng_Hols, str_Weekends)
            If Month(lng_Output) <> Month(lng_RawDate) Then lng_Output = Date_LastBDOfMonth(Date_EOMonth(lng_Output, -1), rng_Hols, str_Weekends)
        Case "UNADJ", ""
            lng_Output = lng_RawDate
    End Select

    Date_ApplyBDC = lng_Output
End Function

Public Function Date_NextCoupon_FromCell(lng_PrevCoupon As Long, str_Term As String, rng_HolDates As Range, str_Weekends As String, _
    int_Multiples As Integer, Optional bln_CheckEOM As Boolean = False, Optional str_BDC As String = "MOD FOLL") As Long
    ' ## Add requested number of multiples of the term to previous coupon date, using specified business day convention

    Dim lng_TermDate As Long: lng_TermDate = Date_AddTerm(lng_PrevCoupon, str_Term, int_Multiples, bln_CheckEOM)
    Dim lng_Output As Long: lng_Output = Date_ApplyBDC(lng_TermDate, str_BDC, rng_HolDates, str_Weekends)
    Date_NextCoupon_FromCell = lng_Output
End Function

Public Function Date_NextCoupon(lng_PrevCoupon As Long, str_Term As String, cal_Calendar As Calendar, _
    int_Multiples As Integer, Optional bln_CheckEOM As Boolean = False, Optional str_BDC As String = "MOD FOLL") As Long
    ' ## Add requested number of multiples of the term to previous coupon date, using specified business day convention

    Dim lng_TermDate As Long: lng_TermDate = Date_AddTerm(lng_PrevCoupon, str_Term, int_Multiples, bln_CheckEOM)
    Dim lng_Output As Long: lng_Output = Date_ApplyBDC(lng_TermDate, str_BDC, cal_Calendar.HolDates, cal_Calendar.Weekends)
    Date_NextCoupon = lng_Output
End Function

Public Function Date_CouponSchedule(lng_StartDate As Long, str_TermToMat As String, cal_Calendar As Calendar, str_CouponFreq As String, _
    Optional str_BDC As String = "MOD FOLL", Optional bln_CheckEOM As Boolean = False) As Long()
    ' ## Provides array of calculation end dates for a swap or bond, based on the specified maturity, frequency and conventions

    Dim int_NumSwapCoupons As Integer: int_NumSwapCoupons = Calc_NumPeriods(str_TermToMat, str_CouponFreq)
    Dim lngArr_Output() As Long: ReDim lngArr_Output(1 To int_NumSwapCoupons) As Long
    Dim int_Ctr As Integer

    For int_Ctr = 1 To int_NumSwapCoupons
        lngArr_Output(int_Ctr) = Date_NextCoupon(lng_StartDate, str_CouponFreq, cal_Calendar, int_Ctr, bln_CheckEOM, str_BDC)
    Next int_Ctr

    Date_CouponSchedule = lngArr_Output
End Function

Public Function Date_NextFutMat(lng_RefDate As Long, str_Generator As String, int_Multiples As Integer, cal_Calendar As Calendar) As Long
    ' ## Get expiry date of the futures contract, based on the specified generator
    ' ## The generator is a custom name specific to the tool.  Additional generators can be set up here

    Dim lng_ShiftedRefDate As Long, lng_CurrentMonthFutDate As Long, lng_Output As Long
    Dim int_ToNextFutMonth As Integer
    Dim int_TotalMonthsToShift As Integer: int_TotalMonthsToShift = 0

    Select Case UCase(str_Generator)
        Case "3M_THIRDWED_FOLL"
            ' Contracts in Mar, Jun, Sep, Dec.  Expiry date is the third Wednesday of the delivery month, next business day if this is a holiday

            int_ToNextFutMonth = 2 - (Month(lng_RefDate) - 1) Mod 3  ' Number of months until next futures month
            If int_ToNextFutMonth = 0 Then
                ' If on or past the maturity for the current month, go to the next contract
                lng_CurrentMonthFutDate = Date_PrevDayOfWeek(lng_RefDate - Day(lng_RefDate) + 21, 3)
                lng_CurrentMonthFutDate = Date_ApplyBDC(lng_CurrentMonthFutDate, "FOLL", cal_Calendar.HolDates, cal_Calendar.Weekends)
                If lng_RefDate >= lng_CurrentMonthFutDate Then int_TotalMonthsToShift = 3
            End If

            int_TotalMonthsToShift = int_TotalMonthsToShift + int_ToNextFutMonth + 3 * (int_Multiples - 1)
            lng_ShiftedRefDate = DateAdd("m", int_TotalMonthsToShift, lng_RefDate)
            lng_Output = Date_PrevDayOfWeek(lng_ShiftedRefDate - Day(lng_ShiftedRefDate) + 21, 3)
            lng_Output = Date_ApplyBDC(lng_Output, "FOLL", cal_Calendar.HolDates, cal_Calendar.Weekends)
        Case "3M_SECONDFRI_FOLL"
            ' Contracts in Mar, Jun, Sep, Dec.  Expiry date is the second Friday of the delivery month, next business day if this is a holiday

            int_ToNextFutMonth = 2 - (Month(lng_RefDate) - 1) Mod 3  ' Number of months until next futures month
            If int_ToNextFutMonth = 0 Then
                ' If on or past the maturity for the current month, go to the next contract
                lng_CurrentMonthFutDate = Date_PrevDayOfWeek(lng_RefDate - Day(lng_RefDate) + 14, 5)
                lng_CurrentMonthFutDate = Date_ApplyBDC(lng_CurrentMonthFutDate, "FOLL", cal_Calendar.HolDates, cal_Calendar.Weekends)
                If lng_RefDate >= lng_CurrentMonthFutDate Then int_TotalMonthsToShift = 3
            End If

            int_TotalMonthsToShift = int_TotalMonthsToShift + int_ToNextFutMonth + 3 * (int_Multiples - 1)
            lng_ShiftedRefDate = DateAdd("m", int_TotalMonthsToShift, lng_RefDate)
            lng_Output = Date_PrevDayOfWeek(lng_ShiftedRefDate - Day(lng_ShiftedRefDate) + 14, 5)
            lng_Output = Date_ApplyBDC(lng_Output, "FOLL", cal_Calendar.HolDates, cal_Calendar.Weekends)
        Case "3M_THUPOST10_FOLL"
            ' Contracts in Mar, Jun, Sep, Dec.  Expiry date is first Thursday after 10th day of the month, next business day if this is a holiday

            int_ToNextFutMonth = 2 - (Month(lng_RefDate) - 1) Mod 3  ' Number of months until next futures month
            If int_ToNextFutMonth = 0 Then
                ' If on or past the maturity for the current month, go to the next contract
                lng_CurrentMonthFutDate = Date_PrevDayOfWeek(lng_RefDate - Day(lng_RefDate) + 17, 4)
                lng_CurrentMonthFutDate = Date_ApplyBDC(lng_CurrentMonthFutDate, "FOLL", cal_Calendar.HolDates, cal_Calendar.Weekends)
                If lng_RefDate >= lng_CurrentMonthFutDate Then int_TotalMonthsToShift = 3
            End If

            int_TotalMonthsToShift = int_TotalMonthsToShift + int_ToNextFutMonth + 3 * (int_Multiples - 1)
            lng_ShiftedRefDate = DateAdd("m", int_TotalMonthsToShift, lng_RefDate)
            lng_Output = Date_PrevDayOfWeek(lng_ShiftedRefDate - Day(lng_ShiftedRefDate) + 17, 4)
            lng_Output = Date_ApplyBDC(lng_Output, "FOLL", cal_Calendar.HolDates, cal_Calendar.Weekends)
    End Select

    Date_NextFutMat = lng_Output
End Function

Public Function Date_Union(ParamArray arr_Ranges() As Variant) As Long()
    ' ## Returns unsorted union of ranges containing dates

    Dim lngArr_Output() As Long
    Dim dic_Union As New Dictionary
    Dim rng_Active As Variant
    Dim int_RowCtr As Integer
    Dim lng_ActiveDate As Long

    ' Create list of unique dates
    For Each rng_Active In arr_Ranges
        For int_RowCtr = 1 To rng_Active.Rows.Count
            lng_ActiveDate = CLng(rng_Active(int_RowCtr, 1).Value)
            If dic_Union.Exists(lng_ActiveDate) = False Then Call dic_Union.Add(lng_ActiveDate, 0)
        Next int_RowCtr
    Next rng_Active

    ' Output unique dates to array
    Dim int_NumRows As Integer: int_NumRows = dic_Union.Count
    ReDim lngArr_Output(1 To int_NumRows, 1 To 1) As Long
    int_RowCtr = 1
    Dim var_Key As Variant
    For Each var_Key In dic_Union.Keys
        lngArr_Output(int_RowCtr, 1) = var_Key
        int_RowCtr = int_RowCtr + 1
    Next var_Key

    Date_Union = lngArr_Output
End Function