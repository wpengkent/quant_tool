Option Explicit

Public Function GetIRStartDate(str_curve As String, str_type As String, str_Term As String, lng_ExDate As Long, _
    lng_BuildDate As Long, int_SettleDays As Integer, cal_Calendar As Calendar, str_BDC As String, str_Currency As String, _
    Optional dic_StaticInfo As Dictionary = Nothing) As Long
    Dim lng_Output As Long

    If dic_StaticInfo Is Nothing Then Set dic_StaticInfo = GetStaticInfo()

    ' Determine how to calculate start date based on rate type
    Select Case str_type
        Case "FXFWDPT"
            Select Case UCase(str_Term)
                Case "1D": lng_Output = lng_BuildDate
                Case "2D": lng_Output = cyGetFXTomDate(str_Currency, lng_BuildDate, dic_StaticInfo)
                Case Else: lng_Output = cyGetFXSpotDate(str_Currency, lng_BuildDate, dic_StaticInfo)
            End Select
        Case "IRBILL"
            Select Case UCase(str_Term)
                Case "1D": lng_Output = lng_BuildDate
                Case "2D": lng_Output = date_workday(lng_BuildDate, 1, cal_Calendar.HolDates, cal_Calendar.Weekends)
                Case Else: lng_Output = date_workday(lng_BuildDate, int_SettleDays, cal_Calendar.HolDates, cal_Calendar.Weekends)
            End Select
        Case "IRFUTB"
            lng_Output = lng_ExDate
        Case "YTM_LONG", "YTM_SHORT", "YTM_LONG", "ZERO", "IRSWAP", "FXFWD", "BASIS_SCCY", "IRBSWP", "YTM_DISC"
            lng_Output = date_workday(lng_BuildDate, int_SettleDays, cal_Calendar.HolDates, cal_Calendar.Weekends)
    End Select

    If str_Term <> "1D" Then lng_Output = Date_ApplyBDC(lng_Output, "FOLL", cal_Calendar.HolDates, cal_Calendar.Weekends)

    GetIRStartDate = lng_Output
End Function

Public Function GetIRMatDate(str_curve As String, str_type As String, str_Term As String, lng_ExDate As Long, lng_BuildDate As Long, _
    lng_startdate As Long, cal_MatHols As Calendar, str_BDC As String, str_FlowsFreq As String, bln_AdjMatForEst As Boolean, _
    bln_EOMRule As Boolean, str_Currency As String, Optional str_FutRule As String = "", Optional dic_StaticInfo As Dictionary = Nothing) As Long

    Dim lng_Output As Long, lng_TermDate As Long, lng_LastEstDate As Long
    If dic_StaticInfo Is Nothing Then Set dic_StaticInfo = GetStaticInfo()
    Dim mas_Mapping As MappingRules: Set mas_Mapping = dic_StaticInfo(StaticInfoType.MappingRules)
    Dim int_SpotDays As Integer: int_SpotDays = mas_Mapping.Lookup_CCYSpotDays(str_Currency)

    ' For determining maturity date of instruments with coupons
    Dim str_FixedLegFreq As String, str_FloatingLegFreq As String
    Dim int_NumPeriodsFixed As Integer, int_NumPeriodsFloating As Integer
    Dim lng_LastFloatingLegStartDate As Long, lng_LastFloatingLegEstEndDate As Long
    Dim lng_CalcEndDate As Long

    ' Determine how to calculate maturity based on rate type
    Select Case str_type
        Case "FXFWDPT"
            ' Add term to the start date (which is the settlement date), using modified following business day convention
            Select Case UCase(str_Term)
                Case "1D"
                    lng_TermDate = cyGetFXTomDate(str_Currency, lng_BuildDate, dic_StaticInfo)
                Case "2D"
                    If int_SpotDays = 1 Then
                        lng_TermDate = date_addterm(lng_startdate, "1D", 1, bln_EOMRule)
                    Else
                        lng_TermDate = cyGetFXSpotDate(str_Currency, lng_BuildDate, dic_StaticInfo)
                    End If
                Case Else
                    ' Add term to the start date (which is the settlement date)
                    lng_TermDate = date_addterm(lng_startdate, str_Term, 1, bln_EOMRule)
            End Select

            lng_Output = Date_ApplyBDC(lng_TermDate, str_BDC, cal_MatHols.HolDates, cal_MatHols.Weekends)
        Case "IRBILL"
            ' Add term to the start date (which is the settlement date), using modified following business day convention
            Select Case UCase(str_Term)
                Case "1D", "2D"
                    lng_TermDate = date_workday(lng_startdate, 1)
                Case Else
                    ' Add term to the start date (which is the settlement date)
                    lng_TermDate = date_addterm(lng_startdate, str_Term, 1, bln_EOMRule)
            End Select

            lng_Output = Date_ApplyBDC(lng_TermDate, str_BDC, cal_MatHols.HolDates, cal_MatHols.Weekends)
        Case "IRFUTB"
            ' Maturity of the underlying is exactly 3 months after the start date
            If str_FutRule = "NII" Then
                Select Case Left(UCase(str_curve), 3)
                    Case "AUD"
                        lng_Output = Date_NextFutMat(lng_startdate, "3M_SECONDFRI_FOLL", 1, cal_MatHols)
                    Case "NZD"
                            lng_Output = Date_NextFutMat(lng_startdate, "3M_THUPOST10_FOLL", 1, cal_MatHols)
                    Case "MYR", "USD", "EUR", "GBP"
                        lng_Output = Date_NextFutMat(lng_startdate, "3M_THIRDWED_FOLL", 1, cal_MatHols)
                End Select
            Else
                lng_Output = date_addterm(lng_startdate, "3M", 1, True)
            End If
        Case "IRBSWP", "BASIS_SCCY"
            ' Add term to start date, using modified following business day convention
            lng_TermDate = date_addterm(lng_startdate, str_Term, 1, bln_EOMRule)
            lng_Output = Date_ApplyBDC(lng_TermDate, str_BDC, cal_MatHols.HolDates, cal_MatHols.Weekends)
        Case "YTM_SHORT", "ZERO", "YTM_DISC"
            ' Add term to start date, using modified following business day convention
            lng_TermDate = date_addterm(lng_startdate, str_Term, 1, bln_EOMRule)
            lng_Output = Date_ApplyBDC(lng_TermDate, str_BDC, cal_MatHols.HolDates, cal_MatHols.Weekends)
        Case "YTM_LONG", "IRSWAP"
            ' Add term to start date, using modified following business day convention
            lng_TermDate = date_addterm(lng_startdate, str_Term, 1, bln_EOMRule)
            lng_CalcEndDate = Date_ApplyBDC(lng_TermDate, str_BDC, cal_MatHols.HolDates, cal_MatHols.Weekends)

            If bln_AdjMatForEst = True Then
                str_FixedLegFreq = ReadSetting_LegA(str_FlowsFreq)
                str_FloatingLegFreq = ReadSetting_LegB(str_FlowsFreq)
                int_NumPeriodsFixed = Calc_NumPeriods(str_Term, str_FixedLegFreq)
                int_NumPeriodsFloating = Calc_NumPeriods(str_Term, str_FloatingLegFreq)

                lng_LastFloatingLegStartDate = Date_NextCoupon(lng_startdate, str_FloatingLegFreq, cal_MatHols, int_NumPeriodsFloating - 1, bln_EOMRule, str_BDC)
                lng_LastFloatingLegEstEndDate = Date_NextCoupon(lng_LastFloatingLegStartDate, str_FloatingLegFreq, cal_MatHols, 1, bln_EOMRule, str_BDC)

                lng_Output = WorksheetFunction.Max(lng_CalcEndDate, lng_LastFloatingLegEstEndDate)
            Else
                lng_Output = lng_CalcEndDate
            End If
    End Select

    GetIRMatDate = lng_Output
End Function