Option Explicit


' ## MEMBER DATA
Private str_name As String, shi_Base As DateShifter, int_days As Integer
Private strLst_Calendars_ToUse As Collection, strLst_Calendars_Orig As Collection
Private bln_IsBusDays As Boolean, str_BDC As String, str_Algorithm As String
Private cas_Calendars As CalendarSet


' ## INITIALIZATION
Public Sub Initialize(fld_Params As DateShifterParams)
    With fld_Params
        str_name = .ShifterName
        Set shi_Base = .BaseShifter
        int_days = .DaysToShift
        bln_IsBusDays = .IsBusDays
        str_BDC = .BDC
        str_Algorithm = .Algorithm
    End With

    ' Gather calendars.  For parallel, these will be split into their components so they can be applied individually then compared
    Set cas_Calendars = GetObject_CalendarSet()
    Select Case str_Algorithm
        Case "PARALLEL"
            If fld_Params.Calendar <> "-" Then
                Set strLst_Calendars_ToUse = Convert_SplitToList(fld_Params.Calendar, "_")
                Set strLst_Calendars_Orig = Convert_SplitToList(fld_Params.Calendar, "_")
            Else
                Set strLst_Calendars_ToUse = New Collection
                Set strLst_Calendars_Orig = New Collection
            End If
        Case "UNION"
            If fld_Params.Calendar <> "-" Then
                Set strLst_Calendars_ToUse = New Collection
                Set strLst_Calendars_Orig = New Collection
                Call strLst_Calendars_ToUse.Add(fld_Params.Calendar)
                Call strLst_Calendars_Orig.Add(fld_Params.Calendar)
            End If
    End Select
End Sub


' ## PROPERTIES
Public Property Get BaseShifter() As DateShifter
    Set BaseShifter = shi_Base
End Property

Public Property Get IsRelShifter() As Boolean
    IsRelShifter = (Not shi_Base Is Nothing)
End Property


' ## METHODS - LOOKUP
Public Function Lookup_ShiftedDate(lng_OrigDate As Long) As Long
    Dim lngLst_ShiftedDates As New Collection
    Dim lng_Output As Long, int_ctr As Integer, cal_Active As Calendar, lng_ActiveShiftedDate As Long
    Dim lng_ShiftedByBase As Long

    If shi_Base Is Nothing Then
        lng_ShiftedByBase = lng_OrigDate
    Else
        lng_ShiftedByBase = shi_Base.Lookup_ShiftedDate(lng_OrigDate)
    End If

    Debug.Assert strLst_Calendars_ToUse.count > 0
    For int_ctr = 1 To strLst_Calendars_ToUse.count
        ' Gather the calendar
        cal_Active = cas_Calendars.Lookup_Calendar(strLst_Calendars_ToUse(int_ctr))

        ' Apply the shift to the original date
        If bln_IsBusDays = True Then
            lng_ActiveShiftedDate = date_workday(lng_ShiftedByBase, int_days, cal_Active.HolDates, cal_Active.Weekends)
        Else
            lng_ActiveShiftedDate = date_addterm(lng_ShiftedByBase, "1D", int_days, False)
            lng_ActiveShiftedDate = Date_ApplyBDC(lng_ActiveShiftedDate, str_BDC, cal_Active.HolDates, cal_Active.Weekends)
        End If

        ' Temporarily store the result
        Call lngLst_ShiftedDates.Add(lng_ActiveShiftedDate)
    Next int_ctr

    ' Obtain the final shifted date.  For union algorithm, there will only be one item in list.  For parallel there will be multiple
    If int_days >= 0 Then
        lng_Output = Examine_MaxValueInList(lngLst_ShiftedDates)
    Else
        lng_Output = Examine_MinValueInList(lngLst_ShiftedDates)
    End If

    Lookup_ShiftedDate = lng_Output
End Function


' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Sub IncludeExternalCalendar(str_External As String)
    ' ## Add the specified calendar
    Select Case str_Algorithm
        Case "PARALLEL"
            Dim strLst_New As Collection
            Set strLst_New = Convert_SplitToList(str_External, "_")
            Set strLst_Calendars_ToUse = Convert_MergeLists(strLst_Calendars_ToUse, strLst_New)
        Case "UNION"
            ' Only allowed one calendar
            Set strLst_Calendars_ToUse = New Collection
            Call strLst_Calendars_ToUse.Add(str_External)
    End Select
End Sub

Public Sub RemoveExternalCalendar()
    ' ## Revert to calendars defined in the shifter
    Set strLst_Calendars_ToUse = Gather_CopyList(strLst_Calendars_Orig)
End Sub
