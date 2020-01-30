Option Explicit


' ## STATIC OBJECTS
Public Function GetObject_ConfigSheet() As ConfigSheet
    ' ## Return object holding main configuration
    Dim cfg_Output As New ConfigSheet
    Call cfg_Output.Initialize(GetSheet_Config())
    Set GetObject_ConfigSheet = cfg_Output
End Function

Public Function GetObject_ScenContainer(str_ContainerName As String, Optional dic_CurveSet As Dictionary = Nothing, _
    Optional sct_Default As ScenContainer = Nothing, Optional dic_GlobalStaticInfo As Dictionary = Nothing) As ScenContainer
    ' ## Returns the specified container, but if this is not found, returns the specified default container
    Dim sct_Output As New ScenContainer
    Dim str_SheetName As String: str_SheetName = "CONT_" & str_ContainerName

    If Examine_WorksheetExists(ThisWorkbook, str_SheetName) = True Then
        Call sct_Output.Initialize(ThisWorkbook.Worksheets("CONT_" & str_ContainerName), dic_CurveSet, dic_GlobalStaticInfo)
    Else
        Set sct_Output = sct_Default
    End If

    Set GetObject_ScenContainer = sct_Output
End Function

Public Function GetObject_RiskEngine() As RiskEngine
    ' ## Return object used to run scenarios on a target trade set
    Dim eng_Output As New RiskEngine
    Call eng_Output.Initialize
    Set GetObject_RiskEngine = eng_Output
End Function

Public Function GetObject_CalendarSet() As CalendarSet
    ' ## Return object holding all holiday calendar and weekend days
    Dim cas_Output As New CalendarSet
    Call cas_Output.Initialize(GetSheet_Holidays())
    Set GetObject_CalendarSet = cas_Output
End Function

Public Function GetObject_IRGeneratorSet() As IRGeneratorSet
    ' ## Return object holding all IR generator configuration
    Dim igs_Output As New IRGeneratorSet
    Call igs_Output.Initialize(GetSheet_Generators_IR())
    Set GetObject_IRGeneratorSet = igs_Output
End Function

Public Function GetObject_YieldGeneratorSet() As YieldGeneratorSet
    ' ## Return object holding all Yield generator configuration
    Dim igs_Output As New YieldGeneratorSet
    Call igs_Output.Initialize(GetSheet_Generators_Yield())
    Set GetObject_YieldGeneratorSet = igs_Output
End Function

Public Function GetObject_DateShifterSet() As DateShifterSet
    ' ## Return object holding all date shifter definitions
    Dim dss_Output As New DateShifterSet
    Call dss_Output.Initialize(GetSheet_DateShifters())
    Set GetObject_DateShifterSet = dss_Output
End Function

Public Function GetObject_IRQuerySet() As IRQuerySet
    ' ## Return object holding all IR database query definitions
    Dim iqs_Output As New IRQuerySet
    Call iqs_Output.Initialize(GetSheet_IRQueries(), GetObject_CalendarSet(), GetObject_MappingRules())
    Set GetObject_IRQuerySet = iqs_Output
End Function

Public Function GetObject_MappingRules() As MappingRules
    ' ## Return object containing the set of all mapping rules
    Dim map_Output As New MappingRules
    Call map_Output.Initialize(GetSheet_Mapping())
    Set GetObject_MappingRules = map_Output
End Function

Public Function GetObject_InstCache(enu_InstType As InstType, bln_StoreInst As Boolean, Optional dic_CurveSetInput As Dictionary = Nothing, _
    Optional dic_StaticInfo As Dictionary = Nothing) As InstrumentCache
    Dim ica_Output As New InstrumentCache
    Call ica_Output.Initialize(enu_InstType, bln_StoreInst, dic_CurveSetInput, dic_StaticInfo)
    Set GetObject_InstCache = ica_Output
End Function


' ## CURVE OBJECTS
Public Function GetObject_EQSpots(bln_DataExists As Boolean, Optional dic_GlobalStaticInfo As Dictionary = Nothing) As Data_EQSpots
    Dim eqs_All As Data_EQSpots: Set eqs_All = New Data_EQSpots
    Call eqs_All.Initialize(ThisWorkbook.Worksheets("EQSPT"), bln_DataExists, dic_GlobalStaticInfo)
    Set GetObject_EQSpots = eqs_All
End Function
Public Function GetObject_EQVols(bln_DataExists As Boolean, Optional dic_GlobalStaticInfo As Dictionary = Nothing) As Data_EQVols
    Dim eqv_All As Data_EQVols: Set eqv_All = New Data_EQVols
    Call eqv_All.Initialize(ThisWorkbook.Worksheets("EQVOL"), bln_DataExists, dic_GlobalStaticInfo)
    Set GetObject_EQVols = eqv_All
End Function

Public Function GetObject_FXSpots(bln_DataExists As Boolean, Optional dic_GlobalStaticInfo As Dictionary = Nothing) As Data_FXSpots
    Dim fxs_All As Data_FXSpots: Set fxs_All = New Data_FXSpots
    Call fxs_All.Initialize(ThisWorkbook.Worksheets("FXSPT"), bln_DataExists, dic_GlobalStaticInfo)
    Set GetObject_FXSpots = fxs_All
End Function

Public Function GetObject_IRCurve(str_curve As String, bln_DataExists As Boolean, bln_AddIfMissing As Boolean, _
    Optional dic_GlobalStaticInfo As Dictionary = Nothing) As Data_IRCurve

    On Error GoTo errHandler
    Dim irc_Output As Data_IRCurve: Set irc_Output = New Data_IRCurve
    Dim wks_Location As Worksheet: Set wks_Location = ThisWorkbook.Worksheets("IRC_" & str_curve)
    Call irc_Output.Initialize(wks_Location, bln_DataExists, dic_GlobalStaticInfo)
    Set GetObject_IRCurve = irc_Output

errHandler:
    Select Case Err
        Case 9  ' Worksheet does not exist
            If bln_AddIfMissing = True Then
                ThisWorkbook.Worksheets("IRC").Copy After:=ThisWorkbook.Worksheets("IRC")
                Set wks_Location = ThisWorkbook.Worksheets("IRC (2)")
                wks_Location.Name = "IRC_" & str_curve
                wks_Location.Visible = xlSheetVisible
                Resume Next
            End If
    End Select
End Function

Public Function GetObject_FXVols(str_Code As String, bln_DataExists As Boolean, bln_AddIfMissing As Boolean, _
    Optional dic_GlobalStaticInfo As Dictionary = Nothing) As Data_FXVols

    On Error GoTo errHandler
    Dim fxv_Output As Data_FXVols: Set fxv_Output = New Data_FXVols
    Dim wks_Location As Worksheet: Set wks_Location = ThisWorkbook.Worksheets("FXV_" & str_Code)
    Call fxv_Output.Initialize(wks_Location, bln_DataExists, dic_GlobalStaticInfo)
    Set GetObject_FXVols = fxv_Output

errHandler:
    Select Case Err
        Case 9  ' Worksheet does not exist
            If bln_AddIfMissing = True Then
                ThisWorkbook.Worksheets("FXV").Copy After:=ThisWorkbook.Worksheets("FXV")
                Set wks_Location = ThisWorkbook.Worksheets("FXV (2)")
                wks_Location.Name = "FXV_" & str_Code
                wks_Location.Visible = xlSheetVisible
                Resume Next
            End If
    End Select
End Function

Public Function GetObject_CapVols(str_Code As String, bln_DataExists As Boolean, bln_AddIfMissing As Boolean, _
    Optional dic_GlobalStaticInfo As Dictionary = Nothing) As Data_CapVolsQJK   'QJK code 16/12/2014


    ''Public Function GetObject_CapVols(str_Code As String, bln_DataExists As Boolean, bln_AddIfMissing As Boolean, _
    Optional dic_GlobalStaticInfo As Dictionary = Nothing) As Data_CapVols

    On Error GoTo errHandler
    'Dim cvl_Output As Data_CapVols: Set cvl_Output = New Data_CapVols
    Dim cvl_Output As Data_CapVolsQJK: Set cvl_Output = New Data_CapVolsQJK 'QJK code 16/12/2014
    Dim wks_Location As Worksheet: Set wks_Location = ThisWorkbook.Worksheets("CVL_" & str_Code)
    Call cvl_Output.Initialize(wks_Location, bln_DataExists, dic_GlobalStaticInfo)
    Set GetObject_CapVols = cvl_Output

errHandler:
    Select Case Err
        Case 0
        Case 9  ' Worksheet does not exist
            If bln_AddIfMissing = True Then
                ThisWorkbook.Worksheets("CVL").Copy After:=ThisWorkbook.Worksheets("CVL")
                Set wks_Location = ThisWorkbook.Worksheets("CVL (2)")
                wks_Location.Name = "CVL_" & str_Code
                wks_Location.Visible = xlSheetVisible
                Resume Next
            End If
        Case Else
            MsgBox "Error code " & Err.Number & ": " & Err.Description
    End Select
End Function

Public Function GetObject_CapVolSurf(str_Code As String, dbl_Strike As Double, bln_DataExists As Boolean, bln_AddIfMissing As Boolean, _
    Optional dic_GlobalStaticInfo As Dictionary = Nothing) As Data_CapVolsQJK

    'Check all worksheets name to get the strike pillars
    Dim int_count As Integer
    Dim arr_allSheetName As Variant
    ReDim arr_allSheetName(Sheets.count)

    For int_count = 1 To Sheets.count
        arr_allSheetName(int_count) = Sheets(int_count).Name
    Next int_count

    'capture the relevant strike pillars
    Dim arr_strike() As Double
    Dim int_count2 As Integer: int_count2 = 0
    Dim str_tempCaption As String

    For int_count = 1 To UBound(arr_allSheetName)
        str_tempCaption = UCase(arr_allSheetName(int_count))
        If InStr(str_tempCaption, "CVL_" & str_Code) > 0 And InStr(str_tempCaption, "=") > 0 And _
        Mid(str_tempCaption, InStr(str_tempCaption, "=") + 1, Len(str_tempCaption) - InStr(str_tempCaption, "=")) <> 0 Then
            int_count2 = int_count2 + 1
            ReDim Preserve arr_strike(int_count2)
            arr_strike(int_count2) = Mid(str_tempCaption, InStr(str_tempCaption, "=") + 1, Len(str_tempCaption) - InStr(str_tempCaption, "="))
        End If
    Next int_count

    'sort strikes in ascending order
    Dim int_sort As Integer, int_sort2 As Integer
    Dim dbl_temp As Double

    For int_sort = 1 To UBound(arr_strike)
        For int_sort2 = int_sort + 1 To UBound(arr_strike)
            If arr_strike(int_sort) > arr_strike(int_sort2) Then
                dbl_temp = arr_strike(int_sort2)
                arr_strike(int_sort2) = arr_strike(int_sort)
                arr_strike(int_sort) = dbl_temp
            End If
        Next int_sort2
    Next int_sort

    'find upper pillar and lower pillar of strike
    Dim bln_isOnPillar As Boolean: bln_isOnPillar = False
    Dim int_StrikePillarCount As Integer

    For int_count = 1 To UBound(arr_strike)
        If dbl_Strike - arr_strike(int_count) = 0 Then
            bln_isOnPillar = True
            int_StrikePillarCount = int_count
            Exit For
        End If
        If dbl_Strike < arr_strike(1) Then
            bln_isOnPillar = True
            int_StrikePillarCount = 1
            Exit For
        ElseIf dbl_Strike > arr_strike(UBound(arr_strike)) Then
            bln_isOnPillar = True
            int_StrikePillarCount = UBound(arr_strike)
            Exit For
        ElseIf dbl_Strike - arr_strike(int_count) < 0 Then
            bln_isOnPillar = False
            int_StrikePillarCount = int_count
            Exit For
        End If
    Next int_count

    Dim cvl_Output As Data_CapVolsQJK: Set cvl_Output = New Data_CapVolsQJK
    Dim wks_Location As Worksheet
    On Error GoTo errHandler
    If bln_isOnPillar = True Then
    'check if off lowest and highest pillar before default: QJK added 08/11/2016
    If dbl_Strike < arr_strike(1) Then
     Set wks_Location = ThisWorkbook.Worksheets("CVL_" & str_Code & "K=" & arr_strike(1))
    ElseIf dbl_Strike > arr_strike(UBound(arr_strike)) Then
     Set wks_Location = ThisWorkbook.Worksheets("CVL_" & str_Code & "K=" & arr_strike(UBound(arr_strike)))
    Else
        Set wks_Location = ThisWorkbook.Worksheets("CVL_" & str_Code & "K=" & arr_strike(int_StrikePillarCount))
    End If
        Call cvl_Output.Initialize(wks_Location, bln_DataExists, dic_GlobalStaticInfo)
        Set GetObject_CapVolSurf = cvl_Output
    Else
        'Strike falls between pillars
        'First Strike Pillar
        Set wks_Location = ThisWorkbook.Worksheets("CVL_" & str_Code & "K=" & arr_strike(int_StrikePillarCount - 1))
        Call cvl_Output.Initialize(wks_Location, bln_DataExists, dic_GlobalStaticInfo, True, True, dbl_Strike)
        'Second Strike Pillar
        Set wks_Location = ThisWorkbook.Worksheets("CVL_" & str_Code & "K=" & arr_strike(int_StrikePillarCount))
        Call cvl_Output.Initialize(wks_Location, bln_DataExists, dic_GlobalStaticInfo, True, False, dbl_Strike)
        Set GetObject_CapVolSurf = cvl_Output
    End If

errHandler:
    Select Case Err
        Case 0
        Case 9  ' Worksheet does not exist
            If bln_AddIfMissing = True Then
                ThisWorkbook.Worksheets("CVL").Copy After:=ThisWorkbook.Worksheets("CVL")
                Set wks_Location = ThisWorkbook.Worksheets("CVL (2)")
                wks_Location.Name = "CVL_" & str_Code
                wks_Location.Visible = xlSheetVisible
                Resume Next
            End If
        Case Else
            MsgBox "Error code " & Err.Number & ": " & Err.Description
    End Select
End Function

'QJK added 04/11/2016
'get Upper and Lower surface for Upper and Lower Strike if in  Between Strikes



Public Function GetObject_CapVolSurfUpperStrike(str_Code As String, dbl_Strike As Double, bln_DataExists As Boolean, bln_AddIfMissing As Boolean, _
    Optional dic_GlobalStaticInfo As Dictionary = Nothing) As Data_CapVolsQJK

    'Check all worksheets name to get the strike pillars
    Dim int_count As Integer
    Dim arr_allSheetName As Variant
    ReDim arr_allSheetName(Sheets.count)

    For int_count = 1 To Sheets.count
        arr_allSheetName(int_count) = Sheets(int_count).Name
    Next int_count

    'capture the relevant strike pillars
    Dim arr_strike() As Double
    Dim int_count2 As Integer: int_count2 = 0
    Dim str_tempCaption As String

    For int_count = 1 To UBound(arr_allSheetName)
        str_tempCaption = UCase(arr_allSheetName(int_count))
        If InStr(str_tempCaption, "CVL_" & str_Code) > 0 And InStr(str_tempCaption, "=") > 0 And _
        Mid(str_tempCaption, InStr(str_tempCaption, "=") + 1, Len(str_tempCaption) - InStr(str_tempCaption, "=")) <> 0 Then
            int_count2 = int_count2 + 1
            ReDim Preserve arr_strike(int_count2)
            arr_strike(int_count2) = Mid(str_tempCaption, InStr(str_tempCaption, "=") + 1, Len(str_tempCaption) - InStr(str_tempCaption, "="))
        End If
    Next int_count

    'sort strikes in ascending order
    Dim int_sort As Integer, int_sort2 As Integer
    Dim dbl_temp As Double

    For int_sort = 1 To UBound(arr_strike)
        For int_sort2 = int_sort + 1 To UBound(arr_strike)
            If arr_strike(int_sort) > arr_strike(int_sort2) Then
                dbl_temp = arr_strike(int_sort2)
                arr_strike(int_sort2) = arr_strike(int_sort)
                arr_strike(int_sort) = dbl_temp
            End If
        Next int_sort2
    Next int_sort

    'find upper pillar and lower pillar of strike
    Dim bln_isOnPillar As Boolean: bln_isOnPillar = False
    Dim int_StrikePillarCount As Integer

    For int_count = 1 To UBound(arr_strike)
        If dbl_Strike - arr_strike(int_count) = 0 Then
            bln_isOnPillar = True
            int_StrikePillarCount = int_count
            Exit For
        End If
        If dbl_Strike < arr_strike(1) Then
            bln_isOnPillar = True
            int_StrikePillarCount = 1
            Exit For
        ElseIf dbl_Strike > arr_strike(UBound(arr_strike)) Then
            bln_isOnPillar = True
            int_StrikePillarCount = UBound(arr_strike)
            Exit For
        ElseIf dbl_Strike - arr_strike(int_count) < 0 Then
            bln_isOnPillar = False
            int_StrikePillarCount = int_count
            Exit For
        End If
    Next int_count

    Dim cvl_Output As Data_CapVolsQJK: Set cvl_Output = New Data_CapVolsQJK
    Dim wks_Location As Worksheet
    On Error GoTo errHandler
    If bln_isOnPillar = True Then
        Set wks_Location = ThisWorkbook.Worksheets("CVL_" & str_Code & "K=" & arr_strike(int_StrikePillarCount))
        Call cvl_Output.Initialize(wks_Location, bln_DataExists, dic_GlobalStaticInfo)
        Set GetObject_CapVolSurf = cvl_Output
    Else
        'Strike falls between pillars
        'First Strike Pillar
        Set wks_Location = ThisWorkbook.Worksheets("CVL_" & str_Code & "K=" & arr_strike(int_StrikePillarCount - 1))
        Call cvl_Output.Initialize(wks_Location, bln_DataExists, dic_GlobalStaticInfo, True, True, dbl_Strike)
        'Second Strike Pillar
        Set wks_Location = ThisWorkbook.Worksheets("CVL_" & str_Code & "K=" & arr_strike(int_StrikePillarCount))
        Call cvl_Output.Initialize(wks_Location, bln_DataExists, dic_GlobalStaticInfo, True, False, dbl_Strike)
        Set GetObject_CapVolSurf = cvl_Output
    End If

errHandler:
    Select Case Err
        Case 0
        Case 9  ' Worksheet does not exist
            If bln_AddIfMissing = True Then
                ThisWorkbook.Worksheets("CVL").Copy After:=ThisWorkbook.Worksheets("CVL")
                Set wks_Location = ThisWorkbook.Worksheets("CVL (2)")
                wks_Location.Name = "CVL_" & str_Code
                wks_Location.Visible = xlSheetVisible
                Resume Next
            End If
        Case Else
            MsgBox "Error code " & Err.Number & ": " & Err.Description
    End Select
End Function

'end of QJK added 04/11/2016
'get Upper and Lower surface for Upper and Lower Strike if in  Between Strikes




Public Function GetObject_SwptVols(str_Code As String, bln_DataExists As Boolean, bln_AddIfMissing As Boolean, _
    Optional dic_GlobalStaticInfo As Dictionary = Nothing) As Data_SwptVols

    On Error GoTo errHandler
    Dim svc_Output As New Data_SwptVols
    Dim wks_Location As Worksheet: Set wks_Location = ThisWorkbook.Worksheets("SVL_" & str_Code)
    Call svc_Output.Initialize(wks_Location, bln_DataExists, dic_GlobalStaticInfo)
    Set GetObject_SwptVols = svc_Output

errHandler:
    Select Case Err
        Case 9  ' Worksheet does not exist
            If bln_AddIfMissing = True Then
                ThisWorkbook.Worksheets("SVL").Copy After:=ThisWorkbook.Worksheets("SVL")
                Set wks_Location = ThisWorkbook.Worksheets("SVL (2)")
                wks_Location.Name = "SVL_" & str_Code
                wks_Location.Visible = xlSheetVisible
                Resume Next
            End If
    End Select
End Function

'## EQSmile Getobject addition
Public Function GetObject_EQSmile(str_Code As String, bln_DataExists As Boolean, bln_AddIfMissing As Boolean, _
    Optional dic_GlobalStaticInfo As Dictionary = Nothing) As Data_EQSmile


    On Error GoTo errHandler
    Dim eqs_Output As New Data_EQSmile
    Dim wks_Location As Worksheet: Set wks_Location = ThisWorkbook.Worksheets("EVL_" & str_Code)
    Call eqs_Output.Initialize(wks_Location, bln_DataExists, dic_GlobalStaticInfo)
    Set GetObject_EQSmile = eqs_Output

errHandler:
    Select Case Err
        Case 9  ' Worksheet does not exist
            If bln_AddIfMissing = True Then
                ThisWorkbook.Worksheets("EVL").Copy After:=ThisWorkbook.Worksheets("EVL")
                Set wks_Location = ThisWorkbook.Worksheets("EVL (2)")
                wks_Location.Name = "EVL_" & str_Code
                wks_Location.Visible = xlSheetVisible
                Resume Next
            End If
    End Select
End Function




' ## OTHER
Public Function GetObject_Calendar(str_CalendarName As String) As Calendar
    ' ## Return specified calendar, contaning holidays and weekend days
    Dim cas_Calendars As CalendarSet: Set cas_Calendars = GetObject_CalendarSet()
    GetObject_Calendar = cas_Calendars.Lookup_Calendar(str_CalendarName)
End Function

Public Function GetObject_DateShifter(str_ShifterName As String) As DateShifter
    ' ## Return specified date shifter
    Dim dss_Shifters As DateShifterSet: Set dss_Shifters = GetObject_DateShifterSet()
    Call dss_Shifters.Initialize(GetSheet_DateShifters())
    Set GetObject_DateShifter = dss_Shifters.Lookup_Shifter(str_ShifterName)
End Function