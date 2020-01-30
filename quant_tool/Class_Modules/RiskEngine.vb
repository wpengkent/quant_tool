Option Explicit

' ## MEMBER DATA
' Static assignments
Private Const bln_ShowTiming As Boolean = True
Private Const int_NumCols_TradeStatic As Integer = 5, int_NumOutputTypes As Integer = 5
Private Const int_NumCols_VaRStress As Integer = 9, int_NumCols_IRSens As Integer = 8, int_NumCols_Vega As Integer = 8
Private wks_Location As Worksheet, wks_Results_VaRStress As Worksheet, wks_Comparison As Worksheet
Private dic_Containers As Dictionary
Private rng_TopLeft_VaRScen As Range, rng_TopLeft_IRSens As Range, rng_TopLeft_Vega As Range
Private rng_TopLeft_Trades As Range, rng_TopLeft_VaROutput As Range
Private rng_TopLeft_MXValues As Range, rng_TopLeft_Comparison As Range
Private rng_TopLeft_DV01Output As Range, rng_TopLeft_DV02Output As Range, rng_TopLeft_VegaOutput As Range
Private dic_GlobalStaticInfo As Dictionary, cfg_Settings As ConfigSheet

' Dynamic variables
Private arrLst_Target As Collection, int_NumSelectedTrades As Integer


' ## INITIALIZATION
Public Sub Initialize()
    Set wks_Location = GetSheet_RiskEngine()
    Set wks_Results_VaRStress = GetSheet_Results_VaRStress()
    Set wks_Comparison = GetSheet_Comparison()
    Set dic_Containers = New Dictionary
    dic_Containers.CompareMode = CompareMethod.TextCompare
    Set dic_GlobalStaticInfo = GetStaticInfo()
    Set cfg_Settings = dic_GlobalStaticInfo(StaticInfoType.ConfigSheet)

    Set rng_TopLeft_Trades = wks_Location.Range("A3")
    Set rng_TopLeft_VaRScen = rng_TopLeft_Trades.Offset(0, 11)
    Set rng_TopLeft_IRSens = rng_TopLeft_VaRScen.Offset(0, 5)
    Set rng_TopLeft_Vega = rng_TopLeft_IRSens.Offset(0, 4)

    Set rng_TopLeft_VaROutput = wks_Results_VaRStress.Range("A2")
    Set rng_TopLeft_MXValues = GetSheet_Results_MX().Range("A2")
    Set rng_TopLeft_Comparison = wks_Comparison.Range("A2")
    Set rng_TopLeft_DV01Output = GetSheet_Results_DV01().Range("A2")
    Set rng_TopLeft_DV02Output = GetSheet_Results_DV02().Range("A2")
    Set rng_TopLeft_VegaOutput = GetSheet_Results_Vega().Range("A2")
End Sub


' ## PROPERTIES
Public Property Get TopLeft_Trades() As Range
    Set TopLeft_Trades = rng_TopLeft_Trades
End Property

Public Property Get NumTrades() As Integer
    NumTrades = Examine_NumRows(rng_TopLeft_Trades)
End Property

Public Property Get NumScen_VaRStress() As Integer
    NumScen_VaRStress = Examine_NumRows(rng_TopLeft_VaRScen)
End Property

Public Property Get NumScen_IRSens() As Integer
    NumScen_IRSens = Examine_NumRows(rng_TopLeft_IRSens)
End Property

Public Property Get SetupSheet() As Worksheet
    Set SetupSheet = wks_Location
End Property


' ## METHODS - BATCH
Public Sub Execute_VaRStress()
    ' ## Batch run - for the specified scenarios and trades, produce risk engine results and output to the result sheet
    Dim fld_AppState_Orig As ApplicationState: fld_AppState_Orig = Gather_ApplicationState(ApplicationStateType.Current)
    Dim fld_AppState_Opt As ApplicationState: fld_AppState_Opt = Gather_ApplicationState(ApplicationStateType.Optimized)
    Call Action_SetAppState(fld_AppState_Opt)
    Const str_DLContainerName As String = "DL~"

    Dim sng_Time_Start As Single: sng_Time_Start = Timer

    ' Set up curves and interdependencies between them
    Application.StatusBar = "Gathering curve set"
    Dim dic_CurveSet As Dictionary: Set dic_CurveSet = GetAllCurves(True, False, dic_GlobalStaticInfo)
    Call FillAllDependencies(dic_CurveSet)

    ' Gather original container
    Dim str_OrigContainer As String: str_OrigContainer = GetRange_ActiveContainer().Value
    Dim sct_Orig As ScenContainer: Set sct_Orig = GetObject_ScenContainer(str_OrigContainer, dic_CurveSet, , dic_GlobalStaticInfo)
    Dim sct_Active As ScenContainer

    ' Instrument cache object
    Application.StatusBar = "Reading instruments"
    '##Matt edit
    'Dim ica_Instruments As InstrumentCache: Set ica_Instruments = GetObject_InstCache(InstType.All, True, dic_CurveSet, dic_GlobalStaticInfo)
    '##Matt edit end


    ' Prepare static variables and initialize output sheet
    Dim enu_RevalType As RevalType: enu_RevalType = RevalType.PnL
    Call FillTarget(enu_RevalType)
    Dim arrLst_Scenarios As Collection: Set arrLst_Scenarios = Gather_Scen_VaRStress()
    Dim int_OrigScen As Integer: int_OrigScen = cfg_Settings.CurrentScen
    Call Action_ClearBelow(rng_TopLeft_VaROutput, int_NumCols_VaRStress)

    ' Determine dimensions of output
    Dim varArr_Output() As Variant
    ReDim varArr_Output(1 To int_NumSelectedTrades, 1 To int_NumCols_VaRStress) As Variant

    ' Prepare variables used in loop
    Dim arr_ActiveTarget() As Variant, arr_ActiveScenInfo() As Variant
    Dim int_FoundTradeIndex As Integer, int_FoundType As Integer, str_FoundSheet As String, str_FoundCell As String
    Dim str_FoundTradeID As String, str_FoundStrategy As String, str_FoundMXID As String
    Dim str_ActiveContainer As String, int_ActiveScen As Integer, int_ActiveMXScen As Integer
    Dim rng_ActiveOutput As Range: Set rng_ActiveOutput = rng_TopLeft_VaROutput.Resize(int_NumSelectedTrades, int_NumCols_VaRStress)
    Dim bln_ActiveIsDownload As Boolean, str_ActiveContSheet As String

    ' Loop through active scenarios and active trades
    Call Me.Clear_VaRStress
    Dim int_ScenCtr As Integer, int_TargetCtr As Integer
    For int_ScenCtr = 1 To arrLst_Scenarios.count
        ' Read scenario information
        arr_ActiveScenInfo = arrLst_Scenarios(int_ScenCtr)
        int_ActiveScen = arr_ActiveScenInfo(2)
        int_ActiveMXScen = arr_ActiveScenInfo(3)

        ' Determine whether to download scenario or use that from sheet
        str_ActiveContainer = arr_ActiveScenInfo(1)
        bln_ActiveIsDownload = Left(str_ActiveContainer, 3) = str_DLContainerName
        If bln_ActiveIsDownload = True Then
            str_ActiveContainer = Right(str_ActiveContainer, Len(str_ActiveContainer) - 3)
            str_ActiveContSheet = str_DLContainerName
        Else
            str_ActiveContSheet = str_ActiveContainer
        End If

        ' Find container or create and store it
        If dic_Containers.Exists(str_ActiveContSheet) Then
            Set sct_Active = dic_Containers(str_ActiveContSheet)
        Else
            Set sct_Active = GetObject_ScenContainer(str_ActiveContSheet, dic_CurveSet, sct_Orig, dic_GlobalStaticInfo)
            Call dic_Containers.Add(str_ActiveContSheet, sct_Active)
        End If

        ' Download scenario if required
        If bln_ActiveIsDownload = True Then
            Application.StatusBar = "Downloading scenario: " & int_ActiveScen
            Call sct_Active.ClearScenarios
            Call sct_Active.DownloadFromDB(str_ActiveContainer, int_ActiveScen, int_ActiveScen)
        End If

        ' Apply scenario and revalue trades
        Application.StatusBar = "Applying scenario: " & int_ActiveScen
        Call sct_Active.LoadScenario(int_ActiveScen)
        '##Matt edit 2
        Dim ica_Instruments As InstrumentCache: Set ica_Instruments = GetObject_InstCache(InstType.All, True, dic_CurveSet, dic_GlobalStaticInfo)
        '##Matt edit 2 end

        Call ica_Instruments.PerformRecalc(sct_Active.ValDate_Current)
        Call ica_Instruments.OutputValues(enu_RevalType)

        ' Process results
        For int_TargetCtr = 1 To arrLst_Target.count
            ' Read from target
            arr_ActiveTarget = arrLst_Target(int_TargetCtr)
            int_FoundTradeIndex = arr_ActiveTarget(1)
            int_FoundType = arr_ActiveTarget(2)
            str_FoundSheet = arr_ActiveTarget(3)
            str_FoundCell = arr_ActiveTarget(4)
            str_FoundTradeID = arr_ActiveTarget(5)
            str_FoundStrategy = arr_ActiveTarget(6)
            str_FoundMXID = arr_ActiveTarget(7)

            ' Store output values
            varArr_Output(int_FoundTradeIndex, 1) = str_FoundTradeID
            varArr_Output(int_FoundTradeIndex, 2) = str_FoundStrategy
            varArr_Output(int_FoundTradeIndex, 3) = str_FoundMXID
            varArr_Output(int_FoundTradeIndex, 4) = str_FoundSheet
            varArr_Output(int_FoundTradeIndex, 5) = str_ActiveContainer
            varArr_Output(int_FoundTradeIndex, 6) = int_ActiveScen
            varArr_Output(int_FoundTradeIndex, 7) = int_ActiveMXScen
            varArr_Output(int_FoundTradeIndex, 7 + int_FoundType) = ThisWorkbook.Worksheets(str_FoundSheet).Range(str_FoundCell).Value2
        Next int_TargetCtr

        ' Output values from storage
        rng_ActiveOutput.Value = varArr_Output
        Set rng_ActiveOutput = rng_ActiveOutput.Offset(int_NumSelectedTrades, 0)
    Next int_ScenCtr

    ' Reset to base scenario if a scenario was run
    If arrLst_Scenarios.count > 0 Then
        Application.StatusBar = "Resetting scenario to original"
        If bln_ActiveIsDownload = True Then Call sct_Active.ClearScenarios
        If str_OrigContainer = str_DLContainerName Then Call sct_Orig.ClearScenarios
        Call sct_Orig.LoadScenario(int_OrigScen)
        Call ica_Instruments.PerformRecalc(sct_Orig.ValDate_Current)
        Call ica_Instruments.OutputValues(enu_RevalType)
    End If

    ' Display timing
    Dim sng_Time_End As Double: sng_Time_End = Timer
    If bln_ShowTiming = True Then
        Debug.Print "Batch Generation - Time elapsed: " & Round(sng_Time_End - sng_Time_Start, 1) & " seconds"
    End If

    Call Action_SetAppState(fld_AppState_Orig)
End Sub

Public Sub Clear_VaRStress()
    ' ## Clear all results from the result sheet
    Call Action_DeleteBelow(rng_TopLeft_VaROutput, int_NumCols_VaRStress)
End Sub

Public Sub Execute_IRSens(enu_RevalType As RevalType)
    ' ## Batch run - for the specified scenarios and trades, produce risk engine results and output to the result sheet
    Dim fld_AppState_Orig As ApplicationState: fld_AppState_Orig = Gather_ApplicationState(ApplicationStateType.Current)
    Dim fld_AppState_Opt As ApplicationState: fld_AppState_Opt = Gather_ApplicationState(ApplicationStateType.Optimized)
    Call Action_SetAppState(fld_AppState_Opt)

    Dim sng_Time_Start As Single: sng_Time_Start = Timer

    ' Set up curves and interdependencies between them
    Application.StatusBar = "Gathering curve set"
    Dim dic_CurveSet As Dictionary: Set dic_CurveSet = GetAllCurves(True, False, dic_GlobalStaticInfo)
    Call FillAllDependencies(dic_CurveSet)

    ' Obtain list of curves to evaluate IR sensitivities
    Dim arrLst_Scen_IRSens As Collection: Set arrLst_Scen_IRSens = Gather_Scen_IRSens()

    ' Instrument cache object
    Application.StatusBar = "Reading instruments"
    Dim ica_Instruments As InstrumentCache: Set ica_Instruments = GetObject_InstCache(InstType.All, True, dic_CurveSet, dic_GlobalStaticInfo)

    ' Prepare static variables and initialize output sheet
    Call FillTarget(enu_RevalType)
    Dim int_LoadedScen As Integer: int_LoadedScen = cfg_Settings.CurrentScen

    ' Determine dimensions of output
    Dim varArr_Output() As Variant
    ReDim varArr_Output(1 To int_NumSelectedTrades, 1 To int_NumCols_IRSens) As Variant

    ' Prepare variables used in loop
    Dim arr_ActiveScen() As Variant, str_ActiveCurve As String, str_ActiveShiftType As String
    Dim irc_Shifted As Data_IRCurve, lng_ActivePillarDate As Long
    Dim rng_ActiveOutput As Range

    Select Case enu_RevalType
        Case RevalType.DV01
            Set rng_ActiveOutput = rng_TopLeft_DV01Output.Resize(int_NumSelectedTrades, int_NumCols_IRSens)
            Call Action_ClearBelow(rng_TopLeft_DV01Output, int_NumCols_IRSens)
        Case RevalType.DV02
            Set rng_ActiveOutput = rng_TopLeft_DV02Output.Resize(int_NumSelectedTrades, int_NumCols_IRSens)
            Call Action_ClearBelow(rng_TopLeft_DV02Output, int_NumCols_IRSens)
        Case Else: Debug.Assert False
    End Select

    ' Loop through active scenarios and active trades
    Dim int_ScenCtr As Integer, int_TargetCtr As Integer, int_PillarCtr As Integer
    Dim int_NumPillars As Integer
    For int_ScenCtr = 1 To arrLst_Scen_IRSens.count
        ' Read scenario information
        arr_ActiveScen = arrLst_Scen_IRSens(int_ScenCtr)
        str_ActiveCurve = arr_ActiveScen(1)
        str_ActiveShiftType = arr_ActiveScen(2)
        Set irc_Shifted = dic_CurveSet(CurveType.IRC)(str_ActiveCurve)

        Select Case str_ActiveShiftType
            Case "ZERO (UNIFORM)"
                Call RunActiveIRSens(ica_Instruments, enu_RevalType, irc_Shifted, 0, int_LoadedScen, rng_ActiveOutput)
                Set rng_ActiveOutput = rng_ActiveOutput.Offset(int_NumSelectedTrades, 0)
            Case "ZERO (BY PILLAR)"
                int_NumPillars = irc_Shifted.NumPoints
                For int_PillarCtr = 1 To int_NumPillars
                    Call RunActiveIRSens(ica_Instruments, enu_RevalType, irc_Shifted, int_PillarCtr, int_LoadedScen, rng_ActiveOutput)
                    Set rng_ActiveOutput = rng_ActiveOutput.Offset(int_NumSelectedTrades, 0)
                Next int_PillarCtr
            Case Else: Debug.Assert False
        End Select
    Next int_ScenCtr

    ' Display timing
    Dim sng_Time_End As Double: sng_Time_End = Timer
    If bln_ShowTiming = True Then
        Debug.Print "Batch Generation - Time elapsed: " & Round(sng_Time_End - sng_Time_Start, 1) & " seconds"
    End If

    Call Action_SetAppState(fld_AppState_Orig)
End Sub

Public Sub Execute_Vega()
    ' ## Batch run - for the specified scenarios and trades, produce risk engine results and output to the result sheet
    Dim fld_AppState_Orig As ApplicationState: fld_AppState_Orig = Gather_ApplicationState(ApplicationStateType.Current)
    Dim fld_AppState_Opt As ApplicationState: fld_AppState_Opt = Gather_ApplicationState(ApplicationStateType.Optimized)
    Call Action_SetAppState(fld_AppState_Opt)

    Dim sng_Time_Start As Single: sng_Time_Start = Timer

    ' Set up curves and interdependencies between them
    Application.StatusBar = "Gathering curve set"
    Dim dic_CurveSet As Dictionary: Set dic_CurveSet = GetAllCurves(True, False, dic_GlobalStaticInfo)
    Call FillAllDependencies(dic_CurveSet)

    ' Obtain list of curves to evaluate vega sensitivities
    Dim arrLst_Scen_Vega As Collection: Set arrLst_Scen_Vega = Gather_Scen_Vega()

    ' Instrument cache object
    Application.StatusBar = "Reading instruments"
    Dim ica_Instruments As InstrumentCache: Set ica_Instruments = GetObject_InstCache(InstType.All, True, dic_CurveSet, dic_GlobalStaticInfo)

    ' Prepare static variables and initialize output sheet
    Dim enu_RevalType As RevalType: enu_RevalType = RevalType.Vega
    Call FillTarget(enu_RevalType)
    Dim int_LoadedScen As Integer: int_LoadedScen = cfg_Settings.CurrentScen

    ' Determine dimensions of output
    Dim varArr_Output() As Variant
    ReDim varArr_Output(1 To int_NumSelectedTrades, 1 To int_NumCols_Vega) As Variant

    ' Prepare variables used in loop
    Dim arr_ActiveTarget() As Variant
    Dim int_FoundTradeIndex As Integer, int_FoundType As Integer, str_FoundSheet As String, str_FoundCell As String
    Dim str_FoundTradeID As String, str_FoundStrategy As String, str_FoundMXID As String
    Dim arr_ActiveInfo() As Variant, enu_ActiveType As CurveType, str_ActiveCurve As String
    Dim rng_ActiveOutput As Range: Set rng_ActiveOutput = rng_TopLeft_VegaOutput.Resize(int_NumSelectedTrades, int_NumCols_Vega)
    Call Action_ClearBelow(rng_TopLeft_VegaOutput, int_NumCols_Vega)

    ' Loop through active scenarios and active trades
    Dim int_ScenCtr As Integer, int_TargetCtr As Integer
    For int_ScenCtr = 1 To arrLst_Scen_Vega.count
        ' Read scenario information
        arr_ActiveInfo = arrLst_Scen_Vega(int_ScenCtr)
        enu_ActiveType = arr_ActiveInfo(1)
        str_ActiveCurve = arr_ActiveInfo(2)

        ' Apply scenario and revalue trades
        Application.StatusBar = "Curve: " & str_ActiveCurve
        Call ica_Instruments.SetVegaCurve(enu_ActiveType, str_ActiveCurve)
        Call ica_Instruments.OutputValues(enu_RevalType)

        ' Process results
        For int_TargetCtr = 1 To arrLst_Target.count
            ' Read from target
            arr_ActiveTarget = arrLst_Target(int_TargetCtr)
            int_FoundTradeIndex = arr_ActiveTarget(1)
            int_FoundType = arr_ActiveTarget(2)
            str_FoundSheet = arr_ActiveTarget(3)
            str_FoundCell = arr_ActiveTarget(4)
            str_FoundTradeID = arr_ActiveTarget(5)
            str_FoundStrategy = arr_ActiveTarget(6)
            str_FoundMXID = arr_ActiveTarget(7)

            ' Store output values
            varArr_Output(int_FoundTradeIndex, 1) = str_FoundTradeID
            varArr_Output(int_FoundTradeIndex, 2) = str_FoundStrategy
            varArr_Output(int_FoundTradeIndex, 3) = str_FoundMXID
            varArr_Output(int_FoundTradeIndex, 4) = str_FoundSheet
            varArr_Output(int_FoundTradeIndex, 5) = int_LoadedScen
            varArr_Output(int_FoundTradeIndex, 6) = GetCurveTypeName(enu_ActiveType)
            varArr_Output(int_FoundTradeIndex, 7) = str_ActiveCurve
            varArr_Output(int_FoundTradeIndex, 8) = ThisWorkbook.Worksheets(str_FoundSheet).Range(str_FoundCell).Value2
        Next int_TargetCtr

        ' Output values from storage
        rng_ActiveOutput.Value = varArr_Output
        Set rng_ActiveOutput = rng_ActiveOutput.Offset(int_NumSelectedTrades, 0)
    Next int_ScenCtr

    ' Display timing
    Dim sng_Time_End As Double: sng_Time_End = Timer
    If bln_ShowTiming = True Then
        Debug.Print "Batch Generation - Time elapsed: " & Round(sng_Time_End - sng_Time_Start, 1) & " seconds"
    End If

    Call Action_SetAppState(fld_AppState_Orig)
End Sub


' ## METHODS - TARGET TRADE SET
Public Sub ClearTrades()
    Call Action_ClearBelow(rng_TopLeft_Trades, int_NumCols_TradeStatic + int_NumOutputTypes)
End Sub

Private Sub FillTarget(enu_RevalType As RevalType)
    ' ## Return array list containing information for only the selected trades and results
    Set arrLst_Target = New Collection
    Const int_OffsetToSheet As Integer = 4
    Dim str_ActiveSheet As String, str_ActiveAddress As String, str_ActiveTradeID As String
    Dim str_ActiveStrategy As String, str_ActiveMXID As String
    Dim int_TradeCtr As Integer, int_TypeCtr As Integer  ' TypeCtr corresponds to values of the enumeration 'OutputType'
    Dim arr_ActiveInfo(1 To 7) As Variant
    Dim int_NextTradeIndex As Integer: int_NextTradeIndex = 0
    Dim intLst_ColsIncluded As New Collection

    Select Case enu_RevalType
        Case RevalType.PnL
            Call intLst_ColsIncluded.Add(1)
            Call intLst_ColsIncluded.Add(2)
        Case RevalType.DV01: Call intLst_ColsIncluded.Add(3)
        Case RevalType.DV02: Call intLst_ColsIncluded.Add(4)
        Case RevalType.Vega: Call intLst_ColsIncluded.Add(5)
    End Select

    Dim bln_ActiveTradeAdded As Boolean
    Dim var_ColCtr As Variant
    For int_TradeCtr = 1 To Me.NumTrades
        ' Store only the selected trades
        If rng_TopLeft_Trades.Offset(int_TradeCtr - 1, 3).Value = "YES" Then
            bln_ActiveTradeAdded = False
            str_ActiveTradeID = rng_TopLeft_Trades.Offset(int_TradeCtr - 1, 0).Value
            str_ActiveStrategy = rng_TopLeft_Trades.Offset(int_TradeCtr - 1, 1).Value
            str_ActiveMXID = rng_TopLeft_Trades.Offset(int_TradeCtr - 1, 2).Value
            str_ActiveSheet = rng_TopLeft_Trades.Offset(int_TradeCtr - 1, int_OffsetToSheet).Value

            ' Store entries for PnL and PnLChg. Assumes these are the first two columns after the sheet name
            ' Store only addresses which are not marked "-"
            For Each var_ColCtr In intLst_ColsIncluded
                str_ActiveAddress = rng_TopLeft_Trades.Offset(int_TradeCtr - 1, int_OffsetToSheet + var_ColCtr).Value
                If str_ActiveAddress <> "-" Then
                    ' Count the unique number of trades relevant for the output type
                    If bln_ActiveTradeAdded = False Then
                        int_NextTradeIndex = int_NextTradeIndex + 1
                        bln_ActiveTradeAdded = True
                    End If

                    arr_ActiveInfo(1) = int_NextTradeIndex
                    arr_ActiveInfo(2) = CInt(var_ColCtr)
                    arr_ActiveInfo(3) = str_ActiveSheet
                    arr_ActiveInfo(4) = str_ActiveAddress
                    arr_ActiveInfo(5) = str_ActiveTradeID
                    arr_ActiveInfo(6) = str_ActiveStrategy
                    arr_ActiveInfo(7) = str_ActiveMXID
                    Call arrLst_Target.Add(arr_ActiveInfo)
                End If
            Next var_ColCtr
        End If
    Next int_TradeCtr

    ' Store number of trades selected
    int_NumSelectedTrades = int_NextTradeIndex
End Sub


' ## METHODS - SCENARIO SET
Private Function Gather_Scen_VaRStress() As Collection
    ' ## Return list containing only the selected scenarios
    Dim arrLst_Output As New Collection
    Dim arr_ActiveInfo(1 To 3) As Variant
    Dim int_ctr As Integer
    For int_ctr = 1 To Me.NumScen_VaRStress
        If rng_TopLeft_VaRScen.Offset(int_ctr - 1, 3).Value = "YES" Then
            arr_ActiveInfo(1) = rng_TopLeft_VaRScen.Offset(int_ctr - 1, 0).Value  ' Container
            arr_ActiveInfo(2) = rng_TopLeft_VaRScen.Offset(int_ctr - 1, 1).Value  ' Scenario
            arr_ActiveInfo(3) = rng_TopLeft_VaRScen.Offset(int_ctr - 1, 2).Value  ' MX scenario
            Call arrLst_Output.Add(arr_ActiveInfo)
        End If
    Next int_ctr

    Set Gather_Scen_VaRStress = arrLst_Output
End Function

Private Function Gather_Scen_IRSens() As Collection
    ' ## Return list containing only the selected curves for DV01 computation
    Dim arrLst_Output As New Collection
    Dim arr_ActiveInfo(1 To 2) As Variant
    Dim int_ctr As Integer
    For int_ctr = 1 To Me.NumScen_IRSens
        If rng_TopLeft_IRSens.Offset(int_ctr - 1, 2).Value = "YES" Then
            arr_ActiveInfo(1) = rng_TopLeft_IRSens.Offset(int_ctr - 1, 0).Value  ' Curve name
            arr_ActiveInfo(2) = UCase(rng_TopLeft_IRSens.Offset(int_ctr - 1, 1).Value)  ' Shift type
            Call arrLst_Output.Add(arr_ActiveInfo)
        End If
    Next int_ctr

    Set Gather_Scen_IRSens = arrLst_Output
End Function

Private Function Gather_Scen_Vega() As Collection
    ' ## Return list containing the selected curves, along with their data type
    Dim arrLst_Output As New Collection

    Dim int_ctr As Integer, arr_Active(1 To 2) As Variant
    For int_ctr = 1 To Me.NumScen_IRSens
        If rng_TopLeft_Vega.Offset(int_ctr - 1, 2).Value = "YES" Then
            arr_Active(1) = GetCurveType(rng_TopLeft_Vega.Offset(int_ctr - 1, 0).Value)  ' Curve type
            arr_Active(2) = rng_TopLeft_Vega.Offset(int_ctr - 1, 1).Value  ' Curve name
            Call arrLst_Output.Add(arr_Active)
        End If
    Next int_ctr

    Set Gather_Scen_Vega = arrLst_Output
End Function


' ## METHODS - SUPPORT
Private Sub RunActiveIRSens(ica_Instruments As InstrumentCache, enu_RevalType As RevalType, irc_Curve As Data_IRCurve, _
    int_Pillar As Integer, int_LoadedScen As Integer, rng_Output_TopLeft As Range)
    ' Apply scenario and revalue trades
    Dim str_curve As String: str_curve = irc_Curve.CurveName
    Application.StatusBar = "Curve: " & str_curve
    ica_Instruments.Curve_IRSens = str_curve
    ica_Instruments.Pillar_IRSens = int_Pillar
    Call ica_Instruments.OutputValues(enu_RevalType)

    ' Determine dimensions of output
    Dim varArr_Output() As Variant
    ReDim varArr_Output(1 To int_NumSelectedTrades, 1 To int_NumCols_IRSens) As Variant

    ' Process results
    Dim arr_ActiveTarget() As Variant
    Dim int_FoundTradeIndex As Integer, int_FoundType As Integer, str_FoundSheet As String, str_FoundCell As String
    Dim str_FoundTradeID As String, str_FoundStrategy As String, str_FoundMXID As String
    Dim int_TargetCtr As Integer
    For int_TargetCtr = 1 To arrLst_Target.count
        ' Read from target
        arr_ActiveTarget = arrLst_Target(int_TargetCtr)
        int_FoundTradeIndex = arr_ActiveTarget(1)
        int_FoundType = arr_ActiveTarget(2)
        str_FoundSheet = arr_ActiveTarget(3)
        str_FoundCell = arr_ActiveTarget(4)
        str_FoundTradeID = arr_ActiveTarget(5)
        str_FoundStrategy = arr_ActiveTarget(6)
        str_FoundMXID = arr_ActiveTarget(7)

        ' Store output values
        varArr_Output(int_FoundTradeIndex, 1) = str_FoundTradeID
        varArr_Output(int_FoundTradeIndex, 2) = str_FoundStrategy
        varArr_Output(int_FoundTradeIndex, 3) = str_FoundMXID
        varArr_Output(int_FoundTradeIndex, 4) = str_FoundSheet
        varArr_Output(int_FoundTradeIndex, 5) = int_LoadedScen
        varArr_Output(int_FoundTradeIndex, 6) = str_curve
        If int_Pillar = 0 Then
            varArr_Output(int_FoundTradeIndex, 7) = "-"
        Else
            varArr_Output(int_FoundTradeIndex, 7) = irc_Curve.Lookup_MaturityFromIndex(int_Pillar)
        End If
        varArr_Output(int_FoundTradeIndex, 8) = ThisWorkbook.Worksheets(str_FoundSheet).Range(str_FoundCell).Value2
    Next int_TargetCtr

    ' Output values from storage
    rng_Output_TopLeft.Value = varArr_Output
End Sub


' ## METHODS - COMPARISON
Public Function GetDict_Comparison() As Dictionary
    ' ## Inner dictionary has MX ID as key
    ' ## Outer dictionary has MX scenario number as key
    Dim dic_output As New Dictionary, dic_ActiveScen As Dictionary
    Dim lng_NumRows As Long: lng_NumRows = Examine_NumRows(rng_TopLeft_VaROutput)
    Dim str_ActiveID_MX As String, str_ActiveScenID As String, dbl_ActiveVal As Double
    Dim varArr_REOutput() As Variant: varArr_REOutput = rng_TopLeft_VaROutput.Resize(lng_NumRows, int_NumCols_VaRStress + 2).Value2
    Dim arr_Comparison(1 To 5) As Variant, arr_Comparison_Existing() As Variant
    Dim lng_RowCtr As Long

    For lng_RowCtr = 1 To lng_NumRows
        str_ActiveID_MX = varArr_REOutput(lng_RowCtr, 3)
        str_ActiveScenID = varArr_REOutput(lng_RowCtr, 5) & varArr_REOutput(lng_RowCtr, 7)  ' Container joined with MX scenario
        arr_Comparison(1) = varArr_REOutput(lng_RowCtr, 1)  ' RE trade ID
        arr_Comparison(2) = varArr_REOutput(lng_RowCtr, 6)  ' RE scenario
        arr_Comparison(3) = varArr_REOutput(lng_RowCtr, 4)  ' Sheet name
        arr_Comparison(4) = varArr_REOutput(lng_RowCtr, 8)  ' PnL
        arr_Comparison(5) = varArr_REOutput(lng_RowCtr, 9)  ' PnLChg

        ' Ensure scenario exists in output
        If dic_output.Exists(str_ActiveScenID) = True Then
            ' Add to existing scenario
            Set dic_ActiveScen = dic_output(str_ActiveScenID)
        Else
            ' Create entry for new scenario
            Set dic_ActiveScen = New Dictionary
            Call dic_output.Add(str_ActiveScenID, dic_ActiveScen)
        End If

        ' Place result in output
        If dic_ActiveScen.Exists(str_ActiveID_MX) Then
            arr_Comparison_Existing = dic_ActiveScen(str_ActiveID_MX)

            ' Replace RE trade ID with strategy and aggregate results
            arr_Comparison(1) = varArr_REOutput(lng_RowCtr, 2)
            arr_Comparison(4) = arr_Comparison_Existing(4) + arr_Comparison(4)
            arr_Comparison(5) = arr_Comparison_Existing(5) + arr_Comparison(5)

            Call dic_ActiveScen.Remove(str_ActiveID_MX)
            Call dic_ActiveScen.Add(str_ActiveID_MX, arr_Comparison)
        Else
            Call dic_ActiveScen.Add(str_ActiveID_MX, arr_Comparison)
        End If
    Next lng_RowCtr

    Set GetDict_Comparison = dic_output
End Function

Public Function GetDict_MXValues() As Dictionary
    ' ## Purpose of storing in dictionary is to aggregate duplicate MX IDs, which can occur if a barrier option knocks in
    ' ## Inner dictionary has MX ID as key
    ' ## Middle dictionary has MX scenario number as key
    ' ## Outer dictionary has container as key
    Dim dic_output As New Dictionary, dic_ActiveCont As Dictionary, dic_ActiveScen As Dictionary
    Dim lng_NumRows As Long: lng_NumRows = Examine_NumRows(rng_TopLeft_MXValues)
    Dim str_ActiveID_MX As String, str_ActiveCont As String, lng_ActiveScenID As Long, dbl_ActiveVal As Double
    Dim varArr_MXOutput() As Variant: varArr_MXOutput = rng_TopLeft_MXValues.Resize(lng_NumRows, 5).Value2
    Dim dblArr_ActiveValues(1 To 2) As Double, dblArr_ActiveValues_Existing() As Double
    Dim lng_RowCtr As Long

    For lng_RowCtr = 1 To lng_NumRows
        ' Gather values from table
        str_ActiveID_MX = varArr_MXOutput(lng_RowCtr, 2)
        lng_ActiveScenID = varArr_MXOutput(lng_RowCtr, 3)
        str_ActiveCont = varArr_MXOutput(lng_RowCtr, 1)
        dblArr_ActiveValues(1) = varArr_MXOutput(lng_RowCtr, 4)  ' Result
        dblArr_ActiveValues(2) = varArr_MXOutput(lng_RowCtr, 5)  ' Result (diff)

        ' Catalogue by container
        If dic_output.Exists(str_ActiveCont) = True Then
            ' Add to existing scenario
            Set dic_ActiveCont = dic_output(str_ActiveCont)
        Else
            ' Create entry for new scenario
            Set dic_ActiveCont = New Dictionary
            Call dic_output.Add(str_ActiveCont, dic_ActiveCont)
        End If

        ' Catalogue by scenario
        If dic_ActiveCont.Exists(lng_ActiveScenID) = True Then
            ' Add to existing scenario
            Set dic_ActiveScen = dic_ActiveCont(lng_ActiveScenID)
        Else
            ' Create entry for new scenario
            Set dic_ActiveScen = New Dictionary
            Call dic_ActiveCont.Add(lng_ActiveScenID, dic_ActiveScen)
        End If

        ' Catalogue by MX trade ID
        If dic_ActiveScen.Exists(str_ActiveID_MX) Then
            dblArr_ActiveValues_Existing = dic_ActiveScen(str_ActiveID_MX)

            ' Add values if previous values already exist
            dblArr_ActiveValues(1) = dblArr_ActiveValues_Existing(1) + dblArr_ActiveValues(1)
            dblArr_ActiveValues(2) = dblArr_ActiveValues_Existing(2) + dblArr_ActiveValues(2)

            Call dic_ActiveScen.Remove(str_ActiveID_MX)
            Call dic_ActiveScen.Add(str_ActiveID_MX, dblArr_ActiveValues)
        Else
            Call dic_ActiveScen.Add(str_ActiveID_MX, dblArr_ActiveValues)
        End If
    Next lng_RowCtr

    Set GetDict_MXValues = dic_output
End Function

Public Sub OutputComparison()
    Dim rng_ActiveMXLine As Range: Set rng_ActiveMXLine = rng_TopLeft_MXValues.Resize(1, 4)
    Dim str_Active_RECont As String, str_Active_MXID As String, lng_ActiveScen_MX As Long
    Dim dbl_PnL_MX As Double, dbl_PnLChg_MX As Double
    Dim dic_Overall_MX As Dictionary: Set dic_Overall_MX = Me.GetDict_MXValues()
    Dim dic_ActiveCont_MX As Dictionary, dic_ActiveScen_MX As Dictionary, dblArr_ActiveMXValues() As Double
    Dim dic_Overall_RE As Dictionary: Set dic_Overall_RE = Me.GetDict_Comparison()
    Dim dic_ByScen_RE As Dictionary, arr_Line_RE() As Variant
    'Dim rng_OutputLine As Range: Set rng_OutputLine = rng_TopLeft_Comparison
    Dim lng_ContainerCtr As Long, lng_ScenCtr As Long, lng_TradeCtr As Long
    Dim lng_NotFoundCtr As Long: lng_NotFoundCtr = 0
    Dim int_ActiveCol As Integer
    Dim dic_Addresses As New Dictionary
    dic_Addresses.CompareMode = CompareMethod.TextCompare
    Dim lng_MaxResults As Long: lng_MaxResults = Examine_NumRows(rng_TopLeft_MXValues)
    Dim varArr_Results() As Variant: ReDim varArr_Results(1 To lng_MaxResults, 1 To 9) As Variant  ' Not all the spaces in the array need to be filled
    Dim lng_ActiveRowNum As Long: lng_ActiveRowNum = 0

    ' Clear existing outputs
    Call Action_ClearBelow(rng_TopLeft_Comparison, 18)

    ' Go through each stored entry from MX results and collect results.  The dictionary is catalogued by container, scenario and trade
    For lng_ContainerCtr = 1 To dic_Overall_MX.count
        str_Active_RECont = dic_Overall_MX.Keys(lng_ContainerCtr - 1)
        Set dic_ActiveCont_MX = dic_Overall_MX.Items(lng_ContainerCtr - 1)

        For lng_ScenCtr = 1 To dic_ActiveCont_MX.count
            lng_ActiveScen_MX = dic_ActiveCont_MX.Keys(lng_ScenCtr - 1)
            Set dic_ActiveScen_MX = dic_ActiveCont_MX.Items(lng_ScenCtr - 1)

            For lng_TradeCtr = 1 To dic_ActiveScen_MX.count
                lng_ActiveRowNum = lng_ActiveRowNum + 1
                str_Active_MXID = dic_ActiveScen_MX.Keys(lng_TradeCtr - 1)
                dblArr_ActiveMXValues = dic_ActiveScen_MX.Items(lng_TradeCtr - 1)
                dbl_PnL_MX = dblArr_ActiveMXValues(1)
                dbl_PnLChg_MX = dblArr_ActiveMXValues(2)

                If dic_Overall_RE.Exists(str_Active_RECont & lng_ActiveScen_MX) Then
                    Set dic_ByScen_RE = dic_Overall_RE(str_Active_RECont & lng_ActiveScen_MX)
                    If dic_ByScen_RE.Exists(str_Active_MXID) Then
                        arr_Line_RE = dic_ByScen_RE(str_Active_MXID)
                        int_ActiveCol = 1
                        varArr_Results(lng_ActiveRowNum, int_ActiveCol) = arr_Line_RE(1)  ' RE trade ID

                        int_ActiveCol = int_ActiveCol + 1
                        varArr_Results(lng_ActiveRowNum, int_ActiveCol) = str_Active_MXID

                        int_ActiveCol = int_ActiveCol + 1
                        varArr_Results(lng_ActiveRowNum, int_ActiveCol) = arr_Line_RE(3)  ' Sheet name

                        int_ActiveCol = int_ActiveCol + 1
                        varArr_Results(lng_ActiveRowNum, int_ActiveCol) = arr_Line_RE(2)  ' RE scenario

                        int_ActiveCol = int_ActiveCol + 1
                        varArr_Results(lng_ActiveRowNum, int_ActiveCol) = lng_ActiveScen_MX

                        int_ActiveCol = int_ActiveCol + 1
                        varArr_Results(lng_ActiveRowNum, int_ActiveCol) = arr_Line_RE(4)  ' RE result

                        int_ActiveCol = int_ActiveCol + 1
                        varArr_Results(lng_ActiveRowNum, int_ActiveCol) = dbl_PnL_MX

                        int_ActiveCol = int_ActiveCol + 1
                        varArr_Results(lng_ActiveRowNum, int_ActiveCol) = arr_Line_RE(5)  ' RE result variation

                        int_ActiveCol = int_ActiveCol + 1
                        varArr_Results(lng_ActiveRowNum, int_ActiveCol) = dbl_PnLChg_MX
                    Else
                        lng_NotFoundCtr = lng_NotFoundCtr + 1
                    End If
                Else
                    lng_NotFoundCtr = lng_NotFoundCtr + 1
                End If
            Next lng_TradeCtr
        Next lng_ScenCtr
    Next lng_ContainerCtr

    ' Output data to sheet and add comparison formulas if there was data output
    If lng_ActiveRowNum >= 1 Then
        rng_TopLeft_Comparison.Resize(lng_ActiveRowNum, 9).Value = varArr_Results

        int_ActiveCol = int_ActiveCol + 1
        rng_TopLeft_Comparison(1, int_ActiveCol).Formula = "=(G2-F2)/R2*10000"
        Call dic_Addresses.Add("D_Result_Abs", rng_TopLeft_Comparison(1, int_ActiveCol).Address(False, False))

        int_ActiveCol = int_ActiveCol + 1
        rng_TopLeft_Comparison(1, int_ActiveCol).Formula = "=IF(F2=0,0,G2/F2-1)"
        Call dic_Addresses.Add("D_Result_Rel", rng_TopLeft_Comparison(1, int_ActiveCol).Address(False, False))

        int_ActiveCol = int_ActiveCol + 1
        rng_TopLeft_Comparison(1, int_ActiveCol).Formula = "=(I2-H2)/R2 *10000"
        Call dic_Addresses.Add("D_ResultV_Abs", rng_TopLeft_Comparison(1, int_ActiveCol).Address(False, False))

        int_ActiveCol = int_ActiveCol + 1
        rng_TopLeft_Comparison(1, int_ActiveCol).Formula = "=IF(H2=0,0,I2/H2-1)"
        Call dic_Addresses.Add("D_ResultV_Rel", rng_TopLeft_Comparison(1, int_ActiveCol).Address(False, False))

        int_ActiveCol = int_ActiveCol + 1
        rng_TopLeft_Comparison(1, int_ActiveCol).Formula = "=IF(AND(ABS(" & dic_Addresses("D_Result_Rel") & ")>$U$2,ABS(" _
            & dic_Addresses("D_Result_Abs") & ")>$U$3),""YES"","""")"
        Call dic_Addresses.Add("Result_Sig", rng_TopLeft_Comparison(1, int_ActiveCol).Address(False, False))

        int_ActiveCol = int_ActiveCol + 1
        rng_TopLeft_Comparison(1, int_ActiveCol).Formula = "=IF(AND(ABS(" & dic_Addresses("D_ResultV_Rel") & ")>$U$2,ABS(" _
            & dic_Addresses("D_ResultV_Abs") & ")>$U$3),""YES"","""")"
        Call dic_Addresses.Add("ResultV_Sig", rng_TopLeft_Comparison(1, int_ActiveCol).Address(False, False))

        int_ActiveCol = int_ActiveCol + 1
        rng_TopLeft_Comparison(1, int_ActiveCol).Formula = "=IF(OR(" & dic_Addresses("Result_Sig") & "=""YES""," _
            & dic_Addresses("ResultV_Sig") & "=""YES""),""YES"","""")"

        int_ActiveCol = int_ActiveCol + 1
        rng_TopLeft_Comparison(1, int_ActiveCol).Formula = "=IF(OR(AND((F2=0)<>(G2=0),ABS(" & dic_Addresses("D_Result_Abs") _
            & ")>$U$3),AND((H2=0)<>(I2=0),ABS(" & dic_Addresses("D_ResultV_Abs") & ")>$U$3)),""YES"","""")"

        int_ActiveCol = int_ActiveCol + 1
        rng_TopLeft_Comparison(1, int_ActiveCol).Formula = 1

        ' Fill down if required
        If lng_ActiveRowNum >= 2 Then
            wks_Comparison.Range("J2:R2").Resize(lng_ActiveRowNum, 9).FillDown
        End If

        wks_Comparison.Calculate
    End If

    If lng_NotFoundCtr > 0 Then
        MsgBox "Error - RE results missing:" & lng_NotFoundCtr
    End If
End Sub