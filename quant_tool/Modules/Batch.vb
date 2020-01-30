Option Explicit

Public Function Batch_OutputTopLeft() As Range
    Set Batch_OutputTopLeft = ThisWorkbook.Worksheets("Batch Output").Range("A2")
End Function

Public Sub RunBatch()
    ' ## Output based on either defined pillars or native pillars
    Dim fld_AppState_Orig As ApplicationState: fld_AppState_Orig = Gather_ApplicationState(ApplicationStateType.Current)
    Dim fld_AppState_Opt As ApplicationState: fld_AppState_Opt = Gather_ApplicationState(ApplicationStateType.Optimized)
    Call Action_SetAppState(fld_AppState_Opt)

    Dim sng_Time_Start As Single: sng_Time_Start = Timer

    Dim wks_BatchInputs As Worksheet: Set wks_BatchInputs = ThisWorkbook.Worksheets("Batch Generation")
    Dim rng_ActiveDates As Range: Set rng_ActiveDates = wks_BatchInputs.Range("A3:C3")
    Dim rng_OutputTopLeft As Range: Set rng_OutputTopLeft = Batch_OutputTopLeft()
    Dim str_DataType As String: str_DataType = UCase(wks_BatchInputs.Range("F2").Value)
    Dim str_OutputType As String: str_OutputType = UCase(wks_BatchInputs.Range("F3").Value)
    Dim bln_QueryDB As Boolean: bln_QueryDB = (UCase(wks_BatchInputs.Range("F4").Value) = "YES")
    Dim lng_ActiveBuildDate As Long, rng_ActiveCurveName As Range, rng_ActiveOutput As Range
    Dim irc_ActiveCurve As Data_IRCurve, cvl_ActiveCurve As Data_CapVols, fxv_Active As Data_FXVols, svl_Active As Data_SwptVols
    Dim dic_StaticInfo As Dictionary: Set dic_StaticInfo = GetStaticInfo()
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    Dim dic_CurveSet As Dictionary: Set dic_CurveSet = GetAllCurves(True, False, dic_StaticInfo)
    Dim fxs_Spots As Data_FXSpots: Set fxs_Spots = dic_CurveSet(CurveType.FXSPT)
    Dim lng_FoundRow As Long

    ' Variable headers
    With rng_OutputTopLeft
        Select Case str_DataType
            Case "FXV"
                .Offset(-1, 0).Value = "Fgn"
                .Offset(-1, 1).Value = "Dom"
                .Offset(-1, 2).Value = "Days"
                .Offset(-1, 3).Value = "Call Delta"
                .Offset(-1, 4).Value = "Data Date"
                .Offset(-1, 5).Value = "Vol"
            Case "SVL"
                .Offset(-1, 0).Value = "Curve Name"
                .Offset(-1, 1).Value = "Data Date"
                .Offset(-1, 2).Value = "Days"
                .Offset(-1, 3).Value = "Value"
                .Offset(-1, 4).Value = "Swap Mat"
                .Offset(-1, 5).Value = "StrikeSprd"
            Case Else
                .Offset(-1, 0).Value = "Curve Name"
                .Offset(-1, 1).Value = "Data Date"
                .Offset(-1, 2).Value = "Days"
                .Offset(-1, 3).Value = "Value"
                .Offset(-1, 4).Value = "Type"
                .Offset(-1, 5).Value = "Pillar Set"
        End Select
    End With

    While rng_ActiveDates(1, 1).Value <> ""
        ' Set system date
        lng_ActiveBuildDate = rng_ActiveDates(1, 1).Value
        cfg_Settings.CurrentBuildDate = lng_ActiveBuildDate
        cfg_Settings.CurrentDataDate = rng_ActiveDates(1, 2).Value
        Call cfg_Settings.SetCurrentValDate(rng_ActiveDates(1, 3).Value)

        Application.StatusBar = "Generating output for: " & Format(lng_ActiveBuildDate, "dd/mm/yyyy")

        Select Case str_DataType
            Case "IRC"
                If bln_QueryDB = True Then
                    Call fxs_Spots.LoadRates
                    Call GenerateSelectedIRCurves(dic_CurveSet(CurveType.IRC), False)
                End If

                Set rng_ActiveCurveName = GetRange_CurveSetup(CurveType.IRC)  ' Reset to top of the list of curves
                While rng_ActiveCurveName(1, 1).Value <> ""
                    If rng_ActiveCurveName(1, 2).Value = "YES" Then
                        Set irc_ActiveCurve = GetObject_IRCurve(rng_ActiveCurveName(1, 1).Value, True, False, dic_StaticInfo)
                        Set rng_ActiveOutput = Gather_RowForAppend(rng_OutputTopLeft)

                        Select Case str_OutputType
                            Case "BOOTSTRAPPED": Call irc_ActiveCurve.Output_ZeroRates(rng_ActiveOutput)
                            Case "RAW": Call irc_ActiveCurve.Output_ParRates(rng_ActiveOutput)
                            Case "BOTH"
                                Call irc_ActiveCurve.Output_ZeroRates(rng_ActiveOutput)
                                Set rng_ActiveOutput = Gather_RowForAppend(rng_OutputTopLeft)
                                Call irc_ActiveCurve.Output_ParRates(rng_ActiveOutput)
                        End Select
                    End If

                    Set rng_ActiveCurveName = rng_ActiveCurveName.Offset(1, 0)
                Wend
            Case "CVL"
                If bln_QueryDB = True Then
                    Call fxs_Spots.LoadRates

                    If str_OutputType = "BOOTSTRAPPED" Then Call GenerateSelectedIRCurves(dic_CurveSet(CurveType.IRC))
                    Call GenerateSelectedCapVolCurves(dic_CurveSet(CurveType.cvl))
                End If

                Set rng_ActiveCurveName = GetRange_CurveSetup(CurveType.cvl)
                While rng_ActiveCurveName(1, 1).Value <> ""
                    If rng_ActiveCurveName(1, 2).Value = "YES" Then
                        Set cvl_ActiveCurve = GetObject_CapVols(rng_ActiveCurveName(1, 1).Value, True, False, dic_StaticInfo)
                        Set rng_ActiveOutput = Gather_RowForAppend(rng_OutputTopLeft)

                        Select Case str_OutputType
                            Case "BOOTSTRAPPED": Call cvl_ActiveCurve.OutputFinalVols(rng_ActiveOutput)
                            Case "RAW": Call cvl_ActiveCurve.OutputOrigCapVols(rng_ActiveOutput)
                            Case "BOTH"
                                Call cvl_ActiveCurve.OutputFinalVols(rng_ActiveOutput)
                                Set rng_ActiveOutput = Gather_RowForAppend(rng_ActiveOutput)
                                Call cvl_ActiveCurve.OutputOrigCapVols(rng_ActiveOutput)
                        End Select
                    End If

                    Set rng_ActiveCurveName = rng_ActiveCurveName.Offset(1, 0)
                Wend
            Case "FXV"
                If bln_QueryDB = True Then Call GenerateSelectedFXVolCurves(dic_CurveSet(CurveType.FXV), False)

                Set rng_ActiveCurveName = GetRange_CurveSetup(CurveType.FXV)
                While rng_ActiveCurveName(1, 1).Value <> ""
                    If rng_ActiveCurveName(1, 2).Value = "YES" Then
                        Set fxv_Active = GetObject_FXVols(rng_ActiveCurveName(1, 1).Value, True, False, dic_StaticInfo)
                        Set rng_ActiveOutput = Gather_RowForAppend(rng_OutputTopLeft)
                        Call fxv_Active.OutputFinalVols(rng_ActiveOutput)
                    End If

                    Set rng_ActiveCurveName = rng_ActiveCurveName.Offset(1, 0)
                Wend
            Case "SVL"
                If bln_QueryDB = True Then Call GenerateSelectedSwptVolCurves(dic_CurveSet(CurveType.SVL))

                Set rng_ActiveCurveName = GetRange_CurveSetup(CurveType.SVL)
                While rng_ActiveCurveName(1, 1).Value <> ""
                    If rng_ActiveCurveName(1, 2).Value = "YES" Then
                        Set svl_Active = GetObject_SwptVols(rng_ActiveCurveName(1, 1).Value, True, False, dic_StaticInfo)
                        Set rng_ActiveOutput = Gather_RowForAppend(rng_OutputTopLeft)
                        Call svl_Active.OutputFinalVols(rng_ActiveOutput)
                    End If

                    Set rng_ActiveCurveName = rng_ActiveCurveName.Offset(1, 0)
                Wend
        End Select

        Set rng_ActiveDates = rng_ActiveDates.Offset(1, 0)
    Wend

    Call GotoSheet(wks_BatchInputs.Name)

    Dim sng_Time_End As Single: sng_Time_End = Timer
    Debug.Print "Batch Generation - Time elapsed: " & Round(sng_Time_End - sng_Time_Start, 1) & " seconds"

    Call Action_SetAppState(fld_AppState_Orig)
End Sub

Public Sub ClearBatch()
    Call Action_ClearBelow(Batch_OutputTopLeft, 6)
End Sub

Public Sub ViewBatchOutput()
    Call GotoSheet("Batch Output")
End Sub