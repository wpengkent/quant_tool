Attribute VB_Name = "GetInstParams"
Option Explicit
' ## Functions to read instrument parameters from a row of a booking sheet
Public Function GetInstParams_RngAcc(rng_Input As Range, dic_StaticInfo As Dictionary, Optional str_TradeID As String = "-") As InstParams_RngAcc
    Dim fld_Output As InstParams_RngAcc
    ' Gather generator level information
    Dim igs_Generators As IRGeneratorSet: Set igs_Generators = dic_StaticInfo(StaticInfoType.IRGeneratorSet)
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    '******************Generator Finder******************************
    Dim str_Gen_LegA As String: str_Gen_LegA = rng_Input(1, 26).Value
    Dim str_Gen_LegA2Digi As String: str_Gen_LegA2Digi = rng_Input(1, 27).Value
    Dim str_Gen_LegB As String: str_Gen_LegB = rng_Input(1, 51).Value
    Dim str_Gen_LegB2Digi As String: str_Gen_LegB2Digi = rng_Input(1, 52).Value
    '****************************************************************
    Dim fld_LegA As IRLegParams: fld_LegA = igs_Generators.Lookup_Generator(str_Gen_LegA)
    Dim fld_LegB As IRLegParams: fld_LegB = igs_Generators.Lookup_Generator(str_Gen_LegB)
    Dim fld_LegA2Digi As IRLegParams: fld_LegA2Digi = igs_Generators.Lookup_Generator(str_Gen_LegA2Digi)
    Dim fld_LegB2Digi As IRLegParams: fld_LegB2Digi = igs_Generators.Lookup_Generator(str_Gen_LegB2Digi)

    Dim dcp_ActiveParams As DictParams
    Dim str_Custom As String, wks_Custom As Worksheet
    Dim int_ctr As Integer

    ' Shared
    fld_LegA.TradeID = str_TradeID
    fld_LegB.TradeID = str_TradeID
    fld_LegA.IsFwdGeneration = True
    fld_LegB.IsFwdGeneration = True
    fld_LegA.IsUniformPeriods = False
    fld_LegB.IsUniformPeriods = False
    fld_Output.GeneratorA = str_Gen_LegA
    fld_Output.GeneratorB = str_Gen_LegB

    int_ctr = 1
    If rng_Input(1, int_ctr).Value = "<SYSTEM>" Then
        fld_LegA.ValueDate = cfg_Settings.CurrentValDate
    Else
        fld_LegA.ValueDate = rng_Input(1, int_ctr).Value
    End If
    fld_LegB.ValueDate = fld_LegA.ValueDate

    int_ctr = int_ctr + 1
    fld_LegA.Swapstart = rng_Input(1, int_ctr).Value
    fld_LegB.Swapstart = fld_LegA.Swapstart

    int_ctr = int_ctr + 1
    fld_LegA.GenerationRefPoint = rng_Input(1, int_ctr).Value
    fld_LegB.GenerationRefPoint = fld_LegA.GenerationRefPoint

    int_ctr = int_ctr + 1
    fld_LegA.Term = rng_Input(1, int_ctr).Value
    fld_LegB.Term = fld_LegA.Term

    int_ctr = int_ctr + 1
    fld_Output.Pay_LegA = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_LegA.PExch_Start = rng_Input(1, int_ctr).Value
    fld_LegB.PExch_Start = fld_LegA.PExch_Start

    int_ctr = int_ctr + 1
    fld_LegA.PExch_Intermediate = rng_Input(1, int_ctr).Value
    fld_LegB.PExch_Intermediate = fld_LegA.PExch_Intermediate

    int_ctr = int_ctr + 1
    fld_LegA.PExch_End = rng_Input(1, int_ctr).Value
    fld_LegB.PExch_End = fld_LegA.PExch_End

    int_ctr = int_ctr + 1
    fld_LegA.FloatEst = rng_Input(1, int_ctr).Value
    fld_LegB.FloatEst = fld_LegA.FloatEst

    int_ctr = int_ctr + 1
    fld_LegA.ForceToMV = rng_Input(1, int_ctr).Value
    fld_LegB.ForceToMV = fld_LegA.ForceToMV

    int_ctr = int_ctr + 1
    fld_Output.CCY_PnL = UCase(rng_Input(1, int_ctr).Value)

    int_ctr = int_ctr + 1
    If fld_LegA.index <> "-" Or fld_LegA.index <> "" Then
        fld_LegA.StubInterpolate = rng_Input(1, int_ctr).Value
    End If

    If fld_LegB.index <> "-" Or fld_LegB.index <> "" Then
        fld_LegB.StubInterpolate = rng_Input(1, int_ctr).Value
    End If

    int_ctr = int_ctr + 1
    If fld_LegA.index <> "-" Or fld_LegA.index <> "" Then
        fld_LegA.FixInArrears = rng_Input(1, int_ctr).Value
    End If

    If fld_LegB.index <> "-" Or fld_LegB.index <> "" Then
        fld_LegB.FixInArrears = rng_Input(1, int_ctr).Value
    End If

    int_ctr = int_ctr + 1
    If fld_LegA.index <> "-" Or fld_LegA.index <> "" Then
        fld_LegA.DisableConvAdj = rng_Input(1, int_ctr).Value
    End If

    If fld_LegB.index <> "-" Or fld_LegB.index <> "" Then
        fld_LegB.DisableConvAdj = rng_Input(1, int_ctr).Value
    End If

    int_ctr = int_ctr + 1
    str_Custom = rng_Input(1, int_ctr).Value
    Select Case str_Custom
        Case "-", ""
        Case Else
            Set wks_Custom = ThisWorkbook.Worksheets(str_Custom)
    End Select

    int_ctr = int_ctr + 1
    If rng_Input(1, int_ctr).Value <> "-" Then
        Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
        Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
        Set fld_LegA.AmortSchedule = Gather_Dictionary(dcp_ActiveParams, False)
        Set fld_LegB.AmortSchedule = fld_LegA.AmortSchedule
    End If

    'KL - 201901 For HW1F
    int_ctr = int_ctr + 1
    fld_Output.IsCallable = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.Callable_LegA = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.VolCurve = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.SpotStep = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.TimeStep = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.MeanRev = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    If rng_Input(1, int_ctr).Value <> "-" Then
        Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
        Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
        Set fld_Output.CallDate = Gather_Dictionary(dcp_ActiveParams, False)

        Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value).Offset(0, 1)
        Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
        Set fld_Output.Swapstart = Gather_Dictionary(dcp_ActiveParams, False)
    End If

    ' Leg A
    With fld_LegA
        int_ctr = int_ctr + 1
         If rng_Input(1, int_ctr).Value <> "-" Then
           .Notional = rng_Input(1, int_ctr).Value
           Else
           .Notional = 0
           End If

        int_ctr = int_ctr + 1
        .CCY = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 3
        .FixedFloat = rng_Input(1, int_ctr).Value
        If rng_Input(1, int_ctr).Value = "Fixed" Then
        .index = "-"
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "-" Then
         .RateOrMargin = rng_Input(1, int_ctr).Value
         Else
         .RateOrMargin = 0
         End If

         int_ctr = int_ctr + 1
         .ExoticType = rng_Input(1, int_ctr).Value
         If (CStr(rng_Input(1, int_ctr).Value) = "Range") Or (CStr(rng_Input(1, int_ctr).Value) = "Callable Range") Then
         .IsDigital = True
         End If

         int_ctr = int_ctr + 1
         .Schedule = rng_Input(1, int_ctr).Value

         int_ctr = int_ctr + 1
         .NbofDays = rng_Input(1, int_ctr).Value

         int_ctr = int_ctr + 1
         .PerDayShifter = rng_Input(1, int_ctr).Value

         int_ctr = int_ctr + 1
         .GlobalShifter = rng_Input(1, int_ctr).Value

         int_ctr = int_ctr + 1
         .ApplyTo = rng_Input(1, int_ctr).Value

         int_ctr = int_ctr + 1
         .RangeIndex = rng_Input(1, int_ctr).Value

          int_ctr = int_ctr + 1
          If rng_Input(1, int_ctr).Value <> "-" Then
          .RateFactor = rng_Input(1, int_ctr).Value
          Else
          .RateFactor = 0
          End If

         int_ctr = int_ctr + 1
         If rng_Input(1, int_ctr).Value <> "-" Then
         .Correl = rng_Input(1, int_ctr).Value
         Else
         .Correl = 0
         End If

         int_ctr = int_ctr + 1
        .AboveUpper = rng_Input(1, int_ctr).Value


          int_ctr = int_ctr + 1
          If rng_Input(1, int_ctr).Value <> "-" Then
         .Upper = rng_Input(1, int_ctr).Value
         Else
         .Upper = 0
         End If


         int_ctr = int_ctr + 1
        .AboveLower = rng_Input(1, int_ctr).Value


         int_ctr = int_ctr + 1
          If rng_Input(1, int_ctr).Value <> "-" Then
         .Lower = rng_Input(1, int_ctr).Value
            Else
         .Lower = 0
         End If

         int_ctr = int_ctr + 1
          If rng_Input(1, int_ctr).Value <> "-" Then
         .Lockout = rng_Input(1, int_ctr).Value
          Else
         .Lockout = "-"
         End If

         int_ctr = int_ctr + 1
          If rng_Input(1, int_ctr).Value <> "-" Then
         .Lockoutmode = rng_Input(1, int_ctr).Value
          Else
         .Lockoutmode = "-"
         End If


         int_ctr = int_ctr + 1
          If rng_Input(1, int_ctr).Value <> "-" Then
         .FirstnLastDay = rng_Input(1, int_ctr).Value
          Else
         .FirstnLastDay = "-"
         End If

          int_ctr = int_ctr + 1
         If rng_Input(1, int_ctr).Value <> "-" Then
          Dim splitter1 As Variant  '# Alvin edit 2/10/2018
          splitter1 = Split(rng_Input(1, int_ctr).Value, "|") '# Alvin edit 2/10/2018

          'For main CF Fixings
          Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(splitter1(0))
          Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
          Set .Fixings = Gather_Dictionary(dcp_ActiveParams, False)

          'For Digitals Fixing
          Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(splitter1(1))
          Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
          Set fld_LegA2Digi.FixingsDigi = Gather_Dictionary(dcp_ActiveParams, False)

          End If

          int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "-" Then
          Dim splitter2 As Variant
          splitter2 = Split(rng_Input(1, int_ctr).Value, "|") '# Alvin edit 2/10/2018

          'For main CF Fixings
          Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(splitter2(0))
          Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
          Set .ModStarts = Gather_Dictionary(dcp_ActiveParams)

          'For Digitals Fixing
          Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(splitter2(1))
          Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
          Set fld_LegA2Digi.ModStartsDigi = Gather_Dictionary(dcp_ActiveParams)
         End If
    End With

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "-" Then
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set fld_LegA.VariableRate = Gather_Dictionary(dcp_ActiveParams, False)

            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 2)
            Set fld_LegA.VariableRange1 = Gather_Dictionary(dcp_ActiveParams, False)

            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 3)
            Set fld_LegA.VariableRange2 = Gather_Dictionary(dcp_ActiveParams, False)

            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 4)
            Set fld_LegA.VariableRange3 = Gather_Dictionary(dcp_ActiveParams, False)

            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 5)
            Set fld_LegA.VariableRange4 = Gather_Dictionary(dcp_ActiveParams, False)
        End If

    ' Leg B
   With fld_LegB
        int_ctr = int_ctr + 1
        .Notional = rng_Input(1, int_ctr).Value
         If rng_Input(1, int_ctr).Value <> "-" Then
           .Notional = rng_Input(1, int_ctr).Value
           Else
           .Notional = 0
           End If

        int_ctr = int_ctr + 1
        .CCY = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 3
        .FixedFloat = rng_Input(1, int_ctr).Value
        If rng_Input(1, int_ctr).Value = "Fixed" Then
        .index = "-"
        End If

            int_ctr = int_ctr + 1
          If rng_Input(1, int_ctr).Value <> "-" Then
           .RateOrMargin = rng_Input(1, int_ctr).Value
           Else
           .RateOrMargin = 0
           End If

          int_ctr = int_ctr + 1
           .ExoticType = rng_Input(1, int_ctr).Value
           If (CStr(rng_Input(1, int_ctr).Value) = "Range") Or (CStr(rng_Input(1, int_ctr).Value) = "Callable Range") Then
           .IsDigital = True
           End If

           int_ctr = int_ctr + 1
           .Schedule = rng_Input(1, int_ctr).Value

           int_ctr = int_ctr + 1
           .NbofDays = rng_Input(1, int_ctr).Value

            int_ctr = int_ctr + 1
           .PerDayShifter = rng_Input(1, int_ctr).Value

           int_ctr = int_ctr + 1
           .GlobalShifter = rng_Input(1, int_ctr).Value

           int_ctr = int_ctr + 1
           .ApplyTo = rng_Input(1, int_ctr).Value

           int_ctr = int_ctr + 1
           .RangeIndex = rng_Input(1, int_ctr).Value


            int_ctr = int_ctr + 1
            If rng_Input(1, int_ctr).Value <> "-" Then
            .RateFactor = rng_Input(1, int_ctr).Value
            Else
            .RateFactor = 0
            End If

           int_ctr = int_ctr + 1
           If rng_Input(1, int_ctr).Value <> "-" Then
           .Correl = rng_Input(1, int_ctr).Value
           Else
           .Correl = 0
           End If

            int_ctr = int_ctr + 1
          .AboveUpper = rng_Input(1, int_ctr).Value

             int_ctr = int_ctr + 1
            If rng_Input(1, int_ctr).Value <> "-" Then
           .Upper = rng_Input(1, int_ctr).Value
           Else
           .Upper = 0
           End If


           int_ctr = int_ctr + 1
          .AboveLower = rng_Input(1, int_ctr).Value

           int_ctr = int_ctr + 1
            If rng_Input(1, int_ctr).Value <> "-" Then
           .Lower = rng_Input(1, int_ctr).Value
              Else
           .Lower = 0
           End If

            int_ctr = int_ctr + 1
            If rng_Input(1, int_ctr).Value <> "-" Then
           .Lockout = rng_Input(1, int_ctr).Value
            Else
           .Lockout = "-"
           End If

           int_ctr = int_ctr + 1
            If rng_Input(1, int_ctr).Value <> "-" Then
           .Lockoutmode = rng_Input(1, int_ctr).Value
            Else
           .Lockoutmode = "-"
           End If


           int_ctr = int_ctr + 1
            If rng_Input(1, int_ctr).Value <> "-" Then
           .FirstnLastDay = rng_Input(1, int_ctr).Value
            Else
           .FirstnLastDay = "-"
           End If

           int_ctr = int_ctr + 1
           If rng_Input(1, int_ctr).Value <> "-" Then
            Dim splitter3 As Variant  '# Alvin edit 2/10/2018
            splitter3 = Split(rng_Input(1, int_ctr).Value, "|") '# Alvin edit 2/10/2018

            'Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value) 'Alvin edit 2/10/2018

            'For main CF Fixings
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(splitter3(0))
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set .Fixings = Gather_Dictionary(dcp_ActiveParams, False)

            'For Digitals Fixing
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(splitter3(1))
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set fld_LegB2Digi.FixingsDigi = Gather_Dictionary(dcp_ActiveParams, False)

            End If

            int_ctr = int_ctr + 1
          If rng_Input(1, int_ctr).Value <> "-" Then
            Dim splitter4 As Variant
            splitter4 = Split(rng_Input(1, int_ctr).Value, "|") '# Alvin edit 2/10/2018

            'For main CF Fixings
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(splitter4(0))
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set .ModStarts = Gather_Dictionary(dcp_ActiveParams)

            'For Digitals Fixing
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(splitter4(1))
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set fld_LegB2Digi.ModStartsDigi = Gather_Dictionary(dcp_ActiveParams)
           End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "-" Then
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set fld_LegB.VariableRate = Gather_Dictionary(dcp_ActiveParams, False)

            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 2)
            Set fld_LegB.VariableRange1 = Gather_Dictionary(dcp_ActiveParams, False)

            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 3)
            Set fld_LegB.VariableRange2 = Gather_Dictionary(dcp_ActiveParams, False)

            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 4)
            Set fld_LegB.VariableRange3 = Gather_Dictionary(dcp_ActiveParams, False)

            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 5)
            Set fld_LegB.VariableRange4 = Gather_Dictionary(dcp_ActiveParams, False)
        End If

End With

    fld_Output.TradeID = str_TradeID

    'Clone LegA necessary info to LegADigi
    fld_LegA2Digi.ExoticType = fld_LegA.ExoticType
    fld_LegA2Digi.IsDigital = fld_LegA.IsDigital
    fld_LegA2Digi.Schedule = fld_LegA.Schedule
    fld_LegA2Digi.ApplyTo = fld_LegA.ApplyTo
    fld_LegA2Digi.RangeIndex = fld_LegA.RangeIndex
    fld_LegA2Digi.Lockout = fld_LegA.Lockout
    fld_LegA2Digi.Lockoutmode = fld_LegA.Lockoutmode
    fld_LegA2Digi.PerDayShifter = fld_LegA.PerDayShifter
    fld_LegA2Digi.FirstnLastDay = fld_LegA.FirstnLastDay
    fld_LegA2Digi.AboveUpper = fld_LegA.AboveUpper
    fld_LegA2Digi.AboveLower = fld_LegA.AboveLower
    fld_LegA2Digi.Upper = fld_LegA.Upper
    fld_LegA2Digi.Lower = fld_LegA.Lower
    fld_LegA2Digi.ValueDate = fld_LegA.ValueDate
    fld_LegA2Digi.NbofDays = fld_LegA.NbofDays
'    fld_LegA2Digi.FixingsDigi = fld_LegA.FixingsDigi
'    fld_LegA2Digi.ModStartsDigi = fld_LegA.ModStartsDigi

    'Clone LegB necessary info to LegBDigi
    fld_LegB2Digi.ExoticType = fld_LegB.ExoticType
    fld_LegB2Digi.IsDigital = fld_LegB.IsDigital
    fld_LegB2Digi.Schedule = fld_LegB.Schedule
    fld_LegB2Digi.ApplyTo = fld_LegB.ApplyTo
    fld_LegB2Digi.RangeIndex = fld_LegB.RangeIndex
    fld_LegB2Digi.Lockout = fld_LegB.Lockout
    fld_LegB2Digi.Lockoutmode = fld_LegB.Lockoutmode
    fld_LegB2Digi.PerDayShifter = fld_LegB.PerDayShifter
    fld_LegB2Digi.FirstnLastDay = fld_LegB.FirstnLastDay
    fld_LegB2Digi.AboveUpper = fld_LegB.AboveUpper
    fld_LegB2Digi.AboveLower = fld_LegB.AboveLower
    fld_LegB2Digi.Upper = fld_LegB.Upper
    fld_LegB2Digi.Lower = fld_LegB.Lower
    fld_LegB2Digi.ValueDate = fld_LegB.ValueDate
    fld_LegB2Digi.NbofDays = fld_LegB.NbofDays
'    fld_LegB2Digi.FixingsDigi = fld_LegB2.FixingsDigi
'    fld_LegB2Digi.ModStartsDigi = fld_LegB2.ModStartsDigi

    fld_Output.LegA = fld_LegA
    fld_Output.LegB = fld_LegB

    fld_Output.LegA2Digi = fld_LegA2Digi
    fld_Output.LegB2Digi = fld_LegB2Digi

    GetInstParams_RngAcc = fld_Output
End Function
Public Function GetInstParams_IRS(rng_Input As Range, dic_StaticInfo As Dictionary, Optional str_TradeID As String = "-") As InstParams_IRS
    Dim fld_Output As InstParams_IRS

    ' Gather generator level information
    Dim igs_Generators As IRGeneratorSet: Set igs_Generators = dic_StaticInfo(StaticInfoType.IRGeneratorSet)
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    Dim str_Gen_LegA As String: str_Gen_LegA = rng_Input(1, 18).Value
    Dim str_Gen_LegB As String: str_Gen_LegB = rng_Input(1, 23).Value
    Dim fld_LegA As IRLegParams: fld_LegA = igs_Generators.Lookup_Generator(str_Gen_LegA)
    Dim fld_LegB As IRLegParams: fld_LegB = igs_Generators.Lookup_Generator(str_Gen_LegB)

    Dim dcp_ActiveParams As DictParams
    Dim str_Custom As String, wks_Custom As Worksheet
    Dim int_ctr As Integer

    ' Shared
    fld_LegA.TradeID = str_TradeID
    fld_LegB.TradeID = str_TradeID
    fld_LegA.IsFwdGeneration = True
    fld_LegB.IsFwdGeneration = True
    fld_LegA.IsUniformPeriods = False
    fld_LegB.IsUniformPeriods = False

    int_ctr = 1
    If rng_Input(1, int_ctr).Value = "<SYSTEM>" Then
        fld_LegA.ValueDate = cfg_Settings.CurrentValDate
    Else
        fld_LegA.ValueDate = rng_Input(1, int_ctr).Value
    End If
    fld_LegB.ValueDate = fld_LegA.ValueDate

    int_ctr = int_ctr + 1
    fld_LegA.Swapstart = rng_Input(1, int_ctr).Value
    fld_LegB.Swapstart = fld_LegA.Swapstart

    int_ctr = int_ctr + 1
    fld_LegA.GenerationRefPoint = rng_Input(1, int_ctr).Value
    fld_LegB.GenerationRefPoint = fld_LegA.GenerationRefPoint

    int_ctr = int_ctr + 1
    fld_LegA.Term = rng_Input(1, int_ctr).Value
    fld_LegB.Term = fld_LegA.Term

    int_ctr = int_ctr + 1
    fld_Output.Pay_LegA = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_LegA.PExch_Start = rng_Input(1, int_ctr).Value
    fld_LegB.PExch_Start = fld_LegA.PExch_Start

    int_ctr = int_ctr + 1
    fld_LegA.PExch_Intermediate = rng_Input(1, int_ctr).Value
    fld_LegB.PExch_Intermediate = fld_LegA.PExch_Intermediate

    int_ctr = int_ctr + 1
    fld_LegA.PExch_End = rng_Input(1, int_ctr).Value
    fld_LegB.PExch_End = fld_LegA.PExch_End

    int_ctr = int_ctr + 1
    fld_LegA.FloatEst = rng_Input(1, int_ctr).Value
    fld_LegB.FloatEst = fld_LegA.FloatEst

    int_ctr = int_ctr + 1
    fld_LegA.ForceToMV = rng_Input(1, int_ctr).Value
    fld_LegB.ForceToMV = fld_LegA.ForceToMV

    int_ctr = int_ctr + 1
    fld_Output.CCY_PnL = UCase(rng_Input(1, int_ctr).Value)

    int_ctr = int_ctr + 1
    If fld_LegA.index <> "-" Or fld_LegA.index <> "" Then
        fld_LegA.StubInterpolate = rng_Input(1, int_ctr).Value
    End If

    If fld_LegB.index <> "-" Or fld_LegB.index <> "" Then
        fld_LegB.StubInterpolate = rng_Input(1, int_ctr).Value
    End If

    int_ctr = int_ctr + 1
    If fld_LegA.index <> "-" Or fld_LegA.index <> "" Then
        fld_LegA.FixInArrears = rng_Input(1, int_ctr).Value
    End If

    If fld_LegB.index <> "-" Or fld_LegB.index <> "" Then
        fld_LegB.FixInArrears = rng_Input(1, int_ctr).Value
    End If

    int_ctr = int_ctr + 1
    If fld_LegA.index <> "-" Or fld_LegA.index <> "" Then
        fld_LegA.DisableConvAdj = rng_Input(1, int_ctr).Value
    End If

    If fld_LegB.index <> "-" Or fld_LegB.index <> "" Then
        fld_LegB.DisableConvAdj = rng_Input(1, int_ctr).Value
    End If

    int_ctr = int_ctr + 1
    str_Custom = rng_Input(1, int_ctr).Value
    Select Case str_Custom
        Case "-", ""
        Case Else
            Set wks_Custom = ThisWorkbook.Worksheets(str_Custom)
    End Select

    int_ctr = int_ctr + 1
    If rng_Input(1, int_ctr).Value <> "-" Then
        Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
        Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
        Set fld_LegA.AmortSchedule = Gather_Dictionary(dcp_ActiveParams, False)
        Set fld_LegB.AmortSchedule = fld_LegA.AmortSchedule
    End If

    ' Leg A
    With fld_LegA
        int_ctr = int_ctr + 1
        .Notional = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 2
        .RateOrMargin = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "-" Then
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set .Fixings = Gather_Dictionary(dcp_ActiveParams, False)
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "-" Then
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set .ModStarts = Gather_Dictionary(dcp_ActiveParams)
        End If
    End With

    ' Leg B
    With fld_LegB
        int_ctr = int_ctr + 1
        .Notional = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 2
        .RateOrMargin = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "-" Then
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set .Fixings = Gather_Dictionary(dcp_ActiveParams, False)
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "-" Then
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set .ModStarts = Gather_Dictionary(dcp_ActiveParams)
        End If
    End With

    fld_Output.TradeID = str_TradeID
    fld_Output.LegA = fld_LegA
    fld_Output.LegB = fld_LegB
    GetInstParams_IRS = fld_Output
End Function

Public Function GetInstParams_CFL(rng_Input As Range, dic_StaticInfo As Dictionary, Optional str_TradeID As String = "-") As InstParams_CFL
    Dim fld_Output As InstParams_CFL, fld_Premium As SCFParams
    Dim dcp_ActiveParams As DictParams

    ' Gather generator level information
    Dim igs_Generators As IRGeneratorSet: Set igs_Generators = dic_StaticInfo(StaticInfoType.IRGeneratorSet)
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    Dim str_Gen As String: str_Gen = rng_Input(1, 9).Value
    Dim fld_Underlying As IRLegParams: fld_Underlying = igs_Generators.Lookup_Generator(str_Gen)

    Dim str_Custom As String, wks_Custom As Worksheet
    Dim int_ctr As Integer: int_ctr = 0

    With fld_Underlying
        .TradeID = str_TradeID
        .PExch_Start = False
        .PExch_Intermediate = False
        .PExch_End = False
        .FloatEst = True
        .RateOrMargin = 0
        .ForceToMV = False
        .IsFwdGeneration = True
        .IsUniformPeriods = False

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "<SYSTEM>" Then
            .ValueDate = cfg_Settings.CurrentValDate
        Else
            .ValueDate = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        .Swapstart = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .GenerationRefPoint = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Term = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Notional = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        fld_Output.BuySell = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        Select Case UCase(rng_Input(1, int_ctr).Value)
            Case "CAP": fld_Output.Direction = OptionDirection.CallOpt
            Case "FLOOR": fld_Output.Direction = OptionDirection.PutOpt
        End Select

        int_ctr = int_ctr + 1
        fld_Output.IsDigital = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 2
        fld_Output.strike = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        fld_Output.CCY_PnL = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        fld_Output.VolCurve = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        str_Custom = rng_Input(1, int_ctr).Value
        Select Case str_Custom
            Case "-", ""
            Case Else
                Set wks_Custom = ThisWorkbook.Worksheets(str_Custom)
        End Select

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "-" Then
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set .AmortSchedule = Gather_Dictionary(dcp_ActiveParams, False)
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "-" Then
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set .Fixings = Gather_Dictionary(dcp_ActiveParams, False)
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "-" Then
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set .ModStarts = Gather_Dictionary(dcp_ActiveParams)
        End If
    End With

    ' Premium
    With fld_Premium
        .TradeID = str_TradeID

        int_ctr = int_ctr + 1
        .Amount = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .PmtDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_Disc = rng_Input(1, int_ctr).Value
        .Curve_SpotDisc = .Curve_Disc
    End With

    fld_Output.TradeID = str_TradeID
    fld_Output.Underlying = fld_Underlying
    fld_Output.Premium = fld_Premium
    GetInstParams_CFL = fld_Output
End Function

Public Function GetInstParams_SWT(rng_Input As Range, dic_StaticInfo As Dictionary, Optional str_TradeID As String = "-") As InstParams_SWT
    Dim fld_Output As InstParams_SWT, fld_Premium As SCFParams
    Dim str_Custom As String, wks_Custom As Worksheet
    Dim dcp_ActiveParams As DictParams
    Dim int_ctr As Integer: int_ctr = 0

    ' Gather generator level information
    Dim igs_Generators As IRGeneratorSet: Set igs_Generators = dic_StaticInfo(StaticInfoType.IRGeneratorSet)
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    Dim str_Gen_LegA As String: str_Gen_LegA = rng_Input(1, 27).Value
    Dim str_Gen_LegB As String: str_Gen_LegB = rng_Input(1, 31).Value
    Dim fld_LegA As IRLegParams: fld_LegA = igs_Generators.Lookup_Generator(str_Gen_LegA)
    Dim fld_LegB As IRLegParams: fld_LegB = igs_Generators.Lookup_Generator(str_Gen_LegB)

    ' Shared
    fld_LegA.TradeID = str_TradeID
    fld_LegB.TradeID = str_TradeID
    fld_LegA.ForceToMV = True
    fld_LegB.ForceToMV = True
    fld_LegA.IsFwdGeneration = True
    fld_LegB.IsFwdGeneration = True
    fld_LegA.IsUniformPeriods = False
    fld_LegB.IsUniformPeriods = False
    fld_Output.GeneratorA = str_Gen_LegA
    fld_Output.GeneratorB = str_Gen_LegB

    int_ctr = int_ctr + 1
    If rng_Input(1, int_ctr).Value = "<SYSTEM>" Then
        fld_Output.ValueDate = cfg_Settings.CurrentValDate
    Else
        fld_Output.ValueDate = rng_Input(1, int_ctr).Value
    End If

    int_ctr = int_ctr + 1
    fld_Output.OptionMat = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_LegA.Swapstart = rng_Input(1, int_ctr).Value
    fld_LegB.Swapstart = rng_Input(1, int_ctr).Value
    fld_LegA.ValueDate = fld_LegA.Swapstart
    fld_LegB.ValueDate = fld_LegB.Swapstart

    int_ctr = int_ctr + 1
    fld_LegA.GenerationRefPoint = rng_Input(1, int_ctr).Value
    fld_LegB.GenerationRefPoint = fld_LegA.GenerationRefPoint

    int_ctr = int_ctr + 1
    fld_LegA.Term = rng_Input(1, int_ctr).Value
    fld_LegB.Term = fld_LegA.Term

    int_ctr = int_ctr + 1
    fld_Output.BuySell = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.Pay_LegA = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.Exercise = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_LegA.PExch_Start = rng_Input(1, int_ctr).Value
    fld_LegB.PExch_Start = fld_LegA.PExch_Start

    int_ctr = int_ctr + 1
    fld_LegA.PExch_Intermediate = rng_Input(1, int_ctr).Value
    fld_LegB.PExch_Intermediate = fld_LegA.PExch_Intermediate

    int_ctr = int_ctr + 1
    fld_LegA.PExch_End = rng_Input(1, int_ctr).Value
    fld_LegB.PExch_End = fld_LegA.PExch_End

    int_ctr = int_ctr + 1
    fld_LegA.FloatEst = rng_Input(1, int_ctr).Value
    fld_LegB.FloatEst = fld_LegA.FloatEst

    int_ctr = int_ctr + 1
    fld_Output.CCY_PnL = UCase(rng_Input(1, int_ctr).Value)

    int_ctr = int_ctr + 1
    fld_Output.IsSmile = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.VolCurve = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    str_Custom = rng_Input(1, int_ctr).Value
    Select Case str_Custom
        Case "-", ""
        Case Else
            Set wks_Custom = ThisWorkbook.Worksheets(str_Custom)
    End Select

    int_ctr = int_ctr + 1
    If rng_Input(1, int_ctr).Value <> "-" Then
        Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
        Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
        Set fld_LegA.AmortSchedule = Gather_Dictionary(dcp_ActiveParams, False)
        Set fld_LegB.AmortSchedule = fld_LegA.AmortSchedule
    End If

    int_ctr = int_ctr + 1
    fld_Output.SpotStep = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.TimeStep = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.MeanRev = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    If rng_Input(1, int_ctr).Value <> "-" Then
        Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
        Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
        Set fld_Output.CallDate = Gather_Dictionary(dcp_ActiveParams, False)

        Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value).Offset(0, 1)
        Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
        Set fld_Output.Swapstart = Gather_Dictionary(dcp_ActiveParams, False)
    End If

    ' Premium
    With fld_Premium
        int_ctr = int_ctr + 1
        .Amount = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .PmtDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_Disc = rng_Input(1, int_ctr).Value
        .Curve_SpotDisc = .Curve_Disc
    End With

    ' Leg A
    With fld_LegA
        int_ctr = int_ctr + 1
        .Notional = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 2
        .RateOrMargin = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "-" Then
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set .ModStarts = Gather_Dictionary(dcp_ActiveParams)
        End If
    End With

    ' Leg B
    With fld_LegB
        int_ctr = int_ctr + 1
        .Notional = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 2
        .RateOrMargin = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "-" Then
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set .ModStarts = Gather_Dictionary(dcp_ActiveParams)
        End If
    End With

    fld_Output.TradeID = str_TradeID
    fld_Output.LegA = fld_LegA
    fld_Output.LegB = fld_LegB
    fld_Output.Premium = fld_Premium
    GetInstParams_SWT = fld_Output
End Function

Public Function GetInstParams_FXF(rng_Input As Range, dic_StaticInfo As Dictionary, Optional str_TradeID As String = "-") As InstParams_FXF
    Dim fld_Output As InstParams_FXF
    Dim fld_FlowA As SCFParams, fld_FlowB As SCFParams
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    Dim int_ctr As Integer: int_ctr = 0

    ' Shared
    fld_FlowA.TradeID = str_TradeID
    fld_FlowB.TradeID = str_TradeID

    int_ctr = int_ctr + 1
    If rng_Input(1, int_ctr).Value = "<SYSTEM>" Then
        fld_Output.ValueDate = cfg_Settings.CurrentValDate
    Else
        fld_Output.ValueDate = rng_Input(1, int_ctr).Value
    End If

    int_ctr = int_ctr + 1
    fld_FlowA.PmtDate = rng_Input(1, int_ctr).Value
    fld_FlowB.PmtDate = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.Pay_FlowA = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.CCY_PnL = UCase(rng_Input(1, int_ctr).Value)

    ' Flow A
    With fld_FlowA
        int_ctr = int_ctr + 1
        .Amount = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .Curve_Disc = rng_Input(1, int_ctr).Value
        .Curve_SpotDisc = .Curve_Disc
    End With

    ' Flow B
    With fld_FlowB
        int_ctr = int_ctr + 1
        .Amount = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .Curve_Disc = rng_Input(1, int_ctr).Value
        .Curve_SpotDisc = .Curve_Disc
    End With

    ' Prepare final output
    fld_Output.TradeID = str_TradeID
    fld_Output.FlowA = fld_FlowA
    fld_Output.FlowB = fld_FlowB

    GetInstParams_FXF = fld_Output
End Function

Public Function GetInstParams_DEP(rng_Input As Range, dic_StaticInfo As Dictionary, Optional str_TradeID As String = "-") As InstParams_DEP
    Dim fld_Output As InstParams_DEP
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    Dim int_ctr As Integer: int_ctr = 0

    With fld_Output
        fld_Output.TradeID = str_TradeID

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "<SYSTEM>" Then
            .ValueDate = cfg_Settings.CurrentValDate
        Else
            .ValueDate = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        .StartDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .MatDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsLoan = (UCase(rng_Input(1, int_ctr).Value) = "BORROW")

        int_ctr = int_ctr + 1
        .Principal = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY_Principal = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .PExch = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Rate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Daycount = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY_PnL = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .Curve_Disc = rng_Input(1, int_ctr).Value
    End With

    GetInstParams_DEP = fld_Output
End Function

Public Function GetInstParams_FRA(rng_Input As Range, dic_StaticInfo As Dictionary, Optional str_TradeID As String = "-") As InstParams_FRA
    Dim fld_Output As InstParams_FRA
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    Dim dic_Fixings As Dictionary
    Dim int_ctr As Integer: int_ctr = 0

    With fld_Output
        fld_Output.TradeID = str_TradeID

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "<SYSTEM>" Then
            .ValueDate = cfg_Settings.CurrentValDate
        Else
            .ValueDate = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        .StartDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsBuy = (UCase(rng_Input(1, int_ctr).Value) = "B")

        int_ctr = int_ctr + 1
        .Notional = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Generator = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Rate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Fixing = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY_PnL = UCase(rng_Input(1, int_ctr).Value)
    End With

    GetInstParams_FRA = fld_Output
End Function

'-------------------------------------------------------------------------------------------
' NAME:    GetInstParams_FVN
'
' PURPOSE: Get FX Vanilla parameters
'
' NOTES:
'
' INPUT OPTIONS:
'
' MODIFIED:
'    30JAN2019 - KW - Add Late Type (Late Cash, Late Delivery, Late Delivery ATM Spot)
'
'-------------------------------------------------------------------------------------------
Public Function GetInstParams_FVN(rng_Input As Range, dic_StaticInfo As Dictionary, Optional str_TradeID As String = "-") As InstParams_FVN
    ' ## Input is a horizontal range containing the parameter values
    Dim fld_Output As InstParams_FVN, fld_Premium As SCFParams
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    Dim int_ctr As Integer: int_ctr = 0

    With fld_Output
        .TradeID = str_TradeID

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "<SYSTEM>" Then
            .ValueDate = cfg_Settings.CurrentValDate
        Else
            .ValueDate = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        .MatDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .DelivDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "-" Then
            .LateType = "STANDARD"
        Else
            .LateType = UCase(rng_Input(1, int_ctr).Value)
        End If

        int_ctr = int_ctr + 1
        .IsBuy = (UCase(rng_Input(1, int_ctr).Value) = "B")

        int_ctr = int_ctr + 1
        .CCY_Fgn = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .CCY_Dom = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .Notional_Fgn = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .ExerciseType = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        Select Case UCase(rng_Input(1, int_ctr).Value)
            Case "C", "CALL": .Direction = OptionDirection.CallOpt
            Case "P", "PUT": .Direction = OptionDirection.PutOpt
        End Select

        int_ctr = int_ctr + 1
        .IsDigital = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .strike = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "-" Then
            .CCY_Payout = .CCY_Dom
        Else
            .CCY_Payout = UCase(rng_Input(1, int_ctr).Value)
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "-" Then
            .QuantoFactor = 1
        Else
            .QuantoFactor = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "-" Then
            .CSFactor = 0
        Else
            .CSFactor = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        .CCY_PnL = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .IsSmile = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsRescaling = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_Disc = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_SpotDisc = rng_Input(1, int_ctr).Value
    End With

    With fld_Premium
        .TradeID = str_TradeID

        int_ctr = int_ctr + 1
        .Amount = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .PmtDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_Disc = rng_Input(1, int_ctr).Value
        .Curve_SpotDisc = .Curve_Disc
    End With

    fld_Output.Premium = fld_Premium
    GetInstParams_FVN = fld_Output
End Function

Public Function GetInstParams_BND(rng_Input As Range, dic_StaticInfo As Dictionary, Optional str_TradeID As String = "-") As InstParams_BND
    ' ## Input is a horizontal range containing the parameter values
    Dim fld_Output As InstParams_BND, fld_PurchaseCost As SCFParams
    Dim str_Custom As String, wks_Custom As Worksheet
    Dim dcp_ActiveParams As DictParams
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    Dim int_ctr As Integer: int_ctr = 0

    With fld_Output
        .TradeID = str_TradeID
        .IsUniformPeriods = False

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "<SYSTEM>" Then
            .ValueDate = cfg_Settings.CurrentValDate
        Else
            .ValueDate = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        .SpotDays = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .StubType = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsLongCpn = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .TradeDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .FirstAccDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .MatDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .RollDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsFwdGeneration = (UCase(rng_Input(1, int_ctr).Value) = "FORWARD")

        int_ctr = int_ctr + 1
        .IsBuy = (UCase(rng_Input(1, int_ctr).Value) = "B")

        int_ctr = int_ctr + 1
        .Principal = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY_Principal = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .index = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .RateOrMargin = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsRoundFlow = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .PmtFreq = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Daycount = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .BDC = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .PaymentSchType = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .EOM = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .EOM2830 = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY_PnL = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .PmtCal = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .estcal = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_Disc = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_Est = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_SpotDisc = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .DiscSpread = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CalType = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .BondMarketPrice = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .YieldGenerator = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        str_Custom = rng_Input(1, int_ctr).Value
        Select Case str_Custom
            Case "-", ""
            Case Else
                Set wks_Custom = ThisWorkbook.Worksheets(str_Custom)
        End Select

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "-" Then
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set .Fixings = Gather_Dictionary(dcp_ActiveParams, False)
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "-" Then
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set .ModStarts = Gather_Dictionary(dcp_ActiveParams)
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "-" Then
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set .AmortSchedule = Gather_Dictionary(dcp_ActiveParams, False)
        End If


    End With

    With fld_PurchaseCost
        .TradeID = str_TradeID

        int_ctr = int_ctr + 1
        .Amount = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .PmtDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_Disc = rng_Input(1, int_ctr).Value
        .Curve_SpotDisc = .Curve_Disc
    End With

    fld_Output.PurchaseCost = fld_PurchaseCost
    GetInstParams_BND = fld_Output
End Function
Public Function GetInstParams_BA(rng_Input As Range, dic_StaticInfo As Dictionary, Optional str_TradeID As String = "-") As InstParams_BND
    Dim str_NotUsed As String: str_NotUsed = "-"
    ' ## Input is a horizontal range containing the parameter values
    'function added by QJK 30092014
    Dim fld_Output As InstParams_BND, fld_PurchaseCost As SCFParams
    Dim str_Custom As String, wks_Custom As Worksheet
    Dim dcp_ActiveParams As DictParams
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    Dim int_ctr As Integer: int_ctr = 0

    With fld_Output
        .TradeID = str_TradeID
        .IsUniformPeriods = False

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "<SYSTEM>" Then
            .ValueDate = cfg_Settings.CurrentValDate
        Else
            .ValueDate = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        .SpotDays = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .StubType = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsLongCpn = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .FirstAccDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .MatDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .RollDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsFwdGeneration = (UCase(rng_Input(1, int_ctr).Value) = "FORWARD")

        int_ctr = int_ctr + 1
        .IsBuy = (UCase(rng_Input(1, int_ctr).Value) = "B")

        int_ctr = int_ctr + 1
        .Principal = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY_Principal = UCase(rng_Input(1, int_ctr).Value)

        'int_Ctr = int_Ctr + 1
        .index = str_NotUsed 'rng_Input(1, int_Ctr).Value

        'int_Ctr = int_Ctr + 1
        .RateOrMargin = 0 'rng_Input(1, int_Ctr).Value

        int_ctr = int_ctr + 1
        .IsRoundFlow = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .PmtFreq = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Daycount = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .BDC = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .PaymentSchType = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .EOM = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .EOM2830 = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY_PnL = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .PmtCal = rng_Input(1, int_ctr).Value

       ' int_Ctr = int_Ctr + 1
        .estcal = str_NotUsed 'rng_Input(1, int_Ctr).Value

        int_ctr = int_ctr + 1
        .Curve_Disc = rng_Input(1, int_ctr).Value

        'int_Ctr = int_Ctr + 1
        .Curve_Est = str_NotUsed 'rng_Input(1, int_Ctr).Value

        int_ctr = int_ctr + 1
        .Curve_SpotDisc = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .DiscSpread = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CalType = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .BondMarketPrice = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .YieldGenerator = rng_Input(1, int_ctr).Value


'not applicable, all deleted
        'int_Ctr = int_Ctr + 1
        'str_Custom = "-" 'rng_Input(1, int_Ctr).Value
       ' Select Case str_Custom
       '     Case "-", ""
        '    Case Else
        '        Set wks_Custom = ThisWorkbook.Worksheets(str_Custom)
       ' End Select

        'int_Ctr = int_Ctr + 1
        'If rng_Input(1, int_Ctr).Value <> "-" Then
        '    Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_Ctr).Value)
        '    Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
        '    Set .Fixings = Gather_Dictionary(dcp_ActiveParams, False)
       ' End If

        'int_Ctr = int_Ctr + 1
        'If rng_Input(1, int_Ctr).Value <> "-" Then
        '    Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_Ctr).Value)
        '    Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
        '    Set .ModStarts = Gather_Dictionary(dcp_ActiveParams)
       ' End If

        'int_Ctr = int_Ctr + 1
        'If rng_Input(1, int_Ctr).Value <> "-" Then
        '    Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_Ctr).Value)
        '    Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
        '    Set .AmortSchedule = Gather_Dictionary(dcp_ActiveParams, False)
        'End If


    End With

    With fld_PurchaseCost
        .TradeID = str_TradeID

        int_ctr = int_ctr + 1
        .Amount = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .PmtDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_Disc = rng_Input(1, int_ctr).Value
        .Curve_SpotDisc = .Curve_Disc
    End With

    fld_Output.PurchaseCost = fld_PurchaseCost
    GetInstParams_BA = fld_Output
End Function
Public Function GetInstParams_NID(rng_Input As Range, dic_StaticInfo As Dictionary, Optional str_TradeID As String = "-") As InstParams_BND
    ' ## Input is a horizontal range containing the parameter values
    Dim fld_Output As InstParams_BND, fld_PurchaseCost As SCFParams
    Dim str_Custom As String, wks_Custom As Worksheet
    Dim dcp_ActiveParams As DictParams
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    Dim int_ctr As Integer: int_ctr = 0

    With fld_Output
        .TradeID = str_TradeID
        .IsUniformPeriods = False

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "<SYSTEM>" Then
            .ValueDate = cfg_Settings.CurrentValDate
        Else
            .ValueDate = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        .SpotDays = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .StubType = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsLongCpn = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .TradeDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .FirstAccDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .MatDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .RollDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsFwdGeneration = (UCase(rng_Input(1, int_ctr).Value) = "FORWARD")

        int_ctr = int_ctr + 1
        .IsBuy = (UCase(rng_Input(1, int_ctr).Value) = "B")

        int_ctr = int_ctr + 1
        .Principal = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY_Principal = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .index = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .RateOrMargin = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsRoundFlow = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .PmtFreq = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Daycount = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .BDC = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .PaymentSchType = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .EOM = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .EOM2830 = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY_PnL = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .PmtCal = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .estcal = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_Disc = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_Est = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_SpotDisc = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .DiscSpread = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CalType = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .BondMarketPrice = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .YieldGenerator = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        str_Custom = rng_Input(1, int_ctr).Value
        Select Case str_Custom
            Case "-", ""
            Case Else
                Set wks_Custom = ThisWorkbook.Worksheets(str_Custom)
        End Select

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "-" Then
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set .Fixings = Gather_Dictionary(dcp_ActiveParams, False)
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "-" Then
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set .ModStarts = Gather_Dictionary(dcp_ActiveParams)
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "-" Then
            Set dcp_ActiveParams.KeysTopLeft = wks_Custom.Range(rng_Input(1, int_ctr).Value)
            Set dcp_ActiveParams.ValuesTopLeft = dcp_ActiveParams.KeysTopLeft.Offset(0, 1)
            Set .AmortSchedule = Gather_Dictionary(dcp_ActiveParams, False)
        End If


    End With

    With fld_PurchaseCost
        .TradeID = str_TradeID

        int_ctr = int_ctr + 1
        .Amount = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .PmtDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_Disc = rng_Input(1, int_ctr).Value
        .Curve_SpotDisc = .Curve_Disc
    End With

    fld_Output.PurchaseCost = fld_PurchaseCost
    GetInstParams_NID = fld_Output
End Function

Public Function GetInstParams_FBR(rng_Input As Range, dic_StaticInfo As Dictionary, Optional str_TradeID As String = "-") As InstParams_FBR
    ' ## Input is a horizontal range containing the parameter values
    Dim fld_Output As InstParams_FBR, fld_Premium As SCFParams
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    Dim int_ctr As Integer: int_ctr = 0

    With fld_Output
        .TradeID = str_TradeID

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "<SYSTEM>" Then
            .ValueDate = cfg_Settings.CurrentValDate
        Else
            .ValueDate = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        .MatDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .DelivDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsBuy = (UCase(rng_Input(1, int_ctr).Value) = "B")

        int_ctr = int_ctr + 1
        .CCY_Fgn = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .CCY_Dom = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .Notional_Fgn = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        Select Case UCase(rng_Input(1, int_ctr).Value)
            Case "C", "CALL": .OptDirection = OptionDirection.CallOpt
            Case "P", "PUT": .OptDirection = OptionDirection.PutOpt
        End Select

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "-" Then
            .strike = -1
        Else
            .strike = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        .IsKnockOut = (UCase(rng_Input(1, int_ctr).Value) = "OUT")

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "-" Then
            .LowerBar = -1
        Else
            .LowerBar = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "-" Then
            .UpperBar = -1
        Else
            .UpperBar = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "-" Then
            .WindowStart = .ValueDate
        Else
            .WindowStart = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "-" Then
            .WindowEnd = .MatDate
        Else
            .WindowEnd = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "-" Then
            .CCY_Payout = .CCY_Dom
        Else
            .CCY_Payout = UCase(rng_Input(1, int_ctr).Value)
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "-" Then
            .QuantoFactor = 1
        Else
            .QuantoFactor = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        .CCY_PnL = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .IsSmile_Orig = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsSmile_IfKnocked = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsRescaling = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_Disc = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_SpotDisc = rng_Input(1, int_ctr).Value
    End With

    With fld_Premium
        .TradeID = str_TradeID

        int_ctr = int_ctr + 1
        .Amount = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .PmtDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_Disc = rng_Input(1, int_ctr).Value
        .Curve_SpotDisc = .Curve_Disc
    End With

    fld_Output.Premium = fld_Premium
    GetInstParams_FBR = fld_Output
End Function

Public Function GetInstParams_FTB(rng_Input As Range, dic_StaticInfo As Dictionary, Optional str_TradeID As String = "-") As InstParams_FTB
    Dim fld_Output As InstParams_FTB
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    Dim int_ctr As Integer

    With fld_Output
        .TradeID = str_TradeID

        int_ctr = 1
        If rng_Input(1, int_ctr).Value = "<SYSTEM>" Then
            .ValueDate = cfg_Settings.CurrentValDate
        Else
            .ValueDate = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        .FutMat = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsBuy = (UCase(rng_Input(1, int_ctr).Value) = "B")

        int_ctr = int_ctr + 1
        .Notional = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY_Notional = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .UndTerm = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Daycount = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "-" Then
            .Price_Mkt = 0  ' Placeholder value
        Else
            .Price_Mkt = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "-" Then
            .AdjUndMatDate = .FutMat  ' Placeholder value
        Else
            .AdjUndMatDate = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        .SettleDays = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsSpreadOn_PnL = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsSpreadOn_DV01 = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .MeasureMode = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .CCY_PnL = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .PmtCal = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_Est = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Price_Orig = rng_Input(1, int_ctr).Value
    End With

    GetInstParams_FTB = fld_Output
End Function

Public Function GetInstParams_FBN(rng_Input As Range, dic_StaticInfo As Dictionary, Optional str_TradeID As String = "-") As InstParams_FBN
    Dim fld_Output As InstParams_FBN
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    Dim int_ctr As Integer

    With fld_Output
        .TradeID = str_TradeID
        .IsUniformPeriods = True

        int_ctr = 1
        If rng_Input(1, int_ctr).Value = "<SYSTEM>" Then
            .ValueDate = cfg_Settings.CurrentValDate
        Else
            .ValueDate = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        .FutMat = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .UndMat = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .SettleDays = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsBuy = (UCase(rng_Input(1, int_ctr).Value) = "B")

        int_ctr = int_ctr + 1
        .Notional = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Generator = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Coupon = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .BDC_Accrual = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Price_Mkt = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "-" Then
            .ConvFac = 100
        Else
            .ConvFac = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        .PriceType = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .IsSpreadOn_PnL = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsSpreadOn_DV01 = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY_PnL = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .Price_Orig = rng_Input(1, int_ctr).Value
    End With

    GetInstParams_FBN = fld_Output
End Function

Public Function GetInstParams_FRE(rng_Input As Range, dic_StaticInfo As Dictionary, Optional str_TradeID As String = "-") As InstParams_FRE
    ' ## Input is a horizontal range containing the parameter values
    Dim fld_Output As InstParams_FRE, fld_Premium As SCFParams
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    Dim int_ctr As Integer: int_ctr = 0

    With fld_Output
        .TradeID = str_TradeID

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "<SYSTEM>" Then
            .ValueDate = cfg_Settings.CurrentValDate
        Else
            .ValueDate = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        .MatDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .DelivDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsBuy = (UCase(rng_Input(1, int_ctr).Value) = "B")

        int_ctr = int_ctr + 1
        .CCY_Fgn = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .CCY_Dom = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .RebateAmt = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY_Rebate = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .IsKnockOut = (UCase(rng_Input(1, int_ctr).Value) = "OUT")

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "-" Then
            .LowerBar = -1
        Else
            .LowerBar = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "-" Then
            .UpperBar = -1
        Else
            .UpperBar = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "-" Then
            .WindowStart = .ValueDate
        Else
            .WindowStart = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "-" Then
            .WindowEnd = .MatDate
        Else
            .WindowEnd = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        .IsInstantRebate = (UCase(rng_Input(1, int_ctr).Value) = "KNOCK TIME")

        int_ctr = int_ctr + 1
        .CCY_PnL = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .IsSmile = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_Disc = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_SpotDisc = rng_Input(1, int_ctr).Value
    End With

    With fld_Premium
        .TradeID = str_TradeID

        int_ctr = int_ctr + 1
        .Amount = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .PmtDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_Disc = rng_Input(1, int_ctr).Value
        .Curve_SpotDisc = .Curve_Disc
    End With

    fld_Output.Premium = fld_Premium
    GetInstParams_FRE = fld_Output
End Function

Public Function GetInstParams_ECS(rng_Input As Range, dic_StaticInfo As Dictionary, Optional str_TradeID As String = "-") As InstParams_ECS
    ' ## Input is a horizontal range containing the parameter values
    Dim fld_Output As InstParams_ECS, fld_PurchaseCost As SCFParams
    Dim str_Custom As String, wks_Custom As Worksheet
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    Dim int_ctr As Integer: int_ctr = 0

    With fld_Output
        .TradeID = str_TradeID

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "<SYSTEM>" Then
            .ValueDate = cfg_Settings.CurrentValDate
        Else
            .ValueDate = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        .SpotDays = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsBuy = (UCase(rng_Input(1, int_ctr).Value) = "B")

        int_ctr = int_ctr + 1
        .Quantity = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Security = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY_Sec = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY_PnL = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .SpotCal = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_SpotDisc = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .OutputType = rng_Input(1, int_ctr).Value

        With fld_PurchaseCost
            .TradeID = str_TradeID

            int_ctr = int_ctr + 1
            .Amount = rng_Input(1, int_ctr).Value

            int_ctr = int_ctr + 1
            .CCY = UCase(rng_Input(1, int_ctr).Value)

            int_ctr = int_ctr + 1
            .PmtDate = rng_Input(1, int_ctr).Value

            int_ctr = int_ctr + 1
            .Curve_Disc = rng_Input(1, int_ctr).Value
            .Curve_SpotDisc = .Curve_Disc
        End With

    End With
    fld_Output.PurchaseCost = fld_PurchaseCost

    GetInstParams_ECS = fld_Output

End Function

Public Function GetInstParams_EQO(rng_Input As Range, dic_StaticInfo As Dictionary, Optional str_TradeID As String = "-") As InstParams_EQO
    ' ## Input is a horizontal range containing the parameter values
    Dim fld_Output As InstParams_EQO, fld_PurchaseCost As SCFParams
    Dim str_Custom As String, wks_Custom As Worksheet
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    Dim int_ctr As Integer: int_ctr = 0

    ' Listing the parameters for EQO in the sheets
    With fld_Output
        .TradeID = str_TradeID

        int_ctr = int_ctr + 1 'new
        If rng_Input(1, int_ctr).Value = "<SYSTEM>" Then
            .ValueDate = cfg_Settings.CurrentValDate
            .OriValueDate = cfg_Settings.OrigValDate
        Else
            .ValueDate = rng_Input(1, int_ctr).Value
            .OriValueDate = cfg_Settings.OrigValDate
        End If

        .Description = rng_Input(1, int_ctr - 3).Value

        int_ctr = int_ctr + 1
        .SpotDays = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .MatDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .DelivDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsFutures = (UCase(rng_Input(1, int_ctr).Value) = "YES") 'new

        int_ctr = int_ctr + 1
        .FuturesContract = UCase(rng_Input(1, int_ctr).Value) 'new

        int_ctr = int_ctr + 1
        .FutMat_Date = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .ExerciseType = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .IsCall = (UCase(rng_Input(1, int_ctr).Value) = "CALL")

        int_ctr = int_ctr + 1
        .IsBuy = (UCase(rng_Input(1, int_ctr).Value) = "B")

        int_ctr = int_ctr + 1
        .LotSize = rng_Input(1, int_ctr).Value 'new

        int_ctr = int_ctr + 1
        .Quantity = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .strike = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Parity = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Security = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .VolCode = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .CCY_Sec = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY_PnL = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .SpotCal = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_SpotDisc = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_MarketDisc = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Settlement = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .SettlementDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .OptionSpot = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Div_Amount = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .DivPayment_Date = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .DivEx_Date = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Div_Type = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .DividendSpot = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .VolSpread = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .PLType = rng_Input(1, int_ctr).Value

        With fld_PurchaseCost
            .TradeID = str_TradeID

            int_ctr = int_ctr + 1
            .Amount = rng_Input(1, int_ctr).Value

            int_ctr = int_ctr + 1
            .CCY = UCase(rng_Input(1, int_ctr).Value)

            int_ctr = int_ctr + 1
            .PmtDate = rng_Input(1, int_ctr).Value

            int_ctr = int_ctr + 1
            .Curve_Disc = rng_Input(1, int_ctr).Value
            .Curve_SpotDisc = .Curve_Disc
        End With

    End With
    fld_Output.PurchaseCost = fld_PurchaseCost

    GetInstParams_EQO = fld_Output

End Function

Public Function GetInstParams_EQF(rng_Input As Range, dic_StaticInfo As Dictionary, Optional str_TradeID As String = "-") As InstParams_EQF
    ' ## Input is a horizontal range containing the parameter values
    Dim fld_Output As InstParams_EQF
    Dim str_Custom As String, wks_Custom As Worksheet
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    Dim int_ctr As Integer: int_ctr = 0

    ' Listing the parameters for EQO in the sheets
    With fld_Output
        .TradeID = str_TradeID

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "<SYSTEM>" Then
            .ValueDate = cfg_Settings.CurrentValDate
            .OriValueDate = cfg_Settings.OrigValDate
        Else
            .ValueDate = rng_Input(1, int_ctr).Value
            .OriValueDate = cfg_Settings.OrigValDate
        End If

        int_ctr = int_ctr + 1
        .SpotDays = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .MatDate = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsBuy = (UCase(rng_Input(1, int_ctr).Value) = "B")

        int_ctr = int_ctr + 1
        .Quantity = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .LotSize = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Futures = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .Fut_ContractPrice = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Fut_MktPrice = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Security = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .CCY_Sec = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY_PnL = UCase(rng_Input(1, int_ctr).Value)

        int_ctr = int_ctr + 1
        .SpotCal = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_Disc = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Curve_Div = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Div_Yield = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .PLType = rng_Input(1, int_ctr).Value

    End With

    GetInstParams_EQF = fld_Output

End Function

Public Function GetInstParams_EQS(rng_Input As Range, dic_StaticInfo As Dictionary, Optional str_TradeID As String = "-") As InstParams_EQS
    ' ## Input is a horizontal range containing the parameter values
    Dim fld_Output As InstParams_EQS
    Dim str_Custom As String, wks_Custom As Worksheet
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    Dim int_ctr As Integer: int_ctr = 0

    ' Listing the parameters for EQS in the sheets
    With fld_Output
        .TradeID = str_TradeID

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value = "<SYSTEM>" Then
            .ValueDate = cfg_Settings.CurrentValDate
        Else
            .ValueDate = rng_Input(1, int_ctr).Value
        End If

        int_ctr = int_ctr + 1
        .Swapstart = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .Term = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsCalcEqualFix = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsEqPayer = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .IsTotalRet = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .EqSpotDays = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .ConstantType = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "" And rng_Input(1, int_ctr).Value <> "-" Then
            .Quantity = rng_Input(1, int_ctr).Value
        Else
            .Quantity = 0
        End If

        int_ctr = int_ctr + 1
        If rng_Input(1, int_ctr).Value <> "" And rng_Input(1, int_ctr).Value <> "-" Then
            .Notional = rng_Input(1, int_ctr).Value
        Else
            .Notional = 0
        End If

        int_ctr = int_ctr + 1
        .CCY_Fix = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .CCY_PnL = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .PLType = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        Dim bln_WksExist As Boolean: bln_WksExist = False
        If Examine_WorksheetExists(ThisWorkbook, rng_Input(1, int_ctr).Value) = True Then
            Set .CustomSheet = ThisWorkbook.Worksheets(rng_Input(1, int_ctr).Value)
            bln_WksExist = True
        End If


        Dim int_DateCnt As Integer: int_DateCnt = 0

        int_ctr = int_ctr + 1
        If bln_WksExist = True And rng_Input(1, int_ctr).Value <> "" Then
            Dim dic_CustomDate As New Dictionary

            Dim rng_target As Range
            Set rng_target = .CustomSheet.Range(rng_Input(1, int_ctr).Value)

            While rng_target.Value <> ""
                int_DateCnt = int_DateCnt + 1

                dic_CustomDate.Add "ID|" & int_DateCnt, rng_target.Value
                dic_CustomDate.Add "NUM|" & int_DateCnt, rng_target.Offset(0, 1).Value
                dic_CustomDate.Add "TYPE|" & int_DateCnt, rng_target.Offset(0, 2).Value
                dic_CustomDate.Add "DATE|" & int_DateCnt, rng_target.Offset(0, 3).Value

                Set rng_target = rng_target.Offset(1, 0)
            Wend

            Set .CustomDate = dic_CustomDate

        End If

        int_ctr = int_ctr + 1
        .Security = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        Dim irg_LegA As IRGeneratorSet
        Set irg_LegA = GetObject_IRGeneratorSet()

        Dim irg_LegA_Det As IRLegParams
        irg_LegA_Det = irg_LegA.Lookup_Generator(rng_Input(1, int_ctr).Value)

        .CCY_Eq = irg_LegA_Det.CCY
        .EqFreq = irg_LegA_Det.PmtFreq
        .EqFixCal = irg_LegA_Det.estcal
        .EqPmtCal = irg_LegA_Det.PmtCal
        .EqEstCurve = irg_LegA_Det.Curve_Est
        .EqDiscCurve = irg_LegA_Det.Curve_Disc
        .EqBDC = irg_LegA_Det.BDC
        .EqIsEOM = irg_LegA_Det.EOM


        int_ctr = int_ctr + 1
        .EqDivType = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1

        Dim int_DivCnt As Integer: int_DivCnt = 0

        If .EqDivType = "Cash" Then
            If bln_WksExist = True Then
                Dim dic_Div As New Dictionary

                Dim rng_Div As Range
                Set rng_Div = .CustomSheet.Range(rng_Input(1, int_ctr).Value)

                int_DivCnt = int_DivCnt + 1
                dic_Div.Add "EX|" & int_DivCnt, rng_Div.Value
                dic_Div.Add "PMT|" & int_DivCnt, rng_Div.Offset(0, 1).Value
                dic_Div.Add "DIV|" & int_DivCnt, rng_Div.Offset(0, 2).Value
                dic_Div.Add "SHIFT|" & int_DivCnt, rng_Div.Offset(0, 3).Value
                Set rng_Div = rng_Div.Offset(1, 0)

                dic_Div.Add "COUNT", dic_Div.count

                While rng_Div.Value <> ""
                    int_DivCnt = int_DivCnt + 1
                    dic_Div.Add "EX|" & int_DivCnt, rng_Div.Value
                    dic_Div.Add "PMT|" & int_DivCnt, rng_Div.Offset(0, 1).Value
                    dic_Div.Add "DIV|" & int_DivCnt, rng_Div.Offset(0, 2).Value
                    dic_Div.Add "SHIFT|" & int_DivCnt, rng_Div.Offset(0, 3).Value
                    Set rng_Div = rng_Div.Offset(1, 0)
                Wend

                Set .EqDiv = dic_Div
            End If
        End If

        int_ctr = int_ctr + 1
        If bln_WksExist = True Then
            If rng_Input(1, int_ctr).Value <> "" And rng_Input(1, int_ctr).Value <> "-" Then
                Dim dic_EqFix As New Dictionary
                Dim rng_EqFix As Range
                Set rng_EqFix = .CustomSheet.Range(rng_Input(1, int_ctr).Value)
                dic_EqFix.Add rng_EqFix.Value, rng_EqFix.Offset(0, 1).Value
                Set rng_EqFix = rng_EqFix.Offset(1, 0)

                While rng_EqFix.Value <> ""
                    dic_EqFix.Add rng_EqFix.Value, rng_EqFix.Offset(0, 1).Value
                    Set rng_EqFix = rng_EqFix.Offset(1, 0)
                Wend

                Set .EqFixing = dic_EqFix
            End If
        End If

        int_ctr = int_ctr + 1
        If bln_WksExist = True Then
            If rng_Input(1, int_ctr).Value <> "" And rng_Input(1, int_ctr).Value <> "-" Then
                Dim dic_FxFix As New Dictionary
                Dim rng_FxFix As Range
                Set rng_FxFix = .CustomSheet.Range(rng_Input(1, int_ctr).Value)
                dic_FxFix.Add rng_FxFix.Value, rng_FxFix.Offset(0, 1).Value
                Set rng_FxFix = rng_FxFix.Offset(1, 0)

                While rng_FxFix.Value <> ""
                    dic_FxFix.Add rng_FxFix.Value, rng_FxFix.Offset(0, 1).Value
                    Set rng_FxFix = rng_FxFix.Offset(1, 0)
                Wend

                Set .FxFixing = dic_FxFix
            End If
        End If

        int_ctr = int_ctr + 1
        Dim irg_LegB As IRGeneratorSet
        Set irg_LegB = GetObject_IRGeneratorSet()

        Dim irg_LegB_Det As IRLegParams
        irg_LegB_Det = irg_LegB.Lookup_Generator(rng_Input(1, int_ctr).Value)

        .CCY_Id = irg_LegB_Det.CCY
        .IdFreq = irg_LegB_Det.PmtFreq
        .IdFixCal = irg_LegB_Det.estcal
        .IdPmtCal = irg_LegB_Det.PmtCal

        If irg_LegB_Det.index = "-" Then
            .IdIsFix = True
        Else
            .IdIsFix = False
        End If

        .IdEstCurve = irg_LegB_Det.Curve_Est
        .IdDiscCurve = irg_LegB_Det.Curve_Disc
        .IdDayCnt = irg_LegB_Det.Daycount
        .IdBDC = irg_LegB_Det.BDC
        .IdIsEOM = irg_LegB_Det.EOM

        int_ctr = int_ctr + 1
        .IdSpotDays = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        .RateOrMargin = rng_Input(1, int_ctr).Value

        int_ctr = int_ctr + 1
        If bln_WksExist = True Then
            If rng_Input(1, int_ctr).Value <> "" And rng_Input(1, int_ctr).Value <> "-" Then
                Dim dic_IdFix As New Dictionary
                Dim rng_IdFix As Range
                Set rng_IdFix = .CustomSheet.Range(rng_Input(1, int_ctr).Value)
                dic_IdFix.Add rng_IdFix.Value, rng_IdFix.Offset(0, 1).Value
                Set rng_IdFix = rng_IdFix.Offset(1, 0)

                While rng_IdFix.Value <> ""
                    dic_IdFix.Add rng_IdFix.Value, rng_IdFix.Offset(0, 1).Value
                    Set rng_IdFix = rng_IdFix.Offset(1, 0)
                Wend

                Set .IdFixing = dic_IdFix
            End If
        End If


    End With

    GetInstParams_EQS = fld_Output

End Function
Public Function GetInstParams_FXFut(rng_Input As Range, dic_StaticInfo As Dictionary, Optional str_TradeID As String = "-") As InstParams_FXFut
    Dim fld_Output As InstParams_FXFut
    Dim fld_FlowA As SCFParams, fld_FlowB As SCFParams
    Dim cfg_Settings As ConfigSheet: Set cfg_Settings = dic_StaticInfo(StaticInfoType.ConfigSheet)
    Dim int_ctr As Integer: int_ctr = 0

    ' Shared
    fld_FlowA.TradeID = str_TradeID
    fld_FlowB.TradeID = str_TradeID

    int_ctr = int_ctr + 1
    If rng_Input(1, int_ctr).Value = "<SYSTEM>" Then
        fld_Output.ValueDate = cfg_Settings.CurrentValDate
    Else
        fld_Output.ValueDate = rng_Input(1, int_ctr).Value
    End If

    int_ctr = int_ctr + 1
    fld_Output.SpotDays = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.SpotCal = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.MatDate = rng_Input(1, int_ctr).Value
    fld_FlowA.PmtDate = rng_Input(1, int_ctr).Value
    fld_FlowB.PmtDate = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.PayCal = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.IsBuy = (UCase(rng_Input(1, int_ctr).Value) = "B")

    int_ctr = int_ctr + 1
    fld_Output.Quantity = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.LotSize = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.LotSizeCCY = UCase(rng_Input(1, int_ctr).Value)

    int_ctr = int_ctr + 1
    fld_Output.Futures = UCase(rng_Input(1, int_ctr).Value)

    int_ctr = int_ctr + 1
    fld_Output.Fut_ContractPrice = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.Fut_MktPrice = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.Underlying = rng_Input(1, int_ctr).Value

    int_ctr = int_ctr + 1
    fld_Output.CCY_PnL = UCase(rng_Input(1, int_ctr).Value)

    'FlowA
    int_ctr = int_ctr + 1
    fld_FlowA.CCY = UCase(rng_Input(1, int_ctr).Value)

    int_ctr = int_ctr + 1
    fld_FlowA.Curve_Disc = rng_Input(1, int_ctr).Value
    fld_FlowA.Curve_SpotDisc = rng_Input(1, int_ctr).Value
    fld_FlowA.Amount = fld_Output.Quantity * fld_Output.LotSize

    'FlowB
    int_ctr = int_ctr + 1
    fld_FlowB.CCY = UCase(rng_Input(1, int_ctr).Value)

    int_ctr = int_ctr + 1
    fld_FlowB.Curve_Disc = rng_Input(1, int_ctr).Value
    fld_FlowB.Curve_SpotDisc = rng_Input(1, int_ctr).Value
    fld_FlowB.Amount = fld_Output.Quantity * fld_Output.LotSize

    ' Prepare final output
    fld_Output.TradeID = str_TradeID
    fld_Output.FlowA = fld_FlowA
    fld_Output.FlowB = fld_FlowB

    GetInstParams_FXFut = fld_Output
End Function