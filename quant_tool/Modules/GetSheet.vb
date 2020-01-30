Option Explicit

Public Function GetSheet_Config() As Worksheet
    Set GetSheet_Config = ThisWorkbook.Worksheets("General Settings")
End Function

Public Function GetSheet_Holidays() As Worksheet
    Set GetSheet_Holidays = ThisWorkbook.Worksheets("Holidays")
End Function

Public Function GetSheet_Mapping() As Worksheet
    Set GetSheet_Mapping = ThisWorkbook.Worksheets("Mapping Rules")
End Function

Public Function GetSheet_RiskEngine() As Worksheet
    Set GetSheet_RiskEngine = ThisWorkbook.Worksheets("RiskEngine")
End Function

Public Function GetSheet_Comparison() As Worksheet
    Set GetSheet_Comparison = ThisWorkbook.Worksheets("Comparison_VaRStress")
End Function

Public Function GetSheet_Results_MX() As Worksheet
    Set GetSheet_Results_MX = ThisWorkbook.Worksheets("Results_MX_VaRStress")
End Function

Public Function GetSheet_Results_VaRStress() As Worksheet
    Set GetSheet_Results_VaRStress = ThisWorkbook.Worksheets("Results_VaRStress")
End Function

Public Function GetSheet_Results_DV01() As Worksheet
    Set GetSheet_Results_DV01 = ThisWorkbook.Worksheets("Results_DV01")
End Function

Public Function GetSheet_Results_DV02() As Worksheet
    Set GetSheet_Results_DV02 = ThisWorkbook.Worksheets("Results_DV02")
End Function

Public Function GetSheet_Results_Vega() As Worksheet
    Set GetSheet_Results_Vega = ThisWorkbook.Worksheets("Results_Vega")
End Function

Public Function GetSheet_Generators_IR() As Worksheet
    Set GetSheet_Generators_IR = ThisWorkbook.Worksheets("IR Generators")
End Function

Public Function GetSheet_DateShifters() As Worksheet
    Set GetSheet_DateShifters = ThisWorkbook.Worksheets("DateShifters")
End Function

Public Function GetSheet_IRQueries() As Worksheet
    Set GetSheet_IRQueries = ThisWorkbook.Worksheets("IRQueries")
End Function

Public Function GetSheet_ScenDownload() As Worksheet
    Set GetSheet_ScenDownload = ThisWorkbook.Worksheets("CONT_DL~")
End Function

Public Function GetSheet_Generators_Yield() As Worksheet
    Set GetSheet_Generators_Yield = ThisWorkbook.Worksheets("Yield Generators")
End Function