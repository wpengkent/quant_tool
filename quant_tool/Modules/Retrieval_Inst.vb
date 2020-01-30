Option Explicit

' ## GENERAL
Public Function GetInstType_String(enu_InstType As InstType) As String
    Dim str_Output As String
    Select Case enu_InstType
        Case InstType.All: str_Output = "<ALL>"
        Case InstType.IRS: str_Output = "IRS"
        Case InstType.CFL: str_Output = "CFL"
        Case InstType.SWT: str_Output = "SWT"
        Case InstType.FXF: str_Output = "FXF"
        Case InstType.DEP: str_Output = "DEP"
        Case InstType.FRA: str_Output = "FRA"
        Case InstType.FVN: str_Output = "FVN"
        Case InstType.BND: str_Output = "BND"
        Case InstType.FBR: str_Output = "FBR"
        Case InstType.FTB: str_Output = "FTB"
        Case InstType.FBN: str_Output = "FBN"
        Case InstType.FRE: str_Output = "FRE"
        Case InstType.ECS: str_Output = "ECS"
        Case InstType.EQO: str_Output = "EQO"
        Case InstType.EQF: str_Output = "EQF"
        Case InstType.EQS: str_Output = "EQS"
        Case InstType.BA: str_Output = "BA"
        Case InstType.NID: str_Output = "NID"
        Case InstType.FXFut: str_Output = "FXFut"
        Case InstType.RngAcc: str_Output = "RngAcc"
    End Select

    GetInstType_String = str_Output
End Function