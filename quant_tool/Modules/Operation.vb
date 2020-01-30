Option Explicit

Public Sub ShowMenu()
    Application.StatusBar = False
    If ThisWorkbook Is ActiveWorkbook Then frm_Menu.Show
End Sub

Public Sub RunRE_VaRStress()
    Dim eng_Risk As RiskEngine: Set eng_Risk = GetObject_RiskEngine()
    Call eng_Risk.Execute_VaRStress
    Call GotoSheet(GetSheet_Results_VaRStress().Name)
End Sub

Public Sub RunRE_DV01()
    Dim eng_Risk As RiskEngine: Set eng_Risk = GetObject_RiskEngine()
    Call eng_Risk.Execute_IRSens(RevalType.DV01)
    Call GotoSheet(GetSheet_Results_DV01().Name)
End Sub

Public Sub RunRE_DV02()
    Dim eng_Risk As RiskEngine: Set eng_Risk = GetObject_RiskEngine()
    Call eng_Risk.Execute_IRSens(RevalType.DV02)
    Call GotoSheet(GetSheet_Results_DV02().Name)
End Sub

Public Sub RunRE_Vega()
    Dim eng_Risk As RiskEngine: Set eng_Risk = GetObject_RiskEngine()
    Call eng_Risk.Execute_Vega
    Call GotoSheet(GetSheet_Results_Vega().Name)
End Sub

Public Sub DisplaySheetNames()
    Call GetObject_ConfigSheet().DisplaySheetNames
End Sub

Public Sub ApplySheetVisibility()
    Call GetObject_ConfigSheet().ApplySheetVisibility
End Sub

Public Sub GotoSheet(str_name As String)
    If Examine_WorksheetExists(ThisWorkbook, str_name) Then
        With ThisWorkbook.Worksheets(str_name)
            .Visible = xlSheetVisible
            .Activate
        End With
    End If
End Sub

Public Sub GenerateComparisonSheet()
    ' ## Populate sheet with formulas comparing Murex and excel VaR/Stress results
    Application.StatusBar = "Generating comparison..."

    Dim eng_Risk As RiskEngine: Set eng_Risk = GetObject_RiskEngine()
    eng_Risk.OutputComparison
    Call GotoSheet(GetSheet_Comparison().Name)

    Application.StatusBar = False
End Sub

Public Sub DownloadToContSheet()
    ' ## Download scenarios for the specified container to the container sheet for downloads
    Application.StatusBar = "Downloading scenarios..."

    Dim wks_Download As Worksheet: Set wks_Download = GetSheet_ScenDownload()
    Dim rng_DBContainer As Range: Set rng_DBContainer = wks_Download.Range("O1")
    Dim int_ScenMin As Integer: int_ScenMin = rng_DBContainer.Offset(1, 0).Value
    Dim int_ScenMax As Integer: int_ScenMax = rng_DBContainer.Offset(2, 0).Value
    Dim str_ContSheet As String: str_ContSheet = Right(wks_Download.Name, Len(wks_Download.Name) - 5)

    Dim sct_Container As ScenContainer: Set sct_Container = GetObject_ScenContainer(str_ContSheet)
    sct_Container.ClearScenarios
    Call sct_Container.DownloadFromDB(rng_DBContainer.Value, int_ScenMin, int_ScenMax)

    Application.StatusBar = False
End Sub
