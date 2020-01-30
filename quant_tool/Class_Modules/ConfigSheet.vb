    Option Explicit

Private Enum SettingName
    RatesDB = 1
    ScenDB
    CurrentContainer
    CurrentScenario
    BuildDate_Curr
    DataDate_Curr
    ValDate_Curr
    BuildDate_Orig
    DataDate_Orig
    ValDate_Orig
    EqRatesDB
End Enum

' ## MEMBER DATA
Private wks_Location As Worksheet
Private dic_Settings As Dictionary
Private str_Path_RatesDB As String, str_Path_ScenDB As String
Private rng_CurrentCont As Range, rng_CurrentScen As Range
Private rng_Date_OrigBuild As Range, rng_Date_OrigData As Range, rng_Date_OrigVal As Range
Private rng_Date_CurrentBuild As Range, rng_Date_CurrentData As Range, rng_Date_CurrentVal As Range
Private rng_TopLeft_Sheets As Range


' ## INITIALIZATION
Public Sub Initialize(wks_LocationInput)
    Set wks_Location = wks_LocationInput
    Set dic_Settings = New Dictionary

    Dim rng_TopLeft As Range: Set rng_TopLeft = wks_Location.Range("A1")
    Dim int_RowOffset As Integer, int_ColOffset As Integer

    With wks_Location.Range("A1")
        int_RowOffset = 0
        int_ColOffset = 1
        Call dic_Settings.Add(SettingName.RatesDB, .Offset(int_RowOffset, int_ColOffset).Value)

        int_RowOffset = int_RowOffset + 1
        Call dic_Settings.Add(SettingName.ScenDB, .Offset(int_RowOffset, int_ColOffset).Value)

        int_RowOffset = int_RowOffset + 1
        Set rng_CurrentCont = .Offset(int_RowOffset, int_ColOffset)
        Call dic_Settings.Add(SettingName.CurrentContainer, rng_CurrentCont.Value)

        int_RowOffset = int_RowOffset + 1
        Set rng_CurrentScen = .Offset(int_RowOffset, int_ColOffset)
        Call dic_Settings.Add(SettingName.CurrentScenario, rng_CurrentScen.Value)

        int_RowOffset = int_RowOffset + 6
        Set rng_Date_OrigBuild = .Offset(int_RowOffset, int_ColOffset)
        Call dic_Settings.Add(SettingName.BuildDate_Orig, rng_Date_OrigBuild.Value)

        int_RowOffset = int_RowOffset + 1
        Set rng_Date_OrigData = .Offset(int_RowOffset, int_ColOffset)
        Call dic_Settings.Add(SettingName.DataDate_Orig, rng_Date_OrigData.Value)

        int_RowOffset = int_RowOffset + 1
        Set rng_Date_OrigVal = .Offset(int_RowOffset, int_ColOffset)
        Call dic_Settings.Add(SettingName.ValDate_Orig, rng_Date_OrigVal.Value)

        int_RowOffset = int_RowOffset - 2
        int_ColOffset = int_ColOffset + 1
        Set rng_Date_CurrentBuild = .Offset(int_RowOffset, int_ColOffset)
        Call dic_Settings.Add(SettingName.BuildDate_Curr, rng_Date_CurrentBuild.Value)

        int_RowOffset = int_RowOffset + 1
        Set rng_Date_CurrentData = .Offset(int_RowOffset, int_ColOffset)
        Call dic_Settings.Add(SettingName.DataDate_Curr, rng_Date_CurrentData.Value)

        int_RowOffset = int_RowOffset + 1
        Set rng_Date_CurrentVal = .Offset(int_RowOffset, int_ColOffset)
        Call dic_Settings.Add(SettingName.ValDate_Curr, rng_Date_CurrentVal.Value)


        int_ColOffset = 1
        int_RowOffset = int_RowOffset - 6
        Call dic_Settings.Add(SettingName.EqRatesDB, .Offset(int_RowOffset, int_ColOffset).Value)


    End With

    Set rng_TopLeft_Sheets = wks_Location.Range("G2")
End Sub


' ## PROPERTIES
Public Property Get RatesDBPath() As String
    RatesDBPath = dic_Settings(SettingName.RatesDB)
End Property
Public Property Get RatesEqDBPath() As String
    RatesEqDBPath = dic_Settings(SettingName.EqRatesDB)
End Property

Public Property Get ScenDBPath() As String
    ScenDBPath = dic_Settings(SettingName.ScenDB)
End Property

Public Property Get CurrentCont() As String
    CurrentCont = dic_Settings(SettingName.CurrentContainer)
End Property

Public Property Get CurrentScen() As Long
    CurrentScen = dic_Settings(SettingName.CurrentScenario)
End Property

Public Property Let CurrentScen(lng_Scen As Long)
    rng_CurrentScen.Value = lng_Scen
    Call dic_Settings.Remove(SettingName.CurrentScenario)
    Call dic_Settings.Add(SettingName.CurrentScenario, lng_Scen)
End Property

Public Property Get CurrentBuildDate() As Long
    CurrentBuildDate = dic_Settings(SettingName.BuildDate_Curr)
End Property

Public Property Get CurrentDataDate() As Long
    CurrentDataDate = dic_Settings(SettingName.DataDate_Curr)
End Property

Public Property Get CurrentValDate() As Long
    CurrentValDate = dic_Settings(SettingName.ValDate_Curr)
End Property

Public Property Get OrigBuildDate() As Long
    OrigBuildDate = dic_Settings(SettingName.BuildDate_Orig)
End Property

Public Property Get OrigDataDate() As Long
    OrigDataDate = dic_Settings(SettingName.DataDate_Orig)
End Property

Public Property Get OrigValDate() As Long
    OrigValDate = dic_Settings(SettingName.ValDate_Orig)
End Property

Public Property Let CurrentBuildDate(lng_date As Long)
    rng_Date_CurrentBuild.Value = lng_date
    Call dic_Settings.Remove(SettingName.BuildDate_Curr)
    Call dic_Settings.Add(SettingName.BuildDate_Curr, lng_date)
End Property

Public Property Let CurrentDataDate(lng_date As Long)
    rng_Date_CurrentData.Value = lng_date
    Call dic_Settings.Remove(SettingName.DataDate_Curr)
    Call dic_Settings.Add(SettingName.DataDate_Curr, lng_date)
End Property


' ## METHODS - CHANGE PARAMETERS / UPDATE
Public Function SetCurrentValDate(lng_date As Long) As Boolean
    ' ## Set valuation date and returns whether or not the date changed
    Dim bln_Output As Boolean
    If rng_Date_CurrentVal.Value <> lng_date Then
        bln_Output = True
        Call dic_Settings.Remove(SettingName.ValDate_Curr)
        Call dic_Settings.Add(SettingName.ValDate_Curr, lng_date)
        rng_Date_CurrentVal.Value = lng_date
    Else
        bln_Output = False
    End If


    SetCurrentValDate = bln_Output
End Function

Public Sub DisplaySheetNames()
    ' ## Display sorted list of sheet names starting from the designated cell
    Call Action_ClearBelow(rng_TopLeft_Sheets, 2)

    ' Output sheet names
    Dim strArr_Sheets() As String: strArr_Sheets = Convert_ListToStrArr(Gather_SheetNames(ThisWorkbook))
    Dim strArr_Sorted() As String: strArr_Sorted = Convert_Sort_String(strArr_Sheets)
    Call Action_OutputArray(strArr_Sorted, rng_TopLeft_Sheets)

    ' Set visibility field according to current state
    Dim rng_Active As Range: Set rng_Active = rng_TopLeft_Sheets
    While rng_Active.Value <> ""
        If ThisWorkbook.Worksheets(rng_Active.Value).Visible = True Then rng_Active.Offset(0, 1).Value = "YES"
        Set rng_Active = rng_Active.Offset(1, 0)
    Wend
End Sub

Public Sub ApplySheetVisibility()
    ' ## Read instructions set on the settings sheet and make only the selected sheets visible
    Dim str_ControlSheet As String: str_ControlSheet = wks_Location.Name
    Dim bln_ScreenUpdating As Boolean: bln_ScreenUpdating = Application.ScreenUpdating
    Dim rng_Active As Range: Set rng_Active = rng_TopLeft_Sheets
    Dim str_Active As String
    Application.ScreenUpdating = False

    While rng_Active.Value <> ""
        str_Active = rng_Active.Value
        If Examine_WorksheetExists(ThisWorkbook, str_Active) Then
            If UCase(rng_Active.Offset(0, 1).Value) = "YES" Then
                ThisWorkbook.Worksheets(str_Active).Visible = xlSheetVisible
            Else
                If rng_Active.Value <> str_ControlSheet Then
                    ThisWorkbook.Worksheets(str_Active).Visible = xlSheetHidden
                End If
            End If
        End If

        Set rng_Active = rng_Active.Offset(1, 0)
    Wend

    Application.ScreenUpdating = bln_ScreenUpdating
End Sub