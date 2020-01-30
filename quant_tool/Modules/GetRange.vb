Option Explicit

Public Function GetRange_ActiveContainer() As Range
    Set GetRange_ActiveContainer = GetSheet_Config().Range("B3")
End Function

Public Function GetRange_CurveSetup(enu_Type As CurveType) As Range
    Dim rng_Output As Range
    With ThisWorkbook.Worksheets("Setup - " & GetCurveTypeName(enu_Type))
        Select Case enu_Type
            Case CurveType.FXV: Set rng_Output = .Range("A4:R4")
            Case CurveType.IRC: Set rng_Output = .Range("A4:S4")
            Case CurveType.cvl: Set rng_Output = .Range("A4:L4")
            Case CurveType.SVL: Set rng_Output = .Range("A4:L4")
            Case CurveType.EVL: Set rng_Output = .Range("A4:D4")
            Case Else: Debug.Assert False
        End Select
    End With

    Set GetRange_CurveSetup = rng_Output
End Function

Public Function GetRange_CurveParams(enu_Type As CurveType, str_Code As String) As Range
    Dim rng_Output As Range
    Dim wks_Location As Worksheet, rng_CodeRef As Range
    Dim int_Row As Integer

    ' Find relevant sheet
    Set wks_Location = ThisWorkbook.Worksheets("Setup - " & GetCurveTypeName(enu_Type))

    ' Find row for code
    Set rng_CodeRef = Gather_RangeBelow(wks_Location.Range("A4"))
    int_Row = Examine_FindIndex(Convert_RangeToList(rng_CodeRef), str_Code)
    Debug.Assert int_Row <> -1

    ' Return params for code
    Select Case enu_Type
        Case CurveType.FXV: Set rng_Output = wks_Location.Range("C3:R3").Offset(int_Row, 0)
        Case CurveType.IRC: Set rng_Output = wks_Location.Range("C3:S3").Offset(int_Row, 0)
        'Case CurveType.CVL: Set rng_Output = wks_Location.Range("C3:L3").Offset(int_Row, 0)
        Case CurveType.cvl: Set rng_Output = wks_Location.Range("C3:N3").Offset(int_Row, 0)  'QJK code 16/12/2014
        Case CurveType.SVL: Set rng_Output = wks_Location.Range("C3:L3").Offset(int_Row, 0)
        Case CurveType.EVL: Set rng_Output = wks_Location.Range("C3:D3").Offset(int_Row, 0)
        Case Else: Debug.Assert False
    End Select

    Set GetRange_CurveParams = rng_Output
End Function