Option Explicit
' ## FUNCTIONS TO EVALUATE A PROPERTY OF THE INPUT


Public Function Examine_Contains(arr_List As Variant, var_TargetVal As Variant) As Boolean
    ' ## Check if array contains target value
    Examine_Contains = (Examine_FindIndex(arr_List, var_TargetVal) <> -1)
End Function

Public Function Examine_FindIndex(arr_List As Variant, var_TargetVal As Variant) As Long
    ' ## Similar to Match() worksheet function, but works on collections and arrays
    ' Determine bounds, catering for several possible input types
    Dim lng_LowerBound As Long: lng_LowerBound = Examine_LowerBoundIndex(arr_List)
    Dim lng_UpperBound As Long: lng_UpperBound = Examine_UpperBoundIndex(arr_List)

    ' Find the first index which matches the target
    Dim lng_Ctr As Long
    For lng_Ctr = lng_LowerBound To lng_UpperBound
        If arr_List(lng_Ctr) = var_TargetVal Then
            Examine_FindIndex = lng_Ctr
            Exit Function
        End If
    Next lng_Ctr

    ' No match
    Examine_FindIndex = -1
End Function

Public Function Examine_LowerBoundIndex(arr_List As Variant) As Long
    ' ## Return the lower bound of the array, collection or range vector
    Dim lng_Output As Long
    Select Case TypeName(arr_List)
        Case "Collection": lng_Output = 1
        Case "Range": lng_Output = 1
        Case Else
            If IsArray(arr_List) Then lng_Output = LBound(arr_List)
    End Select

    Examine_LowerBoundIndex = lng_Output
End Function

Public Function Examine_UpperBoundIndex(arr_List As Variant) As Long
    ' ## Return the upper bound of the array, collection or range vector
    Dim lng_Output As Long
    Select Case TypeName(arr_List)
        Case "Collection": lng_Output = arr_List.Count
        Case "Range"
            If arr_List.Columns.Count = 1 Then
                lng_Output = arr_List.Rows.Count
            Else
                lng_Output = arr_List.Columns.Count
            End If
        Case Else
            If IsArray(arr_List) Then lng_Output = UBound(arr_List)
    End Select

    Examine_UpperBoundIndex = lng_Output
End Function

Public Function Examine_TermType(str_Term As String) As String
    ' ## Read only the type of the term ('Y', 'M', 'W' or 'D'), ignoring the multiple.
    ' ## For example, converts "2M" to "M"
    Examine_TermType = Right(str_Term, 1)
End Function

Public Function Examine_TermQty(str_Term As String) As Integer
    ' ## Read only the multiple of the term, ignoring the type
    ' ## For example, converts "2M" to 2
    Examine_TermQty = CInt(Left(str_Term, Len(str_Term) - 1))
End Function

Public Function Examine_DaysPerYear(str_DayCountConv As String) As Integer
    ' ## Read the number of days in a year for the specified daycount convention.  Act/Act is not supported
    Dim int_Output As Integer

    Select Case str_DayCountConv
        Case "ACT/360", "30/360", "30E/360": int_Output = 360
        Case "ACT/365": int_Output = 365
    End Select

    Examine_DaysPerYear = int_Output
End Function

Public Function Examine_WorksheetExists(wbk_Book As Workbook, str_Name As String) As Boolean
    ' ## Return true if specified worksheet name exists within the specified workbook
    On Error GoTo errHandler

    Dim wks_Found As Worksheet: Set wks_Found = wbk_Book.Worksheets(str_Name)

errHandler:
    Select Case Err
        Case 9: Examine_WorksheetExists = False
        Case 0: Examine_WorksheetExists = True
    End Select
End Function

Public Function Examine_NumRows(rng_Top As Range) As Long
    ' ## Input the first row containing output, not the header
    ' Returns the number of rows, including blank rows
    Dim lng_Output As Long
    Dim rng_NextAfterBottom As Range: Set rng_NextAfterBottom = Gather_RowForAppend(rng_Top)
    If rng_NextAfterBottom Is Nothing Then
        Set rng_NextAfterBottom = rng_Top(1, 1).Offset(rng_Top.Worksheet.Rows.Count - rng_Top.Row, 0)
        lng_Output = rng_NextAfterBottom.Row + 1 - rng_Top.Row
    Else
        lng_Output = rng_NextAfterBottom.Row - rng_Top.Row
    End If
    If lng_Output <= 0 Then lng_Output = 0
    Examine_NumRows = lng_Output
End Function

Public Function Examine_NumCols(rng_Left As Range) As Integer
    ' ## Input the first column containing output, not the header
    ' Returns the number of columns, including blank columns
    Dim int_Output As Integer: int_Output = Gather_ColForAppend(rng_Left).Column - rng_Left.Column
    If int_Output <= 0 Then int_Output = 0
    Examine_NumCols = int_Output
End Function

Function Examine_UpperBoundIndex_Row(arr_2D As Variant) As Long
    ' ## Find the upper bound of the row dimension of a range or 2D array
    Dim lng_Output As Long
    Dim str_Type As String: str_Type = TypeName(arr_2D)
    If str_Type = "Range" Then lng_Output = arr_2D.Rows.Count Else lng_Output = UBound(arr_2D, 1)

    Examine_UpperBoundIndex_Row = lng_Output
End Function

Function Examine_UpperBoundIndex_Col(arr_2D As Variant) As Long
    ' ## Find the upper bound of the column dimension of a range or 2D array
    Dim lng_Output As Long
    Dim str_Type As String: str_Type = TypeName(arr_2D)
    If str_Type = "Range" Then lng_Output = arr_2D.Columns.Count Else lng_Output = UBound(arr_2D, 2)

    Examine_UpperBoundIndex_Col = lng_Output
End Function

Function Examine_LowerBoundIndex_Row(arr_2D As Variant) As Long
    ' ## Find the lower bound of the row dimension of a range or 2D array
    Dim lng_Output As Long
    Dim str_Type As String: str_Type = TypeName(arr_2D)
    If str_Type = "Range" Then lng_Output = 1 Else lng_Output = LBound(arr_2D, 1)

    Examine_LowerBoundIndex_Row = lng_Output
End Function

Function Examine_LowerBoundIndex_Col(arr_2D As Variant) As Long
    ' ## Find the lower bound of the column dimension of a range or 2D array
    Dim lng_Output As Long
    Dim str_Type As String: str_Type = TypeName(arr_2D)
    If str_Type = "Range" Then lng_Output = 1 Else lng_Output = LBound(arr_2D, 2)

    Examine_LowerBoundIndex_Col = lng_Output
End Function

Function Examine_IsArrayInitialized(arr_Input As Variant) As Boolean
    ' ## Returns false if dynamic array has not been re-dimensioned with bounds yet
    On Error GoTo errHandler
    Dim int_UBound As Integer: int_UBound = UBound(arr_Input)

errHandler:
    Select Case Err
        Case 0: Examine_IsArrayInitialized = True
        Case 9: Examine_IsArrayInitialized = False
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description
    End Select
End Function

Public Function Examine_MaxOfPair(var_ValueA As Variant, var_ValueB As Variant) As Variant
    ' ## Return the maximum of the two specified values
    If var_ValueA > var_ValueB Then
        Examine_MaxOfPair = var_ValueA
    Else
        Examine_MaxOfPair = var_ValueB
    End If
End Function

Public Function Examine_MinOfPair(var_ValueA As Variant, var_ValueB As Variant) As Variant
    ' ## Return the maximum of the two specified values
    If var_ValueA < var_ValueB Then
        Examine_MinOfPair = var_ValueA
    Else
        Examine_MinOfPair = var_ValueB
    End If
End Function

Public Function Examine_CountNumBelow(varLst_Input As Collection, var_Value As Variant, bln_UniqueAscending As Boolean) As Integer
    ' ## Return the number of items same as or below the specified value
    ' ## If the array contains unique values sorted in ascending order, this function is much faster
    Dim int_Ctr As Integer
    Dim int_TotalNum As Integer: int_TotalNum = varLst_Input.Count
    Dim int_Output As Integer

    If int_TotalNum = 0 Then
        int_Output = 0
    ElseIf bln_UniqueAscending = True And int_TotalNum > 4 Then
        ' Use shortcut
        ' Initialize bounds
        Dim int_LowerBound As Integer: int_LowerBound = 1
        Dim int_UpperBound As Integer: int_UpperBound = int_TotalNum
        Dim int_ActiveGuess As Integer, var_ActiveLookup As Variant
        Dim bln_Continue As Boolean: bln_Continue = True

        ' Check if value outside bounds already
        If var_Value >= varLst_Input(int_UpperBound) Then
            bln_Continue = False
            int_Output = int_UpperBound
        ElseIf var_Value <= varLst_Input(int_LowerBound) Then
            bln_Continue = False
            If var_Value = varLst_Input(int_LowerBound) Then int_Output = 1 Else int_Output = 0
        End If

        While bln_Continue = True
            ' Generate observation
            int_ActiveGuess = (int_LowerBound + int_UpperBound) / 2
            var_ActiveLookup = varLst_Input(int_ActiveGuess)

            ' Update bounds based on observation
            If var_Value = var_ActiveLookup Then
                Examine_CountNumBelow = int_ActiveGuess
                Exit Function
            ElseIf var_Value < var_ActiveLookup Then
                int_UpperBound = int_ActiveGuess
            ElseIf var_Value > var_ActiveLookup Then
                int_LowerBound = int_ActiveGuess
            End If

            ' Handle case where no more bisection is possible
            If int_UpperBound - int_LowerBound = 0 Then
                bln_Continue = False
                int_Output = int_LowerBound
            ElseIf int_UpperBound - int_LowerBound = 1 Then
                bln_Continue = False
                If var_Value >= varLst_Input(int_UpperBound) Then int_Output = int_UpperBound Else int_Output = int_LowerBound
            End If
        Wend
    Else
        ' Examine entire list
        For int_Ctr = 1 To varLst_Input.Count
            If varLst_Input(int_Ctr) <= var_Value Then int_Output = int_Output + 1
        Next int_Ctr
    End If

    Examine_CountNumBelow = int_Output
End Function

Public Function Examine_MaxValueInList(varLst_Input As Collection) As Variant
    ' ## Return the maximum value contained within the list
    Dim var_Max As Variant: var_Max = varLst_Input(1)
    Dim int_Ctr As Integer
    For int_Ctr = 2 To varLst_Input.Count
        If varLst_Input(int_Ctr) > var_Max Then var_Max = varLst_Input(int_Ctr)
    Next int_Ctr

    Examine_MaxValueInList = var_Max
End Function

Public Function Examine_MinValueInList(varLst_Input As Collection) As Variant
    ' ## Return the minimum value contained within the list
    Dim var_Min As Variant: var_Min = varLst_Input(1)
    Dim int_Ctr As Integer
    For int_Ctr = 2 To varLst_Input.Count
        If varLst_Input(int_Ctr) < var_Min Then var_Min = varLst_Input(int_Ctr)
    Next int_Ctr

    Examine_MinValueInList = var_Min
End Function

Public Function Examine_IsUniform(arr_List As Variant) As Boolean
    ' ## Return true if all the values in the 1D array or collection are identical
    Dim bln_Output As Boolean: bln_Output = True
    Dim lng_LowerBound As Integer: lng_LowerBound = Examine_LowerBoundIndex(arr_List)
    Dim lng_UpperBound As Integer: lng_UpperBound = Examine_UpperBoundIndex(arr_List)
    Dim var_FirstVal As Variant: var_FirstVal = arr_List(lng_LowerBound)
    Dim lng_Ctr As Long

    ' Check any values are different from the first value
    For lng_Ctr = lng_LowerBound To lng_UpperBound
        If arr_List(lng_Ctr) <> var_FirstVal Then bln_Output = False
    Next lng_Ctr

    Examine_IsUniform = bln_Output
End Function