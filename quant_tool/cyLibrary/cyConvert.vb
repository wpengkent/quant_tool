Option Explicit

Public Function Convert_RangeToList(rng_Input As Range, Optional intLst_IndexFilter As Collection = Nothing) As Collection
    ' ## Convert values contained within an Excel range to a list
    ' ## The Excel range must be a one-dimensional vector, either horizontal or vertical
    ' ## Index filter is a list of indices to include, specified in order
    Dim varLst_Output As New Collection

    If TypeName(rng_Input.Value) = "Variant()" Then
        Dim varArr_Values() As Variant: varArr_Values = rng_Input.Value
        If rng_Input.Columns.Count = 1 Then
            Dim lng_RowCtr As Long
            For lng_RowCtr = 1 To rng_Input.Rows.Count
                Call varLst_Output.Add(varArr_Values(lng_RowCtr, 1))
            Next lng_RowCtr
        ElseIf rng_Input.Rows.Count = 1 Then
            Dim lng_ColCtr As Long
            For lng_ColCtr = 1 To rng_Input.Columns.Count
                Call varLst_Output.Add(varArr_Values(1, lng_ColCtr))
            Next lng_ColCtr
        End If
    Else
        Call varLst_Output.Add(rng_Input.Value)
    End If

    If Not intLst_IndexFilter Is Nothing Then
        Set varLst_Output = Convert_FilterList(varLst_Output, intLst_IndexFilter)
    End If

    Set Convert_RangeToList = varLst_Output
End Function

Public Function Convert_FilterList(varLst_Orig As Collection, intLst_IndexFilter As Collection) As Collection
    ' ## Index filter is a list of indices to include, specified in order
    Dim varLst_Output As New Collection
    Dim int_FilterCount As Integer: int_FilterCount = intLst_IndexFilter.Count

    ' Determine whether input range is horizontal or vertical
    Dim int_RowCtr As Integer
    For int_RowCtr = 1 To int_FilterCount
        Call varLst_Output.Add(varLst_Orig(intLst_IndexFilter(int_RowCtr)))
    Next int_RowCtr

    Set Convert_FilterList = varLst_Output
End Function

Public Function Convert_FilterArr(varArr_Orig As Variant, intLst_IndexFilter As Collection) As Variant()
    ' ## Index filter is a list of indices to include, specified in order
    Dim int_FilterCount As Integer: int_FilterCount = intLst_IndexFilter.Count
    Dim varArr_Output() As Variant: ReDim varArr_Output(1 To int_FilterCount) As Variant

    ' Determine whether input range is horizontal or vertical
    Dim int_RowCtr As Integer
    For int_RowCtr = 1 To int_FilterCount
        varArr_Output(int_RowCtr) = varArr_Orig(intLst_IndexFilter(int_RowCtr))
    Next int_RowCtr

    Convert_FilterArr = varArr_Output
End Function

Public Function Convert_RangeToArrDict(rng_Input As Range, Optional intLst_RowFilter As Collection = Nothing, _
    Optional intLst_ColFilter As Collection = Nothing, Optional bln_InvertRows As Boolean = False, _
    Optional bln_InvertCols As Boolean = False) As Dictionary
    ' ## Convert values contained within an Excel range to a dictionary of arrays
    ' ## Index filter is a list of indices to include, specified in order
    Dim dic_Output As Dictionary: Set dic_Output = New Dictionary
    Dim int_FilterCount_Row As Integer, int_FilterCount_Col As Integer

    ' Determine number of rows in output
    Dim bln_Filtered_Row As Boolean, bln_Filtered_Col As Boolean
    If intLst_RowFilter Is Nothing Then
        int_FilterCount_Row = rng_Input.Rows.Count
        bln_Filtered_Row = False
    Else
        int_FilterCount_Row = intLst_RowFilter.Count
        bln_Filtered_Row = True
    End If

    ' Determine number of columns in output
    If intLst_ColFilter Is Nothing Then
        int_FilterCount_Col = rng_Input.Columns.Count
        bln_Filtered_Col = False
    Else
        int_FilterCount_Col = intLst_ColFilter.Count
        bln_Filtered_Col = True
    End If

    ' Temporary variables
    Dim dblArr_ActiveRow() As Double: ReDim dblArr_ActiveRow(1 To int_FilterCount_Col) As Double
    Dim int_ActiveRow As Integer, int_ActiveCol As Integer

    ' Fill output
    Dim int_RowCtr As Integer, int_ColCtr As Integer
    For int_RowCtr = 1 To int_FilterCount_Row
        If bln_Filtered_Row = True Then int_ActiveRow = intLst_RowFilter(int_RowCtr) Else int_ActiveRow = int_RowCtr
        If bln_InvertRows = True Then int_ActiveRow = int_FilterCount_Row + 1 - int_ActiveRow

        ' Fill array with values from the current row
        For int_ColCtr = 1 To int_FilterCount_Col
            If bln_Filtered_Col = True Then int_ActiveCol = intLst_ColFilter(int_ColCtr) Else int_ActiveCol = int_ColCtr
            If bln_InvertCols = True Then int_ActiveCol = int_FilterCount_Col + 1 - int_ActiveCol

            dblArr_ActiveRow(int_ColCtr) = rng_Input(int_ActiveRow, int_ActiveCol).Value
        Next int_ColCtr

        ' Add to list
        Call dic_Output.Add(int_RowCtr, dblArr_ActiveRow)
    Next int_RowCtr

    Set Convert_RangeToArrDict = dic_Output
End Function

Public Function Convert_RangeToDblArr(rng_Input As Range, Optional intLst_IndexFilter As Collection = Nothing) As Double()
    Convert_RangeToDblArr = Convert_ListToDblArr(Convert_RangeToList(rng_Input, intLst_IndexFilter))
End Function

Public Function Convert_RangeToArr2D(rng_Input As Range) As Variant()
    ' ## Convert values contained within any rectangular Excel range to a two-dimensional array of matching dimensions
    Dim varArr_Output() As Variant
    If rng_Input.Rows.Count > 1 Or rng_Input.Columns.Count > 1 Then
        varArr_Output = rng_Input.Value2
    Else
        ' Can't assign 1x1 range to array, assign the value
        ReDim varArr_Output(1 To 1, 1 To 1) As Variant
        varArr_Output(1, 1) = rng_Input.Value2
    End If
    Convert_RangeToArr2D = varArr_Output
End Function

Public Function Convert_RangeToLngArr(rng_Input As Range, Optional intLst_IndexFilter As Collection = Nothing) As Long()
    Convert_RangeToLngArr = Convert_ListToLngArr(Convert_RangeToList(rng_Input, intLst_IndexFilter))
End Function

Public Function Convert_ListToLngArr(lngLst_Input As Collection) As Long()
    ' ## Return an array containing the same values as the list
    Dim lngArr_Output() As Long: ReDim lngArr_Output(1 To lngLst_Input.Count) As Long
    Dim lng_Ctr As Long
    For lng_Ctr = 1 To lngLst_Input.Count
        lngArr_Output(lng_Ctr) = lngLst_Input(lng_Ctr)
    Next lng_Ctr

    Convert_ListToLngArr = lngArr_Output
End Function

Public Function Convert_RangeToIntArr(rng_Input As Range, Optional intLst_IndexFilter As Collection = Nothing) As Integer()
    Convert_RangeToIntArr = Convert_ListToIntArr(Convert_RangeToList(rng_Input, intLst_IndexFilter))
End Function

Public Function Convert_ListToIntArr(intLst_Input As Collection) As Integer()
    ' ## Return an array containing the same values as the list
    Dim intArr_Output() As Integer: ReDim intArr_Output(1 To intLst_Input.Count) As Integer
    Dim int_Ctr As Integer
    For int_Ctr = 1 To intLst_Input.Count
        intArr_Output(int_Ctr) = intLst_Input(int_Ctr)
    Next int_Ctr

    Convert_ListToIntArr = intArr_Output
End Function

Public Function Convert_ListToArr2D(varLst_Input As Collection, Optional bln_TransposeToHorizontal As Boolean = False) As Variant()
    ' ## Return an array containing the same values as the list
    Dim int_NumItems As Integer: int_NumItems = varLst_Input.Count
    Dim varArr_Output() As Variant
    If bln_TransposeToHorizontal = True Then
        ReDim varArr_Output(1 To 1, 1 To int_NumItems) As Variant
    Else
        ReDim varArr_Output(1 To int_NumItems, 1 To 1) As Variant
    End If

    Dim int_Ctr As Integer
    If bln_TransposeToHorizontal = True Then
        For int_Ctr = 1 To int_NumItems
            varArr_Output(1, int_Ctr) = varLst_Input(int_Ctr)
        Next int_Ctr
    Else
        For int_Ctr = 1 To int_NumItems
            varArr_Output(int_Ctr, 1) = varLst_Input(int_Ctr)
        Next int_Ctr
    End If

    Convert_ListToArr2D = varArr_Output
End Function

Public Function Convert_ArrToList(arr_Input() As Double) As Collection
    ' ## Returns a list containing the same values as the array
    Dim varLst_Output As New Collection
    If Examine_IsArrayInitialized(arr_Input) = True Then
        Dim int_Ctr As Integer
        For int_Ctr = LBound(arr_Input) To UBound(arr_Input)
            Call varLst_Output.Add(arr_Input(int_Ctr))
        Next int_Ctr
    End If

    Set Convert_ArrToList = varLst_Output
End Function

Public Function Convert_ListToDblArr(dblLst_Input As Collection) As Double()
    ' ## Return an array containing the same values as the list
    Dim dblArr_Output() As Double: ReDim dblArr_Output(1 To dblLst_Input.Count) As Double
    Dim int_Ctr As Integer
    For int_Ctr = 1 To dblLst_Input.Count
        dblArr_Output(int_Ctr) = CDbl(dblLst_Input(int_Ctr))
    Next int_Ctr

    Convert_ListToDblArr = dblArr_Output
End Function

Public Function Convert_ListToStrArr(strLst_Input As Collection) As String()
    ' ## Return an array containing the same values as the list
    Dim strArr_Output() As String: ReDim strArr_Output(1 To strLst_Input.Count) As String
    Dim int_Ctr As Integer
    For int_Ctr = 1 To strLst_Input.Count
        strArr_Output(int_Ctr) = strLst_Input(int_Ctr)
    Next int_Ctr

    Convert_ListToStrArr = strArr_Output
End Function

Public Function Convert_Array1DTo2D(varArr_1D As Variant, Optional bln_TransposeToHorizontal As Boolean = False) As Variant()
    ' ## Returns a 2 dimensional array with the identical indicies and values as the specified 1 dimensional array
    Dim int_LBound As Integer: int_LBound = LBound(varArr_1D)
    Dim int_UBound As Integer: int_UBound = UBound(varArr_1D)
    Dim varArr_2D() As Variant
    If bln_TransposeToHorizontal = True Then
        ReDim varArr_2D(int_LBound To int_LBound, int_LBound To int_UBound) As Variant
    Else
        ReDim varArr_2D(int_LBound To int_UBound, int_LBound To int_LBound) As Variant
    End If

    Dim int_Ctr As Integer

    For int_Ctr = int_LBound To int_UBound
        If bln_TransposeToHorizontal = True Then
            varArr_2D(int_LBound, int_Ctr) = varArr_1D(int_Ctr)
        Else
            varArr_2D(int_Ctr, int_LBound) = varArr_1D(int_Ctr)
        End If
    Next int_Ctr

    Convert_Array1DTo2D = varArr_2D
End Function

Public Function Convert_Reverse(dblArr_Input() As Double) As Double()
    ' ## Return a copy which has reversed the order of array elements compared to the specified input array

    Dim int_LowerBound As Integer: int_LowerBound = LBound(dblArr_Input)
    Dim int_UpperBound As Integer: int_UpperBound = UBound(dblArr_Input)
    Dim dblArr_Output() As Double: ReDim dblArr_Output(int_LowerBound To int_UpperBound) As Double

    Dim int_Ctr As Integer
    For int_Ctr = int_LowerBound To int_UpperBound
        dblArr_Output(int_Ctr) = dblArr_Input(int_UpperBound + int_LowerBound - int_Ctr)
    Next int_Ctr

    Convert_Reverse = dblArr_Output
End Function

Public Function Convert_Reverse_List(varLst_Input As Collection) As Collection
    ' ## Return a copy which has reversed the order of elements compared to the specified input list
    Dim varLst_Output As New Collection
    Dim int_NumItems As Integer: int_NumItems = varLst_Input.Count
    Dim int_Ctr As Integer
    For int_Ctr = 1 To int_NumItems
        Call varLst_Output.Add(varLst_Input(int_NumItems + 1 - int_Ctr))
    Next int_Ctr

    Set Convert_Reverse_List = varLst_Output
End Function

Public Function Convert_Transpose_DblArr(dblArr_Input() As Double) As Double()
    ' ## Convert single dimensional vector to a two dimensional vector with multiple rows and one column

    Dim int_LowerBound As Integer: int_LowerBound = LBound(dblArr_Input)
    Dim int_UpperBound As Integer: int_UpperBound = UBound(dblArr_Input)
    Dim dblArr_Output() As Double: ReDim dblArr_Output(int_LowerBound To int_UpperBound, 1 To 1) As Double

    Dim int_Ctr As Integer
    For int_Ctr = int_LowerBound To int_UpperBound
        dblArr_Output(int_Ctr, 1) = dblArr_Input(int_Ctr)
    Next int_Ctr

    Convert_Transpose_DblArr = dblArr_Output
End Function

Public Function Convert_Transpose_LngArr(lngArr_Input() As Long) As Long()
    ' ## Convert single dimensional vector to a two dimensional vector with multiple rows and one column

    Dim int_LowerBound As Integer: int_LowerBound = LBound(lngArr_Input)
    Dim int_UpperBound As Integer: int_UpperBound = UBound(lngArr_Input)
    Dim lngArr_Output() As Long: ReDim dblArr_Output(int_LowerBound To int_UpperBound, 1 To 1) As Long

    Dim int_Ctr As Integer
    For int_Ctr = int_LowerBound To int_UpperBound
        lngArr_Output(int_Ctr, 1) = lngArr_Input(int_Ctr)
    Next int_Ctr

    Convert_Transpose_LngArr = lngArr_Output
End Function

Public Function Convert_Join_StrArr(ByVal strArr_1 As Variant, ByVal strArr_2 As Variant) As String()
    ' ## Join together two arrays
    Dim int_Count_2 As Integer: int_Count_2 = UBound(strArr_2) - LBound(strArr_2) + 1
    Dim int_UBound_1 As Integer: int_UBound_1 = UBound(strArr_1)
    Dim int_LBound_2 As Integer: int_LBound_2 = LBound(strArr_2)

    ' Keep elements in array 1
    ReDim Preserve strArr_1(LBound(strArr_1) To int_UBound_1 + int_Count_2) As String

    ' Add elements in array 2 to the end
    Dim int_Ctr As Integer
    For int_Ctr = 1 To int_Count_2
        strArr_1(int_UBound_1 + int_Ctr) = strArr_2(int_Ctr + int_LBound_2 - 1)
    Next int_Ctr

    Convert_Join_StrArr = strArr_1
End Function

Public Function Convert_Sort_String(strArr_Original() As String) As String()
    Dim intArr_SortKey() As Integer: intArr_SortKey = Calc_SortKey(strArr_Original, True)
    Dim int_LBound As Integer: int_LBound = LBound(strArr_Original)
    Dim int_UBound As Integer: int_UBound = UBound(strArr_Original)
    Dim strArr_Output() As String: ReDim strArr_Output(int_LBound To int_UBound) As String

    Dim int_Ctr As Integer
    For int_Ctr = int_LBound To int_UBound
        strArr_Output(intArr_SortKey(int_Ctr)) = strArr_Original(int_Ctr)
    Next int_Ctr

    Convert_Sort_String = strArr_Output
End Function

Public Function Convert_SQLDate(lng_Date As Long) As String
    ' ## Convert an Excel date to a format recognised by SQL queries
    Convert_SQLDate = Format(lng_Date, "mm/d/yyyy")
End Function

Public Function Convert_DFToZero(dbl_DF As Double, lng_StartDate As Long, lng_EndDate As Long, str_RateType As String, _
    Optional str_CouponFreq As String = "6M", Optional bln_IsFwdGeneration As Boolean = False) As Double
    ' ## Convert a discount factor between two specified dates to an equivalent zero rate expressed in the specified convention
    ' ## Returns zero rate expressed as a percentage
    Select Case UCase(str_RateType)
        Case "ZERO": Convert_DFToZero = -Math.Log(dbl_DF) / Calc_YearFrac(lng_StartDate, lng_EndDate, "ACT/365", "") * 100
        Case "ACT/365": Convert_DFToZero = (1 / dbl_DF - 1) / Calc_YearFrac(lng_StartDate, lng_EndDate, "ACT/365", "") * 100
        Case "ACT/360": Convert_DFToZero = (1 / dbl_DF - 1) / Calc_YearFrac(lng_StartDate, lng_EndDate, "ACT/360", "") * 100

        Case "ACT/ACT", "ACT/ACT NM", "ACT/ACT CPN", "ACT/ACT XTE":
        Convert_DFToZero = (1 / dbl_DF - 1) / Calc_YearFrac(lng_StartDate, lng_EndDate, str_RateType, str_CouponFreq, bln_IsFwdGeneration) * 100

        'Other Linear Rate Type
        Case Else: Convert_DFToZero = (1 / dbl_DF - 1) / Calc_YearFrac(lng_StartDate, lng_EndDate, str_RateType, "") * 100
    End Select
End Function

Public Function Convert_SimpToDF(dbl_ParRate As Double, lng_StartDate As Long, lng_EndDate As Long, str_DCC As String, _
    Optional str_CouponFreq As String = "6M") As Double
    ' ## Take a simple interest rate defined using the specified daycount convention and convert to discount factor
    Convert_SimpToDF = 1 / (1 + dbl_ParRate / 100 * Calc_YearFrac(lng_StartDate, lng_EndDate, str_DCC, str_CouponFreq))
End Function

Public Function Convert_DiscToDF(dbl_ParRate As Double, lng_StartDate As Long, lng_EndDate As Long, str_DCC As String, _
    Optional str_CouponFreq As String = "6M") As Double
    ' ## Take a simple discount rate defined using the specified daycount convention and convert to discount factor
    Convert_DiscToDF = 1 - dbl_ParRate / 100 * Calc_YearFrac(lng_StartDate, lng_EndDate, str_DCC, str_CouponFreq)
End Function

Public Function Convert_DiscBondPriceToZeroRate(dbl_DiscBondPrice As Double, lng_StartDate As Long, lng_EndDate As Long, str_DCC As String, _
    Optional str_CouponFreq As String = "6M") As Double
    ' ## Take a simple interest rate defined using the specified daycount convention and convert to discount factor
    Convert_DiscBondPriceToZeroRate = -WorksheetFunction.Ln(dbl_DiscBondPrice / 100) / Calc_YearFrac(lng_StartDate, lng_EndDate, str_DCC, str_CouponFreq) * 100
End Function

Public Function Convert_DiscRateToZeroRate(dbl_DiscRate As Double, lng_StartDate As Long, lng_EndDate As Long, str_DCC As String, _
    Optional str_CouponFreq As String = "6M") As Double
    ' ## Take a simple interest rate defined using the specified daycount convention and convert to discount factor
    Convert_DiscRateToZeroRate = -WorksheetFunction.Ln(1 - dbl_DiscRate / 100 * Calc_YearFrac(lng_StartDate, lng_EndDate, str_DCC, str_CouponFreq)) / Calc_YearFrac(lng_StartDate, lng_EndDate, "ACT/365", str_CouponFreq) * 100
End Function

Public Function Convert_RecordsToList(ByRef rst_Data As ADODB.Recordset, int_Column As Integer, str_DataType As String) As Collection
    ' ## Read query results and convert to a native collection format containing just the values in the specified column
    Dim str_DataTypeUCase As String: str_DataTypeUCase = UCase(str_DataType)
    Dim lst_Output As New Collection

    While Not rst_Data.EOF
        Select Case str_DataTypeUCase
            Case "LONG": Call lst_Output.Add(CLng(rst_Data(int_Column)))
            Case "STRING": Call lst_Output.Add(CStr(rst_Data(int_Column)))
        End Select
        rst_Data.MoveNext
    Wend

    Set Convert_RecordsToList = lst_Output
End Function

Public Function Convert_Split(str_Orig As String, str_Delimiter As String, Optional int_Pos_Base1 As Integer = -1) As Variant
    If int_Pos_Base1 = -1 Then
        Convert_Split = Split(str_Orig, str_Delimiter)
    Else
        Convert_Split = Split(str_Orig, str_Delimiter)(int_Pos_Base1 - 1)
    End If
End Function

Public Function Convert_SplitToList(str_Orig As String, str_Delimiter As String) As Collection
    ' ## Split the string by the specified delimiter and output the results in a list
    Dim varLst_Output As New Collection
    Dim varArr_Split As Variant: varArr_Split = Split(str_Orig, str_Delimiter)
    Dim int_Ctr As Integer

    For int_Ctr = LBound(varArr_Split) To UBound(varArr_Split)
        Call varLst_Output.Add(varArr_Split(int_Ctr))
    Next int_Ctr

    Set Convert_SplitToList = varLst_Output
End Function

Public Function Convert_MergeLists(varLst_Left As Collection, varLst_Right As Collection) As Collection
    ' ## Return a new collection containing the elements of the right appended to those of the left list
    Dim varLst_Output As New Collection, int_Ctr As Integer
    For int_Ctr = 1 To varLst_Left.Count
        Call varLst_Output.Add(varLst_Left(int_Ctr))
    Next int_Ctr

    For int_Ctr = 1 To varLst_Right.Count
        Call varLst_Output.Add(varLst_Right(int_Ctr))
    Next int_Ctr

    Set Convert_MergeLists = varLst_Output
End Function

Public Function Convert_MergeDicts(dic_Left As Dictionary, dic_Right As Dictionary) As Dictionary
    ' ## Return a new dictionary containing the elements of the right appended to those of the left
    Dim dic_Output As New Dictionary: dic_Output.CompareMode = dic_Left.CompareMode
    Dim var_Key As Variant

    For Each var_Key In dic_Left.Keys
        Call dic_Output.Add(var_Key, dic_Left(var_Key))
    Next var_Key

    For Each var_Key In dic_Right.Keys
        If dic_Output.Exists(var_Key) = False Then Call dic_Output.Add(var_Key, dic_Right(var_Key))
    Next var_Key

    Set Convert_MergeDicts = dic_Output
End Function

Public Function Convert_ListToParams(varLst_Items As Collection, Optional str_EnclosingChars As String = "") As String
    ' ## Converts the collection to (Item1, Item2, Item3, ...) format
    ' ## Each item is surrounded by the specified enclosing character
    Dim int_NumItems As Integer: int_NumItems = varLst_Items.Count
    Dim str_SQL As String

    If int_NumItems > 0 Then
        str_SQL = "("

        Dim int_Ctr As Integer
        For int_Ctr = 1 To int_NumItems
            str_SQL = str_SQL & str_EnclosingChars & varLst_Items(int_Ctr) & str_EnclosingChars & ", "
        Next int_Ctr

        ' Cut off final comma and space, then add closing parentheses
        str_SQL = Left(str_SQL, Len(str_SQL) - 2)
        str_SQL = str_SQL & ")"
    Else
        str_SQL = ""
    End If

    Convert_ListToParams = str_SQL
End Function

Public Function Convert_PayoffCode(enu_Payoff As EuropeanPayoff) As String
    ' ## Return a string representation of the payoff type enumeration
    Dim str_Output As String
    Select Case enu_Payoff
        Case EuropeanPayoff.Standard: str_Output = "Standard"
        Case EuropeanPayoff.Digital_CoN: str_Output = "Digital CoN"
        Case EuropeanPayoff.Digital_AoN: str_Output = "Digital AoN"
    End Select
    Convert_PayoffCode = str_Output
End Function

Public Function Convert_ArrListToArr2D(arrLst_Input As Collection, bln_PreserveInput As Boolean, Optional lng_MaxRows As Long = -1) As Variant()
    ' Prepare output array
    Dim lng_NumRows As Long: lng_NumRows = arrLst_Input.Count
    If lng_MaxRows <> -1 And lng_NumRows > lng_MaxRows Then lng_NumRows = lng_MaxRows  ' Limit the number to convert
    Dim int_NumCols As Integer: int_NumCols = UBound(arrLst_Input(1)) - LBound(arrLst_Input(1)) + 1
    Dim varArr_Output() As Variant: ReDim varArr_Output(1 To lng_NumRows, 1 To int_NumCols) As Variant
    Dim lng_RowCtr As Long

    ' Preserve input if required
    Dim arrLst_Preserved As Collection
    If bln_PreserveInput = True Then Set arrLst_Preserved = New Collection

    ' Populate output array
    Dim int_ColCtr As Integer, varArr_Active() As Variant
    For lng_RowCtr = 1 To lng_NumRows
        varArr_Active = arrLst_Input(1)
        If bln_PreserveInput = True Then Call arrLst_Preserved.Add(varArr_Active)
        Call arrLst_Input.Remove(1)  ' Remove from list for performance reasons since Collection is a linked list structure

        For int_ColCtr = 1 To int_NumCols
            varArr_Output(lng_RowCtr, int_ColCtr) = varArr_Active(int_ColCtr)
        Next int_ColCtr
    Next lng_RowCtr

    If bln_PreserveInput = True Then Set arrLst_Input = arrLst_Preserved
    Convert_ArrListToArr2D = varArr_Output
End Function

Public Function Convert_Arr2DToDict(strArr_Keys() As Variant, strArr_Values() As Variant) As Dictionary
    ' ## Places the specified keys and values in a dictionary form
    ' ## Assumes arrays are 2D with indices starting from 1, and keys are unique
    Dim dic_Output As New Dictionary: dic_Output.CompareMode = CompareMethod.TextCompare
    Dim lng_NumItems As Long: lng_NumItems = UBound(strArr_Keys, 1)
    Debug.Assert UBound(strArr_Values, 1) = lng_NumItems

    Dim lng_RowCtr As Long
    For lng_RowCtr = 1 To lng_NumItems
        Call dic_Output.Add(strArr_Keys(lng_RowCtr, 1), strArr_Values(lng_RowCtr, 1))
    Next lng_RowCtr

    Set Convert_Arr2DToDict = dic_Output
End Function