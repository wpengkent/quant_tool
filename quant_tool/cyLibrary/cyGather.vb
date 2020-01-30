Option Explicit
' ## EXECUTE A PROCESS TO RETURN A NON-NUMERIC OUTPUT

Public Function Gather_RecordSet_Sybase(sbp_Params As SybaseParams, str_Database As String, str_SQL As String) As ADODB.Recordset
    ' ## Execute SQL query in Sybase and return results as a recordset object

    ' Build connection string
    Dim str_Conn As String: str_Conn = "DRIVER={Adaptive Server Enterprise};UID=" & sbp_Params.UserID & ";PWD=" _
        & sbp_Params.Password & ";database=" & str_Database & ";server=" & sbp_Params.IPAddress & ";port=" _
        & sbp_Params.Port & ";"

    Dim bln_ExistingMsg As Boolean: bln_ExistingMsg = (Application.StatusBar <> False)
    Dim str_OrigMessage As String: str_OrigMessage = Application.StatusBar
    Application.StatusBar = "Querying database..."

    ' Create record set, run query and return results
    Dim rst_Results As New ADODB.Recordset
    Call rst_Results.Open(str_SQL, str_Conn)

    If bln_ExistingMsg = True Then
        Application.StatusBar = str_OrigMessage
    Else
        Application.StatusBar = False
    End If

    Set Gather_RecordSet_Sybase = rst_Results
End Function

Public Function Gather_Hols(sbp_Params As SybaseParams, str_Database As String, mqp_Params As MxQP_Hols) As Long()
    ' ## Returns annual and special holidays combined over date window

    ' Construct query then read to list - special hols
    Dim str_SQL_Special As String: str_SQL_Special = SQL_MX_SpecialHols(mqp_Params)
    Dim rst_SpecialHols As ADODB.Recordset: Set rst_SpecialHols = Gather_RecordSet_Sybase(sbp_Params, str_Database, str_SQL_Special)
    Dim lst_SpecialHols As New Collection: Set lst_SpecialHols = Convert_RecordsToList(rst_SpecialHols, 0, "Long")

    ' Construct query then read to list - annual hols
    Dim str_SQL_Annual As String: str_SQL_Annual = SQL_MX_AnnualHols(mqp_Params)
    Dim rst_AnnualHols As ADODB.Recordset: Set rst_AnnualHols = Gather_RecordSet_Sybase(sbp_Params, str_Database, str_SQL_Annual)
    Dim lst_AnnualHols As New Collection: Set lst_AnnualHols = Convert_RecordsToList(rst_AnnualHols, 0, "Long")

    Dim int_StartYear As Integer: int_StartYear = Year(mqp_Params.MinDate)
    Dim int_EndYear As Integer: int_EndYear = Year(mqp_Params.MaxDate)
    Dim int_AnnualHolsPerYear As Integer: int_AnnualHolsPerYear = lst_AnnualHols.Count
    Dim int_NumAnnualHols As Integer: int_NumAnnualHols = (int_EndYear - int_StartYear + 1) * int_AnnualHolsPerYear
    Dim int_NumSpecialHols As Integer: int_NumSpecialHols = lst_SpecialHols.Count

    Dim lngArr_Output() As Long: ReDim lngArr_Output(1 To int_NumAnnualHols + int_NumSpecialHols, 1 To 1) As Long
    Dim int_Ctr As Integer
    Dim lng_ActiveHol As Long
    Dim int_ActiveIndex As Integer
    Dim int_ActiveYear As Integer: int_ActiveYear = int_StartYear

    ' Start by creating annual holidays
    For int_Ctr = 1 To int_NumAnnualHols
        int_ActiveIndex = int_Ctr Mod int_AnnualHolsPerYear
        If int_ActiveIndex = 0 Then int_ActiveIndex = int_AnnualHolsPerYear

        ' Apply holiday to current year
        lng_ActiveHol = lst_AnnualHols(int_ActiveIndex)
        lngArr_Output(int_Ctr, 1) = DateAdd("yyyy", int_ActiveYear - Year(lng_ActiveHol), lng_ActiveHol)

        ' Move to next year when past the last holiday for the year
        If int_ActiveIndex = int_AnnualHolsPerYear Then int_ActiveYear = int_ActiveYear + 1
    Next int_Ctr

    ' Append special holidays
    For int_Ctr = 1 To int_NumSpecialHols
        lngArr_Output(int_Ctr + int_NumAnnualHols, 1) = lst_SpecialHols(int_Ctr)
    Next int_Ctr

    Gather_Hols = lngArr_Output
End Function

Public Function Gather_RowForAppend(rng_Top As Range) As Range
    ' ## Returns the row after the bottom row, in the specified column
    ' ## Returns null if bottom row is populated
    Dim rng_BottomRow As Range: Set rng_BottomRow = rng_Top(1, 1).Offset(rng_Top.Worksheet.Rows.Count - rng_Top.Row, 0)
    If rng_BottomRow.Value = "" Then Set Gather_RowForAppend = rng_BottomRow.End(xlUp).Offset(1, 0)
End Function

Public Function Gather_ColForAppend(rng_Left As Range) As Range
    ' ## Returns the column after the rightmost column in the specified row
    Set Gather_ColForAppend = rng_Left(1, 1).Offset(0, rng_Left.Worksheet.Columns.Count - rng_Left.Column).End(xlToLeft).Offset(0, 1)
End Function

Public Function Gather_Dictionary(fld_Start As DictParams, Optional bln_ConvertToString As Boolean = True, _
    Optional bln_CaseSensitive As Boolean = False) As Dictionary
    ' ## From the start positions defined, add lines to dictionary and work one cell down at a time until first blank key

    Dim dic_Output As Dictionary: Set dic_Output = New Dictionary
    If bln_CaseSensitive = False Then dic_Output.CompareMode = CompareMethod.TextCompare
    Dim rng_ActiveKey As Range: Set rng_ActiveKey = fld_Start.KeysTopLeft
    Dim rng_ActiveValue As Range: Set rng_ActiveValue = fld_Start.ValuesTopLeft

    While rng_ActiveKey.Value <> ""
        ' Ignore duplicates, use first definition in the query output from Murex
        If dic_Output.Exists(rng_ActiveKey.Value) = False Then
            If bln_ConvertToString = True Then
                Call dic_Output.Add(CStr(rng_ActiveKey.Value), CStr(rng_ActiveValue.Value))
            Else
                Call dic_Output.Add(rng_ActiveKey.Value, rng_ActiveValue.Value)
            End If
        End If

        Set rng_ActiveKey = rng_ActiveKey.Offset(1, 0)
        Set rng_ActiveValue = rng_ActiveValue.Offset(1, 0)
    Wend

    Set Gather_Dictionary = dic_Output
End Function

Public Function Gather_SheetNames(wbk_Input As Workbook, Optional str_Prefix As String = "-", _
    Optional bln_IncludePrefix As Boolean = True) As Collection
    ' ## List the sheet names in a workbook.  Can optionally filter by a prefix
    Dim strLst_Output As New Collection
    Dim int_NumSheets As Integer: int_NumSheets = wbk_Input.Worksheets.Count
    Dim int_Ctr As Integer

    If str_Prefix = "-" Then
        ' Return all names
        For int_Ctr = 1 To int_NumSheets
            Call strLst_Output.Add(wbk_Input.Worksheets(int_Ctr).Name)
        Next int_Ctr
    Else
        ' Return filtered names
        Dim int_PrefixLength As Integer: int_PrefixLength = Len(str_Prefix)
        Dim str_ActiveName As String
        For int_Ctr = 1 To int_NumSheets
            str_ActiveName = wbk_Input.Worksheets(int_Ctr).Name
            If Left(str_ActiveName, int_PrefixLength) = str_Prefix Then
                If bln_IncludePrefix = True Then
                    Call strLst_Output.Add(str_ActiveName)
                Else
                    Call strLst_Output.Add(Right(str_ActiveName, Len(str_ActiveName) - int_PrefixLength))
                End If
            End If
        Next int_Ctr
    End If

    Set Gather_SheetNames = strLst_Output
End Function

Public Function Gather_RangeBelow(rng_Top As Range) As Range
    ' ## Returns a range in the specified column, starting with the specified cell and continuing to the last populated cell in the column
    ' ## If the specified cell is blank and so are all the rows beneath, returns null
    Dim lng_NumRows As Long: lng_NumRows = Examine_NumRows(rng_Top)
    If lng_NumRows > 0 Then Set Gather_RangeBelow = rng_Top.Resize(lng_NumRows, 1)
End Function

Public Function Gather_CurrencyFormat() As String
    Gather_CurrencyFormat = "$#,##0.00;[Red]-$#,##0.00"
End Function

Public Function Gather_DateFormat() As String
    Gather_DateFormat = "dd/mm/yyyy"
End Function

Public Function Gather_FloatNumFormat() As String
    Gather_FloatNumFormat = "General"
End Function

Public Function Gather_AdjacentPillars(var_RefPoint As Variant, varLst_Pillars As Collection) As Dictionary
    ' ## Return pillar on or immediately earlier/later than the specified reference point
    ' ## Pillar collection should be sorted in ascending order
    ' ## If specified date is earlier than first pillar, returns first pillar.  If later than last pillar, returns last pillar

    Dim dic_Output As Dictionary: Set dic_Output = New Dictionary
    dic_Output.CompareMode = CompareMethod.TextCompare

    Dim var_ActivePillar As Variant
    Dim int_NumPillars As Integer: int_NumPillars = varLst_Pillars.Count
    Dim int_Ctr As Integer

    Dim int_NumPillarsBelowRef As Integer: int_NumPillarsBelowRef = Examine_CountNumBelow(varLst_Pillars, var_RefPoint, True)
    If int_NumPillarsBelowRef = 0 Then
        ' A single pillar or reference point below the lowest pillar
        Call dic_Output.Add("Pillar_Below", varLst_Pillars(1))
        Call dic_Output.Add("Pillar_Above", varLst_Pillars(1))
        Call dic_Output.Add("Index_Below", 1)
        Call dic_Output.Add("Index_Above", 1)
    ElseIf int_NumPillarsBelowRef = int_NumPillars Then
        ' Reference point above the highest pillar
        Call dic_Output.Add("Pillar_Below", varLst_Pillars(int_NumPillars))
        Call dic_Output.Add("Pillar_Above", varLst_Pillars(int_NumPillars))
        Call dic_Output.Add("Index_Below", int_NumPillars)
        Call dic_Output.Add("Index_Above", int_NumPillars)
    Else
        ' Reference within set of pillars
        Call dic_Output.Add("Pillar_Below", varLst_Pillars(int_NumPillarsBelowRef))
        Call dic_Output.Add("Pillar_Above", varLst_Pillars(int_NumPillarsBelowRef + 1))
        Call dic_Output.Add("Index_Below", int_NumPillarsBelowRef)
        Call dic_Output.Add("Index_Above", int_NumPillarsBelowRef + 1)
    End If

    ' Return index of the pillars and the values of the pillars
    Set Gather_AdjacentPillars = dic_Output
End Function

Public Function Gather_ApplicationState(enu_Type As ApplicationStateType) As ApplicationState
    ' ## Return various settings of the application at that point in time
    Dim fld_Output As ApplicationState
    With Application
        Select Case enu_Type
            Case ApplicationStateType.Current
                fld_Output.ScreenUpdating = .ScreenUpdating
                fld_Output.CalculationMode = .Calculation
                fld_Output.EventsEnabled = .EnableEvents
                fld_Output.DisplayAlerts = .DisplayAlerts
                fld_Output.StatusBarMsg = False
            Case ApplicationStateType.Optimized
                fld_Output.ScreenUpdating = False
                fld_Output.CalculationMode = xlCalculationManual
                fld_Output.EventsEnabled = False
                fld_Output.DisplayAlerts = False

                If .StatusBar = False Then
                    fld_Output.StatusBarMsg = "Please wait..."
                Else
                    fld_Output.StatusBarMsg = .StatusBar
                End If
        End Select
    End With

    Gather_ApplicationState = fld_Output
End Function

Public Function Gather_CopyList(varLst_Input As Collection) As Collection
    ' ## Duplicate the input list
    Dim varLst_Output As New Collection
    Dim int_Ctr As Integer
    For int_Ctr = 1 To varLst_Input.Count
        Call varLst_Output.Add(varLst_Input(int_Ctr))
    Next int_Ctr
    Set Gather_CopyList = varLst_Output
End Function

Public Function Gather_CopyLngArr(lngArr_Input() As Long) As Long()
    ' ## Duplicate the input list
    Dim int_LBound As Integer: int_LBound = LBound(lngArr_Input)
    Dim int_UBound As Integer: int_UBound = UBound(lngArr_Input)
    Dim lngArr_Output() As Long: ReDim lngArr_Output(int_LBound To int_UBound) As Long

    Dim int_Ctr As Integer
    For int_Ctr = int_LBound To int_UBound
        lngArr_Output(int_Ctr) = lngArr_Input(int_Ctr)
    Next int_Ctr
    Gather_CopyLngArr = lngArr_Output
End Function

Public Function Gather_Hierarchy_TopDown(strArr_AllParents() As Variant, strArr_AllChildren() As Variant) As Dictionary
    ' ## Return dictionary containing mapping from parent to a list of its children
    ' ## Input arrays are 2D column vectors
    Dim dic_Output As New Dictionary
    Dim lng_NumRows As Long: lng_NumRows = UBound(strArr_AllParents, 1)
    Debug.Assert UBound(strArr_AllChildren, 1) = lng_NumRows
    Dim lng_RowCtr As Long, str_ActiveParent As String, str_ActiveChild As String, strLst_ActiveChildren As Collection

    For lng_RowCtr = 1 To lng_NumRows
        str_ActiveParent = strArr_AllParents(lng_RowCtr, 1)
        str_ActiveChild = strArr_AllChildren(lng_RowCtr, 1)

        ' Obtain list of children for the active parent
        If dic_Output.Exists(str_ActiveParent) Then
            Set strLst_ActiveChildren = dic_Output(str_ActiveParent)
        Else
            Set strLst_ActiveChildren = New Collection
            Call dic_Output.Add(str_ActiveParent, strLst_ActiveChildren)
        End If

        ' Add active child to the list
        Call strLst_ActiveChildren.Add(str_ActiveChild)
    Next lng_RowCtr

    Set Gather_Hierarchy_TopDown = dic_Output
End Function

Public Function Gather_AllDescendants(dic_Hierarchy_TopDown As Dictionary, dic_BaseItems As Dictionary, str_Parent As String, _
    Optional ByRef dic_KnownDescendants As Dictionary = Nothing) As Collection
    ' ## Return a list of all the items in the set of base items which are direct or indirect descendants of the parent node
    Dim strLst_Output As New Collection
    If dic_KnownDescendants Is Nothing Then
        ' Cache stores the parent vs the list of descendants
        Set dic_KnownDescendants = New Dictionary
        dic_KnownDescendants.CompareMode = CompareMethod.TextCompare
    End If

    If dic_KnownDescendants.Exists(str_Parent) Then
        Set strLst_Output = dic_KnownDescendants(str_Parent)
    ElseIf dic_BaseItems.Exists(str_Parent) Then
        Call strLst_Output.Add(str_Parent)
    Else
        ' Check there are actually descendants of the current parent
        If dic_Hierarchy_TopDown.Exists(str_Parent) Then
            Dim strLst_DirectDescs As Collection: Set strLst_DirectDescs = dic_Hierarchy_TopDown(str_Parent)
            Dim lng_Ctr As Long, str_ActiveDesc As String

            For lng_Ctr = 1 To strLst_DirectDescs.Count
                str_ActiveDesc = strLst_DirectDescs(lng_Ctr)

                If dic_BaseItems.Exists(str_ActiveDesc) = True Then
                    Call strLst_Output.Add(str_ActiveDesc)
                Else
                    ' If not a base item, try checking the descendants of this descendant.  Will add any new items to the list
                    Set strLst_Output = Convert_MergeLists(strLst_Output, Gather_AllDescendants(dic_Hierarchy_TopDown, _
                        dic_BaseItems, str_ActiveDesc, dic_KnownDescendants))
                End If
            Next lng_Ctr
        End If

        ' Cache the descendents of this node
        Call dic_KnownDescendants.Add(str_Parent, strLst_Output)
    End If

    Set Gather_AllDescendants = strLst_Output
End Function

Public Function Gather_TopParent_Range(rng_Children As Range, rng_Parents As Range, str_Child As String) As String
    ' ## Returns the highest ancestor of the specified child
    ' ## Assumes the range for the children is a column vector which contains unique values and no spaces in between rows
    Dim fld_Params As DictParams
    Set fld_Params.KeysTopLeft = rng_Children(1, 1)
    Set fld_Params.ValuesTopLeft = rng_Parents(1, 1)
    Dim dic_ChildToParent As Dictionary: Set dic_ChildToParent = Gather_Dictionary(fld_Params)

    Gather_TopParent_Range = Gather_TopParent(dic_ChildToParent, str_Child)
End Function

Public Function Gather_TopParent(dic_ChildToParent As Dictionary, str_Child As String) As String
    ' ## Returns the highest ancestor of the specified child
    ' ## Assumes the range for the children is a column vector which contains unique values and no spaces in between rows
    Dim str_ActiveChild As String: str_ActiveChild = str_Child

    While dic_ChildToParent.Exists(str_ActiveChild)
        str_ActiveChild = dic_ChildToParent(str_ActiveChild)
    Wend

    Gather_TopParent = str_ActiveChild
End Function

Public Function Gather_LibraryPath() As String
    Gather_LibraryPath = ThisWorkbook.FullName
End Function