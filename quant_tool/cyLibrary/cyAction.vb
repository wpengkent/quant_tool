Option Explicit

Public Sub Action_Query_Access(str_DBPath As String, str_SQL As String, rng_OutputTopLeft As Range, _
    Optional bln_WithHeaders As Boolean = False)

    ' ## Run specified SQL code on specified Access database and output query to specified range
    Dim fld_AppState_Orig As ApplicationState: fld_AppState_Orig = Gather_ApplicationState(ApplicationStateType.Current)
    Dim fld_AppState_Opt As ApplicationState: fld_AppState_Opt = Gather_ApplicationState(ApplicationStateType.Optimized)
    Call Action_SetAppState(fld_AppState_Opt)

    Dim dbs_Source As DAO.Database
    Dim rst_Results As DAO.Recordset

    ' Print recordset
    Set dbs_Source = OpenDatabase(str_DBPath)
    Set rst_Results = dbs_Source.OpenRecordset(str_SQL)

    ' Optionally extract headers
    If bln_WithHeaders = True Then
        Dim int_ColCtr As Integer
        For int_ColCtr = 1 To rst_Results.Fields.Count
            rng_OutputTopLeft(1, int_ColCtr).Value = rst_Results.Fields(int_ColCtr - 1).Name
        Next int_ColCtr

        Set rng_OutputTopLeft = rng_OutputTopLeft.Offset(1, 0)
    End If

    Call rng_OutputTopLeft.CopyFromRecordset(rst_Results)

    rst_Results.Close
    dbs_Source.Close

    Call Action_SetAppState(fld_AppState_Orig)
End Sub

Public Sub Action_Query_Access_TableNames(str_DBPath As String, rng_OutputTopLeft As Range)
    ' ## Obtain table names from the specified database
    Dim fld_AppState_Orig As ApplicationState: fld_AppState_Orig = Gather_ApplicationState(ApplicationStateType.Current)
    Dim fld_AppState_Opt As ApplicationState: fld_AppState_Opt = Gather_ApplicationState(ApplicationStateType.Optimized)
    Call Action_SetAppState(fld_AppState_Opt)

    Dim dbs_Source As DAO.Database
    Dim tdf_Active As DAO.TableDef

    ' Gather table names
    Set dbs_Source = OpenDatabase(str_DBPath)

    Dim int_NumTables As Integer: int_NumTables = dbs_Source.TableDefs.Count
    Dim strArr_Output() As String: ReDim strArr_Output(1 To int_NumTables, 1 To 1) As String
    Dim int_Ctr As Integer
    For int_Ctr = 1 To int_NumTables
        Set tdf_Active = dbs_Source.TableDefs(int_Ctr - 1)
        strArr_Output(int_Ctr, 1) = tdf_Active.Name
    Next int_Ctr

    dbs_Source.Close

    ' Output to Excel
    rng_OutputTopLeft.Resize(int_NumTables, 1).Value = strArr_Output

    Call Action_SetAppState(fld_AppState_Orig)
End Sub

Public Sub Action_Query_AccessADO(str_DBPath As String, str_SQL As String, rng_OutputTopLeft As Range, _
    Optional bln_WithHeaders As Boolean = False)

    ' ## Run specified SQL code on specified Access database and output query to specified range
    Dim fld_AppState_Orig As ApplicationState: fld_AppState_Orig = Gather_ApplicationState(ApplicationStateType.Current)
    Dim fld_AppState_Opt As ApplicationState: fld_AppState_Opt = Gather_ApplicationState(ApplicationStateType.Optimized)
    Call Action_SetAppState(fld_AppState_Opt)

    Dim rst_Results As New ADODB.Recordset

    ' Build connection string
    Dim str_Conn As String: str_Conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & str_DBPath

    ' Print recordset
    Call rst_Results.Open(str_SQL, str_Conn)

    ' Optionally extract headers
    If bln_WithHeaders = True Then
        Dim int_ColCtr As Integer
        For int_ColCtr = 1 To rst_Results.Fields.Count
            rng_OutputTopLeft(1, int_ColCtr).Value = rst_Results.Fields(int_ColCtr - 1).Name
        Next int_ColCtr

        Set rng_OutputTopLeft = rng_OutputTopLeft.Offset(1, 0)
    End If

    Call rng_OutputTopLeft.CopyFromRecordset(rst_Results)

    rst_Results.Close

    Call Action_SetAppState(fld_AppState_Orig)
End Sub

Public Sub Action_Query_Sybase(sbp_Params As SybaseParams, str_Database As String, str_SQL As String, rng_Output As Range)
    ' ## Run specified SQL code on specified Sybase database and output query to specified range

    Dim rst_Results As ADODB.Recordset: Set rst_Results = Gather_RecordSet_Sybase(sbp_Params, str_Database, str_SQL)
    Call rng_Output.CopyFromRecordset(rst_Results)  ' Output to range
    rst_Results.Close
End Sub

Public Sub Action_ClearBelow(rng_TopLeft As Range, int_NumCols As Integer)
    ' ## Clear all cells in the column at or below the start
    Dim lng_NumRows As Long
    lng_NumRows = Examine_NumRows(rng_TopLeft)

    If lng_NumRows > 0 Then
        rng_TopLeft.Resize(lng_NumRows, int_NumCols).ClearContents
    End If
End Sub

Public Sub Action_DeleteBelow(ByRef rng_TopLeft As Range, int_NumCols As Integer)
    ' ## Delete all cells in the column at or below the start
    Dim lng_NumRows As Long: lng_NumRows = Examine_NumRows(rng_TopLeft)
    Dim bln_Protected As Boolean: bln_Protected = rng_TopLeft.Worksheet.ProtectContents = True
    Dim wks_Parent As Worksheet: Set wks_Parent = rng_TopLeft.Worksheet
    Dim str_Address As String: str_Address = rng_TopLeft.Address

    If lng_NumRows > 0 Then
        If bln_Protected = True Then wks_Parent.Unprotect
        rng_TopLeft.Resize(lng_NumRows, int_NumCols).Delete Shift:=xlUp
        Set rng_TopLeft = wks_Parent.Range(str_Address)
        If bln_Protected = True Then Call wks_Parent.Protect
    End If
End Sub

Public Sub Action_UpdateNamedRange(str_Name As String, rng_Top As Range)
    ' ## Based on the specified top cell, find the last used cell in the column and set the range between the top and bottom to have the specified range name
    On Error GoTo errHandler

    ' Update range for containers
    Dim lng_NumRows As Long: lng_NumRows = Examine_NumRows(rng_Top)

    ' Remove existing reference to this range name
    Dim wbk_Book As Workbook: Set wbk_Book = rng_Top.Worksheet.Parent
    wbk_Book.Names(str_Name).Delete

    If lng_NumRows = 0 Then
        wbk_Book.Names.Add Name:=str_Name, RefersTo:=rng_Top
    Else
        wbk_Book.Names.Add Name:=str_Name, RefersTo:=rng_Top.Resize(lng_NumRows, 1)
    End If

errHandler:
    Select Case Err
        Case 1004
            ' Don't worry about deleting range if it didn't already exist
            Resume Next
    End Select
End Sub

Public Sub Action_OutputArray(arr_Values As Variant, rng_OutputTopLeft As Range, Optional bln_TransposeToHorizontal As Boolean = False)
    ' ## Output contents of the column array to the specified location
    Dim int_Ctr As Integer
    Dim lng_NumVals As Long: lng_NumVals = UBound(arr_Values) - LBound(arr_Values) + 1
    Dim varArr_Values2D() As Variant: varArr_Values2D = Convert_Array1DTo2D(arr_Values, bln_TransposeToHorizontal)

    If bln_TransposeToHorizontal = True Then
        rng_OutputTopLeft.Resize(1, lng_NumVals).Value = varArr_Values2D
    Else
        rng_OutputTopLeft.Resize(lng_NumVals, 1).Value = varArr_Values2D
    End If
End Sub

Public Sub Action_SetAppState(fld_State As ApplicationState)
    ' ## Return application to the settings specified
    With Application
        .ScreenUpdating = fld_State.ScreenUpdating
        .Calculation = fld_State.CalculationMode
        .EnableEvents = fld_State.EventsEnabled
        .DisplayAlerts = fld_State.DisplayAlerts
        .StatusBar = fld_State.StatusBarMsg
    End With
End Sub

Public Sub Action_RemoveAllSheets(wbk_Target As Workbook, str_Prefix As String, Optional str_Warning As String = "")
    ' ## Remove all sheets which have names beginning with the specified prefix
    Dim bln_Proceed As Boolean: bln_Proceed = True

    ' Show warning if it exists
    If Len(str_Warning) > 0 Then
        Dim int_Result As Integer: int_Result = MsgBox(str_Warning, vbOKCancel)
        If int_Result <> vbOK Then bln_Proceed = False
    End If

    If bln_Proceed = True Then
        Dim fld_AppState_Orig As ApplicationState: fld_AppState_Orig = Gather_ApplicationState(ApplicationStateType.Current)
        Dim fld_AppState_Opt As ApplicationState: fld_AppState_Opt = Gather_ApplicationState(ApplicationStateType.Optimized)
        Call Action_SetAppState(fld_AppState_Opt)

        Dim int_PrefixLength As Integer: int_PrefixLength = Len(str_Prefix)
        Dim wks_Found As Worksheet
        For Each wks_Found In wbk_Target.Worksheets
            If Left(wks_Found.Name, int_PrefixLength) = str_Prefix Then wks_Found.Delete
        Next wks_Found

        Call Action_SetAppState(fld_AppState_Orig)
    End If
End Sub

Public Sub Action_SwapValues(ByRef var_A As Variant, ByRef var_B As Variant)
    ' ## Exchange the values of A and B
    Dim var_Temp As Variant: var_Temp = var_A
    var_A = var_B
    var_B = var_Temp
End Sub

Public Sub Action_PrintMatrix(var_Matrix As Variant)
    ' ## Display a 2D array or range in immediate pane
    Dim int_RowCtr As Integer, int_ColCtr As Integer

    For int_RowCtr = Examine_LowerBoundIndex_Row(var_Matrix) To Examine_UpperBoundIndex_Row(var_Matrix)
        For int_ColCtr = Examine_LowerBoundIndex_Col(var_Matrix) To Examine_UpperBoundIndex_Col(var_Matrix) - 1
            Debug.Print Format(var_Matrix(int_RowCtr, int_ColCtr), "#,##0.0#") & " | ";  ' Semicolon suppresses new line
        Next int_ColCtr
        Debug.Print Format(var_Matrix(int_RowCtr, int_ColCtr), "#,##0.0#")
    Next int_RowCtr

    Debug.Print vbCr  ' Leave space
End Sub