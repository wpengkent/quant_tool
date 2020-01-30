Option Explicit

' ## MEMBER DATA
Private rng_TopLeft As Range
Private dic_Cache_Shifters As Dictionary


' ## INITIALIZATION
Public Sub Initialize(wks_Input As Worksheet)
    Set rng_TopLeft = wks_Input.Range("A3")
    Set dic_Cache_Shifters = New Dictionary
End Sub


' ## METHODS - LOOKUP
Public Function Lookup_Shifter(str_name As String) As DateShifter
    ' ## Return the DateShifter object based on the specified name
    Dim shi_Output As DateShifter

    If dic_Cache_Shifters.Exists(str_name) Then
        ' Read output from cache
        Set shi_Output = dic_Cache_Shifters(str_name)
    Else
        Dim int_NumRows As Integer: int_NumRows = Examine_NumRows(rng_TopLeft)
        Dim strArr_Names() As Variant: strArr_Names = rng_TopLeft.Resize(int_NumRows, 1).Value
        Dim fld_Params As DateShifterParams, shi_Base As DateShifter
        Dim var_ParamsRead() As Variant
        Dim int_ActiveCol As Integer

        Dim int_ctr As Integer
        For int_ctr = 1 To int_NumRows
            If strArr_Names(int_ctr, 1) = str_name Then
                ' Read output from sheet
                var_ParamsRead = rng_TopLeft.Offset(int_ctr - 1, 0).Resize(1, 7).Value

                With fld_Params
                    int_ActiveCol = 1
                    .ShifterName = UCase(var_ParamsRead(1, int_ActiveCol))

                    int_ActiveCol = int_ActiveCol + 1
                    If var_ParamsRead(1, int_ActiveCol) <> "-" Then
                        Set shi_Base = Me.Lookup_Shifter(CStr(var_ParamsRead(1, int_ActiveCol)))
                    End If
                    Set .BaseShifter = shi_Base

                    int_ActiveCol = int_ActiveCol + 1
                    .Calendar = UCase(var_ParamsRead(1, int_ActiveCol))

                    int_ActiveCol = int_ActiveCol + 1
                    .DaysToShift = var_ParamsRead(1, int_ActiveCol)

                    int_ActiveCol = int_ActiveCol + 1
                    .IsBusDays = var_ParamsRead(1, int_ActiveCol)

                    int_ActiveCol = int_ActiveCol + 1
                    .BDC = UCase(var_ParamsRead(1, int_ActiveCol))

                    int_ActiveCol = int_ActiveCol + 1
                    .Algorithm = UCase(var_ParamsRead(1, int_ActiveCol))
                End With

                Set shi_Output = New DateShifter
                Call shi_Output.Initialize(fld_Params)

                ' Store in cache
                Call dic_Cache_Shifters.Add(str_name, shi_Output)
                Exit For
            End If
        Next int_ctr
    End If

    Debug.Assert (Not shi_Output Is Nothing)
'    Set shi_Output.BaseShifter = shi_Base
    Set Lookup_Shifter = shi_Output
End Function

