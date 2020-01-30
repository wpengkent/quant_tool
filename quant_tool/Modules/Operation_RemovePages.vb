Option Explicit

Public Sub RemoveAllFXVolCurves()
    Call Action_RemoveAllSheets(ThisWorkbook, "FXV_", "Remove all FXV curves?")
End Sub

Public Sub RemoveAllRateCurves()
    Call Action_RemoveAllSheets(ThisWorkbook, "IRC_", "Remove all IRC curves?")
End Sub

Public Sub RemoveAllCapVolCurves()
    Call Action_RemoveAllSheets(ThisWorkbook, "CVL_", "Remove all CVL curves?")
End Sub

Public Sub RemoveAllSwptVolCurves()
    Call Action_RemoveAllSheets(ThisWorkbook, "SVL_", "Remove all SVL curves?")
End Sub

Public Sub RemoveAllEqsVolCurves()
    Call Action_RemoveAllSheets(ThisWorkbook, "EVL_", "Remove all EVL curves?")
End Sub

'Public Sub Action_RemoveAllSheets(str_Prefix As String, str_Warning As String)
'    Dim int_Result As Integer: int_Result = MsgBox(str_Warning, vbOKCancel)
'
'    If int_Result = vbOK Then
'        Dim bln_ScreenUpdating As Boolean: bln_ScreenUpdating = Application.ScreenUpdating
'        Application.ScreenUpdating = False
'        Application.DisplayAlerts = False
'
'        Dim int_PrefixLength As Integer: int_PrefixLength = Len(str_Prefix)
'        Dim wks_Found As Worksheet
'        For Each wks_Found In ThisWorkbook.Worksheets
'            If Left(wks_Found.Name, int_PrefixLength) = str_Prefix Then wks_Found.Delete
'        Next wks_Found
'
'        Application.DisplayAlerts = True
'        Application.ScreenUpdating = bln_ScreenUpdating
'    End If
'End Sub
