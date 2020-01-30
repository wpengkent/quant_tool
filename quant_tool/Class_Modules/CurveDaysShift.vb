Option Explicit

' ## MEMBER DATA
Private intLst_Days As Collection, dblLst_Shifts As Collection, enu_ShockType As ShockType


' ## INITIALIZATION
Public Function Initialize(enu_TypeInput As ShockType)
    Set intLst_Days = New Collection
    Set dblLst_Shifts = New Collection
    enu_ShockType = enu_TypeInput
End Function


' ## PROPERTIES
Public Property Get NumShifts() As Integer
    NumShifts = intLst_Days.count
End Property

Public Property Get Days_Arr() As Variant()
    Days_Arr = Convert_ListToArr2D(intLst_Days)
End Property

Public Property Get Shifts_Arr() As Variant()
    Shifts_Arr = Convert_ListToArr2D(dblLst_Shifts)
End Property

Public Property Get ShockType() As ShockType
    ShockType = enu_ShockType
End Property


' ## METHODS - WRITE SHIFTS
Public Sub AddShift(int_numdays As Integer, dbl_ShiftSize As Double)
    ' ## Add shift such that the list of shifts is always sorted by the number of days
    Dim int_NumBelow As Integer: int_NumBelow = Examine_CountNumBelow(intLst_Days, int_numdays, True)
    If int_NumBelow = 0 Then
        ' Add to start of list
        If intLst_Days.count = 0 Then
            Call intLst_Days.Add(int_numdays)
            Call dblLst_Shifts.Add(dbl_ShiftSize)
        Else
            Call intLst_Days.Add(int_numdays, , 1)
            Call dblLst_Shifts.Add(dbl_ShiftSize, , 1)
        End If
    Else
        ' Add to correct position in list
        Call intLst_Days.Add(int_numdays, , , int_NumBelow)
        Call dblLst_Shifts.Add(dbl_ShiftSize, , , int_NumBelow)
    End If
End Sub

Public Sub AddIsolatedShift(int_numdays As Integer, dbl_ShiftSize As Double)
    ' ## Add shift that only affects the specific number of days, with no effect on other pillars
    ' ## Border the shift with zero shifts to prevent other pillars receiving shifts
    If Examine_Contains(intLst_Days, int_numdays - 1) = False Then Call Me.AddShift(int_numdays - 1, 0)

    ' Check if need to replace a zero shift which is bordering the neighbouring pillar (e.g. 1D and 2D)
    Dim int_FoundIndex As Integer: int_FoundIndex = Examine_FindIndex(intLst_Days, int_numdays)
    If int_FoundIndex = -1 Then
        Call Me.AddShift(int_numdays, dbl_ShiftSize)
    Else
        Call intLst_Days.Remove(int_FoundIndex)
        Call dblLst_Shifts.Remove(int_FoundIndex)

        ' Handle all possible cases when trying to add to the correct position in the list
        If int_FoundIndex = 1 Then
            If intLst_Days.count = 0 Then
                Call intLst_Days.Add(int_numdays)
                Call dblLst_Shifts.Add(dbl_ShiftSize)
            Else
                Call intLst_Days.Add(int_numdays, , 1)
                Call dblLst_Shifts.Add(dbl_ShiftSize, , 1)
            End If
        Else
            Call intLst_Days.Add(int_numdays, , , int_FoundIndex - 1)
            Call dblLst_Shifts.Add(dbl_ShiftSize, , , int_FoundIndex - 1)
        End If
    End If

    If Examine_Contains(intLst_Days, int_numdays + 1) = False Then Call Me.AddShift(int_numdays + 1, 0)
End Sub

Public Sub AddUniformShift(dbl_ShiftSize As Double)
    ' ## Add a single shift to all maturities
    If intLst_Days.count > 0 Then
        Set intLst_Days = New Collection
        Set dblLst_Shifts = New Collection
    End If

    Call intLst_Days.Add(0)
    Call dblLst_Shifts.Add(dbl_ShiftSize)
End Sub


' ## METHODS - READ SHIFTS
Public Function ReadShift(int_numdays As Integer) As Double
    ' ## Obtain the shift size at the specified number of days
    Dim dbl_Output As Double
    If Me.NumShifts = 0 Then
        dbl_Output = 0
    Else
        dbl_Output = Interp_Lin(intLst_Days, dblLst_Shifts, int_numdays, True)
    End If

    ReadShift = dbl_Output
End Function