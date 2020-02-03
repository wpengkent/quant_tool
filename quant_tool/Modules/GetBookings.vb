Option Explicit

' ## Functions to return instrument booking information.  Specifies the top range in the booking sheet for each attribute
Public Function GetBookings_BND() As Dictionary
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    Dim wks_Booking As Worksheet: Set wks_Booking = ThisWorkbook.Worksheets("INST_Bond")
    With dic_output
        Call .Add(BookingAttribute.Sheet, wks_Booking)
        Call .Add(BookingAttribute.IDSelection, wks_Booking.Range("A3").Resize(1, 4))
        Call .Add(BookingAttribute.Params, wks_Booking.Range("E3:AQ3"))
        Call .Add(BookingAttribute.Outputs, wks_Booking.Range("AT3").Resize(1, 8))
        Call .Add(BookingAttribute.BaseChg, wks_Booking.Range("BC3").Resize(1, 2))
    End With

    Set GetBookings_BND = dic_output
End Function
Public Function GetBookings_NID() As Dictionary  'QJK code 03102014
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    Dim wks_Booking As Worksheet: Set wks_Booking = ThisWorkbook.Worksheets("INST_NID")
    With dic_output
        Call .Add(BookingAttribute.Sheet, wks_Booking)
        Call .Add(BookingAttribute.IDSelection, wks_Booking.Range("A3").Resize(1, 4))
        Call .Add(BookingAttribute.Params, wks_Booking.Range("E3:AQ3"))
        Call .Add(BookingAttribute.Outputs, wks_Booking.Range("AT3").Resize(1, 8))
        Call .Add(BookingAttribute.BaseChg, wks_Booking.Range("BC3").Resize(1, 2))
    End With

    Set GetBookings_NID = dic_output
End Function
Public Function GetBookings_BA() As Dictionary
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    Dim wks_Booking As Worksheet: Set wks_Booking = ThisWorkbook.Worksheets("INST_BA")
    With dic_output
        Call .Add(BookingAttribute.Sheet, wks_Booking)
        Call .Add(BookingAttribute.IDSelection, wks_Booking.Range("A3").Resize(1, 4))
        Call .Add(BookingAttribute.Params, wks_Booking.Range("E3:AH3"))
        Call .Add(BookingAttribute.Outputs, wks_Booking.Range("AK3").Resize(1, 8))
        Call .Add(BookingAttribute.BaseChg, wks_Booking.Range("AT3").Resize(1, 2))
    End With

    Set GetBookings_BA = dic_output
End Function
Public Function GetBookings_ECS() As Dictionary
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    Dim wks_Booking As Worksheet: Set wks_Booking = ThisWorkbook.Worksheets("INST_EQCash")
    With dic_output
        Call .Add(BookingAttribute.Sheet, wks_Booking)
        Call .Add(BookingAttribute.IDSelection, wks_Booking.Range("A3").Resize(1, 4))
        Call .Add(BookingAttribute.Params, wks_Booking.Range("E3:R3"))
        Call .Add(BookingAttribute.Outputs, wks_Booking.Range("U3").Resize(1, 4))
        Call .Add(BookingAttribute.BaseChg, wks_Booking.Range("Z3").Resize(1, 2))
    End With

    Set GetBookings_ECS = dic_output
End Function


Public Function GetBookings_EQO() As Dictionary
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    Dim wks_Booking As Worksheet: Set wks_Booking = ThisWorkbook.Worksheets("INST_EQOptions")
    With dic_output
        Call .Add(BookingAttribute.Sheet, wks_Booking)
        Call .Add(BookingAttribute.IDSelection, wks_Booking.Range("A3").Resize(1, 4))
        Call .Add(BookingAttribute.Params, wks_Booking.Range("E3:AM3"))
        Call .Add(BookingAttribute.Outputs, wks_Booking.Range("AP3").Resize(1, 4))
        Call .Add(BookingAttribute.BaseChg, wks_Booking.Range("AU3").Resize(1, 2))
    End With

    Set GetBookings_EQO = dic_output
End Function

Public Function GetBookings_EQF() As Dictionary
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    Dim wks_Booking As Worksheet: Set wks_Booking = ThisWorkbook.Worksheets("INST_EQFut")
    With dic_output
        Call .Add(BookingAttribute.Sheet, wks_Booking)
        Call .Add(BookingAttribute.IDSelection, wks_Booking.Range("A3").Resize(1, 4))
        Call .Add(BookingAttribute.Params, wks_Booking.Range("E3:U3"))
        Call .Add(BookingAttribute.Outputs, wks_Booking.Range("X3").Resize(1, 4))
        Call .Add(BookingAttribute.BaseChg, wks_Booking.Range("AC3").Resize(1, 2))
    End With

    Set GetBookings_EQF = dic_output
End Function
Public Function GetBookings_EQS() As Dictionary
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    Dim wks_Booking As Worksheet: Set wks_Booking = ThisWorkbook.Worksheets("INST_EQSwap")
    With dic_output
        Call .Add(BookingAttribute.Sheet, wks_Booking)
        Call .Add(BookingAttribute.IDSelection, wks_Booking.Range("A3").Resize(1, 4))
        Call .Add(BookingAttribute.Params, wks_Booking.Range("E3:AC3"))
        Call .Add(BookingAttribute.Outputs, wks_Booking.Range("AE3").Resize(1, 4))
        Call .Add(BookingAttribute.BaseChg, wks_Booking.Range("AJ3").Resize(1, 2))
    End With

    Set GetBookings_EQS = dic_output
End Function
Public Function GetBookings_IRS() As Dictionary
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    Dim wks_Booking As Worksheet: Set wks_Booking = ThisWorkbook.Worksheets("INST_IRS")
    With dic_output
        Call .Add(BookingAttribute.Sheet, wks_Booking)
        Call .Add(BookingAttribute.IDSelection, wks_Booking.Range("A3").Resize(1, 4))
        Call .Add(BookingAttribute.Params, wks_Booking.Range("E3:AD3"))
        Call .Add(BookingAttribute.Outputs, wks_Booking.Range("AF3").Resize(1, 5))
        Call .Add(BookingAttribute.BaseChg, wks_Booking.Range("AL3").Resize(1, 2))
    End With

    Set GetBookings_IRS = dic_output
End Function

Public Function GetBookings_RngAcc() As Dictionary
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    Dim wks_Booking As Worksheet: Set wks_Booking = ThisWorkbook.Worksheets("INST_RngAcc")
    With dic_output
        Call .Add(BookingAttribute.Sheet, wks_Booking)
        Call .Add(BookingAttribute.IDSelection, wks_Booking.Range("A3").Resize(1, 4))
        Call .Add(BookingAttribute.Params, wks_Booking.Range("E3:BY3"))
        Call .Add(BookingAttribute.Outputs, wks_Booking.Range("CA3").Resize(1, 6))
        Call .Add(BookingAttribute.BaseChg, wks_Booking.Range("CH3").Resize(1, 2))
    End With

    Set GetBookings_RngAcc = dic_output
End Function

Public Function GetBookings_CFL() As Dictionary
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    Dim wks_Booking As Worksheet: Set wks_Booking = ThisWorkbook.Worksheets("INST_CapFloor")
    With dic_output
        Call .Add(BookingAttribute.Sheet, wks_Booking)
        Call .Add(BookingAttribute.IDSelection, wks_Booking.Range("A3").Resize(1, 4))
        Call .Add(BookingAttribute.Params, wks_Booking.Range("E3:X3"))
        Call .Add(BookingAttribute.Outputs, wks_Booking.Range("Z3").Resize(1, 6))
        Call .Add(BookingAttribute.BaseChg, wks_Booking.Range("AG3").Resize(1, 2))
    End With

    Set GetBookings_CFL = dic_output
End Function

Public Function GetBookings_SWT() As Dictionary
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    Dim wks_Booking As Worksheet: Set wks_Booking = ThisWorkbook.Worksheets("INST_Swaption")
    With dic_output
        Call .Add(BookingAttribute.Sheet, wks_Booking)
        Call .Add(BookingAttribute.IDSelection, wks_Booking.Range("A3").Resize(1, 4))
        Call .Add(BookingAttribute.Params, wks_Booking.Range("E3:AK3"))
        Call .Add(BookingAttribute.Outputs, wks_Booking.Range("AM3").Resize(1, 6))
        Call .Add(BookingAttribute.BaseChg, wks_Booking.Range("AT3").Resize(1, 2))
    End With

    Set GetBookings_SWT = dic_output
End Function


Public Function GetBookings_DEP() As Dictionary
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    Dim wks_Booking As Worksheet: Set wks_Booking = ThisWorkbook.Worksheets("INST_SimpleDeposit")
    With dic_output
        Call .Add(BookingAttribute.Sheet, wks_Booking)
        Call .Add(BookingAttribute.IDSelection, wks_Booking.Range("A3").Resize(1, 4))
        Call .Add(BookingAttribute.Params, wks_Booking.Range("E3:O3"))
        Call .Add(BookingAttribute.Outputs, wks_Booking.Range("R3").Resize(1, 5))
        Call .Add(BookingAttribute.BaseChg, wks_Booking.Range("X3").Resize(1, 2))
    End With

    Set GetBookings_DEP = dic_output
End Function

Public Function GetBookings_FRA() As Dictionary
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    Dim wks_Booking As Worksheet: Set wks_Booking = ThisWorkbook.Worksheets("INST_FRA")
    With dic_output
        Call .Add(BookingAttribute.Sheet, wks_Booking)
        Call .Add(BookingAttribute.IDSelection, wks_Booking.Range("A3").Resize(1, 4))
        Call .Add(BookingAttribute.Params, wks_Booking.Range("E3:L3"))
        Call .Add(BookingAttribute.Outputs, wks_Booking.Range("O3").Resize(1, 5))
        Call .Add(BookingAttribute.BaseChg, wks_Booking.Range("U3").Resize(1, 2))
    End With

    Set GetBookings_FRA = dic_output
End Function

Public Function GetBookings_FTB() As Dictionary
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    Dim wks_Booking As Worksheet: Set wks_Booking = ThisWorkbook.Worksheets("INST_Fut_Bill")
    With dic_output
        Call .Add(BookingAttribute.Sheet, wks_Booking)
        Call .Add(BookingAttribute.IDSelection, wks_Booking.Range("A3").Resize(1, 4))
        Call .Add(BookingAttribute.Params, wks_Booking.Range("E3:U3"))
        Call .Add(BookingAttribute.Outputs, wks_Booking.Range("W3").Resize(1, 3))
        Call .Add(BookingAttribute.BaseChg, wks_Booking.Range("AA3").Resize(1, 2))
    End With

    Set GetBookings_FTB = dic_output
End Function

Public Function GetBookings_FXF() As Dictionary
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    Dim wks_Booking As Worksheet: Set wks_Booking = ThisWorkbook.Worksheets("INST_FXFwd")
    With dic_output
        Call .Add(BookingAttribute.Sheet, wks_Booking)
        Call .Add(BookingAttribute.IDSelection, wks_Booking.Range("A3").Resize(1, 4))
        Call .Add(BookingAttribute.Params, wks_Booking.Range("E3:N3"))
        Call .Add(BookingAttribute.Outputs, wks_Booking.Range("Q3").Resize(1, 5))
        Call .Add(BookingAttribute.BaseChg, wks_Booking.Range("W3").Resize(1, 2))
    End With

    Set GetBookings_FXF = dic_output
End Function

Public Function GetBookings_FVN() As Dictionary
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    Dim wks_Booking As Worksheet: Set wks_Booking = ThisWorkbook.Worksheets("INST_FXVanilla")
    With dic_output
        Call .Add(BookingAttribute.Sheet, wks_Booking)
        Call .Add(BookingAttribute.IDSelection, wks_Booking.Range("A3").Resize(1, 4))
        Call .Add(BookingAttribute.Params, wks_Booking.Range("E3:AA3"))
        Call .Add(BookingAttribute.Outputs, wks_Booking.Range("AE3").Resize(1, 8))
        Call .Add(BookingAttribute.BaseChg, wks_Booking.Range("AN3").Resize(1, 2))
    End With

    Set GetBookings_FVN = dic_output
End Function

Public Function GetBookings_FBR() As Dictionary
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    Dim wks_Booking As Worksheet: Set wks_Booking = ThisWorkbook.Worksheets("INST_FXBarrier")
    With dic_output
        Call .Add(BookingAttribute.Sheet, wks_Booking)
        Call .Add(BookingAttribute.IDSelection, wks_Booking.Range("A3").Resize(1, 4))
        Call .Add(BookingAttribute.Params, wks_Booking.Range("E3:AD3"))
        Call .Add(BookingAttribute.Outputs, wks_Booking.Range("AG3").Resize(1, 8))
        Call .Add(BookingAttribute.BaseChg, wks_Booking.Range("AP3").Resize(1, 2))
    End With

    Set GetBookings_FBR = dic_output
End Function

Public Function GetBookings_FBN() As Dictionary
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    Dim wks_Booking As Worksheet: Set wks_Booking = ThisWorkbook.Worksheets("INST_Fut_Bond")
    With dic_output
        Call .Add(BookingAttribute.Sheet, wks_Booking)
        Call .Add(BookingAttribute.IDSelection, wks_Booking.Range("A3").Resize(1, 4))
        Call .Add(BookingAttribute.Params, wks_Booking.Range("E3:T3"))
        Call .Add(BookingAttribute.Outputs, wks_Booking.Range("V3").Resize(1, 3))
        Call .Add(BookingAttribute.BaseChg, wks_Booking.Range("Z3").Resize(1, 2))
    End With

    Set GetBookings_FBN = dic_output
End Function

Public Function GetBookings_FRE() As Dictionary
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    Dim wks_Booking As Worksheet: Set wks_Booking = ThisWorkbook.Worksheets("INST_FXRebate")
    With dic_output
        Call .Add(BookingAttribute.Sheet, wks_Booking)
        Call .Add(BookingAttribute.IDSelection, wks_Booking.Range("A3").Resize(1, 4))
        Call .Add(BookingAttribute.Params, wks_Booking.Range("E3:Z3"))
        Call .Add(BookingAttribute.Outputs, wks_Booking.Range("AC3").Resize(1, 6))
        Call .Add(BookingAttribute.BaseChg, wks_Booking.Range("AJ3").Resize(1, 2))
    End With

    Set GetBookings_FRE = dic_output
End Function


Public Function GetBookings_FXFut() As Dictionary
    Dim dic_output As Dictionary: Set dic_output = New Dictionary
    Dim wks_Booking As Worksheet: Set wks_Booking = ThisWorkbook.Worksheets("INST_FXFut")
    With dic_output
        Call .Add(BookingAttribute.Sheet, wks_Booking)
        Call .Add(BookingAttribute.IDSelection, wks_Booking.Range("A3").Resize(1, 4))
        Call .Add(BookingAttribute.Params, wks_Booking.Range("E3:V3"))
        Call .Add(BookingAttribute.Outputs, wks_Booking.Range("Y3").Resize(1, 5))
        Call .Add(BookingAttribute.BaseChg, wks_Booking.Range("AE3").Resize(1, 2))
    End With

    Set GetBookings_FXFut = dic_output
End Function