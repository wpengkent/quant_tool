Option Explicit

Public Function Solve_Secant(wbk_Caller As Workbook, str_XYFunction As String, dic_StaticParams As Dictionary, _
    dbl_InitialX1 As Double, dbl_InitialX2 As Double, dbl_TargetY As Double, dbl_Tolerance As Double, _
    int_MaxIterations As Integer, dbl_FallBackValue As Double, ByRef dic_SecondaryOutputs As Dictionary) As Double
    ' ## Perform the secant method to solve for the input which sets the function to the target value

    Dim dbl_Output As Double
    Dim dbl_SecantX1 As Double, dbl_SecantX2 As Double, dbl_SecantX3 As Double
    Dim dbl_SecantY1 As Double, dbl_SecantY2 As Double, dbl_SecantY3 As Double
    Call dic_SecondaryOutputs.RemoveAll

    ' Set first initial guess
    dbl_SecantX1 = dbl_InitialX1
    dbl_SecantY1 = Application.Run("'" & wbk_Caller.Name & "'!" & str_XYFunction, dbl_SecantX1, dic_StaticParams, dic_SecondaryOutputs) - dbl_TargetY

    ' Set second intitial guess
    dbl_SecantX2 = dbl_InitialX2
    dbl_SecantY2 = Application.Run("'" & wbk_Caller.Name & "'!" & str_XYFunction, dbl_SecantX2, dic_StaticParams, dic_SecondaryOutputs) - dbl_TargetY

    ' Prepare for iteration
    Dim int_IterCtr As Integer: int_IterCtr = 0
    Dim bln_Solvable As Boolean: bln_Solvable = True

    Do
        If dbl_SecantY2 - dbl_SecantY1 = 0 Then
            ' Allow greater tolerance if having difficulty solving
            If Abs(dbl_SecantY3) > (dbl_Tolerance * 100) Or int_IterCtr = 0 Then
                ' No solution even with looser tolerance
                dbl_SecantY3 = 0
                bln_Solvable = False
            Else
                ' Solved to looser tolerance
                Exit Do
            End If
        End If

        If bln_Solvable = True Then
            int_IterCtr = int_IterCtr + 1

            ' Set new guess
            dbl_SecantX3 = dbl_SecantX2 - dbl_SecantY2 * (dbl_SecantX2 - dbl_SecantX1) / (dbl_SecantY2 - dbl_SecantY1)
            dbl_SecantY3 = Application.Run("'" & wbk_Caller.Name & "'!" & str_XYFunction, dbl_SecantX3, dic_StaticParams, dic_SecondaryOutputs) - dbl_TargetY
        End If

        dbl_SecantX1 = dbl_SecantX2
        dbl_SecantY1 = dbl_SecantY2
        dbl_SecantX2 = dbl_SecantX3
        dbl_SecantY2 = dbl_SecantY3
    Loop Until Abs(dbl_SecantY3) < dbl_Tolerance Or int_IterCtr >= int_MaxIterations

    ' Output final solution if possible, otherwise output the fallback
    If bln_Solvable = True Then dbl_Output = dbl_SecantX3 Else dbl_Output = dbl_FallBackValue

    Solve_Secant = dbl_Output
    Call dic_SecondaryOutputs.Add("NumIterations", int_IterCtr)
    Call dic_SecondaryOutputs.Add("Solvable", bln_Solvable)
End Function

Public Function Solve_FixedPt(wbk_Caller As Workbook, str_XXFunction As String, dic_StaticParams As Dictionary, _
    dbl_InitialGuess As Double, dbl_Tolerance As Double, int_MaxIterations As Integer, ByRef dbl_FallBackValue As Double, _
    ByRef dic_SecondaryOutputs As Dictionary) As Double
    ' ## Perform fixed point iteration to solve for the input which sets the function to the target value

    Dim dbl_ActiveGuess As Double: dbl_ActiveGuess = dbl_InitialGuess
    Dim dbl_PrevGuess As Double
    Call dic_SecondaryOutputs.RemoveAll

    Dim int_IterCtr As Integer
    For int_IterCtr = 1 To int_MaxIterations
        dbl_PrevGuess = dbl_ActiveGuess

        ' Perform iteration steps
        dbl_ActiveGuess = Application.Run("'" & wbk_Caller.Name & "'!" & str_XXFunction, dbl_PrevGuess, dic_StaticParams, dic_SecondaryOutputs)

        ' Stop iterating if convergence criteria is met
        If Abs(dbl_ActiveGuess - dbl_PrevGuess) < dbl_Tolerance Then
            Solve_FixedPt = dbl_ActiveGuess
            Call dic_SecondaryOutputs.Add("NumIterations", int_IterCtr)
            Call dic_SecondaryOutputs.Add("Solvable", True)
            Exit Function
        End If
    Next int_IterCtr

    ' Solution not found, output the fallback
    Solve_FixedPt = dbl_FallBackValue
    Call dic_SecondaryOutputs.Add("NumIterations", int_MaxIterations)
    Call dic_SecondaryOutputs.Add("Solvable", False)
End Function

Public Function Solve_BrentDekker(wbk_Caller As Workbook, str_XYFunction As String, dic_StaticParams As Dictionary, _
    dbl_X_OrigLBound As Double, dbl_X_OrigUBound As Double, dbl_TargetY As Double, dbl_ToleranceX As Double, _
    dbl_ToleranceY As Double, int_MaxIterations As Integer, dbl_FallBackValue As Double, ByRef dic_SecondaryOutputs As Dictionary) As Double
    ' ## Use the Brent-Dekker algorithm to solve for the input which sets the function to the target value
    ' ## It contains a bisection method fallback
    ' ## The X's are the guess variables, the Y's are the distances from the target
    ' ## Y1, Y2 are in order of distance to target, Y3 and Y4 are in order of lag
    Call dic_SecondaryOutputs.RemoveAll

    ' Initialize guesses (X1 is the bound which has the value nearest to the target)
    Dim dbl_Bound_X1 As Double: dbl_Bound_X1 = dbl_X_OrigLBound
    Dim dbl_Bound_X2 As Double: dbl_Bound_X2 = dbl_X_OrigUBound

    ' Function values - target is always 0
    Dim dbl_Bound_Y2 As Double, dbl_Bound_Y1 As Double
    dbl_Bound_Y2 = Application.Run("'" & wbk_Caller.Name & "'!" & str_XYFunction, dbl_Bound_X2, dic_StaticParams, dic_SecondaryOutputs) - dbl_TargetY
    dbl_Bound_Y1 = Application.Run("'" & wbk_Caller.Name & "'!" & str_XYFunction, dbl_Bound_X1, dic_StaticParams, dic_SecondaryOutputs) - dbl_TargetY

    ' Check if inputs have opposite sign and therefore contain the solution
    If (dbl_Bound_Y2 * dbl_Bound_Y1 >= 0) Then
        Solve_BrentDekker = dbl_FallBackValue
        Call dic_SecondaryOutputs.Add("NumIterations", 0)
        Call dic_SecondaryOutputs.Add("Solvable", False)
        Exit Function
    End If

    ' If Y2 is closer to the target, then swap such that Y1 is closer
    If (Abs(dbl_Bound_Y2) < Abs(dbl_Bound_Y1)) Then
        Call Action_SwapValues(dbl_Bound_X2, dbl_Bound_X1)
        Call Action_SwapValues(dbl_Bound_Y2, dbl_Bound_Y1)
    End If

    ' Initialize other variables
    Dim dbl_OldBound_X3 As Double: dbl_OldBound_X3 = dbl_Bound_X2
    Dim dbl_OldBound_Y3 As Double: dbl_OldBound_Y3 = dbl_Bound_Y2
    Dim dbl_OldBound_X4 As Double
    Dim dbl_Latest_X As Double
    Dim dbl_Latest_Y As Double: dbl_Latest_Y = dbl_Bound_Y1
    Dim bln_UseBisection As Boolean
    Dim bln_BisectionOnPrevIter As Boolean: bln_BisectionOnPrevIter = True

    ' Perform iterations
    Dim int_IterCtr As Integer: int_IterCtr = 0
    While ((Abs(dbl_Bound_Y1) > dbl_ToleranceY) And (Abs(dbl_Bound_X1 - dbl_Bound_X2) > dbl_ToleranceX) And (int_IterCtr < int_MaxIterations))
        int_IterCtr = int_IterCtr + 1
        dbl_Bound_Y2 = Application.Run("'" & wbk_Caller.Name & "'!" & str_XYFunction, dbl_Bound_X2, dic_StaticParams, dic_SecondaryOutputs) - dbl_TargetY
        dbl_Bound_Y1 = Application.Run("'" & wbk_Caller.Name & "'!" & str_XYFunction, dbl_Bound_X1, dic_StaticParams, dic_SecondaryOutputs) - dbl_TargetY

        ' Determine latest guess
        If ((dbl_Bound_Y2 <> dbl_OldBound_Y3) And (dbl_Bound_Y1 <> dbl_OldBound_Y3)) Then
            ' Use inverse quadratic interpolation
            dbl_Latest_X = Interp_InvQuad(dbl_Bound_X1, dbl_Bound_X2, dbl_OldBound_X3, dbl_Bound_Y1, dbl_Bound_Y2, _
                dbl_OldBound_Y3, 0)
        Else
            ' Use linear interpolation
             dbl_Latest_X = Interp_Lin_Binary(dbl_Bound_Y1, dbl_Bound_Y2, dbl_Bound_X1, dbl_Bound_X2, 0)
        End If

       ' Determine whether bisection is needed
        bln_UseBisection = False
        If dbl_Latest_X < (3 * dbl_Bound_X2 + dbl_Bound_X1) / 4 Or dbl_Latest_X > dbl_Bound_X1 Then bln_UseBisection = True
        If bln_BisectionOnPrevIter = True And Abs(dbl_Latest_X - dbl_Bound_X1) >= Abs(dbl_Bound_X1 - dbl_OldBound_X3) / 2 Then bln_UseBisection = True
        If bln_BisectionOnPrevIter = False And Abs(dbl_Latest_X - dbl_Bound_X1) >= Abs(dbl_OldBound_X3 - dbl_OldBound_X4) / 2 Then bln_UseBisection = True
        If bln_BisectionOnPrevIter = True And Abs(dbl_Bound_X1 - dbl_OldBound_X3) <= dbl_ToleranceX Then bln_UseBisection = True
        If bln_BisectionOnPrevIter = False And Abs(dbl_OldBound_X3 - dbl_OldBound_X4) <= dbl_ToleranceX Then bln_UseBisection = True

        ' Apply bisection if required
        If bln_UseBisection = True Then
            dbl_Latest_X = (dbl_Bound_X2 + dbl_Bound_X1) / 2
            bln_BisectionOnPrevIter = True
        Else
            bln_BisectionOnPrevIter = False
        End If

        ' Evaluate distance to the target for the latest guess
        dbl_Latest_Y = Application.Run("'" & wbk_Caller.Name & "'!" & str_XYFunction, dbl_Latest_X, dic_StaticParams, dic_SecondaryOutputs) - dbl_TargetY

        ' Store relevant old bounds, then update current bounds
        dbl_OldBound_X4 = dbl_OldBound_X3
        dbl_OldBound_X3 = dbl_Bound_X1
        dbl_OldBound_Y3 = dbl_Bound_Y1
        If dbl_Bound_Y2 * dbl_Latest_Y < 0 Then
            ' Lower bound and new point straddle the solution, so update the upper bound
            dbl_Bound_X1 = dbl_Latest_X
            dbl_Bound_Y1 = dbl_Latest_Y
        Else
            ' Upper bound and new point straddle the solution, so update the upper bound
            dbl_Bound_X2 = dbl_Latest_X
            dbl_Bound_Y2 = dbl_Latest_Y
        End If

        ' Ensure Y1 is the closest to the target
        If Abs(dbl_Bound_Y2) < Abs(dbl_Bound_Y1) Then
            Call Action_SwapValues(dbl_Bound_X2, dbl_Bound_X1)
            Call Action_SwapValues(dbl_Bound_Y2, dbl_Bound_Y1)
        End If
    Wend

    ' Return solution
    If int_IterCtr >= int_MaxIterations Then
        Solve_BrentDekker = dbl_FallBackValue
        Call dic_SecondaryOutputs.Add("NumIterations", int_IterCtr)
        Call dic_SecondaryOutputs.Add("Solvable", False)
    Else
        Solve_BrentDekker = dbl_Bound_X1
        Call dic_SecondaryOutputs.Add("NumIterations", int_IterCtr)
        Call dic_SecondaryOutputs.Add("Solvable", True)
    End If
End Function