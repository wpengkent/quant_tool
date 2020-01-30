Option Explicit
' ## INTERPOLATION FUNCTIONS

Public Function Interp_Lin_Binary(var_X1 As Variant, var_X2 As Variant, var_Y1 As Variant, var_Y2 As Variant, var_XLookup As Variant) As Double
    ' ## Linear interpolation and extrapolation between two points
    If var_X1 = var_X2 Then
        ' Trivial case where point is repeated
        Interp_Lin_Binary = var_Y1
    Else
        Interp_Lin_Binary = var_Y1 + (var_Y2 - var_Y1) * (var_XLookup - var_X1) / (var_X2 - var_X1)
    End If
End Function

Public Function Interp_Lin(arr_X As Variant, arr_Y As Variant, var_LookupX As Variant, bln_FlatEnds As Boolean) As Double
    ' ## Linear interpolation using a reference set of x and y coordinates in array form, where the x coordinates are in ascending order
    ' ## FlatEnds means flat extrapolation, the alternative is linear extrapolation based on the last two points (which will fail if less than 2 points in reference set)
    Dim int_LowerBound As Integer: int_LowerBound = Examine_LowerBoundIndex(arr_X)
    Dim int_UpperBound As Integer: int_UpperBound = Examine_UpperBoundIndex(arr_X)
    Dim dbl_Output As Double, int_RowCtr As Integer

    If int_LowerBound = int_UpperBound Then
        dbl_Output = arr_Y(int_LowerBound)
    ElseIf var_LookupX <= arr_X(int_LowerBound) Then
        If bln_FlatEnds = True Then
            dbl_Output = arr_Y(int_LowerBound)
        Else
            ' Extrapolate based on continuing the line between the first two points
            dbl_Output = Interp_Lin_Binary(arr_X(int_LowerBound), arr_X(int_LowerBound + 1), arr_Y(int_LowerBound), arr_Y(int_LowerBound + 1), var_LookupX)
        End If
    ElseIf var_LookupX < arr_X(int_UpperBound) Then
        ' Interpolation
        For int_RowCtr = int_LowerBound + 1 To int_UpperBound
            If var_LookupX <= arr_X(int_RowCtr) Then
                dbl_Output = Interp_Lin_Binary(arr_X(int_RowCtr - 1), arr_X(int_RowCtr), arr_Y(int_RowCtr - 1), arr_Y(int_RowCtr), var_LookupX)
                Exit For
            End If
        Next int_RowCtr
    Else
        If bln_FlatEnds = True Then
            dbl_Output = arr_Y(int_UpperBound)
        Else
            ' Extrapolate based on continuing the line between the last two points
            dbl_Output = Interp_Lin_Binary(arr_X(int_UpperBound - 1), arr_X(int_UpperBound), arr_Y(int_UpperBound - 1), arr_Y(int_UpperBound), var_LookupX)
        End If
    End If

    Interp_Lin = dbl_Output
End Function

Public Function Interp_Lin_Range(rng_X As Range, rng_Y As Range, var_LookupX As Variant, bln_FlatEnds As Boolean, _
    Optional intLst_IndexFilters As Collection = Nothing) As Double
    ' ## Linear interpolation using Excel ranges as inputs rather than arrays
    Dim dblLst_X As Collection: Set dblLst_X = Convert_RangeToList(rng_X, intLst_IndexFilters)
    Dim dblLst_Y As Collection: Set dblLst_Y = Convert_RangeToList(rng_Y, intLst_IndexFilters)
    Interp_Lin_Range = Interp_Lin(dblLst_X, dblLst_Y, var_LookupX, bln_FlatEnds)
End Function

Public Function Interp_V2t_Binary(var_X1 As Variant, var_X2 As Variant, var_Y1 As Variant, var_Y2 As Variant, _
    var_XBase As Variant, var_XLookup As Variant) As Double
    ' ## Interpolation in Y^2*X and extrapolation, between two points
    Dim dbl_Interp As Double

    dbl_Interp = Interp_Lin_Binary(var_X1, var_X2, var_Y1 ^ 2 * (var_X1 - var_XBase), var_Y2 ^ 2 * (var_X2 - var_XBase), var_XLookup)
    Interp_V2t_Binary = Sqr(dbl_Interp / (var_XLookup - var_XBase))
End Function

Public Function Interp_V2t(arr_X As Variant, arr_Y As Variant, var_XBase As Variant, var_XLookup As Variant) As Double
    ' ## Linear interpolation in (Y^2 * X), with flat extrapolation
    ' ## X usually represents time, Y usually represents volatility
    Dim int_LowerBound As Integer: int_LowerBound = Examine_LowerBoundIndex(arr_X)
    Dim int_UpperBound As Integer: int_UpperBound = Examine_UpperBoundIndex(arr_X)

    If var_XLookup <= arr_X(int_LowerBound) Then
        ' Flat extrapolation
        Interp_V2t = arr_Y(int_LowerBound)
    ElseIf var_XLookup < arr_X(int_UpperBound) Then
        ' Interpolation
        Dim int_RowCtr As Integer
        Dim dbl_V2tLeft As Double, dbl_V2tRight As Double, dbl_V2tInterp As Double
        Dim dbl_DateLeft As Double, dbl_DateRight As Double

        For int_RowCtr = int_LowerBound + 1 To int_UpperBound
            If var_XLookup <= arr_X(int_RowCtr) Then
                Interp_V2t = Interp_V2t_Binary(arr_X(int_RowCtr - 1), arr_X(int_RowCtr), _
                    arr_Y(int_RowCtr - 1), arr_Y(int_RowCtr), var_XBase, var_XLookup)
                Exit Function
            End If
        Next int_RowCtr
    Else
        ' Flat extrapolation
        Interp_V2t = arr_Y(int_UpperBound)
    End If
End Function

Public Function Interp_V2t_Range(rng_Dates As Range, rng_Vols As Range, var_StartDate As Variant, var_LookupDate As Variant, _
    Optional intLst_IndexFilters As Collection = Nothing) As Double
    ' ## V2t interpolation using Excel ranges as inputs rather than arrays

    Dim dblLst_Dates As Collection: Set dblLst_Dates = Convert_RangeToList(rng_Dates, intLst_IndexFilters)
    Dim dblLst_Vols As Collection: Set dblLst_Vols = Convert_RangeToList(rng_Vols, intLst_IndexFilters)
    Interp_V2t_Range = Interp_V2t(dblLst_Dates, dblLst_Vols, var_StartDate, var_LookupDate)
End Function

Public Function Interp_V2t_WeekDays(arr_Dates As Variant, arr_Vols As Variant, lng_StartDate As Long, lng_LookupDate As Long) As Double
    ' ## V2t interpolation but only counting weekdays between two maturity pillars
    Dim int_LowerBound As Integer: int_LowerBound = Examine_LowerBoundIndex(arr_Dates)
    Dim int_UpperBound As Integer: int_UpperBound = Examine_UpperBoundIndex(arr_Dates)

    If lng_LookupDate <= arr_Dates(int_LowerBound) Then
        ' Flat extrapolation
        Interp_V2t_WeekDays = arr_Vols(int_LowerBound)
    ElseIf lng_LookupDate < arr_Dates(int_UpperBound) Then
        ' Interpolation
        Dim int_RowCtr As Integer
        Dim dbl_V2tLeft As Double, dbl_V2tRight As Double, dbl_V2tInterp As Double
        Dim lng_DateLeft As Long, lng_DateRight As Long
        Dim lng_WeekDaysRange As Long, lng_WeekDaysWithinLookup As Long

        For int_RowCtr = int_LowerBound + 1 To int_UpperBound
            If lng_LookupDate <= arr_Dates(int_RowCtr) Then
                lng_DateLeft = arr_Dates(int_RowCtr - 1)
                lng_DateRight = arr_Dates(int_RowCtr)
                dbl_V2tLeft = arr_Vols(int_RowCtr - 1) ^ 2 * (lng_DateLeft - lng_StartDate)
                dbl_V2tRight = arr_Vols(int_RowCtr) ^ 2 * (lng_DateRight - lng_StartDate)
                lng_WeekDaysRange = Calc_WeekDaysBetween(lng_DateLeft, lng_DateRight)
                lng_WeekDaysWithinLookup = Calc_WeekDaysBetween(lng_DateLeft, lng_LookupDate)

                dbl_V2tInterp = dbl_V2tLeft + lng_WeekDaysWithinLookup / lng_WeekDaysRange * (dbl_V2tRight - dbl_V2tLeft)
                Interp_V2t_WeekDays = Sqr(dbl_V2tInterp / (lng_LookupDate - lng_StartDate))
                Exit Function
            End If
        Next int_RowCtr
    Else
        ' Flat extrapolation
        Interp_V2t_WeekDays = arr_Vols(int_UpperBound)
    End If
End Function

Public Function Interp_V2t_WeekDays_Range(rng_Dates As Range, rng_Vols As Range, lng_StartDate As Long, lng_LookupDate As Long, _
    Optional intLst_IndexFilters As Collection = Nothing) As Double
    ' ## V2t interpolation on weekdays using Excel ranges as inputs rather than arrays
    Dim lngLst_Dates As Collection: Set lngLst_Dates = Convert_RangeToList(rng_Dates, intLst_IndexFilters)
    Dim dblLst_Vols As Collection: Set dblLst_Vols = Convert_RangeToList(rng_Vols, intLst_IndexFilters)
    Interp_V2t_WeekDays_Range = Interp_V2t_WeekDays(lngLst_Dates, dblLst_Vols, lng_StartDate, lng_LookupDate)
End Function

Public Function Interp_InvQuad(var_X1 As Variant, var_X2 As Variant, var_X3 As Variant, var_Y1 As Variant, _
    var_Y2 As Variant, var_Y3 As Variant, var_YTarget As Variant) As Double

    Interp_InvQuad = var_X1 * (var_Y2 - var_YTarget) * (var_Y3 - var_YTarget) / ((var_Y2 - var_Y1) * (var_Y3 - var_Y1)) _
        + var_X2 * (var_Y1 - var_YTarget) * (var_Y3 - var_YTarget) / ((var_Y1 - var_Y2) * (var_Y3 - var_Y2)) _
        + var_X3 * (var_Y1 - var_YTarget) * (var_Y2 - var_YTarget) / ((var_Y1 - var_Y3) * (var_Y2 - var_Y3))
End Function

Public Function Interp_Spline(arr_X As Variant, arr_Y As Variant, var_LookupX As Variant) As Double
    ' ## Clamped cubic spline algorithm
    ' ## X inputs need to be in ascending order
    ' ## Burden & Faires - Numerical Analysis, 4th edition, algorithm 3.5
    Dim int_LowerBound As Integer: int_LowerBound = Examine_LowerBoundIndex(arr_X)
    Dim int_UpperBound As Integer: int_UpperBound = Examine_UpperBoundIndex(arr_X)

    ' Handle case where there is only one point
    If int_LowerBound = int_UpperBound Then
        Interp_Spline = arr_Y(int_UpperBound)
        Exit Function
    End If

    Dim FPO As Double: FPO = (arr_Y(int_LowerBound + 1) - arr_Y(int_LowerBound)) / (arr_X(int_LowerBound + 1) - arr_X(int_LowerBound))
    Dim FPN As Double: FPN = (arr_Y(int_UpperBound) - arr_Y(int_UpperBound - 1)) / (arr_X(int_UpperBound) - arr_X(int_UpperBound - 1))

    If var_LookupX < arr_X(int_LowerBound) Then
        Interp_Spline = arr_Y(int_LowerBound) + (var_LookupX - arr_X(int_LowerBound)) * FPO  ' Linear extrapolation
    ElseIf var_LookupX > arr_X(int_UpperBound) Then
        Interp_Spline = arr_Y(int_UpperBound) + (var_LookupX - arr_X(int_UpperBound)) * FPN  ' Linear extrapolation
    Else
        ' ## CREATE SPLINE

        ' Coefficients of cubic polynomial
        Dim a() As Double: ReDim a(int_LowerBound To int_UpperBound) As Double
        Dim b() As Double: ReDim b(int_LowerBound To int_UpperBound) As Double
        Dim c() As Double: ReDim c(int_LowerBound To int_UpperBound) As Double
        Dim d() As Double: ReDim d(int_LowerBound To int_UpperBound) As Double

        ' Other intermediate variables
        Dim h() As Double: ReDim h(int_LowerBound To int_UpperBound) As Double
        Dim alpha() As Double: ReDim alpha(int_LowerBound To int_UpperBound) As Double
        Dim l() As Double: ReDim l(int_LowerBound To int_UpperBound) As Double
        Dim mu() As Double: ReDim mu(int_LowerBound To int_UpperBound) As Double
        Dim z() As Double: ReDim z(int_LowerBound To int_UpperBound) As Double

        'Step 1 - Assign values for a and h
        Dim i As Integer
        For i = int_LowerBound To int_UpperBound - 1
            a(i) = arr_Y(i)
            h(i) = arr_X(i + 1) - arr_X(i)
        Next i
        a(int_UpperBound) = arr_Y(int_UpperBound)

        ' Step 2 - Assign endpoint values for alpha
        alpha(int_LowerBound) = 0
        alpha(int_UpperBound) = 0

        ' Step 3 - Assign intermediate values for alpha
        For i = int_LowerBound + 1 To int_UpperBound - 1
            alpha(i) = 3 * (a(i + 1) * h(i - 1) - a(i) * (arr_X(i + 1) - arr_X(i - 1)) + a(i - 1) * h(i)) / (h(i - 1) * h(i))
        Next i

        ' Step 4 - Assign left endpoint values for l, mu and z
        l(int_LowerBound) = 2 * h(int_LowerBound)
        mu(int_LowerBound) = 0.5
        z(int_LowerBound) = alpha(int_LowerBound) / l(int_LowerBound)

        ' Step 5 - Assign intermediate values for l, mu and z
        For i = int_LowerBound + 1 To int_UpperBound - 1
            l(i) = 2 * (arr_X(i + 1) - arr_X(i - 1)) - h(i - 1) * mu(i - 1)
            mu(i) = h(i) / l(i)
            z(i) = (alpha(i) - h(i - 1) * z(i - 1)) / l(i)
        Next i

        ' Step 6 - Assign right endpoint values for l, z and c
        l(int_UpperBound) = h(int_UpperBound - 1) * (2 - mu(int_UpperBound - 1))
        z(int_UpperBound) = (alpha(int_UpperBound) - h(int_UpperBound - 1) * z(int_UpperBound - 1)) / l(int_UpperBound)
        c(int_UpperBound) = z(int_UpperBound)

        ' Step 7 - Assign values for remainder of the coefficients
        Dim j As Integer
        For j = int_UpperBound - 1 To int_LowerBound Step -1
            c(j) = z(j) - mu(j) * c(j + 1)
            b(j) = (a(j + 1) - a(j)) / h(j) - h(j) * (c(j + 1) + 2 * c(j)) / 3
            d(j) = (c(j + 1) - c(j)) / (3 * h(j))
        Next j

        ' ## READ FROM SPLINE
        ' Determine relevant pillar for computing value of spline
        Dim int_RelevantPillar As Integer
        For i = int_LowerBound To int_UpperBound - 1
            If var_LookupX >= arr_X(i) And var_LookupX < arr_X(i + 1) Then
                int_RelevantPillar = i
                Exit For
            End If
        Next i

        If var_LookupX = arr_X(int_UpperBound) Then int_RelevantPillar = int_UpperBound - 1

        ' Compute value from spline
        Interp_Spline = a(int_RelevantPillar) + b(int_RelevantPillar) * (var_LookupX - arr_X(int_RelevantPillar)) _
            + c(int_RelevantPillar) * (var_LookupX - arr_X(int_RelevantPillar)) ^ 2 + d(int_RelevantPillar) * (var_LookupX - arr_X(int_RelevantPillar)) ^ 3
    End If
End Function

Public Function Interp_Spline_Range(rng_X As Range, rng_Y As Range, dbl_LookupX As Double) As Double
    ' ## Cubic spline interpolation using Excel ranges as inputs rather than arrays
    Interp_Spline_Range = Interp_Spline(Convert_RangeToList(rng_X), Convert_RangeToList(rng_Y), dbl_LookupX)
End Function

Public Function Interp_Poly(arr_X As Variant, arr_Y As Variant, var_LookupX As Variant) As Double
    ' ## Polynomial interpolation
    Interp_Poly = Calc_PolyValue(Calc_PolyCoefs(arr_X, arr_Y), var_LookupX)
End Function

Public Function Interp_Poly_Range(rng_X As Range, rng_Y As Range, dbl_LookupX As Double) As Double
    ' ## Polynomial interpolation using Excel ranges as inputs rather than arrays
    Interp_Poly_Range = Interp_Poly(Convert_RangeToList(rng_X), Convert_RangeToList(rng_Y), dbl_LookupX)
End Function