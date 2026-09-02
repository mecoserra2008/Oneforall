' =============================================================================
' Pull-to-par assertions, executed against the REAL BondPullToParCore source
' lifted out of src/modPNL.bas (see tools/vba_run.py).
'
' Every case is a property the economics must satisfy, not a golden number
' copied out of a previous run - a golden number would just re-bless whatever
' the code did last time.
' =============================================================================

Const NCURVE As Integer = 11

Dim gOut As Integer
Dim gFail As Long
Dim gPass As Long


Sub RunTests
    gOut = FreeFile
    Open "@@OUTPUT@@" For Output As #gOut
    gFail = 0
    gPass = 0

    Test_ForwardReturnEqualsHorizonRate_FlatCurve
    Test_ForwardReturnEqualsHorizonRate_SteepCurve
    Test_ForwardReturnEqualsHorizonRate_InvertedCurve
    Test_ForwardPricingRemovesFakeRollDown
    Test_PremiumBondPullsDown
    Test_DiscountBondPullsUp
    Test_ScalesWithHorizon
    Test_ZeroSpreadVsWideSpread
    Test_RejectsBadInputs
    Test_ScheduleEndsExactlyOnMaturity
    Test_CouponPaidInsideHorizonIsNotDoubleCounted

    Print #gOut, "----"
    Print #gOut, "passed=" & gPass & " failed=" & gFail
    Close #gOut
End Sub


' --- assertion helpers -------------------------------------------------------

' `label` deliberately, not `name`: Name is a Basic statement keyword and using
' it as a parameter is a compile error, not a runtime one.
Sub Ok(ByVal label As String, ByVal condition As Boolean, ByVal detail As String)
    If condition Then
        gPass = gPass + 1
        Print #gOut, "PASS  " & label
    Else
        gFail = gFail + 1
        Print #gOut, "FAIL  " & label & "   " & detail
    End If
End Sub

Function Approx(ByVal a As Double, ByVal b As Double, ByVal tol As Double) As Boolean
    Approx = (Abs(a - b) <= tol)
End Function


' --- curve builders ----------------------------------------------------------

' The tenor grid every test curve uses.  NCURVE nodes, matching the shape of
' the real OIS_Curves grid.
' The builders FILL a caller-allocated array rather than ReDim-ing a ByRef
' parameter, which keeps them portable across Basic dialects.

Sub FillTenors(ByRef xs() As Double)
    xs(1) = 0.25:  xs(2) = 0.5:  xs(3) = 1#:   xs(4) = 2#
    xs(5) = 3#:    xs(6) = 5#:   xs(7) = 7#:   xs(8) = 10#
    xs(9) = 15#:   xs(10) = 20#: xs(11) = 30#
End Sub

' Flat curve at `level` percent across the standard tenor grid.
Sub FlatCurve(ByVal level As Double, ByRef xs() As Double, ByRef ys() As Double)
    Dim i As Integer
    FillTenors xs
    For i = 1 To NCURVE
        ys(i) = level
    Next i
End Sub

' Upward sloping: `shortEnd` percent at the short end, plus `slope` percent per
' year.  Not `base` - Base is a Basic keyword (Option Base) and using it as a
' parameter name is a compile error, which headless shows up only as a hang.
Sub SlopedCurve(ByVal shortEnd As Double, ByVal slope As Double, _
                ByRef xs() As Double, ByRef ys() As Double)
    Dim i As Integer
    FillTenors xs
    For i = 1 To NCURVE
        ys(i) = shortEnd + slope * xs(i)
    Next i
End Sub


' -----------------------------------------------------------------------------
' THE defining identity.
'
' When no coupon falls inside the horizon, the forward dirty price must be the
' spot dirty price grown at the horizon zero rate - exactly, on ANY curve shape.
' That single equality is what makes this "pull to par and roll" rather than
' "free money from a steep curve".
'
' The function returns a CLEAN price change, so put accrued back on both sides:
'
'     pullToPar + (AI(T0) - AI(T-1))  ==  P_dirty(T-1) * ((1+y(h))^h - 1)
'
' Note what this also says about the carry split: the left-hand side is the
' bond's TOTAL carry over the period.  How much of it is labelled coupon and how
' much price is a day-count convention choice; the total is not.
' -----------------------------------------------------------------------------
Sub Test_ForwardReturnEqualsHorizonRate_FlatCurve
    CheckForwardReturnIdentity "flat curve", 3#, 0#, 0.03
End Sub

Sub Test_ForwardReturnEqualsHorizonRate_SteepCurve
    CheckForwardReturnIdentity "steep curve", 1#, 0.15, 0.03
End Sub

Sub Test_ForwardReturnEqualsHorizonRate_InvertedCurve
    CheckForwardReturnIdentity "inverted curve", 5#, -0.08, 0.03
End Sub

Sub CheckForwardReturnIdentity(ByVal label As String, _
                               ByVal shortEnd As Double, _
                               ByVal slope As Double, _
                               ByVal coupon As Double)

    Dim xs(1 To NCURVE) As Double, ys(1 To NCURVE) As Double
    Dim prior As Date, current As Date, maturity As Date
    Dim pull As Variant, spotDirty As Variant
    Dim horizon As Double, horizonRate As Double
    Dim accrualDelta As Double, expected As Double

    If slope = 0# Then
        FlatCurve shortEnd, xs, ys
    Else
        SlopedCurve shortEnd, slope, xs, ys
    End If

    ' 31-Dec to 31-Mar: annual coupons pay on 30-Jun, so none inside the horizon.
    prior = DateSerial(2025, 12, 31)
    current = DateSerial(2026, 3, 31)
    maturity = DateSerial(2032, 6, 30)

    pull = BondPullToParCore(prior, current, maturity, _
                             coupon, 1, 2, 0#, xs, ys, NCURVE)

    spotDirty = BondPullModelDirtyPrice(prior, maturity, coupon, 1, 0#, _
                                        xs, ys, NCURVE)

    horizon = BondPullYears(prior, current)
    horizonRate = BondPullInterp(xs, ys, NCURVE, horizon) / 100#

    accrualDelta = AccruedInterest(current, maturity, coupon, 1, 2) _
                 - AccruedInterest(prior, maturity, coupon, 1, 2)

    expected = CDbl(spotDirty) * ((1 + horizonRate) ^ horizon - 1)

    Ok "forward return equals the horizon rate (" & label & ")", _
        Approx(CDbl(pull) + accrualDelta, expected, 0.0000001), _
        "pull+accrual=" & CStr(CDbl(pull) + accrualDelta) & _
        " expected=" & CStr(expected)
End Sub


' -----------------------------------------------------------------------------
' THE regression this rewrite exists for.
'
' The naive alternative - reprice on the SAME curve at the SHORTER maturity -
' books the roll-down of an upward-sloping curve as profit the desk never
' earned.  Both methods must agree when the curve is flat (there is nothing to
' roll down), and the naive one must be materially richer when it is steep.
' -----------------------------------------------------------------------------
Sub Test_ForwardPricingRemovesFakeRollDown
    Dim xs(1 To NCURVE) As Double, ys(1 To NCURVE) As Double
    Dim prior As Date, current As Date, maturity As Date
    Dim coupon As Double
    Dim fwdFlat As Double, naiveFlat As Double
    Dim fwdSteep As Double, naiveSteep As Double

    prior = DateSerial(2025, 12, 31)
    current = DateSerial(2026, 3, 31)
    maturity = DateSerial(2032, 6, 30)
    coupon = 0.03

    FlatCurve 3#, xs, ys
    fwdFlat = CDbl(BondPullToParCore(prior, current, maturity, _
                                     coupon, 1, 2, 0#, xs, ys, NCURVE))
    naiveFlat = NaiveRollDown(prior, current, maturity, coupon, xs, ys)

    Ok "flat curve: forward and naive agree (nothing to roll down)", _
        Approx(fwdFlat, naiveFlat, 0.002), _
        "fwd=" & CStr(fwdFlat) & " naive=" & CStr(naiveFlat)

    SlopedCurve 1#, 0.15, xs, ys
    fwdSteep = CDbl(BondPullToParCore(prior, current, maturity, _
                                      coupon, 1, 2, 0#, xs, ys, NCURVE))
    naiveSteep = NaiveRollDown(prior, current, maturity, coupon, xs, ys)

    Ok "steep curve: naive books roll-down the forward method does not", _
        naiveSteep - fwdSteep > 0.05, _
        "fwd=" & CStr(fwdSteep) & " naive=" & CStr(naiveSteep)
End Sub

' Clean price change from repricing on the SAME curve one horizon later - the
' method this rewrite replaced.  Present only as the contrast case.
Function NaiveRollDown(ByVal prior As Date, ByVal current As Date, _
                       ByVal maturity As Date, ByVal coupon As Double, _
                       ByRef xs() As Double, ByRef ys() As Double) As Double

    Dim cleanPrior As Double, cleanCurrent As Double

    cleanPrior = CDbl(BondPullModelDirtyPrice(prior, maturity, coupon, 1, 0#, _
                                              xs, ys, NCURVE)) _
               - AccruedInterest(prior, maturity, coupon, 1, 2)

    cleanCurrent = CDbl(BondPullModelDirtyPrice(current, maturity, coupon, 1, 0#, _
                                                xs, ys, NCURVE)) _
                 - AccruedInterest(current, maturity, coupon, 1, 2)

    NaiveRollDown = cleanCurrent - cleanPrior
End Function


' -----------------------------------------------------------------------------
' Coupon above the curve -> the bond trades above par -> its clean price must
' fall towards par as time passes.
' -----------------------------------------------------------------------------
Sub Test_PremiumBondPullsDown
    Dim xs(1 To NCURVE) As Double, ys(1 To NCURVE) As Double
    Dim r As Variant
    FlatCurve 2#, xs, ys

    r = BondPullToParCore( _
        DateSerial(2025, 6, 30), DateSerial(2025, 7, 31), DateSerial(2032, 6, 30), _
        0.06, 1, 2, 0#, xs, ys, NCURVE)

    Ok "premium bond pulls DOWN to par", _
        CDbl(r) < 0#, "expected negative, got " & CStr(r)

    ' one month of a ~4% price premium unwinding over 7 years is cents, not euros
    Ok "premium pull is economically sized", _
        Abs(CDbl(r)) > 0.001 And Abs(CDbl(r)) < 0.5, "got " & CStr(r)
End Sub


' -----------------------------------------------------------------------------
' Coupon below the curve -> discount bond -> clean price rises towards par.
' -----------------------------------------------------------------------------
Sub Test_DiscountBondPullsUp
    Dim xs(1 To NCURVE) As Double, ys(1 To NCURVE) As Double
    Dim r As Variant
    FlatCurve 5#, xs, ys

    r = BondPullToParCore( _
        DateSerial(2025, 6, 30), DateSerial(2025, 7, 31), DateSerial(2032, 6, 30), _
        0.01, 1, 2, 0#, xs, ys, NCURVE)

    Ok "discount bond pulls UP to par", _
        CDbl(r) > 0#, "expected positive, got " & CStr(r)
End Sub


' -----------------------------------------------------------------------------
' Pull-to-par is a time effect, so a longer horizon must move the price further
' in the same direction (monotone, and roughly proportional over short windows).
' -----------------------------------------------------------------------------
Sub Test_ScalesWithHorizon
    Dim xs(1 To NCURVE) As Double, ys(1 To NCURVE) As Double
    Dim r1 As Variant, r2 As Variant
    FlatCurve 2#, xs, ys

    r1 = BondPullToParCore( _
        DateSerial(2025, 6, 30), DateSerial(2025, 7, 30), DateSerial(2032, 6, 30), _
        0.06, 1, 2, 0#, xs, ys, NCURVE)

    r2 = BondPullToParCore( _
        DateSerial(2025, 6, 30), DateSerial(2025, 8, 29), DateSerial(2032, 6, 30), _
        0.06, 1, 2, 0#, xs, ys, NCURVE)

    Ok "two months pulls further than one, same sign", _
        CDbl(r2) < CDbl(r1) And CDbl(r1) < 0#, _
        "1m=" & CStr(r1) & " 2m=" & CStr(r2)

    Ok "roughly linear in the horizon", _
        Approx(CDbl(r2) / CDbl(r1), 2#, 0.15), _
        "ratio " & CStr(CDbl(r2) / CDbl(r1))
End Sub


' -----------------------------------------------------------------------------
' A wide credit spread makes the bond a bigger discount, so it must pull UP
' harder than the same bond at zero spread.
' -----------------------------------------------------------------------------
Sub Test_ZeroSpreadVsWideSpread
    Dim xs(1 To NCURVE) As Double, ys(1 To NCURVE) As Double
    Dim rTight As Variant, rWide As Variant
    FlatCurve 3#, xs, ys

    rTight = BondPullToParCore( _
        DateSerial(2025, 6, 30), DateSerial(2025, 7, 31), DateSerial(2032, 6, 30), _
        0.03, 1, 2, 0#, xs, ys, NCURVE)

    rWide = BondPullToParCore( _
        DateSerial(2025, 6, 30), DateSerial(2025, 7, 31), DateSerial(2032, 6, 30), _
        0.03, 1, 2, 0.02, xs, ys, NCURVE)      ' +200bp

    Ok "wider spread => stronger upward pull", _
        CDbl(rWide) > CDbl(rTight), _
        "tight=" & CStr(rTight) & " wide=" & CStr(rWide)
End Sub


' -----------------------------------------------------------------------------
' Bad inputs must come back as #N/A (Error 2042), never as a number and never as
' a runtime error that would blank the whole PNL column.
' -----------------------------------------------------------------------------
Sub Test_RejectsBadInputs
    Dim xs(1 To NCURVE) As Double, ys(1 To NCURVE) As Double
    Dim r As Variant
    FlatCurve 3#, xs, ys

    ' current date before prior date
    r = BondPullToParCore( _
        DateSerial(2025, 7, 31), DateSerial(2025, 6, 30), DateSerial(2032, 6, 30), _
        0.03, 1, 2, 0#, xs, ys, NCURVE)
    Ok "reverse dates rejected", IsError(r), "got " & CStr(r)

    ' matured bond
    r = BondPullToParCore( _
        DateSerial(2025, 6, 30), DateSerial(2025, 7, 31), DateSerial(2024, 1, 1), _
        0.03, 1, 2, 0#, xs, ys, NCURVE)
    Ok "matured bond rejected", IsError(r), "got " & CStr(r)

    ' unsupported coupon frequency
    r = BondPullToParCore( _
        DateSerial(2025, 6, 30), DateSerial(2025, 7, 31), DateSerial(2032, 6, 30), _
        0.03, 3, 2, 0#, xs, ys, NCURVE)
    Ok "frequency 3 rejected", IsError(r), "got " & CStr(r)

    ' unsupported day count code
    r = BondPullToParCore( _
        DateSerial(2025, 6, 30), DateSerial(2025, 7, 31), DateSerial(2032, 6, 30), _
        0.03, 1, 99, 0#, xs, ys, NCURVE)
    Ok "day count 99 rejected", IsError(r), "got " & CStr(r)

    ' degenerate curve
    r = BondPullToParCore( _
        DateSerial(2025, 6, 30), DateSerial(2025, 7, 31), DateSerial(2032, 6, 30), _
        0.03, 1, 2, 0#, xs, ys, 1)
    Ok "single-node curve rejected", IsError(r), "got " & CStr(r)
End Sub


' -----------------------------------------------------------------------------
' The redemption must land exactly on the maturity date, whatever month-end
' rolling the coupon generator does on the way there.
' -----------------------------------------------------------------------------
Sub Test_ScheduleEndsExactlyOnMaturity
    Dim d() As Date
    Dim n As Long

    n = BondPullSchedule(DateSerial(2025, 3, 17), DateSerial(2031, 8, 31), 2, d)
    Ok "semi-annual schedule ends on maturity", _
        n > 0 And d(n) = DateSerial(2031, 8, 31), _
        "n=" & CStr(n) & " last=" & CStr(d(n))

    n = BondPullSchedule(DateSerial(2025, 3, 17), DateSerial(2031, 8, 31), 4, d)
    Ok "quarterly schedule ends on maturity", _
        n > 0 And d(n) = DateSerial(2031, 8, 31), _
        "n=" & CStr(n) & " last=" & CStr(d(n))

    ' settlement inside the final coupon period -> redemption only
    n = BondPullSchedule(DateSerial(2031, 7, 1), DateSerial(2031, 8, 31), 2, d)
    Ok "stub period yields one redemption flow", _
        n = 1 And d(1) = DateSerial(2031, 8, 31), _
        "n=" & CStr(n)
End Sub


' -----------------------------------------------------------------------------
' A coupon that pays INSIDE the horizon is cash in the bank, reported by
' Carry_Coupon.  It must drop out of the forward price rather than being
' compounded into it, or the carry leg doubles.
'
' Property: crossing a coupon date must not create a discontinuous jump in
' pull-to-par of anything like the coupon's size.
' -----------------------------------------------------------------------------
Sub Test_CouponPaidInsideHorizonIsNotDoubleCounted
    Dim xs(1 To NCURVE) As Double, ys(1 To NCURVE) As Double
    Dim beforeCpn As Variant, afterCpn As Variant
    FlatCurve 3#, xs, ys

    ' maturity 30-Jun-2032, annual coupons on 30-Jun.
    ' horizon ending 29-Jun-2026 pays no coupon; ending 1-Jul-2026 pays one.
    beforeCpn = BondPullToParCore( _
        DateSerial(2025, 12, 31), DateSerial(2026, 6, 29), DateSerial(2032, 6, 30), _
        0.06, 1, 2, 0#, xs, ys, NCURVE)

    afterCpn = BondPullToParCore( _
        DateSerial(2025, 12, 31), DateSerial(2026, 7, 1), DateSerial(2032, 6, 30), _
        0.06, 1, 2, 0#, xs, ys, NCURVE)

    Ok "no coupon-sized jump across a payment date", _
        Abs(CDbl(afterCpn) - CDbl(beforeCpn)) < 1#, _
        "before=" & CStr(beforeCpn) & " after=" & CStr(afterCpn)
End Sub
