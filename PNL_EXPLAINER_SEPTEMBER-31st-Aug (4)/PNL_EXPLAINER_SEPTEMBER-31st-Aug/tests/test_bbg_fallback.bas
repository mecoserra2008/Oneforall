' =============================================================================
' THE BLOOMBERG SECURITY FALLBACK CHAIN
'
' A bond field is asked of one security id after another until one answers.  The
' order matters and the LENGTH is capped by Excel, so both are asserted here.
'
' Order: the desk starts a BDP lookup with the ISIN-suffixed form, so the chain
' is the already-resolved ticker, then `XS123... ISIN`, then /isin/, then the
' venue-qualified forms.
'
' Positions come from the BCOL_ constants, never from a literal RC number.  The
' first version of this test pinned RC72..RC77, and when the candidate block
' moved one column left - to match the sheet it had never actually matched -
' every one of those assertions failed for the right reason and had to be
' hand-edited.  A test that must be edited whenever the thing it guards changes
' legitimately is not guarding it.  What is asserted here is ORDER and IDENTITY.
'
' Length: Excel stops at 64 levels of nested functions.  Every (candidate x
' field) pair costs two, so a column asking for five fields cannot try as many
' ids as one asking for one.  A formula over the limit is not a wrong number -
' Excel refuses it outright and the write fails.
' =============================================================================

Dim gOut As Integer
Dim gPassed As Long
Dim gFailed As Long


' `what` deliberately, not `name`: Name is a Basic statement keyword and using
' it as a parameter is a compile error, not a runtime one.
Sub AssertEq(ByVal got As Variant, ByVal want As Variant, ByVal what As String)
    If CStr(got) = CStr(want) Then
        gPassed = gPassed + 1
        Print #gOut, "PASS  " & what
    Else
        gFailed = gFailed + 1
        Print #gOut, "FAIL  " & what & "   got " & CStr(got) & ", wanted " & CStr(want)
    End If
End Sub


Sub AssertTrue(ByVal cond As Boolean, ByVal what As String, ByVal detail As String)
    If cond Then
        gPassed = gPassed + 1
        Print #gOut, "PASS  " & what
    Else
        gFailed = gFailed + 1
        Print #gOut, "FAIL  " & what & "   " & detail
    End If
End Sub


' Deepest nesting of function calls in a formula - what Excel's 64 limit counts.
Function MaxNesting(ByVal f As String) As Long
    Dim i As Long, d As Long, mx As Long, ch As String
    For i = 1 To Len(f)
        ch = Mid$(f, i, 1)
        If ch = "(" Then
            d = d + 1
            If d > mx Then mx = d
        ElseIf ch = ")" Then
            d = d - 1
        End If
    Next i
    MaxNesting = mx
End Function


Function Occurrences(ByVal hay As String, ByVal needle As String) As Long
    Dim n As Long, p As Long
    p = InStr(1, hay, needle, vbTextCompare)
    Do While p > 0
        n = n + 1
        p = InStr(p + 1, hay, needle, vbTextCompare)
    Loop
    Occurrences = n
End Function


' --- the order --------------------------------------------------------------
Sub TestOrder()

    Dim c As Variant
    c = BondSecurityFallbackCols()

    ' The resolved ticker leads: PARSEKYABLE_DES has already worked out which
    ' form this bond answers to, so it is tried before any guess.
    AssertEq c(0), 41, "1st is the resolved ticker (AO = RC41)"

    ' Then the desk's own starting point for a BDP lookup.
    AssertEq c(1), ColIdx(BCOL_BBG_CAND_ISIN), "2nd is ISIN & "" ISIN"""

    ' Then /isin/, which the module header says the resolver tries first among
    ' the constructed forms.
    AssertEq c(2), ColIdx(BCOL_BBG_CAND_SLASHISIN), "3rd is /isin/"

    ' Then the venue-qualified forms, Corp before Govt.
    AssertEq c(3), ColIdx(BCOL_BBG_CAND_CORP), "4th is Corp"
    AssertEq c(4), ColIdx(BCOL_BBG_CAND_BVAL_CORP), "5th is @BVAL Corp"
    AssertEq c(6), ColIdx(BCOL_BBG_CAND_GOVT), "7th is Govt"

    AssertEq UBound(c) + 1, 11, "eleven candidates in all"

    ' THE REGRESSION.  The chain used to be Array(COL_BBG_TICKER, 75, 77) -
    ' ticker, @BVAL Corp, Govt - with NEITHER ISIN form in it.
    AssertTrue c(1) <> 75, "the 2nd is no longer @BVAL Corp", "still 75"

End Sub


' --- the length cap ---------------------------------------------------------
Sub TestCap()

    ' 6 + 2 * candidates * fields <= 60
    AssertEq BondFallbackColCount(1), 11, "one field tries the whole chain"
    AssertEq BondFallbackColCount(2), 11, "two fields still try the whole chain"
    AssertEq BondFallbackColCount(3), 9, "three fields fit nine"
    AssertEq BondFallbackColCount(4), 6, "four fields fit six"
    AssertEq BondFallbackColCount(5), 5, "five fields fit five"

    ' Never zero, whatever it is asked.
    AssertTrue BondFallbackColCount(0) >= 1, "a zero field count still tries one", ""
    AssertTrue BondFallbackColCount(99) >= 1, "an absurd field count still tries one", ""

    ' And it is an improvement, not just a change: every column used to get
    ' three candidates regardless.
    AssertTrue BondFallbackColCount(1) > 3, "one-field columns now try MORE than before", ""
    AssertTrue BondFallbackColCount(5) > 3, "even the five-field column tries more", ""

End Sub


' --- the formulas that come out ---------------------------------------------
Sub TestRenderedFormula()

    Dim f1 As String, f5 As String

    ' The second argument is Optional with a "" default in VBA.  It is passed
    ' explicitly here because LibreOffice Basic - which is what runs this - does
    ' not apply a VBA optional-parameter default and raises a type mismatch
    ' instead.  A harness limitation, not a difference in the code under test:
    ' "" is exactly what the real callers leave it as.

    ' PX_LAST: one field, so the whole chain fits.
    f1 = BBGFirstBDPMultiFieldFormulaR1C1(Array("PX_LAST"), "")

    ' ModDur: the worst case in the workbook, five fields.
    f5 = BBGFirstBDPMultiFieldFormulaR1C1( _
        Array("YAS_MOD_DUR", "DUR_ADJ_MID", "MOD_DUR_MID", "DUR_MID", "DURATION"), "")

    AssertTrue MaxNesting(f1) <= 64, "one-field formula is inside Excel's limit", _
        "nesting " & MaxNesting(f1)
    AssertTrue MaxNesting(f5) <= 64, "five-field formula is inside Excel's limit", _
        "nesting " & MaxNesting(f5)

    ' The cap is meant to leave headroom, not to sit on the line.
    AssertTrue MaxNesting(f5) <= 60, "five-field formula keeps the headroom", _
        "nesting " & MaxNesting(f5)

    ' Excel also caps a formula at 8192 characters.
    AssertTrue Len(f1) <= 8192, "one-field formula is inside the length limit", _
        "chars " & Len(f1)
    AssertTrue Len(f5) <= 8192, "five-field formula is inside the length limit", _
        "chars " & Len(f5)

    ' One BDP per (candidate x field).
    AssertEq Occurrences(f1, "BDP("), 11, "one field x eleven candidates = 11 BDP"
    AssertEq Occurrences(f5, "BDP("), 25, "five fields x five candidates = 25 BDP"

    ' The ISIN-suffixed candidate is actually referenced, and early.
    AssertTrue InStr(f1, "RC" & ColIdx(BCOL_BBG_CAND_ISIN)) > 0, _
        "the ISIN-suffixed candidate is in the formula", ""
    AssertTrue InStr(f1, "RC" & COL_BBG_TICKER) < InStr(f1, "RC" & ColIdx(BCOL_BBG_CAND_ISIN)), _
        "the ticker is tried before it", ""
    AssertTrue InStr(f1, "RC" & ColIdx(BCOL_BBG_CAND_ISIN)) < InStr(f1, "RC" & ColIdx(BCOL_BBG_CAND_SLASHISIN)), _
        "and it is tried before /isin/", ""

    ' Every formula still guards on a blank ISIN.
    AssertTrue Left$(f1, 8) = "=IF(RC1=", "still guarded on a blank ISIN", Left$(f1, 12)

End Sub


Sub RunTests

    gOut = FreeFile
    Open "@@OUTPUT@@" For Output As #gOut

    gPassed = 0
    gFailed = 0

    TestOrder
    TestCap
    TestRenderedFormula

    Print #gOut, "----"
    Print #gOut, "passed=" & gPassed & " failed=" & gFailed
    Close #gOut

End Sub
