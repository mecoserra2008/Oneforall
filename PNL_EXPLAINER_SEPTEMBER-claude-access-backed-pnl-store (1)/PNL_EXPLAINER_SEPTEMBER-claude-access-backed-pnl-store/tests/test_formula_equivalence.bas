' =============================================================================
' FORMULA-STRING EQUIVALENCE
'
' WritePNLRow used to build these formulas inline.  They now go through the
' builders in modEconFormulas, so exactly one place defines each quantity.
'
' A refactor like that is only safe if the STRING is unchanged - the cells hold
' the generated text, and a difference of one character is a different model.
' Each case below pins the builder's output against the literal the inline code
' produced, character for character.  The expected strings are transcribed by
' hand, not captured from a run of the new code.
'
' The ACTUAL side is built from the real PCOL_/BCOL_ constants, so these cases
' also pin the sheet GEOMETRY: move a column and the generated string stops
' matching the hand-written literal, and the test says which one.  That is the
' failure this file exists to catch - removing a column is exactly when a
' formula quietly starts pointing one column left.
'
' Rows: PNL_Attribution row 5, Bonds row 6 - arbitrary, only the substitution
' pattern matters.
' =============================================================================

Const OUT_PATH As String = "@@OUTPUT@@"

' Sheet prefixes exactly as WritePNLRow builds them.
Const BND As String = "'Bonds'!"
Const CFG As String = "Config!"

Dim gOut As Integer
Dim gPass As Long
Dim gFail As Long


Sub RunTests
    gOut = FreeFile
    Open OUT_PATH For Output As #gOut
    gPass = 0
    gFail = 0

    Test_MarketValueChange
    Test_ResidualPnL
    Test_ResidualPercent
    Test_HedgeEfficiency
    Test_HedgeDv01Gap
    Test_RollToPar

    Print #gOut, "----"
    Print #gOut, "passed=" & gPass & " failed=" & gFail
    Close #gOut
End Sub


Sub Same(ByVal label As String, ByVal expected As String, ByVal actual As String)
    If expected = actual Then
        gPass = gPass + 1
        Print #gOut, "PASS  " & label
    Else
        gFail = gFail + 1
        Print #gOut, "FAIL  " & label
        Print #gOut, "        inline : " & expected
        Print #gOut, "        builder: " & actual
    End If
End Sub


' --- AA  Delta_Dirty_MV_EUR --------------------------------------------------
Sub Test_MarketValueChange
    Same "Delta_Dirty_MV_EUR", _
        "=IF(AND(ISNUMBER('Bonds'!AG6),ISNUMBER('Bonds'!AH6))," & _
        "'Bonds'!AG6-'Bonds'!AH6,"""")", _
        MarketValueChangeFml(BND & BCOL_DIRTYMV_T0_EUR & "6", _
                             BND & BCOL_DIRTYMV_TM1_EUR & "6")
End Sub


' --- BI  Unexplained_Residual_PnL --------------------------------------------
Sub Test_ResidualPnL
    Same "Unexplained_Residual_PnL", _
        "=IF(AND(ISNUMBER(BB5),ISNUMBER(BA5)),BB5-BA5,"""")", _
        ResidualPnLFml(PCOL_OFFICIAL_PNL & "5", PCOL_TOTAL_EXPLAINED & "5")
End Sub


' --- BJ  Unexplained_Residual_Pct --------------------------------------------
Sub Test_ResidualPercent
    Same "Unexplained_Residual_Pct", _
        "=IF(AND(ISNUMBER(BC5),ISNUMBER(BB5),BB5<>0),BC5/ABS(BB5),"""")", _
        ResidualPercentFml(PCOL_RESIDUAL & "5", PCOL_OFFICIAL_PNL & "5")
End Sub


' --- BM  Hedge_Efficiency ----------------------------------------------------
Sub Test_HedgeEfficiency
    Same "Hedge_Efficiency", _
        "=IF(AND(ISNUMBER(N5),ISNUMBER(T5),T5<>0)," & _
        "1-ABS(N5-T5)/ABS(T5),"""")", _
        HedgeEfficiencyFml(PCOL_HEDGE_DV01 & "5", _
                           PCOL_TARGET_HEDGE_DV01 & "5")
End Sub


' Bond_DV01_Change_Approx and Actual_vs_Synthetic_DV01 were pinned here.  Both
' columns have been removed from PnlLayout, so there is nothing left for the
' case to protect - a test that pins a formula the workbook never writes only
' fails when somebody tidies up, which is the opposite of useful.  Dv01ChangeFml
' itself is still in modEconFormulas; nothing calls it.


' --- U  Hedge_DV01_Gap -------------------------------------------------------
Sub Test_HedgeDv01Gap
    Same "Hedge_DV01_Gap", _
        "=IF(AND(ISNUMBER(N5),ISNUMBER(T5)),N5-T5,"""")", _
        HedgeDv01GapFml(PCOL_HEDGE_DV01 & "5", _
                        PCOL_TARGET_HEDGE_DV01 & "5")
End Sub


' --- AT  Carry_RollToPar -----------------------------------------------------
'
' The one case where the builder was not a like-for-like lift: the old
' RollToParFml in the library emitted a straight-line amortisation of the
' premium to maturity, which ignores the curve entirely.  What is pinned here is
' the CURRENT inline construction in WritePNLRow, which already called the
' BondPullToParPrice UDF - so this proves the move into the library changed
' nothing, while the library's superseded straight-line version is gone.
Sub Test_RollToPar
    Dim expected As String
    Dim actual As String

    expected = "=IFERROR(F5*'Bonds'!R6*" & _
        "BondPullToParPrice(" & _
            "Config!B4,Config!B5,'Bonds'!F6,'Bonds'!D6,'Bonds'!CE6," & _
            "'Bonds'!CK6,C5,BZ5," & _
            "BondSpreadTMinus1(BZ5,'Bonds'!AM6,'Bonds'!AK6,'Bonds'!AE6," & _
            "'Bonds'!AD6,'Bonds'!AR6)" & _
        ")/100,"""")"

    actual = RollToParFml(PCOL_NOTIONAL & "5", BND & BCOL_FX_TM1 & "6", _
        PullToParPriceFml(CFG & CFG_T0_DATE, CFG & CFG_T1_DATE, _
            BND & BCOL_MATURITY & "6", _
            BND & BCOL_COUPON & "6", _
            BND & BCOL_COUPONFREQ_NUM & "6", _
            BND & BCOL_BONDDCC_CODE & "6", _
            PCOL_CCY & "5", _
            PCOL_SPREAD_FRAMEWORK_AUTO & "5", _
            PriorSpreadFml(PCOL_SPREAD_FRAMEWORK_AUTO & "5", _
                BND & BCOL_GSPREAD_TM1 & "6", _
                BND & BCOL_ISPREAD_TM1 & "6", _
                BND & BCOL_ASW_TM1 & "6", _
                BND & BCOL_ZSPRD_TM1 & "6", _
                BND & BCOL_OAS_TM1 & "6")))

    Same "Carry_RollToPar", expected, actual
End Sub
