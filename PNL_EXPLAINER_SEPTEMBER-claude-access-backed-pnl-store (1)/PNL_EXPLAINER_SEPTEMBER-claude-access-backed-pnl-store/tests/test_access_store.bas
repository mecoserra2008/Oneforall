' =============================================================================
' THE ACCESS STORE'S PURE HALF
'
' Everything here decides what gets written, without a database in sight:
'
'   the POSITION KEYS      what "the same position" means.  Wrong, and either
'                          two real positions collapse into one (a hedge
'                          silently disappears) or one position becomes two
'                          (its risk is double-counted).
'
'   the FINGERPRINT        what "the same run" means.  Wrong in one direction
'                          and every save is a duplicate; wrong in the other
'                          and a changed book is refused as unchanged.
'
'   the SQL LITERALS       a comma decimal separator turns VALUES (1.5) into
'                          VALUES (1,5), which is not an error - it is a valid
'                          statement writing the wrong columns.
'
'   the DDL                a malformed CREATE TABLE is otherwise only ever
'                          seen as a provider error on somebody else's machine.
' =============================================================================

Dim gOut As Integer
Dim gPassed As Long
Dim gFailed As Long


Sub AssertEq(ByVal got As Variant, ByVal want As Variant, ByVal what As String)
    If CStr(got) = CStr(want) Then
        gPassed = gPassed + 1
        Print #gOut, "PASS  " & what
    Else
        gFailed = gFailed + 1
        Print #gOut, "FAIL  " & what & "   got [" & CStr(got) & "], wanted [" & CStr(want) & "]"
    End If
End Sub


Sub AssertNe(ByVal a As Variant, ByVal b As Variant, ByVal what As String)
    If CStr(a) <> CStr(b) Then
        gPassed = gPassed + 1
        Print #gOut, "PASS  " & what
    Else
        gFailed = gFailed + 1
        Print #gOut, "FAIL  " & what & "   both are [" & CStr(a) & "]"
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


' -----------------------------------------------------------------------------
' POSITION KEYS
' -----------------------------------------------------------------------------
Sub TestBondKeys()

    Dim a As String, b As String, c As String

    a = AccBondPositionKey("XS1234567890", "PORT1", "HTC")
    b = AccBondPositionKey("XS1234567890", "PORT1", "AFS")

    AssertNe a, b, "one ISIN in two accounting categories is two positions"

    AssertEq AccBondPositionKey("xs1234567890", " port1 ", "htc"), a, _
        "case and surrounding blanks do not make a new position"

    AssertNe AccBondPositionKey("XS1234567890", "PORT2", "HTC"), a, _
        "the same bond in another portfolio is another position"

    ' The trap this guards: components joined without a separator let
    ' ("AB","C") and ("A","BC") produce one key.
    AssertNe AccBondPositionKey("AB", "C", "X"), _
             AccBondPositionKey("A", "BC", "X"), _
        "the components cannot run together"

    c = AccBondPositionKey("", "", "")
    AssertTrue Len(c) > 0, "an all-blank bond still yields a key", _
        "empty would be indistinguishable from 'no key at all'"

End Sub


Sub TestSwapKeys()

    Dim plainA As String, plainB As String, synth As String

    ' Swaps!A already carries source|id|LinkedISIN|row.
    plainA = AccSwapPositionKey("PLAIN|EUSW10|XS111|42")
    plainB = AccSwapPositionKey("PLAIN|EUSW10|XS222|42")
    synth = AccSwapPositionKey("SYNTHETIC|EUSW10|XS111|42")

    AssertNe plainA, plainB, _
        "one swap hedging two bonds is two positions"
    AssertNe plainA, synth, _
        "the plain and synthetic alternatives on one row are two positions"

    AssertEq AccSwapPositionKey("  PLAIN|EUSW10|XS111|42  "), plainA, _
        "surrounding blanks do not make a new position"

    AssertEq AccSwapPositionKey(""), "", _
        "a swap row with no mapping id has no key"

End Sub


Sub TestFutureKeys()

    Dim rtj As String, rt As String

    ' THE bug this exists for: the two coverage books number their rows
    ' independently, so row 163 is a different position in each.
    rtj = AccFuturePositionKey("RTJ", "163", "RXZ4", "XS111")
    rt = AccFuturePositionKey("RT", "163", "RXZ4", "XS111")

    AssertNe rtj, rt, _
        "coverage row 163 of Tx Juro is not row 163 of Total"

    AssertNe AccFuturePositionKey("RTJ", "163", "RXZ4", "XS111"), _
             AccFuturePositionKey("RTJ", "163", "RXZ4", "XS222"), _
        "one contract covering two bonds is two positions"

    AssertNe AccFuturePositionKey("RTJ", "163", "RXZ4", "XS111"), _
             AccFuturePositionKey("RTJ", "163", "OEZ4", "XS111"), _
        "two contracts on one coverage row are two positions"

    AssertEq AccFuturePositionKey("rtj", "163", "rxz4", "xs111"), rtj, _
        "case does not make a new position"

End Sub


' -----------------------------------------------------------------------------
' THE FINGERPRINT
' -----------------------------------------------------------------------------
Sub TestFingerprint()

    Dim h1 As String, h2 As String

    AssertEq Len(AccHash32("")), 8, "a hash is eight hex digits"
    AssertEq AccHash32(""), "811c9dc5", "the empty hash is the FNV-1a offset basis"
    AssertEq AccHash32("a"), "e40c292c", "FNV-1a of 'a'"
    AssertEq AccHash32("foobar"), "bf9cf968", "FNV-1a of 'foobar'"

    AssertNe AccHash32("XS111"), AccHash32("XS112"), _
        "one changed character changes the hash"

    h1 = AccFingerprintOfKeys(Array("A", "B", "C"))
    AssertEq AccFingerprintOfKeys(Array("A", "B", "C")), h1, _
        "the same keys fingerprint the same"

    AssertNe AccFingerprintOfKeys(Array("A", "B")), h1, _
        "a dropped position changes the fingerprint"

    AssertNe AccFingerprintOfKeys(Array("A", "C", "B")), h1, _
        "a reordered book is a different book"

    ' The separator is what stops this collapsing.
    h2 = AccFingerprintOfKeys(Array("AB", "C"))
    AssertNe AccFingerprintOfKeys(Array("A", "BC")), h2, _
        "keys cannot run together into one"

    AssertEq AccFingerprintOfKeys(Array()), AccHash32(""), _
        "an empty book fingerprints as the empty hash"

End Sub


' -----------------------------------------------------------------------------
' SQL LITERALS
' -----------------------------------------------------------------------------
Sub TestSqlLiterals()

    AssertEq AccSqlText("PORT1", 32), "'PORT1'", "plain text is quoted"
    AssertEq AccSqlText("O'Brien", 32), "'O''Brien'", "an apostrophe is doubled"
    AssertEq AccSqlText("", 32), "NULL", "blank is NULL, not an empty string"
    AssertEq AccSqlText("   ", 32), "NULL", "so is whitespace"
    AssertEq AccSqlText("ABCDEFGHIJ", 4), "'ABCD'", "text is cut to the column"

    ' The quote survives the truncation as a quote: cutting first and escaping
    ' afterwards is what keeps a half-escaped apostrophe out of the statement.
    AssertEq AccSqlText("AB'CDEF", 4), "'AB''C'", "a cut string is still escaped"

    AssertEq AccSqlNum(1.5), "1.5", "a decimal point, never a comma"
    AssertEq AccSqlNum(-2.25), "-2.25", "negatives keep their sign"
    AssertEq AccSqlNum(0), "0", "zero is zero, not NULL"
    AssertEq AccSqlNum(""), "NULL", "a blank number is NULL"
    AssertEq AccSqlNum("not a number"), "NULL", "text in a number column is NULL"

    AssertEq AccSqlDate(DateSerial(2026, 3, 7)), _
        "#2026-03-07 00:00:00#", "a date is an unambiguous ISO literal"
    AssertEq AccSqlDate(""), "NULL", "a blank date is NULL"
    AssertEq AccSqlDate("not a date"), "NULL", "unparseable is NULL"

End Sub


' -----------------------------------------------------------------------------
' THE DDL
' -----------------------------------------------------------------------------
Sub TestCreateTableSql()

    Dim sql As String

    sql = AccCreateTableSql("Pos_Bond")

    AssertTrue InStr(sql, "CREATE TABLE [Pos_Bond] (") = 1, _
        "the statement names its table", sql
    AssertTrue InStr(sql, "[RunID] LONG") > 0, _
        "RunID is in it", sql
    AssertTrue InStr(sql, "[PositionKey] TEXT(200)") > 0, _
        "PositionKey is sized to hold a key", sql
    AssertTrue InStr(sql, "PRIMARY KEY (RunID, PositionKey)") > 0, _
        "the key that separates runs AND refuses duplicates", sql
    AssertEq Right$(sql, 1), ")", "the statement is closed"

    AssertTrue InStr(AccCreateTableSql("Run"), "[RunID] AUTOINCREMENT") > 0, _
        "the run spine numbers itself", AccCreateTableSql("Run")

    AssertEq AccCreateTableSql("NoSuchTable"), "", _
        "an unknown table yields no statement rather than a broken one"

End Sub


Sub TestKeyLengthFits()

    Dim k As String

    ' A realistic worst case: the longest components the sheets can hold.
    k = AccFuturePositionKey(String$(4, "R"), String$(32, "9"), _
                             String$(64, "X"), String$(12, "Z"))

    AssertTrue Len(k) <= 200, _
        "a worst-case futures key fits PositionKey TEXT(200)", _
        "length " & CStr(Len(k))

End Sub


Sub RunTests()
    gOut = FreeFile
    Open "@@OUTPUT@@" For Output As #gOut

    gPassed = 0
    gFailed = 0

    TestBondKeys
    TestSwapKeys
    TestFutureKeys
    TestFingerprint
    TestSqlLiterals
    TestCreateTableSql
    TestKeyLengthFits

    Print #gOut, "----"
    Print #gOut, "passed=" & gPassed & " failed=" & gFailed
    Close #gOut
End Sub
