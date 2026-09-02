' =============================================================================
' THE RUN STORE'S PURE HALF
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
'   the CSV ROUND TRIP     a bond name with a comma in it is not exotic.  A
'                          writer that does not quote it, or a reader that
'                          splits on commas, works on every file anyone tests
'                          with and then silently shifts one row's columns.
'
'   the FOLDER ORDER       'the latest run' is a string comparison over folder
'                          names, so the name has to sort chronologically.
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

    a = StoreBondPositionKey("XS1234567890", "PORT1", "HTC")
    b = StoreBondPositionKey("XS1234567890", "PORT1", "AFS")

    AssertNe a, b, "one ISIN in two accounting categories is two positions"

    AssertEq StoreBondPositionKey("xs1234567890", " port1 ", "htc"), a, _
        "case and surrounding blanks do not make a new position"

    AssertNe StoreBondPositionKey("XS1234567890", "PORT2", "HTC"), a, _
        "the same bond in another portfolio is another position"

    ' The trap this guards: components joined without a separator let
    ' ("AB","C") and ("A","BC") produce one key.
    AssertNe StoreBondPositionKey("AB", "C", "X"), _
             StoreBondPositionKey("A", "BC", "X"), _
        "the components cannot run together"

    c = StoreBondPositionKey("", "", "")
    AssertTrue Len(c) > 0, "an all-blank bond still yields a key", _
        "empty would be indistinguishable from 'no key at all'"

End Sub


Sub TestSwapKeys()

    Dim plainA As String, plainB As String, synth As String

    ' Swaps!A already carries source|id|LinkedISIN|row.
    plainA = StoreSwapPositionKey("PLAIN|EUSW10|XS111|42")
    plainB = StoreSwapPositionKey("PLAIN|EUSW10|XS222|42")
    synth = StoreSwapPositionKey("SYNTHETIC|EUSW10|XS111|42")

    AssertNe plainA, plainB, _
        "one swap hedging two bonds is two positions"
    AssertNe plainA, synth, _
        "the plain and synthetic alternatives on one row are two positions"

    AssertEq StoreSwapPositionKey("  PLAIN|EUSW10|XS111|42  "), plainA, _
        "surrounding blanks do not make a new position"

    AssertEq StoreSwapPositionKey(""), "", _
        "a swap row with no mapping id has no key"

End Sub


Sub TestFutureKeys()

    Dim rtj As String, rt As String

    ' THE bug this exists for: the two coverage books number their rows
    ' independently, so row 163 is a different position in each.
    rtj = StoreFuturePositionKey("RTJ", "163", "RXZ4", "XS111")
    rt = StoreFuturePositionKey("RT", "163", "RXZ4", "XS111")

    AssertNe rtj, rt, _
        "coverage row 163 of Tx Juro is not row 163 of Total"

    AssertNe StoreFuturePositionKey("RTJ", "163", "RXZ4", "XS111"), _
             StoreFuturePositionKey("RTJ", "163", "RXZ4", "XS222"), _
        "one contract covering two bonds is two positions"

    AssertNe StoreFuturePositionKey("RTJ", "163", "RXZ4", "XS111"), _
             StoreFuturePositionKey("RTJ", "163", "OEZ4", "XS111"), _
        "two contracts on one coverage row are two positions"

    AssertEq StoreFuturePositionKey("rtj", "163", "rxz4", "xs111"), rtj, _
        "case does not make a new position"

End Sub


' -----------------------------------------------------------------------------
' THE FINGERPRINT
' -----------------------------------------------------------------------------
Sub TestFingerprint()

    Dim h1 As String, h2 As String

    AssertEq Len(StoreHash32("")), 8, "a hash is eight hex digits"
    AssertEq StoreHash32(""), "811c9dc5", "the empty hash is the FNV-1a offset basis"
    AssertEq StoreHash32("a"), "e40c292c", "FNV-1a of 'a'"
    AssertEq StoreHash32("foobar"), "bf9cf968", "FNV-1a of 'foobar'"

    AssertNe StoreHash32("XS111"), StoreHash32("XS112"), _
        "one changed character changes the hash"

    h1 = StoreFingerprintOfKeys(Array("A", "B", "C"))
    AssertEq StoreFingerprintOfKeys(Array("A", "B", "C")), h1, _
        "the same keys fingerprint the same"

    AssertNe StoreFingerprintOfKeys(Array("A", "B")), h1, _
        "a dropped position changes the fingerprint"

    AssertNe StoreFingerprintOfKeys(Array("A", "C", "B")), h1, _
        "a reordered book is a different book"

    ' The separator is what stops this collapsing.
    h2 = StoreFingerprintOfKeys(Array("AB", "C"))
    AssertNe StoreFingerprintOfKeys(Array("A", "BC")), h2, _
        "keys cannot run together into one"

    AssertEq StoreFingerprintOfKeys(Array()), StoreHash32(""), _
        "an empty book fingerprints as the empty hash"

End Sub


' -----------------------------------------------------------------------------
' THE DATE TEXT
' -----------------------------------------------------------------------------
Sub TestDateText()

    ' The duplicate-run guard compares the stored AsOfT0 to today's AS TEXT, so
    ' two desks on different regional settings must write the same string for
    ' the same day.  Format$ would not: its separators follow the machine.
    AssertEq StoreDateText(DateSerial(2026, 3, 7)), _
        "2026-03-07 00:00:00", "a date is unambiguous and sortable"
    AssertEq StoreDateText(""), "", "a blank date is blank, not 1899"
    AssertEq StoreDateText("not a date"), "", "unparseable is blank"

End Sub


' -----------------------------------------------------------------------------
' THE CSV ROUND TRIP
' -----------------------------------------------------------------------------
Sub TestCsvRoundTrip()

    Dim v(0 To 2, 0 To 2) As Variant
    Dim back As Variant

    ' `csv`, not `text`: Text is a VBA keyword (Open ... For Text) and using it
    ' as a variable is a COMPILE error, which in a headless host surfaces as a
    ' hang on a dialog nobody can see rather than as a message.
    Dim csv As String

    AssertEq StoreCsvField("plain"), "plain", "an ordinary value is not quoted"
    AssertEq StoreCsvField("a,b"), """a,b""", "a comma forces quotes"
    AssertEq StoreCsvField("say ""hi"""), """say """"hi""""""", _
        "a quote is doubled and the field quoted"
    AssertEq StoreCsvField(" pad "), """ pad """, _
        "leading and trailing space is preserved by quoting"
    AssertEq StoreCsvField(""), "", "empty stays empty, not a pair of quotes"

    ' A header row and two data rows, carrying every character that breaks a
    ' naive writer or a Split-on-comma reader.
    v(0, 0) = "ISIN": v(0, 1) = "Name": v(0, 2) = "Notional"
    v(1, 0) = "XS111": v(1, 1) = "BIG, S.A. 3.5% 2030": v(1, 2) = 1000000
    v(2, 0) = "XS222": v(2, 1) = "He said ""no""" & vbLf & "second line"
    v(2, 2) = ""

    csv = StoreCsvText(v)
    back = StoreCsvParse(csv)

    AssertEq UBound(back, 1) - LBound(back, 1) + 1, 3, "three records survive"
    AssertEq UBound(back, 2) - LBound(back, 2) + 1, 3, "three fields survive"
    AssertEq back(0, 0), "ISIN", "the header round-trips"
    AssertEq back(1, 1), "BIG, S.A. 3.5% 2030", _
        "a comma inside a name does not become a column break"
    AssertEq back(2, 1), "He said ""no""" & vbLf & "second line", _
        "an embedded quote and newline survive intact"
    AssertEq back(1, 2), "1000000", "a number round-trips as its text"
    AssertEq back(2, 2), "", "an empty field stays empty, not zero"

    ' An accented name, because bond names carry them and the file is UTF-8.
    AssertEq StoreCsvParse(StoreCsvText(Accented()))(1, 0), _
        "Varia" & ChrW$(231) & ChrW$(227) & "o", _
        "an accented value survives the round trip"

    ' A file with no trailing newline still has a last record.
    back = StoreCsvParse("a,b" & vbCrLf & "1,2")
    AssertEq back(1, 1), "2", "the last record needs no trailing newline"

End Sub


Function Accented() As Variant
    Dim v(0 To 1, 0 To 0) As Variant
    v(0, 0) = "Header"
    v(1, 0) = "Varia" & ChrW$(231) & ChrW$(227) & "o"
    Accented = v
End Function


' -----------------------------------------------------------------------------
' NUMBERS, AND THE THINGS THAT ONLY LOOK LIKE THEM
'
' The failure this section exists for leaves no trace: a DV01 written as
' "1234,56" on a comma-decimal machine comes back as TEXT, sits in the cell
' looking entirely correct, and contributes nothing to any sum.
' -----------------------------------------------------------------------------
Sub TestNumberText()

    AssertEq StoreNumberText(1234.56), "1234.56", "a decimal point, always"
    AssertEq StoreNumberText(1000000), "1000000", "no thousands separator"
    AssertEq StoreNumberText(-42), "-42", "a negative keeps its sign"
    AssertEq StoreNumberText(0.5), "0.5", "the leading zero is put back"
    AssertEq StoreNumberText(-0.5), "-0.5", "and on a negative fraction too"
    AssertEq StoreNumberText(0), "0", "zero is zero"

    ' What StoreCsvField does with each type, which is the writer's real
    ' contract: no quoting, because there is no separator to quote.
    AssertEq StoreCsvField(1234.56), "1234.56", "a number is written bare"
    AssertEq StoreCsvField(True), "TRUE", "a boolean is written as a word"
    AssertEq StoreCsvField(False), "FALSE", "and so is False"

End Sub


Sub TestTypedValue()

    AssertEq StoreTypedValue("1234.56"), 1234.56, "a decimal reads back numeric"
    AssertTrue IsNumeric(StoreTypedValue("1234.56")), _
        "and it really is a number, not text that looks like one", _
        TypeName(StoreTypedValue("1234.56"))
    AssertEq StoreTypedValue("-42"), -42, "a negative reads back"
    AssertEq StoreTypedValue("1E+20"), 1E+20, "an exponent reads back"
    AssertEq StoreTypedValue(""), "", "an empty field stays empty"
    AssertEq StoreTypedValue("TRUE"), True, "TRUE reads back as a boolean"

    ' Row_Valid feeds DashValidMask, which is --(N(...)=1).  N() of the text
    ' "1" is 0, so a Row_Valid restored as text masks every row out and the
    ' Dashboard reports a book of zero bonds.
    AssertTrue IsNumeric(StoreTypedValue("1")), _
        "Row_Valid restores numeric, so the dashboard mask still sees it", ""

    ' The other direction, and the more dangerous one.
    AssertEq StoreTypedValue("XS1234567890"), "XS1234567890", _
        "an ISIN stays text"
    AssertEq StoreTypedValue("0012"), "0012", _
        "a padded portfolio code is not the number 12"
    AssertEq StoreTypedValue("007"), "007", _
        "nor is a padded coverage row the number 7"
    AssertEq StoreTypedValue("1234,56"), "1234,56", _
        "a comma decimal is not a number - it is a file written wrong"
    AssertEq StoreTypedValue("3M"), "3M", "a tenor is not a number"
    AssertEq StoreTypedValue("-"), "-", "a lone dash is not a number"
    AssertEq StoreTypedValue("1.2.3"), "1.2.3", "a version is not a number"
    AssertEq StoreTypedValue("1E"), "1E", "a truncated exponent is not a number"

    ' Full circle: what the sheet held, through the file, back to a number.
    AssertEq StoreTypeArray(StoreCsvParse(StoreCsvText(Numbers())))(1, 0), _
        1234.56, "a DV01 survives write, read and typing as a number"
    AssertEq StoreTypeArray(StoreCsvParse(StoreCsvText(Numbers())))(0, 0), _
        "DV01", "the header stays text"
    AssertEq StoreTypeArray(StoreCsvParse(StoreCsvText(Numbers())))(1, 1), _
        "0012", "and the padded code beside it is still padded"

End Sub


Function Numbers() As Variant
    Dim v(0 To 1, 0 To 1) As Variant
    v(0, 0) = "DV01": v(0, 1) = "Portfolio"
    v(1, 0) = 1234.56: v(1, 1) = "0012"
    Numbers = v
End Function


' -----------------------------------------------------------------------------
' FOLDER ORDER
' -----------------------------------------------------------------------------
Sub TestRunFolderOrder()

    Dim a As String
    Dim b As String

    a = StoreRunFolderName(DateSerial(2026, 8, 31) + TimeSerial(9, 30, 0))
    b = StoreRunFolderName(DateSerial(2026, 9, 1) + TimeSerial(9, 0, 0))

    AssertEq a, "run_20260831_093000", "the name is fixed-width and big-endian"
    AssertTrue a < b, "September sorts after August", a & " vs " & b

    ' The trap the fixed width exists for: an unpadded month or hour would sort
    ' 2026-8-31 after 2026-09-01, and "the latest run" would serve the wrong day.
    AssertTrue StoreRunFolderName(DateSerial(2026, 9, 1)) > _
               StoreRunFolderName(DateSerial(2026, 8, 31)), _
        "a single-digit month still sorts correctly", ""

    AssertTrue StoreRunFolderName(DateSerial(2026, 3, 7) + TimeSerial(9, 5, 0)) < _
               StoreRunFolderName(DateSerial(2026, 3, 7) + TimeSerial(10, 0, 0)), _
        "09:05 sorts before 10:00", ""

End Sub


Sub TestKeyLengthFits()

    Dim k As String

    ' A realistic worst case: the longest components the sheets can hold.
    k = StoreFuturePositionKey(String$(4, "R"), String$(32, "9"), _
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
    TestDateText
    TestCsvRoundTrip
    TestNumberText
    TestTypedValue
    TestRunFolderOrder
    TestKeyLengthFits

    Print #gOut, "----"
    Print #gOut, "passed=" & gPassed & " failed=" & gFailed
    Close #gOut
End Sub
