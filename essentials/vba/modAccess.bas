Option Explicit

' =============================================================================
' modAccess  -  the run store
'
' WHAT THIS IS FOR
'
'   The workbook holds one day.  Load tomorrow and today is gone: the bond
'   population, the swap mapping, the futures coverage - all of it overwritten
'   in place, because a sheet is a single mutable value and there is nowhere
'   else for it to go.  Every question that starts "what did we hold on..." is
'   therefore unanswerable, and every answer the workbook gives is unauditable,
'   since the inputs that produced it no longer exist.
'
'   This module gives them somewhere else to go: an Access database, one row
'   per position per run, keyed so that a run can never contain the same
'   position twice and can never be confused with another run.
'
' THE THREE GUARANTEES
'
'   1  RUN SEPARATION.  Every stored row carries RunID.  Runs never edit each
'      other; a save inserts and nothing else.  History is therefore a property
'      of the schema rather than a discipline somebody has to keep.
'
'   2  NO DUPLICATE POSITION WITHIN A RUN.  Each position gets a PositionKey
'      built from the fields that make it distinct, and (RunID, PositionKey) is
'      the primary key.  Access refuses the second insert - the guard is in the
'      database, not only in the loop that writes it.  A refused row is not
'      silently dropped: it is written to Run_Issue with its source row.
'
'   3  NO DUPLICATE RUN.  Saving an unchanged workbook twice would otherwise
'      produce two runs holding identical positions.  Each run carries a
'      fingerprint - a hash over every position key it holds - and a save whose
'      fingerprint already exists for the same reporting date is refused.
'
' WHAT IT DOES NOT DO
'
'   It does not store computed columns.  Everything a formula produces is
'   derived from the inputs stored here, so keeping it would be keeping the
'   same fact twice, and restoring it would overwrite the formula that makes
'   it.  Fact_PnlAttribution in access/schema.sql is where computed output will
'   land when that phase arrives; see docs/ACCESS_ARCHITECTURE.md.
'
'   It does not know where a column is.  modPNL answers that, through
'   PnlPositionFieldMap, so a column move needs no edit here.
'
' WHERE THE PIECES ARE
'
'   Config!B41   database path        (blank = PNL_Data.accdb beside the book)
'   Config!B42   autosave TRUE/FALSE  (save at the end of Button 1 and 2)
'   Config!B43   last saved RunID     (written here, do not edit)
'   Config!B44   last store message   (written here, do not edit)
'   Config!B45   runs to keep         (0 = keep everything)
'
' Buttons:  Access_TestConnection, Access_SaveCurrentRun, Access_ShowRunHistory,
'           Access_LoadRunIntoSheets, Access_PurgeOldRuns.
'
' Late binding throughout: no reference has to be added to the VBE, and a
' machine without ACE fails with a message that says so rather than refusing to
' compile the whole project.
' =============================================================================


' --- Config cells this module owns -------------------------------------------
Private Const ACC_SH_CONFIG As String = "Config"
Private Const ACC_CFG_DB_PATH  As String = "B41"
Private Const ACC_CFG_AUTOSAVE As String = "B42"
Private Const ACC_CFG_LAST_RUN As String = "B43"
Private Const ACC_CFG_STATUS   As String = "B44"
Private Const ACC_CFG_KEEP     As String = "B45"
Private Const ACC_CFG_RUN_STAMP As String = "B46"

Private Const ACC_SH_HISTORY As String = "Run_History"

Private Const ACC_DEFAULT_DB As String = "PNL_Data.accdb"

' ONE DATABASE PER RUN
'
' Each run writes its own .accdb, named for the moment the data was RETRIEVED -
' PNL_Run_20260901_173000.accdb - and the index database records where they all
' are.  Three things follow, and the third is the reason:
'
'   A stored run is immutable.  Nothing a later run does can touch it, so the
'   audit trail cannot be edited by accident.
'
'   The 2 GB per-file ceiling stops being a horizon.  It applied to the whole
'   history in one file; it now applies to one day.
'
'   A re-save of the SAME retrieval lands on the SAME file.  The stamp is the
'   retrieval time, not the save time, so pressing the button twice on one pull
'   is idempotent rather than producing a second database that differs only in
'   when somebody clicked.
'
' The cost is that history is an attach-and-union across files rather than one
' query.  AccIndexDatabasePath is what makes that tractable: it holds one row
' per run, so "which files do I need" is a query rather than a directory scan.
Private Const ACC_RUN_DB_PREFIX As String = "PNL_Run_"
Private Const ACC_INDEX_DB      As String = "PNL_Runs_Index.accdb"

' The retrieval stamp of the run in progress, yyyymmdd_hhnnss.  Set once by
' AccBeginRun and reused by every save in that run, so the bond autosave and the
' hedge autosave write to one database rather than two.
Private mRunStamp As String

' Bumped whenever a table or column is added below.  Stamped into Meta_Schema
' after a successful ensure, so a database built by an older module is a
' visible fact rather than a missing-column error somewhere later.
Private Const ACC_SCHEMA_VERSION As Long = 1

' ADO constants, spelled out rather than referenced: late binding means the
' type library is not loaded and adOpenStatic and friends are not defined.
Private Const ACC_ADO_OPEN_KEYSET As Long = 1
Private Const ACC_ADO_OPEN_STATIC As Long = 3
Private Const ACC_ADO_LOCK_OPTIMISTIC As Long = 3
Private Const ACC_ADO_LOCK_READONLY As Long = 1
Private Const ACC_ADO_SCHEMA_TABLES As Long = 20
Private Const ACC_ADO_CMD_TEXT As Long = 1

' Access refuses a duplicate primary key with this provider error.  Caught by
' number rather than by message, which is localised.
Private Const ACC_ERR_DUPLICATE As Long = -2147217887

' A position key longer than the column is a truncation waiting to become a
' false duplicate, so it is refused instead.  Matches TEXT(200) in schema.sql.
Private Const ACC_MAX_KEY_LEN As Long = 200

' What joins the components of a position key.
'
' The separator has to be a character none of the components can contain.  "|"
' would be wrong for swaps, whose mapping-row id is itself built out of "|" -
' two different mappings could then produce one key by accident.  Chr(11) is
' equally impossible in an ISIN, a portfolio code or a contract code, which is
' all it has to be.
'
' vbVerticalTab, not ChrW$(1): a Const must be a constant EXPRESSION, so
' `= ChrW$(1)` is a COMPILE error - the whole project refuses to build, at a
' line that has nothing to do with what you were editing.
Private Const ACC_KEY_SEP As String = vbVerticalTab

' Runaway guard on the issue log: a genuinely broken source could otherwise
' write one row per position, twice over.
Private Const ACC_MAX_ISSUES As Long = 500


' =============================================================================
' THE TABLE SPECIFICATIONS
'
' One array per table, each entry "ColumnName|TYPE".  These are the single
' source of truth for what modAccess creates, and tools/check_access_schema.py
' compares them against access/schema.sql on every test run, so the DDL you can
' read and the DDL that actually runs cannot drift apart.
'
' Keeping them as data rather than as CREATE TABLE strings is what makes
' "add the columns this database is missing" possible at all: the same array
' that builds a new table tells an existing one what it lacks.
' =============================================================================

Private Function AccTableNames() As Variant
    AccTableNames = Array("Run", "Pos_Bond", "Pos_Swap", "Pos_Future", _
                          "Run_Issue", "Meta_Schema")
End Function


Private Function AccTableSpec(ByVal tableName As String) As Variant

    Select Case tableName

    Case "Run"
        AccTableSpec = Array( _
            "RunID|AUTOINCREMENT", _
            "RunStartedUtc|DATETIME", _
            "RunFinishedUtc|DATETIME", _
            "AsOfT0|DATETIME", _
            "AsOfTM1|DATETIME", _
            "RunBy|TEXT(64)", _
            "ModuleVersion|TEXT(32)", _
            "ExcelBitness|TEXT(8)", _
            "RunStatus|TEXT(16)", _
            "RunFingerprint|TEXT(64)", _
            "BondCount|LONG", _
            "SwapCount|LONG", _
            "FutureCount|LONG", _
            "SupersedesRunID|LONG", _
            "Notes|LONGTEXT")

    Case "Pos_Bond"
        AccTableSpec = Array( _
            "RunID|LONG", _
            "PositionKey|TEXT(200)", _
            "ISIN|TEXT(12)", _
            "InstrumentName|TEXT(128)", _
            "CCY|TEXT(8)", _
            "Coupon|DOUBLE", _
            "CouponFreqDesc|TEXT(32)", _
            "Maturity|DATETIME", _
            "Notional|DOUBLE", _
            "AcctgCat|TEXT(32)", _
            "Portfolio|TEXT(32)", _
            "BookVal|DOUBLE", _
            "SourceRow|LONG")

    Case "Pos_Swap"
        AccTableSpec = Array( _
            "RunID|LONG", _
            "PositionKey|TEXT(200)", _
            "CCY|TEXT(8)", _
            "Notional|DOUBLE", _
            "StartDate|DATETIME", _
            "EndDate|DATETIME", _
            "PayFixed|TEXT(4)", _
            "Portfolio|TEXT(64)", _
            "LinkedISIN|TEXT(12)", _
            "Cpty|TEXT(128)", _
            "BBGSwapDirectID|TEXT(64)", _
            "BBGFixedLegID|TEXT(64)", _
            "BBGFloatLegID|TEXT(64)", _
            "SwapIDSource|TEXT(16)", _
            "CoverageRelation|TEXT(64)", _
            "SwapMapSourceRow|TEXT(32)", _
            "SwapMapClass|TEXT(16)", _
            "SwapMapStatus|TEXT(128)", _
            "MapNotional|DOUBLE", _
            "MapCCY|TEXT(8)", _
            "MapCounterparty|TEXT(128)", _
            "NotionalFinal|DOUBLE", _
            "NotionalSource|TEXT(32)", _
            "SourceRow|LONG")

    Case "Pos_Future"
        AccTableSpec = Array( _
            "RunID|LONG", _
            "PositionKey|TEXT(200)", _
            "ContractCode|TEXT(64)", _
            "Exchange|TEXT(32)", _
            "CCY|TEXT(8)", _
            "Contracts|DOUBLE", _
            "Portfolio|TEXT(64)", _
            "LinkedISIN|TEXT(12)", _
            "ImportStatus|TEXT(128)", _
            "CoverageSourceRow|TEXT(32)", _
            "CoverageRelation|TEXT(64)", _
            "FutureLabel|TEXT(128)", _
            "HedgeType|TEXT(32)", _
            "CoverageStartDate|DATETIME", _
            "CoverageInfoD|TEXT(128)", _
            "HedgeSource|TEXT(4)", _
            "CoverageBPV|DOUBLE", _
            "SourceRow|LONG")

    Case "Run_Issue"
        AccTableSpec = Array( _
            "RunID|LONG", _
            "IssueSeq|LONG", _
            "Severity|TEXT(8)", _
            "TableName|TEXT(32)", _
            "PositionKey|TEXT(200)", _
            "SourceRow|LONG", _
            "Detail|TEXT(255)")

    Case "Meta_Schema"
        AccTableSpec = Array( _
            "SchemaVersion|LONG", _
            "AppliedUtc|DATETIME", _
            "AppliedBy|TEXT(64)")

    Case Else
        AccTableSpec = Array()

    End Select

End Function


' The primary key of a table, as the column list for a CONSTRAINT clause.
Private Function AccTablePK(ByVal tableName As String) As String

    Select Case tableName
    Case "Run":         AccTablePK = "RunID"
    Case "Pos_Bond":    AccTablePK = "RunID, PositionKey"
    Case "Pos_Swap":    AccTablePK = "RunID, PositionKey"
    Case "Pos_Future":  AccTablePK = "RunID, PositionKey"
    Case "Run_Issue":   AccTablePK = "RunID, IssueSeq"
    Case "Meta_Schema": AccTablePK = "SchemaVersion"
    Case Else:          AccTablePK = ""
    End Select

End Function


' Secondary indexes, "IndexName|Table|Columns".  Every one of them serves a
' query this module actually runs; an index nothing reads is pure write cost.
Private Function AccIndexSpecs() As Variant
    AccIndexSpecs = Array( _
        "IX_Run_AsOfT0|Run|AsOfT0", _
        "IX_Run_Fingerprint|Run|AsOfT0, RunFingerprint", _
        "IX_Pos_Bond_ISIN|Pos_Bond|RunID, ISIN", _
        "IX_Pos_Swap_Bond|Pos_Swap|RunID, LinkedISIN", _
        "IX_Pos_Future_Bond|Pos_Future|RunID, LinkedISIN")
End Function


' CREATE TABLE, built from the spec above.  Pure string assembly on purpose:
' it is the one part of the schema machinery that can be tested without a
' database, and a malformed DDL is otherwise only visible as a provider error
' on a machine you are not sitting at.
Private Function AccCreateTableSql(ByVal tableName As String) As String

    Dim spec As Variant
    Dim parts() As String
    Dim pk As String
    Dim i As Long

    spec = AccTableSpec(tableName)
    If UBound(spec) < LBound(spec) Then Exit Function

    ReDim parts(LBound(spec) To UBound(spec))

    For i = LBound(spec) To UBound(spec)
        parts(i) = "[" & AccSpecName(CStr(spec(i))) & "] " & _
                   AccSpecType(CStr(spec(i)))
    Next i

    pk = AccTablePK(tableName)

    AccCreateTableSql = "CREATE TABLE [" & tableName & "] (" & _
        Join(parts, ", ") & _
        IIf(Len(pk) = 0, "", ", CONSTRAINT [PK_" & tableName & _
            "] PRIMARY KEY (" & pk & ")") & ")"

End Function


Private Function AccSpecName(ByVal entry As String) As String
    AccSpecName = Split(entry, "|")(0)
End Function


Private Function AccSpecType(ByVal entry As String) As String
    AccSpecType = Split(entry, "|")(1)
End Function


' =============================================================================
' SQL LITERALS
'
' Every one of these is locale-proof by construction, which is not fussiness:
' on a machine with a comma decimal separator CStr(1.5) is "1,5", and
' "VALUES (1,5)" is a syntactically valid INSERT into the wrong columns.  It
' does not fail - it writes the wrong number and says nothing.
'
' Str$ always emits a dot (that is its documented difference from CStr), and
' the date literal is assembled from Year/Month/Day rather than through Format,
' whose separators follow the machine too.
'
' Positions are not written through these at all - they go through a recordset,
' which is both faster and type-safe.  These build the WHERE clauses and the
' handful of single-row statements.
' =============================================================================

Public Function AccSqlText(ByVal v As Variant, ByVal maxLen As Long) As String

    Dim s As String

    If IsNull(v) Or IsEmpty(v) Then
        AccSqlText = "NULL"
        Exit Function
    End If

    If IsError(v) Then
        AccSqlText = "NULL"
        Exit Function
    End If

    s = Trim$(CStr(v))

    If Len(s) = 0 Then
        AccSqlText = "NULL"
        Exit Function
    End If

    If maxLen > 0 Then
        If Len(s) > maxLen Then s = Left$(s, maxLen)
    End If

    AccSqlText = "'" & Replace(s, "'", "''") & "'"

End Function


Public Function AccSqlNum(ByVal v As Variant) As String

    If IsNull(v) Or IsEmpty(v) Then
        AccSqlNum = "NULL"
        Exit Function
    End If

    If IsError(v) Then
        AccSqlNum = "NULL"
        Exit Function
    End If

    If Not IsNumeric(v) Then
        AccSqlNum = "NULL"
        Exit Function
    End If

    ' Str$ prefixes a positive number with a space and never uses a comma.
    AccSqlNum = Trim$(Str$(CDbl(v)))

End Function


Public Function AccSqlDate(ByVal v As Variant) As String

    Dim d As Date

    If IsNull(v) Or IsEmpty(v) Then
        AccSqlDate = "NULL"
        Exit Function
    End If

    If IsError(v) Then
        AccSqlDate = "NULL"
        Exit Function
    End If

    ' The empty-string test comes FIRST and is not redundant.  A blank cell
    ' reaches here as Empty, which the check above catches, but a cell holding
    ' a zero-length string reaches here as "" - and IsDate("") is False in VBA
    ' and True in some Basic hosts, where it would become 30 December 1899.  A
    ' silent 1899 in AsOfT0 breaks the duplicate guard for every future run.
    If Len(Trim$(CStr(v))) = 0 Then
        AccSqlDate = "NULL"
        Exit Function
    End If

    If Not IsDate(v) Then
        AccSqlDate = "NULL"
        Exit Function
    End If

    d = CDate(v)

    AccSqlDate = "#" & AccPad(Year(d), 4) & "-" & AccPad(Month(d), 2) & "-" & _
                 AccPad(Day(d), 2) & " " & AccPad(Hour(d), 2) & ":" & _
                 AccPad(Minute(d), 2) & ":" & AccPad(Second(d), 2) & "#"

End Function


Private Function AccPad(ByVal n As Long, ByVal padTo As Long) As String

    Dim s As String

    s = Trim$(Str$(n))

    Do While Len(s) < padTo
        s = "0" & s
    Loop

    AccPad = s

End Function


' =============================================================================
' THE RUN FINGERPRINT
'
' FNV-1a over the run's position keys, one 32-bit word rendered as eight hex
' digits per kind, concatenated.  Two saves of an unchanged book therefore
' produce the same string and the second is refused.
'
' Done in Double arithmetic rather than Long: VBA's Long is signed, 32-bit
' multiplication overflows it on the second character, and VBA's Mod operator
' coerces to Long and would overflow too.  Every intermediate here stays under
' 2^53, where a Double holds integers exactly.
'
' Hashed per UTF-16 code unit rather than per byte, which is not the textbook
' FNV-1a - it does not need to be.  It needs to be deterministic and to change
' when the positions change, and tests/test_access_store.bas pins it.
' =============================================================================

Public Function AccHash32(ByVal s As String) As String

    Dim h As Double
    Dim i As Long
    Dim c As Long

    h = 2166136261#

    For i = 1 To Len(s)

        ' AscW returns a SIGNED Integer, so anything above U+7FFF comes back
        ' negative.  Folding it back up is what makes the hash the same number
        ' here and in the test harness, which uses a host whose Asc is not
        ' signed - a masking AND would quietly differ between the two.
        c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536

        h = AccXor32(h, CDbl(c))
        h = AccMul32(h, 16777619#)
    Next i

    AccHash32 = AccHex32(h)

End Function


Private Function AccXor32(ByVal a As Double, ByVal b As Double) As Double

    Dim aHi As Long
    Dim aLo As Long
    Dim bHi As Long
    Dim bLo As Long

    aHi = CLng(Int(a / 65536#))
    aLo = CLng(a - aHi * 65536#)
    bHi = CLng(Int(b / 65536#))
    bLo = CLng(b - bHi * 65536#)

    AccXor32 = (aHi Xor bHi) * 65536# + (aLo Xor bLo)

End Function


Private Function AccMul32(ByVal a As Double, ByVal b As Double) As Double

    Dim aHi As Double
    Dim aLo As Double
    Dim hi As Double
    Dim r As Double

    aHi = Int(a / 65536#)
    aLo = a - aHi * 65536#

    ' Only the low 16 bits of the high half survive a 32-bit truncation, so
    ' they are reduced BEFORE being shifted back up - otherwise the product
    ' leaves the range a Double represents exactly.
    hi = aHi * b
    hi = hi - Int(hi / 65536#) * 65536#

    r = hi * 65536# + aLo * b

    AccMul32 = r - Int(r / 4294967296#) * 4294967296#

End Function


Private Function AccHex32(ByVal v As Double) As String

    Dim digits As String
    Dim outS As String
    Dim x As Double
    Dim d As Long
    Dim i As Long

    digits = "0123456789abcdef"
    x = v

    For i = 1 To 8
        d = CLng(x - Int(x / 16#) * 16#)
        outS = Mid$(digits, d + 1, 1) & outS
        x = Int(x / 16#)
    Next i

    AccHex32 = outS

End Function


' The fingerprint of one kind's key list.  Separated from AccHash32 so the
' joining rule - a newline between keys, so "AB" + "C" cannot collide with
' "A" + "BC" - is stated once and tested.
Public Function AccFingerprintOfKeys(ByVal keys As Variant) As String

    Dim s As String
    Dim i As Long

    If Not IsArray(keys) Then
        AccFingerprintOfKeys = AccHash32("")
        Exit Function
    End If

    If UBound(keys) < LBound(keys) Then
        AccFingerprintOfKeys = AccHash32("")
        Exit Function
    End If

    For i = LBound(keys) To UBound(keys)
        s = s & CStr(keys(i)) & vbLf
    Next i

    AccFingerprintOfKeys = AccHash32(s)

End Function


' =============================================================================
' POSITION KEYS
'
' What makes a position distinct, as one string.  These decide what "the same
' position" means, so they are pure functions with no sheet in sight and they
' are the most heavily tested thing in this module.
'
' The separator is a character none of the components can contain.  Using "|"
' would be wrong for swaps, whose mapping-row id is itself built out of "|" -
' two different mappings could then produce one key by accident.
' =============================================================================

Public Function AccBondPositionKey( _
    ByVal isin As Variant, _
    ByVal portfolio As Variant, _
    ByVal acctgCat As Variant) As String

    ' AcctgCat is in the key because one ISIN can be held in two accounting
    ' categories in the same portfolio, and they are two positions with two
    ' book values.  Keying on (ISIN, Portfolio) alone would keep whichever
    ' arrived first and refuse the other as a duplicate.
    AccBondPositionKey = AccJoinKey(Array( _
        UCase$(AccClean(isin)), _
        UCase$(AccClean(portfolio)), _
        UCase$(AccClean(acctgCat))))

End Function


Public Function AccSwapPositionKey(ByVal mappingRowId As Variant) As String

    ' Swaps!A already holds a unique mapping-row id, built by
    ' AppendSwapMapRowsToSwaps out of source|id|LinkedISIN|row precisely so
    ' that one swap hedging two bonds is two rows.  Re-deriving a key here
    ' would be a second opinion on a question the sheet has already answered.
    AccSwapPositionKey = AccClean(mappingRowId)

End Function


Public Function AccFuturePositionKey( _
    ByVal hedgeSource As Variant, _
    ByVal coverageSourceRow As Variant, _
    ByVal contractCode As Variant, _
    ByVal linkedIsin As Variant) As String

    ' HedgeSource leads because it is what scopes the rest: the two Hedge Risco
    ' books number their coverage rows independently, so row 163 means one
    ' thing in Tx Juro and another in Total.  Drop it and one of the two is
    ' discarded as a duplicate of the other - silently, and the bond it hedged
    ' reads as unhedged.
    AccFuturePositionKey = AccJoinKey(Array( _
        UCase$(AccClean(hedgeSource)), _
        AccClean(coverageSourceRow), _
        UCase$(AccClean(contractCode)), _
        UCase$(AccClean(linkedIsin))))

End Function


Private Function AccJoinKey(ByVal parts As Variant) As String

    Dim s As String
    Dim i As Long

    For i = LBound(parts) To UBound(parts)
        If i > LBound(parts) Then s = s & ACC_KEY_SEP
        s = s & CStr(parts(i))
    Next i

    AccJoinKey = s

End Function


Private Function AccClean(ByVal v As Variant) As String

    If IsNull(v) Or IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function

    AccClean = Trim$(CStr(v))

End Function


' The key for one row of a PnlPositionsForStore array, by kind.  This is the
' only place that knows which fields feed which key, so adding a component is
' one edit here plus one in the key function itself.
Private Function AccKeyForRow( _
    ByVal kind As String, _
    ByVal data As Variant, _
    ByVal rowIndex As Long, _
    ByVal cols As Object) As String

    Select Case UCase$(kind)

    Case "BOND"
        AccKeyForRow = AccBondPositionKey( _
            AccCell(data, rowIndex, cols, "ISIN"), _
            AccCell(data, rowIndex, cols, "Portfolio"), _
            AccCell(data, rowIndex, cols, "AcctgCat"))

    Case "SWAP"
        AccKeyForRow = AccSwapPositionKey( _
            AccCell(data, rowIndex, cols, "PositionKey"))

    Case "FUTURE"
        AccKeyForRow = AccFuturePositionKey( _
            AccCell(data, rowIndex, cols, "HedgeSource"), _
            AccCell(data, rowIndex, cols, "CoverageSourceRow"), _
            AccCell(data, rowIndex, cols, "ContractCode"), _
            AccCell(data, rowIndex, cols, "LinkedISIN"))

    End Select

End Function


Private Function AccCell( _
    ByVal data As Variant, _
    ByVal rowIndex As Long, _
    ByVal cols As Object, _
    ByVal fieldName As String) As Variant

    If Not cols.Exists(fieldName) Then Exit Function

    AccCell = data(rowIndex, cols(fieldName))

End Function


' Field name -> column index, for one PnlPositionsForStore array.
Private Function AccColumnIndex(ByVal data As Variant) As Object

    Dim d As Object
    Dim j As Long
    Dim nm As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    For j = LBound(data, 2) To UBound(data, 2)
        nm = Trim$(CStr(data(LBound(data, 1), j)))
        If Len(nm) > 0 Then
            If Not d.Exists(nm) Then d.Add nm, j
        End If
    Next j

    Set AccColumnIndex = d

End Function


' =============================================================================
' THE CONNECTION
'
' Two things go wrong on a desk machine and both are worth naming precisely.
'
' The provider is not registered.  ACE is a separate install from Office, and
' it must match Excel's BITNESS - 64-bit Excel cannot load the 32-bit
' redistributable, and the error message it gives ("provider is not
' registered on the local machine") says nothing about bitness at all.  So the
' bitness is reported here, unasked.
'
' The database is on a network share that is slow or gone.  ACE's default
' timeouts are long enough that Excel looks frozen, so they are set short and
' the failure is a message rather than a hang.
' =============================================================================

Private Function AccProviders() As Variant
    ' 16.0 first: it is what a current Office ships with.  12.0 is the older
    ' redistributable and is still what most desks actually have.
    AccProviders = Array("Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0")
End Function


Private Function AccConnString( _
    ByVal provider As String, _
    ByVal dbPath As String) As String

    AccConnString = "Provider=" & provider & ";Data Source=" & dbPath & ";"

End Function


' Where the database lives.  A blank Config cell means "beside the workbook",
' which is the right default: the two travel together, and nobody has to type a
' path to get started.  A relative path is resolved against the workbook too.
' The folder the run databases live in.
'
' Config!B41 held a FILE path in the single-database build.  It is read as a
' folder now, and a value that still names a .accdb is taken as "the folder that
' file was in" - so an existing configured path keeps working and lands the run
' databases beside the database it used to name, rather than failing or writing
' somewhere surprising.
Private Function AccRunFolder() As String

    Dim ws As Worksheet
    Dim configured As String
    Dim folder As String

    folder = ThisWorkbook.Path
    If Len(folder) = 0 Then folder = CurDir$

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ACC_SH_CONFIG)
    On Error GoTo 0

    If Not ws Is Nothing Then
        configured = Trim$(CStr(ws.Range(ACC_CFG_DB_PATH).value))
    End If

    If Len(configured) = 0 Then
        AccRunFolder = folder
        Exit Function
    End If

    ' A value naming a database file: take its folder.
    If InStr(1, configured, ".accdb", vbTextCompare) > 0 _
       Or InStr(1, configured, ".mdb", vbTextCompare) > 0 Then
        If InStrRev(configured, Application.PathSeparator) > 0 Then
            configured = Left$(configured, _
                InStrRev(configured, Application.PathSeparator) - 1)
        Else
            AccRunFolder = folder
            Exit Function
        End If
    End If

    ' A bare name, or anything without a drive or UNC prefix, is relative.
    If InStr(configured, ":") = 0 And Left$(configured, 2) <> "\\" Then
        AccRunFolder = folder & Application.PathSeparator & configured
    Else
        AccRunFolder = configured
    End If

    Do While Right$(AccRunFolder, 1) = Application.PathSeparator
        AccRunFolder = Left$(AccRunFolder, Len(AccRunFolder) - 1)
    Loop

End Function


' Start a run, and fix the stamp every save in it will use.
'
' Called with the moment the data was retrieved.  Passing 0 means "now", which is
' right for a load that has just finished and wrong for a re-save - so a stamp
' already set is kept unless the caller forces a new one.
Public Sub AccBeginRun( _
    Optional ByVal retrievedAt As Date = 0, _
    Optional ByVal forceNew As Boolean = False)

    Dim ws As Worksheet
    Dim stamped As Date

    If Len(mRunStamp) > 0 And Not forceNew Then Exit Sub

    stamped = retrievedAt
    If stamped = 0 Then stamped = AccUtcNow()

    mRunStamp = Format$(stamped, "yyyymmdd") & "_" & Format$(stamped, "hhnnss")

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ACC_SH_CONFIG)
    If Not ws Is Nothing Then ws.Range(ACC_CFG_RUN_STAMP).value = mRunStamp
    On Error GoTo 0

End Sub


' The stamp of the run in progress.  Recovered from Config!B46 when the module
' has been reset - a lost stamp would otherwise split one run across two files.
Public Function AccRunStamp() As String

    Dim ws As Worksheet

    If Len(mRunStamp) = 0 Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(ACC_SH_CONFIG)
        If Not ws Is Nothing Then
            mRunStamp = Trim$(CStr(ws.Range(ACC_CFG_RUN_STAMP).value))
        End If
        On Error GoTo 0
    End If

    If Len(mRunStamp) = 0 Then AccBeginRun

    AccRunStamp = mRunStamp

End Function


' The database for the run in progress.
Public Function AccDatabasePath() As String

    AccDatabasePath = AccRunFolder() & Application.PathSeparator & _
                      ACC_RUN_DB_PREFIX & AccRunStamp() & ".accdb"

End Function


' The index: one row per run database, so history does not need a folder scan.
Public Function AccIndexDatabasePath() As String

    AccIndexDatabasePath = AccRunFolder() & Application.PathSeparator & ACC_INDEX_DB

End Function


' Open a connection, creating the file and any missing tables on the way.
' Returns Nothing and shows one message on failure - callers check for Nothing
' rather than trapping.
Private Function AccOpen( _
    Optional ByVal ensureSchema As Boolean = True, _
    Optional ByVal quiet As Boolean = False) As Object

    Dim conn As Object
    Dim providers As Variant
    Dim dbPath As String
    Dim lastErr As String
    Dim i As Long

    dbPath = AccDatabasePath()
    providers = AccProviders()

    If Len(Dir$(dbPath)) = 0 Then
        If Not AccCreateDatabaseFile(dbPath, quiet) Then Exit Function
    End If

    For i = LBound(providers) To UBound(providers)

        On Error Resume Next
        Set conn = CreateObject("ADODB.Connection")
        conn.ConnectionTimeout = 15
        conn.CommandTimeout = 120
        conn.Open AccConnString(CStr(providers(i)), dbPath)

        If Err.Number <> 0 Then
            lastErr = Err.Description
            Err.Clear
            Set conn = Nothing
        End If
        On Error GoTo 0

        If Not conn Is Nothing Then Exit For

    Next i

    If conn Is Nothing Then

        AccSetStatus "Could not open " & dbPath & ": " & lastErr

        ' Quiet is for the autosave.  A desk that has not installed the engine
        ' yet would otherwise get a modal provider error on every press of
        ' Button 1 and Button 2 - twice a day, forever, about something the
        ' buttons do not depend on.  Config!B44 still says what happened, and
        ' Access_TestConnection is the button that answers loudly on purpose.
        If quiet Then Exit Function

        MsgBox "Could not open the Access database." & vbCrLf & vbCrLf & _
               "Path:   " & dbPath & vbCrLf & _
               "Excel:  " & AccBitness() & vbCrLf & _
               "Error:  " & lastErr & vbCrLf & vbCrLf & _
               "The Access Database Engine must be installed and must match " & _
               "Excel's bitness (" & AccBitness() & ").", _
               vbCritical, "Access store"
        Exit Function
    End If

    If ensureSchema Then AccEnsureSchema conn

    Set AccOpen = conn

End Function


' Create an empty .accdb.  ADOX is the normal route; DAO is the fallback for
' machines where msadox is missing but the engine itself is fine.
Private Function AccCreateDatabaseFile( _
    ByVal dbPath As String, _
    Optional ByVal quiet As Boolean = False) As Boolean

    Dim cat As Object
    Dim dbe As Object
    Dim providers As Variant
    Dim i As Long

    providers = AccProviders()

    For i = LBound(providers) To UBound(providers)
        On Error Resume Next
        Set cat = CreateObject("ADOX.Catalog")
        If Not cat Is Nothing Then
            cat.Create AccConnString(CStr(providers(i)), dbPath)
        End If
        Err.Clear
        Set cat = Nothing
        On Error GoTo 0
        If Len(Dir$(dbPath)) > 0 Then Exit For
    Next i

    If Len(Dir$(dbPath)) = 0 Then
        On Error Resume Next
        Set dbe = CreateObject("DAO.DBEngine.120")
        If Not dbe Is Nothing Then
            dbe.CreateDatabase dbPath, ";LANGID=0x0409;CP=1252;COUNTRY=0"
        End If
        Err.Clear
        Set dbe = Nothing
        On Error GoTo 0
    End If

    If Len(Dir$(dbPath)) = 0 Then

        AccSetStatus "Could not create " & dbPath & "."
        If quiet Then Exit Function

        MsgBox "Could not create the Access database file." & vbCrLf & vbCrLf & _
               "Path:  " & dbPath & vbCrLf & _
               "Excel: " & AccBitness() & vbCrLf & vbCrLf & _
               "Check the folder is writable and that the Access Database " & _
               "Engine matching " & AccBitness() & " Excel is installed.", _
               vbCritical, "Access store"
        Exit Function
    End If

    AccCreateDatabaseFile = True

End Function


Private Function AccBitness() As String
    #If Win64 Then
        AccBitness = "64-bit"
    #Else
        AccBitness = "32-bit"
    #End If
End Function


' =============================================================================
' MAKING THE DATABASE MATCH THE MODULE
'
' Run on every connect, and cheap when there is nothing to do: one schema
' query, one dictionary of table names, then per table one lookup.
'
' It creates what is missing and adds columns a table lacks; it never drops or
' retypes anything.  A destructive migration is a decision, not a side effect
' of opening a workbook, so if one is ever needed it will be a macro somebody
' has to run on purpose.
' =============================================================================

Private Sub AccEnsureSchema(ByVal conn As Object)

    Dim existing As Object
    Dim tableNames As Variant
    Dim indexSpecs As Variant
    Dim i As Long
    Dim tbl As String

    Set existing = AccExistingTables(conn)
    tableNames = AccTableNames()

    For i = LBound(tableNames) To UBound(tableNames)

        tbl = CStr(tableNames(i))

        If Not existing.Exists(tbl) Then
            AccExecuteQuiet conn, AccCreateTableSql(tbl)
        Else
            AccAddMissingColumns conn, tbl
        End If

    Next i

    indexSpecs = AccIndexSpecs()

    For i = LBound(indexSpecs) To UBound(indexSpecs)
        AccEnsureIndex conn, CStr(indexSpecs(i))
    Next i

    AccStampSchemaVersion conn

End Sub


Private Function AccExistingTables(ByVal conn As Object) As Object

    Dim d As Object
    Dim rs As Object

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    On Error Resume Next
    Set rs = conn.OpenSchema(ACC_ADO_SCHEMA_TABLES)

    If Not rs Is Nothing Then
        Do While Not rs.EOF
            If UCase$(CStr(rs.fields("TABLE_TYPE").value)) = "TABLE" Then
                d(CStr(rs.fields("TABLE_NAME").value)) = True
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If
    Err.Clear
    On Error GoTo 0

    Set AccExistingTables = d

End Function


' Add whatever the table specification has and the table does not.
'
' This is the whole "moving a column is easy" promise on the database side:
' a new stored field is one entry in AccTableSpec and one in
' PnlPositionFieldMap, and every existing database picks it up on the next
' connect with no migration script and no lost history.
Private Sub AccAddMissingColumns(ByVal conn As Object, ByVal tableName As String)

    Dim rs As Object
    Dim have As Object
    Dim spec As Variant
    Dim i As Long
    Dim f As Long
    Dim colName As String

    spec = AccTableSpec(tableName)
    If UBound(spec) < LBound(spec) Then Exit Sub

    Set have = CreateObject("Scripting.Dictionary")
    have.CompareMode = 1

    On Error Resume Next
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM [" & tableName & "] WHERE 1=0", conn, _
            ACC_ADO_OPEN_STATIC, ACC_ADO_LOCK_READONLY

    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    For f = 0 To rs.fields.Count - 1
        have(CStr(rs.fields(f).name)) = True
    Next f

    rs.Close

    For i = LBound(spec) To UBound(spec)

        colName = AccSpecName(CStr(spec(i)))

        If Not have.Exists(colName) Then
            ' AUTOINCREMENT cannot be added to a populated table, and the only
            ' one is the run's own key, which exists from the first create.
            If UCase$(AccSpecType(CStr(spec(i)))) <> "AUTOINCREMENT" Then
                AccExecuteQuiet conn, _
                    "ALTER TABLE [" & tableName & "] ADD COLUMN [" & _
                    colName & "] " & AccSpecType(CStr(spec(i)))
            End If
        End If

    Next i

End Sub


' CREATE INDEX, ignoring "it is already there".  There is no portable way to
' ask ACE whether an index exists that is cheaper than trying.
Private Sub AccEnsureIndex(ByVal conn As Object, ByVal spec As String)

    Dim parts As Variant

    parts = Split(spec, "|")
    If UBound(parts) < 2 Then Exit Sub

    AccExecuteQuiet conn, _
        "CREATE INDEX [" & parts(0) & "] ON [" & parts(1) & "] (" & _
        parts(2) & ")"

End Sub


' Execute a statement, swallowing the error.  Used only where failing is a
' legitimate outcome - the object is already there - and never on the data
' path, where an error must surface.
Private Function AccExecuteQuiet( _
    ByVal conn As Object, _
    ByVal sql As String) As Boolean

    If Len(sql) = 0 Then Exit Function

    On Error Resume Next
    conn.Execute sql
    AccExecuteQuiet = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

End Function


Private Sub AccStampSchemaVersion(ByVal conn As Object)

    Dim rs As Object
    Dim already As Boolean

    On Error Resume Next
    Set rs = conn.Execute( _
        "SELECT COUNT(*) AS n FROM [Meta_Schema] WHERE SchemaVersion = " & _
        CStr(ACC_SCHEMA_VERSION))

    If Err.Number = 0 Then
        If Not rs Is Nothing Then
            already = (CDblSafeAcc(rs.fields("n").value) > 0)
            rs.Close
        End If
    End If
    Err.Clear
    On Error GoTo 0

    If already Then Exit Sub

    AccExecuteQuiet conn, _
        "INSERT INTO [Meta_Schema] (SchemaVersion, AppliedUtc, AppliedBy) " & _
        "VALUES (" & CStr(ACC_SCHEMA_VERSION) & ", " & _
        AccSqlDate(AccUtcNow()) & ", " & AccSqlText(AccUserName(), 64) & ")"

End Sub


Private Function CDblSafeAcc(ByVal v As Variant) As Double

    On Error GoTo Zero

    If IsNull(v) Or IsEmpty(v) Then Exit Function
    CDblSafeAcc = CDbl(v)
    Exit Function

Zero:
    CDblSafeAcc = 0

End Function


' UTC, so runs from machines in different offsets sort against each other.
' Falls back to local time if WMI is unavailable, which is better than
' refusing to save.
Private Function AccUtcNow() As Date

    Dim wmi As Object
    Dim items As Object
    Dim item As Object
    Dim offsetMinutes As Long

    AccUtcNow = Now

    On Error GoTo Local_

    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set items = wmi.ExecQuery("SELECT * FROM Win32_TimeZone")

    For Each item In items
        offsetMinutes = CLng(item.Bias)
        Exit For
    Next item

    AccUtcNow = DateAdd("n", -offsetMinutes, Now)
    Exit Function

Local_:
    AccUtcNow = Now

End Function


Private Function AccUserName() As String

    On Error Resume Next
    AccUserName = Environ$("USERNAME")
    If Len(AccUserName) = 0 Then AccUserName = Application.UserName
    Err.Clear
    On Error GoTo 0

End Function


' =============================================================================
' SAVING A RUN
'
' Order matters, and it is the order of what can still be undone:
'
'   1  read all three sheets into memory and build every key
'   2  fingerprint them, and stop here if this exact set is already stored
'   3  open a transaction, insert the Run row, get its RunID
'   4  insert the positions, and the issues the insert refused
'   5  commit, and only then write the RunID back onto Config
'
' Nothing is written to the database until the duplicate check has passed, and
' nothing is visible to another user until the commit - so a failure halfway
' leaves no half-run behind, which is the whole reason the transaction is here
' rather than three separate autocommitted loops.
' =============================================================================

Public Function Access_SaveCurrentRun( _
    Optional ByVal silent As Boolean = False, _
    Optional ByVal notes As String = "", _
    Optional ByVal quiet As Boolean = False) As Long

    Dim conn As Object
    Dim bonds As Variant
    Dim swaps As Variant
    Dim futures As Variant
    Dim fingerprint As String
    Dim priorRun As Long
    Dim runId As Long
    Dim asOf As Variant
    Dim issues As Object
    Dim savedBonds As Long
    Dim savedSwaps As Long
    Dim savedFutures As Long
    Dim inTrans As Boolean
    Dim msg As String
    Dim errText As String

    bonds = PnlPositionsForStore(STORE_KIND_BOND)
    swaps = PnlPositionsForStore(STORE_KIND_SWAP)
    futures = PnlPositionsForStore(STORE_KIND_FUTURE)

    If AccArrayRowCount(bonds) = 0 And AccArrayRowCount(swaps) = 0 _
       And AccArrayRowCount(futures) = 0 Then
        AccSetStatus "Nothing to save - all three sheets are empty."
        If Not silent Then
            MsgBox "There are no positions on Bonds, Swaps or Futures, so " & _
                   "there is nothing to store.", vbExclamation, "Access store"
        End If
        Exit Function
    End If

    fingerprint = AccRunFingerprint(bonds, swaps, futures)
    asOf = PnlAsOfDates()

    Set conn = AccOpen(ensureSchema:=True, quiet:=quiet)
    If conn Is Nothing Then Exit Function

    On Error GoTo SaveFail

    priorRun = AccRunWithFingerprint(conn, asOf(0), fingerprint)

    If priorRun > 0 Then
        AccSetStatus "Unchanged since run " & CStr(priorRun) & _
                     " - not saved again."
        If Not silent Then
            MsgBox "These are exactly the positions already stored as run " & _
                   CStr(priorRun) & " for the same reporting date, so " & _
                   "nothing was saved." & vbCrLf & vbCrLf & _
                   "Load new positions first, or use Access_PurgeRun if " & _
                   "run " & CStr(priorRun) & " should be replaced.", _
                   vbInformation, "Access store"
        End If
        Access_SaveCurrentRun = priorRun
        GoTo SaveExit
    End If

    conn.BeginTrans
    inTrans = True

    runId = AccInsertRun(conn, asOf(0), asOf(1), fingerprint, notes)

    If runId <= 0 Then
        Err.Raise vbObjectError + 2100, , _
            "The database did not return a RunID for the new run."
    End If

    Set issues = AccNewIssueLog()

    savedBonds = AccSavePositions(conn, runId, "BOND", "Pos_Bond", bonds, issues)
    savedSwaps = AccSavePositions(conn, runId, "SWAP", "Pos_Swap", swaps, issues)
    savedFutures = AccSavePositions(conn, runId, "FUTURE", "Pos_Future", _
                                    futures, issues)

    AccWriteIssues conn, runId, issues

    AccFinishRun conn, runId, savedBonds, savedSwaps, savedFutures, _
                 IIf(issues.Count = 0, "OK", "OK_WITH_ISSUES")

    conn.CommitTrans
    inTrans = False

    AccSetLastRun runId
    msg = "Run " & CStr(runId) & ": " & CStr(savedBonds) & " bonds, " & _
          CStr(savedSwaps) & " swaps, " & CStr(savedFutures) & " futures" & _
          IIf(issues.Count = 0, "", ", " & CStr(issues.Count) & " issue(s)")
    AccSetStatus msg

    Access_SaveCurrentRun = runId

    If Not silent Then
        MsgBox msg & vbCrLf & vbCrLf & _
               "Database: " & AccDatabasePath() & _
               IIf(issues.Count = 0, "", vbCrLf & vbCrLf & _
                   "See the Run_Issue table for what was refused."), _
               IIf(issues.Count = 0, vbInformation, vbExclamation), _
               "Access store"
    End If

SaveExit:
    On Error Resume Next
    conn.Close
    On Error GoTo 0
    Exit Function

SaveFail:
    ' Read Err FIRST.  Everything below either calls a procedure that traps
    ' internally or rolls back under On Error Resume Next, and both reset Err -
    ' so by the time the message box is built there is nothing left to report
    ' and the failure reads as "Saving the run failed and nothing was written."
    ' with a blank second half.
    errText = Err.Description

    If inTrans Then
        On Error Resume Next
        conn.RollbackTrans
        On Error GoTo 0
    End If

    AccSetStatus "Save FAILED: " & errText

    If Not quiet Then
        MsgBox "Saving the run failed and nothing was written." & vbCrLf & _
               vbCrLf & errText, vbCritical, "Access store"
    End If

    Resume SaveExit

End Function


' One string standing for everything the run holds.  The counts are in it as
' well as the keys, so a run that lost a position cannot fingerprint the same
' as one that never had it.
Private Function AccRunFingerprint( _
    ByVal bonds As Variant, _
    ByVal swaps As Variant, _
    ByVal futures As Variant) As String

    AccRunFingerprint = _
        AccFingerprintOfKeys(AccKeysOf("BOND", bonds)) & "-" & _
        AccFingerprintOfKeys(AccKeysOf("SWAP", swaps)) & "-" & _
        AccFingerprintOfKeys(AccKeysOf("FUTURE", futures)) & "-" & _
        AccHash32(CStr(AccArrayRowCount(bonds)) & "/" & _
                  CStr(AccArrayRowCount(swaps)) & "/" & _
                  CStr(AccArrayRowCount(futures)))

End Function


' Every position key of one kind, in sheet order.
Public Function AccKeysOf(ByVal kind As String, ByVal data As Variant) As Variant

    Dim cols As Object
    Dim keys() As String
    Dim n As Long
    Dim i As Long

    n = AccArrayRowCount(data)

    If n = 0 Then
        AccKeysOf = Array()
        Exit Function
    End If

    Set cols = AccColumnIndex(data)
    ReDim keys(1 To n)

    For i = 1 To n
        keys(i) = AccKeyForRow(kind, data, LBound(data, 1) + i, cols)
    Next i

    AccKeysOf = keys

End Function


Public Function AccArrayRowCount(ByVal data As Variant) As Long

    If IsEmpty(data) Then Exit Function
    If Not IsArray(data) Then Exit Function

    On Error GoTo NoRows
    AccArrayRowCount = UBound(data, 1) - LBound(data, 1)
    On Error GoTo 0

    If AccArrayRowCount < 0 Then AccArrayRowCount = 0
    Exit Function

NoRows:
    AccArrayRowCount = 0

End Function


Private Function AccRunWithFingerprint( _
    ByVal conn As Object, _
    ByVal asOfT0 As Variant, _
    ByVal fingerprint As String) As Long

    Dim rs As Object
    Dim sql As String

    sql = "SELECT TOP 1 RunID FROM [Run] WHERE RunFingerprint = " & _
          AccSqlText(fingerprint, 64) & " AND AsOfT0 " & _
          IIf(AccSqlDate(asOfT0) = "NULL", "IS NULL", _
              "= " & AccSqlDate(asOfT0)) & _
          " AND RunStatus <> 'ABANDONED' ORDER BY RunID DESC"

    On Error GoTo NoMatch

    Set rs = conn.Execute(sql)

    If Not rs.EOF Then AccRunWithFingerprint = CLng(rs.fields(0).value)
    rs.Close

    Exit Function

NoMatch:
    AccRunWithFingerprint = 0

End Function


' Insert the Run row and read back the identity it was given.
'
' @@IDENTITY on the SAME connection, which is what makes it safe: it is
' per-connection, so another user saving at the same moment cannot hand this
' one their RunID.
Private Function AccInsertRun( _
    ByVal conn As Object, _
    ByVal asOfT0 As Variant, _
    ByVal asOfTM1 As Variant, _
    ByVal fingerprint As String, _
    ByVal notes As String) As Long

    Dim rs As Object

    conn.Execute _
        "INSERT INTO [Run] (RunStartedUtc, AsOfT0, AsOfTM1, RunBy, " & _
        "ModuleVersion, ExcelBitness, RunStatus, RunFingerprint, Notes) " & _
        "VALUES (" & _
        AccSqlDate(AccUtcNow()) & ", " & _
        AccSqlDate(asOfT0) & ", " & _
        AccSqlDate(asOfTM1) & ", " & _
        AccSqlText(AccUserName(), 64) & ", " & _
        AccSqlText("schema " & CStr(ACC_SCHEMA_VERSION), 32) & ", " & _
        AccSqlText(AccBitness(), 8) & ", " & _
        "'RUNNING', " & _
        AccSqlText(fingerprint, 64) & ", " & _
        AccSqlText(notes, 255) & ")"

    Set rs = conn.Execute("SELECT @@IDENTITY")

    If Not rs.EOF Then AccInsertRun = CLng(rs.fields(0).value)
    rs.Close

End Function


Private Sub AccFinishRun( _
    ByVal conn As Object, _
    ByVal runId As Long, _
    ByVal bondCount As Long, _
    ByVal swapCount As Long, _
    ByVal futureCount As Long, _
    ByVal runStatus As String)

    conn.Execute _
        "UPDATE [Run] SET RunFinishedUtc = " & AccSqlDate(AccUtcNow()) & _
        ", BondCount = " & CStr(bondCount) & _
        ", SwapCount = " & CStr(swapCount) & _
        ", FutureCount = " & CStr(futureCount) & _
        ", RunStatus = " & AccSqlText(runStatus, 16) & _
        " WHERE RunID = " & CStr(runId)

    ' The run database is now complete, so the index can point at it.  Failing to
    ' index is NOT allowed to fail the save: the run itself is stored and correct,
    ' and an index that has fallen behind is repairable by rescanning the folder.
    On Error Resume Next
    AccIndexThisRun conn, runId, bondCount, swapCount, futureCount, runStatus
    On Error GoTo 0

End Sub


' Record this run in the index database.
'
' One row per retrieval stamp.  A re-save of the same pull UPDATEs its row rather
' than adding a second, which is the same rule the run databases follow - the
' stamp is the retrieval moment, so two saves of one pull are one run.
Private Sub AccIndexThisRun( _
    ByVal conn As Object, _
    ByVal runId As Long, _
    ByVal bondCount As Long, _
    ByVal swapCount As Long, _
    ByVal futureCount As Long, _
    ByVal runStatus As String)

    Dim idx As Object
    Dim rs As Object
    Dim stamp As String
    Dim sql As String
    Dim asOfT0 As Variant
    Dim asOfT1 As Variant
    Dim fingerprint As String

    stamp = AccRunStamp()
    If Len(stamp) = 0 Then Exit Sub

    ' Read back what the run recorded about itself, rather than re-deriving it -
    ' the index must agree with the database it points at.
    Set rs = conn.Execute( _
        "SELECT AsOfT0, AsOfT1, RunFingerprint FROM [Run] WHERE RunID = " & CStr(runId))
    If Not rs.EOF Then
        asOfT0 = rs.Fields(0).value
        asOfT1 = rs.Fields(1).value
        fingerprint = AccClean(rs.Fields(2).value)
    End If
    rs.Close

    Set idx = AccOpenIndex()
    If idx Is Nothing Then Exit Sub

    idx.Execute "DELETE FROM Run_Index WHERE RunStamp = " & AccSqlText(stamp, 15)

    sql = "INSERT INTO Run_Index (RunStamp, DbPath, AsOfT0, AsOfT1, " & _
          "RunFinishedUtc, RunStatus, RunFingerprint, BondCount, SwapCount, " & _
          "FutureCount, MachineName, UserName) VALUES (" & _
          AccSqlText(stamp, 15) & ", " & _
          AccSqlText(AccDatabasePath(), 255) & ", " & _
          AccSqlDate(asOfT0) & ", " & _
          AccSqlDate(asOfT1) & ", " & _
          AccSqlDate(AccUtcNow()) & ", " & _
          AccSqlText(runStatus, 16) & ", " & _
          AccSqlText(fingerprint, 32) & ", " & _
          CStr(bondCount) & ", " & CStr(swapCount) & ", " & CStr(futureCount) & ", " & _
          AccSqlText(Environ$("COMPUTERNAME"), 64) & ", " & _
          AccSqlText(AccUserName(), 64) & ")"

    idx.Execute sql
    idx.Close

End Sub


' Open the index database, creating the file and its one table if absent.
'
' Separate from AccOpen because the index is not a run: it has its own schema,
' it outlives every run database, and it must never be created by AccEnsureSchema
' with the run tables in it.
Private Function AccOpenIndex() As Object

    Dim conn As Object
    Dim providers As Variant
    Dim dbPath As String
    Dim i As Long

    dbPath = AccIndexDatabasePath()

    If Len(Dir$(dbPath)) = 0 Then
        If Not AccCreateDatabaseFile(dbPath, quiet:=True) Then Exit Function
    End If

    providers = AccProviders()

    For i = LBound(providers) To UBound(providers)
        On Error Resume Next
        Set conn = CreateObject("ADODB.Connection")
        conn.Open AccConnString(CStr(providers(i)), dbPath)
        If Err.Number <> 0 Then
            Err.Clear
            Set conn = Nothing
        End If
        On Error GoTo 0
        If Not conn Is Nothing Then Exit For
    Next i

    If conn Is Nothing Then Exit Function

    AccEnsureIndexSchema conn

    Set AccOpenIndex = conn

End Function


' Create Run_Index if it is not there.  Mirrors access/index_schema.sql; keep the
' two in step, the same way AccTableSpec mirrors access/schema.sql.
Private Sub AccEnsureIndexSchema(ByVal conn As Object)

    AccExecuteQuiet conn, _
        "CREATE TABLE Run_Index (" & _
        "RunIndexID AUTOINCREMENT NOT NULL, " & _
        "RunStamp TEXT(15) NOT NULL, " & _
        "DbPath TEXT(255) NOT NULL, " & _
        "AsOfT0 DATETIME, " & _
        "AsOfT1 DATETIME, " & _
        "RunStartedUtc DATETIME, " & _
        "RunFinishedUtc DATETIME, " & _
        "RunStatus TEXT(16), " & _
        "RunFingerprint TEXT(32), " & _
        "BondCount LONG, " & _
        "SwapCount LONG, " & _
        "FutureCount LONG, " & _
        "MachineName TEXT(64), " & _
        "UserName TEXT(64), " & _
        "Notes LONGTEXT, " & _
        "CONSTRAINT PK_Run_Index PRIMARY KEY (RunIndexID))"

    AccExecuteQuiet conn, _
        "CREATE UNIQUE INDEX IX_Run_Index_Stamp ON Run_Index (RunStamp)"

    AccExecuteQuiet conn, _
        "CREATE INDEX IX_Run_Index_AsOf ON Run_Index (AsOfT0)"

End Sub


' =============================================================================
' WRITING THE POSITIONS
'
' Through a recordset rather than INSERT statements, for three reasons: no
' value is ever turned into text and back (so no locale and no escaping
' question), the field types come from the table rather than from a guess, and
' ACE appends far faster this way than it parses a thousand statements.
'
' Two duplicate guards, on purpose.  The dictionary catches a repeat before it
' reaches the database, so the issue log can record WHICH row was refused and
' where it sat.  The primary key catches anything the dictionary's idea of
' sameness missed.  The second is the one that cannot be bypassed.
' =============================================================================

Private Function AccSavePositions( _
    ByVal conn As Object, _
    ByVal runId As Long, _
    ByVal kind As String, _
    ByVal tableName As String, _
    ByVal data As Variant, _
    ByVal issues As Object) As Long

    Dim rs As Object
    Dim cols As Object
    Dim spec As Variant
    Dim seen As Object
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim key As String
    Dim colName As String
    Dim srcRow As Variant
    Dim written As Long

    n = AccArrayRowCount(data)
    If n = 0 Then Exit Function

    Set cols = AccColumnIndex(data)
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 0          ' keys are exact strings, case included

    spec = AccTableSpec(tableName)

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM [" & tableName & "] WHERE 1=0", conn, _
            ACC_ADO_OPEN_KEYSET, ACC_ADO_LOCK_OPTIMISTIC

    For i = 1 To n

        key = AccKeyForRow(kind, data, LBound(data, 1) + i, cols)
        srcRow = AccCell(data, LBound(data, 1) + i, cols, "SourceRow")

        If Len(key) = 0 Then

            AccAddIssue issues, "WARN", tableName, key, srcRow, _
                "row has no identifying fields - not stored"

        ElseIf Len(key) > ACC_MAX_KEY_LEN Then

            ' Truncating would make two different positions share a key, so
            ' the row is refused instead.  Never seen in practice; if it ever
            ' is, PositionKey's TEXT(200) is what needs raising.
            AccAddIssue issues, "ERROR", tableName, key, srcRow, _
                "position key is " & CStr(Len(key)) & " characters, over the " & _
                CStr(ACC_MAX_KEY_LEN) & " the column holds"

        ElseIf seen.Exists(key) Then

            AccAddIssue issues, "WARN", tableName, key, srcRow, _
                "duplicate of the row already stored from sheet row " & _
                CStr(seen(key))

        Else

            seen(key) = srcRow

            rs.AddNew
            rs.fields("RunID").value = runId
            rs.fields("PositionKey").value = key

            For j = LBound(spec) To UBound(spec)

                colName = AccSpecName(CStr(spec(j)))

                If colName <> "RunID" And colName <> "PositionKey" Then
                    If cols.Exists(colName) Then
                        rs.fields(colName).value = AccCoerce( _
                            data(LBound(data, 1) + i, cols(colName)), _
                            AccSpecType(CStr(spec(j))))
                    End If
                End If

            Next j

            rs.Update
            written = written + 1

        End If

    Next i

    rs.Close

    AccSavePositions = written

End Function


' A sheet cell into something the column will accept.
'
' Blank is Null, never zero and never "": a bond with no book value has no book
' value, and storing 0 would make it look like a position worth nothing.  A
' cell holding an Excel error - #N/A from a lookup that found nothing - is Null
' for the same reason.
Private Function AccCoerce( _
    ByVal v As Variant, _
    ByVal colType As String) As Variant

    Dim s As String
    Dim upperType As String

    AccCoerce = Null

    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function

    upperType = UCase$(colType)

    If upperType = "DOUBLE" Or upperType = "LONG" Then
        If Not IsNumeric(v) Then Exit Function
        AccCoerce = CDbl(v)
        Exit Function
    End If

    If upperType = "DATETIME" Then
        If Not IsDate(v) Then Exit Function
        AccCoerce = CDate(v)
        Exit Function
    End If

    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function

    ' Text longer than its column would be truncated by the provider without
    ' comment; cutting it here is the same outcome said out loud.
    AccCoerce = Left$(s, AccTextLen(colType))

End Function


Private Function AccTextLen(ByVal colType As String) As Long

    Dim inner As String
    Dim p1 As Long
    Dim p2 As Long

    AccTextLen = 255

    p1 = InStr(colType, "(")
    p2 = InStr(colType, ")")

    If p1 > 0 And p2 > p1 Then
        inner = Mid$(colType, p1 + 1, p2 - p1 - 1)
        If IsNumeric(inner) Then AccTextLen = CLng(inner)
    End If

    If UCase$(colType) = "LONGTEXT" Then AccTextLen = 65535

End Function


' =============================================================================
' THE ISSUE LOG
'
' Held in memory during the save and written inside the same transaction, so
' "the run stored cleanly" is answerable from the database months later rather
' than from whether anyone read a message box.
' =============================================================================

Private Function AccNewIssueLog() As Object
    Set AccNewIssueLog = CreateObject("Scripting.Dictionary")
End Function


Private Sub AccAddIssue( _
    ByVal issues As Object, _
    ByVal severity As String, _
    ByVal tableName As String, _
    ByVal key As String, _
    ByVal srcRow As Variant, _
    ByVal detail As String)

    If issues.Count >= ACC_MAX_ISSUES Then Exit Sub

    issues(issues.Count) = Array(severity, tableName, key, srcRow, detail)

End Sub


Private Sub AccWriteIssues( _
    ByVal conn As Object, _
    ByVal runId As Long, _
    ByVal issues As Object)

    Dim k As Variant
    Dim it As Variant
    Dim seq As Long

    For Each k In issues.keys

        it = issues(k)
        seq = seq + 1

        conn.Execute _
            "INSERT INTO [Run_Issue] (RunID, IssueSeq, Severity, TableName, " & _
            "PositionKey, SourceRow, Detail) VALUES (" & _
            CStr(runId) & ", " & CStr(seq) & ", " & _
            AccSqlText(it(0), 8) & ", " & _
            AccSqlText(it(1), 32) & ", " & _
            AccSqlText(it(2), 200) & ", " & _
            AccSqlNum(it(3)) & ", " & _
            AccSqlText(it(4), 255) & ")"

    Next k

End Sub


' =============================================================================
' READING A RUN BACK
'
' The database returns a recordset; the sheets want the same 2-D array shape
' PnlPositionsForStore produces, header row and all.  Building that shape here
' is what lets one restore function serve all three sheets and lets a run
' stored before a field existed still load - the missing column is simply
' absent from the header, and modPNL leaves that column blank.
' =============================================================================

Public Sub Access_LoadRunIntoSheets(Optional ByVal runIdIn As Long = 0)

    Dim conn As Object
    Dim runId As Long
    Dim answer As String
    Dim bondRows As Long
    Dim swapRows As Long
    Dim futRows As Long
    Dim problems As String
    Dim oldCalc As XlCalculation
    Dim oldScreen As Boolean

    runId = runIdIn

    If runId <= 0 Then
        answer = InputBox( _
            "Which run should be loaded onto Bonds, Swaps and Futures?" & _
            vbCrLf & vbCrLf & _
            "Run Access_ShowRunHistory first if you need the list." & _
            vbCrLf & vbCrLf & _
            "This REPLACES what is on those three sheets.", _
            "Load a stored run", _
            CStr(AccLastRun()))

        If Len(Trim$(answer)) = 0 Then Exit Sub
        If Not IsNumeric(answer) Then Exit Sub
        runId = CLng(answer)
    End If

    Set conn = AccOpen()
    If conn Is Nothing Then Exit Sub

    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating

    On Error GoTo LoadFail

    If Not AccRunExists(conn, runId) Then
        MsgBox "There is no run " & CStr(runId) & " in " & _
               AccDatabasePath() & ".", vbExclamation, "Access store"
        GoTo LoadExit
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Application.StatusBar = "Loading run " & CStr(runId) & "..."

    bondRows = PnlRestorePositions(STORE_KIND_BOND, _
        AccReadPositions(conn, runId, "Pos_Bond"))
    swapRows = PnlRestorePositions(STORE_KIND_SWAP, _
        AccReadPositions(conn, runId, "Pos_Swap"))
    futRows = PnlRestorePositions(STORE_KIND_FUTURE, _
        AccReadPositions(conn, runId, "Pos_Future"))

    Application.Calculation = oldCalc

    Application.StatusBar = "Rewriting formulas for the restored run..."
    problems = PnlRewriteFormulasForRestoredRun()

    Application.StatusBar = False

    AccSetStatus "Loaded run " & CStr(runId) & ": " & CStr(bondRows) & _
                 " bonds, " & CStr(swapRows) & " swaps, " & _
                 CStr(futRows) & " futures."

    MsgBox "Run " & CStr(runId) & " is on the sheets:" & vbCrLf & vbCrLf & _
           CStr(bondRows) & " bonds, " & CStr(swapRows) & " swaps, " & _
           CStr(futRows) & " futures." & vbCrLf & vbCrLf & _
           "The formulas have been rewritten and the book recalculated. " & _
           "Market data is still live, so the prices are today's - what has " & _
           "been restored is the POSITIONS." & _
           IIf(Len(problems) = 0, "", vbCrLf & vbCrLf & "Problems:" & _
               vbCrLf & problems), _
           IIf(Len(problems) = 0, vbInformation, vbExclamation), _
           "Access store"

LoadExit:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Application.StatusBar = False
    On Error Resume Next
    conn.Close
    On Error GoTo 0
    Exit Sub

LoadFail:
    MsgBox "Loading run " & CStr(runId) & " failed:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Access store"
    Resume LoadExit

End Sub


Private Function AccRunExists( _
    ByVal conn As Object, _
    ByVal runId As Long) As Boolean

    Dim rs As Object

    Set rs = conn.Execute( _
        "SELECT COUNT(*) AS n FROM [Run] WHERE RunID = " & CStr(runId))

    If Not rs.EOF Then AccRunExists = (CDblSafeAcc(rs.fields("n").value) > 0)
    rs.Close

End Function


' One table's rows for one run, in the header-plus-data shape modPNL restores
' from.  Ordered by SourceRow so the sheet comes back in the order it went in.
Private Function AccReadPositions( _
    ByVal conn As Object, _
    ByVal runId As Long, _
    ByVal tableName As String) As Variant

    Dim rs As Object
    Dim block As Variant
    Dim outv() As Variant
    Dim fieldCount As Long
    Dim rowCount As Long
    Dim i As Long
    Dim j As Long

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM [" & tableName & "] WHERE RunID = " & CStr(runId) & _
            " ORDER BY SourceRow", conn, ACC_ADO_OPEN_STATIC, _
            ACC_ADO_LOCK_READONLY

    fieldCount = rs.fields.Count

    If rs.EOF Then
        rowCount = 0
    Else
        ' GetRows returns fields-by-rows, which is the transpose of what the
        ' sheet wants - hence the swapped indices below.  One call beats
        ' walking the recordset a field at a time.
        block = rs.GetRows()
        rowCount = UBound(block, 2) - LBound(block, 2) + 1
    End If

    ReDim outv(0 To rowCount, 0 To fieldCount - 1)

    For j = 0 To fieldCount - 1
        outv(0, j) = CStr(rs.fields(j).name)
    Next j

    For i = 1 To rowCount
        For j = 0 To fieldCount - 1
            outv(i, j) = block(j, i - 1)
        Next j
    Next i

    rs.Close

    AccReadPositions = outv

End Function


' =============================================================================
' THE HISTORY SHEET
'
' A plain list of runs on a sheet of its own, rebuilt on demand.  It exists so
' that "which run do I want" is answerable without opening Access, and it is
' never read by anything - no formula depends on it, so rebuilding it cannot
' break the book.
' =============================================================================

Public Sub Access_ShowRunHistory()

    Dim conn As Object
    Dim ws As Worksheet
    Dim rs As Object
    Dim headers As Variant
    Dim n As Long

    Set conn = AccOpen()
    If conn Is Nothing Then Exit Sub

    On Error GoTo HistFail

    Set ws = AccHistorySheet()

    ws.Cells.ClearContents

    ws.Range("A1").value = "Stored runs - " & AccDatabasePath()
    ws.Range("A1").Font.Bold = True

    headers = Array("RunID", "AsOfT0", "AsOfTM1", "Started (UTC)", _
                    "Finished (UTC)", "By", "Status", "Bonds", "Swaps", _
                    "Futures", "Issues", "Fingerprint", "Notes")

    ws.Range("A3").Resize(1, UBound(headers) + 1).value = headers
    ws.Range("A3").Resize(1, UBound(headers) + 1).Font.Bold = True

    Set rs = conn.Execute( _
        "SELECT r.RunID, r.AsOfT0, r.AsOfTM1, r.RunStartedUtc, " & _
        "r.RunFinishedUtc, r.RunBy, r.RunStatus, r.BondCount, r.SwapCount, " & _
        "r.FutureCount, " & _
        "(SELECT COUNT(*) FROM [Run_Issue] i WHERE i.RunID = r.RunID) " & _
        "AS IssueCount, r.RunFingerprint, r.Notes " & _
        "FROM [Run] AS r ORDER BY r.RunID DESC")

    If Not rs.EOF Then
        ws.Range("A4").CopyFromRecordset rs
        n = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 3
    End If

    rs.Close

    ws.Columns("A:M").AutoFit
    ws.Activate

    AccSetStatus CStr(n) & " run(s) listed."

HistExit:
    On Error Resume Next
    conn.Close
    On Error GoTo 0
    Exit Sub

HistFail:
    MsgBox "Could not list the stored runs:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Access store"
    Resume HistExit

End Sub


Private Function AccHistorySheet() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ACC_SH_HISTORY)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = ACC_SH_HISTORY
    End If

    Set AccHistorySheet = ws

End Function


' =============================================================================
' DELETING RUNS
'
' The only destructive operation here, and it is deliberately awkward: it names
' what it is about to remove and requires a Yes.  Positions first, then the run
' row, all inside one transaction, so a half-deleted run cannot exist.
' =============================================================================

Public Sub Access_PurgeRun(Optional ByVal runIdIn As Long = 0)

    Dim conn As Object
    Dim runId As Long
    Dim answer As String
    Dim rs As Object
    Dim summary As String
    Dim inTrans As Boolean

    runId = runIdIn

    If runId <= 0 Then
        answer = InputBox("Which run should be deleted?", "Delete a run", "")
        If Len(Trim$(answer)) = 0 Then Exit Sub
        If Not IsNumeric(answer) Then Exit Sub
        runId = CLng(answer)
    End If

    Set conn = AccOpen()
    If conn Is Nothing Then Exit Sub

    On Error GoTo PurgeFail

    Set rs = conn.Execute( _
        "SELECT AsOfT0, BondCount, SwapCount, FutureCount FROM [Run] " & _
        "WHERE RunID = " & CStr(runId))

    If rs.EOF Then
        rs.Close
        MsgBox "There is no run " & CStr(runId) & ".", vbExclamation, _
               "Access store"
        GoTo PurgeExit
    End If

    summary = "Run " & CStr(runId) & ", as of " & _
              CStr(rs.fields("AsOfT0").value) & ": " & _
              CStr(rs.fields("BondCount").value) & " bonds, " & _
              CStr(rs.fields("SwapCount").value) & " swaps, " & _
              CStr(rs.fields("FutureCount").value) & " futures."
    rs.Close

    If MsgBox(summary & vbCrLf & vbCrLf & _
              "Delete it permanently?", vbYesNo + vbExclamation, _
              "Delete a run") <> vbYes Then GoTo PurgeExit

    conn.BeginTrans
    inTrans = True

    AccDeleteRunRows conn, runId

    conn.CommitTrans
    inTrans = False

    AccSetStatus "Deleted run " & CStr(runId) & "."
    MsgBox "Run " & CStr(runId) & " deleted.", vbInformation, "Access store"

PurgeExit:
    On Error Resume Next
    conn.Close
    On Error GoTo 0
    Exit Sub

PurgeFail:
    If inTrans Then
        On Error Resume Next
        conn.RollbackTrans
        On Error GoTo 0
    End If
    MsgBox "Deleting run " & CStr(runId) & " failed:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Access store"
    Resume PurgeExit

End Sub


Private Sub AccDeleteRunRows(ByVal conn As Object, ByVal runId As Long)

    conn.Execute "DELETE FROM [Pos_Bond] WHERE RunID = " & CStr(runId)
    conn.Execute "DELETE FROM [Pos_Swap] WHERE RunID = " & CStr(runId)
    conn.Execute "DELETE FROM [Pos_Future] WHERE RunID = " & CStr(runId)
    conn.Execute "DELETE FROM [Run_Issue] WHERE RunID = " & CStr(runId)
    conn.Execute "DELETE FROM [Run] WHERE RunID = " & CStr(runId)

End Sub


' Keep the most recent N runs and delete the rest.  Config!B45 holds N; 0 - the
' default - keeps everything, because a store that quietly throws history away
' is worse than one that grows.
Public Sub Access_PurgeOldRuns()

    Dim conn As Object
    Dim keep As Long
    Dim rs As Object
    Dim victims As Collection
    Dim v As Variant
    Dim inTrans As Boolean

    keep = AccKeepRuns()

    If keep <= 0 Then
        MsgBox "Config!" & ACC_CFG_KEEP & " is blank or zero, which means " & _
               "keep every run. Set it to the number of runs to keep and " & _
               "run this again.", vbInformation, "Access store"
        Exit Sub
    End If

    Set conn = AccOpen()
    If conn Is Nothing Then Exit Sub

    On Error GoTo TrimFail

    Set victims = New Collection

    Set rs = conn.Execute( _
        "SELECT RunID FROM [Run] WHERE RunID NOT IN " & _
        "(SELECT TOP " & CStr(keep) & " RunID FROM [Run] ORDER BY RunID DESC)")

    Do While Not rs.EOF
        victims.Add CLng(rs.fields(0).value)
        rs.MoveNext
    Loop
    rs.Close

    If victims.Count = 0 Then
        MsgBox "There are " & CStr(keep) & " runs or fewer, so nothing was " & _
               "deleted.", vbInformation, "Access store"
        GoTo TrimExit
    End If

    If MsgBox(CStr(victims.Count) & " run(s) older than the most recent " & _
              CStr(keep) & " will be deleted permanently." & vbCrLf & _
              vbCrLf & "Continue?", vbYesNo + vbExclamation, _
              "Trim the history") <> vbYes Then GoTo TrimExit

    conn.BeginTrans
    inTrans = True

    For Each v In victims
        AccDeleteRunRows conn, CLng(v)
    Next v

    conn.CommitTrans
    inTrans = False

    AccSetStatus "Deleted " & CStr(victims.Count) & " old run(s)."
    MsgBox CStr(victims.Count) & " run(s) deleted.", vbInformation, _
           "Access store"

TrimExit:
    On Error Resume Next
    conn.Close
    On Error GoTo 0
    Exit Sub

TrimFail:
    If inTrans Then
        On Error Resume Next
        conn.RollbackTrans
        On Error GoTo 0
    End If
    MsgBox "Trimming the history failed:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Access store"
    Resume TrimExit

End Sub


' =============================================================================
' THE TEST BUTTON
'
' Proves the whole path rather than just the connection: opens (creating the
' file if it is not there), applies any missing tables and columns, and reports
' what is actually in the database.  A green answer here means a save will
' work, which is the only useful thing a connection test can say.
' =============================================================================

Public Sub Access_TestConnection()

    Dim conn As Object
    Dim tableNames As Variant
    Dim existing As Object
    Dim report As String
    Dim runCount As Long
    Dim rs As Object
    Dim i As Long
    Dim created As Boolean

    created = (Len(Dir$(AccDatabasePath())) = 0)

    Set conn = AccOpen()
    If conn Is Nothing Then Exit Sub

    On Error GoTo TestFail

    Set existing = AccExistingTables(conn)
    tableNames = AccTableNames()

    For i = LBound(tableNames) To UBound(tableNames)
        report = report & vbCrLf & "  " & _
            IIf(existing.Exists(CStr(tableNames(i))), "OK   ", "MISSING ") & _
            CStr(tableNames(i))
    Next i

    Set rs = conn.Execute("SELECT COUNT(*) AS n FROM [Run]")
    If Not rs.EOF Then runCount = CLng(CDblSafeAcc(rs.fields("n").value))
    rs.Close

    MsgBox "Access store is reachable." & vbCrLf & vbCrLf & _
           "Database: " & AccDatabasePath() & vbCrLf & _
           "Excel:    " & AccBitness() & vbCrLf & _
           "Schema:   version " & CStr(ACC_SCHEMA_VERSION) & vbCrLf & _
           "Runs:     " & CStr(runCount) & _
           IIf(created, vbCrLf & vbCrLf & _
               "The database did not exist and has just been created.", "") & _
           vbCrLf & vbCrLf & "Tables:" & report, _
           vbInformation, "Access store"

TestExit:
    On Error Resume Next
    conn.Close
    On Error GoTo 0
    Exit Sub

TestFail:
    MsgBox "The database opened but the check failed:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Access store"
    Resume TestExit

End Sub


' =============================================================================
' THE AUTOSAVE HOOK
'
' Called at the end of Button 1 and Button 2.  Silent in BOTH directions - no
' box on success and none on failure - because a database that is unreachable
' must not interrupt a load that has already succeeded, and a desk that has not
' installed the engine yet would otherwise get a modal provider error twice a
' day about something the buttons do not depend on.
'
' Config!B44 carries what happened either way, and Access_TestConnection is the
' button that answers loudly on purpose.
'
' Button 1 saves too, not only Button 2, because a bond load that is never
' followed by a hedge load is still a state somebody may need to look at
' again - and the duplicate guard means the Button 2 save that follows minutes
' later supersedes it rather than doubling it.
' =============================================================================

Public Sub Access_AutoSaveAfterLoad(ByVal stepName As String)

    If Not AccAutoSaveEnabled() Then Exit Sub

    On Error GoTo Silent_

    ' The return value is deliberately discarded.  Every branch of the save has
    ' already written an accurate Config!B44 - "Run 12: 380 bonds..." when it
    ' stored one, "Unchanged since run 11 - not saved again" when the
    ' fingerprint matched - and reporting the RunID again here would overwrite
    ' the second of those with "Autosaved as run 11", which is not what
    ' happened.
    Access_SaveCurrentRun _
        silent:=True, notes:="autosave after " & stepName, quiet:=True

    Exit Sub

Silent_:
    AccSetStatus "Autosave after " & stepName & " failed: " & Err.Description

End Sub


' =============================================================================
' CONFIG
'
' Read defensively throughout: a workbook whose Config sheet has not been set
' up yet must still open, and a missing cell means "the default", not a crash.
' =============================================================================

Private Function AccConfigSheet() As Worksheet

    On Error Resume Next
    Set AccConfigSheet = ThisWorkbook.Worksheets(ACC_SH_CONFIG)
    Err.Clear
    On Error GoTo 0

End Function


Public Function AccAutoSaveEnabled() As Boolean

    Dim ws As Worksheet
    Dim v As String

    Set ws = AccConfigSheet()
    If ws Is Nothing Then Exit Function

    v = UCase$(Trim$(CStr(ws.Range(ACC_CFG_AUTOSAVE).value)))

    AccAutoSaveEnabled = (v = "TRUE" Or v = "YES" Or v = "Y" Or v = "1")

End Function


Private Function AccKeepRuns() As Long

    Dim ws As Worksheet
    Dim v As String

    Set ws = AccConfigSheet()
    If ws Is Nothing Then Exit Function

    v = Trim$(CStr(ws.Range(ACC_CFG_KEEP).value))
    If IsNumeric(v) Then AccKeepRuns = CLng(val(v))

End Function


Private Function AccLastRun() As Long

    Dim ws As Worksheet
    Dim v As String

    Set ws = AccConfigSheet()
    If ws Is Nothing Then Exit Function

    v = Trim$(CStr(ws.Range(ACC_CFG_LAST_RUN).value))
    If IsNumeric(v) Then AccLastRun = CLng(val(v))

End Function


Private Sub AccSetLastRun(ByVal runId As Long)

    Dim ws As Worksheet

    Set ws = AccConfigSheet()
    If ws Is Nothing Then Exit Sub

    ws.Range(ACC_CFG_LAST_RUN).value = runId

End Sub


Private Sub AccSetStatus(ByVal msg As String)

    Dim ws As Worksheet

    Set ws = AccConfigSheet()
    If ws Is Nothing Then Exit Sub

    ws.Range(ACC_CFG_STATUS).value = _
        Format$(Now, "yyyy-mm-dd hh:nn") & "  " & msg

End Sub


' Label the cells this module owns.  Called by modPNL's layout setup so a fresh
' workbook shows them named rather than as five anonymous entries under the
' last thing anyone labelled.
Public Sub Access_SetupConfigCells()

    Dim ws As Worksheet

    Set ws = AccConfigSheet()
    If ws Is Nothing Then Exit Sub

    ws.Range("A41").value = "Run database FOLDER (blank = beside this workbook)"
    ws.Range("A42").value = "Autosave every run to Access (TRUE/FALSE)"
    ws.Range("A43").value = "Last saved RunID (auto - do not edit)"
    ws.Range("A44").value = "Last store message (auto - do not edit)"
    ws.Range("A45").value = "Runs to keep when trimming (0 = keep all)"
    ws.Range("A46").value = "Current run stamp (auto - do not edit)"

    If Len(Trim$(CStr(ws.Range(ACC_CFG_AUTOSAVE).value))) = 0 Then
        ws.Range(ACC_CFG_AUTOSAVE).value = "TRUE"
    End If

    If Len(Trim$(CStr(ws.Range(ACC_CFG_KEEP).value))) = 0 Then
        ws.Range(ACC_CFG_KEEP).value = 0
    End If

End Sub
