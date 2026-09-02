"""Identifiers VBA / the Excel object library provide for free.

Only names that can appear *bare* in code need to be here; anything reached
through a member access (`Application.Calculation`) is skipped by the scanner.
"""

KEYWORDS = {
    "addressof", "alias", "and", "any", "as", "attribute", "byref", "byval",
    "call", "case", "cbool", "cbyte", "ccur", "cdate", "cdbl", "cdec", "cint",
    "class", "clng", "clngptr", "clnglng", "compare", "const", "csng", "cstr",
    "curdir", "cvar", "cverr", "date", "debug", "decimal", "declare",
    "defbool", "defbyte", "defcur", "defdate", "defdbl", "defint", "deflng",
    "defobj", "defsng", "defstr", "defvar", "dim", "do", "each", "else",
    "elseif", "empty", "end", "endif", "enum", "eqv", "erase", "error",
    "event", "exit", "explicit", "false", "for", "friend", "function", "get",
    "global", "gosub", "goto", "if", "imp", "implements", "in", "input",
    "is", "let", "lib", "like", "load", "lock", "loop", "lset", "me", "mod",
    "new", "next", "not", "nothing", "null", "on", "option", "optional", "or",
    "paramarray", "preserve", "print", "private", "property", "public",
    "raiseevent", "randomize", "redim", "rem", "resume", "return", "rset",
    "select", "set", "setattr", "single", "static", "step", "stop", "sub",
    "then", "to", "true", "type", "typeof", "unload", "until", "wend",
    "while", "with", "withevents", "write", "xor", "base", "binary", "text",
    "module", "spc", "tab", "output", "append", "random", "read", "shared",
    "len", "line", "close", "open", "seek", "width", "reset", "kill", "name",
}

TYPES = {
    "boolean", "byte", "currency", "date", "double", "integer", "long",
    "longlong", "longptr", "object", "single", "string", "variant", "any",
    "workbook", "worksheet", "range", "chart", "chartobject", "listobject",
    "listrow", "listcolumn", "querytable", "workbookconnection", "name",
    "shape", "collection", "dictionary", "xlcalculation", "series",
    "seriescollection", "hyperlink", "pivottable", "comment", "font",
    "interior", "borders", "border", "validation", "workbooks", "worksheets",
    "sheets", "cells", "characters", "sortfields", "sort", "filedialog",
}

FUNCTIONS = {
    # conversion / maths
    "abs", "asc", "ascb", "ascw", "atn", "chr", "chrb", "chrw", "cos", "exp",
    "fix", "hex", "int", "log", "oct", "rnd", "round", "sgn", "sin", "sqr",
    "tan", "val",
    # strings
    "format", "instr", "instrrev", "join", "lcase", "left", "leftb", "ltrim",
    "mid", "midb", "replace", "right", "rightb", "rtrim", "space", "split",
    "str", "strcomp", "strconv", "string", "strreverse", "trim", "ucase",
    "filter", "formatcurrency", "formatdatetime", "formatnumber",
    "formatpercent",
    # dates
    "dateadd", "datediff", "datepart", "dateserial", "datevalue", "day",
    "hour", "minute", "month", "monthname", "now", "second", "time",
    "timer", "timeserial", "timevalue", "weekday", "weekdayname", "year",
    # arrays / info
    "array", "isarray", "isdate", "isempty", "iserror", "ismissing",
    "isnull", "isnumeric", "isobject", "lbound", "ubound", "typename",
    "vartype", "iif", "choose", "switch", "createobject", "getobject",
    "environ", "command", "doevents", "shell", "appactivate", "sendkeys",
    "msgbox", "inputbox", "beep", "dir", "fileattr", "filedatetime",
    "filelen", "freefile", "loc", "lof", "eof", "curdir", "mkdir", "rmdir",
    "chdir", "chdrive", "partition", "qbcolor", "rgb", "sln", "syd", "ddb",
    "fv", "ipmt", "irr", "mirr", "nper", "npv", "pmt", "ppmt", "pv", "rate",
    "cbool", "cbyte", "ccur", "cdate", "cdbl", "cdec", "cint", "clng",
    "clngptr", "clnglng", "csng", "cstr", "cvar", "cverr", "cvdate",
}

OBJECTS = {
    "application", "thisworkbook", "activeworkbook", "activesheet",
    "activecell", "activewindow", "selection", "workbooks", "worksheets",
    "sheets", "range", "cells", "columns", "rows", "err", "debug",
    "screenupdating", "vba", "excel", "userform", "forms", "clipboard",
    "wscript", "scripting", "fso",
}

# Excel / VBA enumeration members and constants that appear bare.
CONSTANT_PREFIXES = ("xl", "vb", "msg", "mso", "ad", "wd", "ol", "cd", "vbext")

EXTRA_CONSTANTS = {
    "pi", "true", "false", "nothing", "null", "empty",
    # compile-time conditional constants supplied by the VBA host
    "win64", "win32", "win16", "mac", "vba6", "vba7",
}

ALL_BARE = KEYWORDS | TYPES | FUNCTIONS | OBJECTS | EXTRA_CONSTANTS


def is_builtin(name: str) -> bool:
    low = name.lower()
    if low in ALL_BARE:
        return True
    # xlCalculationManual / vbCritical / msoTextOrientation...
    for pref in CONSTANT_PREFIXES:
        if low.startswith(pref) and len(low) > len(pref) and low[len(pref)].isalpha():
            # only treat as a constant if it is camelCase-ish (has an
            # uppercase letter after the prefix in the original spelling)
            if name[len(pref):len(pref) + 1].isupper():
                return True
    return False
