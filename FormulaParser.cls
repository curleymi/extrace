VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormulaParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   Michael Curley
'



' ----- class local fields -----
Private initialized As Boolean
Private specialDelimiter As String
Private operatorTokens() As Variant
Private delimiterTokens() As Variant
Private stringTokens() As Variant
Private sheetNames() As String
Private currentSheetIndex As Integer
Private lastToken As String

' ----- master data set upon initialize -----
Private originCell As String
Private masterFormula As String
Private masterParsedFormula() As String
Private masterSheetIndexes() As String
Private masterGotoReferences() As String
Private masterFormulaSplits() As String

' ----- references split data -----
Private referencesParsedFormula() As String
Private referencesSheetIndexes() As String
Private referencesGotoReferences() As String
Private referencesFormulaSplits() As String

' ----- argument split data -----
Private argumentParsedFormula() As String
Private argumentSheetIndexes() As String
Private argumentGotoReferences() As String
Private argumentFormulaSplits() As String

' ----- the data returned to caller -----
Private specifiedParsedFormula() As String
Private specifiedSheetIndexes() As String
Private specifiedGotoReferences() As String
Private specifiedFormulaSplits() As String





' ----- constructor, set initial values ----------------------------------------

Private Sub Class_Initialize()
    specialDelimiter = vbNewLine ' a token that should never appear in a formula
    operatorTokens = Array("<>", "<=", ">=", "<", ">", "=", "+", "-", "*", "/", "&", "%", "!")
    delimiterTokens = Array(")", "(", ",")
    stringTokens = Array("""", "'")
    sheetNames = parseSheetNames()
    currentSheetIndex = ActiveSheet.Index
    originCell = ActiveCell.Address
    If ActiveCell.HasFormula Then
        parse (ActiveCell.formula)
        initialized = True
        splitByAll ' default split type
    Else
        initialized = False
    End If
End Sub





' ----- privates ---------------------------------------------------------------

' conditionally appends the delimiter if the base or added string is not empty
Private Function appendStr(base As String, str As String) As String
    Dim str2 As String
    str2 = Trim(str)
    If Len(base) = 0 Then
        appendStr = str2
    ElseIf Len(str2) = 0 Then
        appendStr = base
    Else
        appendStr = base & specialDelimiter & str2
    End If
End Function



' returns the string as an array given that the appendStr function was used
Private Function toArr(str As String) As String()
    toArr = Split(str, specialDelimiter)
End Function



' returns if any special characters/strings are in the given string
Private Function containsSpecial(str As String) As Boolean
    containsSpecial = False
    For Each t In operatorTokens
        If InStr(str, t) > 0 Then
            containsSpecial = True
            Exit For
        End If
    Next
    If Not containsSpecial Then
        For Each t In delimiterTokens
            If InStr(str, t) > 0 Then
                containsSpecial = True
                Exit For
            End If
        Next
        If Not containsSpecial Then
            For Each t In stringTokens
                If InStr(str, t) > 0 Then
                    containsSpecial = True
                    Exit For
                End If
            Next
        End If
    End If
End Function



' builds an array of all sheet names
Private Function parseSheetNames() As String()
    Dim name As String
    Dim names As String
    names = ""
    For Each s In Sheets
        name = s.name
        If InStr(name, " ") > 0 Or containsSpecial(name) Then
            name = "'" & name & "'"
        End If
        names = appendStr(names, name)
    Next
    parseSheetNames = toArr(names)
End Function



' generic to check if a string starts with any value of types
' sets lastToken with the value on true
Private Function isOfType(str As String, types As Variant) As Boolean
    isOfType = False
    If Len(str) > 0 Then
        lastToken = ""
        For Each t In types
            If InStr(str, t) = 1 Then
                lastToken = t
                isOfType = True
                Exit For
            End If
        Next t
    End If
End Function



' returns if the string starts with an operator token
Private Function isOperator(str As String) As Boolean
    isOperator = isOfType(str, operatorTokens)
End Function



' returns if the string starts with a delimiting token
Private Function isDelimiter(str As String) As Boolean
    isDelimiter = isOperator(str)
    If Not isDelimiter Then
        isDelimiter = isOfType(str, delimiterTokens)
    End If
End Function



' returns if the string starts with a string token
Private Function isString(str As String) As Boolean
    isString = isOfType(str, stringTokens)
End Function



' returns the index of the sheet if it is one and -1 if not
Private Function indexSheetName(token As String) As Integer
    Dim i As Integer
    i = 1 ' indexing for sheets starts at 1
    indexSheetName = -1
    For Each name In sheetNames
        If InStr(token, name) = 1 Then
            indexSheetName = i
            Exit For
        End If
        i = i + 1
    Next
End Function



' returns the index of the sheet given master/specifiedSheetIndexes
Private Function convertIndex(indexes() As String, i As Integer) As Integer
    convertIndex = CInt(indexes(i))
End Function



' returns a substring of the given string from the given index to end
' indexing starts from 0
Private Function substring(str As String, i As Integer) As String
    substring = Right(str, Len(str) - i)
End Function



' breaks the given string into the tokens defined by local fields
Private Function tokenize(formula As String) As String
    Dim i As Integer
    Dim j As Integer
    Dim d As Integer
    Dim form As String
    Dim curForm As String
    i = 0
    form = formula
    tokenize = ""
    Do While i < Len(form)
        curForm = substring(form, i)
        ' ----- strings ------------------------------------
        If isString(curForm) Then
            tokenize = appendStr(tokenize, Left(form, i))
            d = 0
            j = Len(lastToken)
            i = InStr(j + 1, curForm, lastToken)
            Do While lastToken = """" And i > 0
                If InStr(j + 1, curForm, """""") = i Then
                    d = d + 1
                    j = i + 1
                    i = InStr(j + 1, curForm, """")
                Else
                    ' discrepancy between length and literal double quote identifiers
                    i = i - (2 * d)
                    Exit Do
                End If
            Loop
            If i <= 0 Then
                tokenize = "ERROR"
                form = ""
                Exit Do
            End If
            i = i + j - 1
            tokenize = appendStr(tokenize, Left(curForm, i))
            form = substring(curForm, i)
            i = 0
        ' ----- delimiters ---------------------------------
        ElseIf isDelimiter(curForm) Then
            tokenize = appendStr(tokenize, Left(form, i))
            j = Len(lastToken)
            tokenize = appendStr(tokenize, Left(curForm, j))
            form = substring(curForm, j)
            i = 0
        ' ----- arguments/unknown --------------------------
        Else
            i = i + 1
        End If
    Loop
    tokenize = appendStr(tokenize, form)
End Function



' combines open parens with preceeding function names, assuming they are separated
Private Function combineFunctionNames(tokenized As String) As String
    Dim token As String
    Dim prevToken As String
    prevToken = ""
    combineFunctionNames = ""
    For Each t In toArr(tokenized)
        token = t
        If token = "(" And Not isString(prevToken) And Not isDelimiter(prevToken) Then
            combineFunctionNames = combineFunctionNames & token
            prevToken = prevToken & token
        Else
            combineFunctionNames = appendStr(combineFunctionNames, token)
            prevToken = token
        End If
    Next
End Function



' builds the arrays of tokens and reference links
Private Function combineAndLinkReferences(tokenized As String)
    Dim i As Integer
    Dim refs As String
    Dim token As String
    Dim tokens As String
    Dim prevToken As String
    Dim refFound As Boolean
    i = -1 ' invalid sheet index
    refs = ""
    tokens = ""
    prevToken = ""
    For Each t In toArr(tokenized)
        token = t
        If i >= 0 Then
            ' ! denotes a sheet came before, skip to next to append the range text
            If t <> "!" Then
                tokens = appendStr(tokens, prevToken & "!" & token)
                refs = appendStr(refs, CStr(i))
                i = -1
                prevToken = ""
            End If
        Else
            i = indexSheetName(token)
            If i < 0 Then
                tokens = appendStr(tokens, token)
                refs = appendStr(refs, CStr(currentSheetIndex))
            End If
            prevToken = token
        End If
    Next
    masterParsedFormula = toArr(tokens)
    masterSheetIndexes = toArr(refs)
End Function



' sets the single masterFormula string from masterParsedFormula array
Private Function setMasterFormula()
    masterFormula = ""
    For Each t In masterParsedFormula
        masterFormula = masterFormula & t
    Next
End Function



' populates the masterGotoReference array
Private Function createMasterGotos()
    Dim i As Integer
    Dim j As Integer
    Dim shetIndex As Integer
    Dim gotos As String
    Dim token As String
    i = 0
    gotos = ""
    For Each t In masterParsedFormula
        token = t
        j = indexSheetName(token)
        If containsSpecial(token) And j < 0 Then
            gotos = appendStr(gotos, originCell)
        Else
            If j >= 0 Then
                token = Replace(token, sheetNames(j - 1) & "!", "")
            End If
            On Error Resume Next
            Set test = ActiveWorkbook.Sheets(convertIndex(masterSheetIndexes, i)).Range(token)
            If Err = 0 Then
                gotos = appendStr(gotos, token)
            Else
                Err.Clear
                gotos = appendStr(gotos, originCell)
            End If
        End If
        i = i + 1
    Next
    masterGotoReferences = toArr(gotos)
End Function



' sets masterFormulaSplits, this array is split by any delimiting token or string
Private Function createMasterFormulaSplits()
    Dim i As Integer
    Dim fLen As Integer
    Dim tLen As Integer
    Dim splits As String
    i = 0
    fLen = Len(masterFormula)
    splits = ""
    For Each t In masterParsedFormula
        tLen = Len(t)
        splits = appendStr(splits, i & "," & tLen & "," & fLen - i - tLen)
        i = i + tLen
    Next
    masterFormulaSplits = toArr(splits)
End Function



' sets all argument fields (tokens that are not special)
Private Function createArgumentFormulaSplits()
    Dim i As Integer
    Dim iSplit As Integer
    Dim token As String
    Dim formula As String
    Dim indexes As String
    Dim gotos As String
    Dim splits As String
    Dim fLen As Integer
    Dim tLen As Integer
    i = 0
    iSplit = 0
    formula = ""
    indexes = ""
    gotos = ""
    splits = ""
    fLen = Len(masterFormula)
    For Each t In masterParsedFormula
        token = t
        tLen = Len(token)
        If Not containsSpecial(token) Or isString(token) Or token = "=" Or indexSheetName(token) >= 0 Then
            formula = appendStr(formula, token)
            indexes = appendStr(indexes, masterSheetIndexes(i))
            gotos = appendStr(gotos, masterGotoReferences(i))
            splits = appendStr(splits, iSplit & "," & tLen & "," & fLen - iSplit - tLen)
        End If
        i = i + 1
        iSplit = iSplit + tLen
    Next
    argumentParsedFormula = toArr(formula)
    argumentSheetIndexes = toArr(indexes)
    argumentGotoReferences = toArr(gotos)
    argumentFormulaSplits = toArr(splits)
End Function



' sets all reference fields (tokens that reference any cell other than origin)
Private Function createReferencesFormulaSplits()
    Dim i As Integer
    Dim iSplit As Integer
    Dim token As String
    Dim ref As String
    Dim formula As String
    Dim indexes As String
    Dim gotos As String
    Dim splits As String
    Dim fLen As Integer
    Dim tLen As Integer
    i = 0
    iSplit = 0
    formula = ""
    indexes = ""
    gotos = ""
    splits = ""
    fLen = Len(masterFormula)
    For Each t In masterParsedFormula
        token = t
        tLen = Len(token)
        ref = masterGotoReferences(i)
        If ref <> originCell Or token = "=" Then
            formula = appendStr(formula, token)
            indexes = appendStr(indexes, masterSheetIndexes(i))
            gotos = appendStr(gotos, ref)
            splits = appendStr(splits, iSplit & "," & tLen & "," & fLen - iSplit - tLen)
        End If
        i = i + 1
        iSplit = iSplit + tLen
    Next
    referencesParsedFormula = toArr(formula)
    referencesSheetIndexes = toArr(indexes)
    referencesGotoReferences = toArr(gotos)
    referencesFormulaSplits = toArr(splits)
End Function



' parses the given formula to an array of individual tokens
' sets the masterParsedFormula and masterSheetIndexes class locals
Private Function parse(formula As String)
    combineAndLinkReferences (combineFunctionNames(tokenize(formula)))
    setMasterFormula
    createMasterGotos
    createMasterFormulaSplits
    createArgumentFormulaSplits
    createReferencesFormulaSplits
End Function





' ----- publics ----------------------------------------------------------------

' modifies the return of data to index by all
Public Function splitByAll()
    If initialized Then
        specifiedParsedFormula = masterParsedFormula
        specifiedSheetIndexes = masterSheetIndexes
        specifiedGotoReferences = masterGotoReferences
        specifiedFormulaSplits = masterFormulaSplits
    End If
End Function



' modifies the return of data() to index by references only
Public Function splitByReferences()
    If initialized Then
        specifiedParsedFormula = referencesParsedFormula
        specifiedSheetIndexes = referencesSheetIndexes
        specifiedGotoReferences = referencesGotoReferences
        specifiedFormulaSplits = referencesFormulaSplits
    End If
End Function



' modifies the return of data() to index by arguments only
Public Function splitByArguments()
    If initialized Then
        specifiedParsedFormula = argumentParsedFormula
        specifiedSheetIndexes = argumentSheetIndexes
        specifiedGotoReferences = argumentGotoReferences
        specifiedFormulaSplits = argumentFormulaSplits
    End If
End Function



' returns the master formula
Public Function formula() As String
    If initialized Then
        formula = masterFormula
    Else
        formula = "< NONE >"
    End If
End Function


' the main function to return the data associated with parsing
Public Function data() As String()
    If initialized Then
        data = specifiedParsedFormula
    Else
        data = Split("", ",")
    End If
End Function



' links the current sheet to the value of data(i) index
' returns the array where the data(i) is isolated between two other strings
Public Function gotoDataAndSplitFormula(i As Integer) As String()
    Dim splitFormula(3) As String
    Dim splitIndexes() As String
    Dim i0 As Integer
    Dim i1 As Integer
    Dim i2 As Integer
    If initialized Then
        splitIndexes = Split(specifiedFormulaSplits(i), ",")
        i0 = convertIndex(splitIndexes, 0)
        i1 = convertIndex(splitIndexes, 1)
        i2 = convertIndex(splitIndexes, 2)
        splitFormula(0) = Left(masterFormula, i0)
        splitFormula(1) = Mid(masterFormula, i0 + 1, i1)
        splitFormula(2) = Right(masterFormula, i2)
        Application.Goto (ActiveWorkbook.Sheets(convertIndex(specifiedSheetIndexes, i)).Range(specifiedGotoReferences(i)))
    Else
        splitFormula(0) = ""
        splitFormula(1) = ""
        splitFormula(2) = ""
    End If
    gotoDataAndSplitFormula = splitFormula
End Function





