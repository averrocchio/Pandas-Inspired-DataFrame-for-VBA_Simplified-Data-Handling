'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "DataFrame"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'@ModuleDescription "DataFrame fluente su Variant(,) 1-based: lettura/scrittura da Range/ListObject; operazioni in memoria; diagnostica opzionale."
'@Folder("LibreriaVBA/Facade")
Option Explicit

'==============================================================================
' INDEX (usa Ctrl+F su [SEC] per saltare alle sezioni)
' [SEC] Lifecycle
' [SEC] Properties RO
' [SEC] Properties RW
' [SEC] Loaders
' [SEC] Core Ops
' [SEC] IO
' [SEC] Diagnostics
' [SEC] Helpers (Array/Header/Guards)
'==============================================================================

' ====== Campi privati ========================================================
Private mData As Variant            ' Array 2D (1..R, 1..C)
Private mHeader As Variant          ' Array 1D (1..C) con nomi colonna
Private mKeys As Variant            ' Array 1D di indici colonna (opz.)
Private mNullToken As String        ' Es. "", "-", "NA"
Private mAliasMap As Object         ' Dictionary (incapsulato)
Private mTypes As Variant           ' Array 1D con tipo per colonna (opz.)
Private mMetrics As Object          ' Dictionary per metriche
Private mDebug As Boolean

' ====== Errori & utilità (privati) ==========================================
Private Enum DFErr
    edfInvalidData = 11010
    edfInvalidHeader = 11020
    edfNonLinearHeader = 11030
    edfJaggedRowsMismatch = 11040
    edfInvalidKeysSpec = 11050
    edfHeaderKeysMismatch = 11060
    edfRuntimeError = 11999
End Enum


Public Enum DFHeaderAt
    dfRow = 1
    dfColumn = 2
End Enum


Public Enum DFBounds
    dfliteral = 1
    dfCurrentRegion = 2
    dfTrimOuterEmpty = 3
    dfSmart = 4
End Enum



Private Sub RaiseDf(ByVal src As String, ByVal code As DFErr, ByVal message As String)
    Err.Raise vbObjectError + code, src, message
End Sub

'==============================================================================
' [SEC] Lifecycle
' RD_BOOKMARK: DF: Lifecycle
'==============================================================================
Private Sub Class_Initialize()
    mNullToken = ""                 ' default: cella vuota in output
    Set mMetrics = CreateObject("Scripting.Dictionary")
    Set mAliasMap = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set mMetrics = Nothing
    Set mAliasMap = Nothing
End Sub

'==============================================================================
' [SEC] Properties RO
' RD_BOOKMARK: DF: Props RO
'==============================================================================
'' @Description: Restituisce il numero di righe (0 se mData non è inizializzato)"
Public Property Get RowsCount() As Long
    If IsArray(mData) Then
        RowsCount = UBound(mData, 1) - LBound(mData, 1) + 1
    Else
        RowsCount = 0
    End If
End Property

'@Description "Restituisce il numero di colonne (0 se mData non è inizializzato)"
Public Property Get ColsCount() As Long
    If IsArray(mData) Then
        ColsCount = UBound(mData, 2) - LBound(mData, 2) + 1
    Else
        ColsCount = 0
    End If
End Property

'@Description "Forma R×C in stringa (es. "12×5")"
Public Property Get Shape() As String
    Shape = CStr(RowsCount) & "×" & CStr(ColsCount)
End Property

'@Description "Copia difensiva dell'header 1D 1-based"
Public Property Get header() As Variant
    If Not IsArray(mHeader) Then Exit Property
    Dim lo As Long, hi As Long, n As Long, i As Long
    lo = LBound(mHeader): hi = UBound(mHeader): n = hi - lo + 1
    Dim out As Variant: ReDim out(1 To n)
    For i = 1 To n
        out(i) = mHeader(lo + i - 1)
    Next i
    header = out
End Property

'==============================================================================
' [SEC] Properties RW
' RD_BOOKMARK: DF: Props RW
'==============================================================================
'@Description "Token per null/blank in output"
Public Property Get NullToken() As String
    NullToken = mNullToken
End Property

Public Property Let NullToken(ByVal value As String)
    mNullToken = value
End Property

'@Description "Abilita/disabilita diagnostica"
Public Property Get DebugMode() As Boolean
    DebugMode = mDebug
End Property

Public Property Let DebugMode(ByVal value As Boolean)
    mDebug = value
End Property

'@Description "Chiavi come array 1D di indici (1-based)"
Public Property Get Keys() As Variant
    If IsArray(mKeys) Then
        Keys = mKeys
    End If
End Property

'@Description "Imposta le chiavi da stringa "A,B" o array di nomi/indici"
'  TODO[DF:Keys]: implementare alias, validazione indici 1..ColsCount, normalizzazione nomi
Public Property Let Keys(ByVal value As Variant)
    Dim idx() As Long
    If IsArray(value) Then
        ' Caso: array di numeri o stringhe
        idx = Hdr_ResolveIndicesFromArray(value, mHeader, mAliasMap)
    Else
        ' Caso: stringa separata da virgole
        idx = Hdr_ResolveIndicesFromSpec(CStr(value), mHeader, mAliasMap)
    End If
    If (Not Arr_IsValidIndices(idx, ColsCount)) Then _
        RaiseDf "DataFrame.Keys", edfInvalidKeysSpec, "Chiavi fuori intervallo o duplicate."
    mKeys = idx
End Property

'==============================================================================
' [SEC] Loaders
' RD_BOOKMARK: DF: Loaders
'==============================================================================
'@Description "Crea/riusa l'istanza caricando dati da array + header"
Public Function LoadFromArray(ByRef data As Variant, ByRef header As Variant) As DataFrame

    If Not IsArray(data) Then RaiseDf "DataFrame.LoadFromArray", edfInvalidData, "Parametro 'data' non è un array."
    If Not IsArray(header) Then RaiseDf "DataFrame.LoadFromArray", edfInvalidHeader, "Parametro 'header' non è un array."

    ' DATA: fast-path 2D, altrimenti jagged?2D
    If Arr_Is2D(data) Then
        mData = data
    ElseIf IsArray(data(LBound(data))) Then
        mData = Arr_JaggedTo2D1(data)
    Else
        RaiseDf "DataFrame.LoadFromArray", edfInvalidData, "Array 'data' non riconosciuto (né 2D né jagged)."
    End If

    ' HEADER: 2D lineare ? 1D, altrimenti 1D normalizzato
    If Arr_Is2D(header) Then
        mHeader = Hdr_2DLinearTo1D(header)
    Else
        mHeader = Arr_To1D1(header)
    End If

    ' Invarianti
    G_EnsureDataHeaderAgree "DataFrame.LoadFromArray"

    Set LoadFromArray = Me
End Function

'@Description "Carica da ListObject (opz. riga header)"
'  TODO[DF:Loaders]: implementare lettura da lo.HeaderRowRange / lo.DataBodyRange
'@Description "Carica da ListObject (header da ListColumns; dati da DataBodyRange)"
Public Function LoadFromListObject(ByVal lo As ListObject) As DataFrame
    Dim t0 As Double: t0 = Timer
    
    If lo Is Nothing Then _
        RaiseDf "DataFrame.LoadFromListObject", edfInvalidHeader, "ListObject non valido (Nothing)."
    If lo.ListColumns.Count = 0 Then _
        RaiseDf "DataFrame.LoadFromListObject", edfInvalidHeader, "Il ListObject non ha colonne."

    ' ----- Header: dai nomi delle ListColumns, normalizzati e resi unici -----
    Dim c As Long: c = lo.ListColumns.Count
    Dim hdr As Variant: ReDim hdr(1 To c)
    Dim i As Long, nameRaw As String
    
    For i = 1 To c
        nameRaw = CStr(lo.ListColumns(i).Name)
        hdr(i) = CollapseSpaces(Trim$(nameRaw))
    Next i
    hdr = Hdr_EnsureUnique1D(hdr)  ' suffissi _2, _3... in caso di duplicati (case-insensitive)
    mHeader = hdr

    ' ----- Dati: lettura in blocco; gestisci tabella vuota come 0×C -----
    Dim body As Range
    On Error Resume Next
    Set body = lo.DataBodyRange    ' Nothing se la tabella non ha righe
    On Error GoTo 0
    
    Dim data As Variant
    If body Is Nothing Then
        ReDim data(1 To 0, 1 To c) ' 0 righe, C colonne
    Else
        data = body.Value2         ' sempre 2D per range multi-cella
    End If
    mData = data

    ' ----- Invarianti & metriche -----
    G_EnsureDataHeaderAgree "DataFrame.LoadFromListObject"
    
    Dim dt As Double: dt = Timer - t0
    If dt < 0# Then dt = dt + 24# * 3600#   ' rollover a mezzanotte
    mMetrics("source") = "ListObject:" & lo.Name
    mMetrics("rows") = RowsCount
    mMetrics("cols") = ColsCount
    mMetrics("read_ms") = CLng(dt * 1000#)
    mMetrics("empty_body") = (RowsCount = 0)

    If mDebug Then
        Debug.Print "DF.LoadFromListObject"; _
                    " name=" & lo.Name; _
                    " rows=" & RowsCount; _
                    " cols=" & ColsCount; _
                    " ms=" & CLng(dt * 1000#)
    End If

    Set LoadFromListObject = Me
End Function



'@Description "Carica da Range (hasHeader/headerAt) con bounds: literal | currentRegion | trimOuterEmpty | smart"



Public Function LoadFromRange(ByVal rng As Range, _
                                Optional ByVal hasHeader As Boolean = True, _
                                Optional ByVal headerAt As DFHeaderAt = dfRow, _
                                Optional ByVal bounds As DFBounds = dfliteral, _
                                Optional ByVal maxExtendRows As Long = 100000) As DataFrame


Dim t0 As Double: t0 = Timer

 

    If rng Is Nothing Then _
        RaiseDf "DataFrame.LoadFromRange", edfInvalidData, "Range non valido (Nothing)."
    If rng.Areas.Count > 1 Then _
        RaiseDf "DataFrame.LoadFromRange", edfInvalidData, "Range non contiguo (Areas.Count>1)."

    Dim ws As Worksheet: Set ws = rng.Worksheet
    Dim baseRng As Range, workRng As Range, finalRng As Range
    Dim arr As Variant
    Dim r As Long, c As Long               ' dimensioni array (uso generico)
    Dim i As Long, j As Long               ' indici loop (unica dichiarazione)
    Dim orient As String
    Dim cutTop As Long, cutBottom As Long, cutLeft As Long, cutRight As Long
    Dim extendedRows As Long: extendedRows = 0



    ' -------- 1) Determina range di base secondo bounds --------

    Select Case bounds 'LCase$(bounds)
        Case dfCurrentRegion, dfSmart
            Set baseRng = rng.Cells(1, 1).CurrentRegion
        Case dfliteral, dfTrimOuterEmpty
            Set baseRng = rng
        Case Else
            RaiseDf "DataFrame.LoadFromRange", edfInvalidData, "bounds non riconosciuto: '" & bounds & "'."
    End Select


    ' -------- 2) (eventuale) trim dei bordi esterni vuoti in RAM --------

    'If LCase$(bounds) = "trimouterempty" Or LCase$(bounds) = "smart" Then

    If bounds = dfTrimOuterEmpty Or bounds = dfSmart Then
        arr = baseRng.Value2         ' lettura in blocco
        r = UBound(arr, 1) - LBound(arr, 1) + 1
        c = UBound(arr, 2) - LBound(arr, 2) + 1
        Call DetectOuterEmptyCuts(arr, cutTop, cutBottom, cutLeft, cutRight)
        
        ' Se tutto vuoto, non tagliare oltre il minimo (lascia baseRng)
        If cutTop = 0 And cutBottom = 0 And cutLeft = 0 And cutRight = 0 Then
            Set workRng = baseRng
        Else
            Set workRng = baseRng.Offset(cutTop, cutLeft).Resize( _
                          baseRng.rows.Count - cutTop - cutBottom, _
                          baseRng.Columns.Count - cutLeft - cutRight)
        End If
    Else
        Set workRng = baseRng
    End If


    ' -------- 3) smart: estensione "header-down" solo verticale --------

    ' Vale solo se hasHeader=True e orientamento effettivo = "row"

    orient = ResolveHeaderAt(headerAt, workRng.rows.Count, workRng.Columns.Count)

    'If LCase$(bounds) = "smart" And hasHeader And LCase$(orient) = "row" Then

    If bounds = dfSmart And hasHeader And LCase$(orient) = "row" Then

        Dim headerRow As Long: headerRow = workRng.Row

        Dim leftCol As Long:    leftCol = workRng.Column

        Dim rightCol As Long:   rightCol = workRng.Column + workRng.Columns.Count - 1


        ' cap inferiore: UsedRange o max estensione

        Dim usedBottom As Long

        usedBottom = ws.UsedRange.Row + ws.UsedRange.rows.Count - 1

        Dim bottomCap As Long

        bottomCap = headerRow + maxExtendRows

        If usedBottom < bottomCap Then bottomCap = usedBottom

        If bottomCap < headerRow Then bottomCap = headerRow


        ' Leggi blocco colonne fisse da headerRow..bottomCap e cerca last non-vuota (dal basso)

        Dim bigRng As Range

        Set bigRng = ws.Range(ws.Cells(headerRow, leftCol), ws.Cells(bottomCap, rightCol))

        Dim bigArr As Variant: bigArr = bigRng.Value2


        Dim lastDataRowRel As Long: lastDataRowRel = FindLastNonEmptyRow(bigArr, True) ' True = salta header

        Dim currentBottom As Long: currentBottom = workRng.Row + workRng.rows.Count - 1

        Dim newBottom As Long

     

        If lastDataRowRel = 0 Then

            ' nessun dato sotto l'header: lascia invariato

            Set finalRng = workRng

        Else

            newBottom = headerRow + lastDataRowRel - 1

            If newBottom > currentBottom Then

                Set finalRng = ws.Range(ws.Cells(workRng.Row, workRng.Column), ws.Cells(newBottom, rightCol))

                extendedRows = newBottom - currentBottom

            Else

                Set finalRng = workRng

            End If

        End If

    Else

        Set finalRng = workRng

    End If


    ' -------- 4) Costruzione header + dati in RAM --------

    Dim dataArr As Variant, hdr As Variant

    Dim rowsData As Long, colsData As Long


    Dim finalArr As Variant: finalArr = finalRng.Value2

    Dim rLo As Long, rHi As Long, cLo As Long, cHi As Long

    rLo = LBound(finalArr, 1): rHi = UBound(finalArr, 1)

    cLo = LBound(finalArr, 2): cHi = UBound(finalArr, 2)


    orient = ResolveHeaderAt(headerAt, (rHi - rLo + 1), (cHi - cLo + 1))


    If LCase$(orient) = "row" Then

        colsData = cHi - cLo + 1

        If hasHeader Then

            hdr = HeaderFromRow(finalArr, rLo, cLo, cHi)

            hdr = Hdr_EnsureUnique1D(hdr)

            ' dati = righe successive

            rowsData = (rHi - rLo)

            If rowsData <= 0 Then

                ReDim dataArr(1 To 0, 1 To colsData)

            Else

                ReDim dataArr(1 To rowsData, 1 To colsData)

                For i = 1 To rowsData

                    For j = 1 To colsData

                        dataArr(i, j) = finalArr(rLo + i, cLo + j - 1)

                    Next j

                Next i

            End If

        Else

            ' header generico C1..Cn, dati = tutto il blocco

            hdr = GenHeaderCols(colsData)

            rowsData = (rHi - rLo + 1)

            ReDim dataArr(1 To rowsData, 1 To colsData)

            Dim a As Long, b As Long, rr As Long, cc As Long

            rr = 0

            For a = rLo To rHi

                rr = rr + 1

                cc = 0

                For b = cLo To cHi

                    cc = cc + 1

                    dataArr(rr, cc) = finalArr(a, b)

                Next b

            Next a

        End If


    ElseIf LCase$(orient) = "col" Then

        rowsData = rHi - rLo + 1

        If hasHeader Then

            hdr = HeaderFromCol(finalArr, cLo, rLo, rHi)

            hdr = Hdr_EnsureUnique1D(hdr)

            ' dati = colonne successive

            colsData = (cHi - cLo)

            If colsData <= 0 Then

                ReDim dataArr(1 To rowsData, 1 To 0)

            Else

                ReDim dataArr(1 To rowsData, 1 To colsData)

                For i = 1 To rowsData

                    For j = 1 To colsData

                        dataArr(i, j) = finalArr(rLo + i - 1, cLo + j)

                    Next j

                Next i

            End If

        Else

            ' header generico R1..Rr, dati = tutto il blocco (orientamento col)

            colsData = cHi - cLo + 1

            hdr = GenHeaderCols(colsData) ' coerente con resto libreria

            ReDim dataArr(1 To rowsData, 1 To colsData)

            Dim r2 As Long, c2 As Long, rr2 As Long, cc2 As Long

            rr2 = 0

            For r2 = rLo To rHi

                rr2 = rr2 + 1

                cc2 = 0

                For c2 = cLo To cHi

                    cc2 = cc2 + 1

                    dataArr(rr2, cc2) = finalArr(r2, c2)

                Next c2

            Next r2

        End If

    Else

        RaiseDf "DataFrame.LoadFromRange", edfInvalidHeader, "headerAt non riconosciuto: '" & headerAt & "'."

    End If


    ' Normalizza esposizione header (trim/collapse)

    Dim k As Long

    For k = LBound(hdr) To UBound(hdr)

        hdr(k) = CollapseSpaces(Trim$(CStr(hdr(k))))

    Next k

    hdr = Hdr_EnsureUnique1D(hdr)


    ' -------- 5) Commit in oggetto + invarianti/metriche --------

    mHeader = hdr

    mData = dataArr

    G_EnsureDataHeaderAgree "DataFrame.LoadFromRange"


    ' metriche

    Dim dt As Double: dt = Timer - t0

    If dt < 0# Then dt = dt + 24# * 3600#

    mMetrics("source") = "Range:" & rng.Address(0, 0)

    'mMetrics("bounds_policy") = LCase$(bounds)

    mMetrics("bounds_policy") = bounds

    mMetrics("orientation") = LCase$(orient)

    mMetrics("rows") = RowsCount

    mMetrics("cols") = ColsCount

    mMetrics("extended_rows") = extendedRows

    'If LCase$(bounds) = "trimouterempty" Or LCase$(bounds) = "smart" Then

    If bounds = dfTrimOuterEmpty Or bounds = dfSmart Then

        mMetrics("trim_top") = cutTop

        mMetrics("trim_bottom") = cutBottom

        mMetrics("trim_left") = cutLeft

        mMetrics("trim_right") = cutRight

    End If

    mMetrics("read_ms") = CLng(dt * 1000#)


    If mDebug Then

        Debug.Print "DF.LoadFromRange"; _
                    " addr=" & rng.Address(0, 0); _
                    " bounds=" & bounds; _
                    " orient=" & LCase$(orient); _
                    " rows=" & RowsCount; _
                    " cols=" & ColsCount; _
                    " extRows=" & extendedRows; _
                    " ms=" & CLng(dt * 1000#)

    End If


    Set LoadFromRange = Me

End Function



'==============================================================================
' [SEC] Core Ops
' RD_BOOKMARK: DF: Ops Core
'==============================================================================
'@Description "Filtra righe in base a DSL condizionale"
Public Function Filter(ByVal condition As String) As DataFrame
    ' Mini-spec:
    ' 1) Supporta condizione singola nel formato: "col op valore".
    ' 2) Operatori MVP: =, <>, >, <, contains (case-insensitive per contains).
    ' 3) Preserva header e ordine righe originali (filtro stabile).
    ' 4) Gestisce DataFrame vuoto restituendo schema invariato.
    Dim out As New DataFrame
    G_EnsureDataHeaderAgree "DataFrame.Filter"

    Dim colIdx As Long, op As String, rhs As String
    ParseCondition condition, colIdx, op, rhs

    Dim keep() As Boolean
    ReDim keep(1 To RowsCount)

    Dim i As Long, kept As Long
    For i = 1 To RowsCount
        If EvalCondition(mData(i, colIdx), op, rhs) Then
            keep(i) = True
            kept = kept + 1
        End If
    Next i

    Dim outData As Variant
    If kept = 0 Then
        ReDim outData(1 To 0, 1 To ColsCount)
    Else
        ReDim outData(1 To kept, 1 To ColsCount)
    End If

    Dim rOut As Long, j As Long
    rOut = 0
    For i = 1 To RowsCount
        If keep(i) Then
            rOut = rOut + 1
            For j = 1 To ColsCount
                outData(rOut, j) = mData(i, j)
            Next j
        End If
    Next i

    out.mHeader = Arr_To1D1(mHeader)
    out.mData = outData
    out.mNullToken = mNullToken
    out.mDebug = mDebug
    out.mTypes = mTypes
    out.G_EnsureDataHeaderAgree "DataFrame.Filter"

    Set Filter = out
End Function

'@Description "Ordinamento stabile multi-colonna"
Public Function Sort(ByVal cols As String, Optional ByVal ascending As String = "") As DataFrame
    ' Mini-spec:
    ' 1) Ordina per una o più colonne (spec separata da virgole).
    ' 2) Direzioni opzionali via ascending="asc,desc,..." (default asc).
    ' 3) Implementazione stabile su indice righe (insertion sort MVP).
    ' 4) Nessun side-effect: ritorna nuovo DataFrame.
    Dim out As New DataFrame
    G_EnsureDataHeaderAgree "DataFrame.Sort"

    Dim idx() As Long
    idx = Hdr_ResolveIndicesFromSpec(cols, mHeader, mAliasMap)
    If Not Arr_IsValidIndices(idx, ColsCount) Then
        RaiseDf "DataFrame.Sort", edfInvalidKeysSpec, "Specifica colonne sort non valida."
    End If

    Dim asc() As Boolean
    asc = ParseSortDirections(ascending, UBound(idx) - LBound(idx) + 1)

    Dim ord() As Long
    Dim i As Long, j As Long, k As Long
    ReDim ord(1 To RowsCount)
    For i = 1 To RowsCount
        ord(i) = i
    Next i

    Dim keyRow As Long
    For i = 2 To RowsCount
        keyRow = ord(i)
        j = i - 1
        Do While j >= 1 And CompareRowsByCols(ord(j), keyRow, idx, asc) > 0
            ord(j + 1) = ord(j)
            j = j - 1
        Loop
        ord(j + 1) = keyRow
    Next i

    Dim outData As Variant
    ReDim outData(1 To RowsCount, 1 To ColsCount)
    For i = 1 To RowsCount
        For k = 1 To ColsCount
            outData(i, k) = mData(ord(i), k)
        Next k
    Next i

    out.mHeader = Arr_To1D1(mHeader)
    out.mData = outData
    out.mNullToken = mNullToken
    out.mDebug = mDebug
    out.mTypes = mTypes
    out.G_EnsureDataHeaderAgree "DataFrame.Sort"

    Set Sort = out
End Function

'@Description "Dedup su chiavi (policy: keep_first/keep_last/...)"
Public Function Dedup(Optional ByVal policy As String = "keep_first") As DataFrame
    ' Mini-spec:
    ' 1) Dedup su tutte le colonne o su subset in mKeys (se impostato).
    ' 2) Policy supportate: keep_first, keep_last.
    ' 3) Preserva schema invariato, ritorna nuovo DataFrame.
    Dim out As New DataFrame
    G_EnsureDataHeaderAgree "DataFrame.Dedup"

    Dim keysIdx() As Long
    If IsArray(mKeys) Then
        keysIdx = mKeys
        If Not Arr_IsValidIndices(keysIdx, ColsCount) Then
            RaiseDf "DataFrame.Dedup", edfInvalidKeysSpec, "mKeys non valido."
        End If
    Else
        ReDim keysIdx(1 To ColsCount)
        Dim c As Long
        For c = 1 To ColsCount
            keysIdx(c) = c
        Next c
    End If

    Dim keepLast As Boolean
    policy = LCase$(Trim$(policy))
    Select Case policy
        Case "keep_first": keepLast = False
        Case "keep_last": keepLast = True
        Case Else
            RaiseDf "DataFrame.Dedup", edfInvalidData, "Policy non supportata: " & policy
    End Select

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim keep() As Boolean
    ReDim keep(1 To RowsCount)

    Dim i As Long, key As String, kept As Long
    If keepLast Then
        For i = RowsCount To 1 Step -1
            key = BuildRowKey(mData, i, keysIdx)
            If Not seen.Exists(key) Then
                seen.Add key, True
                keep(i) = True
                kept = kept + 1
            End If
        Next i
    Else
        For i = 1 To RowsCount
            key = BuildRowKey(mData, i, keysIdx)
            If Not seen.Exists(key) Then
                seen.Add key, True
                keep(i) = True
                kept = kept + 1
            End If
        Next i
    End If

    Dim outData As Variant
    ReDim outData(1 To kept, 1 To ColsCount)

    Dim rOut As Long, j As Long
    rOut = 0
    For i = 1 To RowsCount
        If keep(i) Then
            rOut = rOut + 1
            For j = 1 To ColsCount
                outData(rOut, j) = mData(i, j)
            Next j
        End If
    Next i

    out.mHeader = Arr_To1D1(mHeader)
    out.mData = outData
    out.mNullToken = mNullToken
    out.mDebug = mDebug
    out.mTypes = mTypes
    out.G_EnsureDataHeaderAgree "DataFrame.Dedup"

    Set Dedup = out
End Function

'@Description "Proiezione/riordino colonne"
Public Function Project(ByVal cols As String) As DataFrame
    ' Mini-spec:
    ' 1) Accetta una lista colonne separata da virgole (nomi o indici).
    ' 2) Preserva esattamente ordine e subset richiesto nel parametro cols.
    ' 3) Valida esistenza/duplicati tramite resolver header centralizzato.
    ' 4) Copia dati in memoria su nuovo DataFrame senza side-effect Excel.
    ' 5) Gestisce DataFrame vuoto mantenendo schema coerente.
    Dim out As New DataFrame
    G_EnsureDataHeaderAgree "DataFrame.Project"

    Dim idx() As Long
    idx = Hdr_ResolveIndicesFromSpec(cols, mHeader, mAliasMap)
    If Not Arr_IsValidIndices(idx, ColsCount) Then
        RaiseDf "DataFrame.Project", edfInvalidKeysSpec, "Specifica colonne non valida o con duplicati."
    End If

    Dim nCols As Long
    nCols = UBound(idx) - LBound(idx) + 1

    Dim outHeader As Variant
    ReDim outHeader(1 To nCols)

    Dim outData As Variant
    Dim i As Long, j As Long
    ReDim outData(1 To RowsCount, 1 To nCols)

    For j = 1 To nCols
        outHeader(j) = mHeader(idx(LBound(idx) + j - 1))
        For i = 1 To RowsCount
            outData(i, j) = mData(i, idx(LBound(idx) + j - 1))
        Next i
    Next j

    Set out = NewFrameFromParts(outData, outHeader)

    Set Project = out
End Function

'@Description "Rinomina colonne in base a mapping"
Public Function Rename(ByVal map As Variant) As DataFrame
    ' Mini-spec:
    ' 1) Supporta mapping rename da Dictionary oppure stringa "old:new,...".
    ' 2) Le chiavi sono risolte con nomi header (case-insensitive).
    ' 3) Valida sintassi mapping e colonne mancanti con RaiseDf.
    ' 4) Normalizza header rinominato (trim/collapse) e garantisce univocità.
    ' 5) Non muta l'istanza corrente: ritorna nuovo DataFrame.
    Dim out As New DataFrame
    G_EnsureDataHeaderAgree "DataFrame.Rename"

    Dim outHeader As Variant
    outHeader = Arr_To1D1(mHeader)

    Dim pairs As Variant, pair As Variant, kv() As String
    Dim i As Long, oldName As String, newName As String
    Dim idx As Long

    If IsObject(map) Then
        If TypeName(map) <> "Dictionary" Then
            RaiseDf "DataFrame.Rename", edfInvalidHeader, "Mapping oggetto non supportato: usare Scripting.Dictionary."
        End If
        For Each pair In map.Keys
            oldName = CStr(pair)
            newName = CStr(map(pair))
            idx = Hdr_FindIndexByName(oldName, mHeader, mAliasMap)
            outHeader(idx) = CollapseSpaces(Trim$(newName))
        Next pair
    Else
        pairs = Split(CStr(map), ",")
        For i = LBound(pairs) To UBound(pairs)
            pair = Trim$(CStr(pairs(i)))
            If Len(pair) = 0 Then GoTo ContinueLoop
            kv = Split(CStr(pair), ":")
            If UBound(kv) <> 1 Then
                RaiseDf "DataFrame.Rename", edfInvalidHeader, "Mapping rename non valido: '" & pair & "'. Atteso old:new"
            End If
            oldName = Trim$(kv(0))
            newName = CollapseSpaces(Trim$(kv(1)))
            idx = Hdr_FindIndexByName(oldName, mHeader, mAliasMap)
            outHeader(idx) = newName
ContinueLoop:
        Next i
    End If

    outHeader = Hdr_EnsureUnique1D(outHeader)

    Dim dataCopy As Variant
    dataCopy = mData
    Set out = NewFrameFromParts(dataCopy, outHeader)

    Set Rename = out
End Function

'@Description "Append di righe con unione schema per nome colonna"
Public Function Append(ByVal rows As DataFrame) As DataFrame
    ' Mini-spec:
    ' 1) Appende le righe di un secondo DataFrame allineando le colonne per nome.
    ' 2) Lo schema risultato è l'unione: colonne sinistra + nuove colonne di destra.
    ' 3) Le colonne mancanti su un lato vengono valorizzate a vbNullString.
    ' 4) Preserva l'ordine colonne della sinistra, poi aggiunge le colonne extra della destra.
    ' 5) Copia in memoria in blocco su nuovo array 2D 1-based.
    Dim out As New DataFrame
    If rows Is Nothing Then
        RaiseDf "DataFrame.Append", edfInvalidData, "DataFrame da appendere non valido (Nothing)."
    End If

    G_EnsureDataHeaderAgree "DataFrame.Append"

    Dim rightData As Variant
    Dim rightHeader As Variant
    rightData = rows.AsArray
    rightHeader = rows.header

    If (Not IsArray(rightData)) Or (Not IsArray(rightHeader)) Then
        RaiseDf "DataFrame.Append", edfInvalidData, "DataFrame da appendere non inizializzato."
    End If

    Dim leftMap As Object, rightMap As Object
    Set leftMap = HeaderIndexMap(mHeader)
    Set rightMap = HeaderIndexMap(rightHeader)

    Dim outCols As Long
    outCols = ColsCount

    Dim j As Long
    For j = 1 To rows.ColsCount
        If Not leftMap.Exists(LCase$(CStr(rightHeader(j)))) Then
            outCols = outCols + 1
        End If
    Next j

    Dim outHeader As Variant
    ReDim outHeader(1 To outCols)

    Dim outColMap As Object
    Set outColMap = CreateObject("Scripting.Dictionary")

    For j = 1 To ColsCount
        outHeader(j) = mHeader(j)
        outColMap(LCase$(CStr(mHeader(j)))) = j
    Next j

    Dim wCol As Long
    wCol = ColsCount
    For j = 1 To rows.ColsCount
        Dim k As String
        k = LCase$(CStr(rightHeader(j)))
        If Not outColMap.Exists(k) Then
            wCol = wCol + 1
            outHeader(wCol) = rightHeader(j)
            outColMap(k) = wCol
        End If
    Next j

    Dim outRows As Long
    outRows = RowsCount + rows.RowsCount

    Dim outData As Variant
    ReDim outData(1 To outRows, 1 To outCols)

    Dim i As Long
    Dim srcCol As Long

    ' Righe sinistra
    For i = 1 To RowsCount
        For j = 1 To outCols
            srcCol = HeaderIndexOrZero(leftMap, CStr(outHeader(j)))
            If srcCol > 0 Then
                outData(i, j) = mData(i, srcCol)
            Else
                outData(i, j) = vbNullString
            End If
        Next j
    Next i

    ' Righe destra
    For i = 1 To rows.RowsCount
        For j = 1 To outCols
            srcCol = HeaderIndexOrZero(rightMap, CStr(outHeader(j)))
            If srcCol > 0 Then
                outData(RowsCount + i, j) = rightData(i, srcCol)
            Else
                outData(RowsCount + i, j) = vbNullString
            End If
        Next j
    Next i

    Set out = NewFrameFromParts(outData, outHeader)

    Set Append = out
End Function

'@Description "Join con destra, specifica chiavi e tipo"
Public Function JoinRight(ByVal right As DataFrame, ByVal Keys As String, _
                          Optional ByVal how As String = "inner", _
                          Optional ByVal suffixes As String = "_L,_R") As DataFrame
    ' Mini-spec:
    ' 1) MVP: join su chiavi con stessi nomi tra left (Me) e right.
    ' 2) Tipi supportati: inner, right (default inner).
    ' 3) Richiede chiavi univoche nel DataFrame left per evitare ambiguità.
    ' 4) Output: colonne left + colonne right non-key (con suffisso in collisione).
    Dim out As New DataFrame
    If right Is Nothing Then
        RaiseDf "DataFrame.JoinRight", edfInvalidData, "DataFrame right non valido (Nothing)."
    End If

    G_EnsureDataHeaderAgree "DataFrame.JoinRight"
    right.G_EnsureDataHeaderAgree "DataFrame.JoinRight"

    Dim leftKeys() As Long, rightKeys() As Long
    leftKeys = Hdr_ResolveIndicesFromSpec(Keys, mHeader, mAliasMap)
    rightKeys = Hdr_ResolveIndicesFromSpec(Keys, right.mHeader, right.mAliasMap)

    If Not Arr_IsValidIndices(leftKeys, ColsCount) Then _
        RaiseDf "DataFrame.JoinRight", edfInvalidKeysSpec, "Specifica chiavi left non valida."
    If Not Arr_IsValidIndices(rightKeys, right.ColsCount) Then _
        RaiseDf "DataFrame.JoinRight", edfInvalidKeysSpec, "Specifica chiavi right non valida."

    how = LCase$(Trim$(how))
    If how <> "inner" And how <> "right" Then
        RaiseDf "DataFrame.JoinRight", edfInvalidData, "how non supportato: " & how
    End If

    Dim sfxR As String
    sfxR = ParseSuffixRight(suffixes)

    Dim leftMap As Object: Set leftMap = CreateObject("Scripting.Dictionary")
    Dim i As Long, j As Long, key As String
    For i = 1 To RowsCount
        key = BuildRowKey(mData, i, leftKeys)
        If leftMap.Exists(key) Then
            RaiseDf "DataFrame.JoinRight", edfInvalidData, "Chiavi duplicate nel DataFrame left (MVP richiede univocità)."
        End If
        leftMap.Add key, i
    Next i

    Dim rightNonKeyIdx() As Long
    rightNonKeyIdx = BuildNonKeyIndex(right.ColsCount, rightKeys)

    Dim outCols As Long
    outCols = ColsCount + IIf(IsArray(rightNonKeyIdx), UBound(rightNonKeyIdx) - LBound(rightNonKeyIdx) + 1, 0)

    Dim outHeader As Variant
    ReDim outHeader(1 To outCols)

    For j = 1 To ColsCount
        outHeader(j) = mHeader(j)
    Next j

    Dim nRightAdd As Long
    nRightAdd = 0
    If IsArray(rightNonKeyIdx) Then nRightAdd = UBound(rightNonKeyIdx) - LBound(rightNonKeyIdx) + 1

    For j = 1 To nRightAdd
        Dim nm As String
        nm = CStr(right.mHeader(rightNonKeyIdx(LBound(rightNonKeyIdx) + j - 1)))
        If HeaderContains(outHeader, nm, ColsCount) Then nm = nm & sfxR
        outHeader(ColsCount + j) = nm
    Next j
    outHeader = Hdr_EnsureUnique1D(outHeader)

    Dim maxRows As Long
    maxRows = right.RowsCount
    Dim outData As Variant
    ReDim outData(1 To maxRows, 1 To outCols)

    Dim rOut As Long, leftRow As Long
    rOut = 0
    For i = 1 To right.RowsCount
        key = BuildRowKey(right.mData, i, rightKeys)
        If leftMap.Exists(key) Then
            leftRow = CLng(leftMap(key))
        Else
            leftRow = 0
        End If

        If how = "right" Or leftRow > 0 Then
            rOut = rOut + 1
            For j = 1 To ColsCount
                If leftRow > 0 Then
                    outData(rOut, j) = mData(leftRow, j)
                Else
                    outData(rOut, j) = Empty
                End If
            Next j

            For j = 1 To nRightAdd
                outData(rOut, ColsCount + j) = right.mData(i, rightNonKeyIdx(LBound(rightNonKeyIdx) + j - 1))
            Next j
        End If
    Next i

    If rOut < maxRows Then
        Dim trimmed As Variant
        ReDim trimmed(1 To rOut, 1 To outCols)
        For i = 1 To rOut
            For j = 1 To outCols
                trimmed(i, j) = outData(i, j)
            Next j
        Next i
        outData = trimmed
    End If

    out.mHeader = outHeader
    out.mData = outData
    out.mNullToken = mNullToken
    out.mDebug = mDebug
    out.G_EnsureDataHeaderAgree "DataFrame.JoinRight"

    Set JoinRight = out
End Function

'@Description "Pulizia base (trim/collapse/coercioni leggere)"
Public Function Clean(Optional ByVal trimSpaces As Boolean = True, _
                      Optional ByVal CollapseSpaces As Boolean = True, _
                      Optional ByVal softCoerce As Boolean = True) As DataFrame
    ' Mini-spec:
    ' 1) Normalizza stringhe (trim/collapse) su tutte le celle.
    ' 2) Converte token null comuni ("NA","N/A","NULL","-") in Empty.
    ' 3) Soft-coerce numeri/date quando richiesto.
    Dim out As New DataFrame
    G_EnsureDataHeaderAgree "DataFrame.Clean"

    Dim outData As Variant
    ReDim outData(1 To RowsCount, 1 To ColsCount)

    Dim i As Long, j As Long
    For i = 1 To RowsCount
        For j = 1 To ColsCount
            outData(i, j) = CleanValue(mData(i, j), trimSpaces, CollapseSpaces, softCoerce)
        Next j
    Next i

    out.mHeader = Arr_To1D1(mHeader)
    out.mData = outData
    out.mNullToken = mNullToken
    out.mDebug = mDebug
    out.G_EnsureDataHeaderAgree "DataFrame.Clean"

    Set Clean = out
End Function

'@Description "Inferenza soft dei tipi per colonna"
Public Function InferTypes() As DataFrame
    ' Mini-spec:
    ' 1) Inferenza soft per colonna: number/date/text.
    ' 2) Ignora Empty/Null/stringhe vuote durante il conteggio.
    ' 3) Salva il risultato in mTypes e nelle metriche.
    Dim out As New DataFrame
    G_EnsureDataHeaderAgree "DataFrame.InferTypes"

    Dim types As Variant
    ReDim types(1 To ColsCount)

    Dim j As Long, i As Long
    Dim nNum As Long, nDate As Long, nTxt As Long
    Dim v As Variant, s As String

    For j = 1 To ColsCount
        nNum = 0: nDate = 0: nTxt = 0
        For i = 1 To RowsCount
            v = mData(i, j)
            If Not IsNull(v) And Not IsEmptyLike(v) Then
                If IsDate(v) Then
                    nDate = nDate + 1
                ElseIf IsNumeric(v) Then
                    nNum = nNum + 1
                Else
                    s = Trim$(CStr(v))
                    If Len(s) > 0 Then
                        If IsDate(s) Then
                            nDate = nDate + 1
                        ElseIf IsNumeric(s) Then
                            nNum = nNum + 1
                        Else
                            nTxt = nTxt + 1
                        End If
                    End If
                End If
            End If
        Next i

        If nNum >= nDate And nNum >= nTxt Then
            types(j) = "number"
        ElseIf nDate >= nNum And nDate >= nTxt Then
            types(j) = "date"
        Else
            types(j) = "text"
        End If
    Next j

    out.mHeader = Arr_To1D1(mHeader)
    out.mData = mData
    out.mNullToken = mNullToken
    out.mDebug = mDebug
    out.mTypes = types
    out.mMetrics("infer_types_cols") = ColsCount
    out.G_EnsureDataHeaderAgree "DataFrame.InferTypes"

    Set InferTypes = out
End Function

'==============================================================================
' [SEC] IO
' RD_BOOKMARK: DF: IO
'==============================================================================
'@Description "Restituisce l'array dati (valuta copia in futuro)"
Public Function AsArray() As Variant
    AsArray = mData
End Function

'@Description "Scrive su ListObject (replace body; opz. aggiorna header)"
Public Function WriteToListObject(ByVal lo As ListObject) As Long
    On Error GoTo EH

    If lo Is Nothing Then
        RaiseDf "DataFrame.WriteToListObject", edfInvalidData, "ListObject non valido (Nothing)."
    End If
    G_EnsureDataHeaderAgree "DataFrame.WriteToListObject"
    If lo.ListColumns.Count <> ColsCount Then
        RaiseDf "DataFrame.WriteToListObject", edfHeaderKeysMismatch, "Numero colonne ListObject non coerente con DataFrame."
    End If

    Dim i As Long
    For i = 1 To ColsCount
        lo.ListColumns(i).Name = CStr(mHeader(i))
    Next i

    Dim t0 As Double: t0 = Timer

    Dim oldCalc As XlCalculation
    Dim oldScrUpd As Boolean, oldEvents As Boolean
    oldCalc = Application.Calculation
    oldScrUpd = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim anchor As Range
    Set anchor = lo.Range.Cells(1, 1)

    lo.Resize anchor.Resize(RowsCount + 1, ColsCount)
    If RowsCount > 0 Then
        lo.DataBodyRange.Value2 = BuildDataOutArray()
    End If

    Dim dt As Double: dt = Timer - t0
    If dt < 0# Then dt = dt + 24# * 3600#
    mMetrics("write_ms") = CLng(dt * 1000#)
    mMetrics("write_addr") = lo.Range.Address(0, 0)
    mMetrics("write_rows") = RowsCount
    mMetrics("write_cols") = ColsCount

    WriteToListObject = RowsCount
    GoTo CLEANUP

EH:
    RaiseDf "DataFrame.WriteToListObject", edfRuntimeError, "Errore di scrittura su tabella: " & Err.Number & " - " & Err.Description

CLEANUP:
    Application.ScreenUpdating = oldScrUpd
    Application.EnableEvents = oldEvents
    Application.Calculation = oldCalc
End Function

'@Description "Scrive header + dati partendo SEMPRE da rng.Cells(1,1)."
Public Function WriteToRange( _
    ByVal rng As Range, _
    Optional ByVal includeHeader As Boolean = True _
) As DataFrame
    On Error GoTo EH

    If rng Is Nothing Then
        RaiseDf "DataFrame.WriteToRange", edfInvalidData, "Range destinazione non valido (Nothing)."
    End If
    G_EnsureDataHeaderAgree "DataFrame.WriteToRange"
    If ColsCount < 1 Then
        RaiseDf "DataFrame.WriteToRange", edfInvalidData, "Nessuna colonna da scrivere."
    End If

    Dim t0 As Double: t0 = Timer

    ' Switch performance
    Dim oldCalc As XlCalculation
    Dim oldScrUpd As Boolean, oldEvents As Boolean
    oldCalc = Application.Calculation
    oldScrUpd = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Ancorati sempre alla prima cella del range passato
    Dim anchor As Range
    Set anchor = rng.Cells(1, 1)

    ' Costruisci outArr: header opzionale + dati
    Dim rOut As Long, cOut As Long, i As Long, j As Long
    Dim outArr As Variant
    cOut = ColsCount

    If includeHeader Then
        rOut = RowsCount + 1
        ReDim outArr(1 To rOut, 1 To cOut)
        ' header
        For j = 1 To cOut
            outArr(1, j) = mHeader(j)
        Next j
        ' dati
        If RowsCount > 0 Then
            For i = 1 To RowsCount
                For j = 1 To cOut
                    If IsEmpty(mData(i, j)) Or mData(i, j) = vbNullString Then
                        outArr(i + 1, j) = mNullToken
                    Else
                        outArr(i + 1, j) = mData(i, j)
                    End If
                Next j
            Next i
        End If
    Else
        rOut = RowsCount
        If rOut = 0 Then GoTo CLEANUP
        ReDim outArr(1 To rOut, 1 To cOut)
        For i = 1 To rOut
            For j = 1 To cOut
                If IsEmpty(mData(i, j)) Or mData(i, j) = vbNullString Then
                    outArr(i, j) = mNullToken
                Else
                    outArr(i, j) = mData(i, j)
                End If
            Next j
        Next i
    End If

    ' Scrittura in blocco a partire da anchor
    anchor.Resize(rOut, cOut).Value2 = outArr

    ' metriche minime
    Dim dt As Double: dt = Timer - t0
    If dt < 0# Then dt = dt + 24# * 3600#
    mMetrics("write_ms") = CLng(dt * 1000#)
    mMetrics("write_addr") = anchor.Address(0, 0)
    mMetrics("write_rows") = rOut
    mMetrics("write_cols") = cOut

    Set WriteToRange = Me
    GoTo CLEANUP

EH:
    RaiseDf "DataFrame.WriteToRange", edfRuntimeError, "Errore di scrittura: " & Err.Number & " - " & Err.Description

CLEANUP:
    Application.ScreenUpdating = oldScrUpd
    Application.EnableEvents = oldEvents
    Application.Calculation = oldCalc
End Function


'@Description "Appende righe a ListObject"
Public Function AppendTo(ByVal lo As ListObject) As Long
    Dim base As New DataFrame
    Dim merged As DataFrame

    base.LoadFromListObject lo
    Set merged = base.Append(Me)
    AppendTo = merged.WriteToListObject(lo)
End Function

'@Description "Espone metriche come Variant (copia)"
Public Function Metrics() As Variant
    Dim out As Variant
    Dim n As Long, i As Long

    If mMetrics Is Nothing Or mMetrics.Count = 0 Then
        ReDim out(1 To 1, 1 To 2)
        out(1, 1) = "info"
        out(1, 2) = "no_metrics"
        Metrics = out
        Exit Function
    End If

    n = mMetrics.Count
    ReDim out(1 To n, 1 To 2)

    Dim keys As Variant
    keys = mMetrics.Keys
    For i = 1 To n
        out(i, 1) = CStr(keys(i - 1))
        out(i, 2) = mMetrics(keys(i - 1))
    Next i

    Metrics = out
End Function

'@Description "Abilita debug fluente e restituisce Me"
Public Function WithDebug(ByVal enabled As Boolean) As DataFrame
    mDebug = enabled
    Set WithDebug = Me
End Function

'==============================================================================
' [SEC] Diagnostics
' RD_BOOKMARK: DF: Diagnostics
'==============================================================================
' TODO[DF:Diag]: Chiavi standard metriche ? "rows","cols","read_ms","write_ms","strategy"

'==============================================================================
' [SEC] Helpers (Array/Header/Guards)
' RD_BOOKMARK: DF: Helpers
'==============================================================================

'@Description "Collassa spazi multipli in singoli (senza regex)" ------
Private Function CollapseSpaces(ByVal s As String) As String
    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    CollapseSpaces = s
End Function

'@Description "Materializza mData in array 2D sostituendo empty/nullstring con NullToken"
Private Function BuildDataOutArray() As Variant
    Dim out As Variant
    Dim i As Long, j As Long
    ReDim out(1 To RowsCount, 1 To ColsCount)
    For i = 1 To RowsCount
        For j = 1 To ColsCount
            If IsEmpty(mData(i, j)) Or mData(i, j) = vbNullString Then
                out(i, j) = mNullToken
            Else
                out(i, j) = mData(i, j)
            End If
        Next j
    Next i
    BuildDataOutArray = out
End Function

'@Description "Parsing direzioni sort: 'asc,desc' (default asc)"
Private Function ParseSortDirections(ByVal spec As String, ByVal nCols As Long) As Boolean()
    Dim dirs() As Boolean
    Dim i As Long
    ReDim dirs(1 To nCols)

    If Len(Trim$(spec)) = 0 Then
        For i = 1 To nCols: dirs(i) = True: Next i
        ParseSortDirections = dirs
        Exit Function
    End If

    Dim parts() As String
    parts = Split(spec, ",")

    If UBound(parts) - LBound(parts) + 1 = 1 Then
        Dim oneDir As Boolean
        oneDir = ParseOneDirection(CStr(parts(LBound(parts))), 1)
        For i = 1 To nCols: dirs(i) = oneDir: Next i
    ElseIf UBound(parts) - LBound(parts) + 1 = nCols Then
        For i = 1 To nCols
            dirs(i) = ParseOneDirection(CStr(parts(LBound(parts) + i - 1)), i)
        Next i
    Else
        RaiseDf "DataFrame.Sort", edfInvalidData, "Parametro 'ascending' deve avere 1 valore o uno per ogni colonna di sort."
    End If

    ParseSortDirections = dirs
End Function

Private Function ParseOneDirection(ByVal token As String, ByVal pos As Long) As Boolean
    Dim t As String
    t = LCase$(Trim$(token))
    Select Case t
        Case "", "asc", "ascending", "1", "true"
            ParseOneDirection = True
        Case "desc", "descending", "0", "false"
            ParseOneDirection = False
        Case Else
            RaiseDf "DataFrame.Sort", edfInvalidData, "Direzione sort non valida in posizione " & CStr(pos) & ": '" & token & "'."
    End Select
End Function

'@Description "Confronta due righe seguendo la specifica di sort"
Private Function CompareRowsBySpec(ByVal rowA As Long, ByVal rowB As Long, _
                                   ByRef idx() As Long, ByRef dirs() As Boolean) As Long
    Dim k As Long, col As Long, cmp As Long
    For k = LBound(idx) To UBound(idx)
        col = idx(k)
        cmp = CompareValues(mData(rowA, col), mData(rowB, col))
        If cmp <> 0 Then
            If dirs(k - LBound(idx) + 1) Then
                CompareRowsBySpec = cmp
            Else
                CompareRowsBySpec = -cmp
            End If
            Exit Function
        End If
    Next k
    CompareRowsBySpec = 0
End Function

'@Description "Confronto base Variant: Empty/Null minori; numeri e date per valore; fallback lessicografico"
Private Function CompareValues(ByVal a As Variant, ByVal b As Variant) As Long
    If IsNull(a) And IsNull(b) Then CompareValues = 0: Exit Function
    If IsNull(a) Then CompareValues = -1: Exit Function
    If IsNull(b) Then CompareValues = 1: Exit Function

    If IsEmpty(a) And IsEmpty(b) Then CompareValues = 0: Exit Function
    If IsEmpty(a) Then CompareValues = -1: Exit Function
    If IsEmpty(b) Then CompareValues = 1: Exit Function

    If IsDate(a) And IsDate(b) Then
        If CDbl(CDate(a)) < CDbl(CDate(b)) Then
            CompareValues = -1
        ElseIf CDbl(CDate(a)) > CDbl(CDate(b)) Then
            CompareValues = 1
        Else
            CompareValues = 0
        End If
        Exit Function
    End If

    If IsNumeric(a) And IsNumeric(b) Then
        If CDbl(a) < CDbl(b) Then
            CompareValues = -1
        ElseIf CDbl(a) > CDbl(b) Then
            CompareValues = 1
        Else
            CompareValues = 0
        End If
        Exit Function
    End If

    Dim sa As String, sb As String
    sa = CStr(a): sb = CStr(b)
    If StrComp(sa, sb, vbTextCompare) < 0 Then
        CompareValues = -1
    ElseIf StrComp(sa, sb, vbTextCompare) > 0 Then
        CompareValues = 1
    Else
        CompareValues = 0
    End If
End Function





'@Description "Rende unici i nomi header (case-insensitive) aggiungendo suffissi _2, _3, ..."
Private Function Hdr_EnsureUnique1D(ByRef hdr As Variant) As Variant
    Dim out As Variant: out = hdr
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long, key As String, base As String, k As Long
    
    For i = LBound(out) To UBound(out)
        base = CStr(out(i))
        key = LCase$(base)
        If Not seen.Exists(key) Then
            seen.Add key, 1
        Else
            k = seen(key) + 1
            seen(key) = k
            out(i) = base & "_" & CStr(k)
        End If
    Next i
    Hdr_EnsureUnique1D = out
End Function


'=== Helpers: Tipi/forme degli array ==========================================

'==================== Helpers privati per LoadFromRange ====================

'@Description "Risoluzione orientamento header: 'row'|'col'|'auto' ? 'row' di default salvo 1xN/Nx1"
Private Function ResolveHeaderAt(ByVal headerAt As DFHeaderAt, ByVal rows As Long, ByVal cols As Long) As String
    Select Case headerAt
        Case dfRow
            ResolveHeaderAt = "row"
        Case dfColumn
            ResolveHeaderAt = "col"
        Case Else
            ResolveHeaderAt = "row"
    End Select
End Function

'@Description "True se la cella è logicamente vuota (Empty o stringa vuota/spazi)"
Private Function IsEmptyLike(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then
        IsEmptyLike = True
    ElseIf VarType(v) = vbString Then
        IsEmptyLike = (LenB(Trim$(CStr(v))) = 0)
    Else
        IsEmptyLike = False
    End If
End Function

'@Description "Calcola taglio bordi esterni completamente vuoti; restituisce tagli top/bottom/left/right"
Private Sub DetectOuterEmptyCuts(ByRef a As Variant, _
                                 ByRef cutTop As Long, ByRef cutBottom As Long, _
                                 ByRef cutLeft As Long, ByRef cutRight As Long)
    Dim rLo As Long, rHi As Long, cLo As Long, cHi As Long
    rLo = LBound(a, 1): rHi = UBound(a, 1)
    cLo = LBound(a, 2): cHi = UBound(a, 2)
    Dim r As Long, c As Long: r = rHi - rLo + 1: c = cHi - cLo + 1

    Dim rowHas() As Boolean: ReDim rowHas(1 To r)
    Dim colHas() As Boolean: ReDim colHas(1 To c)

    Dim i As Long, j As Long
    For i = 1 To r
        For j = 1 To c
            If Not IsEmptyLike(a(rLo + i - 1, cLo + j - 1)) Then
                rowHas(i) = True
                colHas(j) = True
            End If
        Next j
    Next i

    ' Top
    cutTop = 0
    For i = 1 To r
        If rowHas(i) Then Exit For
        cutTop = cutTop + 1
    Next i
    ' Bottom
    cutBottom = 0
    For i = r To 1 Step -1
        If rowHas(i) Then Exit For
        cutBottom = cutBottom + 1
    Next i
    ' Left
    cutLeft = 0
    For j = 1 To c
        If colHas(j) Then Exit For
        cutLeft = cutLeft + 1
    Next j
    ' Right
    cutRight = 0
    For j = c To 1 Step -1
        If colHas(j) Then Exit For
        cutRight = cutRight + 1
    Next j

    ' Se tutto vuoto, evita di tagliare tutto (lascia zero tagli)
    If cutTop = r Or cutLeft = c Then
        cutTop = 0: cutBottom = 0: cutLeft = 0: cutRight = 0
    End If
End Sub

'@Description "Trova ultima riga non vuota scansionando dal basso; se skipHeader=True salta la prima riga"
Private Function FindLastNonEmptyRow(ByRef a As Variant, ByVal skipHeader As Boolean) As Long
    Dim rLo As Long, rHi As Long, cLo As Long, cHi As Long
    rLo = LBound(a, 1): rHi = UBound(a, 1)
    cLo = LBound(a, 2): cHi = UBound(a, 2)
    Dim startRow As Long: startRow = rLo + IIf(skipHeader, 1, 0)

    Dim i As Long, j As Long
    For i = rHi To startRow Step -1
        For j = cLo To cHi
            If Not IsEmptyLike(a(i, j)) Then
                FindLastNonEmptyRow = (i - rLo + 1) ' indice relativo 1-based nel blocco
                Exit Function
            End If
        Next j
    Next i
    FindLastNonEmptyRow = 0
End Function

Private Sub ParseCondition(ByVal condition As String, ByRef colIdx As Long, ByRef op As String, ByRef rhs As String)
    Dim txt As String
    txt = Trim$(condition)
    If Len(txt) = 0 Then
        RaiseDf "DataFrame.Filter", edfInvalidData, "Condition vuota."
    End If

    Dim p As Long, lhs As String

    p = InStr(1, LCase$(txt), " contains ", vbTextCompare)
    If p > 0 Then
        lhs = Trim$(Left$(txt, p - 1))
        rhs = Trim$(Mid$(txt, p + Len(" contains ")))
        op = "contains"
    Else
        p = InStr(1, txt, "<>", vbBinaryCompare)
        If p > 0 Then
            lhs = Trim$(Left$(txt, p - 1))
            rhs = Trim$(Mid$(txt, p + 2))
            op = "<>"
        Else
            p = InStr(1, txt, "=", vbBinaryCompare)
            If p > 0 Then
                lhs = Trim$(Left$(txt, p - 1))
                rhs = Trim$(Mid$(txt, p + 1))
                op = "="
            Else
                p = InStr(1, txt, ">", vbBinaryCompare)
                If p > 0 Then
                    lhs = Trim$(Left$(txt, p - 1))
                    rhs = Trim$(Mid$(txt, p + 1))
                    op = ">"
                Else
                    p = InStr(1, txt, "<", vbBinaryCompare)
                    If p > 0 Then
                        lhs = Trim$(Left$(txt, p - 1))
                        rhs = Trim$(Mid$(txt, p + 1))
                        op = "<"
                    End If
                End If
            End If
        End If
    End If

    If Len(op) = 0 Then
        RaiseDf "DataFrame.Filter", edfInvalidData, "Operatore non supportato nella condition: " & condition
    End If

    lhs = TrimQuotes(lhs)
    rhs = TrimQuotes(rhs)
    colIdx = Hdr_FindIndexByName(lhs, mHeader, mAliasMap)
End Sub

Private Function TrimQuotes(ByVal s As String) As String
    s = Trim$(s)
    If Len(s) >= 2 Then
        If (Left$(s, 1) = """ And Right$(s, 1) = """) Or (Left$(s, 1) = "'" And Right$(s, 1) = "'") Then
            s = Mid$(s, 2, Len(s) - 2)
        End If
    End If
    TrimQuotes = s
End Function

Private Function EvalCondition(ByVal lhs As Variant, ByVal op As String, ByVal rhs As String) As Boolean
    Dim ltxt As String
    ltxt = CStr(lhs)

    Select Case LCase$(op)
        Case "contains"
            EvalCondition = (InStr(1, LCase$(ltxt), LCase$(rhs), vbTextCompare) > 0)
        Case "="
            EvalCondition = (CompareVariant(lhs, ParseLiteral(rhs)) = 0)
        Case "<>"
            EvalCondition = (CompareVariant(lhs, ParseLiteral(rhs)) <> 0)
        Case ">"
            EvalCondition = (CompareVariant(lhs, ParseLiteral(rhs)) > 0)
        Case "<"
            EvalCondition = (CompareVariant(lhs, ParseLiteral(rhs)) < 0)
        Case Else
            RaiseDf "DataFrame.Filter", edfInvalidData, "Operatore non supportato: " & op
    End Select
End Function

Private Function ParseLiteral(ByVal s As String) As Variant
    Dim t As String
    t = TrimQuotes(Trim$(s))

    If Len(t) = 0 Then
        ParseLiteral = vbNullString
    ElseIf LCase$(t) = "null" Then
        ParseLiteral = Null
    ElseIf IsDate(t) Then
        ParseLiteral = CDate(t)
    ElseIf IsNumeric(t) Then
        ParseLiteral = CDbl(t)
    Else
        ParseLiteral = t
    End If
End Function

Private Function CompareRowsByCols(ByVal rowA As Long, ByVal rowB As Long, ByRef idx() As Long, ByRef asc() As Boolean) As Long
    Dim k As Long, c As Long, cmp As Long
    For k = LBound(idx) To UBound(idx)
        c = idx(k)
        cmp = CompareVariant(mData(rowA, c), mData(rowB, c))
        If cmp <> 0 Then
            If asc(k - LBound(idx) + 1) Then
                CompareRowsByCols = cmp
            Else
                CompareRowsByCols = -cmp
            End If
            Exit Function
        End If
    Next k
    CompareRowsByCols = 0
End Function

Private Function CompareVariant(ByVal a As Variant, ByVal b As Variant) As Long
    On Error GoTo FallbackText

    If IsNull(a) And IsNull(b) Then
        CompareVariant = 0
        Exit Function
    ElseIf IsNull(a) Then
        CompareVariant = -1
        Exit Function
    ElseIf IsNull(b) Then
        CompareVariant = 1
        Exit Function
    End If

    If IsEmptyLike(a) And IsEmptyLike(b) Then
        CompareVariant = 0
    ElseIf IsEmptyLike(a) Then
        CompareVariant = -1
    ElseIf IsEmptyLike(b) Then
        CompareVariant = 1
    ElseIf IsDate(a) And IsDate(b) Then
        If CDate(a) < CDate(b) Then
            CompareVariant = -1
        ElseIf CDate(a) > CDate(b) Then
            CompareVariant = 1
        Else
            CompareVariant = 0
        End If
    ElseIf IsNumeric(a) And IsNumeric(b) Then
        If CDbl(a) < CDbl(b) Then
            CompareVariant = -1
        ElseIf CDbl(a) > CDbl(b) Then
            CompareVariant = 1
        Else
            CompareVariant = 0
        End If
    Else
        GoTo FallbackText
    End If
    Exit Function

FallbackText:
    Dim sa As String, sb As String
    sa = LCase$(CStr(a)): sb = LCase$(CStr(b))
    If sa < sb Then
        CompareVariant = -1
    ElseIf sa > sb Then
        CompareVariant = 1
    Else
        CompareVariant = 0
    End If
End Function

Private Function BuildRowKey(ByRef data As Variant, ByVal r As Long, ByRef keyIdx() As Long) As String
    Dim i As Long, c As Long
    Dim sb As String
    sb = vbNullString
    For i = LBound(keyIdx) To UBound(keyIdx)
        c = keyIdx(i)
        sb = sb & "|" & CanonicalValue(data(r, c))
    Next i
    BuildRowKey = sb
End Function

Private Function CanonicalValue(ByVal v As Variant) As String
    If IsNull(v) Then
        CanonicalValue = "<NULL>"
    ElseIf IsEmptyLike(v) Then
        CanonicalValue = "<EMPTY>"
    ElseIf IsDate(v) Then
        CanonicalValue = "<DATE>" & Format$(CDate(v), "yyyy-mm-dd hh:nn:ss")
    ElseIf IsNumeric(v) Then
        CanonicalValue = "<NUM>" & Replace$(CStr(CDbl(v)), ",", ".")
    Else
        CanonicalValue = "<TXT>" & LCase$(Trim$(CStr(v)))
    End If
End Function

Private Function BuildNonKeyIndex(ByVal totalCols As Long, ByRef keyIdx() As Long) As Long()
    Dim isKey() As Boolean
    ReDim isKey(1 To totalCols)

    Dim i As Long
    For i = LBound(keyIdx) To UBound(keyIdx)
        isKey(keyIdx(i)) = True
    Next i

    Dim cnt As Long
    For i = 1 To totalCols
        If Not isKey(i) Then cnt = cnt + 1
    Next i

    Dim out() As Long
    If cnt = 0 Then
        BuildNonKeyIndex = out
        Exit Function
    End If

    ReDim out(1 To cnt)
    cnt = 0
    For i = 1 To totalCols
        If Not isKey(i) Then
            cnt = cnt + 1
            out(cnt) = i
        End If
    Next i

    BuildNonKeyIndex = out
End Function

Private Function HeaderContains(ByRef hdr As Variant, ByVal name As String, ByVal maxIdx As Long) As Boolean
    Dim i As Long
    For i = 1 To maxIdx
        If LCase$(CStr(hdr(i))) = LCase$(name) Then
            HeaderContains = True
            Exit Function
        End If
    Next i
End Function

Private Function ParseSuffixRight(ByVal suffixes As String) As String
    Dim parts() As String
    parts = Split(CStr(suffixes), ",")
    If UBound(parts) >= 1 Then
        ParseSuffixRight = Trim$(parts(1))
    Else
        ParseSuffixRight = "_R"
    End If
    If Len(ParseSuffixRight) = 0 Then ParseSuffixRight = "_R"
End Function

Private Function CleanValue(ByVal v As Variant, ByVal doTrim As Boolean, ByVal doCollapse As Boolean, ByVal doCoerce As Boolean) As Variant
    If IsNull(v) Then
        CleanValue = Empty
        Exit Function
    End If

    If VarType(v) = vbString Then
        Dim s As String
        s = CStr(v)
        If doTrim Then s = Trim$(s)
        If doCollapse Then s = CollapseSpaces(s)

        Select Case LCase$(s)
            Case "", "na", "n/a", "null", "-", "none"
                CleanValue = Empty
                Exit Function
        End Select

        If doCoerce Then
            If IsDate(s) Then
                CleanValue = CDate(s)
                Exit Function
            ElseIf IsNumeric(s) Then
                CleanValue = CDbl(s)
                Exit Function
            End If
        End If
        CleanValue = s
    Else
        CleanValue = v
    End If
End Function

'@Description "Header da prima riga (1D 1-based)"
Private Function HeaderFromRow(ByRef a As Variant, ByVal r As Long, ByVal cLo As Long, ByVal cHi As Long) As Variant
    Dim n As Long: n = cHi - cLo + 1
    Dim out As Variant: ReDim out(1 To n)
    Dim j As Long
    For j = 1 To n
        out(j) = a(r, cLo + j - 1)
    Next j
    HeaderFromRow = out
End Function

'@Description "Header da prima colonna (1D 1-based)"
Private Function HeaderFromCol(ByRef a As Variant, ByVal c As Long, ByVal rLo As Long, ByVal rHi As Long) As Variant
    Dim n As Long: n = rHi - rLo + 1
    Dim out As Variant: ReDim out(1 To n)
    Dim i As Long
    For i = 1 To n
        out(i) = a(rLo + i - 1, c)
    Next i
    HeaderFromCol = out
End Function

'@Description "Genera header C1..Cn"
Private Function GenHeaderCols(ByVal n As Long) As Variant
    Dim out As Variant: ReDim out(1 To n)
    Dim i As Long
    For i = 1 To n
        out(i) = "C" & CStr(i)
    Next i
    GenHeaderCols = out
End Function


'@Description "True se v è un array 2D reale (ha la 2^ dimensione)"
Private Function Arr_Is2D(ByRef v As Variant) As Boolean
    On Error Resume Next
    Dim ub2 As Long: ub2 = UBound(v, 2)
    Arr_Is2D = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

'@Description "Converte un Array-of-Array (jagged) in un array 2D 1-based; richiede righe equi-lunghe"
Private Function Arr_JaggedTo2D1(ByRef jag As Variant) As Variant
    Dim r0 As Long, r1 As Long, c0 As Long, c1 As Long
    Dim rows As Long, cols As Long
    Dim i As Long, j As Long
    Dim first As Variant, rowv As Variant
    If Not IsArray(jag) Then RaiseDf "DataFrame.Arr_JaggedTo2D1", edfInvalidData, "Input non è un array."

    r0 = LBound(jag): r1 = UBound(jag)
    first = jag(r0)
    If Not IsArray(first) Then RaiseDf "DataFrame.Arr_JaggedTo2D1", edfInvalidData, "Non è un array di array."
    c0 = LBound(first): c1 = UBound(first)

    rows = r1 - r0 + 1
    cols = c1 - c0 + 1

    Dim out As Variant
    ReDim out(1 To rows, 1 To cols)

    For i = 1 To rows
        rowv = jag(r0 + i - 1)
        ' (opz) verifica lunghezze allineate
        If Not IsArray(rowv) Then RaiseDf "DataFrame.Arr_JaggedTo2D1", edfJaggedRowsMismatch, "Riga non è un array."
        If LBound(rowv) <> c0 Or UBound(rowv) <> c1 Then _
            RaiseDf "DataFrame.Arr_JaggedTo2D1", edfJaggedRowsMismatch, "Lunghezze righe non uniformi."
        For j = 1 To cols
            out(i, j) = rowv(c0 + j - 1)
        Next j
    Next i
    Arr_JaggedTo2D1 = out
End Function

'@Description "Normalizza un array 1D (0-based o 1-based) a 1-based"
Private Function Arr_To1D1(ByRef v As Variant) As Variant
    If Not IsArray(v) Then RaiseDf "DataFrame.Arr_To1D1", edfInvalidHeader, "Input non è un array."
    Dim lo As Long, hi As Long, n As Long, i As Long
    lo = LBound(v): hi = UBound(v): n = hi - lo + 1
    Dim out As Variant: ReDim out(1 To n)
    For i = 1 To n
        out(i) = v(lo + i - 1)
    Next i
    Arr_To1D1 = out
End Function

'@Description "Riduce un header 2D ?lineare? a 1D 1-based (1×N o N×1)"
Private Function Hdr_2DLinearTo1D(ByRef hdr As Variant) As Variant
    If Not IsArray(hdr) Then RaiseDf "DataFrame.Hdr_2DLinearTo1D", edfInvalidHeader, "Input non è un array."
    Dim rLo As Long, rHi As Long, cLo As Long, cHi As Long
    Dim n As Long, i As Long
    rLo = LBound(hdr, 1): rHi = UBound(hdr, 1)
    cLo = LBound(hdr, 2): cHi = UBound(hdr, 2)

    If rHi - rLo = 0 Then
        n = cHi - cLo + 1
        Dim outRow As Variant: ReDim outRow(1 To n)
        For i = 1 To n
            outRow(i) = hdr(rLo, cLo + i - 1)
        Next i
        Hdr_2DLinearTo1D = outRow
    ElseIf cHi - cLo = 0 Then
        n = rHi - rLo + 1
        Dim outCol As Variant: ReDim outCol(1 To n)
        For i = 1 To n
            outCol(i) = hdr(rLo + i - 1, cLo)
        Next i
        Hdr_2DLinearTo1D = outCol
    Else
        RaiseDf "DataFrame.Hdr_2DLinearTo1D", edfNonLinearHeader, "Header 2D non lineare (attese 1 riga o 1 colonna)."
    End If
End Function

'@Description "Dallo spec generico (stringa "A,B" o array misto) ricava indici 1-based"
Private Function Hdr_ResolveIndicesFromSpec(ByVal spec As String, _
                                            ByRef hdr As Variant, _
                                            ByVal aliasMap As Object) As Long()
    Dim parts() As String, i As Long
    spec = Trim$(spec)
    If Len(spec) = 0 Then GoTo EmptySpec
    parts = Split(spec, ",")
    Hdr_ResolveIndicesFromSpec = Hdr_ResolveIndicesFromArray(parts, hdr, aliasMap)
    Exit Function
EmptySpec:
    Dim emptyArr() As Long
    Hdr_ResolveIndicesFromSpec = emptyArr
End Function

'@Description "Risolve un array di nomi/indici in indici 1-based (alias opz.)"
Private Function Hdr_ResolveIndicesFromArray(ByVal specArr As Variant, _
                                             ByRef hdr As Variant, _
                                             ByVal aliasMap As Object) As Long()
    Dim out() As Long, n As Long, i As Long
    n = IIf(IsArray(specArr), UBound(specArr) - LBound(specArr) + 1, 0)
    ReDim out(1 To n)
    For i = 1 To n
        Dim v As Variant: v = specArr(LBound(specArr) + i - 1)
        If IsNumeric(v) Then
            out(i) = CLng(v)
        Else
            out(i) = Hdr_FindIndexByName(CStr(v), hdr, aliasMap)
        End If
    Next i
    Hdr_ResolveIndicesFromArray = out
End Function

'@Description "Trova indice di colonna per nome (case-insensitive, alias opz.)"
Private Function Hdr_FindIndexByName(ByVal nameSpec As String, _
                                     ByRef hdr As Variant, _
                                     ByVal aliasMap As Object) As Long
    Dim target As String: target = Trim$(LCase$(nameSpec))

    ' Alias: se presente, risolvi a nome canonico
    If Not aliasMap Is Nothing Then
        If aliasMap.Exists(target) Then target = LCase$(CStr(aliasMap(target)))
    End If

    Dim i As Long
    For i = LBound(hdr) To UBound(hdr)
        If LCase$(CStr(hdr(i))) = target Then
            Hdr_FindIndexByName = i
            Exit Function
        End If
    Next i
    ' Non trovato ? errore
    RaiseDf "DataFrame.Hdr_FindIndexByName", edfHeaderKeysMismatch, _
            "Colonna non trovata: '" & nameSpec & "'"
End Function

'@Description "Verifica che tutti gli indici siano 1..cols e non duplicati"
Private Function Arr_IsValidIndices(ByRef idx() As Long, ByVal cols As Long) As Boolean
    On Error GoTo Fail
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(idx) To UBound(idx)
        If idx(i) < 1 Or idx(i) > cols Then GoTo Fail
        If seen.Exists(CStr(idx(i))) Then GoTo Fail
        seen.Add CStr(idx(i)), True
    Next i
    Arr_IsValidIndices = True
    Exit Function
Fail:
    Arr_IsValidIndices = False
End Function

'=== Guards: invarianti ========================================================

'@Description "Verifica coerenza mData/mHeader (cols uguali, 1-based)"
Private Sub G_EnsureDataHeaderAgree(ByVal src As String)
    If Not IsArray(mData) Then Exit Sub
    If Not IsArray(mHeader) Then Exit Sub
    Dim cols As Long: cols = UBound(mData, 2) - LBound(mData, 2) + 1
    If cols <> (UBound(mHeader) - LBound(mHeader) + 1) Then _
        RaiseDf src, edfHeaderKeysMismatch, _
                "Numero colonne di dati e header non coerenti."
End Sub

'=== Fine Helpers ==============================================================


Private Function NewFrameFromParts(ByRef data As Variant, ByRef header As Variant) As DataFrame
    Dim out As New DataFrame
    out.LoadFromArray data, header
    out.NullToken = mNullToken
    out.DebugMode = mDebug
    Set NewFrameFromParts = out
End Function

Private Function HeaderIndexMap(ByRef hdr As Variant) As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")
    Dim i As Long, k As String
    For i = LBound(hdr) To UBound(hdr)
        k = LCase$(CStr(hdr(i)))
        If Not m.Exists(k) Then m(k) = i
    Next i
    Set HeaderIndexMap = m
End Function

Private Function HeaderIndexOrZero(ByVal map As Object, ByVal nameSpec As String) As Long
    Dim k As String
    k = LCase$(CStr(nameSpec))
    If map.Exists(k) Then
        HeaderIndexOrZero = CLng(map(k))
    Else
        HeaderIndexOrZero = 0
    End If
End Function


