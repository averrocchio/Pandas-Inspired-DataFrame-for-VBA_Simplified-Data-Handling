VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Private mMetrics As Object          ' Dictionary pAer metriche
Private mDebug As Boolean

' ====== Errori & utilità (privati) ==========================================
Private Enum DFErr
    edfInvalidData = 11010
    edfInvalidHeader = 11020
    edfNonLinearHeader = 11030
    edfJaggedRowsMismatch = 11040
    edfInvalidKeysSpec = 11050
    edfHeaderKeysMismatch = 11060
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
    If IsArray(mData) Then RowsCount = UBound(mData, 1) - LBound(mData, 1) + 1
End Property

'@Description "Restituisce il numero di colonne (0 se mData non è inizializzato)"
Public Property Get ColsCount() As Long
    If IsArray(mData) Then ColsCount = UBound(mData, 2) - LBound(mData, 2) + 1
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
    Keys = mKeys  ' restituisce riferimento (accettabile per ora)
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
Attribute LoadFromArray.VB_Description = "Crea/riusa l'istanza caricando dati da array + header"
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

                              Optional ByVal headerAt As String = "row", _

                              Optional ByRef bounds As DFBounds = dfLiteral, _

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

        Case dfLiteral, dfTrimOuterEmpty

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

Enum DFBounds
    dfLiteral = 1
    dfCurrentRegion = 2
    dfTrimOuterEmpty = 3
    dfSmart = 4
End Enum

'==============================================================================
' [SEC] Core Ops
' RD_BOOKMARK: DF: Ops Core
'==============================================================================
'@Description "Filtra righe in base a DSL condizionale"
'  TODO[DF:Ops]: valutare condition riga-per-riga, preservare header
Public Function Filter(ByVal condition As String) As DataFrame
    Dim out As New DataFrame
    ' TODO: out.mData = ... ; out.mHeader = mHeader
    Set Filter = out
End Function

'@Description "Ordinamento stabile multi-colonna"
'  TODO[DF:Ops]: stable sort con specifica colonne e direzioni
Public Function Sort(ByVal cols As String, Optional ByVal ascending As String = "") As DataFrame
    Dim out As New DataFrame
    ' TODO
    Set Sort = out
End Function

'@Description "Dedup su chiavi (policy: keep_first/keep_last/...)"
'  TODO[DF:Ops]: usare mKeys; politiche duplicati
Public Function Dedup(Optional ByVal policy As String = "keep_first") As DataFrame
    Dim out As New DataFrame
    ' TODO
    Set Dedup = out
End Function

'@Description "Proiezione/riordino colonne"
Public Function Project(ByVal cols As String) As DataFrame
    Dim out As New DataFrame
    ' TODO
    Set Project = out
End Function

'@Description "Rinomina colonne in base a mapping"
Public Function Rename(ByVal map As Variant) As DataFrame
    Dim out As New DataFrame
    ' TODO
    Set Rename = out
End Function

'@Description "Append di righe allineando per nome"
Public Function Append(ByVal rows As DataFrame) As DataFrame
    Dim out As New DataFrame
    ' TODO
    Set Append = out
End Function

'@Description "Join con destra, specifica chiavi e tipo"
Public Function JoinRight(ByVal right As DataFrame, ByVal Keys As String, _
                          Optional ByVal how As String = "inner", _
                          Optional ByVal suffixes As String = "_L,_R") As DataFrame
    Dim out As New DataFrame
    ' TODO
    Set JoinRight = out
End Function

'@Description "Pulizia base (trim/collapse/coercioni leggere)"
Public Function Clean(Optional ByVal trimSpaces As Boolean = True, _
                      Optional ByVal CollapseSpaces As Boolean = True, _
                      Optional ByVal softCoerce As Boolean = True) As DataFrame
    Dim out As New DataFrame
    ' TODO
    Set Clean = out
End Function

'@Description "Inferenza soft dei tipi per colonna"
Public Function InferTypes() As DataFrame
    Dim out As New DataFrame
    ' TODO
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
'  TODO[DF:IO]: full replace DataBodyRange con mData, gestione header
Public Function WriteTo(ByVal lo As ListObject) As Long
    ' TODO: ritorna righe scritte
End Function

'@Description "Appende righe a ListObject"
'  TODO[DF:IO]: append in coda
Public Function AppendTo(ByVal lo As ListObject) As Long
    ' TODO
End Function

'@Description "Espone metriche come Variant (copia)"
'  TODO[DF:IO]: decidere formato (array 2D, dizionario serializzato, ecc.)
Public Function Metrics() As Variant
    ' TODO
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
Private Function ResolveHeaderAt(ByVal headerAt As String, ByVal rows As Long, ByVal cols As Long) As String
    Dim ha As String: ha = LCase$(Trim$(headerAt))
    Select Case ha
        Case "row", "col"
            ResolveHeaderAt = ha
        Case "auto"
            If rows = 1 And cols >= 1 Then
                ResolveHeaderAt = "row"
            ElseIf cols = 1 And rows >= 1 Then
                ResolveHeaderAt = "col"
            Else
                ResolveHeaderAt = "row" ' convenzione
            End If
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

'@Description "Riduce un header 2D lineare a 1D 1-based (1×N o N×1)"
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


