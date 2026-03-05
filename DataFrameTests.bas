Attribute VB_Name = "DataFrameTests"
Option Explicit

Private Sub AssertEquals(ByVal testName As String, ByVal actual As Variant, ByVal expected As Variant)
    If CStr(actual) <> CStr(expected) Then
        Err.Raise vbObjectError + 21001, testName, "Expected=" & CStr(expected) & " Actual=" & CStr(actual)
    End If
End Sub

Private Sub AssertTrue(ByVal testName As String, ByVal cond As Boolean, ByVal msg As String)
    If Not cond Then
        Err.Raise vbObjectError + 21002, testName, msg
    End If
End Sub

Private Sub PrintPass(ByVal testName As String)
    Debug.Print "PASS - " & testName
End Sub

Private Sub PrintFail(ByVal testName As String, ByVal errMsg As String)
    Debug.Print "FAIL - " & testName & " - " & errMsg
End Sub

Public Sub Test_LoadFromArray_Basic()
    Const T As String = "Test_LoadFromArray_Basic"
    On Error GoTo EH

    Dim data(1 To 2, 1 To 2) As Variant
    Dim hdr(1 To 2) As Variant
    data(1, 1) = 1: data(1, 2) = "A"
    data(2, 1) = 2: data(2, 2) = "B"
    hdr(1) = "id": hdr(2) = "name"

    Dim df As New DataFrame
    df.LoadFromArray data, hdr

    AssertEquals T, df.RowsCount, 2
    AssertEquals T, df.ColsCount, 2
    AssertEquals T, df.AsArray()(2, 2), "B"

    PrintPass T
    Exit Sub
EH:
    PrintFail T, Err.Description
End Sub

Public Sub Test_LoadFromRange_WithHeader()
    Const T As String = "Test_LoadFromRange_WithHeader"
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)

    ws.Range("AA1:AB3").ClearContents
    ws.Range("AA1").Value2 = "id": ws.Range("AB1").Value2 = "name"
    ws.Range("AA2").Value2 = 1: ws.Range("AB2").Value2 = "Anna"
    ws.Range("AA3").Value2 = 2: ws.Range("AB3").Value2 = "Bruno"

    Dim df As New DataFrame
    df.LoadFromRange ws.Range("AA1:AB3"), True, dfRow, dfliteral

    AssertEquals T, df.RowsCount, 2
    AssertEquals T, df.ColsCount, 2
    AssertEquals T, df.header()(1), "id"
    AssertEquals T, df.AsArray()(2, 2), "Bruno"

    PrintPass T
    Exit Sub
EH:
    PrintFail T, Err.Description
End Sub

Public Sub Test_Project_And_Rename()
    Const T As String = "Test_Project_And_Rename"
    On Error GoTo EH

    Dim data(1 To 3, 1 To 3) As Variant
    Dim hdr(1 To 3) As Variant
    data(1, 1) = 1: data(1, 2) = "Anna": data(1, 3) = "IT"
    data(2, 1) = 2: data(2, 2) = "Bruno": data(2, 3) = "HR"
    data(3, 1) = 3: data(3, 2) = "Carla": data(3, 3) = "IT"
    hdr(1) = "id": hdr(2) = "name": hdr(3) = "dept"

    Dim df As New DataFrame
    Dim projected As DataFrame
    Dim renamed As DataFrame

    df.LoadFromArray data, hdr
    Set projected = df.Project("dept,name")
    Set renamed = projected.Rename("dept:team,name:full_name")

    AssertEquals T, renamed.ColsCount, 2
    AssertEquals T, renamed.header()(1), "team"
    AssertEquals T, renamed.header()(2), "full_name"
    AssertEquals T, renamed.AsArray()(2, 1), "HR"

    PrintPass T
    Exit Sub
EH:
    PrintFail T, Err.Description
End Sub

Public Sub Test_Append_SchemaMismatch_ShouldFail()
    Const T As String = "Test_Append_SchemaMismatch_ShouldFail"
    On Error GoTo EH

    Dim dataA(1 To 1, 1 To 2) As Variant
    Dim hdrA(1 To 2) As Variant
    dataA(1, 1) = 1: dataA(1, 2) = "A"
    hdrA(1) = "id": hdrA(2) = "name"

    Dim dataB(1 To 1, 1 To 2) As Variant
    Dim hdrB(1 To 2) As Variant
    dataB(1, 1) = "A": dataB(1, 2) = 1
    hdrB(1) = "name": hdrB(2) = "code"

    Dim leftDf As New DataFrame
    Dim rightDf As New DataFrame

    leftDf.LoadFromArray dataA, hdrA
    rightDf.LoadFromArray dataB, hdrB

    leftDf.Append rightDf

    PrintFail T, "Expected failure for schema mismatch but operation succeeded"
    Exit Sub
EH:
    If InStr(1, Err.Description, "Colonna non trovata", vbTextCompare) > 0 _
        Or InStr(1, Err.Description, "non coerente", vbTextCompare) > 0 Then
        PrintPass T
    Else
        PrintFail T, Err.Description
    End If
End Sub

Public Sub Test_Filter_Contains()
    Const T As String = "Test_Filter_Contains"
    On Error GoTo EH

    Dim data(1 To 4, 1 To 2) As Variant
    Dim hdr(1 To 2) As Variant
    data(1, 1) = 1: data(1, 2) = "Rome"
    data(2, 1) = 2: data(2, 2) = "Milan"
    data(3, 1) = 3: data(3, 2) = "Roma Nord"
    data(4, 1) = 4: data(4, 2) = "Turin"
    hdr(1) = "id": hdr(2) = "city"

    Dim df As New DataFrame
    Dim out As DataFrame
    df.LoadFromArray data, hdr

    Set out = df.Filter("city contains rom")

    AssertEquals T, out.RowsCount, 2
    AssertEquals T, out.AsArray()(1, 2), "Rome"
    AssertEquals T, out.AsArray()(2, 2), "Roma Nord"

    PrintPass T
    Exit Sub
EH:
    PrintFail T, Err.Description
End Sub

Public Sub Test_Dedup_ByKeys()
    Const T As String = "Test_Dedup_ByKeys"
    On Error GoTo EH

    Dim data(1 To 5, 1 To 3) As Variant
    Dim hdr(1 To 3) As Variant
    data(1, 1) = 1: data(1, 2) = "A": data(1, 3) = "x"
    data(2, 1) = 2: data(2, 2) = "A": data(2, 3) = "y"
    data(3, 1) = 3: data(3, 2) = "B": data(3, 3) = "z"
    data(4, 1) = 4: data(4, 2) = "B": data(4, 3) = "k"
    data(5, 1) = 5: data(5, 2) = "C": data(5, 3) = "m"
    hdr(1) = "id": hdr(2) = "grp": hdr(3) = "val"

    Dim df As New DataFrame
    Dim out As DataFrame
    df.LoadFromArray data, hdr
    df.Keys = "grp"

    Set out = df.Dedup("keep_first")

    AssertEquals T, out.RowsCount, 3
    AssertEquals T, out.AsArray()(1, 2), "A"
    AssertEquals T, out.AsArray()(2, 2), "B"
    AssertEquals T, out.AsArray()(3, 2), "C"

    PrintPass T
    Exit Sub
EH:
    PrintFail T, Err.Description
End Sub

Public Sub Test_Sort_MultiColumn()
    Const T As String = "Test_Sort_MultiColumn"
    On Error GoTo EH

    Dim data(1 To 4, 1 To 3) As Variant
    Dim hdr(1 To 3) As Variant
    data(1, 1) = 1: data(1, 2) = "B": data(1, 3) = 20
    data(2, 1) = 2: data(2, 2) = "A": data(2, 3) = 20
    data(3, 1) = 3: data(3, 2) = "A": data(3, 3) = 10
    data(4, 1) = 4: data(4, 2) = "B": data(4, 3) = 10
    hdr(1) = "id": hdr(2) = "grp": hdr(3) = "score"

    Dim df As New DataFrame
    Dim out As DataFrame
    df.LoadFromArray data, hdr

    Set out = df.Sort("grp,score", "asc,desc")

    AssertEquals T, out.AsArray()(1, 1), 2
    AssertEquals T, out.AsArray()(2, 1), 3
    AssertEquals T, out.AsArray()(3, 1), 1
    AssertEquals T, out.AsArray()(4, 1), 4

    PrintPass T
    Exit Sub
EH:
    PrintFail T, Err.Description
End Sub

Public Sub Test_Clean_And_InferTypes()
    Const T As String = "Test_Clean_And_InferTypes"
    On Error GoTo EH

    Dim data(1 To 3, 1 To 3) As Variant
    Dim hdr(1 To 3) As Variant
    data(1, 1) = " 10 ": data(1, 2) = " 2024-01-01 ": data(1, 3) = " NA "
    data(2, 1) = "11": data(2, 2) = "2024-01-02": data(2, 3) = "hello"
    data(3, 1) = "12": data(3, 2) = "2024-01-03": data(3, 3) = "-"
    hdr(1) = "n": hdr(2) = "d": hdr(3) = "txt"

    Dim df As New DataFrame
    Dim cleaned As DataFrame
    Dim typed As DataFrame

    df.LoadFromArray data, hdr
    Set cleaned = df.Clean(True, True, True)
    Set typed = cleaned.InferTypes()

    AssertEquals T, cleaned.AsArray()(1, 1), 10
    AssertTrue T, IsDate(cleaned.AsArray()(1, 2)), "Expected date conversion"
    AssertTrue T, IsEmpty(cleaned.AsArray()(1, 3)), "Expected null token conversion"

    Dim m As Variant
    m = typed.Metrics()
    AssertTrue T, UBound(m, 1) >= 1, "Expected metrics entries"

    PrintPass T
    Exit Sub
EH:
    PrintFail T, Err.Description
End Sub
Attribute VB_Name = "DataFrameTests"
Option Explicit

Private Sub AssertEquals(ByVal testName As String, ByVal actual As Variant, ByVal expected As Variant)
    If CStr(actual) <> CStr(expected) Then
        Err.Raise vbObjectError + 21001, testName, "Expected=" & CStr(expected) & " Actual=" & CStr(actual)
    End If
End Sub

Private Sub PrintPass(ByVal testName As String)
    Debug.Print "PASS - " & testName
End Sub

Private Sub PrintFail(ByVal testName As String, ByVal errMsg As String)
    Debug.Print "FAIL - " & testName & " - " & errMsg
End Sub

Public Sub Test_LoadFromArray_Basic()
    Const T As String = "Test_LoadFromArray_Basic"
    On Error GoTo EH

    Dim data(1 To 2, 1 To 2) As Variant
    Dim hdr(1 To 2) As Variant
    data(1, 1) = 1: data(1, 2) = "A"
    data(2, 1) = 2: data(2, 2) = "B"
    hdr(1) = "id": hdr(2) = "name"

    Dim df As New DataFrame
    df.LoadFromArray data, hdr

    AssertEquals T, df.RowsCount, 2
    AssertEquals T, df.ColsCount, 2
    AssertEquals T, df.AsArray()(2, 2), "B"

    PrintPass T
    Exit Sub
EH:
    PrintFail T, Err.Description
End Sub

Public Sub Test_Project_And_Rename()
    Const T As String = "Test_Project_And_Rename"
    On Error GoTo EH

    Dim data(1 To 3, 1 To 3) As Variant
    Dim hdr(1 To 3) As Variant
    data(1, 1) = 1: data(1, 2) = "Anna": data(1, 3) = "IT"
    data(2, 1) = 2: data(2, 2) = "Bruno": data(2, 3) = "HR"
    data(3, 1) = 3: data(3, 2) = "Carla": data(3, 3) = "IT"
    hdr(1) = "id": hdr(2) = "name": hdr(3) = "dept"

    Dim df As New DataFrame
    Dim projected As DataFrame
    Dim renamed As DataFrame

    df.LoadFromArray data, hdr
    Set projected = df.Project("dept,name")
    Set renamed = projected.Rename("dept:team,name:full_name")

    AssertEquals T, renamed.ColsCount, 2
    AssertEquals T, renamed.header()(1), "team"
    AssertEquals T, renamed.header()(2), "full_name"
    AssertEquals T, renamed.AsArray()(2, 1), "HR"

    PrintPass T
    Exit Sub
EH:
    PrintFail T, Err.Description
End Sub

Public Sub Test_Append_SchemaMismatch_ShouldFail()
    Const T As String = "Test_Append_SchemaMismatch_ShouldFail"
    On Error GoTo EH

    Dim dataA(1 To 1, 1 To 2) As Variant
    Dim hdrA(1 To 2) As Variant
    dataA(1, 1) = 1: dataA(1, 2) = "A"
    hdrA(1) = "id": hdrA(2) = "name"

    Dim dataB(1 To 1, 1 To 2) As Variant
    Dim hdrB(1 To 2) As Variant
    dataB(1, 1) = "A": dataB(1, 2) = 1
    hdrB(1) = "name": hdrB(2) = "code"

    Dim leftDf As New DataFrame
    Dim rightDf As New DataFrame

    leftDf.LoadFromArray dataA, hdrA
    rightDf.LoadFromArray dataB, hdrB

    leftDf.Append rightDf

    PrintFail T, "Expected failure for schema mismatch but operation succeeded"
    Exit Sub
EH:
    If InStr(1, Err.Description, "Colonna non trovata", vbTextCompare) > 0 _
        Or InStr(1, Err.Description, "non coerente", vbTextCompare) > 0 Then
        PrintPass T
    Else
        PrintFail T, Err.Description
    End If
End Sub

Public Sub Test_Dedup_ByKeys()
    Debug.Print "FAIL - Test_Dedup_ByKeys - TODO (Dedup non ancora implementato)"
End Sub

Public Sub Test_Filter_Contains()
    Debug.Print "FAIL - Test_Filter_Contains - TODO (Filter non ancora implementato)"
End Sub

Public Sub Test_LoadFromRange_WithHeader()
    Debug.Print "FAIL - Test_LoadFromRange_WithHeader - test manuale su worksheet richiesto"
End Sub
