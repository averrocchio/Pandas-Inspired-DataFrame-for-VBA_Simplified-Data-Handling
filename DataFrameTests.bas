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
