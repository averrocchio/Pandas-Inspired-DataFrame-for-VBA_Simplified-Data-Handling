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

Public Sub Test_Append_HeaderUnion()
    Const T As String = "Test_Append_HeaderUnion"
    On Error GoTo EH

    Dim dataA(1 To 2, 1 To 5) As Variant
    Dim hdrA(1 To 5) As Variant
    dataA(1, 1) = "D1": dataA(1, 2) = "D2": dataA(1, 3) = "D4": dataA(1, 4) = "D5": dataA(1, 5) = "D6"
    dataA(2, 1) = "D7": dataA(2, 2) = "D8": dataA(2, 3) = "D9": dataA(2, 4) = "D10": dataA(2, 5) = "D11"
    hdrA(1) = "COL1": hdrA(2) = "COL2": hdrA(3) = "COL3": hdrA(4) = "COL4": hdrA(5) = "COL5"

    Dim dataB(1 To 1, 1 To 4) As Variant
    Dim hdrB(1 To 4) As Variant
    dataB(1, 1) = "D12": dataB(1, 2) = "D13": dataB(1, 3) = "D14": dataB(1, 4) = "D15"
    hdrB(1) = "COL2": hdrB(2) = "COL3": hdrB(3) = "COL5": hdrB(4) = "COL6"

    Dim leftDf As New DataFrame
    Dim rightDf As New DataFrame
    Dim out As DataFrame

    leftDf.LoadFromArray dataA, hdrA
    rightDf.LoadFromArray dataB, hdrB

    Set out = leftDf.Append(rightDf)

    AssertEquals T, out.ColsCount, 6
    AssertEquals T, out.RowsCount, 3
    AssertEquals T, out.header()(6), "COL6"
    AssertEquals T, out.AsArray()(1, 6), ""
    AssertEquals T, out.AsArray()(3, 1), ""
    AssertEquals T, out.AsArray()(3, 2), "D12"
    AssertEquals T, out.AsArray()(3, 6), "D15"

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

    hdr(1) = "id": hdr(2) = "team": hdr(3) = "score"
    data(1, 1) = 3: data(1, 2) = "B": data(1, 3) = 10
    data(2, 1) = 1: data(2, 2) = "A": data(2, 3) = 30
    data(3, 1) = 4: data(3, 2) = "A": data(3, 3) = 20
    data(4, 1) = 2: data(4, 2) = "B": data(4, 3) = 40

    Dim df As New DataFrame
    Dim sorted As DataFrame
    df.LoadFromArray data, hdr

    Set sorted = df.Sort("team,score", "asc,desc")

    AssertEquals T, sorted.AsArray()(1, 1), 1
    AssertEquals T, sorted.AsArray()(2, 1), 4
    AssertEquals T, sorted.AsArray()(3, 1), 2
    AssertEquals T, sorted.AsArray()(4, 1), 3

    PrintPass T
    Exit Sub
EH:
    PrintFail T, Err.Description
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
