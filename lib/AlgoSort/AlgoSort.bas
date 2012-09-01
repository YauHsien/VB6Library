Attribute VB_Name = "AlgoSort"
-Attribute VB_Name = "AlgoSort"
Option Explicit

Public Enum SortOrderEnum
    soeAscending
    soeDescending
End Enum

Public Function QuickSort(Unsort() As Variant, Optional SortComparer As String = "LessThan", Optional SortOrder As SortOrderEnum = SortOrderEnum.soeAscending) As Variant()
    Dim result() As Variant
    If UBound(Unsort) - LBound(Unsort) = 0 Then
        result = Unsort
    Else
        Dim lenResult As Long
        Dim partLeft() As Variant
        Dim lenLeft As Long
        Dim partRight() As Variant
        Dim lenRight As Long
        Dim result1() As Variant
        Dim i As Long
        Dim comparison As New AlgoSortComparison
        lenLeft = 0
        lenRight = 0
        For i = LBound(Unsort) + 1 To UBound(Unsort)
            If CallByName(comparison, SortComparer, VbMethod, CVar(Unsort(i)), CVar(Unsort(0))) Then
                If SortOrder = soeAscending Then
                    lenLeft = lenLeft + 1
                    ReDim Preserve partLeft(lenLeft - 1) As Variant
                    partLeft(lenLeft - 1) = Unsort(i)
                Else '''if sortorder = soeDescending then
                    lenRight = lenRight + 1
                    ReDim Preserve partRight(lenRight - 1) As Variant
                    partRight(lenRight - 1) = Unsort(i)
                End If
            Else '''if unsort(i) >= unsort(0) then
                If SortOrder = soeAscending Then
                    lenRight = lenRight + 1
                    ReDim Preserve partRight(lenRight - 1) As Variant
                    partRight(lenRight - 1) = Unsort(i)
                Else '''if sortorder = soeDescending then
                    lenLeft = lenLeft + 1
                    ReDim Preserve partLeft(lenLeft - 1) As Variant
                    partLeft(lenLeft - 1) = Unsort(i)
                End If
            End If
        Next i
        Set comparison = Nothing
        If lenLeft = 0 Then
            lenResult = 0
        Else
            result = QuickSort(partLeft, SortComparer, SortOrder)
            lenResult = UBound(result) - LBound(result) + 1
        End If
        ReDim Preserve result(lenResult) As Variant
        result(lenResult) = Unsort(0)
        lenResult = lenResult + 1
        If lenRight > 0 Then
            result1 = QuickSort(partRight, SortComparer, SortOrder)
            ReDim Preserve result(lenResult + UBound(result1) - LBound(result1)) As Variant
            For i = LBound(result1) - LBound(result1) To UBound(result1) - LBound(result1)
                result(lenResult + i - LBound(result1)) = result1(i)
            Next i
        End If
    End If
    QuickSort = result
End Function

Public Function test() As Boolean
    Dim result() As Variant
    Dim i As Long
    Dim c As AlgoSortComparison
    On Error GoTo err_test
    Set c = New AlgoSortComparison
    AlgoSortHelper.ConfirmMethodName c, "LengthLessThan"
    result = QuickSort(AlgoData.CVarArr(Array("hello", ",", "world")), c.LengthLessThan, soeAscending)
    For i = LBound(result) To UBound(result)
        Debug.Print result(i); "!",
    Next i
    Debug.Print
    test = True
    Exit Function
err_test:
    MsgBox Err.Description
End Function

