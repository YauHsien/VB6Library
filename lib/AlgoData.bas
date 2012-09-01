Attribute VB_Name = "AlgoData"
Option Explicit

Public Function CVarArr(AnArray As Variant) As Variant()
    Dim result() As Variant
    If IsArray(AnArray) Then
        Dim i As Long
        ReDim result(LBound(AnArray) To UBound(AnArray)) As Variant
        For i = LBound(AnArray) To UBound(AnArray)
            result(i) = AnArray(i)
        Next i
    Else
        ReDim result(0) As Variant
        result(0) = AnArray
    End If
    CVarArr = result
End Function

