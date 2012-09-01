Attribute VB_Name = "AlgoSortHelper"
Option Explicit

Public Function ConfirmMethodName(ObjectInstance As Object, MethodName As String) As Boolean
    Dim MethodName1 As String
    On Error GoTo err_ConfirmMethodName
    MethodName1 = CallByName(ObjectInstance, MethodName, VbMethod)
    If MethodName1 = MethodName Then
        ConfirmMethodName = True
    Else
        On Error GoTo 0
        Err.Raise 5502, "AlgoSortComparison", "所指定方法名稱 " & MethodName & " 不符合方法的自我描述名稱 " & MethodName1 & " ，" & vbCrLf & "請核對物件類別模組 AlgoSortComparison 的定義。"
    End If
    Exit Function
err_ConfirmMethodName:
    On Error GoTo 0
    Err.Raise 5501, "AlgoSortComparison", "找不到指定的方法名稱 " & MethodName & " ，" & vbCrLf & "請核對物件類別模組 AlgoSortComparison 的定義。"
End Function

Public Sub RegisterComparer(A As Variant, B As Variant, MethodName As String, Predicate As Boolean, ByRef ReturnValue As Variant)
    If IsMissing(A) Or IsMissing(B) Then
        ReturnValue = MethodName
    Else
        ReturnValue = Predicate
    End If
End Sub

