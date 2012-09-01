VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlgoSortComparison"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function LessThan(Optional A As Variant, Optional B As Variant) As Variant
    Call AlgoSortHelper.RegisterComparer(A:=A, B:=B, MethodName:="LessThan", Predicate:=(A < B), ReturnValue:=LessThan)
End Function

Public Function LengthLessThan(Optional A As Variant, Optional B As Variant) As Variant
    Call AlgoSortHelper.RegisterComparer(A:=A, B:=B, MethodName:="LengthLessThan", Predicate:=(Len(CStr(A)) < Len(CStr(B))), ReturnValue:=LengthLessThan)
End Function

