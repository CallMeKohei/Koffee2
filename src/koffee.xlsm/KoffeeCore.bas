Attribute VB_Name = "KoffeeCore"
'' --------------------------------------------------------
''  FILE    : KoffeeCore.bas
''  AUTHOR  : callmekohei <callmekohei at gmail.com>
''  License : MIT license
'' --------------------------------------------------------
Option Explicit
Option Private Module

''' Dependencies
'''
'''     IsJagArr
'''         ArrRank(Ariawase)
'''
'''     ArrExplodeImpl
'''         ArrayEx(Ariawase)

Public Function IsJagArr(ByVal arr As Variant) As Boolean

    If Not IsArray(arr) Then GoTo Escape
    On Error GoTo Escape

    If ArrRank(arr) > 1 Then GoTo Escape

    Dim v1 As Variant, v2 As Variant

    For Each v1 In arr
        If Not IsObject(v1) Then
            For Each v2 In v1
                If Not IsObject(v2) Then
                    IsJagArr = True
                    GoTo Escape
                End If
            Next v2
        End If
    Next v1

Escape:
End Function

Public Function Base01(ByVal arr As Variant, _
    Optional ByVal BaseOne As Boolean = False, _
    Optional ByRef acc As Variant, _
    Optional ByRef acc_i As Long, _
    Optional ByVal acc_ub As Long, _
    Optional ByVal level As Long, _
    Optional ByRef leaf As Long) As Variant

    If IsMissing(acc) Then
        If BaseOne = True Then
            ReDim acc(1 To 32)
            acc_i = 1
            acc_ub = 32
        Else
            ReDim acc(32)
            acc_i = 0
            acc_ub = 32
        End If
    End If

    Dim v
    If IsJagArr(arr) Then
        For Each v In arr
            Base01 = Base01(v, BaseOne, acc, acc_i, acc_ub, level + 1, leaf)
        Next v
    Else

        Dim a(), i As Long
        If BaseOne = True Then
            ReDim a(1 To 32)
            i = 1
        Else
            ReDim a(32)
            i = 0
        End If

        Dim ub   As Long: ub = 32

        For Each v In arr

            If ub = i Then
                ub = ub + 1
                ub = -1 + ub + ub
                If BaseOne = True Then
                    ReDim Preserve a(1 To ub - 1)
                Else
                    ReDim Preserve a(ub - 1)
                End If
            End If

            If IsObject(v) Then
                Set a(i) = v
            Else
                Let a(i) = v
            End If

            i = i + 1

        Next v

        If acc_ub = acc_i Then
            acc_ub = acc_ub + 1
            acc_ub = -1 + acc_ub + acc_ub
            If BaseOne = True Then
                ReDim Preserve acc(1 To acc_ub - 1)
                acc_ub = UBound(acc)
            Else
                ReDim Preserve acc(acc_ub - 1)
                acc_ub = UBound(acc)
            End If
        End If

        acc(acc_i) = Truncate(a)
        acc_i = acc_i + 1

    End If

    leaf = IIf(leaf < level, level, leaf)

    If level = 0 Then
        Base01 = ArrUnflatten(Truncate(acc), leaf - 1, BaseOne)
    Else
        Base01 = acc
    End If

End Function

Public Function Truncate(ByVal arr As Variant) As Variant

    Dim lb As Long: lb = LBound(arr)

    Dim i As Long, n As Long: n = UBound(arr)
    For i = 0 To UBound(arr)
        If Not IsEmpty(arr(n - i)) Then
            If lb > 0 Then
                ReDim Preserve arr(1 To n - i)
            Else
                ReDim Preserve arr(n - i)
            End If
            Truncate = arr
            GoTo Escape
        End If
    Next i

Escape:
    If UBound(arr) = -1 Or IsEmpty(Truncate) Then
        Truncate = Array(Empty)
    End If
End Function

Public Function ArrUnflatten(ByVal arr As Variant, _
    Optional ByVal n As Long = 1, _
    Optional ByVal BaseOne As Boolean = False) As Variant

    Dim i As Long
    For i = 0 To n - 1
        If BaseOne Then
            arr = ArrShift(arr, Array(), True)
        Else
            arr = ArrShift(arr, Array())
        End If
    Next i
    ArrUnflatten = arr
End Function

Public Function ArrShift(ByVal val As Variant, ByVal arr As Variant, Optional ByVal BaseOne As Boolean = False) As Variant

    Dim a() As Variant: If BaseOne Then ReDim a(1 To 32) Else ReDim a(32)
    Dim ub As Long:  ub = 32
    Dim i As Long:  If BaseOne Then i = 2 Else i = 1

    If IsObject(val) Then
        If BaseOne Then Set a(1) = val Else Set a(0) = val
    Else
        If BaseOne Then Let a(1) = val Else Let a(0) = val
    End If

    Dim v As Variant
    For Each v In arr

        If ub = i Then
            ub = ub + 1
            ub = -1 + ub + ub
            If BaseOne Then ReDim Preserve a(1 To ub - 1) Else ReDim Preserve a(ub - 1)
        End If

        If IsObject(v) Then
            Set a(i) = v
        Else
            Let a(i) = v
        End If

        i = i + 1
    Next v

    ArrShift = Truncate(a)

End Function

Public Function ArrExplode(ParamArray arr() As Variant) As Variant
    ArrExplode = ArrExplodeImpl(arr).ToArray
End Function

Private Function ArrExplodeImpl(ByVal arr As Variant, Optional ByRef acc As ArrayEx) As ArrayEx
    If acc Is Nothing Then Set acc = New ArrayEx
    If IsArray(arr) Then
        Dim v
        For Each v In arr
            If Not IsArray(v) Then
                If IsObject(v) Then
                    acc.AddObj v
                Else
                    acc.AddVal v
                End If
            End If
            Set ArrExplodeImpl = ArrExplodeImpl(v, acc)
        Next v
    Else
        Set ArrExplodeImpl = acc
    End If
End Function
