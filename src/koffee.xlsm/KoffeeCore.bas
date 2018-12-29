Attribute VB_Name = "KoffeeCore"
''' --------------------------------------------------------
'''  FILE    : KoffeeCore.bas
'''  AUTHOR  : callmekohei <callmekohei at gmail.com>
'''  License : MIT license
''' --------------------------------------------------------
Option Explicit

Public Function RandomBetween(Optional ByVal UpperBound As Variant = 0, Optional ByVal LowerBound As Variant = 9) As Variant
    Randomize
    RandomBetween = CInt((UpperBound - LowerBound + 1) * Rnd()) + LowerBound
End Function

Public Function ArrShuffle(ByVal arr As Variant) As Variant
    Dim i As Long, n As Integer, tmp As Variant
    For i = 0 To UBound(arr)
        Randomize: n = Int((UBound(arr) + 1) * Rnd)
        tmp = arr(i)
        arr(i) = arr(n)
        arr(n) = tmp
    Next
    ArrShuffle = arr
End Function

Public Function FetchSh(ByVal sql As String, Optional ByVal fpath As String = "", Optional ByVal isHeader As Boolean = False) As Variant

    Select Case fpath
        Case Is = "": fpath = ThisWorkbook.Path & "\" & ThisWorkbook.Name
        Case Else:    fpath = fpath
    End Select

    Dim ado As AdoEx: Set ado = New AdoEx
    ado.Init adExcel, fpath
    Select Case isHeader
        Case False: FetchSh = ado.JagArrAdoRS(sql)
        Case True:  FetchSh = Array(ado.JagArrAdoRS(sql), ado.JagArrAdoRsHeader(sql))
    End Select

    Set ado = Nothing
End Function

Public Function FetchCSV(ByVal sql As String, ByVal fpath As String, Optional ByVal isHeader As Boolean = False) As Variant
    Dim ado As AdoEx: Set ado = New AdoEx
    ado.Init adCsv, fpath
    FetchCSV = ado.JagArrAdoRS(sql)

    Select Case isHeader
        Case False: FetchCSV = ado.JagArrAdoRS(sql)
        Case True:  FetchCSV = Array(ado.JagArrAdoRS(sql), ado.JagArrAdoRsHeader(sql))
    End Select
    Set ado = Nothing
End Function

Public Function IsJagArr(ByVal arr As Variant) As Boolean

    If Not IsArray(arr) Then GoTo Escape
    On Error GoTo Escape

    If ArrRank(arr) > 1 Then GoTo Escape

    'Not JagArr -> Err.raise 13
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

Public Function Base01(ByVal arr As Variant, Optional ByVal BaseOne As Boolean = False, _
    Optional ByRef acc As Variant, Optional ByRef acc_i As Long, Optional ByVal acc_ub As Long, _
    Optional ByVal level As Long, Optional ByRef leaf As Long) As Variant

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

Public Function ArrUnflatten(ByVal arr As Variant, Optional ByVal n As Long = 1, Optional ByVal BaseOne As Boolean = False) As Variant
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

Public Function ArrFill(ByVal arr As Variant, ByVal n As Long, Optional ByVal filled As Variant = 0, Optional isRear As Boolean = False) As Variant

    If n <= UBound(arr) Then
        ArrFill = arr
        GoTo Escape
    End If

    Dim a(): ReDim a(32)
    Dim ub As Long:  ub = 32

    Dim i As Long
    For i = 0 To n - UBound(arr) - 1
        If ub = i Then
            ub = ub + 1
            ub = -1 + ub + ub
            ReDim Preserve a(ub - 1)
        End If
        If IsObject(filled) Then
            Set a(i) = filled
        Else
            Let a(i) = filled
        End If

    Next i

    If isRear Then
        ArrFill = ArrConcat(Truncate(a), arr)
    Else
        ArrFill = ArrConcat(arr, Truncate(a))
    End If
Escape:
End Function

Public Function nth(ByVal index As Variant, ByVal arr As Variant) As Variant

    If Not IsArray(arr) Then Err.Raise 13
    If Not IsNumeric(index) Then Err.Raise 13
    If index < LBound(arr) Or index > UBound(arr) Then Err.Raise 13

    If IsObject(arr(index)) Then
        Set nth = arr(index)
    Else
        Let nth = arr(index)
    End If

End Function

Public Function Rest(ByVal arr As Variant) As Variant

    Dim lb As Long: lb = LBound(arr)

    Dim k As Long, cnt As Long, n As Long: n = UBound(arr)
    For k = 0 To UBound(arr)
        If IsEmpty(arr(n - k)) Then
            cnt = cnt + 1
        Else
            Exit For
        End If
    Next k

    Dim a(): ReDim a(32)
    Dim ub As Long:  ub = 32
    Dim j As Long: j = 0


    Dim i As Long
    For i = 1 To UBound(arr)

        If ub = j Then
            ub = ub + 1
            ub = -1 + ub + ub
            ReDim Preserve a(ub - 1)
        End If

        If IsObject(arr(i)) Then
            Set a(j) = arr(i)
        Else
            a(j) = arr(i)
        End If

        j = j + 1

    Next i

    If cnt <> 0 Then
        Rest = ArrConcat(Truncate(a), ArrFill(Array(), cnt, Empty))
    Else
        Rest = Truncate(a)
    End If

End Function

'About Array sort
Public Function ArrSortAsc(ByVal arr As Variant) As Variant
    ArrSort arr, True
    ArrSortAsc = arr
End Function

Public Function ArrSortDec(ByVal arr As Variant) As Variant
    ArrSort arr, False
    ArrSortDec = arr
End Function

Public Function ArrRev2(ByVal arr As Variant) As Variant
    ArrRev arr
    ArrRev2 = arr
End Function

'About Array Cal
Public Function ArrPlus(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant

    If Not (IsArray(arr1) And IsArray(arr2)) Then Err.Raise 13
    If (ArrRank(arr1) <> ArrRank(arr2)) And (UBound(arr1) <> UBound(arr2)) Then Err.Raise 13

    Dim a(): ReDim a(32)
    Dim ub As Long:  ub = 32

    Dim i As Long, arrx As ArrayEx: Set arrx = New ArrayEx
    For i = 0 To UBound(arr1)

        If ub = i Then
            ub = ub + 1
            ub = -1 + ub + ub
            ReDim Preserve a(ub - 1)
        End If

        a(i) = CCur(arr1(i)) + CCur(arr2(i))

    Next i

    ArrPlus = Truncate(a)

End Function

Public Function ArrMinus(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant

    If Not (IsArray(arr1) And IsArray(arr2)) Then Err.Raise 13
    If (ArrRank(arr1) <> ArrRank(arr2)) And (UBound(arr1) <> UBound(arr2)) Then Err.Raise 13

    Dim a(): ReDim a(32)
    Dim ub As Long:  ub = 32

    Dim i As Long, arrx As ArrayEx: Set arrx = New ArrayEx
    For i = 0 To UBound(arr1)

        If ub = i Then
            ub = ub + 1
            ub = -1 + ub + ub
            ReDim Preserve a(ub - 1)
        End If

        a(i) = CCur(arr1(i)) - CCur(arr2(i))

    Next i

    ArrMinus = Truncate(a)

End Function

Public Function ArrDivide(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant

    If Not (IsArray(arr1) And IsArray(arr2)) Then Err.Raise 13
    If (ArrRank(arr1) <> ArrRank(arr2)) And (UBound(arr1) <> UBound(arr2)) Then Err.Raise 13

    Dim a(): ReDim a(32)
    Dim ub As Long:  ub = 32

    Dim i As Long, arrx As ArrayEx: Set arrx = New ArrayEx
    For i = 0 To UBound(arr1)

        If ub = i Then
            ub = ub + 1
            ub = -1 + ub + ub
            ReDim Preserve a(ub - 1)
        End If

        a(i) = CCur(arr1(i)) / CCur(arr2(i))

    Next i

    ArrDivide = Truncate(a)

End Function

'About SET operation
Public Function ArrUnion(ByVal s1 As Variant, ByVal s2 As Variant) As Variant
    If Not (IsArray(s1) And IsArray(s2)) Then Err.Raise 13
    If IsJagArr(s1) Or IsJagArr(s2) Then Err.Raise 13
    If Not (ArrRank(s1) = 1 Or ArrRank(s2) = 1) Then Err.Raise 13

    ArrUnion = ArrUniq(ArrConcat(s1, s2))
End Function

Public Function ArrDiff(ByVal s1 As Variant, s2 As Variant) As Variant
    If Not (IsArray(s1) And IsArray(s2)) Then Err.Raise 13
    If IsJagArr(s1) Or IsJagArr(s2) Then Err.Raise 13
    If Not (ArrRank(s1) = 1 Or ArrRank(s2) = 1) Then Err.Raise 13

        Dim v As Variant, arrx As ArrayEx: Set arrx = New ArrayEx
    For Each v In s1
        If ArrIndexOf(s2, v) = -1 Then arrx.AddVal v
    Next v
    ArrDiff = arrx.ToArray
End Function

Public Function ArrDiff2(ByVal s1 As Variant, ByVal s2 As Variant) As Variant
    If Not (IsArray(s1) And IsArray(s2)) Then Err.Raise 13
    If IsJagArr(s1) Or IsJagArr(s2) Then Err.Raise 13
    If Not (ArrRank(s1) = 1 Or ArrRank(s2) = 1) Then Err.Raise 13

    ArrDiff2 = ArrUniq(ArrConcat(ArrDiff(s1, s2), ArrDiff(s2, s1)))
End Function

Public Function ArrIntersect(ByVal s1 As Variant, s2 As Variant) As Variant
    If Not (IsArray(s1) And IsArray(s2)) Then Err.Raise 13
    If IsJagArr(s1) Or IsJagArr(s2) Then Err.Raise 13
    If Not (ArrRank(s1) = 1 Or ArrRank(s2) = 1) Then Err.Raise 13

    Dim v As Variant, arrx As ArrayEx: Set arrx = New ArrayEx
    For Each v In s1
        If ArrIndexOf(s2, v) > -1 Then arrx.AddVal v
    Next v
    ArrIntersect = arrx.ToArray
End Function

'About Array operation
Public Function ArrShift(ByVal val As Variant, ByVal arr As Variant, Optional ByVal BaseOne As Boolean = False) As Variant

    Dim a(): If BaseOne Then ReDim a(1 To 32) Else ReDim a(32)
    Dim ub As Long:  ub = 32
    Dim i As Long:  If BaseOne Then i = 2 Else i = 1

    If IsObject(val) Then
        If BaseOne Then Set a(1) = val Else Set a(0) = val
    Else
        If BaseOne Then Let a(1) = val Else Let a(0) = val
    End If

    Dim v
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

Public Function ArrUnshift(ByVal arr As Variant) As Variant
    ArrUnshift = Rest(arr)
End Function

Public Function ArrPush(ByVal val As Variant, ByVal arr As Variant) As Variant
    ReDim Preserve arr(UBound(arr) + 1)
    If IsObject(val) Then
        Set arr(UBound(arr)) = val
    Else
        Let arr(UBound(arr)) = val
    End If
    ArrPush = arr
End Function

Public Function ArrPop(ByVal arr As Variant) As Variant
    ReDim Preserve arr(UBound(arr) - 1)
    ArrPop = arr
End Function

'Type Converter
Public Function ArrCLng(ByVal arr As Variant) As Variant

    Dim a(): ReDim a(32)
    Dim ub As Long:  ub = 32
    Dim i As Long: i = 0

    Dim v, arrx As ArrayEx: Set arrx = New ArrayEx
    For Each v In arr

        If ub = i Then
            ub = ub + 1
            ub = -1 + ub + ub
            ReDim Preserve a(ub - 1)
        End If

        a(i) = CLng(v)
        i = i + 1
    Next v
    ArrCLng = Truncate(a)
End Function

Public Function ArrCCur(ByVal arr As Variant) As Variant

    Dim a(): ReDim a(32)
    Dim ub As Long:  ub = 32
    Dim i As Long: i = 0

    Dim v, arrx As ArrayEx: Set arrx = New ArrayEx
    For Each v In arr

        If ub = i Then
            ub = ub + 1
            ub = -1 + ub + ub
            ReDim Preserve a(ub - 1)
        End If

        a(i) = CCur(v)
        i = i + 1
    Next v
    ArrCCur = Truncate(a)
End Function

Public Function StepTotal(ByVal arr As Variant, Optional ByRef acc As Variant) As Variant

    If IsMissing(acc) Then acc = Array()

    Dim v
    If IsJagArr(arr) Then
        For Each v In arr
            StepTotal = StepTotal(v, acc)
        Next v
    Else

        Dim a(): ReDim a(32)
        Dim ub As Long:  ub = 32
        Dim i As Long: i = 0
        Dim tmp As Currency

        For Each v In arr

            If ub = i Then
                ub = ub + 1
                ub = -1 + ub + ub
                ReDim Preserve a(ub - 1)
            End If

            tmp = tmp + CCur(v)
            a(i) = tmp
            i = i + 1
        Next v

        acc = ArrPush(Truncate(a), acc)

    End If

    StepTotal = acc

End Function

Public Function Nz2(ByVal arr As Variant, Optional ByVal ValueIfNull As Variant, Optional ByVal EnableSpeedUp As Boolean = False) As Variant

    Dim Alt
    If IsMissing(ValueIfNull) Then
        Alt = IIf(IsNumericArray(arr, EnableSpeedUp), 0, "")
    Else
        Alt = ValueIfNull
    End If

    Dim v, arrx As ArrayEx: Set arrx = New ArrayEx
    For Each v In arr
        arrx.AddVal IIf(IsNull(v), Alt, v)
    Next v

    Nz2 = arrx.ToArray

End Function

Private Function IsNumericArray(ByVal arr As Variant, Optional ByVal Economy As Boolean = False) As Boolean

    Dim myChoice:      myChoice = ArrShuffle(ArrRange(LBound(arr), UBound(arr)))
    Dim JudgeRange
    If Economy Then
        JudgeRange = ArrSlice(ArrShuffle(ArrRange(LBound(arr), UBound(arr))), 0, Int(ArrLen(arr) * (1 / 3)))
    Else
        JudgeRange = myChoice
    End If

    Dim v
    For Each v In JudgeRange
        If Not IsNumeric(arr(v)) Then
            If Not IsNull(arr(v)) Then
                IsNumericArray = False
                GoTo Escape
            End If
        End If
    Next

    IsNumericArray = True

Escape:
End Function
