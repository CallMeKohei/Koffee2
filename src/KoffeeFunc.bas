Attribute VB_Name = "KoffeeFunc"
''' --------------------------------------------------------
'''  FILE    : KoffeeFunc.bas
'''  AUTHOR  : callmekohei <callmekohei at gmail.com>
'''  License : MIT license
''' --------------------------------------------------------
Option Explicit

Public Function Flow(ByVal seed As Variant, ParamArray funcs() As Variant) As Variant

    If TypeName(funcs(0)) = "Atom" Then
        Select Case funcs(0).GetAddr
            Case Is = VBA.CLng(AddressOf ZipMark)
                funcs(0).FastApply seed, seed(0), seed(1)
                funcs = ArrUnshift(funcs)
            Case Is = VBA.CLng(AddressOf FoldMark), VBA.CLng(AddressOf ScanMark)
                If Not IsArray(seed(0)) And IsArray(seed(1)) Then
                    funcs(0).FastApply seed, seed(1), seed(0)
                    funcs = ArrUnshift(funcs)
                End If
        End Select
    End If

    If Not IsEmpty(funcs(0)) Then
        Dim fun As Variant
        For Each fun In funcs
            fun.FastApply seed, seed
        Next fun
    End If

    If IsObject(seed) Then
        Set Flow = seed
    Else
        Let Flow = seed
    End If

End Function

Public Function Compose(ParamArray funcs() As Variant) As Variant
    Dim a As New Atom: Set Compose = a.AddFunc(Init(New Func, vbVariant, AddressOf ComposeImpl, vbVariant, vbVariant)).Bind(funcs)
End Function

Public Function ComposeImpl(ByVal funcs As Variant, ByVal seed As Variant) As Variant

    If TypeName(funcs(0)) = "Atom" Then
        Select Case funcs(0).GetAddr
            Case Is = VBA.CLng(AddressOf ZipMark)
                funcs(0).FastApply seed, seed(0), seed(1)
                funcs = ArrUnshift(funcs)
            Case Is = VBA.CLng(AddressOf FoldMark), VBA.CLng(AddressOf ScanMark)
                If Not IsArray(seed(0)) And IsArray(seed(1)) Then
                    funcs(0).FastApply seed, seed(1), seed(0)
                    funcs = ArrUnshift(funcs)
                End If
        End Select
    End If

    If Not IsEmpty(funcs(0)) Then
        Dim fun As Variant
        For Each fun In funcs
            fun.FastApply seed, seed
        Next fun
    End If

    If IsObject(seed) Then
        Set ComposeImpl = seed
    Else
        Let ComposeImpl = seed
    End If

End Function

'About Partial
Public Function Partial(ByVal fun As Variant, ParamArray args() As Variant) As Atom

    If Not (TypeName(fun) = "Atom" Or TypeName(fun) = "Func") Then Err.Raise 13
    Dim a As New Atom: a.LetAddr = VBA.CLng(AddressOf PartialMark)

    Dim tmp As Variant
    If UBound(args) = -1 Then
        tmp = Missing
    Else
        If ArrRank(args) = 0 Then
            Set tmp = Tuple2Of(args(0), args(0))
        Else
            Set tmp = Tuple2Of(args, args)
        End If
    End If

    Set Partial = a.AddFunc(Init(New Func, vbVariant, AddressOf PartialImpl, vbVariant, vbVariant)).Bind(fun).Bind(tmp)

End Function

Public Function PartialMark() As Byte
End Function

Private Function PartialImpl(ByVal f As Variant, ByVal arg As Variant) As Variant

    Dim tmp
    If IsArray(arg) Then
        Select Case UBound(arg)
            Case Is < 0: Err.Raise 13
            Case Is = 0:
                If IsArray(arg(0)) Then
                    tmp = arg
                Else
                    tmp = Array(arg)
                End If
            Case Else
                tmp = arg
        End Select
    Else
        tmp = Array(arg)
    End If

    f.CallByPtr PartialImpl, tmp

End Function

'About fnull
Public Function fnull(ByVal fun As Variant, ByVal Alt As Variant) As Atom

    If Not (TypeName(fun) = "Atom" Or TypeName(fun) = "Func") Then Err.Raise 13
    Dim a As New Atom: Set fnull = a.AddFunc(Init(New Func, vbVariant, AddressOf fnullImpl, vbVariant, vbVariant, vbVariant)).Bind(fun).Bind(Alt)

    If TypeName(fun) = "Atom" Then
        If fun.GetAddr = VBA.CLng(AddressOf PartialMark) Then
            a.LetAddr = VBA.CLng(AddressOf fnullPartialMark)
        Else
            a.LetAddr = VBA.CLng(AddressOf fnullMark)
        End If
    Else
            a.LetAddr = VBA.CLng(AddressOf fnullMark)
    End If

End Function

Public Function fnullMark() As Byte
End Function

Public Function fnullPartialMark() As Byte
End Function

Private Function fnullImpl(ByVal f As Variant, ByVal AltArg As Variant, ByVal arg As Variant) As Variant

    If IsMissing(arg) Then
        arg = Array(Missing)
    ElseIf Not IsArray(arg) Then
        arg = Array(arg)
    End If

    f.CallByPtr fnullImpl, Flow(arg, Map(Partial(NLV, AltArg)))

End Function

Public Function Map(ByVal f As Variant) As Atom
    Select Case TypeName(f)
        Case Is = "Func", "Atom", "Long"
            If TypeName(f) = "Long" Then Set f = Init(New Func, vbVariant, f, vbVariant)
            Dim a As New Atom: Set Map = a.AddFunc(Init(New Func, vbVariant, AddressOf MapImpl, vbVariant, vbVariant)).Bind(f)
        Case Else
            Err.Raise 13
    End Select
End Function

Public Function MapImpl(ByVal f As Variant, ByVal arr As Variant) As Variant
    If Not IsArray(arr) Then Err.Raise 13
    Dim lb As Long: lb = LBound(arr)
    Dim ub As Long: ub = UBound(arr)
    Dim ret As Variant
    If ub - lb < 0 Then
        ret = Array()
        GoTo Ending
    End If

    ReDim ret(lb To ub)

    Dim i As Long
    Select Case TypeName(f)
        Case "Func"
            For i = lb To ub: f.FastApply ret(i), arr(i): Next
        Case "Atom"
            For i = lb To ub
                f.FastApply ret(i), arr(i)
                If f.GetAddr = VBA.CLng(AddressOf fnullPartialMark) Then f.Pop
            Next i
        Case Else
            Err.Raise 13
    End Select

Ending:
    MapImpl = ret
End Function

Public Function Filter(ByVal f As Variant) As Atom
    Select Case TypeName(f)
        Case Is = "Func", "Atom", "Long"
            If TypeName(f) = "Long" Then Set f = Init(New Func, vbBoolean, f, vbVariant)
            Dim a As New Atom: Set Filter = a.AddFunc(Init(New Func, vbVariant, AddressOf FilterImpl, vbObject, vbVariant)).Bind(f)
        Case Else
            Err.Raise 13
    End Select
End Function

Public Function FilterImpl(ByVal f As Object, ByVal arr As Variant) As Variant
    If Not IsArray(arr) Then Err.Raise 13
    Dim lb As Long: lb = LBound(arr)
    Dim ub As Long: ub = UBound(arr)
    Dim ret As Variant
    If ub - lb < 0 Then
        ret = Array()
        GoTo Ending
    End If

    ReDim ret(lb To ub)

    Dim flg As Boolean
    Dim ixArr As Long
    Dim ixRet As Long: ixRet = lb

    If IsObject(arr(lb)) Then
        For ixArr = lb To ub
            f.FastApply flg, arr(ixArr)
            If TypeName(f) = "Atom" Then f.Pop
            If flg Then Set ret(IncrPst(ixRet)) = arr(ixArr)
        Next
    Else
        For ixArr = lb To ub
            f.FastApply flg, arr(ixArr)
            If TypeName(f) = "Atom" Then f.Pop
            If flg Then Let ret(IncrPst(ixRet)) = arr(ixArr)
        Next
    End If

    If ixRet > 0 Then
        ReDim Preserve ret(lb To ixRet - 1)
    Else
        ret = Array()
    End If

Ending:
    FilterImpl = ret
End Function

Private Sub ArrFoldPrep( _
    arr As Variant, seedv As Variant, i As Long, stat As Variant, _
    Optional isObj As Boolean _
    )

    If IsObject(seedv) Then
        Set stat = seedv
    Else
        Let stat = seedv
    End If

    If IsMissing(stat) Then
        isObj = IsObject(arr(i))
        If isObj Then
            Set stat = arr(i)
        Else
            Let stat = arr(i)
        End If
        i = i + 1
    End If
End Sub

Public Function Fold(ByVal f As Variant) As Atom
    Select Case TypeName(f)
        Case Is = "Func", "Atom", "Long"
            If TypeName(f) = "Long" Then Set f = Init(New Func, vbVariant, f, vbVariant, vbVariant)
            Dim a As New Atom: Set Fold = a.AddFunc(Init(New Func, vbVariant, AddressOf FoldImpl, vbObject, vbVariant, vbVariant)).Bind(f)
            a.LetAddr = VBA.CLng(AddressOf FoldMark)
        Case Else
            Err.Raise 13
    End Select
End Function

Public Function FoldMark() As Byte
End Function

Public Function FoldImpl( _
    ByVal f As Object, ByVal arr As Variant, Optional ByVal seedv As Variant _
    ) As Variant

    If Not IsArray(arr) Then Err.Raise 13

    Dim stat As Variant
    Dim i As Long: i = LBound(arr)
    ArrFoldPrep arr, seedv, i, stat

    For i = i To UBound(arr)
        If TypeName(f) = "Atom" Then
            If f.GetAddr = VBA.CLng(AddressOf fnullMark) Then
                f.FastApply stat, Array(stat, arr(i))
                f.Pop
            Else
                f.FastApply stat, stat, arr(i)
                'f.Pop.Pop
            End If
        Else
            f.FastApply stat, stat, arr(i)
        End If
    Next i

    If IsObject(stat) Then
        Set FoldImpl = stat
    Else
        Let FoldImpl = stat
    End If
End Function

Public Function Scan(ByVal f As Variant) As Atom
    Select Case TypeName(f)
        Case Is = "Func", "Atom", "Long"
            If TypeName(f) = "Long" Then Set f = Init(New Func, vbVariant, f, vbVariant, vbVariant)
            Dim a As New Atom: Set Scan = a.AddFunc(Init(New Func, vbVariant, AddressOf ScanImpl, vbObject, vbVariant, vbVariant)).Bind(f)
            a.LetAddr = VBA.CLng(AddressOf ScanMark)
        Case Else
            Err.Raise 13
    End Select
End Function

Public Function ScanMark() As Byte
End Function

Public Function ScanImpl( _
    ByVal f As Object, ByVal arr As Variant, Optional ByVal seedv As Variant _
    ) As Variant

    If Not IsArray(arr) Then Err.Raise 13

    Dim isObj As Boolean
    Dim stat As Variant
    Dim i As Long: i = LBound(arr)
    ArrFoldPrep arr, seedv, i, stat, isObj

    Dim stats As ArrayEx: Set stats = New ArrayEx

    If isObj Then
        stats.AddObj stat
        For i = i To UBound(arr)
            If TypeName(f) = "Atom" Then
                If f.GetAddr = VBA.CLng(AddressOf fnullMark) Then
                    f.FastApply stat, Array(stat, arr(i))
                    f.Pop
                Else
                    f.FastApply stat, stat, arr(i)
                    'f.Pop.Pop
                End If
            Else
                f.FastApply stat, stat, arr(i)
            End If
            stats.AddObj stat
        Next
    Else
        stats.addval stat
        For i = i To UBound(arr)
            If TypeName(f) = "Atom" Then
                If f.GetAddr = VBA.CLng(AddressOf fnullMark) Then
                    f.FastApply stat, Array(stat, arr(i))
                    f.Pop
                Else
                    f.FastApply stat, stat, arr(i)
                    f.Pop.Pop
                End If
            Else
                f.FastApply stat, stat, arr(i)
            End If
            stats.addval stat
        Next
    End If

    ScanImpl = stats.ToArray
End Function

Public Function Zip(ByVal f As Variant) As Atom
    Select Case TypeName(f)
        Case Is = "Func", "Atom", "Long"
            If TypeName(f) = "Long" Then Set f = Init(New Func, vbObject, f, vbVariant, vbVariant)
            Dim a As New Atom: Set Zip = a.AddFunc(Init(New Func, vbVariant, AddressOf ZipImpl, vbObject, vbVariant, vbVariant)).Bind(f)
            a.LetAddr = VBA.CLng(AddressOf ZipMark)
        Case Else
            Err.Raise 13
    End Select
End Function

Public Function ZipMark() As Byte
End Function

Public Function ZipImpl( _
    ByVal f As Object, ByVal arr1 As Variant, ByVal arr2 As Variant _
    ) As Variant

    If Not (IsArray(arr1) And IsArray(arr2)) Then Err.Raise 13
    Dim lb1 As Long: lb1 = LBound(arr1)
    Dim lb2 As Long: lb2 = LBound(arr2)
    Dim ub0 As Long: ub0 = UBound(arr1) - lb1
    If ub0 <> UBound(arr2) - lb2 Then Err.Raise 5
    Dim ret As Variant
    If ub0 < 0 Then
        ret = Array()
        GoTo Ending
    End If

    ReDim ret(ub0)

    Dim i As Long
    For i = 0 To ub0
        If TypeName(f) = "Atom" Then
            If f.GetAddr = VBA.CLng(AddressOf fnullMark) Then
                f.FastApply ret(i), Array(arr1(lb1 + i), arr2(lb2 + i))
                f.Pop.PopPop
            Else
                f.FastApply ret(i), arr1(lb1 + i), arr2(lb2 + i)
                f.Pop
            End If
        Else
            f.FastApply ret(i), arr1(lb1 + i), arr2(lb2 + i)
        End If
    Next

Ending:
    ZipImpl = ret
End Function

Public Function Unfold(ByVal f As Variant) As Atom
    Select Case TypeName(f)
        Case Is = "Func", "Atom", "Long"
            If TypeName(f) = "Long" Then Set f = Init(New Func, vbVariant, f, vbObject)
            Dim a As New Atom: Set Unfold = a.AddFunc(Init(New Func, vbVariant, AddressOf ArrUnfold, vbObject, vbVariant)).Bind(f)
        Case Else
            Err.Raise 13
    End Select
End Function

Public Function GroupBy(ByVal f As Variant) As Atom
    Select Case TypeName(f)
        Case Is = "Func", "Atom", "Long"
            If TypeName(f) = "Long" Then Set f = Init(New Func, vbVariant, f, vbObject)
            Dim a As New Atom: Set GroupBy = a.AddFunc(Init(New Func, vbVariant, AddressOf ArrGroupBy, vbObject, vbVariant)).Bind(f)
        Case Else
            Err.Raise 13
    End Select
End Function


Private Sub ArrFoldRPrep( _
    arr As Variant, seedv As Variant, i As Long, stat As Variant, _
    Optional isObj As Boolean _
    )

    If IsObject(seedv) Then
        Set stat = seedv
    Else
        Let stat = seedv
    End If

    If IsMissing(stat) Then
        isObj = IsObject(arr(i))
        If isObj Then
            Set stat = arr(i)
        Else
            Let stat = arr(i)
        End If
        i = i - 1
    End If
End Sub

Public Function FoldR(ByVal f As Variant) As Atom
    Select Case TypeName(f)
        Case Is = "Func", "Atom", "Long"
            If TypeName(f) = "Long" Then Set f = Init(New Func, vbVariant, f, vbVariant, vbVariant)
            Dim a As New Atom: Set FoldR = a.AddFunc(Init(New Func, vbVariant, AddressOf FoldRImpl, vbObject, vbVariant, vbVariant)).Bind(f)
            a.LetAddr = VBA.CLng(AddressOf FoldRMark)
        Case Else
            Err.Raise 13
    End Select
End Function

Public Function FoldRMark() As Byte
End Function

Private Function FoldRImpl( _
    ByVal f As Object, ByVal arr As Variant, Optional ByVal seedv As Variant _
    ) As Variant

    If Not IsArray(arr) Then Err.Raise 13

    Dim stat As Variant
    Dim i As Long: i = UBound(arr)
    ArrFoldRPrep arr, seedv, i, stat

    For i = i To 0 Step -1
        If TypeName(f) = "Atom" Then
            If f.GetAddr = VBA.CLng(AddressOf fnullMark) Then
                f.FastApply stat, Array(stat, arr(i))
                f.Pop
            Else
                f.FastApply stat, stat, arr(i)
            End If
        Else
            f.FastApply stat, stat, arr(i)
        End If
    Next i

    If IsObject(stat) Then
        Set FoldRImpl = stat
    Else
        Let FoldRImpl = stat
    End If
End Function

Public Function ScanR(ByVal f As Variant) As Atom
    Select Case TypeName(f)
        Case Is = "Func", "Atom", "Long"
            If TypeName(f) = "Long" Then Set f = Init(New Func, vbVariant, f, vbVariant, vbVariant)
            Dim a As New Atom: Set ScanR = a.AddFunc(Init(New Func, vbVariant, AddressOf ScanRImpl, vbObject, vbVariant, vbVariant)).Bind(f)
            a.LetAddr = VBA.CLng(AddressOf ScanRMark)
        Case Else
            Err.Raise 13
    End Select
End Function

Public Function ScanRMark() As Byte
End Function

Private Function ScanRImpl( _
    ByVal f As Object, ByVal arr As Variant, Optional ByVal seedv As Variant _
    ) As Variant

    If Not IsArray(arr) Then Err.Raise 13

    Dim isObj As Boolean
    Dim stat As Variant
    Dim i As Long: i = UBound(arr)
    ArrFoldRPrep arr, seedv, i, stat, isObj

    Dim stats As ArrayEx: Set stats = New ArrayEx

    If isObj Then
        stats.AddObj stat
        For i = i To 0 Step -1
            If TypeName(f) = "Atom" Then
                If f.GetAddr = VBA.CLng(AddressOf fnullMark) Then
                    f.FastApply stat, Array(stat, arr(i))
                    f.Pop
                Else
                    f.FastApply stat, stat, arr(i)
                    'f.Pop.Pop
                End If
            Else
                f.FastApply stat, stat, arr(i)
            End If
            stats.AddObj stat
        Next
    Else
        stats.addval stat
        For i = i To 0 Step -1
            If TypeName(f) = "Atom" Then
                If f.GetAddr = VBA.CLng(AddressOf fnullMark) Then
                    f.FastApply stat, Array(stat, arr(i))
                    f.Pop
                Else
                    f.FastApply stat, stat, arr(i)
                    f.Pop.Pop
                End If
            Else
                f.FastApply stat, stat, arr(i)
            End If
            stats.addval stat
        Next
    End If

    ScanRImpl = stats.ToArray
End Function

Public Function Reject(ByVal f As Variant) As Atom
    Select Case TypeName(f)
        Case Is = "Func", "Atom", "Long"
            If TypeName(f) = "Long" Then Set f = Init(New Func, vbBoolean, f, vbVariant)
            Set Reject = Filter(Complement(f))
        Case Else
            Err.Raise 13
    End Select
End Function

Private Function Complement(ByVal pred As Variant) As Atom
    Dim a As New Atom: Set Complement = a.AddFunc(Init(New Func, vbBoolean, AddressOf ComplementImpl, vbVariant, vbVariant)).Bind(pred)
    a.LetAddr = VBA.CLng(AddressOf ComplementMark)
End Function

Public Function ComplementMark() As Byte
End Function

Private Function ComplementImpl(ByVal prd As Variant, ByVal arg As Variant) As Boolean
    ComplementImpl = Not prd.Apply(arg)
End Function


Public Function All(ByVal fun As Variant) As Atom
    Select Case TypeName(fun)
        Case Is = "Func", "Atom", "Long"
            If TypeName(fun) = "Long" Then Set fun = Init(New Func, vbBoolean, fun, vbVariant)
            Dim a As New Atom: Set All = a.AddFunc(Init(New Func, vbBoolean, AddressOf AllImpl, vbVariant, vbVariant)).Bind(fun)
        Case Else
            Err.Raise 13
    End Select
End Function

Private Function AllImpl(ByVal fun As Variant, ByVal arr As Variant) As Boolean

    If Not IsArray(arr) Then Err.Raise 13

    Dim v
    For Each v In arr
        If fun.Apply(v) <> True Then
            AllImpl = False
            GoTo Escape
        End If
    Next v

    AllImpl = True

Escape:
End Function

Public Function Any_(ByVal fun As Variant) As Atom
    Select Case TypeName(fun)
        Case Is = "Func", "Atom", "Long"
            If TypeName(fun) = "Long" Then Set fun = Init(New Func, vbBoolean, fun, vbVariant)
            Dim a As New Atom: Set Any_ = a.AddFunc(Init(New Func, vbBoolean, AddressOf AnyImpl, vbVariant, vbVariant)).Bind(fun)
        Case Else
            Err.Raise 13
    End Select
End Function

Private Function AnyImpl(ByVal fun As Variant, ByVal arr As Variant) As Boolean

    If Not IsArray(arr) Then Err.Raise 13

    Dim v
    For Each v In arr
        If fun.Apply(v) = True Then
            AnyImpl = True
            GoTo Escape
        End If
    Next v

Escape:
End Function

Public Function Find(ByVal fun As Variant) As Atom
    Select Case TypeName(fun)
        Case Is = "Func", "Atom", "Long"
            If TypeName(fun) = "Long" Then Set fun = Init(New Func, vbBoolean, fun, vbVariant)
            Dim a As New Atom: Set Find = a.AddFunc(Init(New Func, vbVariant, AddressOf FindImpl, vbVariant, vbVariant)).Bind(fun)
        Case Else
            Err.Raise 13
    End Select
End Function

Private Function FindImpl(ByVal f As Variant, ByVal arr As Variant) As Variant

    If Not IsArray(arr) Then Err.Raise 13

    Dim v
    For Each v In arr
        If f.Apply(v) = True Then
            If IsObject(v) Then
                Set FindImpl = v
                GoTo Escape
            Else
                Let FindImpl = v
                GoTo Escape
            End If
        End If
    Next v

Escape:
End Function

Public Function AllOf(ParamArray funs() As Variant) As Boolean

    Dim arrx As New ArrayEx
    If UBound(funs) <> -1 Then
        Dim v
        For Each v In funs: arrx.AddObj v: Next v
    End If
    arrx.addval True

    AllOf = Flow(arrx.ToArray, FoldR(Init(New Func, vbBoolean, AddressOf AllOfImpl, vbBoolean, vbVariant)))

End Function

Private Function AllOfImpl(ByVal truth As Boolean, ByVal fun As Variant) As Boolean
    AllOfImpl = truth And fun.Apply
End Function

Public Function AnyOf(ParamArray funs() As Variant) As Boolean

    Dim arrx As New ArrayEx
    If UBound(funs) <> -1 Then
        Dim v
        For Each v In funs: arrx.AddObj v: Next v
    End If
    arrx.addval False

    AnyOf = Flow(arrx.ToArray, FoldR(Init(New Func, vbBoolean, AddressOf AnyOfImpl, vbBoolean, vbVariant)))

End Function

Private Function AnyOfImpl(ByVal truth As Boolean, ByVal fun As Variant) As Boolean
    AnyOfImpl = truth Or fun.Apply
End Function

'About Type converter
Public Function ToByte(ByVal n As Variant) As Variant
    ToByte = CByte(n)
End Function

Public Function ToInt(ByVal n As Variant) As Variant
    ToInt = CInt(n)
End Function

Public Function ToLng(ByVal n As Variant) As Variant
    ToLng = CLng(n)
End Function

Public Function ToSng(ByVal n As Variant) As Variant
    ToSng = CSng(n)
End Function

Public Function ToDbl(ByVal n As Variant) As Variant
    ToDbl = CDbl(n)
End Function

Public Function ToCur(ByVal n As Variant) As Variant
    ToCur = CCur(n)
End Function

Public Function ToVar(ByVal n As Variant) As Variant
    ToVar = CVar(n)
End Function

Public Function ArrToByte(ByVal arr As Variant) As Variant
    ArrToByte = ArrMap(Init(New Func, vbVariant, AddressOf ToByte, vbVariant), arr)
End Function

Public Function ArrToInt(ByVal arr As Variant) As Variant
    ArrToInt = ArrMap(Init(New Func, vbVariant, AddressOf ToInt, vbVariant), arr)
End Function

Public Function ArrToSng(ByVal arr As Variant) As Variant
    ArrToSng = ArrMap(Init(New Func, vbVariant, AddressOf ToSng, vbVariant), arr)
End Function

Public Function ArrToDbl(ByVal arr As Variant) As Variant
    ArrToDbl = ArrMap(Init(New Func, vbVariant, AddressOf ToDbl, vbVariant), arr)
End Function

Public Function ArrToCur(ByVal arr As Variant) As Variant
    ArrToCur = ArrMap(Init(New Func, vbVariant, AddressOf ToCur, vbVariant), arr)
End Function


'About Array operation
Public Function PopA() As Func
    Set PopA = Init(New Func, vbVariant, AddressOf ArrPop, vbVariant)
End Function

Public Function UnShiftA() As Func
    Set UnShiftA = Init(New Func, vbVariant, AddressOf ArrUnshift, vbVariant)
End Function

Public Function PushA() As Func
    Set PushA = Init(New Func, vbVariant, AddressOf ArrPush, vbVariant, vbVariant)
End Function

Public Function ShiftA() As Func
    Set ShiftA = Init(New Func, vbVariant, AddressOf ArrShift2, vbVariant, vbVariant)
End Function

Private Function ArrShift2(ByVal val As Variant, ByVal arr As Variant) As Variant
    Dim clct As New collection: Set clct = ArrToClct(arr)
    Shift clct, val
    ArrShift2 = ClctToArr(clct)
End Function

Public Function ArrSlice2(ByVal arr As Variant, ByVal First As Long, ByVal last As Variant, Optional ByVal n As Long = 11) As Variant
    ArrSlice2 = ArrSlice(ArrFill(arr, n), First, last)
End Function


Public Function First(ByVal arr As Variant) As Variant
    If IsObject(nth(0, arr)) Then
        Set First = nth(0, arr)
    Else
        Let First = nth(0, arr)
    End If
End Function

Public Function Second(ByVal arr As Variant) As Variant
    If IsObject(nth(1, arr)) Then
        Set Second = nth(1, arr)
    Else
        Let Second = nth(1, arr)
    End If
End Function

Public Function Existy(ByVal x As Variant) As Variant

    If IsObject(x) Then
        If x Is Nothing Then
            Existy = False
        Else
            Existy = True
        End If
    Else
        If IsNull(x) Then
            Existy = False
        ElseIf IsEmpty(x) Then
            Existy = False
        ElseIf IsMissing(x) Then
            Existy = False
        Else
            Existy = True
        End If
    End If

End Function

Public Function Truthy(ByVal x As Variant) As Variant
    Select Case IsObject(x)
        Case True:  Truthy = IIf(Existy(x), True, False)
        Case False: Truthy = IIf((x <> False) And Existy(x), True, False)
    End Select
End Function

Public Function Always(ByVal x As Variant) As Atom
    Dim a As New Atom: Set Always = a.AddFunc(Init(New Func, vbVariant, AddressOf AlwaysImpl, vbVariant)).Bind(x): a.LetAddr = VBA.CLng(AddressOf AlwaysImpl)
End Function

Public Function AlwaysImpl(ByVal x As Variant) As Variant
    If IsObject(x) Then
        Set AlwaysImpl = x
    Else
        Let AlwaysImpl = x
    End If
End Function

Public Function Repeat(ByVal n As Variant, ByVal f As Variant) As Variant
    Dim v As Variant, arrx As New ArrayEx
    For Each v In ArrRange(1, n)
        Select Case TypeName(f)
            Case "Func", "Atom"
                arrx.addval f.Apply
            Case Else
                arrx.addval f
        End Select
    Next v
    Repeat = arrx.ToArray
End Function

Public Function NLV() As Func
    Set NLV = Init(New Func, vbVariant, AddressOf NLVImpl, vbVariant, vbVariant)
End Function

Public Function NLVImpl(ByVal AltArg As Variant, ByVal arg As Variant) As Variant
    NLVImpl = IIf(Existy(arg), arg, AltArg)
End Function

Public Function Array2Of() As Atom
    Dim a As New Atom: Set Array2Of = a.AddFunc(Init(New Func, vbVariant, AddressOf Array2OfImpl, vbVariant, vbVariant))
End Function

Private Function Array2OfImpl(ByVal x As Variant, ByVal y As Variant) As Variant
    Array2OfImpl = Array(x, y)
End Function

Public Function ToStr2(ByVal x As Variant) As Variant
    ToStr2 = ToStr(x)
End Function

Public Function Add() As Func
    Set Add = Init(New Func, vbVariant, AddressOf AddImpl, vbVariant, vbVariant)
End Function

Public Function AddImpl(ByVal x As Variant, ByVal y As Variant) As Variant
    AddImpl = x + y
End Function

Public Function Minus() As Func
    Set Minus = Init(New Func, vbVariant, AddressOf MinusImpl, vbVariant, vbVariant)
End Function

Public Function MinusImpl(ByVal x As Variant, ByVal y As Variant) As Variant
    MinusImpl = x - y
End Function

Public Function Mult() As Func
    Set Mult = Init(New Func, vbVariant, AddressOf MultImpl, vbVariant, vbVariant)
End Function

Public Function MultImpl(ByVal x As Variant, ByVal y As Variant) As Variant
    MultImpl = x * y
End Function

Public Function div() As Func
    Set div = Init(New Func, vbVariant, AddressOf DivImpl, vbVariant, vbVariant)
End Function

Public Function DivImpl(ByVal x As Variant, ByVal y As Variant) As Variant
    DivImpl = x / y
End Function

Public Function DivR() As Func
    Set DivR = Init(New Func, vbVariant, AddressOf DivRImpl, vbVariant, vbVariant)
End Function

Public Function DivRImpl(ByVal x As Variant, ByVal y As Variant) As Variant
    DivRImpl = y / x
End Function

Public Function IsOdd() As Func
    Set IsOdd = Init(New Func, vbBoolean, AddressOf IsOddImpl, vbVariant)
End Function

Public Function IsOddImpl(ByVal n As Variant) As Boolean
    IsOddImpl = n Mod 2 = 1
End Function

Public Function IsEven() As Func
    Set IsEven = Init(New Func, vbBoolean, AddressOf IsEvenImple, vbVariant)
End Function

Public Function IsEvenImple(ByVal n As Variant) As Boolean
    IsEvenImple = n Mod 2 = 0
End Function

'About Tuple operation
Public Function GetItem1(ByVal tpl As Tuple) As Variant
    GetItem1 = tpl.Item1
End Function

Public Function GetItem2(ByVal tpl As Tuple) As Variant
    GetItem2 = tpl.Item2
End Function

Public Function GetItem3(ByVal tpl As Tuple) As Variant
    GetItem3 = tpl.Item3
End Function

Public Function GetItem4(ByVal tpl As Tuple) As Variant
    GetItem4 = tpl.Item4
End Function

Public Function TplPlus(ByVal tpl As Tuple) As Variant
    TplPlus = ToCur(tpl.Item1) + ToCur(tpl.Item2)
End Function

Public Function TplMinus(ByVal tpl As Tuple) As Variant
    TplMinus = ToCur(tpl.Item1) - ToCur(tpl.Item2)
End Function

Public Function Tuple2Of(ByVal itm1 As Variant, ByVal itm2 As Variant) As Tuple
    Set Tuple2Of = Init(New Tuple, itm1, itm2)
End Function

Public Function ZipKey(ByVal arr As Variant, ByVal aValue As Variant) As Variant
    Dim v As Variant
    For Each v In arr
        If Equals(v(1), aValue) Then
            ZipKey = v(0)
            GoTo Escape
        End If
    Next v
Escape:
End Function

Public Function ZipValue(ByVal arr As Variant, ByVal aKey As Variant) As Variant
    Dim v As Variant
    For Each v In arr
        If Equals(v(0), aKey) Then
            ZipValue = v(1)
            GoTo Escape
        End If
    Next v
Escape:
End Function


'About Actions
Public Function Actions(ParamArray funcs() As Variant) As Atom
    Dim a As New Atom: Set Actions = a.AddFunc(Init(New Func, vbVariant, AddressOf ActionsImpl, vbVariant, vbVariant)).Bind(funcs)
End Function

Private Function ActionsImpl(ByVal fArgs As Variant, ByVal seed As Variant) As Variant

    Dim tpl As New Tuple
    If TypeName(seed) <> "Tuple" Then
        Set tpl = Init(New Tuple, seed, seed)
    End If

    Dim Status As Variant
    Dim Values As New ArrayEx

    Dim v As Variant, v2 As Variant
    For Each v In fArgs
        On Error Resume Next
            Set tpl = v.Apply(tpl)
            If Err.Number = 0 Then
                Status = tpl.Item1
                If IsArray(tpl.Item2) Then
                    Dim tmpEx As New ArrayEx
                    For Each v2 In tpl.Item2
                        tmpEx.addval v2
                    Next v2
                    Values.addval tmpEx.ToArray
                    Set tmpEx = Nothing
                Else
                    Values.addval tpl.Item2
                End If
            ElseIf Err.Number = 5 Then
                Values.addval Missing
            Else
                Err.Raise 13
            End If
        On Error GoTo 0
    Next v

    Set ActionsImpl = Init(New Tuple, Status, Values.ToArray)

End Function

Public Function Lift(ByVal AnserFun As Variant, Optional ByVal StateFun As Variant) As Atom
    If IsEmpty(StateFun) Then: StateFun = Missing
    Dim a As New Atom: Set Lift = a.AddFunc(Init(New Func, vbObject, AddressOf LiftImpl, vbVariant, vbVariant, vbObject)).Bind(AnserFun).Bind(StateFun)
End Function

Private Function LiftImpl(ByVal AnserFun As Variant, ByVal StateFun As Variant, ByVal tpl As Tuple) As Tuple

    Dim ans As Variant: ans = AnserFun.Apply(tpl.Item1)

    Dim stat As Variant
    If Not IsMissing(StateFun) Then
        stat = StateFun.Apply(ArrPop(tpl.Item2))
    Else
        stat = ans
    End If

    Set LiftImpl = Init(New Tuple, stat, ans)

End Function

Public Function EliminateMissing(ByVal arr As Variant) As Variant
    Dim v As Variant, arrx As New ArrayEx
    For Each v In arr
        If Not IsMissing(v) Then
            If IsObject(v) Then arrx.AddObj v Else arrx.addval v
        End If
    Next v
    EliminateMissing = arrx.ToArray
End Function

Public Function Checker(ParamArray Preds() As Variant) As Atom
    Dim a As New Atom: Set Checker = a.AddFunc(Init(New Func, vbVariant, AddressOf CheckerImpl, vbVariant, vbVariant)).Bind(Preds)
End Function

Private Function CheckerImpl(ByVal prds As Variant, ByVal x As Variant) As Variant
    Dim f, tmp As Variant, arrx As New ArrayEx
    For Each f In prds
            tmp = f.Apply(x)
            If tmp <> True Then arrx.addval tmp
    Next f
    CheckerImpl = arrx.ToArray
End Function

Public Function Validator(ByVal message As String, ByVal pred As Variant) As Atom
    Dim a As New Atom
    Set Validator = a.AddFunc(Init(New Func, vbVariant, AddressOf ValidatorImpl, vbVariant, vbVariant, vbVariant)).Bind(message).Bind(pred)
End Function

Private Function ValidatorImpl(ByVal msg As Variant, ByVal prd As Variant, ByVal x As Variant) As Variant
    Dim bln As Boolean: bln = prd.Apply(x)
    If bln Then
        ValidatorImpl = True
    Else
        ValidatorImpl = msg
    End If
End Function

Public Function IsNumber(ByVal x As Variant) As Boolean
    IsNumber = IsNumeric(x)
End Function

Public Function IsString(ByVal x As Variant) As Boolean
    IsString = TypeName(x) = "String"
End Function


