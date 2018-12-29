Attribute VB_Name = "KoffeePuzzle"
'  +--------------                                         --------------+
'  |||||||||    Koffee2 0.1.0                                            |
'  |: ^_^ :|    Koffee2 is free Library based on Ariawase.               |
'  |||||||||    The Project Page: https://github.com/CallMeKohei/Koffee2 |
'  +--------------                                         --------------+
Option Explicit

'PowerSet and Combinations
Public Function PowerSet(ByVal arr As Variant, Optional ByVal R As Long = 0) As Variant

    Dim a(): ReDim a(32)
    Dim ub As Long:  ub = 32
    Dim i As Long: i = 0

    Dim v, T, j As Long
    For Each v In PowerSetImpl(UBound(arr) + 1, R)
        T = Array()
        For j = 0 To UBound(arr)
            If v(j) <> 0 Then
                T = ArrPush(arr(j), T)
            End If
        Next j

        If ub = i Then
            ub = ub + 1
            ub = -1 + ub + ub
            ReDim Preserve a(ub - 1)
        End If

        a(i) = T
        i = i + 1

    Next v

    PowerSet = Truncate(a)

End Function

Public Function PowerSetImpl(ByVal n As Variant, Optional R As Long = 0) As Variant

    Dim arr(): ReDim arr(32)
    Dim ub As Long:  ub = 32
    Dim i As Long: i = 0

    Dim v, bitflg
    For Each v In ArrRange(0, (2 ^ n) - 1)
        bitflg = Dec2bin(v, n, R)
        If UBound(bitflg) <> -1 Then

            If ub = i Then
                ub = ub + 1
                ub = -1 + ub + ub
                ReDim Preserve arr(ub - 1)
            End If

            arr(i) = bitflg
            i = i + 1

        End If
    Next v

    PowerSetImpl = Truncate(arr)

End Function

Public Function Combin(ByVal arr As Variant, Optional ByVal R As Long = 0) As Variant
    If R <= 0 Or R > UBound(arr) + 1 Then R = UBound(arr) + 1
    Combin = PowerSet(arr, R)
End Function

'Permutations
Public Function Permut(ByVal arr As Variant, ByVal R As Long) As collection
    Dim clct As New collection
    PermutImpl arr, R, clct
    Set Permut = clct
End Function

Private Sub PermutImpl(ByVal a As Variant, ByVal R As Long, ByRef clct As collection)

    Dim ub           As Byte: ub = UBound(a)
    Dim leafLevel    As Byte: If R = 0 Then leafLevel = ub Else leafLevel = R - 1
    Dim arr()        As Variant: ReDim arr(leafLevel)
    Dim currentLevel As Long
    Dim stackPointer As Long
    Dim stack(99)    As Variant
    Dim level(99)    As Byte
    Dim used         As Object: Set used = CreateObject("Scripting.Dictionary")
    Dim i            As Byte

    Dim v
    For Each v In a
        stackPointer = stackPointer + 1
        stack(stackPointer) = v
        level(stackPointer) = 0
        used.Add key:=v, Item:=False
    Next v

    Do While stackPointer > 0

        currentLevel = level(stackPointer): level(stackPointer) = 0
        arr(currentLevel) = stack(stackPointer): stack(stackPointer) = 0
        stackPointer = stackPointer - 1

        If used(arr(currentLevel)) = True Then
            currentLevel = currentLevel - 1
        Else
            used(arr(currentLevel)) = True
            If currentLevel = leafLevel Then
                clct.Add arr
                If stackPointer > 0 Then
                    For i = level(stackPointer) To leafLevel
                        used(arr(i)) = False
                    Next i
                End If
            Else
                For Each v In a
                    If used(v) = False Then
                        stackPointer = stackPointer + 1
                        stack(stackPointer) = v
                        level(stackPointer) = currentLevel + 1
                    End If
                Next v
            End If
        End If
    Loop

End Sub

Public Sub PermutLimited(ByVal fst As Byte, ByVal lst As Byte, ByVal R As Byte, Optional ByRef clct As collection)

    Dim ub           As Byte: ub = lst - fst
    Dim leafLevel    As Byte: If R = 0 Then leafLevel = ub Else leafLevel = R - 1
    Dim arr()        As Byte: ReDim arr(leafLevel)
    Dim currentLevel As Long
    Dim stackPointer As Long
    Dim stack(99)    As Byte
    Dim level(99)    As Byte
    Dim used()       As Byte: ReDim used(lst)
    Dim i            As Byte

    For i = fst To lst
        stackPointer = stackPointer + 1
        stack(stackPointer) = i
        level(stackPointer) = 0
    Next i

    Do While stackPointer > 0

        currentLevel = level(stackPointer): level(stackPointer) = 0
        arr(currentLevel) = stack(stackPointer): stack(stackPointer) = 0
        stackPointer = stackPointer - 1

        If used(arr(currentLevel)) = True Then
            currentLevel = currentLevel - 1
        Else
            used(arr(currentLevel)) = True
            If currentLevel = leafLevel Then
                clct.Add arr
                If stackPointer > 0 Then
                    For i = level(stackPointer) To leafLevel
                        used(arr(i)) = False
                    Next i
                End If
            Else
                For i = fst To lst
                    If used(i) = False Then
                        stackPointer = stackPointer + 1
                        stack(stackPointer) = i
                        level(stackPointer) = currentLevel + 1
                    End If
                Next i
            End If
        End If
    Loop

End Sub

'Repeated Permutations
Public Function ReptPermut(ByVal arr As Variant, Optional ByVal R As Long = 0) As Variant

    Dim v, v2, tmp, i As Long, arrx As New ArrayEx
    For Each v In ReptPermutImpl(UBound(arr) + 1, R)
        tmp = Array()
        For Each v2 In v
            tmp = ArrPush(arr(v2), tmp)
        Next v2
        arrx.addval tmp
    Next v

    ReptPermut = arrx.ToArray

End Function

Public Function ReptPermutImpl(ByVal n As Variant, Optional R As Long = 0) As Variant

    Dim v, bitflg, arrx As New ArrayEx
    For Each v In ArrRange(0, (n ^ R - 1))
        bitflg = SplitStr(Dec2N(v, n))
        If UBound(bitflg) <> -1 Then
            arrx.addval ArrCLng(ArrFill(bitflg, R - 1, , True))
        End If
    Next v

    ReptPermutImpl = arrx.ToArray

End Function

Public Function Dec2N(ByVal val As Long, ByVal n As Long) As String

    If val = 0 Then
        Dec2N = "0"
        GoTo Escape
    End If

    Dim i As Long:      i = 1
    Dim tmp As String:  tmp = ""

    Do While (val >= i)
        tmp = Dec2NImpl((val Mod (i * n)) \ i) & tmp
        i = i * n
    Loop
    Dec2N = tmp

Escape:
End Function

Private Function Dec2NImpl(ByVal val As Long) As String

    If val < 10 Then
        Dec2NImpl = CStr(val)
    Else
        Dec2NImpl = Chr(65 + val - 10)
    End If

End Function

'bit Operations
Public Function Dec2bin(ByVal val As Long, Optional ByVal n As Long = 0, Optional ByVal R As Long = 0) As Variant
    If val < 0 Then
        Dec2bin = BitNOT(SplitStr(Dec2BinImpl(Abs(val + 1), n, R)))
    Else
        Dec2bin = SplitStr(Dec2BinImpl(val, n, R))
    End If
End Function

Public Function Dec2BinImpl(ByVal val As Long, Optional ByVal n As Long = 0, Optional ByVal R As Long = 0) As String

    Dim bit As Long
    Dim tmp As String
    Dim cnt As Long

    Do Until (val < 2 ^ bit)
        If (val And 2 ^ bit) <> 0 Then
            tmp = "1" & tmp
            cnt = cnt + 1
            If R <> 0 And cnt > R Then GoTo Escape
        Else
            tmp = "0" & tmp
        End If

        bit = bit + 1
    Loop

    If tmp = "" Then tmp = 0

    If R = 0 Then
        If n = 0 Then
            Dec2BinImpl = tmp
        Else
            Dec2BinImpl = Right(tmp + 10 ^ n, n)
        End If
    Else
        If cnt = R Then
            Dec2BinImpl = Right(tmp + 10 ^ n, n)
        Else
            Dec2BinImpl = Empty
        End If
    End If

Escape:
End Function

Public Function SplitStr(ByVal str As String) As Variant

    If str = "" Then
        SplitStr = Array()
        GoTo Escape
    End If


    Dim a(): ReDim a(32)
    Dim ub As Long:  ub = 32

    Dim i As Long
    For i = 1 To Len(str)
        If ub = i Then
            ub = ub + 1
            ub = -1 + ub + ub
            ReDim Preserve a(ub - 1)
        End If

        a(i - 1) = Mid(str, i, 1)
    Next i

    SplitStr = Truncate(a)

Escape:
End Function

Public Function BitAnd(ByVal flg1 As Variant, ByVal flg2 As Variant) As Variant

    Dim ub As Long: ub = UBound(flg1)
    If UBound(flg1) > UBound(flg2) Then
        ub = UBound(flg1)
        flg2 = ArrFill(flg2, ub, , True)
    ElseIf UBound(flg1) < UBound(flg2) Then
        ub = UBound(flg2)
        flg1 = ArrFill(flg1, ub, , True)
    End If

    Dim i As Long, arrx As New ArrayEx
    For i = 0 To ub
        arrx.addval IIf(CLng(flg1(i)) = CLng(flg2(i)) And CLng(flg1(i)) = 1, 1, 0)
    Next i
    BitAnd = arrx.ToArray

End Function

Public Function BitOr(ByVal flg1 As Variant, ByVal flg2 As Variant) As Variant

    Dim ub As Long: ub = UBound(flg1)
    If UBound(flg1) > UBound(flg2) Then
        ub = UBound(flg1)
        flg2 = ArrFill(flg2, ub, , True)
    ElseIf UBound(flg1) < UBound(flg2) Then
        ub = UBound(flg2)
        flg1 = ArrFill(flg1, ub, , True)
    End If

    Dim i As Long, arrx As New ArrayEx
    For i = 0 To UBound(flg1)
        arrx.addval IIf(CLng(flg1(i)) = CLng(flg2(i)) And CLng(flg1(i)) = 0, 0, 1)
    Next i
    BitOr = arrx.ToArray

End Function

Public Function BitXor(ByVal flg1 As Variant, ByVal flg2 As Variant) As Variant

    Dim ub As Long: ub = UBound(flg1)
    If UBound(flg1) > UBound(flg2) Then
        ub = UBound(flg1)
        flg2 = ArrFill(flg2, ub, , True)
    ElseIf UBound(flg1) < UBound(flg2) Then
        ub = UBound(flg2)
        flg1 = ArrFill(flg1, ub, , True)
    End If

    Dim i As Long, arrx As New ArrayEx
    For i = 0 To UBound(flg1)
        arrx.addval IIf(CLng(flg1(i)) = CLng(flg2(i)), 0, 1)
    Next i
    BitXor = arrx.ToArray

End Function

Public Function BitNOT(ByVal flg As Variant) As Variant

    Dim i As Long, arrx As New ArrayEx
    For i = 0 To UBound(flg)
        arrx.addval IIf(CLng(flg(i)) = 1, 0, 1)
    Next i
    BitNOT = arrx.ToArray

End Function

Public Function BitFlag2(ByVal flgs As Variant) As Long
    BitFlag2 = 0
    Dim ub As Long: ub = UBound(flgs)

    Dim i As Long
    For i = 0 To ub
        BitFlag2 = BitFlag2 + Abs(flgs(i)) * 2 ^ (ub - i)
    Next
End Function

Public Function BitRShift(ByVal flg As Variant, ByVal n As Long) As Variant
    Dim i As Long
    For i = 1 To n
        flg = ArrShift(0, flg)
        flg = ArrPop(flg)
    Next i
    BitRShift = flg
End Function

Public Function BitLShift(ByVal flg As Variant, ByVal n As Long) As Variant
    Dim i As Long
    For i = 1 To n
        flg = ArrPush(0, flg)
        flg = ArrUnshift(flg)
    Next i
    BitLShift = flg
End Function

Public Function BitComplement(ByVal flg As Variant) As Variant
    BitComplement = BitPlus(BitNOT(flg), Array(1))
End Function

Public Function BitPlus(ByVal flg1 As Variant, flg2 As Variant) As Variant

    If UBound(flg1) > UBound(flg2) Then
        flg2 = ArrFill(flg2, UBound(flg1), , True)
    ElseIf UBound(flg1) < UBound(flg2) Then
        flg1 = ArrFill(flg1, UBound(flg2), , True)
    End If

    BitPlus = Dec2bin(BitFlag2(flg1) + BitFlag2(flg2))
End Function

Public Function BitMinus(ByVal flg1 As Variant, flg2 As Variant) As Variant

    If UBound(flg1) > UBound(flg2) Then
        flg2 = ArrFill(flg2, UBound(flg1), , True)
    ElseIf UBound(flg1) < UBound(flg2) Then
        flg1 = ArrFill(flg1, UBound(flg2), , True)
    End If

    BitMinus = Dec2bin(BitFlag2(flg1) - BitFlag2(flg2))
End Function

Public Function BitMult(ByVal flg1 As Variant, flg2 As Variant) As Variant

    If UBound(flg1) > UBound(flg2) Then
        flg2 = ArrFill(flg2, UBound(flg1), , True)
    ElseIf UBound(flg1) < UBound(flg2) Then
        flg1 = ArrFill(flg1, UBound(flg2), , True)
    End If

    BitMult = Dec2bin(BitFlag2(flg1) * BitFlag2(flg2))
End Function

Public Function BitDiv(ByVal flg1 As Variant, flg2 As Variant) As Variant

    If UBound(flg1) > UBound(flg2) Then
        flg2 = ArrFill(flg2, UBound(flg1), , True)
    ElseIf UBound(flg1) < UBound(flg2) Then
        flg1 = ArrFill(flg1, UBound(flg2), , True)
    End If

    BitDiv = Dec2bin(BitFlag2(flg1) / BitFlag2(flg2))
End Function
