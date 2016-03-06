Sub Sample_KoffeePuzzle20150821()
    
    Dim v
    
    'パワーセット（Powerset）
    For Each v In PowerSet(Array("A", "B", "C"))
        Debug.Print Dump(v)
    Next v
    
    'Array()
    'Array("C")
    'Array("B")
    'Array("B", "C")
    'Array("A")
    'Array("A", "C")
    'Array("A", "B")
    'Array("A", "B", "C")
    
    
    '組み合わせ（Combinations）
    For Each v In Combin(Array("A", "B", "C"), 2)
        Debug.Print Dump(v)
    Next v
    
    'Array("B", "C")
    'Array("A", "C")
    'Array("A", "B")
    
    
    '順列（Permutations）
    For Each v In Permut(Array("A", "B", "C"), 3)
        Debug.Print Dump(v)
    Next v
    
    'Array("C", "B", "A")
    'Array("C", "A", "B")
    'Array("B", "C", "A")
    'Array("B", "A", "C")
    'Array("A", "C", "B")
    'Array("A", "B", "C")
    
    
    '順列（Permutations）型キツキツバージョン（ちょっと速い）
    Dim clct As New collection: PermutLimited 1, 3, 3, clct
    For Each v In clct
        Debug.Print Dump(ArrCLng(v))
    Next v
    
    'Array(3&, 2&, 1&)
    'Array(3&, 1&, 2&)
    'Array(2&, 3&, 1&)
    'Array(2&, 1&, 3&)
    'Array(1&, 3&, 2&)
    'Array(1&, 2&, 3&)
    
    
    '重複順列（Repeated Permutations）
    For Each v In ReptPermut(Array("A", "B", "C"), 2)
        Debug.Print Dump(v)
    Next v
    
    'Array("A", "A")
    'Array("A", "B")
    'Array("A", "C")
    'Array("B", "A")
    'Array("B", "B")
    'Array("B", "C")
    'Array("C", "A")
    'Array("C", "B")
    'Array("C", "C")


 'ビット演算（Bit Operations）
    
    Debug.Print Dump(BitFlag2(Array(1, 0, 0, 1, 1, 1, 0, 0)))
    '156&
    
    Debug.Print Dump(Dec2bin(156))
    'Array("1", "0", "0", "1", "1", "1", "0", "0")
    
    Debug.Print Dump((Dec2bin(-156)))
    'Array(0%, 1%, 1%, 0%, 0%, 1%, 0%, 0%)
    
    Debug.Print Dump(BitAnd(Array(1, 0, 1), Array(1, 1, 0, 0)))
    'Array(0%, 1%, 0%, 0%)
    
    Debug.Print Dump(BitOr(Array(1, 0, 1), Array(1, 1, 0, 0)))
    'Array(1%, 1%, 0%, 1%)

    Debug.Print Dump(BitXor(Array(1, 0, 1), Array(1, 1, 0, 0)))
    'Array(1%, 0%, 0%, 1%)

    Debug.Print Dump(BitNOT(Array(1, 0, 1)))
    'Array(0%, 1%, 0%)
    
    Debug.Print Dump(BitRShift(Array(1, 0, 1, 1), 1))
    'Array(0%, 1%, 0%, 1%)
    
    Debug.Print Dump(BitLShift(Array(1, 0, 1, 1), 1))
    'Array(0%, 1%, 1%, 0%)
    
    Debug.Print Dump(BitPlus(Dec2bin(1), Dec2bin(2)))
    'Array("1", "1")
    
    Debug.Print Dump(BitMinus(Dec2bin(3), Dec2bin(1)))
    'Array("1", "0")
    
    Debug.Print Dump(BitMult(Dec2bin(3), Dec2bin(2)))
    'Array("1", "1", "0")
    
    Debug.Print Dump(BitDiv(Dec2bin(10), Dec2bin(5)))
    'Array("1", "0")
    
End Sub
