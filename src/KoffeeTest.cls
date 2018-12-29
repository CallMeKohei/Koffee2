VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KoffeeTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'  +--------------                                         --------------+
'  |||||||||    Koffee2 0.1.0                                            |
'  |: ^_^ :|    Koffee2 is free Library based on Ariawase.               |
'  |||||||||    The Project Page: https://github.com/CallMeKohei/Koffee2 |
'  +--------------                                         --------------+
Option Explicit

#If False Then
Private Sub KoffeeTestRun()
    Assert.RunTestOf New KoffeeTest
End Sub

Public Function add1(ByVal x As Variant) As Variant
    add1 = x + 1
End Function

Public Function DivTriImpl(ByVal x As Variant, ByVal y As Variant, ByVal z As Variant) As Variant
    DivTriImpl = x / y / z
End Function

Public Function FibFun(ByVal tpl As Tuple) As Variant
    Dim lmt As Long: lmt = 100

    Dim v1 As Long: v1 = tpl.Item2
    Dim v2 As Long: v2 = tpl.Item1 + v1
    If v1 <= lmt Then
        Set FibFun = Init(New Tuple, v1, Init(New Tuple, v1, v2))
    Else
        Let FibFun = Missing
    End If
End Function

Public Function SayVBA() As Variant
    SayVBA = "VBA!"
End Function

Public Function IsOver200(ByVal arr As Variant) As Boolean
    If arr(2)(0) = "PRICE" And arr(2)(1) >= 200 Then
        IsOver200 = True
    End If
End Function

Public Function ArrJoin(ByVal arr As Variant) As Variant
    ArrJoin = Join(arr, " ")
End Function

Public Function zero() As Variant
    Dim a As New Atom: Set zero = a.AddFunc(Init(New Func, vbVariant, AddressOf zeroImpl, vbVariant, vbVariant))
End Function

Private Function zeroImpl(ByVal f As Variant, ByVal x As Variant) As Variant
    zeroImpl = f.Apply(x)
End Function

Public Function fst() As Atom
    Dim a As New Atom: Set fst = a.AddFunc(Init(New Func, vbVariant, AddressOf fstImpl, vbVariant, vbVariant))
End Function

Private Function fstImpl(ByVal f As Variant, ByVal x As Variant) As Variant
    fstImpl = f.Apply(x)
End Function

Public Function snd() As Atom
    Dim a As New Atom: Set snd = a.AddFunc(Init(New Func, vbVariant, AddressOf sndImpl, vbVariant, vbVariant))
End Function

Private Function sndImpl(ByVal f As Variant, ByVal x As Variant) As Variant
    sndImpl = f.Apply(x)
End Function

Public Function Tr() As Boolean
    Tr = True
End Function

Public Function Fa() As Boolean
    Fa = False
End Function

Public Function isGreenBlack(ByVal rcd As Object) As Boolean

    If rcd("Color") = "black" Or rcd("Color") = "green" Then
        isGreenBlack = True
    End If

End Function
#End If

Public Sub KoffeeCore_Test()

    Dim sh As Worksheet: Set sh = AddSheet("myTestTemporarySheet")

    PutVal Array("F1", "A1", "A2", "A3"), sh.Range("A1"), True
    PutVal Array("F2", "B1", "B2", "B3"), sh.Range("B1"), True
    PutVal Array("F3", "C1", "C2", "C3"), sh.Range("C1"), True

    Dim sql As String:   sql = "SELECT * FROM [myTestTemporarySheet$]"

    Dim arrq As Variant:  arrq = FetchSh(sql)

    Assert.AreEqualArr Array("A1", "A2", "A3"), arrq(0)
    Assert.AreEqualArr Array("B1", "B2", "B3"), arrq(1)
    Assert.AreEqualArr Array("C1", "C2", "C3"), arrq(2)


    Dim arrr As Variant: arrr = FetchSh(sql, , True)

    Assert.AreEqualArr Array("F1", "F2", "F3"), arrr(1)

    Application.DisplayAlerts = False
    sh.Delete
    Application.DisplayAlerts = True



    Dim jagArr1 As Variant: jagArr1 = Array(Array(1, 2, 3), Array(4, 5, 6))
    Dim jagArr2 As Variant: jagArr2 = Array(Array(1, 2, 3), Array(4, 5, 6), Array(7, 8, 9))

    Dim jagArr3 As Variant: jagArr3 = Array(Array(1), Array(2, 3, 4), Array(5, 6))
    Dim jagArr4 As Variant: jagArr4 = Array(Array(), Array(1, 2, 3), Array(4, 5))
    Dim jagArr5 As Variant: jagArr5 = Array(ArrToClct(Array()), ArrToClct(Array(0)), ArrToClct(Array(0, 0)))


    Assert.AreEqual True, IsJagArr(jagArr1), "jagArr1"
    Assert.AreEqual True, IsJagArr(jagArr2), "jagArr2"
    Assert.AreEqual True, IsJagArr(jagArr3), "jagArr3"
    Assert.AreEqual True, IsJagArr(jagArr4), "jagArr4"
    Assert.AreEqual False, IsJagArr(jagArr5), "jagArr5"

    Dim arr1D() As Variant
    Dim arr1DInt(0) As Integer
    Dim arr1DClct() As Variant: arr1DClct = Array(ArrToClct(Array(0)))
    Assert.AreEqual False, IsJagArr(arr1D), "arr1D"
    Assert.AreEqual False, IsJagArr(arr1DInt), "arr1DInt"
    Assert.AreEqual False, IsJagArr(arr1DClct), "arr1DClct"

    Dim arr3DInt(0, 0, 0) As Integer
    Dim arr3DClct() As Variant: arr3DClct = JagArrToArr2D(Array(Array(ArrToClct(Array(0))), Array(ArrToClct(Array(0))), Array(ArrToClct(Array(0)))))
    Assert.AreEqual False, IsJagArr(arr3DInt), "arr3DInt"
    Assert.AreEqual False, IsJagArr(arr3DClct), "arr3DClct"


    Assert.AreEqualArr Array(1, 2, 3), Truncate(Array(1, 2, 3, Empty, Empty, Empty))
    Assert.AreEqualArr Array(Empty, 1, 2, 3), Truncate(Array(Empty, 1, 2, 3, Empty, Empty, Empty))

    Assert.AreEqual 0&, LBound(Base01(Array(0, 1)))
    Assert.AreEqual 1&, LBound(Base01(Array(0, 1), True))

    Assert.AreEqualArr Array(1, 2, 3, 4, 5), ArrExplode(Array(1, Array(2, Array(3, Array(4))), 5))

    Assert.AreEqualArr Array(1&, 2&, 3&, 0, 0), ArrFill(Array(1&, 2&, 3&), 4)

    Dim arr As Variant: arr = Array("a", "b", "c", "d")

    Assert.AreEqual "a", nth(0, arr)
    Assert.AreEqual "b", nth(1, arr)
    Assert.AreEqual "c", nth(2, arr)
    Assert.AreEqual "d", nth(3, arr)

    Assert.AreEqualArr Array("b", "c", "d"), Rest(arr)
    Assert.AreEqualArr Array("c", "d"), Rest(Rest(arr))
    Assert.AreEqualArr Array("d"), Rest(Rest(Rest(arr)))


    Dim arrA As Variant: arrA = Array(5, 4, 3, 2, 1)
    Assert.AreEqualArr Array(1, 2, 3, 4, 5), ArrSortAsc(arrA)
    Assert.AreEqualArr Array(5, 4, 3, 2, 1), ArrSortDec(arrA)
    Assert.AreEqualArr Array(1, 2, 3, 4, 5), ArrRev2(arrA)

    Assert.AreEqualArr Array(5@, 7@, 9@), ArrPlus(Array(1, 2, 3), Array(4, 5, 6))
    Assert.AreEqualArr Array(3@, 3@, 3@), ArrMinus(Array(4, 5, 6), Array(1, 2, 3))

    Dim arr1 As Variant: arr1 = Array(1, 2, 3, "a")
    Dim arr2 As Variant: arr2 = Array(3, 4, 5, "a", "b")

    Assert.AreEqualArr Array(1, 2, 3, "a", 4, 5, "b"), ArrUnion(arr1, arr2)
    Assert.AreEqualArr Array(3, 4, 5, "a", "b", 1, 2), ArrUnion(arr2, arr1)
    Assert.AreEqualArr Array(1, 2), ArrDiff(arr1, arr2)
    Assert.AreEqualArr Array(4, 5, "b"), ArrDiff(arr2, arr1)
    Assert.AreEqualArr Array(1, 2, 4, 5, "b"), ArrDiff2(arr1, arr2)
    Assert.AreEqualArr Array(4, 5, "b", 1, 2), ArrDiff2(arr2, arr1)
    Assert.AreEqualArr Array(3, "a"), ArrIntersect(arr1, arr2)
    Assert.AreEqualArr Array(3, "a"), ArrIntersect(arr2, arr1)

    Assert.AreEqualArr Array(9, 1, 2, 3), ArrShift(9, Array(1, 2, 3))
    Assert.AreEqualArr Array(2, 3), ArrUnshift(Array(1, 2, 3))
    Assert.AreEqualArr Array(1, 2, 3, 9), ArrPush(9, Array(1, 2, 3))
    Assert.AreEqualArr Array(1, 2), ArrPop(Array(1, 2, 3))


    Assert.AreEqualArr Array(1&, 2&, 3&), ArrCLng(Array(1, 2, 3))
    Assert.AreEqualArr Array(1@, 2@, 3@), ArrCCur(Array(1, 2, 3))

    Assert.AreEqualArr Array(1@, 3@, 6@), StepTotal(Array(1, 2, 3))(0)
    Assert.AreEqualArr Array(4@, 9@, 15@), StepTotal(Array(Array(1, 2, 3), Array(4, 5, 6)))(1)

    Assert.AreEqualArr Array(1, 0, 3), Nz2(Array(1, Null, 3))
    Assert.AreEqualArr Array(1, "", 3, "a"), Nz2(Array(1, Null, 3, "a"))

End Sub

Public Sub KoffeeTable_Test()

    Dim header:         header = Array("DEN_NO", "FruitName", "Color")
    Dim Denpyo:         Denpyo = Array(719750, 719750, 719750, 719750, 719750, 719750, 719751, 719751, 719751, 719751, 741900, 741900, 741900)
    Dim Fruit:          Fruit = Array("Apple", "DragonFruit", "Papaya", "Blueberry", "Mangosteen", "Grape", "Watermelon", "Cherry", "Peach", "GrapeFruit", "Kiwi", "Avocado", "Litchee")
    Dim FruitColor:     FruitColor = Array("red", "red", "yellow", "blue", "red", "green", "green", "red", "pink", "orange", "green", "dark yellow", "black")
    Dim tbl:            tbl = CreateTable(header, Array(Denpyo, Fruit, FruitColor))

    Assert.AreEqualArr header, tbl(0)
    Assert.AreEqualArr Denpyo, tbl(1)(0)
    Assert.AreEqualArr Fruit, tbl(1)(1)
    Assert.AreEqualArr FruitColor, tbl(1)(2)

    Assert.AreEqualArr header, keys(tbl)

    Assert.AreEqualArr Denpyo, Values(tbl)(0)
    Assert.AreEqualArr Fruit, Values(tbl)(1)
    Assert.AreEqualArr FruitColor, Values(tbl)(2)

    Assert.AreEqualArr FruitColor, Pluck(tbl, "Color")

    Assert.AreEqualArr Array(719750, "Blueberry", "blue"), Record(tbl, 3)
    Assert.AreEqualArr Array(719750, "Blueberry", "blue"), ToRecord(tbl)(3)
    Assert.AreEqualArr header, ToTable(header, ToRecord(tbl))(0)
    Assert.AreEqualArr FruitColor, Project(Array("Color"), tbl)(1)(0)

    Dim f As Func: Set f = Init(New Func, vbBoolean, AddressOf isGreenBlack, vbObject)
    Assert.AreEqualArr Array("Grape", "Watermelon", "Kiwi", "Litchee"), Flow(tbl, WHERE(f))(1)(1)
    Assert.AreEqualArr Array("Grape", "Watermelon", "Kiwi", "Litchee"), Restrict(f, tbl)(1)(1)


    Dim h1:     h1 = Array("A", "B", "C")
    Dim b1:     b1 = Array(1, 2, 3)
    Dim b2:     b2 = Array(11, 22, 33)
    Dim b3:     b3 = Array(111, 222, 333)
    Dim tbl1:   tbl1 = CreateTable(h1, Array(b1, b2, b3))

    Dim h2:     h2 = Array("a", "b", "c")
    Dim bb1:    bb1 = Array(1, 2, 5)
    Dim bb2:    bb2 = Array(11, 123, 554)
    Dim bb3:    bb3 = Array(111, 222, 555)
    Dim tbl2:   tbl2 = CreateTable(h2, Array(bb1, bb2, bb3))

    Assert.AreEqualArr Array(11), Intersect(tbl1, Array(1, 2), tbl2, Array(1, 2))(0)(1)(1)
    Assert.AreEqualArr Array(2, 3), RejectArr(tbl1, tbl2)(1)(0)

End Sub

Sub KoffeePuzzle_Test()

    Assert.AreEqualArr Array(), PowerSet(Array("A", "B", "C"))(0)
    Assert.AreEqualArr Array("C"), PowerSet(Array("A", "B", "C"))(1)
    Assert.AreEqualArr Array("B"), PowerSet(Array("A", "B", "C"))(2)
    Assert.AreEqualArr Array("B", "C"), PowerSet(Array("A", "B", "C"))(3)
    Assert.AreEqualArr Array("A"), PowerSet(Array("A", "B", "C"))(4)
    Assert.AreEqualArr Array("A", "C"), PowerSet(Array("A", "B", "C"))(5)
    Assert.AreEqualArr Array("A", "B"), PowerSet(Array("A", "B", "C"))(6)
    Assert.AreEqualArr Array("A", "B", "C"), PowerSet(Array("A", "B", "C"))(7)

    Assert.AreEqualArr Array("B", "C"), Combin(Array("A", "B", "C"), 2)(0)
    Assert.AreEqualArr Array("A", "C"), Combin(Array("A", "B", "C"), 2)(1)
    Assert.AreEqualArr Array("A", "B"), Combin(Array("A", "B", "C"), 2)(2)

    Assert.AreEqualArr Array("C", "B", "A"), ClctToArr(Permut(Array("A", "B", "C"), 3))(0)
    Assert.AreEqualArr Array("C", "A", "B"), ClctToArr(Permut(Array("A", "B", "C"), 3))(1)
    Assert.AreEqualArr Array("B", "C", "A"), ClctToArr(Permut(Array("A", "B", "C"), 3))(2)
    Assert.AreEqualArr Array("B", "A", "C"), ClctToArr(Permut(Array("A", "B", "C"), 3))(3)
    Assert.AreEqualArr Array("A", "C", "B"), ClctToArr(Permut(Array("A", "B", "C"), 3))(4)
    Assert.AreEqualArr Array("A", "B", "C"), ClctToArr(Permut(Array("A", "B", "C"), 3))(5)

    Dim clct As New collection: PermutLimited 1, 3, 3, clct
    Assert.AreEqualArr Array(CByte(3), CByte(2), CByte(1)), ClctToArr(clct)(0)
    Assert.AreEqualArr Array(CByte(3), CByte(1), CByte(2)), ClctToArr(clct)(1)
    Assert.AreEqualArr Array(CByte(2), CByte(3), CByte(1)), ClctToArr(clct)(2)
    Assert.AreEqualArr Array(CByte(2), CByte(1), CByte(3)), ClctToArr(clct)(3)
    Assert.AreEqualArr Array(CByte(1), CByte(3), CByte(2)), ClctToArr(clct)(4)
    Assert.AreEqualArr Array(CByte(1), CByte(2), CByte(3)), ClctToArr(clct)(5)

    Assert.AreEqualArr Array("A", "A"), ReptPermut(Array("A", "B", "C"), 2)(0)
    Assert.AreEqualArr Array("A", "B"), ReptPermut(Array("A", "B", "C"), 2)(1)
    Assert.AreEqualArr Array("A", "C"), ReptPermut(Array("A", "B", "C"), 2)(2)
    Assert.AreEqualArr Array("B", "A"), ReptPermut(Array("A", "B", "C"), 2)(3)
    Assert.AreEqualArr Array("B", "B"), ReptPermut(Array("A", "B", "C"), 2)(4)
    Assert.AreEqualArr Array("B", "C"), ReptPermut(Array("A", "B", "C"), 2)(5)
    Assert.AreEqualArr Array("C", "A"), ReptPermut(Array("A", "B", "C"), 2)(6)
    Assert.AreEqualArr Array("C", "B"), ReptPermut(Array("A", "B", "C"), 2)(7)
    Assert.AreEqualArr Array("C", "C"), ReptPermut(Array("A", "B", "C"), 2)(8)

End Sub


Public Sub KoffeeExcel_Test()

    '    |  A   B   C   D   E   F   G
    '----+---------------------------
    '  1 |
    '  2 |      1
    '  3 |      2
    '  4 |      3
    '  5 |
    '  6 |      A
    '  7 |      B

    Dim sh As Worksheet: Set sh = AddSheet("myTestTemporarySheet")
    PutVal Array(1#, 2#, 3#, Empty, "A", "B"), sh.Range("B2"), True


    Assert.AreEqual 4&, LastRow(sh.Range("B2"), True)
    Assert.AreEqual 7&, LastRow(sh.Range("B2"))


    Application.DisplayAlerts = False
    sh.Delete
    Application.DisplayAlerts = True



        '    |  A   B   C   D   E   F   G
    '----+---------------------------
    '  1 |
    '  2 |      1   2   3       A   B
    '  3 |
    '  4 |
    '  5 |
    '  6 |
    '  7 |

    Set sh = AddSheet("myTestTemporarySheet")
    PutVal Array(1#, 2#, 3#, Empty, "A", "B"), sh.Range("B2")


    Assert.AreEqual 4&, LastCol(sh.Range("B2"), True)
    Assert.AreEqual 7&, LastCol(sh.Range("B2"))


    Application.DisplayAlerts = False
    sh.Delete
    Application.DisplayAlerts = True


    '    |  A   B   C   D   E   F
    '----+------------------------
    '  1 |
    '  2 |      1   2   3       a
    '  3 |      4   5   6       b
    '  4 |      7   8   9       c
    '  5 |
    '  6 |      d   e   f

    Set sh = AddSheet("myTestTemporarySheet")
    PutVal JagArrToArr2D(Array(Array(1, 2, 3, Empty, "a"), _
                               Array(4, 5, 6, Empty, "b"), _
                               Array(7, 8, 9, Empty, "c"), _
                               Array(Empty, Empty, Empty, Empty, Empty), _
                               Array("d", "e", "f", Empty, Empty))), sh.Range("B2")

    Assert.AreEqualArr Array(1#, 2#, 3#, Empty, "a"), GetVal(sh.Range("B2"), AllHol)(0)
    Assert.AreEqualArr Array(4#, 5#, 6#, Empty, "b"), GetVal(sh.Range("B2"), AllHol)(1)
    Assert.AreEqualArr Array(7#, 8#, 9#, Empty, "c"), GetVal(sh.Range("B2"), AllHol)(2)
    Assert.AreEqualArr Array(Empty, Empty, Empty, Empty, Empty), GetVal(sh.Range("B2"), AllHol)(3)
    Assert.AreEqualArr Array("d", "e", "f", Empty, Empty), GetVal(sh.Range("B2"), AllHol)(4)

    Assert.AreEqualArr Array(1#, 4#, 7#, Empty, "d"), GetVal(sh.Range("B2"), AllVer)(0)
    Assert.AreEqualArr Array(2#, 5#, 8#, Empty, "e"), GetVal(sh.Range("B2"), AllVer)(1)
    Assert.AreEqualArr Array(3#, 6#, 9#, Empty, "f"), GetVal(sh.Range("B2"), AllVer)(2)
    Assert.AreEqualArr Array(Empty, Empty, Empty, Empty, Empty), GetVal(sh.Range("B2"), AllVer)(3)
    Assert.AreEqualArr Array("a", "b", "c", Empty, Empty), GetVal(sh.Range("B2"), AllVer)(4)

    Assert.AreEqualArr Array(1#, 2#, 3#, Empty, "a"), GetVal(sh.Range("B2"), Holizontal)
    Assert.AreEqualArr Array(1#, 4#, 7#, Empty, "d"), GetVal(sh.Range("B2"), Vertical)


    Application.DisplayAlerts = False
    sh.Delete
    Application.DisplayAlerts = True




    '    |  A   B   C   D   E   F
    '----+------------------------
    '  1 |
    '  2 |      1   2   3       a
    '  3 |      4   5   6       b
    '  4 |      7   8   9       c
    '  5 |
    '  6 |      d   e   f

    Set sh = AddSheet("myTestTemporarySheet")
    PutVal JagArrToArr2D(Array(Array(1, 2, 3, Empty, "a"), _
                               Array(4, 5, 6, Empty, "b"), _
                               Array(7, 8, 9, Empty, "c"), _
                               Array(Empty, Empty, Empty, Empty, Empty), _
                               Array("d", "e", "f", Empty, Empty))), sh.Range("B2")

    Assert.AreEqual 1#, sh.Range("B2").Value
    Assert.AreEqual 2#, sh.Range("C2").Value
    Assert.AreEqual 3#, sh.Range("D2").Value
    Assert.AreEqual Empty, sh.Range("E2").Value
    Assert.AreEqual "a", sh.Range("F2").Value

    Assert.AreEqual 4#, sh.Range("B3").Value
    Assert.AreEqual 5#, sh.Range("C3").Value
    Assert.AreEqual 6#, sh.Range("D3").Value
    Assert.AreEqual Empty, sh.Range("E3").Value
    Assert.AreEqual "b", sh.Range("F3").Value

    Assert.AreEqual 7#, sh.Range("B4").Value
    Assert.AreEqual 8#, sh.Range("C4").Value
    Assert.AreEqual 9#, sh.Range("D4").Value
    Assert.AreEqual Empty, sh.Range("E4").Value
    Assert.AreEqual "c", sh.Range("F4").Value

    Assert.AreEqual Empty, sh.Range("B5").Value
    Assert.AreEqual Empty, sh.Range("C5").Value
    Assert.AreEqual Empty, sh.Range("D5").Value
    Assert.AreEqual Empty, sh.Range("E5").Value
    Assert.AreEqual Empty, sh.Range("F5").Value

    Assert.AreEqual "d", sh.Range("B6").Value
    Assert.AreEqual "e", sh.Range("C6").Value
    Assert.AreEqual "f", sh.Range("D6").Value
    Assert.AreEqual Empty, sh.Range("E6").Value
    Assert.AreEqual Empty, sh.Range("F6").Value


    Application.DisplayAlerts = False
    sh.Delete
    Application.DisplayAlerts = True



    Dim v1 As Worksheet, arrx1 As New ArrayEx
    For Each v1 In ThisWorkbook.Worksheets
        arrx1.addval v1.Name
    Next v1

    Assert.AreEqualArr arrx1.ToArray, ArrSheetsName(ThisWorkbook)



    Dim v2 As Worksheet, arrx2 As New ArrayEx
    For Each v2 In ThisWorkbook.Worksheets
        arrx2.addval v2.Name
    Next v2

    Assert.AreEqual True, ExistsSheet(arrx2.ToArray(0))
    Assert.AreEqual False, ExistsSheet("abcdefghijklmnopqrstuvwxyz")


    Set sh = AddSheet("myTestTemporarySheet")

    Dim v3 As Worksheet, arrx3 As New ArrayEx
    For Each v3 In ThisWorkbook.Worksheets
        arrx3.addval v3.Name
    Next v3

    Assert.AreEqual True, ExistsSheet("myTestTemporarySheet")

    Application.DisplayAlerts = False
    sh.Delete
    Application.DisplayAlerts = True



    Set sh = AddSheet("myTestTemporarySheet")
    Dim shCopied As Worksheet: Set shCopied = CopySheet("myTestTemporarySheet", "myTestTemporarySheet2")

    Assert.AreEqual True, ExistsSheet("myTestTemporarySheet2")

    Application.DisplayAlerts = False
    sh.Delete
    shCopied.Delete
    Application.DisplayAlerts = True

End Sub

Public Sub ArrToNumericType_Test()

    Dim arr0 As Variant: arr0 = Array(1, 2, 3)
    Assert.AreEqualArr Array(CByte(1), CByte(2), CByte(3)), ArrToByte(arr0)
    Assert.AreEqualArr Array(1, 2, 3), ArrToInt(arr0)
    Assert.AreEqualArr Array(1&, 2&, 3&), ArrCLng(arr0)
    Assert.AreEqualArr Array(1!, 2!, 3!), ArrToSng(arr0)
    Assert.AreEqualArr Array(1#, 2#, 3#), ArrToDbl(arr0)
    Assert.AreEqualArr Array(1@, 2@, 3@), ArrToCur(arr0)

End Sub

Public Sub Flow_Test()

    Dim arr As Variant: arr = ArrRange(1, 100)

    Assert.AreEqualArr Map(AddressOf add1).Bind(arr).Apply, Flow(arr, Map(AddressOf add1))
    Assert.AreEqualArr Filter(IsOdd).Bind(arr).Apply, Flow(arr, Filter(IsOdd))
    Assert.AreEqualArr Zip(AddressOf Tuple2Of).Bind(Array(1, 2, 3)).Bind(Array("a", "b", "c")).Apply, Flow(Array(Array(1, 2, 3), Array("a", "b", "c")), Zip(AddressOf Tuple2Of))
    Assert.AreEqual Fold(Add).Bind(arr).Apply, Flow(arr, Fold(Add))
    Assert.AreEqual Fold(Add).Bind(arr).Bind(0).Apply, Flow(Array(0, arr), Fold(Add))
    Assert.AreEqualArr Scan(Add).Bind(arr).Apply, Flow(arr, Scan(Add))
    Assert.AreEqualArr Scan(Add).Bind(arr).Bind(0).Apply, Flow(Array(0, arr), Scan(Add))
    Assert.AreEqualArr Unfold(AddressOf FibFun).Bind(Init(New Tuple, 1, 1)).Apply, Flow(Init(New Tuple, 1, 1), Unfold(AddressOf FibFun))

    Dim arr1 As Variant: arr1 = ArrRange(1, 100)
    arr1 = Scan(Add).Bind(arr1).Apply
    arr1 = Map(AddressOf add1).Bind(arr1).Apply
    arr1 = Filter(IsOdd).Bind(arr1).Apply
    arr1 = Fold(Add).Bind(arr1).Apply

    Dim arr2 As Variant: arr2 = ArrRange(1, 100)
    arr2 = Flow(arr2, Scan(Add), Map(AddressOf add1), Filter(IsOdd), Fold(Add))

    Assert.AreEqual arr1, arr2
End Sub

Public Sub Map_Test()
    Dim f As Func: Set f = Init(New Func, vbVariant, AddressOf ToStr2, vbVariant)
    Assert.AreEqualArr Array("1", "2", "3"), Map(f).Bind(Array(1, 2, 3)).Apply

    Dim a As New Atom: Set a = a.AddFunc(Init(New Func, vbVariant, AddressOf ToStr2, vbVariant))
    Assert.AreEqualArr Array("1", "2", "3"), Map(a).Bind(Array(1, 2, 3)).Apply

    Assert.AreEqualArr Array("1", "2", "3"), Map(AddressOf ToStr2).Bind(Array(1, 2, 3)).Apply


    Assert.AreEqualArr Array(100#, 50#, 25#), Flow(Array(1, 2, 4), Map(Partial(div, 100)))
    Assert.AreEqualArr Array(1, 99, 3), Flow(Array(1, Null, 3), Map(Partial(NLV, 99)))

    Dim Jarr: Jarr = Array(Array(1, Null, 3), Array(10, Null, 30))
    Assert.AreEqualArr Array(0, 1, 2, 3), Flow(Jarr, Map(Map(Partial(NLV, 2))), Map(Partial(ShiftA, 0)))(0)
End Sub

Public Sub Filter_Test()
    Dim f As Func: Set f = Init(New Func, vbBoolean, AddressOf IsOddImpl, vbVariant)
    Assert.AreEqualArr Array(1, 3), Filter(f).Bind(Array(1, 2, 3)).Apply

    Assert.AreEqualArr Array(1, 3), Filter(IsOdd).Bind(Array(1, 2, 3)).Apply

    Assert.AreEqualArr Array(1, 3), Filter(AddressOf IsOddImpl).Bind(Array(1, 2, 3)).Apply
End Sub

Public Sub Zip_Test()
    Dim f As Func: Set f = Init(New Func, vbObject, AddressOf Tuple2Of, vbVariant, vbVariant)
    Assert.AreEqualArr _
        Array(Init(New Tuple, "A", 1), Init(New Tuple, "B", 2)), _
        Zip(f).Bind(Array("A", "B")).Bind(Array(1, 2)).Apply

    Dim arr As Variant: arr = Zip(Array2Of).Bind(Array("A", "B")).Bind(Array(1, 2)).Apply
    Assert.AreEqualArr Array("A", 1), arr(0)

    Assert.AreEqualArr _
        Array(Init(New Tuple, "A", 1), Init(New Tuple, "B", 2)), _
        Zip(AddressOf Tuple2Of).Bind(Array("A", "B")).Bind(Array(1, 2)).Apply

End Sub

Public Sub Fold_Test()
    Dim f As Func: Set f = Init(New Func, vbVariant, AddressOf AddImpl, vbVariant, vbVariant)
    Assert.AreEqual 6, Fold(f).Bind(Array(1, 2, 3)).Apply
    Assert.AreEqual 6, Fold(f).Bind(Array(1, 2, 3)).Bind(0).Apply


    Assert.AreEqual 6, Fold(Add).Bind(Array(1, 2, 3)).Apply
    Assert.AreEqual 6, Fold(Add).Bind(Array(1, 2, 3)).Bind(0).Apply

    Assert.AreEqual 6, Fold(AddressOf AddImpl).Bind(Array(1, 2, 3)).Apply
    Assert.AreEqual 6, Fold(AddressOf AddImpl).Bind(Array(1, 2, 3)).Bind(0).Apply
End Sub

Public Sub Scan_Test()
    Dim f As Func: Set f = Init(New Func, vbVariant, AddressOf AddImpl, vbVariant, vbVariant)
    Assert.AreEqualArr Array(1, 3, 6), Scan(f).Bind(Array(1, 2, 3)).Apply
    Assert.AreEqualArr Array(0, 1, 3, 6), Scan(f).Bind(Array(1, 2, 3)).Bind(0).Apply

    Assert.AreEqualArr Array(1, 3, 6), Scan(Add).Bind(Array(1, 2, 3)).Apply
    Assert.AreEqualArr Array(0, 1, 3, 6), Scan(Add).Bind(Array(1, 2, 3)).Bind(0).Apply

    Assert.AreEqualArr Array(1, 3, 6), Scan(AddressOf AddImpl).Bind(Array(1, 2, 3)).Apply
    Assert.AreEqualArr Array(0, 1, 3, 6), Scan(AddressOf AddImpl).Bind(Array(1, 2, 3)).Bind(0).Apply
End Sub

Public Sub Unfold_Test()
    Dim f As Func: Set f = Init(New Func, vbVariant, AddressOf FibFun, vbObject)
    Assert.AreEqualArr Array(1&, 2&, 3&, 5&, 8&, 13&, 21&, 34&, 55&, 89&), _
        Unfold(f).Bind(Init(New Tuple, 1, 1)).Apply

    Assert.AreEqualArr Array(1&, 2&, 3&, 5&, 8&, 13&, 21&, 34&, 55&, 89&), _
        Unfold(AddressOf FibFun).Bind(Init(New Tuple, 1, 1)).Apply
End Sub

Public Sub GroupByA_Test()
    Dim arr As Variant
    arr = Array( _
            Init(New Tuple, "A", 1), Init(New Tuple, "B", 4), Init(New Tuple, "C", 3), _
            Init(New Tuple, "A", 2), Init(New Tuple, "B", 3), Init(New Tuple, "C", 1), _
            Init(New Tuple, "A", 3), Init(New Tuple, "B", 2), Init(New Tuple, "C", 2), _
            Init(New Tuple, "A", 4), Init(New Tuple, "B", 1))

    Dim lkup As Variant

    Dim f As Func: Set f = Init(New Func, vbVariant, AddressOf GetItem1, vbObject)
    lkup = GroupBy(f).Bind(arr).Apply

    Assert.AreEqual 3&, ArrLen(lkup)
    Assert.AreEqual "A", lkup(0).Item1
    Assert.AreEqual "B", lkup(1).Item1
    Assert.AreEqual "C", lkup(2).Item1
    Assert.AreEqualArr _
        Array(Init(New Tuple, "A", 1), Init(New Tuple, "A", 2), Init(New Tuple, "A", 3), Init(New Tuple, "A", 4)), _
        lkup(0).Item2
    Assert.AreEqualArr _
        Array(Init(New Tuple, "B", 4), Init(New Tuple, "B", 3), Init(New Tuple, "B", 2), Init(New Tuple, "B", 1)), _
        lkup(1).Item2
    Assert.AreEqualArr _
        Array(Init(New Tuple, "C", 3), Init(New Tuple, "C", 1), Init(New Tuple, "C", 2)), _
        lkup(2).Item2
End Sub

Public Sub GroupByB_Test()
    Dim arr As Variant
    arr = Array( _
            Init(New Tuple, "A", 1), Init(New Tuple, "B", 4), Init(New Tuple, "C", 3), _
            Init(New Tuple, "A", 2), Init(New Tuple, "B", 3), Init(New Tuple, "C", 1), _
            Init(New Tuple, "A", 3), Init(New Tuple, "B", 2), Init(New Tuple, "C", 2), _
            Init(New Tuple, "A", 4), Init(New Tuple, "B", 1))

    Dim lkup As Variant

    lkup = GroupBy(AddressOf GetItem1).Bind(arr).Apply

    Assert.AreEqual 3&, ArrLen(lkup)
    Assert.AreEqual "A", lkup(0).Item1
    Assert.AreEqual "B", lkup(1).Item1
    Assert.AreEqual "C", lkup(2).Item1
    Assert.AreEqualArr _
        Array(Init(New Tuple, "A", 1), Init(New Tuple, "A", 2), Init(New Tuple, "A", 3), Init(New Tuple, "A", 4)), _
        lkup(0).Item2
    Assert.AreEqualArr _
        Array(Init(New Tuple, "B", 4), Init(New Tuple, "B", 3), Init(New Tuple, "B", 2), Init(New Tuple, "B", 1)), _
        lkup(1).Item2
    Assert.AreEqualArr _
        Array(Init(New Tuple, "C", 3), Init(New Tuple, "C", 1), Init(New Tuple, "C", 2)), _
        lkup(2).Item2
End Sub

Public Sub FoldR_Test()
    Dim f As Func: Set f = Init(New Func, vbVariant, AddressOf DivImpl, vbVariant, vbVariant)
    Assert.AreEqual 3#, FoldR(f).Bind(Array(5, 2, 30)).Apply
    Assert.AreEqual 0.2, FoldR(f).Bind(Array(5, 2, 30)).Bind(60).Apply

    Assert.AreEqual 3#, FoldR(div).Bind(Array(5, 2, 30)).Apply
    Assert.AreEqual 0.2, FoldR(div).Bind(Array(5, 2, 30)).Bind(60).Apply

    Assert.AreEqual 3#, FoldR(AddressOf DivImpl).Bind(Array(5, 2, 30)).Apply
    Assert.AreEqual 0.2, FoldR(AddressOf DivImpl).Bind(Array(5, 2, 30)).Bind(60).Apply
End Sub

Public Sub ScanR_Test()
    Dim f As Func: Set f = Init(New Func, vbVariant, AddressOf AddImpl, vbVariant, vbVariant)
    Assert.AreEqualArr Array(3, 5, 6), ScanR(f).Bind(Array(1, 2, 3)).Apply
    Assert.AreEqualArr Array(0, 3, 5, 6), ScanR(f).Bind(Array(1, 2, 3)).Bind(0).Apply

    Assert.AreEqualArr Array(3, 5, 6), ScanR(Add).Bind(Array(1, 2, 3)).Apply
    Assert.AreEqualArr Array(0, 3, 5, 6), ScanR(Add).Bind(Array(1, 2, 3)).Bind(0).Apply

    Assert.AreEqualArr Array(3, 5, 6), ScanR(AddressOf AddImpl).Bind(Array(1, 2, 3)).Apply
    Assert.AreEqualArr Array(0, 3, 5, 6), ScanR(AddressOf AddImpl).Bind(Array(1, 2, 3)).Bind(0).Apply
End Sub

Public Sub fnull_Test()
    Assert.AreEqual "abc", fnull(Init(New Func, vbVariant, AddressOf AlwaysImpl, vbVariant), 1).Bind("abc").Apply
    Assert.AreEqual "", fnull(Init(New Func, vbVariant, AddressOf AlwaysImpl, vbVariant), 1).Bind("").Apply
    Assert.AreEqual 1, fnull(Init(New Func, vbVariant, AddressOf AlwaysImpl, vbVariant), 1).Bind(1).Apply
    Assert.AreEqual 0, fnull(Init(New Func, vbVariant, AddressOf AlwaysImpl, vbVariant), 1).Bind(0).Apply
    Assert.AreEqual -1, fnull(Init(New Func, vbVariant, AddressOf AlwaysImpl, vbVariant), 1).Bind(-1).Apply
    Assert.AreEqual 1, fnull(Init(New Func, vbVariant, AddressOf AlwaysImpl, vbVariant), 1).Bind(Null).Apply
    Assert.AreEqual 1, fnull(Init(New Func, vbVariant, AddressOf AlwaysImpl, vbVariant), 1).Bind(Empty).Apply
'

    Assert.AreEqualArr Array(20, 2000, 60), Flow(Array(1, Null, 3), Map(fnull(Partial(Mult, 20), 100)))
    Assert.AreEqual 300, Flow(Array(1, Null, 3), Fold(fnull(Mult, 100)))
    Assert.AreEqualArr Array(1, 100, 300), Flow(Array(1, Null, 3), Scan(fnull(Mult, 100)))
    Assert.AreEqualArr Array(1, 3), Flow(Array(1, Null, 3), Filter(fnull(IsOdd, 0)))
    Assert.AreEqualArr Array(100, 4), Flow(Array(Array(1, Null), Array(3, 4)), Zip(fnull(Array2Of, 100)))(1)

End Sub

Public Sub Existy_Test()
    Assert.AreEqual True, Existy("abc")
    Assert.AreEqual True, Existy("")
    Assert.AreEqual True, Existy(1)
    Assert.AreEqual True, Existy(0)
    Assert.AreEqual True, Existy(-1)
    Assert.AreEqual False, Existy(Null)
    Assert.AreEqual False, Existy(Empty)
    Dim obj As Object
    Assert.AreEqual False, Existy(obj)
End Sub

Public Sub Truthy_Test()
    Assert.AreEqual True, Truthy("abc")
    Assert.AreEqual True, Truthy("")
    Assert.AreEqual True, Truthy(1)
    Assert.AreEqual False, Truthy(0)
    Assert.AreEqual True, Truthy(-1)
    Assert.AreEqual False, Truthy(Null)
    Assert.AreEqual False, Truthy(Empty)
    Dim obj As Object
    Assert.AreEqual False, Truthy(obj)
End Sub

Public Sub Always_Test()
    Dim x As Variant: Set x = Always(New Func)
    Dim y As Variant: Set y = Always(New Func)
    Assert.AreEqual True, Eq(x, x)
    Assert.AreEqual False, Eq(x, y)

    Assert.AreEqualArr Array(1, 2, 3), Always(Array(1, 2, 3)).Apply

    Assert.AreEqual "abc", Always("abc").Apply
    Assert.AreEqual "", Always("").Apply
    Assert.AreEqual 1, Always(1).Apply
    Assert.AreEqual 0, Always(0).Apply
    Assert.AreEqual -1, Always(-1).Apply
    Assert.IsNullVal Always(Null).Apply
    Assert.AreEqual Empty, Always(Empty).Apply
End Sub

Public Sub Repeat_Test()
    Assert.AreEqualArr Array("abc", "abc", "abc"), Repeat(3, Always("abc"))
    Assert.AreEqualArr Array("", "", ""), Repeat(3, Always(""))
    Assert.AreEqualArr Array(1, 1, 1), Repeat(3, Always(1))
    Assert.AreEqualArr Array(0, 0, 0), Repeat(3, Always(0))
    Assert.AreEqualArr Array(-1, -1, -1), Repeat(3, Always(-1))
    Assert.AreEqualArr Array(Empty, Empty, Empty), Repeat(3, Always(Empty))

    Dim f As New Func: Set f = Init(New Func, vbVariant, AddressOf SayVBA)
    Assert.AreEqualArr Array("VBA!", "VBA!", "VBA!"), Repeat(3, f)
End Sub

Public Sub NLV_Test()

    Assert.AreEqual "abc", Partial(NLV, 99).Bind("abc").Apply
    Assert.AreEqual "", Partial(NLV, 99).Bind("").Apply
    Assert.AreEqual 1, Partial(NLV, 99).Bind(1).Apply
    Assert.AreEqual 0, Partial(NLV, 99).Bind(0).Apply
    Assert.AreEqual -1, Partial(NLV, 99).Bind(-1).Apply
    Assert.AreEqual 99, Partial(NLV, 99).Bind(Null).Apply
    Assert.AreEqual 99, Partial(NLV, 99).Bind(Empty).Apply
    Dim obj As Object
    Assert.AreEqual 99, Partial(NLV, 99).Bind(obj).Apply
End Sub

Public Sub Array2Of_Test()
    Assert.AreEqualArr Array(1, "A"), Array2Of.Apply(1, "A")
End Sub

Public Sub ToStr2_Test()
    Assert.AreEqual "1", ToStr2(1)
End Sub

Public Sub Add_Test()
    Assert.AreEqual 7, Add.Apply(3, 4)
End Sub

Public Sub Minus_Test()
    Assert.AreEqual -1, Minus.Apply(3, 4)
End Sub

Public Sub Mult_Test()
    Assert.AreEqual 12, Mult.Apply(3, 4)
End Sub

Public Sub Div_Test()
    Assert.AreEqual 5#, div.Apply(50, 10)
End Sub

Public Sub DivR_Test()
    Assert.AreEqual 0.2, DivR.Apply(50, 10)
End Sub

Public Sub Partial_Test()

    Dim AJoin As Func: Set AJoin = Init(New Func, vbVariant, AddressOf ArrJoin, vbVariant)

    Assert.AreEqual "1 2 3", Partial(AJoin).Apply(Array(1, 2, 3))
    Assert.AreEqual "1 2 3", Partial(AJoin, Array(1, 2, 3)).Apply
    Assert.AreEqual "1 2 3", Partial(AJoin).Bind(Array(1, 2, 3)).Apply

    Dim Plus1 As Func: Set Plus1 = Init(New Func, vbVariant, AddressOf add1, vbVariant)

    Assert.AreEqual 11, Partial(Plus1).Apply(10)
    Assert.AreEqual 11, Partial(Plus1).Bind(10).Apply
    Assert.AreEqual 11, Partial(Plus1, 10).Apply

    Assert.AreEqual 2#, Partial(div).Apply(20, 10)
    Assert.AreEqual 2#, Partial(div).Bind(20).Apply(10)
    Assert.AreEqual 2#, Partial(div).Bind(20).Bind(10).Apply
    Assert.AreEqual 2#, Partial(div, 20).Apply(10)
    Assert.AreEqual 2#, Partial(div, 20).Bind(10).Apply
    Assert.AreEqual 2#, Partial(div, 20, 10).Apply
'
    Dim div3 As Func: Set div3 = Init(New Func, vbVariant, AddressOf DivTriImpl, vbVariant, vbVariant, vbVariant)

    Assert.AreEqual 1#, Partial(div3).Apply(10, 2, 5)
    Assert.AreEqual 1#, Partial(div3).Bind(10).Apply(2, 5)
    Assert.AreEqual 1#, Partial(div3).Bind(10).Bind(2).Apply(5)
    Assert.AreEqual 1#, Partial(div3).Bind(10).Bind(2).Bind(5).Apply
    Assert.AreEqual 1#, Partial(div3, 10).Apply(2, 5)
    Assert.AreEqual 1#, Partial(div3, 10).Bind(2).Apply(5)
    Assert.AreEqual 1#, Partial(div3, 10).Bind(2).Bind(5).Apply
    Assert.AreEqual 1#, Partial(div3, 10, 2).Apply(5)
    Assert.AreEqual 1#, Partial(div3, 10, 2).Bind(5).Apply
    Assert.AreEqual 1#, Partial(div3, 10, 2, 5).Apply

    Dim f As Func: Set f = Init(New Func, vbVariant, AddressOf ToStr2, vbVariant)

    Assert.AreEqual 50.2, Flow(50, Partial(div, 10), Partial(Add, 50))
    Assert.AreEqual "1", Flow(1, Partial(f))
'
    Assert.AreEqualArr Array(9, 1, 2, 3), Partial(ShiftA, 9).Apply(Array(1, 2, 3))
    Assert.AreEqualArr Array(1, 2, 3, 9), Partial(PushA, 9).Apply(Array(1, 2, 3))
'
   Assert.AreEqual 101, Flow(1, Partial(zero, Partial(fst, Partial(snd, Partial(Add, 100)))))

End Sub

Public Sub HelloWorld_Test()

    Dim rept As Func: Set rept = Init(New Func, vbVariant, AddressOf Repeat, vbVariant, vbVariant)
    Assert.AreEqualArr Array("Hello, World!", "Hello, World!", "Hello, World!"), Repeat(3, "Hello, World!")
    Assert.AreEqualArr Array("Hello, World!", "Hello, World!", "Hello, World!"), Flow("Hello, World!", Partial(rept, 3))
    Assert.AreEqualArr Array("Hello, World!", "Hello, World!", "Hello, World!"), Flow(Always("Hello, World!").Apply, Partial(rept, 3))
    Assert.AreEqualArr Array("Hello, World!", "Hello, World!", "Hello, World!"), Flow(Array(Null, Null, Null), Map(Partial(NLV, "Hello, World!")))

End Sub

Public Sub Actions_Test()

    Dim StackAction As Atom
    Set StackAction = Actions(Lift(Partial(ShiftA, 99)) _
                            , Lift(Partial(ShiftA, 55)) _
                            , Lift(PopA))

    Dim tpl As Tuple: Set tpl = Flow(Array(1, 2), StackAction)

    Assert.AreEqualArr Array(55, 99, 1), tpl.Item1
    Assert.AreEqualArr Array(99, 1, 2), tpl.Item2(0)

End Sub

Public Sub Compose_Test()

    Dim arr: arr = Array(1, 1, 2, 7, 5, 99, 55)

    Dim uniq As Func:       Set uniq = Init(New Func, vbVariant, AddressOf ArrUniq, vbVariant)
    Dim sort As Func:       Set sort = Init(New Func, vbVariant, AddressOf ArrSortAsc, vbVariant)
    Dim ToString As Func:   Set ToString = Init(New Func, vbVariant, AddressOf ToStr2, vbVariant)

    Dim f As Atom: Set f = Compose(Map(Partial(Add, 100)), uniq, sort, Map(ToString))

    Assert.AreEqualArr Array("101", "102", "105", "107", "155", "199"), f.Apply(arr)
End Sub

Public Sub Reject_Test()
    Dim f As Func: Set f = Init(New Func, vbBoolean, AddressOf IsOddImpl, vbVariant)
    Assert.AreEqualArr Array(2), Reject(f).Bind(Array(1, 2, 3)).Apply
    Assert.AreEqualArr Array(2), Reject(IsOdd).Bind(Array(1, 2, 3)).Apply
    Assert.AreEqualArr Array(2), Reject(AddressOf IsOddImpl).Bind(Array(1, 2, 3)).Apply
End Sub

Public Sub AllAnyFind_Test()

    Assert.AreEqual False, All(AddressOf IsNumber).Bind(Array("a", "b", 3, "d")).Apply
    Assert.AreEqual True, Any_(AddressOf IsNumber).Bind(Array("a", "b", 3, "d")).Apply
    Assert.AreEqual 3, Find(AddressOf IsNumber).Bind(Array("a", "b", 3, "d")).Apply

    Dim T As Func: Set T = Init(New Func, vbBoolean, AddressOf Tr)
    Dim f As Func: Set f = Init(New Func, vbBoolean, AddressOf Fa)

    Assert.AreEqual True, AllOf()
    Assert.AreEqual True, AllOf(T, T)
    Assert.AreEqual False, AllOf(T, T, T, T, f)

    Assert.AreEqual False, AnyOf()
    Assert.AreEqual True, AnyOf(T, T, f)
    Assert.AreEqual False, AnyOf(f, f, f, f)

End Sub
