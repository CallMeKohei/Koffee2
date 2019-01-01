VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KoffeeTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' --------------------------------------------------------
'''  FILE    : KoffeeTest.cls
'''  AUTHOR  : callmekohei <callmekohei at gmail.com>
'''  License : MIT license
''' --------------------------------------------------------
Option Explicit

'Assert.RunTestOf New KoffeeTest

''' --------------------------------------------------------
'''                     KoffeeCore Test
''' --------------------------------------------------------

Public Sub IsJagArr_Test()

    ''' "Array of arrays" is Jagged arrays
    Assert.AreEqual True, IsJagArr(Array(Array(0)))

    ''' Not initalized array is as array.
    Assert.AreEqual True, IsJagArr(Array(Array()))
    Assert.AreEqual True, IsJagArr(Array(Array(), Array(0)))

    ''' All inner elements must be array.
    Assert.AreEqual False, IsJagArr(0)
    Assert.AreEqual False, IsJagArr(New Collection)
    Assert.AreEqual False, IsJagArr(Array())
    Assert.AreEqual False, IsJagArr(Array(Array(0), 0))
    Assert.AreEqual False, IsJagArr(Array(Array(0), New Collection))

    ''' One Dimension
    Dim arr1D(0)
    Assert.AreEqual False, IsJagArr(arr1D)

    Dim nArr1D(-1 To -1)
    Assert.AreEqual False, IsJagArr(nArr1D)

    ''' Two Dimensions
    Dim arr2D(0, 0)
    Assert.AreEqual False, IsJagArr(arr2D)

    Dim nArr2D(-1 To -1, -1 To -1)
    Assert.AreEqual False, IsJagArr(nArr2D)

    ''' More Dimensions
    Dim arr3D(0, 0, 0)
    Assert.AreEqual False, IsJagArr(arr3D)

    Dim nArr3D(-1 To -1, -1 To -1, -1 To -1)
    Assert.AreEqual False, IsJagArr(nArr3D)

End Sub

Public Sub ArrTranspose_Test()

    '''  Before             After
    '''
    '''  1  2  3     --->   1  2  5
    '''  4  5  6            2  4  6

    ''' validation
    ''' not array
    Assert.IsErrFunc 13, Init(New Func, vbVariant, AddressOf ArrTranspose, vbVariant), Array(0)
    ''' not two dimensions
    Dim arr1D(0) As Variant
    Assert.IsErrFunc 13, Init(New Func, vbVariant, AddressOf ArrTranspose, vbVariant), Array(arr1D)
    ''' not two dimensions
    Dim arr3D(0, 0, 0) As Variant
    Assert.IsErrFunc 13, Init(New Func, vbVariant, AddressOf ArrTranspose, vbVariant), Array(arr3D)

    ''' Before ( Base 0 )
    Dim arr As Variant: arr = JagArrToArr2D(Array(Array(1, 2, 3), Array(4, 5, 6)))

    Assert.AreEqual 1, arr(0, 0)
    Assert.AreEqual 2, arr(0, 1)
    Assert.AreEqual 3, arr(0, 2)

    Assert.AreEqual 4, arr(1, 0)
    Assert.AreEqual 5, arr(1, 1)
    Assert.AreEqual 6, arr(1, 2)

    ''' After ( Base 0 )
    Dim arr2 As Variant: arr2 = ArrTranspose(arr)

    Assert.AreEqual 1, arr2(0, 0)
    Assert.AreEqual 4, arr2(0, 1)

    Assert.AreEqual 2, arr2(1, 0)
    Assert.AreEqual 5, arr2(1, 1)

    Assert.AreEqual 3, arr2(2, 0)
    Assert.AreEqual 6, arr2(2, 1)

    ''' Before ( Base 1 )
    Dim arr3 As Variant: arr3 = [{1,2,3;4,5,6}]

    Assert.AreEqual 1#, arr3(1, 1)
    Assert.AreEqual 2#, arr3(1, 2)
    Assert.AreEqual 3#, arr3(1, 3)

    Assert.AreEqual 4#, arr3(2, 1)
    Assert.AreEqual 5#, arr3(2, 2)
    Assert.AreEqual 6#, arr3(2, 3)

    ''' After ( Base 1 )
    Dim arr4 As Variant: arr4 = ArrTranspose(arr3)

    Assert.AreEqual 1#, arr4(1, 1)
    Assert.AreEqual 4#, arr4(1, 2)

    Assert.AreEqual 2#, arr4(2, 1)
    Assert.AreEqual 5#, arr4(2, 2)

    Assert.AreEqual 3#, arr4(3, 1)
    Assert.AreEqual 6#, arr4(3, 2)

End Sub

''' --------------------------------------------------------
'''                     KoffeeExcel Test
''' --------------------------------------------------------

Public Sub ArrSheetsName_Test()

    Dim arr As Variant: arr = ArrSheetsName()

    Dim wb As Workbook: Set wb = Application.ThisWorkbook

    Assert.AreEqual wb.Worksheets.Count, ArrLen(arr)

    Dim ws As Worksheet, i As Long
    For Each ws In wb.Worksheets
        Assert.AreEqual ws.Name, arr(i)
        i = i + 1
    Next ws

End Sub

Public Sub ExistsSheet_Test()

    Assert.AreEqual False, ExistsSheet("abcdefghijklmnopqrstuvwxyz")

    Dim worksheet_name As Variant
    For Each worksheet_name In ArrSheetsName()
        Assert.AreEqual True, ExistsSheet(worksheet_name)
    Next worksheet_name

End Sub

Public Sub DeleteSheet_Test()

    Assert.AreEqual False, ExistsSheet("myTestTemporarySheet")
    Application.ThisWorkbook.Worksheets.Add().Name = "myTestTemporarySheet"
    Assert.AreEqual True, ExistsSheet("myTestTemporarySheet")
    DeleteSheet "myTestTemporarySheet"
    Assert.AreEqual False, ExistsSheet("myTestTemporarySheet")

End Sub

Public Sub AddSheet_Test()

    Assert.AreEqual False, ExistsSheet("myTestTemporarySheet")
    Dim ws As Worksheet: Set ws = AddSheet("myTestTemporarySheet")
    Assert.AreEqual True, ExistsSheet("myTestTemporarySheet")
    DeleteSheet "myTestTemporarySheet"
    Assert.AreEqual False, ExistsSheet("myTestTemporarySheet")

End Sub

Public Sub CopySheet_Test()

    Assert.AreEqual False, ExistsSheet("myTestTemporarySheet")
    Assert.AreEqual False, ExistsSheet("myTestTemporarySheet2")

    Dim ws       As Worksheet: Set ws = AddSheet("myTestTemporarySheet")
    Dim wsCopied As Worksheet: Set wsCopied = CopySheet("myTestTemporarySheet", "myTestTemporarySheet2")

    Assert.AreEqual True, ExistsSheet("myTestTemporarySheet")
    Assert.AreEqual True, ExistsSheet("myTestTemporarySheet2")

    DeleteSheet "myTestTemporarySheet"
    DeleteSheet "myTestTemporarySheet2"

    Assert.AreEqual False, ExistsSheet("myTestTemporarySheet")
    Assert.AreEqual False, ExistsSheet("myTestTemporarySheet2")

End Sub

Public Sub GetVal_Test()

    '    |  A   B   C   D   E
    '----+--------------------
    '  1 |
    '  2 |      1   2       a
    '  3 |      3   4       b
    '  4 |
    '  5 |      c   d
    '  6 |

    Dim ws As Worksheet: Set ws = AddSheet("myTestTemporarySheet")

    ws.Range("B2").Value = 1
    ws.Range("C2").Value = 2
    ws.Range("D2").Value = Empty
    ws.Range("E2").Value = "a"

    ws.Range("B3").Value = 3
    ws.Range("C3").Value = 4
    ws.Range("D3").Value = Empty
    ws.Range("E3").Value = "b"

    ws.Range("B4").Value = Empty
    ws.Range("C4").Value = Empty
    ws.Range("D4").Value = Empty
    ws.Range("E4").Value = Empty

    ws.Range("B5").Value = "c"
    ws.Range("C5").Value = "d"
    ws.Range("D5").Value = Empty
    ws.Range("E5").Value = Empty

    Assert.AreEqualArr Array(1#), GetVal(ws.Range("B2"))

    Assert.AreEqualArr Array(1#, 2#, Empty, "a"), GetVal(ws.Range("B2:E2"))(1)

    Assert.AreEqualArr Array(1#), GetVal(ws.Range("B2:B5"))(1)
    Assert.AreEqualArr Array(3#), GetVal(ws.Range("B2:B5"))(2)
    Assert.AreEqualArr Array(Empty), GetVal(ws.Range("B2:B5"))(3)
    Assert.AreEqualArr Array("c"), GetVal(ws.Range("B2:B5"))(4)

    Assert.AreEqualArr Array(1#, 2#, Empty, "a"), GetVal(ws.Range("B2:E5"))(1)
    Assert.AreEqualArr Array(3#, 4#, Empty, "b"), GetVal(ws.Range("B2:E5"))(2)
    Assert.AreEqualArr Array(Empty, Empty, Empty, Empty), GetVal(ws.Range("B2:E5"))(3)
    Assert.AreEqualArr Array("c", "d", Empty, Empty), GetVal(ws.Range("B2:E5"))(4)

    DeleteSheet "myTestTemporarySheet"

End Sub

Public Sub PutVal_Test()

    Dim ws As Worksheet: Set ws = AddSheet("myTestTemporarySheet")

    ''' --------------------------------------------------------
    PutVal 1, ws.Range("B2")

    '    |  A   B   C   D   E   F   G
    '----+---------------------------
    '  1 |
    '  2 |      1
    '  3 |
    '  4 |
    '  5 |
    '  6 |
    '  7 |

    Assert.AreEqual 1#, ws.Range("B2").Value

    ''' --------------------------------------------------------
    ws.Cells.Clear

    PutVal Array(1#, 2#, 3#, Empty, "A", "B"), ws.Range("B2")


    '    |  A   B   C   D   E   F   G
    '----+---------------------------
    '  1 |
    '  2 |      1   2   3       A   B
    '  3 |
    '  4 |
    '  5 |
    '  6 |
    '  7 |

    Assert.AreEqual 1#, ws.Range("B2").Value
    Assert.AreEqual 2#, ws.Range("C2").Value
    Assert.AreEqual 3#, ws.Range("D2").Value
    Assert.AreEqual Empty, ws.Range("E2").Value
    Assert.AreEqual "A", ws.Range("F2").Value
    Assert.AreEqual "B", ws.Range("G2").Value

    ''' ----------------------------------------------------
    ws.Cells.Clear

    PutVal 1, ws.Range("B2"), True

    '    |  A   B   C   D   E   F   G
    '----+---------------------------
    '  1 |
    '  2 |      1
    '  3 |
    '  4 |
    '  5 |
    '  6 |
    '  7 |

    Assert.AreEqual 1#, ws.Range("B2").Value

    ''' --------------------------------------------------------
    ws.Cells.Clear

    PutVal Array(1#, 2#, 3#, Empty, "A", "B"), ws.Range("B2"), True

    '    |  A   B   C   D   E   F   G
    '----+---------------------------
    '  1 |
    '  2 |      1
    '  3 |      2
    '  4 |      3
    '  5 |
    '  6 |      A
    '  7 |      B

    Assert.AreEqual 1#, ws.Range("B2").Value
    Assert.AreEqual 2#, ws.Range("B3").Value
    Assert.AreEqual 3#, ws.Range("B4").Value
    Assert.AreEqual Empty, ws.Range("B5").Value
    Assert.AreEqual "A", ws.Range("B6").Value
    Assert.AreEqual "B", ws.Range("B7").Value

    ''' --------------------------------------------------------
    ws.Cells.Clear

    PutVal Array( _
        Array(1, 2, 3, Empty, "a"), _
        Array(4, 5, 6, Empty, "b"), _
        Array(7, 8, 9, Empty, "c"), _
        Array(Empty, Empty, Empty, Empty, Empty), _
        Array("d", "e", "f", Empty, Empty) _
    ), ws.Range("B2")

    '    |  A   B   C   D   E   F
    '----+------------------------
    '  1 |
    '  2 |      1   2   3       a
    '  3 |      4   5   6       b
    '  4 |      7   8   9       c
    '  5 |
    '  6 |      d   e   f

    Assert.AreEqual 1#, ws.Range("B2").Value
    Assert.AreEqual 2#, ws.Range("C2").Value
    Assert.AreEqual 3#, ws.Range("D2").Value
    Assert.AreEqual Empty, ws.Range("E2").Value
    Assert.AreEqual "a", ws.Range("F2").Value

    Assert.AreEqual 4#, ws.Range("B3").Value
    Assert.AreEqual 5#, ws.Range("C3").Value
    Assert.AreEqual 6#, ws.Range("D3").Value
    Assert.AreEqual Empty, ws.Range("E3").Value
    Assert.AreEqual "b", ws.Range("F3").Value

    Assert.AreEqual 7#, ws.Range("B4").Value
    Assert.AreEqual 8#, ws.Range("C4").Value
    Assert.AreEqual 9#, ws.Range("D4").Value
    Assert.AreEqual Empty, ws.Range("E4").Value
    Assert.AreEqual "c", ws.Range("F4").Value

    Assert.AreEqual Empty, ws.Range("B5").Value
    Assert.AreEqual Empty, ws.Range("C5").Value
    Assert.AreEqual Empty, ws.Range("D5").Value
    Assert.AreEqual Empty, ws.Range("E5").Value
    Assert.AreEqual Empty, ws.Range("F5").Value

    Assert.AreEqual "d", ws.Range("B6").Value
    Assert.AreEqual "e", ws.Range("C6").Value
    Assert.AreEqual "f", ws.Range("D6").Value
    Assert.AreEqual Empty, ws.Range("E6").Value
    Assert.AreEqual Empty, ws.Range("F6").Value

    ''' --------------------------------------------------------
    ws.Cells.Clear

    PutVal Array( _
        Array(1, 2, 3, Empty, "a"), _
        Array(4, 5, 6, Empty, "b"), _
        Array(7, 8, 9, Empty, "c"), _
        Array(Empty, Empty, Empty, Empty, Empty), _
        Array("d", "e", "f", Empty, Empty) _
    ), ws.Range("B2"), True

    '    |  A   B   C   D   E   F
    '----+------------------------
    '  1 |
    '  2 |      1   4   7       d
    '  3 |      2   5   8       e
    '  4 |      3   6   9       f
    '  5 |
    '  6 |      a   b   c

    Assert.AreEqual 1#, ws.Range("B2").Value
    Assert.AreEqual 2#, ws.Range("B3").Value
    Assert.AreEqual 3#, ws.Range("B4").Value
    Assert.AreEqual Empty, ws.Range("B5").Value
    Assert.AreEqual "a", ws.Range("B6").Value

    Assert.AreEqual 4#, ws.Range("C2").Value
    Assert.AreEqual 5#, ws.Range("C3").Value
    Assert.AreEqual 6#, ws.Range("C4").Value
    Assert.AreEqual Empty, ws.Range("C5").Value
    Assert.AreEqual "b", ws.Range("C6").Value

    Assert.AreEqual 7#, ws.Range("D2").Value
    Assert.AreEqual 8#, ws.Range("D3").Value
    Assert.AreEqual 9#, ws.Range("D4").Value
    Assert.AreEqual Empty, ws.Range("D5").Value
    Assert.AreEqual "c", ws.Range("D6").Value

    Assert.AreEqual Empty, ws.Range("E2").Value
    Assert.AreEqual Empty, ws.Range("E3").Value
    Assert.AreEqual Empty, ws.Range("E4").Value
    Assert.AreEqual Empty, ws.Range("E5").Value
    Assert.AreEqual Empty, ws.Range("E6").Value

    Assert.AreEqual "d", ws.Range("F2").Value
    Assert.AreEqual "e", ws.Range("F3").Value
    Assert.AreEqual "f", ws.Range("F4").Value
    Assert.AreEqual Empty, ws.Range("F5").Value
    Assert.AreEqual Empty, ws.Range("F6").Value

    DeleteSheet "myTestTemporarySheet"

End Sub

Public Sub PutValBase1_Test()

    Dim ws As Worksheet: Set ws = AddSheet("myTestTemporarySheet")
    Dim arrBase1 As Variant

    ''' ----------------------------------------------------
    PutVal 1, ws.Range("B2")

    '    |  A   B   C   D   E   F   G
    '----+---------------------------
    '  1 |
    '  2 |      1
    '  3 |
    '  4 |
    '  5 |
    '  6 |
    '  7 |

    Assert.AreEqual 1#, ws.Range("B2").Value

    ''' ----------------------------------------------------
    ws.Cells.Clear

    PutVal Array(1#, 2#, 3#, Empty, "A", "B"), ws.Range("B2")

    arrBase1 = GetVal(ws.Range("B2:G2"))
    ws.Cells.Clear
    PutVal arrBase1, ws.Range("B2")

    '    |  A   B   C   D   E   F   G
    '----+---------------------------
    '  1 |
    '  2 |      1   2   3       A   B
    '  3 |
    '  4 |
    '  5 |
    '  6 |
    '  7 |

    Assert.AreEqual 1#, ws.Range("B2").Value
    Assert.AreEqual 2#, ws.Range("C2").Value
    Assert.AreEqual 3#, ws.Range("D2").Value
    Assert.AreEqual Empty, ws.Range("E2").Value
    Assert.AreEqual "A", ws.Range("F2").Value
    Assert.AreEqual "B", ws.Range("G2").Value

    ''' --------------------------------------------------------
    ws.Cells.Clear

    PutVal 1, ws.Range("B2"), True

    '    |  A   B   C   D   E   F   G
    '----+---------------------------
    '  1 |
    '  2 |      1
    '  3 |
    '  4 |
    '  5 |
    '  6 |
    '  7 |

    Assert.AreEqual 1#, ws.Range("B2").Value

    ''' --------------------------------------------------------
    ws.Cells.Clear

    PutVal Array(1#, 2#, 3#, Empty, "A", "B"), ws.Range("B2"), True

    arrBase1 = GetVal(ws.Range("B2:B7"))
    ws.Cells.Clear
    PutVal arrBase1, ws.Range("B2")

    '    |  A   B   C   D   E   F   G
    '----+---------------------------
    '  1 |
    '  2 |      1
    '  3 |      2
    '  4 |      3
    '  5 |
    '  6 |      A
    '  7 |      B

    Assert.AreEqual 1#, ws.Range("B2").Value
    Assert.AreEqual 2#, ws.Range("B3").Value
    Assert.AreEqual 3#, ws.Range("B4").Value
    Assert.AreEqual Empty, ws.Range("B5").Value
    Assert.AreEqual "A", ws.Range("B6").Value
    Assert.AreEqual "B", ws.Range("B7").Value

    ''' ----------------------------------------------------
    ws.Cells.Clear

    PutVal Array( _
        Array(1, 2, 3, Empty, "a"), _
        Array(4, 5, 6, Empty, "b"), _
        Array(7, 8, 9, Empty, "c"), _
        Array(Empty, Empty, Empty, Empty, Empty), _
        Array("d", "e", "f", Empty, Empty) _
    ), ws.Range("B2")

    arrBase1 = GetVal(ws.Range("B2:F6"))
    ws.Cells.Clear
    PutVal arrBase1, ws.Range("B2")

    '    |  A   B   C   D   E   F
    '----+------------------------
    '  1 |
    '  2 |      1   2   3       a
    '  3 |      4   5   6       b
    '  4 |      7   8   9       c
    '  5 |
    '  6 |      d   e   f

    Assert.AreEqual 1#, ws.Range("B2").Value
    Assert.AreEqual 2#, ws.Range("C2").Value
    Assert.AreEqual 3#, ws.Range("D2").Value
    Assert.AreEqual Empty, ws.Range("E2").Value
    Assert.AreEqual "a", ws.Range("F2").Value

    Assert.AreEqual 4#, ws.Range("B3").Value
    Assert.AreEqual 5#, ws.Range("C3").Value
    Assert.AreEqual 6#, ws.Range("D3").Value
    Assert.AreEqual Empty, ws.Range("E3").Value
    Assert.AreEqual "b", ws.Range("F3").Value

    Assert.AreEqual 7#, ws.Range("B4").Value
    Assert.AreEqual 8#, ws.Range("C4").Value
    Assert.AreEqual 9#, ws.Range("D4").Value
    Assert.AreEqual Empty, ws.Range("E4").Value
    Assert.AreEqual "c", ws.Range("F4").Value

    Assert.AreEqual Empty, ws.Range("B5").Value
    Assert.AreEqual Empty, ws.Range("C5").Value
    Assert.AreEqual Empty, ws.Range("D5").Value
    Assert.AreEqual Empty, ws.Range("E5").Value
    Assert.AreEqual Empty, ws.Range("F5").Value

    Assert.AreEqual "d", ws.Range("B6").Value
    Assert.AreEqual "e", ws.Range("C6").Value
    Assert.AreEqual "f", ws.Range("D6").Value
    Assert.AreEqual Empty, ws.Range("E6").Value
    Assert.AreEqual Empty, ws.Range("F6").Value

    ''' ----------------------------------------------------
    ws.Cells.Clear

    PutVal Array( _
        Array(1, 2, 3, Empty, "a"), _
        Array(4, 5, 6, Empty, "b"), _
        Array(7, 8, 9, Empty, "c"), _
        Array(Empty, Empty, Empty, Empty, Empty), _
        Array("d", "e", "f", Empty, Empty) _
    ), ws.Range("B2"), True

    arrBase1 = GetVal(ws.Range("B2:F6"))
    ws.Cells.Clear
    PutVal arrBase1, ws.Range("B2")

    '    |  A   B   C   D   E   F
    '----+------------------------
    '  1 |
    '  2 |      1   4   7       d
    '  3 |      2   5   8       e
    '  4 |      3   6   9       f
    '  5 |
    '  6 |      a   b   c

    Assert.AreEqual 1#, ws.Range("B2").Value
    Assert.AreEqual 2#, ws.Range("B3").Value
    Assert.AreEqual 3#, ws.Range("B4").Value
    Assert.AreEqual Empty, ws.Range("B5").Value
    Assert.AreEqual "a", ws.Range("B6").Value

    Assert.AreEqual 4#, ws.Range("C2").Value
    Assert.AreEqual 5#, ws.Range("C3").Value
    Assert.AreEqual 6#, ws.Range("C4").Value
    Assert.AreEqual Empty, ws.Range("C5").Value
    Assert.AreEqual "b", ws.Range("C6").Value

    Assert.AreEqual 7#, ws.Range("D2").Value
    Assert.AreEqual 8#, ws.Range("D3").Value
    Assert.AreEqual 9#, ws.Range("D4").Value
    Assert.AreEqual Empty, ws.Range("D5").Value
    Assert.AreEqual "c", ws.Range("D6").Value

    Assert.AreEqual Empty, ws.Range("E2").Value
    Assert.AreEqual Empty, ws.Range("E3").Value
    Assert.AreEqual Empty, ws.Range("E4").Value
    Assert.AreEqual Empty, ws.Range("E5").Value
    Assert.AreEqual Empty, ws.Range("E6").Value

    Assert.AreEqual "d", ws.Range("F2").Value
    Assert.AreEqual "e", ws.Range("F3").Value
    Assert.AreEqual "f", ws.Range("F4").Value
    Assert.AreEqual Empty, ws.Range("F5").Value
    Assert.AreEqual Empty, ws.Range("F6").Value

    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True

End Sub

Public Sub LastRow_Test()

    Dim ws As Worksheet: Set ws = AddSheet("myTestTemporarySheet")

    '    |  A   B   C   D   E   F   G
    '----+---------------------------
    '  1 |
    '  2 |      1
    '  3 |      2
    '  4 |      3
    '  5 |
    '  6 |      A
    '  7 |      B

    ws.Range("B2").Value = 1
    ws.Range("B3").Value = 2
    ws.Range("B4").Value = 3
    ws.Range("B5").Value = Empty
    ws.Range("B6").Value = "A"
    ws.Range("B7").Value = "B"

    Assert.AreEqual 7&, LastRow(ws.Range("B2"))
    Assert.AreEqual 4&, LastRow(ws.Range("B2"), True)

    DeleteSheet "myTestTemporarySheet"

End Sub

Public Sub LastCol_Test()

    Dim ws As Worksheet: Set ws = AddSheet("myTestTemporarySheet")

    '    |  A   B   C   D   E   F   G
    '----+---------------------------
    '  1 |
    '  2 |      1   2   3       A   B
    '  3 |
    '  4 |
    '  5 |
    '  6 |
    '  7 |

    ws.Range("B2").Value = 1
    ws.Range("C2").Value = 2
    ws.Range("D2").Value = 3
    ws.Range("E2").Value = Empty
    ws.Range("F2").Value = "A"
    ws.Range("G2").Value = "B"

    Assert.AreEqual 7&, LastCol(ws.Range("B2"))
    Assert.AreEqual 4&, LastCol(ws.Range("B2"), True)

    DeleteSheet "myTestTemporarySheet"

End Sub

Public Sub Hankaku_Test()

    Dim ws As Worksheet: Set ws = AddSheet("myTestTemporarySheet")
    ws.Range("A1").Value = "ƒAƒCƒEƒGƒI"
    ws.Range("A2").Value = "‚`‚a‚b‚c‚d"
    ws.Range("A3").Value = "‚P‚Q‚R‚S‚T"

    Hankaku ws

    Assert.AreEqual "±²³´µ", ws.Range("A1").Value
    Assert.AreEqual "ABCDE", ws.Range("A2").Value
    Assert.AreEqual 12345#, ws.Range("A3").Value

    DeleteSheet "myTestTemporarySheet"

End Sub

''' --------------------------------------------------------
'''                        AdoEx Test
''' --------------------------------------------------------

Public Sub AdoEx_Test()

    Dim ws As Worksheet: Set ws = AddSheet("myTestTemporarySheet")
    ws.Cells.Clear

    PutVal Array( _
        Array("Number", 1, 2, 3, 4, 5), _
        Array("Name", "Apple", "Banana", "Cherry", Empty, Empty), _
        Array("Price", 200, 300, 400, Empty, 600), _
        Array("Memo1", Empty, Empty, Empty, Empty, Empty), _
        Array("Memo2", "d", "e", "f", Empty, Empty) _
    ), ws.Range("B2"), True

    '''     |   A        B         C         D        E        F
    ''' ----+------- --------- --------- --------- ------- ---------
    '''   1 |
    '''   2 |          Number    Name      Price    Memo1    Memo2
    '''   3 |            1       Apple      200                d
    '''   4 |            2       Banana     300                e
    '''   5 |            3       Cherry     400                f
    '''   6 |            4
    '''   7 |            5                  600

    Dim adox As AdoEx: Set adox = New AdoEx
    adox.Init dbExcel

    Dim arr As Variant: arr = adox.Select_("select * from [myTestTemporarySheet$B2:F7]")

    Assert.AreEqualArr Array("Number", "Name", "Price", "Memo1", "Memo2"), arr(0)

    Assert.AreEqualArr Array(1#, "Apple", 200#, Null, "d"), arr(1)(0)
    Assert.AreEqualArr Array(2#, "Banana", 300#, Null, "e"), arr(1)(1)
    Assert.AreEqualArr Array(3#, "Cherry", 400#, Null, "f"), arr(1)(2)
    Assert.AreEqualArr Array(4#, Null, Null, Null, Null), arr(1)(3)
    Assert.AreEqualArr Array(5#, Null, 600#, Null, Null), arr(1)(4)

    DeleteSheet "myTestTemporarySheet"

End Sub

Public Sub AdoExHeadless_Test()

    ''' Header as data

    Dim ws As Worksheet: Set ws = AddSheet("myTestTemporarySheet")
    ws.Cells.Clear

    PutVal Array( _
        Array("Number", 1, 2, 3, 4, 5), _
        Array("Name", "Apple", "Banana", "Cherry", Empty, Empty), _
        Array("Price", 200, 300, 400, Empty, 600), _
        Array("Memo1", Empty, Empty, Empty, Empty, Empty), _
        Array("Memo2", "d", "e", "f", Empty, Empty) _
    ), ws.Range("B2"), True

    '''     |   A        B         C         D        E        F
    ''' ----+------- --------- --------- --------- ------- ---------
    '''   1 |
    '''   2 |          Number    Name      Price    Memo1    Memo2
    '''   3 |            1       Apple      200                d
    '''   4 |            2       Banana     300                e
    '''   5 |            3       Cherry     400                f
    '''   6 |            4
    '''   7 |            5                  600

    Dim adox As AdoEx: Set adox = New AdoEx
    adox.Init dbExcel, , False

    Dim arr As Variant: arr = adox.Select_("select * from [myTestTemporarySheet$B2:F7]")

    Assert.AreEqualArr Array("F1", "F2", "F3", "F4", "F5"), arr(0)

    ''' B column type is Double. ---> "Number" strings shows Null.
    ''' D column type is Double. ---> "Price" strings shows Null.
    Assert.AreEqualArr Array(Null, "Name", Null, "Memo1", "Memo2"), arr(1)(0)
    Assert.AreEqualArr Array(1#, "Apple", 200#, Null, "d"), arr(1)(1)
    Assert.AreEqualArr Array(2#, "Banana", 300#, Null, "e"), arr(1)(2)
    Assert.AreEqualArr Array(3#, "Cherry", 400#, Null, "f"), arr(1)(3)
    Assert.AreEqualArr Array(4#, Null, Null, Null, Null), arr(1)(4)
    Assert.AreEqualArr Array(5#, Null, 600#, Null, Null), arr(1)(5)

    DeleteSheet "myTestTemporarySheet"

End Sub
