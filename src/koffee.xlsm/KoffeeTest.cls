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


''' ----- koffeeArray test

Public Sub IsJaggedArray_Test()

    ''' Jagged arrays is "Array of arrays".
    Assert.AreEqual True, IsJaggedArray(Array(Array(0)))

    ''' An inner array which is not initialized is regarded as an array.
    Assert.AreEqual True, IsJaggedArray(Array(Array()))
    Assert.AreEqual True, IsJaggedArray(Array(Array(), Array(0)))

    ''' All inner elements must be array.
    Assert.AreEqual False, IsJaggedArray(0)
    Assert.AreEqual False, IsJaggedArray(New Collection)
    Assert.AreEqual False, IsJaggedArray(Array())
    Assert.AreEqual False, IsJaggedArray(Array(Array(0), 0))
    Assert.AreEqual False, IsJaggedArray(Array(Array(0), New Collection))

    ''' One Dimension
    Dim arr1d(0)
    Assert.AreEqual False, IsJaggedArray(arr1d)

    Dim negativeArr1D(-1 To -1)
    Assert.AreEqual False, IsJaggedArray(negativeArr1D)

    ''' Two Dimensions
    Dim arr2D(0, 0)
    Assert.AreEqual False, IsJaggedArray(arr2D)

    Dim negativeArr2D(-1 To -1, -1 To -1)
    Assert.AreEqual False, IsJaggedArray(negativeArr2D)

    ''' More Dimensions
    Dim arr3D(0, 0, 0)
    Assert.AreEqual False, IsJaggedArray(arr3D)

    Dim negativeArr3D(-1 To -1, -1 To -1, -1 To -1)
    Assert.AreEqual False, IsJaggedArray(negativeArr3D)

End Sub

Public Sub ArrayBase0_Test()
    Dim arr(): ReDim arr(1 To 99)
    sArrayBase0 arr
    Assert.AreEqual 0&, LBound(arr)
End Sub

'Public Sub ArrayBase0_2ndDimension_Test()
'    Dim arr(): ReDim arr(1 To 33, 1 To 44)
'    ArrayBase0_2ndDimension arr
'    Assert.AreEqual 0&, LBound(arr)
'End Sub

Public Sub ArrayColumn_Test()
    Dim arr: arr = Array(Array(1, 2), Array(3, 4), Array(5, 6))
    Assert.AreEqualArr Array(1, 3, 5), ArrayColumn(0, arr)
    Assert.AreEqualArr Array(2, 4, 6), ArrayColumn(1, arr)
End Sub

Public Sub ArraySlice_Test()
    Dim arr: arr = Array("a", "b", "c", "d", "e", "f", "g")
    Assert.AreEqualArr Array("b", "c", "d"), ArraySlice(arr, 1, 3)
End Sub

Public Sub ArrayRegexFilter_Test()
    Dim arr: arr = Array("15.0", "16.0", "16.0", "Common", "Outlook")
    Assert.AreEqualArr Array("15.0", "16.0"), ArrayRegexFilter(arr, "\d\d\.\d")
End Sub

Public Sub ArrayTranspose_Test()

    '''  Before           After
    '''
    '''  A  B  C          A  1
    '''  1  2  3   --->   B  2
    '''                   C  3

    ''' validation
    ''' not array
    Dim arr
    Assert.IsErrFunc 9, Init(New Func, vbVariant, AddressOf ArrayTranspose, vbVariant, vbVariant), Array(0), arr
    ''' not two dimensions
    Dim arr1d(0) As Variant
    Assert.IsErrFunc 9, Init(New Func, vbVariant, AddressOf ArrayTranspose, vbVariant, vbVariant), Array(arr1d), arr
    ''' not two dimensions
    Dim arr3D(0, 0, 0) As Variant
    Assert.IsErrFunc 9, Init(New Func, vbVariant, AddressOf ArrayTranspose, vbVariant, vbVariant), Array(arr3D), arr

    ''' Before ( Base 0 )
    Dim arr2D(1, 2) As Variant
    arr2D(0, 0) = "A"
    arr2D(0, 1) = "B"
    arr2D(0, 2) = "C"

    arr2D(1, 0) = 1
    arr2D(1, 1) = 2
    arr2D(1, 2) = 3

    Assert.AreEqual "A", arr2D(0, 0)
    Assert.AreEqual "B", arr2D(0, 1)
    Assert.AreEqual "C", arr2D(0, 2)

    Assert.AreEqual 1, arr2D(1, 0)
    Assert.AreEqual 2, arr2D(1, 1)
    Assert.AreEqual 3, arr2D(1, 2)

    ''' After ( Base 0 )
    Dim arr2DTrans As Variant: sArrayTranspose arr2D, arr2DTrans

    Assert.AreEqual "A", arr2DTrans(0, 0)
    Assert.AreEqual 1, arr2DTrans(0, 1)

    Assert.AreEqual "B", arr2DTrans(1, 0)
    Assert.AreEqual 2, arr2DTrans(1, 1)

    Assert.AreEqual "C", arr2DTrans(2, 0)
    Assert.AreEqual 3, arr2DTrans(2, 1)

    ''' Before ( Base 1 )
    Dim arr2DBase1 As Variant: arr2DBase1 = [{"A","B","C";1,2,3}]

    Assert.AreEqual "A", arr2DBase1(1, 1)
    Assert.AreEqual "B", arr2DBase1(1, 2)
    Assert.AreEqual "C", arr2DBase1(1, 3)

    Assert.AreEqual 1#, arr2DBase1(2, 1)
    Assert.AreEqual 2#, arr2DBase1(2, 2)
    Assert.AreEqual 3#, arr2DBase1(2, 3)

    ''' After ( Base 1 )
    Dim arr2DBase1Trans As Variant: sArrayTranspose arr2DBase1, arr2DBase1Trans

    Assert.AreEqual "A", arr2DBase1Trans(1, 1)
    Assert.AreEqual 1#, arr2DBase1Trans(1, 2)

    Assert.AreEqual "B", arr2DBase1Trans(2, 1)
    Assert.AreEqual 2#, arr2DBase1Trans(2, 2)

    Assert.AreEqual "C", arr2DBase1Trans(3, 1)
    Assert.AreEqual 3#, arr2DBase1Trans(3, 2)

End Sub

Public Sub ArrayHasEmpties_Test()
    Dim arr: arr = Array(1, 2, Empty, 4, 5)
    Dim ary: ary = Array(1, 2, "", 4, 5)
    Assert.AreEqual True, ArrayHasEmpties(arr)
    Assert.AreEqual True, ArrayHasEmpties(ary)
End Sub

Public Sub ArrayRemoveEmpties_Test()
    Dim arr: arr = Array(1, 2, Empty, 4, 5)
    Dim ary: ary = Array(1, 2, "", 4, 5)
    Assert.AreEqualArr Array(1, 2, 4, 5), ArrayRemoveEmpties(arr)
    Assert.AreEqualArr Array(1, 2, 4, 5), ArrayRemoveEmpties(ary)
End Sub

Public Sub ArrayWindow_Test()
    
'    Assert.IsErrFunc 13, Init(New Func, vbVariant, AddressOf ArrayWindow, vbVariant, vbVariant), Array(Array(1, 2, 3), -99)
'    Assert.IsErrFunc 13, Init(New Func, vbVariant, AddressOf ArrayWindow, vbVariant, vbVariant), Array(Array(1, 2, 3), 0)
    
    Dim arr
    
    sArrayWindow Array(1, 2, 3), 1, arr
    Assert.AreEqualArr Array(1, 2, 3), arr(0)
    
    arr = Empty
    sArrayWindow Array(1, 2, 3), 99, arr
    Assert.AreEqualArr Array(2), arr(1)
    
    
    Dim ary(): ary = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
    
    arr = Empty
    sArrayWindow ary, 2, arr
    Assert.AreEqualArr Array(1, 2, 3, 4, 5), arr(0)
    Assert.AreEqualArr Array(6, 7, 8, 9, 10), arr(1)
    
    arr = Empty
    sArrayWindow ary, 3, arr
    Assert.AreEqualArr Array(1, 2, 3, 4), arr(0)
    Assert.AreEqualArr Array(5, 6, 7), arr(1)
    Assert.AreEqualArr Array(8, 9, 10), arr(2)

End Sub


''' ----- koffeeExcel test

Public Sub ArrSheetsName_Test()

    Dim arr As Variant: arr = ArrSheetsName()

    Dim Wb As Workbook: Set Wb = Application.ThisWorkbook

    Assert.AreEqual Wb.Worksheets.Count, ArrayLen(arr)

    Dim ws As Worksheet, i As Long
    For Each ws In Wb.Worksheets
        Assert.AreEqual ws.Name, arr(i)
        i = i + 1
    Next ws

End Sub

Public Sub ExistsSheet_Test()

    Assert.AreEqual False, ExistsSheet("abcdefghijklmnopqrstuvwxyz")

    Dim WORKSHEET_NAME As Variant
    For Each WORKSHEET_NAME In ArrSheetsName()
        Assert.AreEqual True, ExistsSheet(WORKSHEET_NAME)
    Next WORKSHEET_NAME

End Sub

Public Sub DeleteSheet_Test()

    Assert.AreEqual False, ExistsSheet("myTestTemporarySheet")
    Application.ThisWorkbook.Worksheets.add().Name = "myTestTemporarySheet"
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

    Assert.AreEqualArr Array(1#), GetVal(ws.Range("B2"))(0)

    Assert.AreEqualArr Array(1#, 2#, Empty, "a"), GetVal(ws.Range("B2:E2"))(0)

    Assert.AreEqualArr Array(1#), GetVal(ws.Range("B2:B5"))(0)
    Assert.AreEqualArr Array(3#), GetVal(ws.Range("B2:B5"))(1)
    Assert.AreEqualArr Array(Empty), GetVal(ws.Range("B2:B5"))(2)
    Assert.AreEqualArr Array("c"), GetVal(ws.Range("B2:B5"))(3)

    Assert.AreEqualArr Array(1#, 2#, Empty, "a"), GetVal(ws.Range("B2:E5"))(0)
    Assert.AreEqualArr Array(3#, 4#, Empty, "b"), GetVal(ws.Range("B2:E5"))(1)
    Assert.AreEqualArr Array(Empty, Empty, Empty, Empty), GetVal(ws.Range("B2:E5"))(2)
    Assert.AreEqualArr Array("c", "d", Empty, Empty), GetVal(ws.Range("B2:E5"))(3)

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


''' ----- AdoEx test

Public Sub AdoEx_Test()
    Dim ws As Worksheet
    If Not ExistsSheet("myTestTemporarySheet") Then
        Set ws = AddSheet("myTestTemporarySheet")
    Else
        Set ws = Worksheets("myTestTemporarySheet")
    End If
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

    Dim arr As Variant: arr = koffeeArray.SelectExcelHeader("select * from [myTestTemporarySheet$B2:F7]")

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

    Dim arr As Variant: arr = koffeeArray.SelectExcel("select * from [myTestTemporarySheet$B2:F7]")

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
