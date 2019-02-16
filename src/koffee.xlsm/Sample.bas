Attribute VB_Name = "Sample"
''' --------------------------------------------------------
'''  FILE    : Sample.bas
'''  AUTHOR  : callmekohei <callmekohei at gmail.com>
'''  License : MIT license
''' --------------------------------------------------------

Option Explicit
Option Private Module

''' --------------------------------------------------------
'''                     Page Break Line
''' --------------------------------------------------------

Private Sub PageBreakPreview()

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("foo")

    ''' Reset page break line
    ws.ResetAllPageBreaks

    ''' Vertical page break line
    ws.VPageBreaks.Add Range("O1")

    ''' Holizontal page break line
    Dim v As Variant
    For Each v In ArrRange(6, 26, 5)
        ws.HPageBreaks.Add ws.Cells(v, 1)
    Next v

    ''' Bold page break line
    ws.PageSetup.PrintArea = ws.Range("A1:N25").Address

End Sub

''' --------------------------------------------------------
'''                        WorkBooks
''' --------------------------------------------------------

Private Sub Read_WorkBook()

    ExcelStatus False, xlCalculationManual, False, False, True, False

    Dim srcWb As Workbook:  Set srcWb = CreateWorkBook("\\vmware-host\Shared Folders\tmp_icloud\test.xlsx")
    Dim srcWs As Worksheet: Set srcWs = srcWb.Worksheets("nor")
    Dim srcRng As Range:    Set srcRng = srcWs.Range("A1")

    Debug.Print srcRng.Value ''' Hello World!

    srcWb.Close

    ExcelStatus

End Sub

Private Sub Write_WorkBook()

    ExcelStatus False, xlCalculationManual, False, False, True, False

    Dim srcWb As Workbook:  Set srcWb = CreateWorkBook("\\vmware-host\Shared Folders\tmp_icloud\test.xlsx")
    Dim srcWs As Worksheet: Set srcWs = srcWb.Worksheets("nor")
    Dim srcRng As Range:    Set srcRng = srcWs.Range("A1")

    Debug.Print srcRng.Value ''' Hello World!
    srcRng.Value = "foo! bar! baz!"
    Debug.Print srcRng.Value ''' foo! bar! baz!

    srcWb.Save ''' Point!
    srcWb.Close

    ExcelStatus
End Sub

''' --------------------------------------------------------
'''                        WorkSheets
''' --------------------------------------------------------

Private Sub ReadAndWrite_WorkSheet()

    Dim srcWb As Workbook:  Set srcWb = ThisWorkbook
    Dim srcWs As Worksheet: Set srcWs = srcWb.Worksheets("foo")
    Dim srcRng As Range:    Set srcRng = srcWs.Range("A1")

    Debug.Print srcRng.Value ''' Hello World!
    srcRng.Value = "foo! bar! baz!"
    Debug.Print srcRng.Value ''' foo! bar! baz!

End Sub

''' -------------------------------------------------------
'''                      Copy and Paste
''' --------------------------------------------------------

Private Sub CopyAndPasteWithPutVal()

    '''     |  A   B   C   D        |  A   B   C   D   E
    ''' ----+----------------   ----+--------------------
    '''   1 |  a                  1 |  a       a
    '''   2 |  b                  2 |  b       b
    '''   3 |  c                  3 |  c       c
    '''   4 |                     4 |
    '''   5 |                     5 |

    Dim srcWb As Workbook:  Set srcWb = ThisWorkbook
    Dim srcWs As Worksheet: Set srcWs = srcWb.Worksheets("zzz")
    Dim srcRng As Range:    Set srcRng = srcWs.Range("A1:A3")

    Dim dstWb As Workbook:  Set dstWb = ThisWorkbook
    Dim dstWs As Worksheet: Set dstWs = dstWb.Worksheets("zzz")
    Dim dstRng As Range:    Set dstRng = dstWs.Range("C1")

    ''' Copy
    Dim arr As Variant: arr = GetVal(srcRng, True)

    ''' Paste with PutVal
    PutVal arr, dstRng, True

End Sub

Sub CopyAndPasteWithOffset()

    '''     |  A   B   C   D        |  A   B   C   D
    ''' ----+----------------   ----+----------------
    '''   1 |  a                  1 |  a       a
    '''   2 |  b                  2 |  b
    '''   3 |  c                  3 |  c       b
    '''   4 |                     4 |
    '''   5 |                     5 |          c

    Dim srcWb As Workbook:  Set srcWb = ThisWorkbook
    Dim srcWs As Worksheet: Set srcWs = srcWb.Worksheets("zzz")
    Dim srcRng As Range:    Set srcRng = srcWs.Range("A1:A3")

    Dim dstWb As Workbook:  Set dstWb = ThisWorkbook
    Dim dstWs As Worksheet: Set dstWs = dstWb.Worksheets("zzz")
    Dim dstRng As Range:    Set dstRng = dstWs.Range("C1")


    ''' Copy
    Dim arr As Variant: arr = GetVal(srcRng, True)

    ''' Paste with Offset
    Const Interval As Long = 2
    Dim v As Variant, row_i As Long
    For Each v In arr(1)
        dstRng.offset((IncrPst(row_i)) * Interval, 0).Value = v
    Next v

End Sub

''' xlClipboardFormat enumeration (Excel) : https://docs.microsoft.com/en-us/office/vba/api/excel.xlclipboardformat
''' xlClipboardFormatBitmap is 9&
Sub CopyAsBipmap_PasteAsPicture()

    ''' TODO : more simple

    Dim srcWb As Workbook:  Set srcWb = ThisWorkbook
    Dim srcWs As Worksheet: Set srcWs = srcWb.Worksheets("’ •[")
    Dim srcRng As Range:    Set srcRng = srcWs.Range("B2:N6")

    Dim dstWb As Workbook:  Set dstWb = ThisWorkbook
    Dim dstWs As Worksheet: Set dstWs = dstWb.Worksheets("Pic")
    Dim dstRng As Range:    Set dstRng = dstWs.Range("A1")

    ''' Copy and Paste( needs alternate cause clipboard has only one data )
    Dim row_i As Variant
    For Each row_i In ArrRange2(0, 4, 5)
        srcRng.offset(row_i, 0).CopyPicture xlScreen, xlBitmap
        dstRng.offset(row_i, 0).PasteSpecial
    Next row_i

End Sub

''' --------------------------------------------------------
'''                           misc
''' --------------------------------------------------------

Private Sub foo20190215()

    ''' Worksheet
    '''
    '''     |  A   B   C   D   E
    ''' ----+--------------------
    '''   1 |
    '''   2 |  1   x
    '''   3 |              foo001
    '''   4 |  2   y
    '''   5 |
    '''   6 |  3   z
    '''   7 |              bar002

    ''' Result
    '''
    ''' Array()
    ''' Array(1#, "x", "foo001")
    ''' Array(2#, "y")
    ''' Array(3#, "z", "foo002")

    ''' @param
    Dim ws As Worksheet:     Set ws = Worksheets("suntory")
    Dim rng As Range:        Set rng = ws.Range("A2:D7")
    Dim ptrnFind1 As String: ptrnFind1 = "\d.*"
    Dim ptrnFind2 As String: ptrnFind2 = ".*\d\d\d"

    Dim arrx As ArrayEx: Set arrx = New ArrayEx
    Dim arr As Variant, tmp As Variant, tmpArrx As ArrayEx: Set tmpArrx = New ArrayEx
    For Each arr In GetVal(rng)
        If Not ArrLen(ArrUniq(arr)) = 1 Then
            If ArrLen(ReMatch(arr(1), ptrnFind1)) > 0 Then
                arrx.AddVal tmpArrx.ToArray
                Set tmpArrx = Nothing
                Set tmpArrx = New ArrayEx
                tmpArrx.AddVal arr
            Else
                Dim v As Variant
                For Each v In arr
                    If ArrLen(ReMatch(v, ptrnFind2)) > 0 Then tmpArrx.AddVal arr
                Next v
            End If
        End If
    Next arr

    arrx.AddVal tmpArrx.ToArray

    Dim vv As Variant
    For Each vv In arrx.ToArray
        Debug.Print Dump(ArrRemoveEmpty(ArrFlatten(vv)))
    Next vv

End Sub
