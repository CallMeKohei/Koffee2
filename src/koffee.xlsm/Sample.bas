Attribute VB_Name = "Sample"
Option Explicit

Public Sub Sample_PageBreakPreview()

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("foo")

    ws.ResetAllPageBreaks

    ws.VPageBreaks.Add Range("O1")

    Dim v As Variant
    For Each v In ArrRange(6, 26, 5)
        ws.HPageBreaks.Add ws.Cells(v, 1)
    Next v

    ws.PageSetup.PrintArea = ws.Range("A1:N25").Address

End Sub

Sub foo()

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("bbb")
    
    Dim i As Long
    For i = 1 To 10
        InsertRows xlUpRange(ws.Range("b6")), "\d-\d.*", 3
        Application.Wait [Now()] + 500 / 86400000
        DeleteRows xlUpRange(ws.Range("b6")), "\d-\d.*", 3, -1
        Application.Wait [Now()] + 500 / 86400000
    Next i
    
End Sub

Sub Sample_InsertRows()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("bbb")
    InsertRows xlUpRange(ws.Range("b6")), "\d-\d.*", 3
End Sub

Sub Sample_DeleteRows()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("bbb")
    DeleteRows xlUpRange(ws.Range("b6")), "\d-\d.*", 3, -1
End Sub

Sub Sample_CopyFromOtherWorkBook()

'    ExcelStatus False, xlCalculationManual, False, False, True, False

'    Dim excelApp As Excel.Application
    Dim wb As Workbook: Set wb = CreateWorkBook("\\vmware-host\Shared Folders\tmp_icloud\test.xlsx")

    ''' ---------- Body ----------

    Dim ws As Worksheet: Set ws = wb.Worksheets(1)
    Dim v As Variant
'    For Each v In GetVal2(ws.Range("B6:B" & LastRow(ws.Range("B6"))), , "\d-\d.*")(0).Items
'        Debug.Print v
'    Next v
    
    Debug.Print "---"
    
    For Each v In RegexRanges(xlUpRanges(ws.Range("b6")), "\d-\d.*", True)
        Debug.Print v
    Next v

    ''' --------------------------

    CloseWorkBook wb
'    ExcelStatus

End Sub

Sub Sample_CopyFromOtherWorkBook2()

    Dim wb As Workbook: Set wb = CreateWorkBook("\\vmware-host\Shared Folders\tmp_icloud\test.xlsx")

    ''' ---------- Body ----------

    Dim ws As Worksheet: Set ws = wb.Worksheets("‚ ‚¢‚¤‚¦‚¨")
    Dim srcArr As Variant: srcArr = GetVal(ws.Range("b5:e18"), True)
    
    Dim dstWs As Worksheet: Set dstWs = ThisWorkbook.Worksheets("‚©‚«‚­‚¯‚±")
    Dim dstRng As Range: Set dstRng = dstWs.Range("B2")

    PutVal srcArr, dstRng.offset(0, 2), True
    PutVal ArrPadLeft(srcArr(1)), dstRng.offset(0, 0), True

    ''' --------------------------

    CloseWorkBook wb

End Sub



Sub Sample_CopyAndPaste()

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("’ •[")
    Dim rng As Range: Set rng = ws.Range("B2:N6")
    Const dstInterval As Long = 5

    rng.Copy

    Dim row_i As Long
    For row_i = 0 To 4
        rng.offset(row_i * dstInterval, 0).PasteSpecial xlPasteFormats
    Next row_i

End Sub

Sub Sample_CopyAndPasteFromOtherWorksheet()

    Dim srcWs As Worksheet: Set srcWs = ThisWorkbook.Worksheets("aaa")
'    Dim srcArr As Variant: srcArr = GetVal2(srcWs.Range("D1:D5"), , "\d-\d:.*")(0).Items
    Dim srcArr As Variant: srcArr = RegexRanges(srcWs.Range("D1:D5"), "\d-\d:.*")
    

    Dim dstWs As Worksheet: Set dstWs = ThisWorkbook.Worksheets("’ •[")
    Dim dstRng As Range: Set dstRng = dstWs.Range("B2")
    Const dstInterval As Long = 5

    Dim i As Long, v As Variant
    For Each v In srcArr
        dstRng.offset((IncrPst(i)) * dstInterval, 0).Value = v.Value
    Next v

End Sub

Sub Sample_ModifyOtherWorkBook()

'    ExcelStatus False, xlCalculationManual, False, False, True, False

    'Dim excelApp As Excel.Application
    Dim wb As Workbook: Set wb = CreateWorkBook("\\vmware-host\Shared Folders\tmp_icloud\test.xlsx", False)
'    Dim wb As Workbook: Set wb = CreateWorkBook(excelApp, "\\vmware-host\Shared Folders\tmp_icloud\test.xlsx", False)

    ''' ---------- Body ----------

    Dim ws As Worksheet: Set ws = wb.Worksheets(1)
    Debug.Print (ws.Range("A1").Value)
    ws.Range("A1").Value = "ccc!"
    Debug.Print (ws.Range("A1").Value)

    ''' --------------------------

    SaveCloseWorkBook wb
'    SaveCloseWorkBook excelApp, wb
'    ExcelStatus

End Sub

''' xlClipboardFormat enumeration (Excel) : https://docs.microsoft.com/en-us/office/vba/api/excel.xlclipboardformat
''' xlClipboardFormatBitmap is 9&
Sub SampleCopyAsBipmap_pasteAsPicture()

    Dim srcWs As Worksheet: Set srcWs = ThisWorkbook.Worksheets("’ •[")
    Dim srcRng As Range: Set srcRng = srcWs.Range("B2:N6")

    Dim dstWs As Worksheet: Set dstWs = ThisWorkbook.Worksheets("Pic")
    Dim dstRng As Range: Set dstRng = dstWs.Range("A1")

    Dim row_i As Variant
    For Each row_i In ArrRange2(0, 4, 5)
        srcRng.offset(row_i, 0).CopyPicture xlScreen, xlBitmap
        dstRng.offset(row_i, 0).PasteSpecial
    Next row_i

End Sub










