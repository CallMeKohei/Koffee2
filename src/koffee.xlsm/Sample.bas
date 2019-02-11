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

Sub Sample_CopyFromOtherWorkBook()

    ExcelStatus False, xlCalculationManual, False, False, True, False
    
    Dim excelApp As Excel.Application
    Dim wb As Workbook: Set wb = CreateWorkBook(excelApp, "C:\Users\mochi\tmp\test.xlsx")
    
    ''' ---------- Body ----------
    
    Dim ws As Worksheet: Set ws = wb.Worksheets(1)
    Debug.Print Dump(GetVal2(ws.Range("B6:B" & LastRow(ws.Range("B6"))), , "\d-\d.*")(0).Items)
    
    ''' --------------------------
    
    CloseWorkBook excelApp, wb
    ExcelStatus

End Sub

Sub Sample_ModifyOtherWorkBook()
    
    ExcelStatus False, xlCalculationManual, False, False, True, False
    
    Dim excelApp As Excel.Application
    Dim wb As Workbook: Set wb = CreateWorkBook(excelApp, "C:\Users\mochi\tmp\test.xlsx", False)
        
    ''' ---------- Body ----------
    
    Dim ws As Worksheet: Set ws = wb.Worksheets(1)
    Debug.Print (ws.Range("A1").Value)
    ws.Range("A1").Value = "cccc!"
    Debug.Print (ws.Range("A1").Value)
    
    ''' --------------------------
    
    SaveCloseWorkBook excelApp, wb
    ExcelStatus

End Sub

Sub Sample_CopyAndPaste()

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("’ •[")
    Dim rng As Range: Set rng = ws.Range("B2:N6")
    Const dstInterval As Long = 5

    rng.Copy
    
    Dim row_i As Long
    For row_i = 0 To 4
        rng.Offset(row_i * dstInterval, 0).PasteSpecial xlPasteFormats
    Next row_i

End Sub

Sub Sample_CopyAndPasteFromOtherWorksheet()

    Dim srcWs As Worksheet: Set srcWs = ThisWorkbook.Worksheets("aaa")
    Dim srcArr As Variant: srcArr = GetVal2(srcWs.Range("D1:D5"), , "\d-\d:.*")(0).Items
    
    Dim dstWs As Worksheet: Set dstWs = ThisWorkbook.Worksheets("’ •[")
    Dim dstRng As Range: Set dstRng = dstWs.Range("B2")
    Const dstInterval As Long = 5

    Dim i As Long, v As Variant
    For Each v In srcArr
        dstRng.Offset((IncrPst(i)) * dstInterval, 0).Value = v
    Next v

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
        srcRng.Offset(row_i, 0).CopyPicture xlScreen, xlBitmap
        dstRng.Offset(row_i, 0).PasteSpecial
    Next row_i

End Sub

Sub Sample_GetVal2()
    Dim v As Variant
    For Each v In GetVal2(Worksheets("nor").Range("B3:B13"), , "\d-\d.*")(0).Keys
        Debug.Print Dump(v)
    Next v
    For Each v In GetVal2(Worksheets("nor").Range("B3:B13"), , "\d-\d.*")(0).Items
        Debug.Print Dump(v)
    Next v
End Sub

Sub Sample_insertRows()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("bbb")
    InsertRow ws.Range("B6"), 3, "\d-\d.*"
End Sub

Sub Sample_deleteRows()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("bbb")
    DeletRow ws.Range("B6"), 3, "\d-\d.*"
End Sub




