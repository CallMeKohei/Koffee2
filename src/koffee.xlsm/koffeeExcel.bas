Attribute VB_Name = "koffeeExcel"
''' koffeeExcel.bas
''' written by callmekohei(twitter at callmekohei)
''' MIT license
Option Explicit
Option Compare Text
Option Private Module
Option Base 0

''' ----- Workbook

Public Function CreateWorkBook(ByVal FilePath As String, Optional isReadOnly As Boolean = True) As Workbook
    Dim flg As Boolean: flg = IsWorkBookClosed(FilePath)
    If Not (isReadOnly Or flg) Then Err.Raise 9999, , "WorkBook is already opened."
    If flg Then
            Set CreateWorkBook = Workbooks.Open(FileName:=FilePath, UpdateLinks:=0, readOnly:=isReadOnly, IgnoreReadOnlyRecommended:=True)
    Else
        ''' create anthor Excel application process
        Dim excelApp As Excel.Application: Set excelApp = New Excel.Application
        Set CreateWorkBook = excelApp.Workbooks.Open(FileName:=FilePath, UpdateLinks:=0, readOnly:=isReadOnly, IgnoreReadOnlyRecommended:=True)
    End If
End Function

Private Function IsWorkBookClosed(ByVal FilePath As String) As Boolean
    On Error GoTo Escape
        Open FilePath For Append As #1
        Close #1
    On Error GoTo 0
        IsWorkBookClosed = True
Escape:
End Function


''' ----- Worksheet

Public Function ArrSheetsName(Optional ByVal Wb As Workbook = Nothing) As Variant

    ''' ( Usage )
    ''' Dim wb As Workbook: Set wb = Application.ThisWorkbook
    ''' Debug.Print ArrSheetsName(wb)(0)

    If TypeName(Wb) = "Nothing" Then Set Wb = Application.ThisWorkbook

    Dim arr() As String
    ReDim arr(0 To Wb.Sheets.Count - 1)

    Dim ws As Worksheet, i As Long
    For Each ws In Wb.Worksheets
        arr(i) = ws.Name
        i = i + 1
    Next ws

    ArrSheetsName = arr

End Function

''' Debug.Print ExistsSheet("abc")
Public Function ExistsSheet(ByVal SheetName As String, Optional ByVal Wb As Workbook = Nothing) As Boolean

    If TypeName(Wb) = "Nothing" Then Set Wb = Application.ThisWorkbook

    Dim v As Variant
    For Each v In ArrSheetsName(Wb)
        If SheetName = v Then
            ExistsSheet = True
            GoTo Escape
        End If
    Next v

Escape:
End Function

''' DeleteSheet "abc"
Public Sub DeleteSheet(ByVal SheetName As String, Optional ByVal Wb As Workbook = Nothing)

    If TypeName(Wb) = "Nothing" Then Set Wb = Application.ThisWorkbook

    If Not ExistsSheet(SheetName, Wb) Then GoTo Catch

    Application.DisplayAlerts = False
    Wb.Worksheets(SheetName).Delete
    Application.DisplayAlerts = True
    GoTo Escape

Catch:
    Debug.Print "(DeleteSheet): The SheetName is not exists!"
    Exit Sub

Escape:
End Sub

''' Dim ws As Worksheet: Set ws = AddSheet("abc")
Public Function AddSheet(ByVal SheetName As String, Optional ByVal Wb As Workbook = Nothing) As Worksheet
    If TypeName(Wb) = "Nothing" Then Set Wb = Application.ThisWorkbook
    If ExistsSheet(SheetName, Wb) Then GoTo Catch
    Wb.Worksheets.add(after:=Worksheets(Wb.Worksheets.Count)).Name = SheetName
    Set AddSheet = Wb.Worksheets(SheetName)
    GoTo Escape
Catch:
    Debug.Print "(AddSheet): The SheetName is already exists!"
Escape:
End Function

Public Function CopySheet( _
      ByVal srcWsName As String _
    , Optional ByVal dstWsName As String = "" _
    , Optional ByVal srcWb As Workbook = Nothing _
    , Optional ByVal dstWb As Workbook = Nothing) As Worksheet

    ''' ( Usage )
    ''' Dim wsCopied As Worksheet: Set wsCopied = CopySheet("abc", "abcCopied")

    If TypeName(srcWb) = "Nothing" Then Set srcWb = Application.ThisWorkbook
    If TypeName(dstWb) = "Nothing" Then Set dstWb = Application.ThisWorkbook

    If Not ExistsSheet(srcWsName, srcWb) Then GoTo Catch
    If ExistsSheet(dstWsName, dstWb) Then GoTo Catch2

    srcWb.Worksheets(srcWsName).Copy after:=dstWb.Worksheets(dstWb.Sheets.Count)
    If Not (dstWsName = "") Then
        dstWb.ActiveSheet.Name = dstWsName
        Set CopySheet = dstWb.Worksheets(dstWsName)
        GoTo Escape
    Else
        Set CopySheet = dstWb.Worksheets(dstWb.Sheets.Count)
        GoTo Escape
    End If

Catch:
    Debug.Print "(CopySheet): The SouceSheet is not exists!"
    Exit Function

Catch2:
    Debug.Print "(CopySheet): The dstWsName is already exists!"
    Exit Function

Escape:
End Function

''' TODO: write test code!
Public Function MoveSheet( _
      ByVal srcWsName As String _
    , Optional ByVal dstWsName As String = "" _
    , Optional ByVal srcWb As Workbook = Nothing _
    , Optional ByVal dstWb As Workbook = Nothing) As Worksheet

    ''' ( Usage )
    ''' Dim wsCopied As Worksheet: Set wsCopied = CopySheet("abc", "abcCopied")

    If TypeName(srcWb) = "Nothing" Then Set srcWb = Application.ThisWorkbook
    If TypeName(dstWb) = "Nothing" Then Set dstWb = Application.ThisWorkbook

    If Not ExistsSheet(srcWsName, srcWb) Then GoTo Catch
    If ExistsSheet(dstWsName, dstWb) Then GoTo Catch2

    srcWb.Worksheets(srcWsName).Move after:=dstWb.Worksheets(dstWb.Sheets.Count)
    If Not (dstWsName = "") Then
        dstWb.ActiveSheet.Name = dstWsName
        Set MoveSheet = dstWb.Worksheets(dstWsName)
        GoTo Escape
    Else
        Set MoveSheet = dstWb.Worksheets(dstWb.Sheets.Count)
        GoTo Escape
    End If

Catch:
    Debug.Print "(CopySheet): The SrouceSheet is not exists!"
    Exit Function

Catch2:
    Debug.Print "(CopySheet): The dstWsName is already exists!"
    Exit Function

Escape:
End Function

''' ----- Cells(Ranges)

Public Function GetVal(ByRef rng As Range, Optional isVertical As Boolean = False)
    Dim arr
    sGetVal rng, arr, isVertical
    GetVal = arr
End Function

Public Function sGetVal(ByRef rng As Range, ByRef tmp As Variant, Optional isVertical As Boolean = False)

    Dim arr As Variant: arr = rng.Value

    If Not IsArray(arr) Then
        tmp = Array(Array(arr))
        GoTo Ending
    End If

    If isVertical Then
        Dim tmptmp()
        sArrayTransposeLet arr, tmptmp
        sArray2DToJagArrayLet tmptmp, tmp
    Else
        sArray2DToJagArrayLet arr, tmp
    End If

Ending:
End Function

Public Sub PutVal(ByRef arr As Variant, ByRef rng As Range, Optional isVertical As Boolean = False)

    If IsObject(arr) Then Err.Raise 13

    ''' Value to 2D array
    If Not IsArray(arr) Then
        Dim tmp(1 To 1, 1 To 1) As Variant: tmp(1, 1) = arr
        arr = tmp
    End If

    If ArrayRank(arr) >= 3 Then Err.Raise 13

    ''' 1D array to 2D array
    If ArrayRank(arr) = 1 Then
        Dim tmpArr
        If IsJaggedArray(arr) Then
            JagArrayToArray2DLet arr, tmpArr
        Else
            JagArrayToArray2DLet Array(arr), tmpArr
        End If
        
        PutValImpl tmpArr, rng, isVertical
    
    Else
        PutValImpl arr, rng, isVertical
    End If


End Sub

Private Sub PutValImpl(ByRef arr2D As Variant, ByRef rng As Range, Optional isVertical As Boolean = False)

    If Not ArrayRank(arr2D) = 2 Then Err.Raise 13

    If isVertical Then
        ''' Minimum index Excel's Array is 1
        If LBound(arr2D, 1) = 1 Then
            rng.Resize(UBound(arr2D, 2), UBound(arr2D, 1)).Value = ArrayTranspose(arr2D)
        Else
            rng.Resize(UBound(arr2D, 2) + 1, UBound(arr2D, 1) + 1).Value = ArrayTranspose(arr2D)
        End If
    Else
        ''' Minimum index Excel's Array is 1
        If LBound(arr2D, 1) = 1 Then
            rng.Resize(UBound(arr2D, 1), UBound(arr2D, 2)).Value = arr2D
        Else
            rng.Resize(UBound(arr2D, 1) + 1, UBound(arr2D, 2) + 1).Value = arr2D
        End If
    End If

End Sub

Public Function LastRow(ByVal rng As Range, Optional toDown As Boolean = False) As Long
    If toDown Then
        LastRow = rng.End(xlDown).row
    Else
        LastRow = rng.Worksheet.Cells(rng.Worksheet.Rows.Count, rng.Column).End(xlUp).row
    End If
End Function

Public Function LastCol(ByVal rng As Range, Optional toRight As Boolean = False) As Long
    If toRight Then
        LastCol = rng.End(xlToRight).Column
    Else
        LastCol = rng.Worksheet.Cells(rng.row, rng.Worksheet.Columns.Count).End(xlToLeft).Column
    End If
End Function

Public Sub Hankaku(ByVal ws As Worksheet)
    Dim v As Range
    For Each v In ws.UsedRange
        v.Value = StrConv(v.Value, vbNarrow)
    Next
End Sub

''' ----- General Operation

''' @seealso ScreenUpdating https://docs.microsoft.com/en-us/office/vba/api/excel.application.screenupdating (/ja-jp/office/vba/api/excel.application.statusbar)
''' @seealso Calculation    https://docs.microsoft.com/en-us/office/vba/api/excel.application.calculation    (/ja-jp/office/vba/api/excel.application.calculation)
''' @seealso EnableEvents   https://docs.microsoft.com/en-us/office/vba/api/excel.application.enableevents   (/ja-jp/office/vba/api/excel.application.enableevents)
''' @seealso DisplayAlerts  https://docs.microsoft.com/en-us/office/vba/api/excel.application.statusbar      (/ja-jp/office/vba/api/excel.application.statusbar)
''' @seealso StatusBar      https://docs.microsoft.com/en-us/office/vba/api/excel.application.displayalerts  (/ja-jp/office/vba/api/excel.application.displayalerts)

''' ExcelStatus False,xlCalculationManual,True,True,False,True
Public Sub ExcelStatus( _
    Optional ByVal aScreenUpDating As Boolean = True, _
    Optional ByVal aCalculation As XlCalculation = xlCalculationAutomatic, _
    Optional ByVal aEnableEvents As Boolean = True, _
    Optional ByVal aDisplayAlerts As Boolean = True, _
    Optional ByVal aStatusBar As Boolean = False, _
    Optional ByVal aDisplayStatusBar As Boolean = True)

    Application.ScreenUpdating = aScreenUpDating
    Application.Calculation = aCalculation
    Application.EnableEvents = aEnableEvents
    Application.DisplayAlerts = aDisplayAlerts
    Application.statusBar = aStatusBar
    Application.DisplayStatusBar = aDisplayStatusBar

End Sub

Public Sub ProtectSheet(ByVal ws As Worksheet, Optional myPassword As String = "1234")

    ws.Protect _
        Password:=myPassword, _
        DrawingObjects:=False, _
        contents:=True, _
        Scenarios:=True, _
        userinterfaceonly:=False, _
        AllowFormattingCells:=True, _
        AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, _
        AllowInsertingColumns:=False, _
        AllowInsertingRows:=True, _
        AllowInsertingHyperlinks:=True, _
        AllowDeletingColumns:=True, _
        AllowSorting:=True, _
        AllowFiltering:=True, _
        AllowUsingPivotTables:=True

End Sub


''' --------------------------------------------------------
''' ----- deprecated

'Public Function GetValStr(ByVal rng As Range, Optional isVertical As Boolean = False) As Variant
'
'    ''' All array elements are string type.
'    ''' Array("foo","123","","#2019/5/10#")
'
'    Dim tmp As Variant: tmp = GetVal(rng, isVertical)
'    Dim tmpArr() As Variant: ReDim tmpArr(0 To UBound(tmp) - 1)
'    Dim v As Variant, i As Long
'    For Each v In tmp
'        tmpArr(IncrPst(i)) = ArrMap(init(New Func, vbString, AddressOf ToStr2, vbVariant), v)
'    Next v
'
'    GetValStr = Base0(tmpArr)
'
'End Function
'
'Public Function RegexRanges(ByVal rng As Range _
'    , ByVal ptrnFind As String _
'    , Optional ByVal isVertical As Boolean = True _
'    ) As Variant
'
'    Dim ws As Worksheet: Set ws = rng.Worksheet
'    Dim inArr As Variant, arrx As ArrayEx: Set arrx = New ArrayEx
'    For Each inArr In GetVal(rng, isVertical)
'        Dim i As Long
'        For i = 1 To ArrLen(inArr)
'            If ArrLen(ReMatch(inArr(i), ptrnFind)) > 0 Then arrx.AddObj ws.Cells(rng.Offset(i - 1).row, rng.Column)
'        Next i
'    Next inArr
'
'    RegexRanges = arrx.ToArray
'
'End Function
'
'Public Sub InsertRows(ByVal rng As Range, ByVal ptrnFind As String _
'    , Optional ByVal times As Long = 1, Optional ByVal offsetRow As Long = 0, Optional ByVal offsetColumn As Long = 0)
'
'    ''' EXAMPLES :
'    ''' InsertRows xlUpRange(ws.Range("B6")), "\d-\d.*", 3
'
'    Dim i As Long
'    For i = 1 To times
'        If offsetRow = 0 And offsetColumn = 0 Then
'            UnionRanges(RegexRanges(rng, ptrnFind)).EntireRow.Insert
'        Else
'            UnionRanges(offsetRanges(RegexRanges(rng, ptrnFind), offsetRow, offsetColumn)).EntireRow.Insert
'        End If
'    Next i
'End Sub
'
'Public Sub DeleteRows(ByVal rng As Variant, ByVal ptrnFind As String _
'    , Optional ByVal times As Long = 1, Optional ByVal offsetRow As Long = 0, Optional ByVal offsetColumn As Long = 0)
'
'    ''' EXAMPLES :
'    ''' DeleteRows xlUpRange(ws.Range("B6")), "\d-\d.*", 3, -1
'
'    Dim i As Long
'    For i = 1 To times
'        If offsetRow = 0 And offsetColumn = 0 Then
'            UnionRanges(RegexRanges(rng, ptrnFind)).EntireRow.Delete
'        Else
'            UnionRanges(offsetRanges(RegexRanges(rng, ptrnFind), offsetRow, offsetColumn)).EntireRow.Delete
'        End If
'    Next i
'End Sub
'
'Public Function offsetRanges(ByVal arr As Variant, Optional ByVal offsetRow As Long = 0, Optional ByVal offsetColumn As Long = 0) As Variant
'    Dim rng As Variant, arrx As ArrayEx: Set arrx = New ArrayEx
'    For Each rng In arr
'        arrx.AddObj rng.Offset(offsetRow, offsetColumn)
'    Next rng
'    offsetRanges = arrx.ToArray
'End Function
'
'Public Function xlUpRange(ByVal rng As Range) As Range
'    Set xlUpRange = rng.Worksheet.Range(rng, rng.Worksheet.Cells(rng.Worksheet.rows.Count, rng.Column).End(xlUp))
'End Function
'
'Public Function UnionRanges(ByVal arr As Variant) As Range
'    Dim rng As Variant, uRng As Range
'    For Each rng In arr
'        If uRng Is Nothing Then
'            Set uRng = rng
'        Else
'            Set uRng = Union(uRng, rng)
'        End If
'    Next rng
'    Set UnionRanges = uRng
'End Function
