Attribute VB_Name = "Koffee"
''' --------------------------------------------------------
'''  FILE    : Koffee.bas
'''  AUTHOR  : callmekohei <callmekohei at gmail.com>
'''  License : MIT license
''' --------------------------------------------------------
Option Explicit
Option Private Module


''' Dependencies
'''     Ariawase: https://github.com/vbaidiot/ariawase



''' --------------------------------------------------------
'''                      Util Functions
''' --------------------------------------------------------

''' @param arr As Variant(Of Array(Of Array(Of T)))
''' @return As Boolean
Public Function IsJagArr(ByVal arr As Variant) As Boolean

    ''' @seealso Jagged Arrays
    ''' https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/arrays/jagged-arrays

    On Error GoTo Escape

    ''' check outer array
    If Not IsArray(arr) Then GoTo Escape
    If Not ArrRank(arr) = 1 Then GoTo Escape
    If ArrLen(arr) = 0 Then GoTo Escape

    '' check inner array
    Dim innerArray As Variant
    For Each innerArray In arr
        If Not IsArray(innerArray) Then GoTo Escape
        If IsObject(innerArray) Then GoTo Escape
    Next innerArray

    IsJagArr = True

Escape:
End Function

''' @param arr2D As Variant(Of Array(Of T, T))
''' @return As Variant(Of Array(Of T, T))
Public Function ArrTranspose(ByVal arr2D As Variant) As Variant

    If Not IsArray(arr2D) Then Err.Raise 13
    If Not ArrRank(arr2D) = 2 Then Err.Raise 13

    Dim lb1 As Long: lb1 = LBound(arr2D, 2)
    Dim ub1 As Long: ub1 = UBound(arr2D, 2)
    Dim lb2 As Long: lb2 = LBound(arr2D, 1)
    Dim ub2 As Long: ub2 = UBound(arr2D, 1)

    Dim tmpArr2D() As Variant
    ReDim tmpArr2D(lb1 To ub1, lb2 To ub2)

    Dim ix1 As Long, ix2 As Long
    For ix1 = lb1 To ub1
        For ix2 = lb2 To ub2
            If IsObject(arr2D(ix2, ix1)) Then
                Set tmpArr2D(ix1, ix2) = arr2D(ix2, ix1)
            Else
                Let tmpArr2D(ix1, ix2) = arr2D(ix2, ix1)
            End If
        Next ix2
    Next ix1

    ArrTranspose = tmpArr2D

End Function

Public Function Arr1Dto2D(ByVal arr As Variant) As Variant

    If IsJagArr(arr) Then
        Arr1Dto2D = JagArrToArr2D(arr)
        GoTo Escape
    End If

    If Not LBound(arr) = 1 Then
        Arr1Dto2D = JagArrToArr2D(Array(arr))
        GoTo Escape
    End If

    ''' base1 array is regarded as an array made from Excel worksheet's ranges.
    Dim lb As Long: lb = LBound(arr)
    Dim ub As Long: ub = UBound(arr)

    Dim tmp() As Variant
    ReDim tmp(1 To 1, lb To ub)

    Dim i As Long
    For i = 1 To UBound(arr)
        tmp(1, i) = arr(i)
    Next i

    Arr1Dto2D = tmp

Escape:
End Function

''' @param arr As Variant(Of Array(Of T)
''' @return As Variant(Of Array(Of T))
Public Function ArrPadLeft(ByVal arr As Variant) As Variant

    ''' Array("foo",Empty,"bar",Empty)
    ''' Array("foo","foo","bar","bar")

    Dim v As Variant, tmp As String, arrx As ArrayEx: Set arrx = New ArrayEx
    For Each v In arr
        If Not IsEmpty(v) Then tmp = v
        If IsObject(tmp) Then
            arrx.AddObj (tmp)
        Else
            arrx.AddVal (tmp)
        End If
    Next v
    ArrPadLeft = arrx.ToArray
    Set arrx = Nothing
End Function

''' @param arr As Variant(Of Array(Of T)
''' @return As Variant(Of Array(Of T))
Public Function ArrRemoveEmpty(ByVal arr As Variant) As Variant
    Dim v As Variant, tmp As String, arrx As ArrayEx: Set arrx = New ArrayEx

    ''' Array("foo",Empty,"bar",Empty)
    ''' Array("foo","bar")

    For Each v In arr
        If Not IsEmpty(v) Then
            If IsObject(v) Then
                arrx.AddObj v
            Else
                arrx.AddVal v
            End If
        End If
    Next v
    ArrRemoveEmpty = arrx.ToArray
    Set arrx = Nothing
End Function

''' @param dbType As dbTypeEnum
''' @param sql As String
''' @return As Variant(Of Array(Of T, T))
Public Function Select_(ByVal dbType As dbTypeEnum, ByVal sql As String, _
    Optional ByVal fpath As String = "", _
    Optional ByVal isTableHeader As Boolean = True) As Variant

    ''' This function is helper function for AdoEx class.

    Dim adox As AdoEx: Set adox = New AdoEx
    adox.Init dbType, fpath, isTableHeader
    Dim arr As Variant: arr = adox.Select_(sql)
    If IsEmpty(arr) Then GoTo Escape
    Select_ = Array(arr(0), arr(1))
    Set adox = Nothing

Escape:
End Function


''' --------------------------------------------------------
'''                    General Operation
''' --------------------------------------------------------

''' @seealso ScreenUpdating https://docs.microsoft.com/en-us/office/vba/api/excel.application.screenupdating (/ja-jp/office/vba/api/excel.application.statusbar)
''' @seealso Calculation    https://docs.microsoft.com/en-us/office/vba/api/excel.application.calculation    (/ja-jp/office/vba/api/excel.application.calculation)
''' @seealso EnableEvents   https://docs.microsoft.com/en-us/office/vba/api/excel.application.enableevents   (/ja-jp/office/vba/api/excel.application.enableevents)
''' @seealso DisplayAlerts  https://docs.microsoft.com/en-us/office/vba/api/excel.application.statusbar      (/ja-jp/office/vba/api/excel.application.statusbar)
''' @seealso StatusBar      https://docs.microsoft.com/en-us/office/vba/api/excel.application.displayalerts  (/ja-jp/office/vba/api/excel.application.displayalerts)
Public Sub ExcelStatus( _
    Optional ByVal aScreenUpDating As Boolean = True, _
    Optional ByVal aCalculation As XlCalculation = xlCalculationAutomatic, _
    Optional ByVal aEnableEvents As Boolean = True, _
    Optional ByVal aDisplayAlerts As Boolean = True, _
    Optional ByVal aStatusBar As Boolean = False, _
    Optional ByVal aDisplayStatusBar As Boolean = True)

    ''' ( Usage )
    '''
    ''' sub foo ()
    '''
    '''     ExcelStatus _
    '''         aScreenUpDating:=False, _
    '''         aCaluculation:=xlCalculationManual, _
    '''         aEnableEvents:=False, _
    '''         aDisplayAlerts:=False, _
    '''         aStatusBar:=True, _
    '''         aDisplayStatusBar:=False
    '''
    '''
    '''     // Do Somthig ...
    '''
    '''
    '''     ExcelStatus
    '''
    ''' End Sub

    Application.ScreenUpdating = aScreenUpDating
    Application.Calculation = aCalculation
    Application.EnableEvents = aEnableEvents
    Application.DisplayAlerts = aDisplayAlerts
    Application.statusBar = aStatusBar
    Application.DisplayStatusBar = aDisplayStatusBar

End Sub

''' @param ws As Worksheet
Public Sub ProtectSheet(ByVal ws As Worksheet, Optional myPassword As String = "1234")

    ws.Protect _
        Password:=myPassword, _
        DrawingObjects:=False, _
        Contents:=True, _
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
'''                    Workbook Operation
''' --------------------------------------------------------

''' @param excelApp As Excel.Application
''' @param filePath As String
''' @param isReadOnly As Boolean
''' @return As WorkBook
Public Function CreateWorkBook(ByVal filePath As String, Optional isReadOnly As Boolean = True) As Workbook
    Dim flg As Boolean: flg = IsWorkBookClosed(filePath)
    If Not (isReadOnly Or flg) Then Err.Raise 9999, , "WorkBook is already opened."
    If flg Then
            Set CreateWorkBook = Workbooks.Open(Filename:=filePath, UpdateLinks:=0, ReadOnly:=isReadOnly, IgnoreReadOnlyRecommended:=True)
    Else
        ''' create anthor Excel application process
        Dim excelApp As Excel.Application: Set excelApp = New Excel.Application
        Set CreateWorkBook = excelApp.Workbooks.Open(Filename:=filePath, UpdateLinks:=0, ReadOnly:=isReadOnly, IgnoreReadOnlyRecommended:=True)
    End If
End Function

''' @param filePath As String
''' @return As Boolean
Private Function IsWorkBookClosed(ByVal filePath As String) As Boolean
    On Error GoTo Escape
        Open filePath For Append As #1
        Close #1
    On Error GoTo 0
        IsWorkBookClosed = True
Escape:
End Function

''' @param excelApp As Excel.Application
''' @param wb As Workbook
Public Sub SaveCloseWorkBook(ByVal wb As Workbook)
    wb.Save
    wb.Close   ''' This method contains wb.Application.Quit
End Sub

''' @param excelApp As Excel.Application
''' @param wb As Workbook
Public Sub CloseWorkBook(ByVal wb As Workbook)
    wb.Close   ''' This method contains wb.Application.Quit
End Sub


''' --------------------------------------------------------
'''                   WorkSheet Operation
''' --------------------------------------------------------

''' @return As Variant(Of Array(Of String))
Public Function ArrSheetsName(Optional ByVal wb As Workbook = Nothing) As Variant

    ''' ( Usage )
    ''' Dim wb As Workbook: Set wb = Application.ThisWorkbook
    ''' Debug.Print ArrSheetsName(wb)(0)

    If TypeName(wb) = "Nothing" Then Set wb = Application.ThisWorkbook

    Dim arr() As String
    ReDim arr(0 To wb.Sheets.Count - 1)

    Dim ws As Worksheet, i As Long
    For Each ws In wb.Worksheets
        arr(i) = ws.Name
        i = i + 1
    Next ws

    ArrSheetsName = arr

End Function

''' @param SheetName As String
''' @return As Boolean
Public Function ExistsSheet(ByVal SheetName As String, Optional ByVal wb As Workbook = Nothing) As Boolean

    ''' ( Usage )
    ''' Debug.Print ExistsSheet("abc")

    If TypeName(wb) = "Nothing" Then Set wb = Application.ThisWorkbook

    Dim v As Variant
    For Each v In ArrSheetsName(wb)
        If SheetName = v Then
            ExistsSheet = True
            GoTo Escape
        End If
    Next v

Escape:
End Function

''' @param SheetName As String
Public Sub DeleteSheet(ByVal SheetName As String, Optional ByVal wb As Workbook = Nothing)

    ''' ( Usage )
    ''' DeleteSheet "abc"

    If TypeName(wb) = "Nothing" Then Set wb = Application.ThisWorkbook

    If Not ExistsSheet(SheetName, wb) Then GoTo Catch

    Application.DisplayAlerts = False
    wb.Worksheets(SheetName).Delete
    Application.DisplayAlerts = True
    GoTo Escape

Catch:
    Debug.Print "(DeleteSheet): The SheetName is not exists!"
    Exit Sub

Escape:
End Sub

''' @param SheetName As String
''' @return As Worksheet
Public Function AddSheet(ByVal SheetName As String, Optional ByVal wb As Workbook = Nothing) As Worksheet

    ''' ( Usage )
    ''' Dim ws As Worksheet: Set ws = AddSheet("abc")

    If TypeName(wb) = "Nothing" Then Set wb = Application.ThisWorkbook

    If ExistsSheet(SheetName, wb) Then GoTo Catch

    wb.Worksheets.Add().Name = SheetName
    Set AddSheet = wb.Worksheets(SheetName)
    GoTo Escape

Catch:
    Debug.Print "(AddSheet): The SheetName is already exists!"
    Exit Function

Escape:
End Function

''' @param SourceSheetName As String
''' @param SheetName As String
''' @return As Worksheet
Public Function CopySheet(ByVal SourceSheetName As String, ByVal SheetName As String, Optional ByVal wb As Workbook = Nothing) As Worksheet

    ''' ( Usage )
    ''' Dim wsCopied As Worksheet: Set wsCopied = CopySheet("abc", "abcCopied")

    If TypeName(wb) = "Nothing" Then Set wb = Application.ThisWorkbook

    If Not ExistsSheet(SourceSheetName, wb) Then GoTo Catch
    If ExistsSheet(SheetName, wb) Then GoTo Catch2

    wb.Worksheets(SourceSheetName).Copy after:=wb.Worksheets(SourceSheetName)
    wb.ActiveSheet.Name = SheetName
    Set CopySheet = wb.Worksheets(SheetName)
    GoTo Escape

Catch:
    Debug.Print "(CopySheet): The SrouceSheet is not exists!"
    Exit Function

Catch2:
    Debug.Print "(CopySheet): The SheetName is already exists!"
    Exit Function

Escape:
End Function


''' --------------------------------------------------------
'''                          Ranges
''' --------------------------------------------------------

''' @param rng As Rang
''' @param isVertical As Boolean
''' @return As Variant(Of Array(Of Array(Of T)))
Public Function GetVal(ByVal rng As Range, Optional isVertical As Boolean = False) As Variant

    ''' ( Usage ) Dump() is Ariawase's function

    '''     |  A   B   C   D
    ''' ----+------------------
    '''   1 |
    '''   2 |      1   2
    '''   3 |      3   4
    '''   4 |

    ''' Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("abc")
    ''' Debug.Print Dump( GetVal( ws.Range("B2:C4") ) )
    ''' Array( Array(1,2), Array(3,4) )

    Dim arr As Variant: arr = rng.Value
    If Not IsArray(arr) Then
        Dim tmp(1 To 1, 1 To 1) As Variant: tmp(1, 1) = arr: arr = tmp
    End If
    If isVertical Then
        GetVal = Arr2DToJagArr(ArrTranspose(arr))
    Else
        GetVal = Arr2DToJagArr(arr)
    End If

End Function

''' @param rng As Rang
''' @param ptrnFind As String
''' @param isVertical As Boolean
''' @return As Variant(Of Array(Of Array(Of Range)))
Public Function RegexRanges(ByVal rng As Range _
    , ByVal ptrnFind As String _
    , Optional ByVal isVertical As Boolean = True _
    ) As Variant

    Dim ws As Worksheet: Set ws = rng.Worksheet
    Dim inArr As Variant, arrx As ArrayEx: Set arrx = New ArrayEx
    For Each inArr In GetVal(rng, isVertical)
        Dim i As Long
        For i = 1 To ArrLen(inArr)
            If ArrLen(ReMatch(inArr(i), ptrnFind)) > 0 Then arrx.AddObj ws.Cells(rng.offset(i - 1).Row, rng.Column)
        Next i
    Next inArr

    RegexRanges = arrx.ToArray

End Function

''' @param arr As Variant(Of Array(Of T, T) Or Of Array(Of Array(Of T)) Or T)
''' @param rng As Rang
''' @param isVertical As Boolean
Public Sub PutVal(ByVal arr As Variant, ByVal rng As Range, Optional isVertical As Boolean = False)

    ''' ( Usage )

    ''' Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("abc")

    ''' Hrizontal put
    ''' PutVal Array( Array(A,B), Array(1,2) ), ws.Range("B2")

    '''     |  A   B   C   D
    ''' ----+------------------
    '''   1 |
    '''   2 |      A   B
    '''   3 |      1   2
    '''   4 |

    ''' Vertical put
    ''' PutVal Array( Array(A,B), Array(1,2) ), ws.Range("B2") , True

    '''     |  A   B   C   D
    ''' ----+------------------
    '''   1 |
    '''   2 |      A   1
    '''   3 |      B   2
    '''   4 |

    ''' Value to 2D array
    If IsObject(arr) Then Err.Raise 13
    If Not IsArray(arr) Then
        Dim tmp(1 To 1, 1 To 1) As Variant: tmp(1, 1) = arr
        arr = tmp
    End If

    ''' 1D array to 2D array
    If ArrRank(arr) >= 3 Then Err.Raise 13
    If ArrRank(arr) = 1 Then
        arr = Arr1Dto2D(arr)
    End If

    ''' Exclude non-2-dimensional arrays
    If Not ArrRank(arr) = 2 Then Err.Raise 13

    If isVertical Then
        ''' Minimum index Excel's Array is 1
        If LBound(arr, 1) = 1 Then
            rng.Resize(UBound(arr, 2), UBound(arr, 1)).Value = ArrTranspose(arr)
        Else
            rng.Resize(UBound(arr, 2) + 1, UBound(arr, 1) + 1).Value = ArrTranspose(arr)
        End If
    Else
        ''' Minimum index Excel's Array is 1
        If LBound(arr, 1) = 1 Then
            rng.Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
        Else
            rng.Resize(UBound(arr, 1) + 1, UBound(arr, 2) + 1).Value = arr
        End If
    End If

End Sub

''' @param rng As Rang
''' @return As Long
Public Function LastRow(ByVal rng As Range, Optional toDown As Boolean = False) As Long
    If toDown Then
        LastRow = rng.End(xlDown).Row
    Else
        LastRow = rng.Worksheet.Cells(rng.Worksheet.Rows.Count, rng.Column).End(xlUp).Row
    End If
End Function

''' @param rng As Rang
''' @return As Long
Public Function LastCol(ByVal rng As Range, Optional toRight As Boolean = False) As Long
    If toRight Then
        LastCol = rng.End(xlToRight).Column
    Else
        LastCol = rng.Worksheet.Cells(rng.Row, rng.Worksheet.columns.Count).End(xlToLeft).Column
    End If
End Function

''' @param ws As Worksheet
Public Sub Hankaku(ByVal ws As Worksheet)
    Dim v As Range
    For Each v In ws.UsedRange
        v.Value = StrConv(v.Value, vbNarrow)
    Next
End Sub

''' @param rng As Rang
''' @param ptrnFind As String
''' @param times As Long
''' @param offsetRow As Long
''' @param offsetColumn As Long
Public Sub InsertRows(ByVal rng As Range, ByVal ptrnFind As String _
    , Optional ByVal times As Long = 1, Optional ByVal offsetRow As Long = 0, Optional ByVal offsetColumn As Long = 0)
    Dim i As Long
    For i = 1 To times
        If offsetRow = 0 And offsetColumn = 0 Then
            UnionRanges(RegexRanges(rng, ptrnFind)).EntireRow.Insert
        Else
            UnionRanges(offsetRanges(RegexRanges(rng, ptrnFind), offsetRow, offsetColumn)).EntireRow.Insert
        End If
    Next i
End Sub

''' @param rng As Rang
''' @param ptrnFind As String
''' @param times As Long
''' @param offsetRow As Long
''' @param offsetColumn As Long
Public Sub DeleteRows(ByVal rng As Variant, ByVal ptrnFind As String _
    , Optional ByVal times As Long = 1, Optional ByVal offsetRow As Long = 0, Optional ByVal offsetColumn As Long = 0)
    Dim i As Long
    For i = 1 To times
        If offsetRow = 0 And offsetColumn = 0 Then
            UnionRanges(RegexRanges(rng, ptrnFind)).EntireRow.Delete
        Else
            UnionRanges(offsetRanges(RegexRanges(rng, ptrnFind), offsetRow, offsetColumn)).EntireRow.Delete
        End If
    Next i
End Sub

''' @param arr As Variant(Of Array(Of Array(Of Range))
''' @param ptrnFind As String
''' @param times As Long
''' @param offsetRow As Long
''' @param offsetColumn As Long
''' @return Variant(Of Array(Of Range))
Public Function offsetRanges(ByVal arr As Variant, Optional ByVal offsetRow As Long = 0, Optional ByVal offsetColumn As Long = 0) As Variant
    Dim rng As Variant, arrx As ArrayEx: Set arrx = New ArrayEx
    For Each rng In arr
        arrx.AddObj rng.offset(offsetRow, offsetColumn)
    Next rng
    offsetRanges = arrx.ToArray
End Function

''' @param rng As Range
''' @return Range
Public Function xlUpRange(ByVal rng As Range) As Range
    Set xlUpRange = rng.Worksheet.Range(rng, rng.Worksheet.Cells(rng.Worksheet.Rows.Count, rng.Column).End(xlUp))
End Function

''' @param arr as Vaiant(Of Array(Of Range)
''' @return Range
Public Function UnionRanges(ByVal arr As Variant) As Range
    Dim rng As Variant, uRng As Range
    For Each rng In arr
        If uRng Is Nothing Then
            Set uRng = rng
        Else
            Set uRng = Union(uRng, rng)
        End If
    Next rng
    Set UnionRanges = uRng
End Function


