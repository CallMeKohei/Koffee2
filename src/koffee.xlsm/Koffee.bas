Attribute VB_Name = "Koffee"
''' --------------------------------------------------------
'''  FILE    : Koffee.bas
'''  AUTHOR  : callmekohei <callmekohei at gmail.com>
'''  License : MIT license
''' --------------------------------------------------------
Option Explicit
Option Private Module


''' Dependencies
'''
'''     Ariawase: https://github.com/vbaidiot/ariawase
'''
'''                   | ArrRank | ArrLen | Arr2DToJagArr |
'''     --------------+-------- +------- +---------------|
'''     IsJagArr      |    *    |   *    |               |
'''     ArrTranspose  |         |   *    |               |
'''     GetVal        |         |        |       *       |
'''     PutVal        |         |   *    |       *       |


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
'''                     Cells Operation
''' --------------------------------------------------------

''' @param rng As Rang
''' @return As Variant(Of Array(Of Array(Of T)))
Public Function GetVal(ByVal rng As Range) As Variant

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

    If IsArray(arr) Then
        GetVal = Arr2DToJagArr(arr)
    Else
        GetVal = Array(Array(arr))
    End If

End Function

''' @param arr2D As Variant(Of Array(Of T, T) Or Of Array(Of Array(Of T)) Or T)
''' @param rng As Rang
''' @return As Variant(Of Array(Of Array(Of T)))
Public Sub PutVal(ByVal arr2D As Variant, ByVal rng As Range, Optional isVertical As Boolean = False)

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
    If Not IsArray(arr2D) Then
        If IsObject(arr2D) Then Err.Raise 13
        Dim tmp2DArr(0, 0) As Variant: tmp2DArr(0, 0) = arr2D
        arr2D = tmp2DArr
    End If

    If ArrRank(arr2D) >= 3 Then Err.Raise 13

    ''' 1D array to 2D array
    If ArrRank(arr2D) = 1 Then
        If IsJagArr(arr2D) Then
            arr2D = JagArrToArr2D(arr2D)
        Else
            arr2D = JagArrToArr2D(Array(arr2D))
        End If
    End If

    If Not ArrRank(arr2D) = 2 Then Err.Raise 13  ''' Type mismatch

    If isVertical Then
        ''' Minimum index Excel's Array is 1
        If LBound(arr2D, 1) = 1 Then
            rng.Resize(UBound(arr2D, 2), UBound(arr2D, 1)).Value = ArrTranspose(arr2D)
        Else
            rng.Resize(UBound(arr2D, 2) + 1, UBound(arr2D, 1) + 1).Value = ArrTranspose(arr2D)
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

''' @param rng As Rang
''' @return As Long
Public Function LastRow(ByVal rng As Range, Optional toDown As Boolean = False) As Long

    ''' ( Usage )

    '''     |  A   B   C   D   E   F   G
    ''' ----+---------------------------
    '''   1 |
    '''   2 |      1
    '''   3 |      2
    '''   4 |      3
    '''   5 |
    '''   6 |      A
    '''   7 |      B

    ''' Dim ws As Worksheet: Set ws = AddSheet("abc")
    ''' Debug.Print LastRow( ws.Range("B2") )
    ''' 7
    ''' Debug.Print LastRow( ws.Range("B2") , True)
    ''' 4

    If toDown Then
        LastRow = rng.End(xlDown).Row
    Else
        LastRow = rng.Worksheet.Cells(rng.Worksheet.Rows.Count, rng.Column).End(xlUp).Row
    End If

End Function

''' @param rng As Rang
''' @return As Long
Public Function LastCol(ByVal rng As Range, Optional toRight As Boolean = False) As Long

    ''' ( Usage )

    '''    |  A   B   C   D   E   F   G
    '''----+---------------------------
    '''  1 |
    '''  2 |      1   2   3       A   B
    '''  3 |
    '''  4 |
    '''  5 |
    '''  6 |
    '''  7 |

    ''' Dim ws As Worksheet: Set ws = AddSheet("abc")
    ''' Debug.Print LastCol( ws.Range("B2") )
    ''' 7
    ''' Debug.Print LastCol( ws.Range("B2") , True)
    ''' 4

    If toRight Then
        LastCol = rng.End(xlToRight).Column
    Else
        LastCol = rng.Worksheet.Cells(rng.Row, rng.Worksheet.columns.Count).End(xlToLeft).Column
    End If

End Function

''' @param ws As Worksheet
Public Sub Hankaku(ByVal ws As Worksheet)

    ''' ( Usage )
    ''' Dim ws As Worksheet: Set ws = AddSheet("abc")
    ''' Hankaku ws
    ''' "ƒAƒCƒEƒGƒI" -> "±²³´µ"
    ''' "‚`‚a‚b‚c‚d" -> "ABCDE"
    ''' "‚P‚Q‚R‚S‚T" -> 12345

    Dim v As Range
    For Each v In ws.UsedRange
        v.Value = StrConv(v.Value, vbNarrow)
    Next

End Sub
