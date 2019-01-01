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
'''     IsJagArr
'''         ArrRank(Ariawase)
'''     GetVal
'''         Arr2DToJagArr(Ariawase)
'''     PutVal
'''         Arr2DToJagArr(Ariawase)
'''         ArrRank(Ariawase)
'''         IsJagArr(Koffee.core)

''' --------------------------------------------------------
'''                      Core Functions
''' --------------------------------------------------------

Public Function IsJagArr(ByVal arr As Variant) As Boolean

    If Not IsArray(arr) Then GoTo Escape
    On Error GoTo Escape

    If ArrRank(arr) > 1 Then GoTo Escape

    Dim v1 As Variant, v2 As Variant

    For Each v1 In arr
        If Not IsObject(v1) Then
            For Each v2 In v1
                If Not IsObject(v2) Then
                    IsJagArr = True
                    GoTo Escape
                End If
            Next v2
        End If
    Next v1

Escape:
End Function

Public Function ArrTranspose(ByVal arr As Variant) As Variant

    Dim ub1 As Long: ub1 = UBound(arr, 2)
    Dim ub2 As Long: ub2 = UBound(arr, 1)

    Dim tmp() As Variant: ReDim tmp(ub1, ub2)

    Dim ix1 As Long, ix2 As Long
    For ix1 = 0 To ub1
        For ix2 = 0 To ub2
            tmp(ix1, ix2) = arr(ix2, ix1)
        Next ix2
    Next ix1

    ArrTranspose = tmp

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


    Dim arr As Variant: arr = rng

    If IsArray(arr) Then
        GetVal = Arr2DToJagArr(arr)
    Else
        GetVal = Array(arr)
    End If

End Function

Public Sub PutVal(ByVal arr As Variant, ByVal rng As Range, Optional isVertical As Boolean = False)

    ''' ( Usage )

    ''' Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("abc")
    ''' PutVal Array( Array(1,2), Array(3,4) ), ws.Range("B2")

    '''     |  A   B   C   D
    ''' ----+------------------
    '''   1 |
    '''   2 |      1   2
    '''   3 |      3   4
    '''   4 |

    ''' PutVal Array( Array(1,2), Array(3,4) ), ws.Range("B2") , True

    '''     |  A   B   C   D
    ''' ----+------------------
    '''   1 |
    '''   2 |      1   3
    '''   3 |      2   4
    '''   4 |


    ''' 2D array from value : "foo" ---> array(array("foo")) ---> 2D array
    If Not IsArray(arr) Then arr = JagArrToArr2D(Array(Array(arr)))

    ''' 2D array from 1D array : array(1,2) ---> array(array(1,2)) ---> 2D array
    If Not IsJagArr(arr) Then
        If ArrRank(arr) = 1 Then
            arr = Array(arr)
        End If
    End If

    ''' 2D array from Jag array : array( array(1,2), array(3,4) ) ---> 2D array
    If IsJagArr(arr) Then
        arr = JagArrToArr2D(arr)
    End If

    If ArrRank(arr) <> 2 Then Err.Raise 13  ''' Type mismatch

    If isVertical Then

        Dim wf As WorksheetFunction: Set wf = Application.WorksheetFunction

        ''' Minimum index Excel's Array is 1
        If LBound(arr, 1) = 1 Then
            rng.Resize(UBound(arr, 2), UBound(arr, 1)).Value = wf.Transpose(arr)
        Else
            rng.Resize(UBound(arr, 2) + 1, UBound(arr, 1) + 1).Value = wf.Transpose(arr)
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
        LastRow = rng.End(xlDown).row
    Else
        LastRow = rng.Worksheet.Cells(rng.Worksheet.Rows.Count, rng.Column).End(xlUp).row
    End If

End Function

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
        LastCol = rng.Worksheet.Cells(rng.row, rng.Worksheet.columns.Count).End(xlToLeft).Column
    End If

End Function

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
