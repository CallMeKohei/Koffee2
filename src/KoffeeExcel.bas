Attribute VB_Name = "KoffeeExcel"
'  +--------------                                         --------------+
'  |||||||||    Koffee2 0.1.0                                            |
'  |: ^_^ :|    Koffee2 is free Library based on Ariawase.               |
'  |||||||||    The Project Page: https://github.com/CallMeKohei/Koffee2 |
'  +--------------                                         --------------+
Option Explicit

Public Enum GetValOption
    AllHol = 0
    AllVer = 1
    Holizontal = 2
    Vertical = 3
End Enum

Public Function LastRow(ByVal R As Range, Optional toDonw As Boolean = False) As Long
    Select Case toDonw
        Case True:  LastRow = R.End(xlDown).row
        Case False: LastRow = R.Worksheet.Cells(R.Worksheet.Rows.Count, R.Column).End(xlUp).row
    End Select
End Function

Public Function LastCol(ByVal R As Range, Optional toRight As Boolean = False) As Long
    Select Case toRight
        Case True:  LastCol = R.End(xlToRight).Column
        Case False: LastCol = R.Worksheet.Cells(R.row, R.Worksheet.Columns.Count).End(xlToLeft).Column
    End Select
End Function

Public Function GetVal(ByVal R As Range, ByVal GetValOption As GetValOption, Optional ByVal direction As Boolean = False) As Variant

    Dim v As Variant, arr As Variant, arrx As New ArrayEx
    Dim lRow As Long: lRow = (LastRow(R, direction) - R.row)
    Dim lCol As Long: lCol = (LastCol(R, direction) - R.Column)
    
    Select Case GetValOption
        Case Is = 0
            arr = R.Resize(lRow + 1, lCol + 1)
            arr = Arr2DToJagArr(arr)
            For v = 1 To UBound(arr, 1): arrx.addval arr(v): Next v
        Case Is = 1
            For v = 0 To lCol
                arrx.addval ArrFlatten(Arr2DToJagArr(R.Offset(0, v).Resize(lRow + 1, 1).Value))
            Next v
        Case Is = 2: For v = 0 To lCol: arrx.addval R.Cells(1, v + 1).Value: Next v
        Case Is = 3: For v = 0 To lRow: arrx.addval R.Cells(v + 1, 1).Value: Next v
    End Select
    
    GetVal = arrx.ToArray
    
End Function

Public Sub PutVal(ByVal arr As Variant, ByVal R As Range, Optional isVertical As Boolean = False)
    Dim wf As WorksheetFunction: Set wf = Application.WorksheetFunction
    
    If Not IsArray(arr) Then arr = Array(arr)
    If IsJagArr(arr) Then arr = JagArrToArr2D(arr)
    
    Select Case ArrRank(arr)
        Case Is > 2: Err.Raise 13
        Case Is = 2 And LBound(arr, 1) = 1 ' Array based '1' from Worksheets
            If isVertical = True Then
                R.Resize(UBound(arr, 2), UBound(arr, 1)).Value = wf.Transpose(arr)
            Else
                R.Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
            End If
        Case Is = 1
            If isVertical = True Then
                R.Resize(UBound(arr) + 1, LBound(arr) + 1).Value = wf.Transpose(arr)
            Else
                R.Resize(LBound(arr) + 1, UBound(arr) + 1).Value = arr
            End If
        Case Is = 2
            If isVertical = True Then
                R.Resize(UBound(arr, 2) + 1, UBound(arr, 1) + 1).Value = wf.Transpose(arr)
            Else
                R.Resize(UBound(arr, 1) + 1, UBound(arr, 2) + 1).Value = arr
            End If
        Case Else
            Err.Raise 13
    End Select
End Sub

Public Sub Hankaku(ByVal sh As Worksheet)
    Dim v As Variant
    For Each v In sh.UsedRange
      v.Value = StrConv(v.Value, vbNarrow)
    Next
End Sub

Public Function ArrSheetsName(Optional ByVal bk As Workbook = Nothing) As Variant
    
    If TypeName(bk) = "Nothing" Then Set bk = Application.ThisWorkbook
    
    Dim sh As Worksheet, arrx As New ArrayEx
    For Each sh In bk.Worksheets
        arrx.addval ToStr(sh.Name)
    Next sh
    
    ArrSheetsName = arrx.ToArray
    
End Function

Public Function ExistsSheet(ByVal SheetName As String, Optional ByVal bk As Workbook = Nothing) As Boolean
    If TypeName(bk) = "Nothing" Then Set bk = Application.ThisWorkbook
    
    Select Case ArrIndexOf(ArrSheetsName(bk), SheetName)
        Case -1:   ExistsSheet = False
        Case Else: ExistsSheet = True
    End Select

End Function

Public Function AddSheet(ByVal SheetName As String, Optional ByVal bk As Workbook = Nothing) As Worksheet
    
    If TypeName(bk) = "Nothing" Then Set bk = Application.ThisWorkbook
    If ExistsSheet(SheetName, bk) Then Exit Function
    
    With bk.Worksheets.Add()
        .Name = SheetName
    End With
    
    Set AddSheet = bk.Worksheets(SheetName)
    
End Function

Public Function CopySheet(ByVal SourceSheetName As String, ByVal SheetName As String, Optional ByVal bk As Workbook = Nothing) As Worksheet

    If TypeName(bk) = "Nothing" Then Set bk = Application.ThisWorkbook
    If ExistsSheet(SheetName, bk) Then Exit Function
    
    With bk.Worksheets(SourceSheetName)
        .Copy after:=bk.Worksheets(SourceSheetName)
    End With
    
    With bk.ActiveSheet
        .Name = SheetName
    End With
    
    Set CopySheet = bk.ActiveSheet
    
End Function

Public Sub ProtectSheet(ByVal sh As Worksheet, Optional myPassword As String = "1234")
    sh.Protect _
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

Public Sub ExcelStatus( _
    Optional ByVal aScreenUpDating As Boolean = True, _
    Optional ByVal aCalculation As XlCalculation = xlCalculationAutomatic, _
    Optional ByVal aEnableEvents As Boolean = True, _
    Optional ByVal aDisplayAlerts As Boolean = True, _
    Optional ByVal aStatusBar = False, _
    Optional ByVal aDisplayStatusBar = True)
                    
    With Application
      .ScreenUpdating = aScreenUpDating
      .Calculation = aCalculation
      .EnableEvents = aEnableEvents
      .DisplayAlerts = aDisplayAlerts
      .statusBar = aStatusBar
      .DisplayStatusBar = aDisplayStatusBar
    End With
End Sub

