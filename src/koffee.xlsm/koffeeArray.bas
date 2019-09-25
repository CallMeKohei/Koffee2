Attribute VB_Name = "koffeeArray"
''' koffeeArray.bas
''' written by callmekohei(twitter at callmekohei)
''' MIT license
Option Explicit
Option Compare Text
Option Private Module
Option Base 0

Public Function ArrayLen(ByRef aSourceArray As Variant, Optional ByVal aDimmension As Long = 1) As Long
    If Not IsArray(aSourceArray) Then Err.Raise 13
    On Error Resume Next 'if empty is 0
    ArrayLen = (UBound(aSourceArray, aDimmension) - LBound(aSourceArray, aDimmension)) + 1
End Function

Public Function ArrayRank(ByRef aSourceArray As Variant) As Long
    If Not IsArray(aSourceArray) Then Err.Raise 13
    On Error GoTo Ending
    Dim boundIndex As Long: boundIndex = 1
    Dim tmp As Long
    Do
        tmp = UBound(aSourceArray, boundIndex)
        boundIndex = boundIndex + 1
    Loop
Ending:
    ArrayRank = boundIndex - 1
End Function

Public Function ArraySlice(ByRef aSource1DArray As Variant, Optional ByVal fst As Variant, Optional ByVal lst As Variant) As Variant
    If IsMissing(fst) Then fst = LBound(aSource1DArray)
    If IsMissing(lst) Then lst = UBound(aSource1DArray)
    Dim arr
    sArraySlice aSource1DArray, arr, fst, lst
    ArraySlice = arr
End Function

Public Sub sArraySlice(ByRef aSource1DArray As Variant, ByRef aDestArray As Variant, _
    Optional ByVal fst As Long, Optional ByVal lst As Long)
    
    ''' guard
    If Not IsArray(aSource1DArray) Then Err.Raise 13
    If Not ArrayRank(aSource1DArray) = 1 Then Err.Raise 13
    Dim lb As Long: lb = LBound(aSource1DArray)
    Dim ub As Long: ub = UBound(aSource1DArray)
    If ub < lb Then GoTo Ending
    If IsMissing(fst) Then fst = lb
    If IsMissing(lst) Then lst = ub
    If Not (lb <= fst And lst <= ub) Then Err.Raise 5

    aDestArray = Array(): ReDim aDestArray(lst - fst)
    
    Dim i As Long
    For i = 0 To lst - fst
        If IsObject(aSource1DArray(fst + i)) Then
            Set aDestArray(i) = aSource1DArray(fst + i)
        Else
            Let aDestArray(i) = aSource1DArray(fst + i)
        End If
    Next i
    
Ending:
End Sub

''' jagged Arrays: https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/arrays/jagged-arrays
Public Function IsJaggedArray(ByRef aSourceArray As Variant) As Boolean

    On Error GoTo Escape

    ''' check outer array
    If Not IsArray(aSourceArray) Then GoTo Escape
    If Not ArrayRank(aSourceArray) = 1 Then GoTo Escape
    If ArrLen(aSourceArray) = 0 Then GoTo Escape

    '' check inner array
    Dim innerArray As Variant
    For Each innerArray In aSourceArray
        If Not IsArray(innerArray) Then GoTo Escape
        If IsObject(innerArray) Then GoTo Escape
    Next innerArray

    IsJaggedArray = True

Escape:
End Function

Public Function Array2DToJagArray(ByRef aSourceArray As Variant) As Variant
    Dim arr
    sArray2DToJagArray aSourceArray, arr
    Array2DToJagArray = arr
End Function

Public Sub sArray2DToJagArray(ByRef aSourceArray As Variant, ByRef aDestArray As Variant)
    
    If Not IsArray(aSourceArray) Then Err.Raise 13
    If Not ArrayRank(aSourceArray) = 2 Then Err.Raise 13
    If UBound(aSourceArray, 1) < LBound(aSourceArray, 1) Then GoTo Ending
    
    Dim lb1 As Long: lb1 = LBound(aSourceArray, 1)
    Dim ub1 As Long: ub1 = UBound(aSourceArray, 1)
    Dim lb2 As Long: lb2 = LBound(aSourceArray, 2)
    Dim ub2 As Long: ub2 = UBound(aSourceArray, 2)
    
    aDestArray = Array(): ReDim aDestArray(0 To ub1 - lb1)
    Dim tmp2d: tmp2d = Array(): ReDim tmp2d(0 To ub2 - lb2)
        
    Dim idx1 As Long, idx2 As Long
    For idx1 = lb1 To ub1
        For idx2 = lb2 To ub2
            If IsObject(aSourceArray(idx1, idx2)) Then
                Set tmp2d(idx2 - lb2) = aSourceArray(idx1, idx2)
            Else
                Let tmp2d(idx2 - lb2) = aSourceArray(idx1, idx2)
            End If
        Next idx2
        aDestArray(idx1 - lb1) = tmp2d
    Next idx1

Ending:
End Sub

Public Function Array2DToJagArrayLet(ByRef aSourceArray As Variant) As Variant
    Dim arr
    sArray2DToJagArrayLet aSourceArray, arr
    Array2DToJagArrayLet = arr
End Function

Public Sub sArray2DToJagArrayLet(ByRef aSourceArray As Variant, ByRef aDestArray As Variant)
   
    If Not IsArray(aSourceArray) Then Err.Raise 13
    If Not ArrayRank(aSourceArray) = 2 Then Err.Raise 13
    If UBound(aSourceArray, 1) < LBound(aSourceArray, 1) Then GoTo Ending
    
    Dim lb1 As Long: lb1 = LBound(aSourceArray, 1)
    Dim ub1 As Long: ub1 = UBound(aSourceArray, 1)
    Dim lb2 As Long: lb2 = LBound(aSourceArray, 2)
    Dim ub2 As Long: ub2 = UBound(aSourceArray, 2)
    
    aDestArray = Array(): ReDim aDestArray(0 To ub1 - lb1)
    Dim tmp2d: tmp2d = Array(): ReDim tmp2d(0 To ub2 - lb2)
        
    Dim idx1 As Long, idx2 As Long
    For idx1 = lb1 To ub1
        For idx2 = lb2 To ub2
            tmp2d(idx2 - lb2) = aSourceArray(idx1, idx2)
        Next idx2
        aDestArray(idx1 - lb1) = tmp2d
    Next idx1

Ending:
End Sub

Public Function JagArrayToArray2D(ByRef aSourceJagArray As Variant) As Variant
    Dim arr
    sJagArrayToArray2D aSourceJagArray, arr
    JagArrayToArray2D = arr
End Function

Public Sub sJagArrayToArray2D(ByRef aSourceJagArray As Variant, ByRef aDest2Darray As Variant)

    ''' aSourceJagArray should be square array.

    If Not IsArray(aSourceJagArray) Then Err.Raise 13
    If Not IsJaggedArray(aSourceJagArray) Then Err.Raise 13
    If Not IsEmpty(aDest2Darray) Then aDest2Darray = Empty
    
    Dim lb1 As Long: lb1 = LBound(aSourceJagArray)
    Dim ub1 As Long: ub1 = UBound(aSourceJagArray)
    Dim lb2 As Long: lb2 = LBound(aSourceJagArray(lb1))
    Dim ub2 As Long: ub2 = UBound(aSourceJagArray(lb1))

    ReDim aDest2Darray(lb1 To ub1, lb2 To ub2)

    Dim ix1 As Long, ix2 As Long
    For ix1 = lb1 To ub1
        For ix2 = lb2 To ub2
            If IsObject(aSourceJagArray(ix1)(ix2)) Then
                Set aDest2Darray(ix1, ix2) = aSourceJagArray(ix1)(ix2)
            Else
                Let aDest2Darray(ix1, ix2) = aSourceJagArray(ix1)(ix2)
            End If
        Next ix2
    Next ix1

End Sub

Public Sub JagArrayToArray2DLet(ByRef aSourceJagArray As Variant, ByRef aDest2Darray As Variant)

    ''' aSourceJagArray should be square array.

    If Not IsArray(aSourceJagArray) Then Err.Raise 13
    If Not IsJaggedArray(aSourceJagArray) Then Err.Raise 13
    If Not IsEmpty(aDest2Darray) Then aDest2Darray = Empty
    
    Dim lb1 As Long: lb1 = LBound(aSourceJagArray)
    Dim ub1 As Long: ub1 = UBound(aSourceJagArray)
    Dim lb2 As Long: lb2 = LBound(aSourceJagArray(lb1))
    Dim ub2 As Long: ub2 = UBound(aSourceJagArray(ub1))

    ReDim aDest2Darray(lb1 To ub1, lb2 To ub2)

    Dim ix1 As Long, ix2 As Long
    For ix1 = lb1 To ub1
        For ix2 = lb2 To ub2
            Let aDest2Darray(ix1, ix2) = aSourceJagArray(ix1)(ix2)
        Next ix2
    Next ix1

End Sub

Public Function ArrayTranspose(ByRef aSource2DArray As Variant) As Variant
    Dim arr()
    sArrayTranspose aSource2DArray, arr
    ArrayTranspose = arr
End Function

Public Sub sArrayTranspose(ByRef aSource2DArray As Variant, ByRef aDestArray As Variant)

    If Not IsArray(aSource2DArray) Then Err.Raise 13
    If Not ArrayRank(aSource2DArray) = 2 Then Err.Raise 13
'    If Not IsEmpty(aDestArray) Then aDestArray = Empty
    
    Dim lb1 As Long: lb1 = LBound(aSource2DArray, 2)
    Dim ub1 As Long: ub1 = UBound(aSource2DArray, 2)
    Dim lb2 As Long: lb2 = LBound(aSource2DArray, 1)
    Dim ub2 As Long: ub2 = UBound(aSource2DArray, 1)

    ReDim aDestArray(lb1 To ub1, lb2 To ub2)

    Dim ix1 As Long, ix2 As Long
    For ix1 = lb1 To ub1
        For ix2 = lb2 To ub2
            If IsObject(aSource2DArray(ix2, ix1)) Then
                Set aDestArray(ix1, ix2) = aSource2DArray(ix2, ix1)
            Else
                Let aDestArray(ix1, ix2) = aSource2DArray(ix2, ix1)
            End If
        Next ix2
    Next ix1

End Sub

Public Sub sArrayTransposeLet(ByRef aSource2DArray As Variant, ByRef aDestArray As Variant)

    If Not IsArray(aSource2DArray) Then Err.Raise 13
    If Not ArrayRank(aSource2DArray) = 2 Then Err.Raise 13
    If Not IsEmpty(aDestArray) Then aDestArray = Empty

    Dim lb1 As Long: lb1 = LBound(aSource2DArray, 2)
    Dim ub1 As Long: ub1 = UBound(aSource2DArray, 2)
    Dim lb2 As Long: lb2 = LBound(aSource2DArray, 1)
    Dim ub2 As Long: ub2 = UBound(aSource2DArray, 1)

    ReDim aDestArray(lb1 To ub1, lb2 To ub2)

    Dim ix1 As Long, ix2 As Long
    For ix1 = lb1 To ub1
        For ix2 = lb2 To ub2
            aDestArray(ix1, ix2) = aSource2DArray(ix2, ix1)
        Next ix2
    Next ix1

End Sub

Public Sub sArrayBase0(ByRef aSourceArray As Variant)
    ReDim Preserve aSourceArray(0 To UBound(aSourceArray) - 1)
End Sub

Public Function ArrayColumn(ByVal aColumnIndex As Long, ByRef aSourceArray As Variant) As Variant
    Dim i As Long, tmpArray: ReDim tmpArray(0 To UBound(aSourceArray))
    For i = 0 To UBound(aSourceArray)
        If IsObject(tmpArray(i)) Then
            Set tmpArray(i) = aSourceArray(i)(aColumnIndex)
        Else
            Let tmpArray(i) = aSourceArray(i)(aColumnIndex)
        End If
    Next i
    ArrayColumn = tmpArray
End Function

''' Array("15.0", "16.0", "16.0", "Common", "Outlook") => "\d\d\.\d" => Array("15.0", "16.0")
Public Function ArrayRegexFilter(ByVal arr As Variant, ByVal ptrn As String) As Variant

    Dim regx As Object: Set regx = CreateObject("VBScript.RegExp")
    regx.Pattern = ptrn: regx.ignorecase = True: regx.Global = True

    Dim v, dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    For Each v In arr
        If regx.test(v) Then
            If Not dict.exists(v) Then dict.Add v, ""
        End If
    Next v

    ArrayRegexFilter = dict.Keys

    Set dict = Nothing
    Set regx = Nothing

End Function

Public Function ArrayHasEmpties(ByRef aSource1DArray As Variant) As Boolean

    If Not ArrayRank(aSource1DArray) = 1 Then Err.Raise 13

    Dim i As Long
    For i = LBound(aSource1DArray) To UBound(aSource1DArray)
        If aSource1DArray(i) = "" Or IsEmpty(aSource1DArray(i)) Then
            ArrayHasEmpties = True
            GoTo Escape
        End If
    Next i

Escape:
End Function

''' Array("foo",Empty,"bar",Empty) => Array("foo","bar")
Public Function ArrayRemoveEmpties(ByVal aSourceArray As Variant) As Variant

    If ArrRank(aSourceArray) <> 1 Then
        Err.Raise 9999, , "SourceArray must be a single dimensional array."
    End If

    Dim i As Variant, arrx As ArrayEx: Set arrx = New ArrayEx

    For i = LBound(aSourceArray) To UBound(aSourceArray)
        If IsObject(i) Then
            arrx.AddObj aSourceArray(i)
        ElseIf IsArray(aSourceArray(i)) Then
            arrx.AddVal aSourceArray(i)
        ElseIf Not (aSourceArray(i) = "" Or IsEmpty(aSourceArray(i))) Then
            arrx.AddVal aSourceArray(i)
        End If
    Next i
    ArrayRemoveEmpties = arrx.ToArray
    Set arrx = Nothing
End Function

Public Function ArrayWindow(ByRef aSource1DArray As Variant, ByVal GroupNumber As Variant)
    Dim arr
    sArrayWindow aSource1DArray, GroupNumber, arr
    ArrayWindow = arr
End Function
               
Public Sub sArrayWindow(ByRef aSource1DArray As Variant, ByVal GroupNumber As Variant, ByRef aDestArray As Variant)

    ' Array(1..10) divided by 3
    ' -------------------------
    ' => Array(1%, 2%, 3%, 4%)
    ' => Array(5%, 6%, 7%)
    ' => Array(8%, 9%, 10%)
    
    ''' guard
    If Not IsArray(aSource1DArray) Then Err.Raise 13
    If Not ArrayRank(aSource1DArray) = 1 Then Err.Raise 13
    If LBound(aSource1DArray) < 0 Then Err.Raise 13
    
    ''' guard2( GroupNumber )
    Select Case GroupNumber
        Case Is <= 0: Err.Raise 13
        Case Is = 1:  aDestArray = Array(aSource1DArray): GoTo Ending
        Case Is >= (UBound(aSource1DArray) + 1)
            aDestArray = Array(): ReDim aDestArray(0 To UBound(aSource1DArray))
            Dim idx As Long
            For idx = 0 To UBound(aSource1DArray)
                aDestArray(idx) = Array(aSource1DArray(idx))
            Next idx
            GoTo Ending
        Case Else
            GoTo ArrayWindowImpl
    End Select
   
   
ArrayWindowImpl:

    Dim groupIndex As Long: groupIndex = Int(ArrLen(aSource1DArray) / GroupNumber)
    Dim rest As Long: rest = ArrLen(aSource1DArray) Mod GroupNumber
    
    ''' simple divison : e.g. 8 / 3 => array(2,2,2)
    Dim groupIndexArray: groupIndexArray = Array(): ReDim groupIndexArray(0 To GroupNumber - 1)
    Dim i As Long
    For i = 0 To GroupNumber - 1
        groupIndexArray(i) = groupIndex
    Next i
    
    ''' add weight 1 : e.g. 8 / 3 => array(3,3,2)
    If Not rest = 0 Then
        Dim j As Long
        For j = 0 To rest - 1
            groupIndexArray(j) = groupIndexArray(j) + 1
        Next j
    End If
    
    ''' slice array by group index
    aDestArray = Array(): ReDim aDestArray(0 To GroupNumber - 1)
    Dim K As Long, acc_idx As Long
    For K = 0 To UBound(groupIndexArray)
'        aDestArray(K) = ArraySlice(aSource1DArray, acc_idx, acc_idx + groupIndexArray(K) - 1)
        sArraySlice aSource1DArray, aDestArray(K), acc_idx, acc_idx + groupIndexArray(K) - 1
        acc_idx = acc_idx + groupIndexArray(K)
    Next K
    
Ending:
End Sub

Public Function SelectCsv(ByVal sql As String, Optional ByVal aFolder As String = "") As Variant
On Error GoTo Escape

    If aFolder = "" Then aFolder = ThisWorkbook.Path

    Dim adox As AdoEx: Set adox = New AdoEx
    adox.InitCSVHeader aFolder
    Dim arr As Variant: arr = adox.ArraySelect(sql)
    If IsEmpty(arr) Then GoTo Escape
    SelectCsv = Array(arr(0), arr(1))
    Set adox = Nothing

Escape:
    Set adox = Nothing
End Function

Public Function SelectCsvHeader(ByVal sql As String, Optional ByVal aFolder As String = "") As Variant
On Error GoTo Escape

    If aFolder = "" Then aFolder = ThisWorkbook.Path

    Dim adox As AdoEx: Set adox = New AdoEx
    adox.InitCSVHeader aFolder
    Dim arr As Variant: arr = adox.ArraySelect(sql)
    If IsEmpty(arr) Then GoTo Escape
    SelectCsvHeader = Array(arr(0), arr(1))
    Set adox = Nothing

Escape:
    Set adox = Nothing
End Function

Public Function SelectText(ByVal sql As String, Optional ByVal aFolder As String = "") As Variant
On Error GoTo Escape

    If aFolder = "" Then aFolder = ThisWorkbook.Path

    Dim adox As AdoEx: Set adox = New AdoEx
    adox.InitText aFolder
    Dim arr As Variant: arr = adox.ArraySelect(sql)
    If IsEmpty(arr) Then GoTo Escape
    SelectText = Array(arr(0), arr(1))


Escape:
    Set adox = Nothing
End Function

Public Function SelectTextHeader(ByVal sql As String, Optional ByVal aFolder As String = "") As Variant
On Error GoTo Escape

    If aFolder = "" Then aFolder = ThisWorkbook.Path

    Dim adox As AdoEx: Set adox = New AdoEx
    adox.InitTextHeader aFolder
    Dim arr As Variant: arr = adox.ArraySelect(sql)
    If IsEmpty(arr) Then GoTo Escape
    SelectTextHeader = Array(arr(0), arr(1))

Escape:
    Set adox = Nothing
End Function

Public Function SelectAccess(ByVal sql As String, ByVal aFilePath As String) As Variant
On Error GoTo Escape

    Dim adox As AdoEx: Set adox = New AdoEx
    adox.InitAccess aFilePath
    Dim arr As Variant: arr = adox.ArraySelect(sql)
    If IsEmpty(arr) Then GoTo Escape
    SelectAccess = Array(arr(0), arr(1))
    
Escape:
    Set adox = Nothing
End Function

Public Function SelectExcel(ByVal sql As String, Optional ByVal aFilePath As String = "") As Variant
On Error GoTo Escape

    If aFilePath = "" Then aFilePath = ThisWorkbook.Path & "\" & ThisWorkbook.name

    Dim adox As AdoEx: Set adox = New AdoEx
    adox.InitExcel aFilePath
    Dim arr As Variant: arr = adox.ArraySelect(sql)
    If IsEmpty(arr) Then GoTo Escape
    SelectExcel = Array(arr(0), arr(1))
    
Escape:
    Set adox = Nothing
End Function

Public Function SelectExcelHeader(ByVal sql As String, Optional ByVal aFilePath As String = "") As Variant
On Error GoTo Escape

    If aFilePath = "" Then aFilePath = ThisWorkbook.Path & "\" & ThisWorkbook.name

    Dim adox As AdoEx: Set adox = New AdoEx
    adox.InitExcelHeader aFilePath
    Dim arr As Variant: arr = adox.ArraySelect(sql)
    If IsEmpty(arr) Then GoTo Escape
    SelectExcelHeader = Array(arr(0), arr(1))
    
Escape:
    Set adox = Nothing
End Function


''' This function is helper function for AdoEx class.
'Public Function ArraySelect(ByVal dbType As dbTypeEnum, ByVal sql As String, _
'    Optional ByVal fPath As String = "", _
'    Optional ByVal isTableHeader As Boolean = True) As Variant
'
'    Dim adox As AdoEx: Set adox = New AdoEx
'    adox.Init dbType, fPath, isTableHeader
'    Dim arr As Variant: arr = adox.ArraySelect(sql)
'    If IsEmpty(arr) Then GoTo Escape
'    ArraySelect = Array(arr(0), arr(1))
'    Set adox = Nothing
'
'Escape:
'End Function

''' ----- temporary

'Public Function ArrayGroupV1(ByVal arr As Variant) As Variant
'
'    ''' value x 1
'    '''    array( array( keyA, valA ), array( keyB, valB ) )
'    ''' => array( array( keyA, keyB), array( array( valA ), array( valB ) ) )
'
'    ''' group by
'    Dim v
'    Dim dict As Dictionary: Set dict = New Dictionary
'    Dim arrx As ArrayEx: Set arrx = New ArrayEx
'    Dim i As Long
'    For i = 0 To UBound(arr)
'
'        If i = 0 Then
'            dict.Add arr(i)(0), ""
'            arrx.AddVal arr(i)(1)
'        ElseIf i = UBound(arr) Then
'            If arr(i)(0) = arr(i - 1)(0) Then
'                arrx.AddVal arr(i)(1)
'                dict.Item(arr(i - 1)(0)) = arrx.ToArray()
'            Else
'                dict.Item(arr(i - 1)(0)) = arrx.ToArray()
'                dict.Add arr(i)(0), arr(i)(1)
'            End If
'        ElseIf arr(i)(0) = arr(i - 1)(0) Then
'            arrx.AddVal arr(i)(1)
'        Else
'            dict.Item(arr(i - 1)(0)) = arrx.ToArray()
'            dict.Add arr(i)(0), ""
'            Set arrx = Nothing
'            Set arrx = New ArrayEx
'            arrx.AddVal arr(i)(1)
'        End If
'
'    Next i
'
'    ArrayGroupV1 = Array(dict.Keys, dict.Items)
'    Set arrx = Nothing
'    Set dict = Nothing
'End Function

''' ----- predicated

'Public Function ArrPadLeft(ByVal arr As Variant) As Variant
'
'    ''' Array("foo",Empty,"bar",Empty) => Array("foo","foo","bar","bar")
'    Dim v As Variant, tmp As String, arrx As ArrayEx: Set arrx = New ArrayEx
'    For Each v In arr
'        If Not IsEmpty(v) Then tmp = v
'        If IsObject(tmp) Then
'            arrx.AddObj (tmp)
'        Else
'            arrx.AddVal (tmp)
'        End If
'    Next v
'    ArrPadLeft = arrx.ToArray
'    Set arrx = Nothing
'End Function

'Public Function ArrayBase0_2ndDimension(ByRef aSpreadSheetArray As Variant)
'    ReDim Preserve aSpreadSheetArray(LBound(aSpreadSheetArray) To UBound(aSpreadSheetArray, 1), 0 To UBound(aSpreadSheetArray, 2) - 1)
'    aSpreadSheetArray = Core.Arr2DToJagArr(aSpreadSheetArray)
'    ReDim Preserve aSpreadSheetArray(0 To UBound(aSpreadSheetArray) - 1)
'End Function


''' Array("a", "b", "c", "d") => 1,2 => Array("b","c")
'Public Function ArraySlice(ByVal arr As Variant, ByVal fst As Long, Optional snd As Long = 0) As Variant
'    If snd = 0 Then
'        ArraySlice = Application.index(arr, 0, Array(fst + 1))
'    Else
'        If fst > snd Then Err.Raise 9999
'        Dim ary: ReDim ary(0 To snd - fst) As Long
'        Dim i As Long
'        For i = 0 To (snd - fst): ary(i) = i + (fst + 1): Next i
'
'        ArraySlice = Application.index(arr, 0, ary)
'    End If
'End Function
