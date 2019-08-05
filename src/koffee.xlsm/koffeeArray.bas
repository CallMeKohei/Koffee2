Attribute VB_Name = "koffeeArray"
''' koffeeArray.bas
''' written by callmekohei(twitter at callmekohei)
''' MIT license
Option Explicit
Option Compare Text
Option Private Module
Option Base 0

''' jagged Arrays: https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/arrays/jagged-arrays
Public Function IsJaggedArray(ByVal arr As Variant) As Boolean

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

    IsJaggedArray = True

Escape:
End Function

Public Function ArrayBase0(ByRef aSourceArray As Variant)
    ReDim Preserve aSourceArray(0 To UBound(aSourceArray) - 1)
End Function

Public Function ArrayBase0_2ndDimension(ByRef aSpreadSheetArray As Variant)
    ReDim Preserve aSpreadSheetArray(LBound(aSpreadSheetArray) To UBound(aSpreadSheetArray, 1), 0 To UBound(aSpreadSheetArray, 2) - 1)
    aSpreadSheetArray = Core.Arr2DToJagArr(aSpreadSheetArray)
    ReDim Preserve aSpreadSheetArray(0 To UBound(aSpreadSheetArray) - 1)
End Function

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

''' Array("a", "b", "c", "d") => 1,2 => Array("b","c")
Public Function ArraySlice(ByVal arr As Variant, ByVal fst As Long, Optional snd As Long = 0) As Variant
    If snd = 0 Then
        ArraySlice = Application.index(arr, 0, Array(fst + 1))
    Else
        If fst > snd Then Err.Raise 9999
        Dim ary: ReDim ary(0 To snd - fst) As Long
        Dim i As Long
        For i = 0 To (snd - fst): ary(i) = i + (fst + 1): Next i

        ArraySlice = Application.index(arr, 0, ary)
    End If
End Function

''' Array("15.0", "16.0", "16.0", "Common", "Outlook") => "\d\d\.\d" => Array("15.0", "16.0")
Public Function ArrayRegexFilter(ByVal arr As Variant, ByVal ptrn As String) As Variant

    Dim regx As Object: Set regx = CreateObject("VBScript.RegExp")
    regx.Pattern = ptrn: regx.IgnoreCase = True: regx.Global = True

    Dim v, dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    For Each v In arr
        If regx.TEST(v) Then
            If Not dict.exists(v) Then dict.Add v, ""
        End If
    Next v

    ArrayRegexFilter = dict.keys

    Set dict = Nothing
    Set regx = Nothing

End Function

Public Function ArrayTranspose(ByVal arr2D As Variant) As Variant

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

    ArrayTranspose = tmpArr2D

End Function

Public Function ArrayHasEmpties(ByRef aSourceArray As Variant) As Boolean

    If ArrRank(aSourceArray) <> 1 Then
        Err.Raise 9999, , "SourceArray must be a single dimensional array."
    End If

    Dim i As Long
    For i = LBound(aSourceArray) To UBound(aSourceArray)
        If aSourceArray(i) = "" Or IsEmpty(aSourceArray(i)) Then
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
        If Not (aSourceArray(i) = "" Or IsEmpty(aSourceArray(i))) Then
            If IsObject(i) Then
                arrx.AddObj aSourceArray(i)
            Else
                arrx.AddVal aSourceArray(i)
            End If
        End If
    Next i
    ArrayRemoveEmpties = arrx.ToArray
    Set arrx = Nothing
End Function
               
Public Function ArrayWindow(ByVal arr As Variant, ByVal GroupN As Variant) As Variant

    ' Array(1..10) divided by 3
    ' -------------------------
    ' => Array(1%, 2%, 3%, 4%)
    ' => Array(5%, 6%, 7%)
    ' => Array(8%, 9%, 10%)
    
    ''' dependence: ariawase Core.ArrSlice
    

    ''' guard
    If Not IsArray(arr) Then Err.Raise 13
    If ArrRank(arr) > 1 Then Err.Raise 13
    If LBound(arr) < 0 Then Err.Raise 13
    
    ''' guard2( GroupN )
    Select Case GroupN
        Case Is <= 0: Err.Raise 13
        Case Is = 1:  ArrayWindow = Array(arr): GoTo Ending
        Case Is >= (UBound(arr) + 1)
            Dim tmpArray(): tmpArray = Array(): ReDim tmpArray(0 To UBound(arr))
            Dim idx As Long
            For idx = 0 To UBound(arr)
                tmpArray(idx) = Array(arr(idx))
            Next idx
            ArrayWindow = tmpArray
            GoTo Ending
        Case Else
            GoTo ArrayWindowImpl
    End Select
   
   
ArrayWindowImpl:

    Dim groupIndex As Long: groupIndex = Int(ArrLen(arr) / GroupN)
    Dim rest As Long: rest = ArrLen(arr) Mod GroupN
    
    ''' simple divison : e.g. 8 / 3 => array(2,2,2)
    Dim groupIndexArray(): groupIndexArray = Array(): ReDim groupIndexArray(0 To GroupN - 1)
    Dim i As Long
    For i = 0 To GroupN - 1
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
    Dim ary(): ary = Array(): ReDim ary(0 To GroupN - 1)
    Dim k As Long, acc_idx As Long
    For k = 0 To UBound(groupIndexArray)
        ary(k) = Core.ArrSlice(arr, acc_idx, acc_idx + groupIndexArray(k) - 1)
        acc_idx = acc_idx + groupIndexArray(k)
    Next k
    
    ArrayWindow = ary
    
Ending:
End Function

''' This function is helper function for AdoEx class.
Public Function ArraySelect(ByVal dbType As dbTypeEnum, ByVal sql As String, _
    Optional ByVal fPath As String = "", _
    Optional ByVal isTableHeader As Boolean = True) As Variant

    Dim adox As AdoEx: Set adox = New AdoEx
    adox.init dbType, fPath, isTableHeader
    Dim arr As Variant: arr = adox.ArraySelect(sql)
    If IsEmpty(arr) Then GoTo Escape
    ArraySelect = Array(arr(0), arr(1))
    Set adox = Nothing

Escape:
End Function

''' ----- temporary

Public Function ArrayGroupV1(ByVal arr As Variant) As Variant
    
    ''' value x 1
    '''    array( array( keyA, valA ), array( keyB, valB ) )
    ''' => array( array( keyA, keyB), array( array( valA ), array( valB ) ) )
    
    ''' group by
    Dim v
    Dim dict As Dictionary: Set dict = New Dictionary
    Dim arrx As ArrayEx: Set arrx = New ArrayEx
    Dim i As Long
    For i = 0 To UBound(arr)
        
        If i = 0 Then
            dict.Add arr(i)(0), ""
            arrx.AddVal arr(i)(1)
        ElseIf i = UBound(arr) Then
            If arr(i)(0) = arr(i - 1)(0) Then
                arrx.AddVal arr(i)(1)
                dict.Item(arr(i - 1)(0)) = arrx.ToArray()
            Else
                dict.Item(arr(i - 1)(0)) = arrx.ToArray()
                dict.Add arr(i)(0), arr(i)(1)
            End If
        ElseIf arr(i)(0) = arr(i - 1)(0) Then
            arrx.AddVal arr(i)(1)
        Else
            dict.Item(arr(i - 1)(0)) = arrx.ToArray()
            dict.Add arr(i)(0), ""
            Set arrx = Nothing
            Set arrx = New ArrayEx
            arrx.AddVal arr(i)(1)
        End If

    Next i
        
    ArrayGroupV1 = Array(dict.Keys, dict.Items)
    Set arrx = Nothing
    Set dict = Nothing
End Function

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