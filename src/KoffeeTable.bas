Attribute VB_Name = "KoffeeTable"
'  +--------------                                         --------------+
'  |||||||||    Koffee2 0.1.0                                            |
'  |: ^_^ :|    Koffee2 is free Library based on Ariawase.               |
'  |||||||||    The Project Page: https://github.com/CallMeKohei/Koffee2 |
'  +--------------                                         --------------+
Option Explicit

Public Function CreateTable(ByVal ArrKeys As Variant, ArrValues As Variant) As Variant
    CreateTable = Fill(Array(ArrKeys, ArrValues))
End Function

Public Function CreateTableArray(ByVal header As Variant, ByVal vals As Variant) As Variant
    CreateTableArray = Array(header, vals)
End Function

Public Function keys(ByVal tbl As Variant) As Variant
    keys = tbl(0)
End Function

Public Function Values(ByVal tbl As Variant) As Variant
    Values = ArrFlatten(Rest(tbl))
End Function

Public Function Pluck(ByVal tbl As Variant, ByVal ColumnName As Variant) As Variant
    Pluck = Values(tbl)(ArrIndexOf(keys(tbl), ColumnName))
End Function

Public Function Record(ByVal tbl As Variant, ByVal n As Long) As Variant
    
    Dim a(): ReDim a(32)
    Dim ub As Long:  ub = 32
    Dim i As Long: i = 0
    
    Dim v, arrx As New ArrayEx
    For Each v In Values(tbl)
    
        If ub = i Then
            ub = ub + 1
            ub = -1 + ub + ub
            ReDim Preserve a(ub - 1)
        End If
    
        a(i) = v(n)
        i = i + 1
        
    Next v
    
    Record = Truncate(a)

End Function

Public Function ToRecord(ByVal tbl As Variant, Optional ByVal ExistsKeys As Boolean = True) As Variant

    Dim a(): ReDim a(32)
    Dim ub As Long:  ub = 32
    
    Dim vals, ub2 As Long
    If ExistsKeys Then
        vals = Values(tbl)
    Else
        vals = IIf(IsJagArr(tbl), tbl, Array(tbl))
    End If
    
    Dim v, i As Long, arrx As New ArrayEx
    For i = 0 To UBound(vals(0))
    
        For Each v In vals
            arrx.addval v(i)
        Next v
        
        If ub = i Then
            ub = ub + 1
            ub = -1 + ub + ub
            ReDim Preserve a(ub - 1)
        End If
        
        a(i) = arrx.ToArray
        
        Set arrx = Nothing
    Next i
    ToRecord = Truncate(a)
    
End Function

Public Function ToTable(ByVal ArrKey As Variant, ByVal Records As Variant) As Variant

    Dim a(): ReDim a(32)
    Dim ub As Long:  ub = 32
    
    Dim v, i As Long, arrx As New ArrayEx
    For i = 0 To UBound(ArrKey)
    
        For Each v In Records
            arrx.addval v(i)
        Next v
    
        If ub = i Then
            ub = ub + 1
            ub = -1 + ub + ub
            ReDim Preserve a(ub - 1)
        End If
        
        a(i) = arrx.ToArray
        Set arrx = Nothing
        
    Next i
    
    ToTable = CreateTable(ArrKey, Truncate(a))
    
End Function

Public Function Project(ByVal ArrKeys As Variant, ByVal tbl As Variant) As Variant
    
    Dim v, arrx As New ArrayEx
    For Each v In ArrKeys
        arrx.addval ArrIndexOf(keys(tbl), v)
    Next v
    
    Dim nk, header As New ArrayEx
    For Each nk In arrx.ToArray
        header.addval keys(tbl)(nk)
    Next nk
    
    Dim nv, vals As New ArrayEx
    For Each nv In arrx.ToArray
        vals.addval Values(tbl)(nv)
    Next nv
    
    Project = CreateTable(header.ToArray, vals.ToArray)
    
End Function

Public Function WHERE(ByVal pred As Variant) As Atom
    Dim a As New Atom: Set WHERE = a.AddFunc(Init(New Func, vbVariant, AddressOf Restrict, vbVariant, vbVariant)).Bind(pred)
End Function

Public Function Restrict(ByVal pred As Variant, ByVal tbl As Variant) As Variant
    
    Dim keyTbl: keyTbl = keys(tbl)
    
    Dim rcd, arrx As New ArrayEx
    For Each rcd In ToRecord(tbl)
        
        Dim dict As Object: Set dict = CreateDictionary()
        Dim ky, i As Long: i = 0
        For Each ky In keyTbl
            dict.Add key:=ky, Item:=rcd(IncrPst(i))
        Next ky
    
        If pred.Apply(dict) Then
            arrx.addval rcd
        End If
        
    Next rcd
    
    Restrict = ToTable(keyTbl, arrx.ToArray)
    
End Function

Private Function Fill(ByVal tbl As Variant) As Variant

    Dim a(): ReDim a(32)
    Dim ub As Long:  ub = 32
    Dim i As Long: i = 0
    
    Dim vals: vals = Values(tbl)
    Dim Max As Long
    
    Dim n
    For Each n In vals
        If UBound(n) > Max Then Max = UBound(n)
    Next n
    
    Dim v
    For Each v In vals
        If ub = i Then
            ub = ub + 1
            ub = -1 + ub + ub
            ReDim Preserve a(ub - 1)
        End If
    
        If UBound(v) < Max Then
            a(i) = ArrFill(v, Max, Null)
            i = i + 1
        Else
            a(i) = v
            i = i + 1
        End If
    Next v
    
    Fill = CreateTableArray(keys(tbl), Truncate(a))
    
End Function

Private Function ArrPickup(ByVal arr As Variant, ByVal indexs As Variant) As Variant

    Dim a(): ReDim a(32)
    Dim ub As Long:  ub = 32
    Dim i As Long: i = 0

    Dim v, arrx As New ArrayEx
    For Each v In indexs

        If ub = i Then
            ub = ub + 1
            ub = -1 + ub + ub
            ReDim Preserve a(ub - 1)
        End If

        a(i) = arr(v)
        i = i + 1
        
    Next v
    
    ArrPickup = Truncate(a)
End Function

Public Function RejectArr(ByVal tbl1 As Variant, ByVal tbl2 As Variant) As Variant

    Dim a(): ReDim a(32)
    Dim ub As Long: ub = 32
    Dim i As Long: i = 0
    
    Dim key: key = keys(tbl1)
    Dim rec1: rec1 = ToRecord(tbl1)
    Dim rec2: rec2 = ToRecord(tbl2)
    
    Dim v
    For Each v In rec1
        If Not ArrExists(rec2, v) Then
            
            If ub = i Then
                ub = ub + 1
                ub = -1 + ub + ub
                ReDim Preserve a(ub - 1)
            End If
            
            a(i) = v
            i = i + 1
        End If
    Next v
    
    RejectArr = ToTable(key, Truncate(a))
End Function

Public Function ArrExists(ByVal jagArr As Variant, ByVal arr As Variant) As Boolean
    Dim vArr
    For Each vArr In jagArr
        If ArrEquals(arr, vArr) Then
            ArrExists = True
            GoTo Escape
        End If
    Next vArr
Escape:
End Function

Public Function Intersect(ByVal tbl1 As Variant, ByVal idx1 As Variant, ByVal tbl2 As Variant, ByVal idx2 As Variant) As Variant
    
    Dim key1: key1 = keys(tbl1)
    Dim key2: key2 = keys(tbl2)
    Dim rcd1: rcd1 = ToRecord(tbl1)
    Dim rcd2: rcd2 = ToRecord(tbl2)
    
    Dim v1, v2, comA As New ArrayEx, comB As New ArrayEx, comAB As New ArrayEx
    For Each v1 In rcd1
        For Each v2 In rcd2
            If ArrEquals(ArrPickup(v1, idx1), ArrPickup(v2, idx2)) Then
                comA.addval v1
                comB.addval v2
                comAB.AddObj Init(New Tuple, v1, v2)
            End If
        Next v2
    Next v1
    
    Intersect = Array(ToTable(key1, comA.ToArray), ToTable(key2, comB.ToArray), comAB)
        
End Function

