Sub Sample_KoffeeFunc20150821()

    'Map
    Debug.Print Dump(Flow(Array(1, 2, 3), Map(Partial(Add, 100))))
    'Array(101%, 102%, 103%)
    
    'Fold
    Debug.Print Dump(Flow(Array(1, 2, 3), Fold(div)))
    '0.166666666666667#
    
    'FoldR
    Debug.Print Dump(Flow(Array(1, 2, 3), FoldR(div)))
    '1.5#
    
    'Scan
    Debug.Print Dump(Flow(Array(1, 2, 3), Scan(Add)))
    'Array(1%, 3%, 6%)
    
    'ScanR
    Debug.Print Dump(Flow(Array(1, 2, 3), ScanR(Add)))
    'Array(3%, 5%, 6%)
    
    'Zip
    Debug.Print Dump(Flow(Array(Array(1, 2, 3), Array("A", "B", "C")), Zip(Array2Of)))
    'Array(Array(1%, "A"), Array(2%, "B"), Array(3%, "C"))
    
    'Filter
    Debug.Print Dump(Flow(Array("a", "b", 3, "d"), Filter(AddressOf IsNumber)))
    'Array(3%)
    
    'Reject
    Debug.Print Dump(Flow(Array("a", "b", 3, "d"), Reject(AddressOf IsNumber)))
    'Array("a", "b", "d")
    
    'Find
    Debug.Print Dump(Flow(Array("a", "b", 3, "d"), Find(AddressOf IsNumber)))
    '3
    

    'AllOf, AnyOf
    Dim T As Func: Set T = Init(New Func, vbBoolean, AddressOf Tr)
    Dim f As Func: Set f = Init(New Func, vbBoolean, AddressOf Fa)
    
    Debug.Print Dump(AllOf())       'True
    Debug.Print Dump(AllOf(T, T))   'True
    Debug.Print Dump(AllOf(T, T, T, T, f))  'False
    
    Debug.Print Dump(AnyOf())           'False
    Debug.Print Dump(AnyOf(T, T, f))    'True
    Debug.Print Dump(AnyOf(f, f, f, f)) 'False
    
    
    'All, Any_
    Debug.Print Dump(Flow(Array("a", "b", 3, "d"), All(AddressOf IsNumber)))
    'False
    
    Debug.Print Dump(Flow(Array("a", "b", 3, "d"), Any_(AddressOf IsNumber)))
    'True

    
    'Partial
    Debug.Print Partial(div, 20).Apply(10)
    ' 2
    
    'fnull
    Debug.Print Dump(Flow(Array(1, Null, 3), Fold(fnull(Mult, 100))))
    '300%
    
    
    'Actions
    Dim StackAction As Atom
    Set StackAction = Actions(Lift(Partial(ShiftA, 99)) _
                            , Lift(Partial(ShiftA, 55)) _
                            , Lift(PopA))
                            
    Debug.Print Dump(Flow(Array(1, 2), StackAction))
'    Tuple`2 {
'      Item1 = Array(55, 99, 1)
'    , Item2 = Array(
'                  Array(99%, 1%, 2%)
'                , Array(55%, 99%, 1%, 2%)
'                , Array(55%, 99%, 1%)
'            )
'    }
    
    
End Sub

Private Function Tr() As Boolean
    Tr = True
End Function

Private Function Fa() As Boolean
    Fa = False
End Function

Private Function IsNumber(ByVal x As Variant) As Boolean
    IsNumber = IsNumeric(x)
End Function
