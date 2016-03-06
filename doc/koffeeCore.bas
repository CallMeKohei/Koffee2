Sub Sample_KoffeeCore20150820()

    'ワークシートの2次元表をSQLで取得する
    Debug.Print Dump(FetchSh("Select * From [Sheet1$B3:D6]"))
    'Array(
    '      Array("Apple", "Grape", "Kiwi")
    '    , Array(30#, 20#, 50#)
    '    , Array(300#, 400#, 200#)
    ')

    
    '任意の範囲でランダムな数字をえらぶ
    Debug.Print RandomBetween(100, 110)
    '107（そのときによる）
    
    
    '配列の要素をシャッフルする
    Debug.Print Dump(ArrShuffle(Array("A", "B", "C")))
    'Array("B", "A", "C")（そのときによる）
    
    
    'ジャグ配列かどうか
    Debug.Print IsJagArr(Array(Array(1, 2), 3))
    'True
    
    
    'Emptyを後ろから切り詰める
    Debug.Print Dump(Truncate(Array(Empty, 1, 2, Empty, Empty, Empty)))
    'Array((Empty), 1%, 2%)
    
    
    '配列のベースをゼロにする
    Debug.Print LBound(Base01(Array(1, 2)))
    '0
    
    
    '配列のベースを壱にする
    Debug.Print LBound(Base01(Array(1, 2), True))
    '1
    
    
    '配列の要素をすべて1次元にする
    Debug.Print Dump(ArrExplode(Array(Array(1, Array(2, Array(3))))))
    'Array(1%, 2%, 3%)
    
    
    '配列の要素を任意の要素数に変更して任意の文字・数でうめる
    Debug.Print Dump(ArrFill(Array(1, 2), 4, "A"))
    'Array(1%, 2%, "A", "A", "A")
    
    
    '配列の最初の要素以外の配列をかえす
    Debug.Print Dump(Rest(Array(1, 2, 3, 4, 5)))
    'Array(2%, 3%, 4%, 5%)
    
    
    '配列の要素をソートする
    Debug.Print Dump(ArrSortAsc(Array(3, 4, 1, 5, 2)))
    'Array(1%, 2%, 3%, 4%, 5%)
    
    Debug.Print Dump(ArrSortDec(Array(3, 4, 1, 5, 2)))
    'Array(5%, 4%, 3%, 2%, 1%)
    
    
    '要素数が同じ２つの配列を足す
    Debug.Print Dump(ArrPlus(Array(1, 2, 3), Array(10, 20, 30)))
    'Array(11@, 22@, 33@)
    

    '要素数が同じ２つの配列をひく
    Debug.Print Dump(ArrMinus(Array(1, 2, 3), Array(10, 20, 30)))
    'Array(-9@, -18@, -27@)
    
    
    '集合Aと集合Bをたす
    Debug.Print Dump(ArrUnion(Array("A", "B", "C", "D"), Array("C", "D", "E", "F")))
    'Array("A", "B", "C", "D", "E", "F")
    
    
    '集合Aから集合Bをひく
    Debug.Print Dump(ArrDiff(Array("A", "B", "C", "D"), Array("C", "D", "E", "F")))
    'Array("A", "B")
    
    
    '集合Aと集合Bの共通部分をのぞく
    Debug.Print Dump(ArrDiff2(Array("A", "B", "C", "D"), Array("C", "D", "E", "F")))
    'Array("A", "B", "E", "F")
    
    
    '集合Aと集合Bの共通部分
    Debug.Print Dump(ArrIntersect(Array("A", "B", "C", "D"), Array("C", "D", "E", "F")))
    'Array("C", "D")
    
    
    '配列の先頭に値をいれる
    Debug.Print Dump(ArrShift("A", Array(1, 2, 3)))
    'Array("A", 1%, 2%, 3%)
    
    
    '配列の先頭をのぞく
    Debug.Print Dump(ArrUnshift(Array(1, 2, 3)))
    'Array(2%, 3%)
    
    
    '配列の末尾に値をいれる
    Debug.Print Dump(ArrPush("A", Array(1, 2, 3)))
    'Array(1%, 2%, 3%, "A")
    
    
    '配列の末尾の値をのぞく
    Debug.Print Dump(ArrPop(Array(1, 2, 3)))
    'Array(1%, 2%)


    '配列の要素をロング型にする
    Debug.Print Dump(ArrCLng(Array(1, 2, 3)))
    'Array(1&, 2&, 3&)
    
    
    '配列の要素をカレント型にする
    Debug.Print Dump(ArrCCur(Array(1, 2, 3)))
    'Array(1@, 2@, 3@)


    '累計をとった配列をかえす
    Debug.Print Dump(StepTotal(Array(Array(1, 2, 3), Array(4, 5, 6))))
    'Array(Array(1@, 3@, 6@), Array(4@, 9@, 15@))

    
End Sub
