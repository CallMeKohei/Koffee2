Sub Sample_KoffeeExcel20150821()

    Dim sh As Worksheet: Set sh = Worksheets("Sheet1")
    
    
    '配列の値をワークシートにかきこむ
    PutVal Array("A", "B", (Empty), "Z"), sh.Range("A1"), True
    PutVal Array("A", "B", "C", "D", (Empty), "Z"), sh.Range("A1")
    
    
    'ワークシートの値を配列にする
    Debug.Print Dump(GetVal(sh.Range("A1"), Vertical))   'Array("A", "B", (Empty), "Z")
    Debug.Print Dump(GetVal(sh.Range("A1"), Holizontal)) 'Array("A", "B", "C", "D", (Empty), "Z")
    
    
    '最終行を取得する
    Debug.Print LastRow(sh.Range("A1"))         '4
    Debug.Print LastRow(sh.Range("A1"), True)   '2

    '最終列を取得する
    Debug.Print LastCol(sh.Range("A1"))         '6
    Debug.Print LastCol(sh.Range("A1"), True)   '4
    
    'ワークシートの名前を配列にする
    Debug.Print Dump(ArrSheetsName(ThisWorkbook)) 'Array("Sheet1")
    
    'ワークシートに名前があるかどうか
    Debug.Print ExistsSheet("Sheet1") 'True
    
    'ワークシートをついかする
    AddSheet "test"
    
    'ワークシートをコピーする
    CopySheet "test", "testCopied"
    
End Sub
