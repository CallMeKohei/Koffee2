Sub Sample_KoffeeTable20150821()

    'もとデーター
    Dim Header: Header = Array("DEN_NO", "FruitName")
    Dim Denpyo: Denpyo = Array(719750, 719750, 719750, 719750, 719750, 719750, 719751, 719751, 719751, 719751, 741900, 741900)
    Dim Fruit:  Fruit = Array("りんご", "ドラゴンフルーツ", "パパイア", "ブルーベリー", "マンゴスチン", "ぶどう", "スイカ", "さくらんぼ", "もも", "グレープフルーツ", "キウイ", "アボガド")
    
    
    
    'テーブルをつくる: CreateTable
    Dim tbl: tbl = CreateTable(Header, Array(Denpyo, Fruit))
    Debug.Print Dump(tbl)
    
    'Array(
    '      Array("DEN_NO", "FruitName")
    '    , Array(
    '          Array(719750&, 719750&, 719750&, 719750&, 719750&, 719750&, 719751&, 719751&, ...)
    '        , Array("りんご", "ドラゴンフルーツ", "パパイア", "ブルーベリー", "マンゴスチン", "ぶどう", "スイカ", "さくらんぼ", ...)
    '    )
    ')

    
    'キーの取り出し:Keys
    Debug.Print Dump(keys(tbl))
    
    'Array("DEN_NO", "FruitName")
    
    
    '値(value)の取り出し:Values
    Debug.Print Dump(Values(tbl))
    
    'Array(
    '      Array(719750&, 719750&, 719750&, 719750&, 719750&, 719750&, 719751&, 719751&, ...)
    '    , Array("りんご", "ドラゴンフルーツ", "パパイア", "ブルーベリー", "マンゴスチン", "ぶどう", "スイカ", "さくらんぼ", ...)
    ')
    
    
    
    '任意の列の値だけとりだし:Pluck
    Debug.Print Dump(Pluck(tbl, "FruitName"))
    
    'Array("りんご", "ドラゴンフルーツ", "パパイア", "ブルーベリー", "マンゴスチン", "ぶどう", "スイカ", "さくらんぼ", ...)
    

    '任意の列だけとりだし:Project
    Debug.Print Dump(Project(Array("FruitName"), tbl))

    'Array(
    '      Array("FruitName")
    '    , Array(
    '        Array("りんご", "ドラゴンフルーツ", "パパイア", "ブルーベリー", "マンゴスチン", "ぶどう", "スイカ", "さくらんぼ", ...)
    '    )
    ')
'
'
'
'    '任意の条件でレコードをとりだし:Restrict
    Dim f As Func: Set f = Init(New Func, vbBoolean, AddressOf IsApple, vbObject)
    Debug.Print Dump(Restrict(f, tbl))

    'Array(
    '      Array("DEN_NO", "FruitName")
    '    , Array(
    '         Array(719750&)
    '        , Array("りんご")
    '    )
    ')

End Sub

Private Function IsApple(ByVal rcd As Object) As Boolean
    If rcd("FruitName") = "りんご" Then
        IsApple = True
    End If
End Function
