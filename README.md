## KoffeeVBA

callmekohei の VBA をラクに書くためのコード

### startup script

`gitbash`上で適当なフォルダを作り下記のコードを実行します

```bash
#! /bin/bash

mkdir -p ./src/foo.xlsm/

git clone --depth 1 https://github.com/callmekohei/koffeeVBA
git clone --depth 1 https://github.com/callmekohei/ariawaseModified

cp ariawaseModified/vbac.wsf ./

mv ariawaseModified/src/Ariawase.xlsm/* ./src/foo.xlsm/
mv koffeeVBA/src/koffee.xlsm/* ./src/foo.xlsm/

cscript vbac.wsf combine

rm -rf koffeeVBA
rm -rf ariawaseModified

cd bin
explorer foo.xlsm

```

### Sample code

```vb
Private Sub Sample_koffeeArray()

    'ジャグ配列かどうか
    Debug.Print IsJaggedArray(Array(1, 2, 3))
    ''' False
    Debug.Print IsJaggedArray(Array(Array(1), Array(2)))
    ''' True


    ''' ベース１の２次元配列をベースゼロのジャグ配列にする
    Dim arr: arr = ThisWorkbook.Worksheets("Sheet1").Range("A1:C3")
    koffeeArray.ArrayBase0_2ndDimension arr
    Debug.Print LBound(arr)
    '0


    ''' ジャグ配列から任意の列を取り出す
    Dim jagArr: jagArr = Array(Array(1, 2, 3), Array(4, 5, 6), Array(7, 8, 9))
    Debug.Print Dump(ArrayColumn(1, jagArr))
    ''' Array(2%, 5%, 8%)


    ''' 配列をスライスする
    arr = Array(1, 2, 3, 4, 5)
    Debug.Print Dump((ArraySlice(arr, 1, UBound(arr))))
    ''' Array(2#, 3#, 4#, 5#)


    ''' 配列の値の要素を正規表現でフィルタリングする
    arr = Array("15.0", "16.0", "16.0", "Common", "Outlook")
    Debug.Print Dump(ArrayRegexFilter(arr, "\d\d\.\d"))
    ''' Array("15.0", "16.0")


    ''' ２次元配列をトランスポーズする
    jagArr = Array(Array(1, 2, 3), Array(4, 5, 6), Array(7, 8, 9))
    Debug.Print Dump(Arr2DToJagArr(ArrayTranspose(JagArrToArr2D(jagArr))))
    ''' Array(Array(1%, 4%, 7%), Array(2%, 5%, 8%), Array(3%, 6%, 9%))


    ''' １次元配列にて空白要素もしくは空白文字があるかどうか
    Debug.Print ArrayHasEmpties(Array("a", "", "c"))
    ''' True


    ''' １次元配列にて空白要素もしくは空白文字を削除する
    Debug.Print Dump(ArrayRemoveEmpties(Array("a", "", "c", Empty)))
    ''' Array("a", "c")


    ''' エクセルシートの値を配列にする

    '''     |  A   B   C   D
    ''' ----+----------------
    '''   1 |  X   Y   Z
    '''   2 |  a   b   c
    '''   3 |  1   2   3
    '''   4 |

    Debug.Print Dump(ArraySelect(dbExcel, "select * from [Sheet1$]"))
    ''' Array(Array("X", "Y", "Z"), Array(Array("a", "b", "c"), Array("1", "2", "3")))

End Sub

Private Sub Sample_koffeeExcel()

    Dim ws As Worksheet

    ''' ワークシートがあれば削除する
    If koffeeExcel.ExistsSheet("test_worksheet") Then
        koffeeExcel.DeleteSheet "test_worksheet"
    End If

    'ワークシートをついかする
    Set ws = AddSheet("test_worksheet")

    'ワークシートに名前があるかどうか
    Debug.Print ExistsSheet("test_worksheet") 'True

    '配列の値をワークシートにかきこむ
    PutVal Array("X", "Y"), ws.Range("A1")
    PutVal Array("a", "b"), ws.Range("A2")
    PutVal Array(1, 2), ws.Range("A3")

    'ワークシートの値を配列にする
    Debug.Print Dump(GetVal(ws.Range("A1").CurrentRegion, False))  ''' Array(Array("X", "Y"), Array("a", "b"), Array(1#, 2#))
    Debug.Print Dump(GetVal(ws.Range("A1").CurrentRegion, True))   ''' Array(Array("X", "a", 1#), Array("Y", "b", 2#))


    '最終行を取得する
    Debug.Print LastRow(ws.Range("A1"))         '3
    Debug.Print LastRow(ws.Range("A1"), True)   '3

    '最終列を取得する
    Debug.Print LastCol(ws.Range("A1"))         '2
    Debug.Print LastCol(ws.Range("A1"), True)   '2

    'ワークシートの名前を配列にする
    Debug.Print Dump(ArrSheetsName(ThisWorkbook)) ''' Array("Sheet1", "test_worksheetAdded", "test_worksheetCopied", "test_worksheet")

    ''' ワークシートを追加する
    AddSheet "test_worksheetAdded"

    'ワークシートをコピーする
    CopySheet "test_worksheet", "test_worksheetCopied"

End Sub

Private Sub Sample_koffeeIO()

    Dim arr

    ''' csv ファイルを読み込む
    arr = koffeeIO.ArrCsv("path/to/csv", "UTF-8", adLF)

    ''' shift-jis の csvファイルを読み込む
    arr = koffeeIO.ArrCsvSjis("path/to/sjisCsv")

    ''' テキストファイルを作成して文字を書き込む
    koffeeIO.CreateTextFile "path/to/txt", "hello world"

    ''' テキストファイルに文字を追記する
    koffeeIO.AppendText "path/to/txt", "foo bar baz"

End Sub

Private Sub Sample_koffeeTime()

    ''' タイマースタート
    Dim start As Double: start = koffeetime.MilliSecondsTimer

    ''' ３秒とめる
    Wait 3000

    ''' タイマーアウト
    Debug.Print koffeetime.MilliSecondsTimer - start '''  3008.38700908422

End Sub

```

### その他

こんなコードあるよとか、ここもう少しこうしたらとかあったらぜひ`issue`もしくは`twitter(@callmekohei)`にお願いします

### License

This software is released under the MIT License, see [LICENSE.txt](https://github.com/callmekohei/koffeeVBA/blob/master/LICENSE.txt).

