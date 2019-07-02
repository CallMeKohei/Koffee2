Attribute VB_Name = "SampleKoffee"
''' SampleKoffee.bas
''' written by callmekohei(twitter at callmekohei)
''' MIT license
Option Explicit
Option Compare Text
Option Private Module
Option Base 0

Private Sub Sample_koffeeArray()

    '�W���O�z�񂩂ǂ���
    Debug.Print IsJaggedArray(Array(1, 2, 3))
    ''' False
    Debug.Print IsJaggedArray(Array(Array(1), Array(2)))
    ''' True


    ''' �x�[�X�P�̂Q�����z����x�[�X�[���̃W���O�z��ɂ���
    Dim arr: arr = ThisWorkbook.Worksheets("Sheet1").Range("A1:C3")
    koffeeArray.ArrayBase0_2ndDimension arr
    Debug.Print LBound(arr)
    '0


    ''' �W���O�z�񂩂�C�ӂ̗�����o��
    Dim jagArr: jagArr = Array(Array(1, 2, 3), Array(4, 5, 6), Array(7, 8, 9))
    Debug.Print Dump(ArrayColumn(1, jagArr))
    ''' Array(2%, 5%, 8%)


    ''' �z����X���C�X����
    arr = Array(1, 2, 3, 4, 5)
    Debug.Print Dump((ArraySlice(arr, 1, UBound(arr))))
    ''' Array(2#, 3#, 4#, 5#)


    ''' �z��̒l�̗v�f�𐳋K�\���Ńt�B���^�����O����
    arr = Array("15.0", "16.0", "16.0", "Common", "Outlook")
    Debug.Print Dump(ArrayRegexFilter(arr, "\d\d\.\d"))
    ''' Array("15.0", "16.0")


    ''' �Q�����z����g�����X�|�[�Y����
    jagArr = Array(Array(1, 2, 3), Array(4, 5, 6), Array(7, 8, 9))
    Debug.Print Dump(Arr2DToJagArr(ArrayTranspose(JagArrToArr2D(jagArr))))
    ''' Array(Array(1%, 4%, 7%), Array(2%, 5%, 8%), Array(3%, 6%, 9%))


    ''' �P�����z��ɂċ󔒗v�f�������͋󔒕��������邩�ǂ���
    Debug.Print ArrayHasEmpties(Array("a", "", "c"))
    ''' True


    ''' �P�����z��ɂċ󔒗v�f�������͋󔒕������폜����
    Debug.Print Dump(ArrayRemoveEmpties(Array("a", "", "c", Empty)))
    ''' Array("a", "c")


    ''' �G�N�Z���V�[�g�̒l��z��ɂ���

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

    ''' ���[�N�V�[�g������΍폜����
    If koffeeExcel.ExistsSheet("test_worksheet") Then
        koffeeExcel.DeleteSheet "test_worksheet"
    End If

    '���[�N�V�[�g����������
    Set ws = AddSheet("test_worksheet")

    '���[�N�V�[�g�ɖ��O�����邩�ǂ���
    Debug.Print ExistsSheet("test_worksheet") 'True

    '�z��̒l�����[�N�V�[�g�ɂ�������
    PutVal Array("X", "Y"), ws.Range("A1")
    PutVal Array("a", "b"), ws.Range("A2")
    PutVal Array(1, 2), ws.Range("A3")

    '���[�N�V�[�g�̒l��z��ɂ���
    Debug.Print Dump(GetVal(ws.Range("A1").CurrentRegion, False))  ''' Array(Array("X", "Y"), Array("a", "b"), Array(1#, 2#))
    Debug.Print Dump(GetVal(ws.Range("A1").CurrentRegion, True))   ''' Array(Array("X", "a", 1#), Array("Y", "b", 2#))


    '�ŏI�s���擾����
    Debug.Print LastRow(ws.Range("A1"))         '3
    Debug.Print LastRow(ws.Range("A1"), True)   '3

    '�ŏI����擾����
    Debug.Print LastCol(ws.Range("A1"))         '2
    Debug.Print LastCol(ws.Range("A1"), True)   '2

    '���[�N�V�[�g�̖��O��z��ɂ���
    Debug.Print Dump(ArrSheetsName(ThisWorkbook)) ''' Array("Sheet1", "test_worksheetAdded", "test_worksheetCopied", "test_worksheet")

    ''' ���[�N�V�[�g��ǉ�����
    AddSheet "test_worksheetAdded"

    '���[�N�V�[�g���R�s�[����
    CopySheet "test_worksheet", "test_worksheetCopied"

End Sub

Private Sub Sample_koffeeIO()

    Dim arr

    ''' csv �t�@�C����ǂݍ���
    arr = koffeeIO.ArrCsv("path/to/csv", "UTF-8", adLF)

    ''' shift-jis �� csv�t�@�C����ǂݍ���
    arr = koffeeIO.ArrCsvSjis("path/to/sjisCsv")

    ''' �e�L�X�g�t�@�C�����쐬���ĕ�������������
    koffeeIO.CreateTextFile "path/to/txt", "hello world"

    ''' �e�L�X�g�t�@�C���ɕ�����ǋL����
    koffeeIO.AppendText "path/to/txt", "foo bar baz"

End Sub

Private Sub Sample_koffeeTime()

    ''' �^�C�}�[�X�^�[�g
    Dim start As Double: start = koffeetime.MilliSecondsTimer

    ''' �R�b�Ƃ߂�
    Wait 3000

    ''' �^�C�}�[�A�E�g
    Debug.Print koffeetime.MilliSecondsTimer - start '''  3008.38700908422

End Sub
