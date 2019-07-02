Attribute VB_Name = "koffeeIO"
''' koffeeIO.bas
''' written by callmekohei(twitter at callmekohei)
''' MIT license
Option Explicit
Option Compare Text
Option Private Module
Option Base 0

Public Function ArrCsv(ByVal srcPath As String _
    , ByVal chrset As String _
    , ByVal crrLinsep As LineSeparatorsEnum) As Variant

    ''' RemoveBom srcPath, chrset, crrLinsep

    Dim srcStrm As Object: Set srcStrm = CreateAdoDbStream(adTypeText, chrset, crrLinsep)
    srcStrm.Open
    srcStrm.LoadFromFile srcPath

    Dim dstStrm As Object
    Set dstStrm = ChangeCharset(srcStrm, "Shift-JIS")
    Set dstStrm = ChangeLineSeparator(dstStrm, adCRLF)
    Dim arr0 As Variant: arr0 = Split(dstStrm.ReadText, vbCrLf)
    dstStrm.Close

    ''' remove last array element ( CRLF )
    Dim i As Long, arr1() As Variant: ReDim arr1(0 To UBound(arr0) - 1)
    For i = 0 To UBound(arr0) - 1: arr1(i) = SplitCsv(arr0(i)): Next i
    ArrCsv = arr1

    Set srcStrm = Nothing
    Set dstStrm = Nothing

End Function

Public Function ArrCsvSjis(ByVal fp As String) As Variant

    Dim fileNumber As Integer: fileNumber = FreeFile
    Open fp For Binary As #fileNumber
        Dim buf() As Byte: ReDim buf(1 To LOF(fileNumber))
        Get #fileNumber, , buf
    Close #fileNumber

    Dim tmp As Variant: tmp = Split(StrConv(buf, vbUnicode), vbCrLf)

    Dim i As Long, tmpArr() As Variant: ReDim tmpArr(0 To UBound(tmp))
    For i = 0 To UBound(tmp): tmpArr(i) = SplitCsv(tmp(i)): Next i

    ArrCsvSjis = tmpArr

End Function

Private Function SplitCsv(ByVal csv As String) As Variant
    Dim ptrn As String: ptrn = ",(?=(?:[^""]*""[^""]*"")*[^""]*$)"
    Dim arr As Variant: arr = Split(ReReplace(csv, ptrn, ChrW(-1), "g"), ChrW(-1))
    Dim i As Long
    For i = 0 To UBound(arr): arr(i) = Replace(arr(i), """", ""): Next i
    SplitCsv = arr
End Function

Public Function CreateTextFile(ByVal aFilePath As String, ByVal aText As String) As Boolean

    On Error GoTo Err

    ''' create and override new error log file

    Dim Fso As Object: Set Fso = CreateObject("Scripting.FileSystemObject")
    With Fso.CreateTextFile( _
              FileName:=aFilePath _
            , overwrite:=True _
            , Unicode:=False _
        )
        .WriteLine Now
        .Close
    End With

    Set Fso = Nothing

    CreateTextFile = True
    Exit Function

Err:
End Function

Public Function AppendText(ByVal aFilePath As String, ByVal aText As String) As Boolean

    On Error GoTo Err

    ''' append text to file
    Dim Fso As Object: Set Fso = CreateObject("Scripting.FileSystemObject")
    Dim Ts As Object ''' Is TextStream
    Set Ts = Fso.OpenTextFile( _
          FileName:=aFilePath _
        , IOMode:=OpenFileEnum.ForAppending _
        , Create:=True _
        , Format:=TristateEnum.UseDefault _
    )
    With Ts
        .WriteLine aText
        .Close
    End With

    Set Ts = Nothing
    Set Fso = Nothing

    AppendText = True
    Exit Function

Err:
End Function

