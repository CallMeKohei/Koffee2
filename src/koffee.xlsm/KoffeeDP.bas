Attribute VB_Name = "KoffeeDP"
''' --------------------------------------------------------
'''  FILE    : KoffeeDP.bas
'''  AUTHOR  : callmekohei <callmekohei at gmail.com>
'''  License : MIT license
''' --------------------------------------------------------
Option Explicit

Sub DP(ByVal arr As Variant, Optional Lheader As Variant, Optional ByVal header As Variant, Optional ByVal Title As String, Optional ByVal n As Long = 10)

    If IsArray(arr) Then
        'If IsJagArr(arr) Then
            'If LBound(arr) <> 0 Then
                arr = Base01(arr)
            'End If
        'End If
    End If

    If Title <> "" Then Debug.Print Title

    If Not IsMissing(Lheader) Then

        If IsJagArr(arr) Then
            arr = ArrShift(Lheader, arr)
        Else
            arr = ArrShift(Lheader, Array(arr))
        End If

        If Not IsMissing(header) Then
            header = ArrShift("", header)
        End If

    End If

    If Not IsMissing(header) Then

        Dim v, arrx As ArrayEx: Set arrx = New ArrayEx
        For Each v In header
            arrx.AddVal ArrUnflatten(v)
        Next v

        DPImpl arrx.ToArray, n
        Debug.Print (rept("-", n * (UBound(arr) + 1)) & rept("-", UBound(arr)) & "-")

    End If

    DPImpl arr, n

End Sub


Sub DPImpl(ByVal arr As Variant, Optional ByVal n As Long = 10)

    If Not IsArray(arr) Then
        Debug.Print Dump(arr)
        GoTo Escape
    ElseIf Not IsJagArr(arr) Then
        arr = Array(arr)
    End If

    Dim i As Long, j As Long, buf As String
    For i = 0 To UBound(arr(0))
        For j = 0 To UBound(arr)
            If IsNumeric(arr(j)(i)) Then
                buf = buf & StrFixed(Format2(arr(j)(i), n), n) & "|"
            Else
                buf = buf & StrFixed(arr(j)(i), n) & "|"
            End If
        Next j
        Debug.Print buf
        buf = ""
    Next i

Escape:
End Sub

Public Function StrFixed(ByVal str As Variant, Optional ByVal n As Long) As Variant

    If IsObject(str) Or IsError(str) Then str = TypeName(str)
    If IsNumeric(str) And str < 0.0000000001 Then str = ARound(str, 5)

    Select Case LenB(StrConv(str, vbFromUnicode))
        Case Is < n:
            If IsNumeric(str) Then
                StrFixed = ReptPre(str, " ", n)
            Else
                StrFixed = ReptPst(str, " ", n)
            End If
        Case Is > n
            StrFixed = LeftA(str, n)
        Case Else
            StrFixed = str
    End Select

End Function

Public Function rept(ByVal s As String, ByVal n As Long) As Variant
    rept = Application.WorksheetFunction.rept(s, n)
End Function

Public Function ReptPst(ByVal str As String, ByVal s As String, ByVal n As Variant) As Variant
    ReptPst = str & rept(s, n - StringWidth(str))
End Function

Public Function ReptPre(ByVal str As String, ByVal s As String, ByVal n As Variant) As Variant
    ReptPre = rept(s, n - StringWidth(str)) & str
End Function

Public Function Format2(ByVal n As Variant, Optional ByVal width As Long = 10) As Variant

    If Len(n) <= width Then
        Format2 = n
        GoTo Escape
    End If


    If InStr(n, "E") > 0 Then
        Dim tmp: tmp = SepA(CStr(n), InStr(n, "E") - 1)
        Format2 = LeftA(tmp(0), width - 4) & tmp(1)
        GoTo Escape
    Else

        Dim arr: arr = Split(n, ".")

        If Len(arr(0)) > width Then
            Format2 = Format(arr(0), "0.0E-00")
        Else
            Format2 = LeftA(arr(0) & "." & LeftA(arr(1), width - Len(arr(0))), width - 1) & ">"
        End If

    End If

Escape:
End Function
