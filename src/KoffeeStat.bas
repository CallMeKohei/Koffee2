Attribute VB_Name = "KoffeeStat"
'  +--------------                                         --------------+
'  |||||||||    Koffee2 0.1.0                                            |
'  |: ^_^ :|    Koffee2 is free Library based on Ariawase.               |
'  |||||||||    The Project Page: https://github.com/CallMeKohei/Koffee2 |
'  +--------------                                         --------------+
Option Explicit

'Basic Function
Public Function LogN(ByVal xs As Variant) As Variant

    If Not IsArray(xs) Then
        LogN = Log(xs)
    Else
        Dim v As Variant, arrx As New ArrayEx
        For Each v In xs
            arrx.addval Log(v)
        Next v
        LogN = arrx.ToArray
    End If

End Function

Public Function Sqr2(ByVal xs As Variant) As Variant

    If Not IsArray(xs) Then
        Sqr2 = Sqr(xs)
    Else
        Dim v As Variant, arrx As New ArrayEx
        For Each v In xs
            arrx.addval Sqr(v)
        Next v
        Sqr2 = arrx.ToArray
    End If

End Function

Public Function PowerN(ByVal xs As Variant, ByVal nth As Variant) As Variant

    Dim wf As WorksheetFunction: Set wf = Application.WorksheetFunction

    If Not IsArray(xs) Then
        PowerN = wf.Power(xs, nth)
    Else
        Dim v As Variant, arrx As New ArrayEx
        For Each v In xs
            arrx.addval wf.Power(v, nth)
        Next v
        PowerN = arrx.ToArray
    End If

End Function

Public Function PowerScan(ByVal New_x As Variant, ByVal nth As Long) As Variant
    Dim wf As WorksheetFunction: Set wf = Application.WorksheetFunction
    PowerScan = Arr2DToJagArr(wf.Transpose(Application.Power(wf.Transpose(New_x), ArrRange(1, nth))))
End Function

Public Function LinEst2(ByVal known_y As Variant, ByVal known_xs As Variant) As Variant
    LinEst2 = Base01(Arr2DToJagArr(Application.WorksheetFunction.LinEst(known_y, known_xs, True, True)))
End Function


'Trend Line
Public Function TLTrd(ByVal New_xs As Variant, ByVal known_y, ByVal known_xs) As Variant

    If IsJagArr(New_xs) Then

        Dim ub As Long
        If IsJagArr(known_xs) Then
            ub = UBound(known_xs(0))
        Else
            ub = UBound(known_xs)
        End If

        Dim i As Long, arrx As New ArrayEx, rcd:  rcd = ToRecord(New_xs, False)
        For i = 0 To ub
            arrx.addval WorksheetFunction.Trend(known_y, known_xs, PartialPackage(rcd(i)))(1)
        Next i

        TLTrd = arrx.ToArray

    Else
        TLTrd = ArrFlatten(Base01(WorksheetFunction.Trend(known_y, known_xs, PartialPackage(New_xs))))
    End If

End Function

Private Function PartialPackage(ByVal arr As Variant) As Variant
    Dim v, arrx As New ArrayEx
    For Each v In arr
        arrx.addval Array(v)
    Next v
    PartialPackage = arrx.ToArray
End Function

Public Function TLTrdU(ByVal New_x As Variant, ByVal known_y As Variant, ByVal known_xs As Variant, Optional ByVal dc = 0.95)

    If IsArray(New_x) Then
        If IsJagArr(New_x) And UBound(New_x) = 0 Then New_x = ArrFlatten(New_x)
    End If
    If IsJagArr(known_y) And UBound(known_y) = 0 Then known_y = ArrFlatten(known_y)

    TLTrdU = TLTrd(New_x, known_y, known_xs)(0) + Tvalue(UBound(known_y), UBound(New_x), dc) * SEP(TLTrd(known_xs, known_y, known_xs), known_y, known_xs, New_x)(0)

End Function

Public Function TLTrdL(ByVal New_x As Variant, ByVal known_y As Variant, ByVal known_xs As Variant, Optional ByVal dc = 0.95)

    If IsArray(New_x) Then
        If IsJagArr(New_x) And UBound(New_x) = 0 Then New_x = ArrFlatten(New_x)
    End If
    If IsJagArr(known_y) And UBound(known_y) = 0 Then known_y = ArrFlatten(known_y)

    TLTrdL = TLTrd(New_x, known_y, known_xs)(0) - Tvalue(UBound(known_y), UBound(New_x), dc) * SEP(TLTrd(known_xs, known_y, known_xs), known_y, known_xs, New_x)(0)

End Function

Public Function TLLin(ByVal New_x As Variant, ByVal known_y As Variant, ByVal known_xs As Variant) As Variant

    If Not IsArray(New_x) And IsNumeric(New_x) Then New_x = Array(New_x)
    If Not IsArray(New_x) Then Err.Raise 13

    If IsJagArr(New_x) And UBound(New_x) = 0 Then New_x = ArrFlatten(New_x)

    Dim v As Variant, arrx As New ArrayEx
    For Each v In New_x
        arrx.addval Application.WorksheetFunction.Forecast(v, known_y, known_xs)
    Next v

    TLLin = arrx.ToArray

End Function

Public Function TLLinU(ByVal New_x As Variant, ByVal known_y As Variant, ByVal known_xs As Variant)

    If IsArray(New_x) Then
        If IsJagArr(New_x) And UBound(New_x) = 0 Then New_x = ArrFlatten(New_x)
    End If
    If IsJagArr(known_y) And UBound(known_y) = 0 Then known_y = ArrFlatten(known_y)
    If IsJagArr(known_xs) And UBound(known_xs) = 0 Then known_xs = ArrFlatten(known_xs)

    If ArrEquals(IIf(IsArray(New_x), New_x, Array(New_x)), known_xs) Then
        TLLinU = ArrPlus(TLLin(New_x, known_y, known_xs), SEP(New_x, known_y, known_xs))
    Else
        TLLinU = ArrPlus(TLLin(New_x, known_y, known_xs), SEP(TLLin(known_xs, known_y, known_xs), known_y, known_xs, New_x))
    End If

End Function

Public Function TLLinL(ByVal New_x As Variant, ByVal known_y As Variant, ByVal known_xs As Variant)

    If IsArray(New_x) Then
        If IsJagArr(New_x) And UBound(New_x) = 0 Then New_x = ArrFlatten(New_x)
    End If
    If IsJagArr(known_y) And UBound(known_y) = 0 Then known_y = ArrFlatten(known_y)
    If IsJagArr(known_xs) And UBound(known_xs) = 0 Then known_xs = ArrFlatten(known_xs)

    If ArrEquals(IIf(IsArray(New_x), New_x, Array(New_x)), known_xs) Then
        TLLinL = ArrMinus(TLLin(New_x, known_y, known_xs), SEP(New_x, known_y, known_xs))
    Else
        TLLinL = ArrMinus(TLLin(New_x, known_y, known_xs), SEP(TLLin(known_xs, known_y, known_xs), known_y, known_xs, New_x))
    End If

End Function

Public Function TLExp(ByVal New_x As Variant, ByVal known_y As Variant, ByVal known_xs As Variant) As Variant
    If Not IsArray(New_x) And IsNumeric(New_x) Then New_x = Array(New_x)
    If Not IsArray(New_x) Then Err.Raise 13
    TLExp = ArrFlatten(Base01(Application.WorksheetFunction.Growth(known_y, known_xs, New_x)))
End Function

Public Function TLExpU(ByVal New_x As Variant, ByVal known_y As Variant, ByVal known_xs As Variant)

    If IsArray(New_x) Then
        If IsJagArr(New_x) And UBound(New_x) = 0 Then New_x = ArrFlatten(New_x)
    End If
    If IsJagArr(known_y) And UBound(known_y) = 0 Then known_y = ArrFlatten(known_y)
    If IsJagArr(known_xs) And UBound(known_xs) = 0 Then known_xs = ArrFlatten(known_xs)

    If ArrEquals(IIf(IsArray(New_x), New_x, Array(New_x)), known_xs) Then
        TLExpU = ArrPlus(TLExp(New_x, known_y, known_xs), SEP(New_x, known_y, known_xs))
    Else
        TLExpU = ArrPlus(TLExp(New_x, known_y, known_xs), SEP(TLExp(known_xs, known_y, known_xs), known_y, known_xs, New_x))
    End If

End Function

Public Function TLExpL(ByVal New_x As Variant, ByVal known_y As Variant, ByVal known_xs As Variant)

    If IsArray(New_x) Then
        If IsJagArr(New_x) And UBound(New_x) = 0 Then New_x = ArrFlatten(New_x)
    End If
    If IsJagArr(known_y) And UBound(known_y) = 0 Then known_y = ArrFlatten(known_y)
    If IsJagArr(known_xs) And UBound(known_xs) = 0 Then known_xs = ArrFlatten(known_xs)

    If ArrEquals(IIf(IsArray(New_x), New_x, Array(New_x)), known_xs) Then
        TLExpL = ArrMinus(TLExp(New_x, known_y, known_xs), SEP(New_x, known_y, known_xs))
    Else
        TLExpL = ArrMinus(TLExp(New_x, known_y, known_xs), SEP(TLExp(known_xs, known_y, known_xs), known_y, known_xs, New_x))
    End If

End Function

Public Function TLLog(ByVal New_x As Variant, ByVal known_y As Variant, ByVal known_xs As Variant) As Variant

    If Not IsArray(New_x) And IsNumeric(New_x) Then New_x = Array(New_x)
    If Not IsArray(New_x) Then Err.Raise 13

    Dim v As Variant, arrx As New ArrayEx
    For Each v In New_x
        arrx.addval Application.WorksheetFunction.Forecast(Log(v), known_y, LogN(known_xs))
    Next v

    TLLog = arrx.ToArray

End Function

Public Function TLLogU(ByVal New_x As Variant, ByVal known_y As Variant, ByVal known_xs As Variant)

    If IsArray(New_x) Then
        If IsJagArr(New_x) And UBound(New_x) = 0 Then New_x = ArrFlatten(New_x)
    End If
    If IsJagArr(known_y) And UBound(known_y) = 0 Then known_y = ArrFlatten(known_y)
    If IsJagArr(known_xs) And UBound(known_xs) = 0 Then known_xs = ArrFlatten(known_xs)

    If ArrEquals(IIf(IsArray(New_x), New_x, Array(New_x)), known_xs) Then
        TLLogU = ArrPlus(TLLog(New_x, known_y, known_xs), SEP(New_x, known_y, known_xs))
    Else
        TLLogU = ArrPlus(TLLog(New_x, known_y, known_xs), SEP(TLLog(known_xs, known_y, known_xs), known_y, known_xs, New_x))
    End If

End Function

Public Function TLLogL(ByVal New_x As Variant, ByVal known_y As Variant, ByVal known_xs As Variant)

    If IsArray(New_x) Then
        If IsJagArr(New_x) And UBound(New_x) = 0 Then New_x = ArrFlatten(New_x)
    End If
    If IsJagArr(known_y) And UBound(known_y) = 0 Then known_y = ArrFlatten(known_y)
    If IsJagArr(known_xs) And UBound(known_xs) = 0 Then known_xs = ArrFlatten(known_xs)

    If ArrEquals(IIf(IsArray(New_x), New_x, Array(New_x)), known_xs) Then
        TLLogL = ArrMinus(TLLog(New_x, known_y, known_xs), SEP(New_x, known_y, known_xs))
    Else
        TLLogL = ArrMinus(TLLog(New_x, known_y, known_xs), SEP(TLLog(known_xs, known_y, known_xs), known_y, known_xs, New_x))
    End If

End Function

Public Function TLPwr(ByVal New_x As Variant, ByVal known_y As Variant, ByVal known_xs As Variant) As Variant

    If Not IsArray(New_x) And IsNumeric(New_x) Then New_x = Array(New_x)
    If Not IsArray(New_x) Then Err.Raise 13

    Dim v As Variant, arrx As New ArrayEx
    For Each v In New_x
        arrx.addval Exp(Application.WorksheetFunction.Forecast(Log(v), LogN(known_y), LogN(known_xs)))
    Next v

    TLPwr = arrx.ToArray

End Function

Public Function TLPwrU(ByVal New_x As Variant, ByVal known_y As Variant, ByVal known_xs As Variant)

    If IsArray(New_x) Then
        If IsJagArr(New_x) And UBound(New_x) = 0 Then New_x = ArrFlatten(New_x)
    End If
    If IsJagArr(known_y) And UBound(known_y) = 0 Then known_y = ArrFlatten(known_y)
    If IsJagArr(known_xs) And UBound(known_xs) = 0 Then known_xs = ArrFlatten(known_xs)

    If ArrEquals(IIf(IsArray(New_x), New_x, Array(New_x)), known_xs) Then
        TLPwrU = ArrPlus(TLPwr(New_x, known_y, known_xs), SEP(New_x, known_y, known_xs))
    Else
        TLPwrU = ArrPlus(TLPwr(New_x, known_y, known_xs), SEP(TLPwr(known_xs, known_y, known_xs), known_y, known_xs, New_x))
    End If

End Function

Public Function TLPwrL(ByVal New_x As Variant, ByVal known_y As Variant, ByVal known_xs As Variant)

    If IsArray(New_x) Then
        If IsJagArr(New_x) And UBound(New_x) = 0 Then New_x = ArrFlatten(New_x)
    End If
    If IsJagArr(known_y) And UBound(known_y) = 0 Then known_y = ArrFlatten(known_y)
    If IsJagArr(known_xs) And UBound(known_xs) = 0 Then known_xs = ArrFlatten(known_xs)

    If ArrEquals(IIf(IsArray(New_x), New_x, Array(New_x)), known_xs) Then
        TLPwrL = ArrMinus(TLPwr(New_x, known_y, known_xs), SEP(New_x, known_y, known_xs))
    Else
        TLPwrL = ArrMinus(TLPwr(New_x, known_y, known_xs), SEP(TLPwr(known_xs, known_y, known_xs), known_y, known_xs, New_x))
    End If

End Function


Public Function TLPly(ByVal New_x As Variant, ByVal known_y As Variant, ByVal known_xs As Variant, Optional ByVal n As Long = 2) As Variant

    If Not IsArray(New_x) And IsNumeric(New_x) Then New_x = Array(New_x)
    If Not IsArray(New_x) Then Err.Raise 13

    Dim v As Variant, arrx As New ArrayEx
    For Each v In New_x
        arrx.addval Application.WorksheetFunction.Trend(known_y, PowerScan(known_xs, n), PowerScan(Array(v), n))
    Next v

    TLPly = ArrFlatten(arrx.ToArray)

End Function

Public Function TLPlyU(ByVal New_x As Variant, ByVal known_y As Variant, ByVal known_xs As Variant, Optional ByVal n As Long = 2)

    If IsArray(New_x) Then
        If IsJagArr(New_x) And UBound(New_x) = 0 Then New_x = ArrFlatten(New_x)
    End If
    If IsJagArr(known_y) And UBound(known_y) = 0 Then known_y = ArrFlatten(known_y)
    If IsJagArr(known_xs) And UBound(known_xs) = 0 Then known_xs = ArrFlatten(known_xs)

    If ArrEquals(IIf(IsArray(New_x), New_x, Array(New_x)), known_xs) Then
        TLPlyU = ArrPlus(TLPly(New_x, known_y, known_xs), SEP(New_x, known_y, known_xs))
    Else
        TLPlyU = ArrPlus(TLPly(New_x, known_y, known_xs), SEP(TLPly(known_xs, known_y, known_xs, n), known_y, known_xs, New_x))
    End If

End Function

Public Function TLPlyL(ByVal New_x As Variant, ByVal known_y As Variant, ByVal known_xs As Variant, Optional ByVal n As Long = 2)

    If IsArray(New_x) Then
        If IsJagArr(New_x) And UBound(New_x) = 0 Then New_x = ArrFlatten(New_x)
    End If
    If IsJagArr(known_y) And UBound(known_y) = 0 Then known_y = ArrFlatten(known_y)
    If IsJagArr(known_xs) And UBound(known_xs) = 0 Then known_xs = ArrFlatten(known_xs)

    If ArrEquals(IIf(IsArray(New_x), New_x, Array(New_x)), known_xs) Then
        TLPlyL = ArrMinus(TLPly(New_x, known_y, known_xs), SEP(New_x, known_y, known_xs))
    Else
        TLPlyL = ArrMinus(TLPly(New_x, known_y, known_xs), SEP(TLPly(known_xs, known_y, known_xs, n), known_y, known_xs, New_x))
    End If

End Function

Public Function R2Values(ByVal known_y As Variant, ByVal known_x As Variant) As Variant

    Dim wf As WorksheetFunction: Set wf = Application.WorksheetFunction

    If IsJagArr(known_x) Then
        R2Values = LinEst2(known_y, known_x)(2)(0)
    Else
        Dim Lin As Variant: Lin = wf.RSq(known_y, known_x)
        Dim Exp As Variant: Exp = wf.RSq(LogN(known_y), known_x)
        Dim Log As Variant: Log = wf.RSq(known_y, LogN(known_x))
        Dim Pwr As Variant: Pwr = wf.RSq(LogN(known_y), LogN(known_x))
        Dim ply2 As Variant: ply2 = LinEst2(known_y, PowerScan(known_x, 2))(2)(0)
        Dim ply3 As Variant: ply3 = LinEst2(known_y, PowerScan(known_x, 3))(2)(0)
        R2Values = Array(Lin, Exp, Log, Pwr, ply2, ply3)
    End If



End Function


'Standard Error of Predicts

Public Function SEP(ByVal Pred_y, ByVal known_y, ByVal known_xs, Optional ByVal Setter_xs, Optional ByVal SampleCnt As Long, Optional ByVal Known_xsCnt As Long) As Variant

    Dim flg As Boolean
    If IsMissing(Setter_xs) Then
        Setter_xs = known_xs
        flg = True
    End If

    If SampleCnt = 0 Then
        If IsJagArr(known_y) Then
            SampleCnt = ArrLen(known_y(0))
        Else
            SampleCnt = ArrLen(known_y)
        End If
    End If

    If Known_xsCnt = 0 Then
        If IsJagArr(known_xs) Then
            Known_xsCnt = ArrLen(known_xs)
        Else
            Known_xsCnt = 1
        End If
    End If

    Dim v, v2, arrx As New ArrayEx
    If flg Or IsJagArr(Setter_xs) Then
        For Each v In Setter_xs
            arrx.addval Sqr(Residualv(Pred_y, known_y, SampleCnt, Known_xsCnt) * (1 + 1 / SampleCnt + MahalanobisDist(v, known_xs)(1) / (SampleCnt - 1)))
        Next v
        SEP = arrx.ToArray
    Else
        SEP = Array(Sqr(Residualv(Pred_y, known_y, SampleCnt, Known_xsCnt) * (1 + 1 / SampleCnt + MahalanobisDist(Setter_xs, known_xs)(1) / (SampleCnt - 1))))
    End If

End Function

Public Function Residual(ByVal Pred_y, ByVal xs) As Variant

    Dim wf As WorksheetFunction: Set wf = Application.WorksheetFunction

    If IsArray(Pred_y) Then
        Dim i As Long, arrx As New ArrayEx
        For i = 0 To UBound(xs)
            arrx.addval Pred_y(i) - wf.Average(xs(i))
        Next i
        Residual = arrx.ToArray
    Else
        Residual = Pred_y - wf.Average(xs)
    End If

End Function

Public Function Residualv(ByVal Pred_y, ByVal known_y, ByVal samplCnt As Long, ByVal Known_xsCnt As Long) As Variant
    Residualv = WorksheetFunction.SumXMY2(Pred_y, known_y) / DegreeOfFree(samplCnt, Known_xsCnt)
End Function

Public Function DegreeOfFree(ByVal SampleCnt As Long, ByVal Known_xsCnt As Long) As Long
    DegreeOfFree = SampleCnt - Known_xsCnt - 1
End Function

Public Function Tvalue(ByVal SampleCnt As Long, ByVal Known_xsCnt As Long, Optional ByVal dc = 0.95) As Variant
    Tvalue = WorksheetFunction.TInv(1 - dc, DegreeOfFree(SampleCnt, Known_xsCnt))
End Function

Public Function Zvalue(Optional ByVal dc = 0.95) As Variant
    Zvalue = WorksheetFunction.NormSInv((1 - dc) / 2 + dc)
End Function

Public Function MahalanobisDist(ByVal PredVals As Variant, ByVal xs As Variant) As Variant
    Dim wf As WorksheetFunction: Set wf = Application.WorksheetFunction
    MahalanobisDist = wf.MMult(wf.MMult(Residual(PredVals, xs), wf.MInverse(CovarMatrix(xs))), wf.Transpose(Residual(PredVals, xs)))
End Function

Public Function CovarMatrix(ByVal JagArr As Variant) As Variant

    If Not IsJagArr(JagArr) Then JagArr = Array(JagArr)

    Dim wf As WorksheetFunction: Set wf = Application.WorksheetFunction

    Dim v, v1, arrx As New ArrayEx, tmp As New ArrayEx
    For Each v In JagArr
        For Each v1 In JagArr
            tmp.addval wf.Covar(v, v1) * wf.Count(v) / (wf.Count(v) - 1)
        Next v1
        arrx.addval tmp.ToArray
        Set tmp = Nothing
    Next v

    CovarMatrix = arrx.ToArray

End Function

Public Function TDist(ByVal arr, ByVal DegreeOfFreedom, ByVal Tails) As Variant
    Dim v, arrx As New ArrayEx
    For Each v In arr
        arrx.addval WorksheetFunction.TDist(Abs(v), DegreeOfFreedom, 2)
    Next v
    TDist = arrx.ToArray
End Function

Public Function Correl(ByVal arr As Variant) As Variant

    Dim v, v2, tmp As New ArrayEx, arrx As New ArrayEx

    For Each v In arr
        For Each v2 In arr
            tmp.addval Application.WorksheetFunction.Correl(v, v2)
        Next v2
        arrx.addval tmp.ToArray
        Set tmp = Nothing
    Next v

    Correl = arrx.ToArray

End Function


'About ABC chart data
Public Function TplABC(ByVal ArrName As Variant, ByVal ArrVal As Variant) As Tuple

    Dim Order:       Order = ArrRange(1, ArrLen(ArrName))
    Dim Rate:        Rate = DivideByTotal(ArrVal, Total(ArrVal))
    Dim StackRate:   StackRate = StepTotal(Rate)(0)
    Dim Rank:        Rank = ABCMark(StackRate)

    Set TplABC = Init(New Tuple, Order, Rate, StackRate, Rank)
End Function

Private Function Total(ByVal arr As Variant)
    Dim v, tmp As Currency
    For Each v In arr
        tmp = tmp + v
    Next v
    Total = tmp
End Function

Private Function DivideByTotal(ByVal arr As Variant, ByVal Total As Variant) As Variant
    Dim v, arrx As New ArrayEx
    For Each v In arr
        arrx.addval ARound((v / Total) * 100, 2)
    Next v
    DivideByTotal = arrx.ToArray
End Function

Private Function ABCMark(ByVal arr As Variant) As Variant
    Dim v, arrx As New ArrayEx
    For Each v In arr
        Select Case v
            Case Is < 70: arrx.addval "A"
            Case Is < 90: arrx.addval "B"
            Case Else:    arrx.addval "C"
        End Select
    Next v
    ABCMark = arrx.ToArray
End Function

'About Regression_Overview
Public Sub Regression_Overview(ByVal known_y, ByVal known_x_headerIncluded)

    RegressionStatistics known_y, known_x_headerIncluded

    Debug.Print vbNewLine

    ANOVA known_y, known_x_headerIncluded

    Debug.Print vbNewLine

    LinEstChart known_y, known_x_headerIncluded

End Sub

Private Sub RegressionStatistics(ByVal known_y, ByVal known_x_headerIncluded)

    Dim SampleCnt As Long: SampleCnt = ArrLen(known_y(0))
    Dim xCnt As Long:      xCnt = ArrLen(known_x_headerIncluded(0))

    Dim MultipleR:       MultipleR = (WorksheetFunction.Correl(TLTrd(known_x_headerIncluded(0), known_y(0), known_x_headerIncluded(0)), known_y(0)))
    Dim RSquare:         RSquare = LinEst2(known_y, known_x_headerIncluded(0))(2)(0)
    Dim AdjustedRSquare: AdjustedRSquare = 1 - (SampleCnt - 1) / (SampleCnt - xCnt - 1) * (1 - LinEst2(known_y, known_x_headerIncluded(0))(2)(0))
    Dim StandardError:   StandardError = LinEst2(known_y, known_x_headerIncluded(0))(2)(1)
    Dim Observations:    Observations = SampleCnt

    DP Array(MultipleR, RSquare, AdjustedRSquare, StandardError, Observations) _
        , Array("MultR2", "R2", "adjR2", "StdErr", "Obs") _
        , , "(Regression Statistics)" & vbNewLine & "----------------------"

End Sub

Private Sub ANOVA(ByVal known_y, ByVal known_x_headerIncluded)

    Dim lEst: lEst = LinEst2(known_y, known_x_headerIncluded(0))

    Dim Regdf:   Regdf = ArrLen(known_x_headerIncluded(0))
    Dim Resdf:   Resdf = lEst(3)(1)
    Dim totaldf: totaldf = Regdf + Resdf

    Dim RegSS:   RegSS = lEst(4)(0)
    Dim ResSS:   ResSS = lEst(4)(1)
    Dim TotalSS: TotalSS = RegSS + ResSS

    Dim RegMS:   RegMS = ""
    Dim ResMS:   ResMS = Residualv(TLTrd(known_x_headerIncluded(0), known_y, known_x_headerIncluded(0)), known_y, ArrLen(known_y(0)), ArrLen(known_x_headerIncluded(0)))

    Dim RegF:    RegF = lEst(3)(0)
    Dim SigF:    SigF = WorksheetFunction.FDist(RegF, Regdf, Resdf)

    DP Array(Array(Regdf, Resdf, totaldf), Array(RegSS, ResSS, TotalSS), _
             Array(RegMS, ResMS, "Error"), Array(RegF, "Error", "Error"), Array(SigF, "Error", "Error")) _
             , Array("Regression", "Residual", "Total") _
             , Array("df", "SS", "MS", "F", "SigF") _
             , "(ANOVA)", 7

End Sub

Private Sub LinEstChart(ByVal known_y, ByVal known_x_headerIncluded)

    Dim lEst: lEst = LinEst2(known_y, known_x_headerIncluded(0))
    Dim tStat: tStat = ArrDivide(ArrExplode(lEst(0)), ArrExplode(lEst(1)))
    Dim Pvalue: Pvalue = TDist(tStat, lEst(3)(1), 2)

    DP Array(lEst(0), lEst(1), tStat, Pvalue) _
    , ArrExplode(Array(known_x_headerIncluded(1), "Intercept")) _
    , Array("Coef", "StdErr", "t Stat", "P-value") _
    , , 8

    Debug.Print vbNewLine


    Dim i As Long, Lower95 As New ArrayEx, Upper95 As New ArrayEx, Lower99 As New ArrayEx, Upper99 As New ArrayEx
    For i = 0 To UBound(lEst(0))
        Lower95.addval lEst(0)(i) - (lEst(1)(i) * WorksheetFunction.TInv(1 - 0.95, lEst(3)(1)))
        Upper95.addval lEst(0)(i) + (lEst(1)(i) * WorksheetFunction.TInv(1 - 0.95, lEst(3)(1)))
        Lower99.addval lEst(0)(i) - (lEst(1)(i) * WorksheetFunction.TInv(1 - 0.99, lEst(3)(1)))
        Upper99.addval lEst(0)(i) + (lEst(1)(i) * WorksheetFunction.TInv(1 - 0.99, lEst(3)(1)))
    Next i

    DP Array(Lower95.ToArray, Upper95.ToArray, Lower99.ToArray, Upper99.ToArray) _
    , ArrExplode(Array(known_x_headerIncluded(1), "Intercept")) _
    , Array("Lower95", "Upper95", "Lower99", "Upper99") _
    , , 8

End Sub
