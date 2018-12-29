VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Atom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'  +--------------                                         --------------+
'  |||||||||    Koffee2 0.1.0                                            |
'  |: ^_^ :|    Koffee2 is free Library based on Ariawase.               |
'  |||||||||    The Project Page: https://github.com/CallMeKohei/Koffee2 |
'  +--------------                                         --------------+
Option Explicit

Private fun_    As Object
Private arrx_ As Object
Private addr_ As Long

Private Sub Class_Initialize()
    Set fun_ = New Func
    Set arrx_ = New ArrayEx
End Sub

Private Sub Class_Terminate()
    Set fun_ = Nothing
    Set arrx_ = Nothing
End Sub

Public Function AddFunc(ByVal fun As Func) As Atom
    Set fun_ = fun
    Set AddFunc = Me
End Function

Public Function Bind(ByVal x As Variant) As Atom
    If IsObject(x) Then arrx_.AddObj x Else arrx_.addval x
    Set Bind = Me
End Function

Public Property Let LetAddr(ByVal xAddr As Long)
    addr_ = xAddr
End Property

Public Property Get GetAddr()
    GetAddr = addr_
End Property

Public Function Pop() As Atom
    arrx_.popval
    Set Pop = Me
End Function

Public Function PopPop() As Atom
    Dim arr: arr = arrx_.ToArray
    arr(0).Pop
    Set arrx_ = Nothing
    Set arrx_ = New ArrayEx
    Dim v As Variant
    For Each v In arr
        If IsObject(v) Then arrx_.AddObj v Else arrx_.addval v
    Next v
    Set PopPop = Me
End Function

Public Sub CallByPtr(ByRef returnVal As Variant, ByVal args As Variant)
    fun_.CallByPtr returnVal, ApplyPrep(args)
End Sub

Public Sub FastApply(ByRef returnVal As Variant, ParamArray args() As Variant)

    Dim tmp As Variant
    Select Case UBound(args)
        Case Is = -1: tmp = Missing
        Case Is = 0
            If IsObject(args(0)) Then
                Set tmp = args(0)
            Else
                    tmp = args
            End If
        Case Else:  tmp = args
    End Select
    
    fun_.CallByPtr returnVal, ApplyPrep(tmp)
End Sub

Public Function Apply(ParamArray args() As Variant) As Variant
    
    Dim tmp As Variant
    Select Case UBound(args)
        Case Is = -1: tmp = Missing
        Case Is = 0
            If IsObject(args(0)) Then
                Set tmp = args(0)
            Else
                    tmp = args
            End If
        Case Else:  tmp = args
    End Select
    
    fun_.CallByPtr Apply, ApplyPrep(tmp)
    
End Function

Private Function ApplyPrep(ByVal args As Variant)
    
    Dim ps As Variant
    
    Select Case addr_
        
        Case Is = VBA.CLng(AddressOf PartialMark)
            
            Dim arrx As New ArrayEx
            PartialBindPrep arrx
            MergeArgs2 arrx, args
            ps = arrx.ToArray
            
            Dim RestPs: RestPs = Rest(ps)
            If Not IsArray(RestPs(0)) And UBound(RestPs) = 0 Then
                ApplyPrep = Array(ps(0), RestPs(0))
            Else
                ApplyPrep = Array(ps(0), RestPs)
            End If
    
        Case Is = VBA.CLng(AddressOf FoldMark), VBA.CLng(AddressOf ScanMark), VBA.CLng(AddressOf FoldRMark), VBA.CLng(AddressOf ScanRMark)
        
            ps = MergeArgs(args)
            
            If UBound(ps) = 1 Then
                ReDim Preserve ps(2)
                ps(2) = Missing
            End If
            
            ApplyPrep = ps
            
        Case Else
        
            ApplyPrep = MergeArgs(args)
    
    End Select
    
Escape:
End Function

Private Function MergeArgs(ByVal args As Variant) As Variant
    
    If IsMissing(args) Then args = Array()
    
    Dim v As Variant
    If Not IsArray(args) Then
        If IsObject(args) Then arrx_.AddObj args Else arrx_.addval args
    ElseIf Not UBound(args) = -1 Then
        For Each v In args
            If IsObject(v) Then
                arrx_.AddObj v
            ElseIf Not IsArray(v) Then
                arrx_.addval v
            Else
                If UBound(v) = 0 Then
                    arrx_.addval v
                Else
                    arrx_.addval v
                End If
            End If
        Next v
    End If
    
    MergeArgs = arrx_.ToArray
    
    If addr_ = 0 Then
        arrx_.popval
    End If
    
End Function

Private Sub PartialBindPrep(ByRef arrx As Variant)

    Dim v As Variant
    For Each v In arrx_.ToArray
        If Not IsMissing(v) Then
            If TypeName(v) = "Tuple" Then
                If IsObject(v.Item1) Then
                    arrx.AddObj v.Item1
                Else
                    Dim tmp: tmp = v.Item1
                    If IsArray(tmp) Then
                        If UBound(tmp) <> 0 Then
                            Dim v2
                            For Each v2 In tmp
                                arrx.addval v2
                            Next v2
                        Else
                            If TypeName(tmp(0)) = "Atom" Then
                                arrx.AddObj tmp(0)
                            Else
                                arrx.addval tmp(0)
                            End If
                        End If
                    Else
                        arrx.addval tmp
                    End If
                End If
            ElseIf IsObject(v) Then
                arrx.AddObj v
            Else
                arrx.addval v
            End If
        End If
    Next v
    
End Sub

Private Sub MergeArgs2(ByRef arrx As Variant, ByVal args As Variant)
    
    If IsMissing(args) Then args = Array()
    
    Dim v As Variant
    If Not IsArray(args) Then
        If IsObject(args) Then arrx.AddObj args Else arrx.addval args
    ElseIf Not UBound(args) = -1 Then
        For Each v In args
            If IsObject(v) Then
                arrx.AddObj v
            ElseIf Not IsArray(v) Then
                arrx.addval v
            Else
                If UBound(v) = 0 Then
                    arrx.addval v
                Else
                    arrx.addval v
                End If
            End If
        Next v
    End If
    
End Sub
