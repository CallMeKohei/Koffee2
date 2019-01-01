VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' --------------------------------------------------------
'''  FILE    : AdoEx.cls
'''  AUTHOR  : callmekohei <callmekohei at gmail.com>
'''  License : MIT license
''' --------------------------------------------------------
Option Explicit

''' Dependencies
'''
'''     Select_
'''         ArrayEx(Ariawase)
'''         Arr2DToJagArr(Ariawase)
'''     pos
'''         AddSheet(KoffeeExcel)

Public Enum CursorTypeEnum
    adOpenUnspecified = -1
    adOpenForwardOnly = 0
    adOpenKeyset = 1
    adOpenDynamic = 2
    adOpenStatic = 3
End Enum

Public Enum LockTypeEnum
    adLockUnspecified = -1
    adLockReadOnly = 1
    adLockPessimistic = 2
    adLockOptimistic = 3
    adLockBatchOptimistic = 3
End Enum

Public Enum CommandTypeEnum
    adCmdUnspecified = -1
    adCmdText = 1
    adCmdTable = 2
    adCmdStoredProc = 4
    adCmdUnknown = 8
    adAsyncExecute = 16
    adAsyncFetch = 32
    adCmdFile = 256
    adCmdTableDirect = 512
End Enum

Public Enum dbTypeEnum
    dbText = 0
    dbDBF = 1
    dbAccess = 2
    dbExcel = 3
End Enum

Private cn_ As Object
Private rs_ As Object
Private dbEngine_ As String
Private filePath_ As String

Private Sub Class_Initialize()

    Set cn_ = CreateObject("ADODB.Connection")

    ''' Microsoft Access Database Engine 2016 Redistributable
    ''' https://www.microsoft.com/en-us/download/details.aspx?id=54920
    If is64bitOS() Then
        dbEngine_ = "Microsoft.ACE.OLEDB.12.0"
    Else
        dbEngine_ = "Microsoft.Jet.OLEDB.4.0"
    End If

    filePath_ = ThisWorkbook.Path & "\" & ThisWorkbook.Name

End Sub

Private Sub Class_Terminate()

    On Error Resume Next
    If Not rs_ Is Nothing Then rs_.Close
    Set rs_ = Nothing

    cn_.Close
    Set cn_ = Nothing

End Sub

Public Sub Init(ByVal dbType As dbTypeEnum, Optional ByVal fpath As String = "", Optional isHeader As Boolean = True)

    If fpath = "" Then
        If dbType = dbAccess Then Err.Raise 76 ''' Path not found
        fpath = filePath_
    Else
        fpath = fpath
    End If

    Dim strHDR As String: strHDR = IIf(isHeader, "HDR=YES", "HDR=NO")

    Select Case dbType

        Case 0: ''' Text

            cn_.Provider = dbEngine_
            cn_.Properties("Data Source") = Fso.GetParentFolderName(fpath)
            cn_.Properties("Extended Properties") = Join(Array("TEXT", strHDR), ";")
            cn_.Open

        Case 1: ''' DBF

            cn_.Provider = dbEngine_
            cn_.Properties("Data Source") = Fso.GetParentFolderName(fpath)
            cn_.Properties("Extended Properties") = Join(Array("dBASE IV", "User ID=Admin"), ";")
            cn_.Open

        Case 2: ''' Access

            cn_.Provider = dbEngine_
            cn_.Properties("Data Source") = fpath
            cn_.Open

        Case 3: ''' Excel

            Dim dbTypeExcel As String
            Select Case Fso.GetExtensionName(fpath)
                Case "xls" ''' Excel5.0/95 Workbook(*.xls) or Excel97-2003 WorkBook(*.xls)
                    If InStr(Fso.getfile(fpath).Type, 95) Then
                        dbTypeExcel = "Excel 5.0"
                    Else
                        dbTypeExcel = "Excel 8.0"
                    End If
                Case "xlsb" ''' Excel binary Workbook(*.xlsb)
                    dbTypeExcel = "Excel 12.0"
                Case "xlsm" ''' Excel Macro-Enabled Wrokbook(*.xlsm)
                    dbTypeExcel = "Excel 12.0 Macro"
                Case "xlsx" ''' Excel Workbook(*.xlsx)
                    dbTypeExcel = "Excel 12.0 Xml"
                Case Else
                    dbTypeExcel = ""
            End Select

            cn_.Provider = dbEngine_
            cn_.Properties("Data Source") = fpath
            cn_.Properties("Extended Properties") = Join(Array(dbTypeExcel, strHDR), ";")
            cn_.Open

    End Select

End Sub

Public Function Select_(ByVal sql As String, _
    Optional ByVal isHeader As Boolean = False, _
    Optional ByVal propCursorType As CursorTypeEnum = adOpenStatic, _
    Optional ByVal propLockType As LockTypeEnum = adLockReadOnly, _
    Optional ByVal propCommandType As CommandTypeEnum = adCmdText _
    )

    Set rs_ = CreateObject("ADODB.Recordset")
    rs_.Open Source:=sql, ActiveConnection:=cn_, CursorType:=propCursorType, LockType:=propLockType, Options:=propCommandType

    ''' Should be defaulte date with header?
    Dim v As Variant, arrx As New ArrayEx
    For Each v In rs_.Fields
        arrx.AddVal v.Name
    Next v

    Dim arr As Variant: arr = Arr2DToJagArr(ArrTranspose(rs_.GetRows()))

    If isHeader Then
        Select_ = Array(arr, arrx.ToArray)
    Else
        Select_ = arr
    End If

    On Error Resume Next
    If Not rs_ Is Nothing Then rs_.Close
    Set rs_ = Nothing

End Function

Private Function ArrTranspose(ByVal arr As Variant) As Variant

    Dim ub1 As Long: ub1 = UBound(arr, 2)
    Dim ub2 As Long: ub2 = UBound(arr, 1)

    Dim tmp() As Variant: ReDim tmp(ub1, ub2)

    Dim ix1 As Long, ix2 As Long
    For ix1 = 0 To ub1
        For ix2 = 0 To ub2
            tmp(ix1, ix2) = arr(ix2, ix1)
        Next ix2
    Next ix1

    ArrTranspose = tmp

End Function

Public Function Execute(ByVal sql As String) As Long
    Dim rcdNumbers As Long
    cn_.Execute sql, ARecNum
    Execute = rcdNumbers
End Function

Public Sub BeginTrans()
    cn.BeginTrans
End Sub

Public Sub CommitTrans()
    cn_.CommitTrans
End Sub

Public Sub RollbackTrans()
    cn.RollbackTrans
End Sub

Private Function is64bitOS() As Boolean

  Dim v As Variant
  For Each v In Core.Wmi.connectServer.execQuery("Select * From Win32_OperatingSystem")

    If InStr(v.OSArchitecture, "64") Then
        is64bitOS = True
        Exit Function
    End If

  Next

End Function

Public Sub pos(ByVal sql As String, Optional ByRef rng As Range = Nothing)

    Set rs_ = CreateObject("ADODB.Recordset")
    rs_.Open Source:=sql, ActiveConnection:=cn_, CursorType:=adOpenStatic, LockType:=adLockReadOnly, Options:=adCmdText

    If rs_.EOF Then Debug.Print "No Date!: " & Now(): Exit Sub

    If TypeName(rng) = "Nothing" Then
        AddSheet "test"
        Set rng = ThisWorkbook.Worksheets("test").Range("C5")
    End If

    rng.CopyFromRecordset rs_

End Sub
