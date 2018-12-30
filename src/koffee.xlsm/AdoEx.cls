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

Public Enum SourceTypeEnum
    srcText = 0
    srcDBF = 1
    srcAccess = 2
    srcExcel = 3
End Enum

Private cn_ As Object
Private rs_ As Object
Private dbEngine As String
Private filePath As String

Private Sub Class_Initialize()

    Set cn_ = CreateObject("ADODB.Connection")

    ''' Microsoft Access Database Engine 2016 Redistributable
    ''' https://www.microsoft.com/en-us/download/details.aspx?id=54920
    If is64bitOS() Then
        dbEngine = "Microsoft.ACE.OLEDB.12.0"
    Else
        dbEngine = "Microsoft.Jet.OLEDB.4.0"
    End If

    filePath = ThisWorkbook.Path & "\" & ThisWorkbook.Name

End Sub

Private Sub Class_Terminate()

    On Error Resume Next
    If Not rs_ Is Nothing Then rs_.Close
    Set rs_ = Nothing

    cn_.Close
    Set cn_ = Nothing

End Sub

Public Sub Init(ByVal srcType As SourceTypeEnum, Optional ByVal fpath As String = "")

    If fpath = "" Then
        If srcType = srcAccess Then Err.Raise 76 ''' Path not found
        fpath = filePath
    Else
        fpath = fpath
    End If

    Select Case srcType

        Case 0: ''' Text

            With cn_
                .Provider = dbEngine
                .Properties("Data Source") = Fso.GetParentFolderName(fpath)
                .Properties("Extended Properties") = Join(Array("text", "HDR=yes"), ";")
                .Open
            End With

        Case 1: ''' DBF

            With cn_
                .Provider = dbEngine
                .Properties("Data Source") = Fso.GetParentFolderName(fpath)
                .Properties("Extended Properties") = Join(Array("dBASE IV", "User ID=Admin"), ";")
                .Open
            End With

        Case 2: ''' Access

            With cn_
                .Provider = dbEngine
                .Properties("Data Source") = fpath
                .Open
            End With

        Case 3: ''' Excel

            With cn_
                .Provider = dbEngine
                .Properties("Data Source") = fpath
                .Properties("Extended Properties") = Join(Array("Excel 12.0", "HDR=YES"), ";")
                .Open
            End With

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

    If isHeader Then
        Select_ = Array(Arr2DToJagArr(rs_.GetRows()), arrx.ToArray)
    Else
        Select_ = Arr2DToJagArr(rs_.GetRows())
    End If

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
