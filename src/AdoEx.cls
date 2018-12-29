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
    adCsv = 0
    adExcel = 1
    adSQLite = 2
End Enum

Const PROVIDER01    As String = "PROVIDER=MSDASQL;Extended Properties="""
Const DRIVER01      As String = "DRIVER={Microsoft Text Driver (*.txt; *.csv)};"
Const PATH01        As String = "DBQ="

Const PROVIDER02    As String = "PROVIDER=MSDASQL;"
Const DRIVER02      As String = "DRIVER={Microsoft Excel Driver (*.xls)};"
Const PATH02        As String = "DBQ="

Const PROVIDER03    As String = "PROVIDER=MSDASQL;charset=cp932;"
Const DRIVER03      As String = "Driver=SQLite3 ODBC Driver;"
Const PATH03        As String = ""

Private cn_ As Object

Private Sub Class_Initialize()
    Set cn_ = CreateObject("ADODB.Connection")
End Sub

Private Sub Class_Terminate()
    cn_.Close: Set cn_ = Nothing
End Sub

Public Sub Init(ByVal propSourceType As SourceTypeEnum, ByVal fPath As String)
    Dim str As String: str = ""
    Select Case propSourceType
        Case Is = 0: str = PROVIDER01 & DRIVER01 & PATH01 & Fso.GetParentFolderName(fPath) & "\" & """"
        Case Is = 1: str = PROVIDER02 & DRIVER02 & PATH02 & fPath
        Case Is = 2: str = PROVIDER03 & DRIVER03 & PATH03 & fPath
    End Select
    cn_.ConnectionString = str
    cn_.Open
End Sub

Public Function CreateAdoRS(ByVal sql As String, _
    Optional ByVal propCursorType As CursorTypeEnum = adOpenStatic, _
    Optional ByVal propLockType As LockTypeEnum = adLockReadOnly, _
    Optional ByVal propCommandType As CommandTypeEnum = adCmdText _
    ) As Object

    Set CreateAdoRS = CreateObject("ADODB.Recordset")
    With CreateAdoRS
        .Open _
            Source:=sql, _
            ActiveConnection:=cn_, _
            CursorType:=propCursorType, _
            LockType:=propLockType, _
            Options:=propCommandType
    End With

End Function

Public Function JagArrAdoRsHeader(ByVal sql As String) As Variant
    Dim v As Variant, arrx As New ArrayEx
    For Each v In CreateAdoRS(sql).Fields
        arrx.addval v.Name
    Next v
    JagArrAdoRsHeader = arrx.ToArray
End Function

Public Function JagArrAdoRS(ByVal sql As String) As Variant
    JagArrAdoRS = Arr2DToJagArr(CreateAdoRS(sql).GetRows())
End Function

Public Sub pos(ByVal sql As String, Optional ByVal rng As Range = Nothing)

    If CreateAdoRS(sql).EOF Then Debug.Print "No Date!: " & Now(): Exit Sub

    If TypeName(rng) = "Nothing" Then
        AddSheet "test"
        Set rng = ThisWorkbook.Worksheets("test").Range("C5")
    End If

    rng.CopyFromRecordset CreateAdoRS(sql)
End Sub
