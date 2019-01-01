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
'''     Pos
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

Private Con_  As Object
Private Rcds_ As Object
Private Wmi_  As Object
Private dbEngine_ As String
Private filePath_ As String

Private Sub Class_Initialize()

    Set Con_ = CreateObject("ADODB.Connection")
    Set Wmi_ = CreateObject("WbemScripting.SWbemLocator")

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
    If Not Rcds_ Is Nothing Then Rcds_.Close
    Set Rcds_ = Nothing

    Con_.Close
    Set Con_ = Nothing

End Sub

Public Sub Init(ByVal dbType As dbTypeEnum, Optional ByVal fpath As String = "", Optional isTableHeader As Boolean = True)

    If fpath = "" Then
        If dbType = dbAccess Then Err.Raise 76 ''' Path not found
        fpath = filePath_
    Else
        fpath = fpath
    End If

    Dim tableHeader As String: tableHeader = IIf(isTableHeader, "HDR=YES", "HDR=NO")

    Select Case dbType

        Case 0: ''' Text

            Con_.Provider = dbEngine_
            Con_.Properties("Data Source") = Fso.GetParentFolderName(fpath)
            Con_.Properties("Extended Properties") = Join(Array("TEXT", tableHeader), ";")
            Con_.Open

        Case 1: ''' DBF

            Con_.Provider = dbEngine_
            Con_.Properties("Data Source") = Fso.GetParentFolderName(fpath)
            Con_.Properties("Extended Properties") = Join(Array("dBASE IV", "User ID=Admin"), ";")
            Con_.Open

        Case 2: ''' Access

            Con_.Provider = dbEngine_
            Con_.Properties("Data Source") = fpath
            Con_.Open

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

            Con_.Provider = dbEngine_
            Con_.Properties("Data Source") = fpath
            Con_.Properties("Extended Properties") = Join(Array(dbTypeExcel, tableHeader), ";")
            Con_.Open

    End Select

End Sub

Public Function Select_(ByVal sql As String, _
    Optional ByVal propCursorType As CursorTypeEnum = adOpenStatic, _
    Optional ByVal propLockType As LockTypeEnum = adLockReadOnly, _
    Optional ByVal propCommandType As CommandTypeEnum = adCmdText _
    )

    Set Rcds_ = CreateObject("ADODB.Recordset")
    Rcds_.Open Source:=sql, ActiveConnection:=Con_, CursorType:=propCursorType, LockType:=propLockType, Options:=propCommandType

    Dim v As Variant, arrx As New ArrayEx
    For Each v In Rcds_.Fields
        arrx.AddVal v.Name
    Next v

    Select_ = Array( _
        arrx.ToArray, _
        Arr2DToJagArr(ArrTranspose(Rcds_.GetRows())) _
    )

    On Error Resume Next
    If Not Rcds_ Is Nothing Then Rcds_.Close
    Set Rcds_ = Nothing

End Function

Public Function Execute(ByVal sql As String) As Long
    Dim rcdNumbers As Long
    Con_.Execute sql, ARecNum
    Execute = rcdNumbers
End Function

Public Sub BeginTrans()
    cn.BeginTrans
End Sub

Public Sub CommitTrans()
    Con_.CommitTrans
End Sub

Public Sub RollbackTrans()
    cn.RollbackTrans
End Sub

Private Function is64bitOS() As Boolean

  Dim v As Variant
  For Each v In Wmi_.connectServer.execQuery("Select * From Win32_OperatingSystem")

    If InStr(v.OSArchitecture, "64") Then
        is64bitOS = True
        Exit Function
    End If

  Next

End Function

Public Sub Pos(ByVal sql As String, Optional ByRef rng As Range = Nothing)

    Set Rcds_ = CreateObject("ADODB.Recordset")
    Rcds_.Open Source:=sql, ActiveConnection:=Con_, CursorType:=adOpenStatic, LockType:=adLockReadOnly, Options:=adCmdText

    If Rcds_.EOF Then Debug.Print "No Date!: " & Now(): Exit Sub

    If TypeName(rng) = "Nothing" Then
        AddSheet "test"
        Set rng = ThisWorkbook.Worksheets("test").Range("C5")
    End If

    rng.CopyFromRecordset Rcds_

    On Error Resume Next
    If Not Rcds_ Is Nothing Then Rcds_.Close
    Set Rcds_ = Nothing

End Sub
