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

''' Dependencies
'''
'''     ACE Engine
'''         Microsoft Access Database Engine 2016 Redistributable
'''         https://www.microsoft.com/en-us/download/details.aspx?id=54920
'''
'''     Ariawase
'''         https://github.com/vbaidiot/ariawase
'''         ArrRank ( ArrTranspose )

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

Public Enum dbTypeEnum
    dbText = 0
    dbDBF = 1
    dbAccess = 2
    dbExcel = 3
End Enum

Private Con_  As Object 'Is ADODB.Connection
Private Rcds_ As Object 'Is ADODB.Recordset
Private Wmi_  As Object 'Is WbemScripting.SWbemLocator
Private dbEngine_ As String
Private filePath_ As String

Private Sub Class_Initialize()

    Set Con_ = CreateObject("ADODB.Connection")
    Set Wmi_ = CreateObject("WbemScripting.SWbemLocator")
    filePath_ = ThisWorkbook.Path & "\" & ThisWorkbook.Name

    dbEngine_ = AceString(IIf(is64bitOS(), 64, 32))

    ''' if Ace is not installed...
    If dbEngine_ = "" Then
        ''' Office64 can not use Jet.
        #If VBA7 Then
            Err.Raise 9999, , "Please Install ACE Engine"
        #End If
        ''' Use Jet ( pre-installed on Windows32 and Wow64 )
        dbEngine_ = "Microsoft.Jet.OLEDB.4.0"
    End If

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

''' @param sql As String
''' @return As Variant(Of Array(Of T, T))
Public Function Select_(ByVal sql As String, _
    Optional ByVal propCursorType As CursorTypeEnum = adOpenStatic, _
    Optional ByVal propLockType As LockTypeEnum = adLockReadOnly, _
    Optional ByVal propCommandType As CommandTypeEnum = adCmdText _
    )

    Set Rcds_ = CreateObject("ADODB.Recordset")
    Rcds_.Open Source:=sql, ActiveConnection:=Con_, CursorType:=propCursorType, LockType:=propLockType, Options:=propCommandType

    ''' Table header names
    Dim fieldsCnt As Long: fieldsCnt = Rcds_.Fields.Count
    Dim tableHeaderNamesArray() As String: ReDim tableHeaderNamesArray(fieldsCnt - 1)

    Dim idx As Long
    For idx = 0 To (fieldsCnt - 1)
        tableHeaderNamesArray(idx) = Rcds_.Fields.Item(idx).Name
    Next idx

    ''' array( array( Table Names ),  array( Table Values ) )
    Select_ = Array(tableHeaderNamesArray, ArrTranspose(Rcds_.GetRows()))

    On Error Resume Next
    If Not Rcds_ Is Nothing Then Rcds_.Close
    Set Rcds_ = Nothing

End Function

''' @param sql As String
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

''' --------------------------------------------------------
'''                      Util Functions
''' --------------------------------------------------------

''' @return As Boolean
Private Function is64bitOS() As Boolean
    Dim v As Variant
    For Each v In Wmi_.ConnectServer.ExecQuery("Select * From Win32_OperatingSystem")
      If InStr(v.OSArchitecture, "64") Then
          is64bitOS = True
          GoTo Escape
      End If
    Next
Escape:
End Function

''' @param osBit As Long
''' @return As String
Private Function AceString(ByVal osBit As Long) As String

    On Error GoTo Escape

    Const strComputer As String = "."

    ''' Context Flags
    ''' https://docs.microsoft.com/en-us/windows/desktop/WmiSdk/requesting-wmi-data-on-a-64-bit-platform#context-flags
    Dim wmiCtx As Object: Set wmiCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
    wmiCtx.Add "__ProviderArchitecture", osBit
    wmiCtx.Add "__RequiredArchitecture", True

    Dim wmi    As Object: Set wmi = CreateObject("Wbemscripting.SWbemLocator")
    Dim wmiSrv As Object: Set wmiSrv = wmi.ConnectServer(strComputer, "root\default", "", "", , , , wmiCtx)

    AceString = AceStringImpl(wmiSrv.Get("StdRegProv"))

Escape:
    Set wmiCtx = Nothing
    Set wmi = Nothing
    Set wmiSrv = Nothing
End Function

''' @param reg As Object Is SWbemObjectEx
''' @return As String
Private Function AceStringImpl(ByVal reg As Object) As String

    On Error GoTo Escape

    Const HKLM  As Long = &H80000002
    Const Ace16 As String = "SOFTWARE\Classes\Microsoft.ACE.OLEDB.16.0"
    Const Ace15 As String = "SOFTWARE\Classes\Microsoft.ACE.OLEDB.15.0"
    Const Ace14 As String = "SOFTWARE\Classes\Microsoft.ACE.OLEDB.14.0"
    Const Ace13 As String = "SOFTWARE\Classes\Microsoft.ACE.OLEDB.13.0"
    Const Ace12 As String = "SOFTWARE\Classes\Microsoft.ACE.OLEDB.12.0"

    Dim v As Variant, SubKey As Variant
    For Each v In Array(Ace16, Ace15, Ace12)
        reg.EnumKey HKEY_LOCAL_MACHINE, v, SubKey
        If Not IsNull(SubKey) Then
            Dim arr As Variant: arr = Split(v, "\")
            AceStringImpl = arr(ArrLen(arr) - 1)
            Set reg = Nothing
            Exit Function
        End If
    Next v

Escape:
    AceStringImpl = ""
    Set reg = Nothing
End Function

''' @param arr2D As Variant(Of Array(Of T, T))
''' @return As Variant(Of Array(Of T, T))
Private Function ArrTranspose(ByVal arr2D As Variant) As Variant

    ''' This original function is from KoffeeVBA

    If Not IsArray(arr2D) Then Err.Raise 13
    If Not ArrRank(arr2D) = 2 Then Err.Raise 13

    Dim lb1 As Long: lb1 = LBound(arr2D, 2)
    Dim ub1 As Long: ub1 = UBound(arr2D, 2)
    Dim lb2 As Long: lb2 = LBound(arr2D, 1)
    Dim ub2 As Long: ub2 = UBound(arr2D, 1)

    Dim tmpArr2D() As Variant
    ReDim tmpArr2D(lb1 To ub1, lb2 To ub2)

    Dim ix1 As Long, ix2 As Long
    For ix1 = lb1 To ub1
        For ix2 = lb2 To ub2
            If IsObject(arr2D(ix2, ix1)) Then
                Set tmpArr2D(ix1, ix2) = arr2D(ix2, ix1)
            Else
                Let tmpArr2D(ix1, ix2) = arr2D(ix2, ix1)
            End If
        Next ix2
    Next ix1

    ArrTranspose = tmpArr2D

End Function

''' @param sql As String
''' @param rng As Range
Public Sub CopyFromRecordset(ByVal sql As String, ByVal rng As Range)

    Set Rcds_ = CreateObject("ADODB.Recordset")
    Rcds_.Open Source:=sql, ActiveConnection:=Con_, CursorType:=adOpenStatic, LockType:=adLockReadOnly, Options:=adCmdText

    If Rcds_.EOF Then
        Debug.Print "No Date!: " & Now()
        GoTo Escape
    End If

    rng.CopyFromRecordset Rcds_

    On Error Resume Next
    If Not Rcds_ Is Nothing Then Rcds_.Close
    Set Rcds_ = Nothing

Escape:
End Sub
