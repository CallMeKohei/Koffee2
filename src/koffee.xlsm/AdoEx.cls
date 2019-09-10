VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    Const strComputer As String = "."
''' written by callmekohei(twitter at callmekohei)
''' MIT license
Option Explicit
Option Compare Text
Option Base 0

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
    dbDBF = 0
    dbAccess = 1
    dbText = 2
    dbExcel = 3
End Enum

Private Con_  As Object 'Is ADODB.Connection
Private dbEngine_ As String

Private Sub Class_Initialize()
    Set Con_ = CreateObject("ADODB.Connection")
'    dbEngine_ = Ace
End Sub

Private Sub Class_Terminate()
    Class_TerminateImpl
End Sub

Private Sub Class_TerminateImpl()
    On Error Resume Next
    Con_.Close
    Set Con_ = Nothing
End Sub

Public Sub InitCSV(ByVal aFolderPath As String)
    Con_.Provider = "Microsoft.Jet.OLEDB.4.0"
    Con_.Properties("Data Source") = aFolderPath
    Con_.Properties("Extended Properties") = Join(Array("TEXT", "HDR=NO", "FMT=Delimited"), ";")
    Con_.Open
End Sub

Public Sub InitCSVHeader(ByVal aFolderPath As String)
    Con_.Provider = "Microsoft.Jet.OLEDB.4.0"
    Con_.Properties("Data Source") = aFolderPath
    Con_.Properties("Extended Properties") = Join(Array("TEXT", "HDR=YES", "FMT=Delimited"), ";")
    Con_.Open
End Sub

Public Sub InitText(ByVal aFolderPath As String)
    Con_.Provider = "Microsoft.Jet.OLEDB.4.0"
    Con_.Properties("Data Source") = aFolderPath
    Con_.Properties("Extended Properties") = Join(Array("TEXT", "HDR=NO"), ";")
    Con_.Open
End Sub

Public Sub InitTextHeader(ByVal aFolderPath As String)
    Con_.Provider = "Microsoft.Jet.OLEDB.4.0"
    Con_.Properties("Data Source") = aFolderPath
    Con_.Properties("Extended Properties") = Join(Array("TEXT", "HDR=YES"), ";")
    Con_.Open
End Sub

Public Sub InitAccess(ByVal aFilePath As String)
    Con_.Provider = Ace
    Con_.Properties("Data Source") = aFilePath
    Con_.Open
End Sub

Public Sub InitExcel(ByVal aFilePath As String)

    On Error GoTo ErrHdl

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    Dim dbTypeExcel As String
    Select Case fso.GetExtensionName(aFilePath)
        Case "xls"  ''' Excel5.0/95 Workbook(*.xls) or Excel97-2003 WorkBook(*.xls)
            If InStr(fso.getFile(aFilePath).Type, 95) Then
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
    End Select
    
    Con_.Provider = Ace
    Con_.Properties("Data Source") = aFilePath
    Con_.Properties("Extended Properties") = Join(Array(dbTypeExcel, "HDR=NO"), ";")
    Con_.Open

ErrHdl:
    Set fso = Nothing
End Sub

Public Sub InitExcelHeader(ByVal aFilePath As String)

    On Error GoTo ErrHdl

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    Dim dbTypeExcel As String
    Select Case fso.GetExtensionName(aFilePath)
        Case "xls"  ''' Excel5.0/95 Workbook(*.xls) or Excel97-2003 WorkBook(*.xls)
            If InStr(fso.getFile(aFilePath).Type, 95) Then
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
    End Select
    
    Con_.Provider = Ace
    Con_.Properties("Data Source") = aFilePath
    Con_.Properties("Extended Properties") = Join(Array(dbTypeExcel, "HDR=YES"), ";")
    Con_.Open

ErrHdl:
    Set fso = Nothing
End Sub






'Public Sub Init(ByVal dbType As dbTypeEnum, Optional ByVal fPath As String = "", Optional isTableHeader As Boolean = True)
'
'On Error GoTo ErrHdl
'
'    Dim tableHeader As String: tableHeader = IIf(isTableHeader, "HDR=YES", "HDR=NO")
'
'    If Not dbType = dbTypeEnum.dbText Then GoTo Escape
'    dbEngine_ = "Microsoft.Jet.OLEDB.4.0"
'
'    Con_.Provider = dbEngine_
'    Con_.Properties("Data Source") = fPath
'    If dbType = dbCSV Then
'        Con_.Properties("Extended Properties") = Join(Array("TEXT", tableHeader, "FMT=Delimited"), ";")
'    Else
'        Con_.Properties("Extended Properties") = Join(Array("TEXT", tableHeader), ";")
'    End If
'    Con_.Open
'
'    GoTo Escape
'
'ErrHdl:
'    MsgBox "ÉGÉâÅ[Ç‚Ç≈"
'
'Escape:
'End Sub

'Public Sub Init(ByVal dbType As dbTypeEnum, Optional ByVal fPath As String = "", Optional isTableHeader As Boolean = True)
'
'    ''' File path is current excel book if not specify.
''    If fPath = "" Then
''        If dbType = dbAccess Then Err.Raise 76 ''' Path not found
''        fPath = ThisWorkbook.Path & "\" & ThisWorkbook.Name
''    Else
''        fPath = fPath
''    End If
'
'    ''' Table header is default yes.
'    Dim tableHeader As String: tableHeader = IIf(isTableHeader, "HDR=YES", "HDR=NO")
'dbEngine_ = "Microsoft.Jet.OLEDB.4.0"
'    ''' File extension
''    Dim Fso As Object: Set Fso = CreateObject("Scripting.FileSystemObject")
''    Dim fileExt As String: fileExt = LCase(Fso.GetExtensionName(fPath))
''
''    ''' I can't understand following. But real is real. :-(
''    If fileExt = "dbf" Then dbEngine_ = "Microsoft.Jet.OLEDB.4.0"
''    If fileExt = "mdb" Then dbEngine_ = "Microsoft.Jet.OLEDB.4.0"
''    If fileExt = "txt" Then dbEngine_ = "Microsoft.Jet.OLEDB.4.0"
''    If fileExt = "csv" Then dbEngine_ = "Microsoft.Jet.OLEDB.4.0"
''    If fileExt = "xls" Then dbEngine_ = "Microsoft.Jet.OLEDB.4.0"
'
''    Dim Jet As Variant: Jet = Array( _
''          "dbf" _
''        , "mdb" _
''        , "txt", "csv" _
''        , "xls" _
''    )
''
''    Dim Ace As Variant: Ace = Array( _
''          "dbf", "ndx", "mdx" _
''        , "mdb", "accdb" _
''        , "txt", "csv" _
''        , "xls", "xlsx", "xlsm", "xlsb" _
''    )
'
'    ''' check for readable file type
''    Dim xs As Variant, xFlag As Boolean
''    If dbEngine_ = "Microsoft.Jet.OLEDB.4.0" Then
''        For Each xs In Jet
''            If xs = fileExt Then xFlag = True
''        Next
''        If xFlag = False Then GoTo Catch
''    Else
''        For Each xs In Ace
''            If xs = fileExt Then xFlag = True
''        Next
''        If xFlag = False Then GoTo Catch
''    End If
'
'    Con_.Provider = dbEngine_
'    Con_.Properties("Data Source") = fPath
'    Con_.Properties("Extended Properties") = Join(Array("TEXT", tableHeader, "FMT=Delimited"), ";")
'    Con_.Open
'
'    Select Case dbType
'
'        Case 0: ''' DBF
'
'            Con_.Provider = dbEngine_
'            Con_.Properties("Data Source") = fso.GetParentFolderName(fPath)
'            Con_.Properties("Extended Properties") = Join(Array("dBASE IV"), ";")
'            Con_.Open
'
''        Case 1: ''' Access
''
''            Con_.Provider = dbEngine_
''            Con_.Properties("Data Source") = fPath
''            Con_.Open
''
''        Case 2: ''' Text
''
''            Con_.Provider = dbEngine_
''            Con_.Properties("Data Source") = Fso.GetParentFolderName(fPath)
''            If fileExt = "csv" Then
''                Con_.Properties("Extended Properties") = Join(Array("TEXT", tableHeader, "FMT=Delimited"), ";")
''            Else
''                Con_.Properties("Extended Properties") = Join(Array("TEXT", tableHeader), ";")
''            End If
''            Con_.Open
''
''        Case 3: ''' Excel
'
''            Dim dbTypeExcel As String
''            Select Case fileExt
''                Case "xls" ''' Excel5.0/95 Workbook(*.xls) or Excel97-2003 WorkBook(*.xls)
''                    If InStr(Fso.getFile(fPath).Type, 95) Then
''                        dbTypeExcel = "Excel 5.0"
''                    Else
''                        dbTypeExcel = "Excel 8.0"
''                    End If
''                Case "xlsb" ''' Excel binary Workbook(*.xlsb)
''                    dbTypeExcel = "Excel 12.0"
''                Case "xlsm" ''' Excel Macro-Enabled Wrokbook(*.xlsm)
''                    dbTypeExcel = "Excel 12.0 Macro"
''                Case "xlsx" ''' Excel Workbook(*.xlsx)
''                    dbTypeExcel = "Excel 12.0 Xml"
''            End Select
''
''            Con_.Provider = dbEngine_
''            Con_.Properties("Data Source") = fPath
''            Con_.Properties("Extended Properties") = Join(Array(dbTypeExcel, tableHeader), ";")
''            Con_.Open
''
''    End Select
'
'    GoTo Escape
'
'Catch:
'
'    Class_TerminateImpl
'
'    If dbEngine_ = "Microsoft.Jet.OLEDB.4.0" Then
'        Debug.Print "(AdoEX/Init) Jet can not read file. Please check file extension"
'        Err.Raise 9999, , "(AdoEX/Init) Jet can not read file. Please check file extension"
'    Else
'        Debug.Print "(AdoEX/Init) Ace can not read file. Please check file extension"
'        Err.Raise 9999, , "(AdoEX/Init) Ace can not read file. Please check file extension"
'    End If
'
'Escape:
'End Sub

''' @param sql As String
''' @return As Variant(Of Array(Of T, T))
Public Function ArraySelect(ByVal sql As String, _
    Optional ByVal propCursorType As CursorTypeEnum = adOpenStatic, _
    Optional ByVal propLockType As LockTypeEnum = adLockReadOnly, _
    Optional ByVal propCommandType As CommandTypeEnum = adCmdText _
    ) As Variant

    Dim rcds As Object: Set rcds = CreateObject("ADODB.Recordset")
    rcds.Open Source:=sql, ActiveConnection:=Con_, CursorType:=propCursorType, LockType:=propLockType, options:=propCommandType

    ''' Table header names
    Dim fieldsCnt As Long: fieldsCnt = rcds.Fields.Count
    Dim tableHeaderNamesArray() As String: ReDim tableHeaderNamesArray(fieldsCnt - 1)

    Dim idx As Long
    For idx = 0 To (fieldsCnt - 1)
        tableHeaderNamesArray(idx) = rcds.Fields.Item(idx).Name
    Next idx

    ''' array( array( Table Names ),  array( Table Values ) )
    ArraySelect = Array(tableHeaderNamesArray, Arr2DToJagArr(ArrayTranspose(rcds.GetRows())))

    On Error Resume Next
    If Not rcds Is Nothing Then rcds.Close
    Set rcds = Nothing

End Function

''' @param sql As String
''' @return As Long
'Public Function Execute(ByVal sql As String) As Long
'    Dim rcdNumbers As Long
'    Con_.Execute sql, ARecNum
'    Execute = rcdNumbers
'End Function

'Public Sub BeginTrans()
'    cn.BeginTrans
'End Sub
'
'Public Sub CommitTrans()
'    Con_.CommitTrans
'End Sub
'
'Public Sub RollbackTrans()
'    cn.RollbackTrans
'End Sub


''' ----- Utility functions

Public Sub CopyFromRecordset(ByVal sql As String, ByVal rng As Range)

    Dim rcds As Object: Set rcds = CreateObject("ADODB.Recordset")
    rcds.Open Source:=sql, ActiveConnection:=Con_, CursorType:=adOpenStatic, LockType:=adLockReadOnly, options:=adCmdText

    If rcds.EOF Then
        Debug.Print "No Date!: " & Now()
        GoTo Escape
    End If

    rng.CopyFromRecordset rcds

Escape:
    If Not rcds Is Nothing Then rcds.Close
    Set rcds = Nothing
End Sub

Public Function Ace() As String
    Dim s As String
    s = AceDirect
    If s = "" Then s = AceOffice32ClickToRun
    If s = "" Then s = AceClickToRunOffice32
    If s = "" Then s = "Microsoft.Jet.OLEDB.4.0"
    Ace = s
End Function

''' direct driver
Private Function AceDirect() As String

    ''' The Ace Engine is the same place whether it is 32 bit OS or 64 bit OS
    ''' HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Microsoft.ACE.OLEDB.xx.0
    Dim Wmi        As Object: Set Wmi = CreateObject("Wbemscripting.SWbemLocator")
    Dim wmiSrv     As Object: Set wmiSrv = Wmi.ConnectServer(".", "root\default")
    Dim stdRegProv As Object: Set stdRegProv = wmiSrv.Get("StdRegProv")

    Const HKLM  As Long = &H80000002
    Const Ace16 As String = "SOFTWARE\Classes\Microsoft.ACE.OLEDB.16.0"
    Const Ace15 As String = "SOFTWARE\Classes\Microsoft.ACE.OLEDB.15.0"
    Const Ace14 As String = "SOFTWARE\Classes\Microsoft.ACE.OLEDB.14.0"
    Const Ace13 As String = "SOFTWARE\Classes\Microsoft.ACE.OLEDB.13.0"
    Const Ace12 As String = "SOFTWARE\Classes\Microsoft.ACE.OLEDB.12.0"

    Dim v As Variant, subkey As Variant
    For Each v In Array(Ace16, Ace15, Ace14, Ace13, Ace12)
        stdRegProv.EnumKey HKLM, v, subkey
        If Not IsNull(subkey) Then
            Dim arr As Variant: arr = Split(v, "\")
            AceDirect = arr(UBound(arr, 1))
            GoTo Escape
        End If
    Next v

Escape:
    Set Wmi = Nothing
    Set wmiSrv = Nothing
    Set stdRegProv = Nothing
End Function

''' under office driver ( 64bit win, 32bit office )
Private Function AceOffice32ClickToRun() As String

    On Error GoTo Escape

    Const strComputer As String = "."

    ''' Context Flags
    ''' https://docs.microsoft.com/en-us/windows/desktop/WmiSdk/requesting-wmi-data-on-a-64-bit-platform#context-flags
    Dim wmiCtx As Object: Set wmiCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
    wmiCtx.add "__ProviderArchitecture", 64
    wmiCtx.add "__RequiredArchitecture", True

    Dim Wmi        As Object: Set Wmi = CreateObject("Wbemscripting.SWbemLocator")
    Dim wmiSrv     As Object: Set wmiSrv = Wmi.ConnectServer(strComputer, "root\default", "", "", , , , wmiCtx)
    Dim stdRegProv As Object: Set stdRegProv = wmiSrv.Get("StdRegProv")

    Const HKEY_LOCAL_MACHINE  As Long = &H80000002

    Dim v, officeVersionArray, clct As Collection: Set clct = New Collection
    stdRegProv.EnumKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Office\", officeVersionArray
    For Each v In ArrayRegexFilter(officeVersionArray, "\d\d\.\d")
        clct.add "SOFTWARE\Microsoft\Office\" & v & "\ClickToRun\REGISTRY\MACHINE\Software\Classes\"
    Next v

    Dim subkey, i As Long, j As Long
    For i = 9 To 0 Step -1
        For j = 1 To clct.Count
            stdRegProv.EnumKey HKEY_LOCAL_MACHINE, clct.Item(j) & "Microsoft.ACE.OLEDB.1" & i & ".0", subkey
            If Not IsNull(subkey) Then
                AceOffice32ClickToRun = "Microsoft.ACE.OLEDB.1" & i & ".0"
                GoTo Escape:
            End If
        Next j
    Next i

Escape:
    Set Wmi = Nothing
    Set wmiSrv = Nothing
    Set stdRegProv = Nothing
    Set clct = Nothing
End Function

''' under ClickToRun ( 64bit win, 32bit office )
Private Function AceClickToRunOffice32() As String

    On Error GoTo Escape

    Const strComputer As String = "."

    ''' Context Flags
    ''' https://docs.microsoft.com/en-us/windows/desktop/WmiSdk/requesting-wmi-data-on-a-64-bit-platform#context-flags
    Dim wmiCtx As Object: Set wmiCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
    wmiCtx.add "__ProviderArchitecture", 64
    wmiCtx.add "__RequiredArchitecture", True

    Dim Wmi        As Object: Set Wmi = CreateObject("Wbemscripting.SWbemLocator")
    Dim wmiSrv     As Object: Set wmiSrv = Wmi.ConnectServer(strComputer, "root\default", "", "", , , , wmiCtx)
    Dim stdRegProv As Object: Set stdRegProv = wmiSrv.Get("StdRegProv")

    Const HKEY_LOCAL_MACHINE  As Long = &H80000002

    Const clicktorun As String = "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\"

    Dim subkey, i As Long
    For i = 9 To 0 Step -1
        stdRegProv.EnumKey HKEY_LOCAL_MACHINE, clicktorun & "Microsoft.ACE.OLEDB.1" & i & ".0", subkey
        If Not IsNull(subkey) Then
            AceClickToRunOffice32 = "Microsoft.ACE.OLEDB.1" & i & ".0"
            GoTo Escape:
        End If
    Next i

Escape:
    Set Wmi = Nothing
    Set wmiSrv = Nothing
    Set stdRegProv = Nothing
End Function


