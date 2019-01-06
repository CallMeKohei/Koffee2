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
'''     Ariawase
'''         https://github.com/vbaidiot/ariawase
'''         ArrIndexOf ( Init )
'''         ArrRank    ( ArrTranspose )


'''  --------------------------------------------------------
'''     Please read below if you read xlsx Excel file ....
'''  --------------------------------------------------------
'''
'''  1) VBA can read .xlsx file with only Ace Engine.
'''
'''      ( Readable Files )
'''      Engine | dBase | Access | Text | Excel
'''      ------ +------ +------- +----- +-------
'''      Jet    | .dbf  | .mdb   | .txt | .xls
'''             |       |        | .csv |
'''      Ace    | .dbf  | .mdb   | .txt | .xls
'''             | .ndx  | .accdb | .csv | .xlsx
'''             | .mdx  |        |      | .xlsm
'''             |       |        |      | .xlsb
'''
'''  2) VBA can use installed Ace Engine.
'''
'''      VBA see the following place.
'''          \HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Microsoft.ACE.OLEDB.xx.0
'''
'''      Check code whether Ace Engine already be installed.
'''      Sub ExistAceEngine()
'''          Dim adox As AdoEx: Set adox = New AdoEx
'''          ''' Empty strings shows that VBA can not use Ace Engine.
'''          Debug.Print .adox.AceString
'''      End Sub
'''
'''      @seealso https://support.microsoft.com/en-us/help/2874601/can-t-use-the-access-odbc-driver-or-oledb-provider-outside-office-clic
'''      @seealso https://support.microsoft.com/ja-jp/help/2874601/can-t-use-the-access-odbc-driver-or-oledb-provider-outside-office-clic
'''
'''      Install Ace Engine
'''      Microsoft Access Database Engine 2016 Redistributable
'''      https://www.microsoft.com/en-us/download/details.aspx?id=54920
'''
'''      Is's OK whichever program
'''          AccessDatabaseEngine.exe
'''          AccessDatabaseEngine_X64.exe
'''      because VBA see just the following place.
'''          \HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Microsoft.ACE.OLEDB.xx.0


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
    dbDBF = 0
    dbAccess = 1
    dbText = 2
    dbExcel = 3
End Enum

Private Con_  As Object 'Is ADODB.Connection
Private dbEngine_ As String

Private Sub Class_Initialize()

    Set Con_ = CreateObject("ADODB.Connection")
    dbEngine_ = AceString()
    If dbEngine_ = "" Then dbEngine_ = "Microsoft.Jet.OLEDB.4.0"

End Sub

Private Sub Class_Terminate()

    Class_TerminateImpl

End Sub

Private Sub Class_TerminateImpl()

    On Error Resume Next
    Con_.Close
    Set Con_ = Nothing

End Sub

Public Sub Init(ByVal dbType As dbTypeEnum, Optional ByVal fpath As String = "", Optional isTableHeader As Boolean = True)

    ''' File path is current excel book if not specify.
    If fpath = "" Then
        If dbType = dbAccess Then Err.Raise 76 ''' Path not found
        fpath = ThisWorkbook.Path & "\" & ThisWorkbook.Name
    Else
        fpath = fpath
    End If

    ''' Table header is default yes.
    Dim tableHeader As String: tableHeader = IIf(isTableHeader, "HDR=YES", "HDR=NO")

    ''' File extension
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fileExt As String: fileExt = fso.GetExtensionName(fpath)

    ''' I can't understand following. But real is real. :-(
    If fileExt = "dbf" Then dbEngine_ = "Microsoft.Jet.OLEDB.4.0"
    If fileExt = "mdb" Then dbEngine_ = "Microsoft.Jet.OLEDB.4.0"
    If fileExt = "txt" Then dbEngine_ = "Microsoft.Jet.OLEDB.4.0"
    If fileExt = "csv" Then dbEngine_ = "Microsoft.Jet.OLEDB.4.0"
    If fileExt = "xls" Then dbEngine_ = "Microsoft.Jet.OLEDB.4.0"

    Dim Jet As Variant: Jet = Array( _
          "dbf" _
        , "mdb" _
        , "txt", "csv" _
        , "xls" _
    )

    Dim Ace As Variant: Ace = Array( _
          "dbf", "ndx", "mdx" _
        , "mdb", "accdb" _
        , "txt", "csv" _
        , "xls", "xlsx", "xlsm", "xlsb" _
    )

    ''' check for readable file type
    If dbEngine_ = "Microsoft.Jet.OLEDB.4.0" Then
        If ArrIndexOf(Jet, fileExt) = -1 Then GoTo Catch
    Else
        If ArrIndexOf(Ace, fileExt) = -1 Then GoTo Catch
    End If

    Select Case dbType

        Case 0: ''' DBF

            Con_.Provider = dbEngine_
            Con_.Properties("Data Source") = fso.GetParentFolderName(fpath)
            Con_.Properties("Extended Properties") = Join(Array("dBASE IV"), ";")
            Con_.Open

        Case 1: ''' Access

            Con_.Provider = dbEngine_
            Con_.Properties("Data Source") = fpath
            Con_.Open

        Case 2: ''' Text

            Con_.Provider = dbEngine_
            Con_.Properties("Data Source") = fso.GetParentFolderName(fpath)
            If fileExt = "csv" Then
                Con_.Properties("Extended Properties") = Join(Array("TEXT", tableHeader), ";")
            Else
                Con_.Properties("Extended Properties") = Join(Array("TEXT", tableHeader, "FMT=Delimited"), ";")
            End If
            Con_.Open

        Case 3: ''' Excel

            Dim dbTypeExcel As String
            Select Case fileExt
                Case "xls" ''' Excel5.0/95 Workbook(*.xls) or Excel97-2003 WorkBook(*.xls)
                    If InStr(fso.getfile(fpath).Type, 95) Then
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

            Con_.Provider = dbEngine_
            Con_.Properties("Data Source") = fpath
            Con_.Properties("Extended Properties") = Join(Array(dbTypeExcel, tableHeader), ";")
            Con_.Open

    End Select

    GoTo Escape

Catch:

    Class_TerminateImpl

    If dbEngine_ = "Microsoft.Jet.OLEDB.4.0" Then
        Debug.Print "(AdoEX/Init) Jet can not read file. Please check file extension"
        Err.Raise 9999, , "(AdoEX/Init) Jet can not read file. Please check file extension"
    Else
        Debug.Print "(AdoEX/Init) Ace can not read file. Please check file extension"
        Err.Raise 9999, , "(AdoEX/Init) Ace can not read file. Please check file extension"
    End If

Escape:
End Sub

''' @param sql As String
''' @return As Variant(Of Array(Of T, T))
Public Function Select_(ByVal sql As String, _
    Optional ByVal propCursorType As CursorTypeEnum = adOpenStatic, _
    Optional ByVal propLockType As LockTypeEnum = adLockReadOnly, _
    Optional ByVal propCommandType As CommandTypeEnum = adCmdText _
    ) As Variant

    Dim rcds As Object: Set rcds = CreateObject("ADODB.Recordset")
    rcds.Open Source:=sql, ActiveConnection:=Con_, CursorType:=propCursorType, LockType:=propLockType, Options:=propCommandType

    ''' Table header names
    Dim fieldsCnt As Long: fieldsCnt = rcds.Fields.Count
    Dim tableHeaderNamesArray() As String: ReDim tableHeaderNamesArray(fieldsCnt - 1)

    Dim idx As Long
    For idx = 0 To (fieldsCnt - 1)
        tableHeaderNamesArray(idx) = rcds.Fields.Item(idx).Name
    Next idx

    ''' array( array( Table Names ),  array( Table Values ) )
    Select_ = Array(tableHeaderNamesArray, ArrTranspose(rcds.GetRows()))

    On Error Resume Next
    If Not rcds Is Nothing Then rcds.Close
    Set rcds = Nothing

End Function

''' @param sql As String
''' @return As Long
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

''' @param reg As Object Is SWbemObjectEx
''' @return As String
Public Function AceString() As String

    ''' The Ace Engine is the same place whether it is 32 bit OS or 64 bit OS
    ''' HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Microsoft.ACE.OLEDB.xx.0
    Dim wmi        As Object: Set wmi = CreateObject("Wbemscripting.SWbemLocator")
    Dim wmiSrv     As Object: Set wmiSrv = wmi.ConnectServer(".", "root\default")
    Dim stdRegProv As Object: Set stdRegProv = wmiSrv.Get("StdRegProv")

    Const HKLM  As Long = &H80000002
    Const Ace16 As String = "SOFTWARE\Classes\Microsoft.ACE.OLEDB.16.0"
    Const Ace15 As String = "SOFTWARE\Classes\Microsoft.ACE.OLEDB.15.0"
    Const Ace14 As String = "SOFTWARE\Classes\Microsoft.ACE.OLEDB.14.0"
    Const Ace13 As String = "SOFTWARE\Classes\Microsoft.ACE.OLEDB.13.0"
    Const Ace12 As String = "SOFTWARE\Classes\Microsoft.ACE.OLEDB.12.0"

    Dim v As Variant, SubKey As Variant
    For Each v In Array(Ace16, Ace15, Ace14, Ace13, Ace12)
        stdRegProv.EnumKey HKLM, v, SubKey
        If Not IsNull(SubKey) Then
            Dim arr As Variant: arr = Split(v, "\")
            AceString = arr(ArrLen(arr) - 1)
            GoTo Escape
        End If
    Next v

Escape:
    Set wmi = Nothing
    Set wmiSrv = Nothing
    Set stdRegProv = Nothing
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

    Dim rcds As Object: Set rcds = CreateObject("ADODB.Recordset")
    rcds.Open Source:=sql, ActiveConnection:=Con_, CursorType:=adOpenStatic, LockType:=adLockReadOnly, Options:=adCmdText

    If rcds.EOF Then
        Debug.Print "No Date!: " & Now()
        GoTo Escape
    End If

    rng.CopyFromRecordset rcds

    If Not rcds Is Nothing Then rcds.Close
    Set rcds = Nothing

Escape:
End Sub
