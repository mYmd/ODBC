VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vb_ADODBs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' vb_ADODBs
'Copyright (c) 2017 mmYYmmdd

'*********************************************************************************
' ADODB接続（ほぼINSERTに特化）
' ADO(Microsoft AitiveX Data Objects) および ADOX(Microsoft ADO Ext.)を参照設定
' --------------------------------------------------------------------------------
' Function  connect             接続文字列を発行してデータベースに接続
' Sub       disconnect          接続断
' Function  tableExists         テーブル存在確認
' Function  insert              INSERT（１行）
' Function  insert_             INSERT（複数行）
' Function  tableInstance       属性設定用のADOX.Table オブジェクトを返す
' Function  createTable         テーブル生成
' Function  renameTable         テーブル名前変更
' Function  dropTable           テーブル削除
' Function  truncateTable       テーブル全行削除
' Function  mdb_expr            accessdbの接続文字列
' Function  excel_expr          Excelブックの接続文字列
' Function  sqlServer_expr      SQL Serverの接続文字列
'*********************************************************************************

Private Declare PtrSafe Function SQLGetInstalledDrivers Lib "odbccp32.dll" ( _
                                                ByVal buf As String, _
                                            ByVal cbBufMax As Integer, _
                                        ByRef pcchBufOut As Integer) As Byte

Private Declare PtrSafe Function GetTickCount Lib "kernel32.dll" () As Long

    Private mycon_              As ADODB.Connection
    Const providerExpr As String = "Provider=Microsoft.Ace.OLEDB.12.0;"

Private Sub Class_Initialize()
    Set mycon_ = New ADODB.Connection
End Sub

Private Sub Class_Terminate()
    disconnect
    Set mycon_ = Nothing
End Sub

' 接続文字列を発行してデータベースに接続
Public Function connect(ByVal connectionExpr As String, _
                    Optional ByVal provider As String = "MSDASQL") As Boolean
    connect = False
    disconnect
    On Error Resume Next
    If StrConv(connectionExpr, vbLowerCase) Like "*provider*=*" Then
        mycon_.Open connectionExpr
    Else
        mycon_.Open "Provider=" & provider & ";" & connectionExpr
    End If
    connect = mycon_.State <> adStateClosed
    ' https://msdn.microsoft.com/ja-jp/library/cc426812.aspx
End Function

' 接続断
Public Sub disconnect()
    If mycon_.State <> adStateClosed Then mycon_.Close
End Sub

' テーブル存在確認
Public Function tableExists(ByVal tbname As String) As Boolean
    tableExists = False
    If mycon_.State = adStateClosed Then Exit Function
    Dim z As ADOX.Table
    With New ADOX.Catalog
        .ActiveConnection = mycon_
        For Each z In .Tables
            If z.Name = tbname Then
                tableExists = True
                Exit For
            End If
        Next z
    End With
End Function

' INSERT（１行）
' Attrは列名もしくは列の序数（0スタート）の配列
Public Function insert(ByVal tableName As String, _
                       ByRef attr As Variant, _
                       ByRef values As Variant) As Boolean
    insert = False
    If mycon_.State = adStateClosed Then Exit Function
    If sizeof(attr) <> sizeof(values) Then Exit Function
    If 0 = Dimension(values) Or 2 < Dimension(values) Then Exit Function
    Dim flag As Variant
    flag = mapF(p_is_valid, values)
    With New ADODB.Recordset
        .source = tableName
        .ActiveConnection = mycon_
        .LockType = adLockOptimistic
        .Open
        .AddNew filterR(attr, flag), filterR(values, flag)
        .Update
        .Close
    End With
    insert = True
End Function

' INSERT（複数行）
' Attrは列名もしくは列の序数（0スタート）の配列
Public Function insert_(ByVal tableName As String, _
                        ByRef attr As Variant, _
                        ByRef values As Variant) As Variant
    If mycon_.State = adStateClosed Then Exit Function
    If 0 = Dimension(values) Or 2 < Dimension(values) Then Exit Function
    Dim ret As Variant: ret = makeM(sizeof(values, 1))
    Dim flag As Variant
    Dim i As Long, k As Long: k = 0
    With New ADODB.Recordset
        .source = tableName
        .ActiveConnection = mycon_
        .LockType = adLockOptimistic
        .Open
        Select Case Dimension(values)
        Case 1:
            For i = LBound(values) To UBound(values) Step 1
                If sizeof(attr) = sizeof(values(i)) Then
                    flag = mapF(p_is_valid, values(i))
                    .AddNew filterR(attr, flag), filterR(values(i), flag)
                    ret(k) = True
                Else
                    ret(k) = False
                End If
                k = k + 1
            Next i
            .Update
            .Close
        Case 2:
            If sizeof(attr) <> sizeof(values, 2) Then
                .Close
                Exit Function
            End If
            Dim tmp As Variant
            For i = LBound(values, 1) To UBound(values, 1) Step 1
                tmp = selectRow(values, i)
                flag = mapF(p_is_valid, tmp)
                .AddNew filterR(attr, flag), filterR(tmp, flag)
                ret(k) = True
                k = k + 1
            Next i
            .Update
            .Close
        End Select
    End With
    Call swapVariant(ret, insert_)
End Function

' adBoolean         11      ブール値型
' adCurrency         6      通貨型
' adDate             7      日付時刻型
' adDouble           5      倍精度浮動小数点数型
' adGUID            72      オートナンバー型
' adInteger          3      長整数型
' adLongVarWChar   203      メモ型
' adSingle           4      単精度浮動小数点数型
' adSmallInt         2      整数型
' adVarWChar       202      テキスト型

' 属性設定用のADOX.Table オブジェクトを返す
' （createTableの引数と使用することを想定）
Public Function tableInstance(ByVal tName As String) As ADOX.Table
    Set tableInstance = New ADOX.Table
    tableInstance.Name = tName
    'TB.Columns.append "学籍番号", adInteger
    'TB.Columns.append "氏名", adVarWChar, 30
    'TB.Columns.Item(i).Attributes = adColNullable
End Function

' テーブル生成
Public Function createTable(ByVal tb As ADOX.Table) As Boolean
    createTable = False
    If mycon_.State = adStateClosed Then Exit Function
    With New ADOX.Catalog
        .ActiveConnection = mycon_
        .Tables.append tb
    End With
    createTable = True
End Function

' テーブル名前変更
Public Function renameTable(ByVal tbname_old As String, ByVal tbname_new As String) As BookmarkEnum
    renameTable = False
    If mycon_.State = adStateClosed Then Exit Function
    With New ADOX.Catalog
        .ActiveConnection = mycon_
        .Tables.Item(tbname_old).Name = tbname_new
    End With
    renameTable = True
End Function

' テーブル削除
Public Function dropTable(ByVal tbname As String) As Boolean
    dropTable = False
    If mycon_.State = adStateClosed Then Exit Function
    With New ADOX.Catalog
        .ActiveConnection = mycon_
        .Tables.Delete tbname
    End With
    dropTable = True
End Function

' テーブル全行削除
Public Function truncateTable(ByVal tbname As String) As Long
    If mycon_.State = adStateClosed Then Exit Function
    mycon_.Execute "DELETE * FROM " & tbname, truncateTable, adExecuteNoRecords
    ' （レコードセットをDeleteしてMoveNextいくより速い模様）
End Function

' accessdbの接続文字列
Public Function mdb_expr(ByVal dbName As String) As String
    mdb_expr = providerExpr
    mdb_expr = mdb_expr & "Data Source=" & dbName & ";"
End Function

' Excelブックの接続文字列
Public Function excel_expr(ByVal bookName As String, _
                           Optional header As Boolean = True) As String
    excel_expr = providerExpr & _
                 "Data Source=" & bookName & ";" & _
                 "Extended Properties=""[[EP]];HDR=[[HDR]]"";"
    Dim Extended_Properties As String
    Select Case Right(bookName, 4)
        Case ".xls":    Extended_Properties = "Excel 8.0"
        Case "xlsx":    Extended_Properties = "Excel 12.0 Xml"
        Case "xlsb":    Extended_Properties = "Excel 12.0"
        Case "xlsm":    Extended_Properties = "Excel 12.0 Macro"
    End Select
    excel_expr = Replace(excel_expr, "[[EP]]", Extended_Properties)
    excel_expr = Replace(excel_expr, "[[HDR]]", IIf(header, "YES", "NO"))
End Function

' SQL Serverの接続文字列
Public Function sqlServer_expr(ByVal server As String, _
                               ByVal dbName As String, _
                               Optional ByVal uid As String = "", _
                               Optional ByVal pwd As String = "") As String
    Dim drivers As Variant
    Dim i As Variant
    sqlServer_expr = "Provider=SQLOLEDB;Driver={SQL Server Native Client 11.0};"
    If 0 = Len(uid) Then
        sqlServer_expr = sqlServer_expr & "Trusted_Connection=YES;"
    End If
    sqlServer_expr = sqlServer_expr & "Server=" & server & ";" _
                                    & "DATABASE=" & dbName & ";"
    If 0 < Len(uid) Then
        sqlServer_expr = sqlServer_expr & "UID=" & uid & ";"
    End If
    If 0 < Len(pwd) Then
        sqlServer_expr = sqlServer_expr & "PWD=" & pwd & ";"
    End If
    ' select * from sysobjects where xtype = 'U'
End Function
