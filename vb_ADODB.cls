VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vb_ADODB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' vb_ADODB
'Copyright (c) 2017 mmYYmmdd

'*********************************************************************************
' ADODB接続
' ADO(Microsoft AitiveX Data Objects) および ADOX(Microsoft ADO Ext.)を参照設定
' --------------------------------------------------------------------------------
' Function connect              接続文字列を発行してデータベースに接続
' Sub disconnect                接続断
' Sub clear                     ヘッダ（カラム）データとテーブルデータを消去
' Property Get last_connected   最後に成功した接続文字列の取得
' Property Get last_SQL         最後に発行されたSQL文字列の取得
' Property Get last_elapsed     最後に発行されたSELECT文でかかった時間（ミリ秒）
' Property Get last_execError   最後に発行されたexec文のエラー行
' Function get_header           ヘッダ（カラム）情報を取得する
' Function list_tables          テーブル一覧取得
' Function tableAttribute       テーブルの Attribute
' Function select_flat          Select文を発行して2次元配列としてメンバに格納する
' Function select_rowWise       Select文を発行して行の配列としてメンバに格納する
' Function select_columnWise    Select文を発行して列の配列としてメンバに格納する
' Function exec                 insert文やdelete文を実行する
' Function rowSize              Selectしたレコード数
' Function colSize              Selectしたカラム数
' Function get_data             Selectしたテーブルデータのコピーを返す
' Function move_data            Selectしたテーブルデータをmoveして返す
' Function SQLDrivers           インストールされているODBCドライバ配列を返す
' Function excel_expr           特定のExcelブックに接続するための接続文字列
' Function mdb_expr             accessdbの接続文字列
' Function sqlServer_expr       SQL Serverの接続文字列
' Function insert_expr          insert文
' Function insert_expr_         insert文（jag配列または2次元配列）
'*********************************************************************************

Private Declare PtrSafe Function SQLGetInstalledDrivers Lib "odbccp32.dll" ( _
                                                ByVal buf As String, _
                                            ByVal cbBufMax As Integer, _
                                        ByRef pcchBufOut As Integer) As Byte

Private Declare PtrSafe Function GetTickCount Lib "kernel32.dll" () As Long

    Private mycon_              As ADODB.Connection
    Private connectionString    As String    ' 接続文字列
    Private mySQL_              As String
    Private myheader_           As Variant
    Private mydata_             As Variant
    Private myErrorNo_          As Variant
    Private laptime_            As Long

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
    If connect Then connectionString = connectionExpr
    ' https://msdn.microsoft.com/ja-jp/library/cc426812.aspx
End Function

' 接続断
Public Sub disconnect()
    If mycon_.State <> adStateClosed Then mycon_.Close
End Sub

' ヘッダ（カラム）データとテーブルデータを消去
Public Sub clear()
    myheader_ = Empty
    mydata_ = Empty
End Sub

' 最後に成功した接続文字列の取得
Public Property Get last_connected() As String
    last_connected = connectionString
End Property

' 最後に発行されたSQL文字列の取得
Public Property Get last_SQL() As String
    last_SQL = mySQL_
End Property

' 最後に発行されたSELECT文でかかった時間（ミリ秒）
Public Property Get last_elapsed() As Long
    last_elapsed = laptime_
End Property

'最後に発行されたexec文のエラー行
Public Property Get last_execError() As Variant
    last_execError = myErrorNo_
End Property

' ヘッダ（カラム）情報を取得する
Public Function get_header() As Variant
    get_header = myheader_
End Function

' テーブル一覧取得
Public Function list_tables(Optional ByVal type_ As String = "TABLE") As Variant
    If mycon_.State <> adStateClosed Then
        Dim CAT As ADOX.Catalog
        Set CAT = New ADOX.Catalog
        CAT.ActiveConnection = mycon_
        Dim count As Long:      count = CAT.Tables.count
        Dim ret As Variant:     ret = makeM(count, 2)
        Dim i As Long
        For i = 0 To count - 1 Step 1
            ret(i, 0) = CAT.Tables(i).name
            ret(i, 1) = CAT.Tables(i).Type
        Next i
        Set CAT = Nothing
        If type_ <> "*" Then
            ret = filterR(ret, mapF(p_equal(type_), selectCol(ret, 1)))
        End If
        Call swapVariant(list_tables, ret)
    End If
End Function

' テーブルの Attribute
' 列名、型名、データ長、精度、精度、属性（1:キー, -1:Nullable, 0:他）
' SQLServerはSQLOLEDBでconnectしないと取得できない
Public Function tableAttribute(ByVal tableName As String) As Variant
    If mycon_.State <> adStateClosed Then
        Dim CAT As ADOX.Catalog
        Set CAT = New ADOX.Catalog
        CAT.ActiveConnection = mycon_
        Dim TBL As ADOX.Table
        Dim i As Long, j As Long, count As Long
        Dim ret As Variant
        Set TBL = CAT.Tables(tableName)
        count = TBL.Columns.count
        ret = makeM(count, 6)
        For i = 0 To count - 1 Step 1
            With TBL.Columns(i)
                ret(i, 0) = .name
                ret(i, 1) = DataTypeEnum2Str(.Type)
                ret(i, 2) = .DefinedSize
                ret(i, 3) = .Precision
                ret(i, 4) = .NumericScale
                ret(i, 5) = IIf(.Attributes = adColNullable, -1, 0) '2
            End With
        Next i
        Dim pos As Long
        For i = 0 To TBL.Indexes.count - 1 Step 1
            With TBL.Indexes(i)
                If .PrimaryKey Then
                    For j = 0 To .Columns.count - 1 Step 1
                        pos = find_pred(p_equal(.Columns(j).name), selectCol(ret, 0))
                        If pos < Haskell_2_stdFun.rowSize(ret) Then ret(pos, 5) = 1
                    Next j
                    Exit For
                End If
            End With
        Next i
        Set TBL = Nothing
        Set CAT = Nothing
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        Dim header As Variant
        header = selectCol(read_header_(rs, "SELECT * FROM " & tableName), 0)
        Set rs = Nothing
        Dim ord As Variant
        ord = mapF(p_find_pred(p_equal(, ph_2), selectCol(ret, 0)), header)
        tableAttribute = subM(ret, ord)
    End If
End Function

' Select文を発行して2次元配列としてメンバに格納する
Public Function select_flat(ByVal sqlExpr As String) As vb_ADODB
    Set select_flat = select_rowWise(sqlExpr)
    If Not IsEmpty(mydata_) Then mydata_ = unzip(mydata_, 2)
End Function

' Select文を発行して行の配列としてメンバに格納する
Public Function select_rowWise(ByVal sqlExpr As String) As vb_ADODB
    laptime_ = GetTickCount()
    Set select_rowWise = Me
    Me.clear
    If mycon_.State = adStateClosed Then Exit Function
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    On Error GoTo select_error
    myheader_ = read_header_(rs, sqlExpr)
    If rs.BOF Then
        rs.Close
        Exit Function
    Else
        Dim vec As vh_stdvec
        Set vec = New vh_stdvec
        Dim rec As Variant
        Dim count As Long, i As Long
        count = Haskell_2_stdFun.rowSize(myheader_)
        Do
            rec = makeM(count)
            For i = 0 To count - 1 Step 1
                rec(i) = rs.Fields(i).value
            Next i
            vec.push_back rec
            rs.MoveNext
        Loop Until rs.EOF = True
    End If
    rs.Close
    mydata_ = vec.free
    mySQL_ = sqlExpr
select_error:
    laptime_ = GetTickCount() - laptime_
End Function

' Select文を発行して列の配列としてメンバに格納する
Public Function select_columnWise(ByVal sqlExpr As String) As vb_ADODB
    Set select_columnWise = select_rowWise(sqlExpr)
    If Not IsEmpty(mydata_) Then mydata_ = unzip(mydata_, 1)
End Function

' insert文やdelete文を実行する
Public Function exec(ByRef sqlExpr As Variant) As vb_ADODB
    myErrorNo_ = Empty
    On Error Resume Next
    If mycon_.State <> adStateClosed Then
        Dim vec As vh_stdvec
        Set vec = New vh_stdvec
        Dim recordsAffected As Long
        If IsArray(sqlExpr) Then
            Dim i As Long
            For i = LBound(sqlExpr) To UBound(sqlExpr) Step 1
                recordsAffected = 0
                Call mycon_.Execute(sqlExpr(i), _
                                  recordsAffected, _
                                  adExecuteNoRecords)
                If recordsAffected = 0 Then vec.push_back i
            Next i
        Else
            recordsAffected = 0
            Call mycon_.Execute(sqlExpr, _
                              recordsAffected, _
                              adExecuteNoRecords)
            If recordsAffected = 0 Then vec.push_back 0
        End If
    End If
    Set exec = Me
    If 0 < vec.size Then
        myErrorNo_ = vec.free
    End If
End Function

' テーブルのレコード数
Public Function rowSize() As Long
    rowSize = Haskell_2_stdFun.rowSize(mydata_)
End Function

' テーブルのカラム数
Public Function colSize() As Long
    colSize = Haskell_2_stdFun.colSize(mydata_)
End Function

' Select したテーブルデータのコピーを返す
Public Function get_data() As Variant
    get_data = mydata_
End Function

' Select したテーブルデータをmoveして返す
Public Function move_data() As Variant
    Call swapVariant(move_data, mydata_)
End Function

' インストールされているODBCドライバ配列を返す
Public Function SQLDrivers(Optional ByVal buflen = 0) As Variant
    Dim buf As String, ret As Byte, retL As Integer
    buf = Space(IIf(0 < buflen, buflen, 2 * 1024))
    ret = SQLGetInstalledDrivers(buf, Len(buf), retL)
    SQLDrivers = Split(left(buf, retL - 2), Chr(0))
End Function

' 特定のExcelブックに接続するための接続文字列
Public Function excel_expr(ByVal bookName As String) As String
    Dim drivers As Variant
    Dim i As Variant
    excel_expr = "Microsoft Excel Driver"
    drivers = SQLDrivers
    permutate drivers, reverse(sortIndex(mapF(p_len, drivers)))
    i = find_pred(p_equal(excel_expr), mapF(p_left(, Len(excel_expr)), drivers))
    If i <= UBound(drivers) Then
        excel_expr = "Provider=MSDASQL;DBA=R;Driver={" & drivers(i) & "};DBQ="
        excel_expr = excel_expr & bookName
    Else
        excel_expr = ""
    End If
End Function

' accessdbの接続文字列
Public Function mdb_expr(ByVal dbName As String) As String
    Dim i As Variant
    mdb_expr = "Provider=Microsoft.Ace.OLEDB.12.0;"
    mdb_expr = mdb_expr & "Data Source=" & dbName
    'MSysAccessObjects
    'MSysObjects
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
        sqlServer_expr = sqlServer_expr & " Trusted_Connection=YES;"
    End If
    sqlServer_expr = sqlServer_expr & " Server=" & server & "; DATABASE=" & dbName
    If 0 < Len(uid) Then
        sqlServer_expr = sqlServer_expr & "; UID=" & uid
    End If
    If 0 < Len(pwd) Then
        sqlServer_expr = sqlServer_expr & "; PWD=" & pwd
    End If
    ' select * from sysobjects where xtype = 'U'
End Function

' insert文
Public Function insert_expr(ByVal table_name As String, _
                                ByRef attr As Variant, _
                                    ByRef values As Variant) As String
    Dim tmp As Variant:     tmp = values
    Dim flag As Variant:    flag = makeM(Haskell_2_stdFun.rowSize(tmp))
    Dim i As Long
    For i = LBound(tmp) To UBound(tmp) Step 1
        flag(i) = 1
        If IsNull(tmp(i)) Or IsEmpty(tmp(i)) Or IsObject(tmp(i)) Or IsArray(tmp(i)) Then
            flag(i) = 0
        ElseIf Not IsNumeric(tmp(i)) Or VarType(tmp(i)) = vbString Then
            If 0 < Len(tmp(i)) Then
                tmp(i) = "'" & Replace(tmp(i), "'", "''") & "'"
            Else
                flag(i) = 0
            End If
        End If
    Next i
    insert_expr = "INSERT INTO " & table_name
    If Dimension(attr) = 1 Then
        insert_expr = insert_expr & " (" & _
                    join(filterR(attr, flag), ",") & _
                ")"
    End If
    insert_expr = insert_expr & " VALUES(" & _
                join(filterR(tmp, flag), ",") & _
            ");"
End Function

' insert文（jag配列または2次元配列）
Public Function insert_expr_(ByVal table_name As String, _
                                ByRef attr As Variant, _
                                    ByRef values As Variant) As Variant
    Dim ret As Variant
    ret = makeM(Haskell_2_stdFun.rowSize(values))
    Dim i As Long
    If Dimension(values) = 1 Then
        For i = 0 To UBound(ret) Step 1
            ret(i) = insert_expr(table_name, attr, getNth_b(values, i))
        Next i
    Else
        For i = 0 To UBound(ret) Step 1
            ret(i) = insert_expr(table_name, attr, selectRow_b(values, i))
        Next i
    End If
    Call swapVariant(insert_expr_, ret)
End Function

    ' ヘッダ情報取得
    Private Function read_header_(ByVal rs As ADODB.Recordset, _
                                  ByRef sqlExpr As String) As Variant
        On Error GoTo error__
        With rs
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .Open sqlExpr, _
                    mycon_, _
                        adOpenForwardOnly, _
                            adLockReadOnly, _
                                adCmdUnspecified
        End With
        Dim count As Long, i As Long
        count = rs.Fields.count
        Dim header As Variant
        header = makeM(count, 5)
        For i = 0 To count - 1 Step 1
            With rs.Fields(i)
                header(i, 0) = .name
                header(i, 1) = DataTypeEnum2Str(.Type)
                header(i, 2) = .DefinedSize
                header(i, 3) = .Precision
                header(i, 4) = .NumericScale
            End With
        Next i
        Call swapVariant(read_header_, header)
error__:
    End Function

    ' フィールド型のenum値から型名に変換
    Private Function DataTypeEnum2Str(ByVal n As Long) As String
        Select Case n
            Case adArray:               DataTypeEnum2Str = "ARRAY"
            Case adBigInt:              DataTypeEnum2Str = "BIGINT"
            Case adBinary:              DataTypeEnum2Str = "BINARY"
            Case adBoolean:             DataTypeEnum2Str = "BOOLEAN"
            Case adBSTR:                DataTypeEnum2Str = "BSTR"
            Case adChapter:             DataTypeEnum2Str = "CHAPTER"
            Case adChar:                DataTypeEnum2Str = "CHAR"
            Case adCurrency:            DataTypeEnum2Str = "CURRENCY"
            Case adDate:                DataTypeEnum2Str = "DATE"
            Case adDBDate:              DataTypeEnum2Str = "DBDATE"
            Case adDBTime:              DataTypeEnum2Str = "DBTIME"
            Case adDBTimeStamp:         DataTypeEnum2Str = "DBTIMESTAMP"
            Case adDecimal:             DataTypeEnum2Str = "DECIMAL"
            Case adDouble:              DataTypeEnum2Str = "DOUBLE"
            Case adEmpty:               DataTypeEnum2Str = "EMPTY"
            Case adError:               DataTypeEnum2Str = "ERROR"
            Case adFileTime:            DataTypeEnum2Str = "FILETIME"
            Case adGUID:                DataTypeEnum2Str = "GUID"
            Case adIDispatch:           DataTypeEnum2Str = "IDISPATCH"
            Case adInteger:             DataTypeEnum2Str = "INTEGER"
            Case adIUnknown:            DataTypeEnum2Str = "IUNKNOWN"
            Case adLongVarBinary:       DataTypeEnum2Str = "LONGVARBINARY"
            Case adLongVarChar:         DataTypeEnum2Str = "LONGVARCHAR"
            Case adLongVarWChar:        DataTypeEnum2Str = "LONGVARWCHAR"
            Case adNumeric:             DataTypeEnum2Str = "NUMERIC"
            Case adPropVariant:         DataTypeEnum2Str = "PROPVARIANT"
            Case adSingle:              DataTypeEnum2Str = "SINGLE"
            Case adSmallInt:            DataTypeEnum2Str = "SMALLINT"
            Case adTinyInt:             DataTypeEnum2Str = "TINYINT"
            Case adUnsignedBigInt:      DataTypeEnum2Str = "UNSIGNEDBIGINT"
            Case adUnsignedInt:         DataTypeEnum2Str = "UNSIGNEDINT"
            Case adUnsignedSmallInt:    DataTypeEnum2Str = "UNSIGNEDSMALLINT"
            Case adUnsignedTinyInt:     DataTypeEnum2Str = "UNSIGNEDTINYINT"
            Case adUserDefined:         DataTypeEnum2Str = "USERDEFINED"
            Case adVarBinary:           DataTypeEnum2Str = "VARBINARY"
            Case adVarChar:             DataTypeEnum2Str = "VARCHAR"
            Case adVariant:             DataTypeEnum2Str = "VARIANT"
            Case adVarNumeric:          DataTypeEnum2Str = "VARNUMERIC"
            Case adVarWChar:            DataTypeEnum2Str = "VARWCHAR"
            Case adWChar:               DataTypeEnum2Str = "WCHAR"
        End Select
    End Function
