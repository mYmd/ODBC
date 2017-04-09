VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vb_ODBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' vb_ODBC
'Copyright (c) 2016 mmYYmmdd

'*********************************************************************************
' ODBC接続
' --------------------------------------------------------------------------------
'   Function        connect             接続文字列を発行してデータベースに接続
'   Sub             disconnect          接続断
'   Sub             disconnect_all      すべてのオブジェクトの接続を断
'   Sub             clear               ヘッダ（カラム）データとテーブルデータを消去
'   Property Get    last_connected      最後に成功した接続文字列の取得
'   Property Get    last_SQL            最後に発行されたSQL文字列の取得
'   Property Get    last_elapsed        最後に発行されたSELECT文でかかった時間 (ミリ秒)
'   Property Get    last_execError      最後に発行されたexec文のエラー行
'   Function        tableList           テーブル一覧取得
'   Function        tableAttribute      テーブルの Attribute
'   Function        read_header         Select文を発行してヘッダ（カラム）情報をメンバに格納する
'   Function        get_header          メンバに格納されたヘッダ（カラム）情報を取得する
'   Function        select_flat         Select文を発行して2次元配列としてメンバに格納する
'   Function        select_rowWise      Select文を発行して行の配列としてメンバに格納する
'   Function        select_columnWise   Select文を発行して列の配列としてメンバに格納する
'   Function        exec                insert文やdelete文を実行する
'   Function        rowSize             テーブルのレコード数
'   Function        colSize             テーブルのカラム数
'   Function        get_data            Select したテーブルデータのコピーを返す
'   Function        move_data           Select したテーブルデータをmoveして返す
'   Function        insert_expr         insert文
'   Function        insert_expr_        insert文（jag配列または2次元配列）
'   Function        SQLDrivers          インストールされているODBCドライバ配列を返す
'   Function        excel_expr          Excelブックに接続するための接続文字列
'   Function        mdb_expr            accessdbの接続文字列
'   Function        sqlServer_expr      SQL Serverの接続文字列
'*********************************************************************************

Private Declare PtrSafe Sub terminateODBC Lib "vbODBC.dll" (ByVal n As Long)

Private Declare PtrSafe Sub terminateODBC_all Lib "vbODBC.dll" ()

Private Declare PtrSafe Function initODBC Lib "vbODBC.dll" ( _
                                        ByRef n As Long, _
                                ByRef rawStr As Variant) As Long

Private Declare PtrSafe Function selectODBC Lib "vbODBC.dll" ( _
                                    ByVal n As Long, _
                            ByRef SQL As Variant, _
                        Optional ByVal timeOutSec As Long = 0) As Variant

Private Declare PtrSafe Function selectODBC_rowWise Lib "vbODBC.dll" ( _
                                    ByVal n As Long, _
                                ByRef SQL As Variant, _
                            Optional ByVal timeOutSec As Long = 0) As Variant

Private Declare PtrSafe Function selectODBC_columnWise Lib "vbODBC.dll" ( _
                                    ByVal n As Long, _
                                ByRef SQL As Variant, _
                            Optional ByVal timeOutSec As Long = 0) As Variant

Private Declare PtrSafe Function columnAttributes Lib "vbODBC.dll" _
                                    (ByVal n As Long, _
                                ByRef SQL As Variant, _
                            ByVal getNullable As Long) As Variant

Private Declare PtrSafe Function execODBC Lib "vbODBC.dll" _
                                    (ByVal n As Long, _
                                ByRef SQLs As Variant) As Variant

Private Declare PtrSafe Function table_list_all Lib "vbODBC.dll" ( _
                                    ByVal n As Long, _
                                ByRef schemaName As Variant) As Variant

Private Declare PtrSafe Function columnAttributes_all Lib "vbODBC.dll" ( _
                                    ByVal n As Long, _
                                ByRef schemaName As Variant, _
                            ByRef tableName As Variant) As Variant

Private Declare PtrSafe Function SQLGetInstalledDrivers Lib "odbccp32.dll" ( _
                                            ByVal buf As String, _
                                        ByVal cbBufMax As Integer, _
                                    ByRef pcchBufOut As Integer) As Byte

Private Declare PtrSafe Function GetTickCount Lib "kernel32.dll" () As Long

Private connectionString    As String    ' 接続文字列
Private mySQL_              As String
Private myheader_           As Variant
Private mydata_             As Variant
Private myErrorNo_          As Variant
Private myNo_               As Long
Private laptime_            As Long

Private Sub Class_Initialize()
    myNo_ = -1
End Sub

Private Sub Class_Terminate()
    disconnect
    clear
End Sub

' 接続文字列を発行してデータベースに接続
Public Function connect(ByVal connectionExpr As String) As Boolean
    connect = (0 <= initODBC(myNo_, connectionExpr))
    If connect Then connectionString = connectionExpr
End Function

' 接続断
Public Sub disconnect()
    Call terminateODBC(myNo_)
End Sub

' すべてのオブジェクトを接続断
Public Sub disconnect_all()
    Call terminateODBC_all
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

' テーブル一覧取得
Function tableList(ByRef schemaName As Variant, _
                        Optional ByVal type_ As String = "TABLE") As Variant
    Dim ret As Variant
    ret = table_list_all(myNo_, schemaName)
    If IsEmpty(ret(0)) Then Exit Function
    If type_ <> "*" Then
        Dim fil As Variant
        fil = mapF(p_Like(, type_), ret(1))
        tableList = catC(filterR(ret(0), fil), filterR(ret(1), fil))
    Else
        tableList = catC(ret(0), ret(1))
    End If
End Function

' テーブルの Attribute
' 列名、型名、データ長、精度、属性（1:キー, -1:Nullable, 0:他）
Public Function tableAttribute(ByRef schemaName As Variant, _
                            ByRef tableName As Variant) As Variant
    Dim tmp As Variant
    tmp = columnAttributes_all(myNo_, schemaName, tableName)
    If IsEmpty(tmp(0)) Then Exit Function
    tmp(2) = mapF(p_CLng, tmp(2))
    tmp(3) = mapF(p_mult(-1, p_CLng), tmp(3))
    If 0 < sizeof(tmp(5)) Then
        tmp(3) = zipWith(p_plus, tmp(3), foldl1(p_max, product_set(p_equal, tmp(0), tmp(5)), 2))
    End If
    Dim s As Variant:   s = sortIndex(mapF(p_CLng, tmp(4)))
    tableAttribute = subM(foldl1(p_catC, move_many(tmp(0), tmp(1), tmp(2), tmp(3))), s)
End Function

' Select文を発行してヘッダ（カラム）情報をメンバに格納する
Public Function read_header(ByVal sqlExpr As String) As vb_ODBC
    mySQL_ = sqlExpr
    laptime_ = GetTickCount()
    myheader_ = subM(columnAttributes(myNo_, sqlExpr, 0), , Array(0, 1))
    laptime_ = GetTickCount() - laptime_
    Set read_header = Me
End Function

' メンバに格納されたヘッダ（カラム）情報を取得する
Public Function get_header(Optional ByVal transpose_ As Boolean = False) As Variant
    If transpose_ Then
        get_header = transpose(myheader_)
    Else
        get_header = myheader_
    End If
End Function

' Select文を発行して2次元配列としてメンバに格納する
Public Function select_flat(ByVal sqlExpr As String, Optional ByVal timeOutSec As Long = 0) As vb_ODBC
    mySQL_ = sqlExpr
    read_header sqlExpr
    laptime_ = GetTickCount()
    mydata_ = selectODBC(myNo_, sqlExpr, timeOutSec)
    laptime_ = GetTickCount() - laptime_
    Set select_flat = Me
End Function

' Select文を発行して行の配列としてメンバに格納する
Public Function select_rowWise(ByVal sqlExpr As String, Optional ByVal timeOutSec As Long = 0) As vb_ODBC
    mySQL_ = sqlExpr
    read_header sqlExpr
    laptime_ = GetTickCount()
    mydata_ = selectODBC_rowWise(myNo_, sqlExpr, timeOutSec)
    laptime_ = GetTickCount() - laptime_
    Set select_rowWise = Me
End Function

' Select文を発行して列の配列としてメンバに格納する
Public Function select_columnWise(ByVal sqlExpr As String, Optional ByVal timeOutSec As Long = 0) As vb_ODBC
    mySQL_ = sqlExpr
    read_header sqlExpr
    laptime_ = GetTickCount()
    mydata_ = selectODBC_columnWise(myNo_, sqlExpr, timeOutSec)
    laptime_ = GetTickCount() - laptime_
    Set select_columnWise = Me
End Function

' insert文やdelete文を実行する
Public Function exec(ByRef sqlExpr As Variant) As vb_ODBC
    laptime_ = GetTickCount()
    If IsArray(sqlExpr) Then
        myErrorNo_ = execODBC(myNo_, sqlExpr)
    Else
        myErrorNo_ = execODBC(myNo_, VBA.Array(sqlExpr))
    End If
    laptime_ = GetTickCount() - laptime_
    Set exec = Me
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
    swapVariant insert_expr_, ret
End Function

' インストールされているODBCドライバ配列を返す
Public Function SQLDrivers(Optional ByVal buflen = 0) As Variant
    Dim buf As String, ret As Byte, retL As Integer
    buf = Space(IIf(0 < buflen, buflen, 2 * 1024))
    ret = SQLGetInstalledDrivers(buf, Len(buf), retL)
    SQLDrivers = Split(left(buf, retL - 2), vbNullChar)
End Function

' Excelブックに接続するための接続文字列
Public Function excel_expr(ByVal bookName As String) As String
    Dim drivers As Variant
    Dim i As Variant
    excel_expr = "Microsoft Excel Driver"
    drivers = SQLDrivers
    permutate drivers, reverse(sortIndex(mapF(p_len, drivers)))
    i = find_pred(p_equal(excel_expr, p_left(, Len(excel_expr))), drivers)
    If i <= UBound(drivers) Then
        excel_expr = "DBA=R; Driver={" & drivers(i) & "}; DBQ="
        excel_expr = excel_expr & bookName
    Else
        excel_expr = ""
    End If
    '               SELECT * FROM [Sheet1$A1:D10]
End Function

' accessdbの接続文字列
Public Function mdb_expr(ByVal dbName As String) As String
    Dim drivers As Variant
    Dim i As Variant
    mdb_expr = "DBA=R; Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ="
    mdb_expr = mdb_expr & dbName
    '               SELECT * FROM MSysNameMap
End Function

' SQL Serverの接続文字列
Public Function sqlServer_expr(ByVal server As String, _
                               ByVal dbName As String, _
                               Optional ByVal uid As String = "", _
                               Optional ByVal pwd As String = "") As String
    Dim drivers As Variant
    Dim i As Variant
    sqlServer_expr = "Driver={SQL Server Native Client 11.0};"
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
    '               SELECT * FROM sysobjects WHERE xtype = 'U'
End Function
