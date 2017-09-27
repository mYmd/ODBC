Attribute VB_Name = "VH_Tips"
'VH_Tips
'Copyright (c) 2017 mmYYmmdd
Option Explicit

'**********************************
' よく使いそうなパターン集
'**********************************

'
Sub sample_sequence_by_group()
    printM "------- sample_sequence_by_group ----------"
    Dim m As Variant
    m = Array(9, 9, 9, 9, 2, 2, 2, 3, 3, 3, 3, 3, 3, 3, 3, 8, 8, 5, 5, 5, 5, 5, 5, 5)
    printM m
    printM sequence_by_group_imple(m)
End Sub
    '
    Function sequence_by_group_imple(ByRef m As Variant, _
                                     Optional ByRef equal_fun As Variant) As Variant
        Dim pp As Variant
        '隣り合う要素を等値比較
        pp = self_zipWith(IIf(is_bindFun(equal_fun), equal_fun, p_equal), m, -1)
        pp(0) = 1
        sequence_by_group_imple = scanl1(p_plus(p_mult(ph_1, ph_2), 1), pp)
    End Function

'-------------------------------------------------------------
Sub sample_fill_ditto()
    Debug.Print "---------------------------------------."
    Dim pref As Variant
    Debug.Print "「 〃 」が出てきたら前の値を引き継ぐサンプル"
    pref = Array("大阪", "〃", _
                    "宮崎", "〃", "〃", "〃", _
                        "東京", "〃", "〃", _
                            "新潟", "〃", _
                        "愛媛", "〃", _
                    "岡山", "〃", "〃", _
                "沖縄", "〃")
    printM pref
    printM "         ↓ ↓ ↓"
    printM fill_ditto_imple(pref, "〃")
End Sub

    Function fill_ditto_imple(ByRef v As Variant, ByRef ditto As Variant) As Variant
        Dim tmp As Variant
        ' ditto をいったんEmpty に置き換える
        tmp = mapF(p_try(p_equal(ditto), Empty), v)
        ' 第1引数と第2引数をひっくり返しつつ replaceEmpty をスキャンする
        fill_ditto_imple = scanl1(p_replaceEmpty(yield_2, yield_1), tmp)
    End Function

' Replace関数のオブジェクト化はその場で作れる
Public Function make_p_replace(ByVal from_ As String, ByVal to_ As String) As Variant
    make_p_replace = p_join(p_split(, from_), to_)
End Function

' http://home.b07.itscom.net/m-yamada/VBA/ から.cpp ファイルを落としてくる
Sub sample_downloadCppFiles()
    Dim s As String, txt As String
    s = "http://home.b07.itscom.net/m-yamada/VBA/"
    txt = Y_IO_utiliy.getURLText(s, "", "UTF-8")
    Dim files As Variant
    files = Y_IO_utiliy.getTagsFromHTML(txt, "A", "href")
    files = filter_if(p_Like(, "*.cpp"), files)
    Call Y_IO_utiliy.downloadFiles(vector(files), "C:\tmp")     ' フォルダは適宜
End Sub

#If 0 Then
' vb_ODBC のサンプル
Sub sumple_ODBC()
    Dim oo As vb_ODBC
    Set oo = New vb_ODBC
    If Not oo.connect(oo.mdb_expr("C:\売上管理3.mdb")) Then
        Exit Sub
    End If
    Dim tableList As Variant
    tableList = oo.tableList()
    printS tableList
    printM tableList
        'MSysAccessObjects     SYSTEM TABLE
        'MSysAccessXML         SYSTEM TABLE
        '・・・・・・・・・・・・・
        '社員テーブル Table
    Dim tableAttribute As Variant
    tableAttribute = oo.tableAttribute("", "社員テーブル")
    printS tableAttribute
    printM tableAttribute
        '社員ID        INTEGER   10  -1
        '氏名          VARCHAR   50  -1
        '・・・・・・・・・・・・・
        '  データ更新日  DATETIME  19  -1
    Dim m As Variant
    m = oo.select_flat("SELECT * FROM 社員テーブル").get_data
    printS m
    printM m
        '901004  川村 匡      カワム
        '・・・・・・・・・・・・・
        '870001  小
    oo.disconnect
End Sub
#End If

'本日のTTM
Function MUFG_TTM_LAST(Optional ByVal currencyName = "USD") As Variant
    Dim u As String
    u = "http://www.bk.mufg.jp/gdocs/kinri/list_j/kinri/kawase.html"
    Dim s As Variant, d As Variant
    s = Y_IO_utiliy.getURLText(u, "")
    d = Y_IO_utiliy.getTagsFromHTML(s, "td", "it")
    Dim a As Long, b As Long
    a = find_pred(p_Like(, "最終更新日時*"), d)
    b = find_pred(p_Like(, currencyName & "*"), d)
    Dim dt As Date
    dt = CDate(Right(d(a), Len(d(a)) - InStr(d(a), "：")))
    Dim tts As Currency, ttb As Currency
    tts = CCur(d(b + 1))
    ttb = CCur(d(b + 4))
    MUFG_TTM_LAST = VBA.Array(dt, tts, ttb, (tts + ttb) / 2)
End Function

' RangeオブジェクトのUnion
Function range_union(ByRef range1 As Variant, ByRef range2 As Variant) As Variant
    If range1 Is Nothing Then
        Set range_union = range2
    ElseIf range2 Is Nothing Then
        Set range_union = range1
    Else
        Set range_union = range1.Parent.Parent.Parent.Union(range1, range2)
    End If
End Function
    Function p_range_union(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_range_union = make_funPointer(AddressOf range_union, firstParam, secondParam)
    End Function
    
' RangeオブジェクトのIntersect
Function range_intersect(ByRef range1 As Variant, ByRef range2 As Variant) As Variant
    If range1 Is Nothing Or range2 Is Nothing Then
        Set range_intersect = Nothing
    Else
        Set range_intersect = range1.Parent.Parent.Parent.Intersect(range1, range2)
    End If
End Function
    Function p_range_intersect(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_range_intersect = make_funPointer(AddressOf range_intersect, firstParam, secondParam)
    End Function
    
' 文字列全体にわたるInStr
Function instr_thr(ByRef sourceString As Variant, ByRef targetString As Variant) As Variant
    Dim k As Long:  k = 0
    instr_thr = VBA.Array()
    Do
        k = k + 1
        k = InStr(k, sourceString, targetString)
        If 0 < k Then
            Call push_back(instr_thr, k)
        Else
            Exit Do
        End If
    Loop
End Function
    Function p_instr_thr(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_instr_thr = make_funPointer(AddressOf instr_thr, firstParam, secondParam)
    End Function

' HTML中の about: をドメイン名に置換する関数を返す
Function make_p_replace_about(ByVal targetString1 As String, ByVal url As String) As Variant
    Dim ins As Variant
    ins = instr_thr(targetString1, "/")
    If 2 <= sizeof(ins) Then
        If ins(0) = Len("about:/") And ins(0) + 1 < ins(1) Then
            Dim substr As String
            substr = Mid(targetString1, ins(0), ins(1) - ins(0) + 1)
            Dim k As Long
            k = InStr(url, substr)
            If 0 < k Then
                make_p_replace_about = make_p_replace("about:/", Left(url, k))
            End If
        End If
    End If
End Function

' 1次元配列のランキング作成
'（同順位が複数いても順位に隙間は空けない、0スタート）
Function make_rank(ByRef m As Variant, Optional ByRef compFun As Variant) As Variant
    Dim pred As Variant, si As Variant
    pred = IIf(is_bindFun(compFun), compFun, p_less)
    si = sortIndex_pred(m, pred)
    pred(1) = ph_2: pred(2) = ph_1  ' 引数逆転
    make_rank = self_zipWith(pred, subV(m, si), -1)
    make_rank = scanl1(p_plus, make_rank)
    permutate_back make_rank, si
End Function
    Function p_make_rank(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_make_rank = make_funPointer(AddressOf make_rank, firstParam, secondParam, 2)
    End Function

