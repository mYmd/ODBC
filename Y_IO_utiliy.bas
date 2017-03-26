Attribute VB_Name = "Y_IO_utiliy"
'Y_IO_utility
'Copyright (c) 2016 mmYYmmdd
Option Explicit

'*********************************************************************************
'   IO関連ユーティリティ
'*********************************************************************************
'   Function    sheet2m             Excelシートのセル範囲から配列を取得
'   Sub         m2sheet             配列をExcelシートのセル範囲にペースト
'   Function    getRangeMatrix      Excelシートのセル範囲からRangeオブジェクトの配列を取得
'   Function    getInterior         オブジェクトのInteriorプロパティを取得
'   Function    getTextFile         テキストファイルの配列読み込み
'   Sub         writeTextFile       配列のテキストファイル書き込み
'   Function    getURLText          URLで指定されたテキストの配列読み込み
'   Function    urlEncode           URLエンコード
'   Function    urlDecode           URLデコード
'   Sub         m2Clip              配列（2次元以下）をクリップボードに転送する
'   Function    HTMLDocFromText     HTMLテキストからのHTMLDocumentオブジェクト
'   Function    getTagsFromHTML     HTMLテキストからのタグ抽出
'*********************************************************************************

' Excelシートのセル範囲から配列を取得（値のみ）
' vec = True：1次元配列化
' vec = Fale：2次元配列（デフォルト）
' LBound = 0 の配列となる
Public Function sheet2m(ByVal r As Object, Optional ByVal vec As Boolean = False) As Variant
    If Application.name Like "*Excel*" And TypeName(r) = "Range" Then
        If r.cells.count = 1 Then
            sheet2m = makeM(1, 1, r.value)
        Else
            sheet2m = r.value
        End If
        If vec Then sheet2m = vector(sheet2m)
        changeLBound sheet2m, 0
    End If
End Function

' 配列をExcelシートのセル範囲にペースト（左上のセルを指定）
' vertical = True：1次元配列を縦にペーストする
Public Sub m2sheet(ByRef matrix As Variant, _
                   ByVal r As Object, _
                   Optional ByVal vertical As Boolean = False)
    If Application.name = "Microsoft Excel" And TypeName(r) = "Range" Then
        Select Case Dimension(matrix)
        Case 0:
            r.value = matrix
        Case 1:
            If vertical Then
                r.Resize(sizeof(matrix), 1).value = transpose(matrix)
            Else
                r.Resize(1, sizeof(matrix)).value = matrix
            End If
        Case 2:
            r.Resize(rowSize(matrix), colSize(matrix)).value = matrix
        End Select
    End If
End Sub

' Excelシートのセル範囲からRangeオブジェクトの配列を取得
Public Function getRangeMatrix(ByVal r As Object) As Variant
    If Application.name = "Microsoft Excel" And TypeName(r) = "Range" Then
        Dim i As Long, j As Long, ret As Variant
        With r
            ret = makeM(.rows.count, .Columns.count)
            For i = 0 To rowSize(ret) - 1 Step 1
                For j = 0 To colSize(ret) - 1 Step 1
                    Set ret(i, j) = .cells(i + 1, j + 1)
                Next j
            Next i
        End With
    swapVariant getRangeMatrix, ret
    End If
End Function

' オブジェクトのInteriorプロパティを取得
Public Function getInterior(ByRef Ob As Variant, ByRef dummuy As Variant) As Variant
    Set getInterior = Ob.Interior
End Function
    Public Function p_getInterior(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_getInterior = make_funPointer(AddressOf getInterior, firstParam, secondParam)
    End Function

' テキストファイルの配列読み込み
' Charsetはshift-jisは明示的に指定しないとダメ
' head_n : 試し読み先頭行数指定
Public Function getTextFile(ByVal fileName As String, _
                            Optional ByVal line_end As String = vbCrLf, _
                            Optional ByVal Charset As String = "_autodetect_all", _
                            Optional ByVal head_n As Long = -1) As Variant
    Dim ado As Object:  Set ado = CreateObject("ADODB.Stream")
    On Error GoTo closeAdoStream
    Dim i As Long
    Dim lineS As String
    With ado
        .Open
        .Position = 0
        .Type = 2    'ADODB.Stream.adTypeText
        .Charset = Charset
        .LoadFromFile fileName
        If 0 < head_n Then
            .LineSeparator = IIf(line_end = vbCr, 13, IIf(line_end = vbLf, 10, -1))
            For i = 1 To head_n
                lineS = .ReadText(-2)   'adReadLine
                getTextFile = getTextFile & lineS & line_end
            Next i
        Else
            getTextFile = .ReadText
        End If
    End With
closeAdoStream:
    ado.Close
    Set ado = Nothing
    If 0 < Len(line_end) And VarType(getTextFile) = vbString Then
        getTextFile = Split(getTextFile, line_end)
    End If
End Function

' 配列のテキストファイル書き込み
Public Sub writeTextFile(ByRef data As Variant, _
                        ByVal fileName As String, _
                        ByVal Charset As String, _
                        Optional ByVal line_end As String = vbCrLf)
    Dim ado As Object:  Set ado = CreateObject("ADODB.Stream")
    On Error GoTo closeAdoStream
    Dim i As Long
    Dim lineS As String
    With ado
        .Open
        .Position = 0
        .Type = 2    'ADODB.Stream.adTypeText
        .Charset = Charset
        .LineSeparator = IIf(line_end = vbCr, 13, IIf(line_end = vbLf, 10, -1))
        For i = LBound(data) To UBound(data) Step 1
            .WriteText data(i), 1       ' adWriteLine
        Next i
        .SaveToFile fileName, 2 ' adSaveCreateOverWrite
    End With
closeAdoStream:
    ado.Close
    Set ado = Nothing
End Sub

' URLで指定されたテキストの配列読み込み
' head_n : 試し読み先頭行数指定
Public Function getURLText(ByVal url As String, _
                           Optional ByVal line_end As String = vbCrLf, _
                           Optional ByVal Charset As String = "_autodetect_all", _
                           Optional ByVal head_n As Long = -1) As Variant
    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")
    On Error GoTo closeObjects
    http.Open "GET", url, False
    http.Send
    Dim ado As Object:  Set ado = CreateObject("ADODB.Stream")
    Dim i As Long
    Dim lineS As String
    With ado
        .Open
        .Position = 0
        .Type = 1       'ADODB.Stream.adTypeBinary
        .Write http.responseBody
        .Position = 0
        .Type = 2       'ADODB.Stream.adTypeText
        .Charset = Charset
        If 0 < head_n Then
            .LineSeparator = IIf(line_end = vbCr, 13, IIf(line_end = vbLf, 10, -1))
            For i = 1 To head_n
                lineS = .ReadText(-2)   'adReadLine
                getURLText = getURLText & lineS & line_end
            Next i
        Else
            getURLText = .ReadText
        End If
    End With
closeObjects:
    If Not ado Is Nothing Then ado.Close
    Set ado = Nothing
    Set http = Nothing
    If 0 < Len(line_end) And VarType(getURLText) = vbString Then
        getURLText = Split(getURLText, line_end)
    End If
End Function

' URLエンコード（参考実装）
Public Function urlEncode(ByVal s As String) As String
    Dim ado As Object
    Dim tmp As Variant
    tmp = mapF(p_mid(s), zip(iota(1, Len(s)), repeat(1, Len(s))))
    Set ado = CreateObject("ADODB.Stream")
    ado.Charset = "UTF-8"
    tmp = mapF(p_urlEncode_1(, ado), tmp)
    Set ado = Nothing
    urlEncode = join(tmp, "")
End Function

' URLデコード（参考実装）
Public Function urlDecode(ByVal s As String) As String
    If s Like "*%??%??%??*" Then
        Dim begin As Long, theNext As Long
        begin = 1
        Dim ado As Object
        Set ado = CreateObject("ADODB.Stream")
        Do While begin <= Len(s) And mid(s, begin) Like "*%??%??%??*"
            If mid(s, begin, 9) Like "*%??%??%??*" Then
                urlDecode = urlDecode & urlDecode_imple(mid(s, begin, 9), ado)
                begin = begin + 9
            Else
                theNext = InStr(begin + 1, s, "%")
                If 0 < theNext Then
                    urlDecode = urlDecode & mid(s, begin, theNext - begin)
                    begin = theNext
                Else
                    urlDecode = urlDecode & mid(s, begin)
                    begin = Len(s) + 1
                End If
            End If
        Loop
        Set ado = Nothing
        If begin < Len(s) Then
            urlDecode = urlDecode & mid(s, begin)
        End If
    Else
        urlDecode = s
    End If
End Function

    Private Function isKanaKanji(ByVal s As String) As Boolean
        isKanaKanji = False
        If 0 < Len(s) Then
            If left(s, 1) Like "[ｦ-ﾟ]" Then
                isKanaKanji = True
            ElseIf asc(left(s, 1)) < 0 Then
                isKanaKanji = True
            End If
        End If
    End Function

    ' http://stabucky.com/wp/archives/4237
    Private Function urlEncode_1(ByRef s As Variant, ByRef adodbStream As Variant) As Variant
        Dim chars() As Byte
        If isKanaKanji(s) Then
            With adodbStream
                .Open
                .Type = 2       'ADODB.Stream.adTypeText
                .Position = 0
                .WriteText left(s, 1)
                .Position = 0
                .Type = 1       'ADODB.Stream.adTypeBinary
                .Position = 3
                chars = .Read
                .Close
                urlEncode_1 = "%" & Hex(chars(0)) & "%" & Hex(chars(1)) & "%" & Hex(chars(2))
            End With
        Else
            urlEncode_1 = s
        End If
    End Function
    Private Function p_urlEncode_1(Optional ByRef firstParam As Variant, Optional ByRef secondParam As Variant) As Variant
        p_urlEncode_1 = make_funPointer(AddressOf urlEncode_1, firstParam, secondParam)
    End Function


    Private Function urlDecode_imple(ByVal code As String, ByRef adodbStream As Object) As String
        Dim chars(0 To 2) As Byte
        chars(0) = CLng("&H" & mid(code, 2, 2))
        chars(1) = CLng("&H" & mid(code, 5, 2))
        chars(2) = CLng("&H" & mid(code, 8, 2))
        With adodbStream
            .Open
            .Type = 1       'ADODB.Stream.adTypeBinary
            .Position = 0
            .Write chars
            .Position = 0
            .Type = 2       'ADODB.Stream.adTypeText
            .Charset = "UTF-8"
            urlDecode_imple = .ReadText
            .Close
        End With
    End Function

' 配列（2次元以下）をクリップボードに転送する
Public Sub m2Clip(ByRef data As Variant)
    Dim s As String
    Select Case Dimension(data)
    Case 0
        s = CStr_(data) & vbCrLf
    Case 1
        s = join(mapF(p_CStr, data), vbTab) & vbCrLf
    Case 2
        Dim tmp As Variant
        tmp = zipC(mapF(p_CStr, data))
        tmp = mapF(p_join(, vbTab), tmp)
        s = join(tmp, vbCrLf) & vbCrLf
    End Select
    Dim dOb As Object
    'Set dOb = New MSFORMS.DataObject
    Set dOb = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    With dOb
        .SetText s
        .PutInClipboard
    End With
    Set dOb = Nothing
End Sub

' HTMLテキストからのHTMLDocumentオブジェクト
Public Function HTMLDocFromText(ByVal htmlText As String) As Object    'As HtmlDocument
    Set HTMLDocFromText = CreateObject("htmlfile") 'New MSHTML.HTMLDocument
    HTMLDocFromText.Write htmlText
End Function

' HTMLテキストからのタグ抽出
Public Function getTagsFromHTML(ByVal htmlText As String, ByRef tag As Variant) As Variant
    Dim doc As Object   'As HTMLDocument
    Set doc = HTMLDocFromText(htmlText)
    Dim z As Variant
    Dim ret As Variant: ret = VBA.Array()
    'Debug.Print doc.all.tags(tag).length
    For Each z In doc.all.tags(tag)
        push_back ret, z.innerText      'outerHTML
    Next z
    Set doc = Nothing
    swapVariant ret, getTagsFromHTML
End Function
