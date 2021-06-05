VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub 鉅亨網美股個股爬蟲(url As String, i As Integer, title As String, sheet_name As String)
    Dim httpObject As Object
    Dim jsonObject As Object
    Dim result As Variant
    Dim item As Variant
    Dim columnArray As Variant
    
    Set httpObject = CreateObject("MSXML2.XMLHTTP")
    
    httpObject.Open "GET", url, False
    httpObject.send
    
    ' 解析 JSON 格式
    result = httpObject.responseText
    Set jsonObject = JsonConverter.ParseJson(result)("data")("items")
    
    columnArray = Array("時間", "代碼", "中文名稱", "成交價", "漲跌", "漲幅(%)", "成交量", _
                         "市值(百萬)", "本益比", "殖利率(%)", "預估目標值", "預估卷商家數")
                         
    With ThisWorkbook.Worksheets(sheet_name)
        ' 填入標題和欄位名稱
        .Cells(i - 2, 1) = title
        .Range(.Cells(i - 1, 1), .Cells(i - 1, 12)) = columnArray
        
        ' 解決 Class Range 的 Select 方法失敗
        ThisWorkbook.Worksheets(sheet_name).Select
        ' 文字置中
        .Range("A:L").Select
        
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
        For Each item In jsonObject
            ' 時間(處理可能缺值)
            If IsNumeric(item("200007")) = True Then
                ' 將 Timestamp 轉換成 Date
                .Cells(i, 1) = DateAdd("s", item("200007") + 28800, "1/1/1970")
            Else
                .Cells(i, 1) = item("200007")
            End If
            ' 代碼
            .Cells(i, 2) = Split(item("0"), ":")(1)
            ' 中文名稱
            .Cells(i, 3) = item("200009")
            ' 成交價
            .Cells(i, 4) = Format(item("6"), "#.00")
            ' 漲跌
            .Cells(i, 5) = Format(item("11"), "#.00")
            ' 漲幅(%)
            .Cells(i, 6) = Format(item("56") / 100, "Percent")
            ' 成交量
            .Cells(i, 7) = item("800001")
            ' 市值(百萬)
            .Cells(i, 8) = item("700005")
            ' 本益比
            .Cells(i, 9) = Format(item("700001"), "#.0")
            ' 殖利率(%)
            .Cells(i, 10) = Format(item("700002") / 100, "Percent")
            ' 預估目標值
            .Cells(i, 11) = Format(item("700003"), "#.00")
            ' 預估卷商家數
            .Cells(i, 12) = item("700004")
            
            ' 根據漲跌著色
            If .Cells(i, 5) > 0 Then
                .Range(.Cells(i, 4), .Cells(i, 6)).Font.ColorIndex = 3
            ElseIf Cells(i, 5) < 0 Then
                .Range(.Cells(i, 4), .Cells(i, 6)).Font.ColorIndex = 10
            End If
            
            i = i + 1

        Next item
        
        ' 調整欄寬
        .Columns.AutoFit

    End With
End Sub


Sub 鉅亨網美股ETF_指數爬蟲(url As String, i As Integer, title As String, sheet_name As String)
    Dim httpObject As Object
    Dim jsonObject As Object
    Dim result As Variant
    Dim item As Variant
    Dim columnArray As Variant
    
    Set httpObject = CreateObject("MSXML2.XMLHTTP")
    
    httpObject.Open "GET", url, False
    httpObject.send
    
    ' 解析 JSON 格式
    result = httpObject.responseText
    Set jsonObject = JsonConverter.ParseJson(result)("data")("items")

    columnArray = Array("時間", "代碼", "中文名稱", "成交價", "漲跌", "漲幅(%)")
    
    With ThisWorkbook.Worksheets(sheet_name)
        ' 填入標題和欄位名稱
        .Cells(i - 2, 1) = title
        .Range(.Cells(i - 1, 1), .Cells(i - 1, 6)) = columnArray
        
        ' 解決 Class Range 的 Select 方法失敗
        ThisWorkbook.Worksheets(sheet_name).Select
        ' 文字置中
        .Range("A:F").Select
        
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
    
        For Each item In jsonObject
            ' 時間(處理可能缺值)
            If IsNumeric(item("200007")) = True Then
                ' 將 Timestamp 轉換成 Date
                .Cells(i, 1) = DateAdd("s", item("200007") + 28800, "1/1/1970")
            Else
                .Cells(i, 1) = item("200007")
            End If
            ' 代碼
            .Cells(i, 2) = Split(item("0"), ":")(1)
            ' 中文名稱
            .Cells(i, 3) = item("200009")
            ' 成交價
            .Cells(i, 4) = Format(item("6"), "#.00")
            ' 漲跌
            .Cells(i, 5) = Format(item("11"), "#.00")
            ' 漲幅(%)
            .Cells(i, 6) = Format(item("56") / 100, "Percent")
            
            ' 根據漲跌著色
            If .Cells(i, 5) > 0 Then
                .Range(.Cells(i, 4), .Cells(i, 6)).Font.ColorIndex = 3
            ElseIf .Cells(i, 5) < 0 Then
                .Range(.Cells(i, 4), .Cells(i, 6)).Font.ColorIndex = 10
            End If
            
            i = i + 1

        Next item
        
        ' 調整欄寬
        .Columns.AutoFit

    End With
End Sub


Sub 鉅亨網美國公債殖利率爬蟲(url As String, i As Integer, title As String, sheet_name As String)
    Dim httpObject As Object
    Dim jsonObject As Object
    Dim result As Variant
    Dim item As Variant
    Dim columnArray As Variant
    
    Set httpObject = CreateObject("MSXML2.XMLHTTP")
    
    httpObject.Open "GET", url, False
    httpObject.send
    
    ' 解析 JSON 格式
    result = httpObject.responseText
    Set jsonObject = JsonConverter.ParseJson(result)("data")("items")

    columnArray = Array("時間", "代碼", "中文名稱", "成交價", "漲跌", "漲幅(%)")
    
    With ThisWorkbook.Worksheets(sheet_name)
        ' 填入標題和欄位名稱
        .Cells(i - 2, 1) = title
        .Range(.Cells(i - 1, 1), .Cells(i - 1, 6)) = columnArray
        
        ' 解決 Class Range 的 Select 方法失敗
        ThisWorkbook.Worksheets(sheet_name).Select
        ' 文字置中
        .Range("A:F").Select
        
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
    
        For Each item In jsonObject
            ' 時間(處理可能缺值)
            If IsNumeric(item("200007")) = True Then
                ' 將 Timestamp 轉換成 Date
                .Cells(i, 1) = DateAdd("s", item("200007") + 28800, "1/1/1970")
            Else
                .Cells(i, 1) = item("200007")
            End If
            ' 代碼
            .Cells(i, 2) = Split(item("0"), ":")(1)
            ' 中文名稱
            .Cells(i, 3) = item("200009")
            ' 現價
            .Cells(i, 4) = Format(item("21") / 100, "Percent")
            ' 漲跌
            .Cells(i, 5) = Format(item("11"), "#.00")
            ' 漲幅(%)
            .Cells(i, 6) = Format(item("56") / 100, "Percent")
            
            ' 根據漲跌著色
            If .Cells(i, 5) > 0 Then
                .Range(.Cells(i, 4), .Cells(i, 6)).Font.ColorIndex = 3
            ElseIf .Cells(i, 5) < 0 Then
                Range(.Cells(i, 4), .Cells(i, 6)).Font.ColorIndex = 10
            End If
            
            i = i + 1
        Next item
        
        ' 調整欄寬
        .Columns.AutoFit

    End With
End Sub


Sub 執行前10大明星股_前10大市值美股爬蟲()
    With ThisWorkbook.Worksheets("美股")
        ' 先清空資料
        .Range("A:L").Clear
        ' 前10大明星股
        Call 鉅亨網美股個股爬蟲("https://ws.api.cnyes.com/ws/api/v2/universal/quote?type=USTOP10&column=H", 3, "美股前10大明星股", "美股")
        ' 前10大市值
        Call 鉅亨網美股個股爬蟲("https://ws.api.cnyes.com/ws/api/v2/universal/quote?type=USMV10&column=H", 17, "美股前10大市值", "美股")
    End With
End Sub


Sub 執行道瓊30成分股爬蟲()
    With ThisWorkbook.Worksheets("道瓊30成分股")
        ' 先清空資料
        .Range("A:L").Clear
        Call 鉅亨網美股個股爬蟲("https://ws.api.cnyes.com/ws/api/v2/universal/quote?type=DJI30&column=H", 3, "道瓊30成分股", "道瓊30成分股")
    End With
End Sub


Sub 執行美股半導體龍頭爬蟲()
    With ThisWorkbook.Worksheets("美股半導體龍頭")
        ' 先清空資料
        .Range("A:L").Clear
        Call 鉅亨網美股個股爬蟲("https://ws.api.cnyes.com/ws/api/v2/universal/quote?type=USSEMI10&column=H", 3, "美股半導體龍頭", "美股半導體龍頭")
    End With
End Sub


Sub 執行台股_中概股ADR爬蟲()
    With ThisWorkbook.Worksheets("ADR")
        ' 先清空資料
        .Range("A:L").Clear
        ' 台股 ADR
        Call 鉅亨網美股個股爬蟲("https://ws.api.cnyes.com/ws/api/v2/universal/quote?type=TADR10&column=H", 3, "台股 ADR", "ADR")
        ' 中概股 ADR
        Call 鉅亨網美股個股爬蟲("https://ws.api.cnyes.com/ws/api/v2/universal/quote?type=CADR10&column=H", 16, "中概股 ADR", "ADR")
    End With
End Sub


Sub 執行美股ETF爬蟲()
    With ThisWorkbook.Worksheets("美股ETF")
        ' 先清空資料
        .Range("A:F").Clear
        ' 前10大美股 ETF
        Call 鉅亨網美股ETF_指數爬蟲("https://ws.api.cnyes.com/ws/api/v2/universal/quote?type=TOPETF&column=A", 3, "前10大美股 ETF", "美股ETF")
        ' 固定收益類美股 ETF
        Call 鉅亨網美股ETF_指數爬蟲("https://ws.api.cnyes.com/ws/api/v2/universal/quote?type=FIETF&column=A", 17, "固定收益類美股 ETF", "美股ETF")
        ' 商品類美股 ETF
        Call 鉅亨網美股ETF_指數爬蟲("https://ws.api.cnyes.com/ws/api/v2/universal/quote?type=PRODETF&column=A", 31, "商品類美股 ETF", "美股ETF")
    End With
End Sub


Sub 執行美股指數爬蟲()
    With ThisWorkbook.Worksheets("美股指數")
        ' 先清空資料
        .Range("A:F").Clear
        Call 鉅亨網美股ETF_指數爬蟲("https://ws.api.cnyes.com/ws/api/v2/universal/quote?type=USINDEX&column=A", 3, "美股指數", "美股指數")
    End With
End Sub


Sub 執行美國公債殖利率爬蟲()
    With ThisWorkbook.Worksheets("美國公債殖利率")
        ' 先清空資料
        .Range("A:F").Clear
        Call 鉅亨網美國公債殖利率爬蟲("https://ws.api.cnyes.com/ws/api/v2/universal/quote?type=USTB&column=A", 3, "美國公債殖利率", "美國公債殖利率")
    End With
End Sub


Private Sub Workbook_Open()
    Call 執行前10大明星股_前10大市值美股爬蟲
    Call 執行道瓊30成分股爬蟲
    Call 執行美股半導體龍頭爬蟲
    Call 執行台股_中概股ADR爬蟲
    Call 執行美股ETF爬蟲
    Call 執行美股指數爬蟲
    Call 執行美國公債殖利率爬蟲
    ThisWorkbook.Worksheets("美股").Select
End Sub



