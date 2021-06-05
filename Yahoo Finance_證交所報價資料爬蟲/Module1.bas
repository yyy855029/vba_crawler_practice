Attribute VB_Name = "Module1"
Sub 抓取Yahoo_Finance分鐘報價資料()
    Dim stockCode As String
    Dim url As String
    Dim httpObject As Object
    Dim jsonObject As Object
    Dim result As Variant
    Dim timeObject As Object
    Dim openObject As Object
    Dim highObject As Object
    Dim lowObject As Object
    Dim closeObject As Object
    Dim volumeObject As Object
    Dim columnArray As Variant
    Dim i As Integer
    Dim chartRange As Range
    Dim testChart As Object
    
    ' 股票代號欄位
    stockCode = Range("I1").Value
    url = "https://tw.stock.yahoo.com/_td-stock/api/resource/FinanceChartService.ApacLibraCharts;symbols=%5B%22" & stockCode & ".TW%22%5D;type=tick?bkt=%5B%22tw-qsp-exp-no2-1%22%2C%22test-es-module-production%22%2C%22test-portfolio-stream%22%5D&device=desktop&ecma=modern&feature=ecmaModern%2CshowPortfolioStream&intl=tw&lang=zh-Hant-TW&partner=none&prid=2h3pnulg7tklc&region=TW&site=finance&tz=Asia%2FTaipei&ver=1.2.902&returnMeta=true"
    
    Set httpObject = CreateObject("MSXML2.XMLHTTP")
    
    httpObject.Open "GET", url, False
    httpObject.send
    
    ' 解析 JSON 格式
    result = httpObject.responseText
    Set jsonObject = JsonConverter.ParseJson(result)("data")(1)("chart")
    Set timeObject = jsonObject("timestamp")
    Set openObject = jsonObject("indicators")("quote")(1)("open")
    Set highObject = jsonObject("indicators")("quote")(1)("high")
    Set lowObject = jsonObject("indicators")("quote")(1)("low")
    Set closeObject = jsonObject("indicators")("quote")(1)("close")
    Set volumeObject = jsonObject("indicators")("quote")(1)("volume")
    
    ' 清空數值
    Range("A:F").Clear
    
    ' 填入欄位名稱
    columnArray = Array("時間", "開盤價", "最高價", "最低價", "收盤價", "成交量")
    Range("A1:F1") = columnArray
    
    ' 文字置中
    Range("A:F").Select
        
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
    
    ' 填入資料
    For i = 1 To openObject.Count
        ' 時間(處理可能缺值)
        If IsNumeric(timeObject(i)) = True Then
            ' 將 Timestamp 轉換成 Date
            Cells(1 + i, "A") = DateAdd("s", timeObject(i) + 28800, "1/1/1970")
        Else
            Cells(1 + i, "A") = timeObject(i)
        End If
    
        Cells(1 + i, "B") = openObject(i)
        Cells(1 + i, "C") = highObject(i)
        Cells(1 + i, "D") = lowObject(i)
        Cells(1 + i, "E") = closeObject(i)
        Cells(1 + i, "F") = volumeObject(i)
        
        ' 根據漲跌著色
        If Cells(1 + i, "B").Value < Cells(1 + i, "E").Value Then
            Range(Cells(1 + i, "B"), Cells(1 + i, "F")).Font.ColorIndex = 3
        ElseIf Cells(1 + i, "B").Value > Cells(1 + i, "E").Value Then
            Range(Cells(1 + i, "B"), Cells(1 + i, "F")).Font.ColorIndex = 10
        End If
        
    Next i
    
    ' 若原先存在圖則刪除
    If ActiveSheet.ChartObjects.Count > 0 Then
        ActiveSheet.ChartObjects.Delete
    End If
    
    ' 設定繪圖範圍大小
    Set chartRange = Range("H8:P21")
    
    ActiveSheet.Shapes.AddChart2(201, _
                                xlColumnClustered, _
                                Left:=chartRange(1).Left, _
                                Top:=chartRange(1).Top, _
                                Width:=chartRange.Width, _
                                Height:=chartRange.Height).Select
    
    With ActiveChart
        ' X-Y折線圖資料範圍
        .SetSourceData Source:=Range("當日個股分鐘報價!$A$1:$A$272,當日個股分鐘報價!$E$1:$F$272")
        ' 價格折線圖
        With .FullSeriesCollection(1)
            '設定折線圖
            .ChartType = xlLine
            ' 更改線寬度
            .Format.Line.Weight = 1
        End With
        
        ' 成交量柱狀圖
        With .FullSeriesCollection(2)
            .AxisGroup = 2
            ' 設定柱狀圖
            .ChartType = xlColumnClustered
            ' 更改柱顏色
            .Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
            ' 更改柱透明度
            .Format.Fill.Transparency = 0
        End With
    
        ' 更改字型
        .ChartArea.Font.Name = "微軟正黑體"
        ' 刪除圖例
        .Legend.Delete
        
        ' 橫坐標軸
        With .Axes(xlCategory)
            .CategoryType = xlCategoryScale
            ' 更改時間顯示
            .TickLabels.NumberFormatLocal = "h:mm;@"
            ' 更改字體大小
            .TickLabels.Font.Size = 10
            ' 更改加粗體
            .TickLabels.Font.Bold = msoTrue
        End With
        
        ' 縱主座標軸
        With .Axes(xlValue).TickLabels
            ' 更改字體大小
            .Font.Size = 10
            ' 更改加粗體
            .Font.Bold = msoTrue
        End With
        
        ' 縱副座標軸
        With .Axes(xlValue, xlSecondary).TickLabels
            ' 更改字體大小
            .Font.Size = 10
            ' 更改加粗體
            .Font.Bold = msoTrue
        End With
        
        ' 標題
        With .ChartTitle
            ' 更改標題
            .Text = stockCode & " 當天收盤價和成交量"
            ' 更改字體大小
            .Font.Size = 14
            ' 更改加粗體
            .Font.Bold = msoTrue
        End With
             
    End With

    Range("I1").Select

End Sub


Sub 抓取證交所三大法人買賣超日報()
    Dim stockDate As String
    Dim url As String
    Dim httpObject As Object
    Dim jsonObject As Object
    Dim result As Variant
    Dim dataObject As Object
    Dim columnObject As Object
    Dim i As Integer
    Dim j As Integer

    ' 日期欄位
    stockDate = Format(Range("B1").Value, "yyyyMMdd")
    url = "https://www.twse.com.tw/fund/T86?response=json&date=" & stockDate & "&selectType=ALLBUT0999"
    
    Set httpObject = CreateObject("MSXML2.XMLHTTP")
    
    httpObject.Open "GET", url, False
    httpObject.send

    ' 解析 JSON 格式
    result = httpObject.responseText
    Set jsonObject = JsonConverter.ParseJson(result)
    Set dataObject = jsonObject("data")
    Set columnObject = jsonObject("fields")
    
    ' 清空數值
    Range("A5:S10000").Clear
    
    ' 文字置中
    Range("A:S").Select
        
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
    
    ' 填入欄位名稱
    For i = 1 To columnObject.Count
        Cells(4, i) = columnObject(i)
    Next i
    
    ' 填入資料
    ' 每筆證卷
    For i = 1 To dataObject.Count
        ' 每個欄位
        For j = 1 To dataObject(i).Count
            If j >= 1 And j <= 2 Then
                ' 前 2 欄證卷代號和名稱處理
                Cells(i + 4, j) = Replace(CStr(dataObject(i)(j)), " ", "")
            Else
                ' 第 3 欄以後數字資料處理
                Cells(i + 4, j) = CDbl(dataObject(i)(j))
            End If
        Next j
    Next i
    
    ' 根據三大法人買賣超股數排序
    ActiveWorkbook.Worksheets("三大法人買賣超").Range("A4:S1090").Select
    ActiveWorkbook.Worksheets("三大法人買賣超").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("三大法人買賣超").Sort.SortFields.Add2 Key:=Range("S5:S1090") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("三大法人買賣超").Sort
        .SetRange Range("A4:S1090")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' 自動調整欄寬
    Columns.AutoFit
    Range("B1").Select

End Sub























