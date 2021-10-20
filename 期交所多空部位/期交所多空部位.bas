Attribute VB_Name = "Module1"
Sub 抓取期貨報價資料()
    Dim start_date As Date
    Dim end_date As Date
    Dim date_array As Variant
    Dim dataname_array As Variant
    Dim httpObject As Object
    Dim htmlDoc As Object
    Dim result As Variant
    Dim df As Variant
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    ' 改為手動更新
    Application.Calculation = xlCalculationManual
    
    ' 資料日期
    start_date = Worksheets("圖形").Range("C2").Value
    end_date = Worksheets("圖形").Range("D2").Value
    date_array = Array(start_date, end_date)
    
    ' 期貨三大法人
    dataname_array = Array("對照日期貨三大法人", "最新日期貨三大法人")
    
    For i = 0 To UBound(date_array) - LBound(date_array)
        ' 先清空資料
        Worksheets(dataname_array(i)).Cells.ClearContents
        url = "https://www.taifex.com.tw/cht/3/futContractsDate?queryType=1&doQuery=1&queryDate=" & Format(date_array(i), "yyyy/MM/dd")
        
        Set httpObject = CreateObject("MSXML2.XMLHTTP")
        Set htmlDoc = CreateObject("HtmlFile")
    
        httpObject.Open "GET", url, False
        httpObject.send
        
        result = httpObject.responsetext
        htmlDoc.body.innerhtml = result
        
        Set df = htmlDoc.getelementsbytagname("table")(3)
        j = 1
        For Each nrow In df.Rows
            k = 1
            
            For Each ncol In nrow.Cells
                Worksheets(dataname_array(i)).Cells(j, k) = Replace(ncol.innertext, vbCrLf, "")
                k = k + 1
            Next ncol
            j = j + 1
        Next nrow
    
    Next i
    
    ' 期貨報價
    dataname_array = Array("對照日期貨報價", "最新日期貨報價")
    
    For i = 0 To UBound(date_array) - LBound(date_array)
        ' 先清空資料
        Worksheets(dataname_array(i)).Cells.ClearContents
        url = "https://www.taifex.com.tw/cht/3/futDailyMarketReport?queryType=2&marketCode=1&commodity_id=TX&queryDate=" & Format(date_array(i), "yyyy/MM/dd") & "&MarketCode=1&commodity_idt=TX"
        
        Set httpObject = CreateObject("MSXML2.XMLHTTP")
        Set htmlDoc = CreateObject("HtmlFile")
    
        httpObject.Open "GET", url, False
        httpObject.send
        
        result = httpObject.responsetext
        htmlDoc.body.innerhtml = result
        
        Set df = htmlDoc.getelementsbytagname("table")(4)
        j = 1
        For Each nrow In df.Rows
            k = 1
            
            For Each ncol In nrow.Cells
                ' 排除換行處理
                Worksheets(dataname_array(i)).Cells(j, k) = Replace(ncol.innertext, vbCrLf, "")
                k = k + 1
            Next ncol
            j = j + 1
        Next nrow
    
    Next i
    
    ' 恢復自動更新
    Application.Calculation = xlCalculationAutomatic
    
End Sub


Sub 抓取選擇權報價資料()
    Dim start_date As Date
    Dim end_date As Date
    Dim date_array As Variant
    Dim dataname_array As Variant
    Dim httpObject As Object
    Dim htmlDoc As Object
    Dim result As Variant
    Dim df As Variant
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim check_date As String
    
    ' 改為手動更新
    Application.Calculation = xlCalculationManual
    
    ' 資料日期
    start_date = Worksheets("圖形").Range("C2").Value
    end_date = Worksheets("圖形").Range("D2").Value
    date_array = Array(start_date, end_date)
    
    ' 選擇權三大法人
    dataname_array = Array("對照日選擇權三大法人", "最新日選擇權三大法人")
    
    For i = 0 To UBound(date_array) - LBound(date_array)
        ' 先清空資料
        Worksheets(dataname_array(i)).Cells.ClearContents
        url = "https://www.taifex.com.tw/cht/3/callsAndPutsDate?queryType=1&doQuery=1&queryDate=" & Format(date_array(i), "yyyy/MM/dd")
        
        Set httpObject = CreateObject("MSXML2.XMLHTTP")
        Set htmlDoc = CreateObject("HtmlFile")
    
        httpObject.Open "GET", url, False
        httpObject.send
        
        result = httpObject.responsetext
        htmlDoc.body.innerhtml = result
        
        Set df = htmlDoc.getelementsbytagname("table")(3)
        j = 1
        For Each nrow In df.Rows
            k = 1
            
            For Each ncol In nrow.Cells
                ' 排除換行處理
                Worksheets(dataname_array(i)).Cells(j, k) = Replace(ncol.innertext, vbCrLf, "")
                k = k + 1
            Next ncol
            j = j + 1
        Next nrow
    
    Next i
    
    ' 選擇權報價
    dataname_array = Array("對照日選擇權報價", "最新日選擇權報價")
    
    For i = 0 To UBound(date_array) - LBound(date_array)
        ' 先清空資料
        Worksheets(dataname_array(i)).Cells.ClearContents
        url = "https://www.taifex.com.tw/cht/3/optDailyMarketReport?queryType=2&marketCode=0&commodity_id=TXO&queryDate=" & Format(date_array(i), "yyyy/MM/dd") & "&MarketCode=1&commodity_idt=TXO"
        
        Set httpObject = CreateObject("MSXML2.XMLHTTP")
        Set htmlDoc = CreateObject("HtmlFile")
    
        httpObject.Open "GET", url, False
        httpObject.send
        
        result = httpObject.responsetext
        htmlDoc.body.innerhtml = result
        
        Set df = htmlDoc.getelementsbytagname("table")(4)
        j = 1
        For Each nrow In df.Rows
            k = 1
            
            For Each ncol In nrow.Cells
                ' 排除換行處理
                Worksheets(dataname_array(i)).Cells(j, k) = Replace(ncol.innertext, vbCrLf, "")
                k = k + 1
            Next ncol
            j = j + 1
        Next nrow
    
    Next i
    
    ' 整理選擇權報價
    sort_dataname_array = Array("對照日整理選擇權報價", "最新日整理選擇權報價")
    
    For i = 0 To UBound(date_array) - LBound(date_array)
        ' 先清空資料
        Worksheets(sort_dataname_array(i)).Cells.ClearContents
        ' 欄位名稱
        Worksheets(sort_dataname_array(i)).Range("A1:I1") = Array("合約", "Call 履約價", "Call 成交", "Call 成交量", "Call 未平倉", "Put 履約價", "Put 成交", "Put 成交量", "Put 未平倉")
        
        ' 檢查是否為近月合約
        check_date = Worksheets(dataname_array(i)).Range("B2")
        
        For j = 2 To Worksheets(dataname_array(i)).Range("A1").End(xlDown).Row
            ' Call
            If j Mod 2 = 0 And check_date = Worksheets(dataname_array(i)).Cells(j, 2) Then
                ' 合約
                Worksheets(sort_dataname_array(i)).Cells(j \ 2 + 1, 1) = Worksheets(dataname_array(i)).Cells(j, 2)
                ' 履約價
                Worksheets(sort_dataname_array(i)).Cells(j \ 2 + 1, 2) = Worksheets(dataname_array(i)).Cells(j, 3)
                ' 成交
                Worksheets(sort_dataname_array(i)).Cells(j \ 2 + 1, 3) = Replace(Worksheets(dataname_array(i)).Cells(j, 8), "-", "")
                ' 成交量
                Worksheets(sort_dataname_array(i)).Cells(j \ 2 + 1, 4) = Worksheets(dataname_array(i)).Cells(j, 14)
                ' 未平倉
                Worksheets(sort_dataname_array(i)).Cells(j \ 2 + 1, 5) = Worksheets(dataname_array(i)).Cells(j, 15)
             
            ' Put
            ElseIf check_date = Worksheets(dataname_array(i)).Cells(j, 2) Then
                ' 履約價
                Worksheets(sort_dataname_array(i)).Cells(j \ 2 + 1, 6) = Worksheets(dataname_array(i)).Cells(j, 3)
                ' 成交
                Worksheets(sort_dataname_array(i)).Cells(j \ 2 + 1, 7) = Replace(Worksheets(dataname_array(i)).Cells(j, 8), "-", "")
                ' 成交量
                Worksheets(sort_dataname_array(i)).Cells(j \ 2 + 1, 8) = Worksheets(dataname_array(i)).Cells(j, 14)
                ' 未平倉
                Worksheets(sort_dataname_array(i)).Cells(j \ 2 + 1, 9) = Worksheets(dataname_array(i)).Cells(j, 15)
            Else
                Exit For
            End If
        Next j
    Next i
    
    ' 恢復自動更新
    Application.Calculation = xlCalculationAutomatic
    
End Sub


Sub 執行()
    Call 抓取期貨報價資料
    Call 抓取選擇權報價資料
End Sub
