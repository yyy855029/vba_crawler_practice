Attribute VB_Name = "Module1"
Sub ���Yahoo_Finance�����������()
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
    
    ' �Ѳ��N�����
    stockCode = Range("I1").Value
    url = "https://tw.stock.yahoo.com/_td-stock/api/resource/FinanceChartService.ApacLibraCharts;symbols=%5B%22" & stockCode & ".TW%22%5D;type=tick?bkt=%5B%22tw-qsp-exp-no2-1%22%2C%22test-es-module-production%22%2C%22test-portfolio-stream%22%5D&device=desktop&ecma=modern&feature=ecmaModern%2CshowPortfolioStream&intl=tw&lang=zh-Hant-TW&partner=none&prid=2h3pnulg7tklc&region=TW&site=finance&tz=Asia%2FTaipei&ver=1.2.902&returnMeta=true"
    
    Set httpObject = CreateObject("MSXML2.XMLHTTP")
    
    httpObject.Open "GET", url, False
    httpObject.send
    
    ' �ѪR JSON �榡
    result = httpObject.responseText
    Set jsonObject = JsonConverter.ParseJson(result)("data")(1)("chart")
    Set timeObject = jsonObject("timestamp")
    Set openObject = jsonObject("indicators")("quote")(1)("open")
    Set highObject = jsonObject("indicators")("quote")(1)("high")
    Set lowObject = jsonObject("indicators")("quote")(1)("low")
    Set closeObject = jsonObject("indicators")("quote")(1)("close")
    Set volumeObject = jsonObject("indicators")("quote")(1)("volume")
    
    ' �M�żƭ�
    Range("A:F").Clear
    
    ' ��J���W��
    columnArray = Array("�ɶ�", "�}�L��", "�̰���", "�̧C��", "���L��", "����q")
    Range("A1:F1") = columnArray
    
    ' ��r�m��
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
    
    ' ��J���
    For i = 1 To openObject.Count
        ' �ɶ�(�B�z�i��ʭ�)
        If IsNumeric(timeObject(i)) = True Then
            ' �N Timestamp �ഫ�� Date
            Cells(1 + i, "A") = DateAdd("s", timeObject(i) + 28800, "1/1/1970")
        Else
            Cells(1 + i, "A") = timeObject(i)
        End If
    
        Cells(1 + i, "B") = openObject(i)
        Cells(1 + i, "C") = highObject(i)
        Cells(1 + i, "D") = lowObject(i)
        Cells(1 + i, "E") = closeObject(i)
        Cells(1 + i, "F") = volumeObject(i)
        
        ' �ھں��^�ۦ�
        If Cells(1 + i, "B").Value < Cells(1 + i, "E").Value Then
            Range(Cells(1 + i, "B"), Cells(1 + i, "F")).Font.ColorIndex = 3
        ElseIf Cells(1 + i, "B").Value > Cells(1 + i, "E").Value Then
            Range(Cells(1 + i, "B"), Cells(1 + i, "F")).Font.ColorIndex = 10
        End If
        
    Next i
    
    ' �Y����s�b�ϫh�R��
    If ActiveSheet.ChartObjects.Count > 0 Then
        ActiveSheet.ChartObjects.Delete
    End If
    
    ' �]�wø�Ͻd��j�p
    Set chartRange = Range("H8:P21")
    
    ActiveSheet.Shapes.AddChart2(201, _
                                xlColumnClustered, _
                                Left:=chartRange(1).Left, _
                                Top:=chartRange(1).Top, _
                                Width:=chartRange.Width, _
                                Height:=chartRange.Height).Select
    
    With ActiveChart
        ' X-Y��u�ϸ�ƽd��
        .SetSourceData Source:=Range("���ӪѤ�������!$A$1:$A$272,���ӪѤ�������!$E$1:$F$272")
        ' �����u��
        With .FullSeriesCollection(1)
            '�]�w��u��
            .ChartType = xlLine
            ' ���u�e��
            .Format.Line.Weight = 1
        End With
        
        ' ����q�W����
        With .FullSeriesCollection(2)
            .AxisGroup = 2
            ' �]�w�W����
            .ChartType = xlColumnClustered
            ' ���W�C��
            .Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
            ' ���W�z����
            .Format.Fill.Transparency = 0
        End With
    
        ' ���r��
        .ChartArea.Font.Name = "�L�n������"
        ' �R���Ϩ�
        .Legend.Delete
        
        ' ��жb
        With .Axes(xlCategory)
            .CategoryType = xlCategoryScale
            ' ���ɶ����
            .TickLabels.NumberFormatLocal = "h:mm;@"
            ' ���r��j�p
            .TickLabels.Font.Size = 10
            ' ���[����
            .TickLabels.Font.Bold = msoTrue
        End With
        
        ' �a�D�y�жb
        With .Axes(xlValue).TickLabels
            ' ���r��j�p
            .Font.Size = 10
            ' ���[����
            .Font.Bold = msoTrue
        End With
        
        ' �a�Ʈy�жb
        With .Axes(xlValue, xlSecondary).TickLabels
            ' ���r��j�p
            .Font.Size = 10
            ' ���[����
            .Font.Bold = msoTrue
        End With
        
        ' ���D
        With .ChartTitle
            ' �����D
            .Text = stockCode & " ��Ѧ��L���M����q"
            ' ���r��j�p
            .Font.Size = 14
            ' ���[����
            .Font.Bold = msoTrue
        End With
             
    End With

    Range("I1").Select

End Sub


Sub ����ҥ�ҤT�j�k�H�R��W���()
    Dim stockDate As String
    Dim url As String
    Dim httpObject As Object
    Dim jsonObject As Object
    Dim result As Variant
    Dim dataObject As Object
    Dim columnObject As Object
    Dim i As Integer
    Dim j As Integer

    ' ������
    stockDate = Format(Range("B1").Value, "yyyyMMdd")
    url = "https://www.twse.com.tw/fund/T86?response=json&date=" & stockDate & "&selectType=ALLBUT0999"
    
    Set httpObject = CreateObject("MSXML2.XMLHTTP")
    
    httpObject.Open "GET", url, False
    httpObject.send

    ' �ѪR JSON �榡
    result = httpObject.responseText
    Set jsonObject = JsonConverter.ParseJson(result)
    Set dataObject = jsonObject("data")
    Set columnObject = jsonObject("fields")
    
    ' �M�żƭ�
    Range("A5:S10000").Clear
    
    ' ��r�m��
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
    
    ' ��J���W��
    For i = 1 To columnObject.Count
        Cells(4, i) = columnObject(i)
    Next i
    
    ' ��J���
    ' �C���Ҩ�
    For i = 1 To dataObject.Count
        ' �C�����
        For j = 1 To dataObject(i).Count
            If j >= 1 And j <= 2 Then
                ' �e 2 ���Ҩ��N���M�W�ٳB�z
                Cells(i + 4, j) = Replace(CStr(dataObject(i)(j)), " ", "")
            Else
                ' �� 3 ��H��Ʀr��ƳB�z
                Cells(i + 4, j) = CDbl(dataObject(i)(j))
            End If
        Next j
    Next i
    
    ' �ھڤT�j�k�H�R��W�ѼƱƧ�
    ActiveWorkbook.Worksheets("�T�j�k�H�R��W").Range("A4:S1090").Select
    ActiveWorkbook.Worksheets("�T�j�k�H�R��W").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�T�j�k�H�R��W").Sort.SortFields.Add2 Key:=Range("S5:S1090") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�T�j�k�H�R��W").Sort
        .SetRange Range("A4:S1090")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' �۰ʽվ���e
    Columns.AutoFit
    Range("B1").Select

End Sub























