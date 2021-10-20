Attribute VB_Name = "Module1"
Sub ������f�������()
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
    
    ' �אּ��ʧ�s
    Application.Calculation = xlCalculationManual
    
    ' ��Ƥ��
    start_date = Worksheets("�ϧ�").Range("C2").Value
    end_date = Worksheets("�ϧ�").Range("D2").Value
    date_array = Array(start_date, end_date)
    
    ' ���f�T�j�k�H
    dataname_array = Array("��Ӥ���f�T�j�k�H", "�̷s����f�T�j�k�H")
    
    For i = 0 To UBound(date_array) - LBound(date_array)
        ' ���M�Ÿ��
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
    
    ' ���f����
    dataname_array = Array("��Ӥ���f����", "�̷s����f����")
    
    For i = 0 To UBound(date_array) - LBound(date_array)
        ' ���M�Ÿ��
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
                ' �ư�����B�z
                Worksheets(dataname_array(i)).Cells(j, k) = Replace(ncol.innertext, vbCrLf, "")
                k = k + 1
            Next ncol
            j = j + 1
        Next nrow
    
    Next i
    
    ' ��_�۰ʧ�s
    Application.Calculation = xlCalculationAutomatic
    
End Sub


Sub �������v�������()
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
    
    ' �אּ��ʧ�s
    Application.Calculation = xlCalculationManual
    
    ' ��Ƥ��
    start_date = Worksheets("�ϧ�").Range("C2").Value
    end_date = Worksheets("�ϧ�").Range("D2").Value
    date_array = Array(start_date, end_date)
    
    ' ����v�T�j�k�H
    dataname_array = Array("��Ӥ����v�T�j�k�H", "�̷s�����v�T�j�k�H")
    
    For i = 0 To UBound(date_array) - LBound(date_array)
        ' ���M�Ÿ��
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
                ' �ư�����B�z
                Worksheets(dataname_array(i)).Cells(j, k) = Replace(ncol.innertext, vbCrLf, "")
                k = k + 1
            Next ncol
            j = j + 1
        Next nrow
    
    Next i
    
    ' ����v����
    dataname_array = Array("��Ӥ����v����", "�̷s�����v����")
    
    For i = 0 To UBound(date_array) - LBound(date_array)
        ' ���M�Ÿ��
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
                ' �ư�����B�z
                Worksheets(dataname_array(i)).Cells(j, k) = Replace(ncol.innertext, vbCrLf, "")
                k = k + 1
            Next ncol
            j = j + 1
        Next nrow
    
    Next i
    
    ' ��z����v����
    sort_dataname_array = Array("��Ӥ��z����v����", "�̷s���z����v����")
    
    For i = 0 To UBound(date_array) - LBound(date_array)
        ' ���M�Ÿ��
        Worksheets(sort_dataname_array(i)).Cells.ClearContents
        ' ���W��
        Worksheets(sort_dataname_array(i)).Range("A1:I1") = Array("�X��", "Call �i����", "Call ����", "Call ����q", "Call ������", "Put �i����", "Put ����", "Put ����q", "Put ������")
        
        ' �ˬd�O�_�����X��
        check_date = Worksheets(dataname_array(i)).Range("B2")
        
        For j = 2 To Worksheets(dataname_array(i)).Range("A1").End(xlDown).Row
            ' Call
            If j Mod 2 = 0 And check_date = Worksheets(dataname_array(i)).Cells(j, 2) Then
                ' �X��
                Worksheets(sort_dataname_array(i)).Cells(j \ 2 + 1, 1) = Worksheets(dataname_array(i)).Cells(j, 2)
                ' �i����
                Worksheets(sort_dataname_array(i)).Cells(j \ 2 + 1, 2) = Worksheets(dataname_array(i)).Cells(j, 3)
                ' ����
                Worksheets(sort_dataname_array(i)).Cells(j \ 2 + 1, 3) = Replace(Worksheets(dataname_array(i)).Cells(j, 8), "-", "")
                ' ����q
                Worksheets(sort_dataname_array(i)).Cells(j \ 2 + 1, 4) = Worksheets(dataname_array(i)).Cells(j, 14)
                ' ������
                Worksheets(sort_dataname_array(i)).Cells(j \ 2 + 1, 5) = Worksheets(dataname_array(i)).Cells(j, 15)
             
            ' Put
            ElseIf check_date = Worksheets(dataname_array(i)).Cells(j, 2) Then
                ' �i����
                Worksheets(sort_dataname_array(i)).Cells(j \ 2 + 1, 6) = Worksheets(dataname_array(i)).Cells(j, 3)
                ' ����
                Worksheets(sort_dataname_array(i)).Cells(j \ 2 + 1, 7) = Replace(Worksheets(dataname_array(i)).Cells(j, 8), "-", "")
                ' ����q
                Worksheets(sort_dataname_array(i)).Cells(j \ 2 + 1, 8) = Worksheets(dataname_array(i)).Cells(j, 14)
                ' ������
                Worksheets(sort_dataname_array(i)).Cells(j \ 2 + 1, 9) = Worksheets(dataname_array(i)).Cells(j, 15)
            Else
                Exit For
            End If
        Next j
    Next i
    
    ' ��_�۰ʧ�s
    Application.Calculation = xlCalculationAutomatic
    
End Sub


Sub ����()
    Call ������f�������
    Call �������v�������
End Sub
