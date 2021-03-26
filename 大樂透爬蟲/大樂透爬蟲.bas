Attribute VB_Name = "Module1"
'Regular Expression 匹配字串
Function RegxFunc(strInput As String, regexPattern As String) As String
    Dim regEx As New RegExp
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = regexPattern
    End With

    If regEx.Test(strInput) Then
        Set matches = regEx.Execute(strInput)
        RegxFunc = matches(0).Value
    Else
        RegxFunc = "not matched"
    End If
    
End Function


Sub 單一大樂透下載(i As Integer, first_row As Integer)
Attribute 單一大樂透下載.VB_ProcData.VB_Invoke_Func = " \n14"
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://www.lotto-8.com/listltobigbbk.asp?indexpage=" & CStr(i) & "&orderby=new", _
        Destination:=Range("$A$" & CStr(first_row)))
        .WebFormatting = xlWebFormattingNone
        .WebTables = "5"
        .WebPreFormattedTextToColumns = True
        .Refresh BackgroundQuery:=False
    End With
    
End Sub


Sub 批次大樂透下載()
    Dim i As Integer
    Dim first_row As Integer
    
    Range("A:I").Clear
    
    For i = 1 To 5
        ' 匯入Web 資料
        If Range("A1").Value = "" Then
            first_row = 1
        Else
            first_row = Range("A1").End(xlDown).Row + 1
        End If
        
        Call 單一大樂透下載(i, first_row)
    
    Next i

End Sub


Sub 資料處理()
    Dim i As Integer
    
    Range("D1") = "星期"
    
    ' 刪除每頁上方欄位，只保留第1個欄位
    For i = 2 To Range("A1").End(xlDown).Row
        If Cells(i, "A").Value = "日期" Then
            Rows(i).Delete
        End If
    Next i
    
    ' 將日期保留取代，固定日期格式，增加星期欄位
    For i = 2 To Range("A1").End(xlDown).Row
        If (i - 2) Mod 3 = 0 Then
            Cells(i, "A") = Cells(i + 1, "A")
            Cells(i, "A").NumberFormatLocal = "yyyy/m/d"
            ' 中文字匹配
            Cells(i, "D") = RegxFunc(Cells(i + 2, "A"), "([\u4E00-\u9FFF\u6300-\u77FF\u7800-\u8CFF\u8D00-\u9FFF]+)")
        End If
    Next i
                 
    ' 由最後列往回數，每次間隔3列
    For i = Range("A1").End(xlDown).Row To 4 Step -3
        Rows(i).Delete
        Rows(i - 1).Delete
    Next i
    
    ' 建立每組號碼表頭
    For i = 1 To 6
        Cells(1, 4 + i) = i
    Next i
    
    ' 分割每組號碼
    For i = 2 To Range("A1").End(xlDown).Row
        Range("E" & CStr(i) & ":J" & CStr(i)) = VBA.Split(Cells(i, "B"), ",")
        ' 保留數值格式
        Range("E" & CStr(i) & ":J" & CStr(i)) = Range("E" & CStr(i) & ":J" & CStr(i)).Value
    Next i
    
    ' 刪除B欄
    Columns("B").Delete
    ' 調整特別號(B欄)順序
    Columns("J") = Columns("B").Value
    Columns("B").Delete
    Columns.AutoFit
               
End Sub


Sub 大樂透下載()
    Call 批次大樂透下載
    Call 資料處理

End Sub




