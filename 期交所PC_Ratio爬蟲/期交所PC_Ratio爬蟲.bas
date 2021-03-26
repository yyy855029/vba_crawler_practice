Attribute VB_Name = "Module1"
Sub 虫PC_Ratio更(start_date As Date, end_date As Date, first_row As Integer)
    Dim url As String
    
    url = "https://www.taifex.com.tw/cht/3/pcRatio?&queryStartDate=" & Format(start_date, "yyyy/MM/dd") & "&queryEndDate=" & Format(end_date, "yyyy/MM/dd")
        
    ' –Ω常眖程秨﹍禟戈
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;" & url _
        , Destination:=Range("$A$" & CStr(first_row)))
        .WebFormatting = xlWebFormattingNone
        .WebTables = "4"
        .WebPreFormattedTextToColumns = True
        .Refresh BackgroundQuery:=False
    End With
    
End Sub


Sub уΩPC_Ratio更()
    Dim total_start_date As Date
    Dim start_date As Date
    Dim end_date As Date
    Dim i As Integer
    Dim first_row As Integer
    Dim iter_time As Integer
    Dim total_day_diff As Integer
    Dim day_mod As Integer
    
    ' 睲戈
    Range("A:G").Clear
    
    ' 眔羆癬﹍ら㎝挡ら
    total_start_date = Range("J5").Value
    end_date = Range("J6").Value
    
    ' 璸衡Ω计㎝程Ω畉ぱ计
    total_day_diff = DateDiff("d", total_start_date, end_date)
    iter_time = total_day_diff \ 30
    day_mod = total_day_diff Mod 30
    
    ' 璝畉ぱ计30ぱ
    If total_day_diff < 30 Then
        start_date = DateAdd("d", -day_mod, end_date)
    Else
        start_date = DateAdd("d", -30, end_date)
    End If
    
    ' 蹲Web 戈
    For i = 1 To iter_time - 1
        ' 璸衡程计
        If Range("A1").Value = "" Then
            first_row = 1
        Else
            first_row = Range("A1").End(xlDown).Row + 1
        End If

        ' –Ω常眖程秨﹍禟戈
        Call 虫PC_Ratio更(start_date, end_date, first_row)

        ' 穝癬﹍ら㎝挡ら
        end_date = DateAdd("d", -1, start_date)
        start_date = DateAdd("d", -30, end_date)
        
    Next i
    
    ' 璸衡程Ω挡狦
    ' 璸衡程计
    If Range("A1").Value = "" Then
        first_row = 1
    Else
        first_row = Range("A1").End(xlDown).Row + 1
    End If
    
    ' 璝畉ぱ计30ぱ
    If iter_time <> 0 Then
        ' 璝畉ぱ计ぃ琌30ぱ计
        If day_mod <> 0 Then
            end_date = DateAdd("d", -1, start_date)
            start_date = DateAdd("d", -day_mod + iter_time, end_date)
        Else
            end_date = DateAdd("d", -1, start_date)
            start_date = DateAdd("d", -30 + iter_time, end_date)
        End If
    End If
    
    Call 虫PC_Ratio更(start_date, end_date, first_row)
    
    ' 埃–よ逆玂痙材1逆
    For i = 2 To Range("A1").End(xlDown).Row
        If Cells(i, "A").Value = "ら戳" Then
            Rows(i).Delete
        End If
    Next i
    
End Sub








