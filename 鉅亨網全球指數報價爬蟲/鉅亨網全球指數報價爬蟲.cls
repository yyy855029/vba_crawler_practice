VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub ъ芖计厨基()
    Dim i As Integer
    Dim j As Integer
    
    With ThisWorkbook.Worksheets("芖计厨基")
        ' 睲戈
        .Range("A5:S39").Clear
        .Range("C2").Clear
        
        ' 砞﹚程ら
        If Weekday(Date) = 7 Then
            ' 琍戳せ矪瞶
            .Range("C2") = DateAdd("d", -1, Date)
        ElseIf Weekday(Date) = 1 Then
            ' 琍戳ら矪瞶
            .Range("C2") = DateAdd("d", -2, Date)
        End If
        
        ' 眖筪﹄呼ъ芖计厨基戈
        ' 栋いカ初
        With .QueryTables.Add(Connection:= _
            "URL;https://www.cnyes.com/twstock/Index2Real_idx.htm", Destination:=.Range( _
            "$A$5"))
            .WebFormatting = xlWebFormattingNone
            .WebTables = "1"
            .WebPreFormattedTextToColumns = True
            .Refresh BackgroundQuery:=False
        End With
        
        ' 耫禦カ初
        With .QueryTables.Add(Connection:= _
            "URL;https://www.cnyes.com/twstock/Index2Real_idx.htm", Destination:=.Range( _
            "$K$5"))
            .WebFormatting = xlWebFormattingNone
            .WebTables = "2"
            .WebPreFormattedTextToColumns = True
            .Refresh BackgroundQuery:=False
        End With
        
        ' ㎝琎らΜ絃基ゑ计帝︹
        ' 栋いカ初
        For i = 6 To .Range("D6").End(xlDown).Row
            ' Μ絃基
            If .Cells(i, 4) > 0 Then
                .Range(.Cells(i, 3), .Cells(i, 5)).Font.ColorIndex = 3
            ElseIf Cells(i, 4) < 0 Then
                .Range(.Cells(i, 3), .Cells(i, 5)).Font.ColorIndex = 10
            End If
            
            ' 秨絃程蔼程基
            For j = 6 To 8
                If .Cells(i, j) > .Cells(i, 9) Then
                    .Cells(i, j).Font.ColorIndex = 3
                ElseIf .Cells(i, j) < .Cells(i, 9) Then
                    .Cells(i, j).Font.ColorIndex = 10
                End If
            Next j
            
        Next i
        
        ' 耫禦カ初
        For i = 6 To .Range("N6").End(xlDown).Row
            ' Μ絃基
            If .Cells(i, 14) > 0 Then
                .Range(.Cells(i, 13), .Cells(i, 15)).Font.ColorIndex = 3
            ElseIf Cells(i, 14) < 0 Then
                .Range(.Cells(i, 13), .Cells(i, 15)).Font.ColorIndex = 10
            End If
            
            ' 秨絃程蔼程基
            For j = 16 To 18
                If .Cells(i, j) > .Cells(i, 19) Then
                    .Cells(i, j).Font.ColorIndex = 3
                ElseIf .Cells(i, j) < .Cells(i, 19) Then
                    .Cells(i, j).Font.ColorIndex = 10
                End If
            Next j
        
        Next i
        
    End With
End Sub


Sub ъㄈ瑆计厨基()
    Dim i As Integer
    Dim j As Integer
    
    With ThisWorkbook.Worksheets("ㄈ瑆计厨基")
        ' 睲戈
        .Range("A5:I21").Clear
        .Range("C2").Clear
        .Range("C2") = Date
        
        ' 眖筪﹄呼ъㄈ瑆计厨基戈
        With .QueryTables.Add(Connection:= _
            "URL;https://www.cnyes.com/global/IndexImmediateQuotedPrice/003/QuotationPrice.aspx" _
            , Destination:=.Range("$A$5"))
            .WebFormatting = xlWebFormattingNone
            .WebPreFormattedTextToColumns = True
            .Refresh BackgroundQuery:=False
        End With
        
        ' ㎝琎らΜ絃基ゑ计帝︹
        ' 栋いカ初
        For i = 6 To .Range("D6").End(xlDown).Row
            ' Μ絃基
            If .Cells(i, 4) > 0 Then
                .Range(.Cells(i, 3), .Cells(i, 5)).Font.ColorIndex = 3
            ElseIf .Cells(i, 4) < 0 Then
                .Range(.Cells(i, 3), .Cells(i, 5)).Font.ColorIndex = 10
            End If
            
            ' 秨絃程蔼程基
            For j = 6 To 8
                If .Cells(i, j) > .Cells(i, 9) Then
                    .Cells(i, j).Font.ColorIndex = 3
                ElseIf .Cells(i, j) < .Cells(i, 9) Then
                    .Cells(i, j).Font.ColorIndex = 10
                End If
            Next j
            
        Next i
    
    End With
End Sub


Sub ъ稼瑆计厨基()
    Dim i As Integer
    Dim j As Integer
    
    With ThisWorkbook.Worksheets("稼瑆计厨基")
        ' 睲戈
        .Range("A5:I22").Clear
        .Range("C2").Clear
        .Range("C2") = Date
        
        ' 眖筪﹄呼ъ稼瑆计厨基戈
        With .QueryTables.Add(Connection:= _
            "URL;https://www.cnyes.com/global/IndexImmediateQuotedPrice/004/QuotationPrice.aspx" _
            , Destination:=.Range("$A$5"))
            .WebFormatting = xlWebFormattingNone
            .WebPreFormattedTextToColumns = True
            .Refresh BackgroundQuery:=False
        End With
        
        ' ㎝琎らΜ絃基ゑ计帝︹
        ' 栋いカ初
        For i = 6 To .Range("D6").End(xlDown).Row
            ' Μ絃基
            If .Cells(i, 4) > 0 Then
                .Range(.Cells(i, 3), .Cells(i, 5)).Font.ColorIndex = 3
            ElseIf .Cells(i, 4) < 0 Then
                .Range(.Cells(i, 3), .Cells(i, 5)).Font.ColorIndex = 10
            End If
            
            ' 秨絃程蔼程基
            For j = 6 To 8
                If .Cells(i, j) > .Cells(i, 9) Then
                    .Cells(i, j).Font.ColorIndex = 3
                ElseIf .Cells(i, j) < .Cells(i, 9) Then
                    .Cells(i, j).Font.ColorIndex = 10
                End If
            Next j
            
        Next i
    
    End With
End Sub


Sub ъ瑆计厨基()
    Dim i As Integer
    Dim j As Integer
    
    With ThisWorkbook.Worksheets("瑆计厨基")
        ' 睲戈
        .Range("A5:I16").Clear
        .Range("C2").Clear
        .Range("C2") = Date
        
        ' 眖筪﹄呼ъ瑆计厨基戈
        With .QueryTables.Add(Connection:= _
            "URL;https://www.cnyes.com/global/IndexImmediateQuotedPrice/005/QuotationPrice.aspx" _
            , Destination:=.Range("$A$5"))
            .WebFormatting = xlWebFormattingNone
            .WebPreFormattedTextToColumns = True
            .Refresh BackgroundQuery:=False
        End With
        
        ' ㎝琎らΜ絃基ゑ计帝︹
        ' 栋いカ初
        For i = 6 To .Range("D6").End(xlDown).Row
            ' Μ絃基
            If .Cells(i, 4) > 0 Then
                .Range(.Cells(i, 3), .Cells(i, 5)).Font.ColorIndex = 3
            ElseIf .Cells(i, 4) < 0 Then
                .Range(.Cells(i, 3), .Cells(i, 5)).Font.ColorIndex = 10
            End If
            
            ' 秨絃程蔼程基
            For j = 6 To 8
                If .Cells(i, j) > .Cells(i, 9) Then
                    .Cells(i, j).Font.ColorIndex = 3
                ElseIf .Cells(i, j) < .Cells(i, 9) Then
                    .Cells(i, j).Font.ColorIndex = 10
                End If
            Next j
            
        Next i
    
    End With
End Sub


Private Sub Workbook_Open()
    Call ъ芖计厨基
    Call ъㄈ瑆计厨基
    Call ъ稼瑆计厨基
    Call ъ瑆计厨基
    ThisWorkbook.Worksheets("芖计厨基").Select
End Sub



