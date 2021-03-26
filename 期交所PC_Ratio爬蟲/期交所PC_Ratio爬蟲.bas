Attribute VB_Name = "Module1"
Sub ��@PC_Ratio�U��(start_date As Date, end_date As Date, first_row As Integer)
    Dim url As String
    
    url = "https://www.taifex.com.tw/cht/3/pcRatio?&queryStartDate=" & Format(start_date, "yyyy/MM/dd") & "&queryEndDate=" & Format(end_date, "yyyy/MM/dd")
        
    ' �C�����q�̫�C�}�l�K���
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;" & url _
        , Destination:=Range("$A$" & CStr(first_row)))
        .WebFormatting = xlWebFormattingNone
        .WebTables = "4"
        .WebPreFormattedTextToColumns = True
        .Refresh BackgroundQuery:=False
    End With
    
End Sub


Sub �妸PC_Ratio�U��()
    Dim total_start_date As Date
    Dim start_date As Date
    Dim end_date As Date
    Dim i As Integer
    Dim first_row As Integer
    Dim iter_time As Integer
    Dim total_day_diff As Integer
    Dim day_mod As Integer
    
    ' ���M�Ÿ��
    Range("A:G").Clear
    
    ' ���o�`�_�l��M������
    total_start_date = Range("J5").Value
    end_date = Range("J6").Value
    
    ' �p�⭡�N���ƩM�̫�@���ۮt�Ѽ�
    total_day_diff = DateDiff("d", total_start_date, end_date)
    iter_time = total_day_diff \ 30
    day_mod = total_day_diff Mod 30
    
    ' �Y�ۮt�ѼƤp��30��
    If total_day_diff < 30 Then
        start_date = DateAdd("d", -day_mod, end_date)
    Else
        start_date = DateAdd("d", -30, end_date)
    End If
    
    ' �פJWeb ���
    For i = 1 To iter_time - 1
        ' �p��̫�C�C��
        If Range("A1").Value = "" Then
            first_row = 1
        Else
            first_row = Range("A1").End(xlDown).Row + 1
        End If

        ' �C�����q�̫�C�}�l�K���
        Call ��@PC_Ratio�U��(start_date, end_date, first_row)

        ' ��s�_�l��M������
        end_date = DateAdd("d", -1, start_date)
        start_date = DateAdd("d", -30, end_date)
        
    Next i
    
    ' �p��̫�@�����G
    ' �p��̫�C�C��
    If Range("A1").Value = "" Then
        first_row = 1
    Else
        first_row = Range("A1").End(xlDown).Row + 1
    End If
    
    ' �Y�ۮt�ѼƤp��30��
    If iter_time <> 0 Then
        ' �Y�ۮt�ѼƤ��O30�Ѫ�����
        If day_mod <> 0 Then
            end_date = DateAdd("d", -1, start_date)
            start_date = DateAdd("d", -day_mod + iter_time, end_date)
        Else
            end_date = DateAdd("d", -1, start_date)
            start_date = DateAdd("d", -30 + iter_time, end_date)
        End If
    End If
    
    Call ��@PC_Ratio�U��(start_date, end_date, first_row)
    
    ' �R���C���W�����A�u�O�d��1�����
    For i = 2 To Range("A1").End(xlDown).Row
        If Cells(i, "A").Value = "���" Then
            Rows(i).Delete
        End If
    Next i
    
End Sub








