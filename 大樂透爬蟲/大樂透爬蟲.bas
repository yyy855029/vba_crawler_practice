Attribute VB_Name = "Module1"
'Regular Expression �ǰt�r��
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


Sub ��@�j�ֳz�U��(i As Integer, first_row As Integer)
Attribute ��@�j�ֳz�U��.VB_ProcData.VB_Invoke_Func = " \n14"
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://www.lotto-8.com/listltobigbbk.asp?indexpage=" & CStr(i) & "&orderby=new", _
        Destination:=Range("$A$" & CStr(first_row)))
        .WebFormatting = xlWebFormattingNone
        .WebTables = "5"
        .WebPreFormattedTextToColumns = True
        .Refresh BackgroundQuery:=False
    End With
    
End Sub


Sub �妸�j�ֳz�U��()
    Dim i As Integer
    Dim first_row As Integer
    
    Range("A:I").Clear
    
    For i = 1 To 5
        ' �פJWeb ���
        If Range("A1").Value = "" Then
            first_row = 1
        Else
            first_row = Range("A1").End(xlDown).Row + 1
        End If
        
        Call ��@�j�ֳz�U��(i, first_row)
    
    Next i

End Sub


Sub ��ƳB�z()
    Dim i As Integer
    
    Range("D1") = "�P��"
    
    ' �R���C���W�����A�u�O�d��1�����
    For i = 2 To Range("A1").End(xlDown).Row
        If Cells(i, "A").Value = "���" Then
            Rows(i).Delete
        End If
    Next i
    
    ' �N����O�d���N�A�T�w����榡�A�W�[�P�����
    For i = 2 To Range("A1").End(xlDown).Row
        If (i - 2) Mod 3 = 0 Then
            Cells(i, "A") = Cells(i + 1, "A")
            Cells(i, "A").NumberFormatLocal = "yyyy/m/d"
            ' ����r�ǰt
            Cells(i, "D") = RegxFunc(Cells(i + 2, "A"), "([\u4E00-\u9FFF\u6300-\u77FF\u7800-\u8CFF\u8D00-\u9FFF]+)")
        End If
    Next i
                 
    ' �ѳ̫�C���^�ơA�C�����j3�C
    For i = Range("A1").End(xlDown).Row To 4 Step -3
        Rows(i).Delete
        Rows(i - 1).Delete
    Next i
    
    ' �إߨC�ո��X���Y
    For i = 1 To 6
        Cells(1, 4 + i) = i
    Next i
    
    ' ���ΨC�ո��X
    For i = 2 To Range("A1").End(xlDown).Row
        Range("E" & CStr(i) & ":J" & CStr(i)) = VBA.Split(Cells(i, "B"), ",")
        ' �O�d�ƭȮ榡
        Range("E" & CStr(i) & ":J" & CStr(i)) = Range("E" & CStr(i) & ":J" & CStr(i)).Value
    Next i
    
    ' �R��B��
    Columns("B").Delete
    ' �վ�S�O��(B��)����
    Columns("J") = Columns("B").Value
    Columns("B").Delete
    Columns.AutoFit
               
End Sub


Sub �j�ֳz�U��()
    Call �妸�j�ֳz�U��
    Call ��ƳB�z

End Sub




