Attribute VB_Name = "Module1"
Sub ����()

   fromexcelpath = "Macintosh HD:Users:Ethan:Desktop:" '====����====���Ǹ��ļ�����,·��
   fromexcelname = "excel1.xlsx" '====����====���Ǹ��ļ����ƣ�����
    '��Ҫ���Ƶ���
    Dim fromarr() As Variant
      '�������ݱ������
    fromsheet = 1 '==================����====�ӵڼ���������������
    fromtoarr = Array("A>A", "B>B", "C>C", "E>E", "F>F", "G>G", "I>I") '==================����====�����и��Ƶ�����
    '��Ҫ���Ƶ�ÿ���п�ʼ�ͽ���
    from_beginnum = "1" '==================����====�ӵڼ��п�ʼ����
    from_endnum = "4" '==================����====���ƽ�ֹ��
    
  
   
    'ճ�����ݱ������
    tosheet = 1 '==================����====ճ�����ڼ���������
    '��Ҫճ���еĿ�ʼ
    to_beginnum = "2" '==================����====�ӵڼ��п�ʼճ��



   Response = MsgBox("�峤��ʾ���Ƿ��Ѿ����ã�" + vbCrLf + "���ļ���" + fromexcelpath + fromexcelname + "��������", vbYesNo)
   If Response = vbYes Then '�û����¡��ǡ���ť
    
    Else '�û����¡��񡱰�ť
    Exit Sub
    End If
    Application.ScreenUpdating = False

    Dim fromwb As Workbook
    '����������Դ
    
  
    On Error Resume Next
    Set fromwb = Workbooks(fromexcelname)
    On Error GoTo 0
    If fromwb Is Nothing Then
        Set fromwb = Workbooks.Open(Filename:=fromexcelpath + fromexcelname)
    End If
    
    
    
    Dim from_arrlen As Integer
    from_arrlen = UBound(fromtoarr)

    For i = 0 To from_arrlen
        from_chead = Split(fromtoarr(i), ">")(0)
        to_chead = Split(fromtoarr(i), ">")(1)
        
        from_rangestr = from_chead + from_beginnum + ":" + from_chead + from_endnum
        to_rangstr = to_chead + to_beginnum
        
        fromwb.Sheets(fromsheet).Range(from_rangestr).Copy
        
        ThisWorkbook.Sheets(tosheet).Activate
        ThisWorkbook.Sheets(tosheet).Range(to_rangstr).Select
        Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Next
    MsgBox "�峤��ʾ���������", vbOKOnly

End Sub
