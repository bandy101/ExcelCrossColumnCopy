Attribute VB_Name = "Module1"
Sub 复制()

   fromexcelpath = "Macintosh HD:Users:Ethan:Desktop:" '====雷艳====从那个文件复制,路径
   fromexcelname = "excel1.xlsx" '====雷艳====从那个文件复制，名称
    '需要复制的列
    Dim fromarr() As Variant
      '复制数据表的索引
    fromsheet = 1 '==================雷艳====从第几个工作表复制数据
    fromtoarr = Array("A>A", "B>B", "C>C", "E>E", "F>F", "G>G", "I>I") '==================雷艳====从哪列复制到哪列
    '需要复制的每列行开始和结束
    from_beginnum = "1" '==================雷艳====从第几行开始复制
    from_endnum = "4" '==================雷艳====复制截止行
    
  
   
    '粘贴数据表的索引
    tosheet = 1 '==================雷艳====粘贴到第几个工作表
    '需要粘贴行的开始
    to_beginnum = "2" '==================雷艳====从第几行开始粘贴



   Response = MsgBox("村长提示：是否已经配置？" + vbCrLf + "从文件：" + fromexcelpath + fromexcelname + "复制数据", vbYesNo)
   If Response = vbYes Then '用户按下“是”按钮
    
    Else '用户按下“否”按钮
    Exit Sub
    End If
    Application.ScreenUpdating = False

    Dim fromwb As Workbook
    '复制数据来源
    
  
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
    MsgBox "村长提示：复制完成", vbOKOnly

End Sub
