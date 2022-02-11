'version 9.3

' Sub BlueCiting()
'     For i = 1 To ActiveDocument.Fields.Count '遍历文档所有域
'         If Left(ActiveDocument.Fields(i).Code, 4) = " REF" Or Left(ActiveDocument.Fields(i).Code, 14) = " ADDIN EN.CITE" Then 'Word自带的交叉引用的域代码起始4位是" REF"（注意空格），EndNote插入的引用域代码的起始14为是" ADDIN EN.CITE"。根据需求可添加其他类型。
'             ActiveDocument.Fields(i).Select '选中上述几类域
'             Selection.Font.Color = wdColorRed '设置字体颜色
'         End If
'     Next
' End Sub
Sub BlueCiting()
    For i = 1 To ActiveDocument.Fields.Count '遍历文档所有域
        If Left(ActiveDocument.Fields(i).Code, 4) = " REF" Then 
        'Word自带的交叉引用的域代码起始4位是" REF"（注意空格）
            ActiveDocument.Fields(i).Select '选中上述几类域
            Selection.Font.Color = wdColorBlue '设置字体颜色
        ElseIf Left(ActiveDocument.Fields(i).Code, 14) = " ADDIN EN.CITE" Then 
        ' EndNote插入的引用域代码的起始14位是" ADDIN EN.CITE"。根据需求可添加其他类型。
            ActiveDocument.Fields(i).Select '选中上述几类域
            Selection.Font.Color = wdColorRed '设置字体颜色
        End If
        
    Next
End Sub

