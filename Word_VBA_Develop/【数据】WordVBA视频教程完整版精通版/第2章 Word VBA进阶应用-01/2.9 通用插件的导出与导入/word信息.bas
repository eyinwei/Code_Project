Attribute VB_Name = "word信息"
Option Explicit
Sub 当前日期时间()
MsgBox Now
End Sub

Sub doc转PDF()
Dim fd As FileDialog, f, n, arr(), fl, m%
n = MsgBox("是否选择要转为PDF的word文档", 4)
If n = 6 Then
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Show
    For Each fl In fd.SelectedItems
        m = m + 1
        ReDim Preserve arr(1 To m)
        arr(m) = fl
    Next
    MsgBox "请选择要放置的位置的文件夹"
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    If fd.Show = -1 Then
        For Each f In arr
            Documents.Open f
            ActiveDocument.SaveAs2 fd.SelectedItems(1) & "\" & Split(ActiveDocument.Name, ".")(0), 17
            ActiveDocument.Close
        Next
    Else
        MsgBox "你取消了操作！"
    End If
Else
    MsgBox "你取消了操作！"
End If
End Sub
