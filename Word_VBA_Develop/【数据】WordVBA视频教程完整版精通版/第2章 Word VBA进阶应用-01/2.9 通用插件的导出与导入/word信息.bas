Attribute VB_Name = "word��Ϣ"
Option Explicit
Sub ��ǰ����ʱ��()
MsgBox Now
End Sub

Sub docתPDF()
Dim fd As FileDialog, f, n, arr(), fl, m%
n = MsgBox("�Ƿ�ѡ��ҪתΪPDF��word�ĵ�", 4)
If n = 6 Then
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Show
    For Each fl In fd.SelectedItems
        m = m + 1
        ReDim Preserve arr(1 To m)
        arr(m) = fl
    Next
    MsgBox "��ѡ��Ҫ���õ�λ�õ��ļ���"
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    If fd.Show = -1 Then
        For Each f In arr
            Documents.Open f
            ActiveDocument.SaveAs2 fd.SelectedItems(1) & "\" & Split(ActiveDocument.Name, ".")(0), 17
            ActiveDocument.Close
        Next
    Else
        MsgBox "��ȡ���˲�����"
    End If
Else
    MsgBox "��ȡ���˲�����"
End If
End Sub
