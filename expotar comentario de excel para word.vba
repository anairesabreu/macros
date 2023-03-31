Sub CopyCommentsToWord()
'Update 20140325
Dim xComment As Comment
Dim wApp As Object
On Error Resume Next
Set wApp = GetObject(, "Word.Application")
If Err.Number <> 0 Then
  Err.Clear
  Set wApp = CreateObject("Word.Application")
End If
wApp.Visible = True
wApp.Documents.Add DocumentType:=0
For Each xComment In Application.ActiveSheet.Comments
    wApp.Selection.TypeText xComment.Parent.Address & vbTab & xComment.Text
    wApp.Selection.TypeParagraph
Next
Set wApp = Nothing
End Sub
