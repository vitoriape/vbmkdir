Attribute VB_Name = "makeFolders"
Sub makeFolders()

Dim Rng As Range
Dim maxRows, maxCols, r, c As Integer
Set Rng = Selection
maxRows = Rng.Rows.Count
maxCols = Rng.Columns.Count
For c = 1 To maxCols
r = 1
Do While r <= maxRows
If Len(Dir(ActiveWorkbook.Path & "\" & Rng(r, c), vbDirectory)) = 0 Then
MkDir (ActiveWorkbook.Path & "\" & Rng(r, c))
On Error Resume Next
End If
r = r + 1
Loop
Next c

End Sub
