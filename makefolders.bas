Attribute VB_Name = "makefolders"
Sub makefolders()

    Dim folders As Range '[1]
    Dim maxRows, maxColumns, rows, columns As Integer
    
    Set folders = Selection '[2]
    maxRows = folders.rows.Count
    maxColumns = folders.columns.Count
    
    For columns = 1 To maxColumns '[3]
    rows = 1
    
    Do While rows <= maxRows '[4]
        If Len(Dir(ActiveWorkbook.Path & "\" & folders(rows, columns), vbDirectory)) = 0 Then
        MkDir (ActiveWorkbook.Path & "\" & folders(rows, columns))
        
        On Error Resume Next
        End If
    
        rows = rows + 1 '[5]
    
    Loop
    Next columns
    
    
'[1] - Declaração de variáveis
'[2] - Define folders = células selecionadas
'[3] - Atribui valor 1 para as colunas e linhas até o máximo
'[4] - Loop de criação das pastas
'[5] - Loop de linhas

End Sub
