Attribute VB_Name = "makefolders"
Sub makefolders()

    Dim folders As Range '[1]
    Dim maxRows, maxColumns, r, c As Integer
    
    Set folders = Selection '[2]
    maxRows = folders.rows.Count
    maxColumns = folders.columns.Count
    
    For c = 1 To maxColumns '[3]
    r = 1
    
    Do While r <= maxRows '[4]
        If Len(Dir(ActiveWorkbook.Path & "\" & folders(r, c), vbDirectory)) = 0 Then
        MkDir (ActiveWorkbook.Path & "\" & folders(r, c))
        
        On Error Resume Next
        End If
    
        r = r + 1 '[5]
    
    Loop
    Next c
    
    
'[1] - Declaração de variáveis
'[2] - Define folders = células selecionadas
'[3] - Atribui valor 1 para as colunas e linhas até o máximo
'[4] - Loop de criação das pastas
'[5] - Loop de linhas

End Sub
