Attribute VB_Name = "makefolders"
Sub mkfolders()

    Dim folders As Range '[1]
    Dim maxRows, maxColumns, rw, cs As Integer
    
    Set folders = Selection '[2]
        maxRows = folders.rows.Count
            maxColumns = folders.columns.Count
    
    For cs = 1 To maxColumns '[3]
    rw = 1
    
    Do While rw <= maxRows '[4]
        If Len(Dir(ActiveWorkbook.Path & "\" & folders(rw, cs), vbDirectory)) = 0 Then
            MkDir (ActiveWorkbook.Path & "\" & folders(rw, cs))
        
            On Error Resume Next
            
        End If
    
        rw = rw + 1 '[5]
    
    Loop
    Next cs
    
    
'[1] - Declaração de variáveis
'[2] - Define folders = células selecionadas
'[3] - Atribui valor 1 para as colunas e linhas até o máximo
'[4] - Loop de criação das pastas
'[5] - Loop de linhas

End Sub
