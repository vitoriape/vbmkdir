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
    
    
'[1] - Declara��o de vari�veis
'[2] - Define folders = c�lulas selecionadas
'[3] - Atribui valor 1 para as colunas e linhas at� o m�ximo
'[4] - Loop de cria��o das pastas
'[5] - Loop de linhas

End Sub
