Attribute VB_Name = "makefolders"
Sub mkfolders()

    Dim folders As Range 
    Dim maxRows, maxColumns, rw, cs As Integer
    
    ' [ Assign folders = selected cells ]
    Set folders = Selection 
            maxRows = folders.rows.Count
            maxColumns = folders.columns.Count
    
    ' [ Assign value 1 for columns and rows until max
    For cs = 1 To maxColumns 
    rw = 1
        
    ' [ Loop for folder creation ]
    Do While rw <= maxRows '[4]
        If Len(Dir(ActiveWorkbook.Path & "\" & folders(rw, cs), vbDirectory)) = 0 Then
            MkDir (ActiveWorkbook.Path & "\" & folders(rw, cs))
        
            On Error Resume Next
            
        End If
    
    ' [ Loop of rows ]
    rw = rw + 1 
    
    Loop
    Next cs

    MsgBox "Folders created successfully!", vbOKOnly, "Console"
    
End Sub
