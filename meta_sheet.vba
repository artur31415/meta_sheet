Dim fileNames() As String
Dim snColNames() As String
Dim colCells() As Integer
Dim wbs() As Workbook
Dim fileTemplatePath As String

Sub LoadData()
    Dim row_index As Integer
    Dim col_index As Integer
    Dim meta_wb As Workbook
    
    Dim newPath As String
    Dim newFileName As String
    
    Dim isAlready As Range
    
    Dim col_wb_index As Integer
    
    
    Set meta_wb = Workbooks.Add ' (fileTempaltePath)
    
    fileTemplatePath = "template_file.xlsx"
    
    ReDim Preserve fileNames(ThisWorkbook.Worksheets(1).UsedRange.Rows.Count - 1)
    ReDim Preserve snColNames(ThisWorkbook.Worksheets(1).UsedRange.Rows.Count - 1)
    
    Debug.Print "Number of Files " & ThisWorkbook.Worksheets(1).UsedRange.Rows.Count
    
    For x = 2 To ThisWorkbook.Worksheets(1).UsedRange.Rows.Count
        fileNames(x - 1) = ThisWorkbook.Worksheets(1).Cells(x, 1).Text
        snColNames(x - 1) = ThisWorkbook.Worksheets(1).Cells(x, 2).Text
    Next x
    
    ReDim Preserve wbs(ThisWorkbook.Worksheets(1).UsedRange.Rows.Count - 1)
    ReDim Preserve colCells(ThisWorkbook.Worksheets(1).UsedRange.Rows.Count - 1)
    
    Debug.Print "Load data " & UBound(wbs)
    
    For i = 1 To UBound(fileNames)
        newPath = ThisWorkbook.Path & Application.PathSeparator & fileNames(i) & ".xlsx"
        
        Debug.Print "load at " & newPath
        
        Set wbs(i) = Workbooks.Open(newPath, ReadOnly:=True)
        Debug.Print "rows " & wbs(i).Worksheets(1).UsedRange.Rows.Count
        
    Next i
    
    
    ' look for the cols indexes by their names
    For k = 1 To UBound(wbs)
        colCells(k) = wbs(k).Worksheets(1).Rows(1).Find(What:=snColNames(k), LookIn:=xlValues, _
        LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Column
    Next k
    
    ' copy the headers from all wbs to the new meta
    
    meta_wb.Worksheets(1).Cells(1, 1) = "ID"
    
    col_index = 1
    For m = 1 To UBound(fileNames)
        meta_wb.Worksheets(1).Cells(1, col_index + m) = fileNames(m)
    Next m
    
    col_index = UBound(fileNames) + 2
    
    
    For k = 1 To UBound(wbs)
        wbs(k).Worksheets(1).Range(wbs(k).Worksheets(1).Cells(1, 1), wbs(k).Worksheets(1).Cells(1, wbs(k).Worksheets(1).UsedRange.Columns.Count)).Copy _
        Destination:=meta_wb.Worksheets(1).Cells(1, col_index)
        col_index = col_index + wbs(k).Worksheets(1).UsedRange.Columns.Count
    Next k
    
    
    
    col_wb_index = UBound(fileNames) + 2
    row_index = 2
    ' loop thorugh all wbs's
    For a = 1 To UBound(wbs)
        ' row_index = meta_wb.Worksheets(1).UsedRange.Rows.Count + 1
        For x = 2 To wbs(a).Worksheets(1).UsedRange.Rows.Count
    
        
            ' FIXME: BUG WHEN THE SPECIAL COLLUMN IS EMPTY?
            If a = 1 Then
                Set isAlready = Nothing
            Else
                Set isAlready = meta_wb.Worksheets(1).Range(meta_wb.Worksheets(1).Cells(2, col_wb_index + (colCells(a) - 1)), _
                    meta_wb.Worksheets(1).Cells(meta_wb.Worksheets(1).UsedRange.Rows.Count, col_wb_index + (colCells(a) - 1))).Find( _
                    What:=wbs(a).Worksheets(1).Cells(x, colCells(a)).Text, _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
                    MatchCase:=False, SearchFormat:=False)
                
            End If
            
            
            If isAlready Is Nothing Then
                meta_wb.Worksheets(1).Cells(row_index, 1) = row_index - 1
                wbs(a).Worksheets(1).Range(wbs(a).Worksheets(1).Cells(x, 1), wbs(a).Worksheets(1).Cells(x, wbs(a).Worksheets(1).UsedRange.Columns.Count)).Copy _
                Destination:=meta_wb.Worksheets(1).Cells(row_index, col_wb_index)
                
                meta_wb.Worksheets(1).Cells(row_index, a + 1) = "yes"
                
                If a > 1 Then
                    For b = 1 To a - 1
                        If IsEmpty(meta_wb.Worksheets(1).Cells(row_index, b + 1)) Then
                            meta_wb.Worksheets(1).Cells(row_index, b + 1) = "no"
                        End If
                    Next b
                End If
                
                col_index = UBound(fileNames) + 2 ' wbs(a).Worksheets(1).UsedRange.Columns.Count + 1
                For k = 1 To UBound(wbs)
                    If k > a Then
                        meta_wb.Worksheets(1).Cells(row_index, k + 1) = "no"
                        For x_other = 1 To wbs(k).Worksheets(1).UsedRange.Rows.Count
                            If wbs(a).Worksheets(1).Cells(x, colCells(a)).Text = wbs(k).Worksheets(1).Cells(x_other, colCells(k)).Text Then
                                ' Then this is a match!
                                wbs(k).Worksheets(1).Range(wbs(k).Worksheets(1).Cells(x_other, 1), wbs(k).Worksheets(1).Cells(x_other, wbs(k).Worksheets(1).UsedRange.Columns.Count)).Copy _
                                Destination:=meta_wb.Worksheets(1).Cells(row_index, col_index)
                                
                                meta_wb.Worksheets(1).Cells(row_index, k + 1) = "yes"
                                
                                Exit For
                            End If
                        
                        Next x_other
                    End If
                    
                    col_index = col_index + wbs(k).Worksheets(1).UsedRange.Columns.Count
                Next k
                
                row_index = row_index + 1
            End If
            
            
            
        Next x
        Debug.Print "at row " & row_index
        ' row_index = row_index + wbs(a).Worksheets(1).UsedRange.Rows.Count
        col_wb_index = col_wb_index + wbs(a).Worksheets(1).UsedRange.Columns.Count
    Next a
    
    ' meta_wb.Active
    newFileName = "new_meta_data" & Format(Now(), "yyyy_MM_dd_hh_mm_ss") & ".xls"
    With meta_wb
        .Title = "New Meta WB"
        .Subject = "Meta WB"
        .SaveAs Filename:=newFileName
    End With

End Sub
