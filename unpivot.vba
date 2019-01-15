Sub Unpivot_Large_Excel_To_CSV()

    'Unpivot a very large Excelfile to a CSV file (Excel VBA)
    '15-1-2019 - Laurens Sparrius
    
    Dim sheetName As String
    
    'Unpivot the specified worksheet, number of rows and columns to the specified text file
    sheetName = "Blad4"    'worksheet name
    NumberOfColumns = 196  'including header column
    numberOfRows = 1796    'including header row
    
    Open "c:\users\laurens\desktop\output.txt" For Output As #1
    
    'Write header row
    Print #1, "item" & vbTab & "column" & vbTab & "value"
    
    For currentColumn = 2 To NumberOfColumns
    
        For currentRow = 2 To numberOfRows
        
            Print #1, Sheets(sheetName).Cells(currentRow, 1).Value & vbTab & Sheets(sheetName).Cells(1, currentColumn).Value & vbTab & Sheets(sheetName).Cells(currentRow, currentColumn).Value
        
        Next currentRow
    
    Next currentColumn
    
    Close #1
        

End Sub
