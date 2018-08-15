Dim workBook_GroupName
Dim workBook_Extract 
Dim workBook_Output 
Dim xlo
Dim groupColumn
Dim groupName
Dim outputFileRowCount
Dim inputGroupRowCount

Set xlo = CreateObject("Excel.Application")
Set workBook_Output = xlo.Workbooks.Open("C:\Users\Praya\Desktop\Sathya\MonthlyReport\AD_Group_OutPut_File.xlsx")
Set workBook_GroupName = xlo.Workbooks.Open("C:\Users\Praya\Desktop\Sathya\MonthlyReport\AD_Group_Input_File.xlsx")
Set workBook_Extract = xlo.Workbooks.Open("C:\Users\Praya\Desktop\Sathya\MonthlyReport\input file 07282018.xlsx")
xlo.Visible = True

groupColumn = 5

    inputGroupRowCount = workBook_GroupName.Worksheets("Sheet1").UsedRange.Rows.Count
    For i = 2 To inputGroupRowCount
    
        groupName = workBook_GroupName.Worksheets("Sheet1").Cells(i, "A").Value
		
        workBook_Extract.Worksheets("Sheet1").UsedRange.AutoFilter 5, "*" & groupName & "*", Operator = xlAnd	
        workBook_Extract.Worksheets("Sheet1").UsedRange.SpecialCells(12).Copy
    
        outputFileRowCount = workBook_Output.Worksheets("Sheet1").UsedRange.Rows.Count
        workBook_Output.Worksheets("Sheet1").Cells(outputFileRowCount + 1, 1).PasteSpecial
        workBook_Output.Worksheets("Sheet1").Cells(outputFileRowCount + 1, 1).EntireRow.Delete
        
        nCount = workBook_Output.Worksheets("Sheet1").UsedRange.Rows.Count 'Output file new row count
        
        For j = outputFileRowCount + 1 To nCount
            workBook_Output.Worksheets("Sheet1").Cells(j, groupColumn).Value = groupName
        Next
                
    Next
    
    workBook_GroupName.Close (False)
    workBook_Extract.Close (False)
    workBook_Output.Save

