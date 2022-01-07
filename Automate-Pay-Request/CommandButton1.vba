Private Sub CommandButton1_Click()


    'Prompt user for range
    myDate = InputBox("Enter the date range.", "Get Date", Range("I3").Value)
    Range("I3").Value = myDate



    'Prompt user for range
    Set myRange = Application.Selection
    Set myRange = Application.InputBox("Select the range of data.", "Selecting Range", myRange.Address, Type:=8)
    
    
    'We will sort the data by property name
    myRange.Sort Key1:=Range("I6"), Order1:=xlAscending
    
    
    'Turn screen updating off
    Application.ScreenUpdating = False
    
    'We only care about entries that are "Deployed" and not deployed to "State - City"
    cTextDeployed = "Deployed"
    cTextCity = "State - City"
    
    deletedRows = 0
    
    numRows = myRange.Rows.Count
    numCols = myRange.Columns.Count
    
    'Iterate through the range.
    For i = myRange.Rows.Count To 1 Step -1
    
        'We grab the row we are currently in
        Set myRow = myRange.Rows(i)
        
        'Did we find anything in the row that contains "Deployed"?
        Set findDeployed = myRow.Find(cTextDeployed, LookIn:=xlValues)
        'Did we find anything in the row that contains "State - City"?
        Set findCity = myRow.Find(cTextCity, LookIn:=xlValues)
        
        'If the row does not have "Deployed" in it, we delete it
        If findDeployed Is Nothing Then
           myRow.Delete
           deletedRows = deletedRows + 1
           
        'If the row does have "Deployed" in it, but contains "State - City", we delete
        ElseIf Not findCity Is Nothing Then
                myRow.Delete
                deletedRows = deletedRows + 1
            
        End If
        
    Next
    'Once the iteration is done, we have the data we want
    
    
    



    'Get the number of rows avaliable
    numRows = myRange.Rows.Count
    'What is the PayRequst ID?
    payRequestID = ActiveSheet.Cells(2, "I").Value
    'String placeholder
    sameProperty = Null
    
    'What is the date?
    payRequestDate = ActiveSheet.Cells(3, "I").Value
    
    entryCounter = 0
    
    'Iterate through again
    For j = 1 To myRange.Rows.Count
    
        'We grab the row we are currently in
        Set myRow = myRange.Rows(j)
        
        'We determine the property name at this row
        PayRequestProperty = Cells(myRow.row, "I")
        
        'Does this property match the previous one? If so:
        If PayRequestProperty = sameProperty Then
            'Set the Pay Request ID value to column B
            Range("B" & myRow.row).Value = payRequestID
            
            entryCounter = entryCounter + 1
            
            infoSubtype = Range("E" & myRow.row).Value
            infoModel = Range("F" & myRow.row).Value
            infoUID = Range("G" & myRow.row).Value
            infoQuantity = Range("K" & myRow.row).Value
            infoReference = Range("J" & myRow.row).Value
            

            'Cells(col, row)
            sht.Cells(7 + entryCounter, 4).Value = infoSubtype
            sht.Cells(7 + entryCounter, 5).Value = infoModel
            sht.Cells(7 + entryCounter, 6).Value = infoUID
            sht.Cells(7 + entryCounter, 2).Value = infoQuantity
            sht.Cells(7 + entryCounter, 8).Value = infoReference
            
            'We now determine the GL Account code
            If sht.Cells(7 + entryCounter, 4).Value = "Licenses" Then
                sht.Cells(7 + entryCounter, 7).Value = "11111"
                
            ElseIf sht.Cells(7 + entryCounter, 4).Value = "Printers" Then
                sht.Cells(7 + entryCounter, 7).Value = "22222"
                
            ElseIf sht.Cells(7 + entryCounter, 4).Value = "Networking" Then
                sht.Cells(7 + entryCounter, 7).Value = "33333"
                
            ElseIf sht.Cells(7 + entryCounter, 4).Value = "Telecom" Then
                sht.Cells(7 + entryCounter, 7).Value = "33333"
            
            Else
                sht.Cells(7 + entryCounter, 7).Value = "44444"
            
            End If
            
            
            
        'The property does not match the previous one:
        Else
            'Increment the PayRequestID
            payRequestID = payRequestID + 1
            'Set the Pay Request ID value to column B
            Range("B" & myRow.row).Value = payRequestID
            'Set the sameProperty to check for next comparison\
            sameProperty = PayRequestProperty
            
            ' Get all variables needed
            'Name of Pay Request
            payRequestName = PayRequestProperty + " " + payRequestDate
            
            entryCounter = 0
            
            
            'Duplicate the Pay Request Form for this entry
            Worksheets("PayRequestForm").Copy After:=Sheets("UploadData")
            Set sht = Worksheets(Sheets("UploadData").Index + 1)
            
            
            
            'We need conditions here; remove Apartments, and, Townhomes
            If Len(PayRequestProperty) > 31 Then
                PayRequestProperty = Replace(PayRequestProperty, "Apartments and Townhomes", "", 1, 1)
                PayRequestProperty = Replace(PayRequestProperty, "Apartments", "", 1, 1)
                PayRequestProperty = Replace(PayRequestProperty, "Townhomes", "", 1, 1)
            
            End If
            
            sht.Name = PayRequestProperty 'temp
            
            
            
            
            'Set the Project field to the name of the property
            sht.Range("C3").Value = PayRequestProperty
            'Set the Ship Date field to the date
            sht.Range("F3").Value = payRequestDate
            'Set the ID field to the PayRequestID
            sht.Range("H2").Value = payRequestID
            
            infoSubtype = Range("E" & myRow.row).Value
            infoModel = Range("F" & myRow.row).Value
            infoUID = Range("G" & myRow.row).Value
            infoQuantity = Range("K" & myRow.row).Value
            infoReference = Range("J" & myRow.row).Value
            
            sht.Range("D7").Value = infoSubtype
            sht.Range("E7").Value = infoModel
            sht.Range("F7").Value = infoUID
            sht.Range("B7").Value = infoQuantity
            sht.Range("H7").Value = infoReference
            
            
            'We now determine the GL Account code
            If sht.Range("D7").Value = "Licenses" Then
                sht.Range("G7").Value = "11111"
                
            ElseIf sht.Range("D7").Value = "Printers" Then
                sht.Range("G7").Value = "22222"
                
            ElseIf sht.Range("D7").Value = "Networking" Then
                sht.Range("G7").Value "33333"
                
            ElseIf sht.Range("D7").Value = "Telecom" Then
                sht.Range("G7").Value = "33333"
            
            Else
                sht.Range("G7").Value = "44444"
            
            End If
            
            
            'Move this data to the existing sheet
            'CODE
        End If
     
    Next
    
    
    
    'Update the UploadData sheet
    Worksheets("UploadData").Range("I2") = payRequestID
    
    
    
    'Go through all the sheets, convert to excel file, change name, etc.
    
    Dim ws As Worksheet
    
    'Ignores the "Are you sure you want to delete?" messages
    Application.DisplayAlerts = False
    
    'Let us loop through all the worksheets in this current worksheet
    For Each ws In ThisWorkbook.Worksheets
    
        'Variable of this current worksheet
        Dim thisWb As Workbook
        Workbooks("AutomatePayRequest.xlsm").Activate
        Set thisWb = ActiveWorkbook
        
        'Variable for the new workbook we will create
        Dim wb As Workbook
        
        'Any sheet after the UploadData sheet, we want to convert
        'PROBLEM
        If ws.Index > thisWb.Sheets("UploadData").Index Then
            
            'This part determines the name of the new workbook
            sheetPropertyname = ws.Name
            sheetDate = ws.Range("F3").Value
            newWorkbookName = sheetPropertyname + " " + sheetDate
            
            'We copy the TotalCostCalculator file and the new PayRequestForm sheet
            Set wb = Workbooks.Add
            thisWb.Sheets("TotalCostCalculator").Copy After:=wb.Sheets(1)
            ws.Copy Before:=wb.Sheets(1)
            
            'Since we created a new workbook, it came with a default Sheet1. We delete it
            wb.Sheets("Sheet1").Delete
            
            'We also want to rename the PayRequestForm sheet properly
            wb.Sheets(1).Name = "PayRequestForm"
            
            
            'Turn screen updating on
            Application.ScreenUpdating = True
            
            'New workbook is good; we must now save it
            wb.SaveAs Filename:=thisWb.Path & "\" & newWorkbookName & ".xlsx"
            
            'Close workbook
            wb.Close
            
            
            'Turn screen updating off
            Application.ScreenUpdating = False
            
            'We delete the old worksheet as we have now created and saved a new workbook
            ws.Delete
            
        End If
    Next
    
    'Re-enable the display alerts
    Application.DisplayAlerts = True
    
    'Re-enable screenUpdating
    Application.ScreenUpdating = True
        
  
End Sub

                        
Sub Button5_Click()
    Rows("6:" & Rows.Count).Delete
End Sub
