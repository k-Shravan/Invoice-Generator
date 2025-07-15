Attribute VB_Name = "tracker_and_generator"
Sub updateTracker()

    ' Update the tracker with the invoice data (Button)

    ' Store the necessary values in variable, we can directly paste the values but this is more readable
    invoice_Number = Sheets("Invoice").[F7].Value
    company_Name = Sheets("Invoice").[B12].Value
    amount = Sheets("Invoice").[F34].Value
    Date_issued = Sheets("Invoice").[F5].Value
    Date_due = Sheets("Invoice").[F9].Value
    
    Sheets("Invoice Tracker").Select
    ' select the last cell, Go up(until a value is found), then select the cell down below
    Range("A1048576").End(xlUp).Offset(1, 0).Select
    
    ' offset(row, column)
    With Selection
        .Offset(0, 0) = invoice_Number  ' initial point (0, 0)
        .Offset(0, 1) = company_Name    ' move horizontally by one point (column value changes)
        .Offset(0, 2) = amount
        .Offset(0, 3) = Date_issued
        .Offset(0, 4) = Date_due
        .Offset(0, 6) = "Yes"
        .Offset(0, 7) = "No"            ' pdf generation is "No" (we are just updating the tracker
    End With
    
    Sheets("Invoice").Select
    
    MsgBox "Tracker Has been Updated, pdf generation required"

End Sub

Sub pdfGenerator()
    
    ' same as updatetracker() but here pdf generation value will be "Yes"
    invoice_Number = Sheets("Invoice").[F7].Value
    company_Name = Sheets("Invoice").[B12].Value
    amount = Sheets("Invoice").[F34].Value
    Date_issued = Sheets("Invoice").[F5].Value
    Date_due = Sheets("Invoice").[F9].Value
    
    Sheets("Invoice Tracker").Select
    Range("A1048576").End(xlUp).Offset(1, 0).Select
    
    With Selection
        .Offset(0, 0) = invoice_Number
        .Offset(0, 1) = company_Name
        .Offset(0, 2) = amount
        .Offset(0, 3) = Date_issued
        .Offset(0, 4) = Date_due
        .Offset(0, 6) = "Yes"
        .Offset(0, 7) = "Yes"
    End With
    
    Sheets("Invoice").Select
    ActiveSheet.ExportAsFixedFormat Type:=pdf, Filename:="C:\Users\shrav\Desktop\invoice\Invoice-" & invoice_Number & "-" & company_Name
    
    MsgBox "Invoice tracked and pdf has been generated"

End Sub

Sub openForm()

    ' Opens the UserForm
    CompanyBill.Show

End Sub


Sub clear_invoice()
    
    ' as these cells dont contain any formulas, just use clearcontents
    Range("B12").ClearContents
    Range("F7").ClearContents
    Range("E9").ClearContents
    Range("F32").ClearContents
        
    ' For cells with formulas use method: specialcells(xlCellTypeConstants).clearcontents
    ' xlCellTypeConstants only clears constants and nothing else (Formulas will be preserved)
    
    ' This only checks for the B20 cell, so if there are any values in below cells this wont work
    If Range("B20").Value = "" Then
        MsgBox "No values Found in list"
    Else:
        Sheets("Invoice").Range("B20").Select            ' Select the starting cell
        Range(Selection, Selection.Offset(0, 4)).Select  ' select 4 cells to the right (column movement)
        Range(Selection, Selection.Offset(10, 0)).Select ' select 10 cells down (Row movement)
        Selection.Special Cells(xlCellTypeConstants).ClearContents    ' Clear all the constants from the selected range(Npt formulas)
    End If
    
    [A1].Select
    
End Sub
