VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CompanyBill 
   Caption         =   "UserForm1"
   ClientHeight    =   9240.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8100
   OleObjectBlob   =   "CompanyBill.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CompanyBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Instead of referring to userform as companybill refer to it as Me, as this code is specific to this userform
Private Sub UserForm_Initialize()
    ' Populating the combobox using while loop instead of for loop as more data can be added later

    ' Populating the combobox using  existing data
    i = 2
    Do While Sheets("Client List").Cells(i, 1).Value <> ""
        items = Sheets("Client List").Cells(i, 1).Value
        CompanyBill.CBO_ClientName.AddItem items
        i = i + 1
    Loop
    
    i = 2
    Do While Sheets("Product List").Cells(i, 1).Value <> ""
        Item = Sheets("Product List").Cells(i, 1).Value
        CompanyBill.CBO_ProductName.AddItem Item
        i = i + 1
    Loop
    
    ' populating the combobox using array: use combobox.list
    CompanyBill.CMBO_Terms.List = Array(1, 2, 3, 7, 14, 30)
  
End Sub


Private Sub CMD_InvoiceNumber_Click()

    ' check "invoice Tracker" sheet to find any values, if not found invoice number = 10000
    ' else take the last invoice number and add 1 to it
    If Sheets("Invoice Tracker").[A2] <> "" Then
        Sheets("Invoice").[F7].Value = Sheets("Invoice Tracker").[A1].End(xlDown).Value + 1
    Else:
        Sheets("Invoice").[F7].Value = 10000
    End If
       
End Sub


Private Sub CMD_AddRow_Click()
    
    ' Populate the Invoive with the values from the userform
    With Sheets("Invoice")
        .[B12].Value = CompanyBill.CBO_ClientName.Value
        .[B30].End(xlUp).Offset(1, 0).Value = CompanyBill.CBO_ProductName.Value
        .[B30].End(xlUp).Offset(0, 2) = CompanyBill.TBO_Qty.Value
    End With
    
End Sub


Private Sub CMD_DeleteLatest_Click()
    
    ' Check to see if the invoice list(product List) is empty, if not then remove that row
    If Sheets("Invoice").Range("B19").Offset(1, 0).Value <> "" Then
        Sheets("Invoice").Range("B19").End(xlDown).Select
        Range(Selection, Selection.Offset(0, 4)).SpecialCells(xlCellTypeConstants).ClearContents
    ' If there are no values found then return the msgbox
    Else:
        MsgBox "No Records Detected"
    End If

End Sub

' Option button, if any of the below discound button is checked the discount amount will be applied
Private Sub OPN_Discount5_Click()
    
    Sheets("Invoice").[F32].Value = 5 / 100

End Sub

Private Sub OPN_Discount10_Click()

    Sheets("Invoice").[F32].Value = 10 / 100

End Sub


Private Sub OPN_Discount15_Click()

    Sheets("Invoice").[F32].Value = 15 / 100

End Sub

' Popualte the terms with the value from the combobox
Private Sub CMD_Terms_Click()

    Sheets("Invoice").[E9].Value = CompanyBill.CMBO_Terms.Value

End Sub


Private Sub CBO_Submit_Click()
    ' Confirmation msg
    submit_status = MsgBox("Are You Sure", vbYesNo)
    
    If submit_status = 6 Then ' 6 = "Yes"
        updateTracker    ' function calling
        Unload CompanyBill ' clear all the contents from the userform for fresh entry
    End If
    
End Sub


Private Sub CMD_submit_and_print_Click()

    submit_status = MsgBox("Are You Sure", vbYesNo)
    
    If submit_status = 6 Then
        pdfGenerator
        Unload CompanyBill
    End If

End Sub

Private Sub CBO_Cancel_Click()

    CompanyBill.Hide

End Sub


