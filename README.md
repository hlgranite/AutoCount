# AutoCount
AutoCount related script.

## Stock Card
1. Autocount > Stock > Reports > Stock Card.
2. Set filter date > Inquiry.
3. Print > Download as Excel (*.xlsx).
4. Done.

Prepare the excel datasheet in this columns sequence: Date, Stock, Type, UOM, Document, Customer, Description, Qty, Unit Price, Amount. Then add a VBA script to a button click.

```vb
Sub AutoFill()

    Dim order As String
    Dim stock As String
    Dim desc As String
    Dim uom As String
    
    Dim i As Integer
    For i = 2 To 15000
        If (Not IsEmpty(Range("B" & i))) Then
            stock = Range("B" & i).Value
        End If
        
        If (Not IsEmpty(Range("G" & i))) Then
            desc = Range("G" & i).Value
        End If
        
        If (Not IsEmpty(Range("D" & i))) Then
            uom = Range("D" & i).Value
        End If
        
        order = Range("A" & i).Value
        If (order <> "" And order <> "HQ" And order <> "Item :") Then
            Range("B" & i).Value = stock
            Range("D" & i).Value = uom
            Range("G" & i).Value = desc
        End If
    Next i

End Sub
```

## Sales & Collection
1. Autocount > A/R > Monthly Sales & Collection.
2. Pick date range > Preview > Export to XLSX.
3. Use macro below to parse xlsx into columns for further manipulation.
	- customer
	- month
	- sales
	- collection
4. Done.

```vb
' parse monthly sales and collection into fact records for pivot table uses
Sub ParseSalesCollection()

    Dim source As Worksheet
    Set source = Sheets("source")
    
    ' add new worksheet at last
    Dim output As Worksheet
    Set output = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    output.Name = "output"
    
    Dim dateHeaders(1 To 12) As String
    Dim headers(1 To 12) As Integer
    Dim customer As String
    
    ' retrieve date collection
    Let j = 1
    For i = 2 To 50
        If (source.Cells(12, i).Value = "Total") Then
            Exit For
        End If
        
        If (Not IsEmpty(source.Cells(12, i))) Then
            headers(j) = i
            dateHeaders(j) = source.Cells(12, i).Value
            j = j + 1
        End If
    Next i
    
    ' output row to be insert
    Let x = 1
    
    ' retrieve sales and collection once found new customer
    For i = 13 To 1000
        If (Not IsEmpty(source.Cells(i, 1))) Then
            customer = source.Cells(i, 1).Value
            
            If (customer = "Report Criteria") Then
                Exit For
            End If
            
            For j = 1 To 12
                Dim sales As Double
                Dim collection As Double
                
                ' grab sales amount per customer per month
                sales = source.Cells(i, headers(j)).Value
                
                ' grab collection amount per customer per month
                collection = source.Cells(i + 2, headers(j)).Value
                
                ' only insert non zero record
                If (sales > 0 And collection > 0) Then
                    output.Cells(x, 1) = customer
                    output.Cells(x, 2) = dateHeaders(j)
                    output.Cells(x, 3) = sales
                    output.Cells(x, 4) = collection
                    x = x + 1
                End If
            Next j
        End If
    Next i

End Sub
```
