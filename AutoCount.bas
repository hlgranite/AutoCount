Attribute VB_Name = "Module1"
' Auto fill stock card
' TODO: Set activesheet
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
                output.Cells(x, 1) = customer
                output.Cells(x, 2) = dateHeaders(j)
                ' grab sales amount per customer per month
                output.Cells(x, 3) = source.Cells(i, headers(j)).Value
                ' grab collection amount per customer per month
                output.Cells(x, 4) = source.Cells(i + 2, headers(j)).Value
                x = x + 1
            Next j
        End If
    Next i

End Sub

