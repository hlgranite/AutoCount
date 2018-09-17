Attribute VB_Name = "Module1"
' Auto fill stock card
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
