Sub mysub()
    Dim m As Long, n As Long, rank As Long, col_max As Double, i_max As Long, Row As Long, temp As Long, prev_non_zero As Long, TypeError As Boolean
    Dim A() As Double
    rank = 0
    TypeError = False
    prev_non_zero = Rows.Count + 1

    m = Cells(1, 1).End(xlDown).Row
    n = Cells(1, 1).End(xlToRight).Column
    ReDim A(m, n)
    
    Range(Cells(m + 1, 1), Cells(Rows.Count, 1)).EntireRow.Delete
    Range(Cells(1, n + 1), Cells(1, Columns.Count)).EntireColumn.Delete
    Range(Cells(1, 1), Cells(m, n)).ClearFormats
    
    For Col = 1 To n
        For Row = 1 To m
            If IsNumeric(Cells(Row, Col).Value) And (Not IsEmpty(Cells(Row, Col).Value)) Then
                A(Row, Col) = Cells(Row, Col).Value
            Else
                Cells(Row, Col).Interior.Color = RGB(250, 20, 40)
                Cells(Row, Col).Font.Color = RGB(255, 255, 255)
                TypeError = True
            End If
        Next Row
    Next Col
    
    If TypeError Then
        MsgBox "Not all matrix elements are numbers. Please correct and repeat.", vbInformation + vbOKOnly, "Error"
        Exit Sub
    End If

    Row = 1
    Col = 1

    Do While Col <= n And Row <= m
        i_max = Row
        col_max = 0

        For i = Row To m
            If Abs(A(i, Col)) > col_max Then
                col_max = Abs(A(i, Col))
                i_max = i
            End If
        Next i

        If A(i_max, Col) <> 0 Then

            For j = 1 To n
                temp = A(i_max, j)
                A(i_max, j) = A(Row, j)
                A(Row, j) = temp
            Next j

            For i = Row + 1 To m
                For j = Col + 1 To n
                    A(i, j) = A(i, j) - A(i, Col) * A(Row, j) / A(Row, Col)
                Next j
                A(i, Col) = 0
            Next i

            Row = Row + 1
        End If

        Col = Col + 1
    Loop

    For Row = 1 To m
        temp = 0

        For Col = 1 To n
        
            If A(Row, Col) <> 0 Then
                If temp = 0 Then prev_non_zero = Col
                temp = temp + 1
            ElseIf temp <> 0 Then
                temp = temp + 1
            End If
            
            If temp = 0 Then
                Cells(Row + m + 1, Col).Interior.Color = RGB(189, 215, 238)
                If Col >= prev_non_zero Then Cells(Row + m + 1, Col).Borders(xlEdgeTop).Weight = xlMedium
            End If
            
            If temp = 1 And Col > 1 Then Cells(Row + m + 1, Col - 1).Borders(xlEdgeRight).Weight = xlMedium
            
            Cells(Row + m + 1, Col).Value = A(Row, Col)
        Next Col
        
        If temp > 0 Then
            rank = rank + 1
        Else
            prev_non_zero = 1001
        End If
    Next Row
    
    For Row = 1 To m
        Cells(Row, n).Borders(xlEdgeRight).Weight = xlThick
        Cells(Row + m + 1, n).Borders(xlEdgeRight).Weight = xlThick
    Next Row
    
    For Col = 1 To n
        Cells(m, Col).Borders(xlEdgeBottom).Weight = xlThick
        Cells(m + 1, Col).Borders(xlEdgeBottom).Weight = xlThick
        Cells(2 * m + 1, Col).Borders(xlEdgeBottom).Weight = xlThick
    Next Col
    
    With Cells(2 * m + 3, 1)
        .Value = rank
        .Interior.Color = RGB(228, 255, 88)
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeRight).Weight = xlThick
        .Borders(xlEdgeTop).Weight = xlThick
    End With
    
    MsgBox "The rank is " & rank, vbInformation + vbOKOnly, "The result"
End Sub
