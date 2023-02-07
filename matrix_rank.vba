Sub mysub()
Dim m As Integer, n As Integer, rank As Integer, col_max As Double, i_max As Integer, Row As Integer, temp As Integer, prev_non_zero As Integer
    rank = 0
    m = 0
    n = 0
    prev_non_zero = 1001
    Dim A(1000, 1000) As Double

    For col = 1 To 1000
        If IsEmpty(Cells(1, col)) Then
            n = col - 1
            Exit For
        End If
    Next col

    For Row = 1 To 1000
        If IsEmpty(Cells(Row, 1)) Then
            m = Row - 1
            Exit For
        End If
    Next Row

    For col = 1 To n
        For Row = 1 To m
            A(Row, col) = Cells(Row, col).Value
        Next Row
    Next col

    Row = 1
    col = 1

    Do While col <= n And Row <= m
        i_max = Row
        col_max = 0

        For i = Row To m
            If Abs(A(i, col)) > col_max Then
                col_max = Abs(A(i, col))
                i_max = i
            End If
        Next i

        If A(i_max, col) <> 0 Then

            For j = 1 To n
                temp = A(i_max, j)
                A(i_max, j) = A(Row, j)
                A(Row, j) = temp
            Next j

            For i = Row + 1 To m
                For j = col + 1 To n
                    A(i, j) = A(i, j) - A(i, col) * A(Row, j) / A(Row, col)
                Next j
                A(i, col) = 0
            Next i

            Row = Row + 1
        End If

        col = col + 1
    Loop

    For Row = 1 To m
        temp = 0

        For col = 1 To n
        
            If A(Row, col) <> 0 Then
                If temp = 0 Then prev_non_zero = col
                temp = temp + 1
            ElseIf temp <> 0 Then
                temp = temp + 1
            End If
            
            If temp = 0 Then
                Cells(Row + m + 1, col).Interior.Color = RGB(189, 215, 238)
                If col >= prev_non_zero Then Cells(Row + m + 1, col).Borders(xlEdgeTop).Weight = xlMedium
            End If
            
            If temp = 1 And col > 1 Then Cells(Row + m + 1, col - 1).Borders(xlEdgeRight).Weight = xlMedium
            
            Cells(Row + m + 1, col).Value = A(Row, col)
        Next col
        
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
    
    For col = 1 To n
        Cells(m, col).Borders(xlEdgeBottom).Weight = xlThick
        Cells(m + 1, col).Borders(xlEdgeBottom).Weight = xlThick
        Cells(2 * m + 1, col).Borders(xlEdgeBottom).Weight = xlThick
    Next col

    Cells(2 * m + 3, 1).Value = rank
    Cells(2 * m + 3, 1).Interior.Color = RGB(228, 255, 88)
    Cells(2 * m + 3, 1).Borders(xlEdgeBottom).Weight = xlThick
    Cells(2 * m + 3, 1).Borders(xlEdgeRight).Weight = xlThick
    Cells(2 * m + 3, 1).Borders(xlEdgeTop).Weight = xlThick
End Sub
