Attribute VB_Name = "Modul2"
Option Explicit
Sub main()
    
    Dim StartTime As Double
    Dim EndTime As Double
    StartTime = Timer
        
    Dim C() As Double 'cost matrix
    Dim A() As Integer 'covered matrix
    Dim X() As Integer 'matrix with marked zero-entries
    
    Dim N As Integer 'Dimension
    
    Dim i As Integer, i2 As Integer 'to count rows
    Dim j As Integer, j2 As Integer 'to count columns
    
    Dim min As Integer 'minimum value
    
    'Define dimension of the quadratic matrix by using a cell in your worksheet
    N = Cells(2, 1).Value - 1
    
    
    ReDim C(N, N) As Double
    ReDim A(N, N) As Integer
    ReDim X(N, N) As Integer
    
    Dim zeros As Integer 'to count zeros in rows/columns
    Dim row As Integer 'to save row of found zero
    Dim col As Integer 'to save column of found zero
    Dim covered_zeros As Integer 'to count covered zeros
    Dim marked_zeros As Integer 'to cout the marked zeros
    Dim sum_zeros As Integer 'to count number of 0s in matrix
    Dim found As Boolean 'to check if row/column scanning has changed anything
    
    
    'Read Matrix
    For i = 0 To N
        For j = 0 To N
            C(i, j) = Cells(4 + i, 3 + j).Value
        Next
    Next

'Phase 1: row and column reduction
    'row reduction
    For i = 0 To N
        min = 9999
        For j = 0 To N
            If C(i, j) < min Then
                min = C(i, j)
            End If
        Next j
     
        'subtract min value from each element in this row
        For j = 0 To N
            C(i, j) = C(i, j) - min
        Next j
    Next i

   'column reduction
    For j = 0 To N
        min = 9999
        For i = 0 To N
            If C(i, j) < min Then
                min = C(i, j)
            End If
        Next i
        
        'subtract min value from each element in this column
        For i = 0 To N
            C(i, j) = C(i, j) - min
        Next i
    Next j
    'new matrix is ready; number of zeros is saved
   
   
    'PHASE 2: Optimization
    'Draw minimum number of lines to cover all zeros
    'if we cannot cover all zeros with the following procedure: repeat Step1
    'if Step1 doesn't change anything: diagonal selection rule!
    
Step1:
    'Init A = matrix for covered rows and columns
    'Init X = matrix for marked zero-entries
    For i = 0 To N
        For j = 0 To N
            A(i, j) = 0
            X(i, j) = 0
        Next
    Next
    
    'row scanning
    marked_zeros = 0
    covered_zeros = 0
    sum_zeros = 0

RepeatScanning:
    found = False
    For i = 0 To N
        zeros = 0
        For j = 0 To N
            If C(i, j) = 0 Then
                sum_zeros = sum_zeros + 1 'count all zeros in matrix C
                If A(i, j) = 0 Then 'check wether not yet covered
                    zeros = zeros + 1
                    row = i
                    col = j
                End If
            Else
                'skip value
            End If
        Next j
        If zeros = 1 Then
            'Mark this zero in X
            X(row, col) = 1
            marked_zeros = marked_zeros + 1
            found = True
            'vertical line of in A
            For i2 = 0 To N
                If C(i2, col) = 0 Then
                    covered_zeros = covered_zeros + 1
                End If
                A(i2, col) = A(i2, col) + 1
            Next
           
        Else
            'skip this row
        End If
    Next i
    
    'Check wether all zeros are covered
    
    If covered_zeros = sum_zeros Then
        GoTo Step2
    Else
        'do column checking
        For j = 0 To N
            zeros = 0
            For i = 0 To N
                If C(i, j) = 0 And A(i, j) = 0 Then 'zero found
                    zeros = zeros + 1
                    row = i
                    col = j
                Else
                    'skip this value
                End If
            Next i
        If zeros = 1 Then
            'mark this zero in X
            X(row, col) = 1
            found = True
            marked_zeros = marked_zeros + 1
            'horizontal line in A
            For j2 = 0 To N
                If C(row, j2) = 0 And A(row, j2) = 0 Then
                    covered_zeros = covered_zeros + 1
                End If
                A(row, j2) = A(row, j2) + 1 'double covered entries should now have value of 2
            Next j2
        Else
            'skip this column
        End If
        Next j
    
    End If

    'Column checking done
    
    If covered_zeros = sum_zeros Then
        GoTo Step2
    Else
        If found = True Then
            GoTo RepeatScanning
        Else
            'found = false; scanning did not change anything; Do diagonal selection rule:
            'for each row: if zero found check if there is a diagonal neighbour 0,
            'if yes. mark this 0; cover by vertical line
            'if only one zero is in this row, mark it anyway; cover by vertical line
            'if last row then mark first uncovered zero
            'Go To Step2
            
            For i = 0 To N
                For j = 0 To N
                    If C(i, j) = 0 Then
                        If A(i, j) = 0 Then 'check wether not yet covered
                        
                            row = i
                            col = j
                            If i = N Then
                                col = j
                                row = i
                            ElseIf j <> 0 And j <> N Then 'forward and backward checking
                                If C(i + 1, j + 1) = 0 And A(i + 1, j + 1) = 0 _
                                Or C(i + 1, j - 1) = 0 And A(i + 1, j - 1) = 0 Then 'uncovered forward neighbour zero
                                    row = i
                                    col = j
                                End If
                            Else
                                If j = 0 Then 'only check forward
                                    If C(i + 1, j + 1) = 0 And A(i + 1, j + 1) = 0 Then
                                        row = i
                                        col = j
                                        
                                    End If
                                Else '(j = n); only backward checking
                                    If C(i + 1, j - 1) = 0 And A(i + 1, j - 1) = 0 Then
                                        row = i
                                        col = j
                                        
                                    End If
                                End If
                            End If
                        Else
                        'already covered; skip
                        End If
                    Else
                        'skip value as it is not 0
                    End If
                Next j
                
                X(row, col) = 1 'marked
                For i2 = 0 To N
                    A(i2, col) = A(i2, col) + 1
                Next i2
                Next i
                
        End If
    End If
    GoTo Step5
    
    
Step2:
    If marked_zeros = N + 1 Then 'Done
        GoTo Step5
    Else
        GoTo Step3
    End If
    
Step3:
    'Identify min value of undeleted values
   min = 9999
   For i = 0 To N
        For j = 0 To N
            If A(i, j) = 0 And C(i, j) < min Then
                min = C(i, j)
            Else
                'skip value
            End If
        Next j
    Next i
    
    'subtract min from each uncovered; add min to each double covered value
    For i = 0 To N
        For j = 0 To N
            If A(i, j) = 2 Then
                C(i, j) = C(i, j) + min
            ElseIf A(i, j) = 0 Then
                C(i, j) = C(i, j) - min
            Else
                'value remains same
            End If
        Next j
    Next i
    
    'new matrix ready
    GoTo Step1

Step5: 'Write Solution
'   DONE -> X represents the optimal solution
    For i = 0 To N
        For j = 0 To N
            Cells(N + 8 + i, 3 + j).Value = X(i, j)
            If X(i, j) = 1 Then
                Cells(4 + i, 3 + j).Interior.Color = vbYellow
            End If
        Next
    Next
    
    EndTime = Round(Timer - StartTime, 2)
    Cells(4, 1).Value = "Runtime:"
    Cells(5, 1).Value = EndTime & " sec"
    
    Exit Sub
    
    
End Sub

