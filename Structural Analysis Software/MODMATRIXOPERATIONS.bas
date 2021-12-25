Attribute VB_Name = "modMatrixOperations"
Option Explicit
'option base 1
Dim matAug() As Double
Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long



Sub MatInv(matX() As Double, Result() As Double)
On Error GoTo ErrorHandler

'READ DIMENSIONS OF INPUT MATRIX
'--------------------------------------------------------------
Dim NRowX, NColX As Long
    NRowX = UBound(matX, 1)
    NColX = UBound(matX, 2)
ReDim Result(NRowX, NColX)
'--------------------------------------------------------------

'CHECK IF THE MATRIX IS SQUARE
'--------------------------------------------------------------
If NRowX <> NColX Then
    MsgBox "The Coefficient matrix passed to the Function EqSolve is not a Square Matrix"
    ErrorOccured = True
    Exit Sub
End If

'--------------------------------------------------------------


'FORM THE AUGMENTED MATRIX
'--------------------------------------------------------------
Dim NRowAug, NColAug As Long
NRowAug = NRowX
NColAug = NColX * 2
ReDim matAug(NRowAug, NColAug)
For i = 1 To NRowAug
    For j = 1 To NColAug / 2
        matAug(i, j) = matX(i, j)
    Next j
Next i

For k = 1 To NRowAug
    matAug(k, NColAug / 2 + k) = 1
Next k
'--------------------------------------------------------------

'FORWARD REDUCTION OF AUGMENT MATRIX
'--------------------------------------------------------------
Dim PivotElem, Divisor As Double
For i = 1 To NRowAug
    PivotElem = matAug(i, i)
    For j = i To NColAug
        matAug(i, j) = matAug(i, j) / PivotElem
    Next j
    For m = i + 1 To NRowAug
        Divisor = matAug(m, i)
        For n = i To NColAug
            matAug(m, n) = matAug(m, n) - Divisor * matAug(i, n)
        Next n
    Next m
Next i
'--------------------------------------------------------------

'BACKWARD REDUCTION OF AUGMENT MATRIX
'--------------------------------------------------------------
'Dim PivotElem, Divisor As Double
For i = NRowAug To 1 Step -1
    PivotElem = matAug(i, i)
    For j = NColAug To i Step -1
        matAug(i, j) = matAug(i, j) / PivotElem
    Next j
    For m = i - 1 To 1 Step -1
        Divisor = matAug(m, i)
        For n = NColAug To i Step -1
            matAug(m, n) = matAug(m, n) - Divisor * matAug(i, n)
        Next n
    Next m
Next i
'--------------------------------------------------------------

'EXTRACT THE INVERTED MATRIX ASSIGN TO RESULT
'--------------------------------------------------------------
ReDim Result(NRowX, NColX)
For i = 1 To NRowX
    For j = 1 To NColX
        Result(i, j) = matAug(i, j + NColAug / 2)
    Next j
Next i

'--------------------------------------------------------------


'--------------------------------------------------------------
Exit Sub
ErrorHandler:
Call subDispErrInfo("matrix inversion", Err.Number, Err.Description)
End Sub


Sub MatMult(MatA() As Double, MatB() As Double, Result() As Double)
On Error GoTo ErrorHandler

Dim NRowA, NRowB, NColA, NColB As Long
NRowA = UBound(MatA, 1)
NColA = UBound(MatA, 2)
NRowB = UBound(MatB, 1)
NColB = UBound(MatB, 2)

ReDim Result(NRowA, NColB)

For i = 1 To NRowA
    For j = 1 To NColB
         For k = 1 To NRowA
            Result(i, j) = Result(i, j) + MatA(i, k) * MatB(k, j) '+ z
        Next
    Next
Next

Exit Sub
ErrorHandler:
Call subDispErrInfo("matrix multiplication", Err.Number, Err.Description)
End Sub


Sub MatTranspose(Mat() As Double, Result() As Double)
On Error GoTo ErrorHandler

    Dim NRow, NCol As Long
        NRow = UBound(Mat, 1)
        NCol = UBound(Mat, 2)
    ReDim Result(NCol, NRow)
    
    Dim i, j As Long
        For i = 1 To NRow
            For j = 1 To NCol
                Result(j, i) = Mat(i, j)
            Next j
        Next i
        
Exit Sub
ErrorHandler:
Call subDispErrInfo("calculating transpose of a matrix", Err.Number, Err.Description)
End Sub


Sub EqSolve(matX() As Double, MatB() As Double, Result() As Double)
On Error GoTo ErrorHandler

Dim i, j, k, l, m, n As Long

'READ DIMENSIONS OF INPUT MATRICES
'--------------------------------------------------------------
Dim NRowX, NColX, NRowB, NColB As Long
NRowX = UBound(matX, 1)
NColX = UBound(matX, 2)
NRowB = UBound(MatB, 1)
NColB = UBound(MatB, 2)
'--------------------------------------------------------------

'CHECK FOR COMPATIBILITY
'--------------------------------------------------------------
If NRowX <> NColX Then
    MsgBox "The Coefficient matrix passed to the Function EqSolve is not a Square Matrix"
    ErrorOccured = True
    Exit Sub
End If


If NColX <> NRowB Then
    MsgBox "System of Equations passed to EqSolve Function are dimensionally incompatible"
    ErrorOccured = True
    Exit Sub
End If

If NColB > 1 Then
    MsgBox "The Constants matrix passed to the function EqSolve is not a single column matrix"
    ErrorOccured = True
    Exit Sub
End If
'--------------------------------------------------------------


'FORM THE AUGMENTED MATRIX
'--------------------------------------------------------------
Dim NRowAug, NColAug As Long
NRowAug = NRowX
NColAug = NColX + 1
ReDim matAug(NRowAug, NColAug)
For i = 1 To NRowAug
    For j = 1 To NColAug - 1
        matAug(i, j) = matX(i, j)
    Next j
Next i

For k = 1 To NRowAug
    matAug(k, NColAug) = MatB(k, 1)
Next k
'--------------------------------------------------------------

'FORWARD REDUCTION OF AUGMENT MATRIX
'--------------------------------------------------------------
Dim PivotElem, Divisor As Double
For i = 1 To NRowAug
    PivotElem = matAug(i, i)
    For j = i To NColAug
        matAug(i, j) = matAug(i, j) / PivotElem
    Next j
    For m = i + 1 To NRowAug
        Divisor = matAug(m, i)
        For n = i To NColAug
            matAug(m, n) = matAug(m, n) - Divisor * matAug(i, n)
        Next n
    Next m
Next i
'--------------------------------------------------------------

'BACKWARD REDUCTION OF AUGMENT MATRIX
'--------------------------------------------------------------
'Dim PivotElem, Divisor As Double
For i = NRowAug To 1 Step -1
    PivotElem = matAug(i, i)
    For j = NColAug To i Step -1
        matAug(i, j) = matAug(i, j) / PivotElem
    Next j
    For m = i - 1 To 1 Step -1
        Divisor = matAug(m, i)
        For n = NColAug To i Step -1
            matAug(m, n) = matAug(m, n) - Divisor * matAug(i, n)
        Next n
    Next m
Next i
'--------------------------------------------------------------


'SEND OUTPUT
'--------------------------------------------------------------
ReDim Result(NRowB, 1)
For i = 1 To NRowB
    Result(i, 1) = matAug(i, NColAug)
Next i

'--------------------------------------------------------------
Exit Sub
ErrorHandler:
Call subDispErrInfo("solving system of equations", Err.Number, Err.Description)
End Sub

