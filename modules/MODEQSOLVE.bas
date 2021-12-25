Attribute VB_Name = "modEqSolve"
Option Base 1
Option Explicit
Dim matAug() As Double

Function EqSolve(matX(), matB()) As Double()
Dim i, j, k, l, m, n As Long

'READ DIMENSIONS OF INPUT MATRICES
'--------------------------------------------------------------
Dim NRowX, NColX, NRowB, NColB As Long
NRowX = UBound(matX, 1)
NColX = UBound(matX, 2)
NRowB = UBound(matB, 1)
'NColB = UBound(matB, 2)
'--------------------------------------------------------------

'CHECK FOR COMPATIBILITY
'--------------------------------------------------------------
If NRowX <> NColX Then
    MsgBox "The Coefficient matrix passed to the Function EqSolve is not a Square Matrix"
    'ErrorOccured=True
    Exit Function
End If


If NColX <> NRowB Then
    MsgBox "System of Equations passed to EqSolve Function are dimensionally incompatible"
    'ErrorOccured=True
    Exit Function
End If

If NColB > 1 Then
    MsgBox "The Constants matrix passed to the function EqSolve is not a single column matrix"
    'ErrorOccured=True
    Exit Function
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
    matAug(k, NColAug) = matB(k)
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
EqSolve = matAug()
'--------------------------------------------------------------

End Function
