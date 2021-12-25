Attribute VB_Name = "modMatInv"
Option Base 1
Option Explicit
Dim matAug() As Double
Dim Result() As Double

Function MatInv(matX()) As Double()
Dim i, j, k, l, m, n As Long

'READ DIMENSIONS OF INPUT MATRIX
'--------------------------------------------------------------
Dim NRowX, NColX As Long
NRowX = UBound(matX, 1)
NColX = UBound(matX, 2)

'--------------------------------------------------------------

'CHECK IF THE MATRIX IS SQUARE
'--------------------------------------------------------------
If NRowX <> NColX Then
    MsgBox "The Coefficient matrix passed to the Function EqSolve is not a Square Matrix"
    'ErrorOccured=True
    Exit Function
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
'GoTo OUT

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

'EXTRACT THE INVERTED MATRIX
'--------------------------------------------------------------
ReDim Result(NRowX, NColX)
For i = 1 To NRowX
    For j = 1 To NColX
        Result(i, j) = matAug(i, j + NColAug / 2)
    Next j
Next i

'--------------------------------------------------------------

'SEND OUTPUT
OUT:
'--------------------------------------------------------------
'MatInv = matAug()
MatInv = Result()
'--------------------------------------------------------------

End Function

