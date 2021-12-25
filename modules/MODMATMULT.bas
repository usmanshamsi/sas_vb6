Attribute VB_Name = "modMatMult"
Option Base 1
Private Result() As Double
Dim NRowA, NRowB, NColA, NColB As Long


Function MatMult(MatA() As Double, MatB() As Double) As Double()

NRowA = UBound(MatA, 1)
NColA = UBound(MatA, 2)
NRowB = UBound(MatB, 1)
'NColB = UBound(MatB, 2)

ReDim Result(NRowA, NColB)

For i = 1 To NRowA
    For j = 1 To NColB
         For k = 1 To NRowA
            Result(i, j) = Result(i, j) + MatA(i, k) * MatB(k, j) '+ z
        Next
    Next
Next

MatMult = Result()
    
End Function
