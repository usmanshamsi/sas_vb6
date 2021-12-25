Attribute VB_Name = "modMain"
Option Base 1
Option Explicit
Dim x(3, 3), b(3), res() As Double
Dim sol() As Double
'Dim x(2, 2), b(2), res() As Double

Dim matAug() As Double
Dim txt As String
Sub main()
Dim i, j As Long
'assign arbitrary values to matrices
'x
x(1, 1) = 400
x(1, 2) = -200
x(1, 3) = 0
x(2, 1) = -200
x(2, 2) = 400
x(2, 3) = -200
x(3, 1) = 0
x(3, 2) = -200
x(3, 3) = 400


'b
b(1) = 0
b(2) = 0
b(3) = 4

res = MatMult(x(), x())
MsgBox UBound(res, 1)
MsgBox UBound(res, 2)
MsgBox UBound(b)
'sol = MatMult(res(), b())
'print res23
'MsgBox UBound(b, 1)
For i = 1 To 3
    For j = 1 To 3
        txt = txt & res(i, j) & "       "
    Next j
    txt = txt & vbCrLf
Next i
MsgBox txt

End Sub


