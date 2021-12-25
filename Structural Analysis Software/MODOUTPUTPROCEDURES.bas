Attribute VB_Name = "modOutputProcedures"
Option Explicit
'option base 1
    Public XMin, YMin, XMax, YMax, XMid, YMid As Double
    Public ScreenWidth, ScreenHeight, CircleSize As Double
    Public ScaleX, ScaleY, PlotScale, Margin As Double


Function Max(List() As Double)
On Error GoTo ErrorHandler

    Dim i As Long
    Max = Val(List(1))
    If UBound(List) = 1 Then Exit Function
    For i = 1 To UBound(List)
        If Val(List(i)) > Max Then Max = Val(List(i))
    Next i
    
Exit Function
ErrorHandler:
Call subDispErrInfo("calculating maximum of a list", Err.Number, Err.Description)
End Function

Function Min(List() As Double)
On Error GoTo ErrorHandler
    Dim i As Long
    Min = Val(List(1))
    If UBound(List) = 1 Then Exit Function
    For i = 1 To UBound(List)
        If Val(List(i)) < Min Then Min = Val(List(i))
    Next i
    
Exit Function
ErrorHandler:
Call subDispErrInfo("calculating minimum of a list", Err.Number, Err.Description)
End Function

Sub subPlotStr()
On Error GoTo ErrorHandler

'DO WITH THE MAIN FORM PICTURE CONTROL
'--------------------------------------
With frmMain.Picture1

.Cls    'CLEAR SCREEN


If NoofNodes < 1 Then Exit Sub
        'SCALE THE SCREEN ACCORDING THE GEOMETRY OF STRUCTURE
        '---------------------------------------------------------
            If NoofNodes > 0 Then
                XMax = Val(Max(Xcoor()))
                XMin = Val(Min(Xcoor()))
                YMax = Val(Max(Ycoor()))
                YMin = Val(Min(Ycoor()))
            End If
            
            'IN ORDER TO AVOID DIVISION BY ZERO ERROR
            '---------------------------------------
            If XMax = XMin Then XMax = XMin + 0.01
            If YMax = YMin Then YMax = YMin + 0.01
                    
            ScaleX = .Width / Abs((XMax - XMin))
            ScaleY = .Height / Abs((YMax - YMin))
            XMid = (XMax + XMin) / 2
            YMid = (YMax + YMin) / 2
            
            If ScaleX < ScaleY Then
                Margin = 0.2 * Abs((XMax - XMin))
                ScreenWidth = Abs(XMax - XMin) + 2 * Margin
                ScreenHeight = ScreenWidth * .Height / .Width
                CircleSize = 0.005 * ((ScreenWidth ^ 2 + ScreenHeight ^ 2) ^ 0.5)
            Else
                Margin = 0.2 * Abs(YMax - YMin)
                ScreenHeight = Abs(YMax - YMin) + 2 * Margin
                ScreenWidth = ScreenHeight * .Width / .Height
                CircleSize = 0.005 * ((ScreenWidth ^ 2 + ScreenHeight ^ 2) ^ 0.5)
            End If
                
            .ScaleLeft = XMid - ScreenWidth / 2
            .ScaleWidth = ScreenWidth
            
            .ScaleTop = YMid - ScreenHeight / 2
            .ScaleHeight = ScreenHeight

        
        'PLOT NODES
        '-----------
        Dim i As Long, j As Long
        Dim xi, yi
            For i = 1 To NoofNodes
            xi = Val(Xcoor(i))
            yi = Val(Ycoor(i))
            
            'NODE MARKS
            '-----------
            frmMain.Picture1.Circle (xi, YMirror(yi)), CircleSize, 17000000
                
            'NODE NUMBERS
            '-------------
            If (PlotBMD = False) And (PlotSF = False) And (PlotAxial = False) Then
                frmMain.Picture1.ForeColor = vbWhite
                frmMain.Picture1.CurrentX = xi - 0.5 * .TextWidth(CStr(i))
                frmMain.Picture1.CurrentY = YMirror(yi + .TextHeight(CStr(i)) * 1.25 + CircleSize)
                frmMain.Picture1.Print CStr(i)
            End If
                
            Next i
            
        
        'PLOT ELEMENTS
        '--------------
        
        Dim X1, Y1, X2, Y2, xm, ym
        Dim section
        Dim Color
            For i = 1 To NoOfElements
                
                'PLOT LINES
                '-----------
                X1 = Val(Xcoor(Endi(i)))
                X2 = Val(Xcoor(Endj(i)))
                Y1 = YMirror(Val(Ycoor(Endi(i))))
                Y2 = YMirror(Val(Ycoor(Endj(i))))
                section = Val(AsgnSec(i))
                Color = Val(SecColor(section))
                frmMain.Picture1.Line (X1, Y1)-(X2, Y2), Color
                
                'PRINT NUMBER
                '---------------
                If (PlotBMD = False) And (PlotSF = False) And (PlotAxial = False) Then
                frmMain.Picture1.ForeColor = vbYellow
                    xm = (X2 + X1) / 2
                    ym = (Y1 + Y2) / 2
                    
                    .CurrentX = xm - .TextWidth(CStr(i))
                    .CurrentY = ym - .TextHeight(CStr(i)) - CircleSize / 5
                    
                    frmMain.Picture1.Print CStr(i)
                End If
                
            
            Next i
            

'-----------------------------------------------

'PLOT REACTIONS OR LOADS

If (PlotBMD = False) And (PlotSF = False) And (PlotAxial = False) _
        And (frmMain.menuShowLoads.Checked = True) Then
        
    MaxLoad = 0.000000001
    
    Call subCalcMaxLoad
    
    LoadScale = 0.5 * Margin / MaxLoad
    
    For i = 1 To NoofNodes
    
        If XForce(i) <> 0 Then
            Call PlotArrow(Xcoor(i), Ycoor(i), 0.6 * Margin, (3.14159 / 2 + Sgn(XForce(i)) * 3.14159 / 2), vbWhite, CStr(XForce(i)))
        End If
            If YForce(i) <> 0 Then
            Call PlotArrow(Xcoor(i), Ycoor(i), 0.6 * Margin, -Sgn(YForce(i)) * 3.1415925654 / 2, vbBlue, CStr(Abs(YForce(i))))
        End If
            If ZMom(i) <> 0 Then
            Call PlotCircularArrow(Xcoor(i), Ycoor(i), Sgn(ZMom(i)), vbRed, CStr(Abs(ZMom(i))))
        End If
        '--------------------
               
        '--------------------
        frmMain.Picture1.DrawStyle = 2
        If TxRest(i) = 1 Then
            Call PlotArrow(Xcoor(i), Ycoor(i), Abs(0.6 * Margin), 3.14159, vbGreen, "H" & i)  'CStr(3 * (i - 1) + 1))
        End If
        If TyRest(i) = 1 Then
            Call PlotArrow(Xcoor(i), Ycoor(i), Abs(0.6 * Margin), 3 * 3.14159 / 2, vbGreen, "V" & i)  'CStr(3 * (i - 1) + 2))
        End If
        If RzRest(i) = 1 Then
            Call PlotCircularArrow(Xcoor(i), Ycoor(i), 1, vbGreen, "M" & i)   'CStr(3 * (i - 1) + 3))
        End If
        frmMain.Picture1.DrawStyle = 0
    Next i
    
    For i = 1 To NoOfElements
    
            Dim LoadCurr As Double
            Dim Length As Double, Theta As Double
            Dim Cx, Cy
            Dim X1_ As Double, X2_ As Double, Y1_ As Double, Y2_ As Double
            X1 = Xcoor(Endi(i)): Y1 = Ycoor(Endi(i))
            X2 = Xcoor(Endj(i)): Y2 = Ycoor(Endj(i))
            If (X2 - X1) < 0.00001 Then
                Theta = 3.14159 / 2
            Else
                Theta = Atn((Y2 - Y1) / (X2 - X1))
            End If
            Length = GetElemLength(i)
            Cx = (X2 - X1) / Length: Cy = (Y2 - Y1) / Length
            
            X1_ = X1 + Cy * LoadScale * ElemLoadi(i)
            Y1_ = Y1 - Cx * LoadScale * ElemLoadi(i)
            
            X2_ = X2 + Cy * LoadScale * ElemLoadj(i)
            Y2_ = Y2 - Cx * LoadScale * ElemLoadj(i)
            
            frmMain.Picture1.Line (X1, YMirror(Y1))-(X1_, YMirror(Y1_)), vbWhite
            
            If ElemLoadi(i) <> 0 Or ElemLoadj(i) <> 0 Then
                frmMain.Picture1.Line (X1_, YMirror(Y1_))-(X2_, YMirror(Y2_)), vbWhite
            End If
            
            frmMain.Picture1.Line (X2_, YMirror(Y2_))-(X2, YMirror(Y2)), vbWhite
            
            frmMain.Picture1.ForeColor = vbWhite
            
            If Not (ElemLoadi(i) = 0 And ElemLoadj(i) = 0) Then
                frmMain.Picture1.CurrentX = X1_
                frmMain.Picture1.CurrentY = YMirror(Y1_ + 4 * CircleSize)
                frmMain.Picture1.Print CStr(Format(Abs(ElemLoadi(i)), "0.000"))
                
                frmMain.Picture1.CurrentX = X2_
                frmMain.Picture1.CurrentY = YMirror(Y2_ + 4 * CircleSize)
                frmMain.Picture1.Print CStr(Format(Abs(ElemLoadj(i)), "0.000"))
            End If
            
            
            For j = 0 To 5
                LoadCurr = Round(ElemLoadi(i) + (ElemLoadj(i) - ElemLoadi(i)) * j / 5, 5)
                Call PlotArrow(X1 + (X2 - X1) * j / 5, (Y1 + (Y2 - Y1) * j / 5), LoadScale * Abs(LoadCurr), Theta - Sgn(LoadCurr) * 3.14159 / 2, _
                    vbWhite)
                LoadCurr = Round(ElemALoadi(i) + (ElemALoadj(i) - ElemALoadi(i)) * j / 5, 5)
                Call PlotArrow(X1 + (X2 - X1) * j / 5, (Y1 + (Y2 - Y1) * j / 5), LoadScale * Abs(LoadCurr), Theta + 3.14159 / 2 + Sgn(LoadCurr) * 3.14159 / 2, _
                    vbWhite)
            Next j
    Next i

End If
'-----------------------------------------------

If (PlotDeflectedShape = True) And (Analyzed = True) Then
        frmMain.Picture1.DrawStyle = 2


        'PLOT DEFLECTED NODES
        '-----------

        
            For i = 1 To NoofNodes
            'xi = Val(Xcoor(i)) + DispVector(3 * (i - 1) + 1, 1) * DeflectionScale
            'yi = YMirror(Val(Ycoor(i)) + DispVector(3 * (i - 1) + 2, 1) * DeflectionScale)
            xi = DXcoor(i)
            yi = DYcoor(i)
                
                
                'NODE MARKS
                '-----------
                frmMain.Picture1.Circle (xi, YMirror(yi)), CircleSize, 65280
                
                'VALUE OF DEFLECTION
                '--------------------
                frmMain.Picture1.ForeColor = vbWhite
                frmMain.Picture1.CurrentX = xi - frmMain.Picture1.TextWidth(CStr("dX= " & Format((DXcoor(i) - Xcoor(i)) / DeflectionScale, "0.000") & _
                                            ", dY= " & Format((DYcoor(i) - Ycoor(i)) / DeflectionScale, "0.000"))) / 2
                frmMain.Picture1.CurrentY = YMirror(yi - 3 * CircleSize)
                frmMain.Picture1.Print CStr("dX= " & Format((DXcoor(i) - Xcoor(i)) / DeflectionScale, "0.000") & _
                                            ", dY= " & Format((DYcoor(i) - Ycoor(i)) / DeflectionScale, "0.000"))
                                            
                
            Next i
        '--------------------------
                
            'PLOT DEFLECTED LINES
            '-----------.
            For i = 1 To NoOfElements

            Color = Val(SecColor(Val(AsgnSec(i))))
            For j = 0 To Mesh - 1
                frmMain.Picture1.Line (DeflectionX(i, j), DeflectionY(i, j))- _
                (DeflectionX(i, j + 1), DeflectionY(i, j + 1)), Color

            Next j
            Next i
            frmMain.Picture1.DrawStyle = 0
End If
'---------------------------------------------------
If (PlotBMD = True) And (Analyzed = True) Then
    'PLOT BMD LINES
    '-----------.
    
    For i = 1 To NoOfElements
                Color = Val(SecColor(Val(AsgnSec(i))))
        X1 = Val(Xcoor(Endi(i)))
        X2 = Val(Xcoor(Endj(i)))
        Y1 = YMirror(Val(Ycoor(Endi(i))))
        Y2 = YMirror(Val(Ycoor(Endj(i))))
        
        frmMain.Picture1.Line (X1, Y1)- _
        (BMX(i, 0), BMY(i, 0)), Color
        
        frmMain.Picture1.CurrentX = BMX(i, 0)
        frmMain.Picture1.CurrentY = BMY(i, 0)
        frmMain.Picture1.Print CStr(Round(BendingMoment(i, 0), 3))
                
        For j = 0 To Mesh - 1
            frmMain.Picture1.Line (BMX(i, j), BMY(i, j))- _
            (BMX(i, j + 1), BMY(i, j + 1)), Color

        Next j
        
        frmMain.Picture1.CurrentX = BMX(i, Mesh)
        frmMain.Picture1.CurrentY = BMY(i, Mesh)
        frmMain.Picture1.Print CStr(Round(BendingMoment(i, Mesh), 3))
        
        frmMain.Picture1.Line (BMX(i, Mesh), BMY(i, Mesh))- _
        (X2, Y2), Color
    Next i
    
End If
'--------------------------------------------------
If (PlotSF = True) And (Analyzed = True) Then
    'PLOT SF LINES
    '-----------.

    
    For i = 1 To NoOfElements
        Color = Val(SecColor(Val(AsgnSec(i))))
        X1 = Val(Xcoor(Endi(i)))
        X2 = Val(Xcoor(Endj(i)))
        Y1 = YMirror(Val(Ycoor(Endi(i))))
        Y2 = YMirror(Val(Ycoor(Endj(i))))
        
        frmMain.Picture1.Line (X1, Y1)- _
        (SFX(i, 0), SFY(i, 0)), Color
        
                frmMain.Picture1.CurrentX = SFX(i, 0)
                frmMain.Picture1.CurrentY = SFY(i, 0)
                frmMain.Picture1.Print CStr(Round(ShearForce(i, 0), 3))
        For j = 0 To Mesh - 1
        

            
            frmMain.Picture1.Line (SFX(i, j), SFY(i, j))- _
            (SFX(i, j + 1), SFY(i, j + 1)), Color
            


        Next j
        
                frmMain.Picture1.CurrentX = SFX(i, Mesh)
                frmMain.Picture1.CurrentY = SFY(i, Mesh)
                frmMain.Picture1.Print CStr(Round(ShearForce(i, Mesh), 3))
                
        frmMain.Picture1.Line (SFX(i, Mesh), SFY(i, Mesh))- _
        (X2, Y2), Color
    Next i

    
End If
'---------------------------------------------------
If (PlotAxial = True) And (Analyzed = True) Then
    'PLOT Axial LINES
    '-----------.
    
    For i = 1 To NoOfElements
        Color = Val(SecColor(Val(AsgnSec(i))))
        X1 = Val(Xcoor(Endi(i)))
        X2 = Val(Xcoor(Endj(i)))
        Y1 = YMirror(Val(Ycoor(Endi(i))))
        Y2 = YMirror(Val(Ycoor(Endj(i))))
        
        frmMain.Picture1.Line (X1, Y1)- _
        (AxialX(i, 0), AxialY(i, 0)), Color
        
        frmMain.Picture1.CurrentX = AxialX(i, 0)
        frmMain.Picture1.CurrentY = AxialY(i, 0)
        frmMain.Picture1.Print CStr(Round(AxialForce(i, 0), 3))
                
        For j = 0 To Mesh - 1
            frmMain.Picture1.Line (AxialX(i, j), AxialY(i, j))- _
            (AxialX(i, j + 1), AxialY(i, j + 1)), Color

        Next j
        
        frmMain.Picture1.CurrentX = AxialX(i, Mesh)
        frmMain.Picture1.CurrentY = AxialY(i, Mesh)
        frmMain.Picture1.Print CStr(Round(AxialForce(i, Mesh), 3))
        
        frmMain.Picture1.Line (AxialX(i, Mesh), AxialY(i, Mesh))- _
        (X2, Y2), Color
    Next i
    
End If
'--------------------------------------------------
End With

Exit Sub
ErrorHandler:
Call subDispErrInfo("plotting the structure", Err.Number, Err.Description)

End Sub



Function YMirror(ByVal Y As Double) As Double
On Error GoTo ErrorHandler
    
    YMirror = frmMain.Picture1.ScaleTop + _
              (frmMain.Picture1.ScaleTop + frmMain.Picture1.ScaleHeight - Y)
Exit Function
ErrorHandler:
Call subDispErrInfo("mirroring the ordinate", Err.Number, Err.Description)
End Function


Sub subDisplayMatrix(Matrix() As Double, Name As String, DispFormat As String)
On Error GoTo ErrorHandler
    
    'READ DIMENSIONS OF MATRIX
        Dim Nrows, NCols As Long
        Nrows = UBound(Matrix, 1)
        NCols = UBound(Matrix, 2)
    'ASSEMBLE ENTRIES INTO ONE TEXT VARIABLE
        Dim i, j As Long
        Dim txt As String
        txt = ""
        For i = 1 To Nrows
            For j = 1 To NCols
                txt = txt & Format(Matrix(i, j), DispFormat)
                If Not j = NCols Then txt = txt & "   "
            Next j
            txt = txt & vbCrLf
        Next i
    'DISPLAY THE MATRIX
        MsgBox txt, , Name
Exit Sub
ErrorHandler:
Call subDispErrInfo("displaying matrix in message box", Err.Number, Err.Description)
End Sub


Sub subWriteMatrix(Matrix() As Double, WriteFormat As String, File As Integer)
On Error GoTo ErrorHandler
    'READ DIMENSIONS OF MATRIX
        Dim Nrows, NCols As Long
        Nrows = UBound(Matrix, 1)
        NCols = UBound(Matrix, 2)
        

    'WRITE MATRIX TO FILE
        Dim i, j As Long
        For i = 1 To Nrows
            For j = 1 To NCols
                Print #File, Tab(20 * (j - 1)); Format(Matrix(i, j), WriteFormat);
                'Print #File, Format(Matrix(i, j), WriteFormat),
            Next j
            Print #File,
        Next i
        Print #File,
Exit Sub
ErrorHandler:
Call subDispErrInfo("writing matrix to file", Err.Number, Err.Description)
End Sub


Sub PlotArrow(X As Double, Y As Double, ArrowLength As Double, Theta As Double, Color As Long, Optional Text As String)
On Error GoTo ErrorHandler
    If ArrowLength < 0.00001 Then Exit Sub
    Dim X1 As Double, X2 As Double, X3 As Double, X4 As Double, X5 As Double
    Dim Y1 As Double, Y2 As Double, Y3 As Double, Y4 As Double, Y5 As Double
    Dim temp As Integer
    temp = frmMain.Picture1.DrawStyle
    frmMain.Picture1.DrawStyle = 0
    
    X1 = X
    Y1 = Y
    
    X2 = X1 + (3 * CircleSize) * Cos(Theta + Atn(0.3334))
    Y2 = Y1 + (3 * CircleSize) * Sin(Theta + Atn(0.3334))
    
    X3 = X1 + (3 * CircleSize) * Cos(Theta - Atn(0.3334))
    Y3 = Y1 + (3 * CircleSize) * Sin(Theta - Atn(0.3334))
    
    X4 = X1 + (3 * CircleSize) * Cos(Theta)
    Y4 = Y1 + (3 * CircleSize) * Sin(Theta)
    
    X5 = X1 + (ArrowLength) * Cos(Theta)
    Y5 = Y1 + (ArrowLength) * Sin(Theta)
    
    frmMain.Picture1.Line (X1, YMirror(Y1))-(X2, YMirror(Y2)), Color
    frmMain.Picture1.Line (X1, YMirror(Y1))-(X3, YMirror(Y3)), Color
    frmMain.Picture1.Line (X2, YMirror(Y2))-(X3, YMirror(Y3)), Color
    frmMain.Picture1.DrawStyle = temp
    frmMain.Picture1.Line (X4, YMirror(Y4))-(X5, YMirror(Y5)), Color
    
    If Not IsMissing(Text) Then
        frmMain.Picture1.CurrentX = X5 + 0.5 * CircleSize
        frmMain.Picture1.CurrentY = YMirror(Y5)
        frmMain.Picture1.ForeColor = Color
        frmMain.Picture1.Print Text
    End If
Exit Sub
ErrorHandler:
Call subDispErrInfo("plotting load arrow", Err.Number, Err.Description)
    
End Sub

Sub PlotCircularArrow(X As Double, Y As Double, Direction As Integer, Color As Long, Optional Text As String)
On Error GoTo ErrorHandler

If Direction = 1 Then
    frmMain.Picture1.Circle (X, YMirror(Y)), 4 * CircleSize, Color, 0, 3.141592654 * 0.8
    Call PlotArrow(X - 4 * CircleSize, Y, 3 * CircleSize, 3.14159 / 2.7, Color)
End If
If Direction = -1 Then
    frmMain.Picture1.Circle (X, YMirror(Y)), 4 * CircleSize, Color, 3.141592654 * 0.2, 3.141592654
    Call PlotArrow(X + 4 * CircleSize, Y, 3 * CircleSize, 3.14159 - 3.14159 / 2.7, Color)
End If

    If Not IsMissing(Text) Then
        frmMain.Picture1.CurrentX = X
        frmMain.Picture1.CurrentY = YMirror(Y + 7 * CircleSize)
        frmMain.Picture1.ForeColor = Color
        frmMain.Picture1.Print Text
    End If
Exit Sub
ErrorHandler:
Call subDispErrInfo("plotting moment arrow", Err.Number, Err.Description)
End Sub



