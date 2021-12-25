Attribute VB_Name = "modSolutionProcedures"
Option Explicit



Public MaxDeflection As Double, MaxBM As Double, MaxSF As Double, MaxLoad As Double, MaxAxial As Double



Function GetElemLength(i As Long) As Double
On Error GoTo ErrorHandler
    GetElemLength = ((Xcoor(Endi(i)) - Xcoor(Endj(i))) ^ 2 + _
                (Ycoor(Endi(i)) - Ycoor(Endj(i))) ^ 2) ^ 0.5

Exit Function
ErrorHandler:
Call subDispErrInfo("calculating lenght of element " & i, Err.Number, Err.Description)
End Function

'option base 1
Sub subSolve()
On Error GoTo ErrorHandler

If NoOfElements < 1 Then Exit Sub
If Mesh < 1 Then Mesh = 1
    Call subMakeTrans
    Call subMakeElemLocalStiff
    Call subFormElemGlobalStiff
    Call subMakeRestVector
    Call subMakeSysStiff
    Call subMakeRedSysStiff
    Call subMakeDispVector
    Call subMakeForceVector

'REDUCTION OF FORCE VECTOR TO ONLY UNLOCKED DOFS
'-------------------------------------------------
Dim i As Long, TempF() As Double, Counter As Long
    ReDim TempF(DOFs - PDOFs, 1)
    Counter = 0
    For i = 1 To 3 * NoofNodes
        If RestVector(i, 1) = 0 Then
            Counter = Counter + 1
            TempF(Counter, 1) = ForceVector(i, 1)
        End If
    Next i
    
'SOLVE REDUCED SYSTEM OF EQUATIONS
'----------------------------------
Dim tempDis() As Double
    Call EqSolve(RedSysStiff(), TempF(), tempDis())
    'Call subDisplayMatrix(DispVector, "Displacements after solution", "0.00000000")
    
'FORM THE FULL DISPLACEMENT VECTOR CONTAINING LOCKED AND UNLOCKED DOFS.
'----------------------------------------------------------------------
    Counter = 0
    For i = 1 To 3 * NoofNodes
        If RestVector(i, 1) = 0 Then
            Counter = Counter + 1
            DispVector(i, 1) = tempDis(Counter, 1)
        End If
    Next i
    
    'Call subDisplayMatrix(DispVector, "Full Displacements after solution", "0.00000000")

    Call MatMult(SysStiff(), DispVector(), ForceVector())
    'Call subDisplayMatrix(ForceVector, "Forces after solution", "0.0")
    
'ADJUST FORCE VECTOR FOR ELEMENT END ACTIONS AND DIRECT FORCES
'-------------------------------------------------------------
    For i = 1 To 3 * NoofNodes
        ForceVector(i, 1) = ForceVector(i, 1) - EndActions(i, 1)
        ForceVector(i, 1) = ForceVector(i, 1) - DirectForce(i, 1)
    Next i
    
'CALCULATE INTERNAL FORCES
'-------------------------
    Call SubCalcInternalForces

'SET OUTPUT PLOT SCALES
'----------------------
    MaxDeflection = 0.000001: MaxBM = 0.0000000001
    MaxSF = 0.0000000001: MaxAxial = 0.0000000001
    
    Call subCalcMaxDef
    Call subCalcMaxBM
    Call subCalcMaxSF
    Call subCalcMaxAxial
    
    DeflectionScale = 0.65 * Round(Margin / MaxDeflection, 6)
    BMDScale = 0.65 * Round(Margin / MaxBM, 6)
    SFScale = 0.65 * Round(Margin / MaxSF, 6)
    AxialScale = 0.65 * Round(Margin / MaxAxial, 6)

'CALCULATE DEFORMED SHAPE FOR DEFAULT SCALE
'------------------------------------------
    Call subCalcDefCoor
    Call subCalcDefShapes

'SET MISC VARIABLES
'-----------------------
    Analyzed = True
    PlotDeflectedShape = True
    PlotBMD = False
    PlotSF = False
    PlotAxial = False
    
'REFRESH THE PLOT
'----------------
    Call subPlotStr
    
    DoEvents
    'MsgBox "Analysis Completed Successfully." & vbCrLf & "You can write analysis file to see outputs"
Exit Sub
ErrorHandler:
Call subDispErrInfo("analyzing the structure", Err.Number, Err.Description)
End Sub

Sub subMakeTrans()
On Error GoTo ErrorHandler

Dim Node1, Node2, i, l, m As Long
Dim X1, X2, Y1, Y2 As Double
Dim Length, Cx, Cy, CXsq, CYsq, CxCy As Double
    'REDIMENSION ELEMENT TRANSFORMATION ARRAY
    ReDim ElemTrans(NoOfElements, 6, 6)
    
    'FOR ALL ELEMENTS...
    For i = 1 To NoOfElements
        
        'GET THE END NODES AND THIER COORDINATES
        Node1 = Endi(i)
        Node2 = Endj(i)
        X1 = Xcoor(Node1)
        X2 = Xcoor(Node2)
        Y1 = Ycoor(Node1)
        Y2 = Ycoor(Node2)
        
        'COMPUTE LENTH AND DIRECTION COSINES
        Length = ((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2) ^ 0.5
        Cx = (X2 - X1) / Length
        Cy = (Y2 - Y1) / Length
        'CXsq = CX * CX
        'CYsq = CY * CY
        'CxCy = CX * CY
        
        'MAKE ENTRIES INTO ELEMENT TRANSFORMATION ARRY
            'FIRST ROW
                ElemTrans(i, 1, 1) = Cx
                ElemTrans(i, 1, 2) = Cy
                'REST ARE ZERO
            'SECOND ROW
                ElemTrans(i, 2, 1) = -Cy
                ElemTrans(i, 2, 2) = Cx
                'REST ARE ZERO
            'THIRD ROW
                ElemTrans(i, 3, 3) = 1
                'REST ARE ZERO
            'FOURTH ROW
                ElemTrans(i, 4, 4) = Cx
                ElemTrans(i, 4, 5) = Cy
                'REST ARE ZERO
            'FIFTH ROW
                ElemTrans(i, 5, 4) = -Cy
                ElemTrans(i, 5, 5) = Cx
                'REST ARE ZERO
            'SIXTH ROW
                ElemTrans(i, 6, 6) = 1
                'REST ARE ZEROS
        
        'JUST FOR CHECKING PURPOSES
        '    Dim temp() As Double
        '    temp() = fnGetTrans(i)
        '    'Call subDisplayMatrix(temp(), "Transformation for element " & i, "000.000")
    Next i
        
Exit Sub
ErrorHandler:
Call subDispErrInfo("forming transformation matrix for element " & i, Err.Number, Err.Description)
End Sub


Sub GetTrans(ByVal ElemNo As Long, Result() As Double)
On Error GoTo ErrorHandler

    ReDim Result(6, 6) As Double
    Dim i, j As Long
    For i = 1 To 6
        For j = 1 To 6
            Result(i, j) = ElemTrans(ElemNo, i, j)
        Next j
    Next i
    
Exit Sub
ErrorHandler:
Call subDispErrInfo("retrieving transformation matrix of element " & ElemNo, Err.Number, Err.Description)
End Sub

Sub GetElemLocalStiff(ByVal ElemNo As Long, Result() As Double)
On Error GoTo ErrorHandler

    ReDim Result(6, 6) As Double
    Dim i, j As Long
    For i = 1 To 6
        For j = 1 To 6
            Result(i, j) = ElemLocalStiff(ElemNo, i, j)
        Next j
    Next i
    
Exit Sub
ErrorHandler:
Call subDispErrInfo("retrieving local stiffness matrix of element " & ElemNo, Err.Number, Err.Description)
End Sub

Sub GetElemGlobalStiff(ByVal ElemNo As Long, Result() As Double)
On Error GoTo ErrorHandler

    ReDim Result(6, 6) As Double
    Dim i, j As Long
    For i = 1 To 6
        For j = 1 To 6
            Result(i, j) = ElemGlobalStiff(ElemNo, i, j)
        Next j
    Next i

Exit Sub
ErrorHandler:
Call subDispErrInfo("retrieving global stiffness matrix of element " & ElemNo, Err.Number, Err.Description)
End Sub

Sub subMakeElemLocalStiff()
On Error GoTo ErrorHandler

Dim i, l, m As Long
Dim X1, X2, Y1, Y2 As Double
'Dim CX, CY, CXsq, CYsq, CxCy As Double
    'REDIMENSION ELEMENT STIFFNESS MATRIX ARRAY
    ReDim ElemLocalStiff(NoOfElements, 6, 6)
    
    'FOR ALL ELEMENTS...
    For i = 1 To NoOfElements
        
        'GET THE ELEMENT MATERIAL PROPERTIES
        Dim A, E, Ix, Length As Double
        Dim C1, C2 As Double
        Dim ElemSec, ElemMaterial As Long
        
        ElemSec = AsgnSec(i)
        ElemMaterial = SecMat(ElemSec)
        
        A = SecArea(ElemSec)
        E = MatElasMod(ElemMaterial)
        Ix = SecIx(ElemSec)
            X1 = Xcoor(Endi(i))
            X2 = Xcoor(Endj(i))
            Y1 = Ycoor(Endi(i))
            Y2 = Ycoor(Endj(i))
        Length = ((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2) ^ 0.5
        
        C1 = A * E / Length
        C2 = E * Ix / (Length ^ 3)
        
        'MAKE ENTRIES INTO ELEMENT TRANSFORMATION ARRY
            'FIRST ROW
                ElemLocalStiff(i, 1, 1) = C1
                ElemLocalStiff(i, 1, 4) = -C1
                'REST ARE ZERO
            'SECOND ROW
                ElemLocalStiff(i, 2, 2) = 12 * C2
                ElemLocalStiff(i, 2, 3) = 6 * C2 * Length
                ElemLocalStiff(i, 2, 5) = -12 * C2
                ElemLocalStiff(i, 2, 6) = 6 * C2 * Length
                'REST ARE ZERO
            'THIRD ROW
                ElemLocalStiff(i, 3, 2) = 6 * C2 * Length
                ElemLocalStiff(i, 3, 3) = 4 * C2 * Length * Length
                ElemLocalStiff(i, 3, 5) = -6 * C2 * Length
                ElemLocalStiff(i, 3, 6) = 2 * C2 * Length * Length
                'REST ARE ZERO
            'FOURTH ROW
                ElemLocalStiff(i, 4, 1) = -C1
                ElemLocalStiff(i, 4, 4) = C1
                'REST ARE ZERO
            'FIFTH ROW
                ElemLocalStiff(i, 5, 2) = -12 * C2
                ElemLocalStiff(i, 5, 3) = -6 * C2 * Length
                ElemLocalStiff(i, 5, 5) = 12 * C2
                ElemLocalStiff(i, 5, 6) = -6 * C2 * Length
                'REST ARE ZERO
            'SIXTH ROW
                ElemLocalStiff(i, 6, 2) = 6 * C2 * Length
                ElemLocalStiff(i, 6, 3) = 2 * C2 * Length * Length
                ElemLocalStiff(i, 6, 5) = -6 * C2 * Length
                ElemLocalStiff(i, 6, 6) = 4 * C2 * Length * Length
                'REST ARE ZERO
        
        'JUST FOR CHECKING PURPOSES
        '    Dim temp() As Double
        '    temp() = fnGetElemLocalStiff(i)
        '    'Call subDisplayMatrix(temp(), "Local K for element " & i, "000.000")
    Next i
        
Exit Sub
ErrorHandler:
Call subDispErrInfo("forming local stiffness matrix of element " & i, Err.Number, Err.Description)
End Sub

Sub subFormElemGlobalStiff()
On Error GoTo ErrorHandler

ReDim ElemGlobalStiff(NoOfElements, 6, 6)
Dim i, m, n, p As Long
    Dim tempTrans() As Double
    Dim tempK() As Double
    Dim tempTransTr() As Double
    Dim KT() As Double
    Dim TtKT() As Double
    
    For i = 1 To NoOfElements
        
        Call GetTrans(i, tempTrans())
            ''Call subDisplayMatrix(tempTrans(), "T for element " & i, "000.000")
        
        Call GetElemLocalStiff(i, tempK())
            ''Call subDisplayMatrix(tempK(), "Local K for element " & i, "000.000")
        
        Call MatTranspose(tempTrans(), tempTransTr())
            ''Call subDisplayMatrix(tempTransTr(), "Ttr for element " & i, "000.000")
        
        Call MatMult(tempK(), tempTrans(), KT())
            ''Call subDisplayMatrix(KT(), "K x T for element " & i, "000.000")
        
        Call MatMult(tempTransTr(), KT(), TtKT())
            ''Call subDisplayMatrix(TtKT(), "Global K for element " & i, "000.000")
 
        For m = 1 To 6
            For n = 1 To 6
                ElemGlobalStiff(i, m, n) = TtKT(m, n)
            Next n
        Next m
        
    Next i
    
Exit Sub
ErrorHandler:
Call subDispErrInfo("forming global stiffness matrix of element " & i, Err.Number, Err.Description)
End Sub

Sub subMakeSysStiff()
On Error GoTo ErrorHandler


Dim Node1, Node2 As Long
Dim i, m, n As Long
Dim temp() As Double

ReDim SysStiff(3 * NoofNodes, 3 * NoofNodes)

For i = 1 To NoOfElements
    Node1 = Endi(i)
    Node2 = Endj(i)
    
    Call GetElemGlobalStiff(i, temp())
    
    'FIRST QUADRANT
    '---------------
    For m = (3 * Node1 - 3 + 1) To (3 * Node1 - 3 + 3)
        For n = (3 * Node1 - 3 + 1) To (3 * Node1 - 3 + 3)
            SysStiff(m, n) = SysStiff(m, n) + temp(m - (3 * Node1 - 3), n - (3 * Node1 - 3))
        Next n
    Next m
    
    'SECOND QUADRANT
    '---------------
    For m = (3 * Node1 - 3 + 1) To (3 * Node1 - 3 + 3)
        For n = (3 * Node2 - 3 + 1) To (3 * Node2 - 3 + 3)
            SysStiff(m, n) = SysStiff(m, n) + temp(m - (3 * Node1 - 3), n - (3 * Node2 - 3) + 3)
        Next n
    Next m
    
    'THIRD QUADRANT
    '---------------
    For m = (3 * Node2 - 3 + 1) To (3 * Node2 - 3 + 3)
        For n = (3 * Node1 - 3 + 1) To (3 * Node1 - 3 + 3)
            SysStiff(m, n) = SysStiff(m, n) + temp(m - (3 * Node2 - 3) + 3, n - (3 * Node1 - 3))
        Next n
    Next m
    
    'FOURTH QUADRANT
    '---------------
    For m = (3 * Node2 - 3 + 1) To (3 * Node2 - 3 + 3)
        For n = (3 * Node2 - 3 + 1) To (3 * Node2 - 3 + 3)
            SysStiff(m, n) = SysStiff(m, n) + temp(m - (3 * Node2 - 3) + 3, n - (3 * Node2 - 3) + 3)
        Next n
    Next m
    
Next i

''Call subDisplayMatrix(SysStiff(), "System Stiffness matrix", "0")

Exit Sub
ErrorHandler:
Call subDispErrInfo("creating system global stiffness method", Err.Number, Err.Description)
End Sub
Sub subMakeRestVector()
On Error GoTo ErrorHandler

Dim i As Long
ReDim RestVector(3 * NoofNodes, 1)
PDOFs = 0
For i = 1 To NoofNodes
    If TxRest(i) = 1 Then
        RestVector(3 * (i - 1) + 1, 1) = 1
        PDOFs = PDOFs + 1
    End If
    If TyRest(i) = 1 Then
        RestVector(3 * (i - 1) + 2, 1) = 1
        PDOFs = PDOFs + 1
    End If
    
    If RzRest(i) = 1 Then
        RestVector(3 * (i - 1) + 3, 1) = 1
        PDOFs = PDOFs + 1
    End If
Next i
''Call subDisplayMatrix(RestVector(), "Restrained Vector", "0")

Exit Sub
ErrorHandler:
Call subDispErrInfo("creating restrains vector", Err.Number, Err.Description)
End Sub

Sub subMakeRedSysStiff()
On Error GoTo ErrorHandler

Dim i, j As Long
Dim NodeNo As Long

DOFs = 3 * NoofNodes
ReDim RedSysStiff(DOFs - PDOFs, DOFs - PDOFs)
Dim RowCounter As Long, ColCounter As Long
For i = 1 To DOFs
    If RestVector(i, 1) = 0 Then
        RowCounter = RowCounter + 1
        ColCounter = 0
        For j = 1 To DOFs
            If RestVector(j, 1) = 0 Then
                ColCounter = ColCounter + 1
                RedSysStiff(RowCounter, ColCounter) = SysStiff(i, j)
            End If
        Next j
    End If
Next i


''Call subDisplayMatrix(RedSysStiff(), "Reduced System Stiffness matrix", "0")

Exit Sub
ErrorHandler:
Call subDispErrInfo("reducing system stiffness matrix", Err.Number, Err.Description)
End Sub

Sub subMakeForceVector()
On Error GoTo ErrorHandler

ReDim ForceVector(3 * NoofNodes, 1)
ReDim EndActions(3 * NoofNodes, 1)

Dim i As Long
Dim TempLF(6, 1) As Double, tempGF() As Double
Dim tempTrans() As Double
Dim tempTransT() As Double



'ADD NODAL FORCES
'-----------------
For i = 1 To NoofNodes
    ForceVector(3 * (i - 1) + 1, 1) = Val(XForce(i))
    ForceVector(3 * (i - 1) + 2, 1) = Val(YForce(i))
    ForceVector(3 * (i - 1) + 3, 1) = Val(ZMom(i))
Next i


'ADD ELEMENT END ACTIONS
'------------------------
Dim A1 As Double, A2 As Double, V1 As Double, V2 As Double, M1 As Double, M2 As Double, Total As Double
Dim UDL1 As Double, UDL2 As Double, UDAL1 As Double, UDAL2 As Double
Dim Node1 As Long, Node2 As Long
Dim Length As Double


For i = 1 To NoOfElements
    
    Node1 = Endi(i)
    Node2 = Endj(i)

    Length = GetElemLength(i)
    
    UDL1 = ElemLoadi(i)
    UDL2 = ElemLoadj(i)
    UDAL1 = ElemALoadi(i)
    UDAL2 = ElemALoadj(i)
    
    Total = (UDL1 + UDL2) / 2 * Length
    
    A1 = UDAL2 * Length / 6 + UDAL1 * Length / 3
    A2 = UDAL1 * Length / 6 + UDAL2 * Length / 3
    V1 = UDL2 * Length * 0.15 + UDL1 * Length * 0.35
    V2 = UDL1 * Length * 0.15 + UDL2 * Length * 0.35
    M1 = UDL1 * Length ^ 2 / 20 + UDL2 * Length ^ 2 / 30
    M2 = -(UDL2 * Length ^ 2 / 20 + UDL1 * Length ^ 2 / 30)
    
    TempLF(1, 1) = A1
    TempLF(4, 1) = A2
    TempLF(2, 1) = V1
    TempLF(5, 1) = V2
    TempLF(3, 1) = M1
    TempLF(6, 1) = M2
    
    'Call subDisplayMatrix(TempLF(), "Local end Forces on element " & i, "000.000")
    
    Call GetTrans(i, tempTrans())
    Call MatTranspose(tempTrans(), tempTransT())
    Call MatMult(tempTransT(), TempLF(), tempGF())
    
    'Call subDisplayMatrix(tempGF(), "Global end Forces on element " & i, "000.000")
    
    ForceVector(3 * (Node1 - 1) + 1, 1) = ForceVector(3 * (Node1 - 1) + 1, 1) + tempGF(1, 1)
    ForceVector(3 * (Node1 - 1) + 2, 1) = ForceVector(3 * (Node1 - 1) + 2, 1) + tempGF(2, 1)
    ForceVector(3 * (Node1 - 1) + 3, 1) = ForceVector(3 * (Node1 - 1) + 3, 1) + tempGF(3, 1)
    ForceVector(3 * (Node2 - 1) + 1, 1) = ForceVector(3 * (Node2 - 1) + 1, 1) + tempGF(4, 1)
    ForceVector(3 * (Node2 - 1) + 2, 1) = ForceVector(3 * (Node2 - 1) + 2, 1) + tempGF(5, 1)
    ForceVector(3 * (Node2 - 1) + 3, 1) = ForceVector(3 * (Node2 - 1) + 3, 1) + tempGF(6, 1)
    
    EndActions(3 * (Node1 - 1) + 1, 1) = EndActions(3 * (Node1 - 1) + 1, 1) + tempGF(1, 1)
    EndActions(3 * (Node1 - 1) + 2, 1) = EndActions(3 * (Node1 - 1) + 2, 1) + tempGF(2, 1)
    EndActions(3 * (Node1 - 1) + 3, 1) = EndActions(3 * (Node1 - 1) + 3, 1) + tempGF(3, 1)
    EndActions(3 * (Node2 - 1) + 1, 1) = EndActions(3 * (Node2 - 1) + 1, 1) + tempGF(4, 1)
    EndActions(3 * (Node2 - 1) + 2, 1) = EndActions(3 * (Node2 - 1) + 2, 1) + tempGF(5, 1)
    EndActions(3 * (Node2 - 1) + 3, 1) = EndActions(3 * (Node2 - 1) + 3, 1) + tempGF(6, 1)
    
Next i

'SAVE LOADS DIRECTLY APPLIED TO LOCKED DEGREE OF FREEDOM IN DIRECT FORCES
'-------------------------------------------------------------------------
ReDim DirectForce(3 * NoofNodes, 1)
    For i = 1 To 3 * NoofNodes
        If RestVector(i, 1) = 1 Then DirectForce(i, 1) = _
                    DirectForce(i, 1) + ForceVector(i, 1) - EndActions(i, 1)
    Next i
'---------------------------------------------------------------------

'Call subDisplayMatrix(ForceVector, "Force Vector before solving", "0.0")

Exit Sub
ErrorHandler:
Call subDispErrInfo("creating force vector", Err.Number, Err.Description)
End Sub


Sub subMakeDispVector()
On Error GoTo ErrorHandler

ReDim DispVector(3 * NoofNodes, 1)

Dim i As Long
For i = 1 To NoofNodes
    DispVector(3 * i - 3 + 1, 1) = 0 'For the time being no displacements are applied initially
    DispVector(3 * i - 3 + 2, 1) = 0
    DispVector(3 * i - 3 + 3, 1) = 0
Next i

'Call subDisplayMatrix(DispVector, "Displacement Vector before solving", "0.0")

Exit Sub
ErrorHandler:
Call subDispErrInfo("creating displacement vector", Err.Number, Err.Description)
End Sub

Sub GetDisp(ElemNo As Long, Result() As Double)
On Error GoTo ErrorHandler

Dim Node1 As Long, Node2 As Long
    Node1 = Endi(ElemNo)
    Node2 = Endj(ElemNo)

ReDim Result(6, 1)

    Result(1, 1) = DispVector(3 * (Node1 - 1) + 1, 1)
    Result(2, 1) = DispVector(3 * (Node1 - 1) + 2, 1)
    Result(3, 1) = DispVector(3 * (Node1 - 1) + 3, 1)
    
    Result(4, 1) = DispVector(3 * (Node2 - 1) + 1, 1)
    Result(5, 1) = DispVector(3 * (Node2 - 1) + 2, 1)
    Result(6, 1) = DispVector(3 * (Node2 - 1) + 3, 1)
Exit Sub
ErrorHandler:
Call subDispErrInfo("retrieving global displacements of element " & ElemNo, Err.Number, Err.Description)
End Sub
Sub GetForces(ElemNo As Long, Result() As Double)
On Error GoTo ErrorHandler

Dim Node1 As Long, Node2 As Long
    Node1 = Endi(ElemNo)
    Node2 = Endj(ElemNo)

ReDim Result(6, 1)

    Result(1, 1) = ForceVector(3 * (Node1 - 1) + 1, 1)
    Result(2, 1) = ForceVector(3 * (Node1 - 1) + 2, 1)
    Result(3, 1) = ForceVector(3 * (Node1 - 1) + 3, 1)
    
    Result(4, 1) = ForceVector(3 * (Node2 - 1) + 1, 1)
    Result(5, 1) = ForceVector(3 * (Node2 - 1) + 2, 1)
    Result(6, 1) = ForceVector(3 * (Node2 - 1) + 3, 1)

Exit Sub
ErrorHandler:
Call subDispErrInfo("retrieving global end forces of element " & ElemNo, Err.Number, Err.Description)
End Sub


Sub SubCalcInternalForces()
On Error GoTo ErrorHandler

ReDim ShearForce(NoOfElements, 0 To Mesh)
ReDim BendingMoment(NoOfElements, 0 To Mesh)
ReDim AxialForce(NoOfElements, 0 To Mesh)
ReDim Slope(NoOfElements, 0 To Mesh)
ReDim Deflection(NoOfElements, 0 To Mesh)
ReDim INIDIS1(NoOfElements)
ReDim INIDIS2(NoOfElements)

Dim i As Long, j As Long
Dim tempGF() As Double, tempGDis() As Double, tempGStiff() As Double
Dim TempLF() As Double, tempLDis() As Double, tempLStiff() As Double
Dim tempTran() As Double
Dim UDL1 As Double, UDL2 As Double, UDAL1 As Double, UDAL2 As Double
Dim Length As Double
Dim A1 As Double, A2 As Double
Dim txt As String
Dim V1 As Double, V2 As Double, M1 As Double, M2 As Double, Total As Double
Dim E As Double, Ix As Double, EI As Double
Dim Cx As Double, Cy As Double


For i = 1 To NoOfElements


    Length = GetElemLength(i)
    Call GetDisp(i, tempGDis())
    Call GetTrans(i, tempTran())
    Cx = tempTran(1, 1)
    Cy = tempTran(2, 1)
    Call MatMult(tempTran(), tempGDis(), tempLDis())
    Call GetElemLocalStiff(i, tempLStiff())
    Call MatMult(tempLStiff(), tempLDis(), TempLF())
    
    
    'ADJUST FORCES FOR ELEMENT END ACTIONS
'------------------------
    UDL1 = ElemLoadi(i)
    UDL2 = ElemLoadj(i)
    UDAL1 = ElemALoadi(i)
    UDAL2 = ElemALoadj(i)
    
    Total = (UDL1 + UDL2) / 2 * Length
    
    A1 = UDAL2 * Length / 6 + UDAL1 * Length / 3
    A2 = UDAL1 * Length / 6 + UDAL2 * Length / 3
    V1 = UDL2 * Length * 0.15 + UDL1 * Length * 0.35
    V2 = UDL1 * Length * 0.15 + UDL2 * Length * 0.35
    M1 = UDL1 * Length ^ 2 / 20 + UDL2 * Length ^ 2 / 30
    M2 = -(UDL2 * Length ^ 2 / 20 + UDL1 * Length ^ 2 / 30)
    
    TempLF(1, 1) = TempLF(1, 1) - A1
    TempLF(4, 1) = TempLF(4, 1) - A2
    TempLF(2, 1) = TempLF(2, 1) - V1
    TempLF(5, 1) = TempLF(5, 1) - V2
    TempLF(3, 1) = TempLF(3, 1) - M1
    TempLF(6, 1) = TempLF(6, 1) - M2
    
    'Call subDisplayMatrix(TempLF(), "Internal Force of element " & i, "000.0000")

    
    
    'SHEAR FORCE CALCULATIONS
    '--------------------------
    ShearForce(i, 0) = TempLF(2, 1) '+ (tempLF(3, 1) + tempLF(6, 1)) / Length '- forcevector(3 * Endi(i) - 3 + 2, 1)
    txt = ShearForce(i, 0)
    ShearForce(i, Mesh) = -TempLF(5, 1) ' - (tempLF(3, 1) + tempLF(6, 1)) / Length '+ forcevector(3 * Endj(i) - 3 + 2, 1)
     
    For j = 1 To Mesh - 1
        A1 = (UDL1 / Mesh * (Mesh - (j - 1)) + UDL1 / Mesh * (Mesh - j)) / 2 * Length / Mesh
        A2 = (UDL2 / Mesh * (j - 1) + UDL2 / Mesh * j) / 2 * Length / Mesh
        
        ShearForce(i, j) = ShearForce(i, j - 1) + A1 + A2
        txt = txt & vbCrLf & ShearForce(i, j)
    Next j
    txt = txt & vbCrLf & ShearForce(i, Mesh)
    'MsgBox txt, , "shear force for element " & i
    '--------------------------
    
    'BENDING MOMENT CALCULATIONS
    '--------------------------
    BendingMoment(i, 0) = -TempLF(3, 1)
        txt = BendingMoment(i, 0)
    BendingMoment(i, Mesh) = TempLF(6, 1)
    
    For j = 1 To Mesh - 1
        A1 = ShearForce(i, j - 1) * Length / Mesh / 2
        A2 = ShearForce(i, j) * Length / Mesh / 2
        
        BendingMoment(i, j) = BendingMoment(i, j - 1) + A1 + A2
        txt = txt & vbCrLf & BendingMoment(i, j)
    Next j
    txt = txt & vbCrLf & BendingMoment(i, Mesh)
    
    'MsgBox txt, , "Bending Moment for element " & i
    '--------------------------
    
    'SLOPE CALCULATIONS
    '--------------------------
    Slope(i, 0) = tempLDis(3, 1)
    txt = Slope(i, 0)
    Slope(i, Mesh) = tempLDis(6, 1)
    E = MatElasMod(SecMat(AsgnSec(i)))
    Ix = SecIx(AsgnSec(i))
    EI = E * Ix
    
    For j = 1 To Mesh - 1
        A1 = (BendingMoment(i, j - 1) + BendingMoment(i, j)) / 2 * (Length / Mesh)
        
        Slope(i, j) = Slope(i, j - 1) + A1 / EI
        txt = txt & vbCrLf & Slope(i, j)
    Next j
    txt = txt & vbCrLf & Slope(i, Mesh)
    'MsgBox txt, , "Slopes for element " & i
    '--------------------------
    
    'DEFLECTIONS CALCULATIONS
    '--------------------------
    INIDIS1(i) = tempLDis(2, 1)
    INIDIS2(i) = tempLDis(5, 1)
    
    Deflection(i, 0) = tempLDis(2, 1)
        'DeflectionX(i, 0) = -Cy * Deflection(i, 0)
        'DeflectionY(i, 0) = Cx * Deflection(i, 0)
    txt = Deflection(i, 0)
    Deflection(i, Mesh) = tempLDis(5, 1)
        'DeflectionX(i, mesh) = -Cy * Deflection(i, mesh)
        'DeflectionY(i, mesh) = Cx * Deflection(i, mesh)
    
    For j = 1 To Mesh - 1
        'A1 = Slope(i, j - 1) * Length / mesh / 2
        'A2 = Slope(i, j) * Length / mesh / 2
        A1 = 0.5 * (Slope(i, j - 1) + Slope(i, j)) * Length / Mesh
        Deflection(i, j) = Deflection(i, j - 1) + A1 '+ A2
            'DeflectionX(i, j) = -Cy * Deflection(i, j)
            'DeflectionY(i, j) = Cx * Deflection(i, j)
        txt = txt & vbCrLf & Deflection(i, j)
    Next j
       
    txt = txt & vbCrLf & Deflection(i, Mesh)
    
    'MsgBox txt, , "Deflections for element " & i
    '--------------------------
    
    'AXIAL FORCES CALCULATIONS
    '--------------------------
    AxialForce(i, 0) = TempLF(1, 1) '+ (tempLF(3, 1) + tempLF(6, 1)) / Length '- forcevector(3 * Endi(i) - 3 + 2, 1)
    txt = AxialForce(i, 0)
    AxialForce(i, Mesh) = -TempLF(4, 1) ' - (tempLF(3, 1) + tempLF(6, 1)) / Length '+ forcevector(3 * Endj(i) - 3 + 2, 1)
     
    For j = 1 To Mesh - 1
        A1 = (UDAL1 / Mesh * (Mesh - (j - 1)) + UDAL1 / Mesh * (Mesh - j)) / 2 * Length / Mesh
        A2 = (UDAL2 / Mesh * (j - 1) + UDAL2 / Mesh * j) / 2 * Length / Mesh
        
        AxialForce(i, j) = AxialForce(i, j - 1) + A1 + A2
        txt = txt & vbCrLf & AxialForce(i, j)
    Next j
    txt = txt & vbCrLf & AxialForce(i, Mesh)
    'MsgBox txt, , "Axial force for element " & i
    '--------------------------
    
    
    
    
Next i

Exit Sub
ErrorHandler:
Call subDispErrInfo("calculating internal forces", Err.Number, Err.Description)

End Sub
Sub subCalcMaxDef()
On Error GoTo ErrorHandler

'ReDim DeflectionX(NoOfElements, 0 To mesh)
'ReDim DeflectionY(NoOfElements, 0 To mesh)

'Dim temp As Double
Dim i As Long, j As Long
'Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Length As Double
'Dim Cx, Cy
For i = 1 To NoOfElements
    'X1 = Xcoor(Endi(i)): Y1 = Ycoor(Endi(i))
    'X2 = Xcoor(Endj(i)): Y2 = Ycoor(Endj(i))
    'Length = GetElemLength(i)
    'Cx = (X2 - X1) / Length: Cy = (Y2 - Y1) / Length
    For j = 0 To Mesh
        'DeflectionX(i, j) = (-Cy * 1 * Deflection(i, j))
        'DeflectionY(i, j) = (Cx * 1 * Deflection(i, j))
        'temp = (DeflectionX(i, j) ^ 2 + DeflectionY(i, j) ^ 2) ^ 0.5
        If MaxDeflection < Abs(Deflection(i, j)) Then MaxDeflection = Abs(Deflection(i, j))
    Next j
Next i
Exit Sub
ErrorHandler:
Call subDispErrInfo("calculating maximum deflection", Err.Number, Err.Description)
End Sub

Sub subCalcDefCoor()
On Error GoTo ErrorHandler

Dim i As Long

ReDim DXcoor(NoofNodes)
ReDim DYcoor(NoofNodes)

For i = 1 To NoofNodes
    DXcoor(i) = Xcoor(i) + DeflectionScale * DispVector(3 * (i - 1) + 1, 1)
    DYcoor(i) = Ycoor(i) + DeflectionScale * DispVector(3 * (i - 1) + 2, 1)
Next i
Exit Sub
ErrorHandler:
Call subDispErrInfo("calculating deformed coordinates", Err.Number, Err.Description)
End Sub


Sub subCalcDefShapes()
On Error GoTo ErrorHandler

Dim i As Long, j As Long
Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Length As Double
Dim XX1 As Double, XX2 As Double, YY1 As Double, YY2 As Double
Dim Cx, Cy
ReDim DeflectionX(NoOfElements, 0 To Mesh)
ReDim DeflectionY(NoOfElements, 0 To Mesh)

For i = 1 To NoOfElements
    X1 = DXcoor(Endi(i)): Y1 = DYcoor(Endi(i))
    X2 = DXcoor(Endj(i)): Y2 = DYcoor(Endj(i))
    XX1 = Xcoor(Endi(i)): YY1 = Ycoor(Endi(i))
    XX2 = Xcoor(Endj(i)): YY2 = Ycoor(Endj(i))
    Length = GetElemLength(i)
    Cx = (XX2 - XX1) / Length: Cy = (YY2 - YY1) / Length
    For j = 0 To Mesh
        DeflectionX(i, j) = (X1 + (X2 - X1) * j / Mesh) + (-Cy * DeflectionScale * (Deflection(i, j) - (INIDIS1(i) + (INIDIS2(i) - INIDIS1(i)) * j / Mesh)))
        DeflectionY(i, j) = YMirror((Y1 + (Y2 - Y1) * j / Mesh) + (Cx * DeflectionScale * (Deflection(i, j) - (INIDIS1(i) + (INIDIS2(i) - INIDIS1(i)) * j / Mesh))))
    Next j
Next i
Exit Sub
ErrorHandler:
Call subDispErrInfo("deflected shapes", Err.Number, Err.Description)
End Sub

Sub subCalcMaxBM()
On Error GoTo ErrorHandler

Dim i As Long, j As Long
For i = 1 To NoOfElements
    For j = 0 To Mesh
        If MaxBM < Abs(BendingMoment(i, j)) Then MaxBM = Abs(BendingMoment(i, j))
    Next j
Next i
Exit Sub
ErrorHandler:
Call subDispErrInfo("calculating maximum bending moment", Err.Number, Err.Description)
End Sub

Sub subCalcBMDShapes()
On Error GoTo ErrorHandler

ReDim BMX(NoOfElements, 0 To Mesh)
ReDim BMY(NoOfElements, 0 To Mesh)
Dim i As Long, j As Long
Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Length As Double
Dim Cx, Cy


For i = 1 To NoOfElements
    X1 = Xcoor(Endi(i)): Y1 = Ycoor(Endi(i))
    X2 = Xcoor(Endj(i)): Y2 = Ycoor(Endj(i))
    Length = GetElemLength(i)
    Cx = (X2 - X1) / Length: Cy = (Y2 - Y1) / Length
    For j = 0 To Mesh
        BMX(i, j) = (X1 + (X2 - X1) * j / Mesh) + (-Cy * BMDScale * BendingMoment(i, j))
        BMY(i, j) = YMirror((Y1 + (Y2 - Y1) * j / Mesh) + (Cx * BMDScale * BendingMoment(i, j)))
    Next j
Next i
Exit Sub
ErrorHandler:
Call subDispErrInfo("calculating bending moment shapes", Err.Number, Err.Description)
End Sub

Sub subCalcMaxSF()
On Error GoTo ErrorHandler

Dim i As Long, j As Long
For i = 1 To NoOfElements
    For j = 0 To Mesh
        If MaxSF < Abs(ShearForce(i, j)) Then MaxSF = Abs(ShearForce(i, j))
    Next j
Next i
Exit Sub
ErrorHandler:
Call subDispErrInfo("calculating maximum shear force", Err.Number, Err.Description)
End Sub

Sub subCalcSFShapes()
On Error GoTo ErrorHandler

ReDim SFX(NoOfElements, 0 To Mesh)
ReDim SFY(NoOfElements, 0 To Mesh)
Dim i As Long, j As Long
Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Length As Double
Dim Cx, Cy


For i = 1 To NoOfElements
    X1 = Xcoor(Endi(i)): Y1 = Ycoor(Endi(i))
    X2 = Xcoor(Endj(i)): Y2 = Ycoor(Endj(i))
    Length = GetElemLength(i)
    Cx = (X2 - X1) / Length: Cy = (Y2 - Y1) / Length
    For j = 0 To Mesh
        SFX(i, j) = (X1 + (X2 - X1) * j / Mesh) + (-Cy * SFScale * Round(ShearForce(i, j), 10))
        SFY(i, j) = YMirror((Y1 + (Y2 - Y1) * j / Mesh) + (Cx * SFScale * Round(ShearForce(i, j), 10)))
    Next j
Next i
Exit Sub
ErrorHandler:
Call subDispErrInfo("calculating shear force shapes", Err.Number, Err.Description)
End Sub

Sub subCalcMaxLoad()
On Error GoTo ErrorHandler

Dim i As Long, j As Long
For i = 1 To NoOfElements

    If MaxLoad < Abs(ElemLoadi(i)) Then MaxLoad = Abs(ElemLoadi(i))
    If MaxLoad < Abs(ElemLoadj(i)) Then MaxLoad = Abs(ElemLoadj(i))
    If MaxLoad < Abs(ElemALoadi(i)) Then MaxLoad = Abs(ElemALoadi(i))
    If MaxLoad < Abs(ElemALoadj(i)) Then MaxLoad = Abs(ElemALoadj(i))
Next i
Exit Sub
ErrorHandler:
Call subDispErrInfo("calculating maximum distributed load", Err.Number, Err.Description)
End Sub

Sub subCalcMaxAxial()
On Error GoTo ErrorHandler

Dim i As Long, j As Long
For i = 1 To NoOfElements
    For j = 0 To Mesh
        If MaxAxial < Abs(AxialForce(i, j)) Then MaxAxial = Abs(AxialForce(i, j))
    Next j
Next i
Exit Sub
ErrorHandler:
Call subDispErrInfo("calculating maximum axial force", Err.Number, Err.Description)
End Sub

Sub subCalcAxialShapes()
On Error GoTo ErrorHandler

ReDim AxialX(NoOfElements, 0 To Mesh)
ReDim AxialY(NoOfElements, 0 To Mesh)
Dim i As Long, j As Long
Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Length As Double
Dim Cx, Cy

For i = 1 To NoOfElements
    X1 = Xcoor(Endi(i)): Y1 = Ycoor(Endi(i))
    X2 = Xcoor(Endj(i)): Y2 = Ycoor(Endj(i))
    Length = GetElemLength(i)
    Cx = (X2 - X1) / Length: Cy = (Y2 - Y1) / Length
    For j = 0 To Mesh
        AxialX(i, j) = (X1 + (X2 - X1) * j / Mesh) + (-Cy * AxialScale * AxialForce(i, j))
        AxialY(i, j) = YMirror((Y1 + (Y2 - Y1) * j / Mesh) + (Cx * AxialScale * AxialForce(i, j)))
    Next j
Next i
Exit Sub
ErrorHandler:
Call subDispErrInfo("calculating axial force shapes", Err.Number, Err.Description)
End Sub
