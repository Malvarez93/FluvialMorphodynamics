Attribute VB_Name = "Main"
Sub Exner()
    Application.ScreenUpdating = False
    Dim Path As String
    Reset_US
    Reset_DS
    Reset_Solution
    Call Clear_Sheets
    Sheets("Grid-BC").Select
    
    Dim t As Double
    Dim Iter As Integer
    Dim i As Integer
    Dim Q As Double
    Dim us As Double
    Dim DS As Double
    
    t = 0
    Iter = 1
    i = 0
    
    'Calculates Initial Bedforms and updates the Manning roughness to consider Bedforms
    If Sheets("Grid-BC").Range("T7").Value = 1 Then
            Sheets("Solution").Select
            Calculate_Bedforms
            Call Update_Manning
    End If
    
    'Iteration scheme (t < tmax)
    Do While t < Sheets("Grid-BC").Cells(7, 15).Value
        If t > Sheets("Grid-BC").Cells(11 + i + 1, 14).Value Then
            i = i + 1
            t = Sheets("Grid-BC").Cells(11 + i, 14).Value
        End If
        
        'Reads the boundary conditions
        Q = Sheets("Grid-BC").Cells(11 + i, 15).Value
        us = Sheets("Grid-BC").Cells(11 + i, 16).Value
        DS = Sheets("Grid-BC").Cells(11 + i, 17).Value
        
        'Adds Sea Level Rise to the DS Boundary Condition
        If Sheets("Grid-BC").Range("T9").Value = 1 Then
            us = us + t * Sheets("Grid-BC").Range("O4").Value
        End If
        
        'Solve US
        Sheets("US").Select
        Cells(8, 26).Value = us
        Cells(9, 26).Value = Q
        Application.Calculate
        Solve_US
        
        'Solve DS
        Sheets("DS").Select
        Cells(8, 26).Value = DS
        Cells(9, 26).Value = Q
        Application.Calculate
        Solve_DS
        
        'Save Solution
        Sheets("Solution").Select
        Application.Calculate
        Call Paste_Solution(Iter, t)
        Sheets("Solution").Select
        Calculate_Bedforms
        CalculateZNew
              
        'Export Graphs
        If Worksheets("Grid-BC").Cells(3, 20).Value = 1 Then
            Worksheets("Solution").Select
            ActiveSheet.ChartObjects("Chart 1").Activate
            Path = Worksheets("Grid-BC").Cells(4, 20).Value & "\" & Worksheets("Grid-BC").Cells(5, 20).Value & "\" & Format(t, "0.000") & ".png"
            ActiveChart.Export Path
        End If
        
        'Update Bed and Add Time
        Call Update_Bed
        
        If Sheets("Grid-BC").Range("T7").Value = 1 Then
            Call Update_Manning
        End If
        
        Iter = Iter + 1
        t = t + Sheets("Solution").Cells(6, 2).Value
    Loop
    
    Application.ScreenUpdating = True
End Sub

Sub GenerateGrid()
    'Generates a Calculation Grid based on the inputs from the user
    Application.ScreenUpdating = False
    Sheets("Grid-BC").Select
    Dim Station As Long
    Dim SRow As Integer
    Dim LRow As Integer
    Dim z As Double
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    ' Cleans the Grid
    Range("A27").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    'Calculates Depths (for reference to the user)
    SRow = 4
    LRow = SRow + Cells(1, 10).Value - 1
    
    For i = SRow To LRow
        Cells(i, 9).Value = YN(Cells(14, 2).Value, Cells(15, 6).Value, Cells(i, 8).Value, Cells(i, 5).Value, Cells(i, 6).Value, Cells(i, 7).Value)
        Cells(i, 10).Value = YC(Cells(14, 2).Value, Cells(i, 5).Value, Cells(i, 6).Value, Cells(i, 7).Value)
    Next i
    
    
    'Writes the Grid
    Station = Cells(SRow, 2).Value
    z = Cells(1, 2).Value + Cells(4, 4).Value * Cells(4, 8).Value
    i = 0
    j = 0
    k = 0
    
    Do While SRow <= LRow
        If Station + k * Cells(SRow, 4) <= Cells(SRow, 3) Then
            Cells(27 + i, 1).Value = i + 1
            Cells(27 + i, 2).Value = Station + k * Cells(SRow, 4)
            Cells(27 + i, 3).Value = z - Cells(SRow, 4) * Cells(SRow, 8)
            Cells(27 + i, 4).Value = Cells(27 + i, 3).Value + Cells(1, 6).Value
            Cells(27 + i, 5).Value = Cells(SRow, 5)
            z = Cells(27 + i, 3).Value
            Cells(27 + i, 6).Value = Cells(SRow, 6).Value
            Cells(27 + i, 7).Value = Cells(SRow, 7).Value
            Cells(27 + i, 8).Value = Cells(SRow, 9)
            Cells(27 + i, 9).Value = Cells(27 + i, 3) + Cells(27 + i, 8)
            Cells(27 + i, 10).Value = Cells(27 + i, 3) + Cells(SRow, 9)
            Cells(27 + i, 11).Value = Cells(27 + i, 3) + Cells(SRow, 10)
            Cells(27 + i, 12).Value = 0
            
            i = i + 1
            k = k + 1
        Else
            k = 1
            SRow = SRow + 1
            Station = Cells(SRow, 2).Value
        End If
    Loop
    Cells(18, 2).Value = i + 2
    Application.ScreenUpdating = True
End Sub

Sub Reset_US()
    'Cleans the Upstream Calculation Sheet and defines a size to match the Grid
    Application.ScreenUpdating = False
    Dim LRow As Integer
    LRow = Sheets("Grid-BC").Cells(18, 2).Value
    Sheets("US").Select
    Range("A4:W4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("Grid-BC").Select
    Range("A27:H" & LRow + 24).Select
    Selection.Copy
    Sheets("US").Select
    Range("K3").FormulaR1C1 = "='Grid-BC'!R15C6"
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Range("I3:W3").Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("I3:W" & LRow)
    Range("U32").Select
    Selection.ClearContents
    Range("Z8").FormulaR1C1 = "='Grid-BC'!R[3]C[-10]"
    Range("Z9").FormulaR1C1 = "='Grid-BC'!R[2]C[-11]"
    Range("A1").Select
    Application.ScreenUpdating = True
End Sub

Sub Solve_US()
    'Solves the Backwater Curve forwards from the US Boundary Condition
    Application.ScreenUpdating = False
    Sheets("US").Select
    Dim i As Integer
    Dim LRow As Integer
    Dim hc As Double
    LRow = Sheets("Grid-BC").Cells(18, 2).Value
    Cells(3, 8).Value = Cells(8, 26).Value - Cells(3, 3)
    For i = 3 To LRow - 1
        SolverReset
        SolverAdd CellRef:=Cells(i + 1, 8).Address, Relation:=3, FormulaText:="0.01" 'Water depth above 0.01
        SolverAdd CellRef:=Cells(i + 1, 18).Address, Relation:=3, FormulaText:="1" 'Froude number more than 1
        SolverOk SetCell:=Cells(i, 21).Address, MaxMinVal:=2, ValueOf:=0, ByChange:=Cells(i + 1, 8).Address, Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverSolve True
        
        'If the computed water depth is wrong or an incorrect value,  calculates and writes the critical depth
        If Cells(i + 1, 18) <= 1 Or Cells(i + 1, 8) <= 0 Then
            hc = YC(Cells(i + 1, 10).Value, Cells(i + 1, 5).Value, Cells(i + 1, 6).Value, Cells(i + 1, 7).Value)
            Cells(i + 1, 8).Value = hc
        End If
    Next i
    Application.ScreenUpdating = True
End Sub

Sub Reset_DS()
    'Cleans the Upstream Calculation Sheet and defines a size to match the Grid
    Application.ScreenUpdating = False
    Dim LRow As Integer
    LRow = Sheets("Grid-BC").Cells(18, 2).Value
    Sheets("DS").Select
    Range("A4:W4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("Grid-BC").Select
    Range("A27:H" & LRow + 24).Select
    Selection.Copy
    Sheets("DS").Select
    Range("K3").FormulaR1C1 = "='Grid-BC'!R15C6"
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Range("I3:W3").Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("I3:W" & LRow)
    Range("U32").Select
    Selection.ClearContents
    Range("Z8").FormulaR1C1 = "='Grid-BC'!R[3]C[-9]"
    Range("Z9").FormulaR1C1 = "='Grid-BC'!R[2]C[-11]"
    Range("A1").Select
    Application.ScreenUpdating = True
End Sub

Sub Solve_DS()
    Application.ScreenUpdating = False
    Sheets("DS").Select
    Dim i As Integer
    Dim LRow As Integer
    Dim hc As Double
    LRow = Sheets("Grid-BC").Cells(18, 2).Value
    Cells(LRow, 8).Value = Cells(8, 26).Value - Cells(LRow, 3)
    For i = LRow To 4 Step -1
        SolverReset
        SolverAdd CellRef:=Cells(i - 1, 8).Address, Relation:=3, FormulaText:="0.01" 'Water depth above 0.01
        SolverAdd CellRef:=Cells(i - 1, 18).Address, Relation:=1, FormulaText:="1" 'Froude number less than 1
        SolverOk SetCell:=Cells(i - 1, 21).Address, MaxMinVal:=2, ValueOf:=0, ByChange:=Cells(i - 1, 8).Address, Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverSolve True
        
        'If the computed water depth is wrong or an incorrect value,  calculates and writes the critical depth
        If Cells(i - 1, 18) >= 1 Or Cells(i - 1, 8) <= 0 Then
            hc = YC(Cells(i - 1, 10).Value, Cells(i - 1, 5).Value, Cells(i - 1, 6).Value, Cells(i - 1, 7).Value)
            Cells(i - 1, 8).Value = hc
        End If
    Next i
    Application.ScreenUpdating = True
End Sub

Sub Reset_Solution()
    'Cleans the Solution Sheet and defines a size to match the Grid
    Application.ScreenUpdating = False
    Sheets("Solution").Select
    Dim LRow As Integer
    LRow = Sheets("Grid-BC").Cells(18, 2).Value
    Range("M4:CZ4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    'Different considerations for the sediment discharge (qs) calculation (only qb, qb + qss)
    If Sheets("Grid-BC").Range("T8") = 1 Then
        Range("AW3").FormulaR1C1 = "=RC[-10]+RC[-1]+'Grid-BC'!R[24]C[-37]"
    Else
        Range("AW3").FormulaR1C1 = "=RC[-10]+'Grid-BC'!R[24]C[-37]"
    End If
    
    Range("M3:BD3").Select
    Selection.AutoFill Destination:=Range("M3:BD" & LRow), Type:=xlFillDefault
    Range("AG" & LRow).ClearContents
    Range("AO" & LRow).ClearContents
    Range("B6").Formula = "=MIN(AO3:AO" & LRow & ")" 'Calculates the minimum dt
    ActiveWorkbook.Save
    Application.ScreenUpdating = True
End Sub

Sub CalculateZNew()
    'Calculates a new bed level using the Exner equation <- changes the formula in the boundaries
    Application.ScreenUpdating = False
    Dim i As Integer
    Dim LRow As Integer
    Dim zNew As Double
    Dim dt As Double
    
    LRow = Sheets("Grid-BC").Cells(18, 2).Value
    Range("B6").Formula = "=MIN(AO3:AO" & LRow & ")"
    dt = Range("B6").Value
    For i = 3 To LRow - 1
        Range("AP" & i).Select
        If i = 3 Then
            zNew = Cells(i, 15).Value - ((Cells(i + 1, 49).Value - Cells(i, 49).Value) * dt) / ((1 - Cells(7, 6).Value) * (Cells(i + 1, 14).Value - Cells(i, 14).Value))
            If zNew > Cells(i, 16).Value Then
                ActiveCell.Value = zNew
            Else
                ActiveCell.Value = Cells(i, 16).Value
            End If
        Else
            zNew = 0.5 * (Cells(i + 1, 15).Value + Cells(i - 1, 15).Value) - (Cells(i + 1, 49).Value - Cells(i - 1, 49).Value) * dt / ((1 - Cells(7, 6).Value) * (Cells(i + 1, 14).Value - Cells(i - 1, 14).Value))
            If zNew > Cells(i, 16).Value Then
                ActiveCell.Value = zNew
            Else
                ActiveCell.Value = Cells(i, 16).Value
            End If
        End If
    Next i
    i = LRow
    Range("AP" & i).Select
    zNew = Cells(i, 15).Value - (Cells(i - 1, 49).Value - Cells(i, 49).Value) * dt / ((1 - Cells(7, 6).Value) * (Cells(i - 1, 14).Value - Cells(i, 14).Value))
    If zNew > Cells(i, 16).Value Then
        ActiveCell.Value = zNew
    Else
        ActiveCell.Value = Cells(i, 16).Value
    End If
    Application.ScreenUpdating = True
End Sub

Function Hs(Q As Double, Hs1 As Double, B As Double, Sf As Double, D50 As Double, D90 As Double, g As Double, R As Double)
    'Estimation of Hs based on a starting random Hs; used in the Einstein decomposition as proposed by Bateman (2020)
    Dim ks As Double
    Dim qw As Double
    Dim h As Double
    Dim taus As Double
    Dim tauss As Double
    Dim Hs2 As Double
    
    ks = 3 * D90
    qw = Q / B
    h = qw / (8.32 * Sqr(g * Hs1 * Sf) * ((Hs1 / ks) ^ (1 / 6)))
    'Fr = qw / Sqr(g * H)
    taus = h * Sf / (R * D50)
    tauss = 0.06 + 0.4 * (taus)
    'tauss = 0.05 + 0.7 * (taus * (Fr ^ 0.7)) ^ 0.8
    'Wrigth and Parker can also be used without the correction for stratification
    Hs2 = (tauss * R * D50) / Sf
    Hs = Hs2
End Function

Sub Calculate_Bedforms()
    'Solver to find the correct Hs and H values that fulfill the Engelund-Hansen stress relationship in the Einstein decomposition.
    Application.ScreenUpdating = False
    Dim i As Integer
    Dim LRow As Integer
    LRow = Sheets("Grid-BC").Cells(18, 2).Value
    For i = 3 To LRow
        Cells(i, 43).Value = 1
        SolverReset
        SolverAdd CellRef:=Cells(i, 43).Address, Relation:=3, FormulaText:="0.01"
        SolverOk SetCell:=Cells(i, 45).Address, MaxMinVal:=2, ValueOf:=0, ByChange:=Cells(i, 43).Address, Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverSolve True
    Next i
    Application.ScreenUpdating = True
End Sub

Function Find_TauMIN(Fr As Double)
    'Implementation of the Newton-Raphson scheme given by Parker (2004) to find TauMIN
    Dim TauMIN0, TauMIN1, F, dF, tol As Double
    
    TauMIN0 = 0.4
    Iter = 0
    tol = 1
    
    Do While tol > 0.001
        Iter = Iter + 1
        F = TauMIN0 - 0.05 - 0.7 * (TauMIN0 ^ (4 / 5)) * (Fr ^ (14 / 25))
        dF = 1 - 0.7 * (4 / 5) * (TauMIN0 ^ (-1 / 5)) * (Fr ^ (14 / 25))
        TauMIN1 = TauMIN0 - F / dF
        tol = Abs(2 * (TauMIN1 - TauMIN0) / (TauMIN1 + TauMIN0))
        TauMIN0 = TauMIN1
    Loop
    
    Find_TauMIN = TauMIN0
End Function

Function Rouse(h As Double, vs As Double, us As Double, kc As Double) As Double
'Solves the Rouse-Vanoni sediment concentration profile
    Dim c1, c2, x1, x2 As Double
    Dim i As Double
    Dim Integral As Double
    
    Integral = 0
    For i = 0.05 To 0.99 Step 0.01
        c1 = i
        c2 = i + 0.01
        x1 = ((((1 - c1) / c1) / ((1 - 0.05) / 0.05)) ^ (vs / (0.4 * us))) * Application.WorksheetFunction.Ln(30 * (h / kc) * c1)
        x2 = ((((1 - c2) / c2) / ((1 - 0.05) / 0.05)) ^ (vs / (0.4 * us))) * Application.WorksheetFunction.Ln(30 * (h / kc) * c2)
        Integral = Integral + 0.5 * (x1 + x2) * (c2 - c1)
    Next i
    Rouse = Integral
End Function

Sub Update_Bed()
    'Copies the Znew values from the Solution Sheet to the US and DS Calculation Sheets
    Application.ScreenUpdating = False
    Sheets("Solution").Select
    Range("AP3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("US").Select
    Range("C3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Sheets("DS").Select
    Range("C3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.ScreenUpdating = True
End Sub

Sub Update_Manning()
    'Copies the nnew values from the Solution Sheet to the US and DS Calculation Sheets
    Application.ScreenUpdating = False
    Sheets("Solution").Select
    Range("AU3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("US").Select
    Range("K3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Sheets("DS").Select
    Range("K3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.ScreenUpdating = True
End Sub

Sub Paste_Solution(Column As Integer, t As Double)
    'Saves certain parameters of the Solution for every time step
    Application.ScreenUpdating = False
    Dim LRow As Integer
    LRow = Sheets("Grid-BC").Cells(18, 2).Value
    Sheets("Solution").Select
    Range("O3:O" & LRow).Select
    Selection.Copy
    Sheets("z").Select
    Cells(1, Column).Value = t
    Cells(3, Column).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Sheets("Solution").Select
    Range("AB3:AB" & LRow).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("u").Select
    Cells(1, Column).Value = t
    Cells(3, Column).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Sheets("Solution").Select
    Range("AF3:AF" & LRow).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sf").Select
    Cells(1, Column).Value = t
    Cells(3, Column).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Sheets("Solution").Select
    Range("U3:U" & LRow).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("z+h").Select
    Cells(1, Column).Value = t
    Cells(3, Column).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Sheets("Solution").Select
    Range("AW3:AW" & LRow).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("qt").Select
    Cells(1, Column).Value = t
    Cells(3, Column).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Sheets("Solution").Select
    Range("AJ3:AJ" & LRow).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("T0").Select
    Cells(1, Column).Value = t
    Cells(3, Column).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Sheets("Solution").Select
    Range("AK3:AK" & LRow).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Shields").Select
    Cells(1, Column).Value = t
    Cells(3, Column).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Sheets("Solution").Select
    Range("AI3:AI" & LRow).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("M").Select
    Cells(1, Column).Value = t
    Cells(3, Column).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Sheets("Solution").Select
    Range("AD3:AD" & LRow).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Fr").Select
    Cells(1, Column).Value = t
    Cells(3, Column).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Sheets("Solution").Select
    Range("W3:W" & LRow).Select
    Selection.Copy
    Sheets("n").Select
    Cells(1, Column).Value = t
    Cells(3, Column).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.ScreenUpdating = True
End Sub

Sub Clear_Sheets()
    Application.ScreenUpdating = False
    Sheets("z").Cells.Clear
    Sheets("u").Cells.Clear
    Sheets("Sf").Cells.Clear
    Sheets("z+h").Cells.Clear
    Sheets("Shields").Cells.Clear
    Sheets("qt").Cells.Clear
    Sheets("T0").Cells.Clear
    Sheets("M").Cells.Clear
    Sheets("Fr").Cells.Clear
    Sheets("n").Cells.Clear
    Application.ScreenUpdating = True
End Sub


Sub Default_Case()
    'Default values by Bateman (2020) in the Excel Sheet: BackWaterFRM8-V0
    Range("B1").Value = 10.635
    Range("F1").Value = -0.1
    Range("J1").Value = 2
    Range("A4").Value = 1
    Range("B4").Value = 100
    Range("C4").Value = 170
    Range("D4").Value = 5
    Range("E4").Value = 10
    Range("F4").Value = 0
    Range("G4").Value = 0
    Range("H4").Value = 0.008
    Range("A5").Value = 2
    Range("B5").Value = 170
    Range("C5").Value = 245
    Range("D5").Value = 5
    Range("E5").Value = 10
    Range("F5").Value = 0
    Range("G5").Value = 0
    Range("H5").Value = 0.001
    Range("B11").Value = 9.81
    Range("B15").Value = 1000
    Range("B16").Value = 0.000001
    Range("F14").FormulaR1C1 = "=0.8/1000"
    Range("F16").Value = 2650
    Range("F20").Value = 1.65
    Range("F21").Value = 35
    Range("F22").Value = 0.4
    Range("O3").Value = 0.7 'Sea Level Rise (SLR) from IPCC 8.5
    Range("N11").Select
    Range("N11:Q100").ClearContents
    Range("N11").Value = 0
    Range("O11").Value = 20
    Range("P11").Value = 11.4
    Range("Q11").Value = 10.6
    Range("N12").Value = 1000
    Range("O12").Value = 20
    Range("P12").Value = 11.4
    Range("Q12").Value = 10.6
    Range("F23").FormulaR1C1 = "=8/1000"
    Range("F15").FormulaR1C1 = "=((R[-1]C)^(1/6))/21"
    Range("O7").FormulaR1C1 = "=MAX(R[4]C[-1]:R[93]C[-1])"
    'Range("T3").Value = 0
    'Range("T7").Value = 0
    'Range("T8").Value = 0
    'Range("T9").Value = 0
    Call GenerateGrid
End Sub