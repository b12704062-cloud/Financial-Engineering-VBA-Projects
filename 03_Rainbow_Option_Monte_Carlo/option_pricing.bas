Attribute VB_Name = "Module1"
Option Explicit

'========================================================
' HW3 Rainbow Option VBA
'========================================================

Private SavedZBasic() As Double      ' (rep, sim, asset)
Private SavedZBonus1() As Double     ' (rep, sim, asset)

Private SavedBasicReady As Boolean
Private SavedBonus1Ready As Boolean

Private SavedNumSim As Long
Private SavedNumRep As Long
Private SavedN As Long


'========================================================
' 1. Generate input table
'========================================================

Public Sub Generate_HW3_InputTable()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    Dim n As Long
    n = CLng(ws.Range("B6").Value)  ' num_asset
    
    If n <= 0 Then
        MsgBox "Please enter a valid num_asset in B6.", vbExclamation
        Exit Sub
    End If
    
    Dim i As Long, j As Long
    Dim corrCol As Long, covCol As Long
    
    corrCol = 15                  ' O column
    covCol = corrCol + n + 4      ' covariance matrix starts after correlation matrix
    
    ws.Range(ws.Cells(1, 10), ws.Cells(200, 120)).ClearContents
    
    ' Asset inputs
    ws.Cells(1, 10).Value = "Asset"
    ws.Cells(1, 11).Value = "S0"
    ws.Cells(1, 12).Value = "q"
    ws.Cells(1, 13).Value = "sigma"
    
    For i = 1 To n
        ws.Cells(i + 1, 10).Value = "S" & i
        ws.Cells(i + 1, 11).Value = ""
        ws.Cells(i + 1, 12).Value = ""
        ws.Cells(i + 1, 13).Value = ""
    Next i
    
    ' Correlation matrix
    ws.Cells(1, corrCol).Value = "Correlation matrix rho"
    
    For j = 1 To n
        ws.Cells(2, corrCol + j).Value = "S" & j
        ws.Cells(2 + j, corrCol).Value = "S" & j
    Next j
    
    For i = 1 To n
        For j = 1 To n
            If i = j Then
                ws.Cells(i + 2, corrCol + j).Value = 1
            Else
                ws.Cells(i + 2, corrCol + j).Value = ""
            End If
        Next j
    Next i
    
    ' Covariance matrix
    ws.Cells(1, covCol).Value = "Covariance matrix"
    
    For j = 1 To n
        ws.Cells(2, covCol + j).Value = "S" & j
        ws.Cells(2 + j, covCol).Value = "S" & j
    Next j
    
    For i = 1 To n
        For j = 1 To n
            ws.Cells(i + 2, covCol + j).Value = ""
        Next j
    Next i
    
    SavedBasicReady = False
    SavedBonus1Ready = False
    
End Sub


'========================================================
' 2. Basic
'========================================================

Public Sub Run_HW3_Basic()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    Dim k As Double, R As Double, T As Double
    Dim numSim As Long, numRep As Long, n As Long
    
    ReadGeneralInputs ws, k, R, T, numSim, numRep, n
    
    Dim S0() As Double, q() As Double, sigma() As Double
    ReadAssetInputs ws, n, S0, q, sigma
    
    Dim cov() As Double
    cov = GetCovMatrix(ws, n, sigma)
    
    Dim A() As Double
    A = CholeskyUpper(cov, n)
    
    ' 每次重新跑 Basic，就重設共同亂數
    ReDim SavedZBasic(1 To numRep, 1 To numSim, 1 To n)
    ReDim SavedZBonus1(1 To numRep, 1 To numSim, 1 To n)
    
    SavedNumSim = numSim
    SavedNumRep = numRep
    SavedN = n
    
    SavedBasicReady = True
    SavedBonus1Ready = False
    
    Randomize 12345
    
    Dim rep As Long, sim As Long, asset As Long
    For rep = 1 To numRep
        For sim = 1 To numSim
            For asset = 1 To n
                SavedZBasic(rep, sim, asset) = StdNormal()
            Next asset
        Next sim
    Next rep
    
    Dim repValue() As Double
    ReDim repValue(1 To numRep)
    
    Dim ZCorr() As Double
    
    For rep = 1 To numRep
        ZCorr = MultiplySavedByUpper(SavedZBasic, rep, numSim, n, A)
        repValue(rep) = OneReplicationValue(ZCorr, numSim, n, S0, q, sigma, k, R, T)
    Next rep
    
    Dim result(1 To 3) As Double
    GetFinalResult repValue, numRep, result
    
    ws.Range("D2").Value = result(1) ' option value
    ws.Range("D3").Value = result(2) ' CI upper
    ws.Range("D4").Value = result(3) ' CI lower
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error in Basic: " & Err.Description, vbCritical
End Sub


'========================================================
' 3. Bonus1: Antithetic + Moment Matching
'========================================================

Public Sub Run_HW3_Bonus1()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    Dim k As Double, R As Double, T As Double
    Dim numSim As Long, numRep As Long, n As Long
    
    ReadGeneralInputs ws, k, R, T, numSim, numRep, n
    
    If Not SavedBasicReady Then
        MsgBox "Please run Basic first. Bonus1 uses Basic's common random samples.", vbExclamation
        Exit Sub
    End If
    
    CheckSavedDimension numSim, numRep, n
    
    If numSim Mod 2 <> 0 Then
        MsgBox "Bonus1 requires num_sim to be an even number.", vbExclamation
        Exit Sub
    End If
    
    Dim S0() As Double, q() As Double, sigma() As Double
    ReadAssetInputs ws, n, S0, q, sigma
    
    Dim cov() As Double
    cov = GetCovMatrix(ws, n, sigma)
    
    Dim A() As Double
    A = CholeskyUpper(cov, n)
    
    Dim halfN As Long
    halfN = numSim \ 2
    
    Dim rep As Long, sim As Long, asset As Long
    
    ' 用 Basic 的前半部亂數，後半部直接乘 -1
    For rep = 1 To numRep
        For sim = 1 To halfN
            For asset = 1 To n
                SavedZBonus1(rep, sim, asset) = SavedZBasic(rep, sim, asset)
                SavedZBonus1(rep, sim + halfN, asset) = -SavedZBasic(rep, sim, asset)
            Next asset
        Next sim
        
        MomentMatchSaved SavedZBonus1, rep, numSim, n
    Next rep
    
    SavedBonus1Ready = True
    
    Dim repValue() As Double
    ReDim repValue(1 To numRep)
    
    Dim ZCorr() As Double
    
    For rep = 1 To numRep
        ZCorr = MultiplySavedByUpper(SavedZBonus1, rep, numSim, n, A)
        repValue(rep) = OneReplicationValue(ZCorr, numSim, n, S0, q, sigma, k, R, T)
    Next rep
    
    Dim result(1 To 3) As Double
    GetFinalResult repValue, numRep, result
    
    ws.Range("F2").Value = result(1)
    ws.Range("F3").Value = result(2)
    ws.Range("F4").Value = result(3)
    
    ' Degree of CI = Bonus1 CI width / Basic CI width
    If (ws.Range("D3").Value - ws.Range("D4").Value) <> 0 Then
        ws.Range("F5").Value = (result(2) - result(3)) / (ws.Range("D3").Value - ws.Range("D4").Value)
        ws.Range("F5").NumberFormat = "0.0000%"
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error in Bonus1: " & Err.Description, vbCritical
End Sub


'========================================================
' 4. Bonus2: Antithetic + Moment Matching + Inverse Cholesky
'========================================================

Public Sub Run_HW3_Bonus2()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    Dim k As Double, R As Double, T As Double
    Dim numSim As Long, numRep As Long, n As Long
    
    ReadGeneralInputs ws, k, R, T, numSim, numRep, n
    
    If Not SavedBonus1Ready Then
        MsgBox "Please run Bonus1 first. Bonus2 uses Bonus1's moment matched samples.", vbExclamation
        Exit Sub
    End If
    
    CheckSavedDimension numSim, numRep, n
    
    Dim S0() As Double, q() As Double, sigma() As Double
    ReadAssetInputs ws, n, S0, q, sigma
    
    Dim cov() As Double
    cov = GetCovMatrix(ws, n, sigma)
    
    Dim A() As Double
    A = CholeskyUpper(cov, n)
    
    Dim repValue() As Double
    ReDim repValue(1 To numRep)
    
    Dim rep As Long
    Dim ZFinal() As Double
    
    For rep = 1 To numRep
        ZFinal = ApplyInverseCholeskySaved(SavedZBonus1, rep, numSim, n, A)
        repValue(rep) = OneReplicationValue(ZFinal, numSim, n, S0, q, sigma, k, R, T)
    Next rep
    
    Dim result(1 To 3) As Double
    GetFinalResult repValue, numRep, result
    
    ws.Range("H2").Value = result(1)
    ws.Range("H3").Value = result(2)
    ws.Range("H4").Value = result(3)
    
    ' Degree of CI = Bonus2 CI width / Basic CI width
    If (ws.Range("D3").Value - ws.Range("D4").Value) <> 0 Then
        ws.Range("H5").Value = (result(2) - result(3)) / (ws.Range("D3").Value - ws.Range("D4").Value)
        ws.Range("H5").NumberFormat = "0.0000%"
    End If

    Exit Sub
    
ErrHandler:
    MsgBox "Error in Bonus2: " & Err.Description, vbCritical
End Sub


'========================================================
' Read inputs
'========================================================

Private Sub ReadGeneralInputs( _
    ByRef ws As Worksheet, _
    ByRef k As Double, _
    ByRef R As Double, _
    ByRef T As Double, _
    ByRef numSim As Long, _
    ByRef numRep As Long, _
    ByRef n As Long)
    
    k = CDbl(ws.Range("B1").Value)
    R = CDbl(ws.Range("B2").Value)
    T = CDbl(ws.Range("B3").Value)
    numSim = CLng(ws.Range("B4").Value)
    numRep = CLng(ws.Range("B5").Value)
    n = CLng(ws.Range("B6").Value)
    
    If k <= 0 Or T <= 0 Or numSim <= 1 Or numRep <= 1 Or n <= 0 Then
        Err.Raise 2001, , "Please check K, T, num_sim, num_rep, and num_asset."
    End If
End Sub


Private Sub ReadAssetInputs( _
    ByRef ws As Worksheet, _
    ByVal n As Long, _
    ByRef S0() As Double, _
    ByRef q() As Double, _
    ByRef sigma() As Double)
    
    Dim i As Long
    
    ReDim S0(1 To n)
    ReDim q(1 To n)
    ReDim sigma(1 To n)
    
    For i = 1 To n
        S0(i) = CDbl(ws.Cells(i + 1, 11).Value)      ' K column
        q(i) = CDbl(ws.Cells(i + 1, 12).Value)       ' L column
        sigma(i) = CDbl(ws.Cells(i + 1, 13).Value)   ' M column
        
        If S0(i) <= 0 Or sigma(i) <= 0 Then
            Err.Raise 2002, , "Please check S0 and sigma for asset " & i & "."
        End If
    Next i
End Sub


Private Sub CheckSavedDimension(ByVal numSim As Long, ByVal numRep As Long, ByVal n As Long)
    If numSim <> SavedNumSim Or numRep <> SavedNumRep Or n <> SavedN Then
        Err.Raise 2003, , "Inputs have changed. Please run Basic again."
    End If
End Sub


'========================================================
' Covariance / correlation matrix
'========================================================

Private Function GetCovMatrix(ByRef ws As Worksheet, ByVal n As Long, ByRef sigma() As Double) As Double()
    Dim corrCol As Long, covCol As Long
    
    corrCol = 15
    covCol = corrCol + n + 4
    
    Dim cov() As Double
    ReDim cov(1 To n, 1 To n)
    
    Dim i As Long, j As Long
    Dim useCov As Boolean
    Dim v1 As Variant, v2 As Variant
    Dim covVal As Double
    
    useCov = True
    
    ' 判斷 covariance matrix 是否可用：
    ' 對角線一定要填；非對角線只要其中一邊有填即可。
    For i = 1 To n
        For j = i To n
            v1 = ws.Cells(i + 2, covCol + j).Value
            v2 = ws.Cells(j + 2, covCol + i).Value
            
            If i = j Then
                If v1 = "" Or Not IsNumeric(v1) Then
                    useCov = False
                    Exit For
                End If
            Else
                If (v1 = "" Or Not IsNumeric(v1)) And (v2 = "" Or Not IsNumeric(v2)) Then
                    useCov = False
                    Exit For
                End If
            End If
        Next j
        
        If Not useCov Then Exit For
    Next i
    
    If useCov Then
        ' 使用 covariance matrix
        For i = 1 To n
            For j = 1 To n
                v1 = ws.Cells(i + 2, covCol + j).Value
                v2 = ws.Cells(j + 2, covCol + i).Value
                
                If i = j Then
                    covVal = CDbl(v1)
                ElseIf v1 <> "" And IsNumeric(v1) Then
                    covVal = CDbl(v1)
                ElseIf v2 <> "" And IsNumeric(v2) Then
                    covVal = CDbl(v2)
                Else
                    Err.Raise 2004, , "Missing covariance value."
                End If
                
                cov(i, j) = covVal
            Next j
        Next i
        
    Else
        ' 使用 sigma + correlation matrix
        Dim rho As Double
        
        For i = 1 To n
            For j = 1 To n
                If i = j Then
                    rho = 1
                Else
                    v1 = ws.Cells(i + 2, corrCol + j).Value
                    v2 = ws.Cells(j + 2, corrCol + i).Value
                    
                    If v1 <> "" And IsNumeric(v1) Then
                        rho = CDbl(v1)
                    ElseIf v2 <> "" And IsNumeric(v2) Then
                        rho = CDbl(v2)
                    Else
                        Err.Raise 2005, , "Please fill rho(" & i & "," & j & ") or rho(" & j & "," & i & ")."
                    End If
                End If
                
                cov(i, j) = rho * sigma(i) * sigma(j)
            Next j
        Next i
    End If
    
    GetCovMatrix = cov
End Function


'========================================================
' Option payoff
'========================================================

Private Function OneReplicationValue( _
    ByRef ZCorr() As Double, _
    ByVal numSim As Long, _
    ByVal n As Long, _
    ByRef S0() As Double, _
    ByRef q() As Double, _
    ByRef sigma() As Double, _
    ByVal k As Double, _
    ByVal R As Double, _
    ByVal T As Double) As Double
    
    Dim sim As Long, i As Long
    Dim ST As Double, maxST As Double
    Dim payoff As Double, sumPayoff As Double
    
    sumPayoff = 0
    
    For sim = 1 To numSim
        maxST = -1E+100
        
        For i = 1 To n
            ST = S0(i) * Exp((R - q(i) - 0.5 * sigma(i) ^ 2) * T + Sqr(T) * ZCorr(sim, i))
            
            If ST > maxST Then
                maxST = ST
            End If
        Next i
        
        payoff = Application.WorksheetFunction.Max(maxST - k, 0)
        sumPayoff = sumPayoff + payoff
    Next sim
    
    OneReplicationValue = Exp(-R * T) * sumPayoff / numSim
End Function


'========================================================
' Random number
'========================================================

Private Function StdNormal() As Double
    Dim u1 As Double, u2 As Double
    
    Do
        u1 = Rnd
    Loop While u1 <= 0
    
    u2 = Rnd
    
    StdNormal = Sqr(-2 * Log(u1)) * Cos(2 * WorksheetFunction.Pi() * u2)
End Function


'========================================================
' Moment matching
'========================================================

Private Sub MomentMatchSaved( _
    ByRef ZSaved() As Double, _
    ByVal rep As Long, _
    ByVal numSim As Long, _
    ByVal n As Long)
    
    Dim asset As Long, sim As Long
    Dim meanZ As Double, sdZ As Double
    Dim sumZ As Double, sumSq As Double
    
    For asset = 1 To n
        sumZ = 0
        
        For sim = 1 To numSim
            sumZ = sumZ + ZSaved(rep, sim, asset)
        Next sim
        
        meanZ = sumZ / numSim
        
        sumSq = 0
        
        For sim = 1 To numSim
            sumSq = sumSq + (ZSaved(rep, sim, asset) - meanZ) ^ 2
        Next sim
        
        sdZ = Sqr(sumSq / numSim)
        
        If sdZ = 0 Then
            Err.Raise 3001, , "Zero standard deviation in moment matching."
        End If
        
        For sim = 1 To numSim
            ZSaved(rep, sim, asset) = (ZSaved(rep, sim, asset) - meanZ) / sdZ
        Next sim
    Next asset
End Sub


'========================================================
' Matrix operations
'========================================================

Private Function CholeskyUpper(ByRef C() As Double, ByVal n As Long) As Double()
    Dim A() As Double
    ReDim A(1 To n, 1 To n)
    
    Dim i As Long, j As Long, k As Long
    Dim temp As Double
    
    For i = 1 To n
        For j = i To n
            temp = C(i, j)
            
            For k = 1 To i - 1
                temp = temp - A(k, i) * A(k, j)
            Next k
            
            If i = j Then

                
                A(i, i) = Sqr(temp)
            Else
                A(i, j) = temp / A(i, i)
            End If
        Next j
    Next i
    
    CholeskyUpper = A
End Function


Private Function MultiplySavedByUpper( _
    ByRef ZSaved() As Double, _
    ByVal rep As Long, _
    ByVal numSim As Long, _
    ByVal n As Long, _
    ByRef A() As Double) As Double()
    
    Dim R() As Double
    ReDim R(1 To numSim, 1 To n)
    
    Dim sim As Long, j As Long, k As Long
    Dim temp As Double
    
    For sim = 1 To numSim
        For j = 1 To n
            temp = 0
            
            For k = 1 To j
                temp = temp + ZSaved(rep, sim, k) * A(k, j)
            Next k
            
            R(sim, j) = temp
        Next j
    Next sim
    
    MultiplySavedByUpper = R
End Function


Private Function MultiplyArrayByUpper( _
    ByRef Z() As Double, _
    ByVal numSim As Long, _
    ByVal n As Long, _
    ByRef A() As Double) As Double()
    
    Dim R() As Double
    ReDim R(1 To numSim, 1 To n)
    
    Dim sim As Long, j As Long, k As Long
    Dim temp As Double
    
    For sim = 1 To numSim
        For j = 1 To n
            temp = 0
            
            For k = 1 To j
                temp = temp + Z(sim, k) * A(k, j)
            Next k
            
            R(sim, j) = temp
        Next j
    Next sim
    
    MultiplyArrayByUpper = R
End Function


Private Function ApplyInverseCholeskySaved( _
    ByRef ZSaved() As Double, _
    ByVal rep As Long, _
    ByVal numSim As Long, _
    ByVal n As Long, _
    ByRef A() As Double) As Double()
    
    Dim ZTemp() As Double
    ReDim ZTemp(1 To numSim, 1 To n)
    
    Dim sim As Long, asset As Long
    
    For sim = 1 To numSim
        For asset = 1 To n
            ZTemp(sim, asset) = ZSaved(rep, sim, asset)
        Next asset
    Next sim
    
    Dim sampleCov() As Double
    sampleCov = SampleCovMatrix(ZTemp, numSim, n)
    
    Dim B() As Double
    B = CholeskyUpper(sampleCov, n)
    
    Dim BInv() As Double
    BInv = InverseUpper(B, n)
    
    Dim ZWhite() As Double
    ZWhite = MultiplyArrayByUpper(ZTemp, numSim, n, BInv)
    
    Dim ZFinal() As Double
    ZFinal = MultiplyArrayByUpper(ZWhite, numSim, n, A)
    
    ApplyInverseCholeskySaved = ZFinal
End Function


Private Function SampleCovMatrix(ByRef Z() As Double, ByVal numSim As Long, ByVal n As Long) As Double()
    Dim C() As Double
    ReDim C(1 To n, 1 To n)
    
    Dim meanZ() As Double
    ReDim meanZ(1 To n)
    
    Dim sim As Long, j As Long, k As Long
    
    For j = 1 To n
        meanZ(j) = 0
        
        For sim = 1 To numSim
            meanZ(j) = meanZ(j) + Z(sim, j)
        Next sim
        
        meanZ(j) = meanZ(j) / numSim
    Next j
    
    For j = 1 To n
        For k = j To n
            C(j, k) = 0
            
            For sim = 1 To numSim
                C(j, k) = C(j, k) + (Z(sim, j) - meanZ(j)) * (Z(sim, k) - meanZ(k))
            Next sim
            
            C(j, k) = C(j, k) / numSim
            C(k, j) = C(j, k)
        Next k
    Next j
    
    SampleCovMatrix = C
End Function


Private Function InverseUpper(ByRef U() As Double, ByVal n As Long) As Double()
    Dim Inv() As Double
    ReDim Inv(1 To n, 1 To n)
    
    Dim i As Long, j As Long, k As Long
    Dim sumVal As Double
    
    For i = n To 1 Step -1
        If U(i, i) = 0 Then
            Err.Raise 3003, , "Cannot invert matrix. Zero diagonal element."
        End If
        
        Inv(i, i) = 1 / U(i, i)
        
        For j = i + 1 To n
            sumVal = 0
            
            For k = i + 1 To j
                sumVal = sumVal + U(i, k) * Inv(k, j)
            Next k
            
            Inv(i, j) = -sumVal / U(i, i)
        Next j
    Next i
    
    InverseUpper = Inv
End Function


'========================================================
' Statistics
'========================================================

Private Sub GetFinalResult(ByRef repValue() As Double, ByVal numRep As Long, ByRef result() As Double)
    Dim meanV As Double, sdV As Double
    
    meanV = AverageArray(repValue, numRep)
    sdV = PopulationSdArray(repValue, numRep)
    
    result(1) = meanV
    result(2) = meanV + 2 * sdV
    result(3) = meanV - 2 * sdV
End Sub


Private Function AverageArray(ByRef x() As Double, ByVal n As Long) As Double
    Dim i As Long
    Dim s As Double
    
    s = 0
    
    For i = 1 To n
        s = s + x(i)
    Next i
    
    AverageArray = s / n
End Function


Private Function PopulationSdArray(ByRef x() As Double, ByVal n As Long) As Double
    Dim i As Long
    Dim meanX As Double, s As Double
    
    meanX = AverageArray(x, n)
    s = 0
    
    For i = 1 To n
        s = s + (x(i) - meanX) ^ 2
    Next i
    
    PopulationSdArray = Sqr(s / n)
End Function

