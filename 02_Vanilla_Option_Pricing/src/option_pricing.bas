Attribute VB_Name = "Module1"
Option Explicit

'============================================================
' Inputs  : B2:B10 = S0, K, r, q, sigma, T, num_sim, num_rep, n
' Basic   : outputs to E2:E13
' Bonus 1 : outputs to G2:G5
' Bonus 2 : 8 charts from A13 downward
' Bonus 3 : outputs to I2:I3
'============================================================

Private Type HW2Inputs
    S0 As Double
    K As Double
    r As Double
    q As Double
    sigma As Double
    T As Double
    num_sim As Long
    num_rep As Long
    n As Long
End Type

'========================
' Basic math utilities
'========================
Private Function Ncdf(ByVal x As Double) As Double
    Ncdf = Application.WorksheetFunction.Norm_S_Dist(x, True)
End Function

Private Function Ninv(ByVal p As Double) As Double
    If p <= 0# Then p = 0.0000000001
    If p >= 1# Then p = 0.9999999999
    Ninv = Application.WorksheetFunction.Norm_S_Inv(p)
End Function

Private Function Max2(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then Max2 = a Else Max2 = b
End Function

Private Function Payoff(ByVal S As Double, ByVal K As Double, ByVal isCall As Boolean) As Double
    If isCall Then
        Payoff = Max2(S - K, 0#)
    Else
        Payoff = Max2(K - S, 0#)
    End If
End Function

Private Function BSPrice(ByVal S0 As Double, ByVal K As Double, ByVal r As Double, _
                         ByVal q As Double, ByVal sigma As Double, ByVal T As Double, _
                         ByVal isCall As Boolean) As Double
    Dim d1 As Double, d2 As Double

    If T <= 0# Or sigma <= 0# Then
        BSPrice = Payoff(S0, K, isCall)
        Exit Function
    End If

    d1 = (Log(S0 / K) + (r - q + 0.5 * sigma * sigma) * T) / (sigma * Sqr(T))
    d2 = d1 - sigma * Sqr(T)

    If isCall Then
        BSPrice = S0 * Exp(-q * T) * Ncdf(d1) - K * Exp(-r * T) * Ncdf(d2)
    Else
        BSPrice = K * Exp(-r * T) * Ncdf(-d2) - S0 * Exp(-q * T) * Ncdf(-d1)
    End If
End Function

'========================
' Input / output helpers
'========================
Private Function ReadInputs() As HW2Inputs
    With ReadInputs
        .S0 = CDbl(Range("B2").Value)
        .K = CDbl(Range("B3").Value)
        .r = CDbl(Range("B4").Value)
        .q = CDbl(Range("B5").Value)
        .sigma = CDbl(Range("B6").Value)
        .T = CDbl(Range("B7").Value)
        .num_sim = CLng(Range("B8").Value)
        .num_rep = CLng(Range("B9").Value)
        .n = CLng(Range("B10").Value)
    End With
End Function

Private Function InputsValid(ByRef x As HW2Inputs) As Boolean
    If x.S0 <= 0# Or x.K <= 0# Or x.sigma <= 0# Or x.T <= 0# Then
        MsgBox "S0, K, sigma, and T must be positive.", vbExclamation
        InputsValid = False
        Exit Function
    End If
    If x.num_sim <= 1 Or x.num_rep <= 1 Or x.n <= 0 Then
        MsgBox "num_sim and num_rep must be larger than 1; n must be positive.", vbExclamation
        InputsValid = False
        Exit Function
    End If
    InputsValid = True
End Function

Private Sub ClearBasicOutputs()
    Range("E2:E13").ClearContents
End Sub

Private Sub ClearBonus1Outputs()
    Range("G2:G5").ClearContents
End Sub

Private Sub ClearBonus2Outputs()
    Dim co As ChartObject
    Range("K:AF").ClearContents
    For Each co In ActiveSheet.ChartObjects
        If Left(co.Name, 8) = "B2Chart_" Then co.Delete
    Next co
End Sub

Private Sub ClearBonus3Outputs()
    Range("I2:I3").ClearContents
End Sub

Private Sub draw_chart(xArr, y1Arr, y2Arr, title As String, leftPos As Double)

    Dim cht As ChartObject
    Set cht = ActiveSheet.ChartObjects.Add(Left:=leftPos, Width:=300, Top:=400, Height:=200)
    
    With cht.Chart
        .ChartType = xlLine
        
        .SeriesCollection.NewSeries
        .SeriesCollection(1).XValues = xArr
        .SeriesCollection(1).Values = y1Arr
        .SeriesCollection(1).Name = "CRR"
        
        .SeriesCollection.NewSeries
        .SeriesCollection(2).XValues = xArr
        .SeriesCollection(2).Values = y2Arr
        .SeriesCollection(2).Name = "BBS"
        
        .HasTitle = True
        .chartTitle.Text = title
        
    End With

End Sub

'========================
' Monte Carlo One Reptition
'========================
Private Function MCOneRep(ByRef x As HW2Inputs, ByVal isCall As Boolean) As Double
    Dim i As Long, z As Double, ST As Double, sumPayoff As Double
    Dim drift As Double

    drift = Log(x.S0) + (x.r - x.q - 0.5 * x.sigma * x.sigma) * x.T

    For i = 1 To x.num_sim
        z = Ninv(Rnd())
        ST = Exp(drift + x.sigma * Sqr(x.T) * z)
        sumPayoff = sumPayoff + Payoff(ST, x.K, isCall)
    Next i

    MCOneRep = Exp(-x.r * x.T) * sumPayoff / x.num_sim
End Function

Private Sub MCMeanCI(ByRef x As HW2Inputs, ByVal isCall As Boolean, _
                     ByRef meanOut As Double, ByRef ciLow As Double, ByRef ciUp As Double)
    Dim rep As Long, est As Double
    Dim s1 As Double, s2 As Double, sd As Double

    For rep = 1 To x.num_rep
        est = MCOneRep(x, isCall)
        s1 = s1 + est
        s2 = s2 + est * est
    Next rep

    meanOut = s1 / x.num_rep
    sd = Sqr((s2 - x.num_rep * meanOut * meanOut) / (x.num_rep - 1))
    ciLow = meanOut - 1.96 * sd
    ciUp = meanOut + 1.96 * sd
End Sub

'========================
' Basic CRR: 2-dimensional arrays
'========================
Private Function CRR2D(ByRef x As HW2Inputs, ByVal isCall As Boolean, ByVal isAmerican As Boolean) As Double
    Dim n As Long, i As Long, j As Long
    Dim dt As Double, u As Double, d As Double, p As Double, disc As Double
    Dim S() As Double, V() As Double, cont As Double, exv As Double

    n = x.n
    dt = x.T / n
    u = Exp(x.sigma * Sqr(dt))
    d = 1# / u
    p = (Exp((x.r - x.q) * dt) - d) / (u - d)
    disc = Exp(-x.r * dt)

    If p < 0# Or p > 1# Then
        MsgBox "Risk-neutral probability p is outside [0,1]. Try a larger n.", vbExclamation
    End If

    ReDim S(0 To n, 0 To n)
    ReDim V(0 To n, 0 To n)

    For i = 0 To n
        For j = 0 To i
            S(i, j) = x.S0 * (u ^ (i - j)) * (d ^ j)
        Next j
    Next i

    For j = 0 To n
        V(n, j) = Payoff(S(n, j), x.K, isCall)
    Next j

    For i = n - 1 To 0 Step -1
        For j = 0 To i
            cont = disc * (p * V(i + 1, j) + (1# - p) * V(i + 1, j + 1))
            If isAmerican Then
                exv = Payoff(S(i, j), x.K, isCall)
                V(i, j) = Max2(cont, exv)
            Else
                V(i, j) = cont
            End If
        Next j
    Next i

    CRR2D = V(0, 0)
End Function

'========================
' Bonus 1 CRR: one column vector
'========================
Private Function CRR1D(ByRef x As HW2Inputs, ByVal isCall As Boolean, ByVal isAmerican As Boolean, _
                       Optional ByVal nOverride As Long = 0) As Double
    Dim n As Long, i As Long, j As Long
    Dim dt As Double, u As Double, d As Double, p As Double, disc As Double
    Dim V() As Double, Snode As Double, cont As Double, exv As Double

    If nOverride > 0 Then n = nOverride Else n = x.n
    dt = x.T / n
    u = Exp(x.sigma * Sqr(dt))
    d = 1# / u
    p = (Exp((x.r - x.q) * dt) - d) / (u - d)
    disc = Exp(-x.r * dt)

    ReDim V(0 To n)

    For j = 0 To n
        Snode = x.S0 * (u ^ (n - j)) * (d ^ j)
        V(j) = Payoff(Snode, x.K, isCall)
    Next j

    For i = n - 1 To 0 Step -1
        For j = 0 To i
            cont = disc * (p * V(j) + (1# - p) * V(j + 1))
            If isAmerican Then
                Snode = x.S0 * (u ^ (i - j)) * (d ^ j)
                exv = Payoff(Snode, x.K, isCall)
                V(j) = Max2(cont, exv)
            Else
                V(j) = cont
            End If
        Next j
    Next i

    CRR1D = V(0)
End Function

'========================
' Bonus 2 BBS tree: one column vector
' At i = n - 1, use Black-Scholes value for remaining dt.
' For American options, take max(BS value, immediate exercise value).
'========================
Private Function BBS1D(ByRef x As HW2Inputs, ByVal isCall As Boolean, ByVal isAmerican As Boolean, _
                       ByVal nOverride As Long) As Double
    Dim n As Long, i As Long, j As Long
    Dim dt As Double, u As Double, d As Double, p As Double, disc As Double
    Dim V() As Double, Snode As Double, cont As Double, exv As Double, bs As Double

    n = nOverride
    dt = x.T / n
    u = Exp(x.sigma * Sqr(dt))
    d = 1# / u
    p = (Exp((x.r - x.q) * dt) - d) / (u - d)
    disc = Exp(-x.r * dt)

    ReDim V(0 To n - 1)

    'Penultimate layer: i = n - 1
    For j = 0 To n - 1
        Snode = x.S0 * (u ^ (n - 1 - j)) * (d ^ j)
        bs = BSPrice(Snode, x.K, x.r, x.q, x.sigma, dt, isCall)
        If isAmerican Then
            V(j) = Max2(bs, Payoff(Snode, x.K, isCall))
        Else
            V(j) = bs
        End If
    Next j

    'Backward induction from i = n - 2 to 0
    For i = n - 2 To 0 Step -1
        For j = 0 To i
            cont = disc * (p * V(j) + (1# - p) * V(j + 1))
            If isAmerican Then
                Snode = x.S0 * (u ^ (i - j)) * (d ^ j)
                exv = Payoff(Snode, x.K, isCall)
                V(j) = Max2(cont, exv)
            Else
                V(j) = cont
            End If
        Next j
    Next i

    BBS1D = V(0)
End Function

'========================
' Bonus 3: combinatorial method for European calls and puts
' log[C(n,j) * p^(n-j) * (1-p)^j] is used to avoid overflow.
'========================
Private Function CombinatorialEuropean(ByRef x As HW2Inputs, ByVal isCall As Boolean) As Double
    Dim n As Long, j As Long
    Dim dt As Double, u As Double, d As Double, p As Double
    Dim logProb As Double, prob As Double, Snode As Double
    Dim priceSum As Double, payoffValue As Double

    n = x.n
    dt = x.T / n
    u = Exp(x.sigma * Sqr(dt))
    d = 1# / u
    p = (Exp((x.r - x.q) * dt) - d) / (u - d)

    If p <= 0# Or p >= 1# Then
        MsgBox "Combinatorial method requires p inside (0,1). Try a larger n.", vbExclamation
        CombinatorialEuropean = CVErr(xlErrNum)
        Exit Function
    End If

    For j = 0 To n
        logProb = Application.WorksheetFunction.GammaLn(n + 1) _
                - Application.WorksheetFunction.GammaLn(j + 1) _
                - Application.WorksheetFunction.GammaLn(n - j + 1) _
                + (n - j) * Log(p) + j * Log(1# - p)

        prob = Exp(logProb)
        Snode = x.S0 * Exp((n - j) * Log(u) + j * Log(d))
        payoffValue = Payoff(Snode, x.K, isCall)
        priceSum = priceSum + prob * payoffValue
    Next j

    CombinatorialEuropean = Exp(-x.r * x.T) * priceSum
End Function

'========================
' Public macros: one button per part
'========================
Public Sub Run_Basic()
    Dim x As HW2Inputs
    Dim mcC As Double, cLow As Double, cUp As Double
    Dim mcP As Double, pLow As Double, pUp As Double

    x = ReadInputs()
    If Not InputsValid(x) Then Exit Sub
    Randomize
    ClearBasicOutputs

    Range("E2").Value = BSPrice(x.S0, x.K, x.r, x.q, x.sigma, x.T, True)
    Range("E3").Value = BSPrice(x.S0, x.K, x.r, x.q, x.sigma, x.T, False)

    Call MCMeanCI(x, True, mcC, cLow, cUp)
    Call MCMeanCI(x, False, mcP, pLow, pUp)
    Range("E4").Value = mcC
    Range("E5").Value = mcP
    Range("E6").Value = cUp
    Range("E7").Value = cLow
    Range("E8").Value = pUp
    Range("E9").Value = pLow

    Range("E10").Value = CRR2D(x, True, False)
    Range("E11").Value = CRR2D(x, False, False)
    Range("E12").Value = CRR2D(x, True, True)
    Range("E13").Value = CRR2D(x, False, True)
End Sub

Public Sub Run_Bonus1_ColumnVector()
    Dim x As HW2Inputs
    x = ReadInputs()
    If Not InputsValid(x) Then Exit Sub
    ClearBonus1Outputs

    Range("G2").Value = CRR1D(x, True, False)
    Range("G3").Value = CRR1D(x, False, False)
    Range("G4").Value = CRR1D(x, True, True)
    Range("G5").Value = CRR1D(x, False, True)
End Sub

Public Sub Run_Bonus3_Combinatorial()
    Dim x As HW2Inputs
    x = ReadInputs()
    If Not InputsValid(x) Then Exit Sub
    ClearBonus3Outputs

    Range("I2").Value = CombinatorialEuropean(x, True)
    Range("I3").Value = CombinatorialEuropean(x, False)
End Sub

Public Sub run_bonus2()

    Dim x As HW2Inputs
    x = ReadInputs()
    If Not InputsValid(x) Then Exit Sub

    ' === 額外輸出：用輸入的 n 計算 BBS 的 V(0) ===
    Range("K2").Value = "BBS_E_c_V0"
    Range("L2").Value = BBS1D(x, True, False, x.n)
    
    Range("K3").Value = "BBS_E_p_V0"
    Range("L3").Value = BBS1D(x, False, False, x.n)
    
    Range("K4").Value = "BBS_A_c_V0"
    Range("L4").Value = BBS1D(x, True, True, x.n)
    
    Range("K5").Value = "BBS_A_p_V0"
    Range("L5").Value = BBS1D(x, False, True, x.n)

    Dim n As Long, idx As Long
    
    Dim nList() As Double
    Dim CRR_E_c() As Double, BBS_E_c() As Double
    Dim CRR_E_p() As Double, BBS_E_p() As Double
    Dim CRR_A_c() As Double, BBS_A_c() As Double
    Dim CRR_A_p() As Double, BBS_A_p() As Double
    
    ' === 建立 n list ===
    Dim countN As Long
    countN = (500 - 105) / 5 + 1
    
    ReDim nList(1 To countN)
    ReDim CRR_E_c(1 To countN)
    ReDim BBS_E_c(1 To countN)
    ReDim CRR_E_p(1 To countN)
    ReDim BBS_E_p(1 To countN)
    ReDim CRR_A_c(1 To countN)
    ReDim BBS_A_c(1 To countN)
    ReDim CRR_A_p(1 To countN)
    ReDim BBS_A_p(1 To countN)
    
    idx = 1
    
    For n = 105 To 500 Step 5
        
        nList(idx) = n
        
        ' === CRR（用 Bonus1 方法）===
        CRR_E_c(idx) = CRR1D(x, True, False, n)
        CRR_E_p(idx) = CRR1D(x, False, False, n)
        CRR_A_c(idx) = CRR1D(x, True, True, n)
        CRR_A_p(idx) = CRR1D(x, False, True, n)
        
        ' === BBS ===
        BBS_E_c(idx) = BBS1D(x, True, False, n)
        BBS_E_p(idx) = BBS1D(x, False, False, n)
        BBS_A_c(idx) = BBS1D(x, True, True, n)
        BBS_A_p(idx) = BBS1D(x, False, True, n)
        
        idx = idx + 1
        
    Next n
    
    ' === 清除舊圖 ===
    Dim co As ChartObject
    For Each co In ActiveSheet.ChartObjects
        co.Delete
    Next co
    
    ' === 畫圖（8張）===
    Call draw_chart(nList, CRR_E_c, BBS_E_c, "European Call", 50)
    Call draw_chart(nList, CRR_E_p, BBS_E_p, "European Put", 350)
    Call draw_chart(nList, CRR_A_c, BBS_A_c, "American Call", 650)
    Call draw_chart(nList, CRR_A_p, BBS_A_p, "American Put", 950)

End Sub

Private Sub AddOneLineChart(ByVal chartName As String, ByVal chartTitle As String, _
                            ByVal xRange As Range, ByVal yRange As Range, _
                            ByVal leftPos As Double, ByVal topPos As Double)
    Dim co As ChartObject
    Set co = ActiveSheet.ChartObjects.Add(Left:=leftPos, Top:=topPos, Width:=400, Height:=220)
    co.Name = chartName

    With co.Chart
        .ChartType = xlLine
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = chartTitle
        .SeriesCollection(1).XValues = xRange
        .SeriesCollection(1).Values = yRange
        .HasTitle = True
        .chartTitle.Text = chartTitle
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "n"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Option value"
        .HasLegend = False
    End With
End Sub

