Attribute VB_Name = "Module1"
Option Explicit

Private Const TOL As Double = 0.000000001

'========================
' Buttons
'========================

Public Sub Run_Basic()
    Dim ws As Worksheet: Set ws = ActiveSheet
    
    Dim S0 As Double, r As Double, q As Double, sigma As Double, T As Double
    Dim Smax0 As Double, n As Long, numSim As Long, numRep As Long
    
    ReadInputs ws, S0, r, q, sigma, T, Smax0, n, numSim, numRep
    
    Dim euroBT As Double, amerBT As Double
    Dim mcMean As Double, ciLow As Double, ciHigh As Double
    
    PriceBasicTree_LevelBased S0, r, q, sigma, T, Smax0, n, euroBT, amerBT
    PriceMonteCarlo S0, r, q, sigma, T, Smax0, n, numSim, numRep, mcMean, ciLow, ciHigh
    
    ws.Range("E2").value = euroBT
    ws.Range("E3").value = amerBT
    ws.Range("E4").value = mcMean
    ws.Range("E5").value = ciLow
    ws.Range("E6").value = ciHigh
End Sub

Public Sub Run_Bonus1()
    Dim ws As Worksheet: Set ws = ActiveSheet
    
    Dim S0 As Double, r As Double, q As Double, sigma As Double, T As Double
    Dim Smax0 As Double, n As Long, numSim As Long, numRep As Long
    
    ReadInputs ws, S0, r, q, sigma, T, Smax0, n, numSim, numRep
    
    Dim euroBT As Double, amerBT As Double
    PriceBonus1Tree S0, r, q, sigma, T, Smax0, n, euroBT, amerBT
    
    ws.Range("H2").value = euroBT
    ws.Range("H3").value = amerBT
End Sub

Public Sub Run_Bonus2()
    Dim ws As Worksheet: Set ws = ActiveSheet
    
    Dim S0 As Double, r As Double, q As Double, sigma As Double, T As Double
    Dim Smax0 As Double, n As Long, numSim As Long, numRep As Long
    
    ReadInputs ws, S0, r, q, sigma, T, Smax0, n, numSim, numRep
    
    Dim euroCV As Double, amerCV As Double
    PriceBonus2CV S0, r, q, sigma, T, n, euroCV, amerCV
    
    ws.Range("K2").value = euroCV
    ws.Range("K3").value = amerCV
End Sub

'========================
' Inputs
'========================

Private Sub ReadInputs(ws As Worksheet, _
    ByRef S0 As Double, ByRef r As Double, ByRef q As Double, ByRef sigma As Double, _
    ByRef T As Double, ByRef Smax0 As Double, ByRef n As Long, _
    ByRef numSim As Long, ByRef numRep As Long)

    S0 = CDbl(ws.Range("B1").value)
    r = CDbl(ws.Range("B2").value)
    q = CDbl(ws.Range("B3").value)
    sigma = CDbl(ws.Range("B4").value)
    T = CDbl(ws.Range("B5").value)
    Smax0 = CDbl(ws.Range("B6").value)
    n = CLng(ws.Range("B7").value)
    numSim = CLng(ws.Range("B8").value)
    numRep = CLng(ws.Range("B9").value)
    
    If Smax0 < S0 Then
        MsgBox "Smax0 cannot be smaller than S0."
        Err.Raise vbObjectError + 999
    End If
End Sub

'========================
' Basic: level/index-based Smax storage
'========================

Private Sub PriceBasicTree_LevelBased(S0 As Double, r As Double, q As Double, sigma As Double, T As Double, _
    Smax0 As Double, n As Long, ByRef euroPrice As Double, ByRef amerPrice As Double)

    Dim dt As Double, u As Double, d As Double, mu As Double, p As Double, disc As Double
    dt = T / n
    u = Exp(sigma * Sqr(dt))
    d = 1 / u
    mu = Exp((r - q) * dt)
    p = (mu - d) / (u - d)
    disc = Exp(-r * dt)
    
    Dim offset As Long
    offset = n
    
    Dim maxK As Long
    maxK = 2 * n + 1
    
    ' 先建立所有可能股價(ST)的node
    Dim StockLevel() As Double
    ReDim StockLevel(0 To maxK - 1)
    
    Dim h As Long, k As Long
    For k = 0 To maxK - 1
        h = k - offset
        StockLevel(k) = S0 * (u ^ h)
    Next k
    
    Dim euro() As Double, amer() As Double
    ReDim euro(0 To n, 0 To n, 0 To maxK - 1)
    ReDim amer(0 To n, 0 To n, 0 To maxK - 1)
    
    Dim i As Long, j As Long
    Dim level As Long
    Dim S As Double, Smax As Double
    
    ' Terminal payoff：先算完最後一期
    For j = 0 To n
        level = 2 * j - n
        S = StockLevel(level + offset)
        
        For k = 0 To maxK - 1
            h = k - offset
            Smax = WorksheetFunction.Max(Smax0, StockLevel(k))
            
            euro(n, j, k) = WorksheetFunction.Max(Smax - S, 0)
            amer(n, j, k) = euro(n, j, k)
        Next k
    Next j
    
    Dim curLevel As Long, upLevel As Long, downLevel As Long
    Dim upK As Long, downK As Long
    Dim curS As Double, curSmax As Double
    Dim contE As Double, contA As Double, exercise As Double
    
    ' Backward induction
    For i = n - 1 To 0 Step -1
        For j = 0 To i
            curLevel = 2 * j - i
            curS = StockLevel(curLevel + offset)
            
            For k = 0 To maxK - 1
                h = k - offset
                curSmax = WorksheetFunction.Max(Smax0, StockLevel(k))
                
                upLevel = curLevel + 1
                downLevel = curLevel - 1
                
                ' up 後的 Smax level
                If StockLevel(upLevel + offset) > curSmax + TOL Then
                    upK = upLevel + offset
                Else
                    upK = k
                End If
                
                ' down 後的 Smax level
                If StockLevel(downLevel + offset) > curSmax + TOL Then
                    downK = downLevel + offset
                Else
                    downK = k
                End If
                
                contE = disc * (p * euro(i + 1, j + 1, upK) + (1 - p) * euro(i + 1, j, downK))
                contA = disc * (p * amer(i + 1, j + 1, upK) + (1 - p) * amer(i + 1, j, downK))
                exercise = WorksheetFunction.Max(curSmax - curS, 0)
                
                euro(i, j, k) = contE
                amer(i, j, k) = WorksheetFunction.Max(exercise, contA)
            Next k
        Next j
    Next i
    
    euroPrice = euro(0, 0, offset)
    amerPrice = amer(0, 0, offset)
End Sub

'========================
' Bonus 1
'========================

Private Sub PriceBonus1Tree(S0 As Double, r As Double, q As Double, sigma As Double, T As Double, _
    Smax0 As Double, n As Long, ByRef euroPrice As Double, ByRef amerPrice As Double)

    Dim dt As Double, u As Double, d As Double, mu As Double, p As Double, disc As Double
    dt = T / n
    u = Exp(sigma * Sqr(dt))
    d = 1 / u
    mu = Exp((r - q) * dt)
    p = (mu - d) / (u - d)
    disc = Exp(-r * dt)
    
    Dim maxK As Long
    maxK = n + 2
    
    Dim S() As Double, SmaxList() As Double
    Dim cnt() As Long
    Dim euro() As Double, amer() As Double
    
    ReDim S(0 To n, 0 To n)
    ReDim SmaxList(0 To n, 0 To n, 1 To maxK)
    ReDim cnt(0 To n, 0 To n)
    ReDim euro(0 To n, 0 To n, 1 To maxK)
    ReDim amer(0 To n, 0 To n, 1 To maxK)
    
    Dim i As Long, j As Long, k As Long
    
    For i = 0 To n
        For j = 0 To i
            S(i, j) = S0 * (u ^ j) * (d ^ (i - j))
            BuildDirectSmaxList S0, Smax0, u, i, j, SmaxList, cnt
        Next j
    Next i
    
    For j = 0 To n
        For k = 1 To cnt(n, j)
            euro(n, j, k) = WorksheetFunction.Max(SmaxList(n, j, k) - S(n, j), 0)
            amer(n, j, k) = euro(n, j, k)
        Next k
    Next j
    
    Dim curMax As Double, upMax As Double, downMax As Double
    Dim upIdx As Long, downIdx As Long
    Dim contE As Double, contA As Double, exercise As Double
    
    For i = n - 1 To 0 Step -1
        For j = 0 To i
            For k = 1 To cnt(i, j)
                curMax = SmaxList(i, j, k)
                
                upMax = WorksheetFunction.Max(curMax, S(i + 1, j + 1))
                downMax = WorksheetFunction.Max(curMax, S(i + 1, j))
                
                upIdx = FindSmaxIndex(SmaxList, cnt, i + 1, j + 1, upMax)
                downIdx = FindSmaxIndex(SmaxList, cnt, i + 1, j, downMax)
                
                contE = disc * (p * euro(i + 1, j + 1, upIdx) + (1 - p) * euro(i + 1, j, downIdx))
                contA = disc * (p * amer(i + 1, j + 1, upIdx) + (1 - p) * amer(i + 1, j, downIdx))
                exercise = WorksheetFunction.Max(curMax - S(i, j), 0)
                
                euro(i, j, k) = contE
                amer(i, j, k) = WorksheetFunction.Max(exercise, contA)
            Next k
        Next j
    Next i
    
    euroPrice = euro(0, 0, 1)
    amerPrice = amer(0, 0, 1)
End Sub

Private Sub BuildDirectSmaxList(S0 As Double, Smax0 As Double, u As Double, _
    i As Long, j As Long, ByRef SmaxList() As Double, ByRef cnt() As Long)

    Dim exponentNow As Long
    exponentNow = 2 * j - i
    
    Dim hMin As Long, hMax As Long, h As Long
    hMin = WorksheetFunction.Max(0, exponentNow)
    hMax = j
    
    If i = 0 Then
        AddUniqueSmax SmaxList, cnt, i, j, Smax0
        Exit Sub
    End If
    
    For h = hMin To hMax
        AddUniqueSmax SmaxList, cnt, i, j, WorksheetFunction.Max(Smax0, S0 * (u ^ h))
    Next h
End Sub

'========================
' Bonus 2: Cheuk and Vorst
'========================

Private Sub PriceBonus2CV(S0 As Double, r As Double, q As Double, sigma As Double, T As Double, _
    n As Long, ByRef euroPrice As Double, ByRef amerPrice As Double)

    Dim dt As Double, u As Double, d As Double, mu As Double
    Dim phat As Double, qhat As Double, discQ As Double
    
    dt = T / n
    u = Exp(sigma * Sqr(dt))
    d = 1 / u
    mu = Exp((r - q) * dt)
    
    phat = (1 - mu * d) / (mu * (u - d))
    qhat = (mu * u - 1) / (mu * (u - d))
    
    discQ = Exp(-q * dt)
    
    Dim euro() As Double, amer() As Double
    ReDim euro(0 To n, 0 To n + 1)
    ReDim amer(0 To n, 0 To n + 1)
    
    Dim i As Long, j As Long
    
    For j = 0 To n
        euro(n, j) = u ^ j - 1
        amer(n, j) = euro(n, j)
    Next j
    
    Dim upJ As Long, downJ As Long
    Dim contE As Double, contA As Double, exercise As Double
    
    For i = n - 1 To 0 Step -1
        For j = 0 To i
            If j = 0 Then
                upJ = 0
            Else
                upJ = j - 1
            End If
            
            downJ = j + 1
            
            contE = discQ * (qhat * euro(i + 1, upJ) + phat * euro(i + 1, downJ))
            contA = discQ * (qhat * amer(i + 1, upJ) + phat * amer(i + 1, downJ))
            exercise = u ^ j - 1
            
            euro(i, j) = contE
            amer(i, j) = WorksheetFunction.Max(exercise, contA)
        Next j
    Next i
    
    euroPrice = S0 * euro(0, 0)
    amerPrice = S0 * amer(0, 0)
End Sub

'========================
' Monte Carlo
'========================

Private Sub PriceMonteCarlo(S0 As Double, r As Double, q As Double, sigma As Double, T As Double, _
    Smax0 As Double, n As Long, numSim As Long, numRep As Long, _
    ByRef mcMean As Double, ByRef ciLow As Double, ByRef ciHigh As Double)

    Dim repValue() As Double
    ReDim repValue(1 To numRep)
    
    Dim dt As Double
    dt = T / n
    
    Dim drift As Double
    drift = (r - q - 0.5 * sigma * sigma) * dt
    
    Dim vol As Double
    vol = sigma * Sqr(dt)
    
    Dim rep As Long, sim As Long, stepN As Long
    Dim S As Double, Smax As Double, z As Double, payoffSum As Double
    
    Randomize
    
    For rep = 1 To numRep
        payoffSum = 0
        
        For sim = 1 To numSim
            S = S0
            Smax = Smax0
            
            For stepN = 1 To n
                z = RandNormal()
                S = S * Exp(drift + vol * z)
                If S > Smax Then Smax = S
            Next stepN
            
            payoffSum = payoffSum + WorksheetFunction.Max(Smax - S, 0)
        Next sim
        
        repValue(rep) = Exp(-r * T) * payoffSum / numSim
    Next rep
    
    Dim avg As Double, sd As Double
    
    For rep = 1 To numRep
        avg = avg + repValue(rep)
    Next rep
    avg = avg / numRep
    
    For rep = 1 To numRep
        sd = sd + (repValue(rep) - avg) ^ 2
    Next rep
    sd = Sqr(sd / numRep)
    
    mcMean = avg
    ciLow = avg - 2 * sd
    ciHigh = avg + 2 * sd
End Sub

Private Function RandNormal() As Double
    Dim u1 As Double, u2 As Double
    
    u1 = Rnd
    If u1 < 0.0000000001 Then u1 = 0.0000000001
    
    u2 = Rnd
    
    RandNormal = Sqr(-2 * Log(u1)) * Cos(2 * WorksheetFunction.Pi() * u2)
End Function

'========================
' Smax list helpers for Bonus1
'========================

Private Sub AddUniqueSmax(ByRef SmaxList() As Double, ByRef cnt() As Long, _
    i As Long, j As Long, value As Double)

    Dim k As Long
    
    For k = 1 To cnt(i, j)
        If Abs(SmaxList(i, j, k) - value) < TOL Then Exit Sub
    Next k
    
    cnt(i, j) = cnt(i, j) + 1
    SmaxList(i, j, cnt(i, j)) = value
End Sub

Private Function FindSmaxIndex(ByRef SmaxList() As Double, ByRef cnt() As Long, _
    i As Long, j As Long, value As Double) As Long

    Dim k As Long
    
    For k = 1 To cnt(i, j)
        If Abs(SmaxList(i, j, k) - value) < TOL Then
            FindSmaxIndex = k
            Exit Function
        End If
    Next k
    
    Err.Raise vbObjectError + 1001, , "Cannot find Smax in list."
End Function



