Attribute VB_Name = "Module1"
Option Explicit

Sub Run_Basic()
    Call RunAsianHW5(False, False)
End Sub

Sub Run_Bonus1()
    Call RunAsianHW5(True, False)
End Sub

Sub Run_Bonus2()
    Call RunAsianHW5(False, True)
End Sub

Private Sub RunAsianHW5(ByVal DoBonus1 As Boolean, ByVal DoBonus2 As Boolean)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("HW5")
    
    Dim S0 As Double, K As Double, r As Double, q As Double, sigma As Double
    Dim t0 As Double, tau As Double, m As Long, n As Long
    Dim SaveT As Double, numSim As Long, numRep As Long
    
    S0 = ws.Range("C2").value
    K = ws.Range("C3").value
    r = ws.Range("C4").value
    q = ws.Range("C5").value
    sigma = ws.Range("C6").value
    t0 = ws.Range("C7").value
    tau = ws.Range("C8").value
    m = CLng(ws.Range("C9").value)
    n = CLng(ws.Range("C10").value)
    SaveT = ws.Range("C11").value
    numSim = CLng(ws.Range("C12").value)
    numRep = CLng(ws.Range("C13").value)
    
    If DoBonus2 Then
        Dim eu2 As Double, am2 As Double
        Call TreeAsian(S0, K, r, q, sigma, t0, tau, m, n, SaveT, 1, 5, eu2, am2)
        ws.Range("M3").value = eu2
        ws.Range("M4").value = am2
        Exit Sub
    End If
    
    Dim eu As Double, am As Double
    Call TreeAsian(S0, K, r, q, sigma, t0, tau, m, n, SaveT, 1, 1, eu, am)
    
    If Not DoBonus1 Then
        ws.Range("F3").value = eu
        ws.Range("F4").value = am
        
        Dim mcVal As Double, ciU As Double, ciL As Double
        Call MonteCarloAsian(S0, K, r, q, sigma, t0, tau, n, SaveT, numSim, numRep, mcVal, ciU, ciL)
        ws.Range("F5").value = mcVal
        ws.Range("F6").value = ciU
        ws.Range("F7").value = ciL
    Else
        Dim tStart As Double, tEnd As Double
        
        tStart = Timer
        Call TreeAsian(S0, K, r, q, sigma, t0, tau, m, n, SaveT, 1, 1, eu, am)
        tEnd = Timer
        ws.Range("J4").value = eu
        ws.Range("J5").value = am
        ws.Range("J6").value = tEnd - tStart
        
        tStart = Timer
        Call TreeAsian(S0, K, r, q, sigma, t0, tau, m, n, SaveT, 2, 1, eu, am)
        tEnd = Timer
        ws.Range("J8").value = eu
        ws.Range("J9").value = am
        ws.Range("J10").value = tEnd - tStart
        
        tStart = Timer
        Call TreeAsian(S0, K, r, q, sigma, t0, tau, m, n, SaveT, 3, 1, eu, am)
        tEnd = Timer
        ws.Range("J12").value = eu
        ws.Range("J13").value = am
        ws.Range("J14").value = tEnd - tStart
    End If
End Sub

Private Sub TreeAsian(ByVal S0 As Double, ByVal K As Double, ByVal r As Double, ByVal q As Double, _
                      ByVal sigma As Double, ByVal t0 As Double, ByVal tau As Double, _
                      ByVal m As Long, ByVal n As Long, ByVal SaveT As Double, _
                      ByVal SearchMethod As Long, ByVal SampleGap As Long, _
                      ByRef EUPrice As Double, ByRef AMPrice As Double)

    Dim dt As Double, u As Double, d As Double, p As Double, disc As Double
    dt = tau / n
    u = Exp(sigma * Sqr(dt))
    d = 1# / u
    p = (Exp((r - q) * dt) - d) / (u - d)
    disc = Exp(-r * dt)
    
    Dim obs0 As Long
    If SampleGap = 1 Then
        obs0 = CLng(Round(t0 / dt, 0)) + 1
    Else
        obs0 = CLng(Round(t0 / (SampleGap * dt), 0)) + 1
    End If
    
    Dim a() As Double
    Dim CEU() As Double, CAM() As Double
    ReDim a(0 To n, 0 To n, 0 To m)
    ReDim CEU(0 To n, 0 To n, 0 To m)
    ReDim CAM(0 To n, 0 To n, 0 To m)
    
    Dim i As Long, j As Long, kk As Long
    Dim Amax As Double, Amin As Double
    
    For i = 0 To n
        For j = 0 To i
            Call GetAmaxAmin(S0, u, d, i, j, SaveT, obs0, SampleGap, Amax, Amin)
            
            For kk = 0 To m
                If m = 0 Then
                    a(i, j, kk) = Amax
                Else
                    a(i, j, kk) = Amax - (Amax - Amin) * kk / m
                End If
            Next kk
        Next j
    Next i
    
    For j = 0 To n
        For kk = 0 To m
            CEU(n, j, kk) = MaxD(a(n, j, kk) - K, 0)
            CAM(n, j, kk) = CEU(n, j, kk)
        Next kk
    Next j
    
    Dim Au As Double, Ad As Double, Su As Double, sd As Double
    Dim CuEU As Double, CdEU As Double, CuAM As Double, CdAM As Double
    Dim obsNow As Long, doSample As Boolean
    
    For i = n - 1 To 0 Step -1
        For j = 0 To i
            For kk = 0 To m
                Su = S0 * (u ^ (i + 1 - j)) * (d ^ j)
                sd = S0 * (u ^ (i - j)) * (d ^ (j + 1))
                
                doSample = False
                If SampleGap = 1 Then
                    doSample = True
                    obsNow = obs0 + i
                Else
                    If ((i + 1) Mod SampleGap) = 0 Then
                        doSample = True
                        obsNow = obs0 + (i \ SampleGap)
                    End If
                End If
                
                If doSample Then
                    Au = (obsNow * a(i, j, kk) + Su) / (obsNow + 1)
                    Ad = (obsNow * a(i, j, kk) + sd) / (obsNow + 1)
                Else
                    Au = a(i, j, kk)
                    Ad = a(i, j, kk)
                End If
                
                CuEU = InterpValue(Au, a, CEU, i + 1, j, m, SearchMethod)
                CdEU = InterpValue(Ad, a, CEU, i + 1, j + 1, m, SearchMethod)
                CuAM = InterpValue(Au, a, CAM, i + 1, j, m, SearchMethod)
                CdAM = InterpValue(Ad, a, CAM, i + 1, j + 1, m, SearchMethod)
                
                CEU(i, j, kk) = disc * (p * CuEU + (1 - p) * CdEU)
                CAM(i, j, kk) = MaxD(a(i, j, kk) - K, disc * (p * CuAM + (1 - p) * CdAM))
            Next kk
        Next j
    Next i
    
    EUPrice = CEU(0, 0, 0)
    AMPrice = CAM(0, 0, 0)
End Sub

Private Sub GetAmaxAmin(ByVal S0 As Double, ByVal u As Double, ByVal d As Double, _
                        ByVal i As Long, ByVal j As Long, ByVal SaveT As Double, _
                        ByVal obs0 As Long, ByVal SampleGap As Long, _
                        ByRef Amax As Double, ByRef Amin As Double)

    Dim sumMax As Double, sumMin As Double
    Dim stepIdx As Long, upCount As Long, downCount As Long
    Dim stock As Double
    Dim obsFuture As Long
    
    sumMax = obs0 * SaveT
    sumMin = obs0 * SaveT
    obsFuture = 0
    
    For stepIdx = 1 To i
        If SampleGap = 1 Or (stepIdx Mod SampleGap = 0) Then
            If stepIdx <= i - j Then
                upCount = stepIdx
                downCount = 0
            Else
                upCount = i - j
                downCount = stepIdx - (i - j)
            End If
            stock = S0 * (u ^ upCount) * (d ^ downCount)
            sumMax = sumMax + stock
            
            If stepIdx <= j Then
                upCount = 0
                downCount = stepIdx
            Else
                upCount = stepIdx - j
                downCount = j
            End If
            stock = S0 * (u ^ upCount) * (d ^ downCount)
            sumMin = sumMin + stock
            
            obsFuture = obsFuture + 1
        End If
    Next stepIdx
    
    Amax = sumMax / (obs0 + obsFuture)
    Amin = sumMin / (obs0 + obsFuture)
    
    If Amax < Amin Then
        Dim tmp As Double
        tmp = Amax
        Amax = Amin
        Amin = tmp
    End If
End Sub

Private Function InterpValue(ByVal targetA As Double, ByRef a() As Double, ByRef c() As Double, _
                             ByVal i As Long, ByVal j As Long, ByVal m As Long, _
                             ByVal SearchMethod As Long) As Double
    Dim idx As Long
    
    If targetA >= a(i, j, 0) Then
        InterpValue = c(i, j, 0)
        Exit Function
    End If
    
    If targetA <= a(i, j, m) Then
        InterpValue = c(i, j, m)
        Exit Function
    End If
    
    Select Case SearchMethod
        Case 1
            idx = FindSeq(targetA, a, i, j, m)
        Case 2
            idx = FindBin(targetA, a, i, j, m)
        Case 3
            idx = FindLinear(targetA, a, i, j, m)
        Case Else
            idx = FindSeq(targetA, a, i, j, m)
    End Select
    
    Dim Ahigh As Double, Alow As Double, w As Double
    Ahigh = a(i, j, idx - 1)
    Alow = a(i, j, idx)
    
    If Abs(Ahigh - Alow) < 0.000000000001 Then
        InterpValue = c(i, j, idx)
    Else
        w = (Ahigh - targetA) / (Ahigh - Alow)
        InterpValue = w * c(i, j, idx) + (1 - w) * c(i, j, idx - 1)
    End If
End Function

Private Function FindSeq(ByVal targetA As Double, ByRef a() As Double, _
                         ByVal i As Long, ByVal j As Long, ByVal m As Long) As Long
    Dim kk As Long
    For kk = 1 To m
        If targetA >= a(i, j, kk) Then
            FindSeq = kk
            Exit Function
        End If
    Next kk
    FindSeq = m
End Function

Private Function FindBin(ByVal targetA As Double, ByRef a() As Double, _
                         ByVal i As Long, ByVal j As Long, ByVal m As Long) As Long
    Dim L As Long, r As Long, mid As Long
    L = 1
    r = m
    
    Do While L <= r
        mid = (L + r) \ 2
        If targetA <= a(i, j, mid - 1) And targetA >= a(i, j, mid) Then
            FindBin = mid
            Exit Function
        ElseIf targetA < a(i, j, mid) Then
            L = mid + 1
        Else
            r = mid - 1
        End If
    Loop
    
    FindBin = L
    If FindBin < 1 Then FindBin = 1
    If FindBin > m Then FindBin = m
End Function

Private Function FindLinear(ByVal targetA As Double, ByRef a() As Double, _
                            ByVal i As Long, ByVal j As Long, ByVal m As Long) As Long
    Dim Amax As Double, Amin As Double, pos As Long
    
    Amax = a(i, j, 0)
    Amin = a(i, j, m)
    
    If Abs(Amax - Amin) < 0.000000000001 Then
        FindLinear = 1
        Exit Function
    End If
    
    pos = CLng(Int((Amax - targetA) / (Amax - Amin) * m)) + 1
    
    If pos < 1 Then pos = 1
    If pos > m Then pos = m
    
    FindLinear = pos
End Function

Private Sub MonteCarloAsian(ByVal S0 As Double, ByVal K As Double, ByVal r As Double, ByVal q As Double, _
                            ByVal sigma As Double, ByVal t0 As Double, ByVal tau As Double, _
                            ByVal n As Long, ByVal SaveT As Double, _
                            ByVal numSim As Long, ByVal numRep As Long, _
                            ByRef mcVal As Double, ByRef ciU As Double, ByRef ciL As Double)

    Dim dt As Double
    dt = tau / n
    
    Dim obs0 As Long
    obs0 = CLng(Round(t0 / dt, 0)) + 1
    
    Dim repVal() As Double
    ReDim repVal(1 To numRep)
    
    Dim rep As Long, sim As Long, stepIdx As Long
    Dim S As Double, sumA As Double, payoffSum As Double
    Dim z As Double, drift As Double, vol As Double
    
    drift = (r - q - 0.5 * sigma * sigma) * dt
    vol = sigma * Sqr(dt)
    
    Randomize
    
    For rep = 1 To numRep
        payoffSum = 0
        
        For sim = 1 To numSim
            S = S0
            sumA = obs0 * SaveT
            
            For stepIdx = 1 To n
                z = StdNormal()
                S = S * Exp(drift + vol * z)
                sumA = sumA + S
            Next stepIdx
            
            payoffSum = payoffSum + MaxD(sumA / (obs0 + n) - K, 0)
        Next sim
        
        repVal(rep) = Exp(-r * tau) * payoffSum / numSim
    Next rep
    
    Dim avg As Double, sd As Double
    avg = MeanArray(repVal, numRep)
    sd = SdArray(repVal, numRep)
    
    mcVal = avg
    ciU = avg + 1.96 * sd
    ciL = avg - 1.96 * sd
End Sub

Private Function StdNormal() As Double
    Dim u1 As Double, u2 As Double
    u1 = Rnd
    If u1 <= 0 Then u1 = 0.0000000001
    u2 = Rnd
    StdNormal = Sqr(-2# * Log(u1)) * Cos(2# * WorksheetFunction.PI() * u2)
End Function

Private Function MeanArray(ByRef X() As Double, ByVal n As Long) As Double
    Dim i As Long, S As Double
    For i = 1 To n
        S = S + X(i)
    Next i
    MeanArray = S / n
End Function

Private Function SdArray(ByRef X() As Double, ByVal n As Long) As Double
    Dim i As Long, avg As Double, ss As Double
    avg = MeanArray(X, n)
    
    For i = 1 To n
        ss = ss + (X(i) - avg) ^ 2
    Next i
    
    If n > 1 Then
        SdArray = Sqr(ss / (n - 1))
    Else
        SdArray = 0
    End If
End Function

Private Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then
        MaxD = a
    Else
        MaxD = b
    End If
End Function

