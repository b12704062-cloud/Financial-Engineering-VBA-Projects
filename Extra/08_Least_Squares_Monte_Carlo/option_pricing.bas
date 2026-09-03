Attribute VB_Name = "Module1"
Option Explicit

Sub Run_PlainVanillaPut_LSM()
    RunLSM "PLAIN"
End Sub

Sub Run_LookbackPut_LSM()
    RunLSM "LOOKBACK"
End Sub

Sub Run_ArithmeticAverageCall_LSM()
    RunLSM "AVG"
End Sub

Private Sub RunLSM(ByVal optKind As String)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim numSim As Long, numRep As Long
    numSim = CLng(ReadByLabel(ws, "num_sim", 10000))
    numRep = CLng(ReadByLabel(ws, "num_rep", 20))
    
    Dim repVal() As Double
    ReDim repVal(1 To numRep)
    
    Dim i As Long
    Randomize
    
    For i = 1 To numRep
        Select Case UCase(optKind)
            Case "PLAIN"
                repVal(i) = PricePlainPutOneRep(ws, numSim)
            Case "LOOKBACK"
                repVal(i) = PriceLookbackPutOneRep(ws, numSim)
            Case "AVG"
                repVal(i) = PriceAvgCallOneRep(ws, numSim)
        End Select
    Next i
    
    Dim meanVal As Double, sdVal As Double
    meanVal = MeanArr(repVal)
    sdVal = StdevArr(repVal)
    
    WriteOutput ws, optKind, meanVal, meanVal - 1.96 * sdVal, meanVal + 1.96 * sdVal, sdVal
End Sub

Private Function PricePlainPutOneRep(ws As Worksheet, ByVal numSim As Long) As Double
    Dim S0 As Double, K As Double, r As Double, q As Double, sig As Double, t As Double
    Dim n As Long, dt As Double
    Dim rngInput As Range
    Set rngInput = ws.Range("B:C")

    S0 = ReadByLabelInRange(rngInput, "S0")
    K = ReadByLabelInRange(rngInput, "K")
    r = ReadByLabelInRange(rngInput, "r")
    q = ReadByLabelInRange(rngInput, "q")
    sig = ReadByLabelInRange(rngInput, "sigma")
    t = ReadByLabelInRange(rngInput, "T")
    n = CLng(ReadByLabelInRange(rngInput, "n"))
    
    dt = t / n
    
    Dim S() As Double
    ReDim S(1 To numSim, 0 To n)
    
    Dim i As Long, j As Long, z As Double
    For i = 1 To numSim
        S(i, 0) = S0
        For j = 1 To n
            z = SafeNormInv()
            S(i, j) = S(i, j - 1) * Exp((r - q - 0.5 * sig ^ 2) * dt + sig * Sqr(dt) * z)
        Next j
    Next i
    
    Dim cf() As Double, exTime() As Long
    ReDim cf(1 To numSim)
    ReDim exTime(1 To numSim)
    
    For i = 1 To numSim
        cf(i) = MaxD(K - S(i, n), 0)
        exTime(i) = n
    Next i
    
    Dim ev As Double, hv As Double, pred As Double
    Dim X() As Double, Y() As Double, coef() As Double
    Dim m As Long, p As Long
    p = 3
    
    For j = n - 1 To 1 Step -1
        m = CountITMPlain(S, K, numSim, j)
        If m >= p Then
            ReDim X(1 To m, 1 To p)
            ReDim Y(1 To m)
            
            m = 0
            For i = 1 To numSim
                ev = MaxD(K - S(i, j), 0)
                If ev > 0 Then
                    m = m + 1
                    hv = cf(i) * Exp(-r * dt * (exTime(i) - j))
                    X(m, 1) = 1
                    X(m, 2) = S(i, j)
                    X(m, 3) = S(i, j) ^ 2
                    Y(m) = hv
                End If
            Next i
            
            coef = OLSCoef(X, Y, p, m)
            
            For i = 1 To numSim
                ev = MaxD(K - S(i, j), 0)
                If ev > 0 Then
                    pred = coef(1) + coef(2) * S(i, j) + coef(3) * S(i, j) ^ 2
                    If ev > pred Then
                        cf(i) = ev
                        exTime(i) = j
                    End If
                End If
            Next i
        End If
    Next j
    
    Dim sumPV As Double
    For i = 1 To numSim
        sumPV = sumPV + cf(i) * Exp(-r * dt * exTime(i))
    Next i
    
    PricePlainPutOneRep = sumPV / numSim
End Function

Private Function PriceLookbackPutOneRep(ws As Worksheet, ByVal numSim As Long) As Double
    Dim St As Double, r As Double, q As Double, sig As Double, tau As Double, Smax0 As Double
    Dim n As Long, dt As Double
    Dim rngInput As Range
    Set rngInput = ws.Range("E:F")
    
    St = ReadByLabelInRange(rngInput, "St")
    r = ReadByLabelInRange(rngInput, "r")
    q = ReadByLabelInRange(rngInput, "q")
    sig = ReadByLabelInRange(rngInput, "sigma")
    tau = ReadByLabelInRange(rngInput, "T-t")
    n = CLng(ReadByLabelInRange(rngInput, "n"))
    Smax0 = ReadByLabelInRange(rngInput, "Smax")
    
    dt = tau / n
    
    Dim S() As Double, Smax() As Double
    ReDim S(1 To numSim, 0 To n)
    ReDim Smax(1 To numSim, 0 To n)
    
    Dim i As Long, j As Long, z As Double
    For i = 1 To numSim
        S(i, 0) = St
        Smax(i, 0) = MaxD(Smax0, St)
        For j = 1 To n
            z = SafeNormInv()
            S(i, j) = S(i, j - 1) * Exp((r - q - 0.5 * sig ^ 2) * dt + sig * Sqr(dt) * z)
            Smax(i, j) = MaxD(Smax(i, j - 1), S(i, j))
        Next j
    Next i
    
    Dim cf() As Double, exTime() As Long
    ReDim cf(1 To numSim)
    ReDim exTime(1 To numSim)
    
    For i = 1 To numSim
        cf(i) = MaxD(Smax(i, n) - S(i, n), 0)
        exTime(i) = n
    Next i
    
    Dim ev As Double, hv As Double, pred As Double
    Dim X() As Double, Y() As Double, coef() As Double
    Dim m As Long, p As Long
    p = 6
    
    For j = n - 1 To 1 Step -1
        m = CountITMLookback(S, Smax, numSim, j)
        If m >= p Then
            ReDim X(1 To m, 1 To p)
            ReDim Y(1 To m)
            
            m = 0
            For i = 1 To numSim
                ev = MaxD(Smax(i, j) - S(i, j), 0)
                If ev > 0 Then
                    m = m + 1
                    hv = cf(i) * Exp(-r * dt * (exTime(i) - j))
                    
                    X(m, 1) = 1
                    X(m, 2) = S(i, j)
                    X(m, 3) = S(i, j) ^ 2
                    X(m, 4) = Smax(i, j)
                    X(m, 5) = Smax(i, j) ^ 2
                    X(m, 6) = S(i, j) * Smax(i, j)
                    
                    Y(m) = hv
                End If
            Next i
            
            coef = OLSCoef(X, Y, p, m)
            
            For i = 1 To numSim
                ev = MaxD(Smax(i, j) - S(i, j), 0)
                If ev > 0 Then
                    pred = coef(1) _
                         + coef(2) * S(i, j) _
                         + coef(3) * S(i, j) ^ 2 _
                         + coef(4) * Smax(i, j) _
                         + coef(5) * Smax(i, j) ^ 2 _
                         + coef(6) * S(i, j) * Smax(i, j)
                         
                    If ev > pred Then
                        cf(i) = ev
                        exTime(i) = j
                    End If
                End If
            Next i
        End If
    Next j
    
    Dim sumPV As Double
    For i = 1 To numSim
        sumPV = sumPV + cf(i) * Exp(-r * dt * exTime(i))
    Next i
    
    PriceLookbackPutOneRep = sumPV / numSim
End Function

Private Function PriceAvgCallOneRep(ws As Worksheet, ByVal numSim As Long) As Double
    Dim St As Double, K As Double, r As Double, q As Double, sig As Double
    Dim t0 As Double, tau As Double, Save0 As Double
    Dim n As Long, dt As Double
    Dim rngInput As Range
    Set rngInput = ws.Range("H:I")
    
    St = ReadByLabelInRange(rngInput, "St")
    K = ReadByLabelInRange(rngInput, "K")
    r = ReadByLabelInRange(rngInput, "r")
    q = ReadByLabelInRange(rngInput, "q")
    sig = ReadByLabelInRange(rngInput, "sigma")
    t0 = ReadByLabelInRange(rngInput, "t")
    tau = ReadByLabelInRange(rngInput, "T-t")
    n = CLng(ReadByLabelInRange(rngInput, "n"))
    Save0 = ReadByLabelInRange(rngInput, "Save,t")
    
    dt = tau / n
    
    Dim S() As Double, Save() As Double
    ReDim S(1 To numSim, 0 To n)
    ReDim Save(1 To numSim, 0 To n)
    
    Dim i As Long, j As Long, z As Double, curTime As Double
    For i = 1 To numSim
        S(i, 0) = St
        Save(i, 0) = Save0
        
        For j = 1 To n
            z = SafeNormInv()
            S(i, j) = S(i, j - 1) * Exp((r - q - 0.5 * sig ^ 2) * dt + sig * Sqr(dt) * z)
            
            curTime = t0 + (j - 1) * dt
            If curTime <= 0 Then
                Save(i, j) = S(i, j)
            Else
                Save(i, j) = (Save(i, j - 1) * curTime + S(i, j) * dt) / (curTime + dt)
            End If
        Next j
    Next i
    
    Dim cf() As Double, exTime() As Long
    ReDim cf(1 To numSim)
    ReDim exTime(1 To numSim)
    
    For i = 1 To numSim
        cf(i) = MaxD(Save(i, n) - K, 0)
        exTime(i) = n
    Next i
    
    Dim ev As Double, hv As Double, pred As Double
    Dim X() As Double, Y() As Double, coef() As Double
    Dim m As Long, p As Long
    p = 6
    
    For j = n - 1 To 1 Step -1
        m = CountITMAvg(Save, K, numSim, j)
        If m >= p Then
            ReDim X(1 To m, 1 To p)
            ReDim Y(1 To m)
            
            m = 0
            For i = 1 To numSim
                ev = MaxD(Save(i, j) - K, 0)
                If ev > 0 Then
                    m = m + 1
                    hv = cf(i) * Exp(-r * dt * (exTime(i) - j))
                    
                    X(m, 1) = 1
                    X(m, 2) = S(i, j)
                    X(m, 3) = S(i, j) ^ 2
                    X(m, 4) = Save(i, j)
                    X(m, 5) = Save(i, j) ^ 2
                    X(m, 6) = S(i, j) * Save(i, j)
                    
                    Y(m) = hv
                End If
            Next i
            
            coef = OLSCoef(X, Y, p, m)
            
            For i = 1 To numSim
                ev = MaxD(Save(i, j) - K, 0)
                If ev > 0 Then
                    pred = coef(1) _
                         + coef(2) * S(i, j) _
                         + coef(3) * S(i, j) ^ 2 _
                         + coef(4) * Save(i, j) _
                         + coef(5) * Save(i, j) ^ 2 _
                         + coef(6) * S(i, j) * Save(i, j)
                         
                    If ev > pred Then
                        cf(i) = ev
                        exTime(i) = j
                    End If
                End If
            Next i
        End If
    Next j
    
    Dim sumPV As Double
    For i = 1 To numSim
        sumPV = sumPV + cf(i) * Exp(-r * dt * exTime(i))
    Next i
    
    PriceAvgCallOneRep = sumPV / numSim
End Function

Private Function OLSCoef(X() As Double, Y() As Double, ByVal p As Long, ByVal m As Long) As Double()
    Dim a() As Double, b() As Double
    ReDim a(1 To p, 1 To p)
    ReDim b(1 To p)
    
    Dim i As Long, j As Long, K As Long
    For i = 1 To p
        For j = 1 To p
            For K = 1 To m
                a(i, j) = a(i, j) + X(K, i) * X(K, j)
            Next K
        Next j
        
        For K = 1 To m
            b(i) = b(i) + X(K, i) * Y(K)
        Next K
    Next i
    
    OLSCoef = SolveLinear(a, b, p)
End Function

Private Function SolveLinear(a() As Double, b() As Double, ByVal n As Long) As Double()
    Dim i As Long, j As Long, K As Long, maxRow As Long
    Dim maxVal As Double, tmp As Double, factor As Double
    Dim X() As Double
    ReDim X(1 To n)
    
    For i = 1 To n
        a(i, i) = a(i, i) + 0.0000000001
    Next i
    
    For i = 1 To n
        maxRow = i
        maxVal = Abs(a(i, i))
        
        For K = i + 1 To n
            If Abs(a(K, i)) > maxVal Then
                maxVal = Abs(a(K, i))
                maxRow = K
            End If
        Next K
        
        If maxRow <> i Then
            For j = i To n
                tmp = a(i, j)
                a(i, j) = a(maxRow, j)
                a(maxRow, j) = tmp
            Next j
            
            tmp = b(i)
            b(i) = b(maxRow)
            b(maxRow) = tmp
        End If
        
        If Abs(a(i, i)) < 0.000000000001 Then a(i, i) = 0.000000000001
        
        For K = i + 1 To n
            factor = a(K, i) / a(i, i)
            For j = i To n
                a(K, j) = a(K, j) - factor * a(i, j)
            Next j
            b(K) = b(K) - factor * b(i)
        Next K
    Next i
    
    For i = n To 1 Step -1
        tmp = b(i)
        For j = i + 1 To n
            tmp = tmp - a(i, j) * X(j)
        Next j
        X(i) = tmp / a(i, i)
    Next i
    
    SolveLinear = X
End Function

Private Function ReadByLabel(ws As Worksheet, ByVal labelText As String, Optional ByVal defaultVal As Variant) As Double
    Dim c As Range
    Set c = FindLabel(ws, labelText)
    
    If c Is Nothing Then
        If IsMissing(defaultVal) Then
            Err.Raise vbObjectError + 1, , "找不到輸入欄位：" & labelText
        Else
            ReadByLabel = CDbl(defaultVal)
        End If
    ElseIf IsEmpty(c.Offset(0, 1).value) Or c.Offset(0, 1).value = "" Then
        If IsMissing(defaultVal) Then
            Err.Raise vbObjectError + 2, , "欄位右邊沒有值：" & labelText
        Else
            ReadByLabel = CDbl(defaultVal)
        End If
    Else
        ReadByLabel = CDbl(c.Offset(0, 1).value)
    End If
End Function

Private Function FindLabel(ws As Worksheet, ByVal labelText As String) As Range
    Dim c As Range, target As String, cur As String
    target = CleanLabel(labelText)
    
    For Each c In ws.UsedRange
        cur = CleanLabel(CStr(c.value))
        If cur = target Then
            Set FindLabel = c
            Exit Function
        End If
    Next c
End Function

Private Function CleanLabel(ByVal S As String) As String
    S = LCase(Trim(S))
    S = Replace(S, " ", "")
    S = Replace(S, "_", "")
    S = Replace(S, ",", "")
    S = Replace(S, "，", "")
    CleanLabel = S
End Function

Private Sub WriteOutput(ws As Worksheet, _
                        ByVal optKind As String, _
                        ByVal value As Double, _
                        ByVal ciLow As Double, _
                        ByVal ciHigh As Double, _
                        ByVal sdVal As Double)

    Select Case UCase(optKind)
        Case "PLAIN"
            ws.Range("C15").value = value
            ws.Range("C16").value = ciHigh
            ws.Range("C17").value = ciLow

        Case "LOOKBACK"
            ws.Range("F15").value = value
            ws.Range("F16").value = ciHigh
            ws.Range("F17").value = ciLow

        Case "AVG"
            ws.Range("I15").value = value
            ws.Range("I16").value = ciHigh
            ws.Range("I17").value = ciLow
    End Select

End Sub

Private Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function

Private Function MeanArr(a() As Double) As Double
    Dim i As Long, S As Double
    For i = LBound(a) To UBound(a)
        S = S + a(i)
    Next i
    MeanArr = S / (UBound(a) - LBound(a) + 1)
End Function

Private Function StdevArr(a() As Double) As Double
    Dim i As Long, mu As Double, S As Double, n As Long
    n = UBound(a) - LBound(a) + 1
    mu = MeanArr(a)
    
    If n <= 1 Then
        StdevArr = 0
        Exit Function
    End If
    
    For i = LBound(a) To UBound(a)
        S = S + (a(i) - mu) ^ 2
    Next i
    
    StdevArr = Sqr(S / (n - 1))
End Function

Private Function SafeNormInv() As Double
    Dim u As Double
    
    u = Rnd()
    
    If u <= 0.0000001 Then u = 0.0000001
    If u >= 0.9999999 Then u = 0.9999999
    
    SafeNormInv = WorksheetFunction.Norm_S_Inv(u)
End Function

Private Function ReadByLabelInRange(rng As Range, _
                                    ByVal labelText As String, _
                                    Optional ByVal defaultVal As Variant) As Double

    Dim c As Range
    Dim target As String

    target = CleanLabel(labelText)

    For Each c In rng.Cells
        If CleanLabel(CStr(c.value)) = target Then
            If Trim(c.Offset(0, 1).value) = "" Then
                If IsMissing(defaultVal) Then
                    Err.Raise vbObjectError + 1, , "欄位右側沒有值：" & labelText
                Else
                    ReadByLabelInRange = CDbl(defaultVal)
                End If
            Else
                ReadByLabelInRange = CDbl(c.Offset(0, 1).value)
            End If
            Exit Function
        End If
    Next c

    If IsMissing(defaultVal) Then
        Err.Raise vbObjectError + 2, , "找不到欄位：" & labelText
    Else
        ReadByLabelInRange = CDbl(defaultVal)
    End If
End Function

Private Function CountITMPlain(S() As Double, ByVal K As Double, ByVal numSim As Long, ByVal t As Long) As Long
    Dim i As Long
    For i = 1 To numSim
        If K - S(i, t) > 0 Then CountITMPlain = CountITMPlain + 1
    Next i
End Function

Private Function CountITMLookback(S() As Double, Smax() As Double, ByVal numSim As Long, ByVal t As Long) As Long
    Dim i As Long
    For i = 1 To numSim
        If Smax(i, t) - S(i, t) > 0 Then CountITMLookback = CountITMLookback + 1
    Next i
End Function

Private Function CountITMAvg(Save() As Double, ByVal K As Double, ByVal numSim As Long, ByVal t As Long) As Long
    Dim i As Long
    For i = 1 To numSim
        If Save(i, t) - K > 0 Then CountITMAvg = CountITMAvg + 1
    Next i
End Function





