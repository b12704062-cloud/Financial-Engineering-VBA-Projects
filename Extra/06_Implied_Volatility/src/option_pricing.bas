Attribute VB_Name = "Module1"
Option Explicit

Private Const PI As Double = 3.14159265358979
Private Const NEWTON_H As Double = 0.0000000001
Private Const MAX_ITER As Long = 100
Private Const SIGMA_LOW As Double = 0.01
Private Const SIGMA_HIGH As Double = 5#

Public Sub Run_Extra1()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Extra1")
    
    Dim S0 As Double, K As Double, r As Double, q As Double
    Dim t As Double, marketPrice As Double, convCri As Double
    Dim n As Long
    
    S0 = GetInputValueAny(ws, "S0")
    K = GetInputValueAny(ws, "K")
    r = GetInputValueAny(ws, "r")
    q = GetInputValueAny(ws, "q")
    t = GetInputValueAny(ws, "T")
    marketPrice = GetInputValueAny(ws, "market price", "market_price")
    n = CLng(GetInputValueAny(ws, "n"))
    convCri = GetInputValueAny(ws, "conv_cri", "convergence criterion")
    
    If convCri <= 0 Then convCri = 0.0001
    
    ws.Range("G4").value = ImpliedVol_Bisection(S0, K, r, q, t, marketPrice, n, convCri, "BS", "EC")
    ws.Range("G5").value = ImpliedVol_Bisection(S0, K, r, q, t, marketPrice, n, convCri, "BS", "EP")
    
    ws.Range("G7").value = ImpliedVol_Bisection(S0, K, r, q, t, marketPrice, n, convCri, "TREE", "EC")
    ws.Range("G8").value = ImpliedVol_Bisection(S0, K, r, q, t, marketPrice, n, convCri, "TREE", "EP")
    ws.Range("G9").value = ImpliedVol_Bisection(S0, K, r, q, t, marketPrice, n, convCri, "TREE", "AC")
    ws.Range("G10").value = ImpliedVol_Bisection(S0, K, r, q, t, marketPrice, n, convCri, "TREE", "AP")
    
    ws.Range("K4").value = ImpliedVol_Newton(S0, K, r, q, t, marketPrice, n, convCri, "BS", "EC")
    ws.Range("K5").value = ImpliedVol_Newton(S0, K, r, q, t, marketPrice, n, convCri, "BS", "EP")
    
    ws.Range("K7").value = ImpliedVol_Newton(S0, K, r, q, t, marketPrice, n, convCri, "TREE", "EC")
    ws.Range("K8").value = ImpliedVol_Newton(S0, K, r, q, t, marketPrice, n, convCri, "TREE", "EP")
    ws.Range("K9").value = ImpliedVol_Newton(S0, K, r, q, t, marketPrice, n, convCri, "TREE", "AC")
    ws.Range("K10").value = ImpliedVol_Newton(S0, K, r, q, t, marketPrice, n, convCri, "TREE", "AP")

End Sub

Private Function GetInputValueAny(ws As Worksheet, ParamArray labels()) As Double
    Dim i As Long, f As Range
    
    For i = LBound(labels) To UBound(labels)
        Set f = ws.Cells.Find(What:=CStr(labels(i)), LookAt:=xlWhole, MatchCase:=False)
        If Not f Is Nothing Then
            GetInputValueAny = CDbl(f.Offset(0, 1).value)
            Exit Function
        End If
    Next i
End Function

Private Function ImpliedVol_Bisection( _
    S0 As Double, K As Double, r As Double, q As Double, t As Double, _
    marketPrice As Double, n As Long, convCri As Double, _
    modelType As String, OptType As String) As Double
    
    Dim a As Double, b As Double, X As Double
    Dim fA As Double, fX As Double
    Dim i As Long
    
    a = SIGMA_LOW
    b = SIGMA_HIGH
    
    For i = 1 To MAX_ITER
        X = (a + b) / 2#
        
        fA = OptionPrice(S0, K, r, q, t, a, n, modelType, OptType) - marketPrice
        fX = OptionPrice(S0, K, r, q, t, X, n, modelType, OptType) - marketPrice
        
        If fA * fX < 0 Then
            b = X
        Else
            a = X
        End If
        
        If Abs(b - a) < convCri Then Exit For
    Next i
    
    ImpliedVol_Bisection = (a + b) / 2#

End Function

Private Function ImpliedVol_Newton( _
    S0 As Double, K As Double, r As Double, q As Double, t As Double, _
    marketPrice As Double, n As Long, convCri As Double, _
    modelType As String, OptType As String) As Double
        
    Dim sigmaOld As Double, sigmaNew As Double
    Dim fValue As Double, derivativeValue As Double
    Dim i As Long
    
    sigmaOld = 0.2
    
    For i = 1 To MAX_ITER
        fValue = OptionPrice(S0, K, r, q, t, sigmaOld, n, modelType, OptType) - marketPrice
        derivativeValue = OptionDerivative(S0, K, r, q, t, sigmaOld, n, modelType, OptType)
        
        sigmaNew = sigmaOld - fValue / derivativeValue
        
        If sigmaNew <= SIGMA_LOW Then sigmaNew = SIGMA_LOW
        If sigmaNew > SIGMA_HIGH Then sigmaNew = SIGMA_HIGH
        
        If Abs(sigmaNew - sigmaOld) < convCri Then Exit For
        
        sigmaOld = sigmaNew
    Next i
    
    ImpliedVol_Newton = sigmaNew

End Function

Private Function OptionDerivative( _
    S0 As Double, K As Double, r As Double, q As Double, t As Double, _
    sigma As Double, n As Long, modelType As String, OptType As String) As Double
    
    If UCase(modelType) = "BS" Then
        OptionDerivative = BSVega(S0, K, r, q, t, sigma)
    Else
        OptionDerivative = _
            (BinomialPrice(S0, K, r, q, t, sigma + NEWTON_H, n, OptType) _
            - BinomialPrice(S0, K, r, q, t, sigma, n, OptType)) / NEWTON_H
    End If

End Function

Private Function OptionPrice( _
    S0 As Double, K As Double, r As Double, q As Double, t As Double, _
    sigma As Double, n As Long, modelType As String, OptType As String) As Double
    
    If UCase(modelType) = "BS" Then
        If UCase(OptType) = "EC" Then
            OptionPrice = BSPrice(S0, K, r, q, t, sigma, True)
        Else
            OptionPrice = BSPrice(S0, K, r, q, t, sigma, False)
        End If
    Else
        OptionPrice = BinomialPrice(S0, K, r, q, t, sigma, n, OptType)
    End If

End Function

Private Function BSPrice( _
    S0 As Double, K As Double, r As Double, q As Double, _
    t As Double, sigma As Double, isCall As Boolean) As Double
    
    Dim d1 As Double, d2 As Double
    
    d1 = (Log(S0 / K) + (r - q + 0.5 * sigma * sigma) * t) / (sigma * Sqr(t))
    d2 = d1 - sigma * Sqr(t)
    
    If isCall Then
        BSPrice = S0 * Exp(-q * t) * CND(d1) - K * Exp(-r * t) * CND(d2)
    Else
        BSPrice = K * Exp(-r * t) * CND(-d2) - S0 * Exp(-q * t) * CND(-d1)
    End If

End Function

Private Function BSVega( _
    S0 As Double, K As Double, r As Double, q As Double, _
    t As Double, sigma As Double) As Double
    
    Dim d1 As Double
    
    d1 = (Log(S0 / K) + (r - q + 0.5 * sigma * sigma) * t) / (sigma * Sqr(t))
    
    BSVega = S0 * Exp(-q * t) * Sqr(t) * CNDPrime(d1)

End Function

Private Function BinomialPrice( _
    S0 As Double, K As Double, r As Double, q As Double, t As Double, _
    sigma As Double, n As Long, OptType As String) As Double
    
    Dim dt As Double, u As Double, d As Double, p As Double, disc As Double
    Dim i As Long, j As Long
    Dim V() As Double
    Dim Snode As Double, cont As Double, exv As Double
    Dim isCall As Boolean, isAmerican As Boolean
    
    dt = t / n
    
    u = Exp(sigma * Sqr(dt))
    d = 1# / u
    p = (Exp((r - q) * dt) - d) / (u - d)
    disc = Exp(-r * dt)
    
    isCall = (Right(UCase(OptType), 1) = "C")
    isAmerican = (Left(UCase(OptType), 1) = "A")
    
    ReDim V(0 To n)
    
    For j = 0 To n
        Snode = StockNode_LogSpace(S0, sigma, dt, n, j)
        V(j) = Payoff(Snode, K, isCall)
    Next j
    
    For i = n - 1 To 0 Step -1
        For j = 0 To i
            cont = disc * (p * V(j) + (1# - p) * V(j + 1))
            
            If isAmerican Then
                Snode = StockNode_LogSpace(S0, sigma, dt, i, j)
                exv = Payoff(Snode, K, isCall)
                V(j) = Max2(cont, exv)
            Else
                V(j) = cont
            End If
        Next j
    Next i
    
    BinomialPrice = V(0)

End Function

Private Function StockNode_LogSpace( _
    S0 As Double, sigma As Double, dt As Double, _
    i As Long, j As Long) As Double
    
    StockNode_LogSpace = Exp(Log(S0) + (i - 2# * j) * sigma * Sqr(dt))

End Function

Private Function Payoff(S As Double, K As Double, isCall As Boolean) As Double
    If isCall Then
        Payoff = Max2(S - K, 0#)
    Else
        Payoff = Max2(K - S, 0#)
    End If
End Function

Private Function Max2(a As Double, b As Double) As Double
    If a > b Then
        Max2 = a
    Else
        Max2 = b
    End If
End Function

Private Function CND(X As Double) As Double
    CND = Application.WorksheetFunction.Norm_S_Dist(X, True)
End Function

Private Function CNDPrime(X As Double) As Double
    CNDPrime = Exp(-0.5 * X * X) / Sqr(2# * PI)
End Function


