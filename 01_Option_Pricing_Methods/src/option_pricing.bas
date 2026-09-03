Attribute VB_Name = "Module1"
Option Explicit
'==================================================
' Standard normal CDF
'==================================================
Private Function normal_cdf(x As Double) As Double
    normal_cdf = Application.WorksheetFunction.Norm_S_Dist(x, True)
End Function
'==================================================
' Inverse standard normal CDF
'==================================================
Private Function normal_inv_cdf(p As Double) As Double
    normal_inv_cdf = Application.WorksheetFunction.Norm_S_Inv(p)
End Function
'==================================================
' Black-Scholes d1 and d2
'==================================================
Private Function bs_d1( _
    S0 As Double, r As Double, q As Double, sigma As Double, T As Double, K As Double) As Double
    
    bs_d1 = (Log(S0 / K) + (r - q + 0.5 * sigma * sigma) * T) / (sigma * Sqr(T))
End Function

Private Function bs_d2( _
    S0 As Double, r As Double, q As Double, sigma As Double, T As Double, K As Double) As Double
    
    bs_d2 = bs_d1(S0, r, q, sigma, T, K) - sigma * Sqr(T)
End Function
'==================================================
' Black-Scholes vanilla call price
'==================================================
Public Function bs_call_price( _
    S0 As Double, r As Double, q As Double, sigma As Double, T As Double, K As Double) As Variant
    
    Dim d1 As Double, d2 As Double
    
    If S0 <= 0 Or K <= 0 Or sigma <= 0 Or T <= 0 Then
        bs_call_price = CVErr(xlErrNum)
        Exit Function
    End If
    
    d1 = bs_d1(S0, r, q, sigma, T, K)
    d2 = bs_d2(S0, r, q, sigma, T, K)
    
    bs_call_price = S0 * Exp(-q * T) * normal_cdf(d1) - K * Exp(-r * T) * normal_cdf(d2)
End Function
'==================================================
' Payoff of the target option
'==================================================
Private Function target_payoff( _
    ST As Double, K1 As Double, K2 As Double, K3 As Double, K4 As Double) As Double
    
    Dim slope As Double
    slope = (K2 - K1) / (K4 - K3)
    
    If ST <= K1 Then
        target_payoff = 0
    ElseIf ST <= K2 Then
        target_payoff = ST - K1
    ElseIf ST <= K3 Then
        target_payoff = K2 - K1
    ElseIf ST <= K4 Then
        target_payoff = slope * (K4 - ST)
    Else
        target_payoff = 0
    End If
End Function
'==================================================
' Closed-form price by martingale pricing
'==================================================
Public Function martingale_pricing( _
    S0 As Double, r As Double, q As Double, sigma As Double, T As Double, _
    K1 As Double, K2 As Double, K3 As Double, K4 As Double) As Variant
    
    Dim d11 As Double, d12 As Double, d13 As Double, d14 As Double
    Dim d21 As Double, d22 As Double, d23 As Double, d24 As Double
    Dim part1 As Double, part2 As Double, part3 As Double
    Dim slope As Double
    
    If S0 <= 0 Or sigma <= 0 Or T <= 0 Then
        martingale_pricing = CVErr(xlErrNum)
        Exit Function
    End If
    
    If Not (K1 < K2 And K2 <= K3 And K3 < K4) Then
        martingale_pricing = CVErr(xlErrNum)
        Exit Function
    End If
    
    d11 = bs_d1(S0, r, q, sigma, T, K1)
    d12 = bs_d1(S0, r, q, sigma, T, K2)
    d13 = bs_d1(S0, r, q, sigma, T, K3)
    d14 = bs_d1(S0, r, q, sigma, T, K4)
    
    d21 = bs_d2(S0, r, q, sigma, T, K1)
    d22 = bs_d2(S0, r, q, sigma, T, K2)
    d23 = bs_d2(S0, r, q, sigma, T, K3)
    d24 = bs_d2(S0, r, q, sigma, T, K4)
    
    part1 = S0 * Exp(-q * T) * (normal_cdf(d11) - normal_cdf(d12)) _
          - K1 * Exp(-r * T) * (normal_cdf(d21) - normal_cdf(d22))
    
    part2 = (K2 - K1) * Exp(-r * T) * (normal_cdf(d22) - normal_cdf(d23))
    
    slope = (K2 - K1) / (K4 - K3)
    
    part3 = slope * ( _
              K4 * Exp(-r * T) * (normal_cdf(d23) - normal_cdf(d24)) _
            - S0 * Exp(-q * T) * (normal_cdf(d13) - normal_cdf(d14)) _
            )
    
    martingale_pricing = part1 + part2 + part3
End Function
'==================================================
' Bonus 1: replication by vanilla calls
' V = C(K1) - C(K2) - a*C(K3) + a*C(K4)
'==================================================
Public Function replication_pricing( _
    S0 As Double, r As Double, q As Double, sigma As Double, T As Double, _
    K1 As Double, K2 As Double, K3 As Double, K4 As Double) As Variant
    
    Dim C1 As Variant, C2 As Variant, C3 As Variant, C4 As Variant
    Dim a As Double
    
    If S0 <= 0 Or sigma <= 0 Or T <= 0 Then
        replication_pricing = CVErr(xlErrNum)
        Exit Function
    End If
    
    If Not (K1 < K2 And K2 <= K3 And K3 < K4) Then
        replication_pricing = CVErr(xlErrNum)
        Exit Function
    End If
    
    a = (K2 - K1) / (K4 - K3)
    
    C1 = bs_call_price(S0, r, q, sigma, T, K1)
    C2 = bs_call_price(S0, r, q, sigma, T, K2)
    C3 = bs_call_price(S0, r, q, sigma, T, K3)
    C4 = bs_call_price(S0, r, q, sigma, T, K4)
    
    replication_pricing = C1 - C2 - a * C3 + a * C4
End Function
'==================================================
' One Monte Carlo run
'==================================================
Private Function monte_carlo_one_run( _
    S0 As Double, r As Double, q As Double, sigma As Double, T As Double, _
    K1 As Double, K2 As Double, K3 As Double, K4 As Double, _
    nSamples As Long) As Double
    
    Dim i As Long
    Dim Z As Double, ST As Double
    Dim sumPayoff As Double
    Dim muLog As Double
    
    muLog = Log(S0) + (r - q - 0.5 * sigma * sigma) * T
    
    For i = 1 To nSamples
        Z = normal_inv_cdf(Rnd())
        ST = Exp(muLog + sigma * Sqr(T) * Z)
        sumPayoff = sumPayoff + target_payoff(ST, K1, K2, K3, K4)
    Next i
    
    monte_carlo_one_run = Exp(-r * T) * sumPayoff / nSamples
End Function
'==================================================
' Read inputs from worksheet
'==================================================
Private Sub read_inputs( _
    ByRef S0 As Double, ByRef r As Double, ByRef q As Double, ByRef sigma As Double, ByRef T As Double, _
    ByRef K1 As Double, ByRef K2 As Double, ByRef K3 As Double, ByRef K4 As Double)
    
    S0 = Range("B1").Value
    r = Range("B2").Value
    q = Range("B3").Value
    sigma = Range("B4").Value
    T = Range("B5").Value
    K1 = Range("B6").Value
    K2 = Range("B7").Value
    K3 = Range("B8").Value
    K4 = Range("B9").Value
End Sub
'==================================================
' Validate inputs
'==================================================
Private Function inputs_are_valid( _
    S0 As Double, sigma As Double, T As Double, _
    K1 As Double, K2 As Double, K3 As Double, K4 As Double) As Boolean
    
    If S0 <= 0 Or sigma <= 0 Or T <= 0 Then
        MsgBox "S0, sigma, and T must be positive.", vbExclamation
        inputs_are_valid = False
        Exit Function
    End If
    
    If Not (K1 < K2 And K2 <= K3 And K3 < K4) Then
        MsgBox "You must have K1 < K2 <= K3 < K4.", vbExclamation
        inputs_are_valid = False
        Exit Function
    End If
    
    inputs_are_valid = True
End Function
'==================================================
' Clear output area to avoid overlap
'==================================================
Private Sub clear_output_area()
    Range("A12:B30").ClearContents
End Sub
'==================================================
' Button macro 1: martingale pricing output
' Output area: A12:B13
'==================================================
Public Sub run_martingale_pricing()

    Dim S0 As Double, r As Double, q As Double, sigma As Double, T As Double
    Dim K1 As Double, K2 As Double, K3 As Double, K4 As Double
    Dim price As Variant
    
    Call read_inputs(S0, r, q, sigma, T, K1, K2, K3, K4)
    
    If Not inputs_are_valid(S0, sigma, T, K1, K2, K3, K4) Then Exit Sub
    
    Call clear_output_area
    
    price = martingale_pricing(S0, r, q, sigma, T, K1, K2, K3, K4)
    
    Range("A12").Value = "Martingale pricing"
    Range("B12").Value = price
End Sub
'==================================================
' Button macro 2: replication pricing output
' Output area: A15:B20
'==================================================
Public Sub run_replication_pricing()

    Dim S0 As Double, r As Double, q As Double, sigma As Double, T As Double
    Dim K1 As Double, K2 As Double, K3 As Double, K4 As Double
    Dim C1 As Variant, C2 As Variant, C3 As Variant, C4 As Variant
    Dim alpha As Double
    Dim price As Variant
    
    Call read_inputs(S0, r, q, sigma, T, K1, K2, K3, K4)
    
    If Not inputs_are_valid(S0, sigma, T, K1, K2, K3, K4) Then Exit Sub
    
    Call clear_output_area
    
    alpha = (K2 - K1) / (K4 - K3)
    
    C1 = bs_call_price(S0, r, q, sigma, T, K1)
    C2 = bs_call_price(S0, r, q, sigma, T, K2)
    C3 = bs_call_price(S0, r, q, sigma, T, K3)
    C4 = bs_call_price(S0, r, q, sigma, T, K4)
    
    price = C1 - C2 - alpha * C3 + alpha * C4
    
    Range("A15").Value = "alpha"
    Range("B15").Value = alpha
    
    Range("A16").Value = "Call price at K1"
    Range("B16").Value = C1
    
    Range("A17").Value = "Call price at K2"
    Range("B17").Value = C2
    
    Range("A18").Value = "Call price at K3"
    Range("B18").Value = C3
    
    Range("A19").Value = "Call price at K4"
    Range("B19").Value = C4
    
    Range("A20").Value = "Replication pricing"
    Range("B20").Value = price
End Sub
'==================================================
' Button macro 3: Monte Carlo pricing output
' Output area: A23:B26
'==================================================
Public Sub run_monte_carlo_pricing()

    Dim S0 As Double, r As Double, q As Double, sigma As Double, T As Double
    Dim K1 As Double, K2 As Double, K3 As Double, K4 As Double
    Dim nSamples As Long, nRuns As Long
    Dim i As Long
    
    Dim estimates() As Double
    Dim meanValue As Double, sdValue As Double
    Dim sum1 As Double, sum2 As Double
    
    Call read_inputs(S0, r, q, sigma, T, K1, K2, K3, K4)
    
    If Not inputs_are_valid(S0, sigma, T, K1, K2, K3, K4) Then Exit Sub
    
    Call clear_output_area
    
    nSamples = 10000
    nRuns = 20
    
    ReDim estimates(1 To nRuns)
    
    For i = 1 To nRuns
        estimates(i) = monte_carlo_one_run(S0, r, q, sigma, T, K1, K2, K3, K4, nSamples)
        sum1 = sum1 + estimates(i)
        sum2 = sum2 + estimates(i) * estimates(i)
    Next i
    
    meanValue = sum1 / nRuns
    sdValue = Sqr((sum2 - nRuns * meanValue * meanValue) / (nRuns - 1))
    
    Range("A23").Value = "Monte Carlo mean"
    Range("B23").Value = meanValue
    
    Range("A24").Value = "Monte Carlo standard deviation"
    Range("B24").Value = sdValue
    
    Range("A25").Value = "95% CI lower bound"
    Range("B25").Value = meanValue - 2 * sdValue
    
    Range("A26").Value = "95% CI upper bound"
    Range("B26").Value = meanValue + 2 * sdValue
End Sub

