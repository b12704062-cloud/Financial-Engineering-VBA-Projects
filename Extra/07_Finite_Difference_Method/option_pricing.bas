Attribute VB_Name = "Module1"
Option Explicit

Sub Extra2_Click()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim S0 As Double, K As Double, r As Double, q As Double, sigma As Double, t As Double
    Dim Smin As Double, Smax As Double
    Dim m As Long, n As Long
    
    On Error GoTo ErrorHandler
    
    S0 = ws.Range("C2").value
    K = ws.Range("C3").value
    r = ws.Range("C4").value
    q = ws.Range("C5").value
    sigma = ws.Range("C6").value
    t = ws.Range("C7").value
    
    If IsEmpty(ws.Range("C8").value) Or ws.Range("C8").value = "" Then
        Smin = 0
        ws.Range("C8").value = Smin
    Else
        Smin = ws.Range("C8").value
    End If
    
    If IsEmpty(ws.Range("C9").value) Or ws.Range("C9").value = "" Then
        Smax = 2 * S0
        ws.Range("C9").value = Smax
    Else
        Smax = ws.Range("C9").value
    End If
    
    If IsEmpty(ws.Range("C10").value) Or ws.Range("C10").value = "" Then
        m = CLng(2 * S0)
        ws.Range("C10").value = m
    Else
        m = CLng(ws.Range("C10").value)
    End If
    
    n = CLng(ws.Range("C11").value)
    
    If S0 <= 0 Or K <= 0 Or sigma <= 0 Or t <= 0 Or m < 3 Or n < 1 Then
        MsgBox "Input error: 請確認 S0, K, sigma, T > 0，且 m >= 3, n >= 1。", vbExclamation
        Exit Sub
    End If
    
    If S0 < Smin Or S0 > Smax Then
        MsgBox "Input error: S0 必須介於 Smin 與 Smax 之間。", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Implicit
    ws.Range("G3").value = FDMPrice(S0, K, r, q, sigma, t, Smin, Smax, m, n, "C", "E", "I")
    ws.Range("G4").value = FDMPrice(S0, K, r, q, sigma, t, Smin, Smax, m, n, "P", "E", "I")
    ws.Range("G5").value = FDMPrice(S0, K, r, q, sigma, t, Smin, Smax, m, n, "C", "A", "I")
    ws.Range("G6").value = FDMPrice(S0, K, r, q, sigma, t, Smin, Smax, m, n, "P", "A", "I")
    
    ' Explicit
    ws.Range("G8").value = FDMPrice(S0, K, r, q, sigma, t, Smin, Smax, m, n, "C", "E", "X")
    ws.Range("G9").value = FDMPrice(S0, K, r, q, sigma, t, Smin, Smax, m, n, "P", "E", "X")
    ws.Range("G10").value = FDMPrice(S0, K, r, q, sigma, t, Smin, Smax, m, n, "C", "A", "X")
    ws.Range("G11").value = FDMPrice(S0, K, r, q, sigma, t, Smin, Smax, m, n, "P", "A", "X")
    
    ws.Range("G3:G11").NumberFormat = "0.000000"
    
CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Extra2 error: " & Err.Description, vbCritical
End Sub

Private Function FDMPrice(ByVal S0 As Double, ByVal K As Double, _
                          ByVal r As Double, ByVal q As Double, _
                          ByVal sigma As Double, ByVal t As Double, _
                          ByVal Smin As Double, ByVal Smax As Double, _
                          ByVal m As Long, ByVal n As Long, _
                          ByVal OptType As String, ByVal StyleType As String, _
                          ByVal MethodType As String) As Double
    Dim dS As Double, dt As Double
    Dim oldV() As Double, newV() As Double
    Dim j As Long, stepIdx As Long
    Dim S As Double
    
    dS = (Smax - Smin) / m
    dt = t / n
    
    ReDim oldV(0 To m)
    ReDim newV(0 To m)
    
    ' Maturity payoff
    For j = 0 To m
        S = Smin + j * dS
        oldV(j) = Payoff(S, K, OptType)
    Next j
    
    ' Backward induction
    For stepIdx = n - 1 To 0 Step -1
        If MethodType = "I" Then
            Call ImplicitStep(oldV, newV, K, r, q, sigma, dt, dS, Smin, Smax, m, OptType, StyleType)
        Else
            Call ExplicitStep(oldV, newV, K, r, q, sigma, dt, dS, Smin, Smax, m, OptType, StyleType)
        End If
        
        For j = 0 To m
            oldV(j) = newV(j)
        Next j
    Next stepIdx
    
    FDMPrice = InterpolatePrice(S0, Smin, dS, m, oldV)
End Function

Private Sub ImplicitStep(ByRef oldV() As Double, ByRef newV() As Double, _
                         ByVal K As Double, ByVal r As Double, ByVal q As Double, _
                         ByVal sigma As Double, ByVal dt As Double, ByVal dS As Double, _
                         ByVal Smin As Double, ByVal Smax As Double, ByVal m As Long, _
                         ByVal OptType As String, ByVal StyleType As String)
    Dim Nint As Long
    Nint = m - 1
    
    Dim lower() As Double, diag() As Double, upper() As Double, rhs() As Double
    Dim sol() As Double
    ReDim lower(1 To Nint)
    ReDim diag(1 To Nint)
    ReDim upper(1 To Nint)
    ReDim rhs(1 To Nint)
    ReDim sol(1 To Nint)
    
    Dim j As Long, idx As Long
    Dim aj As Double, bj As Double, cj As Double
    Dim S As Double
    
    newV(0) = BoundaryValue(Smin, K, OptType)
    newV(m) = BoundaryValue(Smax, K, OptType)
    
    For j = 1 To m - 1
        idx = j
        
        aj = ((r - q) * j * dt / 2) - (0.5 * sigma ^ 2 * j ^ 2 * dt)
        bj = 1 + sigma ^ 2 * j ^ 2 * dt + r * dt
        cj = -((r - q) * j * dt / 2) - (0.5 * sigma ^ 2 * j ^ 2 * dt)
        
        lower(idx) = aj
        diag(idx) = bj
        upper(idx) = cj
        rhs(idx) = oldV(j)
    Next j
    
    rhs(1) = rhs(1) - lower(1) * newV(0)
    rhs(Nint) = rhs(Nint) - upper(Nint) * newV(m)
    
    lower(1) = 0
    upper(Nint) = 0
    
    Call SolveTridiagonal(lower, diag, upper, rhs, sol, Nint)
    
    For j = 1 To m - 1
        S = Smin + j * dS
        newV(j) = sol(j)
        
        If StyleType = "A" Then
            newV(j) = WorksheetFunction.Max(newV(j), Payoff(S, K, OptType))
        End If
    Next j
End Sub

Private Sub ExplicitStep(ByRef oldV() As Double, ByRef newV() As Double, _
                         ByVal K As Double, ByVal r As Double, ByVal q As Double, _
                         ByVal sigma As Double, ByVal dt As Double, ByVal dS As Double, _
                         ByVal Smin As Double, ByVal Smax As Double, ByVal m As Long, _
                         ByVal OptType As String, ByVal StyleType As String)
    Dim j As Long
    Dim aj As Double, bj As Double, cj As Double
    Dim S As Double
    
    newV(0) = BoundaryValue(Smin, K, OptType)
    newV(m) = BoundaryValue(Smax, K, OptType)
    
    For j = 1 To m - 1
        aj = 0.5 * sigma ^ 2 * j ^ 2 * dt - 0.5 * (r - q) * j * dt
        bj = 1 - sigma ^ 2 * j ^ 2 * dt - r * dt
        cj = 0.5 * sigma ^ 2 * j ^ 2 * dt + 0.5 * (r - q) * j * dt
        
        newV(j) = aj * oldV(j - 1) + bj * oldV(j) + cj * oldV(j + 1)
        
        S = Smin + j * dS
        If StyleType = "A" Then
            newV(j) = WorksheetFunction.Max(newV(j), Payoff(S, K, OptType))
        End If
    Next j
End Sub

Private Sub SolveTridiagonal(ByRef lower() As Double, ByRef diag() As Double, _
                             ByRef upper() As Double, ByRef rhs() As Double, _
                             ByRef X() As Double, ByVal n As Long)
    Dim cPrime() As Double, dPrime() As Double
    ReDim cPrime(1 To n)
    ReDim dPrime(1 To n)
    
    Dim i As Long
    Dim denom As Double
    
    If Abs(diag(1)) < 0.000000000001 Then Err.Raise vbObjectError + 100, , "Zero pivot in tridiagonal solver."
    
    cPrime(1) = upper(1) / diag(1)
    dPrime(1) = rhs(1) / diag(1)
    
    For i = 2 To n
        denom = diag(i) - lower(i) * cPrime(i - 1)
        If Abs(denom) < 0.000000000001 Then Err.Raise vbObjectError + 101, , "Zero pivot in tridiagonal solver."
        
        cPrime(i) = upper(i) / denom
        dPrime(i) = (rhs(i) - lower(i) * dPrime(i - 1)) / denom
    Next i
    
    X(n) = dPrime(n)
    
    For i = n - 1 To 1 Step -1
        X(i) = dPrime(i) - cPrime(i) * X(i + 1)
    Next i
End Sub

Private Function Payoff(ByVal S As Double, ByVal K As Double, ByVal OptType As String) As Double
    If UCase(OptType) = "C" Then
        Payoff = WorksheetFunction.Max(S - K, 0)
    Else
        Payoff = WorksheetFunction.Max(K - S, 0)
    End If
End Function

Private Function BoundaryValue(ByVal S As Double, ByVal K As Double, ByVal OptType As String) As Double
    ' 簡化邊界條件：
    ' Call: S=0 -> 0, S=Smax -> Smax-K
    ' Put : S=0 -> K, S=Smax -> 0
    If UCase(OptType) = "C" Then
        BoundaryValue = WorksheetFunction.Max(S - K, 0)
    Else
        BoundaryValue = WorksheetFunction.Max(K - S, 0)
    End If
End Function

Private Function InterpolatePrice(ByVal S0 As Double, ByVal Smin As Double, _
                                  ByVal dS As Double, ByVal m As Long, _
                                  ByRef V() As Double) As Double
    Dim pos As Double
    Dim jLow As Long
    Dim w As Double
    
    pos = (S0 - Smin) / dS
    
    If pos <= 0 Then
        InterpolatePrice = V(0)
        Exit Function
    End If
    
    If pos >= m Then
        InterpolatePrice = V(m)
        Exit Function
    End If
    
    jLow = Int(pos)
    w = pos - jLow
    
    InterpolatePrice = (1 - w) * V(jLow) + w * V(jLow + 1)
End Function



