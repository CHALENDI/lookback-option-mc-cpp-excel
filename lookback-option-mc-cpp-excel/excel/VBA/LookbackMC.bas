Attribute VB_Name = "modLookbackMC"
Option Explicit

'========================
' DLL interface
'========================
#If VBA7 Then
    Private Declare PtrSafe Function LookbackMC Lib "LookbackMC.dll" ( _
        ByVal S0 As Double, ByVal r As Double, ByVal sigma As Double, ByVal T As Double, _
        ByVal isCall As Long, ByVal nPaths As Long, ByVal nSteps As Long, _
        ByVal seed As LongLong, _
        ByVal epsS As Double, ByVal epsR As Double, ByVal epsSigma As Double, ByVal epsT As Double, _
        ByRef outArr As Double, ByVal outLen As Long) As Long
#Else
    Private Declare Function LookbackMC Lib "LookbackMC.dll" ( _
        ByVal S0 As Double, ByVal r As Double, ByVal sigma As Double, ByVal T As Double, _
        ByVal isCall As Long, ByVal nPaths As Long, ByVal nSteps As Long, _
        ByVal seed As Double, _
        ByVal epsS As Double, ByVal epsR As Double, ByVal epsSigma As Double, ByVal epsT As Double, _
        ByRef outArr As Double, ByVal outLen As Long) As Long
#End If

'========================
' Sheet layout (named ranges expected)
'========================
' Inputs sheet named ranges:
'   calcDate, maturityDate, optType, S0, r, sigma, nPaths, nSteps, seed
'   epsS, epsR, epsSigma, epsT
' Outputs sheet named ranges:
'   outPrice, outDelta, outGamma, outTheta, outRho, outVega
'
' Curves sheet expects:
'   curveS (column of S), curvePrice (price), curveDelta (delta)
'
'========================

Public Function YearFrac_ACT365(ByVal d1 As Date, ByVal d2 As Date) As Double
    YearFrac_ACT365 = (CDbl(d2) - CDbl(d1)) / 365#
End Function

Private Function GetIsCall(ByVal s As String) As Long
    s = LCase$(Trim$(s))
    If s = "call" Or s = "c" Then
        GetIsCall = 1
    Else
        GetIsCall = 0
    End If
End Function

Public Sub RunPricing()
    On Error GoTo EH

    Dim wsIn As Worksheet, wsOut As Worksheet
    Set wsIn = ThisWorkbook.Worksheets("Inputs")
    Set wsOut = ThisWorkbook.Worksheets("Outputs")

    Dim d0 As Date, dT As Date, T As Double
    Dim S0 As Double, r As Double, sigma As Double
    Dim nPaths As Long, nSteps As Long
    Dim seed As LongLong
    Dim epsS As Double, epsR As Double, epsSigma As Double, epsT As Double
    Dim isCall As Long

    If IsEmpty(Range("calcDate").Value) Or Range("calcDate").Value = 0 Then
        d0 = Date
        Range("calcDate").Value = d0
    Else
        d0 = CDate(Range("calcDate").Value)
    End If

    dT = CDate(Range("maturityDate").Value)
    T = YearFrac_ACT365(d0, dT)
    If T <= 0# Then Err.Raise vbObjectError + 100, , "Maturity must be after calculation date."

    S0 = CDbl(Range("S0").Value)
    r = CDbl(Range("r").Value)
    sigma = CDbl(Range("sigma").Value)

    nPaths = CLng(Range("nPaths").Value)
    nSteps = CLng(Range("nSteps").Value)
    seed = CLngLng(Range("seed").Value)

    epsS = CDbl(Range("epsS").Value)
    epsR = CDbl(Range("epsR").Value)
    epsSigma = CDbl(Range("epsSigma").Value)
    epsT = CDbl(Range("epsT").Value)

    isCall = GetIsCall(CStr(Range("optType").Value))

    Dim outArr(0 To 5) As Double
    Dim rc As Long
    rc = LookbackMC(S0, r, sigma, T, isCall, nPaths, nSteps, seed, epsS, epsR, epsSigma, epsT, outArr(0), 6)

    If rc <> 0 Then
        Err.Raise vbObjectError + 200, , "C++ DLL returned error code: " & rc
    End If

    Range("outPrice").Value = outArr(0)
    Range("outDelta").Value = outArr(1)
    Range("outGamma").Value = outArr(2)
    Range("outTheta").Value = outArr(3)
    Range("outRho").Value = outArr(4)
    Range("outVega").Value = outArr(5)

    Exit Sub

EH:
    MsgBox "RunPricing error: " & Err.Description, vbCritical
End Sub

Public Sub GenerateCurves()
    On Error GoTo EH

    Application.ScreenUpdating = False

    Dim wsC As Worksheet
    Set wsC = ThisWorkbook.Worksheets("Curves")

    ' Read base inputs
    Dim d0 As Date, dT As Date, T As Double
    If IsEmpty(Range("calcDate").Value) Or Range("calcDate").Value = 0 Then
        d0 = Date
        Range("calcDate").Value = d0
    Else
        d0 = CDate(Range("calcDate").Value)
    End If
    dT = CDate(Range("maturityDate").Value)
    T = YearFrac_ACT365(d0, dT)

    Dim S0 As Double, r As Double, sigma As Double
    S0 = CDbl(Range("S0").Value)
    r = CDbl(Range("r").Value)
    sigma = CDbl(Range("sigma").Value)

    Dim nPaths As Long, nSteps As Long
    nPaths = CLng(Range("nPaths").Value)
    nSteps = CLng(Range("nSteps").Value)

    Dim seed As LongLong
    seed = CLngLng(Range("seed").Value)

    Dim epsS As Double, epsR As Double, epsSigma As Double, epsT As Double
    epsS = CDbl(Range("epsS").Value)
    epsR = CDbl(Range("epsR").Value)
    epsSigma = CDbl(Range("epsSigma").Value)
    epsT = CDbl(Range("epsT").Value)

    Dim isCall As Long
    isCall = GetIsCall(CStr(Range("optType").Value))

    ' Build S grid around S0: 50% to 150% by default
    Dim nPts As Long: nPts = 31
    Dim i As Long
    Dim Smin As Double, Smax As Double
    Smin = 0.5 * S0
    Smax = 1.5 * S0

    wsC.Range("A1:C1").Value = Array("S", "Price", "Delta")

    For i = 0 To nPts - 1
        Dim S As Double
        S = Smin + (Smax - Smin) * (CDbl(i) / CDbl(nPts - 1))

        Dim outArr(0 To 5) As Double
        Dim rc As Long
        rc = LookbackMC(S, r, sigma, T, isCall, nPaths, nSteps, seed, epsS, epsR, epsSigma, epsT, outArr(0), 6)
        If rc <> 0 Then Err.Raise vbObjectError + 210, , "C++ DLL returned error code: " & rc

        wsC.Cells(i + 2, 1).Value = S
        wsC.Cells(i + 2, 2).Value = outArr(0)
        wsC.Cells(i + 2, 3).Value = outArr(1)
    Next i

    ' Create / refresh charts
    Call BuildCharts(wsC, nPts)

    Application.ScreenUpdating = True
    Exit Sub

EH:
    Application.ScreenUpdating = True
    MsgBox "GenerateCurves error: " & Err.Description, vbCritical
End Sub

Private Sub BuildCharts(ByVal ws As Worksheet, ByVal nPts As Long)
    Dim lastRow As Long: lastRow = nPts + 1

    ' Delete existing charts on Curves sheet
    Dim co As ChartObject
    For Each co In ws.ChartObjects
        co.Delete
    Next co

    ' Price chart
    Dim ch1 As ChartObject
    Set ch1 = ws.ChartObjects.Add(Left:=300, Top:=20, Width:=520, Height:=260)
    ch1.Chart.ChartType = xlXYScatterLinesNoMarkers
    ch1.Chart.SetSourceData Source:=ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 2))
    ch1.Chart.HasTitle = True
    ch1.Chart.ChartTitle.Text = "Lookback Option Price P(S, T0)"
    ch1.Chart.Axes(xlCategory).HasTitle = True
    ch1.Chart.Axes(xlCategory).AxisTitle.Text = "Underlying S0"
    ch1.Chart.Axes(xlValue).HasTitle = True
    ch1.Chart.Axes(xlValue).AxisTitle.Text = "Price"

    ' Delta chart
    Dim ch2 As ChartObject
    Set ch2 = ws.ChartObjects.Add(Left:=300, Top:=300, Width:=520, Height:=260)
    ch2.Chart.ChartType = xlXYScatterLinesNoMarkers
    ch2.Chart.SetSourceData Source:=ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 3))
    ch2.Chart.HasTitle = True
    ch2.Chart.ChartTitle.Text = "Lookback Option Delta Δ(S, T0)"
    ch2.Chart.Axes(xlCategory).HasTitle = True
    ch2.Chart.Axes(xlCategory).AxisTitle.Text = "Underlying S0"
    ch2.Chart.Axes(xlValue).HasTitle = True
    ch2.Chart.Axes(xlValue).AxisTitle.Text = "Delta"
End Sub
