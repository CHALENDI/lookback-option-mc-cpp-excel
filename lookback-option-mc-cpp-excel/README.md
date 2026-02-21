# Lookback Option Pricing – Monte Carlo (C++ + Excel DLL)

## Overview

This project implements a floating-strike lookback option pricing engine using Monte Carlo simulation under the Black–Scholes model.

The computational core is written in C++ for performance and compiled as a Windows DLL.  
Excel is used as a user interface via VBA.

The project was developed as part of an M2 Quantitative Finance course.

---

## Model Description

We simulate geometric Brownian motion under the risk-neutral measure:

dS_t = r S_t dt + σ S_t dW_t

The floating-strike lookback payoff is:

Call:  S(T) − min S(t)
Put:   max S(t) − S(T)

The price is computed as the discounted expected payoff.

---

## Greeks

Greeks are computed via central finite differences using Common Random Numbers (CRN) for variance reduction:

- Delta  (∂P/∂S)
- Gamma  (∂²P/∂S²)
- Theta  (∂P/∂t)
- Rho    (∂P/∂r)
- Vega   (∂P/∂σ)

---

## Architecture

Excel (UI)
↓
VBA
↓
C++ DLL (LookbackMC.dll)
↓
Monte Carlo Engine

This architecture separates the computational engine from the interface, consistent with professional quantitative development practices.

---

## How to Build

1. Open "x64 Native Tools Command Prompt for Visual Studio"
2. Navigate to the cpp directory
3. Run:

   cmake -S . -B build -A x64
   cmake --build build --config Release

4. Copy LookbackMC.dll into the Excel folder

---

## How to Use

1. Open the Excel file (.xlsm)
2. Fill the input parameters
3. Run:
   - RunPricing
   - GenerateCurves

Graphs of P(S,T0) and Δ(S,T0) are automatically generated.

---

## Technologies

- C++
- Monte Carlo Simulation
- Black–Scholes Model
- Excel VBA
- DLL integration
- CMake

---

## Author

CHALENDI Mohamed