# 📘 Lookback Option Pricing — Monte Carlo Engine (C++ DLL + Excel VBA)

## Overview

This project implements a floating-strike lookback option pricing engine using Monte Carlo simulation under the Black–Scholes framework.

The computational core is written in C++ for performance and compiled into a Windows Dynamic Link Library (DLL). The DLL is accessed from Excel via VBA, allowing Excel to serve as a user interface while delegating heavy numerical computations to optimized C++ code.

This project was developed as part of an M2 Quantitative Finance program.

---
## Mathematical Framework

### Risk-Neutral Dynamics

Under the risk-neutral measure \( \mathbb{Q} \), the underlying asset follows a Geometric Brownian Motion:

\[
dS_t = r S_t dt + \sigma S_t dW_t
\]

Its discretized solution used in the Monte Carlo simulation is:

\[
S_{t+\Delta t} = S_t \exp \left( \left(r - \frac{1}{2}\sigma^2\right)\Delta t + \sigma \sqrt{\Delta t} Z \right)
\]

where:

- \( r \) is the risk-free rate  
- \( \sigma \) is the volatility  
- \( Z \sim \mathcal{N}(0,1) \)

---

## Floating-Strike Lookback Payoff

At maturity \( T \), the payoff is defined as:

### Call
\[
\text{Payoff} = S_T - \min_{0 \le t \le T} S_t
\]

### Put
\[
\text{Payoff} = \max_{0 \le t \le T} S_t - S_T
\]

The option price is computed as the discounted expected payoff:

\[
P = e^{-rT} \mathbb{E}^{\mathbb{Q}}[\text{Payoff}]
\]

---

## Greeks Estimation

Greeks are computed using central finite differences combined with Common Random Numbers (CRN) for variance reduction.

\[
\Delta \approx \frac{P(S+h) - P(S-h)}{2h}
\]

\[
\Gamma \approx \frac{P(S+h) - 2P(S) + P(S-h)}{h^2}
\]

\[
\rho \approx \frac{P(r+h) - P(r-h)}{2h}
\]

\[
\text{Vega} \approx \frac{P(\sigma+h) - P(\sigma-h)}{2h}
\]

\[
\Theta = \frac{\partial P}{\partial t} = -\frac{\partial P}{\partial T}
\]

Using Common Random Numbers significantly reduces Monte Carlo noise in sensitivity estimates.

---

## Architecture

The project follows a layered architecture:

Excel Interface (User Inputs & Charts)  
↓  
VBA Layer (Function Calls & Data Handling)  
↓  
C++ DLL (LookbackMC.dll)  
↓  
Monte Carlo Engine  

This design separates the interface from the computational engine, reflecting professional quantitative development practices used in financial institutions.

---

## Build Instructions

1. Open "x64 Native Tools Command Prompt for Visual Studio"
2. Navigate to the `cpp` directory
3. Run:
cmake -S . -B build -A x64

cmake --build build --config Release
4. Copy `LookbackMC.dll` into the Excel project folder.

---

## Usage

1. Open the Excel file (`.xlsm`)
2. Fill in the model parameters:
   - S0
   - r
   - sigma
   - T0
   - T
   - Number of paths
   - Number of steps
3. Run:
   - `RunPricing`
   - `GenerateCurves`

The spreadsheet automatically generates:

- \( P(S, T_0) \): Price as a function of the underlying price  
- \( \Delta(S, T_0) \): Delta as a function of the underlying price  

---

## Technologies

- C++
- Monte Carlo Simulation
- Black–Scholes Model
- Finite Difference Greeks
- Common Random Numbers (Variance Reduction)
- Excel VBA
- Windows DLL
- CMake

---

## Author

CHALENDI Mohamed
