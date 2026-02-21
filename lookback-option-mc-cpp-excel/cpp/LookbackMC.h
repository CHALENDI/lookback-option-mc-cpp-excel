#pragma once

#ifdef _WIN32
  #define DLL_EXPORT extern "C" __declspec(dllexport)
  #define STDCALL __stdcall
#else
  #define DLL_EXPORT extern "C"
  #define STDCALL
#endif

// out[0]=Price
// out[1]=Delta dP/dS
// out[2]=Gamma d2P/dS2
// out[3]=Theta dP/dt  
// out[4]=Rho   dP/dr
// out[5]=Vega  dP/dsigma
DLL_EXPORT int STDCALL LookbackMC(
    double S0,
    double r,
    double sigma,
    double T,          // time to maturity in years
    int isCall,        // 1=Call, 0=Put
    int nPaths,
    int nSteps,
    unsigned long long seed,
    double epsS,
    double epsR,
    double epsSigma,
    double epsT,
    double* out,       // must point to at least 6 doubles
    int outLen          // must be >= 6
);
