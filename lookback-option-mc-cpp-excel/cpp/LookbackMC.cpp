#include "LookbackMC.h"

#include <vector>
#include <random>
#include <cmath>
#include <algorithm>

static inline double price_with_normals(
    double S0, double r, double sigma, double T, bool isCall,
    int nPaths, int nSteps,
    const std::vector<double>& Z
) {
    if (T <= 0.0 || nPaths <= 0 || nSteps <= 0) return 0.0;

    const double dt = T / static_cast<double>(nSteps);
    const double drift = (r - 0.5 * sigma * sigma) * dt;
    const double vol = sigma * std::sqrt(dt);
    const double disc = std::exp(-r * T);

    double sumPayoff = 0.0;

    size_t k = 0;
    for (int i = 0; i < nPaths; ++i) {
        double S = S0;
        double minS = S0;
        double maxS = S0;

        for (int j = 0; j < nSteps; ++j, ++k) {
            const double incr = drift + vol * Z[k];
            S *= std::exp(incr);
            if (S < minS) minS = S;
            if (S > maxS) maxS = S;
        }

        const double payoff = isCall ? std::max(0.0, S - minS)
                                     : std::max(0.0, maxS - S);
        sumPayoff += payoff;
    }

    return disc * (sumPayoff / static_cast<double>(nPaths));
}

extern "C" __declspec(dllexport)
int __stdcall LookbackMC(
    double S0,
    double r,
    double sigma,
    double T,
    int isCallInt,
    int nPaths,
    int nSteps,
    unsigned long long seed,
    double epsS,
    double epsR,
    double epsSigma,
    double epsT,
    double* out,
    int outLen
) {
    if (!out || outLen < 6) return -1;
    if (S0 <= 0.0 || sigma <= 0.0 || T < 0.0) return -2;
    if (nPaths <= 0 || nSteps <= 0) return -3;

    // Default finite-difference bumps if user provides 0.
    const double bumpS = (epsS > 0.0) ? epsS : (1e-4 * std::max(1.0, S0));
    const double bumpR = (epsR > 0.0) ? epsR : 1e-4;
    const double bumpV = (epsSigma > 0.0) ? epsSigma : 1e-4;
    const double bumpT = (epsT > 0.0) ? epsT : (1.0 / 365.0); // ~1 day in years

    const bool isCall = (isCallInt != 0);

    // Generate common random numbers (CRN) once
    const size_t nZ = static_cast<size_t>(nPaths) * static_cast<size_t>(nSteps);
    std::vector<double> Z(nZ);

    std::mt19937_64 rng(seed);
    std::normal_distribution<double> nd(0.0, 1.0);
    for (size_t i = 0; i < nZ; ++i) Z[i] = nd(rng);

    auto P = [&](double S, double rr, double sig, double TT) -> double {
        if (TT <= 0.0) {
            // maturity now: payoff depends on min/max over empty interval -> min=max=S0 -> payoff 0
            return 0.0;
        }
        return price_with_normals(S, rr, sig, TT, isCall, nPaths, nSteps, Z);
    };

    const double price = P(S0, r, sigma, T);

    // Central differences with CRN for variance reduction
    const double price_upS   = P(S0 + bumpS, r, sigma, T);
    const double price_dnS   = P(std::max(1e-12, S0 - bumpS), r, sigma, T);
    const double delta = (price_upS - price_dnS) / (2.0 * bumpS);
    const double gamma = (price_upS - 2.0 * price + price_dnS) / (bumpS * bumpS);

    const double price_upR = P(S0, r + bumpR, sigma, T);
    const double price_dnR = P(S0, r - bumpR, sigma, T);
    const double rho = (price_upR - price_dnR) / (2.0 * bumpR);

    const double price_upV = P(S0, r, sigma + bumpV, T);
    const double price_dnV = P(S0, r, std::max(1e-12, sigma - bumpV), T);
    const double vega = (price_upV - price_dnV) / (2.0 * bumpV);

    // Theta with respect to calendar time t: Theta = dP/dt = - dP/dT
    // because T here is "time to maturity".
    const double price_upT = P(S0, r, sigma, T + bumpT);
    const double price_dnT = P(S0, r, sigma, std::max(1e-12, T - bumpT));
    const double dPdT = (price_upT - price_dnT) / (2.0 * bumpT);
    const double theta = -dPdT;

    out[0] = price;
    out[1] = delta;
    out[2] = gamma;
    out[3] = theta;
    out[4] = rho;
    out[5] = vega;

    return 0;
}
