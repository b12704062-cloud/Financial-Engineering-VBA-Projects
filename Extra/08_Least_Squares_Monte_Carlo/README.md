# American Option Pricing with Least-Squares Monte Carlo

## Overview

This project implements the Least-Squares Monte Carlo (LSMC) method
in Excel VBA to price American-style options.

Unlike European options, American options require an exercise decision
at each time step. LSMC estimates the continuation value through
cross-sectional regression and compares it with the immediate exercise value.

Three option types are implemented:
1. Plain-Vanilla Put
2. Lookback Put
3. Arithmetic Average Call

## 1. Least-Squares Monte Carlo

Risk-neutral stock-price paths are first generated using Monte Carlo simulation.

Starting from maturity, the algorithm moves backward through time and:
- Identifies in-the-money simulation paths
- Discounts future realized cash flows to the current time step
- Estimates continuation values using OLS regression
- Compares estimated continuation values with immediate exercise payoffs
- Updates the optimal exercise time for each simulated path

The resulting cash flows are discounted to obtain the American option value.

## 2. Regression Models

Different state variables are used depending on the option's payoff structure.

### Plain-Vanilla Put

The continuation value is estimated using:

1, S, S²

### Lookback Put

Because the payoff depends on both the current stock price and running maximum,
the regression uses:

1, S, S², Smax, Smax², S × Smax

### Arithmetic Average Call

Because the payoff depends on the running arithmetic average,
the regression uses:

1, S, S², Savg, Savg², S × Savg

These additional state variables allow the regression model to incorporate
the path-dependent information relevant to each option.

## Tools & Concepts

- Excel VBA
- Least-Squares Monte Carlo
- American Options
- Monte Carlo Simulation
- OLS Regression
- Continuation Value Estimation
- Optimal Early Exercise
- Lookback Options
- Arithmetic Average Options
- Path-Dependent Derivatives

## Implementation

The VBA program includes:
- Risk-neutral path simulation
- Backward exercise decisions
- In-the-money path selection
- OLS regression implemented directly in VBA
- Gaussian elimination for solving regression coefficients
- Option-specific regression basis functions
- Exercise-time tracking
- Repeated Monte Carlo estimation
- Confidence interval calculation

## Results

<img width="706" height="622" alt="image" src="https://github.com/user-attachments/assets/84cbdc96-03f6-4419-b403-b32758e69561" />

| Option Type | LSMC Value | CI Lower | CI Upper |
| --- | ---: | ---: | ---: |
| American Plain-Vanilla Put | 4.989362 | 4.882975 | 5.095749 |
| American Lookback Put | 11.3040 | 11.21179 | 11.3962 |
| American Arithmetic Average Call | 2.4560 | 2.3860 | 2.5270 |

The LSMC framework is applied to three different American-style
derivatives using option-specific state variables in the continuation
value regression.

For the plain-vanilla put, the regression is based on the current
stock price. For the two path-dependent options, additional state
variables are included to capture the running maximum or arithmetic
average relevant to their payoff structures.

The results demonstrate how the same LSMC framework can be extended
from a standard American option to more complex path-dependent
derivatives by modifying the regression state variables.


## Source Code

[View VBA Source Code](src/option_pricing.bas)
