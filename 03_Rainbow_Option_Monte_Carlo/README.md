# Rainbow Option Pricing with Monte Carlo Simulation

## Overview

This project implements Monte Carlo simulation in Excel VBA
to price a maximum rainbow option written on multiple correlated assets.

The option payoff is:

Payoff = max(max(S1T, S2T, ..., SnT) - K, 0)

Three simulation approaches are implemented:
1. Cholesky Decomposition
2. Antithetic Variates with Moment Matching
3. Antithetic Variates with Moment Matching and Inverse Cholesky

The methods are compared based on their option value estimates
and the width of the resulting confidence intervals.

## 1. Cholesky Decomposition

The basic Monte Carlo method uses Cholesky decomposition to generate
correlated asset returns from independent standard normal random variables.

The implementation:
- Supports an arbitrary number of underlying assets
- Constructs the covariance matrix from asset volatilities and correlations
- Performs Cholesky decomposition of the covariance matrix
- Generates correlated random variables
- Simulates terminal prices for all underlying assets
- Calculates the maximum terminal asset price and option payoff
- Discounts the average payoff to obtain the option value

## 2. Antithetic Variates & Moment Matching

Variance reduction is introduced by combining antithetic variates
with moment matching.

The implementation:
- Uses the first half of the random samples from the basic simulation
- Generates antithetic samples by multiplying them by -1
- Standardizes the simulated samples through moment matching
- Applies the Cholesky transformation to generate correlated returns
- Recalculates the option value and confidence interval

Common random samples are reused from the basic simulation to make
the comparison between methods more consistent.

## 3. Inverse Cholesky Method

The third approach further applies the inverse Cholesky method
to reduce sampling error in the simulated correlation structure.

The implementation:
- Starts from the antithetic and moment-matched samples
- Calculates their sample covariance matrix
- Performs Cholesky decomposition on the sample covariance matrix
- Applies the inverse Cholesky transformation to whiten the samples
- Transforms the samples again using the target covariance matrix
- Recalculates the rainbow option value and confidence interval

The confidence interval widths of the three methods are compared
to evaluate the effectiveness of variance reduction.

## Tools & Concepts

- Excel VBA
- Rainbow Options
- Monte Carlo Simulation
- Multivariate Simulation
- Cholesky Decomposition
- Covariance & Correlation Matrices
- Antithetic Variates
- Moment Matching
- Inverse Cholesky Transformation
- Variance Reduction
- Common Random Numbers
- Risk-Neutral Valuation

## Implementation

The VBA program includes:
- Dynamic input tables for multiple underlying assets
- Correlation and covariance matrix handling
- Cholesky decomposition implemented in VBA
- Correlated multivariate normal simulation
- Maximum rainbow option payoff calculation
- Antithetic variate generation
- Moment matching
- Sample covariance matrix calculation
- Upper-triangular matrix inversion
- Inverse Cholesky transformation
- Common random sample reuse across methods
- Confidence interval comparison between simulation methods
- Input validation and separate macros for each pricing method

## Results

<img width="1582" height="463" alt="image" src="https://github.com/user-attachments/assets/8661e200-37a0-478b-8c31-3cc64f1b1c46" />

A useful comparison is:

| Method | Option Value | CI Lower | CI Upper | Relative CI Width |
| --- | ---: | ---: | ---: | ---: |
| Cholesky | 25.6359 | 24.9319 | 26.3399 | 100% |
| Antithetic + Moment Matching | 25.6597 | 25.3187 | 26.0007 | 48.43% |
| Inverse Cholesky | 25.6414 | 25.5557 | 25.7272 | 12.18% |

## Source Code

[View VBA Source Code](src/option_pricing.bas)
