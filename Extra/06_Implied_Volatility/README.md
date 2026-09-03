# Implied Volatility Solver

## Overview

This project implements an implied volatility solver in Excel VBA
using multiple option pricing models and numerical root-finding methods.

Given an observed market option price, the program solves for the
volatility that makes the theoretical option value match the market price.

The implementation combines:
1. Black-Scholes Model
2. CRR Binomial Tree
3. Bisection Method
4. Newton's Method

## 1. Pricing Models

The Black-Scholes model is used for:
- European calls
- European puts

The CRR binomial tree is used for:
- European calls
- European puts
- American calls
- American puts

For American options, early exercise is evaluated during backward induction.

## 2. Implied Volatility Solvers

### Bisection Method

The volatility search interval is repeatedly divided in half until
the specified convergence criterion is satisfied.

This provides a stable numerical approach to solving implied volatility.

### Newton's Method

Newton's method iteratively updates the volatility estimate using
the sensitivity of the option price to changes in volatility.

For the Black-Scholes model, analytical Vega is used.

For the binomial-tree model, the derivative is approximated numerically
using a finite difference.

## Tools & Concepts

- Excel VBA
- Implied Volatility
- Black-Scholes Model
- CRR Binomial Tree
- European & American Options
- Bisection Method
- Newton's Method
- Vega
- Numerical Root Finding

## Implementation

The VBA program allows the user to select:
- Call or Put
- Black-Scholes or Binomial Tree pricing
- Bisection or Newton's Method

The solver then iteratively searches for the implied volatility until
the specified convergence criterion is satisfied.

## Results

<img width="845" height="596" alt="image" src="https://github.com/user-attachments/assets/ee9b48d4-1f53-4d6c-8224-f80b74d6ce22" />

## Source Code

[View VBA Source Code](src/option_pricing.bas)
