# Option Pricing with Finite Difference Methods

## Overview

This project implements finite difference methods in Excel VBA
to price European and American call and put options.

The Black-Scholes partial differential equation is solved numerically
on a discretized stock-price and time grid.

Two numerical schemes are implemented:
1. Implicit Finite Difference Method
2. Explicit Finite Difference Method

## 1. Implicit Finite Difference Method

The implicit method constructs a system of linear equations at each
time step.

The implementation:
- Discretizes stock price and time into a finite grid
- Constructs a tridiagonal system for each backward time step
- Solves the system using a tridiagonal matrix algorithm
- Applies backward induction from maturity to the valuation date
- Supports European and American calls and puts

For American options, the continuation value is compared with the
immediate exercise payoff at each grid point.

## 2. Explicit Finite Difference Method

The explicit method calculates each option value directly from
neighboring values at the following time step.

The implementation uses the same stock-price and time grid as the
implicit method and supports:
- European calls
- European puts
- American calls
- American puts

Early exercise is incorporated for American options by comparing
the calculated continuation value with the intrinsic value.

## Tools & Concepts

- Excel VBA
- Black-Scholes PDE
- Finite Difference Method
- Explicit Scheme
- Implicit Scheme
- Tridiagonal Matrix
- Backward Induction
- Linear Interpolation
- European & American Options
- Early Exercise

## Implementation

The VBA program includes:
- Stock-price and time grid construction
- Implicit and explicit finite difference schemes
- Tridiagonal system solver
- Boundary conditions for calls and puts
- Early-exercise handling for American options
- Linear interpolation to estimate the option value at S0

Default grid settings are provided when not specified:
- Smin = 0
- Smax = 2 × S0
- m = 2 × S0

## Results

<img width="450" height="515" alt="image" src="https://github.com/user-attachments/assets/5739195a-bbb9-4927-a72f-3889c0ccc7fe" />

## Source Code

[View VBA Source Code](src/option_pricing.bas)
