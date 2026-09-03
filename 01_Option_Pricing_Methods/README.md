# Option Pricing Methods

## Overview

This project implements three different approaches in Excel VBA
to price a customized option with a piecewise payoff structure.

The three approaches are:
1. Martingale pricing
2. Replicating portfolio
3. Monte Carlo simulation

## 1. Martingale Pricing

A closed-form pricing formula is implemented based on risk-neutral
martingale pricing.

The program calculates Black-Scholes d1 and d2 values for four strike
prices and evaluates the expected discounted payoff across each payoff region.

## 2. Replicating Portfolio

The target payoff is replicated using a portfolio of vanilla call options:

V = C(K1) - C(K2) - αC(K3) + αC(K4)

where:

α = (K2 - K1) / (K4 - K3)

Each vanilla call is priced using the Black-Scholes model.

## 3. Monte Carlo Simulation

Terminal stock prices are simulated under the risk-neutral distribution:

ST = S0 × exp[(r - q - 0.5σ²)T + σ√T Z]

The implementation:
- Generates 10,000 simulated terminal stock prices per run
- Calculates the customized option payoff for each simulated price
- Discounts the average payoff to obtain the option value
- Repeats the simulation 20 times
- Reports the mean estimate, standard deviation, and confidence interval

## Tools & Concepts

- Excel VBA
- Black-Scholes Model
- Martingale Pricing
- Replicating Portfolio
- Risk-Neutral Valuation
- Monte Carlo Simulation
- Confidence Interval Estimation

## Implementation

The VBA program includes:
- Standard normal CDF and inverse CDF
- Black-Scholes call pricing
- Customized piecewise payoff calculation
- Input validation
- Separate macros for analytical, replication, and simulation-based pricing

## Results

<img width="1003" height="605" alt="image" src="https://github.com/user-attachments/assets/e903879a-8384-4e38-876e-92325ee59e5d" />


## Source Code

`src/option_pricing.bas`
