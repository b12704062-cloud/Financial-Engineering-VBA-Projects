# Vanilla Option Pricing Methods

## Overview

This project implements multiple approaches in Excel VBA
to price plain-vanilla call and put options.

The implemented approaches include:
1. Black-Scholes Model
2. Monte Carlo Simulation
3. CRR Binomial Tree
4. Binomial Black-Scholes Tree
5. Combinatorial Method

Both European and American options are considered where applicable.

## 1. Black-Scholes Pricing

The Black-Scholes model is implemented to calculate analytical prices
for European call and put options.

The program calculates d1 and d2 and evaluates the corresponding
Black-Scholes pricing formulas.

## 2. Monte Carlo Simulation

Terminal stock prices are simulated under the risk-neutral distribution:

ST = S0 × exp[(r - q - 0.5σ²)T + σ√T Z]

The implementation:
- Generates a user-specified number of simulated terminal stock prices
- Calculates call and put payoffs for each simulated price
- Discounts the average payoff to obtain the option value
- Repeats the simulation for a user-specified number of repetitions
- Reports the mean estimate and confidence interval

## 3. CRR Binomial Tree

The Cox-Ross-Rubinstein (CRR) binomial tree is implemented using
two-dimensional arrays.

The model:
- Constructs the underlying stock price tree
- Calculates terminal option payoffs
- Uses backward induction to determine option values
- Supports both European and American calls and puts
- Compares continuation value with immediate exercise value for American options

## 4. One-Dimensional CRR Tree

A memory-efficient version of the CRR model is implemented using
a one-dimensional column vector.

Instead of storing the entire stock and option value trees,
the algorithm updates option values recursively during backward induction.

## 5. Binomial Black-Scholes Tree

The Binomial Black-Scholes (BBS) tree is implemented to improve
the convergence behavior of the standard CRR model.

At the penultimate layer of the tree, Black-Scholes values are used
instead of terminal option payoffs before backward induction.

CRR and BBS pricing results are compared for:

n = 105, 110, 115, ..., 500

The program generates charts to visualize the convergence behavior
of both models for European and American calls and puts.

## 6. Combinatorial Method

A combinatorial approach is implemented to price European call
and put options directly from terminal binomial probabilities.

Logarithmic binomial probabilities are used to reduce numerical
overflow when calculating combinations for larger binomial trees.

## Tools & Concepts

- Excel VBA
- Black-Scholes Model
- Monte Carlo Simulation
- CRR Binomial Tree
- American Option Pricing
- Risk-Neutral Valuation
- Backward Induction
- Binomial Black-Scholes Tree
- Combinatorial Method
- Numerical Convergence Analysis

## Implementation

The VBA program includes:
- Black-Scholes call and put pricing
- Risk-neutral Monte Carlo simulation
- Two-dimensional CRR binomial tree
- One-dimensional CRR implementation
- European and American option pricing
- BBS tree implementation
- CRR and BBS convergence comparison
- Automatic generation of convergence charts
- Combinatorial pricing for European options
- Input validation and separate macros for each pricing method

## Results

<img width="1320" height="657" alt="image" src="https://github.com/user-attachments/assets/e2d083c6-cdd4-41d5-a681-fe809b71ef5f" />

## Source Code

[View VBA Source Code](src/option_pricing.bas)
