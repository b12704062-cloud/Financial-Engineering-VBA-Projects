# Arithmetic Average Option Pricing

## Overview

This project implements numerical methods in Excel VBA
to price European and American arithmetic average call options.

The option payoff depends on the arithmetic average of the underlying
stock prices observed from the issue date to maturity:

Payoff = max(Savg,T - K, 0)

Because the arithmetic average introduces an additional path-dependent
state variable, the binomial-tree implementation requires discretization
and interpolation of possible average-price states.

The project focuses on:
1. Binomial Tree with Average-Price State Discretization
2. Monte Carlo Simulation
3. Sequential, Binary, and Direct Linear Search for Interpolation
4. Discrete Sampling of the Arithmetic Average

## 1. Binomial Tree with Average-Price States

The basic binomial-tree implementation extends the CRR framework by
introducing the running arithmetic average as an additional state variable.

At each stock-price node, a range of possible average prices is
approximated by a discrete grid between Amax and Amin.

The implementation:
- Constructs CRR stock-price dynamics
- Determines the maximum and minimum feasible average prices at each node
- Discretizes the average-price state into m grid points
- Calculates future average prices after upward and downward movements
- Uses interpolation to obtain continuation values
- Applies backward induction to price European and American calls
- Compares continuation and immediate exercise values for American options

## 2. Monte Carlo Simulation

Monte Carlo simulation is implemented as an alternative method
for pricing the European arithmetic average call.

For each simulated path:
- The stock price evolves under the risk-neutral process
- Stock prices along the path are accumulated
- The arithmetic average is calculated at maturity
- The terminal payoff is evaluated
- The payoff is discounted to obtain its present value

The simulation is repeated multiple times to estimate the option value
and confidence interval.

## 3. Search Methods for Interpolation

During backward induction, the future average prices Au and Ad generally
do not fall exactly on the discretized average-price grid.

The program therefore locates the two neighboring grid points and uses
linear interpolation to estimate the corresponding continuation values.

Three methods are implemented to locate the interpolation interval:

### Sequential Search

The grid is searched sequentially until the interval containing the
target average price is identified.

### Binary Search

Binary search reduces the number of comparisons required to locate
the appropriate interpolation interval.

### Direct Linear Positioning

Because the average-price grid is evenly spaced between Amax and Amin,
the approximate grid position can be calculated directly from the
target average price.

This avoids repeatedly searching through the grid and substantially
reduces computational time.

## 4. Discrete Average Sampling

The basic model assumes that the arithmetic average is updated at
every time step.

An alternative implementation prices an arithmetic average call when
the stock price is sampled only once every five time steps:

0, 5Δt, 10Δt, ..., T

The running average is updated only at observation dates while the
underlying stock price continues to evolve at every binomial-tree step.

This allows the model to represent discretely monitored Asian options.

## Tools & Concepts

- Excel VBA
- Asian Options
- Arithmetic Average Options
- Path-Dependent Derivatives
- CRR Binomial Tree
- Monte Carlo Simulation
- Risk-Neutral Valuation
- Backward Induction
- State-Space Discretization
- Linear Interpolation
- Sequential Search
- Binary Search
- Computational Efficiency
- Discrete Monitoring

## Implementation

The VBA program includes:
- CRR stock-price dynamics
- Three-dimensional storage of average-price states
- Maximum and minimum feasible average calculations
- Uniform discretization of average-price states
- Linear interpolation of continuation values
- European and American arithmetic average call pricing
- Sequential search for interpolation intervals
- Binary search for interpolation intervals
- Direct calculation of interpolation positions
- Built-in execution-time comparison between search methods
- Monte Carlo path simulation
- Box-Muller standard normal generation
- Discrete observation intervals for arithmetic averaging
- Separate macros for Basic, Bonus 1, and Bonus 2 methods

## Results

<img width="1132" height="546" alt="image" src="https://github.com/user-attachments/assets/2416f317-4359-45cd-8afb-dee3db0925d8" />

### Pricing Validation

| Method | European Call | American Call |
| --- | ---: | ---: |
| Binomial Tree | 2.379472 | 2.507866 |
| Monte Carlo Simulation | 2.317857 | — |

The binomial-tree estimate for the European arithmetic average call
is close to the Monte Carlo estimate.

The Monte Carlo estimate of 2.317857 is obtained with a confidence
interval of [2.249099, 2.386615], which contains the binomial-tree
estimate of 2.379472.

### Search Method Efficiency

| Search Method | European Call | American Call | Execution Time |
| --- | ---: | ---: | ---: |
| Sequential Search | 2.379472 | 2.507866 | 1.761719 sec |
| Binary Search | 2.379472 | 2.507866 | 1.273438 sec |
| Direct Linear Positioning | 2.379472 | 2.507866 | 0.992188 sec |

All three search methods produce identical option values while
requiring different amounts of computation time.

Compared with sequential search, binary search reduces execution time
by approximately 27.7%, while direct linear positioning reduces
execution time by approximately 43.7%.

The results demonstrate that exploiting the uniformly spaced
average-price grid allows the interpolation interval to be located
more efficiently without changing the pricing result.

### Discrete Sampling

| Sampling Method | European Call | American Call |
| --- | ---: | ---: |
| Every time step | 2.379472 | 2.507866 |
| Every 5Δt | 2.400270 | 2.541809 |

The discretely sampled implementation updates the arithmetic average
only once every five time steps while the underlying stock price
continues to evolve at every step.

This produces different option values because the payoff depends on
the set of stock prices included in the arithmetic average.

## Source Code

[View VBA Source Code](src/option_pricing.bas)
