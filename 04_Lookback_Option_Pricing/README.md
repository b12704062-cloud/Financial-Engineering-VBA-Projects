# Lookback Option Pricing

## Overview

This project implements several binomial-tree approaches in Excel VBA
to price European and American lookback put options.

Unlike vanilla options, the value of a lookback option depends on the
historical maximum asset price, which introduces an additional state
variable and substantially increases computational complexity.

The project focuses on different approaches to representing the running
maximum and reducing the computational cost of path-dependent pricing.

The implemented approaches are:
1. Binomial Tree with Explicit Smax State Tracking
2. Direct Smax List Construction
3. Cheuk-Vorst Transformed-State Binomial Method

Monte Carlo simulation is also implemented as an alternative method
for pricing European lookback puts.

## 1. Binomial Tree with Explicit Smax Tracking

The basic implementation extends the standard CRR binomial tree by
introducing the historical maximum price Smax as an additional state variable.

Possible Smax states are propagated through the tree and stored together
with each stock-price node.

This approach provides a direct representation of the path-dependent
payoff but requires a relatively large state space and greater
computational effort.

## 2. Monte Carlo Simulation

Monte Carlo simulation is implemented to price the European lookback put.

For each simulated price path:
- The stock price evolves under the risk-neutral process
- The maximum stock price along the path is recorded
- The terminal lookback payoff is calculated
- The payoff is discounted to obtain the present value

The simulation is repeated multiple times to estimate the option value
and confidence interval.

## 3. Direct Smax List Construction

The second binomial implementation improves the handling of the
path-dependent state variable.

Instead of constructing each node's Smax states by inheriting information
from parent nodes, the feasible Smax values are determined directly
from the location of the node.

This reduces the amount of state propagation required and provides
a more efficient way to construct the lookback-option lattice while
producing the same option value as the basic approach.

## 4. Cheuk-Vorst Binomial Method

The Cheuk-Vorst method provides a more compact representation of the
lookback-option state space by transforming the state variable into
the relative distance between the current stock price and its running maximum.

This avoids explicitly storing the full set of Smax values at every node
and substantially reduces the dimensionality of the binomial lattice.

In this implementation, the Cheuk-Vorst method is applied under the
initial condition:

Smax,t = St

and therefore should not be directly compared with the basic and
direct-Smax methods when the historical maximum differs from the
current stock price.

## Tools & Concepts

- Excel VBA
- Lookback Options
- Path-Dependent Derivatives
- CRR Binomial Tree
- Monte Carlo Simulation
- Risk-Neutral Valuation
- Backward Induction
- State-Space Reduction
- Early Exercise
- Cheuk-Vorst Method

## Implementation

The VBA program includes:
- CRR stock-price dynamics
- Historical maximum price tracking
- Three-dimensional state storage for the basic binomial method
- European and American lookback put pricing
- Monte Carlo path simulation
- Box-Muller normal random number generation
- Direct construction of feasible Smax states
- Duplicate-state detection and state matching
- Cheuk-Vorst transformed-state lattice
- Separate macros for Basic, Bonus 1, and Bonus 2 methods

## Results

<img width="991" height="462" alt="image" src="https://github.com/user-attachments/assets/1c260aa7-5916-4e2c-9888-909dd2ed4480" />
<img width="342" height="130" alt="image" src="https://github.com/user-attachments/assets/7e97b304-2334-4de8-9e15-e3ed9fac5137" />

### Computational Efficiency

| Method | Average Execution Time |
| --- | ---: |
| Basic Binomial Tree | 65.91504 sec |
| Direct Smax Construction | 1.188477 sec |
| Cheuk-Vorst Method* | 0.01367188 sec |

The direct Smax construction substantially reduces computation time
by avoiding the propagation of Smax states from parent nodes.

The Cheuk-Vorst transformed-state approach further reduces the
dimensionality of the pricing problem and provides a substantial
improvement in computational efficiency.

*The Cheuk-Vorst implementation is applied under the initial condition
Smax,t = St.

## Source Code

[View VBA Source Code](src/option_pricing.bas)
