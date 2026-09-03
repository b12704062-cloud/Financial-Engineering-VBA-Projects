# Financial Engineering with Excel VBA

A collection of financial engineering projects implemented in Excel VBA,
covering derivative pricing, Monte Carlo simulation, numerical methods,
and path-dependent option valuation.

These projects were developed through financial engineering coursework
and focus on translating financial models into numerical algorithms
and executable pricing programs.

## Projects

| Project | Derivatives / Application | Key Methods |
| --- | --- | --- |
| [01. Option Pricing Methods](01_Option_Pricing_Methods/) | Customized Piecewise-Payoff Option | Martingale Pricing, Replicating Portfolio, Monte Carlo |
| [02. Vanilla Option Pricing](02_Vanilla_Option_Pricing/) | European & American Options | Black-Scholes, CRR Binomial Tree, BBS Tree, Monte Carlo, Combinatorial Method |
| [03. Rainbow Option Monte Carlo](03_Rainbow_Option_Monte_Carlo/) | Multi-Asset Rainbow Option | Cholesky Decomposition, Antithetic Variates, Moment Matching, Inverse Cholesky |
| [04. Lookback Option Pricing](04_Lookback_Option_Pricing/) | European & American Lookback Puts | Path-Dependent Binomial Tree, State-Space Reduction, Cheuk-Vorst Method |
| [05. Arithmetic Average Option](05_Arithmetic_Average_Option/) | European & American Asian Calls | Binomial Tree, Monte Carlo, State Discretization, Interpolation |
| [06. Implied Volatility](06_Implied_Volatility/) | European & American Options | Black-Scholes, CRR Tree, Bisection, Newton's Method |
| [07. Finite Difference Method](07_Finite_Difference_Method/) | European & American Options | Black-Scholes PDE, Explicit FDM, Implicit FDM, Tridiagonal Solver |
| [08. Least-Squares Monte Carlo](08_Least_Squares_Monte_Carlo/) | Vanilla, Lookback & Arithmetic Average American Options | Monte Carlo, OLS Regression, Early-Exercise Optimization |

## Key Topics

### Derivatives Pricing
- Black-Scholes Model
- European and American Options
- Rainbow Options
- Lookback Options
- Arithmetic Average (Asian) Options
- Implied Volatility

### Numerical & Quantitative Methods
- Monte Carlo Simulation
- CRR Binomial Tree
- Finite Difference Methods
- Least-Squares Monte Carlo
- Cholesky Decomposition
- Variance Reduction
- Numerical Root Finding
- Linear Interpolation
- OLS Regression

### Computational Techniques
- Path-dependent state tracking
- State-space reduction
- Backward induction
- Matrix decomposition
- Numerical optimization
- Computational efficiency comparison

## Selected Highlights

### Monte Carlo Variance Reduction

For a multi-asset rainbow option, variance-reduction techniques were
implemented and compared using common random samples.

The combination of antithetic variates and moment matching reduced
the confidence interval width to approximately **48.43%** of the basic
Cholesky simulation, while the inverse Cholesky approach further reduced
it to approximately **12.18%**.

→ [View Rainbow Option Project](03_Rainbow_Option_Monte_Carlo/)

### Path-Dependent Option Pricing

Lookback and arithmetic average options were implemented by extending
standard binomial-tree frameworks with additional state variables.

Different approaches to state construction, interpolation, and search
were explored to improve computational efficiency while preserving
pricing results.

→ [View Lookback Option Project](04_Lookback_Option_Pricing/)

→ [View Arithmetic Average Option Project](05_Arithmetic_Average_Option/)

### American Options with LSMC

Least-Squares Monte Carlo was implemented to estimate continuation values
and optimal early-exercise decisions for American-style derivatives.

The framework was extended from plain-vanilla puts to path-dependent
lookback puts and arithmetic average calls by incorporating additional
state variables into the regression model.

→ [View LSMC Project](08_Least_Squares_Monte_Carlo/)

## Tools

- **Programming:** Excel VBA
- **Financial Models:** Black-Scholes, CRR, BBS, Cheuk-Vorst
- **Simulation:** Monte Carlo, Variance Reduction, LSMC
- **Numerical Methods:** FDM, Root Finding, Interpolation
- **Statistical Methods:** OLS Regression, Confidence Interval Estimation

## Repository Structure

Each project folder contains:

- `README.md` — methodology, implementation, and results
- `src/` — VBA source code
- `images/` — selected Excel outputs and numerical results

## About

These projects were completed as part of financial engineering coursework.

The repository is organized as a portfolio of my implementations,
with emphasis on understanding the financial models, translating them
into numerical algorithms, and evaluating their computational behavior.
