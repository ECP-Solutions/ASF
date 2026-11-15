### Compile

| Category | Workload | stdLambda | ASF | ASF / stdLambda |
|---|---|---|---|---|
| Simple expressions | Sum and product | 4.954 ms | 1.847 ms | 0.37× |
| Simple expressions | Power and modulo | 5.770 ms | 1.590 ms | 0.28× |
| Simple expressions | Comparison and logic | 7.293 ms | 1.948 ms | 0.27× |
| Simple expressions | String concatenation | 7.813 ms | 1.208 ms | 0.15× |
| Math functions | Rounding and roots | 6.816 ms | 1.786 ms | 0.26× |
| Math functions | Trigonometry and logarithms | 9.753 ms | 3.409 ms | 0.35× |
| Math functions | String functions | 7.564 ms | 1.707 ms | 0.23× |
| Script functions | One function | 10.983 ms | 1.973 ms | 0.18× |
| Script functions | Nested calls | 26.789 ms | 4.448 ms | 0.17× |
| Script functions | Recursion | 16.944 ms | 1.993 ms | 0.12× |
| COM objects | Range properties | 3.558 ms | 1.654 ms | 0.47× |
| COM objects | Range methods | 6.051 ms | 2.163 ms | 0.36× |
| COM objects | FileSystemObject methods | 9.359 ms | 1.821 ms | 0.19× |
| COM collections | Collection items | 3.832 ms | 1.304 ms | 0.34× |
| COM collections | Dictionary lookups | 5.432 ms | 1.856 ms | 0.34× |
| COM collections | Collection loop | 9.727 ms | 2.714 ms | 0.28× |
| Short programs | Variables and branch | 9.548 ms | 2.160 ms | 0.23× |
| Short programs | Loop of 5, called 100 times | 8.263 ms | 2.360 ms | 0.29× |
| Short programs | String building | 13.127 ms | 2.909 ms | 0.22× |
| Long programs | Number theory | 68.847 ms | 18.708 ms | 0.27× |
| Long programs | Text report | 85.041 ms | 12.863 ms | 0.15× |
| **All** | **Geometric mean** | | | **0.25×** |

| Category | ASF / stdLambda (geometric mean) |
|---|---|
| Simple expressions | 0.26× |
| Math functions | 0.27× |
| Script functions | 0.15× |
| COM objects | 0.32× |
| COM collections | 0.32× |
| Short programs | 0.24× |
| Long programs | 0.20× |

### Run

| Category | Workload | stdLambda | ASF | ASF / stdLambda |
|---|---|---|---|---|
| Simple expressions | Sum and product | 26.9 µs | 59.7 µs | 2.22× |
| Simple expressions | Power and modulo | 24.8 µs | 59.3 µs | 2.39× |
| Simple expressions | Comparison and logic | 28.0 µs | 39.2 µs | 1.40× |
| Simple expressions | String concatenation | 22.1 µs | 46.1 µs | 2.08× |
| Math functions | Rounding and roots | 3.674 ms | 73.2 µs | 0.02× |
| Math functions | Trigonometry and logarithms | 5.912 ms | 372.0 µs | 0.06× |
| Math functions | String functions | 4.249 ms | 178.3 µs | 0.04× |
| Script functions | One function | 29.4 µs | 61.3 µs | 2.09× |
| Script functions | Nested calls | 58.9 µs | 54.6 µs | 0.93× |
| Script functions | Recursion | 1.519 ms | 325.1 µs | 0.21× |
| COM objects | Range properties | 112.7 µs | 149.6 µs | 1.33× |
| COM objects | Range methods | 169.7 µs | 422.9 µs | 2.49× |
| COM objects | FileSystemObject methods | 3.543 ms | 350.7 µs | 0.10× |
| COM collections | Collection items | 50.3 µs | 112.9 µs | 2.24× |
| COM collections | Dictionary lookups | 3.294 ms | 4.610 ms | 1.40× |
| COM collections | Collection loop | 347.9 µs | 725.6 µs | 2.09× |
| Short programs | Variables and branch | 24.7 µs | 59.2 µs | 2.40× |
| Short programs | Loop of 5, called 100 times | 9.080 ms | 6.557 ms | 0.72× |
| Short programs | String building | 116.1 µs | 100.4 µs | 0.87× |
| Long programs | Number theory | 11.155 ms | 4.278 ms | 0.38× |
| Long programs | Text report | 59.379 ms | 835.1 µs | 0.01× |
| **All** | **Geometric mean** | | | **0.56×** |

| Category | ASF / stdLambda (geometric mean) |
|---|---|
| Simple expressions | 1.98× |
| Math functions | 0.04× |
| Script functions | 0.75× |
| COM objects | 0.69× |
| COM collections | 1.87× |
| Short programs | 1.14× |
| Long programs | 0.07× |
