### Compile

| Category | Workload | stdLambda | ASF | ASF / stdLambda |
|---|---|---|---|---|
| Simple expressions | Sum and product | 5.352 ms | 2.028 ms | 0.38× |
| Simple expressions | Power and modulo | 6.716 ms | 2.279 ms | 0.34× |
| Simple expressions | Comparison and logic | 7.620 ms | 2.654 ms | 0.35× |
| Simple expressions | String concatenation | 5.078 ms | 1.766 ms | 0.35× |
| Math functions | Rounding and roots | 9.344 ms | 2.660 ms | 0.28× |
| Math functions | Trigonometry and logarithms | 9.593 ms | 3.189 ms | 0.33× |
| Math functions | String functions | 8.989 ms | 3.148 ms | 0.35× |
| Script functions | One function | 16.940 ms | 2.989 ms | 0.18× |
| Script functions | Nested calls | 30.272 ms | 5.042 ms | 0.17× |
| Script functions | Recursion | 17.565 ms | 3.581 ms | 0.20× |
| COM objects | Range properties | 5.536 ms | 1.757 ms | 0.32× |
| COM objects | Range methods | 6.664 ms | 2.303 ms | 0.35× |
| COM objects | FileSystemObject methods | 10.060 ms | 2.518 ms | 0.25× |
| COM collections | Collection items | 5.253 ms | 1.738 ms | 0.33× |
| COM collections | Dictionary lookups | 6.282 ms | 1.524 ms | 0.24× |
| COM collections | Collection loop | 9.000 ms | 2.179 ms | 0.24× |
| Short programs | Variables and branch | 9.783 ms | 1.724 ms | 0.18× |
| Short programs | Loop of 5, called 100 times | 8.894 ms | 2.027 ms | 0.23× |
| Short programs | String building | 8.895 ms | 2.234 ms | 0.25× |
| Long programs | Number theory | 56.444 ms | 12.260 ms | 0.22× |
| Long programs | Text report | 43.404 ms | 8.684 ms | 0.20× |
| **All** | **Geometric mean** | | | **0.26×** |

| Category | ASF / stdLambda (geometric mean) |
|---|---|
| Simple expressions | 0.35× |
| Math functions | 0.32× |
| Script functions | 0.18× |
| COM objects | 0.30× |
| COM collections | 0.27× |
| Short programs | 0.22× |
| Long programs | 0.21× |

### Run

| Category | Workload | stdLambda | ASF | ASF / stdLambda |
|---|---|---|---|---|
| Simple expressions | Sum and product | 32.5 µs | 67.2 µs | 2.06× |
| Simple expressions | Power and modulo | 40.9 µs | 72.0 µs | 1.76× |
| Simple expressions | Comparison and logic | 38.8 µs | 69.4 µs | 1.79× |
| Simple expressions | String concatenation | 36.4 µs | 77.5 µs | 2.13× |
| Math functions | Rounding and roots | 3.962 ms | 79.6 µs | 0.02× |
| Math functions | Trigonometry and logarithms | 6.465 ms | 254.3 µs | 0.04× |
| Math functions | String functions | 6.245 ms | 376.5 µs | 0.06× |
| Script functions | One function | 45.1 µs | 73.2 µs | 1.62× |
| Script functions | Nested calls | 75.2 µs | 69.8 µs | 0.93× |
| Script functions | Recursion | 1.797 ms | 482.8 µs | 0.27× |
| COM objects | Range properties | 91.9 µs | 139.4 µs | 1.52× |
| COM objects | Range methods | 187.6 µs | 458.9 µs | 2.45× |
| COM objects | FileSystemObject methods | 3.980 ms | 388.6 µs | 0.10× |
| COM collections | Collection items | 65.1 µs | 116.4 µs | 1.79× |
| COM collections | Dictionary lookups | 3.281 ms | 4.267 ms | 1.30× |
| COM collections | Collection loop | 340.8 µs | 623.0 µs | 1.83× |
| Short programs | Variables and branch | 23.1 µs | 35.8 µs | 1.54× |
| Short programs | Loop of 5, called 100 times | 6.336 ms | 3.885 ms | 0.61× |
| Short programs | String building | 71.9 µs | 62.4 µs | 0.87× |
| Long programs | Number theory | 8.071 ms | 1.995 ms | 0.25× |
| Long programs | Text report | 52.813 ms | 919.9 µs | 0.02× |
| **All** | **Geometric mean** | | | **0.53×** |

| Category | ASF / stdLambda (geometric mean) |
|---|---|
| Simple expressions | 1.93× |
| Math functions | 0.04× |
| Script functions | 0.74× |
| COM objects | 0.71× |
| COM collections | 1.62× |
| Short programs | 0.94× |
| Long programs | 0.07× |
