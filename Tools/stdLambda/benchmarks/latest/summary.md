### Compile

| Category | Workload | stdLambda | ASF | ASF / stdLambda |
|---|---|---|---|---|
| Simple expressions | Sum and product | 2.969 ms | 1.299 ms | 0.44× |
| Simple expressions | Power and modulo | 5.083 ms | 1.571 ms | 0.31× |
| Simple expressions | Comparison and logic | 3.633 ms | 2.196 ms | 0.60× |
| Simple expressions | String concatenation | 2.030 ms | 667.9 µs | 0.33× |
| Math functions | Rounding and roots | 3.375 ms | 1.098 ms | 0.33× |
| Math functions | Trigonometry and logarithms | 4.407 ms | 1.621 ms | 0.37× |
| Math functions | String functions | 3.795 ms | 1.020 ms | 0.27× |
| Script functions | One function | 5.734 ms | 1.384 ms | 0.24× |
| Script functions | Nested calls | 14.576 ms | 3.945 ms | 0.27× |
| Script functions | Recursion | 8.742 ms | 2.206 ms | 0.25× |
| COM objects | Range properties | 2.221 ms | 1.250 ms | 0.56× |
| COM objects | Range methods | 3.502 ms | 1.007 ms | 0.29× |
| COM objects | FileSystemObject methods | 4.604 ms | 1.220 ms | 0.27× |
| COM collections | Collection items | 2.020 ms | 727.0 µs | 0.36× |
| COM collections | Dictionary lookups | 3.727 ms | 1.217 ms | 0.33× |
| COM collections | Collection loop | 5.902 ms | 2.888 ms | 0.49× |
| Short programs | Variables and branch | 6.232 ms | 2.153 ms | 0.35× |
| Short programs | Loop of 5, called 100 times | 5.191 ms | 2.815 ms | 0.54× |
| Short programs | String building | 5.430 ms | 2.686 ms | 0.49× |
| Long programs | Number theory | 72.705 ms | 15.586 ms | 0.21× |
| Long programs | Text report | 69.165 ms | 17.008 ms | 0.25× |
| **All** | **Geometric mean** | | | **0.34×** |

| Category | ASF / stdLambda (geometric mean) |
|---|---|
| Simple expressions | 0.40× |
| Math functions | 0.32× |
| Script functions | 0.25× |
| COM objects | 0.35× |
| COM collections | 0.39× |
| Short programs | 0.45× |
| Long programs | 0.23× |

### Run

| Category | Workload | stdLambda | ASF | ASF / stdLambda |
|---|---|---|---|---|
| Simple expressions | Sum and product | 17.4 µs | 18.7 µs | 1.07× |
| Simple expressions | Power and modulo | 12.8 µs | 21.7 µs | 1.69× |
| Simple expressions | Comparison and logic | 13.0 µs | 18.9 µs | 1.45× |
| Simple expressions | String concatenation | 11.3 µs | 23.7 µs | 2.10× |
| Math functions | Rounding and roots | 1.247 ms | 26.6 µs | 0.02× |
| Math functions | Trigonometry and logarithms | 2.380 ms | 102.6 µs | 0.04× |
| Math functions | String functions | 1.607 ms | 110.1 µs | 0.07× |
| Script functions | One function | 14.6 µs | 24.3 µs | 1.66× |
| Script functions | Nested calls | 41.0 µs | 31.5 µs | 0.77× |
| Script functions | Recursion | 842.0 µs | 324.2 µs | 0.39× |
| COM objects | Range properties | 50.2 µs | 59.5 µs | 1.19× |
| COM objects | Range methods | 93.6 µs | 192.0 µs | 2.05× |
| COM objects | FileSystemObject methods | 1.285 ms | 173.5 µs | 0.13× |
| COM collections | Collection items | 24.7 µs | 50.7 µs | 2.05× |
| COM collections | Dictionary lookups | 1.289 ms | 1.833 ms | 1.42× |
| COM collections | Collection loop | 229.9 µs | 341.4 µs | 1.48× |
| Short programs | Variables and branch | 15.9 µs | 19.6 µs | 1.24× |
| Short programs | Loop of 5, called 100 times | 4.303 ms | 2.739 ms | 0.64× |
| Short programs | String building | 47.1 µs | 40.3 µs | 0.86× |
| Long programs | Number theory | 5.239 ms | 1.694 ms | 0.32× |
| Long programs | Text report | 21.240 ms | 1.313 ms | 0.06× |
| **All** | **Geometric mean** | | | **0.55×** |

| Category | ASF / stdLambda (geometric mean) |
|---|---|
| Simple expressions | 1.53× |
| Math functions | 0.04× |
| Script functions | 0.79× |
| COM objects | 0.69× |
| COM collections | 1.63× |
| Short programs | 0.88× |
| Long programs | 0.14× |
