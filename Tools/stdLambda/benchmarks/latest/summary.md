### Compile

| Workload | stdLambda | ASF | ASF / stdLambda |
|---|---|---|---|
| Arithmetic | 5.184 ms | 1.758 ms | 0.34× |
| Math operations | 6.596 ms | 2.243 ms | 0.34× |
| Conditional branches | 4.218 ms | 1.441 ms | 0.34× |
| For/While loops | 7.814 ms | 1.735 ms | 0.22× |
| Function calls | 7.449 ms | 1.518 ms | 0.20× |
| Fibonacci recursion | 9.454 ms | 1.963 ms | 0.21× |
| Array transformations | 2.858 ms | 1.992 ms | 0.70× |
| Object property access | 3.144 ms | 900.6 µs | 0.29× |
| COM property access | 2.719 ms | 933.4 µs | 0.34× |
| **Geometric mean** | | | **0.31×** |

### Run

| Workload | stdLambda | ASF | ASF / stdLambda |
|---|---|---|---|
| Arithmetic | 22.7 µs | 50.1 µs | 2.21× |
| Math operations | 1.788 ms | 60.0 µs | 0.03× |
| Conditional branches | 13.6 µs | 41.6 µs | 3.07× |
| For/While loops | 6.033 ms | 3.650 ms | 0.60× |
| Function calls | 19.5 µs | 38.8 µs | 1.99× |
| Fibonacci recursion | 961.3 µs | 339.0 µs | 0.35× |
| Array transformations | 605.5 µs | 350.6 µs | 0.58× |
| Object property access | 951.5 µs | 31.5 µs | 0.03× |
| COM property access | 30.5 µs | 130.0 µs | 4.26× |
| **Geometric mean** | | | **0.58×** |
