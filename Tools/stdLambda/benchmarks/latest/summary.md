### Compile

| Workload | stdLambda | ASF | ASF / stdLambda |
|---|---|---|---|
| Arithmetic | 5.115 ms | 1.980 ms | 0.39× |
| Math operations | 5.898 ms | 1.752 ms | 0.30× |
| Conditional branches | 3.574 ms | 1.339 ms | 0.37× |
| For/While loops | 9.843 ms | 2.184 ms | 0.22× |
| Function calls | 10.113 ms | 1.573 ms | 0.16× |
| Fibonacci recursion | 18.325 ms | 3.839 ms | 0.21× |
| Array transformations | 3.121 ms | 2.088 ms | 0.67× |
| Object property access | 2.671 ms | 888.1 µs | 0.33× |
| COM property access | 3.165 ms | 1.003 ms | 0.32× |
| **Geometric mean** | | | **0.30×** |

### Run

| Workload | stdLambda | ASF | ASF / stdLambda |
|---|---|---|---|
| Arithmetic | 23.0 µs | 42.7 µs | 1.86× |
| Math operations | 2.916 ms | 58.7 µs | 0.02× |
| Conditional branches | 21.3 µs | 39.2 µs | 1.84× |
| For/While loops | 5.559 ms | 4.041 ms | 0.73× |
| Function calls | 30.5 µs | 64.7 µs | 2.12× |
| Fibonacci recursion | 1.778 ms | 416.8 µs | 0.23× |
| Array transformations | 631.9 µs | 385.9 µs | 0.61× |
| Object property access | 1.862 ms | 30.4 µs | 0.02× |
| COM property access | 40.7 µs | 92.7 µs | 2.28× |
| **Geometric mean** | | | **0.44×** |
