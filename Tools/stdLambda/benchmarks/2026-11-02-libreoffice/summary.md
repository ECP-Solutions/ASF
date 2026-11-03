### Compile

| Workload | stdLambda | ASF | ASF / stdLambda |
|---|---|---|---|
| Arithmetic expression | 218.500 ms | 5.375 ms | 0.02× |
| Predicate | 201.000 ms | 4.812 ms | 0.02× |
| String functions | 275.000 ms | 7.500 ms | 0.03× |
| Conditional | 296.000 ms | 8.313 ms | 0.03× |
| Local variables | 322.500 ms | 12.625 ms | 0.04× |
| Object members | 266.000 ms | 7.125 ms | 0.03× |
| Loop | 320.500 ms | 19.250 ms | 0.06× |
| While and branches | 499.000 ms | 40.500 ms | 0.08× |
| Recursion | 406.500 ms | 16.750 ms | 0.04× |
| Callback over 300 items | 230.000 ms | 14.125 ms | 0.06× |
| **Geometric mean** | | | **0.04×** |

### Run

| Workload | stdLambda | ASF | ASF / stdLambda |
|---|---|---|---|
| Arithmetic expression | 281.3 µs | 2.562 ms | 9.11× |
| Predicate | 429.7 µs | 2.187 ms | 5.09× |
| String functions | 734.4 µs | 3.062 ms | 4.17× |
| Conditional | 293.0 µs | 1.687 ms | 5.76× |
| Local variables | 304.7 µs | 2.844 ms | 9.33× |
| Object members | 773.4 µs | 3.187 ms | 4.12× |
| Loop | 732.000 ms | 4018.500 ms | 5.49× |
| While and branches | 363.000 ms | 2124.500 ms | 5.85× |
| Recursion | 277.500 ms | 3329.500 ms | 12.00× |
| Callback over 300 items | 46.500 ms | 484.500 ms | 10.42× |
| **Geometric mean** | | | **6.66×** |
