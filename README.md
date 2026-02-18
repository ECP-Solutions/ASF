# Advanced Scripting Framework (ASF)

![ASF](/docs/assets/img/ASF%20logo.png)

[![Tests (Rubberduck)](https://img.shields.io/badge/tests-Rubberduck-brightgreen)](https://rubberduckvba.com/)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/ECP-Solutions/ASF?style=plastic)](https://github.com/ECP-Solutions/ASF/releases/latest)
[![Mentioned in Awesome VBA](https://awesome.re/mentioned-badge.svg)](https://github.com/sancarn/awesome-vba)

> **JavaScript-like scripting language implemented entirely in VBA.** Zero COM dependencies. Native Office object model integration with runtime monkey patching capabilities.

```vb
Dim engine As New ASF
engine.Run engine.Compile("return [1,2,3,4,5].filter(fun(x){ return x > 2 }).reduce(fun(a,x){ return a+x }, 0)")
Debug.Print engine.OUTPUT_  ' => 12
```

---

## Official Extension

[ASF VS Code extension](https://marketplace.visualstudio.com/items?itemName=ECPSolutions.asf-language) provides syntax highlighting and language support.

:heavy_check_mark:|**Live Coding** :rocket:
:-----:|:-----:
VS Code Integration|![ASF-code](/docs/assets/img/ASF.gif)

---

## 📚 Documentation

- [API Reference](https://ecp-solutions.github.io/ASF/API.html)
- [Language Syntax](https://ecp-solutions.github.io/ASF/Syntax.html)
- [Complete Language Reference](https://ecp-solutions.github.io/ASF/Language%20reference.html)
- [COM Prototyping demo](https://ecp-solutions.github.io/ASF/examples/com-prototyping-demo.html)

## Core Features

**Language Features:**
- Full JavaScript-like syntax with lexical scoping
- First-class functions, closures, and arrow function equivalents
- ES6+ array methods (`map`, `filter`, `reduce`, `find`, etc.)
- Object/array literals with spread/rest operators
- Classes with inheritance (`extends`, `super`)
- Built-in regex engine with lookarounds
- Template literals and destructuring assignment

**VBA Integration:**
- Seamless VBA Object model integration.
- A lot of advanced function available through VBA-Expressions  via `@(...)` syntax
- Variable injection between VBA and ASF contexts
- Native array/collection marshalling
- Error containment and debugging support

**Architecture:**
- Pure VBA implementation (no external dependencies)
- Sandboxed VM execution with call stack tracing and toggle configuration for object model accessibility
- 32-bit and 64-bit Office compatibility
- Module system with import/export support

## Advanced Office Integration

### COM Object Model Extension

ASF can **dynamically extend native Office objects** at runtime using prototype-based monkey patching:

```javascript
// Extend Excel ListRow objects with custom methods
prototype.COM.ListRow asDictionary() {
    let headers = this.parent.listcolumns;
    let values = this.range.value2;
    let result = {};
    for (let i = 1, i <= headers.count, i+=1) {
        result.set(headers.item(i).name, values[1][i]);
    };
    return result;
};

// Usage: ListRows now have .asDictionary() method
let data = $1.Sheets(1).ListObjects('Table1').ListRows.map(
    fun(row) { return row.asDictionary().get('email'); }
);
```

This enables:
- **Runtime method injection** into COM objects
- **JavaScript-like method chaining** on Office objects  
- **Functional programming patterns** on traditional VBA collections
- **Context-aware `this` binding** preserving COM object references

### Collection Method Override

The `OverrideCollMethods` option replaces native VBA collection behavior with JavaScript array semantics:

```vb
' Enable collection override (USE WITH CAUTION)
engine.OverrideCollMethods = True
```

This transforms VBA Collections into JavaScript-like arrays with:
- Array iteration methods (`map`, `filter`, `forEach`)
- Immutable operations (`slice`, `concat`, `toSorted`)
- Modern array methods (`find`, `includes`, `at`)
- Destructuring and spread operator support

**⚠️ Warning:** This modifies fundamental Office object behavior and may break existing VBA code expectations.

## Quick Start (5 Minutes)

### Step 1: Installation

1. Download the [latest release](https://github.com/ECP-Solutions/ASF/releases/latest)
2. Import these class modules into your VBA project:
   - `ASF.cls` (main engine)
   - `ASF_Compiler.cls`
   - `ASF_VM.cls`
   - `ASF_Parser.cls`
   - `ASF_Globals.cls`
   - `ASF_ScopeStack.cls`
   - `ASF_Map.cls`
   - `ASF_RegexEngine.cls`
   - `UDFunctions.cls`
   - `VBAcallBack.cls`
   - `VBAexpressions.cls`
   - `VBAexpressionsScope.cls`

### Step 2: Hello World

```vb
Sub HelloASF()
    Dim engine As New ASF
    Dim idx As Long
    
    ' Compile ASF code
    idx = engine.Compile("print('Hello from ASF!')")
    
    ' Run it
    engine.Run idx
End Sub
```

### Step 3: Real-World Example

**Scenario:** Filter and sum sales data over $1000

```vb
Sub ProcessSalesData()
    Dim engine As New ASF
    Dim script As String
    
    script = _
        "sales = [850, 1200, 950, 2500, 1100, 600];" & _
        "highValue = sales.filter(fun(x) { return x > 1000 });" & _
        "total = highValue.reduce(fun(sum, item) { return sum + item }, 0);" & _
        "return total;"
    
    engine.Run engine.Compile(script)
    Debug.Print "High-value sales total: $" & engine.OUTPUT_  ' => $4800
End Sub
```

### Step 4: VBA Integration

**Call VBA-Expressions functions from ASF using `@(...)`:**

```vb
' In UDFunctions.cls
Public Function GetTaxRate(fakeArg As Variant) As Variant
    GetTaxRate = 0.08
End Function

' In your VBA code
Sub CalculateTotal()
    Dim asfGlobals As New ASF_Globals
    asfGlobals.ASF_InitGlobals
    asfGlobals.gExprEvaluator.DeclareUDF "GetTaxRate", "UserDefFunctions"
    
    Dim engine As New ASF
    engine.SetGlobals asfGlobals
    
    Dim script As String
    script = _
        "price = 100;" & _
        "tax = price * @(GetTaxRate());" & _
        "return price + tax;"
    
    engine.Run engine.Compile(script)
    Debug.Print "Total with tax: $" & engine.OUTPUT_  ' => $108
End Sub
```

**That's it!** You're now using modern scripting in VBA.

> **NOTE**: VBA function calls require a single `Variant` parameter as per [VBA-Expressions](https://github.com/ECP-Solutions/VBA-Expressions) requirements.

---

## Technical Architecture

### Compilation Pipeline

1. **Lexical Analysis:** Pure VBA tokenizer with support for ES6+ syntax
2. **Parsing:** Recursive descent parser generating AST nodes
3. **Compilation:** AST transformation with scope analysis and optimization
4. **Execution:** Stack-based virtual machine with proper `this` binding

### VM Features

- **Lexical Scoping:** Proper closure capture with environment chains
- **Prototype Chain:** JavaScript-like prototype inheritance
- **Call Stack Management:** Proper tail call handling and stack overflow protection  
- **Memory Management:** Automatic garbage collection for ASF objects
- **Error Handling:** Contained exceptions with full stack traces

## When to Use ASF

### ✅ Optimal Use Cases

- **Complex Data Processing:** ETL pipelines, JSON parsing, nested transformations
- **Dynamic Business Logic:** User-configurable rules without VBA recompilation  
- **Functional Programming:** When `map`/`filter`/`reduce` patterns improve readability
- **Office Object Extension:** Adding custom methods to native Excel/Word objects
- **Developer Productivity:** Rapid prototyping with familiar JavaScript syntax

### ⚠️ Use Native VBA For

- **Performance-Critical Code:** Tight loops processing millions of items
- **Direct COM Interop:** Extensive Office API manipulation  
- **Memory-Constrained Environments:** When every KB matters
- **Type Safety Requirements:** Compile-time checking and IntelliSense

### 🤝 Hybrid Approach

Use ASF for complex logic and VBA for performance hotspots:

```vb
' ASF for business logic
Dim businessRules As String: businessRules = "return data.filter(fun(x) { return x.status == 'active' }).map(fun(x) { return x.value * 1.08; })"
Dim processedData As Variant: processedData = engine.Run(engine.Compile(businessRules))

' VBA for bulk Excel operations
Range("A1").Resize(UBound(processedData), 1).Value2 = Application.Transpose(processedData)
```

## Debugging & Troubleshooting

### Common Issues

#### Type mismatch with arrays
```vb
' ❌ Won't work
engine.Compile("arr = " & Join(myArray, ","))

' ✅ Correct
engine.InjectVariable "arr", myArray
engine.Compile("print(arr[1])")
```

#### COM object access
```vb
' ❌ Won't work
script = "Range('A1').Value2 = 10"

' ✅ Use placeholder pseudo injection 
script = "$1.Range('A1').Value2 = 10"
pid = engine.Compile(script)
engine.Run pid, ThisWorkbook.Sheets(1)
```

## Contributing

- **Issues:** [GitHub Issues](https://github.com/ECP-Solutions/ASF/issues)
- **Pull Requests:** Include comprehensive tests
- **Documentation:** Examples and technical deep-dives welcome
- **Enterprise Support:** Tag issues with "Enterprise"

## License

**MIT License** - Copyright © 2026 W. García

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=ECP-Solutions/ASF&type=Date)](https://star-history.com/#ECP-Solutions/ASF&Date)

---

<div align="center">

**Modern scripting for VBA. Zero compromises.**

[📚 Documentation](docs/Language%20reference.md) • [🐛 Report Bug](https://github.com/ECP-Solutions/ASF/issues) • [💡 Request Feature](https://github.com/ECP-Solutions/ASF/issues) • [⭐ Star on GitHub](https://github.com/ECP-Solutions/ASF)

</div>
