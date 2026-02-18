# Advanced Scripting Framework (ASF)

![ASF](/docs/assets/img/ASF%20logo.png)

[![Tests (Rubberduck)](https://img.shields.io/badge/tests-Rubberduck-brightgreen)](https://rubberduckvba.com/)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/ECP-Solutions/ASF?style=plastic)](https://github.com/ECP-Solutions/ASF/releases/latest)
[![Mentioned in Awesome VBA](https://awesome.re/mentioned-badge.svg)](https://github.com/sancarn/awesome-vba)

> **Modern JavaScript-like scripting for VBA projects.** No COM dependencies. No migration required. Just drop in and start using `map`/`filter`/`reduce`, classes with inheritance, closures, regex, modules import/export—all inside your existing Excel, Access, or Office VBA code.

```vb
' Before: 30+ lines of VBA boilerplate with ScriptControl
' After: Clean, readable ASF
Dim engine As New ASF
engine.Run engine.Compile("return [1,2,3,4,5].filter(fun(x){ return x > 2 }).reduce(fun(a,x){ return a+x }, 0)")
Debug.Print engine.OUTPUT_  ' => 12
```

---

## Official extension
Users can now take advantage of syntax highlighting using the [official ASF VS Code extension](https://marketplace.visualstudio.com/items?itemName=ECPSolutions.asf-language). 

:heavy_check_mark:|**Check it out!** :rocket:
:-----:|:-----:
Live coding|![ASF-code](/docs/assets/img/ASF.gif)

---
## 📋 Table of Contents

- [Documentation](#documentation)
- [Why ASF?](#why-asf)
- [When to Use ASF](#when-to-use-asf)
- [Quick Start (5 Minutes)](#quick-start-5-minutes)
- [Feature Showcase](#feature-showcase)
- [Performance & Benchmarks](#performance--benchmarks)
- [ASF vs. Alternatives](#asf-vs-alternatives)
- [Language Reference](#language-reference)
- [Examples & Use Cases](#examples--use-cases)
- [Debugging & Troubleshooting](#debugging--troubleshooting)
- [Contributing](#contributing)
- [License](#license)

---

## 📚 Documentation

- [API](https://ecp-solutions.github.io/ASF/API.html)
- [Syntax](https://ecp-solutions.github.io/ASF/Syntax.html)
- [Complete Language Reference](https://ecp-solutions.github.io/ASF/Language%20reference.html)
- [Object Methods vs JavaScript Comparison](https://ecp-solutions.github.io/ASF/examples/object-methods-comparison.html)

## Why ASF?

**VBA is stuck in the 1990s.** But millions of business-critical workflows still run on it. ASF brings modern scripting ergonomics to VBA without forcing you to abandon your existing codebase. ASF isn't about making VBA into something else. It's about giving developers who already work in both web and Office environments a shared vocabulary. Also, ASF is for motivate VBA developers to learn/use modern ergonomics without exit the VBA IDE. 

### The Problem ASF Solves

**Scenario:** You need to process JSON from a web API, filter data, apply transformations, and handle complex regex patterns in Excel.

**Native VBA approach:** 
- 🔴 ScriptControl (32-bit only, COM dependency, security issues)
- 🔴 Clunky Collection/Dictionary manipulation
- 🔴 200+ lines of nested loops and string parsing
- 🔴 No first-class functions = copy-paste everywhere

**ASF approach:**
- ✅ Works in 32-bit AND 64-bit Office
- ✅ Zero external dependencies (pure VBA implementation)
- ✅ JavaScript-like syntax familiar to web developers
- ✅ 20 lines of expressive, chainable operations
- ✅ First-class functions, closures, modern array methods

### Core Benefits

| Feature | Native VBA | ASF |
|---------|-----------|-----|
| **Array filtering** | Manual loops with counters | `.filter(fun(x) { return x > 10 })` |
| **Data transformation** | ReDim + index juggling | `.map(fun(x) { return x * 2 })`. Array destructuring with rest operator support `[first, second, ...rest] = [1, 2, 3, 4, 5]`|
| **Regex support** | COM + error-prone patterns | Built-in regex engine with lookarounds |
| **JSON handling** | Parse character-by-character | Object/array literals: `{a: 1, b: [2, 3]}` |
| **Classes & inheritance** | ❌ Not inheritance support | Full OOP with `extends` and `super` |
| **Anonymous functions** | ❌ Not supported | `fun(x) { return x * 2 }` |
| **64-bit compatibility** | ScriptControl breaks | ✅ Works everywhere |
| **Security** | Late-bound COM risks | Sandboxed VM execution |

---

## When to Use ASF

### ✅ Perfect For

- **Complex data transformations** - ETL workflows, data cleaning, report generation
- **JSON/API integration** - Parse responses, build requests, handle nested structures
- **User-defined formulas** - Let users safely write custom logic without VBA access
- **Config-driven workflows** - Business rules that change without recompiling
- **Regex-heavy tasks** - Pattern matching, text extraction, validation
- **Functional programming patterns** - When `map`/`filter`/`reduce` beats loops
- **Web developer transitioning to Office** - Familiar JavaScript-like syntax

### ⚠️ Use Native VBA When

- Simple macros (e.g., `Range("A1").Value = 10`)
- Performance-critical tight loops (millions of iterations)
- Interacting with COM or application host objects extensively
- You need VBA's compile-time type checking

### 🤝 Use BOTH When

- Complex business logic + Excel automation (ASF for logic an non heavy API, VBA for speed boost)
- You need rapid iteration on algorithms but final performance matters (prototype in ASF, optimize hotspots in VBA)

**Think of ASF like Excel's `LAMBDA`:** Not a replacement for formulas, but a powerful addition when complexity demands it.

---

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

**Call VBA functions from ASF using `@(...)`:**

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

> **NOTE**: the calls to VBA is intentionally conditioned to string data type. Each UDF MUST be declared with ONLY one `Variant` argument as per [VBA-Expressions rules](https://github.com/ECP-Solutions/VBA-Expressions?tab=readme-ov-file#user-defined-functions-udf). This aligns perfectly with ASF's sandbox VM execution. 

---

## Feature Showcase

### 🎯 Before/After Comparisons

#### Regex Matching

<details>
<summary><strong>❌ Native VBA (40+ lines with ScriptControl)</strong></summary>

```vb
Private Static Function JsRmatch(Pattern As String, Modifiers As String, InputData As String, ReturnVal As Variant) As Boolean
'Matches regular expression using Javascript and returns array of match and submatches
    
    Dim Scontrol As Object
    Dim Mcount As Long
    Dim I As Long
    Const Script As String = "var matches;" & _
        "  function reMatch(thisPattern, thisModifiers, thisData) {" & _
        "  var regex = new RegExp(thisPattern, thisModifiers);" & _
        "  matches = regex.exec(thisData);" & _
        "  if(matches) {" & _
        "    return matches.length;" & _
        "  } else {" & _
        "    return 0;" & _
        " }}"
        
    If Scontrol Is Nothing Then
        Set Scontrol = CreateObject("ScriptControl")
        With Scontrol
            .Language = "JScript"
            .AddCode Script
        End With
    End If

    With Scontrol
        On Error Resume Next
        Mcount = .Run("reMatch", Pattern, Modifiers, InputData)
        If Err = 0 Then
            On Error GoTo 0
        Else
            On Error GoTo 0
            MsgBox "Error matching regular expression " & Pattern
            ReturnVal = Array()
            Exit Function
        End If
        
        If Mcount = 0 Then
            ReturnVal = Array()
        Else
            ReDim ReturnVal(0 To Mcount - 1) As String
            For I = 0 To Mcount - 1
                ReturnVal(I) = Scontrol.Eval("matches[" & I & "]")
            Next I
        End If
    End With
    JsRmatch = True
End Function
```
</details>

**✅ With ASF (13 lines)**

```vb
Private Function JsRmatch(pattern As String, InputData As String, ReturnVal As Variant) As Boolean
    Dim engine As ASF
    Dim pidx As Long

    Set engine = New ASF
    With engine
        .InjectVariable "inputData", InputData
        .InjectVariable "pattern", pattern
        pidx = .Compile("return(inputData.match(pattern));")
        .Run pidx
        JsRmatch = IsArray(.OUTPUT_)
        ReturnVal = .OUTPUT_
    End With
End Function
```

**Result:** 69% less code. No COM dependency. Works in 64-bit. Easier to maintain.

---

#### Array Filtering & Transformation

<details>
<summary><strong>❌ Native VBA (20+ lines)</strong></summary>

```vb
Function ProcessScores(scores As Variant) As Double
    Dim filtered() As Double
    Dim count As Long
    Dim i As Long
    Dim total As Double
    
    ' Filter scores > 70
    count = 0
    ReDim filtered(1 To UBound(scores))
    For i = LBound(scores) To UBound(scores)
        If scores(i) > 70 Then
            count = count + 1
            filtered(count) = scores(i)
        End If
    Next i
    ReDim Preserve filtered(1 To count)
    
    ' Double each score
    For i = 1 To count
        filtered(i) = filtered(i) * 2
    Next i
    
    ' Sum
    total = 0
    For i = 1 To count
        total = total + filtered(i)
    Next i
    
    ProcessScores = total
End Function
```
</details>

**✅ With ASF (6 lines)**

```vb
Function ProcessScores(scores As Variant) As Double
    Dim engine As New ASF
    engine.InjectVariable "scores", scores
    engine.Run engine.Compile("return scores.filter(fun(x){return x>70}).map(fun(x){return x*2}).reduce(fun(a,x){return a+x},0)")
    ProcessScores = engine.OUTPUT_
End Function
```

---

### 🚀 Advanced Features

#### 1. Classes with Inheritance & Polymorphism

```js
class Animal {
    field name, species;
    
    constructor(name, species) {
        this.name = name;
        this.species = species;
    }
    
    speak() {
        return this.name + ' makes a sound';
    }
}

class Dog extends Animal {
    field breed;
    
    constructor(name, breed) {
        super(name, 'Dog');
        this.breed = breed;
    }
    
    speak() {
        return this.name + ' barks!';
    }
}

let dog = new Dog('Rex', 'Labrador');
print(dog.speak());  // => 'Rex barks!'
print(dog.species);  // => 'Dog'
```

#### 2. Closures & First-Class Functions

```js
fun makeCounter() {
    let count = 0;
    return fun() {
        count = count + 1;
        return count;
    };
}

let counter = makeCounter();
print(counter());  // => 1
print(counter());  // => 2
print(counter());  // => 3
```

#### 3. Template Literals

```js
let user = { name: 'Alice', score: 95 };
let message = `Hello ${user.name}, your score is ${user.score}!`;
print(message);  // => 'Hello Alice, your score is 95!'
```

#### 4. Advanced Regex with Lookarounds

```js
// Extract numbers followed by 'px'
let text = 'width: 100px, height: 50px, margin: 20em';
let pixels = text.match(`/(\d+)(?=px)/g`);
print(pixels);  // => ['100', '50']
```

#### 5. Object & Array Chaining

```js
let data = {
    users: [
        { name: 'Alice', age: 25, active: true },
        { name: 'Bob', age: 30, active: false },
        { name: 'Charlie', age: 35, active: true }
    ]
};

let activeNames = data.users
    .filter(fun(u) { return u.active })
    .map(fun(u) { return u.name })
    .join(', ');

print(activeNames);  // => 'Alice, Charlie'
```

#### 6. Spread/rest operator support

```js
//rest argument
fun greetAll(greeting, ...names) {
    result = greeting + ': ';
    for (name of names) {
        result = result + name + ', ';
    };
    return result.slice(0, -2);
};
msg = greetAll('Hello', 'Alice', 'Bob', 'Charlie');
return msg //=> Hello: Alice, Bob, Charlie
```

```js
//spread operator on arrays
begin = [1]; middle = [2, 3, 4]; end = [5]; combined = [...begin, ...middle, ...end]; 
print(combined); //=> [ 1, 2, 3, 4, 5 ]

//spread operator on objects
obj1 = {a: 1, b: 2}; obj2 = {c: 3, ...obj1, d: 4};
return `${obj2.a}; ${obj2.b}; ${obj2.c}; ${obj2.d}` //=> 1; 2; 3; 4
```

#### 7. Array destructuring

```js
[first, second, ...rest] = [1, 2, 3, 4, 5];
return `${rest}` //=> [ 3, 4, 5 ]
```

#### 8. Module system

Starting with `ASFv3.0.0`, users can perform secure, shared script invocation without exposing raw VBA host access through safe `import`/`export` keywords. For example, place this code in a module named `math.vas`

```js
// Named exports
fun add(a, b) {
    return a + b;
};

fun multiply(a, b) {
    return a * b;
};

PI = 3.14159;

export { add, multiply, PI };
```

In a file named `main_math.vas` place this code

```js
scwd(wd); 
import { add, multiply, PI } from './math.vas';
result = add(5, 3);
area = PI * multiply(5, 5);
return `5 + 3 = ${result}, Circle area: ${area}`;// => 78.53975
```

Next, in the VBA IDE, place this code in a new Excel workbook and save it in the same folder as `math.vas` and `main_math.vas` files

```vb
Private Sub math_test()
    Dim wd As String
    Dim eng As New ASF
	Dim result As String
	
    wd = ThisWorkbook.path
    With eng
        .InjectVariable "wd", wd
        result = CStr(.Execute(wd & "\main_math.vas"))
    End With
End Sub
```

After execution, the code returns

```
5 + 3 = 8, Circle area: 78.53975
```

For more details about ASF's module system visit the [documentation](https://ecp-solutions.github.io/ASF/Language%20reference.html#module-system).

---

## Performance & Benchmarks

ASF runs on a VM, so there's overhead vs. native VBA. Here's what you need to know:

### When Performance Matters

**❌ Don't use ASF for:**
- Tight loops processing millions of cells
- Real-time event handlers firing 100x/second
- Pixel-level graphics manipulation

**✅ ASF overhead is negligible for:**
- Data import/export workflows
- Report generation (hundreds of records)
- User-triggered automation
- API request/response handling
- Config-driven business logic

### Performance Tips

```js
// ❌ Slow: Repeated property access in loop
for (let i = 1, i <= arr.length, i += 1) {
    // arr.length evaluated 10k times
}

// ✅ Fast: Cache length
let len = arr.length;
for (let i = 1, i <= len, i += 1) {
    // len evaluated once
}

// ❌ Slow: Building arrays element by element
let result = [];
for (let i = 1, i <= 1000, i += 1) {
    result.push(data[i]);
}

// ✅ Fast: Use slice for ranges
let result = data.slice(1, 1000);
```

**Bottom line:** ASF adds overhead for typical tasks. If your workflow takes seconds anyway, ASF's expressiveness is worth it.

---

## ASF vs. Alternatives

### ASF vs. ScriptControl (MSScriptControl.ScriptControl)

| Feature | ScriptControl | ASF |
|---------|--------------|-----|
| 64-bit support | ❌ 32-bit only | ✅ Works everywhere |
| COM dependency | ❌ Yes (security risk) | ✅ Pure VBA |
| Sandboxed execution | ❌ Full system access | ✅ Controlled via VBA |
| Modern features | ❌ ES3-era JavaScript | ✅ Classes, closures, etc. |
| Maintained | ❌ Deprecated | ✅ Active development |
| Installation | ❌ Requires registration | ✅ Drop-in library |

**Verdict:** If you're using ScriptControl, switch to ASF immediately.

---

### ASF vs. twinBASIC

| Feature | twinBASIC | ASF |
|---------|-----------|-----|
| Approach | Compile-time modern VB | Runtime scripting layer |
| Migration required | ✅ Yes (port codebase) | ❌ No (drop-in library) |
| OOP features | ✅ Full VB.NET-like | ✅ Classes + inheritance |
| Use in locked files | ❌ Requires recompilation | ✅ Works in locked Excel |
| Functional programming | ⚠️ Limited | ✅ First-class functions, closures |
| Best for | New projects, greenfield | Legacy codebases, add-ins |

**Verdict:** twinBASIC and ASF are complementary. Use twinBASIC for complete rewrites, ASF for enhancing existing VBA.

---

### ASF vs. Python in Excel

| Feature | Python in Excel | ASF |
|---------|----------------|-----|
| Office version | Microsoft 365 only | ✅ Office 2007+ |
| Subscription required | ✅ Yes | ❌ No |
| Offline use | ⚠️ Limited | ✅ Full offline |
| VBA integration | ⚠️ Indirect | ✅ Native via `@(...)` |
| Cell formulas | ✅ Directly in cells | ❌ VBA macros only |
| Learning curve | Python syntax | JavaScript-like syntax |

**Verdict:** Python in Excel is for calculations in worksheets. ASF is for automation, workflows, and scripting VBA behavior.

---

### Decision Matrix

```
Choose ASF if:
├─ You have existing VBA code to enhance
├─ You need modern syntax without migration
├─ Your Office version is older than Microsoft 365
├─ You want functional programming patterns
└─ You're building add-ins or automation tools

Choose twinBASIC if:
├─ You're starting a new project from scratch
├─ You can rewrite your entire VBA codebase
├─ You need full VB.NET-like features
└─ Compile-time performance is critical

Choose Python in Excel if:
├─ You only need worksheet calculations
├─ You have Microsoft 365 subscription
├─ Your team knows Python but not VBA
└─ You're not building automation macros

Use Native VBA if:
├─ Simple, straightforward tasks
├─ Heavy Excel object model interaction
└─ No complex data transformations needed
```

---

## Language Reference

See the comprehensive [Language Reference](docs/Language%20reference.md) for complete syntax documentation.

### Quick Syntax Guide

```js
// Variables
let x = 10;
x += 5;

// Arrays
let arr = [1, 2, 3, 4, 5];
arr.push(6);
arr.filter(fun(x) { return x > 3 });  // => [4, 5, 6]

// Objects
let person = { name: 'John', age: 30 };
person.email = 'john@example.com';
person.keys();  // => ['name', 'age', 'email']

// Functions
fun greet(name) {
    return 'Hello, ' + name + '!';
}

// Anonymous functions
let square = fun(x) { return x * x };

// Control flow
if (x > 10) {
    print('Big');
} elseif (x > 5) {
    print('Medium');
} else {
    print('Small');
}

// Loops
for (let i = 0, i < 5, i += 1) {
    print(i);
}

for (let val of arr) {
    print(val);
}

// Classes
class Rectangle {
    field width = 0, height = 0;
    
    constructor(w, h) {
        this.width = w;
        this.height = h;
    }
    
    getArea() {
        return this.width * this.height;
    }
}

// Template literals
let name = 'Alice';
let greeting = `Hello, ${name}!`;

// Regex
let re = regex(`/\d+/g`);
let numbers = 'abc123def456'.match(re);  // => ['123', '456']

// Try-catch
try {
    riskyOperation();
} catch {
    print('Error occurred');
}

// VBA integration
let taxRate = @(GetTaxRate());  // Calls VBA function
```

---

## Examples & Use Cases

### Real-World Example: JSON API Response Processing

```vb
Sub ProcessAPIResponse()
    Dim engine As New ASF
    Dim jsonResponse As String
    
    ' Simulate API response
    jsonResponse = _
        "{" & _
        "  users: [" & _
        "    { id: 1, name: 'Alice', sales: 15000, active: true }," & _
        "    { id: 2, name: 'Bob', sales: 8000, active: false }," & _
        "    { id: 3, name: 'Charlie', sales: 22000, active: true }" & _
        "  ]" & _
        "};"
    
    Dim script As String
    script = _
        "let response = " & jsonResponse & _
        "let topSellers = response.users" & _
        "  .filter(fun(u) { return u.active && u.sales > 10000 })" & _
        "  .map(fun(u) { return { name: u.name, bonus: u.sales * 0.1 } })" & _
        "  .sort(fun(a, b) {" & _
        "    if (a.bonus > b.bonus) { return -1 };" & _
        "    if (a.bonus < b.bonus) { return 1 };" & _
        "    return 0;" & _
        "  });" & _
        "print(topSellers); return topSellers;"
    
    engine.Run engine.Compile(script)
    ' Output for further processing
    Dim result As Variant
    result = engine.OUTPUT_
    ' result => [{ name: 'Charlie', bonus: 2200 }, { name: 'Alice', bonus: 1500 }]
End Sub
```

### Example: User-Defined Formula Engine

```vb
Function EvaluateUserFormula(formula As String, inputs As Variant) As Variant
    Dim engine As New ASF
    
    ' Inject user inputs
    engine.InjectVariable "x", inputs(0)
    engine.InjectVariable "y", inputs(1)
    
    ' User provides formula like: "(x * 2) + (y ^ 2)"
    engine.Run engine.Compile("return " & formula)
    EvaluateUserFormula = engine.OUTPUT_
End Function

' Usage
Debug.Print EvaluateUserFormula("(x * 2) + (y ^ 2)", Array(5, 3))  ' => 19
```

More examples in the [`examples/`](/examples/) directory.

---

## Debugging & Troubleshooting

### Enable Call Stack Tracing

```vb
Dim engine As New ASF
engine.EnableCallTrace = True

Dim script As String
script = _
    "fun add(a, b) { return a + b; }" & _
    "fun multiply(a, b) { return a * b; }" & _
    "x = add(3, 4);" & _
    "y = multiply(x, 3);" & _
    "print(y)"

engine.Run engine.Compile(script)

' Print call stack trace
Debug.Print "=== Call Stack Trace ==="
Debug.Print engine.GetCallStackTrace()
```

**Output:**
```
=== Call Stack Trace ===
CALL: add(3, 4) -> 7
CALL: multiply(7, 3) -> 21
```

### Reading Scripts from Files

VBA has a limit on line continuations. For large scripts, use:

```vb
Dim script As String
script = ASF.ReadTextFile("C:\path\to\script.asf")

Dim engine As New ASF
engine.Run engine.Compile(script)
```

### Common Issues

#### Issue: "Type mismatch" when passing arrays

**Problem:** VBA arrays don't automatically convert to ASF arrays.

**Solution:** Use `InjectVariable`:

```vb
' ❌ Won't work
engine.Compile("arr = " & Join(myArray, ","))

' ✅ Correct
engine.InjectVariable "arr", myArray
engine.Compile("print(arr[1])")
```

---

#### Issue: "Object doesn't support this property or method"

**Problem:** Trying to use VBA objects directly in ASF.

**Solution:** Use `@(...)` to call VBA functions:

```vb
' ❌ Won't work
script = "Range('A1').Value = 10"

' ✅ Correct - call VBA wrapper
Public Function SetCellValue(args As Variant) As Variant
    Dim lb As Long: lb = LBound(args)
    Range(args(lb)).Value = args(lb+1)
    SetCellValue = True
End Function

' In UDFunctions.cls, then:
script = "@(SetCellValue('A1', 10))"
```

---

#### Issue: Slow performance

**Problem:** Processing thousands of cells in tight loop.

**Solution:** Batch operations or use native VBA for critical sections:

```js
// ❌ Slow - processes cells one by one
for (let i = 1, i <= 10000, i += 1) {
    @(ProcessCell(i))  // VBA call in loop
}

// ✅ Fast - batch data, process once
let data = @(GetRangeData());  // Get all at once
let processed = data.map(fun(x) { return x * 2 });
@(WriteRangeData(processed));  // Write all at once
```

---

### Reporting Bugs

When reporting issues, please include:

1. **ASF version** (check releases)
2. **Office version** (Excel 2016, 2019, 365, etc.)
3. **Minimal reproducible example**
4. **Expected vs. actual behavior**
5. **Error messages or stack traces**

[Open an issue on GitHub](https://github.com/ECP-Solutions/ASF/issues)

---

## Running the Test Suite

1. Open `ASF [VERSION NAME].xlsm` workbook
2. Install [Rubberduck VBA](https://rubberduckvba.com/)
3. In Rubberduck, navigate to the test explorer
4. Run all tests - they should pass ✅

The test suite covers:
- Arithmetic & operators
- Control flow (if/for/while/switch)
- Functions & closures
- Arrays & objects (literals, methods, chaining)
- Classes & inheritance
- String manipulation
- Spread/rest operator
- Regex engine
- VBA integration
- Error handling

---

## Contributing

### How to Contribute

- **Read the [contributing](/CONTRIBUTING.md) guide**
- **Report bugs** via [GitHub Issues](https://github.com/ECP-Solutions/ASF/issues)
- **Request features** with clear use cases
- **Submit PRs** with tests covering behavior changes
- **Improve docs** - examples, tutorials, translations
- **Share your ASF projects** - help others learn

### Community

- ⭐ Star this repo if ASF helps you!
- 💬 Discuss on [r/vba](https://reddit.com/r/vba)
- 📧 Enterprise inquiries: open an issue with "Enterprise" tag

---

## License

**MIT License**

Copyright © 2026 W. García

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

---

## Acknowledgments

- Built with inspiration from JavaScript, Python, and modern language design
- Powered by [VBA-Expressions](https://github.com/ECP-Solutions/VBA-Expressions)
- Tested with [Rubberduck VBA](https://rubberduckvba.com/)
- Community feedback from [r/vba](https://reddit.com/r/vba)

---

<div align="center">

**Made with ❤️ for the VBA community**

[📚 Documentation](docs/Language%20reference.md) • [🐛 Report Bug](https://github.com/ECP-Solutions/ASF/issues) • [💡 Request Feature](https://github.com/ECP-Solutions/ASF/issues) • [⭐ Star on GitHub](https://github.com/ECP-Solutions/ASF)

</div>

---

## FAQ

### Can ASF replace all my VBA code?

No, and that's not the goal. ASF complements VBA by handling complex logic, data transformations, and dynamic workflows while VBA continues to handle Excel/Office API interactions. Think of ASF as adding `LAMBDA` functions to your VBA—not replacing it, enhancing it.

### Is ASF production-ready?

Yes. ASF has a comprehensive test suite and is being used in real-world applications. However, as with any tool, test thoroughly in your specific environment before deploying to production.

### How does ASF handle errors in user-provided scripts?

ASF uses `try/catch` blocks for error handling. Errors are contained within the VM and won't crash your VBA application. You can enable call tracing with `EnableCallTrace = True` to debug issues.

### Can I use ASF in protected/locked workbooks?

Yes! Unlike solutions that require recompilation (like twinBASIC), ASF works as a runtime library in locked Excel files, making it perfect for add-ins and distributed tools.

### Does ASF work with Excel Online / Office Web Apps?

No. ASF requires the VBA runtime, which is only available in desktop Office applications (Windows/Mac). For web-based solutions, consider server-side alternatives.

### Can I contribute to ASF development?

Absolutely! We welcome contributions via pull requests. Please include tests for any new features or bug fixes. See the Contributing section above.

### How is ASF different from writing JavaScript in Office Add-ins?

Office Add-ins run in a separate JavaScript context and interact with Office through APIs. ASF runs directly in the VBA runtime, allowing seamless integration with existing VBA code, no web server required, and works offline.

### Is there commercial support available?

ASF is open-source under MIT license. For enterprise integration assistance or custom development, open a GitHub issue tagged "Enterprise" to discuss your needs.

### Can ASF access the Windows filesystem, registry, or external DLLs?

ASF itself is sandboxed within the VBA environment. To access external resources, create VBA wrapper functions and call them via `@(...)` syntax. This keeps ASF scripts secure while allowing controlled access to system features when needed.

---

## Testimonials

> "It's a very interesting project, like woodworking equivalent of coding." — *Excel Automation Developer*

> "It's mostly already working in tB, plus or minus a tB bug and missing feature." — *Developer*

*Have you used ASF in a project? Share your experience via GitHub issues!*

---

## Star History

If ASF has helped you, consider starring the repository to help others discover it!

[![Star History Chart](https://api.star-history.com/svg?repos=ECP-Solutions/ASF&type=Date)](https://star-history.com/#ECP-Solutions/ASF&Date)

---

## What's Next?

Now that you understand what ASF can do, here are some next steps:

1. **[Install ASF](#quick-start-5-minutes)** and try the Hello World example
2. **[Read the Language Reference](docs/Language%20reference.md)** for complete syntax documentation
3. **[Explore Examples](examples/)** to see real-world use cases
4. **Join the Discussion** on [r/vba](https://reddit.com/r/vba) - search for "ASF"
5. **Star the Repository** to follow development and help others discover ASF

---

**Happy Scripting! 🚀**
