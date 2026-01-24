---
layout: default
title: Language Documentation
has_children: false
nav_order: 4
description: "ASF Language Documentation."
---

# ASF Language Documentation

Version 1.0 | Complete Language Reference

---

## Table of Contents

1. [Introduction](#introduction)
2. [Getting Started](#getting-started)
3. [Basic Syntax](#basic-syntax)
4. [Data Types](#data-types)
5. [Variables](#variables)
6. [Operators](#operators)
7. [Control Flow](#control-flow)
8. [Functions](#functions)
9. [Arrays](#arrays)
10. [Objects](#objects)
11. [Object Member Methods](#object-member-methods)
12. [Classes](#classes)
13. [Built-in Functions](#built-in-functions)
14. [String Methods](#string-methods)
15. [Array Methods](#array-methods)
16. [Template Literals](#template-literals)
17. [Regular Expressions](#regular-expressions)
18. [Error Handling](#error-handling)
19. [VBA Integration](#vba-integration)
20. [Best Practices](#best-practices)

---

## Introduction

**ASF (Advanced Scripting Framework)** is a JavaScript-like scripting language implemented in VBA (Visual Basic for Applications). It provides modern programming features within Excel, Access, and other Office applications.

### Key Features

- **JavaScript-like syntax** - Familiar syntax for web developers
- **Object-oriented programming** - Classes with inheritance
- **Functional programming** - First-class functions and closures
- **Modern array methods** - map, filter, reduce, and more
- **Template literals** - String interpolation with backticks
- **Regular expressions** - Pattern matching support
- **VBA integration** - Seamless integration with VBA code

### Why ASF?

- Write more expressive code in Office applications
- Leverage JavaScript patterns in VBA environment
- Build complex logic with modern language features
- Share code logic between web and Office platforms

---

## Getting Started

### Installation

1. Import the ASF class modules into your VBA project:
   - `ASF.cls`
   - `ASF_Compiler.cls`
   - `ASF_Globals.cls`
   - `ASF_Map.cls`
   - `ASF_Parser.cls`
   - `ASF_ScopeStack.cls`
   - `ASF_VM.cls`
   - `ASF_RegexEngine.cls`
   - `UDFunctions.cls`
   - `VBAcallBack.cls`
   - `VBAexpressions.cls`
   - `VBAexpressionsScope.cls`

### Hello World

```vba
Sub HelloWorld()
    Dim engine As New ASF
    Dim code As String
    
    code = "print('Hello, World!');"
    
    Dim idx As Long
    idx = engine.Compile(code)
    engine.Run idx
End Sub
```

### Basic Usage Pattern

```vba
Sub RunASFCode()
    ' Create engine instance
    Dim engine As New ASF
    
    ' Write ASF code
    Dim code As String
    code = "let x = 10; print(x * 2);"
    
    ' Compile and run
    Dim idx As Long
    idx = engine.Compile(code)
    engine.Run idx
    
    ' Get output (if function returns a value)
    Dim result As Variant
    result = engine.OUTPUT_
End Sub
```

---

## Basic Syntax

### Comments

```javascript
// Single-line comment

/* Multi-line
   comment */

# Python-style comment (also supported)
```

### Statements

Statements are terminated by semicolons (`;`):

```javascript
let x = 10;
print(x);
```

Semicolons are optional at the end of the script but required at end of statements blocks:

```javascript
if (x > 5) {
    print('Greater');
};  // Semicolon mandatory here

let y = 20;  // Semicolon optional here
```

### Case Sensitivity

ASF is case-sensitive:

```javascript
let myVar = 10;
let MyVar = 20;  // Different variable
```

### Whitespace

Whitespace is generally ignored:

```javascript
let x=10;        // Valid
let y = 20;      // More readable
```

---

## Data Types

### Primitive Types

#### Number

All numbers are floating-point:

```javascript
let integer = 42;
let decimal = 3.14159;
let negative = -100;
let scientific = 1.5e10;
```

#### String

Strings are enclosed in single quotes:

```javascript
let name = 'John Doe';
let message = 'Hello, World!';
let empty = '';
```

Template literals and regex patterns use backticks:

```javascript
let name = 'Alice';
let greeting = `Hello, ${name}!`;  // "Hello, Alice!"
let arr = 'test1test2'.match(`/t(e)(st(\d?))/g`)
```

#### Boolean

```javascript
let isTrue = true;
let isFalse = false;
```

#### Null

Represents intentional absence of value:

```javascript
let nothing = null;
```

### Composite Types

#### Array

```javascript
let numbers = [1, 2, 3, 4, 5];
let mixed = [1, 'two', true, null];
let nested = [[1, 2], [3, 4]];
let empty = [];
```

#### Object

```javascript
let person = {
    name: 'John',
    age: 30,
    email: 'john@example.com'
};

let nested = {
    user: {
        name: 'Alice',
        address: {
            city: 'Boston'
        }
    }
};
```

### Type Checking

```javascript
typeof 42;           // 'number'
typeof 'hello';      // 'string'
typeof true;         // 'boolean'
typeof null;         // 'null'
typeof [];           // 'array'
typeof {};           // 'object'
typeof fun() {};     // 'function'

// Built-in functions
isArray([1, 2, 3]);  // true
isNumeric(42);       // true
isNumeric('42');     // true
isNumeric('hello');  // false
```

---

## Variables

### Declaration

Variables do not need to be declared; in any case, the interpreter supports the use of the `let` keyword to assign variables:

```javascript
x = 10;
let y = 20;      // Also valid (converted to simple assignment)
```

### Scope

Functions shares scope variables, outer variables can be mutated:

```javascript
let x = 10;

fun test() {
    let x = 20;  /* Same variable*/
    print(x)    /* Outputs: 20 */
};

test();
print(x);        // Outputs: 20
```

### Assignment

```javascript
let x = 10;
x = 20;          /* Reassignment also accepts let x = 20 (does not behave like JavaScript)*/

let arr = [1, 2, 3];
arr[0] = 10;     /* Array element assignment */

let obj = { name: 'John' };
obj.name = 'Jane';  /* Property assignment */
```

### Compound Assignment

```javascript
let x = 10;
x += 5;    // x = x + 5  (15)
x -= 3;    // x = x - 3  (12)
x *= 2;    // x = x * 2  (24)
x /= 4;    // x = x / 4  (6)
x %= 4;    // x = x % 4  (2)
x ^= 3;    // x = x ^ 3  (8)
x &= 7;    // x = x & 7  (string concat)
x |= 2;    // x = x | 2  (bitwise OR)
```

---

## Operators

### Arithmetic Operators

```javascript
let a = 10, b = 3;

a + b;    // 13 (addition)
a - b;    // 7  (subtraction)
a * b;    // 30 (multiplication)
a / b;    // 3.333... (division)
a % b;    // 1  (modulus/remainder)
a ^ b;    // 1000 (exponentiation)
```

### Comparison Operators

```javascript
let x = 10, y = 20;

x == y;   // false (equal)
x != y;   // true  (not equal)
x < y;    // true  (less than)
x > y;    // false (greater than)
x <= y;   // true  (less than or equal)
x >= y;   // false (greater than or equal)
```

### Logical Operators

```javascript
let a = true, b = false;

a && b;   // false (AND)
a || b;   // true  (OR)
!a;       // false (NOT)
```

### String Concatenation

```javascript
'Hello' + ' ' + 'World';  // 'Hello World'
'Value: ' + 42;           // 'Value: 42'
'Count' & ': ' & 10;      // 'Count: 10' (alternative)
```

### Bitwise Operators

```javascript
let x = 5;   // Binary: 101
let y = 3;   // Binary: 011

x << 1;      // 10 (left shift)
x >> 1;      // 2  (right shift)
```

Compound bitwise assignment:

```javascript
x <<= 2;     // x = x << 2
x >>= 1;     // x = x >> 1
```

### Ternary Operator

```javascript
let age = 18;
let status = (age >= 18) ? 'adult' : 'minor';
print(status);  // 'adult'

// Nested ternary
let score = 85;
let grade = (score >= 90) ? 'A' :
            (score >= 80) ? 'B' :
            (score >= 70) ? 'C' : 'F';
```

### Operator Precedence

From highest to lowest:

1. Parentheses `( )`
2. Unary `!`, `-`, `typeof`
3. Exponentiation `^`
4. Multiplication/Division `*`, `/`, `%`
5. Addition/Subtraction `+`, `-`
6. Bitwise Shift `<<`, `>>`
7. Comparison `<`, `>`, `<=`, `>=`
8. Equality `==`, `!=`
9. Logical AND `&&`
10. Logical OR `||`
11. Ternary `? :`
12. Assignment `=`, `+=`, `-=`, etc.

---

## Control Flow

### If Statement

```javascript
let x = 10;

if (x > 5) {
    print('Greater than 5');
};

if (x > 15) {
    print('Greater than 15');
} else {
    print('Not greater than 15');
};

// Multiple conditions
if (x > 20) {
    print('Greater than 20');
} elseif (x > 10) {
    print('Greater than 10');
} elseif (x > 5) {
    print('Greater than 5');
} else {
    print('5 or less');
};
```

### Switch Statement

```javascript
let day = 3;

switch (day) {
    case 1 {
        print('Monday');
    }
    case 2 {
        print('Tuesday');
    }
    case 3 {
        print('Wednesday');
    }
    default {
        print('Other day');
    };
};
```

**Note:** ASF switch statements don't fall through - no `break` needed.

### For Loop

#### Classic For Loop

```javascript
// Standard C-style for loop 
for (let i = 0, i < 5, i += 1) {
    print(i);
};

```

#### For-In Loop (Iterate Indices/Keys)

```javascript
// Array indices
let arr = [10, 20, 30];
for (let i in arr) {
    print(i);  /* 1, 2, 3 (indices, 1-based) */
};

// Object keys
let obj = { name: 'John', age: 30 };
for (let key in obj) {
    print(key);  /* 'name', 'age' */
};

// String indices
let str = 'ABC';
for (let i in str) {
    print(i);  // 1, 2, 3
};
```

#### For-Of Loop (Iterate Values)

```javascript
// Array values
let arr = [10, 20, 30]
for (let val of arr) {
    print(val);  /* 10, 20, 30 */
};

// String characters
let str = 'ABC';
for (let char of str) {
    print(char);  /* 'A', 'B', 'C' */
};

// Object values
let obj = { name: 'John', age: 30 };
for (let val of obj) {
    print(val);  /* 'John', 30 */
};
```

### While Loop

```javascript
let i = 0;
while (i < 5) {
    print(i);
    i = i + 1;
}

// Infinite loop with break
let count = 0;
while (true) {
    if (count >= 10) {
        break;
    };
    count = count + 1;
}; print(count); //-->10
```

### Break and Continue

```javascript
// Break: exit loop
for (let i = 0, i < 10, i += 1) {
    if (i == 5) {
        break;  // Exit loop when i is 5
    };
    print(i);
};

// Continue: skip to next iteration
for (let i = 0, i < 10, i += 1) {
    if (i % 2 == 0) {
        continue;  // Skip even numbers
    };
    print(i);  // Prints odd numbers only
};
```

---

## Functions

### Function Declaration

```javascript
fun greet(name) {
    print('Hello, ' + name + '!');
};

greet('Alice');  // Hello, Alice!
```

### Function with Return Value

```javascript
fun add(a, b) {
    return a + b;
};

let result = add(5, 3);  // 8
```

### Function Expressions

```javascript
let square = fun(x) {
    return x * x;
};

print(square(5));  // 25
```

### Anonymous Functions

```javascript
let numbers = [1, 2, 3, 4, 5];

// Anonymous function in map
let squared = numbers.map(fun(x) {
    return x * x;
}); //--> [ 1, 4, 9, 16, 25 ]
```

### Closures

Functions can capture variables from their enclosing scope:

```javascript
fun makeCounter() {
	let count = 0;
	return fun() {
		count = count + 1;
		return count;
	};
};
let counter = makeCounter();
print(counter());  // 1
print(counter());  // 2
print(counter());  // 3
```

### Higher-Order Functions

Functions that accept or return other functions:

```javascript
fun operate(a, b, operation) {
    return operation(a, b)
};

let add = fun(x, y) { return x + y };
let multiply = fun(x, y) { return x * y };

print(operate(5, 3, add));       // 8
print(operate(5, 3, multiply));  // 15
```

### Recursion

```javascript
fun factorial(n) {
    if (n <= 1) {
        return 1;
    };
    return n * factorial(n - 1);
};

print(factorial(5));  // 120

// Fibonacci
fun fib(n) {
    if (n <= 1) {
        return n;
    };
    return fib(n - 1) + fib(n - 2);
};

print(fib(10));  // 55
```

### Default Parameters (Workaround)

```javascript
fun greet(name) {
    if (typeof name == 'undefined') {
        name = 'Guest';
    };
    print('Hello, ' + name + '!');
};

greet();          // Hello, Guest!
greet('Alice');   // Hello, Alice!
```

### Variable Arguments (Workaround)

```javascript
// Using array as parameter
fun sum(numbers) {
    let total = 0;
    for (let i = 1, i <= numbers.length, i += 1) {
        total += numbers[i];
    };
    return total;
};

print(sum([1, 2, 3, 4]));  // 10
```

---

## Arrays

### Creating Arrays

```javascript
let empty = [];
let numbers = [1, 2, 3, 4, 5];
let mixed = [1, 'two', true, null, [5, 6]];
```

### Array Indexing

**Important:** ASF uses 1-based indexing by default (EXPERIMENTAL: configurable with `option base`).

```javascript
let arr = [10, 20, 30, 40];

// 1-based indexing (default)
print(arr[1]);  // 10 (first element)
print(arr[4]);  // 40 (last element)

// Assignment
arr[2] = 25;
print(arr[2]);  // 25
```

### Array Properties

```javascript
let arr = [1, 2, 3, 4, 5];
print(arr.length);  // 5
```

### Nested Arrays

```javascript
let matrix = [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9]
];

print(matrix[1][1]);  // 1
print(matrix[2][3]);  // 6
```

### Array Utilities

```javascript
// Check if array
isArray([1, 2, 3]);  // true
isArray('hello');    // false

// Flatten nested arrays
let nested = [1, [2, 3], [4, [5, 6]]];
let flat = flatten(nested);  // [1, 2, 3, 4, 5, 6]

// Flatten with depth limit
let partial = flatten(nested, 1);  // [1, 2, 3, 4, [5, 6]]

// Clone array (deep copy)
let original = [1, 2, [3, 4]];
let copy = clone(original);
```

---

## Array Methods

### Transforming Methods

#### map

Transform each element:

```javascript
let numbers = [1, 2, 3, 4, 5];
let doubled = numbers.map(fun(x) {
    return x * 2;
});
print(doubled);  // [2, 4, 6, 8, 10]

// With index
let indexed = numbers.map(fun(val, idx, arr) {
    return val + idx;
});
```

#### filter

Select elements that match a condition:

```javascript
let numbers = [1, 2, 3, 4, 5, 6];
let evens = numbers.filter(fun(x) {
    return x % 2 == 0;
});
print(evens);  // [2, 4, 6]
```

#### reduce

Reduce array to single value:

```javascript
let numbers = [1, 2, 3, 4, 5];
let sum = numbers.reduce(fun(acc, val) {
    return acc + val;
}, 1);
print(sum);  // 16

// Without initial value (uses first element)
let product = numbers.reduce(fun(acc, val) {
    return acc * val;
});
print(product);  // 120
```

### Searching Methods

#### find

Find first matching element:

```javascript
let users = [
    { name: 'John', age: 25 },
    { name: 'Jane', age: 30 },
    { name: 'Bob', age: 35 }
];

let user = users.find(fun(u) {
    return u.age > 28;
});
print(user.name);  // 'Jane'
```

#### findIndex

Find index of first matching element:

```javascript
let numbers = [10, 20, 30, 40, 50];
let idx = numbers.findIndex(fun(x) {
    return x > 25;
});
print(idx);  // 3 (30 is at index 3, 1-based)
```

#### findLast / findLastIndex

Find from end of array:

```javascript
let numbers = [10, 20, 30, 20, 10];
let last = numbers.findLast(fun(x) {
    return x == 20;
});
print(last);  // 20 (last occurrence)
```

#### indexOf / lastIndexOf

```javascript
let arr = [1, 2, 3, 2, 1];
print(arr.indexOf(2));      // 2 (first occurrence)
print(arr.lastIndexOf(2));  // 4 (last occurrence)
print(arr.indexOf(5));      // -1 (not found)
```

#### includes

```javascript
let fruits = ['apple', 'banana', 'orange'];
print(fruits.includes('banana'));  // true
print(fruits.includes('grape'));   // false
```

### Mutating Methods

#### push

Add elements to end:

```javascript
let arr = [1, 2, 3];
arr.push(4);
arr.push(5, 6);
print(arr);  // [1, 2, 3, 4, 5, 6]
```

#### pop

Remove last element:

```javascript
let arr = [1, 2, 3, 4];
let last = arr.pop();
print(last);  // 4
print(arr);   // [1, 2, 3]
```

#### shift

Remove first element:

```javascript
let arr = [1, 2, 3, 4];
let first = arr.shift();
print(first);  // 1
print(arr);    // [2, 3, 4]
```

#### unshift

Add elements to beginning:

```javascript
let arr = [3, 4];
arr.unshift(1, 2);
print(arr);  // [1, 2, 3, 4]
```

#### splice

Remove/insert elements:

```javascript
let arr = [1, 2, 3, 4, 5];

// Remove 2 elements starting at index 2
let removed = arr.splice(2, 2);
print(removed);  // [2, 3]
print(arr);      // [1, 4, 5]

// Insert elements
arr = [1, 2, 5];
arr.splice(3, 0, 3, 4);  // At index 3, remove 0, insert 3, 4
print(arr);  // [1, 2, 3, 4, 5]

// Replace elements
arr = [1, 2, 3, 4, 5];
arr.splice(2, 2, 99);  // Remove 2 elements, insert 99
print(arr);  // [1, 99, 5]
```

#### reverse

Reverse array in place:

```javascript
let arr = [1, 2, 3, 4, 5];
arr.reverse();
print(arr);  // [5, 4, 3, 2, 1]
```

#### sort

Sort array in place:

```javascript
let numbers = [3, 1, 4, 1, 5, 9];
numbers.sort();
print(numbers);  // [1, 1, 3, 4, 5, 9]

// Custom comparator
let words = ['banana', 'apple', 'cherry'];
words.sort(fun(a, b) {
    if (a < b) { return -1;};
    if (a > b) { return 1;};
    return 0;
});
print(words);  // ['apple', 'banana', 'cherry']
```

### Non-Mutating Methods

#### slice

Extract portion of array:

```javascript
let arr = [1, 2, 3, 4, 5];
let sub = arr.slice(2, 4);  // From index 2 to 4 (exclusive)
print(sub);  // [2, 3]

// Negative indices (from end)
let last2 = arr.slice(-2);
print(last2);  // [4, 5]
```

#### concat

Combine arrays:

```javascript
let arr1 = [1, 2];
let arr2 = [3, 4];
let combined = arr1.concat(arr2, [5, 6]);
print(combined);  // [1, 2, 3, 4, 5, 6]
```

#### toSorted / toReversed / toSpliced

Non-mutating versions (return new array):

```javascript
let arr = [3, 1, 4, 1, 5];
let sorted = arr.toSorted();
print(sorted);  // [1, 1, 3, 4, 5]
print(arr);     // [3, 1, 4, 1, 5] (unchanged)

let reversed = arr.toReversed();
print(reversed);  // [5, 1, 4, 1, 3]
print(arr);       // [3, 1, 4, 1, 5] (unchanged)
```

#### with

Create copy with one element changed:

```javascript
let arr = [1, 2, 3, 4];
let modified = arr.with(2, 99);  // Change index 2 to 99
print(modified);  // [1, 99, 3, 4]
print(arr);       // [1, 2, 3, 4] (unchanged)
```

### Iteration Methods

#### forEach

Execute function for each element:

```javascript
let numbers = [1, 2, 3, 4, 5];
numbers.forEach(fun(val, idx) {
    print('Index ' + idx + ': ' + val);
});
```

#### every

Test if all elements match condition:

```javascript
let numbers = [2, 4, 6, 8];
let allEven = numbers.every(fun(x) {
    return x % 2 == 0;
});
print(allEven);  // true
```

#### some

Test if any element matches condition:

```javascript
let numbers = [1, 3, 5, 8];
let hasEven = numbers.some(fun(x) {
    return x % 2 == 0;
});
print(hasEven);  // true
```

### Utility Methods

#### unique

Remove duplicates:

```javascript
let arr = [1, 2, 2, 3, 3, 3, 4];
let unique = arr.unique();
print(unique);  // [1, 2, 3, 4]
```

#### at

Access with negative indices:

```javascript
let arr = [1, 2, 3, 4, 5];
print(arr.at(1));   // 1 (first element)
print(arr.at(-1));  // 5 (last element)
print(arr.at(-2));  // 4 (second to last)
```

#### entries

Get [index, value] pairs:

```javascript
let arr = ['a', 'b', 'c'];
let pairs = arr.entries();
// [[1, 'a'], [2, 'b'], [3, 'c']]
```

#### join / toString

Convert to string:

```javascript
let arr = [1, 2, 3];
print(arr.join());       // '1, 2, 3'
print(arr.join('-'));    // '1-2-3'
print(arr.toString());   // '1, 2, 3'
```

#### from

Create array from iterable:

```javascript
let str = 'hello';
let chars = [].from(str);
print(chars);  // ['h', 'e', 'l', 'l', 'o']

// With mapping function
let doubled = [].from([1, 2, 3], fun(x) {
    return x * 2;
});
print(doubled);  // [2, 4, 6]
```

#### of

Create array from arguments:

```javascript
let arr = [].of(1, 2, 3, 4);
print(arr);  // [1, 2, 3, 4]
```

#### copyWithin

Copy elements within array:

```javascript
let arr = [1, 2, 3, 4, 5];
arr.copyWithin(1, 3);  // Copy from index 3 to index 1
print(arr);  // [1, 3, 4, 4, 5]
```

#### delete

Remove element at index:

```javascript
let arr = [1, 2, 3, 4, 5];
arr.delete(3);  // Remove element at index 3
print(arr);  // [1, 2, 4, 5]
```

---

## Objects

### Creating Objects

```javascript
let person = {
    name: 'John Doe',
    age: 30,
    email: 'john@example.com'
};
```

### Accessing Properties

```javascript
// Dot notation
print(person.name);  // 'John Doe'

// Bracket notation
print(person['age']);  // 30

// Dynamic property access
let prop = 'email';
print(person[prop]);  // 'john@example.com'
```

### Setting Properties

```javascript
person.age = 31;
person['city'] = 'Boston';
person.country = 'USA';  // Add new property
```

### Nested Objects

```javascript
let company = {
    name: 'Tech Corp',
    address: {
        street: '123 Main St',
        city: 'Boston',
        zip: '02101'
    },
    employees: [
        { name: 'Alice', role: 'Developer' },
        { name: 'Bob', role: 'Designer' }
    ]
};

print(company.address.city);        // 'Boston'
print(company.employees[1].name);   // 'Alice' (1-based)
```

### Methods in Objects

```javascript
let calculator = {
    add: fun(a, b) {
        return a + b;
    },
    multiply: fun(a, b) {
        return a * b;
    }
};

print(calculator.add(5, 3));       // 8
print(calculator.multiply(4, 7));  // 28
```

### Dynamic Property Names

```javascript
let key = 'status';
let obj = {};
obj[key] = 'active';
print(obj.status);  // 'active'
```

## Object Member Methods

### Property Enumeration

#### keys

Get array of all property names:

```javascript
let person = { name: 'John', age: 30, city: 'Boston' };
let props = person.keys();
print(props);  // ['name', 'age', 'city']

// Iterate over keys
person.keys().forEach(fun(key) {
    print(key + ': ' + person[key]);
});
```

#### values

Get array of all property values:

```javascript
let scores = { math: 85, english: 92, science: 78 };
let vals = scores.values();
print(vals);  // [85, 92, 78]

// Calculate total
let total = scores.values().reduce(fun(sum, val) {
    return sum + val;
}, 0);
print(total);  // 255
```

#### entries

Get array of [key, value] pairs:

```javascript
let config = { host: 'localhost', port: 8080, ssl: true };
let pairs = config.entries();
print(pairs);  // [['host', 'localhost'], ['port', 8080], ['ssl', true]]

// Convert to different format
config.entries().forEach(fun(pair) {
    print(pair[1] + '=' + pair[2]);
});
// Output:
// host=localhost
// port=8080
// ssl=true
```

### Property Management

#### size / length

Get number of properties:

```javascript
let obj = { a: 1, b: 2, c: 3 };
print(obj.size());    // 3
print(obj.length());  // 3 (alias)

let empty = {};
print(empty.size());  // 0
```

#### hasKey / has

Check if property exists:

```javascript
let user = { name: 'Alice', email: 'alice@example.com' };

print(user.hasKey('name'));     // true
print(user.has('email'));       // true (alias)
print(user.hasKey('phone'));    // false

// Safe property access
if (user.has('address')) {
    print(user.address.city);
} else {
    print('No address available');
};
```

#### isEmpty

Check if object has no properties:

```javascript
let obj1 = {};
print(obj1.isEmpty());  // true

let obj2 = { x: 1 };
print(obj2.isEmpty());  // false

// Clear and check
obj2.clear();
print(obj2.isEmpty());  // true
```

#### get

Get property value with optional default:

```javascript
let config = { timeout: 30, retry: 3 };

// Get existing property
print(config.get('timeout'));  // 30

// Get with default for missing property
print(config.get('maxSize', 100));  // 100 (returns default)

// Without default returns empty
print(config.get('missing'));  // '' (empty)

// Use in conditionals
let port = config.get('port', 8080);
print('Port: ' + port);  // Port: 8080
```

#### set

Set property value:

```javascript
let user = { name: 'John' };

// Add new property
user.set('age', 30);
print(user.age);  // 30

// Update existing property
user.set('name', 'Jane');
print(user.name);  // Jane

// Chain multiple sets
user.set('city', 'Boston').set('country', 'USA');

// Set with dynamic key
let key = 'status';
user.set(key, 'active');
print(user.status);  // active
```

#### delete / remove

Remove property:

```javascript
let person = { name: 'Alice', age: 25, temp: 'delete-me' };

// Delete property
person.delete('temp');
print(person.hasKey('temp'));  // false

// Using alias
person.remove('age');
print(person);  // { name: 'Alice' }

// Delete non-existent property (safe)
person.delete('notThere');  // No error
```

#### clear

Remove all properties:

```javascript
let data = { a: 1, b: 2, c: 3, d: 4 };
print(data.size());  // 4

data.clear();
print(data.size());  // 0
print(data.isEmpty());  // true

// Object is now empty but still usable
data.set('x', 10);
print(data.x);  // 10
```

### Object Transformation

#### clone

Create deep copy of object:

```javascript
let original = {
    name: 'John',
    scores: [85, 90, 78],
    address: {
        city: 'Boston',
        zip: '02101'
    }
};

let copy = original.clone();

// Modify copy
copy.name = 'Jane';
copy.scores[1] = 95;
copy.address.city = 'New York';

// Original unchanged
print(original.name);          // John
print(original.scores[1]);     // 85 (1-based indexing)
print(original.address.city);  // Boston

// Copy modified
print(copy.name);          // Jane
print(copy.scores[1]);     // 95
print(copy.address.city);  // New York
```

#### merge

Merge another object's properties:

```javascript
let defaults = {
    timeout: 30,
    retry: 3,
    verbose: false
};

let userConfig = {
    timeout: 60,
    cache: true
};

// Merge userConfig into defaults (mutates defaults)
defaults.merge(userConfig);
print(defaults);
// {
//   timeout: 60,    // Overwritten
//   retry: 3,       // Preserved
//   verbose: false, // Preserved
//   cache: true     // Added
// }

// Nested merge (overwrites nested objects)
let obj1 = { a: 1, nested: { x: 10, y: 20 } };
let obj2 = { b: 2, nested: { y: 30, z: 40 } };

obj1.merge(obj2);
print(obj1);
// {
//   a: 1,
//   b: 2,
//   nested: { y: 30, z: 40 }  // obj2.nested replaces obj1.nested
// }

// Safe merge pattern (preserve original)
let merged = original.clone().merge(updates);
```

### Iteration Methods

#### forEach

Execute function for each property:

```javascript
let scores = { math: 85, english: 92, science: 78 };

// With value only
scores.forEach(fun(val) {
    print(val);
});
// Output: 85, 92, 78

// With value and key
scores.forEach(fun(val, key) {
    print(key + ': ' + val);
});
// Output:
// math: 85
// english: 92
// science: 78

// Modify during iteration (affects original)
let data = { a: 1, b: 2, c: 3 };
data.forEach(fun(val, key) {
    data[key] = val * 2;
});
print(data);  // { a: 2, b: 4, c: 6 }

// Count values matching condition
let count = 0;
scores.forEach(fun(val) {
    if (val > 80) {
        count = count + 1;
    }
});
print(count);  // 2
```

#### map

Transform all values, return new object:

```javascript
let prices = { apple: 1.50, banana: 0.75, orange: 2.00 };

// Double all prices
let doubled = prices.map(fun(val) {
    return val * 2;
});
print(doubled);  // { apple: 3, banana: 1.5, orange: 4 }

// With key parameter
let formatted = prices.map(fun(val, key) {
    return key + ': $' + val;
});
print(formatted);
// { apple: 'apple: $1.5', banana: 'banana: $0.75', orange: 'orange: $2' }

// Transform nested objects
let users = {
    user1: { name: 'John', age: 30 },
    user2: { name: 'Jane', age: 25 }
};

let names = users.map(fun(user) {
    return user.name;
});
print(names);  // { user1: 'John', user2: 'Jane' }

// Original unchanged (non-mutating)
print(prices.apple);  // 1.5
```

#### filter

Filter properties by condition, return new object:

```javascript
let scores = { math: 85, english: 92, science: 78, history: 95 };

// Filter passing grades (>= 90)
let passing = scores.filter(fun(val) {
    return val >= 90;
});
print(passing);  // { english: 92, history: 95 }

// Filter by key
let user = {
    name: 'John',
    _id: 12345,
    email: 'john@example.com',
    _internal: true
};

let publicFields = user.filter(fun(val, key) {
    return !key.startsWith('_');
});
print(publicFields);  // { name: 'John', email: 'john@example.com' }

// Complex filtering
let products = {
    item1: { name: 'Widget', price: 10, inStock: true },
    item2: { name: 'Gadget', price: 25, inStock: false },
    item3: { name: 'Tool', price: 15, inStock: true }
};

let available = products.filter(fun(item) {
    return item.inStock && item.price < 20;
});
print(available);  // { item1: {...}, item3: {...} }
```

#### some

Test if any property matches condition:

```javascript
let scores = { math: 85, english: 92, science: 78 };

// Any score above 90?
let hasExcellent = scores.some(fun(val) {
    return val > 90;
});
print(hasExcellent);  // true

// Any failing grade?
let hasFailing = scores.some(fun(val) {
    return val < 60;
});
print(hasFailing);  // false

// With key
let config = {
    debugMode: false,
    logging: true,
    verbose: false
};

let anyEnabled = config.some(fun(val, key) {
    return val == true && key.startsWith('log');
});
print(anyEnabled);  // true

// Validation example
let user = { name: '', email: 'valid@example.com' };
let hasEmptyField = user.some(fun(val) {
    return val == '';
});
print(hasEmptyField);  // true
```

#### every

Test if all properties match condition:

```javascript
let scores = { math: 85, english: 92, science: 88 };

// All passing (>= 60)?
let allPassing = scores.every(fun(val) {
    return val >= 60;
});
print(allPassing);  // true

// All excellent (>= 90)?
let allExcellent = scores.every(fun(val) {
    return val >= 90;
});
print(allExcellent);  // false

// Validation example
let requiredFields = { name: 'John', email: 'john@example.com', age: 30 };
let allPresent = requiredFields.every(fun(val) {
    return val != null && val != '';
});
print(allPresent);  // true

// Type checking
let numbers = { a: 1, b: 2, c: 3 };
let allNumeric = numbers.every(fun(val) {
    return typeof val == 'number';
});
print(allNumeric);  // true
```

### Method Chaining

Combine object methods for complex transformations:

```javascript
let inventory = {
    apple: { price: 1.50, quantity: 10 },
    banana: { price: 0.75, quantity: 5 },
    orange: { price: 2.00, quantity: 0 },
    grape: { price: 3.50, quantity: 8 }
};

// Filter in-stock items, then get total value
let totalValue = inventory
    .filter(fun(item) {
        return item.quantity > 0;
    })
    .map(fun(item) {
        return item.price * item.quantity;
    })
    .values()
    .reduce(fun(sum, val) {
        return sum + val;
    }, 0);

print(totalValue);  // 43.75

// Transform and validate
let users = {
    user1: { name: 'John', age: 30, active: true },
    user2: { name: 'Jane', age: 25, active: false },
    user3: { name: 'Bob', age: 35, active: true }
};

let activeUserNames = users
    .filter(fun(u) { return u.active; })
    .map(fun(u) { return u.name; })
    .values();

print(activeUserNames);  // ['John', 'Bob']
```

---

## Classes

### Basic Class Definition

```javascript
class Person {
    constructor(name, age) {
        this.name = name;
        this.age = age;
    }
    
    greet() {
        print('Hello, I am ' + this.name);
    }
    
    getAge() {
        return this.age;
    }
}

let person = new Person('Alice', 25);
person.greet();              // Hello, I am Alice
print(person.getAge());      // 25
print(person.name);          // Alice
```

### Class Fields

Declare instance fields with the `field` keyword:

```javascript
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

let rect = new Rectangle(10, 5);
print(rect.getArea());  // 50
```

### Field Syntax

```javascript
// Single field
field x;
field x = 10;

// Multiple fields (comma-separated)
field x, y, z;
field x = 1, y = 2, z = 3;

// Mixed with/without initializers
field x = 1, y, z = 3;

// Multiple declarations
field width = 0, height = 0;
field color = 'black';
field name;
```

### Field Initialization Order

1. Parent class fields (if inheritance)
2. Current class fields
3. Constructor body

```javascript
class Point {
    field x = 0, y = 0;
    
    constructor(x, y) {
        // Fields already initialized to 0
        this.x = x;  // Now set to parameter value
        this.y = y;
    }
}
```

### Static Methods

Methods that belong to the class, not instances:

```javascript
class MathHelper {
    static add(a, b) {
        return a + b;
    }
    
    static multiply(a, b) {
        return a * b;
    }
    
    static PI = 3.14159;  // Note: static fields via method
}

// Call without creating instance
print(MathHelper.add(5, 3));        // 8
print(MathHelper.multiply(4, 7));   // 28
```

### Inheritance

Classes can extend other classes:

```javascript
class Animal {
    field name, species;
    
    constructor(name, species) {
        this.name = name;
        this.species = species;
    }
    
    speak() {
        print(this.name + ' makes a sound');
    }
}

class Dog extends Animal {
    field breed;
    
    constructor(name, breed) {
        super(name, 'Dog');  // Call parent constructor
        this.breed = breed;
    }
    
    speak() {
        print(this.name + ' barks!');
    }
    
    getBreed() {
        return this.breed;
    }
}

let dog = new Dog('Rex', 'Labrador');
dog.speak();              // Rex barks!
print(dog.getBreed());    // Labrador
print(dog.species);       // Dog
```

### Super Keyword

Call parent constructor:

```javascript
class Shape {
    field x, y;
    
    constructor(x, y) {
        this.x = x;
        this.y = y;
    }
}

class Circle extends Shape {
    field radius;
    
    constructor(x, y, r) {
        super(x, y);  // Must call parent constructor
        this.radius = r;
    }
    
    getArea() {
        return 3.14159 * this.radius * this.radius;
    }
}
```

### This Binding

The `this` keyword refers to the current instance:

```javascript
class Counter {
    field count = 0;
    
    increment() {
        this.count = this.count + 1;
    }
    
    getValue() {
        return this.count;
    }
}

let c = new Counter();
c.increment();
c.increment();
print(c.getValue());  // 2
```

### Complete Class Example

```javascript
class BankAccount {
    field balance = 0;
    field accountNumber;
    field owner;
    
    constructor(owner, accountNum) {
        this.owner = owner;
        this.accountNumber = accountNum;
    }
    
    deposit(amount) {
        if (amount > 0) {
            this.balance = this.balance + amount;
            return true;
        }
        return false;
    }
    
    withdraw(amount) {
        if (amount > 0 && amount <= this.balance) {
            this.balance = this.balance - amount;
            return true;
        }
        return false;
    }
    
    getBalance() {
        return this.balance;
    }
    
    getInfo() {
        return this.owner + ' (' + this.accountNumber + '): $' + this.balance;
    }
    
    static create(owner) {
        let accNum = 'ACC' + Math.floor(Math.random() * 10000);
        return new BankAccount(owner, accNum);
    }
}

let account = new BankAccount('John Doe', 'ACC12345');
account.deposit(1000);
account.withdraw(250);
print(account.getInfo());  // John Doe (ACC12345): $750
```

---

## Built-in Functions

### Output

#### print

Output to debug/console:

```javascript
print('Hello, World!');
print(42);
print([1, 2, 3]);
print({ name: 'John', age: 30 });
```

### Type Checking

#### typeof

Return type as string:

```javascript
typeof 42;          // 'number'
typeof 'hello';     // 'string'
typeof true;        // 'boolean'
typeof null;        // 'null'
typeof undefined;   // 'undefined'
typeof [1, 2, 3];   // 'array'
typeof {};          // 'object'
typeof fun() {};    // 'function'
```

#### isArray

Check if value is array:

```javascript
isArray([1, 2, 3]);  // true
isArray('hello');    // false
isArray(null);       // false
```

#### isNumeric

Check if value is numeric:

```javascript
isNumeric(42);       // true
isNumeric('42');     // true
isNumeric('hello');  // false
isNumeric(true);     // false
```

### Array Utilities

#### range

Create array of numbers:

```javascript
range(5);          // [0, 1, 2, 3, 4]
range(1, 6);       // [1, 2, 3, 4, 5]
range(0, 10, 2);   // [0, 2, 4, 6, 8]
range(10, 0, -2);  // [10, 8, 6, 4, 2]
```

#### flatten

Flatten nested arrays:

```javascript
let nested = [1, [2, 3], [4, [5, 6]]];
flatten(nested);      // [1, 2, 3, 4, 5, 6]
flatten(nested, 1);   // [1, 2, 3, 4, [5, 6]]
```

#### clone

Deep copy arrays/objects:

```javascript
let original = [1, 2, [3, 4]];
let copy = clone(original);
copy[3][1] = 99;
print(original[3][1]);  // 4 (unchanged)
```

### Iteration Helper

#### foreach

Iterate over array:

```javascript
let nums = [1, 2, 3, 4, 5];
foreach(nums, fun(val, idx, arr) {
    print('Index ' + idx + ': ' + val)
});
```

Iterate over object:

```javascript
let scores = {math: 85, english: 92, science: 78};
foreach(scores, fun(val, key) {
    s = s + val; c += 1
}); return('Average: ' + s/c); // Average: 85
```

### Regular Expression Helper

#### regex

Create regex engine:

```javascript
let re = regex(`/hello/i`);  // Case-insensitive
let re2 = regex(`\\d+`);      // Matches digits
```

---

## String Methods

### Accessing Characters

#### charAt / charCodeAt

```javascript
let str = 'Hello';
print(str.charAt(0));      // 'H' (0-based like JavaScript)
print(str.charCodeAt(0));  // 72 (ASCII code of 'H')
```

#### at

Access with negative indices:

```javascript
let str = 'Hello';
print(str.at(0));   // 'H' (first char)
print(str.at(-1));  // 'o' (last char)
print(str.at(-2));  // 'l' (second to last)
```

### String Properties

```javascript
let str = 'Hello';
print(str.length);  // 5
```

### String Transformation

#### toLowerCase / toUpperCase

```javascript
let str = 'Hello World';
print(str.toLowerCase());  // 'hello world'
print(str.toUpperCase());  // 'HELLO WORLD'
```

#### trim / trimStart / trimEnd

```javascript
let str = '  hello  ';
print(str.trim());       // 'hello'
print(str.trimStart());  // 'hello  '
print(str.trimEnd());    // '  hello'
```

#### repeat

```javascript
let str = 'abc';
print(str.repeat(3));  // 'abcabcabc'
```

#### padStart / padEnd

```javascript
let str = '5';
print(str.padStart(3, '0'));  // '005'
print(str.padEnd(3, '0'));    // '500'
```

### String Searching

#### indexOf / lastIndexOf

```javascript
let str = 'hello world hello';
print(str.indexOf('hello'));      // 0 (first occurrence)
print(str.lastIndexOf('hello'));  // 12 (last occurrence)
print(str.indexOf('xyz'));        // -1 (not found)
```

#### includes

```javascript
let str = 'hello world';
print(str.includes('world'));  // true
print(str.includes('xyz'));    // false
```

#### startsWith / endsWith

```javascript
let str = 'hello world';
print(str.startsWith('hello'));  // true
print(str.endsWith('world'));    // true
print(str.startsWith('world'));  // false
```

### String Extraction

#### slice

```javascript
let str = 'hello world';
print(str.slice(0, 5));   // 'hello'
print(str.slice(6));      // 'world'
print(str.slice(-5));     // 'world' (last 5 chars)
```

#### substring

```javascript
let str = 'hello world';
print(str.substring(0, 5));   // 'hello'
print(str.substring(6, 11));  // 'world'
```

### String Manipulation

#### concat

```javascript
let str1 = 'Hello';
let str2 = 'World';
print(str1.concat(' ', str2));  // 'Hello World'
```

#### split

```javascript
let str = 'apple,banana,orange';
let fruits = str.split(',');
print(fruits);  // ['apple', 'banana', 'orange']

let chars = 'hello'.split('');
print(chars);  // ['h', 'e', 'l', 'l', 'o']

// With limit
let limited = 'a,b,c,d'.split(',', 2);
print(limited);  // ['a', 'b']
```

#### replace / replaceAll

```javascript
let str = 'hello world hello';

// Replace first occurrence
print(str.replace('hello', 'hi'));     // 'hi world hello'

// Replace all occurrences
print(str.replaceAll('hello', 'hi'));  // 'hi world hi'

// With function
let result = str.replace('hello', fun(match) {
    return match.toUpperCase();
});
print(result);  // 'HELLO world hello'

// With regex
let str2 = 'hello123world456';
let result2 = str2.replace(`/\\d+/`, 'X');  // Replace digits
print(result2);  // 'helloXworld456'
```

### Pattern Matching

#### match / matchAll

```javascript
let str = 'The price is $10 and $20';

// Match first occurrence
let match = str.match('$10');
print(match);  // ['$10']

// Match with regex (all digits)
let numbers = str.match(`/\\d+/g`);
print(numbers);  // ['10', '20']

// matchAll returns array of all matches
let all = str.matchAll(`/\\d+/g`);
// [['10'], ['20']]
```

### Conversion

#### toString / valueOf

```javascript
let str = 'hello';
print(str.toString());  // 'hello'
print(str.valueOf());   // 'hello'
```

#### fromCharCode

```javascript
print(''.fromCharCode(72));  // 'H'
```

### Comparison

#### localeCompare

```javascript
let a = 'apple';
let b = 'banana';
print(a.localeCompare(b));  // -1 (a < b)
print(b.localeCompare(a));  // 1  (b > a)
print(a.localeCompare(a));  // 0  (equal)
```

---

## Template Literals

### Basic Syntax

Use backticks (`` ` ``) for template literals:

```javascript
let name = 'Alice';
let greeting = `Hello, ${name}!`;
print(greeting);  // Hello, Alice!
```

### Expressions in Templates

Any expression can be embedded:

```javascript
let a = 5, b = 10;
print(`${a} + ${b} = ${a + b}`);  // 5 + 10 = 15

let items = [1, 2, 3];
print(`Array length: ${items.length}`);  // Array length: 3
```

### Multi-line (Simulated)

```javascript
let message = `Line 1
Line 2
Line 3`;
// Note: Actual newline handling depends on VBA
```

### Nested Templates

```javascript
let user = { name: 'John', role: 'admin' };
let status = `User ${user.name} (${user.role == 'admin' ? 'Administrator' : 'User'})`;
print(status);  // User John (Administrator)
```

### Complex Expressions

```javascript
let users = [
    { name: 'Alice', score: 85 },
    { name: 'Bob', score: 92 }
];

for (let i = 1, i <= users.length, i += 1) {
    let user = users[i];
    print(`${user.name}: ${user.score >= 90 ? 'A' : 'B'}`);
}
// Output:
// Alice: B
// Bob: A
```

---

## Regular Expressions

### Creating Regex

#### Slash Notation (Recommended)

```javascript
let re = regex(`/pattern/flags`);
let digitPattern = regex(`/\\d+/`);
let caseInsensitive = regex(`/hello/i`);
```

#### Constructor Notation

```javascript
let re = regex(`pattern`);
re.init(`\\d+`);                    // Just pattern
re.init(`hello`, true);             // Pattern + case-insensitive
re.init(`hello`, true, false, true); // Pattern + flags
```

### Regex Flags

- `i` - Case-insensitive
- `g` - Global (find all matches)
- `m` - Multiline
- `s` - Dot matches newline (dotAll)

### Regex Methods

#### test

Test if pattern matches:

```javascript
let re = regex(`\\d+`);
print(re.test('hello123'));  // true
print(re.test('hello'));     // false
```

#### exec

Execute regex and get first match:

```javascript
let re = regex(`(\\d+)-(\\d+)`);
let result = re.exec('Phone: 555-1234');
print(result);  // ['555-1234', '555', '1234']
// result[1] = full match, result[2+] = capture groups
```

#### execAll

Get all matches (global flag required):

```javascript
let re = regex(`\\d+`);
let matches = re.execAll('Numbers: 10, 20, 30');
print(matches);  // [['10'], ['20'], ['30']]
```

#### replace

Replace pattern matches:

```javascript
let re = regex(`\\d+`);
let str = 'Price: $10 and $20';
let result = re.replace(str, 'XX');
print(result);  // Price: $XX and $XX
```

#### split

Split string by pattern:

```javascript
let re = regex(`[,;]`);
let str = 'apple,banana;orange';
let parts = re.split(str);
print(parts);  // ['apple', 'banana', 'orange']
```

#### escape

Escape special regex characters, and the first character of the given string:

```javascript
let escaped = regex().escape('a.b*c?');
print(escaped);  // '\a\.b\*c\?'
```

### Regex Properties

```javascript
let re = regex(`hello`);

// Get/Set pattern
print(re.getPattern());  // 'hello'
re.setPattern(`world`);
print(re.getPattern());  // 'world'

// Get/Set flags
print(re.getIgnoreCase());  // true
re.setIgnoreCase(false);

print(re.getMultiline());
re.setMultiline(true);

print(re.getDotAll());
re.setDotAll(true);
```

### String Methods with Regex

#### match

```javascript
let str = 'The numbers are 10 and 20';
let nums = str.match('/\\d+/g');
print(nums);  // ['10', '20']
```

#### replace

```javascript
let str = 'hello123world456';

// Replace with string
let r1 = str.replace('/\\d+/', 'X');
print(r1);  // 'helloXworld456' (first only)

// Replace all with global flag
let r2 = str.replace('/\\d+/g', 'X');
print(r2);  // 'helloXworldX'

// Replace with function
let r3 = str.replace('/\\d+/g', fun(match) {
    return '[' + match + ']'
});
print(r3);  // 'hello[123]world[456]'
```

### Common Patterns

```javascript
// Email validation (simple)
let email = regex(`^[^@]+@[^@]+\\.[^@]+$`);
print(email.test('user@example.com'));  // true

// Phone number
let phone = regex(`^\\d{3}-\\d{3}-\\d{4}$`);
print(phone.test('555-123-4567'));  // true

// URL
let url = regex(`^https?:\\/\\/`);
print(url.test('https://example.com'));  // true

// Hexadecimal color
let color = regex(`^#[0-9a-fA-F]{6}$`);
print(color.test('#FF5733'));  // true

// Extract all words
let words = regex(`\\w+`);
let text = 'Hello, World! 123';
print(words.execAll(text));  // [['Hello'], ['World'], ['123']]
```

---

## Error Handling

### Try-Catch

```javascript
try {
    let x = 10 / 0;  // May cause error in some contexts
    print('Result: ' + x);
} catch {
    print('An error occurred!');
}
```

### Nested Try-Catch

```javascript
try {
    try {
        // Inner operation
        let result = riskyOperation();
    } catch {
        print('Inner error');
    };
    
    // Outer operation
    anotherOperation();
} catch {
    print('Outer error');
};
```

### Error Recovery

```javascript
fun safeDivide(a, b) {
    try {
        if (b == 0) {
            return null;
        }
        return a / b;
    } catch {
        return null;
    };
}

let result = safeDivide(10, 2);
if (result == null) {
    print('Division failed');
} else {
    print('Result: ' + result);
};
```

### Validation Pattern

```javascript
fun validateUser(user) {
    try {
        if (typeof(user.name) != 'string') {
            return false;
        }
        if (typeof(user.age) != 'number') {
            return false;
        }
        if (user.age < 0) {
            return false;
        }
        return true;
    } catch {
        return false;
    };
};

let user = { name: 'John', age: 30 };
if (validateUser(user)) {
    print('User is valid');
} else {
    print('Invalid user data');
};
```

---

## VBA Integration

### Evaluating with VBA-Expressions library

Use `@(...)` syntax to evaluate with [VBA-Expressions](https://github.com/ECP-Solutions/VBA-Expressions):

```javascript
// Call VBA-Expressions functions
let a = @({1;0;4});
let b = @({1;1;6});
let c = @({-3;0;-10});
print(@(MROUND(LUDECOMP(ARRAY(a;b;c));4)))
```

### Call to native VBA Functions
Functions with a single `Variant` argument can be invoked from ASF through the VBA-Expressions library

```vb
Dim asfGlobals As New ASF_Globals
Dim progIdx  As Long
With asfGlobals
	.ASF_InitGlobals
	.gExprEvaluator.DeclareUDF "ThisWBname", "UserDefFunctions"
End With
Set scriptEngine = New ASF
With scriptEngine
	.SetGlobals asfGlobals
	progIdx = .Compile("/*Get Thisworkbook name*/ return(@(ThisWBname()))")
	.Run progIdx
	Debug.Print CStr(.OUTPUT_)
End With
```

>**Notes**: In the above example, the `ThisWBname` function is coded in the `UDFunctions.cls` class module. And the callback is defined in the `VBAcallBack.cls` class module as `Public UserDefFunctions As New UDFunctions`. Each UDF will receive a single evaluated array of strings, containing all the arguments given. In this way, the ASF remains sandboxed and secure at runtime. Users can add new class module with custom functions and declare them with the `DeclareUDF` method as `<exprEvaluator>.DeclareUDF <functionName>, <alias>`; where `<alias>` is a custom name given to Variable holding an instance of the target class module.

### Injecting VBA Values

From VBA, inject values into ASF:

```vb
Dim engine As New ASF
engine.InjectVariable "userData", Array("John", 30, "john@example.com")

Dim code As String
code = "print(userData[1]);"  ' Prints: John

Dim idx As Long
idx = engine.Compile(code)
engine.Run idx
```

### Getting Results in VBA

```vb
Dim engine As New ASF
Dim code As String
code = "let x = 10; let y = 20; return(x + y);"

Dim idx As Long
idx = engine.Compile(code)
engine.Run idx

' Get output
Dim result As Variant
result = engine.OUTPUT_
Debug.Print result  ' Prints: 30
```

---

## Best Practices

### Code Organization

#### Use Functions for Reusability

```javascript
// Good
fun calculateTotal(items) {
    return items.reduce(fun(sum, item) {
        return sum + item.price;
    }, 0)
};

// Bad
let total = 0;
for (let i = 1, i <= items.length, i = i + 1) {
    total = total + items[i].price;
};
```

#### Modular Design

```javascript
// Separate concerns
fun validateInput(data) {
    // Validation logic
}

fun processData(data) {
    if (!validateInput(data)) {
        return null;
    }
    // Processing logic
}

fun formatOutput(result) {
    // Formatting logic
};
```

### Naming Conventions

```javascript
// Variables and functions: camelCase
let userName = 'John';
fun calculateTotal() { }

// Classes: PascalCase
class UserAccount { }
class BankTransaction { }

// Constants: UPPER_CASE (by convention)
let MAX_SIZE = 100;
let DEFAULT_TIMEOUT = 30;

// Boolean variables: is/has prefix
let isValid = true;
let hasPermission = false;
```

### Performance Tips

#### Use Array Methods Instead of Loops

```javascript
// Good - functional approach
let doubled = numbers.map(fun(x) { return x * 2; });
let evens = numbers.filter(fun(x) { return x % 2 == 0; });

// Slower - manual loops
let doubled = [];
for (let i = 1, i <= numbers.length, i = i + 1) {
    doubled.push(numbers[i] * 2);
};
```

#### Cache Length in Loops

```javascript
// Good
let len = arr.length;
for (let i = 1, i <= len, i = i + 1) {
    // Process arr[i]
}

// Less efficient
for (let i = 1, i <= arr.length, i = i + 1) {
    // arr.length evaluated each iteration
};
```

#### Use Local Variables

```javascript
// Good
fun processItems(items) {
    let total = 0;
    let count = items.length;
    // Process locally
    return { total: total, count: count };
}

// Less efficient (global access)
let globalTotal = 0;
fun processItems(items) {
    // Access global repeatedly
}
```

### Error Handling

#### Always Validate Input

```javascript
fun divide(a, b) {
    if (typeof a != 'number' || typeof b != 'number') {
        return null;
    };
    if (b == 0) {
        return null;
    };
    return a / b;
};
```

#### Use Try-Catch for Risky Operations

```javascript
fun parseUserData(jsonString) {
    try {
        // Risky operation
        return JSON.parse(jsonString);
    } catch {
        return null;
    };
};
```

### Memory Management

#### Clear Large Arrays

```javascript
let largeArray = range(1, 100000);
// Use array
processData(largeArray);
// Clear when done
largeArray = [];
```

#### Avoid Deep Nesting

```javascript
// Good - flat structure
fun processUser(user) {
    if (!user) return null;
    if (!user.name) return null;
    if (!user.email) return null;
    return formatUser(user);
}

// Bad - deep nesting
fun processUser(user) {
    if (user) {
        if (user.name) {
            if (user.email) {
                return formatUser(user);
            };
        };
    };
    return null;
};
```

---

## Appendix

### Option Base (EXPERIMENTAL)

Control array indexing (0-based or 1-based):

```javascript
// Set at program start
option base 0;  // Use 0-based indexing
option base 1;  // Use 1-based indexing (default)

let arr = [10, 20, 30];
print(arr[0]);  // Depends on option base setting
```

### Reserved Keywords

The following words are reserved and cannot be used as variable names:

- `class`
- `constructor`
- `extends`
- `super`
- `static`
- `field`
- `new`
- `fun`
- `return`
- `if`
- `else`
- `elseif`
- `for`
- `while`
- `switch`
- `case`
- `default`
- `break`
- `continue`
- `try`
- `catch`
- `typeof`
- `true`
- `false`
- `null`
- `let`
- `print`

### Limitations

- No `async/await` support
- No `Promise` or callback patterns
- No spread operator (`...`)
- No destructuring assignment
- No arrow functions (`=>`)
- No `const` declaration (use `let`)
- No `var` (use `let`)
- Limited regex features compared to JavaScript
- No module system (import/export)

### Future Enhancements

Potential future additions:

- Enhanced error handling with error objects
- More ES6+ features
- Performance optimizations

---

## Contributing

For bug reports, feature requests, or contributions, please contact the ASF development team.

## License

Copyright 2026 W. García

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

---

**End of Documentation**

Version 1.0 | Last Updated:January 2026
