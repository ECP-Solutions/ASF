---
layout: default
title: Home
has_children: true
nav_order: 1
description: "Modern JavaScript-like scripting for VBA projects."
---

# Introductory things
{: .fs-6 }

No COM dependencies. No migration required. Just drop in and start using `map`/`filter`/`reduce`, classes with inheritance, closures, regex, and more—all inside your existing Excel, Access, or Office VBA code.
{: .fs-4 .fw-300 }

```vb
' Before: 30+ lines of VBA boilerplate with ScriptControl
' After: Clean, readable ASF
Dim engine As New ASF
engine.Run engine.Compile("return [1,2,3,4,5].filter(fun(x){ return x > 2 }).reduce(fun(a,x){ return a+x }, 0)")
Debug.Print engine.OUTPUT_  ' => 12
```
