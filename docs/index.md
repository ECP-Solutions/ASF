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

## Ready-made add-ins

- **[Excel add-in and LibreOffice extension](Add-ins.html):** `ASF.xlam` runs ASF scripts in worksheet functions (`=ASF_RUN("return $1 * 2;", A1)`), from a script library and from the ASF tab of the ribbon; `ASF.oxt` brings the engine to LibreOffice. Installation in a few steps.
- **[Python](Python.html):** functions written in Python as worksheet functions, with numpy and pandas, running inside Excel for Windows. The add-in carries its own Python: one file to install.
