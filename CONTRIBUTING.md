# Contributing to ASF (Advanced Scripting Framework)

Thank you for your interest in contributing to ASF! This document provides guidelines and instructions for contributing to the project.

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [How Can I Contribute?](#how-can-i-contribute)
- [Getting Started](#getting-started)
- [Development Setup](#development-setup)
- [Coding Standards](#coding-standards)
- [Testing Guidelines](#testing-guidelines)
- [Submitting Changes](#submitting-changes)
- [Issue Guidelines](#issue-guidelines)
- [Community](#community)

---

## Code of Conduct

We are committed to providing a welcoming and inclusive environment for all contributors. By participating in this project, you agree to:

- Be respectful and considerate in your communication
- Welcome newcomers and help them get started
- Accept constructive criticism gracefully
- Focus on what's best for the community and the project
- Show empathy towards other community members

Unacceptable behavior includes harassment, trolling, insulting comments, or personal attacks. Project maintainers have the right to remove comments, commits, or contributions that don't align with this Code of Conduct.

---

## How Can I Contribute?

There are many ways to contribute to ASF:

### 🐛 Report Bugs

Found a bug? Help us fix it by:
- Checking if the issue already exists in [GitHub Issues](https://github.com/ECP-Solutions/ASF/issues)
- If not, open a new issue with a clear title and description
- Include a minimal reproducible example
- Specify your environment (Office version, 32/64-bit, Windows/Mac)

### 💡 Suggest Features

Have an idea for a new feature?
- Check existing issues to avoid duplicates
- Open a new issue tagged with `enhancement`
- Describe the use case and why it would be valuable
- Provide examples of how it would work

### 📝 Improve Documentation

Documentation improvements are always welcome:
- Fix typos, clarify explanations, or add examples
- Write tutorials or how-to guides
- Create video walkthroughs or demos
- Translate documentation to other languages

### 🔧 Submit Code

Ready to write code?
- Start with issues tagged `good first issue` or `help wanted`
- Discuss major changes in an issue before starting work
- Follow our coding standards (see below)
- Include tests for new features or bug fixes
- Update documentation to reflect your changes

### 🧪 Add Test Cases

Improve test coverage:
- Add tests for edge cases
- Write integration tests for real-world scenarios
- Improve test documentation

### 🎨 Create Examples

Show others how to use ASF:
- Add examples to the `examples/` directory
- Create real-world use case demonstrations
- Share your ASF projects (we may feature them!)

---

## Getting Started

### Prerequisites

- **Office Application** with VBA support (Excel, Access, Word, etc.)
  - Office 2007 or later (Windows or Mac)
  - Both 32-bit and 64-bit supported
- **[Rubberduck VBA](https://rubberduckvba.com/)** for running tests
- **Git** for version control
- **GitHub account** for submitting contributions

### Fork and Clone

1. Fork the repository on GitHub
2. Clone your fork locally:
   ```bash
   git clone https://github.com/YOUR-USERNAME/ASF.git
   cd ASF
   ```
3. Add the upstream repository:
   ```bash
   git remote add upstream https://github.com/ECP-Solutions/ASF.git
   ```

### Keep Your Fork Synced

Before starting work, sync with upstream:
```bash
git fetch upstream
git checkout main
git merge upstream/main
```

---

## Development Setup

### Import Modules into VBA Project

1. Open Excel (or your preferred Office app)
2. Press `Alt+F11` to open the VBA editor
3. Import the following class modules:
   - `ASF.cls`
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

### Set Up Rubberduck

1. Install [Rubberduck VBA](https://rubberduckvba.com/)
2. Import test modules from `tests/`
3. In Rubberduck, navigate to the Test Explorer
4. Run all tests to verify your setup

### Project Structure

```
ASF/
├── src/                    # Source code
│   ├── ASF.cls            # Main engine class
│   ├── ASF_Compiler.cls   # Compiler
│   ├── ASF_VM.cls         # Virtual machine
│   ├── ASF_Parser.cls     # Parser
│   └── ...
├── tests/                  # Test suite
│   ├── TestRunner.bas     # Rubberduck tests
│   └── ...
├── examples/              # Usage examples
│   ├── basic/
│   ├── advanced/
│   └── real-world/
├── docs/                  # Documentation
│   ├── Language reference.md
│   └── assets/
├── README.md
├── CONTRIBUTING.md
└── LICENSE
```

---

## Coding Standards

### VBA Code Style

#### Naming Conventions

```vb
' Classes: PascalCase
Class ASF_Compiler
Class ASF_VM

' Public methods: PascalCase
Public Function CompileExpression(expr As String) As Long

' Private methods: PascalCase with prefix
Private Function ParseToken(token As String) As Variant

' Variables: camelCase
Dim currentToken As String
Dim tokenIndex As Long

' Constants: UPPER_SNAKE_CASE
Const MAX_STACK_SIZE As Long = 1000
Const DEFAULT_TIMEOUT As Long = 30

' Parameters: camelCase
Function ProcessData(inputArray As Variant, filterFunc As Object) As Variant
```

#### Code Formatting

```vb
' Indent with 4 spaces (not tabs)
Public Function Example(param As String) As String
    Dim result As String
    
    If Len(param) > 0 Then
        result = "Valid"
    Else
        result = "Invalid"
    End If
    
    Example = result
End Function

' Use blank lines to separate logical sections
Public Sub ProcessItems()
    ' Variable declarations
    Dim items As Variant
    Dim i As Long
    
    ' Initialization
    items = GetItems()
    
    ' Processing loop
    For i = LBound(items) To UBound(items)
        ProcessItem items(i)
    Next i
    
    ' Cleanup
    Set items = Nothing
End Sub
```

#### Comments

```vb
' Use single-line comments for brief explanations
Dim counter As Long  ' Tracks number of iterations

' Use multi-line comments for complex logic
'------------------------------------------------------------
' Algorithm: Boyer-Moore string search
' Time complexity: O(n/m) average case
' Space complexity: O(1)
'------------------------------------------------------------
Private Function SearchString(text As String, pattern As String) As Long
    ' Implementation
End Function

' Document public APIs with structured comments
'------------------------------------------------------------
' Compiles ASF source code into executable bytecode
'
' Parameters:
'   sourceCode - The ASF script to compile
'
' Returns:
'   Program index for use with Run method
'
' Throws:
'   Compilation error if syntax is invalid
'------------------------------------------------------------
Public Function Compile(sourceCode As String) As Long
```

### ASF Script Style

For scripts in examples and tests:

```js
// Use descriptive variable names
let userCount = 10;  // ✅ Good
let x = 10;          // ❌ Avoid (unless in math contexts)

// Use consistent indentation (2 or 4 spaces)
fun processData(items) {
    let result = items
        .filter(fun(x) { return x > 0 })
        .map(fun(x) { return x * 2 });
    return result;
}

// Add comments for complex logic
fun calculateScore(data) {
    // Weighted average: 70% performance, 30% attendance
    let perfScore = data.performance * 0.7;
    let attScore = data.attendance * 0.3;
    return perfScore + attScore;
}

// Use template literals for readability
let message = `User ${user.name} has ${user.points} points`;  // ✅
let message = 'User ' + user.name + ' has ' + user.points + ' points';  // ❌
```

---

## Testing Guidelines

### Test Organization

Tests are organized using Rubberduck's testing framework:

```vb
'@TestModule
'@Folder("ASF.Tests")

Option Explicit
Option Private Module

Private Assert As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@TestMethod("Arithmetic")
Private Sub TestAddition()
    On Error GoTo TestFail
    
    ' Arrange
    Dim engine As New ASF
    Dim script As String
    script = "return 2 + 3;"
    
    ' Act
    engine.Run engine.Compile(script)
    
    ' Assert
    Assert.AreEqual 5, engine.OUTPUT_
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: " & Err.Description
End Sub
```

### Test Categories

Use Rubberduck annotations to categorize tests:

```vb
'@TestMethod("Arithmetic")     ' Basic math operations
'@TestMethod("Arrays")         ' Array operations
'@TestMethod("Functions")      ' Function definitions and calls
'@TestMethod("Classes")        ' OOP features
'@TestMethod("ControlFlow")    ' if/for/while/switch
'@TestMethod("Integration")    ' VBA integration tests
'@TestMethod("Regression")     ' Bug fix verification
```

### Writing Good Tests

#### Test One Thing

```vb
' ✅ Good - tests one specific behavior
'@TestMethod("Arrays")
Private Sub TestArrayFilter_RemovesItemsNotMatchingPredicate()
    Dim engine As New ASF
    engine.Run engine.Compile("return [1,2,3,4,5].filter(fun(x){return x>2})")
    Assert.AreEqual "[3, 4, 5]", FormatArray(engine.OUTPUT_)
End Sub

' ❌ Bad - tests multiple unrelated things
'@TestMethod("Arrays")
Private Sub TestArrayMethods()
    ' Tests filter, map, reduce all in one test
    ' Hard to debug if it fails
End Sub
```

#### Use Descriptive Names

```vb
' ✅ Good - clear what's being tested
Private Sub TestRegex_MatchesDigitsWithGlobalFlag()
Private Sub TestClass_InheritedMethodOverridesParent()
Private Sub TestClosure_CapturesVariableByReference()

' ❌ Bad - unclear what's being tested
Private Sub TestRegex1()
Private Sub TestClass()
Private Sub TestCase42()
```

#### Test Edge Cases

```vb
' Test normal cases
Private Sub TestArraySlice_ReturnsSubset()

' Test edge cases
Private Sub TestArraySlice_WithNegativeIndices()
Private Sub TestArraySlice_WithEmptyArray()
Private Sub TestArraySlice_WithOutOfBoundsIndex()
Private Sub TestArraySlice_WithStartGreaterThanEnd()
```

#### Arrange-Act-Assert Pattern

```vb
Private Sub TestExample()
    On Error GoTo TestFail
    
    actual = CStr(GetResult("return(1 + 2 * 3);"))
    expected = "7"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
```

```vb
Private Sub if_multiline()
    On Error GoTo TestFail
    Dim globals As ASF_Globals
    GetResult "a=3;" & _
                    "if (a==1) {" & _
                    "  print('one')" & _
                    "} elseif (a==2) {" & _
                    "  print('two')" & _
                    "} elseif (a==3) {" & _
                    "  print('three')" & _
                    "} else {" & _
                    "  print('other')" & _
                    "};" & _
                    "print('end');", True
    Set globals = scriptEngine.GetGlobals
    With globals
        actual = CStr(.gRuntimeLog(.gRuntimeLog.count - 1)) & ", " & CStr(.gRuntimeLog(.gRuntimeLog.count))
    End With
    expected = "PRINT:'three', PRINT:'end'"
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
    Resume TestExit
End Sub
```

### Running Tests

```vb
' Run all tests
' In Rubberduck > Test Explorer > Run All Tests

' Run specific test module
' Right-click module > Run Tests

' Run single test
' Right-click test method > Run Test
```

### Test Coverage

When adding new features:
- Write tests BEFORE implementing the feature (TDD approach)
- Aim for 80%+ code coverage for new functionality
- Test both success and failure paths
- Test edge cases and boundary conditions

---

## Submitting Changes

### Branch Naming

Use descriptive branch names:

```bash
# Feature branches
git checkout -b feature/add-spread-operator
git checkout -b feature/async-patterns

# Bug fix branches
git checkout -b fix/regex-escaping-bug
git checkout -b fix/memory-leak-in-closures

# Documentation branches
git checkout -b docs/update-quick-start
git checkout -b docs/add-performance-guide
```

### Commit Messages

Write clear, descriptive commit messages:

```bash
# ✅ Good commit messages
git commit -m "Add support for spread operator in array literals"
git commit -m "Fix regex escaping bug in replace method"
git commit -m "Update quick start guide with VBA integration example"

# ❌ Bad commit messages
git commit -m "Fix bug"
git commit -m "Update docs"
git commit -m "WIP"
```

**Format:**
```
<type>: <short summary (50 chars or less)>

<optional detailed description>

<optional footer with issue references>
```

**Types:**
- `feat:` - New feature
- `fix:` - Bug fix
- `docs:` - Documentation changes
- `test:` - Adding or updating tests
- `refactor:` - Code refactoring
- `perf:` - Performance improvements
- `style:` - Code style changes (formatting, etc.)

**Example:**
```
feat: Add support for destructuring assignment

Implements basic array destructuring syntax:
  let [a, b, c] = [1, 2, 3]

Does not yet support object destructuring or rest parameters.

Closes #123
```

### Pull Request Process

1. **Update your branch** with latest upstream:
   ```bash
   git fetch upstream
   git rebase upstream/main
   ```

2. **Run all tests** and ensure they pass

3. **Update documentation** if you've changed functionality

4. **Push to your fork:**
   ```bash
   git push origin feature/your-feature-name
   ```

5. **Create Pull Request** on GitHub with:
   - Clear title describing the change
   - Description explaining what and why
   - Reference related issues (e.g., "Closes #123")
   - Screenshots/examples if applicable

6. **Respond to feedback** from code review

7. **Squash commits** if requested before merge

### Pull Request Template

When creating a PR, include:

```markdown
## Description
Brief description of what this PR does

## Type of Change
- [ ] Bug fix
- [ ] New feature
- [ ] Breaking change
- [ ] Documentation update

## Testing
- [ ] All existing tests pass
- [ ] Added new tests for this change
- [ ] Tested manually in Excel/Office

## Checklist
- [ ] Code follows project style guidelines
- [ ] Documentation updated
- [ ] No new warnings introduced
- [ ] Self-review completed

## Related Issues
Closes #(issue number)

## Additional Notes
Any other relevant information
```

---

## Issue Guidelines

### Before Opening an Issue

- Search existing issues to avoid duplicates
- Check if it's already fixed in the latest version
- Gather all relevant information

### Bug Reports

Include:

```markdown
**Description**
Clear description of the bug

**To Reproduce**
1. Create ASF engine
2. Run script: `[paste script here]`
3. Observe error

**Expected Behavior**
What should happen

**Actual Behavior**
What actually happens

**Environment**
- Office Version: Excel 2019
- 32-bit or 64-bit: 64-bit
- Operating System: Windows 10
- ASF Version: v2.0.0

**Additional Context**
- Error messages
- Stack traces (if EnableCallTrace was enabled)
- Screenshots
```

### Feature Requests

Include:

```markdown
**Use Case**
Describe the problem this feature would solve

**Proposed Solution**
How you envision this working

**Example**
```js
// Example of how the feature would be used
let result = newFeature(data);
```

**Alternatives Considered**
Other approaches you've thought about

**Additional Context**
Why this would be valuable to ASF users
```

---

## Community

### Getting Help

- **Documentation:** [Language Reference](docs/Language%20reference.md)
- **Examples:** Browse the `examples/` directory
- **Discussions:** [GitHub Discussions](https://github.com/ECP-Solutions/ASF/discussions)
- **Reddit:** [r/vba](https://reddit.com/r/vba) - search for "ASF"

### Sharing Your Work

Built something cool with ASF? We'd love to feature it:
- Open a "Show and Tell" discussion on GitHub
- Tag your project with `#ASF` on social media
- Consider contributing it to `examples/`

### Recognition

Contributors are recognized in:
- GitHub contributors list
- Release notes for significant contributions
- Project README (for major features)

---

## Release Process

(For maintainers)

1. Update version numbers in code
2. Update CHANGELOG.md
3. Run full test suite
4. Create GitHub release with:
   - Version tag (e.g., v2.1.0)
   - Release notes
   - Compiled examples
5. Announce in community channels

---

## Questions?

Don't hesitate to ask:
- Open a GitHub Discussion for general questions
- Comment on relevant issues
- Tag maintainers if you need specific guidance

**Thank you for contributing to ASF!** 🎉

Your efforts help make VBA development more modern and enjoyable for everyone.

---

*Last updated: January 2026*