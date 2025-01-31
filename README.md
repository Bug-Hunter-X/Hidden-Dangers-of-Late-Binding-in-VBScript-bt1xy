# Hidden Dangers of Late Binding in VBScript

This repository demonstrates a common yet easily overlooked error in VBScript: the pitfalls of late binding. Late binding, while offering flexibility, introduces runtime errors if objects or their members are unexpectedly unavailable.

The `bug.vbs` file showcases a scenario where late binding leads to a runtime failure. The `bugSolution.vbs` file provides a corrected version with improved error handling.

## How to Reproduce

1.  Clone this repository.
2.  Run `bug.vbs` (you'll likely encounter a runtime error).
3.  Run `bugSolution.vbs` to see the corrected version.

## Learning Points

- Always handle potential errors when using late binding in VBScript.
- Consider early binding for better performance and error detection during development.
- Employ robust error handling to gracefully manage unexpected situations.