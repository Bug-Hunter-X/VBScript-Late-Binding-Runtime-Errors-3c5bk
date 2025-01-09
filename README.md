# VBScript Late Binding Runtime Error Example

This repository demonstrates a common runtime error in VBScript caused by late binding. Late binding, while offering flexibility, can lead to unexpected errors if an object doesn't support a called method. The `bug.vbs` file shows the problematic code.  The solution, `bugSolution.vbs`, demonstrates using early binding to prevent this error. 

## How to Reproduce

1.  Save the code in `bug.vbs`.
2.  Run the script.  Observe the runtime error.

## Solution

The `bugSolution.vbs` file provides a solution using early binding, which requires explicit type declaration, preventing runtime errors from unsupported methods.