# The `spec` folder

## Overview

Each integration test consists of one or more *specs*, or specifications.

## Launching tests

The tests may be initiated by running directly by double clicking or from a command prompt, type `cscript <filename>`.
'Suites' or collections of tests may be run using the `TestLauncher*` files in the [suite](suite) folder.

## Different kinds of tests

- Files named *\*.spec.sk.vbs* use the SendKeys method and should be used with caution because the tests simulate user input/keystrokes.
- Files named *\*.spec.elev.vbs* are intended to be run from an elevated command prompt.
- Files named *\*.spec.std+elev.vbs* are intended to be run from either an elevated command prompt or from a non-elevated command prompt.
- Files named *\*.spec.wow.vbs* are intended to be run on a 64-bit system to test the specified class with the 32-bit executable(s).
