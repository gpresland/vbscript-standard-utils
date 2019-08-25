Option Explicit
On Error Resume Next

Wscript.Echo "--------------------------------------------------------------------------------"
Wscript.Echo " MathUtils Tests"
Wscript.Echo "--------------------------------------------------------------------------------"

'-------------------------------------------------------------------------------
Assert Clamp(0, 1, 10) = 1, "Clamp should return 1"
Assert Clamp(5, 1, 10) = 5, "Clamp should return 5"
Assert Clamp(11, 1, 10) = 10, "Clamp should return 10"
Assert Clamp(0.5, 0, 1) = 0.5, "Clamp should return 0.5"
Assert Clamp(0.5, 0.25, 0.75) = 0.5, "Clamp should return 0.5"
Assert Clamp(1, 0.25, 0.75) = 0.75, "Clamp should return 0.75"
Assert Clamp(0, 0.25, 0.75) = 0.25, "Clamp should return 0.25"
