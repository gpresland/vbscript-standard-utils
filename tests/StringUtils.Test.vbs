Option Explicit
On Error Resume Next

Wscript.Echo "--------------------------------------------------------------------------------"
Wscript.Echo " StringUtils Tests"
Wscript.Echo "--------------------------------------------------------------------------------"

'-------------------------------------------------------------------------------
Assert Not IsNullOrEmpty("abc"), "IsNullOrEmpty should return false"
Assert Not IsNullOrEmpty("   "), "IsNullOrEmpty should return false"
Assert IsNullOrEmpty(""), "IsNullOrEmpty should return true"
Assert IsNullOrEmpty(null), "IsNullOrEmpty should return true"
Assert IsNullOrEmpty(Empty), "IsNullOrEmpty should return true"
