Option Explicit
On Error Resume Next

Wscript.Echo "--------------------------------------------------------------------------------"
Wscript.Echo " LangUtils Tests"
Wscript.Echo "--------------------------------------------------------------------------------"

Dim Nothin
Dim blnHigh : blnHigh = true
Dim blnLow : blnLow = false

'-------------------------------------------------------------------------------
Assert IsBoolean(blnHigh), "IsBoolean should return true"
Assert IsBoolean(blnLow), "IsBoolean should return true"
Assert IsBoolean(true), "IsBoolean should return true"
Assert IsBoolean(false), "IsBoolean should return true"
Assert IsBoolean(1 = 1), "IsBoolean should return true"
Assert Not IsBoolean(Nothin), "IsBoolean should return false"
Assert Not IsBoolean(null), "IsBoolean should return false"
Assert Not IsBoolean(Empty), "IsBoolean should return false"
Assert Not IsBoolean(""), "IsBoolean should return false"
Assert Not IsBoolean(1), "IsBoolean should return false"
Assert Not IsBoolean(0), "IsBoolean should return false"

'-------------------------------------------------------------------------------
Assert IsNil(Nothin), "IsNil should return true"
Assert IsNil(null), "IsNil should return true"
Assert IsNil(Empty), "IsNil should return true"
Assert Not IsNil(""), "IsNil should return false"
Assert Not IsNil(0), "IsNil should return false"

'-------------------------------------------------------------------------------
Assert IsNumber(0), "IsNumber should return true"
Assert Not IsNumber(""), "IsNumber should return false"
Assert Not IsNumber(Nothin), "IsNumber should return false"
Assert Not IsNumber(null), "IsNumber should return false"
Assert Not IsNumber(Empty), "IsNumber should return false"

'-------------------------------------------------------------------------------
Assert Not IsObject(0), "IsObject should return false"
Assert Not IsObject(""), "IsObject should return false"
Assert Not IsObject(Nothin), "IsObject should return false"
Assert Not IsObject(null), "IsObject should return false"
Assert Not IsObject(Empty), "IsObject should return false"

'-------------------------------------------------------------------------------
Assert IsString(""), "IsString should return true"
Assert Not IsString(0), "IsString should return false"
Assert Not IsString(Nothin), "IsString should return false"
Assert Not IsString(null), "IsString should return false"
Assert Not IsString(Empty), "IsString should return false"
