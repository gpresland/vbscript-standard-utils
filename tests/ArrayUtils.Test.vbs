Option Explicit
On Error Resume Next

Wscript.Echo "--------------------------------------------------------------------------------"
Wscript.Echo " ArrayUtils Tests"
Wscript.Echo "--------------------------------------------------------------------------------"

Dim arrNumbersArray(5)
arrNumbersArray(0) = "A"
arrNumbersArray(1) = "B"
arrNumbersArray(2) = "C"
arrNumbersArray(3) = "D"
arrNumbersArray(4) = "E"

Dim arrEmptyArray(0)

'-------------------------------------------------------------------------------
Assert Head(arrNumbersArray) = "A", "Head should return 'A'"
Assert IsNull(Head(arrEmptyArray)), "Head should return Null"
Head("")
Assert Err.Number = 1, "Head should handle invalid argument"
Err.Clear
Head(Null)
Assert Err.Number = 1, "Head should handle null argument"
Err.Clear

'-------------------------------------------------------------------------------
Assert Includes(arrNumbersArray, "A"), "Includes should return true"
Assert Not Includes(arrNumbersArray, 999), "Includes should return false"
Assert Not Includes(arrEmptyArray, 999), "Includes should return false for empty array"
Call Includes(Null, 999)
Assert Err.Number = 1, "Includes should handle invalid argument"
Err.Clear

'-------------------------------------------------------------------------------
Assert IndexOf(arrNumbersArray, "A") = 0, "IndexOf should return 0"
Assert IndexOf(arrNumbersArray, "E") = 4, "IndexOf should return 4"
Assert IndexOf(arrNumbersArray, 999) = -1, "IndexOf should return -1"
Assert IndexOf(arrEmptyArray, 999) = -1, "IndexOf should handle empty array"
Call IndexOf("", 999)
Assert Err.Number = 1, "IndexOf should handle invalid argument"
Err.Clear

'-------------------------------------------------------------------------------
Assert Last(arrNumbersArray) = "E", "Last should return 5"
Assert IsNull(Last(arrEmptyArray)), "Last should handle empty array"
Err.Clear
Call Last("")
Assert Err.Number = 1, "Last should handle invalid argument"
Err.Clear
Last(Null)
Assert Err.Number = 1, "Last should handle null argument"
Err.Clear

'-------------------------------------------------------------------------------
Dim arrTail : arrTail = Tail(arrNumbersArray)
Assert UBound(arrTail) = UBound(arrNumbersArray) - 1, "Tail should return 1 less than the original array"
Assert arrTail(0) = arrNumbersArray(1), "Tail should have a first element value of " & arrNumbersArray(1)
Assert UBound(Tail(arrEmptyArray)) = 0, "Tail should handle empty array"
Err.Clear
Call Tail("")
Assert Err.Number = 1, "Tail should handle invalid argument"
Err.Clear
Tail(Null)
Assert Err.Number = 1, "Tail should handle null argument"
Err.Clear
