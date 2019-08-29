'###############################################################################
'#
'#  Description:
'#
'#  Generic Array functions.
'#
'#  Author(s)   : Greg Presland
'#
'#  Created     : 24 Aug 2019
'#  Last Edited : 24 Aug 2019
'#
'###############################################################################

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Gets the first element of array.
'     Arr : The array to query.
' Returns the first element of array, otherwise null if no element.
'
Public Function Head(Arr)
    If Not IsArray(Arr) Then
        Err.Raise 1, "Head: Argument 'Arr' must be of type Array. Was " & TypeName(Arr) & "."
        Exit Function
    End If
    If UBound(Arr) = 0 Then
        Head = null
        Exit Function
    End If
    Head = Arr(LBound(Arr))
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Checks if value is in array.
'     Arr   : The array to search.
'     Value : The value to search for.
' Returns true if value is found, else false.
'
Public Function Includes(Arr, Value)
    If Not IsArray(Arr) Then
        Err.Raise 1, "Includes: Argument 'Arr' must be of type Array. Was " & TypeName(Arr) & "."
        Exit Function
    End If
    Includes = UBound(Filter(Arr, Value)) > -1
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Gets the index at which the first occurrence of value is found in array.
'     Arr   : The array to search.
'     Value : The value to search for.
' Returns the index of the matched value, else -1.
'
Public Function IndexOf(Arr, Value)
    If Not IsArray(Arr) Then
        Err.Raise 1, "IndexOf: Argument 'Arr' must be of type Array. Was " & TypeName(Arr) & "."
        Exit Function
    End If
    Dim i
    For i = 0 To UBound(Arr, 1)
        If Arr(i) = Value Then
            IndexOf = i
            Exit Function
        End If
    Next
    IndexOf = -1
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Gets the last element of array.
'     Arr : The array to query.
' Returns the last element of array.
'
Public Function Last(Arr)
    If Not IsArray(Arr) Then
        Err.Raise 1, "Last: Argument 'Arr' must be of type Array. Was " & TypeName(Arr) & "."
        Exit Function
    End If
    If UBound(Arr) = 0 Then
        Last = null
        Exit Function
    End If
    Last = Arr(UBound(Arr) - 1)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Gets all but the first element of array.
'     Arr : The array to query.
' Returns the slice of array.
'
Public Function Tail(Arr)
    If Not IsArray(Arr) Then
        Err.Raise 1, "Tail: Argument 'Arr' must be of type Array. Was " & TypeName(Arr) & "."
        Exit Function
    End If
    Dim intStartIndex : intStartIndex = LBound(Arr, 1)
    Dim intEndIndex : intEndIndex = UBound(Arr, 1)
    If intStartIndex >= intEndIndex Then
        Exit Function
    End If
    Dim intLength : intLength = intEndIndex - intStartIndex - 1
    Dim arrResult() : Redim arrResult(intLength)
    Dim i
    For i = intStartIndex + 1 To intEndIndex
        arrResult(i - 1) = Arr(i)
    Next
    Tail = arrResult
End Function
'###############################################################################
'#
'#  Description:
'#
'#  Generic langauge functions.
'#
'#  Author(s)   : Greg Presland
'#
'#  Created     : 24 Aug 2019
'#  Last Edited : 24 Aug 2019
'#
'###############################################################################

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Checks if value is classified as a boolean primitive or object.
'     Value : The value to check.
' Returns true if value is a boolean, else false.
'
Public Function IsBoolean(Value)
    IsBoolean = TypeName(Value) = "Boolean"
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Checks if value is null or undefined.
'     Value : The value to check.
' Returns true if value is nullish, else false.
'
Public Function IsNil(Value)
    IsNil = _
        IsEmpty(Value) Or _
        IsNull(Value)' Or _
        ' Value Is Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Checks if value is classified as a Number primitive or object.
'     Value : The value to check.
' Returns true if value is a number, else false.
'
Public Function IsNumber(Value)
    Dim name: name = TypeName(Value)
    IsNumber = _
    	name = "Byte" Or _
        name = "Decimal" Or _
        name = "Double" Or _
        name = "Integer" Or _
        name = "Long" Or _
        name = "SByte" Or _
        name = "Short" Or _
        name = "Single" Or _
        name = "UInteger" Or _
        name = "Ulong" Or _
        name = "UShort"
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Checks if value is the language type of Object.
'     Value : The value to check.
' Returns true if value is an object, else false.
'
Public Function IsObject(Value)
    IsObject = TypeName(Value) = "Object"
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Checks if value is classified as a String primitive or object.
'     Value : The value to check.
' Returns true if value is a string, else false.
'
Public Function IsString(Value)
    IsString = TypeName(Value) = "String"
End Function
'###############################################################################
'#
'#  Description:
'#
'#  Generic math functions.
'#
'#  Author(s)   : Greg Presland
'#
'#  Created     : 24 Aug 2019
'#  Last Edited : 24 Aug 2019
'#
'###############################################################################

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Clamps number within the inclusive lower and upper bounds.
'     Number : The number to clamp.
'     Lower  : The lower bound.
'     Upper  : The upper bound.
' Returns the first element of array.
'
Public Function Clamp(Number, Lower, Upper)
    If Number < Lower Then
        Clamp = Lower
    ElseIf Number > Upper Then
        Clamp = Upper
    Else
        Clamp = Number
    End If
End Function
'###############################################################################
'#
'#  Description:
'#
'#  Generic String functions.
'#
'#  Author(s)   : Greg Presland
'#
'#  Created     : 24 Aug 2019
'#  Last Edited : 24 Aug 2019
'#
'###############################################################################

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Checks if a string is null or empty.
'     Value : The string to check.
' Returns true if empty, otherwise false.
'
Function IsNullOrEmpty(Value)
    IsNullOrEmpty = _
    	IsEmpty(Value) Or _
    	IsNull(Value) Or _
    	(TypeName(Value) = "String" And Len(Value) = 0)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Checks if a string is null or whitespace.
'     Value : The string to check.
' Returns true if empty, otherwise false.
'
Function IsNullOrWhiteSpace(Value)
    IsNullOrWhiteSpace = _
    	IsEmpty(Value) Or _
    	IsNull(Value) Or _
    	(TypeName(Value) = "String" And Len(Trim(Value)) = 0)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' The PadEnd() method pads the current string with another string (multiple
' times, if needed) until the resulting string reaches the given length. The
' padding is applied from the start of the current string.
'     Value     : The string to pad.
'     Length    : The length to pad to.
'     Character : The Character to use for padding.
' Returns the padded string.
'
Function PadEnd(Value, Length, Character)
    If Len(Character) < 1 Then
    	Raise
    End If
    Value = CStr(Value)
    Do While Len(Value) < Length
        Value = Value & Character
    Loop
    PadEnd = Value
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' The PadStart() method pads the current string with another string (multiple
' times, if needed) until the resulting string reaches the given length. The
' padding is applied at the end of the current string.
'     Value     : The string to pad.
'     Length    : The length to pad to.
'     Character : The character to use for padding.
' Returns the padded string.
'
Function PadStart(Value, Length, Character)
    If Len(Character) < 1 Then
        Raise
    End If
    Value = CStr(Value)
    Do While Len(Value) < Length
        Value = Character & Value
    Loop
    PadStart = Value
End Function
