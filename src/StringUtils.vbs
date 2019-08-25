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
    	Value Is Nothing Or _
    	(TypeName Is "String" And Len(Trim(Value)) Is 0)
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
