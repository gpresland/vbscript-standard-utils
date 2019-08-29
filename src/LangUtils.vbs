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
        IsNull(Value)
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
