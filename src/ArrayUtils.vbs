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
