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
