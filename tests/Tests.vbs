Option Explicit

Sub Assert(blnExpression, strDescription)
    If blnExpression Then
        Wscript.Echo "PASS: " & strDescription
    Else
        Wscript.Echo "FAIL: " & strDescription
    End If
End Sub

Sub Include(strScriptName)
    Const ForReading = 1, ForWriting = 2
    Dim objFS : Set objFS = CreateObject("Scripting.FileSystemObject")
    Dim objFile : Set objFile = objFS.OpenTextFile(strScriptName, ForReading)
    ExecuteGlobal objFile.ReadAll()
    objFile.Close
End Sub

Include("dist\StandardUtils.vbs")

Include("tests\ArrayUtils.Test.vbs")
Include("tests\LangUtils.Test.vbs")
Include("tests\MathUtils.Test.vbs")
Include("tests\StringUtils.Test.vbs")
