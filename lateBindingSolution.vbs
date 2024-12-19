Dim fso
On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
If Err.Number <> 0 Then
  WScript.Echo "Error: Could not create FileSystemObject.  Check if the Scripting Runtime is correctly installed."
  Err.Clear
  WScript.Quit
End If

' ... rest of your code using fso ...
Set fso = Nothing
On Error GoTo 0