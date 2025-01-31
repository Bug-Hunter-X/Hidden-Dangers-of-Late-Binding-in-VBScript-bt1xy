Early Binding and Comprehensive Error Handling:
```vbscript
On Error Resume Next  'Better practice to handle errors specifically
Dim obj As Object
Set obj = GetObject("Some.Object") 'Use GetObject to safely get existing object
If Err.Number <> 0 Then
  MsgBox "Error creating or getting object: " & Err.Description, vbCritical
  Err.Clear
  Exit Sub 'Or take another appropriate action
End If

If TypeName(obj) = "Some.Object" Then
    If obj.DoSomething Then
        'Process Result
    Else
        MsgBox "DoSomething method failed: " & Err.Description, vbExclamation
    End If 
Else
    MsgBox "Object type mismatch", vbExclamation
End If
Set obj = Nothing
On Error GoTo 0
```
Note: Replace "Some.Object" with the actual object and adjust the error handling to your specific application.