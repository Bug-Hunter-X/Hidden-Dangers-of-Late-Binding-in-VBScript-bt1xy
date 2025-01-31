Late Binding: VBScript allows late binding, meaning that you don't have to declare object variables before using them. While convenient, this can lead to runtime errors if the object isn't available or doesn't have the method or property you're trying to access.  Example:
```vbscript
Set obj = CreateObject("Some.Object")
If Err.Number <> 0 Then
   ' Handle Error 
End If
obj.DoSomething  'Error if Some.Object or DoSomething doesn't exist
```