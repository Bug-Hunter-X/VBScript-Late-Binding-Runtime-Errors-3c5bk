Late Binding: In VBScript, objects are often used without explicitly declaring their type. This can lead to runtime errors if the object isn't what you expect.  For example, attempting to use a method on an object that doesn't support it.  This can be hard to debug because the error might not appear until runtime. 
```vbscript
Dim obj
Set obj = CreateObject("Some.Object")
'Assume Some.Object doesn't have a method called "DoSomething"
 obj.DoSomething 
```