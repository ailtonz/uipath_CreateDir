# Emulação da função abaixo:
~~~VB
Public Function CreateDir(strPath As String)  
Dim elm As Variant  
Dim strCheckPath As String: strCheckPath = ""  
    For Each elm In Split(strPath, "\")  
    strCheckPath = strCheckPath & elm & "\"  
    If Len(Dir(strCheckPath, vbDirectory)) = 0 Then MkDir strCheckPath  
    Next  
End Function
~~~
