Attribute VB_Name = "ModuleMain"
Option Explicit

Sub Main()
    Dim fs As Object
    
    Set fs = CreateObject("Scripting.FileSystemObject")

    fs.CopyFile "C:\texto.txt", "c:\texto2.txt", True
    
    
End Sub
