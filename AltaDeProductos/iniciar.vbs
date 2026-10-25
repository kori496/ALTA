' Abre Alta de productos sin mostrar la ventana negra de la consola.
Set fso = CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
If Not fso.FileExists(dir & "\node\node.exe") Then
  MsgBox "Falta node\node.exe en " & dir & ". Vuelve a descomprimir el .zip completo.", 16, "Alta de productos"
  WScript.Quit 1
End If
Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = dir
sh.Run """" & dir & "\node\node.exe"" """ & dir & "\app\server.js""", 0, False
