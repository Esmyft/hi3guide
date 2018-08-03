Option Explicit

Dim scriptdir
Dim fso
set fso = CreateObject("Scripting.FileSystemObject")
dim CurrentDirectory
CurrentDirectory = fso.GetAbsolutePathName(".")

wscript.echo "lol"
scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
wscript.echo scriptdir
wscript.echo CurrentDirectory 
