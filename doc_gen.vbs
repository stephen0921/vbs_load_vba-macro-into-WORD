Dim WshShell
set WshShell = CreateObject("wscript.Shell")
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Word\Security\AccessVBOM",1,"REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Word\Security\AccessVBOM",1,"REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\13.0\Word\Security\AccessVBOM",1,"REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\Security\AccessVBOM",1,"REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Word\Security\AccessVBOM",1,"REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Word\Security\AccessVBOM",1,"REG_DWORD"
Dim oWord, oModule
Set oWord = CreateObject("Word.application")
oWord.Visible = True
oWord.DisplayAlerts = True
oWord.Documents.Add
Set docActive = oWord.ActiveDocument
Set oModule = docActive.VBProject.VBComponents.Add(1)
set fso=createobject("scripting.filesystemobject")  
set file=fso.opentextfile("word.vb")  
strCopy = file.readall  
file.close
strCode = strCopy
oModule.CodeModule.AddFromString strCode
oWord.run "all_module_regs"
oWord.Quit
