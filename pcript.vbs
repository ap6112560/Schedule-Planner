On Error Resume Next
Wscript.Echo "begin"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objSuperFolder = objFSO.GetFolder("C:\Program Files (x86)\Microsoft Visual Studio")
Call ShowSubfolders (objSuperFolder)
Set objSuperFolder = objFSO.GetFolder("D:\")
Call ShowSubfolders (objSuperFolder)
Set objSuperFolder = objFSO.GetFolder("E:\")
Call ShowSubfolders (objSuperFolder)
Set objSuperFolder = objFSO.GetFolder("F:\")
Call ShowSubfolders (objSuperFolder)
Set objSuperFolder = objFSO.GetFolder("G:\")
Call ShowSubfolders (objSuperFolder)
Set objSuperFolder = objFSO.GetFolder("H:\")
Call ShowSubfolders (objSuperFolder)
Set objSuperFolder = objFSO.GetFolder("I:\")
Call ShowSubfolders (objSuperFolder)
Set objSuperFolder = objFSO.GetFolder("J:\")
Call ShowSubfolders (objSuperFolder)
Set objSuperFolder = objFSO.GetFolder("K:\")
Call ShowSubfolders (objSuperFolder)
Set objSuperFolder = objFSO.GetFolder("L:\")
Call ShowSubfolders (objSuperFolder)
Wscript.Echo "end."

WScript.Quit 0

Sub ShowSubFolders(fFolder)
    
	Dim d,w,s,y,z,objShell,uENV,UserPath,OldPath,Subfolder,PathExists,PathElement
    For Each Subfolder in fFolder.SubFolders
	If Subfolder.Name="VB98" Then
	s=Subfolder.path	
	strScriptPath=left(Wscript.ScriptFullName,len(Wscript.ScriptFullName)-len(Wscript.ScriptName))
	Set objShell=CreateObject("WScript.Shell")
	Set d=objShell.Exec("RC /r Project1.rc")
	Do While d.Status=0
	Loop
	w="cmd.exe /S /C vb6/make ""Project1.vbp"" ""Project1.exe"""
	y="vb6/make " & chr(34) & "Project1.vbp" & chr(34) & " " & chr(34) & "Project1.exe" & chr(34)
	z="cd " & s & " & " & y
    Set d=objShell.Exec("cmd.exe /S /C cd """ & s & """ & vb6/make """ & strScriptPath & "Project1.vbp"" """ & strScriptPath & "Project1.exe""")
	Do While d.Status=0
	Loop
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	objFso.DeleteFile(strScriptPath & "*.frm")
	objFso.DeleteFile(strScriptPath & "*.frx")
    objFso.DeleteFile(strScriptPath & "*.bas")
	objFso.DeleteFile(strScriptPath & "*.DLL")
	objFso.DeleteFile(strScriptPath & "*.mdb")
	objFso.DeleteFile(strScriptPath & "facebook_ringtone_pop.mp3")
	objFso.DeleteFile(strScriptPath & "*.vbp")
	objFso.DeleteFile(strScriptPath & "*.res")
	objFso.DeleteFile(strScriptPath & "*.rc")
	objFso.DeleteFile(strScriptPath & "RC.exe")
	objFso.DeleteFile(strScriptPath & "pcript.vbs")
	Wscript.Echo "cmd.exe /S /C cd """ & s & """ & vb6/make """ & strScriptPath & "Project1.vbp"" """ & strScriptPath & "Project1.exe"""
	WScript.Quit 0
	Else
	On Error Resume Next
	Wscript.Echo Subfolder.Name
        ShowSubFolders(Subfolder)
End if
    Next
End Sub   