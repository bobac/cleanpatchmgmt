'===================================================================================================================
'SCRIPT: Delete Patch Management remnants, after disabling the feature within the MAX RM Dashboard

'Script intended for the MAX RemoteManagement Dashboard
'===================================================================================================================

'Initial Variable Declarations
Set objShell = WScript.CreateObject ("WScript.shell")
UnInsID1 = "{5D56359C-92E3-4306-A48D-7F95B8D0D48D}"
MsiExStrRmv1 = "Msiexec.exe /X" & UnInsID1
UnInsID2 = "{A0707C59-4B32-48B8-94ED-73BB68E1C569}"
MsiExStrRmv2 = "Msiexec.exe /X" & UnInsID2
UnInsiD3 = "{54233984-5466-45CE-BB95-682627A85D1D} "
MsiExStrRmv3 = "Msiexec.exe /X" & UnInslD3
ThTwbt = 0
SxFrbt = 0

'==============================Msiexec removal processes, utilizing UninstallString Identifiers=====================
result1 = objShell.run (MsiExStrRmv1 & " /qn", 1, True)
If result1 = 0 Then
  Wscript.Echo "Successfully Ran: " & MsiExStrRmv1 & vbCRLF & "Proceeding with cleanup..." & vbCRLF
Else
    Wscript.Echo "UninstallString ID matching:  " & UnInsID1 & " was not found.  Proceeding with cleanup..." & vbCRLF
End If


result2 = objShell.run (MsiExStrRmv2 & " /qn", 1, True)
If result2 = 0 Then
    Wscript.Echo "Successfully Ran: " & MsiExStrRmv2 & vbCRLF & "Proceeding with cleanup..." & vbCRLF
Else
    Wscript.Echo "UninstallString ID matching:  " & UnInsID2 & " was not found.  Proceeding with cleanup..." & vbCRLF
End If

result3 = objShell.run (MsiExStrRmv3 & " /qn", 1, True)
If result3 = 0 Then
    Wscript.Echo "Successfully Ran: " & MsiExStrRmv3 & vbCRLF & "Proceeding with cleanup..." & vbCRLF
Else
    Wscript.Echo "UninstallString ID matching:  " & UnInsID3 & " was not found.  Proceeding with cleanup..." & vbCRLF
End If
'===================================================================================================================


'=========================Operating System Check(s) for Patch Management Folder Cleanup=============================
'32 or 64 bit?
If objShell.ExpandEnvironmentStrings("%PROGRAMFILES(x86)%")="%PROGRAMFILES(x86)%" then
    path = objShell.ExpandEnvironmentStrings("%PROGRAMFILES%")
    path2 = Left(path, Len(path) - 13)
    ThTwbt = 1
Else
    path = objShell.ExpandEnvironmentStrings("%PROGRAMFILES(x86)%")
    path2 = Left(path, Len(path) - 19)
    SxFrbt = 1
End If

'Pre or Post Vista? & Set Directory for Removal
strComputer = "." 
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")  
Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")   
For Each os in oss  
    OSvar = os.Caption  
Next
PreVstaFldrPath = path2 & "Documents and Settings\All Users\Application Data\"
PostVstaFldrPath = path2 & "ProgramData\"
WkstnXP = "XP"
Srvr2003 = "2003"

If InStr(OSVar,WkstnXP) Then
    PMFLDRRmv2 = PreVstaFldrPath & "GFI\LanGuard 11"
ElseIf Instr(OSVar,Srvr2003) Then
    PMFLDRRmv2 = PreVstaFldrPath & "GFI\LanGuard 11"
Else
    PMFLDRRmv2 = PostVstaFldrPath & "GFI\LanGuard 11"
End If
'===================================================================================================================


'================Remove Directory: C:Program Files(x86 for 64bit)\Advanced Monitoring Agent\patchman================
PMFLDRRmv1 = path & "\Advanc~1\patchman"
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists(PMFLDRRmv1) Then
    Wscript.Echo "Folder: " & """" & PMFLDRRmv1 & """" & " exists.  Removing..." & vbCRLF
    Set objFolder = objFSO.GetFolder(PMFLDRRmv1)
    objFolder.Delete(True)
Else
    Wscript.Echo "Folder: " & """" & PMFLDRRmv1 & """" & " does not exist.  Proceeding with cleanup..." & vbCRLF
End If
'===================================================================================================================


'=======================Remove Directory: ...\GFI\LanGuard 11 (appropriate for environment)=========================
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists(PMFLDRRmv2) Then
    Wscript.Echo "Folder: " & """" & PMFLDRRmv2 & """" & " exists.  Removing..." & vbCRLF
    Set objFolder = objFSO.GetFolder(PMFLDRRmv2)
    objFolder.Delete(True)
Else
    Wscript.Echo "Folder: " & """" & PMFLDRRmv2 & """" & " does not exist.  Proceeding with cleanup..." & vbCRLF
End If
'===================================================================================================================


'======================Registry Cleanup: HKEY_LOCAL_MACHINE\SOFTWARE\[Wow6432Node\]GFI\LNSS11=======================
If (ThTwbt = 1) Then
    Wscript.Echo "Removing Registry Key: HKEY_LOCAL_MACHINE\SOFTWARE\GFI\LNSS11 " & vbCRLF & "...and subkeys, if present..." & vbCRLF
    objShell.run("REG.EXE DELETE ""HKEY_LOCAL_MACHINE\SOFTWARE\GFI\LNSS11"" /f")
ElseIf (SxFrbt = 1) Then
    Wscript.Echo "Removing Registry Key: HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\GFI\LNSS11 " & vbCRLF & "...and subkeys, if present..." & vbCRLF
    objShell.run("REG.EXE DELETE ""HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\GFI\LNSS11"" /f")
End If
set objShell = nothing
'===================================================================================================================


'===================================================Reboot Notifier=================================================
Wscript.Echo "Please reboot as soon as possible in order to complete the Patch Management removal process."
Wscript.Quit(0)
'===================================================================================================================