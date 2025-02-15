' ***************************************************************************
'
' This script lets show the screen with the graphics card (GPU)
' unscaled and centered
' Dieses Skript l�sst den Bildschirm mit der Grafikkarte (GPU)
' unskaliert und mittig darstellen
'
' Author: Dragodraki
' Last updated: 2024
' Homepage: https://github.com/Dragodraki
' OS: Windows Vista+
' Hardware: nearly all graphic cards from AMD, NVIDIA and Intel
' (appropriate driver need to be installed)
' Use at your own risk - I assume no liability for damages of any kind
' caused by this script or associated components!
'
' ***************************************************************************

Option Explicit
On Error Resume Next

' Nicht als Variable definieren: wscript, ScriptFullName !!!
' F�r Textbox statt "wscript.echo" den Befehl "msgbox" verwenden !!!
' Der msgbox/popup-Parameter +131072 erh�ht zwar die Textaufl�sung, aber verhindert die nach definierter Zeit Schlie�en-Funktion!


Dim fso, warningbox, file, folder, folder2, wshell, wshshell, objShell, myKey1, myKey2, strPfad, Systemdrive, Username, AppData, oFSO, oFolder, sDirectoryPath, oFileCollection, oFile, sDir, iDaysOld, sExt, text, text2, fcreate, fdesc, WshSheg, WshShell1, WshShell2, WshShell3, ws, oShell, oWindowList, EXPLORERUMGEBUNGSVARIABLE, EXPLORERUMGEBUNGSVARIABLE2, oWindow, KeinExplorer, user, objFSO, RIGHTHERE
Dim objAllUsersProgramsFolder, strAllUsersProgramsPath, str, objfolder, objFolderItem, colVerbs, objVerb, objShortcut, objFileSystem, sRegFile, TEST, owindowsFolderPath, newDIR, ziel, MyComputer, strComputer, objWMIService, colOperatingSystems, objOperatingSystem, colProcess, objProcess, strProcess, strWMIQuery
Dim sProcessName, sComputer, oWmi, colProcessList, Repeater, objFile, objReg, objSystemInfo, strHostname, startexe, objstream, strArgs, target_folder, fs, TEMP, USERPROFILE, ALLUSERSPROFILE, LOCALAPPDATA, SYSTEMROOT, RIGHTHEREnoBACKSLASH, desktopPath, link, filesys, PARENT_RIGHTHERE, strRoot, strModify, shell, FULLSCRIPTNAME, SCRIPTFILENAME, key, output, gameexe, RIGHTHEREDoubleSlash, XPmode, exelostbox, KompatibilitaetValue, ScriptingHostProcessName, oReg, strKeyPath, linkname, linkname2, linkname3, linkname4, linkname5, PARENT_RIGHTHEREDoubleSlash, search, skriptfilename, iconpfad, strValueName, dwValue, AlleRunasRobEintraege, NeueRunasRobEintraege
dim ResetteRegistrySchluessel, NeuerStringName_1, NeuerStringWert_1, NeuerStringName_2, NeuerStringWert_2, NeuerStringName_3, NeuerStringWert_3, NeuerStringName_4, NeuerStringWert_4, NeuerStringName_5, NeuerStringWert_5, NeuerDwordName_1, NeuerDwordWert_1, NeuerDwordName_2, NeuerDwordWert_2, NeuerDwordName_3, NeuerDwordWert_3, NeuerDwordName_4, NeuerDwordWert_4, NeuerDwordName_5, NeuerDwordWert_5, Aktiviere_HKCU, Aktiviere_HKLM, strValue, Scriptplayer, DeleteAproDings, strContents, strFirstLine, strNewContents, Aktiviere_ALL_HKCU, oNetwork, NeuerStringWert_1A, NeuerStringWert_2A, NeuerStringWert_3A, NeuerDwordWert_1A, NeuerDwordWert_2A, NeuerDwordWert_3A, drive, GetDriveLetter, label, labelfilename ,mount_disc, GetDriveLetterRaw, Resolution, RIGHTHEREDoubleSlashNoBackslash, WindowsVersionName, WMI, OSs, OS, dtmConvertedDate, objOperatingSyst, value, NewStringWindowsVersion, MajorWindowsVersion, objArgsPScript, objArgsWScript, objArgsCScript, objArgs, MitParameterGestartet, ArgumentTotal, Arg, f, app, ProzessUNDTitelOffen, ProzessUNDPfadOffen, ExternalPathCheck, TitleCheck, ProzessEgal_PfadTitle_Warten, AufAlleProzesseVBSpfadWarten, TempFolderName, BenutzerAusListe, infobox, USERDOMAIN, LASTSHOWEDUSER
dim LastResolutionShowed, AllResolutionList, AndereAufloesung, Mindestbreite, Mindesthoehe, Seitenverhaeltnis, Result, Aufloesungstest, Horizontale, Vertikale, AufloesungCheckpoint, AufloesungAlternative, AufloesungAlternativeCheckpoint, FullPath, arr, dir, path, Counter, pReg, qReg, rReg, sReg, objRegistry, sKeyPath, sSubkey, tSubkey, uSubkey, oSubkey, pSubkey, aSubKeys, bSubKeys, cSubKeys, dSubKeys, eSubKeys, sTmpValueName, sTmpKeyName1, sTmpKeyName2, sTmpKeyName3, values, Return, Log, logpath, logSubkey, logvalues, bArray, BinaryPositionNumberToChange, BinaryReplacement, AMDPath, ScalingProzedurNotwendig
dim ColOSs, OSType, service, AppName, Process, AppName2, strDelete, arrProcesses, arrWindowTitles, i, intProcesses, strProcList, TITLENAME2KILL, Informationsbox, strApplication, colItems, ProcessName, Apptitle, objApp, PID, ProcessPfad, nPos, strList, p, q, strKonto, strFileName, NvidiaTest_1, NvidiaTest_2, RegistryFolderToUnlock, StikyNotRestart, ProcessPfad2, AMDLaunchTask, NVIDIALaunchTask, NVIDIAPath, isServiceInstalled, svcQry, objOutParams, objSvc, SucheServiceName, ClassicNotes, BuildInNotes




'--- ERKENNE WINDOWS-VERSION (BEI WINDOWS XP GEHE IN DEN XP-MODUS)

Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set objFSO = CreateObject("Scripting.FileSystemObject")


' Variante A - Pr�fe per Versionsnummer (eigentlich sicheres Verfahren, solange OS-Nummer nicht gr��er als 100 Millionen ist)

'System nicht in Abfrage enthalten! Windows-Version wird trotzdem angezeigt
Const COMPUTER = "."
Set WMI = GetObject("winmgmts:{impersonationLevel=impersonate}\\" & COMPUTER & "\root\cimv2")
Set OSs = WMI.ExecQuery("Select * from Win32_OperatingSystem")
For Each OS in OSs 
If IsNull(WindowsVersionName) Or IsEmpty(WindowsVersionName) Or WindowsVersionName= "" Then
    'msgbox "System nicht in Abfrage enthalten!" & vbCrLf & "Betriebssystem: " & OS.Caption
else
    'msgbox WindowsVersionName
end if
Next
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
strComputer = "."
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
For Each os in oss
  value = os.Version
NewStringWindowsVersion="."
MajorWindowsVersion=left(value, instr(value, NewStringWindowsVersion)-1)
'msgbox MajorWindowsVersion
MajorWindowsVersion = MajorWindowsVersion + 100000006
if MajorWindowsVersion < "100000012" then
XPmode = "Ja"
'msgbox "Windows XP"
else
end if
Next
'msgbox "Ergebnis der Variante A: " & XPmode


' Variante B - Pr�fe per Namen (wenn hier Windows 7/8/10 vorliegt, �berschreibe Ergebnis der Variante A)

if (instr(MyComputer, "SV") OR (instr(MyComputer, "TB"))) then
		Else	
		strComputer = "."
		Set objWMIService = GetObject("winmgmts:" _
			& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

		Set colOperatingSystems = objWMIService.ExecQuery _
			("Select * from Win32_OperatingSystem")
'Windows 7
		For Each objOperatingSystem in colOperatingSystems
			  if instr(objOperatingSystem.Caption, "7") Then
			WindowsVersionName = "Windows 7"
			XPmode = "Nein"
			'msgbox "Win 7 detected"
			Else
			End If
		Next
'Windows 8
		For Each objOperatingSystem in colOperatingSystems
			  if instr(objOperatingSystem.Caption, "8") Then
			WindowsVersionName = "Windows 8"
			XPmode = "Nein"
			'msgbox "Win 8 detected"
			Else
			End If
		Next
'Windows 10
		For Each objOperatingSystem in colOperatingSystems
			  if instr(objOperatingSystem.Caption, "10") Then
			WindowsVersionName = "Windows 10"
			XPmode = "Nein"
			'msgbox "Win 10 detected"
			Else
			End If
			Next
'Windows XP
		For Each objOperatingSystem in colOperatingSystems
			  if instr(objOperatingSystem.Caption, "XP") Then
			XPmode = "Ja"
			WindowsVersionName = "Windows XP"
			'msgbox "Win XP detected"
			Else
			End If
		Next

end if
'msgbox "Ergebnis der Variante B: " & XPmode





'--- DEFINIERE UMGEBUNGSVARIABLEN
SET WshShell = CreateObject( "WScript.Shell" )
Userdomain = WshShell.Environment("Process")("USERDOMAIN")
SET WshShell = CreateObject ("WScript.Shell")
AppData = WshShell.Environment("Process")("APPDATA")
SET WshShell = CreateObject ("WScript.Shell")
Systemdrive = WshShell.Environment("Process")("SYSTEMDRIVE")
SET WshShell = CreateObject ("WScript.Shell")
Username = WshShell.Environment("Process")("USERNAME")
SET objShell = Createobject("wscript.shell")
Userprofile = WshShell.Environment("Process")("USERPROFILE")
SET WshShell = CreateObject ("WScript.Shell")
Allusersprofile = WshShell.Environment("Process")("ALLUSERSPROFILE")
SET objShell = CreateObject("Shell.Application")
Localappdata = objShell.Namespace(&H1c&).Self.Path	' so ist LOCALAPPDATA auch in Windows XP m�glich :)
Set objShell  = CreateObject("Scripting.FileSystemObject")
Temp = CreateObject("Scripting.FileSystemObject").GetParentFolderName(LOCALAPPDATA)
Temp = TEMP & "\Temp"				' TEMP-Ordner f�r Windows XP (wechselt automatisch bei h�herem OS)
if TEMP = USERPROFILE & "\AppData\Temp" then
' msgbox "Laut Variableninhalt handelt es sich um Windows > XP und daher wird die TEMP-Umgebungsvariable angepasst."
SET WshShell = CreateObject ("WScript.Shell")
Temp = WshShell.Environment("Process")("TEMP")		' TEMP-Ordner f�r Windows 7,8,10
else
' XPmode = "Ja"
end if
SET WshShell = CreateObject ("WScript.Shell")
Systemroot = WshShell.Environment("Process")("SYSTEMROOT")
SET WshShell = CreateObject ("WScript.Shell")
FULLSCRIPTNAME = ScriptFullName
if FULLSCRIPTNAME = "" then 
FULLSCRIPTNAME = Wscript.ScriptFullName
end if
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(ScriptFullName)
SCRIPTFILENAME = objFSO.GetFileName(objFile)
if SCRIPTFILENAME = "" then 
SCRIPTFILENAME = Wscript.ScriptName
end if
set WshShell = Createobject("wscript.shell")
RIGHTHERE = WshShell.CurrentDirectory & "\"
if XPmode = "Ja" then
search="\" & SCRIPTFILENAME
'RIGHTHERE=left (FULLSCRIPTNAME, instr(FULLSCRIPTNAME, search)-0)
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(RIGHTHERE) Then
  Set f = fso.GetFile(RIGHTHERE)
ElseIf fso.FolderExists(RIGHTHERE) Then
  Set f = fso.GetFolder(RIGHTHERE)
Else
  'path doesn't exist
End If
Set app = CreateObject("Shell.Application")
' Mache in XP aus der verk�rzten Pfadangabe RIGHTHERE den vollen...
RIGHTHERE = app.NameSpace(f.ParentFolder.Path).ParseName(f.Name).Path & "\"
end if
SET WshShell = CreateObject ("WScript.Shell")
RIGHTHEREnoBACKSLASH = Left(RIGHTHERE, Len(RIGHTHERE) -1)
Set filesys = CreateObject("Scripting.FileSystemObject")
PARENT_RIGHTHERE = CreateObject("Scripting.FileSystemObject").GetParentFolderName(RIGHTHERE)
RIGHTHEREDoubleSlash=Replace(RIGHTHERE,"\","\\")
PARENT_RIGHTHEREDoubleSlash=Replace(PARENT_RIGHTHERE,"\","\\")
strKonto = CreateObject("WScript.Network").UserName
SET WshShell = CreateObject ("WScript.Shell")
RIGHTHEREDoubleSlashNoBackslash = Left(RIGHTHEREDoubleSlash, Len(RIGHTHEREDoubleSlash) -2)





' --- Adminrechte pr�fen

ArgumentTotal = ""
If CSI_IsAdmin Then output = "YES"
If NOT CSI_IsAdmin Then output = "NO"
' msgbox output	'Message-Box, ob ich Administrator-Privilegien habe oder nicht (YES=Adminrechte, NO=Userrechte)
Function CSI_IsAdmin()
  CSI_IsAdmin = False
  On Error Resume Next
  key = CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
  If err.number = 0 Then CSI_IsAdmin = True
End Function

if not output = "YES" then
Set fso = CreateObject("Scripting.FileSystemObject") 
warningbox = MsgBox("Bitte das Skript nur mit Adminrechten starten, dessen EXE (wscript.exe) im selben Pfad liegt wie die VBS-Datei!",0+16+4096+256+65536+131072,"Debugging-Error")

else


if output = "YES" then

RequireAdmin






' ------------------------------------------------------------------------------------------
' DAS EIGENTLICHE SKRIPT
' ------------------------------------------------------------------------------------------


' MUSS MIT ADMINISTRATORRECHTEN LAUFEN!



'--- Variablen setzen und Funktionenzugeh�rigkeiten erstellen
Const cHKCU = &H80000001
Const cHKLM = &H80000002
Set fso = CreateObject("Scripting.FileSystemObject")
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
Set pReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
Set qReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
Set rReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
Set sReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
Set oShell = CreateObject("WScript.Shell")


'--------------------------------------
'AMD

sKeyPath = "SYSTEM\CurrentControlSet\Control\Class"

oReg.EnumKey cHKLM, sKeyPath, aSubKeys
If Not (isnull(aSubKeys)) Then
    For Each sSubkey In aSubKeys
        sTmpKeyName1 = sKeyPath & "\" & sSubKey
        pReg.EnumKey cHKLM, sTmpKeyName1, bSubKeys
        If Not (isnull(bSubKeys)) Then
        For Each tSubkey In bSubKeys
	'if tSubkey = "0000" then
	sTmpKeyName2 =  sTmpKeyName1 & "\" & tSubkey
	'msgbox sTmpKeyName2

               qReg.EnumValues cHKLM, sTmpKeyName2, cSubKeys
               If Not (isnull(cSubKeys)) Then
               For Each oSubkey In cSubKeys
	     'oSubkey = "GPUScaling"		' plus eine unbekannte Zahl dahinter (z.B. "GPUScaling84")
	     if Left(oSubkey, 10) = "GPUScaling" then
	     sTmpValueName =  sTmpKeyName2 & "\" & oSubkey	 
	     'msgbox sTmpValueName

	    ' Lese zuerst den Binary-String aus und schaue, ob �nderung notwendig ist
	    Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    path = sTmpKeyName2
	    oReg.GetBinaryValue cHKLM, path, oSubkey, bArray
	    'msgbox bArray(1 - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    If not bArray(1 - 1) = 1 Then
	    ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung �ndern."
		Set filesys = CreateObject("Scripting.FileSystemObject")
		RegistryFolderToUnlock = CreateObject("Scripting.FileSystemObject").GetParentFolderName(sTmpValueName)
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & RegistryFolderToUnlock & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & RegistryFolderToUnlock & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
	    values = Array(01,00,00,00)
	    Return = 1
	    Return = objRegistry.SetBinaryValue(cHKLM, path, oSubkey, values)
	    If (Return = 0) Then
    	    'msgbox "Binary value added successfully"
	    Else
    	    ' An error occurred
    	    'msgbox "Binary could not be added!"
	    End If	    'Meldung, ob Bin�rwert in Registry geschrieben wurde, wurde ausgegeben
	    End if	    ' Pr�fung, ob Skalierung ge�ndert werden muss, abgeschlossen
	     else	'Erste 10 Buchstaben des Werts sind nicht "GPUScaling"...
	     End if     'egal, ob erste 10 Buchstaben des Werts = "GPUScaling" oder nicht
	     Next
	     End if
	'End if
        Next
        End if 
    Next
End if


'--------------------------------------
'AMD

sKeyPath = "SYSTEM\CurrentControlSet\Control\Class"

oReg.EnumKey cHKLM, sKeyPath, aSubKeys
If Not (isnull(aSubKeys)) Then
    For Each sSubkey In aSubKeys
        sTmpKeyName1 = sKeyPath & "\" & sSubKey
        pReg.EnumKey cHKLM, sTmpKeyName1, bSubKeys
        If Not (isnull(bSubKeys)) Then
        For Each tSubkey In bSubKeys
	sTmpKeyName2 =  sTmpKeyName1 & "\" & tSubkey & "\DAL2_DATA__2_0"
	'msgbox sTmpKeyName2
        
        	qReg.EnumKey cHKLM, sTmpKeyName2, cSubKeys
        	If Not (isnull(cSubKeys)) Then
        	For Each uSubkey In cSubKeys
		'uSubkey = "DisplayPath_"		' plus eine unbekannte einstellige Zahl dahinter (z.B. "DisplayPath_8")
		if Left(uSubkey, 12) = "DisplayPath_" then
		sTmpKeyName3 =  sTmpKeyName2 & "\" & uSubkey & "\Option"
		'msgbox sTmpKeyName3

               	rReg.EnumValues cHKLM, sTmpKeyName3, dSubKeys
               	If Not (isnull(dSubKeys)) Then
               	For Each oSubkey In dSubKeys
	     	if oSubkey = "BestViewOption" then
	     	sTmpValueName =  sTmpKeyName3 & "\" & oSubkey	 
	     	'msgbox sTmpValueName 

	   	' Lese zuerst den Binary-String aus und schaue, ob �nderung notwendig ist
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName3
		BinaryPositionNumberToChange = 5
		BinaryReplacement = 1
	    	oReg.GetBinaryValue cHKLM, path, "BestViewOption", bArray
	   	 'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement Then
	    	ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung �ndern."
		Set filesys = CreateObject("Scripting.FileSystemObject")
		RegistryFolderToUnlock = CreateObject("Scripting.FileSystemObject").GetParentFolderName(sTmpValueName)
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & RegistryFolderToUnlock & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & RegistryFolderToUnlock & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
		'�ndere eine bestimmte Position der Binary-Strings
		'Changing value
		'msgbox "String lautet derzeit: " & bArray(BinaryPositionNumberToChange - 1) & " und wird im folgenden ersetzt durch " & BinaryReplacement 	'Mit dieser �nderung stimmt die Position exakt, w�rde sonst die erste 00 nicht mitz�hlen
		bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement
		'Write information back
	    	Return = 1
		Return = oReg.SetBinaryValue(cHKLM, path, "BestViewOption", bArray)
	    	If (Return = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Bin�rwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Pr�fung, ob Skalierung ge�ndert werden muss, abgeschlossen
	     	else
	     	End if
	    	Next
	     	End if
		End if	
        	Next
        	End if
        Next
        End if
    Next
End if


'--------------------------------------
'AMD

sKeyPath = "SYSTEM\CurrentControlSet\Control\Video"

oReg.EnumKey cHKLM, sKeyPath, aSubKeys
If Not (isnull(aSubKeys)) Then
    For Each sSubkey In aSubKeys
        sTmpKeyName1 = sKeyPath & "\" & sSubKey
        pReg.EnumKey cHKLM, sTmpKeyName1, bSubKeys
        If Not (isnull(bSubKeys)) Then
        For Each tSubkey In bSubKeys
	'if tSubkey = "0000" then
	sTmpKeyName2 =  sTmpKeyName1 & "\" & tSubkey
	'msgbox sTmpKeyName2

               qReg.EnumValues cHKLM, sTmpKeyName2, cSubKeys
               If Not (isnull(cSubKeys)) Then
               For Each oSubkey In cSubKeys
	     'oSubkey = "GPUScaling"		' plus eine unbekannte Zahl dahinter (z.B. "GPUScaling84")
	     if Left(oSubkey, 10) = "GPUScaling" then
	     sTmpValueName =  sTmpKeyName2 & "\" & oSubkey	 
	     'msgbox sTmpValueName 

	    ' Lese zuerst den Binary-String aus und schaue, ob �nderung notwendig ist
	    Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    path = sTmpKeyName2
	    oReg.GetBinaryValue cHKLM, path, oSubkey, bArray
	    'msgbox bArray(1 - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    If not bArray(1 - 1) = 1 Then
	    ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung �ndern."
		Set filesys = CreateObject("Scripting.FileSystemObject")
		RegistryFolderToUnlock = CreateObject("Scripting.FileSystemObject").GetParentFolderName(sTmpValueName)
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & RegistryFolderToUnlock & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & RegistryFolderToUnlock & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
	    values = Array(01,00,00,00)
	    Return = 1
	    Return = objRegistry.SetBinaryValue(cHKLM, path, oSubkey, values)
	    If (Return = 0) Then
    	    'msgbox "Binary value added successfully"
	    Else
    	    ' An error occurred
    	    'msgbox "Binary could not be added!"
	    End If	    'Meldung, ob Bin�rwert in Registry geschrieben wurde, wurde ausgegeben
	    End if	    ' Pr�fung, ob Skalierung ge�ndert werden muss, abgeschlossen
	     else	'Erste 10 Buchstaben des Werts sind nicht "GPUScaling"...
	     End if     'egal, ob erste 10 Buchstaben des Werts = "GPUScaling" oder nicht
	     Next
	     End if
	'End if
        Next
        End if 
    Next
End if


'--------------------------------------
'AMD

sKeyPath = "SYSTEM\CurrentControlSet\Control\Video"

oReg.EnumKey cHKLM, sKeyPath, aSubKeys
If Not (isnull(aSubKeys)) Then
    For Each sSubkey In aSubKeys
        sTmpKeyName1 = sKeyPath & "\" & sSubKey
        pReg.EnumKey cHKLM, sTmpKeyName1, bSubKeys
        If Not (isnull(bSubKeys)) Then
        For Each tSubkey In bSubKeys
	sTmpKeyName2 =  sTmpKeyName1 & "\" & tSubkey & "\DAL2_DATA__2_0"
	'msgbox sTmpKeyName2
        
        	qReg.EnumKey cHKLM, sTmpKeyName2, cSubKeys
        	If Not (isnull(cSubKeys)) Then
        	For Each uSubkey In cSubKeys
		'uSubkey = "DisplayPath_"		' plus eine unbekannte einstellige Zahl dahinter (z.B. "DisplayPath_8")
		if Left(uSubkey, 12) = "DisplayPath_" then
		sTmpKeyName3 =  sTmpKeyName2 & "\" & uSubkey & "\Option"
		'msgbox sTmpKeyName3

               	rReg.EnumValues cHKLM, sTmpKeyName3, dSubKeys
               	If Not (isnull(dSubKeys)) Then
               	For Each oSubkey In dSubKeys
	     	if oSubkey = "BestViewOption" then
	     	sTmpValueName =  sTmpKeyName3 & "\" & oSubkey	 
	     	'msgbox sTmpValueName 

	   	' Lese zuerst den Binary-String aus und schaue, ob �nderung notwendig ist
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName3
		BinaryPositionNumberToChange = 5
		BinaryReplacement = 1
	    	oReg.GetBinaryValue cHKLM, path, "BestViewOption", bArray
	   	 'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement Then
	    	ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung �ndern."
		Set filesys = CreateObject("Scripting.FileSystemObject")
		RegistryFolderToUnlock = CreateObject("Scripting.FileSystemObject").GetParentFolderName(sTmpValueName)
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & RegistryFolderToUnlock & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & RegistryFolderToUnlock & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
		'�ndere eine bestimmte Position der Binary-Strings
		'Changing value
		'msgbox "String lautet derzeit: " & bArray(BinaryPositionNumberToChange - 1) & " und wird im folgenden ersetzt durch " & BinaryReplacement 	'Mit dieser �nderung stimmt die Position exakt, w�rde sonst die erste 00 nicht mitz�hlen
		bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement
		'Write information back
	    	Return = 1
		Return = oReg.SetBinaryValue(cHKLM, path, "BestViewOption", bArray)
	    	If (Return = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Bin�rwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Pr�fung, ob Skalierung ge�ndert werden muss, abgeschlossen
	     	else
	     	End if
	    	Next
	     	End if
		End if	
        	Next
        	End if

        Next
        End if
    Next
End if


'--------------------------------------
' NVIDIA

sKeyPath = "SYSTEM\CurrentControlSet\services\nvlddmkm\DisplayDatabase"

oReg.EnumKey cHKLM, sKeyPath, aSubKeys
If Not (isnull(aSubKeys)) Then
    For Each sSubkey In aSubKeys
        sTmpKeyName1 = sKeyPath & "\" & sSubKey

               	rReg.EnumValues cHKLM, sTmpKeyName1, dSubKeys
               	If Not (isnull(dSubKeys)) Then
               	For Each oSubkey In dSubKeys
	     	if oSubkey = "ScalingConfig" then
	     	sTmpValueName =  sTmpKeyName1 & "\" & oSubkey	 
	     	'msgbox sTmpValueName


	   	' Lese zuerst den Binary-String aus und schaue, ob �nderung notwendig ist
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName1
		BinaryPositionNumberToChange = 9
		BinaryReplacement = 9		'09 in hexadezimal ist das Ziel! (Angabe muss erfolgen: ohne Anf�hrungsstriche = dezimal!)
	    	oReg.GetDwordValue cHKLM, path, "ScalingConfig", bArray	' Neue Methode von Nvidia (Dword-Wert: 0=Anzeige, 1=GPU)
		if bArray = 1 then
		NvidiaTest_1 = "NeueMethodeDwordKorrekt"
		End if
		if bArray = 0 then	'Falls "ScalingConfig" ein DWORD-Wert ist und den Wert 0 hat
		Set filesys = CreateObject("Scripting.FileSystemObject")
		RegistryFolderToUnlock = CreateObject("Scripting.FileSystemObject").GetParentFolderName(sTmpValueName)
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & RegistryFolderToUnlock & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & RegistryFolderToUnlock & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
		Return = oReg.SetDwordValue(cHKLM, path, "ScalingConfig", 1)
			if errorlevel = 0 then
	    		ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung �ndern."
			NvidiaTest_2 = "Aus"
			End If
		Else	'schaue stattdessen nach Binary-Eintrag
	    	oReg.GetBinaryValue cHKLM, path, "ScalingConfig", bArray
	   	 'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement AND not NvidiaTest_1 = "NeueMethodeDwordKorrekt" AND not NvidiaTest_2 = "Aus" Then
		Set filesys = CreateObject("Scripting.FileSystemObject")
		RegistryFolderToUnlock = CreateObject("Scripting.FileSystemObject").GetParentFolderName(sTmpValueName)
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & RegistryFolderToUnlock & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & RegistryFolderToUnlock & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
		'�ndere eine bestimmte Position der Binary-Strings
		'Changing value
		'msgbox "String lautet derzeit: " & bArray(BinaryPositionNumberToChange - 1) & " und wird im folgenden ersetzt durch " & BinaryReplacement 	'Mit dieser �nderung stimmt die Position exakt, w�rde sonst die erste db nicht mitz�hlen
		bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement
		'Write information back
	    	Return = 1
		Return = oReg.SetBinaryValue(cHKLM, path, "ScalingConfig", bArray)
	    	If (Return = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung �ndern."
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Bin�rwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Pr�fung, ob Skalierung ge�ndert werden muss, abgeschlossen


	   	' Lese zuerst den Binary-String aus und schaue, ob �nderung notwendig ist
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName1
		BinaryPositionNumberToChange = 13
		BinaryReplacement = 245		'f5 in hexadezimal ist das Ziel! (Angabe muss erfolgen: ohne Anf�hrungsstriche = dezimal!)
	    	oReg.GetBinaryValue cHKLM, path, "ScalingConfig", bArray
	   	 'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement AND not NvidiaTest_1 = "NeueMethodeDwordKorrekt" AND not NvidiaTest_2 = "Aus" Then
		'�ndere eine bestimmte Position der Binary-Strings
		'Changing value
		'msgbox "String lautet derzeit: " & bArray(BinaryPositionNumberToChange - 1) & " und wird im folgenden ersetzt durch " & BinaryReplacement 	'Mit dieser �nderung stimmt die Position exakt, w�rde sonst die erste db nicht mitz�hlen
		bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement
		'Write information back
	    	Return = 1
		Return = oReg.SetBinaryValue(cHKLM, path, "ScalingConfig", bArray)
	    	If (Return = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung �ndern."
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Bin�rwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Pr�fung, ob Skalierung ge�ndert werden muss, abgeschlossen


	   	' Lese zuerst den Binary-String aus und schaue, ob �nderung notwendig ist
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName1
		BinaryPositionNumberToChange = 14
		BinaryReplacement = 0		'00 in hexadezimal ist das Ziel! (Angabe muss erfolgen: ohne Anf�hrungsstriche = dezimal!)
	    	oReg.GetBinaryValue cHKLM, path, "ScalingConfig", bArray
	   	 'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement AND not NvidiaTest_1 = "NeueMethodeDwordKorrekt" AND not NvidiaTest_2 = "Aus" Then
		'�ndere eine bestimmte Position der Binary-Strings
		'Changing value
		'msgbox "String lautet derzeit: " & bArray(BinaryPositionNumberToChange - 1) & " und wird im folgenden ersetzt durch " & BinaryReplacement 	'Mit dieser �nderung stimmt die Position exakt, w�rde sonst die erste db nicht mitz�hlen
		bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement
		'Write information back
	    	Return = 1
		Return = oReg.SetBinaryValue(cHKLM, path, "ScalingConfig", bArray)
	    	If (Return = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung �ndern."
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Bin�rwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Pr�fung, ob Skalierung ge�ndert werden muss, abgeschlossen


		End if	    ' Pr�fung, ob NVIDIA Dword oder Binary Eintrag benutzt, abgeschlossen

	     	else
	     	End if

        Next
        End if
    Next
End if

'--------------------------------------
' NVIDIA

sKeyPath = "SYSTEM\CurrentControlSet\services\nvlddmkm\State\DisplayDatabase"

oReg.EnumKey cHKLM, sKeyPath, aSubKeys
If Not (isnull(aSubKeys)) Then
    For Each sSubkey In aSubKeys
        sTmpKeyName1 = sKeyPath & "\" & sSubKey

               	rReg.EnumValues cHKLM, sTmpKeyName1, dSubKeys
               	If Not (isnull(dSubKeys)) Then
               	For Each oSubkey In dSubKeys
	     	if oSubkey = "ScalingConfig" then
	     	sTmpValueName =  sTmpKeyName1 & "\" & oSubkey	 
	     	'msgbox sTmpValueName 


	   	' Lese zuerst den Binary-String aus und schaue, ob �nderung notwendig ist
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName1
		BinaryPositionNumberToChange = 9
		BinaryReplacement = 137		'89 in hexadezimal ist das Ziel! (Angabe muss erfolgen: ohne Anf�hrungsstriche = dezimal!)
	    	oReg.GetDwordValue cHKLM, path, "ScalingConfig", bArray	' Neue Methode von Nvidia (Dword-Wert: 0=Anzeige, 1=GPU)
		if bArray = 1 then
		NvidiaTest_1 = "NeueMethodeDwordKorrekt"
		End if
		if bArray = 0 then	'Falls "ScalingConfig" ein DWORD-Wert ist und den Wert 0 hat
		Set filesys = CreateObject("Scripting.FileSystemObject")
		RegistryFolderToUnlock = CreateObject("Scripting.FileSystemObject").GetParentFolderName(sTmpValueName)
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & RegistryFolderToUnlock & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & RegistryFolderToUnlock & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
		Return = oReg.SetDwordValue(cHKLM, path, "ScalingConfig", 1)
			if errorlevel = 0 then
	    		ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung �ndern."
			NvidiaTest_2 = "Aus"
			End If
		Else	'schaue stattdessen nach Binary-Eintrag
	    	oReg.GetBinaryValue cHKLM, path, "ScalingConfig", bArray
	   	 'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement AND not NvidiaTest_1 = "NeueMethodeDwordKorrekt" AND not NvidiaTest_2 = "Aus" Then
		Set filesys = CreateObject("Scripting.FileSystemObject")
		RegistryFolderToUnlock = CreateObject("Scripting.FileSystemObject").GetParentFolderName(sTmpValueName)
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & RegistryFolderToUnlock & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & RegistryFolderToUnlock & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
		'�ndere eine bestimmte Position der Binary-Strings
		'Changing value
		'msgbox "String lautet derzeit: " & bArray(BinaryPositionNumberToChange - 1) & " und wird im folgenden ersetzt durch " & BinaryReplacement	'Mit dieser �nderung stimmt die Position exakt, w�rde sonst die erste db nicht mitz�hlen
		bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement
		'Write information back
	    	Return = 1
		Return = oReg.SetBinaryValue(cHKLM, path, "ScalingConfig", bArray)
	    	If (Return = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung �ndern."
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Bin�rwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Pr�fung, ob Skalierung ge�ndert werden muss, abgeschlossen


	   	' Lese zuerst den Binary-String aus und schaue, ob �nderung notwendig ist
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName1
		BinaryPositionNumberToChange = 13
		BinaryReplacement = 117		'75 in hexadezimal ist das Ziel! (Angabe muss erfolgen: ohne Anf�hrungsstriche = dezimal!)
	    	oReg.GetBinaryValue cHKLM, path, "ScalingConfig", bArray
	   	'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement AND not NvidiaTest_1 = "NeueMethodeDwordKorrekt" AND not NvidiaTest_2 = "Aus" Then
		'�ndere eine bestimmte Position der Binary-Strings
		'Changing value
		'msgbox "String lautet derzeit: " & bArray(BinaryPositionNumberToChange - 1) & " und wird im folgenden ersetzt durch " & BinaryReplacement 	'Mit dieser �nderung stimmt die Position exakt, w�rde sonst die erste db nicht mitz�hlen
		bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement
		'Write information back
	    	Return = 1
		Return = oReg.SetBinaryValue(cHKLM, path, "ScalingConfig", bArray)
	    	If (Return = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung �ndern."
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Bin�rwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Pr�fung, ob Skalierung ge�ndert werden muss, abgeschlossen


	   	' Lese zuerst den Binary-String aus und schaue, ob �nderung notwendig ist
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName1
		BinaryPositionNumberToChange = 14
		BinaryReplacement = 0		'00 in hexadezimal ist das Ziel! (Angabe muss erfolgen: ohne Anf�hrungsstriche = dezimal!)
	    	oReg.GetBinaryValue cHKLM, path, "ScalingConfig", bArray
	   	'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement AND not NvidiaTest_1 = "NeueMethodeDwordKorrekt" AND not NvidiaTest_2 = "Aus" Then
		'�ndere eine bestimmte Position der Binary-Strings
		'Changing value
		'msgbox "String lautet derzeit: " & bArray(BinaryPositionNumberToChange - 1) & " und wird im folgenden ersetzt durch " & BinaryReplacement 	'Mit dieser �nderung stimmt die Position exakt, w�rde sonst die erste db nicht mitz�hlen
		bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement
		'Write information back
	    	Return = 1
		Return = oReg.SetBinaryValue(cHKLM, path, "ScalingConfig", bArray)
	    	If (Return = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung �ndern."
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Bin�rwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Pr�fung, ob Skalierung ge�ndert werden muss, abgeschlossen


		End if	    ' Pr�fung, ob NVIDIA Dword oder Binary Eintrag benutzt, abgeschlossen

	     	else
	     	End if

        Next
        End if
    Next
End if

'--------------------------------------
' Intel

sKeyPath = "SYSTEM\CurrentControlSet\Services\ialm"

oReg.EnumKey cHKLM, sKeyPath, aSubKeys
If Not (isnull(aSubKeys)) Then
    For Each sSubkey In aSubKeys
        if Left(sSubkey, 6) = "Device" then
        sTmpKeyName1 = sKeyPath & "\" & sSubKey
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpKeyName1 & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpKeyName1 & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
               	rReg.EnumValues cHKLM, sTmpKeyName1, dSubKeys
               	If Not (isnull(dSubKeys)) Then
               	For Each oSubkey In dSubKeys
	     	if oSubkey = "PreferredDisplayScaling" then

		' HIERBEI KEIN VERMERK ZUP RELOAD DER GPU MACHEN!!!
		' Weil Intel die Werte immer neu erstellt und nur neu l�dt, falls sie gel�scht werden, w�rde dies sonst dazu f�hren, dass die GPU jedes Mal neu gestartet wird - ganz gleich ob die GPU-Einstellung schon vorher richtig war oder nicht! Also l�sche ich diese Registry-Werte einfach ohne GPU-Reload, der Indikator f�r einen notwendigen GPU-Reload liegt f�r Intel nur unter "SYSTEM\CurrentControlSet\Control\GraphicsDrivers\Configuration"! Das reicht aus.
	     	sTmpValueName =  sTmpKeyName1 & "\" & oSubkey	 
	     	'msgbox sTmpValueName 
	    	path = sTmpKeyName1
	    	Return = 1
		Return = oReg.DeleteValue(cHKLM, path, oSubkey)
	    	If (Return = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Bin�rwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Pr�fung, ob Skalierung ge�ndert werden muss, abgeschlossen

	     	Next
	     	End if

        End if
    Next
End if

sKeyPath = "SYSTEM\CurrentControlSet\Control\GraphicsDrivers\Configuration"

oReg.EnumKey cHKLM, sKeyPath, aSubKeys
If Not (isnull(aSubKeys)) Then
    For Each sSubkey In aSubKeys
        sTmpKeyName1 = sKeyPath & "\" & sSubKey
        pReg.EnumKey cHKLM, sTmpKeyName1, bSubKeys
        If Not (isnull(bSubKeys)) Then
        For Each tSubkey In bSubKeys
	if tSubkey = "00" then
	sTmpKeyName2 =  sTmpKeyName1 & "\" & tSubkey & "\00"
	'msgbox sTmpKeyName2
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpKeyName2 & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpKeyName2 & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
               	rReg.EnumValues cHKLM, sTmpKeyName2, cSubKeys
               	If Not (isnull(cSubKeys)) Then
               	For Each oSubkey In cSubKeys

		' HIERBEI KEIN VERMERK ZUP RELOAD DER GPU MACHEN!!!
		' Weil Intel die Werte immer neu erstellt und nur neu l�dt, falls sie gel�scht werden, w�rde dies sonst dazu f�hren, dass die GPU jedes Mal neu gestartet wird - ganz gleich ob die GPU-Einstellung schon vorher richtig war oder nicht! Also l�sche ich diese Registry-Werte einfach ohne GPU-Reload, der Indikator f�r einen notwendigen GPU-Reload liegt f�r Intel nur unter "SYSTEM\CurrentControlSet\Control\GraphicsDrivers\Configuration"! Das reicht aus.

	     	if oSubkey = "DwmClipBox.bottom" then
	     	sTmpValueName =  sTmpKeyName2 & "\" & oSubkey	 
	     	'msgbox sTmpValueName     
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName2
	    	Return = 1
		Return = oReg.DeleteValue(cHKLM, path, oSubkey)
	    	If (Return = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Bin�rwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Pr�fung, ob Skalierung ge�ndert werden muss, abgeschlossen

	     	if oSubkey = "DwmClipBox.left" then
	     	sTmpValueName =  sTmpKeyName2 & "\" & oSubkey	 
	     	'msgbox sTmpValueName     
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName2
	    	Return = 1
		Return = oReg.DeleteValue(cHKLM, path, oSubkey)
	    	If (Return = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Bin�rwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Pr�fung, ob Skalierung ge�ndert werden muss, abgeschlossen

	     	if oSubkey = "DwmClipBox.right" then
	     	sTmpValueName =  sTmpKeyName2 & "\" & oSubkey	 
	     	'msgbox sTmpValueName     
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName2
	    	Return = 1
		Return = oReg.DeleteValue(cHKLM, path, oSubkey)
	    	If (Return = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Bin�rwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Pr�fung, ob Skalierung ge�ndert werden muss, abgeschlossen

	     	if oSubkey = "DwmClipBox.top" then
	     	sTmpValueName =  sTmpKeyName2 & "\" & oSubkey	 
	     	'msgbox sTmpValueName     
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName2
	    	Return = 1
		Return = oReg.DeleteValue(cHKLM, path, oSubkey)
	    	If (Return = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Bin�rwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Pr�fung, ob Skalierung ge�ndert werden muss, abgeschlossen

        Next
        End if
        End if
        Next
        End if
    Next
End if


'--------------------------------------
'ALLE HERSTELLER ### UNTERER TEIL MIT AMD FEHLT NOCH!!!

sKeyPath = "SYSTEM\CurrentControlSet\Control\GraphicsDrivers\Configuration"

oReg.EnumKey cHKLM, sKeyPath, aSubKeys
If Not (isnull(aSubKeys)) Then
    For Each sSubkey In aSubKeys
        sTmpKeyName1 = sKeyPath & "\" & sSubKey
        pReg.EnumKey cHKLM, sTmpKeyName1, bSubKeys
        If Not (isnull(bSubKeys)) Then
        For Each tSubkey In bSubKeys
	if tSubkey = "00" then
	sTmpKeyName2 =  sTmpKeyName1 & "\" & tSubkey & "\00"
	'msgbox sTmpKeyName2
        
               	rReg.EnumValues cHKLM, sTmpKeyName2, cSubKeys
               	If Not (isnull(cSubKeys)) Then
               	For Each oSubkey In cSubKeys
	     	if oSubkey = "Scaling" then
	     	sTmpValueName =  sTmpKeyName2 & "\" & oSubkey	 
	     	'msgbox sTmpValueName     

	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName2
	    	values = "2"
	    	oReg.GetDwordValue cHKLM, path, oSubkey, bArray
	   	'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray = 2 Then
	    	ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung �ndern."
		Set filesys = CreateObject("Scripting.FileSystemObject")
		RegistryFolderToUnlock = CreateObject("Scripting.FileSystemObject").GetParentFolderName(sTmpValueName)
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & RegistryFolderToUnlock & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & RegistryFolderToUnlock & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
	    	Return = 1
	    	Return = objRegistry.SetDwordValue(cHKLM, path, oSubkey, values)
	    	If (Return = 0) Then
    	    		'msgbox "Binary value added successfully"

		logpath = "Software\ATI\ACE\Settings\Graphics\DeviceDFP"
               	sReg.EnumValues cHKCU, logpath, eSubKeys
               	If Not (isnull(eSubKeys)) Then
               	For Each pSubkey In eSubKeys
		if Left(pSubkey, 17) = "CurrentImageScale" then		' (z.B: "CurrentImageScale,6,4")
		'msgbox pSubkey
		logvalues = ""		' keine Ahnung mehr
	    	Log = objRegistry.SetStringValue(cHKCU, logpath, pSubkey, logvalues)
		End if
		Next
		End if

		logpath = "Software\AMD\CN\DisplayScaling"
               	sReg.EnumValues cHKCU, logpath, eSubKeys
               	If Not (isnull(eSubKeys)) Then
               	For Each pSubkey In eSubKeys
		if Left(pSubkey, 17) = "CurrentImageScale" then		' (z.B: "CurrentImageScale,8,4")
		'msgbox pSubkey
		logvalues = "2"
	    	Log = objRegistry.SetDwordValue(cHKCU, logpath, pSubkey, logvalues)
		End if
		Next
		End if

	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Bin�rwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Pr�fung, ob Skalierung ge�ndert werden muss, abgeschlossen

	     	else
	     	End if
	    	Next
	     	End if
		End if  	
        	Next
        	End if
    Next
End if



' ------------------------------------------------------------------------------------------
' AKTUALISIERE GRAFIKKARTE
' ------------------------------------------------------------------------------------------

If not ScalingProzedurNotwendig = "Ja" AND not ScalingProzedurNotwendig = "ja" Then
Else	' Aktivieren der GPU-Skalierung ist noch ausstehend


' --- Sichere Desktop-Icon-Position mit zus�tzlichem Programm

  FullPath = TEMP & "\DesktopOK\Profiles"
  Set fso = CreateObject("Scripting.FileSystemObject")
  arr = split(FullPath, "\")
  path = ""
  For Each dir In arr
    If path <> "" Then path = path & "\"
    path = path & dir
    If fso.FolderExists(path) = False Then fso.CreateFolder(path)
  Next

Set objShell = CreateObject("WScript.Shell")
objShell.Run Chr(34) & RIGHTHERE & "DesktopOK\DesktopOK.exe" & Chr(34) & " /save /silent " & Chr(34) & TEMP & "\DesktopOK\Profiles\DesktopOKProfile.dok" & Chr(34), 0, True


' --- Sichere Windows-Fensterpositionen
Set objShell = CreateObject("WScript.Shell")
objShell.Run Chr(34) & RIGHTHERE & "Windowstate\windowstate.exe" & Chr(34) & " --file " & Chr(34) & TEMP & "\Windowstate\windowstate.json" & Chr(34) & " save", 0, True


' --- AMD: Ggf. Laufende Instanz f�r sp�ter merken
Counter = 0
strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = 'RadeonSoftware.exe'")
For Each objProcess in colProcess
Counter = + 1
objprocess.ExecutablePath = LCase(objprocess.ExecutablePath)
objProcess.Name = LCase(objProcess.Name)
search="\" & objProcess.Name
AMDPath=left(objprocess.ExecutablePath, instr(objprocess.ExecutablePath, search)-0)
Next
If Counter > 0 Then
Set fso = CreateObject("Scripting.FileSystemObject")
If not (fso.FileExists(AMDPath & "cncmd.exe")) Then
Else
AMDLaunchTask = "Ja"
End if
End if


' --- NVIDIA: Ggf. Laufende Instanz f�r sp�ter merken
isServiceInstalled = nothing
ProcessPfad = "nvidia"
SucheServiceName = "NVDisplay.ContainerLocalSystem"
Counter = 0
strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = 'nvcplui.exe'")
For Each objProcess in colProcess
Counter = + 1
objprocess.ExecutablePath = LCase(objprocess.ExecutablePath)
objProcess.Name = LCase(objProcess.Name)
search="\" & objProcess.Name
NVIDIAPath=left(objprocess.ExecutablePath, instr(objprocess.ExecutablePath, search)-0)
    ' Beende NVIDIA, da ansonsten das Programmoberfl�che die �nderungen nicht anzeigt
    ' msgbox "Prozesspfad: " & objProcess.ExecutablePath
    nPos = InStr(objProcess.ExecutablePath, ProcessPfad)  'binary comparison
    if nPos="0" Then
    ' msgbox "Debugginginfo: String nicht gefunden, nicht l�schen..."
    Else
    ' msgbox "Debugginginfo: String ist Teil des ausgef�hrten Pfads - Prozess wird gekillt..."
    objProcess.Terminate
    End if
Next
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
    ' Obtain an instance of the the class 
    ' using a key property value.
    isServiceInstalled = FALSE
    svcQry = "SELECT * from Win32_Service"
    Set objOutParams = objWMIService.ExecQuery(svcQry)
    For Each objSvc in objOutParams
        Select Case objSvc.Name
            Case SucheServiceName
                isServiceInstalled = TRUE
        End Select
    Next
If Counter > 0 AND isServiceInstalled Then
NVIDIALaunchTask = "Ja"
End if


' --- Kille StikyNot.exe, wenn Pfad und Prozessname stimmt (und Kurznotizen �berhaupt laufen)
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcess = objWMIService.ExecQuery("Select * From Win32_Process Where Name = 'StikyNot.exe'")
If colProcess.Count > 0 Then
' msgbox "L�uft..."
StikyNotRestart = "Ja"
Else
StikyNotRestart = "Nein"
End if
If StikyNotRestart = "Ja" Then
' Hilfe -Processname- : Dateiname+Endung, z.B. "Meine App.exe" - Leer lassen, um Prozesskillen ganz zu deaktivieren
' Hilfe -Apptitle- : Fenster�berschrift (komplett oder Teil davon), z.B. "Brennvorgang Free 0% abges" [...]
' Hilfe -ProcessPfad- : Ordnerpfadpfad (komplett oder Teil davon), z.B. "AppData\Local\Temp\"
ProcessName = "StikyNot.exe"
'Apptitle = ""
ProcessPfad = SYSTEMDRIVE & "\Windows\system32\"
ProcessPfad2 = "Classic Sticky Notes\"
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
  ' Connect to Windows Management Instrumentation (WMI) object using a moniker
  Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
  ' Execute a query to get all running instances of specified application
  Set colProcess = objWMIService.ExecQuery("Select * From Win32_Process Where Name = '" & strApplication & "'")
  ' Terminate all running instances of the application
 Set WshShell = WScript.CreateObject ("WScript.Shell")
call KillStickyNotes(ProcessName)
Function KillStickyNotes(ProcessName) 'Hier ist ProcessName keine Variable, sondern ein feststehender Begriff!
    Dim objWMIService, colProcess
    Dim strComputer, strList, p, q
    Dim i :i= 0 
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
    Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = '" & ProcessName & "'")
    For Each q in colProcess
    ' msgbox "Prozesspfad: " & q.ExecutablePath
    nPos = InStr(LCase(q.ExecutablePath), LCase(ProcessPfad)) OR InStr(LCase(q.ExecutablePath), LCase(ProcessPfad2))  'binary comparison
    if nPos="0" Then
    ' msgbox "Debugginginfo: String nicht gefunden, nicht l�schen..."
    Else
    ' msgbox "Debugginginfo: String ist Teil des ausgef�hrten Pfads - Prozess wird gekillt..."
    if InStr(LCase(q.ExecutablePath), LCase(ProcessPfad)) = 0 then
    else
        BuildInNotes = "Ja"
    end if
    if InStr(LCase(q.ExecutablePath), LCase(ProcessPfad2)) = 0 then
    else
        ClassicNotes = "Ja"
    end if
    q.Terminate
    End if
    Next
    i = i+1        
End Function
End if


' --- Pausiere aktive VMware-VMs, um Abst�rze zu verhindern
' delete earlier txt file, if not happened yet
Set fs = CreateObject("Scripting.Filesystemobject")
fs.DeleteFile(TEMP & "\vmrun_temp\VMwareVMs2Resume.txt"), True
' remove empty folder structure through a trick with robocopy
Set objShell = CreateObject("WScript.Shell")
objShell.Run Chr(34) & SYSTEMROOT & "\system32\robocopy.exe" & Chr(34) & " " & Chr(34) & TEMP & "\vmrun_temp" & Chr(34) & " " & Chr(34) & TEMP & "\vmrun_temp" & Chr(34) & " /s /move", 0, True
' build new folder structure
oFolder = TEMP & "\vmrun_temp\test"
BuildFullPath oFolder
Sub BuildFullPath(ByVal oFolder)
  If Not objFSO.FolderExists(oFolder) Then
    BuildFullPath objFSO.GetParentFolderName(oFolder)
    objFSO.CreateFolder oFolder
  End If
End Sub
' export active VMs list in a file
path = TEMP & "\vmrun_temp\VMwareVMs2Resume.txt"		'Ausgabedatei
Set colItems = CreateObject("Scripting.FileSystemObject")
Set objFile = colItems.CreateTextFile(path, True)
Set objShell = CreateObject("WScript.Shell")
Set key = objShell.Exec(RIGHTHERE & "vmrun\vmrun.exe list")			'Befehl, dessen Antwort abgefangen wird
Set oShell = key.StdOut
While Not oShell.AtEndOfStream
   strFirstLine = oShell.ReadLine
   If InStr(strFirstLine,"vmx") Then				'Inhaltsfilter zum Schreiben in Datei
       'msgbox strFirstLine
       objFile.WriteLine(strFirstLine)
' suspend active VM
       Set objShell = CreateObject("WScript.Shell")
       objShell.Run Chr(34) & RIGHTHERE & "vmrun\vmrun.exe" & Chr(34) & " suspend "  & Chr(34) & strFirstLine & Chr(34), 0, True
   End If
Wend
objFile.Close


' --- Starte Grafikkarte neu
Set objShell = CreateObject("WScript.Shell")
objShell.Run Chr(34) & "restart-only.exe" & Chr(34) & "", 0, False


' ! Wartezeit, sehr wichtig
Sleep "6000"	'Sleep-Modus exklusiv f�r "PScript.exe" (= Portable VBS)
WScript.Sleep "6000"	'Sleep-Modus f�r windowseigenen Script-Host "wscript.exe" und "cscript.exe"


' --- Kille Fehlermeldungen (WerFault.exe), wenn Pfad und Prozessname stimmt
' Hilfe -Processname- : Dateiname+Endung, z.B. "Meine App.exe" - Leer lassen, um Prozesskillen ganz zu deaktivieren
' Hilfe -Apptitle- : Fenster�berschrift (komplett oder Teil davon), z.B. "Brennvorgang Free 0% abges" [...]
' Hilfe -ProcessPfad- : Ordnerpfadpfad (komplett oder Teil davon), z.B. "AppData\Local\Temp\"
ProcessName = "WerFault.exe"
'Apptitle = ""
ProcessPfad = SYSTEMDRIVE & "\Windows\system32\"
ProcessPfad2 = SYSTEMDRIVE & "\Windows\SysWOW64\"
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
  ' Connect to Windows Management Instrumentation (WMI) object using a moniker
  Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
  ' Execute a query to get all running instances of specified application
  Set colProcess = objWMIService.ExecQuery("Select * From Win32_Process Where Name = '" & strApplication & "'")
  ' Terminate all running instances of the application
 Set WshShell = WScript.CreateObject ("WScript.Shell")
call KillAll(ProcessName)
Function KillAll(ProcessName) 'Hier ist ProcessName keine Variable, sondern ein feststehender Begriff!
    Dim objWMIService, colProcess
    Dim strComputer, strList, p, q
    Dim i :i= 0 
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
    Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name like '" & ProcessName & "'")
    For Each q in colProcess
    ' msgbox "Prozesspfad: " & q.ExecutablePath
    nPos = InStr(q.ExecutablePath, ProcessPfad) OR InStr(q.ExecutablePath, ProcessPfad2)  'binary comparison
    if nPos="0" Then
    ' msgbox "Debugginginfo: String nicht gefunden, nicht l�schen..."
    Else
    ' msgbox "Debugginginfo: String ist Teil des ausgef�hrten Pfads - Prozess wird gekillt..."
    q.Terminate
    End if
    Next
    i = i+1        
End Function


' --- Starte vimage.exe erneut, wenn aktive Instanz in Pfad und Prozessname vorhanden
' Hilfe -Processname- : Dateiname+Endung, z.B. "Meine App.exe" - Leer lassen, um Prozesskillen ganz zu deaktivieren
' Hilfe -Apptitle- : Fenster�berschrift (komplett oder Teil davon), z.B. "Brennvorgang Free 0% abges" [...]
' Hilfe -ProcessPfad- : Ordnerpfadpfad (komplett oder Teil davon), z.B. "AppData\Local\Temp\"
ProcessName = "vimage.exe"
'Apptitle = ""
ProcessPfad = "\VImage\"
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
  ' Connect to Windows Management Instrumentation (WMI) object using a moniker
  Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
  ' Execute a query to get all running instances of specified application
  Set colProcess = objWMIService.ExecQuery("Select * From Win32_Process Where Name = '" & strApplication & "'")
 Set WshShell = WScript.CreateObject ("WScript.Shell")
call VimageTest(ProcessName)
Function VimageTest(ProcessName) 'Hier ist ProcessName keine Variable, sondern ein feststehender Begriff!
    Dim objWMIService, colProcess
    Dim strComputer, strList, p, q
    Dim i :i= 0 
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
    Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = '" & ProcessName & "'")
    For Each q in colProcess
    ' msgbox "Prozesspfad: " & q.ExecutablePath
    nPos = InStr(q.ExecutablePath, ProcessPfad) 'binary comparison
    if nPos="0" Then
    ' msgbox "Debugginginfo: String nicht gefunden, nicht l�schen..."
    Else
    ' msgbox "Debugginginfo: String ist Teil des ausgef�hrten Pfads - Prozess wird gekillt..."
    'q.Terminate
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run Chr(34) & RIGHTHERE & "VImage\vimage.exe" & Chr(34) & " " & Chr(34) & RIGHTHERE & "VImage\loading-waiting.gif" & Chr(34) & " -clickThrough -alwaysOnTop -taskbarIcon", 1, False
    End if
    Next
    i = i+1        
End Function


' --- Starte Kurznotizen neu
If StikyNotRestart = "Ja" Then
  Set fso = CreateObject("Scripting.FileSystemObject")
  If BuildInNotes = "Ja" Then
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run Chr(34) & RIGHTHERE & "StartStikyNot\StartStikyNot.exe" & Chr(34), 0, True
  Else
    ' File not exists or Sticky Notes wasn't open before
  End If
  Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FileExists(SYSTEMDRIVE & "\Program Files\Classic Sticky Notes\StikyNot.exe") AND ClassicNotes = "Ja" Then
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run Chr(34) & SYSTEMDRIVE & "\Program Files\Classic Sticky Notes\StikyNot.exe" & Chr(34), 1, False
  Else
    ' File not exists or Sticky Notes wasn't open before
  End If
End if


' --- Wiederherstellen der Windows-Fensterpositionen
Set objShell = CreateObject("WScript.Shell")
objShell.Run Chr(34) & RIGHTHERE & "Windowstate\Windowstate_Restore.exe" & Chr(34) & " /VERYSILENT /NORESTART /SUPPRESSMSGBOXES", 0, True


' --- Wiederherstellen der Desktop-Icon-Position mit zus�tzlichem Programm
Set fso = CreateObject("Scripting.FileSystemObject")
If (fso.FileExists(TEMP & "\DesktopOK\Profiles\DesktopOKProfile.dok")) Then
Set objShell = CreateObject("WScript.Shell")
objShell.Run Chr(34) & RIGHTHERE & "DesktopOK\DesktopOK.exe" & Chr(34) & " /load /silent " & Chr(34) & TEMP & "\DesktopOK\Profiles\DesktopOKProfile.dok" & Chr(34), 0, True
Else
' msgbox "Datei fehlt - Bei der Sicherung zuvor ist etwas schiefgelaufen."
End if


' --- Starte AMD-Programm im Hintergrund, falls vorher gelaufen ist und sich nun abgeschaltet hat
if AMDLaunchTask = "Ja" then
' Verwalte RadeonSoftware.exe im Hintegergrund - daf�r braucht es cncmd.exe
Set objShell = CreateObject("WScript.Shell")
objShell.Run Chr(34) & AMDPath & "cncmd.exe" & Chr(34) & " restart", 1, False
' other params for cncmd.exe: exit launcher, hide launcher
end if


' --- Starte NVIDIA-Programm, falls vorher gelaufen ist und sich nun abgeschaltet hat
if NVIDIALaunchTask = "Ja" then
Set objShell = CreateObject("WScript.Shell")
objShell.Run Chr(34) & NVIDIAPath & "nvcplui.exe" & Chr(34) & "", 1, False
end if


' --- Starte NVIDIA-Service erneut, um ggf. abgeschaltetes Hintergrundprogramm neu zu starten
Set objShell = CreateObject("WScript.Shell")
'objShell.Run Chr(34) & SYSTEMDRIVE & "\Windows\system32\net.exe" & Chr(34) & " stop NVDisplay.ContainerLocalSystem", 0, True
'objShell.Run Chr(34) & SYSTEMDRIVE & "\Windows\system32\net.exe" & Chr(34) & " start NVDisplay.ContainerLocalSystem", 0, True


' --- Stelle gespeicherte pausierte VMware-VMs wieder her
' restart remembered VMs
path = TEMP & "\vmrun_temp\VMwareVMs2Resume.txt"		'Ausgabedatei
Set objfso = CreateObject("Scripting.FileSystemObject")
Set objFile = objfso.OpenTextFile(path)
Do Until objFile.AtEndOfStream
  'msgbox objFile.ReadLine
       Set objShell = CreateObject("WScript.Shell")
       objShell.Run Chr(34) & RIGHTHERE & "vmrun\vmrun.exe" & Chr(34) & " start "  & Chr(34) & objFile.ReadLine & Chr(34), 0, True
Loop
objFile.Close
' delete txt file
Set fs = CreateObject("Scripting.Filesystemobject")
fs.DeleteFile(TEMP & "\vmrun_temp\VMwareVMs2Resume.txt"), True
' remove empty folder structure through a trick with robocopy
Set objShell = CreateObject("WScript.Shell")
objShell.Run Chr(34) & SYSTEMROOT & "\system32\robocopy.exe" & Chr(34) & " " & Chr(34) & TEMP & "\vmrun_temp" & Chr(34) & " " & Chr(34) & TEMP & "\vmrun_temp" & Chr(34) & " /s /move", 0, True


End if	' �berpr�fen der ausstehenden GPU-Skalierung abgeschlossen



' --- Kille vimage.exe, wenn Pfad und Prozessname stimmt
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcess = objWMIService.ExecQuery("Select * From Win32_Process Where Name = 'myapp.exe'")
' Hilfe -Processname- : Dateiname+Endung, z.B. "Meine App.exe" - Leer lassen, um Prozesskillen ganz zu deaktivieren
' Hilfe -Apptitle- : Fenster�berschrift (komplett oder Teil davon), z.B. "Brennvorgang Free 0% abges" [...]
' Hilfe -ProcessPfad- : Ordnerpfadpfad (komplett oder Teil davon), z.B. "AppData\Local\Temp\"
ProcessName = "vimage.exe"
'Apptitle = ""
ProcessPfad = "\VImage\"
Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
  ' Connect to Windows Management Instrumentation (WMI) object using a moniker
  Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
  ' Execute a query to get all running instances of specified application
  Set colProcess = objWMIService.ExecQuery("Select * From Win32_Process Where Name = '" & strApplication & "'")
  ' Terminate all running instances of the application
 Set WshShell = WScript.CreateObject ("WScript.Shell")
call KillVImage(ProcessName)
Function KillVImage(ProcessName) 'Hier ist ProcessName keine Variable, sondern ein feststehender Begriff!
    Dim objWMIService, colProcess
    Dim strComputer, strList, p, q
    Dim i :i= 0 
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
    Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = '" & ProcessName & "'")
    For Each q in colProcess
    ' msgbox "Prozesspfad: " & q.ExecutablePath
    nPos = InStr(q.ExecutablePath, ProcessPfad)  'binary comparison
    if nPos="0" Then
    ' msgbox "Debugginginfo: String nicht gefunden, nicht l�schen..."
    Else
    ' msgbox "Debugginginfo: String ist Teil des ausgef�hrten Pfads - Prozess wird gekillt..."
    q.Terminate
    End if
    Next
    i = i+1        
End Function



Set fso = CreateObject("Scripting.FileSystemObject") 
'warningbox = MsgBox("Fertig. Zum Testen bitte eine niedrige Aufl�sung ausprobieren.",0+64+4096+256+65536+131072,"Bildschirmskalierung")


end if	' Ab hier ist es nicht mehr wichtig, dass das Skript mit Adminrechten l�uft (output muss nicht mehr = YES sein)

End if	' Ab hier werden Befehle unabh�ngig der Antwort aus der MessageBox ausgef�hrt
