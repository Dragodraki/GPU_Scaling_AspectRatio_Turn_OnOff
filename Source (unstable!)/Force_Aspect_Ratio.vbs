' Dieses Skript schaltet die GPU-Skalierung von AMD-Grafikkarten automatisch an und stellt den Modus auf AspectRatio

Option Explicit
On Error Resume Next

' Nicht als Variable definieren: wscript, ScriptFullName !!!
' Für Textbox statt "wscript.echo" den Befehl "msgbox" verwenden !!!
' Der msgbox/popup-Parameter +131072 erhöht zwar die Textauflösung, aber verhindert die nach definierter Zeit Schließen-Funktion!


Dim fso, warningbox, file, folder, folder2, wshell, wshshell, objShell, myKey1, myKey2, strPfad, Systemdrive, Username, AppData, oFSO, oFolder, sDirectoryPath, oFileCollection, oFile, sDir, iDaysOld, sExt, text, text2, fcreate, fdesc, WshSheg, WshShell1, WshShell2, WshShell3, ws, oShell, oWindowList, EXPLORERUMGEBUNGSVARIABLE, EXPLORERUMGEBUNGSVARIABLE2, oWindow, KeinExplorer, user, objFSO, RIGHTHERE
Dim objAllUsersProgramsFolder, strAllUsersProgramsPath, str, objfolder, objFolderItem, colVerbs, objVerb, objShortcut, objFileSystem, sRegFile, TEST, owindowsFolderPath, newDIR, ziel, MyComputer, strComputer, objWMIService, colOperatingSystems, objOperatingSystem, colProcess, objProcess, strProcess, strWMIQuery
Dim sProcessName, sComputer, oWmi, colProcessList, Repeater, objFile, objReg, objSystemInfo, strHostname, startexe, objstream, strArgs, target_folder, fs, TEMP, USERPROFILE, ALLUSERSPROFILE, LOCALAPPDATA, SYSTEMROOT, RIGHTHEREnoBACKSLASH, desktopPath, link, filesys, PARENT_RIGHTHERE, strRoot, strModify, shell, FULLSCRIPTNAME, SCRIPTFILENAME, key, output, gameexe, RIGHTHEREDoubleSlash, XPmode, exelostbox, KompatibilitaetValue, ScriptingHostProcessName, oReg, strKeyPath, linkname, linkname2, linkname3, linkname4, linkname5, PARENT_RIGHTHEREDoubleSlash, search, skriptfilename, iconpfad, strValueName, dwValue, AlleRunasRobEintraege, NeueRunasRobEintraege
dim ResetteRegistrySchluessel, NeuerStringName_1, NeuerStringWert_1, NeuerStringName_2, NeuerStringWert_2, NeuerStringName_3, NeuerStringWert_3, NeuerStringName_4, NeuerStringWert_4, NeuerStringName_5, NeuerStringWert_5, NeuerDwordName_1, NeuerDwordWert_1, NeuerDwordName_2, NeuerDwordWert_2, NeuerDwordName_3, NeuerDwordWert_3, NeuerDwordName_4, NeuerDwordWert_4, NeuerDwordName_5, NeuerDwordWert_5, Aktiviere_HKCU, Aktiviere_HKLM, strValue, Scriptplayer, DeleteAproDings, strContents, strFirstLine, strNewContents, Aktiviere_ALL_HKCU, oNetwork, NeuerStringWert_1A, NeuerStringWert_2A, NeuerStringWert_3A, NeuerDwordWert_1A, NeuerDwordWert_2A, NeuerDwordWert_3A, drive, GetDriveLetter, label, labelfilename ,mount_disc, GetDriveLetterRaw, Resolution, RIGHTHEREDoubleSlashNoBackslash, WindowsVersionName, WMI, OSs, OS, dtmConvertedDate, objOperatingSyst, value, NewStringWindowsVersion, MajorWindowsVersion, objArgsPScript, objArgsWScript, objArgsCScript, objArgs, MitParameterGestartet, ArgumentTotal, Arg, f, app, ProzessUNDTitelOffen, ProzessUNDPfadOffen, ExternalPathCheck, TitleCheck, ProzessEgal_PfadTitle_Warten, AufAlleProzesseVBSpfadWarten, TempFolderName, BenutzerAusListe, infobox, USERDOMAIN, LASTSHOWEDUSER
dim LastResolutionShowed, AllResolutionList, AndereAufloesung, Mindestbreite, Mindesthoehe, Seitenverhaeltnis, Result, Aufloesungstest, Horizontale, Vertikale, AufloesungCheckpoint, AufloesungAlternative, AufloesungAlternativeCheckpoint, FullPath, arr, dir, path, Counter, pReg, qReg, rReg, sReg, objRegistry, sKeyPath, sSubkey, tSubkey, uSubkey, oSubkey, pSubkey, aSubKeys, bSubKeys, cSubKeys, dSubKeys, eSubKeys, sTmpValueName, sTmpKeyName1, sTmpKeyName2, sTmpKeyName3, values, Return, Log, logpath, logSubkey, logvalues, bArray, BinaryPositionNumberToChange, BinaryReplacement, AMDPath, ScalingProzedurNotwendig
dim ColOSs, OSType, service, AppName, Process, AppName2, strDelete, arrProcesses, arrWindowTitles, i, intProcesses, strProcList, TITLENAME2KILL, Informationsbox, strApplication, colItems, ProcessName, Apptitle, objApp, PID, ProcessPfad, nPos, strList, p, q, strKonto, strFileName




'--- ERKENNE WINDOWS-VERSION (BEI WINDOWS XP GEHE IN DEN XP-MODUS)

Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set objFSO = CreateObject("Scripting.FileSystemObject")


' Variante A - Prüfe per Versionsnummer (eigentlich sicheres Verfahren, solange OS-Nummer nicht größer als 100 Millionen ist)

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


' Variante B - Prüfe per Namen (wenn hier Windows 7/8/10 vorliegt, überschreibe Ergebnis der Variante A)

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
Localappdata = objShell.Namespace(&H1c&).Self.Path	' so ist LOCALAPPDATA auch in Windows XP möglich :)
Set objShell  = CreateObject("Scripting.FileSystemObject")
Temp = CreateObject("Scripting.FileSystemObject").GetParentFolderName(LOCALAPPDATA)
Temp = TEMP & "\Temp"				' TEMP-Ordner für Windows XP (wechselt automatisch bei höherem OS)
if TEMP = USERPROFILE & "\AppData\Temp" then
' msgbox "Laut Variableninhalt handelt es sich um Windows > XP und daher wird die TEMP-Umgebungsvariable angepasst."
SET WshShell = CreateObject ("WScript.Shell")
Temp = WshShell.Environment("Process")("TEMP")		' TEMP-Ordner für Windows 7,8,10
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
' Mache in XP aus der verkürzten Pfadangabe RIGHTHERE den vollen...
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





' --- Adminrechte prüfen

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



'--- Variablen setzen und Funktionenzugehörigkeiten erstellen
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

	    ' Lese zuerst den Binary-String aus und schaue, ob Änderunjg notwendig ist
	    Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    path = sTmpKeyName2
	    oReg.GetBinaryValue cHKLM, path, oSubkey, bArray
	    'msgbox bArray(1 - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    If not bArray(1 - 1) = 1 Then
	    ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung ändern."
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
	    values = Array(01,00,00,00)
	    Return = objRegistry.SetBinaryValue(cHKLM, path, oSubkey, values)
	    If (Return = 0) And (Err.Number = 0) Then
    	    'msgbox "Binary value added successfully"
	    Else
    	    ' An error occurred
    	    'msgbox "Binary could not be added!"
	    End If	    'Meldung, ob Binärwert in Registry geschrieben wurde, wurde ausgegeben
	    End if	    ' Prüfung, ob Skalierung geändert werden muss, abgeschlossen
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

	   	' Lese zuerst den Binary-String aus und schaue, ob Änderunjg notwendig ist
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName3
		BinaryPositionNumberToChange = 5
		BinaryReplacement = 1
	    	oReg.GetBinaryValue cHKLM, path, "BestViewOption", bArray
	   	 'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement Then
	    	ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung ändern."
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
		'Ändere eine bestimmte Position der Binary-Strings
		'Changing value
		'msgbox "String lautet derzeit: " & bArray(BinaryPositionNumberToChange - 1) & " und wird im folgenden ersetzt durch " & BinaryReplacement 	'Mit dieser Änderung stimmt die Position exakt, würde sonst die erste 00 nicht mitzählen
		bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement
		'Write information back
		Return = oReg.SetBinaryValue(cHKLM, path, "BestViewOption", bArray)

	    	If (Return = 0) And (Err.Number = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Binärwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Prüfung, ob Skalierung geändert werden muss, abgeschlossen
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

	    ' Lese zuerst den Binary-String aus und schaue, ob Änderunjg notwendig ist
	    Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    path = sTmpKeyName2
	    oReg.GetBinaryValue cHKLM, path, oSubkey, bArray
	    'msgbox bArray(1 - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    If not bArray(1 - 1) = 1 Then
	    ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung ändern."
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
	    values = Array(01,00,00,00)
	    Return = objRegistry.SetBinaryValue(cHKLM, path, oSubkey, values)
	    If (Return = 0) And (Err.Number = 0) Then
    	    'msgbox "Binary value added successfully"
	    Else
    	    ' An error occurred
    	    'msgbox "Binary could not be added!"
	    End If	    'Meldung, ob Binärwert in Registry geschrieben wurde, wurde ausgegeben
	    End if	    ' Prüfung, ob Skalierung geändert werden muss, abgeschlossen
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

	   	' Lese zuerst den Binary-String aus und schaue, ob Änderunjg notwendig ist
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName3
		BinaryPositionNumberToChange = 5
		BinaryReplacement = 1
	    	oReg.GetBinaryValue cHKLM, path, "BestViewOption", bArray
	   	 'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement Then
	    	ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung ändern."
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
		'Ändere eine bestimmte Position der Binary-Strings
		'Changing value
		'msgbox "String lautet derzeit: " & bArray(BinaryPositionNumberToChange - 1) & " und wird im folgenden ersetzt durch " & BinaryReplacement 	'Mit dieser Änderung stimmt die Position exakt, würde sonst die erste 00 nicht mitzählen
		bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement
		'Write information back
		Return = oReg.SetBinaryValue(cHKLM, path, "BestViewOption", bArray)

	    	If (Return = 0) And (Err.Number = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Binärwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Prüfung, ob Skalierung geändert werden muss, abgeschlossen
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

	   	' Lese zuerst den Binary-String aus und schaue, ob Änderunjg notwendig ist
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName1
		BinaryPositionNumberToChange = 9
		BinaryReplacement = 1		'01 in hexadezimal ist das Ziel! (Angabe muss erfolgen: ohne Anführungsstriche = dezimal!)
	    	oReg.GetBinaryValue cHKLM, path, "ScalingConfig", bArray
	   	 'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement Then
	    	ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung ändern."
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
		'Ändere eine bestimmte Position der Binary-Strings
		'Changing value
		'msgbox "String lautet derzeit: " & bArray(BinaryPositionNumberToChange - 1) & " und wird im folgenden ersetzt durch " & BinaryReplacement 	'Mit dieser Änderung stimmt die Position exakt, würde sonst die erste db nicht mitzählen
		bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement
		'Write information back
		Return = oReg.SetBinaryValue(cHKLM, path, "ScalingConfig", bArray)

	    	If (Return = 0) And (Err.Number = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Binärwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Prüfung, ob Skalierung geändert werden muss, abgeschlossen

	   	' Lese zuerst den Binary-String aus und schaue, ob Änderunjg notwendig ist
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName1
		BinaryPositionNumberToChange = 13
		BinaryReplacement = 237		'ed in hexadezimal ist das Ziel! (Angabe muss erfolgen: ohne Anführungsstriche = dezimal!)
	    	oReg.GetBinaryValue cHKLM, path, "ScalingConfig", bArray
	   	 'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement Then
	    	ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung ändern."
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
		'Ändere eine bestimmte Position der Binary-Strings
		'Changing value
		'msgbox "String lautet derzeit: " & bArray(BinaryPositionNumberToChange - 1) & " und wird im folgenden ersetzt durch " & BinaryReplacement 	'Mit dieser Änderung stimmt die Position exakt, würde sonst die erste db nicht mitzählen
		bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement
		'Write information back
		Return = oReg.SetBinaryValue(cHKLM, path, "ScalingConfig", bArray)

	    	If (Return = 0) And (Err.Number = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Binärwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Prüfung, ob Skalierung geändert werden muss, abgeschlossen



	     	else
	     	End if

        Next
        End if
    Next
End if


'--------------------------------------
' NVIDIA

sKeyPath = "SYSTEM\ControlSet001\services\nvlddmkm\DisplayDatabase"

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

	   	' Lese zuerst den Binary-String aus und schaue, ob Änderunjg notwendig ist
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName1
		BinaryPositionNumberToChange = 9
		BinaryReplacement = 1		'01 in hexadezimal ist das Ziel! (Angabe muss erfolgen: ohne Anführungsstriche = dezimal!)
	    	oReg.GetBinaryValue cHKLM, path, "ScalingConfig", bArray
	   	 'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement Then
	    	ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung ändern."
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
		'Ändere eine bestimmte Position der Binary-Strings
		'Changing value
		'msgbox "String lautet derzeit: " & bArray(BinaryPositionNumberToChange - 1) & " und wird im folgenden ersetzt durch " & BinaryReplacement 	'Mit dieser Änderung stimmt die Position exakt, würde sonst die erste db nicht mitzählen
		bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement
		'Write information back
		Return = oReg.SetBinaryValue(cHKLM, path, "ScalingConfig", bArray)

	    	If (Return = 0) And (Err.Number = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Binärwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Prüfung, ob Skalierung geändert werden muss, abgeschlossen

	   	' Lese zuerst den Binary-String aus und schaue, ob Änderunjg notwendig ist
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName1
		BinaryPositionNumberToChange = 13
		BinaryReplacement = 237		'ed in hexadezimal ist das Ziel! (Angabe muss erfolgen: ohne Anführungsstriche = dezimal!)
	    	oReg.GetBinaryValue cHKLM, path, "ScalingConfig", bArray
	   	 'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement Then
	    	ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung ändern."
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
		'Ändere eine bestimmte Position der Binary-Strings
		'Changing value
		'msgbox "String lautet derzeit: " & bArray(BinaryPositionNumberToChange - 1) & " und wird im folgenden ersetzt durch " & BinaryReplacement 	'Mit dieser Änderung stimmt die Position exakt, würde sonst die erste db nicht mitzählen
		bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement
		'Write information back
		Return = oReg.SetBinaryValue(cHKLM, path, "ScalingConfig", bArray)

	    	If (Return = 0) And (Err.Number = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Binärwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Prüfung, ob Skalierung geändert werden muss, abgeschlossen



	     	else
	     	End if

        Next
        End if
    Next
End if

'--------------------------------------
' NVIDIA

sKeyPath = "SYSTEM\ControlSet001\services\nvlddmkm\State\DisplayDatabase"

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

	   	' Lese zuerst den Binary-String aus und schaue, ob Änderunjg notwendig ist
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName1
		BinaryPositionNumberToChange = 9
		BinaryReplacement = 129		'81 in hexadezimal ist das Ziel! (Angabe muss erfolgen: ohne Anführungsstriche = dezimal!)
	    	oReg.GetBinaryValue cHKLM, path, "ScalingConfig", bArray
	   	'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement Then
	    	ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung ändern."
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
		'Ändere eine bestimmte Position der Binary-Strings
		'Changing value
		'msgbox "String lautet derzeit: " & bArray(BinaryPositionNumberToChange - 1) & " und wird im folgenden ersetzt durch " & BinaryReplacement	'Mit dieser Änderung stimmt die Position exakt, würde sonst die erste db nicht mitzählen
		bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement
		'Write information back
		Return = oReg.SetBinaryValue(cHKLM, path, "ScalingConfig", bArray)

	    	If (Return = 0) And (Err.Number = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Binärwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Prüfung, ob Skalierung geändert werden muss, abgeschlossen

	   	' Lese zuerst den Binary-String aus und schaue, ob Änderunjg notwendig ist
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName1
		BinaryPositionNumberToChange = 13
		BinaryReplacement = 109		'6d in hexadezimal ist das Ziel! (Angabe muss erfolgen: ohne Anführungsstriche = dezimal!)
	    	oReg.GetBinaryValue cHKLM, path, "ScalingConfig", bArray
	   	'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement Then
	    	ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung ändern."
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
		'Ändere eine bestimmte Position der Binary-Strings
		'Changing value
		'msgbox "String lautet derzeit: " & bArray(BinaryPositionNumberToChange - 1) & " und wird im folgenden ersetzt durch " & BinaryReplacement 	'Mit dieser Änderung stimmt die Position exakt, würde sonst die erste db nicht mitzählen
		bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement
		'Write information back
		Return = oReg.SetBinaryValue(cHKLM, path, "ScalingConfig", bArray)

	    	If (Return = 0) And (Err.Number = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Binärwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Prüfung, ob Skalierung geändert werden muss, abgeschlossen



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

	   	' Lese zuerst den Binary-String aus und schaue, ob Änderunjg notwendig ist
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName1
		BinaryPositionNumberToChange = 9
		BinaryReplacement = 129		'81 in hexadezimal ist das Ziel! (Angabe muss erfolgen: ohne Anführungsstriche = dezimal!)
	    	oReg.GetBinaryValue cHKLM, path, "ScalingConfig", bArray
	   	'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement Then
	    	ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung ändern."
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
		'Ändere eine bestimmte Position der Binary-Strings
		'Changing value
		'msgbox "String lautet derzeit: " & bArray(BinaryPositionNumberToChange - 1) & " und wird im folgenden ersetzt durch " & BinaryReplacement	'Mit dieser Änderung stimmt die Position exakt, würde sonst die erste db nicht mitzählen
		bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement
		'Write information back
		Return = oReg.SetBinaryValue(cHKLM, path, "ScalingConfig", bArray)

	    	If (Return = 0) And (Err.Number = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Binärwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Prüfung, ob Skalierung geändert werden muss, abgeschlossen

	   	' Lese zuerst den Binary-String aus und schaue, ob Änderunjg notwendig ist
	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName1
		BinaryPositionNumberToChange = 13
		BinaryReplacement = 109		'6d in hexadezimal ist das Ziel! (Angabe muss erfolgen: ohne Anführungsstriche = dezimal!)
	    	oReg.GetBinaryValue cHKLM, path, "ScalingConfig", bArray
	   	'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement Then
	    	ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung ändern."
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
		'Ändere eine bestimmte Position der Binary-Strings
		'Changing value
		'msgbox "String lautet derzeit: " & bArray(BinaryPositionNumberToChange - 1) & " und wird im folgenden ersetzt durch " & BinaryReplacement 	'Mit dieser Änderung stimmt die Position exakt, würde sonst die erste db nicht mitzählen
		bArray(BinaryPositionNumberToChange - 1) = BinaryReplacement
		'Write information back
		Return = oReg.SetBinaryValue(cHKLM, path, "ScalingConfig", bArray)

	    	If (Return = 0) And (Err.Number = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Binärwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Prüfung, ob Skalierung geändert werden muss, abgeschlossen



	     	else
	     	End if

        Next
        End if
    Next
End if


'--------------------------------------
'ALLE HERSTELLER

sKeyPath = "SYSTEM\ControlSet001\Control\GraphicsDrivers\Configuration"

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
	    	values = "4"
	    	oReg.GetDwordValue cHKLM, path, oSubkey, bArray
	   	'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray = 4 Then
	    	ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung ändern."
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
	    	Return = objRegistry.SetDwordValue(cHKLM, path, oSubkey, values)
	    	If (Return = 0) And (Err.Number = 0) Then
    	    		'msgbox "Binary value added successfully"

		logpath = "Software\ATI\ACE\Settings\Graphics\DeviceDFP"
               	sReg.EnumValues cHKCU, logpath, eSubKeys
               	If Not (isnull(eSubKeys)) Then
               	For Each pSubkey In eSubKeys
		if Left(pSubkey, 17) = "CurrentImageScale" then		' (z.B: "CurrentImageScale,6,4")
		'msgbox pSubkey
		logvalues = "ScaleMaintainAspectRatio"
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
		logvalues = "0"
	    	Log = objRegistry.SetDwordValue(cHKCU, logpath, pSubkey, logvalues)
		End if
		Next
		End if

	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Binärwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Prüfung, ob Skalierung geändert werden muss, abgeschlossen

	     	else
	     	End if
	    	Next
	     	End if
		End if  	
        	Next
        	End if
    Next
End if


'--------------------------------------
'ALLE HERSTELLER

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
	    	values = "4"
	    	oReg.GetDwordValue cHKLM, path, oSubkey, bArray
	   	'msgbox bArray(BinaryPositionNumberToChange - 1)	' Erste Zahl ist auszulesene Binary-Position und zweite korrigiert einen Darstellungsfehler
	    	If not bArray = 4 Then
	    	ScalingProzedurNotwendig = "Ja" 	' "Muss die Skalierung ändern."
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn setowner -ownr " & Chr(34) & "n:S-1-5-32-544" & Chr(34), 0, true
		set oShell = CreateObject("WScript.Shell")
		oShell.Run "SetACL.exe" & " -on HKEY_LOCAL_MACHINE\" & sTmpValueName & " -ot reg -actn ace -ace " & Chr(34) & "n:S-1-5-32-544;p:full" & Chr(34), 0, true
	    	Return = objRegistry.SetDwordValue(cHKLM, path, oSubkey, values)
	    	If (Return = 0) And (Err.Number = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Binärwert in Registry geschrieben wurde, wurde ausgegeben
	    	End if	    ' Prüfung, ob Skalierung geändert werden muss, abgeschlossen

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


' --- Sichere Desktop-Icon-Position mit zusätzlichem Programm

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



' Starte Grafikkarte neu
Set objShell = CreateObject("WScript.Shell")
objShell.Run Chr(34) & "restart-only.exe" & Chr(34) & "", 0, False


' ! Wartezeit, sehr wichtig
Sleep "5000"	'Sleep-Modus exklusiv für "PScript.exe" (= Portable VBS)
WScript.Sleep "5000"	'Sleep-Modus für windowseigenen Script-Host "wscript.exe" und "cscript.exe"


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _
& strComputer & "\root\cimv2")
Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name like '" & "RadeonSoftware.exe" & "'")
For Each objProcess in colProcess
objprocess.ExecutablePath = LCase(objprocess.ExecutablePath)
objProcess.Name = LCase(objProcess.Name)
search="\" & objProcess.Name
AMDPath=left (objprocess.ExecutablePath, instr(objprocess.ExecutablePath, search)-0)
Next
If not (fso.FileExists(AMDPath & "RadeonSoftware.exe")) Then
Else
If not (fso.FileExists(AMDPath & "cncmd.exe")) Then
Else
Set objShell = CreateObject("WScript.Shell")
objShell.Run Chr(34) & AMDPath & "RadeonSoftware.exe" & Chr(34) & "", 1, False
Sleep "300"	'Sleep-Modus exklusiv für "PScript.exe" (= Portable VBS)
WScript.Sleep "300"	'Sleep-Modus für windowseigenen Script-Host "wscript.exe" und "cscript.exe"
objShell.Run Chr(34) & AMDPath & "RadeonSoftware.exe" & Chr(34) & "", 1, False
End if
End if


' ! Wartezeit, sehr wichtig
Sleep "5000"	'Sleep-Modus exklusiv für "PScript.exe" (= Portable VBS)
WScript.Sleep "5000"	'Sleep-Modus für windowseigenen Script-Host "wscript.exe" und "cscript.exe"



' --- Wiederherstellen der Desktop-Icon-Position mit zusätzlichem Programm

Set fso = CreateObject("Scripting.FileSystemObject")
If (fso.FileExists(TEMP & "\DesktopOK\Profiles\DesktopOKProfile.dok")) Then
Set objShell = CreateObject("WScript.Shell")
objShell.Run Chr(34) & RIGHTHERE & "DesktopOK\DesktopOK.exe" & Chr(34) & " /load /silent " & Chr(34) & TEMP & "\DesktopOK\Profiles\DesktopOKProfile.dok" & Chr(34), 0, True
Else
' msgbox "Datei fehlt - Bei der Sicherung zuvor ist etwas schiefgelaufen."
End if



' Öffne AMD-Oberfläche und schließe sie sofort wieder, damit AMD die Registryänderungen anerkennt und keinen Fehler beim Spielen auswirft
Counter = 0
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _
& strComputer & "\root\cimv2")
Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name like '" & "RadeonSoftware.exe" & "'")
For Each objProcess in colProcess
Counter = + 1
objprocess.ExecutablePath = LCase(objprocess.ExecutablePath)
objProcess.Name = LCase(objProcess.Name)
search="\" & objProcess.Name
AMDPath=left (objprocess.ExecutablePath, instr(objprocess.ExecutablePath, search)-0)
Next
If Counter > 0 Then
If not (fso.FileExists(AMDPath & "RadeonSoftware.exe")) Then
Else
If not (fso.FileExists(AMDPath & "cncmd.exe")) Then
Else
Set objShell = CreateObject("WScript.Shell")
objShell.Run Chr(34) & AMDPath & "RadeonSoftware.exe" & Chr(34) & "", 1, False
Sleep "300"	'Sleep-Modus exklusiv für "PScript.exe" (= Portable VBS)
WScript.Sleep "300"	'Sleep-Modus für windowseigenen Script-Host "wscript.exe" und "cscript.exe"
objShell.Run Chr(34) & AMDPath & "RadeonSoftware.exe" & Chr(34) & "", 1, False
Sleep "5000"	'Sleep-Modus exklusiv für "PScript.exe" (= Portable VBS)
WScript.Sleep "5000"	'Sleep-Modus für windowseigenen Script-Host "wscript.exe" und "cscript.exe"
Set objShell = CreateObject("WScript.Shell")
objShell.Run Chr(34) & AMDPath & "cncmd.exe" & Chr(34) & " hide launcher", 0, True
objShell.Run Chr(34) & AMDPath & "cncmd.exe" & Chr(34) & " exit launcher", 0, True
End if
End if
End if




End if	' Überprüfen der ausstehenden GPU-Skalierung abgeschlossen

Set fso = CreateObject("Scripting.FileSystemObject") 
'warningbox = MsgBox("Fertig. Zum Testen bitte eine niedrige Auflösung ausprobieren.",0+64+4096+256+65536+131072,"Bildschirmskalierung")


end if	' Ab hier ist es nicht mehr wichtig, dass das Skript mit Adminrechten läuft (output muss nicht mehr = YES sein)

End if	' Ab hier werden Befehle unabhängig der Antwort aus der MessageBox ausgeführt
