' Dieses Skript schaltet die GPU-Skalierung von AMD-Grafikkarten automatisch an und stellt den Modus auf AspectRatio

'#############
' HINWEIS für mich:
' Gilt nur für AMD (ggf. auch nur für den AMD Catalyst)
'#############

Option Explicit
'On Error Resume Next


dim oReg
dim pReg
dim qReg
dim rReg
dim objRegistry
dim oShell
dim sKeyPath
dim sSubkey
dim tSubkey
dim uSubkey
dim oSubkey
dim aSubKeys
dim bSubKeys
dim cSubKeys
dim dSubKeys
dim sTmpValueName
dim sTmpKeyName1
dim sTmpKeyName2
dim sTmpKeyName3
dim path
dim values
dim Return


' MUSS MIT ADMINISTRATORRECHTEN LAUFEN!

'--- Intel Grafiktreiber: Skaliere GPU bei Einhaltung des Seitenverhältnis
Const cHKLM = &H80000002
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
Set pReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
Set qReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
Set rReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
Set oShell = CreateObject("WScript.Shell")
 
sKeyPath = "SYSTEM\CurrentControlSet\Control\Class"
oReg.EnumKey cHKLM, sKeyPath, aSubKeys
 
If Not (isnull(aSubKeys)) Then
    For Each sSubkey In aSubKeys
        sTmpKeyName1 = "SYSTEM\CurrentControlSet\Control\Class\" & sSubKey
        pReg.EnumKey cHKLM, sTmpKeyName1, bSubKeys
        If Not (isnull(bSubKeys)) Then
        For Each tSubkey In bSubKeys
	if tSubkey = "0000" then
	sTmpKeyName2 =  sTmpKeyName1 & "\" & tSubkey
	'msgbox sTmpKeyName2

               qReg.EnumValues cHKLM, sTmpKeyName2, cSubKeys
               If Not (isnull(cSubKeys)) Then
               For Each oSubkey In cSubKeys
	     'oSubkey = "GPUScaling"		' plus eine unbekannte Zahl dahinter (z.B. "GPUScaling84")
	     if Left(oSubkey, 10) = "GPUScaling" then
	     sTmpValueName =  sTmpKeyName2 & "\" & oSubkey	 
	     'msgbox sTmpValueName     

	    Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    path = sTmpKeyName2
	    values = Array(01,00,00,00)
	    Return = objRegistry.SetBinaryValue(cHKLM, path, oSubkey, values)
	    If (Return = 0) And (Err.Number = 0) Then
    	    'msgbox "Binary value added successfully"
	    Else
    	    ' An error occurred
    	    'msgbox "Binary could not be added!"
	    End If	    'Meldung, ob Binärwert in Registry geschrieben wurde, wurde ausgegeben
	     else	'Erste 10 Buchstaben des Werts sind nicht "GPUScaling"...
	     End if     'egal, ob erste 10 Buchstaben des Werts = "GPUScaling" oder nicht
	     Next
	     End if
	End if
        Next
        End if
        'msgbox "Attempting to write x value to: """ & sTmpKeyName1 & """"
        '''oShell.RegWrite sTmpValueName, x, "REG_DWORD"
        If Err.Number <> 0 Then
            'msgbox "Write failed with errors: " & Err.Number
        Else
            'msgbox "Write succedded."
        End If     
    Next
End if


'--------------------------------------

sKeyPath = "SYSTEM\CurrentControlSet\Control\Video"
oReg.EnumKey cHKLM, sKeyPath, aSubKeys
 
If Not (isnull(aSubKeys)) Then
    For Each sSubkey In aSubKeys
        sTmpKeyName1 = "SYSTEM\CurrentControlSet\Control\Video\" & sSubKey
        pReg.EnumKey cHKLM, sTmpKeyName1, bSubKeys
        If Not (isnull(bSubKeys)) Then
        For Each tSubkey In bSubKeys
	if tSubkey = "0000" then
	sTmpKeyName2 =  sTmpKeyName1 & "\" & tSubkey
	'msgbox sTmpKeyName2

               qReg.EnumValues cHKLM, sTmpKeyName2, cSubKeys
               If Not (isnull(cSubKeys)) Then
               For Each oSubkey In cSubKeys
	     'oSubkey = "GPUScaling"		' plus eine unbekannte Zahl dahinter (z.B. "GPUScaling84")
	     if Left(oSubkey, 10) = "GPUScaling" then
	     sTmpValueName =  sTmpKeyName2 & "\" & oSubkey	 
	     'msgbox sTmpValueName     

	    Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    path = sTmpKeyName2
	    values = Array(01,00,00,00)
	    Return = objRegistry.SetBinaryValue(cHKLM, path, oSubkey, values)
	    If (Return = 0) And (Err.Number = 0) Then
    	    'msgbox "Binary value added successfully"
	    Else
    	    ' An error occurred
    	    'msgbox "Binary could not be added!"
	    End If	    'Meldung, ob Binärwert in Registry geschrieben wurde, wurde ausgegeben

	     else	'Erste 10 Buchstaben des Werts sind nicht "GPUScaling"...
	     End if     'egal, ob erste 10 Buchstaben des Werts = "GPUScaling" oder nicht
	     Next
	     End if
	End if
        Next
        End if
        'msgbox "Attempting to write x value to: """ & sTmpKeyName1 & """"
        '''oShell.RegWrite sTmpValueName, x, "REG_DWORD"
        If Err.Number <> 0 Then
            'msgbox "Write failed with errors: " & Err.Number
        Else
            'msgbox "Write succedded."
        End If     
    Next
End if


'--------------------------------------

sKeyPath = "SYSTEM\CurrentControlSet\Control\Video"

oReg.EnumKey cHKLM, sKeyPath, aSubKeys
If Not (isnull(aSubKeys)) Then
    For Each sSubkey In aSubKeys
        sTmpKeyName1 = "SYSTEM\CurrentControlSet\Control\Video\" & sSubKey
        pReg.EnumKey cHKLM, sTmpKeyName1, bSubKeys
        If Not (isnull(bSubKeys)) Then
        For Each tSubkey In bSubKeys
	if tSubkey = "0000" then
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

	    	Set objRegistry = GetObject("Winmgmts:root\default:StdRegProv")
	    	path = sTmpKeyName3
	    	values = Array(00,00,00,00,01,00,00,00,03,00,00,00,01,00,00,00,08,00,00,00,00,00,00,00)
	    	Return = objRegistry.SetBinaryValue(cHKLM, path, oSubkey, values)
	    	If (Return = 0) And (Err.Number = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Binärwert in Registry geschrieben wurde, wurde ausgegeben
	     	else
	     	End if
	    	Next
	     	End if
		End if	
        	Next
        	End if
        'msgbox "Attempting to write x value to: """ & sTmpKeyName1 & """"
        '''oShell.RegWrite sTmpValueName, x, "REG_DWORD"
        If Err.Number <> 0 Then
            'msgbox "Write failed with errors: " & Err.Number
        Else
            'msgbox "Write succedded."
        End If    
        End if
        Next
        End if
    Next
End if

'--------------------------------------

sKeyPath = "SYSTEM\ControlSet001\Control\GraphicsDrivers\Configuration"

oReg.EnumKey cHKLM, sKeyPath, aSubKeys
If Not (isnull(aSubKeys)) Then
    For Each sSubkey In aSubKeys
        sTmpKeyName1 = "SYSTEM\ControlSet001\Control\GraphicsDrivers\Configuration\" & sSubKey
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
	    	Return = objRegistry.SetDwordValue(cHKLM, path, oSubkey, values)
	    	If (Return = 0) And (Err.Number = 0) Then
    	    		'msgbox "Binary value added successfully"
	    		Else
    	    		' An error occurred
    	    		'msgbox "Binary could not be added!"
	    	End If	    'Meldung, ob Binärwert in Registry geschrieben wurde, wurde ausgegeben

	     	else
	     	End if
	    	Next
	     	End if
		End if  	
        	Next
        	End if
        'msgbox "Attempting to write x value to: """ & sTmpKeyName1 & """"
        '''oShell.RegWrite sTmpValueName, x, "REG_DWORD"
        If Err.Number <> 0 Then
            'msgbox "Write failed with errors: " & Err.Number
        Else
            'msgbox "Write succedded."
        End If    
    Next
End if
