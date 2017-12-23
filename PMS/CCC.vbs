On error resume next
strToolbarName = "CALLCENTRETOOLBAR.ETB" 'Toolbar to make visible

Dim strMargins

'Map G: and O: drives
Dim WshNet
Set WshNet = CreateObject("WScript.Network")
'WshNet.MapNetworkDrive "G:", "\\melfiler5\cgumel$","True"
'WshNet.MapNetworkDrive "O:", "\\melfiler5\BP_Partners$","True"

Dim aLettersMapped(26)

Dim objNetwork, objDrive, intDrive, intNetLetter 

Set objNetwork = CreateObject("WScript.Network") 
Set objDrive = objNetwork.EnumNetworkDrives 

For intDrive = 0 to objDrive.Count -1 Step 2
	intNetLetter = IntNetLetter +1

	aLettersMapped(asc(objDrive.Item(intDrive)) - 65) = -1 'A mapped means aLettersMapped(0) = -1
	If Ucase(objDrive.Item(intDrive +1)) = "\\MELFILER5\CGUMEL$" then
		intG = -1
	end if
	If Ucase(objDrive.Item(intDrive +1)) = "\\MELFILER5\BP_PARTNERS$" then
		intO = -1
	end if
Next

If intG <> -1 then 'G not yet mapped
	'Try to map to G
	If aLettersMapped(asc("G")-65) <> -1 then
		strMapGTo = "G"
	Else
		intCheckLetters = asc("J") - 65 'Start checking at J
		Do While aLettersMapped(intCheckLetters) = -1
			intCheckLetters = intCheckLetters + 1
		Loop
		WshNet.MapNetworkDrive chr(intCheckLetters + 65) & ":", "\\melfiler5\cgumel$","True"
	End If
End If

If intO <> -1 then 'O not yet mapped
	'Try to map to O
	If aLettersMapped(asc("O")-65) <> -1 then
		strMapGTo = "O"
	Else
		intCheckLetters = asc("P") - 65 'Start checking at P
		Do While aLettersMapped(intCheckLetters) = -1
			intCheckLetters = intCheckLetters + 1
		Loop
		WshNet.MapNetworkDrive chr(intCheckLetters + 65) & ":", "\\melfiler5\BP_Partners$","True"
	End If
End If

'Create a scripting object
Dim WSHShell
Set WSHShell = WScript.CreateObject("WScript.Shell")

'Write the macro location to the registry
WSHShell.RegWrite "HKCU\Software\Attachmate\EXTRA!\WorkstationUser\Preferences\RemoteSchemePath", "\\melfiler5\cgumel$\SharedMacros"
WSHShell.RegWrite "HKCU\Software\Attachmate\EXTRA!\WorkstationUser\Preferences\SaveSettingsOnClose", "YES"
'Set standard keyboard to allow Ctrl-C, Ctrl-V, Ctrl-X
WSHShell.RegWrite "HKCU\Software\Attachmate\EXTRA!\WorkstationUser\Preferences\Keyboard", "key101"

'Fix Explorer margins for Easysure printing
strMargins = strMargins & WSHShell.RegRead("HKCU\Software\Microsoft\Internet Explorer\PageSetup\header")
strMargins = strMargins & WSHShell.RegRead("HKCU\Software\Microsoft\Internet Explorer\PageSetup\footer")
strMargins = strMargins & WSHShell.RegRead("HKCU\Software\Microsoft\Internet Explorer\PageSetup\margin_bottom")
strMargins = strMargins & WSHShell.RegRead("HKCU\Software\Microsoft\Internet Explorer\PageSetup\margin_left")
strMargins = strMargins & WSHShell.RegRead("HKCU\Software\Microsoft\Internet Explorer\PageSetup\margin_right")
strMargins = strMargins & WSHShell.RegRead("HKCU\Software\Microsoft\Internet Explorer\PageSetup\margin_top")

If strMargins <> "1.222050.750000.750001.22205" then
	WSHShell.RegWrite "HKCU\Software\Microsoft\Internet Explorer\PageSetup\header",""
	WSHShell.RegWrite "HKCU\Software\Microsoft\Internet Explorer\PageSetup\footer",""
	WSHShell.RegWrite "HKCU\Software\Microsoft\Internet Explorer\PageSetup\margin_bottom","1.22205"
	WSHShell.RegWrite "HKCU\Software\Microsoft\Internet Explorer\PageSetup\margin_left","0.75000"
	WSHShell.RegWrite "HKCU\Software\Microsoft\Internet Explorer\PageSetup\margin_right","0.75000"
	WSHShell.RegWrite "HKCU\Software\Microsoft\Internet Explorer\PageSetup\margin_top","1.22205"
End If

'Close Extra! so the new macros appear
Set objSystem = CreateObject("EXTRA.System")
objSystem.Quit

'Wait till user reopens Extra!
Do While intCount = 0
WSHShell.Popup "Please reopen your mainframe session", 6
Set objSystem = CreateObject("EXTRA.System")
intCount = objSystem.Sessions.Count
Loop

'Reconnect to Extra!
Set objSystem = CreateObject("EXTRA.System")
Set objSession = objSystem.ActiveSession

'Hide all other toolbars
Set toolbarsExtraToolbars = objSession.Toolbars
For intEachToolbar = 0 To toolbarsExtraToolbars.Count
	if toolbarsExtraToolbars(intEachToolbar).Visible = True then
		intHideEach = WSHShell.Popup ("Hide toolbar " & toolbarsExtraToolbars(intEachToolbar).Name & "?", , , 4+32)
		If intHideEach = 6 then
			toolbarsExtraToolbars(intEachToolbar).Visible = False
		end if
	End if
Next ' intEachToolbar
    
'Display the required toolbar
objSession.Toolbars(strToolbarName).Visible = True

objSession.Save

'Done
WSHShell.Popup "Toolbar installation complete."