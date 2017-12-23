On error resume next
strToolbarName = "MelbCID.ETB" 'Toolbar to make visible


Dim strMargins

'Map G: and O: drives
Dim WshNet
Set WshNet = CreateObject("WScript.Network")
'WshNet.MapNetworkDrive "U:", "\\melfiler5\Metrocid$\Commercial VIC","True"

Dim aLettersMapped(26)

Dim objNetwork, objDrive, intDrive, intNetLetter 

Set objNetwork = CreateObject("WScript.Network") 
Set objDrive = objNetwork.EnumNetworkDrives 

For intDrive = 0 to objDrive.Count -1 Step 2
	intNetLetter = IntNetLetter +1

	aLettersMapped(asc(objDrive.Item(intDrive)) - 65) = -1 'A mapped means aLettersMapped(0) = -1
	If Ucase(objDrive.Item(intDrive +1)) = "\\melfiler5\Metrocid$\Commercial VIC" then
		intG = -1
	end if
Next

If intG <> -1 then 'U not yet mapped
	'Try to map to U
	If aLettersMapped(asc("U")-65) <> -1 then
		strMapGTo = "U"
	Else
		intCheckLetters = asc("J") - 65 'Start checking at J
		Do While aLettersMapped(intCheckLetters) = -1
			intCheckLetters = intCheckLetters + 1
		Loop
		WshNet.MapNetworkDrive chr(intCheckLetters + 65) & ":", "\\melfiler5\Metrocid$\Commercial VIC","True"
	End If
End If



'Create a scripting object
Dim WSHShell
Set WSHShell = WScript.CreateObject("WScript.Shell")

'Write the macro location to the registry
WSHShell.RegWrite "HKCU\Software\Attachmate\EXTRA!\WorkstationUser\Preferences\RemoteSchemePath", "U:\EDI\Macros\"
WSHShell.RegWrite "HKCU\Software\Attachmate\EXTRA!\WorkstationUser\Preferences\SaveSettingsOnClose", "YES"

'Set standard keyboard to allow Ctrl-C, Ctrl-V, Ctrl-X
WSHShell.RegWrite "HKCU\Software\Attachmate\EXTRA!\WorkstationUser\Preferences\Keyboard", "key101"


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