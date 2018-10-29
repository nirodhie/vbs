On Error Resume Next

Dim objNetwork, objDrives, objShell
Dim strSubst, strSubstVal, strSubstName, strEnumDrive

Set objNetwork = CreateObject("WScript.Network")
Set objShell = CreateObject("Shell.Application")
Set objDrives = objNetwork.EnumNetworkDrives

For i = 0 to objDrives.Count - 1 Step 2
	strSubst = objShell.NameSpace(objDrives.Item(i) & Chr(92)).Self.Name 
	strSubstVal = inStr(1,strSubst, Chr(40)) - 2
	strSubstName = Mid(strSubst, 1, strSubstVal)
	strEnumDrive = strEnumDrive & "Drive Letter: " & objDrives.Item(i) & vbCrlF & "Drive Path: " &  _
		objDrives.Item(i+1) & vbCrLf & vbCrLf
Next
MsgBox strEnumDrive ,, "Your mapped Drives"
