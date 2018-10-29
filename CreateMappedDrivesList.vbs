strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objOutFile = objFSO.CreateTextFile(".\ReconnectSharedDrives.bat")

Set colDrives = objWMIService.ExecQuery _
    ("Select * From Win32_LogicalDisk Where DriveType = 4")

For Each objDrive in colDrives
    objOutFile.WriteLine("net use " &objDrive.DeviceID & " " &  _
      objDrive.ProviderName )
Next

objOutFile.Close