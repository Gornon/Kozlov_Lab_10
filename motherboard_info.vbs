On Error Resume Next
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("D:\Info.txt")
Set objService=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
If Err.Number <> 0 Then
	WScript.Echo Err.Number & ": " & Err.Description
	WScript.Quit
End If
For Each objObject In objService.ExecQuery("SELECT * FROM Win32_BaseBoard")

	info= info& "Caption "& objObject.Caption &chr(10)
	info= info& "Manufacturer  "& objObject.Manufacturer &chr(10)
Next

For Each objMoth In objService.ExecQuery("SELECT * FROM Win32_MotherboardDevice")
	info= info& "Primary Bus Type "& objMoth.PrimaryBusType &chr(10)
	info= info& "Secondary Bus Type  "& objMoth.SecondaryBusType &chr(10)
Next
For Each objObject In objService.ExecQuery("SELECT * FROM Win32_Bus")
	info= info& "Bus Type  "& objObject.BusType &chr(10)
Next
objFile.WriteLine info 
WScript.Echo info






