Sub Main
	Dim objSh, strPOPD, FSO

	hs.WriteLog "Harmony", "PowerOff Woonkamer TV"
	
	HarmonyHUBIP = "192.168.111.64"
	HarmonyHUBCommand = "PowerOff"
	HarmonyHUBCLIFolder = "C:\Program Files (x86)\HomeSeer HSPRO\Scripts\harmonyHubCLI"
	HarmonyHUBCLI = HarmonyHUBCLIFolder & "\harmonyHubCLI.js"
	HarmonyHUBLog = "C:\Program Files (x86)\HomeSeer HSPRO\Logs\PowerOffWoonkamerTV.log"
	DoubleQuote = Chr(34)
	GreaterThan = Chr(62)
 
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set OutPutFile = FSO.OpenTextFile(HarmonyHUBLog ,2 , True)
	OutPutFile.WriteLine(now & " Running node script...")
	OutPutFile.Close

	command = "%comspec% /c NODE.EXE " & DoubleQuote & HarmonyHUBCLI & DoubleQuote _
		& " -l " & HarmonyHUBIP & " -a " & DoubleQuote & HarmonyHUBCommand & DoubleQuote _
		& " " & GreaterThan & GreaterThan & " " & DoubleQuote & HarmonyHUBLog & DoubleQuote & " 2" & GreaterThan & "&1"
	
	Set objSh = CreateObject("Wscript.Shell")
	strPOPD = objSh.CurrentDirectory
	objSh.CurrentDirectory = HarmonyHUBCLIFolder
	objSh.Run command, 0, True
	objSh.CurrentDirectory = strPOPD
	Set objSh = Nothing

	Set OutPutFile = FSO.OpenTextFile(HarmonyHUBLog ,8 , True)
	OutPutFile.WriteLine(now & " Done running node script...")
	Set FSO = Nothing
End Sub

	' FileSystemObject parameter
	' 
	' 1 = read
	' 2 = write
	' 8 = append
