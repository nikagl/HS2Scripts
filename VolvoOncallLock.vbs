Sub Main
	Dim objSh, strPOPD, FSO

	hs.WriteLog "Volvo","Lock"
	
	VolvoOncallFolder = "C:\Program Files (x86)\HomeSeer HSPRO\scripts\volvooncall"
	VolvoOncall = VolvoOncallFolder & "\voc"
	VolvoOncallCommand = "lock"
	VolvoOncallLog = "C:\Program Files (x86)\HomeSeer HSPRO\Logs\VolvoOncallLock.log"
	DoubleQuote = Chr(34)
	GreaterThan = Chr(62)
 
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set OutPutFile = FSO.OpenTextFile(VolvoOncallLog ,2 , True)
	OutPutFile.WriteLine(now & " Running python script...")
	OutPutFile.Close

	command = "%comspec% /c PYTHON.EXE " & DoubleQuote & VolvoOncall & DoubleQuote _
		& " -vv " & VolvoOncallCommand _
		& " " & GreaterThan & GreaterThan & " " & DoubleQuote & VolvoOncallLog & DoubleQuote & " 2" & GreaterThan & "&1"
	
	Set objSh = CreateObject("Wscript.Shell")
	strPOPD = objSh.CurrentDirectory
	objSh.CurrentDirectory = VolvoOncallFolder
	objSh.Run command, 0, True
	objSh.CurrentDirectory = strPOPD
	Set objSh = Nothing

	Set OutPutFile = FSO.OpenTextFile(VolvoOncallLog ,8 , True)
	OutPutFile.WriteLine(now & " Done running python script...")
	Set FSO = Nothing
End Sub

	' FileSystemObject parameter
	' 
	' 1 = read
	' 2 = write
	' 8 = append
