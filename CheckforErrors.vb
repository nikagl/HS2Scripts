Public Sub Main(ByVal Parms As Object)
	Dim s As System.IO.StreamReader 
	Dim theData As String
	Dim ignoreError As String
	Dim SendToAddress As String
	Dim ReplyToAddress As String
	Dim LogFile as String
	Dim ErrorLine As String
	Dim Debug As String
	Dim ErrorFound As String
	Dim ErrorText As String
	Debug=0
	ErrorFound=0
	
	If Debug=1 then hs.writelog("checklogDebug", "Log check started...")

	SendToAddress = hs.GetINISetting("Settings", "smtp_to", "", "settings.ini")
	ReplyToAddress = hs.GetINISetting("Settings", "smtp_from", "", "settings.ini")
	LogFile = "C:\Program Files (x86)\HomeSeer HSPRO\Logs\AH.log"
	ErrorText="ERROR:"
	ignoreError = "Sensor Multilevel Report for"
	
    If System.IO.File.Exists(LogFile) = True Then
		If Debug=1 then hs.writelog("checklogDebug", "Logfile found...")
		
		s = System.IO.File.OpenText(LogFile)
		Do While s.Peek >= 0
			theData = s.ReadLine()

			If InStr(Ucase(theData), Ucase(ErrorText)) > 0 AND InStr(Ucase(theData), Ucase(ignoreError)) = 0 Then
				ErrorFound=1
				ErrorLine=theData
			End If
		Loop
		s.Close

		If ErrorFound = 1 Then
			If Debug=1 then hs.writelog("checklogDebug", "Error found!")
			hs.SendEmail(SendToAddress, ReplyToAddress, "Homeseer, error found!", ErrorLine)
		Else
			If Debug=1 then hs.writelog("checklogDebug", "No error found!")
		End If
	Else
		If Debug=1 then hs.writelog("checklogDebug", "Logfile not found...")
	End If
	
	If Debug=1 then hs.writelog("checklogDebug", "Log check done...")
End Sub
