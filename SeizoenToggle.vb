Public Sub Main(ByVal Parms As String) 
	Dim devString As String
	Dim theDevice As String 

	theDevice = "V2"
	devString = hs.DeviceString(theDevice) 

	hs.writelog ("DEBUG", "devString = " & devString)

	'' Toggle it 
	If devString = "Winter" Then
		hs.SetDeviceString(theDevice, "Zomer") 
	Else
		hs.SetDeviceString(theDevice, "Winter") 
	End If

	devString = hs.DeviceString(theDevice) 
	hs.writelog ("DEBUG", "devString = " & devString)
End Sub  
