Public Sub Main(ByVal Parms As String) 
	Dim devString As String
	Dim theDevice As String 

	theDevice = "V3"
	devString = hs.DeviceString(theDevice) 

	hs.writelog ("DEBUG", "devString = " & devString)

	'' Toggle it 
	If devString = "Vakantie" Then
		hs.SetDeviceString(theDevice, "Geen Vakantie") 
	Else
		hs.SetDeviceString(theDevice, "Vakantie") 
	End If

	devString = hs.DeviceString(theDevice) 
	hs.writelog ("DEBUG", "devString = " & devString)
End Sub  
