Sub Main(Byval parm as Object)
		Dim pluginname As String = parm.ToString.Split("#")(0)
		Dim buttonname As String = parm.ToString.Split("#")(1)
		Dim deviceID As String = parm.ToString.Split("#")(2)
		
'		hs.writelog("Debug PressButton.vb","Running script on plugin " & pluginname & ", button " & buttonname & " and device " & deviceID)
		
		Dim dev As Scheduler.Classes.DeviceClass = hs.GetDeviceByRef(deviceID)
		
		Dim SupportsExtButtons As Boolean = False
		Try
				SupportsExtButtons = hs.Plugin(pluginname).SupportsExtendedButtons
				Catch
		End Try
		
		If SupportsExtButtons Then
'				hs.writelog("Debug PressButton.vb","Supports ExtButtons")
				hs.Plugin(pluginname).ButtonPressEx(buttonname, dev)
		Else
'				hs.writelog("Debug PressButton.vb","Does not support ExtButtons")
				hs.Plugin(pluginname).ButtonPress(buttonname)
		End If
End Sub
