Sub Main
	dim en
	dim dv

        Set en = hs.GetDeviceEnumerator
        if IsObject(en) then
        else
        	hs.WriteLog "Enumerator","----------- The device enumerator is invalid ---------"
        end if
        
        Do while not en.Finished
		if enCountChanged then
			hs.WriteLog "Enumerator","----------- The device count has changed ---------"
		end if

		Set dv = en.GetNext
		if not dv is nothing then
			if dv.dev_type_string = "Powermax Zone: Delay 1" or dv.dev_type_string = "Powermax Zone: Perimeter" then
'				hs.WriteLog "DEBUG", "Device Name: " & dv.name 
'				hs.WriteLog "DEBUG", "Device Code: " & dv.hc&dv.dc
'				hs.WriteLog "DEBUG", "Device String: " & hs.DeviceString(dv.hc&dv.dc)

				if hs.DeviceString(dv.hc&dv.dc) = "Closed" then
					 hs.SetDeviceValue dv.hc&dv.dc, 1
				else
					 hs.SetDeviceValue dv.hc&dv.dc, 2
				end if
	                end if
		end if
	Loop
End Sub