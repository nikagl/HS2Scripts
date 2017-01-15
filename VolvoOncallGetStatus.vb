Imports System.IO
Imports System.Diagnostics
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public logName As String = "VolvoPrint"													'set log type for HS logging
Public Debug As Boolean = False															'set to True if the script give you errors and it will provide additional info in the log to help troubleshoot

Sub SetDevice(ByVal deviceID As String, _
				ByVal deviceName As String, _
				ByVal deviceValue As String, _
				Optional ByVal deviceRoom As String = "_volvo", _
				Optional ByVal deviceFloor As String = "_volvo")						'use correct parameters = SetDevice("O1", "Device Name", "Device Value", "Room", "Floor")

	dim dv
	
	If Debug Then
		hs.writelog(logName, "Getting device: " & deviceName & " | " & deviceID)		'write the device name and id to log
		hs.writelog(logName, "Room: " & deviceRoom & " & deviceFloor: " & deviceID)		'write the device room and floor to log
		hs.writelog(logName, "Value: " & deviceValue)									'write the device value and id to log
	End If
	dv = hs.DeviceExistsRef(deviceID)													'check if the device exists
	If Debug Then hs.writelog(logName, " dv: " & dv)
	If dv = -1 then
		If Debug Then
			hs.writelog(logName, "Creating device: " _
									& Left(deviceID, 1) _
									& Mid(deviceID, 2, len(deviceID) - 1))				'write the device id again to log
		End If
		dv = hs.NewDeviceEx(deviceName)													'create the device
		dv.location = deviceFloor														'set the floor
		dv.location2 = deviceRoom														'set the room
		dv.hc = Left(deviceID, 1)														'set the HomeCode (letter)
		dv.dc = Mid(deviceID, 2, len(deviceID) - 1)										'set the DeviceCode (number)
		dv.misc = "&h10"																'set it to status-only device (cannot be controlled).
		dv.dev_type_string="Virtual" 													'set it to be a Virtual Device
	End If
	
	If Debug Then hs.writelog(logName, "Setting device: " & deviceName)					'write the device name again (which is going to be set)
																						'set the device, replace "null" with N/A, false with yes, true with no and remove double quotes and spaces (before and after)
	hs.SetDeviceString(deviceID, _
		deviceValue.Replace(CHR(34), "").Replace("null", "N/A").Replace("false", "No").Replace("true", "Yes").Trim())
End Sub

Sub Main(ByVal Parms As String)
	Dim App = "C:\Python27\python.exe" 													'set Python executable path
	Dim AppPath = "C:\Program Files (x86)\HomeSeer HSPRO\scripts\volvooncall\"			'set Python script path
	Dim AppParam = "voc print"															'parameters to run the app with									voc print washerFluidLevel
	
	logName = logName & " "																'redefine log type for HS log
	If Debug Then hs.writelog(logName, "Running: " & App & " " & AppPath & AppParam)	'write the result to log
	
	Dim strOutput As String
	
	Try
		Dim myproc As New Process
		With myproc
			.StartInfo.WorkingDirectory = AppPath										'set the folder
			.StartInfo.FileName = App													'set the app to start
			.StartInfo.Arguments = AppParam												'set the parameters to use
			.StartInfo.RedirectStandardOutput = True									'make sure standard output is returned to script
			.StartInfo.UseShellExecute = False											'execute from shell
			.StartInfo.CreateNoWindow = False											'without creating a new window
			.Start()
		End With
		
		Dim StdOut As StreamReader = myproc.StandardOutput 								'open file
		strOutput = StdOut.ReadToEnd() 													'read file
		myproc.WaitForExit() 															'wait for full read
		myproc.Close() 																	'close file
		
		Dim deserialized = JsonConvert.DeserializeObject(strOutput)						'get the json results from the python script from standard output
		
		SetDevice("O1", "Volvo averageFuelConsumption", _
		SetDevice("O2", "Volvo averageSpeed", _
		SetDevice("O3", "Volvo brakeFluid", _
		SetDevice("O4", "Volvo carLocked", _
		SetDevice("O5", "Volvo connectionStatus", _
		SetDevice("O6", "Volvo distanceToEmpty", _
		SetDevice("O7", "Volvo engineRunning", _
		SetDevice("O8", "Volvo fuelAmount", _
		SetDevice("O9", "Volvo fuelAmountLevel", _
		SetDevice("O10", "Volvo fuelTankVolume", _
		SetDevice("O11", "Volvo grossWeight", _
		SetDevice("O12", "Volvo odometer", _
		SetDevice("O13", "Volvo remoteClimatizationStatus", _
		SetDevice("O14", "Volvo serviceWarningStatus", _
		SetDevice("O15", "Volvo tripMeter1", _
		SetDevice("O16", "Volvo tripMeter2", _
		SetDevice("O17", "Volvo hoodOpen", _
		SetDevice("O18", "Volvo frontRightDoorOpen", _
		SetDevice("O19", "Volvo frontLeftDoorOpen", _
		SetDevice("O20", "Volvo rearRightDoorOpen", _
		SetDevice("O21", "Volvo rearLeftDoorOpen", _
		SetDevice("O22", "Volvo tailgateOpen", _
		SetDevice("O23", "Volvo hvBatteryLevel", _
		SetDevice("O24", "Volvo distanceToHVBatteryEmpty", _
		SetDevice("O25", "Volvo distanceToHVBatteryEmptyTimestamp", _
		SetDevice("O26", "Volvo timeToHVBatteryFullyChargedTimestamp", _
		SetDevice("O27", "Volvo hvBatteryChargeModeStatus", _
		SetDevice("O28", "Volvo hvBatteryChargeStatus", _
		SetDevice("O29", "Volvo hvBatteryChargeStatusDerived", _
		SetDevice("O30", "Volvo hvBatteryChargeWarning", _
		SetDevice("O31", "Volvo frontLeftTyrePressure", _
		SetDevice("O32", "Volvo frontRightTyrePressure", _
		SetDevice("O33", "Volvo rearLeftTyrePressure", _
		SetDevice("O34", "Volvo rearRightTyrePressure", _
		SetDevice("O35", "Volvo frontLeftWindowOpen", _
		SetDevice("O36", "Volvo frontRightWindowOpen", _
		SetDevice("O37", "Volvo rearLeftWindowOpen", _
		SetDevice("O38", "Volvo rearRightWindowOpen", _
		SetDevice("O39", "Volvo longitude", _
		SetDevice("O40", "Volvo latitude", _
		SetDevice("O41", "Volvo speed", _
		SetDevice("O42", "Volvo heading", _
		SetDevice("O43", "Volvo Google Maps", _
					"<a href='https://maps.google.com/maps?&z=18&q=" _
					& deserialized("position")("latitude").Tostring() & "," _
					& deserialized("position")("longitude").Tostring() & "&ll=" _
					& deserialized("position")("latitude").Tostring() & "," _
					& deserialized("position")("longitude").Tostring() & "&ll=" _
					& "' target='_blank'>Click here!</a>", _
	Catch pEx As Exception
		strOutput = pEx.Message
		If Debug Then hs.WriteLog(logName, "Error: " & strOutput) 						'write the error to log
	End Try
	
	hs.WriteLog("Event", "Done running script...") 										'write the exit to log
End Sub