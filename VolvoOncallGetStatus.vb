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
		
		SetDevice("O1", "Volvo averageFuelConsumption", _					deserialized("averageFuelConsumption").Tostring(), _					"_volvo", _					"_volvoFuel")														'call SetDevice subroutine with value from the json results
		SetDevice("O2", "Volvo averageSpeed", _					deserialized("averageSpeed").Tostring(), _					"_volvo", _					"_volvoMeters")
		SetDevice("O3", "Volvo brakeFluid", _					deserialized("brakeFluid").Tostring(), _					"_volvo", _					"_volvoMeters")
		SetDevice("O4", "Volvo carLocked", _					deserialized("carLocked").Tostring(), _					"_volvo", _					"_volvoMeters")
		SetDevice("O5", "Volvo connectionStatus", _					deserialized("connectionStatus").Tostring(), _					"_volvo", _					"_volvoBattery")
		SetDevice("O6", "Volvo distanceToEmpty", _					deserialized("distanceToEmpty").Tostring(), _					"_volvo", _					"_volvoFuel")
		SetDevice("O7", "Volvo engineRunning", _					deserialized("engineRunning").Tostring(), _					"_volvo", _					"_volvoMeters")
		SetDevice("O8", "Volvo fuelAmount", _					deserialized("fuelAmount").Tostring(), _					"_volvo", _					"_volvoFuel")
		SetDevice("O9", "Volvo fuelAmountLevel", _					deserialized("fuelAmountLevel").Tostring(), _					"_volvo", _					"_volvoFuel")
		SetDevice("O10", "Volvo fuelTankVolume", _					deserialized("fuelTankVolume").Tostring(), _					"_volvo", _					"_volvoFuel")
		SetDevice("O11", "Volvo grossWeight", _					deserialized("grossWeight").Tostring(), _					"_volvo", _					"_volvoMeters")
		SetDevice("O12", "Volvo odometer", _					deserialized("odometer").Tostring() / 1000, _					"_volvo", _					"_volvoMeters")													'meter results are in meters
		SetDevice("O13", "Volvo remoteClimatizationStatus", _					deserialized("remoteClimatizationStatus").Tostring(), _					"_volvo", _					"_volvoBattery")
		SetDevice("O14", "Volvo serviceWarningStatus", _					deserialized("serviceWarningStatus").Tostring(), _					"_volvo", _					"_volvoMeters")
		SetDevice("O15", "Volvo tripMeter1", _					deserialized("tripMeter1").Tostring() / 1000, _					"_volvo", _					"_volvoMeters")													'meter results are in meters
		SetDevice("O16", "Volvo tripMeter2", _					deserialized("tripMeter2").Tostring() / 1000, _					"_volvo", _					"_volvoMeters")													'meter results are in meters
		SetDevice("O17", "Volvo hoodOpen", _					deserialized("doors")("hoodOpen").Tostring(), _					"_volvo", _					"_volvoDoors")
		SetDevice("O18", "Volvo frontRightDoorOpen", _					deserialized("doors")("frontRightDoorOpen").Tostring(), _					"_volvo", _					"_volvoDoors")
		SetDevice("O19", "Volvo frontLeftDoorOpen", _					deserialized("doors")("frontLeftDoorOpen").Tostring(), _					"_volvo", _					"_volvoDoors")
		SetDevice("O20", "Volvo rearRightDoorOpen", _					deserialized("doors")("rearRightDoorOpen").Tostring(), _					"_volvo", _					"_volvoDoors")
		SetDevice("O21", "Volvo rearLeftDoorOpen", _					deserialized("doors")("rearLeftDoorOpen").Tostring(), _					"_volvo", _					"_volvoDoors")
		SetDevice("O22", "Volvo tailgateOpen", _					deserialized("doors")("tailgateOpen").Tostring(), _					"_volvo", _					"_volvoDoors")
		SetDevice("O23", "Volvo hvBatteryLevel", _					deserialized("hvBattery")("hvBatteryLevel").Tostring(), _					"_volvo", _					"_volvoBattery")
		SetDevice("O24", "Volvo distanceToHVBatteryEmpty", _					deserialized("hvBattery")("distanceToHVBatteryEmpty").Tostring(), _					"_volvo", _					"_volvoBattery")
		SetDevice("O25", "Volvo distanceToHVBatteryEmptyTimestamp", _					deserialized("hvBattery")("distanceToHVBatteryEmptyTimestamp").Tostring(), _					"_volvo", _					"_volvoBattery")
		SetDevice("O26", "Volvo timeToHVBatteryFullyChargedTimestamp", _					deserialized("hvBattery")("timeToHVBatteryFullyChargedTimestamp").Tostring(), _					"_volvo", _					"_volvoBattery")
		SetDevice("O27", "Volvo hvBatteryChargeModeStatus", _					deserialized("hvBattery")("hvBatteryChargeModeStatus").Tostring(), _					"_volvo", _					"_volvoBattery")
		SetDevice("O28", "Volvo hvBatteryChargeStatus", _					deserialized("hvBattery")("hvBatteryChargeStatus").Tostring(), _					"_volvo", _					"_volvoBattery")
		SetDevice("O29", "Volvo hvBatteryChargeStatusDerived", _					deserialized("hvBattery")("hvBatteryChargeStatusDerived").Tostring(), _					"_volvo", _					"_volvoBattery")
		SetDevice("O30", "Volvo hvBatteryChargeWarning", _					deserialized("hvBattery")("hvBatteryChargeWarning").Tostring(), _					"_volvo", _					"_volvoBattery")
		SetDevice("O31", "Volvo frontLeftTyrePressure", _					deserialized("tyrePressure")("frontLeftTyrePressure").Tostring(), _					"_volvo", _					"_volvoTires")
		SetDevice("O32", "Volvo frontRightTyrePressure", _					deserialized("tyrePressure")("frontRightTyrePressure").Tostring(), _					"_volvo", _					"_volvoTires")
		SetDevice("O33", "Volvo rearLeftTyrePressure", _					deserialized("tyrePressure")("rearLeftTyrePressure").Tostring(), _					"_volvo", _					"_volvoTires")
		SetDevice("O34", "Volvo rearRightTyrePressure", _					deserialized("tyrePressure")("rearRightTyrePressure").Tostring(), _					"_volvo", _					"_volvoTires")
		SetDevice("O35", "Volvo frontLeftWindowOpen", _					deserialized("windows")("frontLeftWindowOpen").Tostring(), _					"_volvo", _					"_volvoWindows")
		SetDevice("O36", "Volvo frontRightWindowOpen", _					deserialized("windows")("frontRightWindowOpen").Tostring(), _					"_volvo", _					"_volvoWindows")
		SetDevice("O37", "Volvo rearLeftWindowOpen", _					deserialized("windows")("rearLeftWindowOpen").Tostring(), _					"_volvo", _					"_volvoWindows")
		SetDevice("O38", "Volvo rearRightWindowOpen", _					deserialized("windows")("rearRightWindowOpen").Tostring(), _					"_volvo", _					"_volvoWindows")
		SetDevice("O39", "Volvo longitude", _					deserialized("position")("longitude").Tostring(), _					"_volvo", _					"_volvoLocation")
		SetDevice("O40", "Volvo latitude", _					deserialized("position")("latitude").Tostring(), _					"_volvo", _					"_volvoLocation")
		SetDevice("O41", "Volvo speed", _					deserialized("position")("speed").Tostring(), _					"_volvo", _					"_volvoLocation")
		SetDevice("O42", "Volvo heading", _					deserialized("position")("heading").Tostring(), _					"_volvo", _					"_volvoLocation")
		SetDevice("O43", "Volvo Google Maps", _
					"<a href='https://maps.google.com/maps?&z=18&q=" _
					& deserialized("position")("latitude").Tostring() & "," _
					& deserialized("position")("longitude").Tostring() & "&ll=" _
					& deserialized("position")("latitude").Tostring() & "," _
					& deserialized("position")("longitude").Tostring() & "&ll=" _
					& "' target='_blank'>Click here!</a>", _					"_volvo", _					"_volvoLocation")													'add google maps link as device
	Catch pEx As Exception
		strOutput = pEx.Message
		If Debug Then hs.WriteLog(logName, "Error: " & strOutput) 						'write the error to log
	End Try
	
	hs.WriteLog("Event", "Done running script...") 										'write the exit to log
End Sub
