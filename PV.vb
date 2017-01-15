Public Sub Main(ByVal Parms As String)
	' Make sure to create the Virtual Devices P1 to P8
		' P1 = Energy consumption
		' P2 = Power consumption
		' P3 = Energy generation
		' P4 = Power generation
		' P5 = Weekly Energy consumption
		' P6 = Weekly Energy generation
		' P7 = Monthly Energy consumption
		' P8 = Monthly Energy generation
	
	' Enter your API Key and SystemId from PVOutput here:
	Dim APIKey As String = "8a54f46ca35b4c953ccf4f57d99950060ff35148"
	Dim SystemId As String = "29758"
	Dim Debug As String
	Debug=0

	' Define variables
	Dim URL as String = "/service/r2/getstatus.jsp?key=" & APIKey & "&sid=" & SystemId & ""
	If Debug=1 then hs.writelog("Debug PV.vb","URL = http://pvoutput.org" & URL)

	hs.WaitSecs(1)
	Dim SystemSummary = hs.GetURLex("http://pvoutput.org", URL, "", "80", "TRUE", "FALSE", "")
	If Debug=1 then hs.writelog("Debug PV.vb","SystemSummary = " & SystemSummary)

	Dim strArr As String () = SystemSummary.Split(",")
	Dim pdate as String =  strArr(0)
	If Debug=1 then hs.writelog("Debug PV.vb","pdate = " & pdate)

	Dim ptime as String =  strArr(1)
	If Debug=1 then hs.writelog("Debug PV.vb","ptime = " & ptime)

	Dim Egeneration as Double =  strArr(2)
	If Debug=1 then hs.writelog("Debug PV.vb","Egeneration = " & Egeneration)

	Dim Pgeneration as Double =  strArr(3)
	If Debug=1 then hs.writelog("Debug PV.vb","Pgeneration = " & Pgeneration)

	Dim Econsumption as Double =  strArr(4)
	If Debug=1 then hs.writelog("Debug PV.vb","Econsumption = " & Econsumption)

	Dim Pconsumption as Double =  strArr(5)
	If Debug=1 then hs.writelog("Debug PV.vb","Pconsumption = " & Pconsumption)

	Dim efficiency as String =  strArr(6)
	If Debug=1 then hs.writelog("Debug PV.vb","efficiency = " & efficiency)

	Dim temperature as String =  strArr(7)
	If Debug=1 then hs.writelog("Debug PV.vb","temperature = " & temperature)

	Dim voltage as String =  strArr(8)
	If Debug=1 then hs.writelog("Debug PV.vb","voltage = " & voltage)

	Dim CurrentDate As DateTime = DateTime.Now
	If Debug=1 then hs.writelog("Debug PV.vb","CurrentDate = " & CurrentDate)

	Dim DT As String = DateTime.Now.ToString("yyyyMMdd")
	If Debug=1 then hs.writelog("Debug PV.vb","DT = " & DT)

	Dim MonthDF As String = CurrentDate.Year & CurrentDate.ToString("MM") & "01"
	If Debug=1 then hs.writelog("Debug PV.vb","MonthDF = " & MonthDF)

	Dim WeekDF As String = DateTime.Now.AddDays(-7).ToString("yyyyMMdd")
	If Debug=1 then hs.writelog("Debug PV.vb","WeekDF = " & WeekDF)

	Dim MonthToDateURLString As String = "/service/r2/getstatistic.jsp?df=" & MonthDF & "&dt=" & DT & "&key=" & APIKey & "&sid=" & SystemId & "&c=1"
	If Debug=1 then hs.writelog("Debug PV.vb","MonthToDateURLString = http://pvoutput.org" & MonthToDateURLString)

	hs.WaitSecs(1)
	Dim MonthToDateSystemStatistic = hs.GetURLex("http://pvoutput.org", MonthToDateURLString, "", "80", "TRUE", "FALSE", "")
	If Debug=1 then hs.writelog("Debug PV.vb","MonthToDateSystemStatistic = " & MonthToDateSystemStatistic)

	Dim PastWeekURLString As String = "/service/r2/getstatistic.jsp?df=" & WeekDF & "&dt=" & DT & "&key=" & APIKey & "&sid=" & SystemId & "&c=1"
	If Debug=1 then hs.writelog("Debug PV.vb","PastWeekURLString = http://pvoutput.org" & PastWeekURLString)

	hs.WaitSecs(1)
	Dim PastWeekSystemStatistic = hs.GetURLex("http://pvoutput.org", PastWeekURLString, "", "80", "TRUE", "FALSE", "")
	If Debug=1 then hs.writelog("Debug PV.vb","PastWeekSystemStatistic = " & PastWeekSystemStatistic)

	Dim strArrM As String () = MonthToDateSystemStatistic.Split(",")
	Dim MEconsumption as Double =  strArrM(11)
	If Debug=1 then hs.writelog("Debug PV.vb","MEconsumption = " & MEconsumption)

	Dim MEgeneration as Double =  strArrM(0)
	If Debug=1 then hs.writelog("Debug PV.vb","MEgeneration = " & MEgeneration)

	Dim strArrW As String () = PastWeekSystemStatistic.Split(",")
	Dim WEconsumption as Double =  strArrW(11)
	If Debug=1 then hs.writelog("Debug PV.vb","WEconsumption = " & WEconsumption)

	Dim WEgeneration as Double =  strArrW(0)
	If Debug=1 then hs.writelog("Debug PV.vb","WEgeneration = " & WEgeneration)

	' Write string to Virtual Devices
	If Debug=1 then hs.writelog("Debug PV.vb","Setting virtual devices!")
	hs.SetDeviceString("P1",Econsumption.ToString("#,##0") & " Watt")
	hs.SetDeviceString("P2",Pconsumption.ToString("#,##0") & " Watt")
	hs.SetDeviceString("P3",Egeneration.ToString("#,##0") & " Watt")
	hs.SetDeviceString("P4",Pgeneration.ToString("#,##0") & " Watt")
	hs.SetDeviceString("P5",WEconsumption.ToString("#,##0") & " Watt")
	hs.SetDeviceString("P6",WEgeneration.ToString("#,##0") & " Watt")
	hs.SetDeviceString("P7",MEconsumption.ToString("#,##0") & " Watt")
	hs.SetDeviceString("P8",MEgeneration.ToString("#,##0") & " Watt")
	hs.SetDeviceString("P9",CurrentDate)
	If Debug=1 then hs.writelog("Debug PV.vb","Done setting devices!")
End Sub
