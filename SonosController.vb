Public Function GetArtWork(ByVal parm) As String
    Dim pi As Object 'As HSPI_SONOSCONTROLLER.HSPI
    pi = hs.Plugin("SONOSCONTROLLER")
    If pi Is Nothing Then
        hs.writelog("Script", "empty")
        Exit Function
    End If
    Dim MusicApi As Object 'As HSMusicAPI
    MusicApi = pi.GetMusicAPI(parm.tostring)
    Return MusicApi.CurrentArtworkFile("")
End Function

Public Function GetNextArtWork(ByVal parm) As String
    Dim pi As Object 'As HSPI_SONOSCONTROLLER.HSPI
    pi = hs.Plugin("SONOSCONTROLLER")
    If pi Is Nothing Then
        hs.writelog("Script", "empty")
        Exit Function
    End If
    Dim MusicApi As Object 'As HSMusicAPI
    Try
        MusicApi = pi.GetMusicAPI(parm.tostring)
        Return "http://Localhost" & MusicApi.NextAlbumArtPath()
    Catch ex As Exception
    End Try
End Function

Public Function GetZoneName(ByVal Parm) As String
    Dim pi As Object 'As HSPI_SONOSCONTROLLER.HSPI
    pi = hs.Plugin("SONOSCONTROLLER")
    If pi Is Nothing Then
        hs.writelog("Script", "empty")
        Exit Function
    End If
    Dim InstanceName As Object 'As HSMusicAPI
    InstanceName = pi.GetInstanceName(parm.tostring)
    Return InstanceName
End Function

Public Function GetLinkgroupDestInfo(ByVal Parm) As String
    Dim pi As Object 'As HSPI_SONOSCONTROLLER.HSPI
    pi = hs.Plugin("SONOSCONTROLLER")
    If pi Is Nothing Then
        hs.writelog("Script", "empty")
        Exit Function
    End If
    hs.writelog("Script", "GetLinkgroupDestInfo called with Parm = " & Parm.ToString)
    Return pi.GetLinkgroupZoneDestination(Parm)
End Function

Public Function GetLinkgroupSourceInfo(ByVal Parm) As String
    Dim pi As Object 'As HSPI_SONOSCONTROLLER.HSPI
    pi = hs.Plugin("SONOSCONTROLLER")
    If pi Is Nothing Then
        hs.writelog("Script", "empty")
        Exit Function
    End If
    Return pi.GetLinkgroupZoneSource(Parm)
End Function



'this is an example of how to call the plugin 
' drop something like this in the text field in HST
' [$SCRIPT=&hs.runex("SonosController.vb" , "PlayRadioStation" , "Kitchen;Pandora - Linkin Park Radio")]
' or directly linked to a button &hs.runex "SonosController.vb" , "PlayRadioStation" , "Kitchen;Pandora - Linkin Park Radio"

Sub PlayRadioStation(ByVal parm)
    Dim pi As Object ' As HSPI_SONOSCONTROLLER.HSPI
    pi = hs.Plugin("SONOSCONTROLLER")
    If pi Is Nothing Then
        hs.writelog("Sonos Controller Script", "empty")
    End If
    Dim MusicApi  'As HSMusicAPI
    Dim Parms() = {}
    Dim ZoneName, RadioStationName
    Dim AlreadyExist As Boolean = False
    Try
        Parms = Split(Parm, ";")
    Catch ex As Exception
        hs.writelog("Sonos Controller Script", "Error in PlayRadioStation with error = " & Ex.Message)
        Exit Sub
    End Try
    ZoneName = Parms(0)
    RadioStationName = Parms(1)
    Try
        MusicApi = pi.GetMusicAPI(ZoneName) ' specify here which Zone you want to use. You can use ZoneName or ZoneIndex
        MusicApi.PlayRadioStation(RadioStationName)
    Catch ex As Exception
        hs.writelog("Sonos Controller Script", "Error in PlayRadioStation with error = " & Ex.Message)
    End Try
End Sub

Sub PlayPlayList(ByVal parm)
    Dim pi As Object ' As HSPI_SONOSCONTROLLER.HSPI
    pi = hs.Plugin("SONOSCONTROLLER")
    If pi Is Nothing Then
        hs.writelog("Sonos Controller Script", "empty")
    End If
    Dim MusicApi  'As HSMusicAPI
    Dim Parms() = {}
    Dim ZoneName, PlayListName
    Dim AlreadyExist As Boolean = False
    Try
        Parms = Split(Parm, ";")
    Catch ex As Exception
        hs.writelog("Sonos Controller Script", "Error in PlayPlayList with error = " & Ex.Message)
        Exit Sub
    End Try
    ZoneName = Parms(0)
    PlayListName = Parms(1)
    Try
        MusicApi = pi.GetMusicAPI(ZoneName) ' specify here which Zone you want to use. You can use ZoneName or ZoneIndex
        MusicApi.PlayMusic("", "", PlayListName) ' accepts input as follows (Artist,Optional Album,Optional PlayList,Optional Genre,Optional Track
    Catch ex As Exception
        hs.writelog("Sonos Controller Script", "Error in PlayPlayList with error = " & Ex.Message)
    End Try
End Sub

Function GetZoneSource(ByVal parm) As String
    ' Example to use in HST Text field [$SCRIPT=&hs.runex ("SonosController.vb" , "GetZoneSource" , "Master Bedroom" )]
    Dim pi As Object ' As HSPI_SONOSCONTROLLER.HSPI
    pi = hs.Plugin("SONOSCONTROLLER")
    If pi Is Nothing Then
        hs.writelog("Sonos Controller Script", "empty")
    End If
    Dim MusicApi  'As HSMusicAPI
    Try
        MusicApi = pi.GetMusicAPI(Parm) ' specify here which Zone you want to use. You can use ZoneName or ZoneIndex
        GetZoneSource = MusicApi.LinkedZoneSource()
    Catch ex As Exception
        hs.writelog("Sonos Controller Script", "Error in GetZoneSource with error = " & Ex.Message)
    End Try
End Function


Function GetZoneDestination(ByVal parm) As String
    ' Example to use in HST Text field [$SCRIPT=&hs.runex ("SonosController.vb" , "GetZoneDestination" , "Master Bedroom" )]
    Dim pi As Object ' As HSPI_SONOSCONTROLLER.HSPI
    pi = hs.Plugin("SONOSCONTROLLER")
    If pi Is Nothing Then
        hs.writelog("Sonos Controller Script", "empty")
    End If
    Dim MusicApi  'As HSMusicAPI
    Try
        MusicApi = pi.GetMusicAPI(Parm) ' specify here which Zone you want to use. You can use ZoneName or ZoneIndex
        GetZoneDestination = MusicApi.LinkedZoneDestination()
    Catch ex As Exception
        hs.writelog("Sonos Controller Script", "Error in GetZoneDestination with error = " & Ex.Message)
    End Try
End Function


Function GetZoneCount(ByVal parm) As Integer
    ' Example to use in HST Text field [$SCRIPT=&hs.runex ("SonosController.vb" , "GetZoneCount" , "" )]
    Dim pi As Object ' As HSPI_SONOSCONTROLLER.HSPI
    pi = hs.Plugin("SONOSCONTROLLER")
    If pi Is Nothing Then
        hs.writelog("Sonos Controller Script", "empty")
    End If
    Dim MultiZoneAPI  'As HSMultizoneAPI
    Try
        MultiZoneAPI = pi.GetMultiZoneAPI() ' Get the MultiZoneAPI, there is only one so no parameters necessary
        GetZoneCount = MultiZoneAPI.GetZoneCount()
    Catch ex As Exception
        hs.writelog("Sonos Controller Script", "Error in GetZoneCount with error = " & Ex.Message)
    End Try
End Function


Sub AllZonesOn(ByVal parm)
    ' Example to use in HST Text field [$SCRIPT=&hs.runex "SonosController.vb" , "AllZonesOn" , "" ]
    Dim pi As Object ' As HSPI_SONOSCONTROLLER.HSPI
    pi = hs.Plugin("SONOSCONTROLLER")
    If pi Is Nothing Then
        hs.writelog("Sonos Controller Script", "empty")
    End If
    Dim MultiZoneAPI  'As HSMultizoneAPI
    Try
        MultiZoneAPI = pi.GetMultiZoneAPI() ' Get the MultiZoneAPI, there is only one so no parameters necessary
        MultiZoneAPI.AllZonesOn()
    Catch ex As Exception
        hs.writelog("Sonos Controller Script", "Error in AllZonesOn with error = " & Ex.Message)
    End Try
End Sub


Sub AllZonesOff(ByVal parm)
    ' Example to use in HST Text field [$SCRIPT=&hs.runex "SonosController.vb" , "AllZonesOff" , "" ]
    Dim pi As Object ' As HSPI_SONOSCONTROLLER.HSPI
    pi = hs.Plugin("SONOSCONTROLLER")
    If pi Is Nothing Then
        hs.writelog("Sonos Controller Script", "empty")
    End If
    Dim MultiZoneAPI  'As HSMusicAPI
    Try
        MultiZoneAPI = pi.GetMultiZoneAPI() ' Get the MultiZoneAPI, there is only one so no parameters necessary
        MultiZoneAPI.AllZonesOff()
    Catch ex As Exception
        hs.writelog("Sonos Controller Script", "Error in AllZonesOff with error = " & Ex.Message)
    End Try
End Sub


Function GetZoneCTLyrics(ByVal parm) As String
    ' Example to use in HST Text field [$SCRIPT=&hs.runex ("SonosController.vb" , "GetZoneCTLyrics" , "Master Bedroom" )]
    Dim pi As Object ' As HSPI_SONOSCONTROLLER.HSPI
    pi = hs.Plugin("SONOSCONTROLLER")
    If pi Is Nothing Then
        hs.writelog("Sonos Controller Script", "empty")
    End If
    Dim MusicApi  'As HSMusicAPI
    Dim ZoneSourceInfo As String
    Try
        MusicApi = pi.GetMusicAPI(Parm) ' specify here which Zone you want to use. You can use ZoneName or ZoneIndex
        ZoneSourceInfo = MusicApi.CTLyrics
        Dim MidScorePosition As Integer
        MidScorePosition = Instr(ZoneSourceInfo, "-")
        If MidScorePosition <> 0 Then
            ' remove the pre-part
            ZoneSourceInfo = ZoneSourceInfo.remove(0, MidScorePosition)
            ZoneSourceInfo = ZoneSourceInfo.Trim()
            MidScorePosition = Instr(ZoneSourceInfo, "-")
            If MidScorePosition <> 0 Then
                ' the first remove was the linked info, now is the radio station info
                ZoneSourceInfo = ZoneSourceInfo.remove(0, MidScorePosition)
                ZoneSourceInfo = ZoneSourceInfo.Trim()
            End If
        End If
        GetZoneCTLyrics = ZoneSourceInfo
    Catch ex As Exception
        hs.writelog("Sonos Controller Script", "Error in GetZoneCTLyrics with error = " & Ex.Message)
    End Try
End Function

Sub LinkAZoneToAudioInput(ByVal SourceZoneName As String, ByVal DestinationZoneName As String)
    Dim pi As Object
    pi = hs.Plugin("SONOSCONTROLLER")
    If pi Is Nothing Then
        hs.writelog("Script", "empty")
        Exit Sub
    End If
    Dim MusicApi
    Try
        MusicApi = pi.GetMusicAPI(DestinationZoneName) ' specify here which Zone you want to use as stream the input to
    Catch ex As Exception
        hs.writelog("Script", "Music API not found")
        Exit Sub
    End Try
    Try
        MusicApi.PlayLineInput(SourceZoneName) ' specify here which Zone you want to use the audio input from
    Catch ex As Exception
        hs.writelog("Script", "Error in calling PlayLineInput with error = " & ex.message)
        Exit Sub
    End Try
End Sub

Sub LinkAGroupOfZones(ByVal LinkGroupName As String, ByVal SourceZoneName As String, ByVal DestinationZonesString As String)
    Dim pi As Object
    'Example Dim LinkgroupName As String = "Test"
    'Example Dim SourceZoneName As String = "Family Room"
    'Example Dim DestinationZonesString As String = "Master Bedroom;20;1|Kitchen;20;1"
    pi = hs.Plugin("SONOSCONTROLLER")
    If pi Is Nothing Then
        hs.writelog("Script", "empty")
        Exit Sub
    End If
    Try
        pi.SetLinkgroupZoneSource(LinkgroupName, SourceZoneName) ' store the zone source info
    Catch ex As Exception
        hs.writelog("Script", "Error in calling SetLinkgroupZoneSource with error = " & ex.message)
        Exit Sub
    End Try
    Try
        pi.SetLinkgroupZoneDestination(LinkgroupName, DestinationZonesString) ' store the zone destination info
    Catch ex As Exception
        hs.writelog("Script", "Error in calling DestinationZonesString with error = " & ex.message)
        Exit Sub
    End Try
    Try
        pi.HandleLinking(LinkgroupName, True) ' Link!
    Catch ex As Exception
        hs.writelog("Script", "Error in calling HandleLinking with error = " & ex.message)
        Exit Sub
    End Try
End Sub

Sub UnLinkAGroupOfZones(ByVal LinkGroupName As String)
    Dim pi As Object
    pi = hs.Plugin("SONOSCONTROLLER")
    If pi Is Nothing Then
        hs.writelog("Script", "empty")
        Exit Sub
    End If
    Try
        pi.HandleLinking(LinkgroupName, False) ' UnLink!
    Catch ex As Exception
        hs.writelog("Script", "Error in calling HandleLinking with error = " & ex.message)
        Exit Sub
    End Try
End Sub


Sub SaveAllZones(ByVal parm)
    ' Example to use in HST Text field [$SCRIPT=&hs.runex "SonosController.vb" , "SaveAllZones" , "" ]
    Dim pi As Object ' As HSPI_SONOSCONTROLLER.HSPI
    pi = hs.Plugin("SONOSCONTROLLER")
    If pi Is Nothing Then
        hs.writelog("Sonos Controller Script", "empty")
    End If
    Try
        pi.SaveAllPlayersState()
    Catch ex As Exception
        hs.writelog("Sonos Controller Script", "Error in AllZonesOn with error = " & Ex.Message)
    End Try
End Sub


Sub RestoreAllZones(ByVal parm)
    ' Example to use in HST Text field [$SCRIPT=&hs.runex "SonosController.vb" , "RestoreAllZones" , "" ]
    Dim pi As Object ' As HSPI_SONOSCONTROLLER.HSPI
    pi = hs.Plugin("SONOSCONTROLLER")
    If pi Is Nothing Then
        hs.writelog("Sonos Controller Script", "empty")
    End If
    Try
        pi.RestoreAllPlayersState()
    Catch ex As Exception
        hs.writelog("Sonos Controller Script", "Error in AllZonesOff with error = " & Ex.Message)
    End Try
End Sub












