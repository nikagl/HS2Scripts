Public Function WaitForTTS(ByVal parm)
    Dim pi As Object 'As HSPI_SONOSCONTROLLER.HSPI
    pi = hs.Plugin("SONOSCONTROLLER")
    If pi Is Nothing Then
        hs.writelog("Script", "empty")
        Exit Function
    End If
    Dim MaxWait As Integer
    While Not pi.LinkgroupState
        MaxWait = MaxWait + 1
        hs.WaitSecs(0.5)
        If Maxwait > 120 Then ' Max 1 minute
            Exit Function
        End If
    End While
    hs.writelog("SONOSCONTOLLER", "Script: TTSON")
End Function

















