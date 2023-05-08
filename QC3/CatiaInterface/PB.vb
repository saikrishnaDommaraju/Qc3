Imports INFITF
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome

Public Class PB

    Public oStrutPage As Object
    'Public oIEX As Chrome
    Public showState As CatVisPropertyShow

    Public Sub WaitTillPageGetsLoaded(ByVal oIEX As Object)

        'Do While ((oIEX.readyState <> 4) Or (oIEX.Busy = True))
        '    DoEvents
        'Loop

        'Do Until oIEX.Busy = False
        '    DoEvents
        'Loop

        'Do Until (oIEX.readyState = READYSTATE_COMPLETE)
        '    DoEvents
        'Loop


        'Dim PauseTime
        'Dim start
        'PauseTime = 4.0#  ' Set duration.
        'start = Timer ' Set start time.
        'Do While Timer < start + PauseTime
        '    DoEvents ' Yield to other processes.
        'Loop

    End Sub

End Class
