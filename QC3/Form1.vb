Public Class Form1


    Public _TypeOfMaterial As Integer   '1=Composite; 2=Metalic ; 3=Sheet Metal
    Public _TypeOfSection As Integer   '1=Section-13; 2=Section-16
    Public _TypeOfProgram As Integer   '1=Section-13; 2=Section-16
    Public _TypeOfDrawing As Integer    '1=General; 2=SinglePart; 3=Brkt-Instl; 4=Primary-Instl; 5=S16-18-BrktInstal

    Private Sub btn_Cancel_Click(sender As Object, e As EventArgs) Handles btn_Cancel.Click
        iExit = True
        Me.Close()
    End Sub

    Private Sub btn_run_Click(sender As Object, e As EventArgs) Handles btn_run.Click
        Try
            FormDetails()
            Cades_Dwg_Qchecker.TypeOfMaterial = _TypeOfMaterial '1=Composite; 2=Metalic ; 3=Sheet Metal
            Cades_Dwg_Qchecker.TypeOfSection = _TypeOfSection '1=Section-13; 2=Section-16
            Cades_Dwg_Qchecker.TypeOfProgram = _TypeOfProgram  '1=Section-13; 2=Section-16
            Cades_Dwg_Qchecker.TypeOfDrawing = _TypeOfDrawing
            Cades_Dwg_Qchecker._singlepart = rb_single.Checked
            Main()

        Catch ex As Exception
            frmStart.Close()
            _log.Fatal(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name + "::" +
                                 System.Reflection.MethodInfo.GetCurrentMethod.Name + "()", ex)
        End Try

    End Sub

    Public Sub FormDetails()
        Try
            iExit = False
            initializeLog()
            If iExit = True Then Exit Sub

            'Type of Material
            If rb_Composite.Checked = True Then
                _TypeOfMaterial = 1
            ElseIf rb_metallic.Checked = True Then
                _TypeOfMaterial = 2
            Else
                _TypeOfMaterial = 3
            End If

            'Type of Drawing
            If rb_generalsheetCheck.Checked = True Then
                _TypeOfDrawing = 1
            ElseIf rb_single.Checked = True Then
                _TypeOfDrawing = 2
            ElseIf rb_Bracketinstallation.Checked = True Then
                _TypeOfDrawing = 3
                'ElseIf Me.optPriInstl = True Then
                '    TypeOfDrawing = 4
            End If

            ''Type of Section
            If rb_Section1314.Checked = True Then
                _TypeOfSection = 1
            Else
                _TypeOfSection = 2
            End If

            'Check for Type of Drawing
            If TypeOfSection = 2 Then
                If TypeOfDrawing = 3 Then
                    _TypeOfDrawing = 5
                End If
            End If

            'Open Report Status
            If chkReport.Checked = True Then
                OpenReport = True
            Else
                OpenReport = False
            End If

            'Check Program(-900 or -1000)
            If rb_A350_900.Checked = True Then
                _TypeOfProgram = 1
            Else
                _TypeOfProgram = 2
            End If

            iExit = False
            Me.Hide()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Public Sub initializeLog()
        _dev = True
        Dim strConfig As String = AppDomain.CurrentDomain.BaseDirectory
        log4net.Config.XmlConfigurator.Configure(New System.IO.FileInfo(strConfig + "\config\Log4Net.config"))
        If _dev Then
            _log = log4net.LogManager.GetLogger("Dev")
        Else
            _log = log4net.LogManager.GetLogger("Prod")
        End If
    End Sub
End Class
