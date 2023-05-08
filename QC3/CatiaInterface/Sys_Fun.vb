Public Class Sys_Fun

    Dim ComputerNameLen As Long, Result As Long
    Public Function GetDomainName() As String
        Try
            Dim buffer As String
            Dim size As Long
            size = 255
            buffer = Space(size)
            GetComputerNameEx(6, buffer, size)
            GetDomainName = Left$(buffer, size)
        Catch ex As Exception

        End Try
    End Function

    Private Sub GetComputerNameEx(v As Integer, buffer As String, size As Long)
        Throw New NotImplementedException()
    End Sub

    Public Function GetUserName()
        GetUserName = Environ$("username")
    End Function

    Public Function FolderFileCheck() As Boolean
        Try
            Dim chkfiles As String

            chkfiles = Dir("C:\CADESQC\TEMPLATES\QC_TEMPLATE_V1.0.xlsm")
            If chkfiles = "" Then
                FolderFileCheck = True
                Exit Function
            Else
                FolderFileCheck = False
            End If

            chkfiles = Dir("C:\CADESQC\REPORTS", vbDirectory)
            If chkfiles = "" Then
                FolderFileCheck = True
                Exit Function
            Else
                FolderFileCheck = False
            End If
        Catch ex As Exception

        End Try
    End Function

    'EXCEL REPORT CODE STARTS HERE
    Public Function FillSingleCheckReport()

        Dim myExcel As Object
        Dim WorkBook As Object
        Dim WorkSheet As Object
        Dim FileLocation As String
        Dim tempFileName As String
        Dim I As Integer
        Dim myCP As clsCHECKPoint
        If CheckList.CountList >= 1 Then
            On Error Resume Next
            FileLocation = "C:\CADESQC\TEMPLATES\QC_TEMPLATE_V1.0.xlsm"
            myExcel = CreateObject("Excel.Application")
            myExcel.Visible = False
            myExcel.displayalerts = False
            WorkBook = myExcel.Workbooks.Open(FileLocation, ReadOnly:=True, Password:="cades@airbus")
            WorkSheet = WorkBook.Worksheets.Item(1)

            'Write the Heading
            With WorkSheet
                .cells(1, 1) = "CADES (PAG) Quality Check Report -"
                .cells(3, 3) = CheckList.DrawingNumber
                .cells(4, 3) = CheckList.DrawingName
                .cells(5, 3) = CheckList.DrawingState
                .cells(6, 3) = CheckList.DrawingVersion
                .cells(7, 3) = UCase(GetUserName)
                .cells(8, 3) = Format(Now(), "yyyy-mm-dd <> hh:mm:ss")
            End With

            'Write all the Values
            For I = 1 To CheckList.CountList
                myCP = CheckList.GetCheckPoint(I)
                WorkSheet.cells(I + 9, 1) = I
                WorkSheet.cells(I + 9, 2) = myCP.Discription
                WorkSheet.cells(I + 9, 3) = myCP.CheckOKorKO
                WorkSheet.cells(I + 9, 4) = myCP.Comment
            Next I
            tempFileName = "C:\CADESQC\REPORTS\CADES_QC_REPORT_" & CheckList.DrawingNumber & "_" & Format(Now(), "yyyy-mm-dd_hh-mm-ss") & ".pdf"
            WorkSheet.ExportAsFixedFormat(Type:=0, FileName:=tempFileName, IgnorePrintAreas:=False, OpenAfterPublish:=OpenReport)
            myExcel.Workbooks.Close
            myExcel.Quit
            myExcel.displayalerts = True
        End If
        myExcel = Nothing
    End Function

    Public Function GetApprovedList()
        Dim myExcel As Object
        Dim WorkBook As Object
        Dim WorkSheet As Object
        Dim FileLocation As String
        Dim I As Integer

        On Error Resume Next

        FileLocation = "C:\CADESQC\TEMPLATES\QC_TEMPLATE_V1.0.xlsm"
        myExcel = CreateObject("Excel.Application")
        myExcel.Visible = False

        WorkBook = myExcel.Workbooks.Open(FileLocation, ReadOnly:=True, Password:="cades@airbus")
        WorkSheet = WorkBook.Worksheets("SECURITY")

        numOfDomains = WorkSheet.cells(1, 1).Value
        numOfUsers = WorkSheet.cells(1, 2).Value
        ReDim myDomains(numOfDomains)
        ReDim myUsers(numOfUsers)

        For I = 1 To numOfDomains
            myDomains(I) = WorkSheet.cells(I + 1, 1)
        Next I
        For I = 1 To numOfUsers
            myUsers(I) = WorkSheet.cells(I + 1, 2)
        Next I

        myExcel.Workbooks.Close
        myExcel.Quit
        myExcel = Nothing
    End Function


    'To Check the Authourization
    Public Sub CheckAuthourization()
        Try
            Dim domainName As String
            Dim userName As String
            Dim I As Integer
            Dim toExit As Boolean

            'Check Security
            GetApprovedList()

            domainName = GetDomainName()
            userName = GetUserName()

            toExit = True
            For I = 1 To numOfDomains
                If LCase(domainName) = LCase(myDomains(I)) Then
                    toExit = False
                    Exit For
                End If
            Next I
            If toExit = False Then
                toExit = True
                For I = 1 To numOfUsers
                    If LCase(userName) = LCase(myUsers(I)) Then
                        toExit = False
                        Exit For
                    End If
                Next I
            End If

            If toExit = True Then
                MsgBox("You are Not Authorized to Use this Program...", vbCritical, "Unauthorized Use Detected...")
                Exit Sub
            End If
        Catch ex As Exception

        End Try
    End Sub
End Class
