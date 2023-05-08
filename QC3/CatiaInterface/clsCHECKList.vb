Public Class clsCHECKList

    Public List As New Collection
    Public CountList As Integer
    Public DrawingNumber As String
    Public DrawingName As String
    Public DrawingState As String
    Public DrawingVersion As String

    Public Sub AddCompareCheckPoint(ByVal input1 As Object,
                    ByVal input2 As Object,
                    ByVal Discribe As String,
                    ByVal OKcomment As String,
                    ByVal NOKcomment As String)

        Dim chkPnt As New clsCHECKPoint

        chkPnt.SetQCByComparision(input1, input2, Discribe, OKcomment, NOKcomment)
        List.Add(chkPnt)
        CountList = List.Count
    End Sub

    Public Sub AddWarningCompareCheckPoint(ByVal input1 As Object,
                    ByVal input2 As Object,
                    ByVal Discribe As String,
                    ByVal OKcomment As String,
                    ByVal NOKcomment As String)

        Dim chkPnt As New clsCHECKPoint

        chkPnt.SetQCByWarningComparision(input1, input2, Discribe, OKcomment, NOKcomment)
        List.Add(chkPnt)
        CountList = List.Count
    End Sub

    Public Sub AddBooleanCheckPoint(ByVal Check As Boolean,
                    ByVal Discribe As String,
                    ByVal OKcomment As String,
                    ByVal NOKcomment As String)

        Dim chkPnt As New clsCHECKPoint

        chkPnt.SetQCByBooleanCheck(Check, Discribe, OKcomment, NOKcomment)
        List.Add(chkPnt)
        CountList = List.Count
    End Sub

    Public Sub AddWarningBooleanCheckPoint(ByVal Check As Boolean,
                    ByVal Discribe As String,
                    ByVal OKcomment As String,
                    ByVal NOKcomment As String)

        Dim chkPnt As New clsCHECKPoint

        chkPnt.SetQCByWarningBooleanCheck(Check, Discribe, OKcomment, NOKcomment)
        List.Add(chkPnt)
        CountList = List.Count
    End Sub

    Public Sub AddWarning2BooleanCheckPoint(ByVal Check1 As Boolean,
                    ByVal Check2 As Boolean,
                    ByVal Discribe As String,
                    ByVal OKcomment As String,
                    ByVal WAcomment As String,
                    ByVal NOKcomment As String)

        Dim chkPnt As New clsCHECKPoint

        chkPnt.SetQCByWarning2BooleanCheck(Check1, Check2, Discribe, OKcomment, WAcomment, NOKcomment)
        List.Add(chkPnt)
        CountList = List.Count
    End Sub

    Public Function GetCheckPoint(ByVal indx As Integer) As clsCHECKPoint
        If indx <= CountList Then
            GetCheckPoint = List.Item(indx)
        Else
            Exit Function
        End If
    End Function
End Class
