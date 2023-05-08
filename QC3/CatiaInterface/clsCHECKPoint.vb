Public Class clsCHECKPoint

    Public CheckOKorKO As String
    Public Comment As String
    Public Discription As String

    Public Sub SetQCByComparision(ByVal input1 As Object,
                ByVal input2 As Object,
                ByVal Discribe As String,
                ByVal OKcomment As String,
                ByVal NOKcomment As String)

        Discription = Discribe
        If input1 = input2 Then
            CheckOKorKO = "OK"
            Comment = OKcomment
        Else
            CheckOKorKO = "KO"
            Comment = NOKcomment & " ( " & input1 & " <AND> " & input2 & " ) "
        End If
    End Sub

    Public Sub SetQCByWarningComparision(ByVal input1 As Object,
                ByVal input2 As Object,
                ByVal Discribe As String,
                ByVal OKcomment As String,
                ByVal NOKcomment As String)

        Discription = Discribe
        If input1 = input2 Then
            CheckOKorKO = "OK"
            Comment = OKcomment
        Else
            CheckOKorKO = "WA"
            Comment = NOKcomment & " ( " & input1 & " <AND> " & input2 & " ) "
        End If
    End Sub

    Public Sub SetQCByBooleanCheck(ByVal Check As Boolean,
                ByVal Discribe As String,
                ByVal OKcomment As String,
                ByVal NOKcomment As String)

        Discription = Discribe
        If Check Then
            CheckOKorKO = "OK"
            Comment = OKcomment
        Else
            CheckOKorKO = "KO"
            Comment = NOKcomment
        End If
    End Sub

    Public Sub SetQCByWarningBooleanCheck(ByVal Check As Boolean,
                ByVal Discribe As String,
                ByVal OKcomment As String,
                ByVal NOKcomment As String)

        Discription = Discribe
        If Check Then
            CheckOKorKO = "OK"
            Comment = OKcomment
        Else
            CheckOKorKO = "WA"
            Comment = NOKcomment
        End If
    End Sub

    Public Sub SetQCByWarning2BooleanCheck(ByVal Check1 As Boolean,
                ByVal Check2 As Boolean,
                ByVal Discribe As String,
                ByVal OKcomment As String,
                ByVal WAcomment As String,
                ByVal NOKcomment As String)

        Discription = Discribe
        If Check1 = True And Check2 = True Then
            CheckOKorKO = "OK"
            Comment = OKcomment
        ElseIf Check1 = True And Check2 = False Then
            CheckOKorKO = "WA"
            Comment = WAcomment
        Else
            CheckOKorKO = "KO"
            Comment = NOKcomment
        End If
    End Sub


End Class
