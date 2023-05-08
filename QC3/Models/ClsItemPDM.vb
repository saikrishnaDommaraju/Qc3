Public Class ClsItemPDM

    Public PartNumber As String
    Public Quantity As Integer          'Set by Function
    Public ManagedByCad As Boolean
    Public Context As String            'To Check if it is Standard Part Or Not
    Public ItemNumber As String
    Public ItemZone As String
    Public AlternateCode As String      'Set by Function
    Public isPrimary As Boolean         'Set by Function
    Public IPnumber1 As String          'Installation Prop
    Public IPnumber2 As String          'Installation Prop

    'PUBLIC: Set Quantity

    Public Sub SetQuantity(QuantityStr As String)
        Dim var As Object

        If QuantityStr <> "" Then
            var = Split(QuantityStr, " ")
            Quantity = CInt(var(0))
        Else
            Quantity = 0
        End If
    End Sub

    'PUBLIC: Set Alternate Code

    Public Sub SetPrimary(AlternateCodeStr As String)
        If AlternateCodeStr <> "" Then
            AlternateCode = AlternateCodeStr
            If Right(AlternateCodeStr, 1) = 1 Then
                isPrimary = True
            Else
                isPrimary = False
            End If
        Else
            AlternateCode = ""
        End If
    End Sub


End Class
