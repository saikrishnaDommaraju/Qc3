
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace QC3.Checks
    Class Clschecks
        Public check_ID As String
        Public desc As String
        Public passVal As String
        Public check_name As String
        Public disppassVal As String

        Public ReadOnly Property CheckPoint As String
            Get
                Return check_ID
            End Get
        End Property

        Public ReadOnly Property CheckStatus As String
            Get
                Return desc
            End Get
        End Property

        Public ReadOnly Property Remarks As String
            Get
                Return disppassVal
            End Get
        End Property
    End Class
End Namespace