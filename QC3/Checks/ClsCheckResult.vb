Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace QC3.Checks
    Class ClsCheckResult
        Public check_ID As String
        Public descr As String
        Public passVal As String
        Public actualValue As String
        Public measured As String
        Public msg As String
        Public isPass As Boolean

        Public ReadOnly Property CheckID As String
            Get
                Return check_ID
            End Get
        End Property

        Public ReadOnly Property CheckDescription As String
            Get
                Return descr
            End Get
        End Property

        Public ReadOnly Property PassValue As String
            Get
                Return passVal
            End Get
        End Property

        Public ReadOnly Property CalculatedValue As String
            Get
                Return actualValue
            End Get
        End Property

        Public ReadOnly Property MeasuredValue As String
            Get
                Return measured
            End Get
        End Property

        Public ReadOnly Property Result As String
            Get
                Return msg
            End Get
        End Property
    End Class
End Namespace