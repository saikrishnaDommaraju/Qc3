Public Class ClsDrawShtINSTProp

    Public DrawDocsCol As New Collection
    Public GenPropCol As New Collection
    Public PDMpropCol As New Collection
    Public ItemNumberList As New Collection
    Public IPtable As Object
    Public FlagNoteList As New Collection
    Public XlistCount As New Collection

    Public PartNumSht1 As ClsText
    Public PartNumSht2 As ClsText
    Public PartNumsZones As String
    Public PartNumExtra As ClsText
    Public PartNumExtraBol As Boolean
    Public PartNumExtraErrMsg As String
    Public PartNumMatchBol As Boolean
    Public PartNumMatchErrMsg As String

    Public GenProps As clsDrawShtGENProp
End Class
