Imports System.Web.UI.HtmlControls
Imports System.Windows.Forms


Public Class ClsPDMLink3D

    Dim PDM As New ClsPDMLink
    Public PARTProTreatCode As String
    Public PARTDomesticTitle As String
    Public PARTName As String
    Public PARTManuIndx As String
    Public PARTIssIndx As String
    Public PARTWeight As Double
    Public PartMaterial As String
    Public PARTMatQty As String
    Public PARTLen As Double
    Public PARTWid As Double
    Public PARTThk As Double
    Public PARTItemZone As String
    Public PARTMarking As String
    Public PARTfc As String
    Public PARTid As String
    Public PARTMatStruct As String

    'For CheckList
    Public PART_Number As String
    Public PART_State As String
    Public PART_Version As String


    'Main Function for 3D
    Public Function GetPartProperties(ByVal HTZnum As String) As Boolean

        'PDM.StartPDMLink

        'If PDM.SearchHTZ(HTZnum) Then
        '    GetPartProperties = True
        '    PDM.GetLatestVersionProperties
        '    PDM.SwapSRPtoDRWSETProperties
        '    PART_Number = HTZnum
        '    PART_State = PDM.SRPstate
        '    PART_Version = PDM.SRPversion

        '    Call GetGeneralData()
        'Else
        '    GetPartProperties = False
        '    Exit Function
        'End If

        'Call GetPartWeight                 ' WIP  ReleasedByNextVersion
        'Call GetPartMaterialProperties     ' WIP  ReleasedByNextVersion

        'PDM.ClosePDMLink

    End Function

    'Get Part Related Data From General Sheet
    Public Function GetGeneralData()
        'Dim HtResultDoc As HtmlDocument
        'Dim trElems As HtmlElementCollection
        'Dim tblRow As HtmlTableRow
        'Dim idone As Boolean

        'HtResultDoc = PDM.IEX.Document

        'On Error Resume Next

        'trElems = HtResultDoc.All.GetElementsByName("tr")
        ''HtResultDoc.All.tags("tr")
        'idone = False
        'For Each tblRow In trElems

        '    ' If tblRow.className = "tableRowSelSSCI" Or tblRow.className = "tableRowNSelSSCI" Then
        '    ' Debug.Print tblRow.innerText
        '    If tblRow.innerText Like "*Protective Treatment Code:*" Then
        '        PARTProTreatCode = Trim(tblRow.cells.Item(8).innerText)
        '    End If
        '    If tblRow.innerText Like "*Domestic Title:*" Then
        '        PARTDomesticTitle = Trim(tblRow.cells.Item(4 + 1).innerText)
        '    End If
        '    If tblRow.innerText Like "*Manufacturing Index:*" Then
        '        PARTManuIndx = Trim(tblRow.cells.Item(5).innerText)
        '    End If
        '    If tblRow.innerText Like "*Item Zone:*" Then
        '        PARTItemZone = Trim(tblRow.cells.Item(5).innerText)
        '    End If
        '    If tblRow.innerText Like "*Marking:*" Then
        '        PARTMarking = Trim(tblRow.cells.Item(5).innerText)

        '    End If
        '    If tblRow.innerText Like "*Functional Class:*" Then
        '        PARTfc = Trim(tblRow.cells.Item(2).innerText)
        '    End If

        '    If tblRow.innerText Like "*Material Structure:*" Then
        '        PARTMatStruct = Trim(tblRow.cells.Item(5).innerText)
        '    End If
        '    If tblRow.innerText Like "*Identifiable Part:*" Then
        '        PARTid = Trim(tblRow.cells.Item(5).innerText)
        '    End If
        '    If tblRow.innerText Like "*Revision (Issue Index):*" Then
        '        PARTIssIndx = Trim(tblRow.cells.Item(2).innerText)
        '    End If

        '    ' End If

        '    If idone = False Then
        '        If tblRow.innerText Like "*English Title:*" Then
        '            PARTName = Trim(tblRow.cells.Item(4 + 1).innerText)
        '            idone = True
        '        End If
        '    End If
        'Next



    End Function

    'Get Part Weight
    Public Function GetPartWeight()

        'Dim HtResultDoc As HtmlDocument
        'Dim trElems As IHTMLElementCollection
        'Dim tblRow As IHTMLTableRow

        ''GoTo Weight Sheet
        'PDM.ClickByID("infoPageinfoPanelID__infoPage_myTab_object_weightAndBalanceTab")

        'HtResultDoc = PDM.IEX.Document
        'On Error Resume Next
        'trElems = HtResultDoc.All.tags("tr")
        'For Each tblRow In trElems
        '    If tblRow.cells.Length = 2 Then
        '        If tblRow.className = "tableRowSelSSCI" Or tblRow.className = "tableRowNSelSSCI" Then
        '            If tblRow.cells.Item(0).innerText Like "*Manual*" Then
        '                PARTWeight = Trim(tblRow.cells.Item(1).innerText)
        '                Exit Function
        '            End If
        '        End If
        '    End If
        'Next

    End Function

    'Get Part Material Properties
    Public Function GetPartMaterialProperties()
        'Try
        '    Dim HtResultDoc As HtmlDocument
        '    Dim trElems As IHTMLElementCollection
        '    Dim tblRow As IHTMLTableRow
        '    Dim iRead As Boolean

        '    'GoTo Weight Sheet
        '    PDM.ClickByID("ObjectPropNavBar:TextLink")

        '    HtResultDoc = PDM.IEX.Document
        '    'On Error Resume Next
        '    trElems = HtResultDoc.All.tags("tr")
        '    iRead = False
        '    For Each tblRow In trElems
        '        If tblRow.cells.Length = 14 Then
        '            If iRead = True Then
        '                PartMaterial = Trim(tblRow.cells.Item(0).innerText)
        '                PARTMatQty = Trim(tblRow.cells.Item(6).innerText)
        '                PARTWid = Trim(tblRow.cells.Item(9).innerText)
        '                PARTLen = Trim(tblRow.cells.Item(10).innerText)
        '                PARTThk = Trim(tblRow.cells.Item(12).innerText)
        '                Exit Function
        '            End If
        '            'To Skip the Heading
        '            If tblRow.cells.Item(0).innerText Like "*Number*" Then
        '                iRead = True
        '            End If
        '        End If
        '    Next
        'Catch ex As Exception

        'End Try

    End Function

    'Get Properties for CheckList
    Public Function GetPartName()

    End Function




End Class
