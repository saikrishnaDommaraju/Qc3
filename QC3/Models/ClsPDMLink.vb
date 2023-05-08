Imports OpenQA
Imports System.Collections
Imports System.Web.UI.HtmlControls
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Text
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome
Imports System.Threading

Public Class ClsPDMLink

    Private PDMLinkPath As String
    Public IEX As ChromeDriver

    'Public Global Variables
    'SRP => Search Result Page
    Public SRPversion As String
    Public SRPstate As String
    Public SRPlastupdated As String
    'SHT => SHT Properties, DRW => Drawing
    Public DRWSHTname As String
    Public DRWSHTnatco As String
    Public DRWSHTissueIndex As String
    Public DRWSHTcancellationIndex As String
    Public DRWSHTdrawingSize As String
    Public DRWSHTversion As String
    Public DRWSHTstate As String
    Public DRWSHTlastupdated As String
    'SET => DRAWING SET Properties, DRW => Drawing
    Public DRWSETname As String
    Public DRWSETnatco As String
    Public DRWSETissueIndex As String
    Public DRWSETcancellationIndex As String
    Public DRWSETversion As String
    Public DRWSETstate As String
    Public DRWSETlastupdated As String
    Public DRWSETenglishTitle As String
    Public DRWSETdomesticTitle As String
    Public DRWSETremarksList As Object
    'SET => PART Properties, PART => Part Document Properties
    Public PARTzone As String
    Public PARTidentification As String
    'SET => ASSY Properties, ASSY => Assembly Document Properties
    Public ASSYbomItems As Collection
    'SET => INST Properties, INST => Installation Document Properties
    Public INSTbomItems As Collection
    Public chromedriver As ChromeDriver

    'Parameter for Checking FN and GN is Used in Remarks
    'Public DRWSETremarksType As Boolean
    'Private Const WM_GETTEXTLENGTH As Integer = &HE
    'Private Const WM_GETTEXT As Integer = &HD

    'PUBLIC: Start the PDM Link
    Public Sub StartPDMLink()


        Try
            Process.Start("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe", "https://pass-ssi-a350.eu.airbus.corp:1443/D850/app/")

            Dim PDMLinkPath As String = "https://pass-ssi-a350.eu.airbus.corp:1443/D850/app/"
            'Do While wb.ReadyState <> WebBrowserReadyState.Complete
            '    Application.DoEvents()
            'Loop

            'Dim chromedriver As IWebDriver = New ChromeDriver()
            'chromedriver.Navigate().GoToUrl("http://www.google.com/")
            'chromedriver.Quit()

            'Dim webAddress As String = "http://www.google.com/"

            '          driver.get("http://www.google.com/");    

            'Thread.sleep(5000);  // Let the user actually see something!     

            'WebElement searchBox = driver.findElement(By.Name("q"));

            'searchBox.sendKeys("ChromeDriver");     

            'searchBox.submit();    

            'Thread.sleep(5000);  // Let the user actually see something!     

            'driver.quit();  
            'Process.Start(webAddress)
            'Process.Start(url)
            'Dim ass As Security.SecureString = (New System.Net.NetworkCredential("", "AXISCADES@753159")).SecurePassword
            'Dim theSecureString As secure = New NetworkCredential("", "myPass").SecurePassword
            'Process.Start(url, "CBEHEU56", ass, "")
            'Process.GetProcessesByName()
            'MsgBox("The Folder Structure is Not Set-Up Properly...")





            'Path of HTML doc inside a Frame with ID = internal_search_frame
            '' old link
            ''PDMLinkPath = "http://ssia350prod.eu.airbus.corp:8020/SSCIA350/wtcore/jsp/com/ptc/windchill/search/DcaGatewayDelegate.jsp?" ' Commented/by_Rana
            'Dim Driver As New OpenQA.Selenium.ChromeDriver
            'Driver = CreateObject("Selenium.ChromeDriver")
            'Driver.start
            'Driver.Get("https://www.google.com")
            'Driver.Quit
            ' Application.Wait Now + TimeValue("00:00:20")



            '    Dim chrobj As New Selenium.ChromeDriver
            '    chrobj.Get "chrome"
            '    chrobj.Window.Maximize
            '    PDMLinkPath = "https://pass-ssi-a350.eu.airbus.corp:1443/D850/app/"



            'IEX = CreateObject("ChromeExplorer.Application")
            'PB.oIEX = IEX
            'IEX.Height = 1000
            'IEX.Width = 1200
            'IEX.Left = 0
            'IEX.Top = 0
            'IEX.Visible = True
            navigateToPage(PDMLinkPath)

        Catch ex As Exception

        End Try
    End Sub
    Public Function WaitToLoad(ByVal id As String, ByVal driver As ChromeDriver) As Boolean
        Dim i As Integer = 0

        While i < 600
            i += 1
            Thread.Sleep(100)

            Try
                Dim by As By = By.Id(id)
                driver.FindElement(by)
                Exit While
            Catch
            End Try
        End While

        If i = 600 Then
            Return False
        Else
            Return True
        End If
    End Function

    'Private Property pageready As Boolean = False

#Region "Page Loading Functions"
    'Private Sub WaitForPageLoad()
    '    AddHandler ChromeDriver.DocumentCompleted, New WebBrowserDocumentCompletedEventHandler(AddressOf PageWaiter)
    '    While Not pageready
    '        Application.DoEvents()
    '    End While
    '    pageready = False
    'End Sub

    'Private Sub PageWaiter(ByVal sender As Object, ByVal e As WebBrowserDocumentCompletedEventArgs)
    '    If whatbrowser.ReadyState = WebBrowserReadyState.Complete Then
    '        pageready = True
    '        RemoveHandler whatbrowser.DocumentCompleted, New WebBrowserDocumentCompletedEventHandler(AddressOf PageWaiter)
    '    End If
    'End Sub

#End Region

    Private Function GetCurrentUrl(mainWindowHandle As IntPtr, browserName As String, className As String, comboBox1 As Object) As String
        Throw New NotImplementedException()
    End Function


    Private Function GetBrowser(ByVal appName) As System.Diagnostics.Process
        Dim pList() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessesByName(appName)
        For Each proc As System.Diagnostics.Process In pList
            If proc.ProcessName = appName Then
                Return proc
            End If
        Next
        Return Nothing
    End Function
    'Private Shared Function FindWindowEx(parentHandle As IntPtr, childAfter As IntPtr, className As String, windowTitle As String) As IntPtr
    'End Function

    'Private Shared Function SendMessage(hWnd As IntPtr, Msg As UInteger, wParam As Integer, lParam As StringBuilder) As Integer
    'End Function
    'Private Shared Function SendMessage(hWnd As IntPtr, Msg As UInteger, wParam As Integer, lParam As Integer) As Integer
    'End Function
    'Public Shared Function GetChromeHandle() As IntPtr
    '    Dim ChromeHandle As IntPtr = Nothing
    '    Dim Allpro() As Process = Process.GetProcesses()
    '    For Each pro As Process In Allpro
    '        If pro.ProcessName = "chrome" Then
    '            ChromeHandle = pro.MainWindowHandle
    '            Exit For
    '        End If
    '    Next
    '    Return ChromeHandle
    'End Function

    'Public Shared Function getChromeUrl(winHandle As IntPtr) As String
    '    Dim browserUrl As String = Nothing
    '    Dim urlHandle As IntPtr = FindWindowEx(winHandle, IntPtr.Zero, "Chrome_AutocompleteEditView", Nothing)
    '    Const nChars As Integer = 256
    '    Dim Buff As New StringBuilder(nChars)
    '    Dim length As Integer = SendMessage(urlHandle, WM_GETTEXTLENGTH, 0, 0)
    '    If length > 0 Then
    '        SendMessage(urlHandle, WM_GETTEXT, nChars, Buff)
    '        browserUrl = Buff.ToString()

    '        Return browserUrl
    '    Else
    '        Return browserUrl
    '    End If

    'End Function

    'Private Function GetBrowser(ByVal appName) As System.Diagnostics.Process
    '    Dim pList() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessesByName(appName)
    '    For Each proc As System.Diagnostics.Process In pList
    '        'If proc.ProcessName = appName Then
    '        '    Return proc
    '        'End If
    '        Dim prsurl As String = proc.ProcessName
    '        If proc IsNot Nothing Then
    '            Dim browserName As String = "chrome"
    '            Dim className As String = "Edit"
    '            Dim s As String = GetCurrentUrl(proc.MainWindowHandle, browserName, className, "")
    '            If s <> prsurl Then
    '                'TextBox1.Text = s
    '                'ComboBox1.SelectedIndex = 0 'Window list
    '                MsgBox("sucess ")
    '            Else
    '                'Label1.Text = "Current URL is not available"
    '                MsgBox("fail")
    '            End If
    '        Else
    '            'Label1.Text = browserName & " is not available"
    '        End If
    '    Next
    '    Return Nothing
    'End Function
    Public Sub navigateToPage(ByVal PathToNavigate As String)

        Dim chromedr As ChromeDriver = New ChromeDriver()
        chromedriver = chromedr
        Dim id As String = "gloabalSearchField"
        WaitToLoad(id, chromedriver)
        'chromedriver.Navigate().GoToUrl(PathToNavigate)

        'Dim objShell As Object, objIE As Object, objShellWindows As Object
        'Dim iCount As Integer

        'objShell = CreateObject("Shell.Application")
        'objShellWindows = objShell.Windows

        'On Error Resume Next
        'For i = 0 To objShellWindows.Count - 1
        '    If InStr(1, objIE.locationurl, "https://pass-ssi-a350.eu.airbus.corp:1443/D850/app/") Then
        '        MsgBox("Found idp.feide.no Website!")
        '    End If
        'Next
        'Dim chromedriver As IWebDriver = New ChromeDriver()
        'chromedriver.Navigate().GoToUrl(PathToNavigate)

        'Do Until IEX. = False
        '    '    DoEvents
        'Loop

        'Do Until IEX.readyState = READYSTATE_COMPLETE
        '    DoEvents
        'Loop
        'Dim PauseTime
        'Dim start
        'PauseTime = 6.0#  ' Set duration.
        'start = Timer ' Set start time.
        'Do While Timer < start + PauseTime
        '    DoEvents ' Yield to other processes.
        'Loop


    End Sub


    'PRIVATE: Click Link, GoTo Latest Page
    'Modified
    Private Sub ClickSearchLink(ByVal iNavNext As String, ByVal skipStatus As Boolean)

        'Dim htdoc As HtmlDocument
        'Dim templink As HTMLLinkElement
        'Dim link As HTMLLinkElement
        'Dim I As Integer

        'iNavNext = Trim(iNavNext)
        'htdoc = IEX.Document
        'For I = 10 To htdoc.Links.Count - 1
        '    templink = htdoc.Links.Item(I)
        '    'If templink.innerText Like iNavNext Then               '>> Commented by Rana
        '    If Trim(templink.innerText) = iNavNext Then
        '        If skipStatus = False Then
        '            'Set link = HTdoc.Links.Item(i + 1)             '>> Commented by Rana
        '            link = htdoc.Links.Item(I)
        '            Exit For
        '        Else
        '            skipStatus = False
        '        End If
        '    End If
        'Next I

        'link.Click

        'Do While IEX.readyState <> 4 Or IEX.Busy = True
        '    DoEvents
        'Loop

    End Sub

    ''PUBLIC: Close the PDM Link
    Public Sub ClosePDMLink()
        ''Exit Internet Explorer
        'IEX.Quit
        'IEX = Nothing
    End Sub

    'PUBLIC: Search for a HTZ Number and Return False if Not Found

    Public Function SearchHTZ(ByVal HTZnum As String) As Boolean

        'Dim HtSearchDoc As HtmlDocument
        'Dim HtResultDoc As HtmlDocument
        'Dim tdElems As IHTMLElementCollection
        'Dim tdCell As IHTMLTableCell
        'Dim isFound As Boolean
        'Dim I As Integer

        'Call NavigateToPage(PDMLinkPath)

        'HtSearchDoc = IEX.Document
        'HtSearchDoc.GetElementById("gloabalSearchField").Value = HTZnum
        ''''HtSearchDoc.getElementById("customSearchButton").Click '''>>Commented by Rana
        'HtSearchDoc.getElementsByClassName("x-form-trigger global-search-trigger")(0).Click

        'Call WaitTillPageGetsLoaded(IEX)

        'HtResultDoc = IEX.Document
        'Dim SearchWin
        'SearchWin = HtResultDoc.GetElementById("object_search_navigation")
        'tdElems = SearchWin.all.tags("tr")
        'isFound = True
        'I = 1

        'Dim iCount
        'iCount = HtResultDoc.getElementsByClassName("tableCount").Item(0).innerText

        'If iCount = "( 0 objects )" Then
        '    IEX.Visible = False
        '    isFound = False
        'End If

        ''Call WaitTillPageGetsLoaded(IEx)
        'SearchHTZ = isFound

    End Function

    'Public: GoTo Search Results Page And Get Data MoDified
    Public Sub GetLatestVersionProperties()
        'Dim HtResultDoc As HtmlDocument
        'Dim tdElems As IHTMLElementCollection
        'Dim tbl As IHTMLTable
        'Dim tblRow As IHTMLTableRow
        'Dim tblCell As IHTMLTableCell
        'Dim NavNext As String
        'Dim skip As Boolean
        'Dim oSearchResults
        'Dim trElems As Object
        'Dim trEle As HtmlTableRow
        'Dim oColl As New Collection
        'Dim I, VerColumn, StateColumn, DDSColumn, Rownumber


        'HtResultDoc = IEX.Document
        'WaitTillPageGetsLoaded(IEX)
        'oSearchResults = HtResultDoc.GetElementById("advancedNavigation")
        'trElems = oSearchResults.all.tags("tr")
        'I = 0
        'ReDim arrVersions(I)



        ''
        ''    For Each trEle In trElems
        ''        If trEle.cells.Length >= 12 Then
        ''            Dim oTableRowCells
        ''            Set oTableRowCells = trEle.cells
        ''            If Not oTableRowCells.Item(11).innerText = "Version" Then
        ''                'If oTableRowCells.Item(9).all.tags("img").Item(0).Title = "Design Data Set" Or oTableRowCells.Item(9).all.tags("img").Item(0).Title = "Drawing Sheet" Or oTableRowCells.Item(9).all.tags("img").Item(0).Title = "Dev Detailed Part" Or oTableRowCells.Item(9).all.tags("img").Item(0).Title = "ADAP-DS" Then
        ''                If InStr(1, oTableRowCells.Item(4).all.tags("img").Item(0).outerHTML, "Design Data Set", vbTextCompare) > 1 Or InStr(1, oTableRowCells.Item(4).all.tags("img").Item(0).outerHTML, "Drawing Sheet", vbTextCompare) Or InStr(1, oTableRowCells.Item(4).all.tags("img").Item(0).outerHTML, "Dev Detailed Part", vbTextCompare) Or InStr(1, oTableRowCells.Item(4).all.tags("img").Item(0).outerHTML, "ADAP-DS", vbTextCompare) Then
        ''                    If InStr(1, oTableRowCells.Item(12).innerText, "Released", vbTextCompare) >= 1 Then
        ''                        arrVersions(i) = oTableRowCells.Item(11).innerText
        ''                        Call oColl.Add(trEle)
        ''                        i = i + 1
        ''                        ReDim Preserve arrVersions(i)
        ''                    End If
        ''                End If
        ''            End If
        ''        End If
        ''    Next

        ''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>Changing code for Released Dwgs to Wip Dwgs

        'For Each trEle In trElems
        '    Dim oTableRowCells1
        '    oTableRowCells1 = trEle.Cells
        '    For I = 0 To oTableRowCells1.Length - 1
        '        If InStr(1, oTableRowCells1.Item(I).innerText, "Version", vbTextCompare) >= 1 Then
        '            VerColumn = I
        '        End If
        '        If InStr(1, oTableRowCells1.Item(I).innerText, "State", vbTextCompare) >= 1 Then
        '            StateColumn = I
        '        End If
        '        If InStr(1, oTableRowCells1.Item(I).innerText, "Object Type Indicator", vbTextCompare) >= 1 Then
        '            DDSColumn = I
        '        End If
        '    Next
        'Next

        'For Each trEle In trElems
        '    If trEle.Cells.Length >= 12 Then
        '        Dim oTableRowCells
        '        oTableRowCells = trEle.Cells
        '        If Not oTableRowCells.Item(VerColumn).innerText = "Version" Then
        '            'If oTableRowCells.Item(9).all.tags("img").Item(0).Title = "Design Data Set" Or oTableRowCells.Item(9).all.tags("img").Item(0).Title = "Drawing Sheet" Or oTableRowCells.Item(9).all.tags("img").Item(0).Title = "Dev Detailed Part" Or oTableRowCells.Item(9).all.tags("img").Item(0).Title = "ADAP-DS" Then
        '            'If InStr(1, oTableRowCells.Item(9).all.tags("img").Item(0).outerHTML, "Design Data Set", vbTextCompare) > 1 Or InStr(1, oTableRowCells.Item(9).all.tags("img").Item(0).outerHTML, "Drawing Sheet", vbTextCompare) Or InStr(1, oTableRowCells.Item(9).all.tags("img").Item(0).outerHTML, "Dev Detailed Part", vbTextCompare) Or InStr(1, oTableRowCells.Item(9).all.tags("img").Item(0).outerHTML, "ADAP-DS", vbTextCompare) Then
        '            'If InStr(1, CStr(oTableRowCells.Item(3).innerText), CStr(Sheet1.Cells(RowNumber, 4).Value), vbTextCompare) >= 1 Then
        '            'If InStr(CStr(oTableRowCells.Item(VerColumn).innerText), CStr(Trim(Sheet1.cells(Rownumber, 4).Value))) >= 1 Then
        '            If InStr(1, oTableRowCells.Item(VerColumn).nextSibling.innerText, "Work in Progress", vbTextCompare) >= 1 Then 'Work in Progress
        '                arrVersions(I) = oTableRowCells.Item(VerColumn).innerText
        '                Call oColl.Add(trEle)
        '                I = I + 1
        '                ReDim Preserve arrVersions(I)
        '            End If
        '            'End If
        '        End If
        '    End If
        'Next

        'Call Alpha_ArraySorting(arrVersions)

        'Dim oSearchLinkobj As HtmlTableRow
        'oSearchLinkobj = getActualRowFromSearchResults(oColl, CStr(arrVersions(0)))
        'Dim oLink2
        'oLink2 = oSearchLinkobj.all.tags("A").Item(0)
        'oLink2.Click
        'Call WaitTillPageGetsLoaded(IEX)
        'tblCell = oSearchLinkobj.Cells.Item(3)
        'SRPversion = CStr(tblCell.innerText)
        'tblCell = oSearchLinkobj.Cells.Item(4)
        'SRPstate = CStr(tblCell.innerText)
        'tblCell = oSearchLinkobj.Cells.Item(13)
        'SRPlastupdated = CStr(tblCell.innerText)

        'WaitTillPageGetsLoaded(IEX)


    End Sub

    Private Function getActualRowFromSearchResults(oColl As Collection, sToFind As String)

        Dim I As Integer
        For I = 1 To oColl.Count
            Dim oRowOBJ As Object
            oRowOBJ = oColl.Item(I)
            Dim sRowOBJinnertext
            sRowOBJinnertext = oRowOBJ.innerText
            If InStr(sRowOBJinnertext, sToFind) > 0 Then
                getActualRowFromSearchResults = oRowOBJ
                Exit For
            End If
        Next

    End Function
    Public Sub Alpha_ArraySorting(arr)
        'PURPOSE: Sort an array filled with strings in reverse alphabetical order (Z-A)
        'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

        Dim X As Long, Y As Long
        Dim sTemp1 As String
        Dim sTemp2 As String

        For X = LBound(arr) To UBound(arr)
            For Y = X To UBound(arr)
                If UCase(arr(Y)) > UCase(arr(X)) Then
                    sTemp1 = arr(X)
                    sTemp2 = arr(Y)
                    arr(X) = sTemp2
                    arr(Y) = sTemp1
                End If
            Next Y
        Next X

    End Sub



    'Public Sub Sort_Through_Web()


    '    Dim HtResultDoc As HtmlDocument
    '    Dim oVersionObj As Object
    '    Dim oVersionObjName As HTMLAnchorElement
    '    Dim oListMenu1 As Object
    '    Dim oName As Object
    '    Dim oNameString As HTMLAnchorElement
    '    Dim oListMenu2 As Object


    '    HtResultDoc = IEX.Document
    '    Call WaitTillPageGetsLoaded(IEX)

    '    oVersionObj = HtResultDoc.getElementsByClassName("x-grid3-hd-inner x-grid3-hd-version")
    '    oVersionObjName = HtResultDoc.GetElementById(oVersionObj.Item(0).innerText)
    '    oVersionObjName.Click

    '    Call WaitTillPageGetsLoaded(IEX)

    '    oListMenu1 = HtResultDoc.getElementsByClassName("x-menu-list")
    '    oListMenu1.Item(0).childNodes.Item(1).Click
    '    Call WaitTillPageGetsLoaded(IEX)

    '    oName = HtResultDoc.getElementsByClassName("x-grid3-hd-inner x-grid3-hd-name")
    '    oNameString = HtResultDoc.GetElementById(oName.Item(0).innerText)
    '    oNameString.Click
    '    Call WaitTillPageGetsLoaded(IEX)

    '    oListMenu2 = HtResultDoc.getElementsByClassName("x-menu-list")
    '    oListMenu2.Item(0).childNodes.Item(0).Click
    '    Call WaitTillPageGetsLoaded(IEX)

    'End Sub





    'PUBLIC: Swap SRP Properties to DRWSHT Properties
    Public Sub SwapSRPtoDRWSHTProperties()

        DRWSHTversion = SRPversion
        DRWSHTstate = SRPstate
        DRWSHTlastupdated = SRPlastupdated

    End Sub

    'PUBLIC: Swap SRP Properties to DRWSET Properties

    Public Sub SwapSRPtoDRWSETProperties()
        DRWSETversion = SRPversion
        DRWSETstate = SRPstate
        DRWSETlastupdated = SRPlastupdated
    End Sub

    'PUBLIC: GoTo General Page and Get Data

    'Public Sub GetDrawingGeneralPageProperties()
    '    Dim HtResultDoc As HtmlDocument
    '    Dim trElems As IHTMLElementCollection
    '    Dim tblRow 'As IHTMLTableRow
    '    Dim I As Long
    '    Dim X As Boolean
    '    Dim oMainAttributes
    '    Dim oMainAttributesRows


    '    Call WaitTillPageGetsLoaded(IEX)
    '    HtResultDoc = IEX.Document

    '    On Error Resume Next
    '    I = 1
    '    trElems = HtResultDoc.All.tags("tr")
    '    oMainAttributes = HtResultDoc.GetElementById("infoPage_myTab_object_drawingSheetDetailsTab")

    '    oMainAttributesRows = oMainAttributes.all.tags("td")

    '    For tblRow = 0 To oMainAttributesRows.Length - 1


    '        If oMainAttributesRows.Item(tblRow).innerText Like "*Number:*" Then
    '            DRWSHTname = Trim(oMainAttributesRows.Item(tblRow + 1).innerText)
    '        End If

    '        If oMainAttributesRows.Item(tblRow).innerText Like "*Natco:*" Then
    '            DRWSHTnatco = Trim(oMainAttributesRows.Item(tblRow + 1).innerText)
    '        End If

    '        If oMainAttributesRows.Item(tblRow).innerText Like "*Revision (Issue Index):*" Then
    '            DRWSHTissueIndex = Trim(oMainAttributesRows.Item(tblRow + 1).innerText)
    '        End If

    '        If oMainAttributesRows.Item(tblRow).innerText Like "*Cancellation Index:*" Then
    '            DRWSHTcancellationIndex = Trim(oMainAttributesRows.Item(tblRow + 1).innerText)
    '        End If

    '        If oMainAttributesRows.Item(tblRow).innerText Like "*Drawing Size:*" Then
    '            DRWSHTdrawingSize = Trim(oMainAttributesRows.Item(tblRow + 1).innerText)
    '        End If



    '        I = I + 1
    '    Next

    '    If HtResultDoc.GetElementById("infoPageIdentityObjectIdentifier").In Like "*Number:*" Then
    '        DRWSHTname = Trim(oMainAttributesRows.Item(tblRow + 1).innerText)
    '    End If

    '    'Swap the Cancellation Index to respect Check Point Class
    '    X = DRWSHTcancellationIndex
    '    DRWSHTcancellationIndex = Not X

    'End Sub


    'PUBLIC: GoTo Drawing Set Link Page and Get Data

    Public Sub GetDrawingSetLinkProperties(ByVal HTZnumber As String)
        'If Me.SearchHTZ(HTZnumber) Then
        '    'Goto Latest Version Drawing Set
        '    'Call GetLatestVersionProperties()
        '    'Get Properties
        '    Call GetDrawingSETGeneralPageProperties()
        'End If
    End Sub

    'PUBLIC: GoTo Drawing Set Components and Get Remarks

    'Public Sub GetDrawingSetComponentProperties(ByVal HTZnumber As String)
    '    'GoTo Drawing Set Components
    '    ' ClickByID ("ObjectPropNavBar:TextLink:ViewComponents")
    '    '    ClickByID ("infoPageinfoPanelID__infoPage_myTab_object_workSetComponents")

    '    Dim htdoc As HtmlDocument
    '    Dim trElems As IHTMLElementCollection
    '    Dim link As HTMLLinkElement

    '    htdoc = oIEX.Document

    '    'Get Drawing Set Remarks

    '    GetDrawingSETRemarksProperties()

    'End Sub


    'PRIVATE: GoTo Drawing Set General Page and Get Data

    Private Sub GetDrawingSETGeneralPageProperties()

        'Dim HtResultDoc As HtmlDocument
        'Dim trElems As IHTMLElementCollection
        'Dim tblRow 'As IHTMLTableRow
        'Dim I As Long
        'Dim X As Boolean
        'Dim oMainAndAdditionalAttributes

        'HtResultDoc = IEX.Document

        'On Error Resume Next
        'I = 1
        'oMainAndAdditionalAttributes = HtResultDoc.GetElementById("infoPage_myTab_object_workSetDetailsTab")
        'trElems = oMainAndAdditionalAttributes.all.tags("td")

        'For tblRow = 0 To trElems.Length

        '    If trElems.Item(tblRow).innerText Like "*English Title:*" Then
        '        DRWSETname = Trim(trElems.Item(tblRow + 1).innerText)
        '    End If

        '    If trElems.Item(tblRow).innerText Like "*Natco:*" Then
        '        DRWSETnatco = Trim(trElems.Item(tblRow + 1).innerText)
        '    End If

        '    If trElems.Item(tblRow).innerText Like "*Issue Index:*" Then
        '        DRWSETissueIndex = Trim(trElems.Item(tblRow + 1).innerText)
        '    End If

        '    If trElems.Item(tblRow).innerText Like "*Cancellation Index:*" Then
        '        DRWSETcancellationIndex = Trim(trElems.Item(tblRow + 1).innerText)
        '    End If

        '    If trElems.Item(tblRow).innerText Like "*English Title:*" Then
        '        DRWSETenglishTitle = Trim(trElems.Item(tblRow + 1).innerText)
        '    End If

        '    If trElems.Item(tblRow).innerText Like "*Domestic Title:*" Then
        '        DRWSETdomesticTitle = Trim(trElems.Item(tblRow + 1).innerText)
        '    End If

        '    I = I + 1
        'Next


        ''Swap the Cancellation Index to Respect the Check Point Class
        'X = DRWSETcancellationIndex
        'DRWSETcancellationIndex = Not X


    End Sub

    'PUBLIC: Click a Link by its ID Name

    'Public Sub ClickByID(ByVal clickID As String)
    '    Dim htdoc As HtmlDocument
    '    Dim link As HTMLLinkElement

    '    htdoc = IEX.Document
    '    link = htdoc.GetElementById(clickID)

    '    link.Click
    '    With link.firstChild
    '        .focus
    '        .FireEvent("onclick")
    '    End With
    '    ' link.Click

    '    Dim I
    '    Dim link1
    '    link1 = htdoc.GetElementById("infoPageinfoPanelID__infoPage_myTab_object_partInfoDetailsTab")
    '    'Set link1 = HTdoc.getElementById("ext -gen397")


    '    link.TabIndex = link1.TabIndex
    '    '
    '    '    Do While IEx.readyState <> 4 Or IEx.Busy = True
    '    '        DoEvents
    '    '    Loop
    'End Sub

    'PRIVATE: Get Drawing Set Remarks

    'Private Sub GetDrawingSETRemarksProperties()
    '    Dim HtResultDoc As HtmlDocument
    '    Dim trElems As IHTMLElementCollection
    '    Dim tblRow As IHTMLTableRow
    '    Dim tempNum As Integer
    '    Dim str As String
    '    Dim I As Long
    '    'Dim j As Long 'declared by sai 

    '    HtResultDoc = IEX.Document
    '    DRWSETremarksType = True

    '    On Error Resume Next
    '    I = 0
    '    trElems = HtResultDoc.All.tags("A")
    '    Dim oLink2
    '    '    Set oLink2 = TableRow.all.tags("A").Item(0)
    '    For Each Item In trElems
    '        If (InStr(Item.innerText, "Components")) Then
    '            oLink2 = Item.innerHTML
    '            oLink2.Click
    '        End If
    '    Next

    '    For j = 1 To HtResultDoc.Links.Count
    '        If InStr(HtResultDoc.Links.Item(j).InnerText, "Components") Then

    '            oLink2 = HtResultDoc.Links.Item(j).InnerText
    '            oLink2.Click
    '        End If
    '    Next








    '    Dim j As Integer = 0
    '    While j < HtResultDoc.Links.Count And link Is Nothing ' changed to length to count 
    '        If InStr(HtResultDoc.Links(j).InnerText, "Components") > 0 Then link = HtResultDoc.Links(j)
    '        j = j + 1
    '    Wend
    '    If Not link Is Nothing Then
    '        link.Click
    '    End If
    '    For Each tblRow In trElems
    '        If tblRow.cells.Length = 5 Then
    '            If tblRow.cells.Item(0).innerHTML Like "*Remark*" Then
    '                tempNum = CInt(Trim(tblRow.cells.Item(1).innerText))
    '                If tempNum >= 100 Then
    '                    I = I + 1
    '                    '******************
    '                    'New Point Addition
    '                    '------------------
    '                    str = Left(Trim(tblRow.cells.Item(3).innerText), 2)
    '                    If str = "FN" Then

    '                    Else
    '                        DRWSETremarksType = False
    '                    End If
    '                Else
    '                    str = Left(Trim(tblRow.cells.Item(3).innerText), 2)
    '                    If str = "GN" Then

    '                    Else
    '                        DRWSETremarksType = False
    '                    End If
    '                End If
    '            End If
    '        End If
    '    Next

    '    ReDim DRWSETremarksList(1 To I, 1 To 2)
    '    I = 1
    '    For Each tblRow In trElems
    '        If tblRow.cells.Length = 5 Then
    '            If tblRow.cells.Item(0).innerHTML Like "*Remark*" Then
    '                tempNum = CInt(Trim(tblRow.cells.Item(1).innerText))
    '                If tempNum >= 100 Then
    '                    DRWSETremarksList(I, 1) = tempNum
    '                    DRWSETremarksList(I, 2) = Trim(tblRow.cells.Item(3).innerText)
    '                    I = I + 1
    '                End If
    '            End If
    '        End If
    '    Next
    'End Sub

    'PUBLIC: GoTo PART General Page and Get Data

    'Public Sub GetPARTGeneralPageProperties(ByVal isAssy As Boolean)
    '    Dim HtResultDoc As HtmlDocument
    '    Dim trElems As IHTMLElementCollection
    '    Dim tblRow As IHTMLTableRow
    '    Dim I As Long
    '    Dim col As Integer
    '    Dim aa
    '    HtResultDoc = IEX.Document

    '    On Error Resume Next
    '    If isAssy = True Then
    '        col = 0
    '    Else
    '        col = 3
    '    End If
    '    I = 1
    '    trElems = HtResultDoc.All.tags("tr")
    '    'For Each tblRow In trElems
    '    'Debug.Print tblRow.innerText
    '    ' If tblRow.className = "tableRowSelSSCI" Or tblRow.className = "tableRowNSelSSCI" Then
    '    'If tblRow.innerText Like "*marking:*" Then

    '    For I = 1 To trElems.Length
    '        If InStr(trElems.Item(I).innerText, "Item Zone:") Then
    '            aa = Split(trElems.Item(I).innerText, ":")
    '            PARTzone = Left(aa(2), 5)
    '            Exit For
    '        End If
    '        'End If
    '    Next

    '    For I = 1 To trElems.Length
    '        If InStr(trElems.Item(I).innerText, "Marking:") Then
    '            aa = Split(trElems.Item(I).innerText, ":")
    '            PARTidentification = Left(aa(3), 1)
    '            Exit For
    '        End If
    '    Next


    '    '            If InStr(tblRow.innerText, "item Zone:") Then
    '    '             PARTzone = Trim(tblRow.cells.Item(5).innerText)
    '    '            End If
    '    '
    '    '
    '    '            'End If
    '    '            If InStr(tblRow.innerText, "Marking:") Then
    '    '                PARTidentification = Trim(tblRow.cells.Item(5).innerText)
    '    '            End If

    '    '            If tblRow.innerText Like "Item Zone:" Then
    '    '                PARTzone = Trim(tblRow.cells.Item(5).innerText)
    '    '            End If
    '    ' End If
    '    'I = I + 1
    '    'Next


    'End Sub

    'PUBLIC: Get Assy BOM Data

    'Public Sub GetASSYbomItems()

    '    Dim HtResultDoc As HtmlDocument
    '    Dim Structure_Doc As HtmlDocument
    '    Dim trElems As IHTMLElementCollection
    '    Dim tblRow As IHTMLTableRow
    '    Dim bomItem As ClsItemPDM
    '    Dim I As Long
    '    Dim link As HTMLAnchorElement
    '    Dim l
    '    'GoTo PRODUCT STRUCTURE
    '    ' ClickByID ("infoPageinfoPanelID__infoPage_myTab_psb_productStructureGWT")

    '    'ClickByID ("ObjectPropNavBar:TextLink:ProductStructure")
    '    ' ClickByID ("infoPageinfoPanelID__infoPage_myTab_psb_productStructureGWT")

    '    HtResultDoc = IEX.Document
    '    ASSYbomItems = New Collection
    '    '
    '    '
    '    On Error Resume Next



    '    link = Nothing

    '    For I = 2 To HtResultDoc.Links.Count ' changed to length to count 
    '        If InStr(HtResultDoc.Links.Item(I).InnerText, "Structure") Then
    '            'If InStr(HtResultDoc.links.Item(I).innerText, "HtResultDoc.links.Item(I).Id") Then
    '            link = HtResultDoc.GetElementById(HtResultDoc.Links.Item(I).Id)
    '            'Set link = HtResultDoc.getElementByclass("x-tab-left")
    '            'Set link = HtResultDoc.findelementbyclass("x-tab-left")
    '            Exit For
    '        End If

    '    Next

    '    If Not link Is Nothing Then
    '        Call link.Click
    '    End If


    '    'Set trElems = HtResultDoc.getElementsByClassName("x-component x-tab-strip-active") ''''original code

    '    'Set trElems = HtResultDoc.getElementsByClassName("x-component x-tab-strip-over")
    '    'Set trElems = HtResultDoc.getElementsByClassName("x-component x-tab-strip-active x-tab-strip-over")
    '    trElems = HtResultDoc.GetElementById("x-auto-6__PSB.uses")


    '    link = trElems
    '    link.Click

    '    trElems = HtResultDoc.Links.Item(I).All.tags("tr")
    '    'For Each tblRow In trElems

    '    If tblRow.cells.Length = 7 Then '15


    '        If I > 1 Then
    '            If InStr(1, Trim(tblRow.cells.Item(3).innerText), "-STD") = 0 Then
    '                bomItem = New ClsItemPDM
    '                bomItem.PartNumber = Trim(tblRow.cells.Item(3).innerText)
    '                bomItem.SetQuantity(Trim(tblRow.cells.Item(7).innerText))
    '                bomItem.ManagedByCad = Trim(tblRow.cells.Item(8).innerText)
    '                bomItem.ItemNumber = Trim(tblRow.cells.Item(11).innerText)
    '                bomItem.ItemZone = Trim(tblRow.cells.Item(12).innerText)
    '                bomItem.SetPrimary(Trim(tblRow.cells.Item(13).innerText))
    '                'Add to the Collection
    '                ASSYbomItems.Add(bomItem)
    '            End If
    '        End If
    '        I = I + 1
    '    End If
    '    ' Next
    'End Sub

    'PUBLIC: Get Installation BOM Data

    'Public Sub GetINSTbomItems()

    '    Dim HtResultDoc As HtmlDocument
    '    Dim trElems As IHTMLElementCollection
    '    Dim tblRow As IHTMLTableRow
    '    Dim tblRow2 As IHTMLTableRow
    '    Dim bomItem As ClsItemPDM
    '    Dim bomBfh As ClsItemPDM
    '    Dim bomBfh2 As ClsItemPDM
    '    Dim I As Long
    '    Dim j As Long
    '    Dim k As Long
    '    Dim chkbox As HTMLCheckbox
    '    Dim lastItm As Boolean
    '    Dim stopRun As Boolean

    '    Dim test_id As HTMLAnchorElement
    '    Dim test_id2

    '    'GoTo PRODUCT STRUCTURE
    '    '    ClickByID ("ObjectPropNavBar:TextLink:ProductStructure")
    '    '    ClickByID (" x-tab-panel infoPage-tabs x-tab-panel-noborder x-border-panel")

    '    'GoTo PRODUCT STRUCTURE Cant able to open through macro href with Span
    '    'ClickByID ("psbproductStructureGWT")

    '    '    Set INSTbomItems = New Collection
    '    '
    '    '    ClickByID ("ObjectPropNavBar:TextLink:ProductStructure")
    '    On Error Resume Next
    '    I = 0
    '    Dim trElems1
    '    Dim tdElems1
    '    Dim tblRow1, oTrc


    '    IEX.refresh
    '    Call WaitTillPageGetsLoaded(IEX)
    '    Call WaitTillPageGetsLoaded(IEX)
    '    Call WaitTillPageGetsLoaded(IEX)
    '    Call WaitTillPageGetsLoaded(IEX)

    '    HtResultDoc = IEX.Document
    '    '    Set tdElems1 = HtResultDoc.getElementById("psbIFrame").Document.all.tags("tr")
    '    trElems = HtResultDoc.All.tags("tr")


    '    'Set tdElems1 = HtResultDoc.getElementsByClassName("  ext-ie ext-ie8 ext-ie10 ie-docMode8 ext-windows")
    '    '    Set tdElems1 = HtResultDoc.getElementsByClassName("x-tab-strip x-tab-strip-top").Item(1).Children.Item(1)
    '    Dim Row_Table
    '    Row_Table = tdElems1.getElementsByClassName("PSBTree")


    '    For Each tblRow1 In tdElems1
    '        Dim oTr
    '        Debug.Print(tblRow1.innerText)
    '        oTr = tblRow1.all.tags("td")
    '        Debug.Print(oTrc.innerText)
    '        For Each oTrc In oTr
    '            Debug.Print(oTrc.innerText)
    '        Next

    '    Next

    '    If I = 1 Then
    '        If InStr(1, Trim(tblRow.cells.Item(3).innerText), "-STD") = 0 Then
    '            bomItem = New ClsItemPDM
    '            bomItem.PartNumber = Trim(tblRow.cells.Item(3).innerText)
    '            bomItem.Context = Trim(tblRow.cells.Item(5).innerText)
    '            bomItem.SetQuantity(Trim(tblRow.cells.Item(7).innerText))
    '            bomItem.ManagedByCad = Trim(tblRow.cells.Item(8).innerText)
    '            bomItem.ItemNumber = Trim(tblRow.cells.Item(11).innerText)
    '            If InStr(1, bomItem.Context, "Standard-Part") > 0 Then
    '                bomItem.ItemZone = "NA"
    '            Else
    '                bomItem.ItemZone = Trim(tblRow.cells.Item(12).innerText)
    '            End If
    '            bomItem.SetPrimary(Trim(tblRow.cells.Item(13).innerText))
    '            If InStr(1, bomItem.PartNumber, "-BFH") > 0 Then
    '                If TypeOfDrawing <> 5 Then
    '                    ClickCheckBoxByID("ProductStructureTable_checkbox_checkbox__TreeTableNode_" & I - 1)
    '                End If
    '            End If
    '            'Add to the Collection
    '            INSTbomItems.Add(bomItem)
    '        End If
    '    End If
    '    I = I + 1



    '    If TypeOfDrawing <> 5 Then
    '        'Expand the Product Structure to Get the IP Numbers
    '        ClickLinkByInnerHTML("ExpandOne", 0)
    '        'Get The IP Numbers
    '        I = 0
    '        j = 0
    '        lastItm = False
    '        trElems = HtResultDoc.All.tags("tr")
    '        For I = 1 To trElems.Length
    '            tblRow = trElems.Item(I)
    '            If tblRow.cells.Length = 15 Then
    '                For j = 1 To INSTbomItems.Count
    '                    bomBfh = INSTbomItems.Item(j)
    '                    If bomBfh.ManagedByCad = True Then
    '                        If InStr(1, tblRow.cells.Item(3).innerText, bomBfh.PartNumber) > 0 Then
    '                            If j < INSTbomItems.Count Then
    '                                bomBfh2 = INSTbomItems.Item(j + 1)
    '                            Else
    '                                lastItm = True
    '                            End If
    '                            'Carefulllllll Dude
    '                            k = 0
    '                            Do
    '                                k = k + 1
    '                                I = I + 1
    '                                tblRow2 = trElems.Item(I)
    '                                'Check till when to Run
    '                                If lastItm = False Then
    '                                    If InStr(1, tblRow2.cells.Item(3).innerText, bomBfh2.PartNumber) > 0 Then
    '                                        stopRun = True
    '                                        I = I - 1
    '                                    End If
    '                                Else
    '                                    If tblRow2.cells.Length <> 15 Then
    '                                        stopRun = True
    '                                        Exit For
    '                                    End If
    '                                End If
    '                                'Get all IP Numbers
    '                                If InStr(1, tblRow2.cells.Item(6).innerText, "Agreed") > 0 Then
    '                                    If bomBfh.IPnumber1 = "" Then
    '                                        bomBfh.IPnumber1 = Trim(tblRow2.cells.Item(3).innerText)
    '                                    ElseIf bomBfh.IPnumber2 = "" Then
    '                                        bomBfh.IPnumber2 = Trim(tblRow2.cells.Item(3).innerText)
    '                                    End If
    '                                End If
    '                            Loop While Not stopRun
    '                            'Check for HOLE IP
    '                            If k = 2 Then
    '                                bomBfh.ItemNumber = "HOLE"
    '                            End If
    '                            If stopRun = True Then
    '                                stopRun = False
    '                                Exit For
    '                            End If
    '                        End If
    '                    End If
    '                Next j
    '            End If
    '            If stopRun = True Then Exit For
    '        Next I

    '        ClickByID("checkbox3_checkboxheader")
    '        ClickCheckBoxByID("ProductStructureTable_checkbox_checkbox__TreeTableNode_0")
    '        'Colapse the Product Structure for Next use
    '        ClickLinkByInnerHTML("Collapse", 1)
    '        'javascript:submitTableSelectionAction('Collapse')
    '    End If

    'End Sub

    'PRIVATE: Click the Check Box using the ID

    Private Sub ClickCheckBoxByID(ByVal clickID As String)
        'Dim htdoc As HtmlDocument
        'Dim chkbox

        'htdoc = IEX.Document
        'chkbox = htdoc.GetElementById(clickID)
        ''Swap Values
        'If chkbox.Checked = False Then
        '    chkbox.Checked = True
        'Else
        '    chkbox.Checked = False
        'End If

    End Sub


    'PRIVATE: Click Link by Inner HTML Name (Expand BFH Nodes in PRODUCT STRUCTURE)

    'Private Sub ClickLinkByInnerHTML(ByVal iNavNext As String, ByVal skipStep As Integer)
    '    Dim htdoc As HtmlDocument
    '    Dim templink As HTMLLinkElement
    '    Dim link As HTMLLinkElement
    '    Dim I As Integer

    '    iNavNext = "*" & iNavNext & "*"
    '    htdoc = IEX.Document
    '    I = 0
    '    For Each templink In htdoc.Links
    '        If templink.innerHTML Like iNavNext Then
    '            If I >= skipStep Then
    '                link = templink
    '                Exit For
    '            Else
    '                I = I + 1
    '            End If
    '        End If
    '    Next

    '    link.Click
    '    Do While IEX.readyState <> 4 Or IEX.Busy = True
    '        DoEvents
    '    Loop
    'End Sub



End Class
