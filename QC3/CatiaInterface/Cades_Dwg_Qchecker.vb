Imports DRAFTINGITF
Imports INFITF
Imports QC3.QC3.Checks

Module Cades_Dwg_Qchecker

    'form Global variable 
    Public frmStart As Form1
    Public frmStatus As frmStatus
    Public numOfDomains As Integer
    Public numOfUsers As Integer
    Public myDomains() As String
    Public myUsers() As String
    Public TypeOfMaterial As Integer   '1=Composite; 2=Metalic ; 3=Sheet Metal
    Public TypeOfSection As Integer   '1=Section-13; 2=Section-16
    Public TypeOfProgram As Integer   '1=Section-13; 2=Section-16
    Public TypeOfDrawing As Integer    '1=General; 2=SinglePart; 3=Brkt-Instl; 4=Primary-Instl; 5=S16-18-BrktInstal
    Public OpenReport As Boolean
    Public BLANK_SHEET As Boolean
    Public _singlepart As Boolean

    'General Globals
    Public PDM As clsPDMLink
    Public PDM2 As clsPDMLink
    Public drawingProperties As clsDrawShtGENProp
    Public drawingSinglePart As clsDrawShtSNGProp
    Public drawingInstalDraw As clsDrawShtINSTProp

    Public iExit As Boolean
    Public GetRemarks As Boolean        'Flag to Get the Remarks from Drawing Set
    Public CheckList As clsCHECKList
    Public CheckListCol As Collection

    Public IdentMarking1 As Boolean
    Public IdentMarking2 As Boolean
    Public FlagNoteExisting As Boolean
    Public FlagNoteErrorMsg As String
    Public FlagNoteExistDwg As Boolean
    Public FlagNoteErrMsgDwg As String
    Public BomItemNumbers As Boolean
    Public BomItemNumErrMsg As String
    Public BomItemQty As Boolean
    Public BomItemQtyErrMsg As String
    Public BomItemNumInPDM As Boolean
    Public BomItemNumInPDMErrMsg As String
    Public ExtraItemNumInDwgBol As Boolean
    Public ExtraItemNumInDwgErrMsg As String
    Public IPtableAndPDMBol As Boolean
    Public IPtableAndPDMErrMsg As String
    Public Drawingsheetno As Integer
    Public Osel As Selection
    Public DRWSETremarksType As Boolean
    Public CATIA As Application

    Public _log As log4net.ILog
    Public _dev As Boolean

    Sub Main()
        Try
            Dim docOpen As Document
            Dim drawDoc As DrawingDocument
            'CATIA = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Catia.Application"), INFITF.Application)
            'Osel = CATIA.ActiveDocument.Selection
            'Osel.Clear()
            'Dim genfun As Sys_Fun = New Sys_Fun
            'If genfun.FolderFileCheck() = True Then
            '    MsgBox("The Folder Structure is Not Set-Up Properly..." & vbCrLf & "Set-Up all the Folders as specified in the Set-Up Doucment.", vbCritical, "Set-Up Error...")
            '    Exit Sub
            'End If
            'If iExit = True Then Exit Sub
            'docOpen = CATIA.ActiveDocument
            'If TypeName(docOpen) <> "DrawingDocument" Then
            '    MsgBox("Wrong Document Type, Open a Drawing and Then Run This Macro...", vbCritical, "Wrong Document Type...")
            '    Exit Sub
            'End If

            drawDoc = docOpen
            docOpen = Nothing

            If iExit = False Then
                frmStatus = New frmStatus
                frmStatus.Show()
                If TypeOfDrawing = 1 Then  'Initialize the Global CheckList
                    CheckList = New clsCHECKList
                    'Run General Check
                    Call GetGeneralDrawingPDMdata(drawDoc)
                ElseIf TypeOfDrawing = 2 Then
                    CheckList = New clsCHECKList
                    'Run Single / Equipped Part Check
                    'GetSinglePartDrawingPDMdata(drawDoc)
                ElseIf TypeOfDrawing = 3 Then
                    ' CheckListCol = New Collection
                    GetInstallationDrawingPDMdata() 'Run Installation Check
                ElseIf TypeOfDrawing = 4 Then
                    ' CheckListCol = New Collection
                    GetInstallationDrawingPDMdata() 'Run Installation Check
                ElseIf TypeOfDrawing = 5 Then
                    'CheckListCol = New Collection
                    ' GetInstallationDrawingPDMdata2() 'Run Installation Check
                End If

                If iExit = False Then
                    '        MsgBox "Quality Check Completed..." & vbCrLf & vbCrLf & _
                    MsgBox("The Report is Saved in C:\CADESQC\REPORTS\ with File Name Ending with Today's Date and Current Time.  " &
                    vbCrLf & vbCrLf & "In Case of any Issues Contact : Prashanth Reddy", vbInformation, "Quality Check Completed... (CADES Digitech)")

                    ShowStatus("Writing Check-List...", 99, True)
                    'If CheckListCol Is Nothing And CheckList.List.Count >= 1 Then
                    '    genfun.FillSingleCheckReport() 'Write CheckList Data to Excel Report
                    'Else
                    '    For Each CheckList In CheckListCol
                    '        genfun.FillSingleCheckReport() 'Write CheckList Data to Excel Report
                    '    Next

                    'End If
                    ShowStatus("Finalizing Quality Check...", 100, True)
                    frmStatus.Hide()
                    frmStatus.Close()
                    PDM = Nothing
                    PDM2 = Nothing
                    drawingProperties = Nothing
                    drawingSinglePart = Nothing
                    drawingInstalDraw = Nothing
                    drawingProperties = Nothing
                    PDM = Nothing
                    CheckListCol = Nothing
                    CheckList = Nothing
                ElseIf Not TypeOfDrawing = 3 Then

                    MsgBox("Critical Error..." & vbCrLf & vbCrLf & "Contact: Prashanth Reddy.", vbCritical, "Critical Error...")
                    ShowStatus("Finalizing Quality Check...", 100, True)
                    frmStatus.Hide()
                End If
            End If
            PDM = Nothing
            PDM2 = Nothing
            drawingProperties = Nothing
            drawingSinglePart = Nothing
            drawingInstalDraw = Nothing
            drawingProperties = Nothing
            PDM = Nothing
            CheckListCol = Nothing
            CheckList = Nothing
            MsgBox("Qc completed for  " & drawDoc.Name, vbExclamation, "AXISCADES")
        Catch ex As Exception
            frmStatus.Close()
        End Try
    End Sub

    Public Sub ShowStatus(ByVal status As String, ByVal percent As Integer, ByVal toUpdate As Boolean)
        frmStatus.txtlabel.Text = status
        If toUpdate = True Then
            frmStatus.barStatus.Width = 200 * (percent / 100)
            frmStatus.txtPercent.Text = percent & "%"
        End If
        'frmStatus.Repaint
        frmStatus.txtlabel2.Text = "   " & Drawingsheetno & "\" & CATIA.Documents.Count

    End Sub


    Public Sub GenerateaReport()
        Try

            'Call WaitTillPageGetsLoaded(oIEX)

            'PB.oStrutPage.GetINSTbomItems
            'PB.oStrutPage.DRWSETremarksList = drawingInstalDraw.PDMpropCol.Item(1).DRWSETremarksList
            'PB.oStrutPage.ClosePDMLink
            ''*************
            'ShowStatus("Verifying Data...", 90, True)
            ''*************
            'PDM2 = oStrutPage

            CheckFilledBOM(PDM2.INSTbomItems)
            CompareINSTBomItemNumbers()
            CompareINSTflagNotes()
            CompareIPtableBOM()

            '-----------------------
            'Populate the Check List
            '-----------------------
            'New Installation CheckList
            CheckList = New clsCHECKList
            'Set Other CheckList Details
            Dim htz As String
            Dim str As String
            htz = drawingInstalDraw.GenPropCol.Item(1).DRWhtzNumber
            str = drawingInstalDraw.GenPropCol.Item(1).DRWsheetNumber
            Dim I
            For I = 2 To drawingInstalDraw.GenPropCol.Count
                str = str & ", " & drawingInstalDraw.GenPropCol.Item(I).DRWsheetNumber
            Next I
            '*************
            ShowStatus("Updating Check-List...", 95, True)
            '*************
            CheckList.DrawingNumber = htz & "-" & str
            CheckList.DrawingName = drawingInstalDraw.PDMpropCol.Item(1).DRWSETname
            CheckList.DrawingState = "-"
            CheckList.DrawingVersion = "-"

            '1.Item Zone in Drawing Sheet and PDM
            CheckList.AddCompareCheckPoint(drawingInstalDraw.PartNumsZones,
                        PDM2.PARTzone,
                        "Installation Item Number Zone in Drawing and PASS",
                        "OK",
                        "Item Zone in PASS is Not Updated.")
            '2.Identification Marking in PDM Must be Empty
            CheckList.AddCompareCheckPoint("",
                        PDM2.PARTidentification,
                        "Identification Marking in PASS",
                        "OK",
                        "Identification Marking in PASS Must be Empty.")

            If TypeOfProgram = 1 Then
                '3.Sht-1 & Sht-2 Baloon Match
                CheckList.AddBooleanCheckPoint(drawingInstalDraw.PartNumMatchBol,
                            "Installation Item Numbers in Sheet-1 and Sheet-2 are Matching",
                            "OK",
                            drawingInstalDraw.PartNumMatchErrMsg)
            ElseIf TypeOfProgram = 2 Then
                '3.Sht-1 & Sht-2 Baloon Match
                CheckList.AddWarningBooleanCheckPoint(False,
                            "Installation Item Numbers in Sheet-1 and Sheet-2 are Matching",
                            "",
                            "Installation Item Numbers in Sheet-1 and Sheet-2 is Not Checked for A350-1000. Manual Check must be performed.")
            End If

            '4.Extra Installation Baloon
            CheckList.AddBooleanCheckPoint(drawingInstalDraw.PartNumExtraBol,
                        "Extra Installation Item Number in other Drawing Sheets",
                        "OK",
                        drawingInstalDraw.PartNumExtraErrMsg)
            '5.Check if the BOM in PDM is filled
            CheckList.AddBooleanCheckPoint(BomItemNumInPDM,
                        "Assembly BOM in PASS is Completly Filled",
                        "OK",
                        "Some Feilds in PASS BOM are Not Filled.")
            '6.Compare BOM Item Numbers.
            CheckList.AddBooleanCheckPoint(BomItemNumbers,
                        "Assembly BOM Item Numbers & Zones in PASS and Drawing" & vbCrLf & "(STD parts are Ignored, Check STD Item Numbers Manually)",
                        "OK",
                        BomItemNumErrMsg)
            '7.Compare BOM Item Numbers Quantity.
            CheckList.AddBooleanCheckPoint(BomItemQty,
                        "Assembly BOM Item Quantity in PASS and Drawing Sheet" & vbCrLf & "(STD parts are Ignored, Check STD Item Numbers Manually)",
                        "OK",
                        BomItemQtyErrMsg)
            '8.Check if Extra Item Numbers are present in Drawing.
            CheckList.AddBooleanCheckPoint(ExtraItemNumInDwgBol,
                        "Extra Item Numbers are Present in Drawing than PASS",
                        "OK",
                        ExtraItemNumInDwgErrMsg)

            If FlagNoteErrorMsg <> "NA" Then
                '9.Flag Notes in PDM are Present in Drawing
                CheckList.AddWarningBooleanCheckPoint(FlagNoteExisting,
                            "Flag Notes in PASS are Present in Drawing Sheets",
                            "OK",
                            FlagNoteErrorMsg)
                '10.Flag Notes in Drawing are Present in PDM
                CheckList.AddBooleanCheckPoint(FlagNoteExistDwg,
                            "Flag Notes in Drawing Sheets are Present in PASS",
                            "OK",
                            FlagNoteErrMsgDwg)
            Else
                '9.Flag Notes in PDM are Present in Drawing
                CheckList.AddWarningBooleanCheckPoint(False,
                            "Flag Notes in PASS are Present in Drawing Sheets",
                            "",
                            "There are No Flag Notes in PASS")
                '10.Flag Notes in Drawing are Present in PDM
                CheckList.AddWarningBooleanCheckPoint(False,
                            "Flag Notes in Drawing Sheets are Present in PASS",
                            "",
                            "There are No Flag Notes in Drawing Sheets")
            End If

            '11.Check if Remarks in PDM Link starting with FN / GN.
            CheckList.AddBooleanCheckPoint(DRWSETremarksType,
                        "Remarks in PDM Link starting with FN / GN",
                        "OK",
                        "Some Remarks are not starting with FN / GN or the Type of Remarks is wrong.")
            CheckList.AddWarningBooleanCheckPoint(DRWSETremarksType,
                        "Remarks in PASS starting with FN / GN",
                        "OK",
                        "Some Remarks are not starting with FN / GN or the Type of Remarks is wrong.") 'CHANGED TO WA AS PER HARISH

            '12.Compare PDM Bom and IP table in Drawing
            CheckList.AddWarningBooleanCheckPoint(IPtableAndPDMBol,
                        "IP Table in 1st Sheet and PASS BOM Check",
                        "OK",
                        IPtableAndPDMErrMsg)

            '13.Baloon Size
            CheckList.AddBooleanCheckPoint("BaloonSize",
                        "Baloon Size",
                        "OK",
                        "Following Baloons Size are not correct " & "WrongBaloonSizeText") 'added made it as string

            '14.Flag Note
            CheckList.AddBooleanCheckPoint("Flagnotesize",
                        "FLANG NOTE SIZE",
                        "OK",
                        "Following FLANG NOTE  Size are not correct " & "WrongFlangnoteSizeText") ' added made it as string 


            'Add to Check List Collection
            CheckListCol.Add(CheckList)
            'Generating The Report
            Call GenerateaReport()

            If iExit = False Then
                '        MsgBox "Quality Check Completed..." & vbCrLf & vbCrLf & _
                MsgBox("The Report is Saved in C:\CADESQC\REPORTS\ with File Name Ending with Today's Date and Current Time.  " &
            vbCrLf & vbCrLf & "In Case of any Issues Contact : Prashanth Reddy", vbInformation, "Quality Check Completed... (CADES Digitech)")
                '*************
                ShowStatus("Writing Check-List...", 99, True)
                '*************
                Dim genfun As Sys_Fun = New Sys_Fun
                If CheckListCol Is Nothing Then
                    'Write CheckList Data to Excel Report
                    genfun.FillSingleCheckReport()
                Else
                    For Each CheckList In CheckListCol
                        genfun.FillSingleCheckReport()
                    Next
                End If
            Else
                MsgBox("Critical Error..." & vbCrLf & vbCrLf & "Contact: Prashanth Reddy.", vbCritical, "Critical Error...")
            End If


            ShowStatus("Finalizing Quality Check...", 100, True)
            frmStatus.Hide()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub GetInstallationDrawingPDMdata()
        Try
            Dim openDocs As Documents
            Dim tempDoc As Document
            Dim drawDoc As DrawingDocument
            'Dim drawShts As DrawingSheets
            'Dim drawSht As DrawingSheet
            'Dim drawViews As DrawingViews
            'Dim drawView As DrawingView
            'Dim drawTexts As DrawingTexts
            'Dim drawText As DrawingText
            'Dim GenProps As ClsDrawShtGENProp
            Dim str As String
            Dim htz As String
            Dim pass As Boolean

            Dim tempCollection As Collection
            Dim I As Integer
            Dim sss

            '*************
            frmStatus.Show()
            ShowStatus("Reading CATIA Drawing...", 1, True)
            '*************

            openDocs = CATIA.Documents
            drawingInstalDraw = New ClsDrawShtINSTProp

            'Get HTZ of the Active Sheet
            tempDoc = CATIA.ActiveDocument
            If TypeName(tempDoc) = "DrawingDocument" Then
                drawDoc = tempDoc
                sss = Split(drawDoc.Sheets.Name, "\")
                htz = Left(sss(UBound(sss)), 9)
            End If

            'Add Drawing Documents to Collection
            pass = False
            For Each tempDoc In openDocs
                If TypeName(tempDoc) = "DrawingDocument" Then
                    drawDoc = tempDoc
                    sss = Split(drawDoc.Sheets.Name, "\")
                    str = Left(sss(UBound(sss)), 12)
                    If InStr(1, str, htz) > 0 Then
                        pass = True
                    End If
                    If pass = True Then
                        'Add to Collection
                        tempCollection.Add(drawDoc, str)
                        pass = False
                    End If
                End If
            Next
            'Check if all sheets are Open
            If tempCollection.Count < 3 Then
                iExit = True
                MsgBox("Minimum 3 Number of Drawing Sheets are Required to To Check the Installation Drawings.", vbCritical, "Critical Error...")
                Exit Sub
            End If
            If iExit = True Then
                MsgBox("Minimum 3 Number of Drawing Sheets are Required to To Check the Installation Drawings.", vbCritical, "Critical Error...")
                Exit Sub
            End If

            'Sort the Drawing Docs in Ascending Order
            Call SortDrawingDocsAsc(tempCollection)

            drawingInstalDraw.DrawDocsCol = tempCollection
            tempCollection = Nothing

            '*************
            ShowStatus("Reading CATIA Drawing...", 10, True)
            '*************

            'General Check for all the Sheets
            For I = 1 To drawingInstalDraw.DrawDocsCol.Count
                drawDoc = drawingInstalDraw.DrawDocsCol.Item(I)
                '*************
                ShowStatus("Reading Drawing & PASS Data... [ " & Left(drawDoc.Sheets.Name, 14) & " ]", 10 + (50 * (I / drawingInstalDraw.DrawDocsCol.Count)), True)
                '*************
                'New CheckList for Each Drawing Sheet
                CheckList = New clsCHECKList
                'Get All General Properties From All Sheets
                If I = 1 Then
                    GetRemarks = True
                Else
                    GetRemarks = False
                End If
                GetGeneralDrawingPDMdata(drawDoc)
                'Add to Collections
                CheckListCol.Add(CheckList)
                drawingInstalDraw.GenPropCol.Add(drawingProperties)
                drawingInstalDraw.PDMpropCol.Add(PDM)
            Next

            '*************
            ShowStatus("Reading Drawing & PASS Data...[Installation Data]", 63, True)
            '*************

            'Get Installation Drawing Properties
            'Call GetINSTDrawingSetProperties(drawingInstalDraw)

            '-------------------------------------------------------------
            ''''        'Run PDM only if Part/Assy Number is Found in the Drawing
            ''''        If Not drawingInstalDraw.PartNumSht1 Is Nothing Then
            ''''        '*************
            ''''         ShowStatus "Extracting PDM Link Data...", 65, True
            ''''        '*************
            ''''            '------------------------------------
            ''''            'START PDM FOR INSTALLATION PROPERTIES
            PDM2 = New ClsPDMLink
            ''''            'Get All PDM2 Data of the Drawing
            PDM2.StartPDMLink()

            'If PDM2.SearchHTZ(drawingInstalDraw.GenPropCol.Item(1).DRWhtzNumber & "000") Then
            '    'PDM2.GetLatestVersionProperties()
            '    '*************
            '    ShowStatus("Extracting PASS Data...", 70, True)
            '    '*************
            'Else
            '    MsgBox("Drawing Can Not Find the Drawing Number in PASS.", vbCritical, "Quiting Macro...")
            '    Exit Sub
            'End If
            'PDM2.GetPARTGeneralPageProperties(True)
            ''''        '*************
            ''''         ShowStatus "Extracting PDM Link BOM Data... [ " & drawingInstalDraw.GenPropCol.Item(1).DRWhtzNumber & drawingInstalDraw.PartNumSht1.TextContent & " ]", 80, True
            ''''        '*************
            ''''            PDM2.GetINSTbomItems
            ''''            PDM2.DRWSETremarksList = drawingInstalDraw.PDMpropCol.Item(1).DRWSETremarksList
            'PDM2.ClosePDMLink()
            ''''            '------------------------------------
            ''''        End If
            ''''        '*************
            ShowStatus("Verifying Data...", 90, True)
            ''''        '*************
            ''''
            ''''        CheckFilledBOM PDM2.INSTbomItems
            ''''        CompareINSTBomItemNumbers
            ''''        CompareINSTflagNotes
            ''''        CompareIPtableBOM

            '-----------------------
            'Populate the Check List
            '-----------------------
            'New Installation CheckList
            CheckList = New clsCHECKList
            'Set Other CheckList Details
            htz = drawingInstalDraw.GenPropCol.Item(1).DRWhtzNumber
            str = drawingInstalDraw.GenPropCol.Item(1).DRWsheetNumber
            For I = 2 To drawingInstalDraw.GenPropCol.Count
                str = str & ", " & drawingInstalDraw.GenPropCol.Item(I).DRWsheetNumber
            Next I
            '*************
            ShowStatus("Updating Check-List...", 95, True)
            '*************
            CheckList.DrawingNumber = htz & "-" & str
            CheckList.DrawingName = drawingInstalDraw.PDMpropCol.Item(1).DRWSETname
            CheckList.DrawingState = "-"
            CheckList.DrawingVersion = "-"

            '1.Item Zone in Drawing Sheet and PDM
            '''         CheckList.AddCompareCheckPoint (drawingInstalDraw.PartNumsZones, PDM2.PARTzone, "Installation Item Number Zone in Drawing and PDMLink", _
            '''                    "OK", _
            '''                    "Item Zone in PDMLink is Not Updated.")
            CheckList.AddCompareCheckPoint("",
        "",
        "Installation Item Number Zone in Drawing and PASS",
        "OK",
        "Item Zone in PDMLink is Not Updated.")

            '2.Identification Marking in PDM Must be Empty
            CheckList.AddCompareCheckPoint("",
                    PDM2.PARTidentification,
                    "Identification Marking in PASS",
                    "OK",
                    "Identification Marking in PDMLink Must be Empty.")

            If TypeOfProgram = 1 Then
                '3.Sht-1 & Sht-2 Baloon Match
                CheckList.AddBooleanCheckPoint(True,
                        "Installation Item Numbers in Sheet-1 and Sheet-2 are Matching",
                        "OK",
                        drawingInstalDraw.PartNumMatchErrMsg)


                ''        ElseIf TypeOfProgram = 2 Then
                ''            '3.Sht-1 & Sht-2 Baloon Match
                ''            CheckList.AddWarningBooleanCheckPoint False, _
                ''                        "Installation Item Numbers in Sheet-1 and Sheet-2 are Matching", _
                ''                        "", _
                ''                        "Installation Item Numbers in Sheet-1 and Sheet-2 is Not Checked for A350-1000. Manual Check must be performed."
            End If

            '4.Extra Installation Baloon
            CheckList.AddBooleanCheckPoint(drawingInstalDraw.PartNumExtraBol,
                    "Extra Installation Item Number in other Drawing Sheets",
                    "OK",
                    drawingInstalDraw.PartNumExtraErrMsg)
            '5.Check if the BOM in PDM is filled
            CheckList.AddBooleanCheckPoint(BomItemNumInPDM,
                    "Assembly BOM in PDMLink is Completly Filled",
                    "OK",
                    "Some Feilds in PDMLink BOM are Not Filled.")
            '6.Compare BOM Item Numbers.
            CheckList.AddBooleanCheckPoint(BomItemNumbers,
                    "Assembly BOM Item Numbers & Zones in PDMLink and Drawing" & vbCrLf & "(STD parts are Ignored, Check STD Item Numbers Manually)",
                    "OK",
                    BomItemNumErrMsg)
            '7.Compare BOM Item Numbers Quantity.
            CheckList.AddBooleanCheckPoint(BomItemQty,
                    "Assembly BOM Item Quantity in PDMLink and Drawing Sheet" & vbCrLf & "(STD parts are Ignored, Check STD Item Numbers Manually)",
                    "OK",
                    BomItemQtyErrMsg)
            '8.Check if Extra Item Numbers are present in Drawing.
            CheckList.AddBooleanCheckPoint(ExtraItemNumInDwgBol,
                    "Extra Item Numbers are Present in Drawing than PDM Link",
                    "OK",
                    ExtraItemNumInDwgErrMsg)

            '        If FlagNoteErrorMsg <> "NA" Then
            '            '9.Flag Notes in PDM are Present in Drawing
            '            CheckList.AddWarningBooleanCheckPoint FlagNoteExisting, _
            '                        "Flag Notes in PDM Link are Present in Drawing Sheets", _
            '                        "OK", _
            '                        FlagNoteErrorMsg
            '            '10.Flag Notes in Drawing are Present in PDM
            '            CheckList.AddBooleanCheckPoint FlagNoteExistDwg, _
            '                        "Flag Notes in Drawing Sheets are Present in PDM Link", _
            '                        "OK", _
            '                        FlagNoteErrMsgDwg
            '        Else
            '            '9.Flag Notes in PDM are Present in Drawing
            '            CheckList.AddWarningBooleanCheckPoint False, _
            '                        "Flag Notes in PDM Link are Present in Drawing Sheets", _
            '                        "", _
            '                        "There are No Flag Notes in PDM Link"
            '            '10.Flag Notes in Drawing are Present in PDM
            '            CheckList.AddWarningBooleanCheckPoint False, _
            '                        "Flag Notes in Drawing Sheets are Present in PDM Link", _
            '                        "", _
            '                        "There are No Flag Notes in Drawing Sheets"
            '        End If

            '11.Check if Remarks in PDM Link starting with FN / GN.
            '        CheckList.AddBooleanCheckPoint DRWSETremarksType, _
            '"Remarks in PDM Link starting with FN / GN", _
            '  "OK", _
            '"Some Remarks are not starting with FN / GN or the Type of Remarks is wrong."

            '12.Compare PDM Bom and IP table in Drawing
            '        CheckList.AddWarningBooleanCheckPoint IPtableAndPDMBol, _
            '"IP Table in 1st Sheet and PDM Link BOM Check", _
            '"OK", _
            'IPtableAndPDMErrMsg

            'Add to Check List Collection
            CheckListCol.Add(CheckList)


        Catch ex As Exception

        End Try
    End Sub

    '************************************************************


    '---------------------------------
    ' SINGLE PART DRAWING SHEET CHECK
    '---------------------------------

    Private Sub GetSinglePartDrawingPDMdata(ByRef drawDoc As DrawingDocument)
        Try
            '*************
            frmStatus.Show()
            ShowStatus("Reading CATIA Drawing...", 10, True)
            '*************
            '----------------------------
            'Get All General Check Points
            GetRemarks = True
            Call GetGeneralDrawingPDMdata(drawDoc)
            '----------------------------
            '*************
            ShowStatus("Reading CATIA Drawing...", 25, True)
            '*************
            '-------------------------------------------------------------
            'Initialize the Ojects and Assign General Props to Single Part
            drawingSinglePart = New ClsDrawShtSNGProp
            drawingSinglePart.GENprop = drawingProperties

            drawingSinglePart.GetSNGDrawingSheetProperties(drawDoc)
            '*************
            ShowStatus("Reading CATIA Drawing...", 50, True)
            '*************
            '-------------------------------------------------------------
            'Run PDM only if Part/Assy Number is Found in the Drawing
            If Not drawingSinglePart.PartNumber Is Nothing Then
                '*************
                ShowStatus("Extracting PASS Data...", 60, True)
                '*************
                '------------------------------------
                'START PDM FOR SINGLE PART PROPERTIES
                PDM2 = New ClsPDMLink
                'Get All PDM2 Data of the Drawing
                PDM2.StartPDMLink()

                'If PDM2.SearchHTZ(drawingSinglePart.GENprop.DRWhtzNumber & drawingSinglePart.PartNumber.TextContent) Then
                '    'PDM2.GetLatestVersionProperties()
                'Else
                '    MsgBox("Drawing Can Not Find the Drawing Number in PASS.", vbCritical, "Quiting Macro...")
                '    Exit Sub
                'End If
                'PDM2.GetPARTGeneralPageProperties(drawingSinglePart.isAssembly)

                'if the drawing sheet is an Assy the Get BOM Properties
                If drawingSinglePart.isAssembly = True Then
                    '*************
                    ShowStatus("Extracting PASS BOM Data..." & " [ " & drawingSinglePart.GENprop.DRWhtzNumber & drawingSinglePart.PartNumber.TextContent & " ]", 70, True)
                    '*************
                    'PDM2.GetASSYbomItems()
                    'PDM2.GetINSTbomItems
                End If

                'PDM2.ClosePDMLink()
                '------------------------------------
            End If
            '*************
            ShowStatus("Verifying Data...", 80, True)
            '*************

            If Not drawingSinglePart.PartNumber Is Nothing Then
                'Compare Identfication Marking
                'CompareIdentificationMarking
                'Compare Flag Notes
                CompareFlagNotes()
            End If

            If drawingSinglePart.isAssembly = True Then 'AFTER ASSEMBLEY COME HAER CHANADAN
                'Check if the BOM in PDM is filled
                CheckFilledBOM(PDM2.ASSYbomItems)
                'Compare BOM Item Numbers
                CompareBomItemNumbers()
            End If

            '*************
            ShowStatus("Updating Check-List...", 90, True)
            '*************
            '-----------------------
            'Populate the Check List
            '-----------------------
            'If Flag Note is Missing in both PDM and Drawing
            If FlagNoteErrorMsg <> "NA" Then
                '30.Flag Notes in PDM and Drawing
                CheckList.AddBooleanCheckPoint(FlagNoteExisting,
                        "Flag Notes in PASS present in Drawing Sheet",
                        "OK",
                        FlagNoteErrMsgDwg)
            Else
                '30.Flag Notes in PDM and Drawing
                CheckList.AddBooleanCheckPoint(True,
                        "Flag Notes in PASS present in Drawing Sheet",
                        "Flag Notes Were Not Found in Drawing and PASS.",
                        "")
            End If

            '31/03/2015
            'Newly added Point
            'If Flag Note is Missing in both PDM and Drawing
            If FlagNoteErrMsgDwg <> "" Then
                '31.Flag Notes in PDM and Drawing
                CheckList.AddBooleanCheckPoint(FlagNoteExistDwg,
                        "Flag Notes in Drawing Sheet present in PASS",
                        "OK",
                        FlagNoteErrMsgDwg)
            Else
                '31.Flag Notes in PDM and Drawing
                CheckList.AddBooleanCheckPoint(True,
                        "Flag Notes in Drawing Sheet present in PASS",
                        "Flag Notes Were Not Found in Drawing and PASS.",
                        "")
            End If


            '31.Check if Remarks in PDM Link starting with FN / GN.
            'CheckList.AddBooleanCheckPoint( DRWSETremarksType, _
            '"Remarks in PDM Link starting with FN / GN", _
            '"OK", _
            '"Some Remarks are not starting with FN / GN or the Type of Remarks is wrong.")
            CheckList.AddWarningBooleanCheckPoint(DRWSETremarksType,
                    "Remarks in PASS starting with FN / GN",
                    "OK",
                    "Some Remarks are not starting with FN / GN or the Type of Remarks is wrong.") 'CHANGED TO WA AS Per HARISH

            'Execute these Check Points only if Part Number is Present in the Drawing
            If Not drawingSinglePart.PartNumber Is Nothing Then
                '32.Item Zone in Drawing Sheet and PDM
                CheckList.AddCompareCheckPoint(drawingSinglePart.PartNumber.TextZone,
                        PDM2.PARTzone,
                        "Part/Assembly Item Number Zone in Drawing and PASS",
                        "OK",
                        "Item Zone in PASS is Not Updated.")
                '            '33.Compare Identification Marking in Drawing and PDM "Commented on 13-11-2022 Requeested by chetan
                '            CheckList.AddWarning2BooleanCheckPoint IdentMarking1, _
                '                        IdentMarking2, _
                '                        "Identification Marking in Drawing and PASS", _
                '                        "OK", _
                '                        "N/A", _
                '                        "Identification Marking in Drawing and PASS Does Not Match. " & drawingSinglePart.GENprop.TBLKidentMarkComb & " <AND> " & PDM2.PARTidentification

                'These Points Valid Only for Equiped Parts
                If drawingSinglePart.isAssembly = True Then
                    '34.Check if the BOM in PDM is filled
                    CheckList.AddBooleanCheckPoint(BomItemNumInPDM,
                            "Assembly BOM in PASS is Completly Filled",
                            "OK",
                            BomItemNumInPDMErrMsg)
                    '35.Compare BOM Item Numbers.
                    CheckList.AddBooleanCheckPoint(BomItemNumbers,
                            "Assembly BOM Item Numbers & Zones in PASS and Drawing",
                            "OK",
                            BomItemNumErrMsg)
                    '36.Compare BOM Item Numbers Quantity.
                    CheckList.AddBooleanCheckPoint(BomItemQty,
                            "Assembly BOM Item Quantity in PASS and Drawing Sheet",
                            "OK",
                            BomItemQtyErrMsg)
                    '37.Check if Extra Item Numbers are present in Drawing.
                    CheckList.AddBooleanCheckPoint(ExtraItemNumInDwgBol,
                            "Extra Item Numbers are Present in Drawing than PASS",
                            "OK",
                            ExtraItemNumInDwgErrMsg)
                End If
            Else
                '31.If Item Number is missing in Drawing Sheet.
                CheckList.AddBooleanCheckPoint(False,
                        "Part / Assembly Number Baloon",
                        "",
                        "Part / Assembly Number Missing in Drawing Sheet.")
            End If


            '32.Flag Notes in PDM and Drawing
            CheckList.AddBooleanCheckPoint(drawingProperties.BaloonSize,
                "Baloon Size",
                "OK",
                "Following Baloons Size are not correct " & drawingProperties.WrongBaloonSizeText)



        Catch ex As Exception

        End Try

    End Sub

    '---------------------------------
    ' GENERAL DRAWING SHEET CHECK
    '---------------------------------

    Private Sub GetGeneralDrawingPDMdata(ByRef drawDoc As DrawingDocument)
        Try

            PDM = New ClsPDMLink                        '> 'Initialize the Ojects
            drawingProperties = New ClsDrawShtGENProp   '>
            'testing purpose
            PDM.StartPDMLink()
            '*************
            If TypeOfDrawing = 1 Then                       '>
                frmStatus.Show()                                  '>
                ShowStatus("Reading CATIA Drawing...", 25, True)
            End If
            '*************

            'Get All Drawing Sheet Data
            'Call drawingProperties.GetGENDrawingSheetProperties(drawDoc) ' Solved

            '*************
            If TypeOfDrawing = 1 Then
                ShowStatus("Extracting PDM Link Data...", 50, True)
            End If
            '*************
            'Get All PDM Data of the Drawing
            PDM.StartPDMLink()     ' Solved


            'If PDM.SearchHTZ(drawingProperties.DRWhtzNumber & "-" & drawingProperties.DRWsheetNumber) Then                  ' SearchHTZ Solved

            '    'Call PDM.GetLatestVersionProperties()

            'Else
            '    MsgBox("Drawing Can Not Find the Drawing Number in PDM Link.", vbCritical, "Quiting Macro...")
            '    Exit Sub
            'End If


            ' -----------------------------------

            'Swap the Variables from SR
            PDM.SwapSRPtoDRWSHTProperties()


            'PDM.GetDrawingGeneralPageProperties()


            PDM.GetDrawingSetLinkProperties(drawingProperties.DRWhtzNumber)
            If GetRemarks = True Then
                '    ''NA
                '    'Need click on the Components tab to get ComponentProperties
                '    'Pending click is not working
                '    'PDM.GetDrawingSetComponentProperties drawingProperties.DRWhtzNumber
            End If
            'PDM.SwapSRPtoDRWSETProperties()
            'PDM.ClosePDMLink()

            '*************
            If TypeOfDrawing = 1 Then
                ShowStatus("Updating Check-List...", 80, True)
            End If
            '*************
            'Populate the Check List

            '1.Sheet size TB and Actual
            CheckList.AddCompareCheckPoint(drawingProperties.TBLKsheetSize,
                    drawingProperties.DRWactualSheetSize,
                    "Sheet Size in Title Block and Actual Sheet Size in CATIA",
                    "OK",
                    "Title Block Not Updated.")
            '2.Sheet size TB and PDM
            'CheckList.AddCompareCheckPoint(drawingProperties.TBLKsheetSize,
            '            PDM.DRWSHTdrawingSize,
            '            "Sheet Size in Title Block and in PASS",
            '            "OK",
            '            "Sheet Size in PASS Not Updated.")
            '3.Sheet Title Name TB and PDM
            'CheckList.AddCompareCheckPoint(drawingProperties.TBLKenglishTitle,
            '            PDM.DRWSETname,
            '            "Drawing English Title in Title Block and Drawing Name in PASS",
            '            "OK",
            '            "Name in Title Block and PASS does not Match.")
            '4.Sheet Title Local Name TB and PDM
            CheckList.AddCompareCheckPoint(drawingProperties.TBLKdomesticTitle,
                    "",
                    "Drawing Domestic Title in Title Block Should be Empty",
                    "OK",
                    "Domestic Title in Title Block is Not Empty.")
            '5.Sheet Issue in TB and PDM
            'CheckList.AddCompareCheckPoint(drawingProperties.TBLKfirstIssue,
            '            PDM.DRWSHTissueIndex,
            '            "Drawing Issue in Title Block and PASS",
            '            "OK",
            '            "Issue in Title Block is Wrong.")
            '6.Drawing Second Issue in TB
            CheckList.AddCompareCheckPoint(drawingProperties.TBLKsecondIssue,
                    "",
                    "Drawing Second Issue in Title Block Should be Empty",
                    "OK",
                    "Second Issue in Title Block should be Empty.")
            '7.Drawing Number TB and FileName
            CheckList.AddCompareCheckPoint(drawingProperties.TBLKdrawingNumber,
                    drawingProperties.DRWhtzNumber,
                    "Drawing Number in Title Block and File Name",
                    "OK",
                    "Drawing Number is Not Same as the File Name.")
            '8.Sheet Number TB and FileName
            CheckList.AddCompareCheckPoint(drawingProperties.TBLKsheetNumber,
                    drawingProperties.DRWsheetNumber,
                    "Drawing Sheet Number in Title Block and File Name",
                    "OK",
                    "Drawing Sheet Number is Not Same as the File Name.")
            '9.Drawing Airbus Germany is used
            CheckList.AddBooleanCheckPoint(drawingProperties.TBLKairbusGmbH,
                    "Airbus Germany Title Block Used",
                    "OK",
                    "Wrong Natco Used in Title Block.")
            '10.Drawing Year in Title Block
            CheckList.AddBooleanCheckPoint(drawingProperties.TBLKyearBol,
                    "Copyright Year in Title Block",
                    "OK",
                    "Wrong Copyright Year used in Title Block.")
            '11.Drawing Signatures in Title Block
            CheckList.AddBooleanCheckPoint(drawingProperties.TBLKsignatures,
                    "Signatures in Title Block",
                    "OK",
                    "Wrong Signatures in Title Block.")
            '12.Drawing Interchangability in Title Block
            CheckList.AddBooleanCheckPoint(drawingProperties.TBLKinterchangeable,
                    "Interchangeable Part in Title Block",
                    "OK",
                    "Wrong Interchangeable Part in Title Block.")
            '13.Drawing ABD0001 in Title Block
            CheckList.AddBooleanCheckPoint(drawingProperties.TBLKabd0001Bol,
                    "Limits Not Stated in Title Block",
                    "OK",
                    "Wrong ABD0001 Code Used in Title Block.")
            '14.Geo Location in Title Block should be Empty (REMOVED AS PER CHETAN TAMKER)
            'CheckList.AddCompareCheckPoint drawingProperties.TBLKgeoLocation, _
            '"", _
            '"Drawing Geographic Reference in Title Block must be Empty", _
            '"OK", _
            '"Geographic Reference '" & drawingProperties.TBLKgeoLocation & "' in Title Block must be Empty."
            '15.Drawing Surface Finish in Title Block (REMOVED AS PER CHETAN TAMKER)
            'CheckList.AddWarning2BooleanCheckPoint drawingProperties.TBLKsurfaceFinishBol1, _
            'drawingProperties.TBLKsurfaceFinishBol2,
            '"Surface Finish (ABD0002) in Title Block",
            '"OK",
            '"Verify if Surface Finish in Title Block is Correct. ( " & drawingProperties.TBLKsurfaceFinish & " )",
            '"Surface Finish in Title Block is Not as per Standard."
            '16.Standard Scale Used in Title Block
            CheckList.AddBooleanCheckPoint(drawingProperties.TBLKscaleBol,
                    "Standard Scale in Title Block",
                    "OK",
                    "Scale Used in Title Block is Not as per Airbus Standard (Select From Title Block Tool).")
            '17.Views Scale and Title Block Scale same
            CheckList.AddBooleanCheckPoint(drawingProperties.TBLKscaleBol2,
                    "Main View Scale and Title Block Scale Check",
                    "OK",
                    "Scale Used in Title Block is Not Matching with the Main View Scale.")
            '18.Drawing Cancellation Index Status
            'CheckList.AddBooleanCheckPoint(PDM.DRWSHTcancellationIndex,
            '            "Drawing Cancellation Index Status in PASS",
            '            "OK",
            '            "Cancellation Index Status in PASS is True.")
            '19.Drawing Frame Grid Check (REMOVED AS PER CHETAN TAMKER)
            'CheckList.AddBooleanCheckPoint drawingProperties.DRWframeGridUsed, _
            '"Drawing Frame Grid as per New Airbus Standard", _
            '"OK", _
            '"Drawing Frame Grid is Not as per Airbus Standard."

            If BLANK_SHEET <> True Then
                '20.Drawing View Lock Check
                CheckList.AddBooleanCheckPoint(drawingProperties.DRWviewLockStatus,
                        "All Drawing Views are Locked",
                        "OK",
                        "All Drawing Views are Not Locked.")
                '21.Drawing Views are in Show Mode
                CheckList.AddWarningBooleanCheckPoint(drawingProperties.DRWviewHideShowStatus,
                        "All Drawing Views are in Show Mode",
                        "OK",
                        "Some of the Views are Hidden.")

                If TypeOfDrawing <> 5 Then
                    '22.Drawing View Line Type Check
                    CheckList.AddWarningBooleanCheckPoint(drawingProperties.DRWlinePropType,
                            "All Lines are of Correct Line Type",
                            "OK",
                            "Following Lines are Not as per Airbus Std. " & drawingProperties.DRWlinePropTypeLoc)
                    '23.Drawing View Line Thickness Check
                    CheckList.AddWarningBooleanCheckPoint(drawingProperties.DRWlinePropThk,
                            "All Lines are of Correct Line Thickness",
                            "OK",
                            "Following Lines are Not as per Airbus Std. " & drawingProperties.DRWlinePropThkLoc)
                End If

                '24.Drawing View Font Size Check
                CheckList.AddBooleanCheckPoint(drawingProperties.DRWtextSize,
                        "All Font Sizes are of Size  3.5 and 5.0 and 7.0",
                        "OK",
                        "Following Text's does Not have Correct Font Size " & drawingProperties.DRWtextSizeLoc)
                '25.Fake Dimensions used
                CheckList.AddWarningBooleanCheckPoint(drawingProperties.DRWVWfakedim,
                        "Fake Dimensions in Drawing",
                        "OK",
                        "Following Dimensions are Fake " & drawingProperties.DRWVWfakedimLocation)
                '26.Fake Hidden Dimensions used
                CheckList.AddWarningBooleanCheckPoint(drawingProperties.DRWVWfakedimHid,
                        "Dimension Texts Hidden to Fake Dimensions in Drawing",
                        "OK",
                        "Following Dimension Texts are Hidden to Fake " & drawingProperties.DRWVWfakedimHidLocation)
                '27.Dimension Precision is = 0.1
                CheckList.AddBooleanCheckPoint(drawingProperties.DRWVWdimPrec1,
                        "Dimension Precision of All Dimension is Equal to 0.1",
                        "OK",
                        "Following Dimensions Precision is Greater than 0.1 " & drawingProperties.DRWVWdimPrecLocation1)
                '28.Dimension Precision is = 1
                CheckList.AddWarningBooleanCheckPoint(drawingProperties.DRWVWdimPrec2,
                        "Dimension Precision is Equal to 1.0",
                        "OK",
                        "Following Dimensions Precision is is Equal to 1.0 " & drawingProperties.DRWVWdimPrecLocation2)
            End If

            '29.Sheet (Main View) Must be Empty
            CheckList.AddWarningBooleanCheckPoint(drawingProperties.DRWemptyMainView,
                    "Sheet.1 (Root of All Views) Must Not Contain any Elements or Texts",
                    "OK",
                    "There are Some Elements in 'Sheet.1 View' (Root of All Views)")

            '30.Baloon Size
            CheckList.AddBooleanCheckPoint(drawingProperties.BaloonSize,
                    "Baloon Size",
                    "OK",
                    "Following Baloons Size are not correct " & drawingProperties.WrongBaloonSizeText)

            '31.Flagnotes

            If _singlepart = True Then  'frmStart.optSingle.Visible = True

                CheckList.AddBooleanCheckPoint(drawingProperties.Flagnotesize,
            "All Flag Notes Sizes are of Size 5 ",
            "OK",
            "Following Flag Notes Size are not correct " & drawingProperties.WrongFlagnoteSizeText)

            Else
                CheckList.AddBooleanCheckPoint(drawingProperties.Flagnotesize,
            "All Flag Notes Sizes are of Size 3",
            "OK",
            "Following Flag Notes Size are not correct " & drawingProperties.WrongFlagnoteSizeText)

            End If

            'Set Other Details
            CheckList.DrawingNumber = drawingProperties.DRWhtzNumber & "-" & drawingProperties.DRWsheetNumber
            CheckList.DrawingName = "asdaw" 'PDM.DRWSETname
            CheckList.DrawingState = "dasd" 'PDM.DRWSHTstate
            CheckList.DrawingVersion = "dsfsafsa" 'PDM.DRWSHTversion
        Catch ex As Exception
            _log.Fatal(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name + "::" +
                                System.Reflection.MethodInfo.GetCurrentMethod.Name + "()", ex)
        End Try

    End Sub


    Public Function CheckFilledBOM(ByRef BOMcollection As Collection)
        Try
            Dim ptxt As ClsItemPDM
            Dim I As Integer

            I = 1
            BomItemNumInPDM = True
            For Each ptxt In BOMcollection
                If ptxt.ItemNumber = "" Or ptxt.ItemZone = "" Or ptxt.Quantity = 0 Then
                    BomItemNumInPDM = False
                    If I = 1 Then
                        BomItemNumInPDMErrMsg = "The Following BOM Items in PDM is not Filled " & "<" & ptxt.PartNumber & ">"
                        I = I + 1
                    Else
                        BomItemNumInPDMErrMsg = BomItemNumInPDMErrMsg & ", " & "<" & ptxt.PartNumber & ">"
                    End If
                End If
            Next
        Catch ex As Exception

        End Try
    End Function

    Public Function CompareIdentificationMarking()
        Try
            Dim idpdm As Integer
            IdentMarking1 = False
            IdentMarking2 = False
            'Check if Identification Marking is Present in PDM and Drawing
            If PDM2.PARTidentification = "" And drawingSinglePart.GENprop.TBLKidentMark1 = "" Then
                IdentMarking1 = True
                IdentMarking2 = False
                Exit Function
            ElseIf PDM2.PARTidentification = "" And drawingSinglePart.GENprop.TBLKidentMark1 <> "" Then
                IdentMarking1 = False
                IdentMarking2 = False
                Exit Function
            End If

            idpdm = Len(PDM2.PARTidentification)
            If idpdm = 7 Then
                idpdm = 4
            ElseIf idpdm = 5 Then
                idpdm = 3
            ElseIf idpdm = 3 Then
                idpdm = 2
            End If

            If idpdm = 1 Then
                If drawingSinglePart.GENprop.TBLKidentMark1 <> "" Then
                    If InStr(1, PDM2.PARTidentification, drawingSinglePart.GENprop.TBLKidentMark1) > 0 Then
                        IdentMarking1 = True
                        IdentMarking2 = True
                        Exit Function
                    End If
                End If
            End If
            If idpdm = 2 Then
                If drawingSinglePart.GENprop.TBLKidentMark1 <> "" And drawingSinglePart.GENprop.TBLKidentMark2 <> "" Then
                    If InStr(1, PDM2.PARTidentification, drawingSinglePart.GENprop.TBLKidentMark1) > 0 _
              And InStr(1, PDM2.PARTidentification, drawingSinglePart.GENprop.TBLKidentMark2) > 0 Then
                        IdentMarking1 = True
                        IdentMarking2 = True
                        Exit Function
                    End If
                End If
            End If
            If idpdm = 3 Then
                If drawingSinglePart.GENprop.TBLKidentMark1 <> "" And drawingSinglePart.GENprop.TBLKidentMark2 <> "" _
          And drawingSinglePart.GENprop.TBLKidentMark3 <> "" Then
                    If InStr(1, PDM2.PARTidentification, drawingSinglePart.GENprop.TBLKidentMark1) > 0 _
              And InStr(1, PDM2.PARTidentification, drawingSinglePart.GENprop.TBLKidentMark2) > 0 _
              And InStr(1, PDM2.PARTidentification, drawingSinglePart.GENprop.TBLKidentMark3) > 0 Then
                        IdentMarking1 = True
                        IdentMarking2 = True
                        Exit Function
                    End If
                End If
            End If
            If idpdm = 4 Then
                If drawingSinglePart.GENprop.TBLKidentMark1 <> "" And drawingSinglePart.GENprop.TBLKidentMark2 <> "" _
          And drawingSinglePart.GENprop.TBLKidentMark3 <> "" And drawingSinglePart.GENprop.TBLKidentMark4 <> "" Then
                    If InStr(1, PDM2.PARTidentification, drawingSinglePart.GENprop.TBLKidentMark1) > 0 _
              And InStr(1, PDM2.PARTidentification, drawingSinglePart.GENprop.TBLKidentMark2) > 0 _
              And InStr(1, PDM2.PARTidentification, drawingSinglePart.GENprop.TBLKidentMark3) > 0 _
              And InStr(1, PDM2.PARTidentification, drawingSinglePart.GENprop.TBLKidentMark4) > 0 Then
                        IdentMarking1 = True
                        IdentMarking2 = True
                        Exit Function
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Function

    Public Function CompareFlagNotes()
        Try
            Dim I As Long
            Dim j As Long
            Dim ii As Long
            Dim jj As Long
            Dim k As Integer
            Dim idwg As Boolean
            Dim ipdm As Boolean
            Dim isPresentPD() As Boolean
            Dim isPresentDG() As Boolean
            Dim fn As ClsText

            idwg = True
            ipdm = True
            FlagNoteErrorMsg = ""
            FlagNoteErrMsgDwg = ""
            ii = drawingSinglePart.FlagNoteList.Count
            If TypeName(PDM.DRWSETremarksList) <> "Empty" Then
                jj = UBound(PDM.DRWSETremarksList, 1)
            Else
                jj = 0
            End If

            If ii <= 0 Then idwg = False
            If jj <= 0 Then ipdm = False

            If idwg = False And ipdm = False Then
                FlagNoteErrorMsg = "NA"
                Exit Function
            End If
            If idwg = False Then
                FlagNoteErrorMsg = "There Are No Flag Notes in Drawing Sheet(s)."
                Exit Function
            ElseIf ipdm = False Then
                FlagNoteErrorMsg = "There Are No Flag Notes in PDM Link for this Drawing Set."
                Exit Function
            End If

            ReDim isPresentDG(ii)
            ReDim isPresentPD(jj)
            For I = 1 To ii
                isPresentDG(I) = False
            Next I
            For j = 1 To jj
                isPresentPD(j) = False
            Next j
            'Check Fn in PDM with Drawing
            For j = 1 To jj
                For Each fn In drawingSinglePart.FlagNoteList
                    If fn.TextContent = PDM.DRWSETremarksList(j, 1) Then
                        isPresentPD(j) = True
                        Exit For
                    End If
                Next
            Next j
            'Check Fn in Drawing with PDM
            I = 1
            For Each fn In drawingSinglePart.FlagNoteList
                For j = 1 To jj
                    If fn.TextContent = PDM.DRWSETremarksList(j, 1) Then
                        isPresentDG(I) = True
                        Exit For
                    End If
                Next j
                I = I + 1
            Next
            'Check Fn in PDM with Drawing
            FlagNoteExisting = True
            k = 0
            For j = 1 To jj
                If isPresentPD(j) = False Then
                    FlagNoteExisting = False
                    If k = 0 Then
                        FlagNoteErrorMsg = "Flag Note/s Missing in Drawing " & "<" & PDM.DRWSETremarksList(j, 1) & ">"
                        k = k + 1
                    Else
                        FlagNoteErrorMsg = FlagNoteErrorMsg & ", " & "<" & PDM.DRWSETremarksList(j, 1) & ">"
                    End If
                End If
            Next j

            'Check Fn in Drawing with PDM
            FlagNoteExistDwg = True
            k = 0
            For I = 1 To ii
                fn = drawingSinglePart.FlagNoteList.Item(I)
                If isPresentDG(I) = False Then
                    FlagNoteExistDwg = False
                    If k = 0 Then
                        FlagNoteErrMsgDwg = "Flag Note/s Missing in PDM Link " & "<" & fn.TextContent & " @ Zone " & fn.TextZone & ">"
                        k = k + 1
                    Else
                        FlagNoteErrMsgDwg = FlagNoteErrMsgDwg & ", " & "<" & fn.TextContent & " @ Zone " & fn.TextZone & ">"
                    End If
                End If
            Next I
        Catch ex As Exception

        End Try
    End Function

    Public Function CompareBomItemNumbers()
        Try

            Dim I As Integer
            Dim j As Integer
        Dim k As Integer
        Dim m As Integer
        Dim bomCount As Integer
        Dim ptxt As ClsItemPDM
        Dim dtxt As ClsText
        Dim isPresent() As Boolean
        Dim isQtyOk() As Boolean
        Dim found As Boolean

        j = 0
        For Each ptxt In PDM2.ASSYbomItems
            If ptxt.ManagedByCad = True Then
                j = j + 1
            End If
        Next
        bomCount = j

        ReDim isPresent(j)
        ReDim isQtyOk(j)
        For I = 1 To j
            isPresent(I) = False
            isQtyOk(I) = False
        Next I
        'Compare Each Item in PDM with Drawing
        I = 1
        For Each ptxt In PDM2.ASSYbomItems
            If ptxt.ManagedByCad = True Then
                For Each dtxt In drawingSinglePart.ItemNumberList
                    If ptxt.ItemNumber = dtxt.TextContent And dtxt.Qcheck = False Then
                        If ptxt.ItemZone = dtxt.TextZone Then
                            If dtxt.S16Qcheck = True Then isPresent(I) = True
                            If ptxt.Quantity = dtxt.ItemCount Then
                                isQtyOk(I) = True
                                dtxt.Qcheck = True
                                Exit For
                            Else
                                Exit For
                            End If
                        End If
                    End If
                Next
                I = I + 1
            End If
        Next

        'Check for Item Numbers
        BomItemNumbers = True
        m = 1
        k = 1
        For I = 1 To j
            If isPresent(I) = False Then
                BomItemNumbers = False
                For Each ptxt In PDM2.ASSYbomItems
                    If ptxt.ManagedByCad = True Then
                        If k = I Then
                            If m = 1 Then
                                BomItemNumErrMsg = "Error Found at Node & Item Number " & "<" & ptxt.PartNumber & " - " & ptxt.ItemNumber & ">"
                                m = m + 1
                            Else
                                BomItemNumErrMsg = BomItemNumErrMsg & ", " & "<" & ptxt.PartNumber & "-" & ptxt.ItemNumber & ">"
                            End If
                            Exit For
                        Else
                            k = k + 1
                        End If
                    End If
                Next
            End If
        Next I
        'Check for Quantity
        BomItemQty = True
        m = 1
        k = 1
        For I = 1 To j
            If isQtyOk(I) = False Then
                BomItemQty = False
                For Each ptxt In PDM2.ASSYbomItems
                    If ptxt.ManagedByCad = True Then
                        If k = I Then
                            If m = 1 Then
                                BomItemQtyErrMsg = "Quantity Does Not Match at Node & Item Number = Quantity " & "<" & ptxt.PartNumber & " - " & ptxt.ItemNumber & " = " & ptxt.Quantity & ">"
                                m = m + 1
                            Else
                                BomItemQtyErrMsg = BomItemQtyErrMsg & ", " & "<" & ptxt.PartNumber & " - " & ptxt.ItemNumber & " = " & ptxt.Quantity & ">"
                            End If
                            Exit For
                        Else
                            k = k + 1
                        End If
                    End If
                Next
            End If
        Next I
        'Check If Extra Item Numbers Exists in Drawing Sheet
        ExtraItemNumInDwgBol = False
        ExtraItemNumInDwgErrMsg = ""
        found = False
        I = 0
        If bomCount <> drawingSinglePart.ItemNumberList.Count Then
            For Each dtxt In drawingSinglePart.ItemNumberList
                If dtxt.Qcheck = False Then
                    For Each ptxt In PDM2.ASSYbomItems
                        If ptxt.ManagedByCad = False Then
                            If dtxt.TextContent = ptxt.ItemNumber Then
                                found = True
                                Exit For
                            End If
                        End If
                    Next
                    If found = True Then
                        found = False
                    Else
                        If I = 0 Then
                            ExtraItemNumInDwgErrMsg = "Item Number/s Not in PDM Link BOM " & "<" & dtxt.TextContent & " @ Zone " & dtxt.TextZone & ">"
                            I = I + 1
                        Else
                            ExtraItemNumInDwgErrMsg = ExtraItemNumInDwgErrMsg & ", " & "<" & dtxt.TextContent & " @ Zone " & dtxt.TextZone & ">"
                        End If
                    End If
                End If
            Next
        End If

        If ExtraItemNumInDwgErrMsg = "" Then
            ExtraItemNumInDwgBol = True
        Else
            ExtraItemNumInDwgBol = False
        End If

        Catch ex As Exception

        End Try
    End Function

    Public Function CompareINSTBomItemNumbers()
        Try

            Dim I As Integer
        Dim j As Integer
        Dim k As Integer
        Dim m As Integer
        Dim bomCount As Integer
        Dim ptxt As ClsItemPDM
        Dim dtxt As ClsText
        Dim isPresent() As Boolean
        Dim isQtyOk() As Boolean
        Dim found As Boolean

        j = 0
        For Each ptxt In PDM2.INSTbomItems
            If ptxt.ManagedByCad = True And ptxt.Context <> "Standard-Part" Then
                j = j + 1
            End If
        Next
        bomCount = j

        ReDim isPresent(j)
        ReDim isQtyOk(j)
        For I = 1 To j
            isPresent(I) = False
            isQtyOk(I) = False
        Next I
        'Compare Each Item in PDM with Drawing
        I = 1
        For Each ptxt In PDM2.INSTbomItems
            If ptxt.ManagedByCad = True And ptxt.Context <> "Standard-Part" Then
                For Each dtxt In drawingInstalDraw.ItemNumberList
                    'Exception for HOLE
                    If ptxt.ItemNumber = "HOLE" Then
                        isPresent(I) = True
                        isQtyOk(I) = True
                        Exit For
                    Else
                        If ptxt.ItemNumber = dtxt.TextContent And dtxt.Qcheck = False Then
                            If ptxt.ItemZone = dtxt.TextZone Then
                                isPresent(I) = True
                                If ptxt.Quantity = dtxt.ItemCount Then
                                    isQtyOk(I) = True
                                    dtxt.Qcheck = True
                                    Exit For
                                Else
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next
                I = I + 1
            End If
        Next

        k = PDM2.INSTbomItems.Count - j
        'Check for Item Numbers
        BomItemNumbers = True
        m = 1
        For I = 1 To j
            If isPresent(I) = False Then
                BomItemNumbers = False
                ptxt = PDM2.INSTbomItems.Item(k + I)
                If ptxt.ManagedByCad = True And ptxt.Context <> "Standard-Part" Then
                    If m = 1 Then
                        BomItemNumErrMsg = "Error Found at BFH Node & Item Number " & "<" & ptxt.PartNumber & " - " & ptxt.ItemNumber & " @ Zone " & ptxt.ItemZone & ">"
                        m = m + 1
                    Else
                        BomItemNumErrMsg = BomItemNumErrMsg & ", " & "<" & ptxt.PartNumber & " - " & ptxt.ItemNumber & " @ Zone " & ptxt.ItemZone & ">"
                    End If
                End If
            End If
        Next I
        'Check for Quantity
        BomItemQty = True
        m = 1
        For I = 1 To j
            If isQtyOk(I) = False Then
                BomItemQty = False
                ptxt = PDM2.INSTbomItems.Item(k + I)
                If ptxt.ManagedByCad = True Then
                    If m = 1 Then
                        BomItemQtyErrMsg = "Quantity Does Not Match at BFH Node & Item Number = Quantity " & "<" & ptxt.PartNumber & " - " & ptxt.ItemNumber & " = " & ptxt.Quantity & ">"
                        m = m + 1
                    Else
                        BomItemQtyErrMsg = BomItemQtyErrMsg & ", " & "<" & ptxt.PartNumber & " - " & ptxt.ItemNumber & " = " & ptxt.Quantity & ">"
                    End If
                End If
            End If
        Next I
        'Check If Extra Item Numbers Exists in Drawing Sheet
        ExtraItemNumInDwgBol = False
        ExtraItemNumInDwgErrMsg = ""
        found = False
        I = 0
        If bomCount <> drawingInstalDraw.ItemNumberList.Count Then
            For Each dtxt In drawingInstalDraw.ItemNumberList
                If dtxt.Qcheck = False Then
                    For Each ptxt In PDM2.INSTbomItems
                        If ptxt.ManagedByCad = False Then
                            If dtxt.TextContent = ptxt.ItemNumber Then
                                found = True
                                Exit For
                            End If
                        End If
                    Next
                    If found = True Then
                        found = False
                    Else
                        If I = 0 Then
                            ExtraItemNumInDwgErrMsg = "Item Number/s Not in PDM Link BOM " & "<" & dtxt.TextContent & " @ Zone " & dtxt.TextZone & ">"
                            I = I + 1
                        Else
                            ExtraItemNumInDwgErrMsg = ExtraItemNumInDwgErrMsg & ", " & "<" & dtxt.TextContent & " @ Zone " & dtxt.TextZone & ">"
                        End If
                    End If
                End If
            Next
        End If

            If ExtraItemNumInDwgErrMsg = "" Then
                ExtraItemNumInDwgBol = True
            Else
                ExtraItemNumInDwgBol = False
            End If

        Catch ex As Exception

        End Try
    End Function
    Public Function CompareINSTflagNotes()
        Try
            Dim I, j, ii, jj As Long
            Dim k As Integer
            Dim idwg, ipdm As Boolean
            Dim isPresentPD() As Boolean
            Dim isPresentDG() As Boolean
            Dim fn As ClsText

            idwg = True
            ipdm = True
            ii = drawingInstalDraw.FlagNoteList.Count
            If TypeName(PDM2.DRWSETremarksList) <> "Empty" Then
                jj = UBound(PDM2.DRWSETremarksList, 1)
            Else
                jj = 0
            End If

            If ii <= 0 Then idwg = False
            If jj <= 0 Then ipdm = False

            If idwg = False And ipdm = False Then
                FlagNoteErrorMsg = "NA"
                Exit Function
            End If
            If idwg = False Then
                FlagNoteErrorMsg = "There Are No Flag Notes in Drawing Sheet(s)."
                FlagNoteErrMsgDwg = ""
                Exit Function
            ElseIf ipdm = False Then
                FlagNoteErrorMsg = "There Are No Flag Notes in PDM Link for this Drawing Set."
                For I = 1 To ii
                    fn = drawingInstalDraw.FlagNoteList.Item(I)
                    If I = 1 Then
                        FlagNoteErrMsgDwg = "Flag Note/s Missing in PDM Link " & ", <" & fn.TextContent & "@ Zone " & fn.TextZone & "> "
                    Else
                        FlagNoteErrMsgDwg = FlagNoteErrMsgDwg & ", <" & fn.TextContent & "@ Zone " & fn.TextZone & "> "
                    End If
                Next I
                Exit Function
            End If

            ReDim isPresentDG(ii)
            ReDim isPresentPD(jj)
            For I = 1 To ii
                isPresentDG(I) = False
            Next I
            For j = 1 To jj
                isPresentPD(j) = False
            Next j
            'Check Fn in PDM with Drawing
            For j = 1 To jj
                For Each fn In drawingInstalDraw.FlagNoteList
                    If fn.TextContent = PDM2.DRWSETremarksList(j, 1) Then
                        isPresentPD(j) = True
                        Exit For
                    End If
                Next
            Next j
            'Check Fn in Drawing with PDM
            I = 1
            For Each fn In drawingInstalDraw.FlagNoteList
                For j = 1 To jj
                    If fn.TextContent = PDM2.DRWSETremarksList(j, 1) Then
                        isPresentDG(I) = True
                        Exit For
                    End If
                Next j
                I = I + 1
            Next
            'Check Fn in PDM with Drawing
            FlagNoteExisting = True
            k = 0
            For j = 1 To jj
                If isPresentPD(j) = False Then
                    FlagNoteExisting = False
                    If k = 0 Then
                        FlagNoteErrorMsg = "Flag Note/s Missing in Drawing " & "<" & PDM2.DRWSETremarksList(j, 1) & ">"
                        k = k + 1
                    Else
                        FlagNoteErrorMsg = FlagNoteErrorMsg & ", " & "<" & PDM2.DRWSETremarksList(j, 1) & ">"
                    End If
                End If
            Next j

            'Check Fn in Drawing with PDM
            FlagNoteExistDwg = True
            k = 0
            For I = 1 To ii
                fn = drawingInstalDraw.FlagNoteList.Item(I)
                If isPresentDG(I) = False Then
                    FlagNoteExistDwg = False
                    If k = 0 Then
                        FlagNoteErrMsgDwg = "Flag Note/s Missing in PDM Link " & "<" & fn.TextContent & " @ Zone " & fn.TextZone & ">"
                        k = k + 1
                    Else
                        FlagNoteErrMsgDwg = FlagNoteErrMsgDwg & ", " & "<" & fn.TextContent & " @ Zone " & fn.TextZone & ">"
                    End If
                End If
            Next I
        Catch ex As Exception

        End Try
    End Function

    Public Sub CompareIPtableBOM()
        Try
            Dim I, j, k As Integer
            Dim row As Long
            'Dim col As Long
            Dim itm As ClsItemPDM
            Dim itmCol As Collection
            Dim isPresent As Object
            Dim fchk, pass As Boolean
            Dim tempIPnum As String

            itmCol = PDM2.INSTbomItems
            row = UBound(drawingInstalDraw.IPtable, 1)
            I = 0
            For Each itm In itmCol
                If itm.ManagedByCad = True Then
                    If InStr(1, itm.PartNumber, "-BFH") > 0 Then
                        I = I + 1
                    End If
                End If
            Next
            ReDim isPresent(I)
            j = 1
            For Each itm In itmCol
                If itm.ManagedByCad = True Then
                    If InStr(1, itm.PartNumber, "-BFH") > 0 Then
                        For I = 3 To row
                            pass = False
                            If drawingInstalDraw.IPtable(I, 1) = itm.IPnumber1 Then
                                If drawingInstalDraw.IPtable(I, 2) = itm.ItemNumber Then
                                    If drawingInstalDraw.IPtable(I, 3) = itm.ItemZone Then
                                        isPresent(j) = True
                                        fchk = True
                                        pass = True
                                        If itm.IPnumber2 = "" Then Exit For
                                    End If
                                End If
                            ElseIf drawingInstalDraw.IPtable(I, 1) = itm.IPnumber2 And itm.IPnumber2 <> "" Then
                                If drawingInstalDraw.IPtable(I, 2) = itm.ItemNumber Then
                                    If drawingInstalDraw.IPtable(I, 3) = itm.ItemZone Then
                                        isPresent(j) = True
                                        fchk = False
                                        pass = True
                                        'If itm.IPnumber1 = "" Then Exit For
                                    End If
                                End If
                            End If
                            If pass = True Then
                                pass = False
                                If fchk = True Then
                                    tempIPnum = itm.IPnumber2
                                Else
                                    tempIPnum = itm.IPnumber1
                                End If
                                If drawingInstalDraw.IPtable(I + 1, 1) = tempIPnum Then
                                    If drawingInstalDraw.IPtable(I, 2) = itm.ItemNumber Then
                                        If drawingInstalDraw.IPtable(I, 3) = itm.ItemZone Then
                                            isPresent(j) = True
                                            'Skip 1 Row
                                            I = I + 1
                                            Exit For
                                        End If
                                    End If
                                End If
                                I = I + 1
                                isPresent(j) = False
                                Exit For
                            End If
                        Next I
                        j = j + 1
                    End If
                End If
            Next
            IPtableAndPDMBol = True
            k = 1
            I = 1
            For Each itm In itmCol
                If itm.ManagedByCad = True Then
                    If InStr(1, itm.PartNumber, "-BFH") > 0 Then
                        If isPresent(I) = False Then
                            IPtableAndPDMBol = False
                            If k = 1 Then
                                IPtableAndPDMErrMsg = "MissMatch / Error Found at Node in PDM Link BOM " & "<" & itm.PartNumber & ">"
                                k = k + 1
                            Else
                                IPtableAndPDMErrMsg = IPtableAndPDMErrMsg & ", " & "<" & itm.PartNumber & ">"
                            End If
                        End If
                        I = I + 1
                    End If
                End If
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ConsolidateItemNumbersS16()
        Try
            Dim txt As ClsText
            Dim dtxt As ClsText
            Dim col As Collection ' modified by sai 
            Dim I As Long
            Dim isExisting As Boolean
            Dim bomItem As ClsItemPDM

            I = 1
            isExisting = False
            For Each dtxt In drawingInstalDraw.ItemNumberList
                If I > 1 Then
                    For Each txt In col
                        If dtxt.TextContent = txt.TextContent Then
                            isExisting = True
                            Exit For
                        End If
                    Next
                    If isExisting = False Then
                        col.Add(dtxt)
                    Else
                        isExisting = False
                    End If
                Else
                    col.Add(dtxt)
                End If
                I = I + 1
            Next

            txt = Nothing
            dtxt = Nothing

            '/////////////////
            'Assign the PDM Zones
            For Each txt In col
                For Each bomItem In PDM2.INSTbomItems
                    If txt.TextContent = bomItem.ItemNumber Then
                        txt.S16TextZone = bomItem.ItemZone
                        Exit For
                    End If
                Next
            Next

            'Check the fucking Drawing Sheet Zones
            For Each dtxt In col
                dtxt.S16Qcheck = False
                For Each txt In drawingInstalDraw.ItemNumberList
                    If dtxt.TextContent = txt.TextContent Then
                        If dtxt.S16TextZone = txt.TextZone Then
                            dtxt.S16Qcheck = True
                            Exit For
                        End If
                    End If
                Next
            Next

            txt = Nothing
            dtxt = Nothing

            I = 1
            For Each dtxt In drawingInstalDraw.ItemNumberList
                If I > 1 Then
                    For Each txt In col
                        If dtxt.TextContent = txt.TextContent Then
                            dtxt.TextZone = txt.S16TextZone
                            Exit For
                        End If
                    Next
                End If
                I = I + 1
            Next
        Catch ex As Exception

        End Try
    End Sub
    Private Sub SortDrawingDocsAsc(ByRef tmpDocs As Collection)
        Try
            Dim tdoc As DrawingDocument
            Dim I As Integer
            Dim j As Integer
            Dim sss
            Dim sss1

            For I = 1 To tmpDocs.Count - 1
                sss = Split(tmpDocs.Item(I).Sheets.Name, "\")
                For j = I + 1 To tmpDocs.Count
                    sss1 = Split(tmpDocs.Item(I).Sheets.Name, "\")
                    If CInt(Mid(sss(UBound(sss)), 11, 2)) > CInt(Mid(sss1(UBound(sss1)), 11, 2)) Then

                        tdoc = tmpDocs.Item(j)
                        tmpDocs.Remove(j)
                        tmpDocs.Add(tdoc, Before:=I)

                    End If
                Next j
            Next I
        Catch ex As Exception

        End Try
    End Sub

End Module
