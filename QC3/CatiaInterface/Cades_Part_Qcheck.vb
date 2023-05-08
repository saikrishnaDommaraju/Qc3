Imports INFITF
Imports KnowledgewareTypeLib
Imports MECMOD
Imports SPATypeLib

Public Class Cades_Part_Qcheck

    Public RUN_PART As Boolean
    Public TYPE_OF_PART As Integer  '1=Sht, 2=Mach, 3=Comp, 4=PPS
    Public PDM3D As ClsPDMLink3D

    Dim PART_Weight As Double
    Dim PART_L As Double
    Dim PART_W As Double
    Dim PART_T As Double
    Dim PART_MainBodyName As String
    Dim PART_MAT_TYPE As Integer    '1=Al, 2=Ti, 3=CFRP, 4=PPS, 5=CFRI, 6=AlLi

    Dim PART_Material As ClsMaterialProp
    Dim MatListCol As Collection  '= New Collection need to intialize 

    'CheckList Items
    Dim PARTbolColor As Boolean '
    Dim PARTcomColor As String '
    Dim PARTbolLayer As Boolean '
    Dim PARTcomLayer As String '
    Dim PARTbolUBColor As Boolean '
    Dim PARTcomUBColor As String '
    Dim PARTbolUBLayer As Boolean '
    Dim PARTcomUBLayer As String '
    Dim PARTbolProtCode As Boolean '
    Dim PARTcomProtCode As String '
    Dim PARTbolQty As Boolean '
    Dim PARTcomQty As String '
    Dim PARTbolMI As Boolean
    Dim PARTcomMI As String
    Dim PARTbolIZ As Boolean '
    Dim PARTcomIZ As String '
    Dim PARTbolMrk As Boolean '
    Dim PARTcomMrk As String '
    Dim PARTbolID As Boolean '
    Dim PARTcomID As String '
    Dim PARTbolFC As Boolean '
    Dim PARTcomFC As String '
    Dim PARTbolMS As Boolean '
    Dim PARTcomMS As String '
    Dim PARTbolLEN As Boolean '
    Dim PARTcomLEN As String '
    Dim PARTbolWID As Boolean '
    Dim PARTcomWID As String '

    Sub CATMAIN()

        Dim docOpen As Document
        Dim doc As PartDocument
        Dim UnBody As Body
        Dim sel As Selection
        Dim lyr As Long
        Dim pars As Parameters
        Dim par As Parameter
        Dim p1 As Parameter, p2 As Parameter, p3 As Parameter
        Dim iner As Inertia
        Dim mBodyName As String
        Dim rootName As String
        Dim str As String
        Dim I As Integer
        Dim R As Long, G As Long, B As Long
        Dim mas As Double
        Dim isSheetmetal As Boolean
        Dim var

        'To Change the CheckList Headings
        RUN_PART = True

        '''iExit = False
        'To Check the Authourization
        '''Call CheckAuthourization

        '''If iExit = True Then Exit Sub

        docOpen = CATIA.ActiveDocument
        'Test the Document
        If TypeName(docOpen) <> "PartDocument" Then
            MsgBox("Wrong Document Type, Open a Part and Then Run This Macro...", vbCritical, "Wrong Document Type...")
            Exit Sub
        End If

        '*************
        'frmStartPrt.Show ' need to check the form 
        '*************
        If iExit = True Then Exit Sub

        doc = docOpen
        docOpen = Nothing
        PART_Material = New ClsMaterialProp
        sel = doc.Selection
        sel.Clear()

        '*************
        frmStatus.Show()
        ShowStatus("Reading Part Data From CATIA...", 10, True)
        '*************

        'Fill Material Data for comparision
        Call FillMatListCol()

        mBodyName = doc.Part.MainBody.Name

        '1. CHECK IF SHEETMETAL OR NOT
        '-----------------------------
        On Error Resume Next

        Call ChecktheifSheetMetalOrMilled(doc.Part)

        '''''        If doc.Part.SheetMetalParameters Is Nothing Then
        '''''            PART_Material.isSheetmetal = False
        '''''>>>>>>>>> Comment By Rana
        '''''        Else
        '''''            PART_Material.isSheetmetal = True
        '''''        End If

        '2. CHECK THE APPLIED MATERIAL TO MAINBODY
        '-----------------------------------------
        pars = doc.Part.Parameters
        rootName = pars.Parent.Name

        par = pars.GetItem(mBodyName & "\Material")
        PART_Material.Name = UCase(par.ValueAsString)

        '3. LAYER AND COLOR OF MAINBODY
        '------------------------------
        sel.Add(doc.Part.MainBody)
        'sel.VisProperties.GetLayer(catVisLayerBasic, lyr)
        sel.VisProperties.GetRealColor(R, G, B)

        PART_Material.Layer = lyr
        PART_Material.R = R
        PART_Material.G = G
        PART_Material.B = B

        sel.Clear()

        '4.CHECK LAYER AND COLOR OF UNFOLDED BODY IF SHEETMETAL
        '------------------------------------------------------
        If PART_Material.isSheetmetal = True Then
            If TYPE_OF_PART = 1 Or TYPE_OF_PART = 4 Then
                On Error GoTo NOUNFOLDED
                sel.Add(doc.Part.Bodies.GetItem("Unfolded"))

                'sel.VisProperties.GetLayer(VisPropertySet.catVisLayerBasic, lyr)
                sel.VisProperties.GetRealColor(R, G, B)

                PART_Material.UnfoldedLayer = lyr
                PART_Material.UnfoldedR = R
                PART_Material.UnfoldedG = G
                PART_Material.UnfoldedB = B

                sel.Clear()
            End If
        End If

        '*************
        ShowStatus("Reading Part Data From CATIA...", 30, True)
        '*************

        On Error GoTo 0

        '5. MEASURE PART WEIGHT
        '----------------------
        iner = GetBodyInertia(doc.Part, doc.Part.MainBody)
        If iner Is Nothing Then
            MsgBox("Inertia Method Failed...", vbCritical, "Inertia Failed")
            Exit Sub
        Else
            mas = Math.Round(iner.Mass, 3)
            PART_Weight = CDbl(mas)
        End If

        '6. MEASURE LWH FOR MAINBODY
        '---------------------------
        If PART_Material.isSheetmetal = False Then
            sel.Add(doc.Part.MainBody)
            CATIA.StartCommand("Measure Inertia")

            pars = Nothing
            pars = doc.Part.Parameters
            par = pars.Item(pars.Count)
            var = CStr(par.Name)
            var = Split(var, "\")

            p1 = pars.GetItem(rootName & "\" & var(1) & "\BBLx")
            p2 = pars.GetItem(rootName & "\" & var(1) & "\BBLy")
            p3 = pars.GetItem(rootName & "\" & var(1) & "\BBLz")

            PART_L = CDbl(p1.Value)
            PART_W = CDbl(p2.Value)
            PART_T = Math.Round(CDbl(p3.Value), 3)

            '        MsgBox "LENGTH = " & PART_L & vbCrLf & _
            '            "WIDTH = " & PART_W & vbCrLf & _
            '            "THICKNESS = " & PART_T & vbCrLf
            '        Exit Sub


            CATIA.ActiveWindow.Close()
            CATIA.StartCommand("Select")
            sel.Clear()
        End If

        '7. MEASURE LWH FOR UNFOLDED BODY
        '--------------------------------
        If PART_Material.isSheetmetal = True Then
            sel.Add(doc.Part.Bodies.GetItem("Unfolded"))
            CATIA.StartCommand("Measure Inertia")

            pars = Nothing
            pars = doc.Part.Parameters
            par = pars.Item(pars.Count)
            var = CStr(par.Name)
            var = Split(var, "\")

            p1 = pars.GetItem(rootName & "\" & var(1) & "\BBLx")
            p2 = pars.GetItem(rootName & "\" & var(1) & "\BBLy")
            p3 = pars.GetItem(rootName & "\" & var(1) & "\BBLz")

            PART_L = CDbl(p1.Value)
            PART_W = CDbl(p2.Value)
            PART_T = Math.Round(CDbl(p3.Value), 3)

            CATIA.ActiveWindow.Close()
            CATIA.StartCommand("Select")
            sel.Clear()
        End If

        ShowStatus("Reading Part Data From PDM Link...", 50, True)
        'PDM Link Data Extract Start....
        Call PDMLink3DExtract()

        If iExit = True Then Exit Sub
        ShowStatus("Reading Part Data From PDM Link...", 80, True)
        '*************

        'Sort the L W T based on Values
        SortLWT()
        'Compare Part Materail with Database
        GetPartMatType()
        'Checking & Comparision Functions
        CheckPDMPartProperties()


        '*************
        ShowStatus("Writing Check List...", 95, True)
        '*************

        '*******************
        'Fill the CheckList
        '*******************
        'New Installation CheckList
        CheckList = New clsCHECKList
        'Set Other CheckList Details
        CheckList.DrawingNumber = PDM3D.PART_Number
        CheckList.DrawingName = PDM3D.PARTName
        CheckList.DrawingState = PDM3D.PART_State
        CheckList.DrawingVersion = PDM3D.PART_Version

        '1.PartBody Color
        CheckList.AddBooleanCheckPoint(PARTbolColor,
                    "Part Body (Main Body) Color",
                    "OK",
                    PARTcomColor)
        '2.PartBody Layer
        CheckList.AddBooleanCheckPoint(PARTbolLayer,
                    "Part Body (Main Body) Layer",
                    "OK",
                    PARTcomLayer)
        If TYPE_OF_PART = 1 Or TYPE_OF_PART = 4 Then
            '3.Unfolded Body Color
            CheckList.AddBooleanCheckPoint(PARTbolUBColor,
                        "UnFolded Body Color",
                        "OK",
                        PARTcomUBColor)
            '4.Unfolded Body Layer
            CheckList.AddBooleanCheckPoint(PARTbolUBLayer,
                        "UnFolded Body Layer",
                        "OK",
                        PARTcomUBLayer)
        End If
        '5.Main Body Name should be PartBody
        CheckList.AddWarningCompareCheckPoint(mBodyName,
                    "PartBody",
                    "Main Body Name should be 'PartBody'",
                    "OK",
                    "Main Body Name not as per Standard. ")
        '6.Compare PDM Domestic Title and Part Name
        CheckList.AddCompareCheckPoint(PDM3D.PARTDomesticTitle,
                    PDM3D.PARTName,
                    "The 'Domestic Title' and 'Name' in PDMLink must be same",
                    "OK",
                    "Mismatch in Domestic Title and Name in PDMLink. ")
        '7.Units of Qty Part in PDM Link
        ''@WIP        CheckList.AddBooleanCheckPoint (PARTbolQty, _
        '"Unit of Quantity of Material Part in PDM Link", _
        '"OK", _
        'PARTcomQty)
        '8.Compare PDM Weight and CATIA Weight
        ''@WIP           CheckList.AddCompareCheckPoint (PDM3D.PARTWeight, _
        'PART_Weight,
        '            "Weight in PDMLink and CATIA Weight must be same",
        '            "OK",
        '            "Mismatch in Weights. ")
        '9.Compare Thickness from PDM and CATIA
        ''@WIP          CheckList.AddCompareCheckPoint( PDM3D.PARTThk, _
        'PART_T,
        '            "Part Thickness in PDMLink and CATIA Comparision",
        '            "OK",
        '            "Mismatch in Thickness. ")
        '10.Compare Length from PDM and CATIA
        ''@WIP           CheckList.AddBooleanCheckPoint( PARTbolLEN, _
        '"Part Length in PDMLink and CATIA Comparision", _
        '"OK", _
        'PARTcomLEN)
        '11.Compare Width from PDM and CATIA
        ''@WIP           CheckList.AddBooleanCheckPoint (PARTbolWID, _
        '"Part Width in PDMLink and CATIA Comparision", _
        '"OK", _
        'PARTcomWID)
        '12.Check Protective Treatment Code in PDM
        CheckList.AddBooleanCheckPoint(PARTbolProtCode,
                    "Check Protective Treatment Code in PDMLink",
                    "OK",
                    PARTcomProtCode)
        '13.Check Manufacturing Index in PDM
        CheckList.AddBooleanCheckPoint(PARTbolMI,
                    "Check Manufacturing Index in PDMLink",
                    "OK",
                    PARTcomMI)
        '14.Item Zone in PDM Link
        CheckList.AddBooleanCheckPoint(PARTbolIZ,
                    "Item Zone in PDM Link is Filled*** (Verify if the Parameter is Correct)",
                    "OK",
                    PARTcomIZ)
        '15.Marking in PDM Link
        CheckList.AddBooleanCheckPoint(PARTbolMrk,
                    "Part Marking in PDM Link is Filled*** (Verify if the Parameter is Correct)",
                    "OK",
                    PARTcomMrk)
        '16.FC in PDM Link
        CheckList.AddBooleanCheckPoint(PARTbolFC,
                    "Functional Class of Part in PDM Link is Filled*** (Verify if the Parameter is Correct)",
                    "OK",
                    PARTcomFC)
        '17.MS in PDM Link
        CheckList.AddBooleanCheckPoint(PARTbolMS,
                    "Material Structure of Part in PDM Link is Filled*** (Verify if the Parameter is Correct)",
                    "OK",
                    PARTcomMS)
        '18.Identifiable Part in PDM Link
        CheckList.AddBooleanCheckPoint(PARTbolID,
                    "Identifiable Part in PDM Link is Filled*** (Verify if the Parameter is Correct)",
                    "OK",
                    PARTcomID)

        '*******************

        '*************
        ShowStatus("Finalizing Quality Check...", 100, True)
        frmStatus.Hide()

        '*************

        'Fills the CheckReport
        Dim genfun As Sys_Fun = New Sys_Fun
        genfun.FillSingleCheckReport()

        'Normal Exit
        Exit Sub

        'Error Handling
NOUNFOLDED:
        '*************
        ShowStatus("Aborting Due to Error...", 100, True)
        frmStatus.Hide()
        '*************
        MsgBox("Unfolded Body Not Found..." & vbCrLf & "Verify if the Unfolded Body's Name is Correct (i.e. 'Unfolded').", vbCritical, "Wrong Option Selected...")
    End Sub

    Function ChecktheifSheetMetalOrMilled(ByRef oPart As Part)

        Dim I As Integer
        Dim oPara As Parameter
        For I = 1 To oPart.Parameters.Count
            oPara = oPart.Parameters.Item(I)
            If InStr(oPara.Name, "Sheet Metal Parameter") > 0 Then
                PART_Material.isSheetmetal = True
                Exit For
            End If
        Next

    End Function

    'Function For Getting Inertia of a Body
    Function GetBodyInertia(ByRef iPart As Part, ByRef iBody As Body) As Inertia
        Dim objSPAWorkbench As Workbench
        Dim objInertia As Inertia

        On Error Resume Next
        objSPAWorkbench = iPart.Parent.GetWorkbench("SPAWorkbench")
        objInertia = objSPAWorkbench.Inertias.Add(iBody)

        If Err.Number = 0 Then
            GetBodyInertia = objInertia
        Else
            GetBodyInertia = Nothing
        End If
    End Function

    '-------------------
    'PDM LINK FUNCTIONS
    '-------------------
    Sub PDMLink3DExtract()

        Dim I As Long
        Dim j As Long
        Dim partDoc As PartDocument
        Dim params As Parameters
        Dim rootName As String

        PDM3D = New ClsPDMLink3D
        partDoc = CATIA.ActiveDocument
        params = partDoc.Part.Parameters

        rootName = params.Parent.Name

        If Not PDM3D.GetPartProperties(Left(rootName, Len(rootName) - 2)) Then
            MsgBox(rootName & "Part Not Found...", vbCritical)
            Exit Sub
            iExit = True
        End If
    End Sub


    'Fill Material Database
    Sub FillMatListCol()
        Dim I As Integer
        Dim mat As ClsMaterialProp
        Dim nam As Object
        Dim lay As Object
        Dim Re As Object
        Dim Gr As Object
        Dim Bl As Object
        Dim PMT As Object

        nam = {"TITANIUM_4-55", "ALU-LITHIUM_2-66", "ALUMINUM_2-85", "ALUMINIUM", "TITANIUM", "AL", "TI", "CFK-1.58", "CFK-1.58", "CFK-RI-1.55", "CFK-1.59"}
        lay = {"134", "132", "131", "131", "134", "131", "134", "124", "126", "125", "126"}
        Re = {"0", "128", "0", "0", "0", "0", "0", "128", "0", "255", "0"}
        Gr = {"255", "0", "128", "128", "255", "128", "255", "64", "128", "255", "128"}
        Bl = {"255", "255", "255", "255", "255", "255", "255", "64", "0", "170", "0"}
        '1=Al, 2=Ti, 3=CFRP,PPS,CFRI,
        PMT = {"2", "1", "1", "1", "2", "1", "2", "3", "3", "3", "3"}

        For I = 0 To UBound(nam)
            mat = New ClsMaterialProp
            mat.Name = nam(I)
            mat.Layer = lay(I)
            mat.R = Re(I)
            mat.G = Gr(I)
            mat.B = Bl(I)
            mat.PART_MAT_TYPE = PMT(I)
            MatListCol.Add(mat)
        Next I
    End Sub

    '***********************************************
    'WIP
    '***********************************************
    'Get the Enum of Material Type
    Sub GetPartMatType()

        Dim I As Integer
        Dim mat As ClsMaterialProp
        Dim match As Boolean
        Dim varSht As Object
        Dim varMach As Object
        Dim varComp As Object
        Dim varPPS As Object

        varSht = {0, 2, 3, 4, 5, 6}
        varMach = {0, 1, 2, 3, 4, 5, 6}
        varComp = {7, 9}
        varPPS = {8, 10}

        match = False
        '1=Sht, 2=Mach, 3=Comp, 4=PPS
        If TYPE_OF_PART = 1 Then
            For I = 0 To UBound(varSht)
                mat = MatListCol.Item(CLng(varSht(I)) + 1)
                If PART_Material.Name = mat.Name Then
                    match = True
                    If PART_Material.R = mat.R Then
                        If PART_Material.G = mat.G Then
                            If PART_Material.B = mat.B Then
                                PARTbolColor = True
                                PARTcomColor = "OK"
CHECKLAYER:
                                'PART_MAT_TYPE is assigned
                                PART_Material.PART_MAT_TYPE = mat.PART_MAT_TYPE
                                If PART_Material.Layer = mat.Layer Then
                                    PARTbolLayer = True
                                    PARTcomLayer = "OK"
                                    GoTo UNFOLDEDCLR
                                Else
                                    PARTbolLayer = False
                                    PARTcomLayer = "The Layer of the Body should be " & mat.Layer
                                    GoTo UNFOLDEDCLR
                                End If
                            Else
                                GoTo FALSECOLOR
                            End If
                        Else
                            GoTo FALSECOLOR
                        End If
                    Else
FALSECOLOR:
                        PARTbolColor = False
                        PARTcomColor = "The PartBody Color should be R=" & mat.R & " G=" & mat.G & " B=" & mat.B
                        GoTo CHECKLAYER
                    End If
UNFOLDEDCLR:
                    If PART_Material.UnfoldedR = mat.R Then
                        If PART_Material.UnfoldedG = mat.G Then
                            If PART_Material.UnfoldedB = mat.B Then
                                PARTbolUBColor = True
                                PARTcomUBColor = "OK"
CHECKUBLAYER:
                                'PART_MAT_TYPE is assigned
                                PART_Material.PART_MAT_TYPE = mat.PART_MAT_TYPE
                                If PART_Material.UnfoldedLayer = mat.Layer Or PART_Material.UnfoldedLayer = 0 Then
                                    PARTbolUBLayer = True
                                    PARTcomUBLayer = "OK"
                                    Exit Sub
                                Else
                                    PARTbolUBLayer = False
                                    PARTcomUBLayer = "The Layer of the Body should be " & mat.Layer & " OR 0"
                                    Exit Sub
                                End If
                            Else
                                GoTo FALSECOLORUB
                            End If
                        Else
                            GoTo FALSECOLORUB
                        End If
                    Else
FALSECOLORUB:
                        PARTbolUBColor = False
                        PARTcomUBColor = "The PartBody Color should be R=" & mat.R & " G=" & mat.G & " B=" & mat.B
                        GoTo CHECKUBLAYER
                    End If
                End If
                If match = True Then Exit Sub
            Next I

            '1=Sht, 2=Mach, 3=Comp, 4=PPS
        ElseIf TYPE_OF_PART = 2 Then
            For I = 0 To UBound(varMach)
                mat = MatListCol.Item(CLng(varMach(I)) + 1)
                If PART_Material.Name = mat.Name Then
                    match = True
                    If PART_Material.R = mat.R Then
                        If PART_Material.G = mat.G Then
                            If PART_Material.B = mat.B Then
                                PARTbolColor = True
                                PARTcomColor = "OK"
CHECKLAYER2:
                                'PART_MAT_TYPE is assigned
                                PART_Material.PART_MAT_TYPE = mat.PART_MAT_TYPE
                                If PART_Material.Layer = mat.Layer Then
                                    PARTbolLayer = True
                                    PARTcomLayer = "OK"
                                    Exit Sub
                                Else
                                    PARTbolLayer = False
                                    PARTcomLayer = "The Layer of the Body should be " & mat.Layer
                                    Exit Sub
                                End If
                            Else
                                GoTo FALSECOLOR2
                            End If
                        Else
                            GoTo FALSECOLOR2
                        End If
                    Else
FALSECOLOR2:
                        PARTbolColor = False
                        PARTcomColor = "The PartBody Color should be R=" & mat.R & " G=" & mat.G & " B=" & mat.B
                        GoTo CHECKLAYER2
                    End If
                End If
                If match = True Then Exit Sub
            Next I

            '1=Sht, 2=Mach, 3=Comp, 4=PPS
        ElseIf TYPE_OF_PART = 3 Then
            For I = 0 To UBound(varComp)
                mat = MatListCol.Item(CLng(varComp(I)) + 1)
                If PART_Material.Name = mat.Name Then
                    match = True
                    If PART_Material.R = mat.R Then
                        If PART_Material.G = mat.G Then
                            If PART_Material.B = mat.B Then
                                PARTbolColor = True
                                PARTcomColor = "OK"
CHECKLAYER3:
                                'PART_MAT_TYPE is assigned
                                PART_Material.PART_MAT_TYPE = mat.PART_MAT_TYPE
                                If PART_Material.Layer = mat.Layer Then
                                    PARTbolLayer = True
                                    PARTcomLayer = "OK"
                                    Exit Sub
                                Else
                                    PARTbolLayer = False
                                    PARTcomLayer = "The Layer of the Body should be " & mat.Layer
                                    Exit Sub
                                End If
                            Else
                                GoTo FALSECOLOR3
                            End If
                        Else
                            GoTo FALSECOLOR3
                        End If
                    Else
FALSECOLOR3:
                        PARTbolColor = False
                        PARTcomColor = "The PartBody Color should be R=" & mat.R & " G=" & mat.G & " B=" & mat.B
                        GoTo CHECKLAYER3
                    End If
                End If
                If match = True Then Exit Sub
            Next I
            '1=Sht, 2=Mach, 3=Comp, 4=PPS
        ElseIf TYPE_OF_PART = 4 Then
            For I = 0 To UBound(varPPS)
                mat = MatListCol.Item(CLng(varPPS(I)) + 1)
                If PART_Material.Name = mat.Name Then
                    match = True
                    If PART_Material.R = mat.R Then
                        If PART_Material.G = mat.G Then
                            If PART_Material.B = mat.B Then
                                PARTbolColor = True
                                PARTcomColor = "OK"
CHECKLAYER4:
                                'PART_MAT_TYPE is assigned
                                PART_Material.PART_MAT_TYPE = mat.PART_MAT_TYPE
                                If PART_Material.Layer = mat.Layer Then
                                    PARTbolLayer = True
                                    PARTcomLayer = "OK"
                                    GoTo UNFOLDEDCLR4
                                Else
                                    PARTbolLayer = False
                                    PARTcomLayer = "The Layer of the Body should be " & mat.Layer
                                    GoTo UNFOLDEDCLR4
                                End If
                            Else
                                GoTo FALSECOLOR4
                            End If
                        Else
                            GoTo FALSECOLOR4
                        End If
                    Else
FALSECOLOR4:
                        PARTbolColor = False
                        PARTcomColor = "The PartBody Color should be R=" & mat.R & " G=" & mat.G & " B=" & mat.B
                        GoTo CHECKLAYER4
                    End If
UNFOLDEDCLR4:
                    If PART_Material.UnfoldedR = mat.R Then
                        If PART_Material.UnfoldedG = mat.G Then
                            If PART_Material.UnfoldedB = mat.B Then
                                PARTbolUBColor = True
                                PARTcomUBColor = "OK"
CHECKUBLAYER4:
                                'PART_MAT_TYPE is assigned
                                PART_Material.PART_MAT_TYPE = mat.PART_MAT_TYPE
                                If PART_Material.UnfoldedLayer = mat.Layer Or PART_Material.UnfoldedLayer = 0 Then
                                    PARTbolUBLayer = True
                                    PARTcomUBLayer = "OK"
                                    Exit Sub
                                Else
                                    PARTbolUBLayer = False
                                    PARTcomUBLayer = "The Layer of the Body should be " & mat.Layer & " OR 0"
                                    Exit Sub
                                End If
                            Else
                                GoTo FALSECOLORUB4
                            End If
                        Else
                            GoTo FALSECOLORUB4
                        End If
                    Else
FALSECOLORUB4:
                        PARTbolUBColor = False
                        PARTcomUBColor = "The PartBody Color should be R=" & mat.R & " G=" & mat.G & " B=" & mat.B
                        GoTo CHECKUBLAYER4
                    End If
                End If
                If match = True Then Exit Sub
            Next I
        End If
    End Sub

    'Sort L W T based on Values
    Sub SortLWT()
        Dim arr(3) As Double
        Dim temp As Double
        Dim I As Integer, j As Integer

        arr(1) = PART_L
        arr(2) = PART_W
        arr(3) = PART_T

        For I = LBound(arr) To UBound(arr)
            For j = I + 1 To UBound(arr)
                If arr(I) < arr(j) Then
                    temp = arr(j)
                    arr(j) = arr(I)
                    arr(I) = temp
                End If
            Next j
        Next I
    End Sub


    'PDM Properties Checking
    Sub CheckPDMPartProperties()
        Dim l As Double
        Dim w As Double
        Dim X As Double
        'Dim LCHK As Boolean
        'Dim WCHK As Boolean

        'Check if Item Zone is Filled or Not
        If PDM3D.PARTItemZone <> "" Then
            PARTbolIZ = True
            PARTcomIZ = "OK"
        Else
            PARTbolIZ = False
            PARTcomIZ = "Item Zone is Not Filled in PDM Link."
        End If
        'Check if Marking is Filled or Not
        If PDM3D.PARTMarking <> "" Then
            PARTbolMrk = True
            PARTcomMrk = "OK"
        Else
            PARTbolMrk = False
            PARTcomMrk = "Marking is Not Filled in PDM Link."
        End If
        'Check if Identifiable Part is Filled as 'N/A'
        If PDM3D.PARTid = "N/A" Then
            PARTbolID = True
            PARTcomID = "OK"
        Else
            PARTbolID = False
            PARTcomID = "Identifiable Part should be 'N/A' in PDM Link."
        End If
        'Check if Qty in Part is Filled as 'each'
        If InStr(1, PDM3D.PARTMatQty, "each") > 0 Then
            PARTbolQty = True
            PARTcomQty = "OK"
        Else
            PARTbolQty = False
            PARTcomQty = "Unit of Quantity in Material Part should be 'EACH', This Part has " & PDM3D.PARTMatQty
        End If
        'Check if FC is Filled or Not
        If PDM3D.PARTfc <> "" Then
            PARTbolFC = True
            PARTcomFC = "OK"
        Else
            PARTbolFC = False
            PARTcomFC = "Functional Class is Not Filled in PDM Link."
        End If
        'Check if MS is Filled or Not
        If PDM3D.PARTMatStruct <> "" Then
            PARTbolMS = True
            PARTcomMS = "OK"
        Else
            PARTbolMS = False
            PARTcomMS = "Material Structure is Not Filled in PDM Link."
        End If

        'Check the Protective Treatment Code
        If PDM3D.PARTProTreatCode <> "" Then
            If PART_Material.PART_MAT_TYPE = 1 Then
                If Left(PDM3D.PARTProTreatCode, 2) = "AA" Then
                    PARTbolProtCode = True
                    PARTcomProtCode = "OK"
                Else
                    GoTo WRONGPROTCODE
                End If
            ElseIf PART_Material.PART_MAT_TYPE = 2 Then
                If Left(PDM3D.PARTProTreatCode, 2) = "TA" Then
                    PARTbolProtCode = True
                    PARTcomProtCode = "OK"
                Else
                    GoTo WRONGPROTCODE
                End If
            ElseIf PART_Material.PART_MAT_TYPE = 3 Then
                If Left(PDM3D.PARTProTreatCode, 2) = "NA" Then
                    PARTbolProtCode = True
                    PARTcomProtCode = "OK"
                Else
                    GoTo WRONGPROTCODE
                End If
            Else
                GoTo WRONGPROTCODE
            End If
        Else
WRONGPROTCODE:
            PARTbolProtCode = False
            PARTcomProtCode = "Protective Treatment Code is Missing / Wrong."
        End If

        'Check the Manufacturing Index
        If PDM3D.PARTManuIndx = PDM3D.PARTIssIndx Then
            PARTbolMI = True
            PARTcomMI = "OK"
        ElseIf PDM3D.PARTManuIndx = "" Then
            PARTbolMI = True
            PARTcomMI = "OK"
        Else
            PARTbolMI = False
            PARTcomMI = "Manufacturing Index Should be Same as Issue Index or Blank."
        End If

        'Check Length of Raw Material
        PART_L = RoundUp(PART_L)
        If PDM3D.PARTLen > PART_L Then
            X = PART_L Mod 5
            If X = 0 Then
                l = PART_L + 5
                GoTo LENCOMPARE
            End If
            X = 10 - X
            If X < 5 Then
                X = X + 5
            End If
            l = PART_L + X
LENCOMPARE:
            If PDM3D.PARTLen >= l And PDM3D.PARTLen <= l + 1 Then
                PARTbolLEN = True
                PARTcomLEN = "OK"
            Else
                PARTbolLEN = False
                PARTcomLEN = "Part Length in PDM Link to be Rounded Off to < " & l & " >. < PDM=" & PDM3D.PARTLen & " AND CATIA=" & PART_L & " >"
            End If
        Else
            PARTbolLEN = False
            PARTcomLEN = "Part Length in PDM Link should be Greater Than CATIA Length. < PDM=" & PDM3D.PARTLen & " AND CATIA=" & PART_L & " >"
        End If

        'Check Width of Raw Material
        PART_W = RoundUp(PART_W)
        If PDM3D.PARTWid > PART_W Then
            X = PART_W Mod 5
            If X = 0 Then
                w = PART_W + 5
                GoTo WIDCOMPARE
            End If
            X = 10 - X
            If X < 5 Then
                X = X + 5
            End If
            w = PART_W + X
WIDCOMPARE:
            If PDM3D.PARTWid >= l And PDM3D.PARTWid <= l + 1 Then
                PARTbolWID = True
                PARTcomWID = "OK"
            Else
                PARTbolWID = False
                PARTcomWID = "Part Width in PDM Link to be Rounded Off to < " & w & " >. < PDM=" & PDM3D.PARTWid & " AND CATIA=" & PART_W & " >"
            End If
        Else
            PARTbolWID = False
            PARTcomWID = "Part Width in PDM Link should be Greater Than CATIA Width. < PDM=" & PDM3D.PARTWid & " AND CATIA=" & PART_W & " >"
        End If
    End Sub


    'My Coustom RoundUp Function
    Private Function RoundUp(ByVal Param As Double) As Double
        Dim I As Double

        I = Math.Round(Param, 3) - Math.Round(Param, 0)
        If I = 0 Then
            RoundUp = Math.Round(Param)
        ElseIf I > 0 Then
            RoundUp = Math.Round(Param) + 1
        Else
            RoundUp = Math.Round(Param)
        End If
    End Function




End Class
