Imports DRAFTINGITF
Imports INFITF
Imports MECMOD

Public Class ClsDrawShtGENProp

    Public TBLKenglishTitle As String
    Public TBLKdomesticTitle As String
    Public TBLKdrawingNumber As String
    Public TBLKsheetNumber As String
    Public TBLKfirstIssue As String
    Public TBLKsecondIssue As String
    Public TBLKsheetSize As String
    Public TBLKsignatures As Boolean         'Master
    Public TBLKsignatureDOsite As String
    Public TBLKsignatureDOcode As String
    Public TBLKsignature1 As String
    Public TBLKsignature2 As String
    Public TBLKsignature3 As String
    Public TBLKsignature4 As String
    Public TBLKsignature5 As String
    Public TBLKsignature6 As String
    Public TBLKsignature7 As String
    Public TBLKscale As String
    Public TBLKscaleBol As Boolean           'Master
    Public TBLKscaleBol2 As Boolean          'Master
    Public TBLKsurfaceFinish As String
    Public TBLKsurfaceFinishBol1 As Boolean  'Master
    Public TBLKsurfaceFinishBol2 As Boolean  'Master is ment to Flag Warning
    Public TBLKabd0001 As String
    Public TBLKabd0001Bol As Boolean         'Master
    Public TBLKyear As String
    Public TBLKyearBol As Boolean            'Master
    Public TBLKairbusGmbH As Boolean
    Public TBLKinterchangeable As Boolean
    Public TBLKgeoLocation As String
    Public TBLKidentMark1 As String
    Public TBLKidentMark2 As String
    Public TBLKidentMark3 As String
    Public TBLKidentMark4 As String
    Public TBLKidentMarkComb As String

    'PUBLIC: Other Properties
    Public DRWactualSheetSize As String
    Public DRWlimitX As Double
    Public DRWlimitY As Double
    Public DRWhtzNumber As String
    Public DRWsheetNumber As String
    Public DRWviewLockStatus As Boolean
    Public DRWviewHideShowStatus As Boolean
    Public DRWframeGridUsed As Boolean
    Public DRWlinePropType As Boolean
    Public DRWlinePropTypeLoc As String
    Public DRWlinePropThk As Boolean
    Public DRWlinePropThkLoc As String
    Public DRWtextSize As Boolean
    Public DRWtextSizeLoc As String
    Public DRWemptyMainView As Boolean
    Public DRWVWfakedim As Boolean
    Public DRWVWfakedimLocation As String
    Public DRWVWfakedimHid As Boolean
    Public DRWVWfakedimHidLocation As String
    Public DRWVWdimPrec1 As Boolean
    Public DRWVWdimPrec2 As Boolean
    Public DRWVWdimPrecLocation1 As String
    Public DRWVWdimPrecLocation2 As String
    Public BaloonSize As Boolean
    Public WrongBaloonSizeText As String
    Public WrongFlagnoteSizeText As String
    Public Flagnotesize As Boolean


    'PUBLIC: Run this to Get All Drawing Sheet Related Data

    Public Sub GetGENDrawingSheetProperties(ByRef DrawingDoc As DrawingDocument)
        Try

            Dim drawShts As DrawingSheets
        Dim drawSht As DrawingSheet
        Dim drawViews As DrawingViews
        Dim drawView As DrawingView
        Dim drawTexts As DrawingTexts
        'Dim drawText As DrawingText
        Dim drawGeoElems As GeometricElements

        drawShts = DrawingDoc.Sheets
        drawSht = drawShts.Item("Sheet.1")
        drawViews = drawSht.Views
        drawView = drawViews.Item(2)
        drawTexts = drawView.Texts
        drawGeoElems = drawView.GeometricElements

        'Set Public Properties
        'DRWhtzNumber = Left(drawShts.Name, 9)
        'DRWsheetNumber = Mid(drawShts.Name, 11, 2)

        Dim sss
        sss = Split(drawShts.Name, "\")
        DRWhtzNumber = Left(sss(UBound(sss)), 9)
        DRWsheetNumber = Mid(sss(UBound(sss)), 11, 2)

        'Get Public Title Block Properties
        TitleBlockProperties(drawTexts)
        'Get Drawing Sheet Actual Size
        DRWactualSheetSize = GetSheetSize(drawSht.GetPaperWidth)
        DRWlimitX = drawSht.GetPaperWidth
        DRWlimitY = drawSht.GetPaperHeight
        'Get Airbus Text
        TBLKairbusGmbH = GetAirbusTextStatus(drawTexts)
        'Get Year Status
        GetYearStatus()
        'Get ABD0001 Status
        GetABD0001Status()
        'Get Signature Status
        GetSignatureStatus()
        'Get the Surface Finish
        GetSurfaceFinishStatus()
        'Get Standard Scales
        TBLKscaleBol = DrawingStdScaleCheck()
        'Check the Scales of Views and Title Block
        TBLKscaleBol2 = DrawingViewAndTblkScaleCheck(drawViews)
        'Get Interchangeable Status
        TBLKinterchangeable = GetInterchangeable(drawGeoElems)
        'Get the 4 Identification Markings
        GetIdentificationMarking(drawTexts)

        'Check if latest Frame Grid is Used
        DRWframeGridUsed = DrawingFrameCheck(drawTexts)

        'Check if any View are there
        If drawViews.Count <= 2 Then
            BLANK_SHEET = True
            Exit Sub
        Else
            BLANK_SHEET = False
        End If

        'Check if all views are lock
        DRWviewLockStatus = DrawingViewLockCheck(drawViews)
        'Check if all views are in Show Mode
        DRWviewHideShowStatus = DrawingViewHideShowCheck(drawViews, DrawingDoc)
        'If TypeOfDrawing <> 5 Then
        'Check if All Line Types are Correct
        GetAllLineTypes(drawViews, DrawingDoc)
        'End If
        'Check if All Font Sizes are Correct
        GetAllFontSizes(drawViews)
        'Check All the Dimensions are Rounded to 3 decimal place
        GetDimensionsProperties(drawViews)
        'Check Main View Elements are Empty or Not
        CheckMainViewElements(drawViews)
            'Check BaloonSize
            BaloonSizeCheck(drawViews)

            'Check Flagnotes are in height of 3.5 mm
            FlagNotesSizeCheck(drawViews)

        Catch ex As Exception

        End Try
    End Sub

    Private Function TitleBlockProperties(ByVal drawTexts As DrawingTexts)
        Try
            For Each drawText In drawTexts
                If drawText.Name = "AUKTbkText_ENG_ALL_SIZE" And TBLKsheetSize = "" Then
                    TBLKsheetSize = drawText.Text
                ElseIf drawText.Name = "AUKTbkText_ENG_ALL_ISSUE1" And TBLKfirstIssue = "" Then
                    TBLKfirstIssue = drawText.Text
                ElseIf drawText.Name = "AUKTbkText_ENG_ALL_ISSUE2" And TBLKsecondIssue = "" Then
                    TBLKsecondIssue = drawText.Text
                ElseIf drawText.Name = "AUKTbkText_ENG_ALL_TITLE" And TBLKenglishTitle = "" Then
                    TBLKenglishTitle = drawText.Text
                ElseIf drawText.Name = "AUKTbkText_ENG_ALL_LOCAL_TITLE" And TBLKdomesticTitle = "" Then
                    TBLKdomesticTitle = drawText.Text
                ElseIf drawText.Name = "AUKTbkText_ENG_ALL_DRAWINGNUMBER" And TBLKdrawingNumber = "" Then
                    TBLKdrawingNumber = drawText.Text
                ElseIf drawText.Name = "AUKTbkText_ENG_ALL_SHEET" And TBLKsheetNumber = "" Then
                    TBLKsheetNumber = drawText.Text
                ElseIf drawText.Name = "AUKTbkText_ENG_ALL_DO_ORIGIN_SITE" And TBLKsignatureDOsite = "" Then
                    TBLKsignatureDOsite = drawText.Text
                ElseIf drawText.Name = "AUKTbkText_ENG_ALL_DO_ORIGIN_CODE" And TBLKsignatureDOcode = "" Then
                    TBLKsignatureDOcode = drawText.Text
                ElseIf drawText.Name = "AUKTbkText_ENG_ALL_APPD" And TBLKsignature1 = "" Then
                    TBLKsignature1 = drawText.Text
                ElseIf drawText.Name = "AUKTbkText_ENG_ALL_WTS" And TBLKsignature2 = "" Then
                    TBLKsignature2 = drawText.Text
                ElseIf drawText.Name = "AUKTbkText_ENG_ALL_PROCESS" And TBLKsignature3 = "" Then
                    TBLKsignature3 = drawText.Text
                ElseIf drawText.Name = "AUKTbkText_ENG_ALL_SYSTEM" And TBLKsignature4 = "" Then
                    TBLKsignature4 = drawText.Text
                ElseIf drawText.Name = "AUKTbkText_ENG_ALL_STRESS" And TBLKsignature5 = "" Then
                    TBLKsignature5 = drawText.Text
                ElseIf drawText.Name = "AUKTbkText_ENG_ALL_CHKD" And TBLKsignature6 = "" Then
                    TBLKsignature6 = drawText.Text
                ElseIf drawText.Name = "AUKTbkText_ENG_ALL_DRN" And TBLKsignature7 = "" Then
                    TBLKsignature7 = drawText.Text
                ElseIf drawText.Name = "AUKTbkText_ENG_ALL_SCALE" And TBLKscale = "" Then
                    TBLKscale = drawText.Text
                ElseIf drawText.Name = "AUKTbkText_ENG_ALL_SURFACE_FINISH" And TBLKsurfaceFinish = "" Then
                    TBLKsurfaceFinish = drawText.Text
                ElseIf drawText.Name = "AUKTbkText_ENG_ALL_ABD0001_SUFFIX" And TBLKabd0001 = "" Then
                    TBLKabd0001 = drawText.Text
                ElseIf drawText.Name = "AUKTbkText_ENG_MAP_COPYRIGHT" And TBLKyear = "" Then
                    TBLKyear = drawText.Text
                ElseIf drawText.Name = "AUKTbkText_ENG_ALL_GEOG_REF" And TBLKgeoLocation = "" Then
                    TBLKgeoLocation = drawText.Text
                End If
            Next
        Catch ex As Exception

        End Try
    End Function


    'PRIVATE: Get Sheet Size based on Sheet Actual Size

    Private Function GetSheetSize(ByVal ShtSize As Double) As String
        If ShtSize = 1189 Then
            GetSheetSize = "A0"
        ElseIf ShtSize = 841 Then
            GetSheetSize = "A1"
        ElseIf ShtSize = 594 Then
            GetSheetSize = "A2"
        ElseIf ShtSize = 420 Then
            GetSheetSize = "A3"
        Else
            GetSheetSize = "XX"
        End If
    End Function

    'PRIVATE: Check if Using Latest Frame Grid

    Private Function DrawingFrameCheck(ByRef AllDrawingTexts As DrawingTexts) As Boolean
        Dim drawText As DrawingText
        Dim str As String
        Dim I As Integer
        Dim ShtDB As Object
        Dim VerDB As Object
        Dim HorDB As Object
        Dim ver As Boolean
        Dim hor As Boolean

        ShtDB = {"A0", "A1", "A2", "A3"}
        VerDB = {"R", "M", "H", "F"}
        HorDB = {"24", "16", "12", "8"}

        If DRWactualSheetSize = "A0" Then I = 0
        If DRWactualSheetSize = "A1" Then I = 1
        If DRWactualSheetSize = "A2" Then I = 2
        If DRWactualSheetSize = "A3" Then I = 3

        DrawingFrameCheck = False
        ver = False
        hor = False
        For Each drawText In AllDrawingTexts
            str = drawText.Text
            If DRWactualSheetSize = ShtDB(I) Then
                If Len(str) = 1 Then
                    If str = VerDB(I) Then
                        ver = True
                    End If
                ElseIf Len(str) = 2 Then
                    If str = HorDB(I) Then
                        hor = True
                    End If
                End If
            End If
        Next

        If ver = True And hor = True Then
            DrawingFrameCheck = True
        Else
            DrawingFrameCheck = False
        End If
    End Function

    'PRIVATE: Check if all views are Locked

    Private Function DrawingViewLockCheck(ByRef AllDrawingViews As DrawingViews) As Boolean
        Dim drwVw As DrawingView
        Dim Check() As Boolean
        Dim I As Integer

        ReDim Check(AllDrawingViews.Count - 3)

        DrawingViewLockCheck = True
        I = 1
        For Each drwVw In AllDrawingViews
            If I > 2 Then
                If drwVw.LockStatus Then
                    Check(I - 3) = True
                Else
                    Check(I - 3) = False
                End If
            End If
            I = I + 1
        Next

        For I = 0 To AllDrawingViews.Count - 3
            If Check(I) = False Then
                DrawingViewLockCheck = False
                Exit Function
            End If
        Next I
    End Function

    'PRIVATE: Check if Interchangeable YES OR NO is ticked; YES= TRUE, NO= FALSE

    Private Function GetInterchangeable(ByRef AllDrawingGeoElements As GeometricElements) As Boolean
        Dim geoElement As GeometricElement

        For Each geoElement In AllDrawingGeoElements
            If geoElement.Name = "AUKTbkLine_ENG_ALL_IC_LINE_THRO_YES" Then
                GetInterchangeable = True
                Exit Function
            ElseIf geoElement.Name = "AUKTbkLine_ENG_ALL_IC_LINE_THRO_NO" Then
                GetInterchangeable = False
                Exit Function
            Else
                GetInterchangeable = False
            End If
        Next
    End Function

    'PRIVATE: Get SINGLE Text Based on Coordinates Range

    Private Function GetSingleTextByCoordinates(ByRef AllDrawingTexts As DrawingTexts, ByVal x1 As Double,
                                        ByVal x2 As Double, ByVal y1 As Double, ByVal y2 As Double) As String
        Dim drwText As DrawingText

        For Each drwText In AllDrawingTexts
            If drwText.x > x1 And drwText.x < x2 And drwText.y > y1 And drwText.y < y2 Then
                GetSingleTextByCoordinates = drwText.Text
                Exit Function
            End If
        Next
    End Function

    'PRIVATE: Get AIRBUS GmbH Text

    Private Function GetAirbusTextStatus(ByRef AllDrawingTexts As DrawingTexts) As Boolean
        Dim str As String

        If DRWactualSheetSize = "A0" Then
            str = GetSingleTextByCoordinates(AllDrawingTexts, 1070, 1090, 110, 115)
        ElseIf DRWactualSheetSize = "A1" Then
            str = GetSingleTextByCoordinates(AllDrawingTexts, 720, 740, 110, 115)
        ElseIf DRWactualSheetSize = "A2" Then
            str = GetSingleTextByCoordinates(AllDrawingTexts, 470, 490, 110, 115)
        ElseIf DRWactualSheetSize = "A3" Then
            str = GetSingleTextByCoordinates(AllDrawingTexts, 300, 320, 110, 115)
        End If

        If InStr(1, str, "GmbH") > 0 Then
            GetAirbusTextStatus = True
        Else
            GetAirbusTextStatus = False
        End If
    End Function

    'PRIVATE: Get Title Block Year

    Private Function GetYearStatus()
        If TBLKyear = Year(Now()) Then
            TBLKyearBol = True
        Else
            TBLKyearBol = False
        End If
    End Function

    'PRIVATE: Get Signature Status

    Private Function GetSignatureStatus()
        TBLKsignatures = False
        If TBLKsignatureDOsite = "9V" And TBLKsignatureDOcode = "" And TBLKsignature1 = "SEE ECN" And
        TBLKsignature2 = "SEE ECN" And TBLKsignature3 = "SEE ECN" And TBLKsignature4 = "SEE ECN" And
        TBLKsignature5 = "SEE ECN" And TBLKsignature6 = "SEE ECN" And TBLKsignature7 = "SEE ECN" Then

            TBLKsignatures = True
        Else
            TBLKsignatures = False
        End If
    End Function

    'PRIVATE: Get ABD0001 Type

    Private Function GetABD0001Status()
        TBLKabd0001Bol = False
        If TypeOfMaterial = 1 Then
            If TBLKabd0001 = "ABD0001-6" Then
                TBLKabd0001Bol = True
            End If
        ElseIf TypeOfMaterial = 2 Then
            If TBLKabd0001 = "ABD0001-2" Then
                TBLKabd0001Bol = True
            End If
        ElseIf TypeOfMaterial = 3 Then
            If TBLKabd0001 = "ABD0001-7" Then 'NEW UPDATE AS PER CHETAN TAMKER 11-11-2022
                ' If TBLKabd0001 = "ABD0001-2" Then 'NEW UPDATE AS PER CHETAN TAMKER
                TBLKabd0001Bol = True
            End If
        End If
    End Function

    'PRIVATE: Check if all views are in Show Mode

    Private Function DrawingViewHideShowCheck(ByRef AllDrawingViews As DrawingViews, ByRef drawDoc As DrawingDocument) As Boolean
        Dim drwVw As DrawingView
        Dim Check() As Boolean
        Dim I As Integer
        Dim showState As CatVisPropertyShow
        Dim VisProp
        Dim sel As Selection

        ReDim Check(AllDrawingViews.Count - 3)
        sel = drawDoc.Selection

        DrawingViewHideShowCheck = True
        I = 1
        For Each drwVw In AllDrawingViews
            If I > 2 Then
                sel.Add(drwVw)
                VisProp = sel.VisProperties
                VisProp.GetShow(showState)
                If showState = 0 Then
                    Check(I - 3) = True
                Else
                    Check(I - 3) = False
                End If
                sel.Clear
                VisProp = Nothing
            End If
            I = I + 1
        Next

        For I = 0 To AllDrawingViews.Count - 3
            If Check(I) = False Then
                DrawingViewHideShowCheck = False
                Exit Function
            End If
        Next I
    End Function

    'PRIVATE: Check if Scale in Title Block is Standard one

    Private Function DrawingStdScaleCheck() As Boolean
        Dim stdScales As Object
        Dim I As Integer

        stdScales = {"NOT TO SCALE", "50:1", "20:1", "10:1", "5:1", "2:1", "1:1",
                        "3:5", "1:2", "1:5", "1:10", "1:20", "1:50", "1:3", "2:5"}     'ADDED 1:3 AND 2:5 AS PER CHETAN TAMKER
        DrawingStdScaleCheck = False
        For I = 0 To UBound(stdScales)
            If stdScales(I) = TBLKscale Then
                DrawingStdScaleCheck = True
                Exit For
            End If
        Next I
    End Function

    'PRIVATE: Check if Scale in Title Block is Standard one

    Private Function DrawingViewAndTblkScaleCheck(ByRef AllDrawingViews As DrawingViews) As Boolean
        Dim drwView As DrawingView
        Dim thisScale As Object
        Dim firstVal As Double
        Dim secondVal As Double
        Dim fracScale As Double
        Dim I As Integer

        DrawingViewAndTblkScaleCheck = False

        If TBLKscale = "NOT TO SCALE" Then
            DrawingViewAndTblkScaleCheck = True
        ElseIf InStr(1, TBLKscale, ":") > 0 Then
            thisScale = Split(TBLKscale, ":")
            firstVal = CDbl(thisScale(0))
            secondVal = CDbl(thisScale(1))
            fracScale = firstVal / secondVal
            I = 1
            For Each drwView In AllDrawingViews
                If I > 2 Then
                    If fracScale = drwView.Scale2 Then
                        DrawingViewAndTblkScaleCheck = True
                    End If
                End If
                I = I + 1
            Next
        End If
    End Function

    'PRIVATE: Get ABD0001 Type

    Private Function GetSurfaceFinishStatus()
        TBLKsurfaceFinishBol1 = False
        TBLKsurfaceFinishBol2 = False
        If TypeOfMaterial = 1 Then
            If TBLKsurfaceFinish = "" Then
                TBLKsurfaceFinishBol1 = True
                TBLKsurfaceFinishBol2 = True
            ElseIf TBLKsurfaceFinish = "6,3" Then
                TBLKsurfaceFinishBol1 = True
                TBLKsurfaceFinishBol2 = False
            Else
                TBLKsurfaceFinishBol1 = False
                TBLKsurfaceFinishBol2 = False
            End If
        ElseIf TypeOfMaterial = 2 Then
            If TBLKsurfaceFinish = "" Then
                TBLKsurfaceFinishBol1 = True
                TBLKsurfaceFinishBol2 = True
            ElseIf TBLKsurfaceFinish = "3,2" Or TBLKsurfaceFinish = "6,3" Then
                TBLKsurfaceFinishBol1 = True
                TBLKsurfaceFinishBol2 = False
            Else
                TBLKsurfaceFinishBol1 = False
                TBLKsurfaceFinishBol2 = False
            End If
        ElseIf TypeOfMaterial = 3 Then
            If TBLKsurfaceFinish = "" Then
                TBLKsurfaceFinishBol1 = True
                TBLKsurfaceFinishBol2 = True
            ElseIf TBLKsurfaceFinish = "3,2" Then
                TBLKsurfaceFinishBol1 = True
                TBLKsurfaceFinishBol2 = False
            Else
                TBLKsurfaceFinishBol1 = False
                TBLKsurfaceFinishBol2 = False
            End If
        End If
    End Function

    'PRIVATE: Check if Line Type and Thickness is Correct

    Private Sub GetAllLineTypes(ByRef AllDrawingViews As DrawingViews, ByRef drawDoc As DrawingDocument)
        Dim geoElements As GeometricElements
        Dim geoElement As GeometricElement
        Dim drawView As DrawingView
        Dim sel As Selection
        Dim I As Integer
        Dim j1 As Integer
        Dim j2 As Integer
        Dim lineType As Long
        Dim lineThick As Long

        j1 = 1
        j2 = 1
        DRWlinePropType = True
        DRWlinePropThk = True
        sel = drawDoc.Selection
        For I = 3 To AllDrawingViews.Count
            drawView = AllDrawingViews.Item(I)
            geoElements = drawView.GeometricElements
            If geoElements.Count < 1000 Then
                For Each geoElement In geoElements
                    If geoElement.GeometricType = CatGeometricType.catGeoTypeLine2D _
                  Or geoElement.GeometricType = CatGeometricType.catGeoTypeCircle2D _
                  Or geoElement.GeometricType = CatGeometricType.catGeoTypeHyperbola2D _
                  Or geoElement.GeometricType = CatGeometricType.catGeoTypeParabola2D _
                  Or geoElement.GeometricType = CatGeometricType.catGeoTypeSpline2D _
                  Or geoElement.GeometricType = CatGeometricType.catGeoTypeEllipse2D Then
                        sel.Add(geoElement)
                        sel.VisProperties.GetRealLineType(lineType)
                        sel.VisProperties.GetRealWidth(lineThick)
                        sel.Clear()

                        If lineType <> 1 And lineType <> 4 And lineType <> 5 And lineType <> 0 And lineType <> 8 Then
                            DRWlinePropType = False
                            If j1 = 1 Then
                                DRWlinePropTypeLoc = " <" & geoElement.Name & " @ " & drawView.Name & ">"
                                j1 = j1 + 1
                            Else
                                DRWlinePropTypeLoc = DRWlinePropTypeLoc & ", " & "<" & geoElement.Name & " @ " & drawView.Name & ">"
                            End If
                        ElseIf lineThick <> 1 And lineThick <> 2 And lineThick <> 3 Then
                            DRWlinePropThk = False
                            If j2 = 1 Then
                                DRWlinePropThkLoc = " <" & geoElement.Name & " @ " & drawView.Name & ">"
                                j2 = j2 + 1
                            Else
                                DRWlinePropThkLoc = DRWlinePropThkLoc & ", " & "<" & geoElement.Name & " @ " & drawView.Name & ">"
                            End If
                        End If
                    End If
                Next
            End If
        Next
    End Sub

    'PRIVATE: Check if Font Size is Correct

    Private Sub GetAllFontSizes(ByRef AllDrawingViews As DrawingViews)
        Dim txts As DrawingTexts
        Dim txt As DrawingText
        Dim drawView As DrawingView
        Dim I As Integer
        Dim j As Integer
        Dim fontSize As Double
        Dim subTxtFontSize As Double
        Dim subTxtLookup As Double

        j = 1
        DRWtextSize = True
        For I = 3 To AllDrawingViews.Count

            drawView = AllDrawingViews.Item(I)
            If drawView.Texts.Count > 0 Then
                txts = drawView.Texts
                For Each txt In txts
                    fontSize = txt.TextProperties.FontSize
                    subTxtLookup = Len(txt.Text)
                    'Check for Text Sizes of Sub Text
                    If subTxtLookup > 3 Then
                        subTxtFontSize = txt.GetParameterOnSubString(CatTextProperty.catFontSize, subTxtLookup - 3, 3) / 1000
                    Else
                        subTxtFontSize = 100
                    End If
                    'Compare the Results
                    If fontSize = 5 Or fontSize = 7 Or fontSize = 3.5 Then
                        If subTxtFontSize <> 100 Then
                            If subTxtFontSize <> 5 And subTxtFontSize <> 7 And subTxtFontSize <> 3.5 Then
                                DRWtextSize = False
                                If j = 1 Then
                                    DRWtextSizeLoc = "<" & "'" & txt.Text & "'" & " @ " & drawView.Name & ">"
                                    j = j + 1
                                Else
                                    DRWtextSizeLoc = DRWtextSizeLoc & ", " & "<" & "'" & txt.Text & "'" & " @ " & drawView.Name & ">"
                                End If
                            End If
                        End If
                    Else
                        DRWtextSize = False
                        If j = 1 Then
                            DRWtextSizeLoc = "<" & "'" & txt.Text & "'" & " @ " & drawView.Name & ">"
                            j = j + 1
                        Else
                            DRWtextSizeLoc = DRWtextSizeLoc & ", " & "<" & "'" & txt.Text & "'" & " @ " & drawView.Name & ">"
                        End If
                    End If
                Next
            End If
        Next
    End Sub

    'PRIVATE: Get Dimension Properties

    Private Sub GetDimensionsProperties(ByRef AllDrawingViews As DrawingViews)

        Dim AllDrawingDimensions As DrawingDimensions
        Dim drawDim As DrawingDimension
        Dim drawView As DrawingView
        Dim dimen As DrawingDimValue
        Dim fakeDim As Integer
        Dim dimPrecision As Double
        Dim I As Integer
        Dim j As Integer
        Dim k1 As Integer
        Dim k2 As Integer
        Dim k3 As Integer
        Dim k4 As Integer
        'Dim q As Boolean

        I = 1
        j = 1
        k1 = 1
        k2 = 1
        k3 = 1
        k4 = 1
        DRWVWfakedim = True
        DRWVWfakedimHid = True
        DRWVWdimPrec1 = True
        DRWVWdimPrec2 = True
        For Each drawView In AllDrawingViews
            If I > 2 Then
                AllDrawingDimensions = drawView.Dimensions
                For j = 1 To AllDrawingDimensions.Count
                    drawDim = AllDrawingDimensions.Item(j)
                    dimen = drawDim.GetValue
                    'Check for fake Dimension
                    fakeDim = dimen.FakeDimType
                    If fakeDim <> 0 Then
                        DRWVWfakedim = False
                        If k1 = 1 Then
                            DRWVWfakedimLocation = "<" & dimen.GetFakeDimValue(1) & " @ " & drawView.Name & ">"
                            k1 = k1 + 1
                            GoTo IfFakeSkipAboveSteps
                        Else
                            DRWVWfakedimLocation = DRWVWfakedimLocation & ", " & "<" & dimen.GetFakeDimValue(1) & " @ " & drawView.Name & ">"
                            GoTo IfFakeSkipAboveSteps
                        End If
                    End If
                    'Check for Hidden Dimensions
                    If drawDim.ValueDisplay = 2 Then
                        DRWVWfakedimHid = False
                        If k4 = 1 Then
                            DRWVWfakedimHidLocation = "<" & drawDim.Name & " @ " & drawView.Name & ">"
                            k4 = k4 + 1
                            GoTo IfFakeSkipAboveSteps
                        Else
                            DRWVWfakedimHidLocation = DRWVWfakedimHidLocation & ", " & "<" & drawDim.Name & " @ " & drawView.Name & ">"
                            GoTo IfFakeSkipAboveSteps
                        End If
                    End If
                    'Check for Dimension Precision
                    dimPrecision = dimen.GetFormatPrecision(1)
                    If dimPrecision < 0.1 Then
                        DRWVWdimPrec1 = False
                        If k2 = 1 Then
                            DRWVWdimPrecLocation1 = "<" & Left(dimen.Value, 6) & " @ " & drawView.Name & ">"
                            k2 = k2 + 1
                        Else
                            DRWVWdimPrecLocation1 = DRWVWdimPrecLocation1 & ", " & "<" & Left(dimen.Value, 6) & " @ " & drawView.Name & ">"
                        End If
                    End If
                    If dimPrecision = 1 Then
                        DRWVWdimPrec2 = False
                        If k3 = 1 Then
                            DRWVWdimPrecLocation2 = "<" & Left(dimen.Value, 6) & " @ " & drawView.Name & ">"
                            k3 = k3 + 1
                        Else
                            DRWVWdimPrecLocation2 = DRWVWdimPrecLocation2 & ", " & "<" & Left(dimen.Value, 6) & " @ " & drawView.Name & ">"
                        End If
                    End If

IfFakeSkipAboveSteps:

                Next j
            End If
            I = I + 1
        Next
    End Sub

    'PRIVATE: Get MULTIPLE Text Based on Coordinates Range

    Private Function GetMultipleTextByCoordinates(ByRef AllDrawingTexts As DrawingTexts, ByVal x1 As Double,
                                        ByVal x2 As Double, ByVal y1 As Double, ByVal y2 As Double) As Object
        Dim drwText As DrawingText
        Dim var As Object
        Dim I As Integer
        Dim j As Integer

        j = 0
        For Each drwText In AllDrawingTexts
            If drwText.x > x1 And drwText.x < x2 And drwText.y > y1 And drwText.y < y2 Then
                j = j + 1
            End If
        Next
        If j <> 0 Then
            ReDim var(j)
            I = 1
            For Each drwText In AllDrawingTexts
                If drwText.x > x1 And drwText.x < x2 And drwText.y > y1 And drwText.y < y2 Then
                    var(I) = drwText.Text
                    If I = j Then Exit For
                    I = I + 1
                End If
            Next
            GetMultipleTextByCoordinates = var
        Else
            GetMultipleTextByCoordinates = ""
        End If
    End Function

    'PRIVATE: Get Identification Marking

    Private Sub GetIdentificationMarking(ByRef AllDrawingTexts As DrawingTexts)
        Dim var As Object
        Dim I As Integer

        If DRWactualSheetSize = "A0" Then
            var = GetMultipleTextByCoordinates(AllDrawingTexts, 1019, 1039, 60, 75)
        ElseIf DRWactualSheetSize = "A1" Then
            var = GetMultipleTextByCoordinates(AllDrawingTexts, 671, 690, 60, 75)  '671, 691, 60, 75)
        ElseIf DRWactualSheetSize = "A2" Then
            var = GetMultipleTextByCoordinates(AllDrawingTexts, 424, 444, 60, 75)
        ElseIf DRWactualSheetSize = "A3" Then
            var = GetMultipleTextByCoordinates(AllDrawingTexts, 250, 270, 60, 75)
        End If

        If TypeName(var) = "Variant()" Then
            For I = 1 To UBound(var)
                If I = 1 Then TBLKidentMark1 = var(1)
                If I = 2 Then TBLKidentMark2 = var(2)
                If I = 3 Then TBLKidentMark3 = var(3)
                If I = 4 Then TBLKidentMark4 = var(4)
                If I = 1 Then
                    TBLKidentMarkComb = var(I)
                Else
                    TBLKidentMarkComb = TBLKidentMarkComb & "+" & var(I)
                End If
            Next I
        End If
    End Sub

    'PRIVATE: Main View (Sheet.1) Must be Empty

    Private Sub CheckMainViewElements(ByRef AllDrawingViews As DrawingViews)
        'Dim I As Integer
        Dim mainView As DrawingView
        Dim iArrows As DrawingArrows
        Dim iComps As DrawingComponents
        Dim iDims As DrawingDimensions
        Dim iGoeEles As GeometricElements
        Dim iTables As DrawingTables
        Dim iTexts As DrawingTexts

        mainView = AllDrawingViews.Item(1)
        iArrows = mainView.Arrows
        iComps = mainView.Components
        iDims = mainView.Dimensions
        iGoeEles = mainView.GeometricElements
        iTables = mainView.Tables
        iTexts = mainView.Texts

        DRWemptyMainView = True
        If iArrows.Count <> 0 Then DRWemptyMainView = False
        If iComps.Count <> 0 Then DRWemptyMainView = False
        If iDims.Count <> 0 Then DRWemptyMainView = False
        If iGoeEles.Count <> 1 Then DRWemptyMainView = False
        If iTables.Count <> 0 Then DRWemptyMainView = False
        If iTexts.Count <> 0 Then DRWemptyMainView = False
    End Sub

    'PRIVATE: Check baloon size

    Private Function BaloonSizeCheck(ByRef AllDrawingViews As DrawingViews) As Boolean

        'Dim drawView As DrawingView
        Dim CurrentText As DrawingText
        Dim I, j As Integer
        Dim currentView As DrawingView

        For I = 3 To AllDrawingViews.Count
            currentView = AllDrawingViews.Item(I)
            For j = 1 To currentView.Texts.Count
                CurrentText = currentView.Texts.Item(j)

                Dim visProperty_Set As VisPropertySet
                If CurrentText.FrameName = "Circle" Or CurrentText.FrameName = "Variable_Circle" Then
                    Call Osel.Clear()
                    Call Osel.Add(CurrentText)
                    visProperty_Set = Osel.VisProperties
                    'visProperty_Set.GetShow(showState)
                    'Debug.Print(showState)
                    Dim showstate As Integer
                    If showState = 0 Then
                        If CInt(CurrentText.Text) < 100 And CurrentText.FrameName = "Variable_Circle" Then
                            'BaloonSize = False
                            If Not WrongBaloonSizeText = "" Then
                                WrongBaloonSizeText = WrongBaloonSizeText & "<" & CurrentText.Text & " @ " & currentView.Name & ">"
                            Else
                                WrongBaloonSizeText = "<" & CurrentText.Text & " @ " & currentView.Name & ">"
                            End If
                        ElseIf CInt(CurrentText.Text) > 100 And CurrentText.FrameName = "Circle" Then
                            'BaloonSize = False
                            If Not WrongBaloonSizeText = "" Then
                                WrongBaloonSizeText = WrongBaloonSizeText & "<" & CurrentText.Text & " @ " & currentView.Name & ">"
                            Else
                                WrongBaloonSizeText = "<" & CurrentText.Text & " @ " & currentView.Name & ">"
                            End If
                        End If
                    Else
                        'BaloonSize = True
                    End If

                End If

            Next
        Next

        If WrongBaloonSizeText = "" Then
            BaloonSize = True
        Else
            BaloonSize = False
        End If
        Osel.Clear()

    End Function

    'PRIVATE: Check Flag Notes Size

    Private Function FlagNotesSizeCheck(ByRef AllDrawingViews As DrawingViews)
        Dim I As Integer
        Dim j As Integer
        Dim drawVw As DrawingView
        Dim dTexts As DrawingTexts
        Dim dtxt As DrawingText
        Dim flagNote As ClsText

        Dim fontSize

        I = 1
        For Each drawVw In AllDrawingViews
            If I > 2 Then
                dTexts = drawVw.Texts
                For j = 1 To dTexts.Count
                    dtxt = dTexts.Item(j)
                    'Check if the Text is Numeric
                    If IsNumeric(dtxt.Text) Then


                        If dtxt.FrameType = 17 And Len(dtxt.Text) = 3 Then
                            flagNote = New ClsText
                            flagNote.TextContent = dtxt.Text
                            flagNote.FrameType = CInt(dtxt.FrameType)
                            fontSize = dtxt.TextProperties.FontSize
                            Dim int As Integer = 4
                            If int = 4 = True Then 'frmStart.optSingle.Visible

                                If fontSize = "5" Then
                                    If Not WrongFlagnoteSizeText = "" Then
                                        WrongFlagnoteSizeText = WrongFlagnoteSizeText & "<" & dtxt.Text & " @ " & drawVw.Name & ">"
                                    Else
                                        WrongFlagnoteSizeText = ""
                                    End If

                                Else
                                    WrongFlagnoteSizeText = WrongFlagnoteSizeText & "<" & dtxt.Text & " @ " & drawVw.Name & ">"
                                End If
                            Else

                                If fontSize = "3.5" Then
                                    If Not WrongFlagnoteSizeText = "" Then
                                        WrongFlagnoteSizeText = WrongFlagnoteSizeText & "<" & dtxt.Text & " @ " & drawVw.Name & ">"
                                    Else
                                        WrongFlagnoteSizeText = ""
                                    End If

                                Else
                                    WrongFlagnoteSizeText = WrongFlagnoteSizeText & "<" & dtxt.Text & " @ " & drawVw.Name & ">"
                                End If
                            End If

                        End If
                    End If
                Next j
            End If
            I = I + 1
            dTexts = Nothing
        Next

        If WrongFlagnoteSizeText = "" Then
            Flagnotesize = True
        Else
            Flagnotesize = False
        End If


    End Function

End Class
