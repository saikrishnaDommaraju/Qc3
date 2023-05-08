Imports DRAFTINGITF

Public Class Cls_Func

    Public InstlPropPub As ClsDrawShtINSTProp

    'PUBLIC: Run this to Get All Drawing Sheet Related Data

    Public Sub GetINSTDrawingSetProperties(ByRef InstlProp As ClsDrawShtINSTProp)

        'Dim myDrawDoc As DrawingDocument
        'Dim drawTexts As DrawingTexts
        Dim drawDoc As DrawingDocument
        Dim drawShts As DrawingSheets
        Dim drawSht As DrawingSheet
        Dim drawViews As DrawingViews
        Dim drawView As DrawingView

        Dim I As Integer

        InstlPropPub = InstlProp

        For I = 1 To InstlProp.DrawDocsCol.Count
            InstlProp.GenProps = InstlProp.GenPropCol.Item(I)
            drawDoc = InstlProp.DrawDocsCol.Item(I)
            drawShts = drawDoc.Sheets
            drawSht = drawShts.Item("Sheet.1")
            drawViews = drawSht.Views
            'Get All Item Numbers
            GetAllItemNumbers(drawViews, I, InstlProp.GenProps)
            'Get All Flag Notes
            GetAllFlagNotes(drawViews, InstlProp.GenProps)
            If I = 1 And TypeOfDrawing = 5 Then
                GetInstlItemNumber(drawViews)
            End If
        Next I

        'Get IP Table Form 1st Sheet and Compare with PDM and Item Numbers
        If TypeOfDrawing <> 5 Then
            drawView = InstlProp.DrawDocsCol.Item(1).Sheets.Item(1).Views.Item(2)
            GetIPtable(drawView)
        End If

        'Process the 2x Text beside the Item Number
        GetItemCountbyText()

        'Check if Instl Item Numbers are OK
        If TypeOfDrawing <> 5 Then
            CheckPartItem()  '//////////Have to write a replacement function for this
        End If

        'Consolidate Drawing Item Numbers.
        ConsolidateDrwItemNum()

    End Sub


    'PRIVATE: Get All Item Numbers in Drawing Sheet

    Private Sub GetAllItemNumbers(ByRef AllDrawingViews As DrawingViews, ByVal SheetCount As Integer, ByRef GenProps As ClsDrawShtGENProp)
        Dim I As Integer
        Dim j As Integer
        Dim drawVw As DrawingView
        Dim dTexts As DrawingTexts
        Dim dtxt As DrawingText
        Dim assosiatedTxt As DrawingText
        Dim itmNum As ClsText
        Dim xTxt As ClsText
        Dim handCheck As Integer

        I = 1
        For Each drawVw In AllDrawingViews
            If I > 1 Then
                dTexts = drawVw.Texts
                For j = 1 To dTexts.Count
                    dtxt = dTexts.Item(j)
                    'Check if the Text is Numeric
                    If IsNumeric(dtxt.Text) Then
                        If dtxt.FrameType = CatTextFrameType.catCircle Or dtxt.FrameType = 15 Or dtxt.FrameType = 18 Or dtxt.FrameType = 19 _
                     Or dtxt.FrameType = 53 Or dtxt.FrameType = 65 Or dtxt.FrameType = 68 Or dtxt.FrameType = 69 _
                     And Len(dtxt.Text) <= 4 Then
                            itmNum = New ClsText
                            itmNum.TextContent = dtxt.Text
                            itmNum.FrameType = CInt(dtxt.FrameType)
                            'Call the Function to Map the Item No Coord to Sheet Coord
                            itmNum.MapTextCoordinates(dtxt.x, dtxt.y, drawVw.xAxisData, drawVw.yAxisData, drawVw.Scale2, drawVw.Angle)
                            itmNum.AnchorPosType = CInt(dtxt.AnchorPosition)
                            itmNum.ItemCountL = dtxt.Leaders.Count
                            'check if Item Number is inside Sheet and Add them in the Collection
                            If itmNum.TextSx <= GenProps.DRWlimitX And itmNum.TextSY <= GenProps.DRWlimitX _
                         And itmNum.TextSx > 0 And itmNum.TextSY > 0 Then
                                itmNum.GetTextZone(GenProps.DRWsheetNumber, GenProps.DRWactualSheetSize)
                                'Check if Double Baloon
                                If dtxt.FrameType = 15 Or dtxt.FrameType = 65 Then
                                    handCheck = CInt(itmNum.TextContent) Mod 2
                                    If handCheck = 0 Then
                                        If InstlPropPub.PartNumSht1 Is Nothing And SheetCount = 1 Then
                                            InstlPropPub.PartNumSht1 = itmNum
                                        ElseIf InstlPropPub.PartNumSht2 Is Nothing And SheetCount = 2 Then
                                            InstlPropPub.PartNumSht2 = itmNum
                                        ElseIf SheetCount >= 3 Then
                                            InstlPropPub.PartNumExtraBol = True
                                            InstlPropPub.PartNumExtra = itmNum
                                        End If
                                    End If
                                Else
                                    If itmNum.ItemCountL = 0 Then
                                        If itmNum.TextContent = "6230" Or itmNum.TextContent = "6260" Or itmNum.TextContent = "6261" Then
                                            GoTo GoOn
                                        End If
                                        On Error GoTo GoOn
                                        'If Not dtxt.AssociativeElement Is Nothing Then
                                        assosiatedTxt = dtxt.AssociativeElement
                                        itmNum.ItemCountL = assosiatedTxt.Leaders.Count
                                        'End If
                                    End If
GoOn:
                                    On Error GoTo - 1
                                    If TypeOfDrawing = 5 And itmNum.TextContent = "7100" Then
                                        '//DONT ADD TO ITEMNUMBER LIST
                                    Else
                                        InstlPropPub.ItemNumberList.Add(itmNum)
                                    End If
                                End If
                            End If
                        End If
                    End If
                    'Count the Text ending with 'X'
                    If dtxt.FrameType = CatTextFrameType.catNone Then
                        If Right(dtxt.Text, 1) = "x" Or Right(dtxt.Text, 1) = "X" Then
                            If IsNumeric(Left(dtxt.Text, Len(dtxt.Text) - 1)) Then
                                xTxt = New ClsText
                                xTxt.TextContent = dtxt.Text
                                xTxt.FrameType = CInt(dtxt.FrameType)
                                'Call the Function to Map the Item No Coord to Sheet Coord
                                xTxt.MapTextCoordinates(dtxt.x, dtxt.y, drawVw.xAxisData, drawVw.yAxisData, drawVw.Scale2, drawVw.Angle)
                                xTxt.AnchorPosType = CInt(dtxt.AnchorPosition)
                                InstlPropPub.XlistCount.Add(xTxt)
                            End If
                        End If
                    End If
                Next j
            End If
            I = I + 1
            dTexts = Nothing
        Next
    End Sub

    'PRIVATE: Get All Flag Notes in Drawing Sheet

    Private Sub GetAllFlagNotes(ByRef AllDrawingViews As DrawingViews, ByRef GenProps As ClsDrawShtGENProp)
        Dim I As Integer
        Dim j As Integer
        Dim drawVw As DrawingView
        Dim dTexts As DrawingTexts
        Dim dtxt As DrawingText
        Dim flagNote As ClsText

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
                            'Call the Function to Map the Item No Coord to Sheet Coord
                            flagNote.MapTextCoordinates(dtxt.x, dtxt.y, drawVw.xAxisData, drawVw.yAxisData, drawVw.Scale2, drawVw.Angle)
                            flagNote.AnchorPosType = CInt(dtxt.AnchorPosition)
                            'check if Text is inside Sheet and Add them in the Collection
                            If flagNote.TextSx <= GenProps.DRWlimitX And flagNote.TextSY <= GenProps.DRWlimitX _
                         And flagNote.TextSx > 0 And flagNote.TextSY > 0 Then
                                flagNote.GetTextZone(GenProps.DRWsheetNumber, GenProps.DRWactualSheetSize)
                                InstlPropPub.FlagNoteList.Add(flagNote)
                            End If
                        End If
                    End If
                Next j
            End If
            I = I + 1
            dTexts = Nothing
        Next
    End Sub

    'PRIVATE: Get the Item Count by Text "2x" etc.

    Private Sub GetItemCountbyText()
        Dim xTxt As ClsText
        Dim iTxt As ClsText
        Dim X As Double
        Dim Y As Double
        Dim xoff As Double
        Dim yoff As Double
        Dim x1Bound As Double
        Dim y1Bound As Double
        Dim x2Bound As Double
        Dim y2Bound As Double

        If InstlPropPub.ItemNumberList.Count > 0 And InstlPropPub.XlistCount.Count > 0 Then
            For Each iTxt In InstlPropPub.ItemNumberList
                X = iTxt.TextSx
                Y = iTxt.TextSY
                If iTxt.AnchorPosType = 1 Then
                    If Len(iTxt.TextContent) = 4 Then
                        xoff = 8.5
                        yoff = 4.5
                    ElseIf Len(iTxt.TextContent) = 3 Then
                        xoff = 6.5
                        yoff = 4.5
                    ElseIf Len(iTxt.TextContent) = 2 Then
                        xoff = 4.5
                        yoff = 4.5
                    ElseIf Len(iTxt.TextContent) = 1 Then
                        xoff = 2
                        yoff = 4.5
                    End If
                    X = X + xoff
                    Y = Y - yoff
                    x1Bound = X - 30
                    x2Bound = X + 30
                    y1Bound = Y - 10
                    y2Bound = Y + 10
                End If
                For Each xTxt In InstlPropPub.XlistCount
                    If xTxt.TextSx > x1Bound And xTxt.TextSx < x2Bound Then
                        If xTxt.TextSY > y1Bound And xTxt.TextSY < y2Bound Then
                            iTxt.ItemCountT = Left(xTxt.TextContent, Len(xTxt.TextContent) - 1)
                        End If
                    End If
                Next
                iTxt.SetTotalItemCount()
            Next
        ElseIf InstlPropPub.ItemNumberList.Count > 0 And InstlPropPub.XlistCount.Count = 0 Then
            For Each iTxt In InstlPropPub.ItemNumberList
                iTxt.SetTotalItemCount()
            Next
        End If
    End Sub

    'PRIVATE: Check if Part Number in 1st and 2nd Sheet are Same and if Extra Items are present.

    Private Sub CheckPartItem()
        If InstlPropPub.PartNumSht2 Is Nothing Then
            InstlPropPub.PartNumMatchBol = False
            InstlPropPub.PartNumMatchErrMsg = "Installation Item Number in Sheet-2 is Missing"
        ElseIf InstlPropPub.PartNumSht1.TextContent <> InstlPropPub.PartNumSht2.TextContent Then
            InstlPropPub.PartNumMatchBol = False
            InstlPropPub.PartNumMatchErrMsg = "Installation Item Number in Sheet-1 & Sheet-2 Does Not Match"
        Else
            InstlPropPub.PartNumMatchBol = True
        End If

        'Join the Item Zones for Checking with PDM Zones
        'If InstlPropPub.PartNumSht2 Is Nothing Then
        'InstlPropPub.PartNumsZones = InstlPropPub.PartNumSht1.TextZone
        'Else
        '  InstlPropPub.PartNumsZones = InstlPropPub.PartNumSht1.TextZone & " " & InstlPropPub.PartNumSht2.TextZone
        ' End If

        'Check if Extra is there
        If InstlPropPub.PartNumExtraBol = True Then
            InstlPropPub.PartNumExtraBol = False
            InstlPropPub.PartNumExtraErrMsg = "Double Baloons are Used More than 2 Times. < " & InstlPropPub.PartNumExtra.TextContent & " @ " & InstlPropPub.PartNumExtra.TextZone & " >"
        Else
            InstlPropPub.PartNumExtraBol = True
        End If
    End Sub

    'PRIVATE: Consolidate Drawing Item Numbers.

    Private Sub ConsolidateDrwItemNum()
        Dim I As Integer
        Dim itmNum As ClsText
        Dim tmpNum As ClsText

        For Each itmNum In InstlPropPub.ItemNumberList
            If itmNum.ItemCount > 1 Then
                For I = 1 To itmNum.ItemCount - 1
                    tmpNum = New ClsText
                    tmpNum.AnchorPosType = itmNum.AnchorPosType
                    tmpNum.FrameType = itmNum.FrameType
                    tmpNum.ItemCountL = itmNum.ItemCountL
                    tmpNum.ItemCountT = itmNum.ItemCountT
                    tmpNum.Qcheck = itmNum.Qcheck
                    tmpNum.TextContent = itmNum.TextContent
                    tmpNum.TextSx = itmNum.TextSx
                    tmpNum.TextSY = itmNum.TextSY
                    tmpNum.TextZone = itmNum.TextZone
                    tmpNum.ItemCount = 1

                    itmNum.ItemCount = itmNum.ItemCount - 1
                    InstlPropPub.ItemNumberList.Add(tmpNum)
                Next I
            End If
        Next
    End Sub

    'PRIVATE: Get IP Table Sheet-1.

    Private Sub GetIPtable(ByRef dView As DrawingView)
        Dim I As Integer
        Dim j As Integer
        Dim row As Long
        Dim col As Long
        Dim exists As Boolean
        Dim tables As DrawingTables
        Dim table As DrawingTable
        Dim str As String
        Dim IPTable1 As Object

        tables = dView.Tables
        'Exit if No Tables
        If tables.Count = 0 Then Exit Sub
        I = 1
        For Each table In tables
            If table.NumberOfColumns = 3 Then
                If InStr(1, table.GetCellString(1, 1), "IP") > 0 Then
                    exists = True
                    Exit For
                End If
            End If
            I = I + 1
        Next

        If exists = False Then Exit Sub

        table = tables.Item(I)
        row = table.NumberOfRows
        col = table.NumberOfColumns

        ReDim IPTable1(row, col)

        For I = 1 To row
            For j = 1 To col
                IPTable1(I, j) = Trim(CStr(table.GetCellString(I, j)))
            Next j
        Next I
        'Consolidate Table to include CFG Hole with Multiple IP's
        For I = 1 To row
            If IPTable1(I, 1) = "" And IPTable1(I, 2) = "" And IPTable1(I, 3) <> "" Then
                str = IPTable1(I, 3)
                IPTable1(I - 1, 3) = IPTable1(I - 1, 3) & " " & str
            End If
        Next I

        InstlPropPub.IPtable = IPTable1

    End Sub



    'For Section-16-18************************************
    'PRIVATE: Get Install Item Number form Drawing Sheet-1

    Private Sub GetInstlItemNumber(ByRef AllDrawingViews As DrawingViews)
        'Dim itmNum As ClsText
        'Dim xTxt As ClsText
        Dim I As Integer
        Dim j As Integer
        Dim drawVw As DrawingView
        Dim dTexts As DrawingTexts
        Dim dtxt As DrawingText


        I = 1
        For Each drawVw In AllDrawingViews
            If I > 2 Then
                dTexts = drawVw.Texts
                For j = 1 To dTexts.Count
                    dtxt = dTexts.Item(j)
                    'Find the Install Item Number
                    If InStr(1, dtxt.Text, "VIEW LOOKING") > 0 Then
                        InstlPropPub.PartNumSht1 = New ClsText
                        'Added in Version 4.2
                        If IsNumeric(Right(dtxt.Text, 3)) Then
                            InstlPropPub.PartNumSht1.TextContent = Right(dtxt.Text, 3)
                        Else
                            MsgBox("Installtaion Variant Number is Missing in the 'VIEW LOOKING XXxBOARD' Text as per 16-18 Standard.", vbCritical, "Critical Error...")
                        End If
                        Exit Sub
                    End If
                Next j
            End If
            I = I + 1
            dTexts = Nothing
        Next
    End Sub


End Class
