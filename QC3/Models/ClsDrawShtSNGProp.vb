Imports DRAFTINGITF

Public Class ClsDrawShtSNGProp

    Public GENprop As ClsDrawShtGENProp
    Public PartNumber As ClsText
    Public PartNumberRH As ClsText
    Public ItemNumberList As New Collection
    Public XlistCount As New Collection
    Public FlagNoteList As New Collection
    Public isAssembly As Boolean
    Public isAssemblyRH As Boolean


    'PUBLIC: Run this to Get All Drawing Sheet Related Data

    Public Sub GetSNGDrawingSheetProperties(ByRef DrawingDoc As DrawingDocument)

        Dim drawShts As DrawingSheets
        Dim drawSht As DrawingSheet
        Dim drawViews As DrawingViews
        Dim drawView As DrawingView
        Dim drawTexts As DrawingTexts

        drawShts = DrawingDoc.Sheets
        drawSht = drawShts.Item("Sheet.1")
        drawViews = drawSht.Views
        drawView = drawViews.Item(2)
        drawTexts = drawView.Texts

        'Get All the item Numbers
        GetAllItemNumbers(drawViews)
        'Count the Number of Items
        GetItemCountbyText()
        'Get All the Flag Notes Form Drawing
        GetAllFlagNotes(drawViews)

    End Sub


    'PRIVATE: Get All Item Numbers in Drawing Sheet

    Private Sub GetAllItemNumbers(ByRef AllDrawingViews As DrawingViews)
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
            If I > 2 Then
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
                            If itmNum.TextSx <= GENprop.DRWlimitX And itmNum.TextSY <= GENprop.DRWlimitX _
                         And itmNum.TextSx > 0 And itmNum.TextSY > 0 Then
                                itmNum.GetTextZone(GENprop.DRWsheetNumber, GENprop.DRWactualSheetSize)
                                'Check if Double Baloon
                                If dtxt.FrameType = 15 Or dtxt.FrameType = 65 Then
                                    handCheck = CInt(itmNum.TextContent) Mod 2
                                    If handCheck = 0 Then
                                        PartNumber = itmNum
                                        If CInt(PartNumber.TextContent) < 200 Then
                                            isAssembly = True
                                        Else
                                            isAssembly = False
                                        End If
                                    Else
                                        PartNumberRH = itmNum
                                        If CInt(PartNumberRH.TextContent) < 200 Then
                                            isAssemblyRH = True
                                        Else
                                            isAssemblyRH = False
                                        End If
                                        'ItemNumberList.Add itmNum
                                    End If
                                Else
                                    If itmNum.ItemCountL = 0 Then
                                        On Error GoTo GoOn
                                        'If Not dtxt.AssociativeElement Is Nothing Then
                                        assosiatedTxt = dtxt.AssociativeElement
                                        itmNum.ItemCountL = assosiatedTxt.Leaders.Count
                                        'End If
                                    End If
GoOn:
                                    ItemNumberList.Add(itmNum)
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
                                XlistCount.Add(xTxt)
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

    Private Function GetAllFlagNotes(ByRef AllDrawingViews As DrawingViews)
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
                            If flagNote.TextSx <= GENprop.DRWlimitX And flagNote.TextSY <= GENprop.DRWlimitX _
                         And flagNote.TextSx > 0 And flagNote.TextSY > 0 Then
                                flagNote.GetTextZone(GENprop.DRWsheetNumber, GENprop.DRWactualSheetSize)
                                FlagNoteList.Add(flagNote)
                            End If
                        End If
                    End If
                Next j
            End If
            I = I + 1
            dTexts = Nothing
        Next
    End Function

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

        If ItemNumberList.Count > 0 And XlistCount.Count > 0 Then
            For Each iTxt In ItemNumberList
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
                For Each xTxt In XlistCount
                    If xTxt.TextSx > x1Bound And xTxt.TextSx < x2Bound Then
                        If xTxt.TextSY > y1Bound And xTxt.TextSY < y2Bound Then
                            iTxt.ItemCountT = Left(xTxt.TextContent, Len(xTxt.TextContent) - 1)
                        End If
                    End If
                Next
                iTxt.SetTotalItemCount()
            Next
        ElseIf ItemNumberList.Count > 0 And XlistCount.Count = 0 Then
            For Each iTxt In ItemNumberList
                iTxt.SetTotalItemCount()
            Next
        End If
    End Sub


End Class
