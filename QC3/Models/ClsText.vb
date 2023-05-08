Public Class ClsText
    Public TextContent As String
    Public TextSx As Double         'Set by Function
    Public TextSY As Double         'Set by Function
    Public TextZone As String
    Public AnchorPosType As Integer
    Public FrameType As Integer
    Public AtBorder As Boolean
    Public ItemCountL As Integer
    Public ItemCountT As Integer
    Public ItemCount As Integer
    Public Qcheck As Boolean
    Public S16Qcheck As Boolean
    Public S16TextZone As String


    'PUBLIC: Convert Text Coordinates from View Coordinates to Sheet Coordinated

    Public Sub MapTextCoordinates(txtX As Double, txtY As Double, viewX As Double, viewY As Double, viewScale As Double, viewAngle As Double)
        Dim x1 As Double
        Dim y1 As Double

        x1 = (txtX * (Math.Cos(viewAngle))) - (txtY * (Math.Sin(viewAngle)))
        y1 = (txtX * (Math.Sin(viewAngle))) + (txtY * (Math.Cos(viewAngle)))
        TextSx = (x1 * viewScale) + viewX
        TextSY = (y1 * viewScale) + viewY
    End Sub

    'PUBLIC: Get Zone Based on the XY-Coordinates

    Public Sub GetTextZone(SheetNum As String, SheetSize As String)
        'Dim ShtNum As String
        Dim xLim As Object
        Dim xLoc As Object
        Dim yLim As Object
        Dim yLoc As Object
        Dim x1 As Long
        Dim x2 As Long
        Dim xi As Long
        Dim y1 As Long
        Dim y2 As Long
        Dim yi As Long
        Dim xdir As String
        Dim ydir As String
        Dim I As Long

        x1 = 20
        y1 = 10
        If SheetSize = "A0" Then
            xLim = {20, 44.5, 94.5, 144.5, 194.5, 244.5, 294.5, 344.5, 394.5, 444.5, 494.5, 544.5, 594.5, 644.5, 694.5, 744.5, 794.5, 844.5, 894.5, 944.5, 994.5, 1044.5, 1094.5, 1144.5, 1180}
            xLoc = {24, 23, 22, 21, 20, 19, 18, 17, 16, 15, 14, 13, 12, 11, 10, "09", "08", "07", "06", "05", "04", "03", "02", "01"}
            yLim = {10, 70.5, 120.5, 170.5, 220.5, 270.5, 320.5, 370.5, 420.5, 470.5, 520.5, 570.5, 620.5, 670.5, 720.5, 770.5, 832}
            yLoc = {"A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "L", "M", "N", "P", "Q", "R"}
            x2 = 1179
            xi = 23
            y2 = 831
            yi = 15
        ElseIf SheetSize = "A1" Then
            xLim = {20, 70.5, 120.5, 170.5, 220.5, 270.5, 320.5, 370.5, 420.5, 470.5, 520.5, 570.5, 620.5, 670.5, 720.5, 770.5, 832}
            xLoc = {16, 15, 14, 13, 12, 11, 10, "09", "08", "07", "06", "05", "04", "03", "02", "01"}
            yLim = {10, 47, 97, 147, 197, 247, 297, 347, 397, 447, 497, 547, 585}
            yLoc = {"A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "L", "M"}
            x2 = 831
            xi = 15
            y2 = 584
            yi = 11
        ElseIf SheetSize = "A2" Then
            xLim = {20, 47, 97, 147, 197, 247, 297, 347, 397, 447, 497, 547, 585}
            xLoc = {12, 11, 10, "09", "08", "07", "06", "05", "04", "03", "02", "01"}
            yLim = {10, 60, 110, 160, 210, 260, 310, 360, 412}
            yLoc = {"A", "B", "C", "D", "E", "F", "G", "H"}
            x2 = 584
            xi = 11
            y2 = 411
            yi = 7
        ElseIf SheetSize = "A3" Then
            xLim = {20, 60, 110, 160, 210, 260, 310, 360, 412}
            xLoc = {"08", "07", "06", "05", "04", "03", "02", "01"}
            yLim = {10, 48.5, 98.5, 148.5, 198.5, 248.5, 288}
            yLoc = {"A", "B", "C", "D", "E", "F"}
            x2 = 411
            xi = 7
            y2 = 287
            yi = 5
        End If
        If TextSx >= x1 And TextSx <= x2 Then
            For I = 0 To xi
                If TextSx >= CDbl(xLim(I)) And TextSx < CDbl(xLim(I + 1)) Then
                    xdir = CStr(xLoc(I))
                    Exit For
                End If
            Next I
        Else
            xdir = "XX"
        End If
        If TextSY >= y1 And TextSY <= y2 Then
            For I = 0 To yi
                If TextSY >= CDbl(yLim(I)) And TextSY < CDbl(yLim(I + 1)) Then
                    ydir = CStr(yLoc(I))
                    Exit For
                End If
            Next I
        Else
            ydir = "X"
        End If
        TextZone = SheetNum & ydir & xdir
    End Sub

    'PUBLIC: Set ItemCount by Adding Leaders Items and Text Items

    Public Sub SetTotalItemCount()
        If ItemCount = 0 Then
            If ItemCountL > 0 And ItemCountT > 0 Then
                ItemCount = ItemCountL - 1 + ItemCountT
            ElseIf ItemCountL = 0 And ItemCountT > 0 Then
                ItemCount = ItemCountT
            ElseIf ItemCountL > 0 And ItemCountT = 0 Then
                ItemCount = ItemCountL
            Else
                ItemCount = 1
            End If
        End If
    End Sub

End Class

