Option Explicit

'=========================================================
' Youzhi PPT Kit
' modGraphics
'
' 图形处理模块
'
' 图片：
'   1. 加阴影
'   2. 白底图片边框
'
' 方框：
'   3. 红色虚线框
'   4. 红色实线框
'   5. 蓝色虚线框
'   6. 黑色虚线框
'
' 直线：
'   7. 黑色虚线
'   8. 黑色实线
'   9. 灰色虚线
'  10. 灰色实线
'
' 箭头：
'  11. 红色箭头
'  12. 蓝色箭头
'=========================================================



'=========================================================
' 一、基础工具
'=========================================================

Private Function CmToPt(ByVal cmValue As Double) As Single

    CmToPt = CSng(cmValue * 28.3464567)

End Function


Private Function GetCurrentSlide() As Slide

    On Error Resume Next

    If ActiveWindow Is Nothing Then
        Set GetCurrentSlide = Nothing
        Exit Function
    End If

    Set GetCurrentSlide = ActiveWindow.View.Slide

    On Error GoTo 0

End Function


Private Function HasShapeSelection() As Boolean

    HasShapeSelection = False

    If ActiveWindow Is Nothing Then Exit Function

    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        HasShapeSelection = True
    End If

End Function



'=========================================================
' 二、公共：创建矩形方框
'=========================================================

Private Sub AddStandardRectangle( _
    ByVal lineColor As Long, _
    ByVal lineWeight As Single, _
    ByVal dashStyle As MsoLineDashStyle)

    Dim sld As Slide
    Dim shp As Shape

    Dim slideW As Single
    Dim slideH As Single

    Dim boxW As Single
    Dim boxH As Single


    Set sld = GetCurrentSlide()

    If sld Is Nothing Then

        MsgBox "请先打开并选中一张幻灯片。", _
               vbInformation, _
               "Youzhi PPT Kit"

        Exit Sub

    End If


    slideW = ActivePresentation.PageSetup.SlideWidth
    slideH = ActivePresentation.PageSetup.SlideHeight

    '默认 10 cm × 5 cm
    boxW = CmToPt(10)
    boxH = CmToPt(5)


    Set shp = sld.Shapes.AddShape( _
        Type:=msoShapeRectangle, _
        Left:=(slideW - boxW) / 2, _
        Top:=(slideH - boxH) / 2, _
        Width:=boxW, _
        Height:=boxH)


    '无填充
    shp.Fill.Visible = msoFalse


    With shp.Line

        .Visible = msoTrue
        .ForeColor.RGB = lineColor
        .Transparency = 0
        .Weight = lineWeight
        .dashStyle = dashStyle

    End With


    shp.Shadow.Visible = msoFalse

    shp.Select

End Sub



'=========================================================
' 三、公共：创建普通直线
'=========================================================

Private Sub AddStandardLine( _
    ByVal lineColor As Long, _
    ByVal lineWeight As Single, _
    ByVal dashStyle As MsoLineDashStyle)

    Dim sld As Slide
    Dim shp As Shape

    Dim slideW As Single
    Dim slideH As Single

    Dim lineLength As Single


    Set sld = GetCurrentSlide()

    If sld Is Nothing Then

        MsgBox "请先打开并选中一张幻灯片。", _
               vbInformation, _
               "Youzhi PPT Kit"

        Exit Sub

    End If


    slideW = ActivePresentation.PageSetup.SlideWidth
    slideH = ActivePresentation.PageSetup.SlideHeight

    '默认长度 8 cm
    lineLength = CmToPt(8)


    Set shp = sld.Shapes.AddLine( _
        BeginX:=slideW / 2 - lineLength / 2, _
        BeginY:=slideH / 2, _
        EndX:=slideW / 2 + lineLength / 2, _
        EndY:=slideH / 2)


    With shp.Line

        .Visible = msoTrue

        .ForeColor.RGB = lineColor
        .Transparency = 0

        .Weight = lineWeight
        .dashStyle = dashStyle

        '普通直线，不带箭头
        .BeginArrowheadStyle = msoArrowheadNone
        .EndArrowheadStyle = msoArrowheadNone

    End With


    shp.Select

End Sub



'=========================================================
' 四、公共：创建箭头
'=========================================================

Private Sub AddStandardArrow(ByVal lineColor As Long)

    Dim sld As Slide
    Dim shp As Shape

    Dim slideW As Single
    Dim slideH As Single

    Dim arrowLength As Single


    Set sld = GetCurrentSlide()

    If sld Is Nothing Then

        MsgBox "请先打开并选中一张幻灯片。", _
               vbInformation, _
               "Youzhi PPT Kit"

        Exit Sub

    End If


    slideW = ActivePresentation.PageSetup.SlideWidth
    slideH = ActivePresentation.PageSetup.SlideHeight

    arrowLength = CmToPt(8)


    Set shp = sld.Shapes.AddLine( _
        BeginX:=slideW / 2 - arrowLength / 2, _
        BeginY:=slideH / 2, _
        EndX:=slideW / 2 + arrowLength / 2, _
        EndY:=slideH / 2)


    With shp.Line

        .Visible = msoTrue

        .ForeColor.RGB = lineColor
        .Transparency = 0

        '4磅
        .Weight = 4

        .dashStyle = msoLineSolid

        '起点无箭头
        .BeginArrowheadStyle = msoArrowheadNone

        '末端单向箭头
        .EndArrowheadStyle = msoArrowheadTriangle

        '中等箭头
        .EndArrowheadLength = msoArrowheadLengthMedium
        .EndArrowheadWidth = msoArrowheadWidthMedium

    End With


    shp.Select

End Sub



'=========================================================
' 五、图片处理
'=========================================================


'---------------------------------------------------------
' 图片加阴影
'
' 边线：
' 0.75磅
' #E6E6E6
'
' 阴影：
' 外部左下
'---------------------------------------------------------

Public Sub YZK_PictureShadow()

    Dim sel As Selection
    Dim shp As Shape
    Dim i As Long


    If HasShapeSelection() = False Then

        MsgBox "请先选中一张或多张图片。", _
               vbInformation, _
               "Youzhi PPT Kit"

        Exit Sub

    End If


    Set sel = ActiveWindow.Selection


    For i = 1 To sel.ShapeRange.Count

        Set shp = sel.ShapeRange(i)


        With shp.Line

            .Visible = msoTrue
            .ForeColor.RGB = YZK_ColorLightGray()
            .Transparency = 0
            .Weight = 0.75
            .dashStyle = msoLineSolid

        End With


        With shp.Shadow

            .Visible = msoTrue

            .ForeColor.RGB = YZK_ColorBlack()

            .Transparency = 0.65

            .Blur = 4

            '左下
            .OffsetX = -3
            .OffsetY = 3

        End With

    Next i

End Sub



'---------------------------------------------------------
' 白底图片边框
'
' 黑色
' 0.75磅
' 无阴影
'---------------------------------------------------------

Public Sub YZK_WhitePictureBorder()

    Dim sel As Selection
    Dim shp As Shape
    Dim i As Long


    If HasShapeSelection() = False Then

        MsgBox "请先选中一张或多张图片。", _
               vbInformation, _
               "Youzhi PPT Kit"

        Exit Sub

    End If


    Set sel = ActiveWindow.Selection


    For i = 1 To sel.ShapeRange.Count

        Set shp = sel.ShapeRange(i)


        With shp.Line

            .Visible = msoTrue
            .ForeColor.RGB = YZK_ColorBlack()
            .Transparency = 0
            .Weight = 0.75
            .dashStyle = msoLineSolid

        End With


        shp.Shadow.Visible = msoFalse

    Next i

End Sub



'=========================================================
' 六、方框
'=========================================================


'---------------------------------------------------------
' 红色虚线方框
' #C00000
' 4磅
'---------------------------------------------------------

Public Sub YZK_RedDashedBox()

    AddStandardRectangle _
        YZK_ColorRed(), _
        4, _
        msoLineDash

End Sub


'---------------------------------------------------------
' 红色实线方框
'---------------------------------------------------------

Public Sub YZK_RedSolidBox()

    AddStandardRectangle _
        YZK_ColorRed(), _
        4, _
        msoLineSolid

End Sub


'---------------------------------------------------------
' 蓝色虚线方框
' #2E54A1
' 4磅
'---------------------------------------------------------

Public Sub YZK_BlueDashedBox()

    AddStandardRectangle _
        YZK_ColorBlue(), _
        4, _
        msoLineDash

End Sub


'---------------------------------------------------------
' 黑色虚线方框
' 4磅
'---------------------------------------------------------

Public Sub YZK_BlackDashedBox()

    AddStandardRectangle _
        YZK_ColorBlack(), _
        4, _
        msoLineDash

End Sub



'=========================================================
' 七、普通直线
'=========================================================


'---------------------------------------------------------
' 黑色虚线
' 4磅
'---------------------------------------------------------

Public Sub YZK_BlackDashedLine()

    AddStandardLine _
        YZK_ColorBlack(), _
        4, _
        msoLineDash

End Sub


'---------------------------------------------------------
' 黑色实线
' 4磅
'---------------------------------------------------------

Public Sub YZK_BlackSolidLine()

    AddStandardLine _
        YZK_ColorBlack(), _
        4, _
        msoLineSolid

End Sub


'---------------------------------------------------------
' 灰色虚线
'
' #D9D9D9
' 1磅
'---------------------------------------------------------

Public Sub YZK_GrayDashedLine()

    AddStandardLine _
        YZK_ColorGray(), _
        1, _
        msoLineDash

End Sub


'---------------------------------------------------------
' 灰色实线
'
' #D9D9D9
' 4磅
'---------------------------------------------------------

Public Sub YZK_GraySolidLine()

    AddStandardLine _
        YZK_ColorGray(), _
        4, _
        msoLineSolid

End Sub



'=========================================================
' 八、箭头
'=========================================================


'---------------------------------------------------------
' 红色箭头
' #C00000
' 4磅
'---------------------------------------------------------

Public Sub YZK_RedArrow()

    AddStandardArrow YZK_ColorRed()

End Sub


'---------------------------------------------------------
' 蓝色箭头
' #2E54A1
' 4磅
'---------------------------------------------------------

Public Sub YZK_BlueArrow()

    AddStandardArrow YZK_ColorBlue()

End Sub

