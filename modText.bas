Option Explicit


'====================================================
' 公共：给文本框设置字体
'====================================================

Private Sub ApplyBasicTextFormat( _
    ByVal shp As Shape, _
    ByVal fontSize As Single, _
    ByVal fontColor As Long)

    If shp.HasTextFrame <> msoTrue Then Exit Sub
    If shp.TextFrame2.HasText <> msoTrue Then Exit Sub

    With shp.TextFrame2.TextRange.Font

        .NameFarEast = FONT_CN
        .NameAscii = FONT_EN
        .NameComplexScript = FONT_EN

        .Size = fontSize
        .Bold = msoTrue

        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = fontColor

    End With

End Sub


'====================================================
' 1. 标准字体
'
' 中文：微软雅黑
' 英文：Times New Roman
' 18 pt
' 黑色
' 加粗
' 无背景
'====================================================

Public Sub YZK_StandardText()

    Dim sel As Selection
    Dim shp As Shape
    Dim i As Long

    Set sel = ActiveWindow.Selection

    If sel.Type <> ppSelectionShapes Then

        MsgBox "请先选中文本框边框。", _
               vbInformation, _
               "Youzhi PPT Kit"

        Exit Sub

    End If


    For i = 1 To sel.ShapeRange.Count

        Set shp = sel.ShapeRange(i)

        ApplyBasicTextFormat _
            shp, _
            FONT_STANDARD, _
            YZK_Black()

        shp.Fill.Visible = msoFalse

    Next i

End Sub


'====================================================
' 2. 强调字体
'
' 24 pt
' #C00000
' 加粗
' 无背景
'====================================================

Public Sub YZK_EmphasisText()

    Dim sel As Selection
    Dim shp As Shape
    Dim i As Long

    Set sel = ActiveWindow.Selection

    If sel.Type <> ppSelectionShapes Then

        MsgBox "请先选中文本框边框。", _
               vbInformation, _
               "Youzhi PPT Kit"

        Exit Sub

    End If


    For i = 1 To sel.ShapeRange.Count

        Set shp = sel.ShapeRange(i)

        ApplyBasicTextFormat _
            shp, _
            FONT_EMPHASIS, _
            YZK_Red()

        shp.Fill.Visible = msoFalse

    Next i

End Sub


'====================================================
' 3. 列表字体
'
' 18 pt
' 黑色
' 加粗
' 无背景
' 空心方形项目符号 □
'====================================================

Public Sub YZK_ListText()

    Dim sel As Selection
    Dim shp As Shape
    Dim i As Long

    Set sel = ActiveWindow.Selection

    If sel.Type <> ppSelectionShapes Then

        MsgBox "请先选中文本框边框。", _
               vbInformation, _
               "Youzhi PPT Kit"

        Exit Sub

    End If


    For i = 1 To sel.ShapeRange.Count

        Set shp = sel.ShapeRange(i)

        ApplyBasicTextFormat _
            shp, _
            FONT_LIST, _
            YZK_Black()

        shp.Fill.Visible = msoFalse


        If shp.HasTextFrame = msoTrue Then

            If shp.TextFrame2.HasText = msoTrue Then

                With shp.TextFrame2.TextRange.ParagraphFormat.Bullet

                    .Visible = msoTrue

                    'Unicode 空心方框 □
                    .Character = 9633

                End With

            End If

        End If

    Next i

End Sub


'====================================================
' 4. 结论字体
'
' 24 pt
' 白色
' 蓝底 #2E54A1
' 加粗
' 水平居中
' 垂直居中
'====================================================

Public Sub YZK_Conclusion()

    Dim sel As Selection
    Dim shp As Shape
    Dim i As Long

    Set sel = ActiveWindow.Selection

    If sel.Type <> ppSelectionShapes Then

        MsgBox "请先选中结论文本框。", _
               vbInformation, _
               "Youzhi PPT Kit"

        Exit Sub

    End If


    For i = 1 To sel.ShapeRange.Count

        Set shp = sel.ShapeRange(i)

        ApplyBasicTextFormat _
            shp, _
            FONT_CONCLUSION, _
            YZK_White()


        If shp.HasTextFrame = msoTrue Then

            If shp.TextFrame2.HasText = msoTrue Then

                shp.TextFrame2.TextRange.ParagraphFormat.Alignment = _
                    msoAlignCenter

                shp.TextFrame2.VerticalAnchor = _
                    msoAnchorMiddle

            End If

        End If


        With shp.Fill

            .Visible = msoTrue
            .Solid
            .ForeColor.RGB = YZK_Blue()
            .Transparency = 0

        End With


        shp.Line.Visible = msoFalse

    Next i

End Sub
