Option Explicit

'=========================================================
' Youzhi PPT Kit
' Ribbon Callback
'=========================================================

Public gRibbon As Office.IRibbonUI


Public Sub RibbonOnLoad(ribbon As Office.IRibbonUI)

    Set gRibbon = ribbon

End Sub



'=========================================================
' 工具
'=========================================================

Public Sub Ribbon_Test(control As Office.IRibbonControl)

    MsgBox _
        "Youzhi PPT Kit 工作正常！" & vbCrLf & _
        "Version " & YZK_VERSION, _
        vbInformation, _
        "Youzhi PPT Kit"

End Sub



'=========================================================
' 文字
'=========================================================

Public Sub Ribbon_StandardText(control As Office.IRibbonControl)

    YZK_StandardText

End Sub


Public Sub Ribbon_EmphasisText(control As Office.IRibbonControl)

    YZK_EmphasisText

End Sub


Public Sub Ribbon_ListText(control As Office.IRibbonControl)

    YZK_ListText

End Sub


Public Sub Ribbon_Conclusion(control As Office.IRibbonControl)

    YZK_Conclusion

End Sub



'=========================================================
' 图片
'=========================================================

Public Sub Ribbon_PictureShadow(control As Office.IRibbonControl)

    YZK_PictureShadow

End Sub


Public Sub Ribbon_WhitePictureBorder(control As Office.IRibbonControl)

    YZK_WhitePictureBorder

End Sub



'=========================================================
' 方框
'=========================================================

Public Sub Ribbon_RedDashedBox(control As Office.IRibbonControl)

    YZK_RedDashedBox

End Sub


Public Sub Ribbon_RedSolidBox(control As Office.IRibbonControl)

    YZK_RedSolidBox

End Sub


Public Sub Ribbon_BlueDashedBox(control As Office.IRibbonControl)

    YZK_BlueDashedBox

End Sub


Public Sub Ribbon_BlackDashedBox(control As Office.IRibbonControl)

    YZK_BlackDashedBox

End Sub



'=========================================================
' 直线
'=========================================================

Public Sub Ribbon_BlackDashedLine(control As Office.IRibbonControl)

    YZK_BlackDashedLine

End Sub


Public Sub Ribbon_BlackSolidLine(control As Office.IRibbonControl)

    YZK_BlackSolidLine

End Sub


Public Sub Ribbon_GrayDashedLine(control As Office.IRibbonControl)

    YZK_GrayDashedLine

End Sub


Public Sub Ribbon_GraySolidLine(control As Office.IRibbonControl)

    YZK_GraySolidLine

End Sub



'=========================================================
' 箭头
'=========================================================

Public Sub Ribbon_RedArrow(control As Office.IRibbonControl)

    YZK_RedArrow

End Sub


Public Sub Ribbon_BlueArrow(control As Office.IRibbonControl)

    YZK_BlueArrow

End Sub

