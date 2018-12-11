Attribute VB_Name = "Excel_Module"
Option Explicit

Public Enum XlLineStyle
    xlContinuous = 1
    xlDashDot = 4
    xlDashDotDot = 5
    xlSlantDashDot = 13
    xlDash = -4115
    xlDot = -4118
    xlDouble = -4119
    xlLineStyleNone = -4142
End Enum

Public Enum XlBorderWeight
    xlHairline = 1
    xlThin = 2
    xlMedium = -4138
    xlThick = 4
End Enum

Public Enum XlBordersIndex
    xlDiagonalDown = 5
    xlDiagonalUp = 6
    xlEdgeLeft = 7
    xlEdgeTop = 8
    xlEdgeBottom = 9
    xlEdgeRight = 10
    xlInsideVertical = 11
    xlInsideHorizontal = 12
End Enum

Public Enum XlPageOrientation
    xlPortrait = 1
    xlLandscape = 2
End Enum

Public Enum XlPictureAppearance
    xlPrinter = 2
    xlScreen = 1
End Enum

Public Enum XlCopyPictureFormat
    xlBitmap = 2
    xlPicture = -4147
End Enum

Public Const xlLeft = -4131
Public Const xlRight = -4152
Public Const xlCenter = -4108
Public Const xlCenterAcrossSelection = 7

