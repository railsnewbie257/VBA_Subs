<pre>
'-----------------------------------------------
'
' Standard Colors
'
Public currentColor As Long
'
'
Public Const RUST = 192
Public Const RED = 255
Public Const HILITERED = 393372
Public Const ORANGE = 49407
Public Const YELLOW = 65535
Public Const LIGHTGREEN = 5296274
Public Const GREEN = 5287936
Public Const LIGHTBLUE = 15773696
Public Const BLUE = 12611584
Public Const BLACK = 10
Public Const DARKBLUE = 6299648
Public Const PURPLE = 10498160
Public Const PINK = 13395711

Public Const NOCOLOR = 16777215
Public Const LIGHTPINK = 13421823 ' 0.599993896298105
Public Const HILITEPINK = 13551615

Public Const GREY = 9868950
Public Const LIGHTGREY = 14540253

Sub ColorRange(useRange, Optional useColor)

    If useColor = APTCOLLAPSE Then
        With useRange.Interior
            .Pattern = xlPatternLinearGradient
            .Gradient.Degree = 270
            .Gradient.ColorStops.Clear
        End With
        With useRange.Interior.Gradient.ColorStops.Add(0)
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        With useRange.Interior.Gradient.ColorStops.Add(1)
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0
        End With
    ElseIf useColor = GREYSPECKLE Then
        With useRange.Interior
            .Pattern = xlGray16
            .PatternColorIndex = xlAutomatic
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    ElseIf useColor = LIGHTGREY Then
        With useRange.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = useColor
            .ThemeColor = xlThemeColorDark2
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    Else '------------------------------------- regular color
        With useRange.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = useColor
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End If
End Sub

</pre>
